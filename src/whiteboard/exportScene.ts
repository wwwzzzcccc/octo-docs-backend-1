/**
 * Server-side whiteboard (Excalidraw) scene → image serialization (W3).
 *
 * The board has no persisted image on the server; the front-end exports it in
 * the browser with Excalidraw's own canvas renderer. To offer a bot/API export
 * without shipping a headless browser (see the W3 renderer-selection spike), we
 * serialize the decoded `{elements, files}` scene to SVG directly here, and
 * rasterize that SVG to PNG with `@napi-rs/canvas` — a napi library backed by
 * Skia that ships prebuilt musl binaries (no system libraries, no resident
 * process, no network), matching the node:22-alpine deploy image and the same
 * short-lived, sandboxed philosophy as the Typst PDF path.
 *
 * One Skia quirk shapes the PNG path: its SVG engine does not rasterize <image>
 * elements whose href is a data-URI, so the board's embedded images (attachId →
 * data-URI, see renderImage) come out blank when the SVG is decoded. The PNG
 * path therefore composites those image bytes directly onto the canvas with
 * ctx.drawImage after the SVG is drawn (rasterizeSvgToPng), while the SVG output
 * keeps the data-URI embed (browsers render it fine).
 *
 * The element set mirrors `WB_ELEMENT_TYPES` (the Excalidraw v1 subset the
 * front-end binding emits). This is a faithful STRUCTURAL export — shapes,
 * geometry, colors, text and embedded images (resolved by attachId) — rendered
 * with clean geometric strokes rather than Excalidraw's Rough.js hand-drawn
 * look. Fill styles (hachure / cross-hatch) are approximated as a solid fill.
 * Raising stroke fidelity to the sketchy on-canvas style (via roughjs's
 * generator, which is pure-JS and container-safe) is a documented follow-up.
 */

/** Raw decoded scene (from decodeBoardSnapshot in collab/versionRestore.ts). */
export interface SceneInput {
  elements: Array<Record<string, unknown>>
  files: Record<string, Record<string, unknown>>
}

/** Resolved image bytes for an image element, keyed by the element's fileId. */
export interface ResolvedSceneImage {
  mime: string
  bytes: Buffer
}

// ── small typed readers over the untyped element maps ─────────────────────────

function numOf(v: unknown, fallback = 0): number {
  return typeof v === 'number' && Number.isFinite(v) ? v : fallback
}

function strOf(v: unknown, fallback = ''): string {
  return typeof v === 'string' ? v : fallback
}

/** XML/SVG text-content and attribute escaping. */
function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

/** Round to 2dp so the SVG stays compact and deterministic. A coordinate large
 * enough to overflow to a non-finite value under the *100 scaling collapses to
 * 0, so the serializer can never emit width="Infinity" / a malformed attribute. */
function r2(n: number): number {
  const r = Math.round(n * 100) / 100
  return Number.isFinite(r) ? r : 0
}

/** A CSS/SVG color is safe to inline only if it is a simple color token. */
function safeColor(v: unknown, fallback: string): string {
  const s = strOf(v).trim()
  if (!s) return fallback
  if (/^#[0-9a-fA-F]{3,8}$/.test(s)) return s
  if (/^(rgb|rgba|hsl|hsla)\([0-9.,%\s/]+\)$/.test(s)) return s
  if (/^[a-zA-Z]{1,20}$/.test(s)) return s // named color (transparent, white, …)
  return fallback
}

interface Pt {
  x: number
  y: number
}

/** Element `points` are stored relative to the element origin (x, y). */
function absolutePoints(el: Record<string, unknown>): Pt[] {
  const ox = numOf(el.x)
  const oy = numOf(el.y)
  const raw = Array.isArray(el.points) ? (el.points as unknown[]) : []
  const out: Pt[] = []
  for (const p of raw) {
    if (Array.isArray(p) && p.length >= 2) {
      out.push({ x: ox + numOf(p[0]), y: oy + numOf(p[1]) })
    }
  }
  return out
}

interface Bounds {
  minX: number
  minY: number
  maxX: number
  maxY: number
}

/**
 * DoS bounds for the rendered OUTPUT. This endpoint is reader-accessible and the
 * scene geometry is fully attacker-controlled: a single element placed far from
 * the origin (e.g. y = 5e7) makes the scene bounding box — and therefore the
 * raster — arbitrarily large, so a pure GET could otherwise drive the canvas to
 * allocate a multi-GB bitmap and OOM the process. Both the SVG canvas
 * dimensions and the PNG output pixel area are bounded. Defaults here are the
 * standalone fallback; the route overrides them from config.boardExport.
 */
export const DEFAULT_MAX_SVG_DIMENSION = 12_000
export const DEFAULT_MAX_PNG_PIXELS = 40_000_000
/**
 * Max DECODED pixel area (width*height) of a *source* image element before it is
 * composited onto the PNG. Distinct from DEFAULT_MAX_PNG_PIXELS: that one caps
 * the OUTPUT canvas, this one caps each SOURCE bitmap the compositor decodes.
 * The source is bounded only by compressed bytes (maxImageBytes, 10MB) upstream,
 * and a highly compressed image (e.g. a 30000x30000 PNG that zips to <10MB)
 * decodes to a multi-GB RGBA bitmap — so `loadImage` on it OOMs the process, and
 * any reader can trigger it through `?format=png` (a decompression bomb). We read
 * the intrinsic dimensions from the file header (see readImagePixelArea) BEFORE
 * decoding, and fall back to a placeholder when they exceed this cap, so the
 * multi-GB allocation never happens. ~40M px ≈ 160MB RGBA peak per image (images
 * are decoded one at a time in the composite loop, so this is the peak, not the
 * sum). Best-effort: an oversize image degrades to a placeholder, never aborts.
 */
export const DEFAULT_MAX_SOURCE_IMAGE_PIXELS = 40_000_000

/** Clamp a raw canvas extent to [1, maxDim]; a non-finite extent collapses to maxDim. */
function clampDim(raw: number, maxDim: number): number {
  if (!Number.isFinite(raw)) return maxDim
  return Math.min(Math.max(1, raw), maxDim)
}

/**
 * Read the intrinsic pixel area (width*height) of an encoded raster from its
 * header bytes, WITHOUT decoding the pixels — so an attacker-supplied bomb never
 * forces the full bitmap allocation just to learn how big it is. Covers the
 * three formats sniffImageMime admits (PNG / JPEG / GIF); the declared header
 * dimensions are exactly what a decoder would allocate (width*height*4 for RGBA).
 * Returns null when the format/marker is unrecognized or the header is truncated,
 * in which case the caller proceeds to decode (best-effort, matching the rest of
 * this path). `mime` narrows the parser; it is the value sniffed from the same
 * bytes, so it is trusted here.
 */
export function readImagePixelArea(bytes: Buffer, mime: string): number | null {
  if (mime === 'image/png') {
    // 8-byte signature, then the first chunk MUST be IHDR: 4-byte length,
    // "IHDR", then width/height as big-endian uint32 at offsets 16 and 20.
    if (bytes.length < 24) return null
    if (bytes.toString('ascii', 12, 16) !== 'IHDR') return null
    return bytes.readUInt32BE(16) * bytes.readUInt32BE(20)
  }
  if (mime === 'image/gif') {
    // Logical Screen Descriptor: width/height as little-endian uint16 at
    // offsets 6 and 8 (right after the 6-byte "GIF87a"/"GIF89a" header).
    if (bytes.length < 10) return null
    return bytes.readUInt16LE(6) * bytes.readUInt16LE(8)
  }
  if (mime === 'image/jpeg') {
    // Walk the marker segments to the first Start-Of-Frame (SOF0..SOF15, minus
    // the non-SOF markers DHT/JPG/DAC at C4/C8/CC): height/width are big-endian
    // uint16 at marker payload offsets 3 and 5. Skip other segments by length.
    let off = 2 // past the SOI (FF D8)
    while (off + 8 < bytes.length) {
      if (bytes[off] !== 0xff) return null // desynced — bail rather than scan blindly
      const marker = bytes[off + 1]!
      const isSof =
        marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc
      const segLen = bytes.readUInt16BE(off + 2)
      if (isSof) return bytes.readUInt16BE(off + 5) * bytes.readUInt16BE(off + 7)
      if (segLen < 2) return null
      off += 2 + segLen
    }
    return null
  }
  return null
}

/** Draw the light dashed placeholder used for an image that can't be composited. */
function drawImagePlaceholder(
  ctx: import('@napi-rs/canvas').SKRSContext2D,
  dx: number,
  dy: number,
  dw: number,
  dh: number,
): void {
  // Mirrors renderImage's SVG placeholder so PNG and SVG exports look identical.
  ctx.fillStyle = '#f1f3f5'
  ctx.fillRect(dx, dy, dw, dh)
  ctx.strokeStyle = '#ced4da'
  ctx.lineWidth = 1
  ctx.setLineDash([4, 4])
  ctx.strokeRect(dx, dy, dw, dh)
}

function isDeleted(el: Record<string, unknown>): boolean {
  return el.isDeleted === true
}

function renderableType(el: Record<string, unknown>): string {
  return strOf(el.type)
}

/** Scene bounding box across every non-deleted element (points included). */
function sceneBounds(elements: Array<Record<string, unknown>>): Bounds | null {
  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity
  let any = false
  const grow = (x: number, y: number): void => {
    any = true
    if (x < minX) minX = x
    if (y < minY) minY = y
    if (x > maxX) maxX = x
    if (y > maxY) maxY = y
  }
  for (const el of elements) {
    if (isDeleted(el) || !renderableType(el)) continue
    const x = numOf(el.x)
    const y = numOf(el.y)
    const w = numOf(el.width)
    let h = numOf(el.height)
    // A text element occasionally arrives without a computed height (the editor
    // normally sets it). Estimate from its line count so multi-line text is not
    // clipped by a too-tight scene box.
    if (renderableType(el) === 'text') {
      const fontSize = numOf(el.fontSize, 20)
      const lineHeight = numOf(el.lineHeight, 1.25) || 1.25
      const lines = strOf(el.text).split('\n').length
      h = Math.max(h, lines * fontSize * lineHeight)
    }
    grow(x, y)
    grow(x + w, y + h)
    for (const p of absolutePoints(el)) grow(p.x, p.y)
  }
  return any ? { minX, minY, maxX, maxY } : null
}

/** Dash pattern for strokeStyle, scaled by stroke width (Excalidraw-ish). */
function dashArray(strokeStyle: string, sw: number): string | null {
  if (strokeStyle === 'dashed') return `${r2(sw * 4)} ${r2(sw * 4)}`
  if (strokeStyle === 'dotted') return `${r2(sw * 1)} ${r2(sw * 3)}`
  return null
}

/** Common presentation attributes shared by the shape primitives. */
function shapeAttrs(el: Record<string, unknown>): string {
  const stroke = safeColor(el.strokeColor, '#1e1e1e')
  const bg = strOf(el.backgroundColor, 'transparent')
  const fill = bg && bg !== 'transparent' ? safeColor(bg, 'none') : 'none'
  const sw = Math.max(0, numOf(el.strokeWidth, 1))
  const parts = [
    `fill="${fill}"`,
    `stroke="${stroke}"`,
    `stroke-width="${r2(sw)}"`,
    'stroke-linecap="round"',
    'stroke-linejoin="round"',
  ]
  const dash = dashArray(strOf(el.strokeStyle, 'solid'), sw || 1)
  if (dash) parts.push(`stroke-dasharray="${dash}"`)
  return parts.join(' ')
}

/** Wrap a fragment in a <g> that applies opacity and rotation about its center. */
function wrap(el: Record<string, unknown>, inner: string): string {
  const opacityRaw = numOf(el.opacity, 100)
  const opacity = Math.max(0, Math.min(1, opacityRaw / 100))
  const angle = numOf(el.angle, 0)
  const attrs: string[] = []
  if (opacity < 1) attrs.push(`opacity="${r2(opacity)}"`)
  if (angle) {
    const cx = numOf(el.x) + numOf(el.width) / 2
    const cy = numOf(el.y) + numOf(el.height) / 2
    const deg = (angle * 180) / Math.PI
    attrs.push(`transform="rotate(${r2(deg)} ${r2(cx)} ${r2(cy)})"`)
  }
  return attrs.length ? `<g ${attrs.join(' ')}>${inner}</g>` : inner
}

function renderRectangle(el: Record<string, unknown>): string {
  const x = numOf(el.x)
  const y = numOf(el.y)
  const w = numOf(el.width)
  const h = numOf(el.height)
  const rounded = el.roundness != null
  const rx = rounded ? Math.min(32, Math.min(Math.abs(w), Math.abs(h)) * 0.25) : 0
  const rxAttr = rx > 0 ? ` rx="${r2(rx)}" ry="${r2(rx)}"` : ''
  return `<rect x="${r2(x)}" y="${r2(y)}" width="${r2(w)}" height="${r2(h)}"${rxAttr} ${shapeAttrs(el)} />`
}

function renderEllipse(el: Record<string, unknown>): string {
  const w = numOf(el.width)
  const h = numOf(el.height)
  const cx = numOf(el.x) + w / 2
  const cy = numOf(el.y) + h / 2
  return `<ellipse cx="${r2(cx)}" cy="${r2(cy)}" rx="${r2(Math.abs(w) / 2)}" ry="${r2(Math.abs(h) / 2)}" ${shapeAttrs(el)} />`
}

function renderDiamond(el: Record<string, unknown>): string {
  const x = numOf(el.x)
  const y = numOf(el.y)
  const w = numOf(el.width)
  const h = numOf(el.height)
  const pts = [
    { x: x + w / 2, y },
    { x: x + w, y: y + h / 2 },
    { x: x + w / 2, y: y + h },
    { x, y: y + h / 2 },
  ]
    .map((p) => `${r2(p.x)},${r2(p.y)}`)
    .join(' ')
  return `<polygon points="${pts}" ${shapeAttrs(el)} />`
}

/** Arrowhead triangle at the last point, oriented along the final segment. */
function arrowHead(points: Pt[], color: string, sw: number): string {
  if (points.length < 2) return ''
  const p = points[points.length - 1]!
  const q = points[points.length - 2]!
  const dx = p.x - q.x
  const dy = p.y - q.y
  const len = Math.hypot(dx, dy)
  if (len === 0) return ''
  const ux = dx / len
  const uy = dy / len
  const size = Math.max(8, sw * 4)
  // Two base points offset perpendicular from a point `size` back along the line.
  const bx = p.x - ux * size
  const by = p.y - uy * size
  const nx = -uy
  const ny = ux
  const half = size * 0.4
  const a = `${r2(p.x)},${r2(p.y)}`
  const b = `${r2(bx + nx * half)},${r2(by + ny * half)}`
  const c = `${r2(bx - nx * half)},${r2(by - ny * half)}`
  return `<polygon points="${a} ${b} ${c}" fill="${color}" stroke="none" />`
}

function renderLinear(el: Record<string, unknown>, withArrow: boolean): string {
  const points = absolutePoints(el)
  if (points.length < 2) return ''
  const stroke = safeColor(el.strokeColor, '#1e1e1e')
  const sw = Math.max(0, numOf(el.strokeWidth, 1)) || 1
  const dash = dashArray(strOf(el.strokeStyle, 'solid'), sw)
  const dashAttr = dash ? ` stroke-dasharray="${dash}"` : ''
  const d = points.map((p) => `${r2(p.x)},${r2(p.y)}`).join(' ')
  const line = `<polyline points="${d}" fill="none" stroke="${stroke}" stroke-width="${r2(sw)}" stroke-linecap="round" stroke-linejoin="round"${dashAttr} />`
  // Default Excalidraw arrows carry an end arrowhead; honor an explicit null.
  const wantHead = withArrow && el.endArrowhead !== null
  return line + (wantHead ? arrowHead(points, stroke, sw) : '')
}

function renderFreedraw(el: Record<string, unknown>): string {
  const points = absolutePoints(el)
  if (points.length < 2) return ''
  const stroke = safeColor(el.strokeColor, '#1e1e1e')
  const sw = Math.max(1, numOf(el.strokeWidth, 1) * 1.5)
  const d = points.map((p) => `${r2(p.x)},${r2(p.y)}`).join(' ')
  return `<polyline points="${d}" fill="none" stroke="${stroke}" stroke-width="${r2(sw)}" stroke-linecap="round" stroke-linejoin="round" />`
}

/** Map Excalidraw fontFamily codes to a generic family Skia/fontconfig can match. */
function fontFamily(code: unknown): string {
  const c = numOf(code, 1)
  if (c === 3) return 'monospace'
  // Keep a concrete CJK-capable fallback first for Skia's SVG renderer. Generic
  // sans-serif alone can resolve to a Latin-only face even after system fonts
  // are registered, producing □ for otherwise valid Chinese text.
  if (c === 2) return 'Arial Unicode MS, Heiti SC, Noto Sans CJK SC, sans-serif'
  return 'Arial Unicode MS, Heiti SC, Noto Sans CJK SC, sans-serif' // 1 = Virgil
}

function renderText(el: Record<string, unknown>): string {
  const raw = strOf(el.text)
  if (!raw) return ''
  const x = numOf(el.x)
  const y = numOf(el.y)
  const fontSize = numOf(el.fontSize, 20)
  const lineHeight = numOf(el.lineHeight, 1.25) || 1.25
  const fill = safeColor(el.strokeColor, '#1e1e1e')
  const align = strOf(el.textAlign, 'left')
  const anchor = align === 'center' ? 'middle' : align === 'right' ? 'end' : 'start'
  const w = numOf(el.width)
  const anchorX = align === 'center' ? x + w / 2 : align === 'right' ? x + w : x
  const lines = raw.split('\n')
  const step = fontSize * lineHeight
  const tspans = lines
    .map((line, i) => {
      // First baseline sits ~one ascent below the element's top edge.
      const dy = i === 0 ? fontSize * 0.9 : step
      return `<tspan x="${r2(anchorX)}" dy="${r2(dy)}">${esc(line)}</tspan>`
    })
    .join('')
  return `<text x="${r2(anchorX)}" y="${r2(y)}" font-family="${fontFamily(el.fontFamily)}" font-size="${r2(fontSize)}" fill="${fill}" text-anchor="${anchor}" xml:space="preserve">${tspans}</text>`
}

function renderImage(el: Record<string, unknown>, image: ResolvedSceneImage | undefined): string {
  const x = numOf(el.x)
  const y = numOf(el.y)
  const w = numOf(el.width)
  const h = numOf(el.height)
  if (!image) {
    // Unresolved binary (dev synthetic host, missing/oversize) → light placeholder.
    return `<rect x="${r2(x)}" y="${r2(y)}" width="${r2(w)}" height="${r2(h)}" fill="#f1f3f5" stroke="#ced4da" stroke-width="1" stroke-dasharray="4 4" />`
  }
  const b64 = image.bytes.toString('base64')
  // `mime` is a hardcoded literal from sniffImageMime today, but escape it at
  // the attribute sink anyway so a future change that sources the mime from
  // attachment/DB metadata can never break out of the data-URI attribute.
  const href = `data:${esc(image.mime)};base64,${b64}`
  return `<image x="${r2(x)}" y="${r2(y)}" width="${r2(w)}" height="${r2(h)}" preserveAspectRatio="none" href="${href}" />`
}

function renderFrame(el: Record<string, unknown>): string {
  const x = numOf(el.x)
  const y = numOf(el.y)
  const w = numOf(el.width)
  const h = numOf(el.height)
  const name = strOf(el.name)
  const label = name
    ? `<text x="${r2(x)}" y="${r2(y - 6)}" font-family="${fontFamily(2)}" font-size="14" fill="#868e96">${esc(name)}</text>`
    : ''
  return `${label}<rect x="${r2(x)}" y="${r2(y)}" width="${r2(w)}" height="${r2(h)}" fill="none" stroke="#adb5bd" stroke-width="1.5" rx="6" ry="6" />`
}

function renderEmbeddable(el: Record<string, unknown>): string {
  const x = numOf(el.x)
  const y = numOf(el.y)
  const w = numOf(el.width)
  const h = numOf(el.height)
  return `<rect x="${r2(x)}" y="${r2(y)}" width="${r2(w)}" height="${r2(h)}" fill="#f8f9fa" stroke="#868e96" stroke-width="1.5" rx="6" ry="6" />`
}

function renderElement(
  el: Record<string, unknown>,
  imagesByFileId: Map<string, ResolvedSceneImage>,
): string {
  switch (renderableType(el)) {
    case 'rectangle':
      return renderRectangle(el)
    case 'ellipse':
      return renderEllipse(el)
    case 'diamond':
      return renderDiamond(el)
    case 'line':
      return renderLinear(el, false)
    case 'arrow':
      return renderLinear(el, true)
    case 'freedraw':
      return renderFreedraw(el)
    case 'text':
      return renderText(el)
    case 'image': {
      const fileId = strOf(el.fileId)
      return renderImage(el, fileId ? imagesByFileId.get(fileId) : undefined)
    }
    case 'frame':
      return renderFrame(el)
    case 'embeddable':
      return renderEmbeddable(el)
    default:
      return '' // unknown/unrenderable type — skip (mirrors repair's whitelist)
  }
}

/**
 * Geometry of the serialized SVG canvas: the viewBox origin and the clamped
 * width/height. The SVG serializer and the PNG rasterizer share it so canvas
 * image compositing maps scene coordinates onto the raster identically.
 */
export interface SceneLayout {
  minX: number
  minY: number
  width: number
  height: number
}

/**
 * Compute the SVG canvas layout (viewBox origin + clamped dimensions) for a
 * scene. Each dimension is clamped to a finite ceiling: without this a scene
 * whose bounding box is, say, 100 x 5e7 (any editor can move one element far
 * down) would emit width/height in the tens of millions — or "Infinity" for a
 * coordinate that overflows — and the raster step would try to allocate a
 * multi-GB bitmap. Beyond the cap the far edge is cropped rather than fatal.
 */
export function computeSceneLayout(
  scene: SceneInput,
  opts: { maxDimension?: number } = {},
): SceneLayout {
  const elements = scene.elements.filter((el) => !isDeleted(el) && renderableType(el))
  const bounds = sceneBounds(elements)
  const pad = 24
  const maxDim = opts.maxDimension && opts.maxDimension > 0 ? opts.maxDimension : DEFAULT_MAX_SVG_DIMENSION
  const minX = bounds ? bounds.minX - pad : 0
  const minY = bounds ? bounds.minY - pad : 0
  const width = clampDim(bounds ? bounds.maxX - bounds.minX + pad * 2 : 100, maxDim)
  const height = clampDim(bounds ? bounds.maxY - bounds.minY + pad * 2 : 100, maxDim)
  return { minX, minY, width, height }
}

/**
 * Serialize a decoded board scene to a standalone SVG document string.
 * `imagesByFileId` carries the resolved bytes for image elements (best-effort;
 * a missing entry renders a light placeholder). A scene with no renderable
 * element yields a small blank white canvas (a valid, openable export).
 */
export function serializeSceneToSvg(
  scene: SceneInput,
  imagesByFileId: Map<string, ResolvedSceneImage> = new Map(),
  opts: { maxDimension?: number } = {},
): string {
  const elements = scene.elements.filter((el) => !isDeleted(el) && renderableType(el))
  const { minX, minY, width, height } = computeSceneLayout(scene, opts)

  const body = elements.map((el) => wrap(el, renderElement(el, imagesByFileId))).join('')

  return (
    `<svg xmlns="http://www.w3.org/2000/svg" ` +
    `width="${r2(width)}" height="${r2(height)}" ` +
    `viewBox="${r2(minX)} ${r2(minY)} ${r2(width)} ${r2(height)}">` +
    `<rect x="${r2(minX)}" y="${r2(minY)}" width="${r2(width)}" height="${r2(height)}" fill="#ffffff" />` +
    body +
    `</svg>`
  )
}

/** Register the system font directory once so the alpine deploy image resolves
 * CJK / emoji glyphs (font-noto-cjk / font-noto-emoji, installed in the
 * Dockerfile). On Linux @napi-rs/canvas also reads fontconfig, but loading the
 * directory explicitly makes glyph coverage independent of fontconfig state.
 * Best-effort and memoized: a missing directory (e.g. a dev host) is ignored. */
let systemFontsRegistered = false
function registerSystemFonts(globalFonts: { loadFontsFromDir(dir: string): number }): void {
  if (systemFontsRegistered) return
  systemFontsRegistered = true
  // Load common Linux and macOS system font roots. The local acceptance server
  // runs on macOS; without these directories Skia falls back to a sparse Latin
  // font and renders valid CJK glyphs as tofu boxes in Board PNG exports.
  for (const dir of ['/usr/share/fonts', '/System/Library/Fonts', '/System/Library/Fonts/Supplemental', '/Library/Fonts']) {
    try {
      globalFonts.loadFontsFromDir(dir)
    } catch {
      // ignore — fall back to whatever fonts the platform resolves
    }
  }
}

/**
 * Rasterize an SVG document to PNG bytes with `@napi-rs/canvas` (Skia). Lazy-
 * loaded so the SVG path never depends on the native binary being resolvable,
 * and so a platform without a prebuilt binary fails only on PNG (not on module
 * import).
 *
 * Skia's SVG engine does not render <image> elements whose href is a data-URI,
 * so the board's embedded images (renderImage → data-URI) come out blank when
 * the SVG is decoded. When `composite` is supplied, their decoded bytes are
 * drawn directly onto the raster at the same scene position (scaled/offset by
 * the shared layout, honoring angle/opacity) after the SVG is painted. This is
 * best-effort: an image that fails to decode falls back to the same light
 * placeholder the SVG uses for an unresolved image and never aborts the PNG.
 * Composited images are drawn on top of the full SVG raster, so a shape that
 * overlaps an image renders behind it here (an accepted fidelity trade-off).
 */
export async function rasterizeSvgToPng(
  svg: string,
  opts: {
    fitWidth?: number
    maxPixels?: number
    maxSourceImagePixels?: number
    composite?: {
      elements: Array<Record<string, unknown>>
      imagesByFileId: Map<string, ResolvedSceneImage>
      layout: SceneLayout
    }
  } = {},
): Promise<Buffer> {
  const { createCanvas, loadImage, GlobalFonts } = await import('@napi-rs/canvas')
  registerSystemFonts(GlobalFonts)
  const maxPixels = opts.maxPixels && opts.maxPixels > 0 ? opts.maxPixels : DEFAULT_MAX_PNG_PIXELS
  const maxSourcePixels =
    opts.maxSourceImagePixels && opts.maxSourceImagePixels > 0
      ? opts.maxSourceImagePixels
      : DEFAULT_MAX_SOURCE_IMAGE_PIXELS

  // Decode the serialized SVG once to read its intrinsic pixel size (the
  // document width/height, which the serializer already clamped to a finite
  // ceiling). We do not raster this copy — bitmap-upscaling it to the target
  // width would blur shapes and text; instead we re-render the vector at the
  // target size below so the output stays crisp (matching the old resvg path).
  const probe = await loadImage(Buffer.from(svg))
  const natW = Math.max(1, probe.width)
  const natH = Math.max(1, probe.height)
  // fitWidth pins the output width; height follows the intrinsic aspect ratio.
  // Even after the SVG canvas is clamped, an extreme-aspect scene — or the
  // natural-size path (fitWidth=0) on a large board — can still exceed the pixel
  // budget, so cap the target width so width * height <= maxPixels, downscaling
  // uniformly when it would exceed it.
  const desiredW = opts.fitWidth && opts.fitWidth > 0 ? Math.round(opts.fitWidth) : natW
  const areaCapW = Math.floor(Math.sqrt((maxPixels * natW) / natH))
  const outW = Math.max(1, Math.min(desiredW, areaCapW))
  const outH = Math.max(1, Math.round((natH * outW) / natW))

  // Render the vector at the target size: rewrite only the <svg> tag's width/
  // height (viewBox is untouched, so Skia scales the geometry to fill), then
  // decode that. Reuse the probe when the target already equals the intrinsic
  // size to avoid a second decode.
  const svgImg =
    outW === natW && outH === natH
      ? probe
      : await loadImage(
          Buffer.from(
            svg.replace(
              /(<svg\b[^>]*?)width="[^"]*"\s+height="[^"]*"/,
              `$1width="${outW}" height="${outH}"`,
            ),
          ),
        )

  const canvas = createCanvas(outW, outH)
  const ctx = canvas.getContext('2d')
  // White backdrop, matching the previous rasterizer's `background: 'white'`.
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, outW, outH)
  ctx.drawImage(svgImg, 0, 0, outW, outH)

  const composite = opts.composite
  if (composite) {
    const { minX, minY, width, height } = composite.layout
    const scaleX = width > 0 ? outW / width : 1
    const scaleY = height > 0 ? outH / height : 1
    for (const el of composite.elements) {
      if (isDeleted(el) || renderableType(el) !== 'image') continue
      const fileId = strOf(el.fileId)
      const image = fileId ? composite.imagesByFileId.get(fileId) : undefined
      if (!image) continue // unresolved → the SVG already drew a dashed placeholder
      const dx = (numOf(el.x) - minX) * scaleX
      const dy = (numOf(el.y) - minY) * scaleY
      const dw = numOf(el.width) * scaleX
      const dh = numOf(el.height) * scaleY
      const opacity = Math.max(0, Math.min(1, numOf(el.opacity, 100) / 100))
      const angle = numOf(el.angle, 0)
      ctx.save()
      ctx.globalAlpha = opacity
      if (angle) {
        const cx = dx + dw / 2
        const cy = dy + dh / 2
        ctx.translate(cx, cy)
        ctx.rotate(angle)
        ctx.translate(-cx, -cy)
      }
      // Decompression-bomb guard: read the source's intrinsic pixel area from
      // its header and skip the decode entirely when it exceeds the cap. This
      // MUST happen before loadImage — that call allocates the full RGBA bitmap,
      // so a 30000x30000 (≈3.6GB) source would OOM the process on decode, not
      // after. An unparseable header (area === null) falls through to decode,
      // still bounded by the compressed-byte budget upstream.
      const area = readImagePixelArea(image.bytes, image.mime)
      if (area !== null && area > maxSourcePixels) {
        drawImagePlaceholder(ctx, dx, dy, dw, dh)
        ctx.restore()
        continue
      }
      try {
        const bmp = await loadImage(image.bytes)
        ctx.drawImage(bmp, dx, dy, dw, dh)
      } catch {
        // Decode failed → light dashed placeholder (mirrors renderImage's SVG one).
        drawImagePlaceholder(ctx, dx, dy, dw, dh)
      }
      ctx.restore()
    }
  }

  return canvas.toBuffer('image/png')
}
