/**
 * Embedded-image handling (commit ⑥): OOXML w:drawing → image node, then an
 * async upload step that turns the embedded bytes into a real attachId.
 *
 * Split in two because the body walker is pure/sync and has no docId / uploader
 * context:
 *
 *   1. WALKER (sync)   emits an image node carrying the drawing's relationship
 *      id in a private `_embedRel` attr, plus any alt text.
 *   2. UPLOAD (async)  resolveImages() walks the produced doc, and for every
 *      image node: resolves _embedRel -> media entry, validates the bytes by
 *      magic number (never trust the extension), uploads to object storage via
 *      the presigned PUT, registers a doc_attachment row, and rewrites the node
 *      to { type:'image', attrs:{ attachId, alt? } }.
 *
 * Degradation (red猫's requirement): a media entry that is missing, too large,
 * of an unknown / disallowed type, or whose upload fails does NOT sink the
 * import — the node degrades to a real fileAttachment block atom carrying the
 * original file name. One bad image never takes down the document.
 */
import type { PmNode } from './types.js'
import type { ExtractedEntry } from './extract.js'
import { sanitizeSvg } from '../../util/sanitizeSvg.js'

/** image magic numbers → mime. Extension is never trusted. */
const IMAGE_MAGIC: Array<{ mime: string; test: (b: Buffer) => boolean }> = [
  { mime: 'image/png', test: (b) => b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47 },
  { mime: 'image/jpeg', test: (b) => b.length > 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff },
  { mime: 'image/gif', test: (b) => b.length > 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 },
  {
    mime: 'image/webp',
    test: (b) =>
      b.length > 12 &&
      b.toString('ascii', 0, 4) === 'RIFF' &&
      b.toString('ascii', 8, 12) === 'WEBP',
  },
  {
    mime: 'image/bmp',
    test: (b) => b.length > 2 && b[0] === 0x42 && b[1] === 0x4d,
  },
]

/** Sniff an image mime by magic number, or null if it isn't a known image. */
export function sniffImageMime(bytes: Buffer): string | null {
  for (const { mime, test } of IMAGE_MAGIC) if (test(bytes)) return mime
  // This is only a candidate check. SVG has no magic number, so the import
  // path MUST pass candidates through sanitizeSvg before storage.
  const prefix = bytes.subarray(0, Math.min(bytes.length, 4096)).toString('utf8').replace(/^\uFEFF/, '').trimStart()
  if (/<svg(?:\s|>)/i.test(prefix) && !/<html(?:\s|>)/i.test(prefix)) return 'image/svg+xml'
  return null
}

/** The private attr the walker uses to reference a drawing's embed rId. */
export const EMBED_REL_ATTR = '_embedRel'
export const FALLBACK_EMBED_REL_ATTR = '_fallbackEmbedRel'

/** Build the placeholder image node the walker emits (pre-upload). */
export function imagePlaceholder(embedRelId: string, alt: string | null, fallbackRelId?: string | null): PmNode {
  const attrs: Record<string, unknown> = { [EMBED_REL_ATTR]: embedRelId }
  if (fallbackRelId && fallbackRelId !== embedRelId) attrs[FALLBACK_EMBED_REL_ATTR] = fallbackRelId
  if (alt) attrs.alt = alt
  return { type: 'image', attrs }
}

/** Context the async upload step needs (injected by the route, not imported). */
export interface MediaUploadCtx {
  /** docId the attachments belong to. */
  docId: string
  /** uid recorded as the attachment creator. */
  uid: string
  /** Max bytes for a single embedded image. */
  maxImageBytes: number
  /**
   * Upload the bytes and register the attachment; resolves to the new attachId.
   * Supplied by the route so this module stays free of storage/db imports.
   */
  upload: (input: { bytes: Buffer; mime: string; fileName: string }) => Promise<string>
}

/** rId -> media entry, built from rels (image relationships) + extracted media. */
export type MediaResolver = (embedRelId: string) => ExtractedEntry | null

/**
 * Walk the doc and replace every image placeholder with a real image node
 * (attachId) or a degraded fileAttachment/text node. Mutates + returns the doc,
 * collecting warnings for degraded images. Sequential (one upload at a time),
 * best-effort.
 */
export async function resolveImages(
  doc: PmNode,
  resolveMedia: MediaResolver,
  ctx: MediaUploadCtx,
  warnings: string[],
): Promise<PmNode> {
  await walk(doc, resolveMedia, ctx, warnings)
  return doc
}

async function walk(
  node: PmNode,
  resolveMedia: MediaResolver,
  ctx: MediaUploadCtx,
  warnings: string[],
): Promise<void> {
  if (!node.content) return
  for (let i = 0; i < node.content.length; i++) {
    const child = node.content[i]!
    if (child.type === 'image' && child.attrs && EMBED_REL_ATTR in child.attrs) {
      node.content[i] = await resolveOne(child, resolveMedia, ctx, warnings)
    } else {
      await walk(child, resolveMedia, ctx, warnings)
    }
  }
}

async function resolveOne(
  image: PmNode,
  resolveMedia: MediaResolver,
  ctx: MediaUploadCtx,
  warnings: string[],
): Promise<PmNode> {
  const relId = String(image.attrs![EMBED_REL_ATTR])
  const fallbackRel = typeof image.attrs![FALLBACK_EMBED_REL_ATTR] === 'string'
    ? String(image.attrs![FALLBACK_EMBED_REL_ATTR])
    : null
  const alt = typeof image.attrs!.alt === 'string' ? (image.attrs!.alt as string) : null
  const media = resolveMedia(relId)

  const degrade = (reason: string): PmNode => {
    warnings.push(`an image could not be imported (${reason})`)
    const fileName = media?.name.split('/').pop() ?? alt ?? 'image'
    // Degrade to a real fileAttachment block atom (schema node 14) carrying the
    // original file name — a proper block node, never smuggled into inline flow.
    return {
      type: 'fileAttachment',
      attrs: {
        attachId: null,
        fileName,
        mime: media ? sniffImageMime(media.data) : null,
        sizeBytes: media ? media.data.length : null,
      },
    }
  }

  const fallback = async (reason: string): Promise<PmNode> => {
    if (!fallbackRel) return degrade(reason)
    warnings.push(`a native SVG could not be imported (${reason}); its raster fallback was used`)
    return resolveOne(imagePlaceholder(fallbackRel, alt), resolveMedia, ctx, warnings)
  }

  if (!media) return fallback('source missing')
  if (media.data.length > ctx.maxImageBytes) return fallback('too large')

  const mime = sniffImageMime(media.data)
  if (!mime) return fallback('unrecognised image format')

  let bytes = media.data
  if (mime === 'image/svg+xml') {
    try {
      // OOXML is untrusted input. Never persist original SVG bytes: active XML,
      // external references and scripts must be removed by the shared sanitizer.
      bytes = sanitizeSvg(media.data)
    } catch {
      return fallback('invalid or unsafe SVG')
    }
  }
  try {
    const fileName = media.name.split('/').pop() ?? 'image'
    const attachId = await ctx.upload({ bytes, mime, fileName })
    const attrs: Record<string, unknown> = { attachId }
    if (alt) attrs.alt = alt
    return { type: 'image', attrs }
  } catch {
    return fallback('upload failed')
  }
}
