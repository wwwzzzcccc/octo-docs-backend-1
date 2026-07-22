/**
 * Shared types + safety helpers for the docx-import pipeline.
 *
 * PmNode mirrors the ProseMirror-JSON shape the editor's schema uses (and the
 * exact shape the frontend markdown importer emits), so the doc we build here
 * drops straight into `prosemirrorToYDoc` / setContent with no adaptation.
 */

/** ProseMirror-JSON node. */
export interface PmNode {
  type: string
  attrs?: Record<string, unknown>
  content?: PmNode[]
  text?: string
  marks?: PmMark[]
}

export interface PmMark {
  type: string
  attrs?: Record<string, unknown>
}

/**
 * Href scheme whitelist — byte-aligned to the backend export side
 * (renderHtml.isSafeHref). Permits http(s)/mailto/tel + relative/anchor; drops
 * javascript:/data:/vbscript: and control-char ("java\tscript:") bypasses.
 * Returns the original href when safe, else null.
 */
export function safeHref(raw: string | null | undefined): string | null {
  if (!raw) return null
  // Relationship targets occasionally arrive already XML-escaped in source
  // PM (`...?a=1&amp;b=2`). The DOCX writer escapes XML again, so without
  // canonicalization every round-trip grows another `amp;`. Decode only the
  // predefined amp entity (possibly repeated) before validating; this is URL
  // attribute normalization, not general HTML/entity parsing.
  const href = raw.trim().replace(/&(?:amp;)+/gi, '&')
  if (!href) return null
  // eslint-disable-next-line no-control-regex
  const cleaned = href.replace(/[\u0000-\u0020]+/g, '').toLowerCase()
  const scheme = cleaned.match(/^([a-z][a-z0-9+.-]*):/)
  if (!scheme) return href // relative / #anchor — safe
  return ['http', 'https', 'mailto', 'tel'].includes(scheme[1] ?? '') ? href : null
}

/** CSS colour whitelist (mirrors renderHtml.isSafeCssColor) — single token only. */
export function safeCssColor(v: string | null | undefined): string | null {
  if (!v) return null
  const s = v.trim().toLowerCase()
  if (/^#[0-9a-f]{3,8}$/.test(s)) return s
  if (/^rgba?\(\s*[\d.%,\s/]+\)$/.test(s)) return s
  if (/^hsla?\(\s*[\d.%,\s/deg]+\)$/.test(s)) return s
  if (/^[a-z]+$/.test(s)) return s
  return null
}

/** A hex OOXML colour (`RRGGBB`, no `#`) → CSS `#rrggbb`, or null. */
export function ooxmlHexColor(v: string | null | undefined): string | null {
  if (!v) return null
  const s = v.trim().toLowerCase()
  if (s === 'auto') return null
  if (/^[0-9a-f]{6}$/.test(s)) return `#${s}`
  return null
}

/**
 * Decode the five predefined XML entities + numeric character references in a
 * text run. The parser runs with `processEntities: false` as defence-in-depth
 * against DTD-declared entity bombs; that also leaves the harmless predefined
 * entities (&amp; &lt; &gt; &quot; &apos;) and numeric refs (&#NN; / &#xNN;)
 * literal in `w:t` text. Those cannot expand recursively, so decoding ONLY this
 * fixed set here is safe and never reintroduces the entity-bomb risk (custom
 * DTD entities are still never resolved). `&amp;` is applied LAST so a decoded
 * value like `&lt;` is not re-decoded.
 */
export function decodeXmlText(s: string): string {
  if (!s || s.indexOf('&') === -1) return s
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, hex) => codePointOrEmpty(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_m, dec) => codePointOrEmpty(parseInt(dec, 10)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/** Safe String.fromCodePoint: drop out-of-range / invalid code points. */
function codePointOrEmpty(cp: number): string {
  if (!Number.isFinite(cp) || cp < 0 || cp > 0x10ffff) return ''
  // Skip surrogate range and NUL — not valid standalone document text.
  if (cp === 0 || (cp >= 0xd800 && cp <= 0xdfff)) return ''
  try {
    return String.fromCodePoint(cp)
  } catch {
    return ''
  }
}
