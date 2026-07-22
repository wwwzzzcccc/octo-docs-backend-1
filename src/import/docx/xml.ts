/**
 * XXE-safe XML parsing for OOXML parts.
 *
 * fast-xml-parser does NOT resolve external entities or DTDs by default — there
 * is no network/file fetch, so XXE / billion-laughs external expansion is off by
 * construction. We additionally disable internal DOCTYPE entity processing and
 * keep attribute values as-is (no HTML entity re-expansion) so a crafted part
 * can't smuggle expanded content past the size ceilings enforced upstream.
 *
 * OOXML is attribute-heavy and namespace-prefixed (`w:`, `m:`, `a:`, `r:`);
 * we preserve prefixes (they are semantically meaningful here) and surface
 * attributes under a stable `@_` prefix.
 */
import { XMLParser } from 'fast-xml-parser'

/** A parsed XML node tree (loosely typed — the walkers narrow as they descend). */
export type XmlNode = Record<string, unknown>

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  // Preserve `w:`, `m:`, `a:`, `r:` prefixes — they disambiguate OOXML elements.
  removeNSPrefix: false,
  // Keep every repeated child as an array element even when it appears once, so
  // the walkers never have to branch on "single object vs array".
  isArray: () => true,
  // Do NOT process DTDs / internal entities (defence in depth vs entity bombs).
  processEntities: false,
  htmlEntities: false,
  // Preserve whitespace-only text (OOXML uses xml:space="preserve" runs).
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
})

/** Parse an OOXML part buffer into a plain object tree. Never throws on entities. */
export function parseXml(buffer: Buffer): XmlNode {
  return parser.parse(buffer.toString('utf8')) as XmlNode
}

// A second parser in preserveOrder mode, used only where INTRA-element child
// ORDER is semantically significant (a run mixing w:t / w:br / w:tab, or a body
// interleaving w:p / w:tbl). preserveOrder yields an ordered array of
// single-key objects: [{ 'w:t': [...] , ':@': {attrs} }, { 'w:br': [] }, ...].
const orderedParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: false,
  preserveOrder: true,
  processEntities: false,
  htmlEntities: false,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  // Structured OMML can legitimately be very deep: the real acceptance file
  // 11-docx-fixed.docx is valid OOXML with a measured XML element depth of 361.
  // Keep a finite ceiling (512, with ~40% headroom) for hostile input while
  // accepting that document and our own deeply nested export format.
  maxNestedTags: 512,
})

/** An ordered-mode node: exactly one tag key plus an optional `:@` attrs group. */
export type OrderedNode = Record<string, unknown>

/** Parse in order-preserving mode (see orderedParser). Returns the top array. */
export function parseXmlOrdered(buffer: Buffer): OrderedNode[] {
  return orderedParser.parse(buffer.toString('utf8')) as OrderedNode[]
}

/** The single tag name of an ordered-mode node (ignoring the `:@` attrs slot). */
export function orderedTag(node: OrderedNode): string | null {
  for (const k of Object.keys(node)) if (k !== ':@') return k
  return null
}

/** Read a `@_`-prefixed attribute from an ordered-mode node's `:@` group. */
export function orderedAttr(node: OrderedNode, name: string): string | null {
  const group = node[':@']
  if (!group || typeof group !== 'object') return null
  const v = (group as Record<string, unknown>)[`@_${name}`]
  return v == null ? null : String(v)
}

/** Escape text for XML content / attribute values.
 *
 * The parser runs with `processEntities:false`, so a node's `#text` still holds
 * any original XML entities verbatim (e.g. `&lt;` from `<m:t>&lt;</m:t>`).
 * Re-escaping the bare `&` there would double-encode it to `&amp;lt;`, which the
 * downstream OMML→LaTeX converter turns into the literal characters `& l t ;`.
 * So we escape a `&` only when it does NOT already begin a valid entity. */
function xmlEscape(s: string): string {
  return s
    .replace(/&(?!#[0-9]+;|#x[0-9a-fA-F]+;|[a-zA-Z][a-zA-Z0-9]*;)/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

/**
 * Serialise an ordered-mode node (or array of them) back to an XML string. Used
 * to hand an m:oMath subtree to the OMML→MathML converter, which needs real XML
 * rather than the parsed object tree. Attributes ride in the `:@` group.
 */
export function orderedToXml(node: OrderedNode | OrderedNode[]): string {
  const nodes = Array.isArray(node) ? node : [node]
  let out = ''
  for (const n of nodes) {
    const tag = orderedTag(n)
    if (tag == null || tag === '#text') {
      const text = (n as Record<string, unknown>)['#text']
      if (text != null) out += xmlEscape(String(text))
      continue
    }
    const children = n[tag]
    const attrsGroup = (n[':@'] as Record<string, unknown> | undefined) ?? {}
    let attrStr = ''
    for (const [k, v] of Object.entries(attrsGroup)) {
      if (k.startsWith('@_')) attrStr += ` ${k.slice(2)}="${xmlEscape(String(v))}"`
    }
    const inner = Array.isArray(children) ? orderedToXml(children as OrderedNode[]) : ''
    if (inner) out += `<${tag}${attrStr}>${inner}</${tag}>`
    else out += `<${tag}${attrStr}/>`
  }
  return out
}

/** Coerce a possibly-array/undefined child slot into a definite array. */
export function asArray<T = XmlNode>(value: unknown): T[] {
  if (value == null) return []
  return Array.isArray(value) ? (value as T[]) : [value as T]
}

/** Read a `@_`-prefixed attribute as a string, or null. */
export function attr(node: unknown, name: string): string | null {
  if (!node || typeof node !== 'object') return null
  const v = (node as Record<string, unknown>)[`@_${name}`]
  return v == null ? null : String(v)
}
