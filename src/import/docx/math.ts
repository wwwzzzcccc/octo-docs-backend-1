/**
 * Math reconstruction (commit ⑤): OOXML OMML (m:oMath) → LaTeX.
 *
 * This is the exact reverse of the docx EXPORT math pipeline, which is:
 *
 *   LaTeX ──(MathJax TeX input)──▶ MathML ──(custom mathmlToOmml)──▶ OMML
 *
 * so the import pipeline is:
 *
 *   OMML ──(hand-written ommlToMathmlString)──▶ MathML ──(mathml-to-latex)──▶ LaTeX
 *
 * The recovered LaTeX is stored on the editor's blockMath / inlineMath node
 * (`attrs.latex`) — the same shape the editor and the exporter use, so an
 * export→import round-trip reconstructs the formula. That symmetry is the key
 * validation goal: our own exported docx must round-trip its formulas back to
 * (semantically) the same LaTeX.
 *
 * `ommlToMathmlString` is a self-contained OMML→MathML transformer built on
 * `fast-xml-parser` (already a project dep). It replaces the former
 * `omml2mathml` + classic `xmldom` dependency (and its `get-dom` → `jsdom@9` →
 * `request` transitive chain, 13 security advisories), producing the same
 * MathML shapes so the downstream `mathml-to-latex` output is unchanged. Any
 * failure degrades to null and the caller keeps the raw text / drops the node
 * rather than failing the import.
 */
import { MathMLToLaTeX } from 'mathml-to-latex'
import { parseXmlOrdered, orderedTag, orderedAttr, type OrderedNode } from './xml.js'
import type { PmNode } from './types.js'

const OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

/**
 * Upper bound on the OMML XML handed to a single `ommlToLatex` call. The block
 * walk / inline run loop already tick a shared wall-clock deadline between
 * formulas, but that cannot interrupt the work packed INSIDE one conversion:
 * a single crafted `m:oMath` (deeply nested or with a huge run count) would
 * expand through `ommlToMathmlString` + `MathMLToLaTeX.convert` in one
 * uninterrupted synchronous pass and pin the event loop. Bounding the input
 * size caps that per-formula work; a formula this large is not legitimate
 * round-trip content, so we degrade to null (drop the node + warn) exactly
 * like any other unconvertible formula. 256 KB of OMML XML is orders of
 * magnitude above any real equation.
 */
const MAX_OMML_BYTES = 256 * 1024
const MATH_NS = 'http://www.w3.org/1998/Math/MathML'

/**
 * Convert an OMML `<m:oMath>` XML string to LaTeX. Returns null on any failure
 * (unparseable, unsupported constructs) so the caller can degrade gracefully.
 */
export function ommlToLatex(ommlXml: string): string | null {
  // Reject an oversized single formula before any parse/convert work: this is
  // the per-formula size bound that the between-formula deadline cannot give.
  if (ommlXml.length > MAX_OMML_BYTES) return null
  const xml = ensureMathNs(stripDoctype(ommlXml))
  try {
    // Some Office producers store an already-linearized TeX equation in one
    // m:t run, then surround it with hundreds of empty script/radical/matrix
    // layout boxes. Feeding those boxes through MathML invents a forest of
    // roots and brackets that is not present in the formula. When there is one
    // and only one non-empty text run and it contains unmistakable TeX control
    // words, that run is the authoritative semantic representation.
    const linear = singleLinearLatexRun(xml)
    if (linear) return normalizeRecoveredLatex(linear)
    const mmlStr = ommlToMathmlString(xml)
    if (!mmlStr) return null
    const latex = MathMLToLaTeX.convert(restoreDefaultFences(mmlStr))
    const trimmed = typeof latex === 'string' ? latex.trim() : ''
    if (!trimmed.length) return null
    const normalized = normalizeRecoveredLatex(trimmed)
    // The conversion drops OMML run color; recover a single uniform color
    // (as written by `\color`/`\textcolor` on export) by scanning the raw OMML.
    const color = uniformRunColor(xml)
    return color ? `\\textcolor{#${color}}{${normalized}}` : normalized
  } catch {
    return null
  }
}

function singleLinearLatexRun(xml: string): string | null {
  const runs = [...xml.matchAll(/<m:t\b[^>]*>([\s\S]*?)<\/m:t>/g)]
    .map((match) => {
      const raw = match[1]!.trim()
      // This shortcut accepts only literal TeX text. Any XML entity could decode
      // into markup or change token meaning, so route it through the structured
      // OMML converter rather than decoding/stripping it here.
      return raw.includes('&') || raw.includes('<') || raw.includes('>') ? '' : raw
    })
    .filter(Boolean)
  if (runs.length !== 1 || !/\\[a-zA-Z]+(?:_|\b)/.test(runs[0]!)) return null
  return runs[0]!
}

/* ────────────────────────────────────────────────────────────────────────
 * OMML → MathML transformer (replaces omml2mathml + xmldom)
 * ──────────────────────────────────────────────────────────────────────── */

/**
 * Operator characters that `omml2mathml` tokenised as `<mo>` (everything else
 * non-numeric becomes `<mi>`; a run of ASCII digits becomes one `<mn>`).
 * Range-compressed from that package's `operators.js` table so classification
 * — and therefore the recovered LaTeX — is byte-for-byte identical.
 */
const OPERATOR_RANGES =
  '21-23,26,28-29,2b-2f,3a-40,5b-60,7b-7e,a1,a6,a8,ac,af-b4,b7,b9,bf,d7,f7,2c7,' +
  '2d8-2d9,2dc-2dd,300-338,33f,2000-2006,2009-200a,2010,2012-2014,2016,2020-2022,' +
  '2024-2026,2032-2034,203c,2040,2044,204e-2050,2057,2061-2063,2070,2074-207e,' +
  '2080-208e,20d0-20e1,20e4-20ea,2140,2146,2190-21b3,21b6-21b7,21ba-21e9,21f3-2204,' +
  '2206-220d,220f-221d,2223-223e,2240-22a3,22a5-22bd,22c0-22ff,2305-2306,2308-230b,' +
  '231c-231f,2322-2323,2329-232a,233d,233f,23b0-23b1,23dc-23e0,2502,251c,2524,252c,' +
  '2534,2581,2588,2592,25a0-25a1,25ad,25b2-25b9,25bc-25c5,25ca-25cb,25e6,25eb-25ec,' +
  '25f8-25ff,2605-2606,2772-2773,27d1-27eb,27f0-27ff,2900-2980,2982-299a,29b6-29b9,' +
  '29c0-29c1,29c4-29c8,29ce-29db,29df,29e1-29e6,29eb,29f4-2ae0,2ae2-2af0,2af2-2aff,' +
  '2b04,2b06-2b07,2b0c-2b0d,3014-3019,ff01,ff06,ff08-ff09,ff0b-ff0f,ff1a-ff20,' +
  'ff3b-ff3f,ff5b-ff5d'

const OPERATOR_SET: ReadonlySet<number> = (() => {
  const set = new Set<number>()
  for (const part of OPERATOR_RANGES.split(',')) {
    const dash = part.indexOf('-')
    if (dash === -1) {
      set.add(parseInt(part, 16))
    } else {
      set.add(parseInt(part.slice(0, dash), 16))
      const hi = parseInt(part.slice(dash + 1), 16)
      for (let cp = parseInt(part.slice(0, dash), 16) + 1; cp <= hi; cp++) set.add(cp)
    }
  }
  return set
})()

function isOperatorChar(ch: string): boolean {
  return OPERATOR_SET.has(ch.codePointAt(0)!)
}

/**
 * Escape text destined for MathML element content.
 *
 * The OMML is parsed with `processEntities:false`, so an `<m:t>`'s text still
 * holds any original XML entities verbatim (e.g. `&gt;` from `<m:t>&gt;</m:t>`).
 * Re-escaping the bare `&` there would double-encode it to `&amp;gt;`, which the
 * downstream MathML→LaTeX step then renders as the literal `\& g t ;`. So only
 * escape a bare `&` that does NOT already begin a valid entity (mirrors the
 * escaper in ./xml.ts).
 */
function esc(s: string): string {
  return s
    .replace(/&(?!#[0-9]+;|#x[0-9a-fA-F]+;|[a-zA-Z][a-zA-Z0-9]*;)/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

/** Child element nodes of an ordered-mode node (its single tag slot). */
function kids(n: OrderedNode): OrderedNode[] {
  const t = orderedTag(n)
  if (t == null) return []
  const c = n[t]
  return Array.isArray(c) ? (c as OrderedNode[]) : []
}

/** First direct child element with the given OMML tag name. */
function child(n: OrderedNode, name: string): OrderedNode | null {
  for (const c of kids(n)) if (orderedTag(c) === name) return c
  return null
}

/** All direct child elements with the given OMML tag name. */
function childrenNamed(n: OrderedNode, name: string): OrderedNode[] {
  return kids(n).filter((c) => orderedTag(c) === name)
}

/**
 * Tokenise the text of an `<m:t>` run into MathML `<mi>`/`<mn>`/`<mo>` tokens,
 * mirroring `omml2mathml`'s `parseMT`: a leading digit-run → one `<mn>`; a
 * leading operator char → one `<mo>`; otherwise one `<mi>` (a single char,
 * except inside `m:fName`, where the leading non-operator/non-digit run is one
 * multi-char italic `<mi>`). Whitespace inside identifiers becomes NBSP.
 */
function tokenizeText(text: string, inFName: boolean): string {
  let out = ''
  let rest = text
  while (rest.length) {
    const first = rest[0]!
    if (isOperatorChar(first)) {
      out += `<mo>${esc(first)}</mo>`
      rest = rest.slice(1)
    } else if (/\d/.test(first)) {
      const m = /^\d+/.exec(rest)!
      out += `<mn>${esc(m[0])}</mn>`
      rest = rest.slice(m[0].length)
    } else {
      let take = 1
      if (inFName) {
        let i = 0
        while (i < rest.length && !isOperatorChar(rest[i]!) && !/\d/.test(rest[i]!)) i++
        take = Math.max(1, i)
      }
      const chunk = rest.slice(0, take).replace(/\s/g, '\u00a0')
      const attrs = inFName && take > 1 ? ' mathvariant="italic"' : ''
      out += `<mi${attrs}>${esc(chunk)}</mi>`
      rest = rest.slice(take)
    }
  }
  return out
}

/** Convert a sequence of OMML child nodes to MathML. */
function convertSeq(nodes: OrderedNode[], inFName: boolean): string {
  let out = ''
  for (const n of nodes) {
    const t = orderedTag(n)
    if (t == null || t === '#text') continue
    out += convertNode(n, inFName)
  }
  return out
}

/** Convert the body of an OMML slot (`m:e`/`m:num`/… may hold several runs). */
function convertSlot(slot: OrderedNode | null, inFName = false): string {
  if (!slot) return ''
  return convertSeq(kids(slot), inFName)
}

/** Wrap a slot's MathML in a single `<mrow>` (matching omml2mathml). */
function mrow(slot: OrderedNode | null, inFName = false): string {
  return `<mrow>${convertSlot(slot, inFName)}</mrow>`
}

/** Convert a single OMML element node to MathML. */
function convertNode(n: OrderedNode, inFName: boolean): string {
  const t = orderedTag(n)!
  switch (t) {
    case 'm:r':
      return convertRun(n, inFName)
    case 'm:sSup':
      return convertScript(n, 'sup')
    case 'm:sSub':
      return convertScript(n, 'sub')
    case 'm:sSubSup':
      return convertSubSup(n)
    case 'm:f':
      return convertFraction(n)
    case 'm:rad':
      return convertRadical(n)
    case 'm:nary':
      return convertNary(n)
    case 'm:acc':
      return convertAccent(n)
    case 'm:bar':
      return convertBar(n)
    case 'm:groupChr':
      return convertGroupChr(n)
    case 'm:d':
      return convertDelimiter(n)
    case 'm:m':
      return convertMatrix(n)
    case 'm:func':
      return convertFunc(n)
    case 'm:limUpp':
      return `<mover>${mrow(child(n, 'm:e'))}${mrow(child(n, 'm:lim'))}</mover>`
    case 'm:limLow':
      return `<munder>${mrow(child(n, 'm:e'))}${mrow(child(n, 'm:lim'))}</munder>`
    default:
      // Unknown wrapper: descend into its children (lenient walk).
      return convertSeq(kids(n), inFName)
  }
}

/** Empty Office script boxes are layout debris, not exponents/subscripts. */
function convertScript(n: OrderedNode, kind: 'sub' | 'sup'): string {
  const base = convertSlot(child(n, 'm:e'))
  const script = convertSlot(child(n, `m:${kind}`))
  if (!script) return `<mrow>${base}</mrow>`
  return `<m${kind}>${`<mrow>${base}</mrow>`}<mrow>${script}</mrow></m${kind}>`
}

function convertSubSup(n: OrderedNode): string {
  const base = convertSlot(child(n, 'm:e'))
  const sub = convertSlot(child(n, 'm:sub'))
  const sup = convertSlot(child(n, 'm:sup'))
  if (!sub && !sup) return `<mrow>${base}</mrow>`
  if (!sub) return `<msup><mrow>${base}</mrow><mrow>${sup}</mrow></msup>`
  if (!sup) return `<msub><mrow>${base}</mrow><mrow>${sub}</mrow></msub>`
  return `<msubsup><mrow>${base}</mrow><mrow>${sub}</mrow><mrow>${sup}</mrow></msubsup>`
}

/**
 * Decode the XML entities that survive in `<m:t>` text because the OMML parser
 * runs with `processEntities:false`. Without this, a run holding `&gt;` reaches
 * the char-by-char tokeniser as the four chars `&`,`g`,`t`,`;` and is emitted as
 * garbage; decoding it to the single `>` lets it tokenise as one `<mo>` operator
 * (which `esc` then re-escapes to `&gt;` in the MathML). Mirrors the entity set
 * xmldom decoded implicitly in the old pipeline.
 */
function decodeXmlEntities(s: string): string {
  if (s.indexOf('&') === -1) return s
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h: string) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#([0-9]+);/g, (_, d: string) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/** `<m:r>` run: emit tokenised `<m:t>` text (ignoring run properties). */
function convertRun(n: OrderedNode, inFName: boolean): string {
  let out = ''
  for (const c of kids(n)) {
    if (orderedTag(c) !== 'm:t') continue
    let text = ''
    for (const tc of kids(c)) {
      if (orderedTag(tc) === '#text') {
        const v = (tc as Record<string, unknown>)['#text']
        if (v != null) text += String(v)
      }
    }
    out += tokenizeText(decodeXmlEntities(text), inFName)
  }
  return out
}

/** `<m:f>` fraction. `m:type` noBar → linethickness 0pt; skw/lin → bevelled. */
function convertFraction(n: OrderedNode): string {
  const props = child(n, 'm:fPr')
  const type = props ? orderedAttr(child(props, 'm:type') ?? {}, 'm:val') : null
  const lower = type ? type.toLowerCase() : null
  let attrs = ''
  if (lower === 'nobar') attrs = ' linethickness="0pt"'
  else if (lower === 'skw' || lower === 'lin') attrs = ' bevelled="true"'
  const numerator = convertSlot(child(n, 'm:num'))
  const denominator = convertSlot(child(n, 'm:den'))
  // Word/third-party producers occasionally leave a structurally present but
  // empty denominator while repeatedly wrapping an equation. Rendering those
  // malformed `m:f` nodes as real fractions creates a tower of zero-height
  // denominators; Typst then collapses the equation into overlapping, illegible
  // glyphs. An empty denominator carries no mathematical value, so preserve the
  // numerator directly. Proper fractions (including an empty numerator) retain
  // their normal conversion.
  if (denominator.length === 0) return `<mrow>${numerator}</mrow>`
  return `<mfrac${attrs}><mrow>${numerator}</mrow><mrow>${denominator}</mrow></mfrac>`
}

/** `<m:rad>` radical: msqrt when the degree is hidden/empty, else mroot. */
function convertRadical(n: OrderedNode): string {
  const props = child(n, 'm:radPr')
  const degHide = props ? orderedAttr(child(props, 'm:degHide') ?? {}, 'm:val') : null
  const hidden = degHide === '1' || degHide === 'true' || degHide === 'on'
  const deg = child(n, 'm:deg')
  const degInner = convertSlot(deg)
  if (hidden || !deg || degInner.length === 0) {
    return `<msqrt>${convertSlot(child(n, 'm:e'))}</msqrt>`
  }
  return `<mroot>${mrow(child(n, 'm:e'))}${mrow(deg)}</mroot>`
}

/**
 * `<m:nary>` n-ary operator (∑ / ∏ / ∫ …). limLoc `undOvr` → munderover
 * (limits above/below), else msubsup. Missing sub/sup slots collapse the
 * script wrapper accordingly.
 */
function convertNary(n: OrderedNode): string {
  const props = child(n, 'm:naryPr')
  const chr = props ? orderedAttr(child(props, 'm:chr') ?? {}, 'm:val') : null
  const limLoc = props ? orderedAttr(child(props, 'm:limLoc') ?? {}, 'm:val') : null
  const op = chr && chr.length ? chr : '\u222b'
  const mo = `<mo stretchy="false">${esc(op)}</mo>`

  const subNode = child(n, 'm:sub')
  const supNode = child(n, 'm:sup')
  const hasSub = subNode != null && convertSlot(subNode).length > 0
  const hasSup = supNode != null && convertSlot(supNode).length > 0
  const under = limLoc === 'undOvr'

  let scripted: string
  if (hasSub && hasSup) {
    const tagName = under ? 'munderover' : 'msubsup'
    scripted = `<${tagName}>${mo}${mrow(subNode)}${mrow(supNode)}</${tagName}>`
  } else if (hasSub) {
    const tagName = under ? 'munder' : 'msub'
    scripted = `<${tagName}>${mo}${mrow(subNode)}</${tagName}>`
  } else if (hasSup) {
    const tagName = under ? 'mover' : 'msup'
    scripted = `<${tagName}>${mo}${mrow(supNode)}</${tagName}>`
  } else {
    scripted = mo
  }
  return `<mrow>${scripted}${mrow(child(n, 'm:e'))}</mrow>`
}

/** `<m:acc>` accent → `<mover accent="true">` with the accent char as `<mo>`. */
function convertAccent(n: OrderedNode): string {
  const props = child(n, 'm:accPr')
  const chr = props ? orderedAttr(child(props, 'm:chr') ?? {}, 'm:val') : null
  const accent = chr && chr.length ? chr : '\u0302'
  return `<mover accent="true">${mrow(child(n, 'm:e'))}<mo>${esc(accent)}</mo></mover>`
}

/** `<m:bar>` overline/underline. pos top → mover ¯, bot → munder _. */
function convertBar(n: OrderedNode): string {
  const props = child(n, 'm:barPr')
  const pos = props ? orderedAttr(child(props, 'm:pos') ?? {}, 'm:val') : null
  if (pos === 'bot') {
    return `<munder underaccent="false">${mrow(child(n, 'm:e'))}<mo>\u005f</mo></munder>`
  }
  return `<mover accent="false">${mrow(child(n, 'm:e'))}<mo>\u00af</mo></mover>`
}

/**
 * `<m:groupChr>` grouping character (wide over/under arrow, brace, hat …).
 * pos bot → munder, else mover, with the grouping char as the `<mo>`.
 */
function convertGroupChr(n: OrderedNode): string {
  const props = child(n, 'm:groupChrPr')
  const chr = props ? orderedAttr(child(props, 'm:chr') ?? {}, 'm:val') : null
  const pos = props ? orderedAttr(child(props, 'm:pos') ?? {}, 'm:val') : null
  const glyph = chr && chr.length ? chr : '\u23df'
  if (pos === 'bot') {
    return `<munder accentunder="false">${mrow(child(n, 'm:e'))}<mo>${esc(glyph)}</mo></munder>`
  }
  return `<mover accent="false">${mrow(child(n, 'm:e'))}<mo>${esc(glyph)}</mo></mover>`
}

/**
 * `<m:d>` delimiter → `<mfenced>`. begChr/endChr default to `(`/`)`; the
 * default paren is left OFF the emitted attrs (matching omml2mathml) so
 * restoreDefaultFences re-injects it downstream; a non-paren fence (`[`/`{`/…)
 * is annotated explicitly. Each `<m:e>` becomes one `<mrow>` child.
 */
function convertDelimiter(n: OrderedNode): string {
  const props = child(n, 'm:dPr')
  const beg = props ? orderedAttr(child(props, 'm:begChr') ?? {}, 'm:val') : null
  const end = props ? orderedAttr(child(props, 'm:endChr') ?? {}, 'm:val') : null
  const isDefault = (beg == null || beg === '(') && (end == null || end === ')')
  const attrs = isDefault ? '' : ` open="${esc(beg ?? '(')}" close="${esc(end ?? ')')}"`
  const parts = childrenNamed(n, 'm:e')
    .map((e) => `<mrow>${convertSlot(e)}</mrow>`)
    .join('')
  // Several Office producers wrap an expression in repeated one-cell matrices,
  // each carrying another delimiter. Those matrices have no row/column meaning;
  // preserving every shell creates gigantic, invalid TeX while changing no
  // mathematical content. Remove only this exact no-op shape. Real vectors,
  // matrices and ordinary delimiters are untouched.
  if (childrenNamed(n, 'm:e').length === 1) {
    const e = childrenNamed(n, 'm:e')[0]!
    const visible = kids(e).filter((c) => orderedTag(c) !== 'm:ctrlPr')
    if (
      visible.length === 1 &&
      orderedTag(visible[0]!) === 'm:m' &&
      isOneCellMatrix(visible[0]!) &&
      isRedundantMatrixShell(n)
    ) {
      return convertSlot(child(childrenNamed(visible[0]!, 'm:mr')[0]!, 'm:e'))
    }
  }
  return `<mfenced${attrs}>${parts}</mfenced>`
}

function isOneCellMatrix(n: OrderedNode): boolean {
  const rows = childrenNamed(n, 'm:mr')
  return rows.length === 1 && childrenNamed(rows[0]!, 'm:e').length === 1
}

function isRedundantMatrixShell(n: OrderedNode): boolean {
  const props = child(n, 'm:dPr')
  const beg = props ? orderedAttr(child(props, 'm:begChr') ?? {}, 'm:val') : null
  const end = props ? orderedAttr(child(props, 'm:endChr') ?? {}, 'm:val') : null
  // Keep explicit square/curly/bar delimiters: those carry real semantics even
  // around one cell. The pathological fixtures repeat default parenthesis
  // shells (and occasionally omit the properties entirely).
  return (beg == null || beg === '(') && (end == null || end === ')')
}

/** `<m:m>` matrix → `<mtable><mtr><mtd>…`. Each `<m:e>` in a row is one cell. */
function convertMatrix(n: OrderedNode): string {
  if (isOneCellMatrix(n)) {
    const row = childrenNamed(n, 'm:mr')[0]!
    const cell = childrenNamed(row, 'm:e')[0]!
    const body = convertSlot(cell)
    // A one-cell matrix nested around another structural object is an Office
    // layout shell. Keep a genuine scalar 1×1 matrix, but discard wrappers
    // whose sole cell is another matrix/delimiter/radical/script.
    const structural = kids(cell).filter((c) => {
      const tag = orderedTag(c)
      return tag != null && tag !== '#text' && tag !== 'm:ctrlPr'
    })
    if (structural.length === 1 && orderedTag(structural[0]!) !== 'm:r') return `<mrow>${body}</mrow>`
  }
  let rows = ''
  for (const mr of childrenNamed(n, 'm:mr')) {
    let cells = ''
    for (const e of childrenNamed(mr, 'm:e')) cells += `<mtd>${convertSlot(e)}</mtd>`
    rows += `<mtr>${cells}</mtr>`
  }
  return `<mtable>${rows}</mtable>`
}

/**
 * `<m:func>` function apply → `<mrow><mrow>fName</mrow><mo>&ApplyFunction;</mo>
 * <mrow>arg</mrow></mrow>`; the fName tokenises its letters into one multi-char
 * italic `<mi>` (matches omml2mathml).
 */
function convertFunc(n: OrderedNode): string {
  const name = `<mrow>${convertSlot(child(n, 'm:fName'), true)}</mrow>`
  const arg = mrow(child(n, 'm:e'))
  return `<mrow>${name}<mo>\u2061</mo>${arg}</mrow>`
}

/**
 * Convert an `<m:oMath>` XML string to a MathML `<math>` string, or null when
 * the OMML is unparseable / empty. Sets `<math display>` to block for an
 * `m:oMathPara` wrapper, inline otherwise (matching omml2mathml).
 */
export function ommlToMathmlString(ommlXml: string): string | null {
  let parsed: OrderedNode[]
  try {
    parsed = parseXmlOrdered(Buffer.from(ommlXml, 'utf8'))
  } catch {
    return null
  }

  let oMath: OrderedNode | null = null
  let display = 'inline'
  const findMath = (nodes: OrderedNode[], block: boolean): void => {
    for (const nd of nodes) {
      const t = orderedTag(nd)
      if (t === 'm:oMath') {
        if (!oMath) {
          oMath = nd
          if (block) display = 'block'
        }
        return
      }
      if (t === 'm:oMathPara') {
        findMath(kids(nd), true)
        return
      }
      if (t != null && t !== '#text') findMath(kids(nd), block)
    }
  }
  findMath(parsed, false)
  if (!oMath) return null

  const body = convertSeq(kids(oMath), false)
  if (body.length === 0) return null
  return `<math xmlns="${MATH_NS}" display="${display}">${body}</math>`
}

/* ────────────────────────────────────────────────────────────────────────
 * Post-processing (unchanged: color recovery, LaTeX normalization, fences)
 * ──────────────────────────────────────────────────────────────────────── */

/**
 * Return the single uniform run color (6-hex, no `#`) shared by EVERY colored
 * run in the OMML, or null when there is no color or the coloring is mixed.
 * Every `<m:r>` that carries text must share the same `<w:color w:val>` for us
 * to safely wrap the whole formula in one `\textcolor`; mixed/partial coloring
 * is left uncolored rather than risk corrupting the structure.
 */
function uniformRunColor(ommlXml: string): string | null {
  const runs = ommlXml.match(/<m:r\b[\s\S]*?<\/m:r>/g)
  if (!runs || runs.length === 0) return null
  let color: string | null = null
  let sawTextRun = false
  for (const r of runs) {
    // Only consider runs that actually carry visible text.
    const t = r.match(/<m:t[^>]*>([\s\S]*?)<\/m:t>/)
    if (!t || t[1]!.length === 0) continue
    sawTextRun = true
    const m = r.match(/<w:color\b[^>]*\bw:val="([0-9a-fA-F]{6})"/)
    const c = m ? m[1]!.toUpperCase() : null
    if (c === null) return null // an uncolored text run → not uniform
    if (color === null) color = c
    else if (color !== c) return null // mixed colors → bail
  }
  return sawTextRun ? color : null
}

/**
 * Clean up LaTeX recovered by `mathml-to-latex` so the editor's math renderer
 * (KaTeX) accepts it and it matches the original source more closely.
 *
 * Two classes of fix, both conservative (exact patterns only):
 *  1. `mathml-to-latex` renders standard accent operators (`<m:acc>` → MathML
 *     `<mover>`) as generic `\overset{…}{…}`. Map the well-known ones back to
 *     their dedicated commands (`\dot`, `\ddot`, `\vec`, `\tilde`) so they render
 *     as true accents rather than a raised symbol. (`\hat` already round-trips.)
 *  2. It wraps a multi-letter function identifier under a script in
 *     `\left(…\right)` and space-separates the letters, e.g. `\sin^2\theta`
 *     becomes `\left(s i n\right)^{2}\theta`. Collapse a parenthesised run of
 *     spaced letters that forms a known function name back to `\name`.
 */
export function normalizeRecoveredLatex(latex: string): string {
  let out = latex

  // A literal Office circumflex followed by a braced group is recovered by
  // mathml-to-latex as a bare `\hat` command. Bare accents are invalid TeX and
  // render as red source text. Bind the following balanced visible-brace group
  // as the accent argument; importantly this remains an accent and must never
  // be rewritten as the exponent x^2.
  out = out.replace(
    /\\hat\s+(\\left\\\{(?:[^{}]|\{[^{}]*\})*\\right\\\})/g,
    '\\widehat{$1}',
  )

  // mathml-to-latex uses empty \textrm boxes for Office spacing runs. They do
  // not carry semantics and make otherwise valid matrix expressions enormous.
  out = out.replace(/\\textrm\{\s*\}/g, ' ').replace(/ {2,}/g, ' ')

  // 0. `\hdots` is not a KaTeX/LaTeX command; map it to `\dots`. (Also handled on
  //    export, but repair it here too for docx produced elsewhere.)
  out = out.replace(/\\hdots(?![a-zA-Z])/g, '\\dots')

  // 1. Accent operators emitted as \overset{diacritic}{base} → dedicated command.
  //    `mathml-to-latex` recovers some accents as backslash operators
  //    (`\cdot`, `\rightarrow`, …) and others as the raw Unicode combining mark
  //    (e.g. the macron U+0304 for `\bar`, which then breaks KaTeX as a bare
  //    `\overset{̄}{y}`). Map both forms back to the dedicated accent command.
  const ACCENTS: Array<[RegExp, string]> = [
    [/\\overset\{\\cdot\\cdot\}/g, '\\ddot'],
    [/\\overset\{\\cdot\}/g, '\\dot'],
    [/\\overset\{\\sim\}/g, '\\tilde'],
    // Raw Unicode combining marks (U+03xx) and standalone accent glyphs the
    // converter leaves inside \overset. Order: double-mark before single.
    [/\\overset\{\u0308\}/g, '\\ddot'], // combining diaeresis
    [/\\overset\{\u0307\}/g, '\\dot'], // combining dot above
    [/\\overset\{[\u0304\u0305\u00af]\}/g, '\\bar'], // macron / overline / macron glyph
    [/\\overset\{[\u0303\u02dc\u007e]\}/g, '\\tilde'], // combining tilde / small tilde / ~
  ]
  for (const [re, cmd] of ACCENTS) out = out.replace(re, cmd)

  // 1b. Width-aware accents. An arrow/hat/tilde over a MULTI-token base is a
  //     wide (stretchy) accent — `\overrightarrow` / `\widehat` / `\widetilde`;
  //     over a single token it is the narrow `\vec` / `\hat`. `mathml-to-latex`
  //     collapses both to `\overset{arrow|^}{base}`, so pick the command by the
  //     base width (a base with a space or >1 char token is "wide").
  const widthAware: Array<[RegExp, string, string]> = [
    // arrow (\rightarrow, \to, combining/standalone →) → \overrightarrow | \vec
    [/\\overset\{(?:\\rightarrow|\\to|[\u20d7\u2192])\}\{((?:[^{}]|\{[^{}]*\})*)\}/g, '\\overrightarrow', '\\vec'],
    // circumflex (^, combining) → \widehat | \hat
    [/\\overset\{[\u0302\u005e]\}\{((?:[^{}]|\{[^{}]*\})*)\}/g, '\\widehat', '\\hat'],
  ]
  for (const [re, wide, narrow] of widthAware) {
    out = out.replace(re, (_m, base: string) => {
      const b = base.trim()
      // "Wide" when the base is more than a single character/token (contains a
      // space, or has length > 1 after stripping a single \command).
      const isWide = /\s/.test(b) || b.replace(/^\\[a-zA-Z]+/, '').length > 1 || b.length > 1
      return `${isWide ? wide : narrow}{${base}}`
    })
  }

  // 1c. Overline of a multi-char base comes back as \overset{―|‾}{…}; the
  //     correct construct is \overline{…} (KaTeX cannot render the bar glyph).
  out = out.replace(/\\overset\{[\u2015\u2014\u203e]\}/g, '\\overline')

  // 1d. Underbar (`m:bar pos="bot"`, Word "bar below"). convertBar emits a
  //     `<munder>…<mo>_</mo></munder>`, which `mathml-to-latex` turns into
  //     `\underset{\underline}{X}`. That is invalid KaTeX: `\underline` is a
  //     one-argument macro handed bare as `\underset`'s symbol slot, so KaTeX
  //     raises "Unexpected end of input in a macro argument" and the whole
  //     formula renders as red raw text. The correct construct is `\underline{X}`.
  //     (Mirrors the overline repair at 1c; the pos="top" path already yields a
  //     valid `\bar{X}`.)
  out = out.replace(
    /\\underset\{\\underline\}\{((?:[^{}]|\{[^{}]*\})*)\}/g,
    '\\underline{$1}',
  )

  // 2. Multi-letter operator/function names that the library space-separates
  //    (`l i m`, `s i n`, `m o d`) → the dedicated command. This covers both a
  //    bare occurrence and one wrapped by the library in `\left(…\right)`
  //    (e.g. the operand of `\pmod`). Longest names first so `\arcsin` wins over
  //    `\sin`. `mod` maps to `\bmod`.
  const FUNC_NAMES = [
    'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh', 'coth',
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
    'log', 'ln', 'lg', 'exp', 'lim', 'det', 'gcd', 'max', 'min', 'mod',
  ]
  for (const name of FUNC_NAMES) {
    const spaced = name.split('').join(' ') // e.g. 's i n'
    const cmd = name === 'mod' ? '\\bmod' : `\\${name}`
    // (a) inside \left(…\right): collapse the spaced name but keep the parens,
    //     e.g. `\left(m o d n\right)` → `\left(\bmod n\right)` (balanced).
    out = out.replace(
      new RegExp(`(\\\\left\\()${escapeRe(spaced)}(?![a-zA-Z])`, 'g'),
      `$1${cmd}`,
    )
    // (b) bare, delimited by non-letters (or string ends)
    out = out.replace(
      new RegExp(`(^|[^a-zA-Z])${escapeRe(spaced)}(?![a-zA-Z])`, 'g'),
      (_m, pre: string) => `${pre}${cmd}`,
    )
  }

  // 2b. Unwrap `\left(\fn\right)^{…}` / `_{…}` → `\fn^{…}`: the library parenthesises
  //     a scripted function name, but `\sin^2` needs no parens. Only when the
  //     content is exactly one function command and a script immediately follows.
  out = out.replace(
    /\\left\((\\[a-z]+)\\right\)(?=[\^_])/g,
    (m, cmd: string) => (FUNC_NAMES.includes(cmd.slice(1)) ? cmd : m),
  )

  // 3. Repair a dangling `\right` that has no delimiter argument (the converter
  //    can drop the `.` from `\right.`), which is a hard KaTeX parse error and
  //    renders the whole formula as red raw text. A valid `\right` is followed
  //    by a delimiter: a closer `) ] }`, a bar `|`, an explicit `.`, or a
  //    command like `\rangle`. Only inject `.` when it is followed by none of
  //    those (end of string, or an ordinary token such as `=`, `+`, a letter).
  out = out.replace(/\\right(?![a-zA-Z])(?=\s*$)/g, '\\right.')
  out = out.replace(/\\right(?![a-zA-Z])(\s*)(?![)\]}|.\\])/g, '\\right.$1')

  // 4. Multi-line n-ary limits. A stacked subscript/superscript on a big operator
  //    (\sum_{i=1 \\ j=1}) comes back from OMML as `_{\begin{matrix}…\end{matrix}}`,
  //    which renders the limits at full body size and misaligned. The correct
  //    construct is `\substack{…}`, which typesets small, centered, stacked
  //    limits. Rewrite a matrix that is the DIRECT argument of a `_`/`^` script.
  out = out.replace(
    /([_^])\{\s*\\begin\{matrix\}([\s\S]*?)\\end\{matrix\}\s*\}/g,
    (_m, script: string, body: string) => `${script}{\\substack{${body.trim()}}}`,
  )

  // 5. `\pmod` round-trips through OMML as a parenthesised delimiter wrapping a
  //    `mod` operator: `\left(\bmod n\right)`. In that form the leading spacing
  //    of `\bmod` collapses against `(`, rendering `(modn)`. Restore the proper
  //    `\pmod{n}` (which typesets `(mod n)` with correct spacing).
  out = out.replace(
    /\\left\(\s*\\bmod\s+([^()]*?)\s*\\right\)/g,
    (_m, operand: string) => `\\pmod{${operand.trim()}}`,
  )

  return out.trim()
}

/** Escape a string for use inside a RegExp. */
function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * Restore the default parenthesis fence on `<mfenced>` elements that the
 * transformer left without `open`/`close` attributes.
 *
 * OMML delimiters (`<m:d>` with `<m:begChr>`/`<m:endChr>`) always carry explicit
 * characters, but we emit `<mfenced>` WITHOUT open/close when the delimiter is
 * the MathML default — the parenthesis `(`/`)`. Downstream, `mathml-to-latex`
 * then renders a bare `<mfenced>` around a matrix as `\begin{bmatrix}` (its own
 * default), silently turning `\begin{pmatrix}` into square brackets on
 * round-trip. Injecting the explicit `(`/`)` restores the intended parentheses
 * (→ `\begin{pmatrix}`) without affecting `[`/`{`/`|` fences, which are already
 * annotated.
 */
function restoreDefaultFences(mml: string): string {
  return mml.replace(/<mfenced\b([^>]*)>/g, (whole, attrs: string) => {
    if (/\bopen\s*=/.test(attrs)) return whole // explicit fence already present
    return `<mfenced open="(" close=")"${attrs}>`
  })
}

/** Ensure the m: prefix is bound; wrap/inject the namespace when missing. */
function ensureMathNs(ommlXml: string): string {
  if (ommlXml.includes('xmlns:m=')) return ommlXml
  // Inject the namespace on the first m:-prefixed element.
  return ommlXml.replace(/<m:([a-zA-Z]+)/, `<m:$1 xmlns:m="${OMML_NS}"`)
}

/**
 * Remove any DOCTYPE / internal DTD subset before the per-formula OMML is
 * parsed. The shared fast-xml-parser is hardened (processEntities: false) so it
 * does not expand internal DTD entities, but we still strip the DOCTYPE (and any
 * stray custom entity references) up front so a crafted `<m:oMath>` carrying a
 * DOCTYPE + nested-entity bomb is neutralized before parsing; well-formed OMML
 * from Word never contains a DOCTYPE.
 */
function stripDoctype(xml: string): string {
  // Drop a DOCTYPE declaration, including any `[ ... ]` internal subset.
  let out = xml.replace(/<!DOCTYPE[^>[]*(\[[\s\S]*?\])?\s*>/gi, '')
  // Neutralize any remaining custom entity references (&foo;), keeping the
  // standard XML entities intact. Unresolved custom entities would otherwise
  // survive verbatim into the recovered text; replacing them keeps output bounded.
  out = out.replace(/&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)[A-Za-z_][\w.-]*;/g, '')
  return out
}

/**
 * Build a math PM node from recovered LaTeX. `display` selects block vs inline
 * (matches the editor's blockMath / inlineMath nodes, attr `latex`).
 */
export function mathNode(latex: string, display: boolean): PmNode {
  return display
    ? { type: 'blockMath', attrs: { latex } }
    : { type: 'inlineMath', attrs: { latex } }
}
