/**
 * document.xml body walker (commit ②): OOXML paragraphs + runs → ProseMirror.
 *
 * We parse in ORDER-PRESERVING mode (parseXmlOrdered) throughout, because OOXML
 * is order-significant at every level the importer cares about: body blocks
 * (w:p / w:tbl interleave), paragraph inline (runs / hyperlinks / bookmarks),
 * and even within a single run (w:t / w:br / w:tab). Order-preserving parsing is
 * also the right foundation for the list (③) and table (④) commits.
 *
 * Scope of THIS commit: the text backbone —
 *   - w:p                    → paragraph, or heading (via pStyle Heading1..6)
 *   - w:pPr/w:jc             → textAlign attr
 *   - w:r + w:rPr            → text nodes with marks
 *   - w:b/i/u/strike/vertAlign/highlight/color → bold/italic/underline/strike/
 *                              superscript/subscript/highlight/textStyle marks
 *   - w:hyperlink + rels     → link mark (href via safeHref)
 *   - w:br / w:tab / w:t     → hardBreak / text
 *
 * Out of scope here (later commits): ③ lists, ④ tables, ⑤ math, ⑥ images.
 * Those are recognised and skipped so this walker always emits a schema-valid
 * doc; they are wired in as their commits land.
 */
import type { PmNode, PmMark } from './types.js'
import { safeHref, ooxmlHexColor, safeCssColor, decodeXmlText } from './types.js'
import { parseXmlOrdered, orderedTag, orderedAttr, orderedToXml, type OrderedNode } from './xml.js'
import type { Numbering } from './numbering.js'
import type { StyleMap } from './styles.js'
import { buildList, type ListLine } from './list.js'
import { mapTable } from './table.js'
import { ommlToLatex, mathNode } from './math.js'
import { imagePlaceholder } from './media.js'
import { DocxUnsafeError } from './extract.js'

/**
 * Wall-clock budget for the synchronous walk + OMML convert. The extractor's
 * timeout only bounds the zip inflate; this bounds the pure-memory recursion
 * that runs afterwards. `check()` throws a `timeout` DocxUnsafeError (mapped to
 * 400 by the route) once the budget is spent, so a wide / formula-dense
 * document can no longer pin the event loop unboundedly.
 */
export interface Deadline {
  check(): void
}

/** Build a deadline `budgetMs` from now. A non-positive budget never trips. */
export function makeDeadline(budgetMs: number): Deadline {
  const at = budgetMs > 0 ? Date.now() + budgetMs : Infinity
  return {
    check(): void {
      if (Date.now() > at) {
        throw new DocxUnsafeError('docx parse exceeded its time budget', 'timeout')
      }
    },
  }
}

/** Relationship id → target URL, from word/_rels/document.xml.rels. */
export type RelMap = Map<string, string>

export interface WalkResult {
  content: PmNode[]
  warnings: string[]
}

/** Children array of an ordered-mode node under a given tag. */
function kids(node: OrderedNode, tag: string): OrderedNode[] {
  const v = node[tag]
  return Array.isArray(v) ? (v as OrderedNode[]) : []
}

/**
 * Maximum list nesting level. OOXML bounds `w:ilvl` to 0..8; clamping to this
 * range is lossless for valid documents and stops a crafted deep level from
 * making buildList open an unbounded number of nested frames.
 */
const MAX_LIST_LEVEL = 8

/**
 * Maximum block nesting depth (tables inside tables inside …). The XML parser
 * already caps total nested tags, but that throws hard; this lower, table-level
 * cap degrades gracefully (drop + warn) well before a pathological document
 * could stress the synchronous recursion. Real documents nest at most a few
 * levels, so 20 is far above any legitimate use.
 */
const MAX_BLOCK_DEPTH = 20

/** Find the first child of `node` with the given tag (ordered mode). */
function firstChild(children: OrderedNode[], tag: string): OrderedNode | undefined {
  return children.find((c) => orderedTag(c) === tag)
}

/** Parse document.xml + rels into the top-level block content array. */
export function walkDocument(
  documentXml: Buffer,
  rels: RelMap,
  numbering?: Numbering,
  deadline?: Deadline,
  styles?: StyleMap,
): WalkResult {
  const warnings: string[] = []
  const top = parseXmlOrdered(documentXml)
  const documentNode = firstChild(top, 'w:document')
  const body = documentNode ? firstChild(kids(documentNode, 'w:document'), 'w:body') : undefined
  if (!body) {
    warnings.push('missing w:body — empty document')
    return { content: [emptyParagraph()], warnings }
  }

  const content = mapBlockChildren(kids(body, 'w:body'), rels, numbering, warnings, undefined, 0, deadline, styles)
  if (content.length === 0) content.push(emptyParagraph())
  return { content, warnings }
}

/**
 * Map an ordered array of block-level children (w:p / w:tbl) into PM blocks.
 * Consecutive list paragraphs (③) are grouped and rebuilt into nested lists;
 * tables (④) are converted recursively (cells contain their own block content).
 * Shared by the body and by table cells.
 */
function mapBlockChildren(
  children: OrderedNode[],
  rels: RelMap,
  numbering: Numbering | undefined,
  warnings: string[],
  parentFillsWidth?: boolean,
  depth = 0,
  deadline?: Deadline,
  styles?: StyleMap,
): PmNode[] {
  const content: PmNode[] = []
  let listRun: ListLine[] = []
  const flushList = (): void => {
    if (listRun.length === 0) return
    content.push(...buildList(listRun))
    listRun = []
  }

  // Consecutive styled paragraphs (CodeBlock / Callout*) are grouped, because a
  // single code block or callout is exported as one styled paragraph PER LINE.
  // We buffer a run of same-kind styled paragraphs and coalesce on style change.
  let styledRun:
    | { kind: 'code'; variant?: undefined; ps: OrderedNode[] }
    | { kind: 'callout'; variant: string; ps: OrderedNode[] }
    | { kind: 'blockquote'; variant?: undefined; ps: OrderedNode[] }
    | null = null
  // Depth of open BlockQuoteStart markers; while > 0, inner BlockQuote-styled
  // paragraphs are treated as plain paragraphs inside the marker frame.
  let bqDepth = 0
  const flushStyled = (): void => {
    if (!styledRun) return
    if (styledRun.kind === 'code') content.push(codeBlockFrom(styledRun.ps))
    else if (styledRun.kind === 'blockquote')
      content.push(blockquoteFrom(styledRun.ps, rels, warnings, deadline))
    else content.push(calloutFrom(styledRun.variant, styledRun.ps, rels, warnings, deadline))
    styledRun = null
  }

  for (const child of children) {
    // Wall-clock guard: the walk + OMML convert is synchronous and the
    // extractor's timeout only covers zip inflate, not this pure-memory
    // recursion. Check the shared deadline once per block so a wide /
    // formula-dense document trips a `timeout` DocxUnsafeError (mapped to 400)
    // instead of pinning the event loop unboundedly.
    deadline?.check()
    const tag = orderedTag(child)
    if (tag === 'w:p') {
      // A paragraph whose flow is a display-math wrapper (m:oMathPara) becomes
      // a block-level math node rather than a paragraph — UNLESS it is also a
      // numbered/bulleted list item (carries w:numPr). A numbered formula
      // ("1. x^2 = ...") must stay a list item (the math rendered inline inside
      // it) so its number survives; treating it as standalone blockMath drops
      // the list entirely (the "12345 都没了" bug).
      const blockMath = hasNumPr(child) ? null : blockMathOf(child, warnings)
      if (blockMath) {
        flushList()
        flushStyled()
        content.push(blockMath)
        continue
      }
      // Code block boundary marker: flush the current CodeBlock run so an
      // immediately-following code block starts a NEW codeBlock node instead of
      // merging. The marker itself carries no content and is dropped.
      if (codeBlockEndMarker(child, styles)) {
        flushStyled()
        continue
      }
      // Code block / callout paragraphs (recognised by pStyle). Group runs of
      // the same kind; a callout run also requires the same variant.
      const styled = styledKindOf(child, styles)
      // Blockquote boundary markers (BlockQuoteStart / BlockQuoteEnd) bracket a
      // quote's content; foldBlockquote() rebuilds the quote (and nesting) from
      // them. Handle these BEFORE the pStyle grouping: inside a frame, a
      // `BlockQuote`-styled paragraph is a plain quote paragraph, not a separate
      // grouped blockquote node (the frame is the wrapper).
      const bqMarker = blockquoteMarkerOf(child, styles)
      if (bqMarker === 'start') {
        flushList()
        flushStyled()
        bqDepth++
        content.push({ type: '__blockquoteStart' } as unknown as PmNode)
        continue
      }
      if (bqMarker === 'end') {
        flushList()
        flushStyled()
        if (bqDepth > 0) bqDepth--
        content.push({ type: '__blockquoteEnd' } as unknown as PmNode)
        continue
      }
      if (styled && styled.kind === 'blockquote' && bqDepth > 0) {
        // Inside a marker frame: emit as a plain paragraph so foldBlockquote
        // wraps it once (avoids double-wrapping via the legacy pStyle grouping).
        flushList()
        flushStyled()
        content.push(...mapParagraph(child, rels, warnings, deadline, styles))
        continue
      }
      if (styled) {
        flushList()
        if (
          styledRun &&
          styledRun.kind === styled.kind &&
          (styled.kind !== 'callout' || styledRun.variant === styled.variant)
        ) {
          styledRun.ps.push(child)
        } else {
          flushStyled()
          styledRun =
            styled.kind === 'code'
              ? { kind: 'code', ps: [child] }
              : styled.kind === 'blockquote'
                ? { kind: 'blockquote', ps: [child] }
                : { kind: 'callout', variant: styled.variant, ps: [child] }
        }
        continue
      }
      flushStyled()
      const listLine = listLineOf(child, rels, numbering, warnings, deadline)
      if (listLine) {
        listRun.push(listLine)
      } else {
        flushList()
        // Details boundary markers (DetailsStart / DetailsEnd) and the summary
        // line are emitted as sentinel nodes; foldDetails() rebuilds the nested
        // details tree from them after the linear walk.
        const marker = detailsMarkerOf(child, styles)
        if (marker === 'start') {
          content.push({ type: '__detailsStart' } as unknown as PmNode)
        } else if (marker === 'end') {
          content.push({ type: '__detailsEnd' } as unknown as PmNode)
        } else if (marker === 'summary') {
          const { inline } = collectInline(kids(child, 'w:p'), rels, warnings, deadline)
          content.push({
            type: 'detailsSummary',
            content: stripLeadingDetailsToggle(inline),
          })
        } else {
          content.push(...mapParagraph(child, rels, warnings, deadline, styles))
        }
      }
    } else if (tag === 'm:oMathPara' || tag === 'm:oMath') {
      // Display math can also sit directly at block level.
      flushList()
      flushStyled()
      const latex = ommlToLatex(orderedToXml(child))
      if (latex) content.push(mathNode(latex, true))
      else warnings.push('a block formula could not be converted')
    } else if (tag === 'w:tbl') {
      flushList()
      flushStyled()
      // Cap nested-table depth. The walk is synchronous and runs outside the
      // extractor's wall-clock timeout, so a deeply nested `<w:tbl>` (sub-MB,
      // passes every zip cap) could otherwise overflow the stack. Beyond the
      // cap the table is dropped with a warning rather than recursed into.
      if (depth >= MAX_BLOCK_DEPTH) {
        warnings.push('a deeply nested table was dropped (exceeded max nesting depth)')
        continue
      }
      const tblChildren = kids(child, 'w:tbl')
      content.push(
        mapTable(
          tblChildren,
          (tcChildren, fillsWidth) =>
            mapBlockChildren(tcChildren, rels, numbering, warnings, fillsWidth, depth + 1, deadline, styles),
          parentFillsWidth,
          deadline,
        ),
      )
    }
  }
  flushList()
  flushStyled()
  return foldDetails(foldBlockquote(content))
}

/** Paragraph-style ids our exporter uses for code blocks / callouts. */
const CALLOUT_STYLE_VARIANT: Record<string, string> = {
  calloutinfo: 'info',
  calloutwarn: 'warn',
  calloutwarning: 'warn',
  callouttip: 'tip',
  calloutsuccess: 'success',
}

/** Leading emoji glyphs our exporter prepends to a callout's first line. */
const CALLOUT_PREFIX_GLYPHS = ['ℹ️', '⚠️', '💡', '✅', 'ℹ', '⚠']

/**
 * Classify a paragraph as a details boundary marker or summary line by its
 * pStyle (DetailsStart / DetailsEnd / DetailsSummary), or null otherwise.
 */
function detailsMarkerOf(p: OrderedNode, styles?: StyleMap): 'start' | 'end' | 'summary' | null {
  const pChildren = kids(p, 'w:p')
  const pPr = firstChild(pChildren, 'w:pPr')
  if (!pPr) return null
  const styleId = orderedAttr(firstChild(kids(pPr, 'w:pPr'), 'w:pStyle') ?? {}, 'w:val')
  if (!styleId) return null
  const has = (token: string): boolean =>
    styles ? styles.matches(styleId, token) : styleId.toLowerCase().replace(/\s+/g, '') === token
  if (has('detailsstart')) return 'start'
  if (has('detailsend')) return 'end'
  if (has('detailssummary')) return 'summary'
  return null
}

/**
 * Classify a paragraph as a blockquote boundary marker by its pStyle
 * (BlockQuoteStart / BlockQuoteEnd), or null otherwise. These invisible markers
 * bracket a quote's content so any inner block (list, nested quote, …) stays
 * inside the quote on re-import.
 */
function blockquoteMarkerOf(p: OrderedNode, styles?: StyleMap): 'start' | 'end' | null {
  const pChildren = kids(p, 'w:p')
  const pPr = firstChild(pChildren, 'w:pPr')
  if (!pPr) return null
  const styleId = orderedAttr(firstChild(kids(pPr, 'w:pPr'), 'w:pStyle') ?? {}, 'w:val')
  if (!styleId) return null
  const has = (token: string): boolean =>
    styles ? styles.matches(styleId, token) : styleId.toLowerCase().replace(/\s+/g, '') === token
  if (has('blockquotestart')) return 'start'
  if (has('blockquoteend')) return 'end'
  return null
}

/**
 * True if the paragraph is the invisible `CodeBlockEnd` boundary marker the
 * exporter appends after each code block (so adjacent code blocks do not merge).
 */
function codeBlockEndMarker(p: OrderedNode, styles?: StyleMap): boolean {
  const pChildren = kids(p, 'w:p')
  const pPr = firstChild(pChildren, 'w:pPr')
  if (!pPr) return false
  const styleId = orderedAttr(firstChild(kids(pPr, 'w:pPr'), 'w:pStyle') ?? {}, 'w:val')
  if (!styleId) return false
  return styles
    ? styles.matches(styleId, 'codeblockend')
    : styleId.toLowerCase().replace(/\s+/g, '') === 'codeblockend'
}

/**
 * Fold a linear block list containing blockquote sentinels
 * (`__blockquoteStart` / …content… / `__blockquoteEnd`) into nested `blockquote`
 * nodes. Mirrors foldDetails: a stack handles arbitrary nesting, and malformed
 * markers degrade gracefully (unmatched End dropped, open frames flushed).
 */
function foldBlockquote(blocks: PmNode[]): PmNode[] {
  const hasMarker = blocks.some(
    (b) => b.type === '__blockquoteStart' || b.type === '__blockquoteEnd',
  )
  if (!hasMarker) return blocks

  const out: PmNode[] = []
  const stack: PmNode[][] = []
  const emit = (node: PmNode): void => {
    if (stack.length > 0) stack[stack.length - 1]!.push(node)
    else out.push(node)
  }
  const closeFrame = (content: PmNode[]): PmNode => ({
    type: 'blockquote',
    // Fold any details sentinels that were nested inside this quote.
    content:
      content.length > 0 ? foldDetails(content) : [{ type: 'paragraph', content: [] }],
  })

  for (const b of blocks) {
    if (b.type === '__blockquoteStart') {
      stack.push([])
    } else if (b.type === '__blockquoteEnd') {
      const frame = stack.pop()
      if (!frame) continue // unmatched End: drop
      emit(closeFrame(frame))
    } else {
      emit(b)
    }
  }
  // Flush any frames left open by malformed input (missing End markers).
  while (stack.length > 0) {
    const frame = stack.pop()!
    const node = closeFrame(frame)
    if (stack.length > 0) stack[stack.length - 1]!.push(node)
    else out.push(node)
  }
  return out
}

/**
 * Strip the leading "▸ " toggle glyph the exporter prepends to a details
 * summary line (and any immediately following whitespace) so it does not show
 * up as literal text — the editor renders its own disclosure triangle.
 */
function stripLeadingDetailsToggle(inline: PmNode[]): PmNode[] {
  if (inline.length === 0) return inline
  const first = inline[0]
  if (first?.type === 'text' && typeof first.text === 'string') {
    const stripped = first.text.replace(/^\s*▸\s*/, '')
    if (stripped !== first.text) {
      const rest = inline.slice(1)
      return stripped ? [{ ...first, text: stripped }, ...rest] : rest
    }
  }
  return inline
}

/**
 * Fold a linear block list containing details sentinels
 * (`__detailsStart` / `detailsSummary` / …content… / `__detailsEnd`) into
 * nested `details` nodes. A stack handles arbitrary nesting: each Start opens a
 * frame, the first following detailsSummary becomes the summary, remaining
 * blocks fill detailsContent, and End closes the frame (attaching it to its
 * parent frame or the output). Unbalanced or malformed markers degrade
 * gracefully: a stray summary becomes a paragraph, an unmatched End is dropped,
 * and any frames left open at the end are flushed as best-effort details.
 */
function foldDetails(blocks: PmNode[]): PmNode[] {
  const hasMarker = blocks.some(
    (b) => b.type === '__detailsStart' || b.type === '__detailsEnd' || b.type === 'detailsSummary',
  )
  if (!hasMarker) return blocks

  interface Frame {
    summary: PmNode | null
    content: PmNode[]
  }
  const out: PmNode[] = []
  const stack: Frame[] = []
  const emit = (node: PmNode): void => {
    if (stack.length > 0) stack[stack.length - 1]!.content.push(node)
    else out.push(node)
  }

  const summaryToParagraph = (summary: PmNode | null): PmNode => ({
    type: 'paragraph',
    content: (summary?.content as PmNode[]) ?? [],
  })

  const closeFrame = (frame: Frame): PmNode => {
    const summary: PmNode = frame.summary ?? { type: 'detailsSummary', content: [] }
    const content = frame.content.length > 0 ? frame.content : [{ type: 'paragraph', content: [] }]
    return {
      type: 'details',
      attrs: { open: false },
      // Fold any blockquote sentinels nested inside this details content.
      content: [summary, { type: 'detailsContent', content: foldBlockquote(content) }],
    }
  }

  for (const b of blocks) {
    if (b.type === '__detailsStart') {
      stack.push({ summary: null, content: [] })
    } else if (b.type === '__detailsEnd') {
      const frame = stack.pop()
      if (!frame) continue // unmatched End: drop
      emit(closeFrame(frame))
    } else if (b.type === 'detailsSummary') {
      const top = stack[stack.length - 1]
      if (top && top.summary === null) top.summary = b
      else emit(summaryToParagraph(b)) // stray summary: degrade to paragraph
    } else {
      emit(b)
    }
  }
  // Flush any frames left open by malformed input (missing End markers).
  while (stack.length > 0) {
    const frame = stack.pop()!
    const node = closeFrame(frame)
    if (stack.length > 0) stack[stack.length - 1]!.content.push(node)
    else out.push(node)
  }
  return out
}

/**
 * Classify a paragraph by its pStyle: a CodeBlock line, a Callout* line, or
 * neither. Consecutive same-kind paragraphs are later coalesced into one node.
 */
function styledKindOf(
  p: OrderedNode,
  styles?: StyleMap,
): { kind: 'code' } | { kind: 'callout'; variant: string } | { kind: 'blockquote' } | null {
  const pChildren = kids(p, 'w:p')
  const pPr = firstChild(pChildren, 'w:pPr')
  if (!pPr) return null
  const styleId = orderedAttr(firstChild(kids(pPr, 'w:pPr'), 'w:pStyle') ?? {}, 'w:val')
  if (!styleId) return null
  const has = (token: string): boolean =>
    styles ? styles.matches(styleId, token) : styleId.toLowerCase().replace(/\s+/g, '') === token
  if (has('codeblock')) return { kind: 'code' }
  if (has('blockquote')) return { kind: 'blockquote' }
  for (const [token, variant] of Object.entries(CALLOUT_STYLE_VARIANT)) {
    if (has(token)) return { kind: 'callout', variant }
  }
  return null
}

/**
 * Build a single `codeBlock` node from a run of CodeBlock paragraphs (one per
 * source line). codeBlock content is `text*` with NO marks, so we join the
 * plain text of each line with newlines. `language` is not recoverable from the
 * exported docx (the exporter drops it), so it stays null.
 */
function codeBlockFrom(ps: OrderedNode[]): PmNode {
  const lines = ps.map((p) => paragraphPlainText(p))
  const text = lines.join('\n')
  return {
    type: 'codeBlock',
    attrs: { language: null },
    content: text ? [{ type: 'text', text }] : [],
  }
}

/**
 * Build a `callout` node (content `block+`) from a run of Callout* paragraphs.
 * Each source paragraph becomes a child paragraph; the emoji prefix the
 * exporter prepended to the FIRST line is stripped so it does not double up
 * (the editor renders the variant icon itself).
 */
function calloutFrom(
  variant: string,
  ps: OrderedNode[],
  rels: RelMap,
  warnings: string[],
  deadline?: Deadline,
): PmNode {
  const inner: PmNode[] = ps.map((p, i) => {
    const { inline } = collectInline(kids(p, 'w:p'), rels, warnings, deadline)
    const content = i === 0 ? stripLeadingCalloutIcon(inline) : inline
    return { type: 'paragraph', content }
  })
  if (inner.length === 0) inner.push({ type: 'paragraph', content: [] })
  return { type: 'callout', attrs: { variant }, content: inner }
}

/**
 * Build a `blockquote` node (content `block+`) from a run of paragraphs styled
 * with the exporter's `BlockQuote` pStyle. Each source paragraph becomes a
 * child paragraph, preserving inline marks (unlike code blocks).
 */
function blockquoteFrom(
  ps: OrderedNode[],
  rels: RelMap,
  warnings: string[],
  deadline?: Deadline,
): PmNode {
  const inner: PmNode[] = ps.map((p) => {
    const { inline } = collectInline(kids(p, 'w:p'), rels, warnings, deadline)
    return { type: 'paragraph', content: inline }
  })
  if (inner.length === 0) inner.push({ type: 'paragraph', content: [] })
  return { type: 'blockquote', content: inner }
}

/** Concatenated plain text of a paragraph (w:t across all runs), ignoring marks. */
/**
 * A thematic break (horizontal rule). The exporter emits an empty paragraph
 * carrying only a bottom border. Detect that, and (for older exports) a
 * paragraph whose entire visible text is a run of horizontal-line glyphs
 * (─ box-drawing, — em dash, - hyphen, _ underscore), with no other content.
 */
function isHorizontalRule(pPrChildren: OrderedNode[], p: OrderedNode): boolean {
  // (a) empty paragraph with a bottom border (`w:pBdr/w:bottom`).
  const pBdr = firstChild(pPrChildren, 'w:pBdr')
  if (pBdr) {
    const bottom = firstChild(kids(pBdr, 'w:pBdr'), 'w:bottom')
    const val = bottom ? orderedAttr(bottom, 'w:val') : null
    const hasBorder = !!bottom && val !== 'none' && val !== 'nil'
    if (hasBorder && paragraphPlainText(p).trim() === '') return true
  }
  // (b) back-compat: text is only line glyphs (≥ 3 chars), nothing else.
  const text = paragraphPlainText(p).trim()
  if (text.length >= 3 && /^[\u2500\u2014\u2015\-_]+$/.test(text)) {
    // Guard: the paragraph must not carry images or other embedded blocks.
    return true
  }
  return false
}

function paragraphPlainText(p: OrderedNode): string {
  let s = ''
  for (const child of kids(p, 'w:p')) {
    if (orderedTag(child) !== 'w:r') continue
    for (const rc of kids(child, 'w:r')) {
      const tag = orderedTag(rc)
      if (tag === 'w:t') s += orderedText(rc)
      else if (tag === 'w:tab') s += '\t'
    }
  }
  return s
}

/**
 * Remove the leading callout icon glyph (+ any immediately following space)
 * that the exporter prepends to the first callout paragraph, so the imported
 * text does not carry a duplicate icon.
 */
function stripLeadingCalloutIcon(inline: PmNode[]): PmNode[] {
  if (inline.length === 0) return inline
  // The exporter's iconPrefix() emits up to three leading runs in some order:
  // the emoji glyph, an (often empty) bold spacer run, and then the body. Strip
  // the icon glyph wherever it leads, then trim ALL leading whitespace-only
  // text nodes from the result so no spacer run survives.
  const trimLeadingBlank = (nodes: PmNode[]): PmNode[] => {
    let out = nodes
    while (
      out.length > 0 &&
      out[0]?.type === 'text' &&
      typeof out[0]?.text === 'string' &&
      out[0].text.trim() === ''
    ) {
      out = out.slice(1)
    }
    return out
  }

  let nodes = trimLeadingBlank(inline)
  if (nodes.length === 0) return nodes
  const first = nodes[0]
  if (first && first.type === 'text' && typeof first.text === 'string') {
    let text: string = first.text
    for (const glyph of CALLOUT_PREFIX_GLYPHS) {
      if (text.startsWith(glyph)) {
        text = text.slice(glyph.length)
        break
      }
    }
    text = text.replace(/^(?:\u200d|\ufe0f)+/, '').replace(/^\s+/, '')
    const rest = nodes.slice(1)
    nodes = text ? [{ ...first, text }, ...rest] : rest
  }
  // Final pass: drop any spacer run that led the body after the glyph removal.
  return trimLeadingBlank(nodes)
}

/**
 * If this paragraph is a list item, return the extracted ListLine; else null.
 * A list paragraph is one that either carries w:numPr (numbered/bulleted) or a
 * w14:checkbox content control (our exported task items).
 */
function listLineOf(
  p: OrderedNode,
  rels: RelMap,
  numbering: Numbering | undefined,
  warnings: string[],
  deadline?: Deadline,
): ListLine | null {
  const pChildren = kids(p, 'w:p')
  const pPr = firstChild(pChildren, 'w:pPr')
  const pPrChildren = pPr ? kids(pPr, 'w:pPr') : []

  // Task item: a checkbox content control (w:sdt with w14:checkbox) anywhere in
  // the paragraph. Exported task lists use this instead of numbering.
  const checkbox = findCheckbox(pChildren)
  if (checkbox) {
    const { inline, blocks } = collectInline(pChildren, rels, warnings, deadline)
    if (blocks.length) warnings.push('an image inside a task item was dropped')
    // Our exporter writes one layout spacer immediately after the checkbox.
    // It is not document content, so remove exactly one leading whitespace
    // character. Do not trim all whitespace: a user-authored leading space in
    // the task text is meaningful and must remain.
    const taskInline = stripTaskCheckboxSpacer(inline)
    return { ilvl: indentLevel(pPrChildren), kind: 'task', checked: checkbox.checked, inline: taskInline }
  }

  // Numbered/bulleted item: w:pPr/w:numPr with a numId.
  const numPr = firstChild(pPrChildren, 'w:numPr')
  if (numPr) {
    const numPrChildren = kids(numPr, 'w:numPr')
    const numId = orderedAttr(firstChild(numPrChildren, 'w:numId') ?? {}, 'w:val')
    const ilvl = Number(orderedAttr(firstChild(numPrChildren, 'w:ilvl') ?? {}, 'w:val') ?? '0')
    // In OOXML `w:numId="0"` is the explicit "remove numbering" sentinel: Word
    // emits it so a paragraph that would inherit a numPr from its paragraph
    // style renders with NO list. Treat it as not-a-list (fall through to a
    // plain paragraph), not a bullet.
    if (numId != null && numId !== '0') {
      // Clamp the level to the OOXML-legal range (w:ilvl is 0..8) BEFORE it
      // reaches any consumer. A crafted `<w:ilvl w:val="1e999"/>` parses to
      // Infinity and a huge finite value burns CPU; numbering.kindOf walks a
      // descending `for (l = ilvl; l >= 0; l--)` loop, so an unclamped level is
      // an output-side resource-exhaustion DoS the input byte caps do not
      // bound. Clamping here is lossless for valid documents and protects
      // kindOf, startOf, and buildList alike. Mirrors the task path's
      // Math.min(depth, 8).
      const rawIlvl = Number.isFinite(ilvl) && ilvl >= 0 ? ilvl : 0
      const safeIlvl = Math.min(rawIlvl, MAX_LIST_LEVEL)
      const kind = numbering ? numbering.kindOf(numId, safeIlvl) : 'bullet'
      const { inline, blocks } = collectInline(pChildren, rels, warnings, deadline)
      if (blocks.length) warnings.push('an image inside a list item was dropped')
      // An ordered list's first number (w:start / w:startOverride) so a list
      // beginning at 20/41 is not silently reset to 1 on import.
      const start =
        kind === 'ordered' && numbering ? numbering.startOf(numId, safeIlvl) : undefined
      // Carry numId so buildList can break the run when two adjacent same-level,
      // same-kind lists have different numIds (the second restarts numbering).
      return { ilvl: safeIlvl, kind, inline, start, numId }
    }
  }

  return null
}

function stripTaskCheckboxSpacer(inline: PmNode[]): PmNode[] {
  const first = inline[0]
  if (first?.type !== 'text' || typeof first.text !== 'string' || !/^\s/.test(first.text)) return inline
  const text = first.text.slice(1)
  return text ? [{ ...first, text }, ...inline.slice(1)] : inline.slice(1)
}

/**
 * Locate a checkbox content control in a paragraph's children. Returns its
 * checked state, or null when none is present. We look for a w:sdt whose
 * w:sdtPr contains a w14:checkbox, and read w14:checked/@w14:val.
 */
function findCheckbox(pChildren: OrderedNode[]): { checked: boolean } | null {
  for (const child of pChildren) {
    if (orderedTag(child) !== 'w:sdt') continue
    const sdtChildren = kids(child, 'w:sdt')
    const sdtPr = firstChild(sdtChildren, 'w:sdtPr')
    if (!sdtPr) continue
    const sdtPrChildren = kids(sdtPr, 'w:sdtPr')
    const cb = firstChild(sdtPrChildren, 'w14:checkbox')
    if (!cb) continue
    const cbChildren = kids(cb, 'w14:checkbox')
    const checkedVal = orderedAttr(firstChild(cbChildren, 'w14:checked') ?? {}, 'w14:val')
    return { checked: checkedVal === '1' || checkedVal === 'true' }
  }
  return null
}

/** Derive a nesting level from the paragraph indent (task lists have no ilvl).
 * NOTE: the 720-twip step is how OUR OWN exporter indents task items
 * (720*(depth+1)); it is only precise for self-exported docx. Third-party docx
 * use arbitrary indents, so foreign task nesting may be mis-levelled (acceptable
 * — content is preserved, only the visual depth may flatten). */
function indentLevel(pPrChildren: OrderedNode[]): number {
  const ind = firstChild(pPrChildren, 'w:ind')
  if (!ind) return 0
  const left = Number(orderedAttr(ind, 'w:left') ?? orderedAttr(ind, 'w:start') ?? '0')
  if (!Number.isFinite(left) || left <= 0) return 0
  // Export indents task items by 720*(depth+1) twips; invert that, clamped.
  const depth = Math.round(left / 720) - 1
  return depth > 0 ? Math.min(depth, MAX_LIST_LEVEL) : 0
}

/** w:p → paragraph/heading + any extracted block-level images as siblings. */
function mapParagraph(
  p: OrderedNode,
  rels: RelMap,
  warnings: string[],
  deadline?: Deadline,
  styles?: StyleMap,
): PmNode[] {
  const pChildren = kids(p, 'w:p')
  const pPr = firstChild(pChildren, 'w:pPr')
  const pPrChildren = pPr ? kids(pPr, 'w:pPr') : []
  const styleId = pPr ? orderedAttr(firstChild(pPrChildren, 'w:pStyle') ?? {}, 'w:val') : null
  const align = readAlign(pPrChildren)

  // Thematic break (horizontal rule). The exporter writes an empty paragraph
  // whose only property is a bottom border (`w:pBdr/w:bottom`). Older exports
  // used a centered run of box-drawing dashes instead, so also recognise a
  // paragraph whose entire text is a run of ─ / — / - / _ glyphs. Either form
  // maps back to a single `horizontalRule` node.
  if (isHorizontalRule(pPrChildren, p)) {
    return [{ type: 'horizontalRule' }]
  }

  // Images are block-level atoms in the schema, so they cannot live in the
  // paragraph's inline content. collectInline separates them out; we emit them
  // as sibling blocks AFTER the paragraph (their document position within the
  // paragraph flow is not representable, so trailing-sibling is the closest).
  const { inline, blocks } = collectInline(
    pChildren,
    rels,
    warnings,
    deadline,
    styles?.defaultRunStyle,
  )

  const headingLevel = headingLevelFromStyle(styleId, styles)
  const attrs: Record<string, unknown> = {}
  if (align) attrs.textAlign = align

  const out: PmNode[] = []
  // Only emit the text block when it has inline content, OR when there are no
  // extracted image blocks (so an empty paragraph is still preserved).
  if (inline.length > 0 || blocks.length === 0) {
    if (headingLevel) {
      out.push({ type: 'heading', attrs: { ...attrs, level: headingLevel }, content: inline })
    } else {
      out.push({
        type: 'paragraph',
        ...(Object.keys(attrs).length ? { attrs } : {}),
        content: inline,
      })
    }
  }
  out.push(...blocks)
  return out
}

/** "Heading1".."Heading6" (or "heading 1" variants) → 1..6, else null. */
function headingLevelFromStyle(styleId: string | null, styles?: StyleMap): number | null {
  if (!styleId) return null
  // Resolve a numeric styleId (e.g. "2") to its style name via styles.xml first,
  // the same indirection the code/quote/callout/details classifiers use; Word
  // renumbers named heading styleIds into bare integers on copy/re-save.
  // Resolve a numeric styleId (e.g. "2") to its style name via styles.xml, but
  // also accept the raw styleId token (Heading1..6): resolving must not stop a
  // recognised token from classifying when its <w:name> is localized/third-party.
  for (let level = 1; level <= 6; level++) {
    const token = `heading${level}`
    const hit = styles
      ? styles.matches(styleId, token)
      : styleId.toLowerCase().replace(/\s+/g, '') === token
    if (hit) return level
  }
  return null
}

/**
 * If the paragraph is a display-math wrapper (contains an m:oMathPara, or a
 * lone m:oMath with no text runs), return a blockMath node; else null.
 */
function blockMathOf(p: OrderedNode, warnings: string[]): PmNode | null {
  const pChildren = kids(p, 'w:p')
  const mathPara = firstChild(pChildren, 'm:oMathPara')
  const loneMath = firstChild(pChildren, 'm:oMath')
  const hasRuns = pChildren.some((c) => orderedTag(c) === 'w:r' || orderedTag(c) === 'w:hyperlink')
  const target = mathPara ?? (loneMath && !hasRuns ? loneMath : undefined)
  if (!target) return null
  const latex = ommlToLatex(orderedToXml(target))
  if (latex) return mathNode(latex, true)
  warnings.push('a block formula could not be converted')
  return null
}

/** True when a paragraph carries a list numbering reference (w:pPr/w:numPr). */
function hasNumPr(p: OrderedNode): boolean {
  const pPr = firstChild(kids(p, 'w:p'), 'w:pPr')
  if (!pPr) return false
  return firstChild(kids(pPr, 'w:pPr'), 'w:numPr') != null
}

function readAlign(pPrChildren: OrderedNode[]): string | null {
  const jc = orderedAttr(firstChild(pPrChildren, 'w:jc') ?? {}, 'w:val')
  if (!jc) return null
  switch (jc) {
    case 'center':
      return 'center'
    case 'right':
    case 'end':
      return 'right'
    case 'both':
    case 'distribute':
      return 'justify'
    case 'left':
    case 'start':
      return 'left'
    default:
      return null
  }
}

/**
 * Collect the inline children of a paragraph in DOCUMENT ORDER (runs, inline
 * math, hyperlinks) AND separate out block-level images. Images are block atoms
 * in the schema, so they cannot sit in a paragraph's inline content; they are
 * returned in `blocks` for the caller to emit as sibling blocks.
 */
function collectInline(
  pChildren: OrderedNode[],
  rels: RelMap,
  warnings: string[],
  deadline?: Deadline,
  defaultRunStyle?: { fontFamily?: string; fontSize?: string },
): { inline: PmNode[]; blocks: PmNode[] } {
  const inline: PmNode[] = []
  const blocks: PmNode[] = []

  // The per-block deadline check in mapBlockChildren fires only between blocks.
  // A single paragraph / list-item / cell can still hold a run-storm (tens of
  // thousands of w:r / w:hyperlink children) that would otherwise expand in one
  // uninterrupted synchronous pass. Check the shared deadline every N runs so a
  // pathological single block cannot stall the event loop unbounded.
  let sinceCheck = 0
  const tick = (): void => {
    if (deadline && ++sinceCheck >= 512) {
      sinceCheck = 0
      deadline.check()
    }
  }

  for (const child of pChildren) {
    tick()
    const tag = orderedTag(child)
    if (tag === 'w:r') {
      mapRun(child, [], inline, blocks, defaultRunStyle)
    } else if (tag === 'm:oMath') {
      // Inline math embedded directly in the paragraph flow.
      const latex = ommlToLatex(orderedToXml(child))
      if (latex) inline.push(mathNode(latex, false))
      else warnings.push('an inline formula could not be converted')
    } else if (tag === 'm:oMathPara') {
      // A display-math wrapper (m:oMathPara) can appear inside a list item's
      // paragraph flow. blockMathOf only fires for non-list paragraphs, so a
      // numbered/bulleted item holding a display formula reaches collectInline
      // instead. Unwrap it and keep the formula as inline math so it survives
      // rather than being silently dropped.
      const latex = ommlToLatex(orderedToXml(child))
      if (latex) inline.push(mathNode(latex, false))
      else warnings.push('a formula could not be converted')
    } else if (tag === 'w:hyperlink') {
      const relId = orderedAttr(child, 'r:id')
      const anchor = orderedAttr(child, 'w:anchor')
      let href: string | null = null
      if (relId && rels.has(relId)) href = safeHref(rels.get(relId))
      else if (anchor) href = safeHref(`#${anchor}`)
      const linkMark: PmMark[] = href ? [{ type: 'link', attrs: { href } }] : []
      for (const run of kids(child, 'w:hyperlink')) {
        tick()
        if (orderedTag(run) === 'w:r') mapRun(run, linkMark, inline, blocks, defaultRunStyle)
      }
    }
  }

  return { inline, blocks }
}

/**
 * w:r → inline text/hardBreak nodes (into `inline`) carrying the run's marks
 * (+ inherited). Embedded images (w:drawing) are block atoms and go into
 * `blocks`, not the inline flow. Children are processed in document order.
 */
function mapRun(
  run: OrderedNode,
  inherited: PmMark[],
  inline: PmNode[],
  blocks: PmNode[],
  defaultRunStyle?: { fontFamily?: string; fontSize?: string },
): void {
  const runChildren = kids(run, 'w:r')
  const rPr = firstChild(runChildren, 'w:rPr')
  const marks = [...inherited, ...marksFromRunProps(rPr, defaultRunStyle)]

  for (const child of runChildren) {
    const tag = orderedTag(child)
    if (tag === 'w:t') {
      const text = orderedText(child)
      if (text) inline.push({ type: 'text', text, ...(marks.length ? { marks } : {}) })
    } else if (tag === 'w:tab') {
      inline.push({ type: 'text', text: '\t', ...(marks.length ? { marks } : {}) })
    } else if (tag === 'w:br') {
      inline.push({ type: 'hardBreak' })
    } else if (tag === 'w:drawing') {
      const img = drawingToImage(child)
      if (img) blocks.push(img)
    }
  }
}

/**
 * Extract an embedded-image placeholder from a w:drawing. We locate the
 * a:blip/@r:embed relationship id (works for both wp:inline and wp:anchor) and
 * any alt text (wp:docPr/@descr or @title). The bytes are resolved + uploaded
 * later by the async media step (⑥); here we only emit the placeholder.
 */
function drawingToImage(drawing: OrderedNode): PmNode | null {
  // Office stores a PNG fallback on a:blip and the original vector relation on
  // asvg:svgBlip. Prefer the native SVG relation. We still emit exactly one
  // placeholder, and older documents without svgBlip retain raster fallback.
  const svgRel = findSvgEmbedRel(drawing)
  const rasterRel = findEmbedRel(drawing)
  const embedRel = svgRel ?? rasterRel
  if (!embedRel) return null
  const alt = findDocPrDescr(drawing)
  return imagePlaceholder(embedRel, alt, svgRel ? rasterRel : null)
}

function findSvgEmbedRel(node: OrderedNode, depth = 0): string | null {
  if (depth > MAX_DRAWING_DEPTH) return null
  const tag = orderedTag(node)
  if (tag === 'asvg:svgBlip' || tag?.endsWith(':svgBlip')) {
    return orderedAttr(node, 'r:embed') ?? null
  }
  if (tag) {
    for (const child of kids(node, tag)) {
      const found = findSvgEmbedRel(child, depth + 1)
      if (found) return found
    }
  }
  return null
}

/** Max recursion depth for the in-memory drawing DFS. The extract-layer CPU
 * timeout does NOT cover this pure-memory recursion, so a maliciously deep
 * drawing subtree could blow the stack; this cap aligns with the archive's
 * nesting defence. Real drawings nest ~6 levels; 64 is generous. */
const MAX_DRAWING_DEPTH = 64

/** Depth-first search for the first a:blip/@r:embed within a drawing subtree. */
function findEmbedRel(node: OrderedNode, depth = 0): string | null {
  if (depth > MAX_DRAWING_DEPTH) return null
  const tag = orderedTag(node)
  if (tag === 'a:blip') {
    const embed = orderedAttr(node, 'r:embed') ?? orderedAttr(node, 'r:link')
    if (embed) return embed
  }
  if (tag) {
    for (const child of kids(node, tag)) {
      const found = findEmbedRel(child, depth + 1)
      if (found) return found
    }
  }
  return null
}

/** Depth-first search for a wp:docPr/@descr (or @title) alt string. */
function findDocPrDescr(node: OrderedNode, depth = 0): string | null {
  if (depth > MAX_DRAWING_DEPTH) return null
  const tag = orderedTag(node)
  if (tag === 'wp:docPr') {
    return orderedAttr(node, 'descr') ?? orderedAttr(node, 'title')
  }
  if (tag) {
    for (const child of kids(node, tag)) {
      const found = findDocPrDescr(child, depth + 1)
      if (found) return found
    }
  }
  return null
}

/** w:rPr → the PM marks it implies. */
function marksFromRunProps(
  rPr: OrderedNode | undefined,
  defaults?: { fontFamily?: string; fontSize?: string },
): PmMark[] {
  const c = rPr ? kids(rPr, 'w:rPr') : []
  const marks: PmMark[] = []

  const b = firstChild(c, 'w:b')
  if (b && toggleOnOrdered(b)) marks.push({ type: 'bold' })
  const i = firstChild(c, 'w:i')
  if (i && toggleOnOrdered(i)) marks.push({ type: 'italic' })
  const u = firstChild(c, 'w:u')
  if (u && orderedAttr(u, 'w:val') !== 'none') marks.push({ type: 'underline' })
  const strike = firstChild(c, 'w:strike')
  if (strike && toggleOnOrdered(strike)) marks.push({ type: 'strike' })

  const vertAlign = orderedAttr(firstChild(c, 'w:vertAlign') ?? {}, 'w:val')
  if (vertAlign === 'superscript') marks.push({ type: 'superscript' })
  else if (vertAlign === 'subscript') marks.push({ type: 'subscript' })

  const highlight = orderedAttr(firstChild(c, 'w:highlight') ?? {}, 'w:val')
  if (highlight && highlight !== 'none') {
    const col = safeCssColor(highlight)
    marks.push(col ? { type: 'highlight', attrs: { color: col } } : { type: 'highlight' })
  } else {
    // Exporter emits an arbitrary highlight colour as `w:shd` (shading fill), not
    // `w:highlight` (which is limited to ~16 named colours). Read the shading
    // fill back into the highlight mark so a self-exported doc round-trips its
    // exact background colour. `w:shd/@w:fill` of `auto` / absent means no fill.
    // Guard: inline `code` runs are exported as `w:shd fill=F0F0F0` + the code
    // font; that shading is the code chip background, NOT a user highlight, so
    // skip it when the run carries the code font.
    const shdFill = ooxmlHexColor(orderedAttr(firstChild(c, 'w:shd') ?? {}, 'w:fill'))
    const runFont = orderedAttr(firstChild(c, 'w:rFonts') ?? {}, 'w:ascii')
    const isCodeFont =
      runFont != null && /^(consolas|courier new)$/i.test(runFont.trim())
    if (shdFill && !(isCodeFont && shdFill === '#f0f0f0')) {
      marks.push({ type: 'highlight', attrs: { color: shdFill } })
    }
  }

  const color = ooxmlHexColor(orderedAttr(firstChild(c, 'w:color') ?? {}, 'w:val'))
  // w:sz is in half-points; convert to CSS px (1pt = 4/3px), which makes both
  // PDF (px -> pt) and DOCX (px -> half-points) preserve the source size.
  const szHalfPts = Number(orderedAttr(firstChild(c, 'w:sz') ?? {}, 'w:val'))
  const fontSize =
    Number.isFinite(szHalfPts) && szHalfPts > 0 ? `${szHalfPts * 2 / 3}px` : defaults?.fontSize
  const rFonts = firstChild(c, 'w:rFonts') ?? {}
  const fontFamily = orderedAttr(rFonts, 'w:ascii')
    ?? orderedAttr(rFonts, 'w:hAnsi')
    ?? orderedAttr(rFonts, 'w:eastAsia')
    ?? defaults?.fontFamily
  if (color || fontSize || fontFamily) {
    const attrs: Record<string, unknown> = {}
    if (color) attrs.color = color
    if (fontSize) attrs.fontSize = fontSize
    if (fontFamily) attrs.fontFamily = fontFamily
    marks.push({ type: 'textStyle', attrs })
  }

  return marks
}

/** OOXML boolean toggle (ordered node): absent w:val => on. */
function toggleOnOrdered(node: OrderedNode): boolean {
  const v = orderedAttr(node, 'w:val')
  if (v == null) return true
  return !['0', 'false', 'off'].includes(v.toLowerCase())
}

/** Read text out of an ordered w:t node: its `#text` child (entity-decoded). */
function orderedText(t: OrderedNode): string {
  const textKids = kids(t, 'w:t')
  let s = ''
  for (const k of textKids) {
    const v = (k as Record<string, unknown>)['#text']
    if (v != null) s += String(v)
  }
  return decodeXmlText(s)
}

function emptyParagraph(): PmNode {
  return { type: 'paragraph', content: [] }
}
