/**
 * Tests for math reconstruction (commit ⑤): OMML → LaTeX, and the round-trip
 * validation: our own exported formulas must import back to (semantically) the
 * same LaTeX.
 */
import { describe, it, expect } from 'vitest'
import { ommlToLatex, mathNode } from '../src/import/docx/math.js'
import { walkDocument } from '../src/import/docx/document.js'

const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'

/** OMML fixtures shaped like what the exporter (MathJax→mathmlToOmml) emits. */
const OMML = {
  // x^2
  superscript: `<m:oMath ${M}><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>`,
  // a/b
  fraction: `<m:oMath ${M}><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>`,
  // a_i (subscript)
  subscript: `<m:oMath ${M}><m:sSub><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub></m:oMath>`,
}

describe('ommlToLatex — conversion', () => {
  it('converts a superscript', () => {
    expect(ommlToLatex(OMML.superscript)).toBe('x^{2}')
  })

  it('converts a fraction', () => {
    expect(ommlToLatex(OMML.fraction)).toBe('\\frac{a}{b}')
  })

  it('unwraps a malformed fraction with an empty denominator', () => {
    const malformed = `<m:oMath ${M}><m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den/></m:f></m:oMath>`
    expect(ommlToLatex(malformed)).toBe('x')
  })

  it('converts a subscript', () => {
    expect(ommlToLatex(OMML.subscript)).toBe('a_{i}')
  })

  it('injects the m: namespace when missing', () => {
    const noNs = OMML.fraction.replace(` ${M}`, '')
    expect(ommlToLatex(noNs)).toBe('\\frac{a}{b}')
  })

  it('returns null for unparseable / empty OMML', () => {
    expect(ommlToLatex('not omml at all')).toBeNull()
    expect(ommlToLatex(`<m:oMath ${M}></m:oMath>`)).toBeNull()
  })

  it('degrades an oversized single formula to null without converting (per-formula DoS bound)', () => {
    // A crafted formula packed into ONE m:oMath (deeply nested / run-storm) is
    // interruptible by the between-formula deadline; the per-formula size bound
    // caps the work done inside one conversion. An OMML string past the byte
    // cap must degrade to null (drop + warn) rather than expanding synchronously.
    const runs = `<m:r><m:t>x</m:t></m:r>`.repeat(20000) // ~460 KB, over the 256 KB cap
    const huge = `<m:oMath ${M}>${runs}</m:oMath>`
    const start = Date.now()
    expect(ommlToLatex(huge)).toBeNull()
    // Rejected up front, so it returns near-instantly rather than doing the
    // full parse+convert on the giant input.
    expect(Date.now() - start).toBeLessThan(50)
  })

  it('decodes XML entities in run text so >/< operators are not garbled', () => {
    // OMML is parsed with processEntities:false, so `<m:t>&gt;</m:t>` arrives as
    // the literal 4 chars `&gt;`. The run text must be entity-decoded before the
    // char tokeniser runs, or `>` renders as the broken `\& g t ;`.
    const gt = `<m:oMath ${M}><m:r><m:t>x</m:t></m:r><m:r><m:t>&gt;</m:t></m:r><m:r><m:t>0</m:t></m:r></m:oMath>`
    expect(ommlToLatex(gt)).toBe('x > 0')
    const lt = `<m:oMath ${M}><m:r><m:t>a</m:t></m:r><m:r><m:t>&lt;</m:t></m:r><m:r><m:t>b</m:t></m:r></m:oMath>`
    expect(ommlToLatex(lt)).toBe('a < b')
  })

  it('strips a DOCTYPE / entity bomb before parsing and still converts the math', () => {
    // Security: the xmldom path must not expand internal DTD entities. A crafted
    // OMML carrying a DOCTYPE + nested-entity bomb must be neutralized; the
    // math still converts, and no entity expansion occurs.
    const bomb =
      '<?xml version="1.0"?>' +
      '<!DOCTYPE m:oMath [<!ENTITY a "AAAA"><!ENTITY b "&a;&a;&a;&a;"><!ENTITY c "&b;&b;&b;&b;">]>' +
      `<m:oMath ${M}><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>`
    const latex = ommlToLatex(bomb)
    expect(latex).toBe('x^{2}')
    // The custom entity reference must not have expanded into the output.
    expect(latex).not.toContain('AAAA')
  })
})

/**
 * Round-trip: normalise LaTeX (strip whitespace + redundant single-char braces)
 * so semantically equal expressions compare equal. The export side emits e.g.
 * `x^{2}`; a hand-simplified `x^2` is the same formula.
 */
function normLatex(s: string): string {
  return s.replace(/\s+/g, '').replace(/\{([a-zA-Z0-9])\}/g, '$1')
}

describe('round-trip — exported LaTeX survives OMML import', () => {
  // (exported-LaTeX, OMML the exporter produces for it). We assert the imported
  // LaTeX is SEMANTICALLY the original. OMML fixtures mirror mathmlToOmml output.
  const cases: Array<[string, string]> = [
    ['x^2', OMML.superscript],
    ['\\frac{a}{b}', OMML.fraction],
    ['a_i', OMML.subscript],
  ]

  for (const [original, omml] of cases) {
    it(`round-trips ${original}`, () => {
      const imported = ommlToLatex(omml)
      expect(imported).not.toBeNull()
      expect(normLatex(imported!)).toBe(normLatex(original))
    })
  }
})

describe('mathNode', () => {
  it('builds block vs inline math nodes', () => {
    expect(mathNode('x^2', true)).toEqual({ type: 'blockMath', attrs: { latex: 'x^2' } })
    expect(mathNode('x^2', false)).toEqual({ type: 'inlineMath', attrs: { latex: 'x^2' } })
  })
})

describe('walkDocument — math integration', () => {
  function doc(inner: string): Buffer {
    return Buffer.from(
      `<?xml version="1.0"?><w:document xmlns:w="http://x" ${M}><w:body>${inner}</w:body></w:document>`,
      'utf8',
    )
  }

  it('maps inline m:oMath inside a paragraph to an inlineMath node', () => {
    const inner = `<w:p><w:r><w:t>see </w:t></w:r><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></w:p>`
    const out = walkDocument(doc(inner), new Map())
    const para = out.content[0]!
    expect(para.type).toBe('paragraph')
    expect(para.content![0]).toEqual({ type: 'text', text: 'see ' })
    expect(para.content![1]).toEqual({ type: 'inlineMath', attrs: { latex: '\\frac{a}{b}' } })
  })

  it('maps a lone m:oMath paragraph to a blockMath node', () => {
    const inner = `<w:p><m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></w:p>`
    const out = walkDocument(doc(inner), new Map())
    expect(out.content[0]).toEqual({ type: 'blockMath', attrs: { latex: 'x^{2}' } })
  })

  it('maps m:oMathPara at block level to blockMath', () => {
    const inner = `<m:oMathPara><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>`
    const out = walkDocument(doc(inner), new Map())
    expect(out.content[0]).toEqual({ type: 'blockMath', attrs: { latex: '\\frac{a}{b}' } })
  })
})

describe('ommlToLatex — recovered-LaTeX normalization', () => {
  const M2 = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'

  it('round-trips a matrix (m:m) to \\begin{matrix}', () => {
    const mat = `<m:oMath ${M2}><m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr><m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr></m:m></m:oMath>`
    const latex = ommlToLatex(mat)
    expect(latex).not.toBeNull()
    expect(latex).toContain('\\begin{matrix}')
    expect(latex).toContain('a')
    expect(latex).toContain('d')
  })

  it('normalizes accent operators to dedicated commands', () => {
    const acc = (chr: string) =>
      `<m:oMath ${M2}><m:acc><m:accPr><m:chr m:val="${chr}"/></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc></m:oMath>`
    expect(ommlToLatex(acc('\u0307'))).toBe('\\dot{x}') // combining dot above
    expect(ommlToLatex(acc('\u20d7'))).toBe('\\vec{x}') // combining right arrow above
    expect(ommlToLatex(acc('~'))).toBe('\\tilde{x}')
    expect(ommlToLatex(acc('\u0304'))).toBe('\\bar{x}') // combining macron -> \bar
    expect(ommlToLatex(acc('\u0308'))).toBe('\\ddot{x}') // combining diaeresis -> \ddot
  })

  it('converts m:bar to valid KaTeX for both orientations', () => {
    // Word "bar" accent: pos="top" is an overline (\bar), pos="bot" is an
    // underbar. The bot path used to yield the invalid `\underset{\underline}{X}`
    // (KaTeX parse error -> whole formula renders red); it must normalize to
    // the valid `\underline{X}`. The top path stays `\bar{X}`.
    const bar = (pos: string, ...bases: string[]) =>
      `<m:oMath ${M2}><m:bar><m:barPr><m:pos m:val="${pos}"/></m:barPr><m:e>` +
      bases.map((b) => `<m:r><m:t>${b}</m:t></m:r>`).join('') +
      `</m:e></m:bar></m:oMath>`
    expect(ommlToLatex(bar('top', 'y'))).toBe('\\bar{y}')
    expect(ommlToLatex(bar('bot', 'y'))).toBe('\\underline{y}')
    // Multi-token base under the bar must stay a single \underline group.
    expect(ommlToLatex(bar('bot', 'A', 'B'))).toBe('\\underline{A B}')
    // No invalid \underset{\underline} leaks through.
    expect(ommlToLatex(bar('bot', 'y'))).not.toContain('\\underset')
  })

  it('recovers a scripted function name from \\left(…\\right)^{…}', () => {
    const sinsup = `<m:oMath ${M2}><m:sSup><m:e><m:r><m:t>sin</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>`
    const latex = ommlToLatex(sinsup)
    expect(latex).not.toBeNull()
    expect(latex).not.toContain('\\left(')
    expect(latex).toContain('\\sin')
  })

  it('recovers a uniform run color as \\textcolor', () => {
    const colored =
      `<m:oMath ${M2} xmlns:w="http://w">` +
      `<m:r><m:rPr><w:color w:val="FF0000"/></m:rPr><m:t>E</m:t></m:r>` +
      `<m:r><m:rPr><w:color w:val="FF0000"/></m:rPr><m:t>=</m:t></m:r>` +
      `<m:r><m:rPr><w:color w:val="FF0000"/></m:rPr><m:t>m</m:t></m:r>` +
      `</m:oMath>`
    const latex = ommlToLatex(colored)
    expect(latex).not.toBeNull()
    expect(latex).toMatch(/^\\textcolor\{#FF0000\}\{/)
  })

  it('does not wrap in \\textcolor when coloring is mixed / partial', () => {
    const mixed =
      `<m:oMath ${M2} xmlns:w="http://w">` +
      `<m:r><m:rPr><w:color w:val="FF0000"/></m:rPr><m:t>a</m:t></m:r>` +
      `<m:r><m:t>b</m:t></m:r>` + // uncolored run
      `</m:oMath>`
    const latex = ommlToLatex(mixed)
    expect(latex).not.toBeNull()
    expect(latex).not.toContain('\\textcolor')
  })

  it('rewrites stacked n-ary limits (matrix in a script) to \\substack', () => {
    // \sum_{i=1 \\ j=1}^{n} a_{ij}: OMML nary with a 2-row matrix subscript. The
    // converter emits `_{\begin{matrix}…\end{matrix}}`, which renders the limits
    // full-size/misaligned; the correct construct is `_{\substack{…}}`.
    const nary =
      `<m:oMath ${M2}><m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr>` +
      `<m:sub><m:m><m:mr><m:e><m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>1</m:t></m:r></m:e></m:mr>` +
      `<m:mr><m:e><m:r><m:t>j</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>1</m:t></m:r></m:e></m:mr></m:m></m:sub>` +
      `<m:sup><m:r><m:t>n</m:t></m:r></m:sup>` +
      `<m:e><m:sSub><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r><m:r><m:t>j</m:t></m:r></m:sub></m:sSub></m:e></m:nary></m:oMath>`
    const latex = ommlToLatex(nary)
    expect(latex).not.toBeNull()
    expect(latex).toContain('\\substack{')
    expect(latex).not.toContain('\\begin{matrix}')
  })

  it('restores \\pmod from a parenthesised \\bmod delimiter', () => {
    // \pmod{n} round-trips through OMML as `\left(\bmod n\right)`, where the
    // \bmod spacing collapses to `(modn)`. Recover the proper `\pmod{n}`.
    const pmod =
      `<m:oMath ${M2}><m:r><m:t>a</m:t></m:r><m:r><m:t>≡</m:t></m:r><m:r><m:t>b</m:t></m:r>` +
      `<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr>` +
      `<m:e><m:r><m:t>mod</m:t></m:r><m:r><m:t>n</m:t></m:r></m:e></m:d></m:oMath>`
    const latex = ommlToLatex(pmod)
    expect(latex).not.toBeNull()
    expect(latex).toContain('\\pmod{n}')
    expect(latex).not.toContain('\\left(')
  })

  it('keeps a parenthesised matrix as \\begin{pmatrix} (round parens, not [])', () => {
    // OMML delimiter `(` `)` around a matrix. omml2mathml drops the default
    // paren fence, so mathml-to-latex would emit \begin{bmatrix}; restoring the
    // fence keeps it \begin{pmatrix}.
    const pmat =
      `<m:oMath ${M2}><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr>` +
      `<m:e><m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr>` +
      `<m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr></m:m></m:e></m:d></m:oMath>`
    const latex = ommlToLatex(pmat)
    expect(latex).not.toBeNull()
    expect(latex).toContain('\\begin{pmatrix}')
    expect(latex).not.toContain('bmatrix')
  })

  it('keeps a bracketed matrix as \\begin{bmatrix} ([] unaffected by the fence fix)', () => {
    const bmat =
      `<m:oMath ${M2}><m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>` +
      `<m:e><m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e></m:mr></m:m></m:e></m:d></m:oMath>`
    const latex = ommlToLatex(bmat)
    expect(latex).not.toBeNull()
    expect(latex).toContain('\\begin{bmatrix}')
    expect(latex).not.toContain('pmatrix')
  })

  it('recovers a wide arrow/hat over a multi-char base as \\overrightarrow / \\widehat', () => {
    // \overrightarrow{AB} exports as <m:groupChr chr="→">; omml2mathml gives
    // <mover><mrow>AB</mrow><mo>→</mo></mover> -> \overset{\rightarrow}{A B}.
    // A multi-token base must recover the WIDE command, not narrow \vec.
    const arrow =
      `<m:oMath ${M2}><m:groupChr><m:groupChrPr><m:chr m:val="→"/><m:pos m:val="top"/><m:vertJc m:val="bot"/></m:groupChrPr>` +
      `<m:e><m:r><m:t>A</m:t></m:r><m:r><m:t>B</m:t></m:r></m:e></m:groupChr></m:oMath>`
    const latex = ommlToLatex(arrow)
    expect(latex).not.toBeNull()
    expect(latex).toContain('\\overrightarrow{')
    expect(latex).not.toContain('\\vec')
  })

  it('keeps a single-char accent narrow (\\vec / \\hat, not the wide form)', () => {
    const acc = (chr: string, base: string) =>
      `<m:oMath ${M2}><m:acc><m:accPr><m:chr m:val="${chr}"/></m:accPr><m:e><m:r><m:t>${base}</m:t></m:r></m:e></m:acc></m:oMath>`
    expect(ommlToLatex(acc('\u20d7', 'v'))).toBe('\\vec{v}')
    expect(ommlToLatex(acc('\u0302', 'x'))).toBe('\\hat{x}')
    expect(ommlToLatex(acc('\u0304', 'y'))).toBe('\\bar{y}')
  })
})
