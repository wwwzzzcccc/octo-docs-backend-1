import { describe, it, expect } from 'vitest'
import { renderTypst, __test } from '../src/export/renderTypst.js'
import type { ResolvedAttachment } from '../src/export/renderTypst.js'

// ── helpers ────────────────────────────────────────────────────────────────
const doc = (content: unknown[]) => ({ type: 'doc', content })
const para = (content: unknown[]) => ({ type: 'paragraph', content })
const text = (t: string, marks?: unknown[]) => ({ type: 'text', text: t, ...(marks ? { marks } : {}) })

function typ(content: unknown[], attachments = new Map<string, ResolvedAttachment>()): string {
  return renderTypst(doc(content), { title: 'T', attachments })
}

const m = __test.latexToTypstMath

// ── LaTeX -> Typst math conversion ───────────────────────────────────────────
describe('renderTypst — latexToTypstMath', () => {
  it('converts fractions without spurious visible parens', () => {
    expect(m('\\frac{a}{b}')).toBe('frac(a, b)')
    expect(m('\\frac{1}{n^2}')).toBe('frac(1, n^2)')
  })

  it('converts scripts, grouping multi-atom exponents only', () => {
    expect(m('x^2')).toBe('x^2')
    expect(m('x^{2n}')).toBe('x^(2n)')
    expect(m('x_{i}')).toBe('x_i')
  })

  it('maps greek letters and common operators', () => {
    expect(m('\\alpha')).toBe('alpha')
    expect(m('\\pi')).toBe('pi')
    expect(m('a \\leq b')).toBe('a lt.eq b')
    expect(m('a \\times b')).toBe('a times b')
    expect(m('\\infty')).toBe('infinity')
  })

  it('converts sqrt, nth-root, and binom', () => {
    expect(m('\\sqrt{x+1}')).toBe('sqrt(x+1)')
    expect(m('\\sqrt[3]{x}')).toBe('root(3, x)')
    expect(m('\\binom{n}{k}')).toBe('binom(n, k)')
  })

  it('converts sum with limits', () => {
    expect(m('\\sum_{n=1}^{\\infty}')).toBe('sum_(n=1)^(infinity)')
  })

  it('wraps \\text{...} as a Typst string literal', () => {
    expect(m('\\text{hello}')).toBe('"hello"')
  })

  it('strips ASCII control chars so a corrupted formula cannot garble the PDF', () => {
    // Regression ("导出pdf乱码"): a formula authored in a non-raw JS string
    // collapsed `\to`→TAB (U+0009) and `\frac`→FORM-FEED (U+000C). The stray
    // control bytes desynced the command scanner and emitted a bare `cs`
    // identifier that Typst rejects as an unknown variable, failing the whole
    // compile. Control chars must be normalized/stripped before parsing.
    const corrupt = 'lim_{x \u0009o 0} \u000crac{x}{y}'
    const out = m(corrupt)
    // No form-feed / TAB survive into the emitted Typst math.
    // eslint-disable-next-line no-control-regex -- assert control chars are stripped
    expect(out).not.toMatch(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/)
    // A clean formula is unaffected.
    expect(m('\\lim_{x \\to 0} \\frac{\\sin x}{x} = 1')).toBe(
      'lim_(x arrow.r 0) frac(sin x, x) = 1',
    )
  })

  it('handles \\color / \\textcolor in math without garbling (regression: 彩色公式导出乱码)', () => {
    // Unhandled \color/\textcolor used to spell out as `"color"r e d…`, emitting
    // stray identifiers (dx, ey) that Typst rejects as unknown variables and
    // failing the whole PDF compile. They must become Typst `#text(fill: ...)`.
    expect(m('\\color{red}{x^2}')).toBe('#text(fill: red)[$x^2$]')
    expect(m('\\textcolor{blue}{y}')).toBe('#text(fill: blue)[$y$]')
    expect(m('\\textcolor{#ff0000}{z}')).toBe('#text(fill: rgb("#ff0000"))[$z$]')
    // xcolor names not in Typst's palette map to a near equivalent.
    expect(m('\\textcolor{cyan}{w}')).toBe('#text(fill: aqua)[$w$]')
    // The colour name must NOT be math-converted to spaced letters.
    expect(m('\\color{red}{a}')).not.toContain('r e d')
    // `\color{c}{x}` colours only the following group — trailing content stays
    // uncoloured (regression: rest-of-group form used to swallow `+ z` / `= c`).
    expect(m('x + \\color{red}{y} + z')).toBe('x + #text(fill: red)[$y$] + z')
    expect(m('\\color{red}{a} + \\color{blue}{b}')).toBe(
      '#text(fill: red)[$a$] + #text(fill: blue)[$b$]',
    )
    // Rest-of-scope form must run trailing content through the full converter so
    // adjacent letters get spaced (mc -> m c); otherwise Typst reads `mc` as an
    // unknown variable and the whole PDF export falls back to raw-LaTeX verbatim
    // (regression: colored formula garbled the entire export).
    expect(m('\\color{red} E = mc^2')).toBe('#text(fill: red)[$E = m c^2$]')
  })

  it('converts a pmatrix to a bracketed Typst mat() (regression: 公式没组合)', () => {
    // Previously \begin/\end and the & / \\\\ separators were discarded, mushing
    // every cell into one flat run with no delimiters. The matrix is wrapped in
    // lr(size: #88%, ...) so the brackets hug the two rows instead of towering.
    expect(m('\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}')).toBe(
      'lr(size: #88%, ( mat(delim: #none, a, b; c, d) ))',
    )
  })

  it('converts bmatrix / vmatrix / matrix with correct delimiters', () => {
    expect(m('\\begin{bmatrix} 1 & 0 \\\\ 0 & 1 \\end{bmatrix}')).toBe(
      'lr(size: #88%, [ mat(delim: #none, 1, 0; 0, 1) ])',
    )
    expect(m('\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}')).toBe(
      'lr(size: #88%, | mat(delim: #none, a, b; c, d) |)',
    )
    expect(m('\\begin{matrix} a & b \\\\ c & d \\end{matrix}')).toBe('mat(delim: #none, a, b; c, d)')
  })

  it('strips the array column-spec so it does not leak into the first cell', () => {
    // Regression: array carries a LaTeX column-spec arg (e.g. {cc}) that is
    // layout metadata, not content. Without stripping it, cc leaked into the
    // first cell. It must be consumed.
    const out = m('\\begin{array}{cc} 1 & 2 \\\\ 3 & 4 \\end{array}')
    expect(out).toBe('mat(delim: #none, 1, 2; 3, 4)')
    expect(out).not.toMatch(/\bcc\b/)
  })

  it('converts a cases environment to Typst cases() with a quad gap', () => {
    expect(m('\\begin{cases} x^2 & x \\ge 0 \\\\ -x & x < 0 \\end{cases}')).toBe(
      'cases(x^2 quad x gt.eq 0, -x quad x < 0)',
    )
  })

  it('converts a left-braced matrix piecewise function to Typst cases()', () => {
    const latex = 'f \\left(x\\right) = \\left\\{\\begin{matrix} 1 & x > 0 \\\\ 0 & x = 0 \\\\ - 1 & x < 0 \\end{matrix}\\right.'
    const out = m(latex)
    expect(out).toBe('f (x) = cases(1 quad x > 0, 0 quad x = 0, - 1 quad x < 0)')
    expect(out).not.toContain('\\begin{matrix}')

    // Do not overmatch an ordinary unfenced matrix: it remains mat(), not cases().
    expect(m('\\begin{matrix} 1 & 2 \\\\ 3 & 4 \\end{matrix}')).toBe(
      'mat(delim: #none, 1, 2; 3, 4)',
    )
  })

  it('multiplies two matrices without dropping either (adjacent envs)', () => {
    expect(m('\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}\\begin{pmatrix} x \\\\ y \\end{pmatrix}')).toBe(
      'lr(size: #88%, ( mat(delim: #none, a, b; c, d) ))lr(size: #88%, ( mat(delim: #none, x; y) ))',
    )
  })

  it('quotes unknown multi-letter commands as text so Typst never errors', () => {
    // A bare multi-letter ident is an "unknown variable" error in Typst math, so
    // unknown commands are emitted as an upright text string (always compiles).
    expect(m('\\foobar')).toBe('"foobar"')
  })

  it('drops alignment tabs and comments', () => {
    expect(m('a & b')).toBe('a b')
    expect(m('a % comment')).toBe('a')
  })

  it('never emits a raw #/$/backslash that would break out of $...$', () => {
    // Regression for the math converter breaking out of math markup into Typst
    // code mode (a raw `#` reaching name resolution) or leaving an unclosed
    // delimiter — both crashed the whole compile with a 500.
    expect(m('\\left#\\sqrt{x}')).not.toMatch(/[^\\]#/) // # only ever escaped
    expect(m('x^#')).toBe('x^\\#')
    expect(m('x_$')).toBe('x_\\$')
    expect(m('\\left$ x')).toContain('\\$')
    // a base-less script gets a zero-width base so `^2` compiles
    expect(m('^2')).toBe('zws^2')
  })

  it('emits raw LaTeX as a quoted string in verbatim math mode', () => {
    // The route retries in verbatim mode when a convert-pass compile fails, so a
    // malformed formula (e.g. `a__b`, `(^2`) degrades to visible source text
    // instead of 500-ing the whole export. Verbatim output is always a Typst
    // string literal, which compiles regardless of the LaTeX shape.
    const doc = { type: 'doc', content: [{ type: 'blockMath', attrs: { latex: 'a__b' } }] }
    const out = renderTypst(doc, { title: 'T', attachments: new Map(), mathMode: 'verbatim' })
    expect(out).toContain('$ "a__b" $')
  })
})

// ── inline mark rendering ────────────────────────────────────────────────────
describe('renderTypst — marks', () => {
  it('adds deterministic Hangul fallbacks after an explicit source font', () => {
    const out = __test.renderTextNode({
      type: 'text', text: '한국어',
      marks: [{ type: 'textStyle', attrs: { fontFamily: 'Times New Roman', fontSize: '16px' } }],
    })
    expect(out).toContain('font: ("Times New Roman", "Noto Sans CJK KR", "Apple SD Gothic Neo")')
    expect(out).toContain('size: 12.00pt')
  })

  it('renders bold/italic/underline/strike/sup/sub as Typst functions', () => {
    expect(__test.wrapMark({ type: 'bold' }, 'x')).toBe('#text(weight: "bold", stroke: 0.02em)[x]')    // Italic on CJK: skew-synthesized slant wrapping an emph (Latin still italic).
    expect(__test.wrapMark({ type: 'italic' }, 'x')).toBe('#box(skew(ax: -12deg)[#emph[x]])')
    expect(__test.wrapMark({ type: 'underline' }, 'x')).toBe('#underline[x]')
    expect(__test.wrapMark({ type: 'strike' }, 'x')).toBe('#strike[x]')
    expect(__test.wrapMark({ type: 'superscript' }, 'x')).toBe('#super[x]')
    expect(__test.wrapMark({ type: 'subscript' }, 'x')).toBe('#sub[x]')
  })

  it('escapes Typst comment triggers and line-start markup chars in text', () => {
    // `//` must not survive as a Typst line comment: prose or a link-text URL
    // containing `//` would otherwise drop trailing text or comment out a
    // closing delimiter and fail the whole compile.
    const t = (s: string) => __test.renderTextNode({ type: 'text', text: s })
    expect(t('a // b')).toContain('\\/\\/')
    expect(t('https://example.com')).toContain('https:\\/\\/example.com')
    // line-start heading/list markers are neutralised so a paragraph starting
    // `- x` / `= x` / `+ x` is not reinterpreted as a list/heading.
    expect(t('- x')).toContain('\\-')
    expect(t('= x')).toContain('\\=')
    expect(t('+ x')).toContain('\\+')
  })

  it('only emits safe link hrefs, else plain text', () => {
    expect(__test.wrapMark({ type: 'link', attrs: { href: 'https://x.com' } }, 'a')).toBe('#link("https://x.com")[a]')
    // javascript: is unsafe -> plain text
    expect(__test.wrapMark({ type: 'link', attrs: { href: 'javascript:alert(1)' } }, 'a')).toBe('a')
    expect(__test.wrapMark({ type: 'link', attrs: { href: 'java\tscript:alert(1)' } }, 'a')).toBe('a')
  })

  it('applies code mark innermost so bold+code does not leak Typst source', () => {
    // Regression: a run with BOTH bold and code marks must not feed generated
    // markup like #text(weight:"bold")[..] into #raw(); the code chip should
    // wrap the plain text, then bold wraps the chip.
    const out = __test.renderTextNode({
      type: 'text',
      text: 'BoldCode',
      marks: [{ type: 'bold' }, { type: 'code' }],
    })
    expect(out).toContain('#raw("BoldCode")')
    expect(out).not.toContain('#raw("#text')
    expect(out.startsWith('#text(weight: "bold", stroke: 0.02em)[')).toBe(true)
  })

  it('whitelists highlight/textStyle colours and sizes', () => {
    expect(__test.wrapMark({ type: 'highlight', attrs: { color: '#ff0' } }, 'x')).toContain('#highlight(fill: rgb("#ff0"))')
    // unsafe colour -> plain highlight, no injection
    expect(__test.wrapMark({ type: 'highlight', attrs: { color: 'red;position:fixed' } }, 'x')).toBe('#highlight[x]')
    const ts = __test.wrapMark({ type: 'textStyle', attrs: { color: 'red', fontSize: '16px' } }, 'x')
    expect(ts).toContain('fill: red')
    expect(ts).toContain('size: 12.00pt') // 16px * 0.75
  })

  it('whitelists the v16 textStyle fontFamily into a Typst font: argument', () => {
    // A single named family becomes a quoted string; a comma list becomes a tuple
    // with the generic CSS keyword (sans-serif/serif/…) dropped.
    const single = __test.wrapMark({ type: 'textStyle', attrs: { fontFamily: 'Georgia' } }, 'x')
    expect(single).toBe('#text(font: "Georgia")[x]')
    const list = __test.wrapMark({ type: 'textStyle', attrs: { fontFamily: 'Inter, sans-serif' } }, 'x')
    expect(list).toBe('#text(font: "Inter")[x]')
    const multi = __test.wrapMark({ type: 'textStyle', attrs: { fontFamily: '"Helvetica Neue", Arial, sans-serif' } }, 'x')
    expect(multi).toBe('#text(font: ("Helvetica Neue", "Arial"))[x]')
    // Font + size + colour compose in one #text() call.
    const combined = __test.wrapMark(
      { type: 'textStyle', attrs: { color: 'red', fontSize: '16px', fontFamily: 'Inter' } },
      'x',
    )
    expect(combined).toBe('#text(fill: red, size: 12.00pt, font: "Inter")[x]')
  })

  it('cssFontFamilyToTypst drops unsafe/generic-only families (no injection)', () => {
    expect(__test.cssFontFamilyToTypst('Inter')).toBe('"Inter"')
    expect(__test.cssFontFamilyToTypst('sans-serif')).toBeNull()
    expect(__test.cssFontFamilyToTypst('Arial"); #set page(')).toBeNull()
    expect(__test.cssFontFamilyToTypst('')).toBeNull()
  })

  it('maps CJK font names to the embedded OSS CJK families (direction A, not passthrough)', () => {
    // The v16 fontFamily fix: CJK names used to fail the ASCII-only whitelist,
    // get dropped, and the text silently fell back to the preamble default
    // (octo-docs-backend#62). They now resolve to a guaranteed-present embedded
    // OSS family so the user's serif-vs-sans choice actually renders in the PDF.
    // 宋体 / 仿宋 -> serif (Noto Serif CJK SC == Source Han Serif).
    expect(__test.cssFontFamilyToTypst('宋体')).toBe('"Noto Serif CJK SC"')
    expect(__test.cssFontFamilyToTypst('SimSun')).toBe('"Noto Serif CJK SC"')
    expect(__test.cssFontFamilyToTypst('仿宋')).toBe('"Noto Serif CJK SC"')
    expect(__test.cssFontFamilyToTypst('思源宋体')).toBe('"Noto Serif CJK SC"')
    // 黑体 / 微软雅黑 / 苹方 -> sans (Noto Sans CJK SC == Source Han Sans).
    expect(__test.cssFontFamilyToTypst('黑体')).toBe('"Noto Sans CJK SC"')
    expect(__test.cssFontFamilyToTypst('微软雅黑')).toBe('"Noto Sans CJK SC"')
    expect(__test.cssFontFamilyToTypst('Microsoft YaHei')).toBe('"Noto Sans CJK SC"')
    expect(__test.cssFontFamilyToTypst('PingFang SC')).toBe('"Noto Sans CJK SC"')
    expect(__test.cssFontFamilyToTypst('思源黑体')).toBe('"Noto Sans CJK SC"')
    // An unmapped CJK face still honours the serif/sans intent by classifier
    // (华文楷体 -> no serif marker -> sans; 华文中宋 -> 宋 -> serif) and never
    // leaks the raw name into the source.
    expect(__test.cssFontFamilyToTypst('华文中宋')).toBe('"Noto Serif CJK SC"')
    expect(__test.cssFontFamilyToTypst('未知字体')).toBe('"Noto Sans CJK SC"')
  })

  it('maps and de-duplicates a mixed CJK + ASCII stack', () => {
    // A real editor stack: quoted CJK primary, ASCII fallback, generic keyword.
    // The two sans CJK names collapse onto one embedded family (de-duped), the
    // Latin fallback survives, the generic keyword is dropped.
    expect(__test.cssFontFamilyToTypst('"微软雅黑", "PingFang SC", sans-serif')).toBe(
      '"Noto Sans CJK SC"',
    )
    expect(__test.cssFontFamilyToTypst('宋体, Georgia, serif')).toBe(
      '("Noto Serif CJK SC", "Georgia")',
    )
  })

  it('CJK passthrough stays injection-safe (mapping, never the raw name)', () => {
    // A malicious CJK-bearing family name must never reach the Typst source: the
    // name is classified (contains 宋 -> serif) and the fixed literal is emitted,
    // so the injection payload is discarded entirely.
    const attack = '宋体"); #import "/etc/passwd": *; ('
    const out = __test.cssFontFamilyToTypst(attack)
    expect(out).toBe('"Noto Serif CJK SC"')
    expect(out).not.toContain('#import')
    expect(out).not.toContain('passwd')
    // Rendered into a mark it produces only the safe #text() call.
    expect(__test.wrapMark({ type: 'textStyle', attrs: { fontFamily: attack } }, 'x')).toBe(
      '#text(font: "Noto Serif CJK SC")[x]',
    )
    // A CJK name whose only ASCII is an injection attempt: the ASCII segment
    // fails the plain-name whitelist and is dropped, the CJK segment maps.
    expect(__test.cssFontFamilyToTypst('黑体, Arial"); #set page(')).toBe('"Noto Sans CJK SC"')
  })

  it('cssColorToTypst rejects unknown/dangerous values', () => {
    expect(__test.cssColorToTypst('#abc')).toBe('rgb("#abc")')
    expect(__test.cssColorToTypst('rgb(1,2,3)')).toBe('rgb(1, 2, 3)')
    expect(__test.cssColorToTypst('red')).toBe('red')
    expect(__test.cssColorToTypst('expression(alert(1))')).toBeNull()
  })

  it('separates a mapped symbol / greek letter from a following digit', () => {
    // Regression: mapped symbols and GREEK emitted a LEADING-space-only
    // separator, so a following digit fused into an invalid token
    // (`x\\to0` -> `arrow.r0`, `\\alpha2` -> `alpha2`) that Typst rejects as an
    // unknown variable. They now emit a trailing space too.
    expect(m('x\\to0')).toBe('x arrow.r 0')
    expect(m('\\alpha2')).toBe('alpha 2')
    // Still separated from a preceding identifier as before.
    expect(m('\\cos\\theta')).toBe('cos theta')
  })

  it('preserves widehat semantics instead of rendering the command name literally', () => {
    expect(m('x \\widehat{\\left\\{2\\right\\}}')).toBe('x hat(brace.l 2 brace.r)')
  })

  it('bounds deeply nested brace groups without overflowing the stack', () => {
    // Regression: readBraceGroup used a local `depth` brace-balance counter that
    // shadowed the outer recursion-depth parameter, so the MAX_MATH_NESTING
    // guard tested the (post-scan zero) balance counter and never fired, and the
    // recursive call always passed depth 1. Deeply nested `{{{...}}}` recursed
    // unbounded into a RangeError (stack overflow) that escapes the compile
    // catch and 500s the export. The guard now uses the real recursion depth.
    const deep = '{'.repeat(200) + 'x' + '}'.repeat(200)
    expect(() => m(deep)).not.toThrow()
  })
})

// ── escaping / injection safety ──────────────────────────────────────────────
describe('renderTypst — escaping', () => {
  it('escapes Typst markup special chars in body text', () => {
    const out = typ([para([text('a *b* #c $d$ [e] `f`')])])
    // The literal asterisks / hash / dollar / brackets / backtick must be escaped
    // so they render as text, not Typst syntax.
    expect(out).toContain('\\*b\\*')
    expect(out).toContain('\\#c')
    expect(out).toContain('\\$d\\$')
    expect(out).toContain('\\[e\\]')
    expect(out).toContain('\\`f\\`')
  })

  it('escapes double-quotes/backslashes in link href string literal', () => {
    const out = __test.wrapMark({ type: 'link', attrs: { href: 'https://x.com/"a\\b' } }, 't')
    expect(out).toBe('#link("https://x.com/\\"a\\\\b")[t]')
  })

  it('escapes title in preamble', () => {
    const out = renderTypst(doc([]), { title: 'He said "hi" \\ bye', attachments: new Map() })
    expect(out).toContain('#set document(title: "He said \\"hi\\" \\\\ bye")')
  })
})

// ── node coverage ────────────────────────────────────────────────────────────
describe('renderTypst — nodes', () => {
  it('renders headings by level', () => {
    expect(typ([{ type: 'heading', attrs: { level: 2 }, content: [text('H')] }])).toContain('== H')
  })

  it('pins a distinct font size per heading level in the preamble', () => {
    // Regression: without explicit per-level sizing, Typst's default heading
    // scale is nearly indistinguishable with a CJK body font, so H1..H6 all
    // looked the same size in the exported PDF.
    const out = typ([{ type: 'heading', attrs: { level: 1 }, content: [text('H')] }])
    expect(out).toContain('#show heading.where(level: 1): set text(size: 22pt)')
    expect(out).toContain('#show heading.where(level: 2): set text(size: 18pt)')
    expect(out).toContain('#show heading.where(level: 6): set text(size: 11pt)')
  })

  it('inserts break opportunities into a long unbroken run so it does not overflow', () => {
    // Regression: a 100-char `aaaa...` (or a long URL) will not break in a
    // justified paragraph and runs off the right margin. escContent injects a
    // zero-width space (U+200B) every 20 chars into runs of 40+ non-space chars.
    const longWord = 'a'.repeat(100)
    const out = typ([para([text(longWord)])])
    expect(out).toContain('\u200B')
    // A short word is left untouched (no stray break opportunities).
    const shortOut = typ([para([text('hello')])])
    expect(shortOut).not.toContain('\u200B')
  })

  it('never splits an escape sequence when inserting break opportunities', () => {
    // Regression: escContent escapes special chars (# -> \#) BEFORE inserting the
    // zero-width breaks. A naive fixed-width chunk could drop the break between a
    // backslash and the char it escapes, so the backslash would escape the ZWSP
    // instead and re-expose `#`/`$`/`[` to Typst markup (compile failure / stray
    // parsing). The break must never land right after a lone escaping backslash.
    const out = __test.escContent('x' + '#'.repeat(40))
    expect(out).not.toMatch(/\\\u200B/) // no backslash immediately before a ZWSP
    // Every `#` is still escaped as a `\#` pair.
    expect(out).not.toMatch(/(^|[^\\])#/)
  })

  it('renders bullet and ordered lists', () => {
    const b = typ([{ type: 'bulletList', content: [
      { type: 'listItem', content: [para([text('one')])] },
      { type: 'listItem', content: [para([text('two')])] },
    ] }])
    expect(b).toContain('#list(')
    expect(b).toContain('one')
    expect(b).toContain('two')
    const o = typ([{ type: 'orderedList', attrs: { start: 3 }, content: [
      { type: 'listItem', content: [para([text('x')])] },
    ] }])
    expect(o).toContain('#enum(')
    expect(o).toContain('start: 3')
  })

  it('renders task items with checkboxes', () => {
    const out = typ([{ type: 'taskList', content: [
      { type: 'taskItem', attrs: { checked: true }, content: [para([text('done')])] },
      { type: 'taskItem', attrs: { checked: false }, content: [para([text('todo')])] },
    ] }])
    expect(out).toContain('☑')
    expect(out).toContain('☐')
  })

  it('renders code blocks as fenced raw with sanitized language', () => {
    const out = typ([{ type: 'codeBlock', attrs: { language: 'js; rm -rf' }, content: [text('const x=1')] }])
    // Unsafe language dropped, fence still present with the code text.
    expect(out).toContain('```')
    expect(out).toContain('const x=1')
    expect(out).not.toContain('rm -rf')
  })

  it('renders tables with header cells bold and correct column count', () => {
    const out = typ([{ type: 'table', content: [
      { type: 'tableRow', content: [
        { type: 'tableHeader', content: [para([text('A')])] },
        { type: 'tableHeader', content: [para([text('B')])] },
      ] },
      { type: 'tableRow', content: [
        { type: 'tableCell', content: [para([text('1')])] },
        { type: 'tableCell', content: [para([text('2')])] },
      ] },
    ] }])
    expect(out).toContain('table.header(')
    expect(out).toContain('#text(weight: "bold")[A]')
  })

  it('renders callouts with variant fill', () => {
    const out = typ([{ type: 'callout', attrs: { variant: 'warn' }, content: [para([text('careful')])] }])
    expect(out).toContain('#block(fill: rgb("#fef3e6")')
    expect(out).toContain('careful')
  })

  it('renders block and inline math', () => {
    const out = typ([
      { type: 'blockMath', attrs: { latex: '\\frac{1}{2}' } },
      para([{ type: 'inlineMath', attrs: { latex: 'x^2' } }]),
    ])
    expect(out).toContain('$ frac(1, 2) $')
    expect(out).toContain('$x^2$')
  })

  it('resolves emoji shortcodes to glyphs', () => {
    const out = typ([para([{ type: 'emoji', attrs: { name: 'rocket' } }])])
    expect(out).toContain('🚀')
  })

  it('drops image nodes without a resolved local path', () => {
    const out = typ([{ type: 'image', attrs: { attachId: 'nope' } }])
    // No figure is emitted for an unresolved image (the preamble still defines
    // the __capImage helper, so we assert no #figure(__capImage(...)) call).
    expect(out).not.toContain('__capImage("')
    expect(out).not.toContain('#figure(')
  })

  it('uses the image align attr for raster and SVG assets instead of Typst figure centering', () => {
    const attachments = new Map<string, ResolvedAttachment>([
      ['png-left', { url: '', fileName: 'left.png', mime: 'image/png', sizeBytes: 1 }],
      ['svg-center', { url: '', fileName: 'center.svg', mime: 'image/svg+xml', sizeBytes: 1 }],
      ['svg-right', { url: '', fileName: 'right.svg', mime: 'image/svg+xml', sizeBytes: 1 }],
    ])
    const out = renderTypst(doc([
      { type: 'image', attrs: { attachId: 'png-left', align: 'left' } },
      { type: 'image', attrs: { attachId: 'svg-center', align: 'center', width: 80 } },
      { type: 'image', attrs: { attachId: 'svg-right', align: 'right' } },
    ]), {
      title: 'T',
      attachments,
      imagePaths: new Map([
        ['png-left', 'left.png'],
        ['svg-center', 'center.svg'],
        ['svg-right', 'right.svg'],
      ]),
    })

    expect(out).toContain('#align(left)[#figure(__capImage("left.png"))]')
    expect(out).toContain('#align(center)[#figure(__capImage("center.svg", w: 60.0pt))]')
    expect(out).toContain('#align(right)[#figure(__capImage("right.svg"))]')
  })

  it('defaults an image without align to the editor left alignment, not Typst center', () => {
    const out = renderTypst(doc([
      { type: 'image', attrs: { attachId: 'plain' } },
    ]), {
      title: 'T',
      attachments: new Map([['plain', { url: '', fileName: 'plain.png', mime: 'image/png', sizeBytes: 1 }]]),
      imagePaths: new Map([['plain', 'plain.png']]),
    })
    expect(out).toContain('#align(left)[#figure(__capImage("plain.png"))]')
  })

  it('inherits paragraph alignment for legacy/imported image children when image align is absent', () => {
    const out = renderTypst(doc([
      { type: 'paragraph', attrs: { textAlign: 'right' }, content: [
        { type: 'image', attrs: { attachId: 'legacy' } },
      ] },
    ]), {
      title: 'T',
      attachments: new Map([['legacy', { url: '', fileName: 'legacy.svg', mime: 'image/svg+xml', sizeBytes: 1 }]]),
      imagePaths: new Map([['legacy', 'legacy.svg']]),
    })
    expect(out).toContain('#align(right)[#figure(__capImage("legacy.svg"))]')
  })

  it('clamps pathological table span attrs (DoS guard)', () => {
    const out = typ([{ type: 'table', content: [
      { type: 'tableRow', content: [
        { type: 'tableCell', attrs: { colspan: 1e9 }, content: [para([text('x')])] },
      ] },
    ] }])
    // colspan clamped to <=100, so the column count stays sane (<=100 `1fr`
    // tracks), never 1e9 tracks.
    const match = out.match(/columns: \(([^)]*)\)/)
    expect(match).not.toBeNull()
    const trackCount = match![1]!.split(',').length
    expect(trackCount).toBeLessThanOrEqual(100)
  })

  it('clamps lopsided nested-table column widths so narrow columns are not crushed', () => {
    // A nested table (rendered inside a cell) with page-scale, lopsided editor
    // colwidths (e.g. an image column 600px vs 100/90px siblings) previously
    // became 15fr/2.5fr/2.25fr (~75/13/12) and crushed columns 2 & 3 to slivers.
    // The clamp bounds each column to [0.6, 1.6]x an equal share.
    const wcell = (t: string, w: number) => ({
      type: 'tableCell',
      attrs: { colwidth: [w] },
      content: [para([text(t)])],
    })
    const inner = {
      type: 'table',
      content: [
        { type: 'tableRow', content: [wcell('a', 600), wcell('b', 100), wcell('c', 90)] },
      ],
    }
    const out = typ([
      {
        type: 'table',
        content: [
          {
            type: 'tableRow',
            content: [
              { type: 'tableCell', content: [para([text('x')])] },
              { type: 'tableCell', content: [inner] },
            ],
          },
        ],
      },
    ])
    // The nested table's columns arg is the last one emitted.
    const cols = [...out.matchAll(/columns: \(([^)]*)\)/g)].map((m) => m[1]!)
    const innerCols = cols[cols.length - 1]!
    const frs = innerCols
      .split(',')
      .map((s) => parseFloat(s.trim().replace('fr', '')))
    expect(frs.length).toBe(3)
    // Every column within [0.6, 1.6]x an equal (1fr) share.
    for (const f of frs) {
      expect(f).toBeGreaterThanOrEqual(0.6 - 0.01)
      expect(f).toBeLessThanOrEqual(1.6 + 0.01)
    }
    // The widest column (a=600) hits the ceiling; the two narrow ones hit the floor.
    expect(Math.max(...frs)).toBeCloseTo(1.6, 1)
    expect(Math.min(...frs)).toBeCloseTo(0.6, 1)
  })
})

describe('renderTypst — v17 paragraph spacing (lineHeight / spaceBefore / spaceAfter)', () => {
  it('maps lineHeight to par(leading) as (L-1)em', () => {
    const out = typ([{ type: 'paragraph', attrs: { lineHeight: '1.5' }, content: [text('hi')] }])
    expect(out).toContain('#set par(leading: 0.500em)')
  })

  it('maps spaceBefore/spaceAfter (px) to a block above/below in pt', () => {
    const out = typ([
      { type: 'paragraph', attrs: { spaceBefore: '8px', spaceAfter: '12px' }, content: [text('hi')] },
    ])
    expect(out).toContain('#block(above: 6.0pt, below: 9.0pt)')
  })

  it('leaves a plain paragraph free of block/leading wrapping', () => {
    const out = typ([para([text('plain')])])
    expect(out).not.toContain('#set par(leading')
    expect(out).not.toContain('#block(above')
  })

  it('drops out-of-range spacing (>1000) to match the schema sanitizer cap', () => {
    const out = typ([
      { type: 'paragraph', attrs: { spaceBefore: '10000px', spaceAfter: '1001em' }, content: [text('hi')] },
    ])
    // 10000px would map to 7500.0pt and 1001em to 1001em if not capped; both must be dropped.
    expect(out).not.toContain('7500.0pt')
    expect(out).not.toContain('1001em')
    expect(out).not.toContain('#block(above')
  })

  it('keeps spacing at the 1000 boundary (px → pt)', () => {
    const out = typ([
      { type: 'paragraph', attrs: { spaceBefore: '1000px' }, content: [text('hi')] },
    ])
    expect(out).toContain('#block(above: 750.0pt)')
  })

  it('composes alignment with spacing (align stays innermost)', () => {
    const out = typ([
      { type: 'paragraph', attrs: { textAlign: 'center', lineHeight: '2' }, content: [text('c')] },
    ])
    expect(out).toContain('#set par(leading: 1.000em); #align(center)')
  })
})
