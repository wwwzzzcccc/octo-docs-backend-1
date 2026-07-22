import { describe, it, expect } from 'vitest'
import { execFileSync } from 'node:child_process'
import { renderTypst } from '../src/export/renderTypst.js'
import { compileTypst } from '../src/export/typstService.js'
import { probeFailingFormulas } from '../src/api/routes/export.js'

// The typst binary may not be present in every CI image; probe once and skip the
// real-compile suite when it's missing (the pure-transform tests in
// renderTypst.test.ts still run everywhere).
function typstAvailable(): boolean {
  try {
    execFileSync(process.env.TYPST_EXPORT_BINARY || 'typst', ['--version'], { stdio: 'ignore' })
    return true
  } catch {
    return false
  }
}

const HAS_TYPST = typstAvailable()
const d = HAS_TYPST ? describe : describe.skip

const doc = (content: unknown[]) => ({ type: 'doc', content })
const para = (content: unknown[]) => ({ type: 'paragraph', content })
const text = (t: string, marks?: unknown[]) => ({ type: 'text', text: t, ...(marks ? { marks } : {}) })

d('typstService — real compile (integration)', () => {
  it('embeds extractable Korean text instead of dropping Hangul glyphs', async () => {
    const korean = '한국어 문서 글꼴 테스트'
    const src = renderTypst(
      doc([para([text(korean, [{ type: 'textStyle', attrs: { fontFamily: 'Times New Roman', fontSize: '16px' } }])])]),
      { title: 'Korean regression', attachments: new Map() },
    )
    const pdf = await compileTypst(src)
    expect(pdf.subarray(0, 5).toString('latin1')).toBe('%PDF-')
    const extracted = execFileSync('pdftotext', ['-', '-'], { input: pdf }).toString('utf8')
    expect(extracted).toContain(korean)
  })

  it('embeds a sanitized SVG image in the compiled PDF', async () => {
    const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="80" height="40"><rect width="80" height="40" fill="#e11d48"/><circle cx="40" cy="20" r="12" fill="white"/></svg>')
    const src = renderTypst(
      doc([{ type: 'image', attrs: { attachId: 'svg-1', width: '80px' } }]),
      {
        title: 'SVG image',
        attachments: new Map([['svg-1', { url: '', fileName: 'shape.svg', mime: 'image/svg+xml', sizeBytes: svg.length }]]),
        imagePaths: new Map([['svg-1', 'img_0.svg']]),
      },
    )
    expect(src).toContain('__capImage("img_0.svg"')
    const pdf = await compileTypst(src, [{ fileName: 'img_0.svg', bytes: svg }])
    expect(pdf.subarray(0, 5).toString('latin1')).toBe('%PDF-')
    expect(pdf.length).toBeGreaterThan(1000)
  })

  it('compiles a rich CJK + math + table + code document to a valid PDF', async () => {
    const src = renderTypst(
      doc([
        { type: 'heading', attrs: { level: 2 }, content: [text('章节 Section')] },
        para([
          text('中文正文，含 '),
          text('粗体', [{ type: 'bold' }]),
          text(' 与 '),
          text('斜体', [{ type: 'italic' }]),
          text(' 及链接 '),
          text('Octo', [{ type: 'link', attrs: { href: 'https://example.com' } }]),
        ]),
        { type: 'blockMath', attrs: { latex: '\\sum_{n=1}^{\\infty} \\frac{1}{n^2} = \\frac{\\pi^2}{6}' } },
        { type: 'bulletList', content: [
          { type: 'listItem', content: [para([text('第一项')])] },
          { type: 'listItem', content: [para([text('第二项')])] },
        ] },
        { type: 'callout', attrs: { variant: 'warn' }, content: [para([text('注意事项')])] },
        { type: 'table', content: [
          { type: 'tableRow', content: [
            { type: 'tableHeader', content: [para([text('列A')])] },
            { type: 'tableHeader', content: [para([text('列B')])] },
          ] },
          { type: 'tableRow', content: [
            { type: 'tableCell', content: [para([text('数据1')])] },
            { type: 'tableCell', content: [para([text('数据2')])] },
          ] },
        ] },
        { type: 'codeBlock', attrs: { language: 'js' }, content: [text('const x = 1\nconsole.log(x)')] },
      ]),
      { title: '测试文档', attachments: new Map() },
    )

    const pdf = await compileTypst(src)
    // Valid PDF magic header + non-trivial size.
    expect(pdf.subarray(0, 5).toString('latin1')).toBe('%PDF-')
    expect(pdf.length).toBeGreaterThan(1000)
  })

  it('never emits a document that fails to compile for exotic math (fallback path)', async () => {
    const src = renderTypst(
      doc([{ type: 'blockMath', attrs: { latex: '\\weirdunknownmacro{\\foo}_{\\bar}^{42}' } }]),
      { title: 't', attachments: new Map() },
    )
    const pdf = await compileTypst(src)
    expect(pdf.subarray(0, 5).toString('latin1')).toBe('%PDF-')

  })

  it('compiles a left-braced matrix piecewise formula as native Typst cases', async () => {
    const latex = 'f \\left(x\\right) = \\left\\{\\begin{matrix} 1 & x > 0 \\\\ 0 & x = 0 \\\\ - 1 & x < 0 \\end{matrix}\\right.'
    const src = renderTypst(
      doc([{ type: 'blockMath', attrs: { latex } }]),
      { title: 'piecewise', attachments: new Map() },
    )
    expect(src).toContain('$ f (x) = cases(1 quad x > 0, 0 quad x = 0, - 1 quad x < 0) $')
    expect(src).not.toContain('\\begin{matrix}')
    const pdf = await compileTypst(src)
    expect(pdf.subarray(0, 5).toString('latin1')).toBe('%PDF-')
    expect(pdf.length).toBeGreaterThan(1000)
  })

  it('keeps valid math typeset in a 30-formula document with several malformed/corrupt formulas', async () => {
    const valid = Array.from({ length: 27 }, (_, i) => `\\frac{x_${i + 1}}{${i + 2}} + \\alpha`)
    // x^ and y_ are Typst-invalid after conversion. The third case also carries
    // a persisted form-feed before alpha, exercising raw-key matching after the
    // renderer sanitizes control bytes.
    const malformed = ['x^', 'y_', `z^^ + \u000calpha`]
    const formulas = [
      ...valid.slice(0, 9), malformed[0]!,
      ...valid.slice(9, 18), malformed[1]!,
      ...valid.slice(18), malformed[2]!,
    ]
    expect(formulas).toHaveLength(30)

    const result = await probeFailingFormulas(formulas, 'math30', { maxProbes: 40, budgetMs: 15_000 })
    expect(result.exhausted).toBe(false)
    expect(result.failing).toEqual(new Set(malformed))

    const src = renderTypst(
      doc(formulas.map((latex) => ({ type: 'blockMath', attrs: { latex } }))),
      { title: 'math30', attachments: new Map(), verbatimFormulas: result.failing },
    )
    // Valid formulas remain native Typst math, while only the exact bad raw keys
    // become quoted text. No control byte may leak into the generated source.
    expect(src).toContain('$ frac(x_1, 2) + alpha $')
    expect(src).toContain('$ "x^" $')
    expect(src).toContain('$ "y_" $')
    expect(src).toContain('$ "z^^ + alpha" $')
    // eslint-disable-next-line no-control-regex -- regression assertion
    expect(src).not.toMatch(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/)
    const pdf = await compileTypst(src)
    expect(pdf.subarray(0, 5).toString('latin1')).toBe('%PDF-')

    // Even when the probe allowance expires, a batch already proven valid must
    // stay native math. With two probes: the full set fails, the valid left half
    // succeeds, and only the unresolved right half is degraded.
    const budgeted = await probeFailingFormulas(
      [...valid.slice(0, 15), malformed[0]!, ...valid.slice(15), 'q_1', 'q_2'],
      'math30-budgeted',
      { maxProbes: 2, budgetMs: 15_000 },
    )
    expect(budgeted.exhausted).toBe(true)
    expect(budgeted.failing).not.toContain(valid[0])
    expect(budgeted.failing).toContain(malformed[0])
    expect(budgeted.failing.size).toBe(15)
  })

  it('compiles CJK text styled with mapped CJK fonts (embedded OSS faces resolve)', async () => {
    // Regression for octo-docs-backend#62: a document whose text picks CJK fonts
    // (宋体 -> serif, 微软雅黑 -> sans) must map to the embedded OSS families and
    // compile cleanly against the runtime font book (no missing-font failure,
    // no silent fallback). The source is asserted to carry the mapped families.
    const src = renderTypst(
      doc([
        para([text('宋体正文', [{ type: 'textStyle', attrs: { fontFamily: '宋体' } }])]),
        para([text('黑体标题', [{ type: 'textStyle', attrs: { fontFamily: '微软雅黑' } }])]),
      ]),
      { title: '字体测试', attachments: new Map() },
    )
    expect(src).toContain('font: "Noto Serif CJK SC"')
    expect(src).toContain('font: "Noto Sans CJK SC"')
    const pdf = await compileTypst(src)
    expect(pdf.subarray(0, 5).toString('latin1')).toBe('%PDF-')
    expect(pdf.length).toBeGreaterThan(1000)
  })
})

import userFileMath from './fixtures/user-file-math-regressions.json'
import { createCanvas, loadImage } from '@napi-rs/canvas'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

async function renderedInkBounds(pdf: Buffer): Promise<{ width: number; height: number }> {
  const dir = await mkdtemp(join(tmpdir(), 'typst-visual-'))
  try {
    const pdfPath = join(dir, 'page.pdf')
    const pngPath = join(dir, 'page.png')
    await writeFile(pdfPath, pdf)
    execFileSync('pdftoppm', ['-f', '1', '-singlefile', '-r', '144', '-png', pdfPath, join(dir, 'page')])
    const image = await loadImage(await readFile(pngPath))
    const canvas = createCanvas(image.width, image.height)
    const ctx = canvas.getContext('2d')
    ctx.drawImage(image, 0, 0)
    const pixels = ctx.getImageData(0, 0, image.width, image.height).data
    let minX = image.width; let minY = image.height; let maxX = -1; let maxY = -1
    // Ignore the generated title area. Regression content starts below ~110 px at 144 dpi.
    for (let y = 115; y < image.height; y++) for (let x = 0; x < image.width; x++) {
      const i = (y * image.width + x) * 4
      if (pixels[i]! < 235 || pixels[i + 1]! < 235 || pixels[i + 2]! < 235) {
        minX = Math.min(minX, x); minY = Math.min(minY, y); maxX = Math.max(maxX, x); maxY = Math.max(maxY, y)
      }
    }
    return { width: maxX < 0 ? 0 : maxX - minX + 1, height: maxY < 0 ? 0 : maxY - minY + 1 }
  } finally { await rm(dir, { recursive: true, force: true }) }
}

d('fresh user-file visual regressions', () => {
  it('renders case 03 Office matrix at a meaningful size without literal textrm', async () => {
    const src = renderTypst(doc([para([{ type: 'inlineMath', attrs: { latex: userFileMath.case03 } }])]), { title: 'case03', attachments: new Map() })
    expect(src).not.toContain('"textrm"')
    const pdf = await compileTypst(src)
    const extracted = execFileSync('pdftotext', ['-', '-'], { input: pdf }).toString('utf8')
    expect(extracted).not.toMatch(/textrm/i)
    const bounds = await renderedInkBounds(pdf)
    expect(bounds.width).toBeGreaterThan(120)
    expect(bounds.height).toBeGreaterThan(35)
  })

  it('renders case 11 as semantic math rather than raw LaTeX', async () => {
    const src = renderTypst(doc([{ type: 'blockMath', attrs: { latex: userFileMath.case11 } }]), { title: 'case11', attachments: new Map() })
    const pdf = await compileTypst(src)
    const extracted = execFileSync('pdftotext', ['-', '-'], { input: pdf }).toString('utf8')
    expect(extracted).not.toMatch(/\\(?:sqrt|begin|textrm|backslash)/)
    expect(extracted).toMatch(/log|ln/)
    const bounds = await renderedInkBounds(pdf)
    expect(bounds.width).toBeGreaterThan(150)
    expect(bounds.height).toBeGreaterThan(30)
  })

  it('upscales a tiny intrinsic image to a visible bounded size', async () => {
    const canvas = createCanvas(40, 40)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#e11d48'; ctx.fillRect(0, 0, 40, 40)
    const image = canvas.toBuffer('image/png')
    const src = renderTypst(doc([{ type: 'image', attrs: { attachId: 'tiny' } }]), {
      title: 'case06', attachments: new Map([['tiny', { url: '', fileName: 'tiny.png', mime: 'image/png', sizeBytes: image.length }]]),
      imagePaths: new Map([['tiny', 'tiny.png']]),
    })
    const pdf = await compileTypst(src, [{ fileName: 'tiny.png', bytes: image }])
    const bounds = await renderedInkBounds(pdf)
    expect(bounds.width).toBeGreaterThanOrEqual(90)
    expect(bounds.height).toBeGreaterThanOrEqual(90)
  })
})
