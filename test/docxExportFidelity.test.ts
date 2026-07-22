import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { exportDocx } from '../src/export/docx.js'
import type { PmNode } from '../src/export/markdown.js'
import acceptanceFixtures from './fixtures/acceptance-semantic-gaps.json'

function maxXmlDepth(xml: string): number {
  let depth = 0
  let max = 0
  for (const match of xml.matchAll(/<\/?[\w:.-]+(?:\s[^<>]*?)?\s*\/?>/g)) {
    if (match[0].startsWith('</')) depth--
    else if (!match[0].endsWith('/>')) max = Math.max(max, ++depth)
  }
  return max
}

async function parts(doc: PmNode): Promise<{ document: string; styles: string; rels: string }> {
  const zip = await JSZip.loadAsync(await exportDocx(doc))
  return {
    document: await zip.file('word/document.xml')!.async('string'),
    styles: await zip.file('word/styles.xml')!.async('string'),
    rels: await zip.file('word/_rels/document.xml.rels')!.async('string'),
  }
}

describe('DOCX OOXML export fidelity', () => {
  it('emits structured OMML for fractions, matrices, and piecewise cases', async () => {
    const { document } = await parts({ type: 'doc', content: [
      { type: 'blockMath', attrs: { latex: '\\frac{a}{b}' } },
      { type: 'blockMath', attrs: { latex: '\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}' } },
      { type: 'blockMath', attrs: { latex: '\\begin{cases}x&x>0\\\\-x&x\\leq0\\end{cases}' } },
    ] })
    expect(document).toContain('<m:f>')
    expect(document).toContain('<m:m>')
    expect(document).toContain('<m:begChr m:val="("')
    expect(document).toContain('<m:begChr m:val="{"')
    expect(document).not.toContain('OCTOMATH')
    expect(document).not.toContain('\\frac')
  })

  it('reimports exporter-generated OMML deeper than the parser default without losing math or styles', async () => {
    let latex = 'x'
    for (let i = 1; i <= 50; i++) latex = `\\frac{${latex}}{${i}}`
    const marks = [
      { type: 'bold' },
      { type: 'textStyle', attrs: { color: '#123456', fontSize: '18px', fontFamily: 'Arial' } },
    ]
    const source: PmNode = { type: 'doc', content: [{ type: 'paragraph', content: [
      { type: 'text', text: 'styled', marks },
      { type: 'inlineMath', attrs: { latex } },
    ] }] }

    const bytes = await exportDocx(source)
    const zip = await JSZip.loadAsync(bytes)
    const document = await zip.file('word/document.xml')!.async('string')
    expect(maxXmlDepth(document)).toBeGreaterThan(100)

    const restored = (await importDocx(Buffer.from(bytes))).doc
    expect(restored.content?.[0]?.content).toEqual(source.content?.[0]?.content)
  })

  it('round-trips the Unicode floor formula that previously fell back to Consolas text', async () => {
    const latex = '\\left⌊x\\right.⌋ \\leq x < \\left⌊x\\right.⌋ + 1'
    const source: PmNode = { type: 'doc', content: [{ type: 'blockMath', attrs: { latex } }] }
    const bytes = await exportDocx(source)
    const zip = await JSZip.loadAsync(bytes)
    const document = await zip.file('word/document.xml')!.async('string')
    expect(document).toContain('<m:oMath')
    expect(document).not.toContain('w:ascii="Consolas"')

    const restored = (await importDocx(Buffer.from(bytes))).doc
    expect(restored.content).toEqual([{ type: 'blockMath', attrs: { latex } }])
  })

  it('preserves the case 03 Office hat formula as one inline OMML node', async () => {
    // Exact minimized PM formula imported from 03-docx-fixed.docx. Office emits
    // the delimited accent argument without braces, which MathJax otherwise
    // rejects and the DOCX serializer degrades to a plain Consolas run.
    const latex = String.raw`x \hat \left\{2\right\}`
    const source: PmNode = { type: 'doc', content: [{ type: 'paragraph', content: [
      { type: 'text', text: 'formula ' },
      { type: 'inlineMath', attrs: { latex } },
    ] }] }
    const bytes = await exportDocx(source)
    const zip = await JSZip.loadAsync(bytes)
    const document = await zip.file('word/document.xml')!.async('string')

    expect(document.match(/<m:oMath[ >]/g)).toHaveLength(1)
    expect(document).not.toContain('w:ascii="Consolas"')

    const restored = (await importDocx(Buffer.from(bytes))).doc
    const math = restored.content?.[0]?.content?.filter((node) => node.type === 'inlineMath') ?? []
    expect(math).toHaveLength(1)
    expect(String(math[0]?.attrs?.latex)).toMatch(/^x \\wide?hat\{\\left\\\{2\\right\\\}\}$/)
  })

  it('preserves heading levels, safe text colors/highlights, and hyperlink relationships', async () => {
    const { document, rels } = await parts({ type: 'doc', content: [
      { type: 'heading', attrs: { level: 2 }, content: [{ type: 'text', text: 'Colored', marks: [{ type: 'textStyle', attrs: { color: '#1d72b8', fontSize: '18pt' } }] }] },
      { type: 'paragraph', content: [{ type: 'text', text: 'link', marks: [{ type: 'link', attrs: { href: 'https://example.com/a?b=1' } }, { type: 'highlight', attrs: { color: '#fff3a0' } }] }] },
    ] })
    expect(document).toContain('w:pStyle w:val="Heading2"')
    expect(document).toContain('w:color w:val="1D72B8"')
    expect(document).toContain('w:shd w:fill="FFF3A0"')
    expect(document).toContain('<w:hyperlink')
    expect(rels).toContain('Target="https://example.com/a?b=1"')
  })

  it('normalizes repeatedly escaped ampersands in hyperlink relationship attributes', async () => {
    const source = acceptanceFixtures.group07SpecialText as PmNode
    const bytes = await exportDocx(source)
    const zip = await JSZip.loadAsync(bytes)
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('Target="https://example.com/very/long/path?a=1&amp;b=2&amp;c=3"')
    expect(rels).not.toContain('&amp;amp;')

    const restored = (await importDocx(Buffer.from(bytes))).doc
    expect(canonicalTextInventory(restored).find((item) => item.text.startsWith('https://'))?.marks).toContainEqual({
      type: 'link', attrs: { href: 'https://example.com/very/long/path?a=1&b=2&c=3' },
    })
  })

  it('round-trips nested tasks without adding exporter layout whitespace', async () => {
    const source = acceptanceFixtures.group02NestedTask as PmNode
    const restored = (await importDocx(Buffer.from(await exportDocx(source)))).doc
    expect(restored).toEqual(source)
  })

  it('writes fixed table grid/cell widths, merged cells, row heights, and header shading', async () => {
    const { document } = await parts({ type: 'doc', content: [{ type: 'table', content: [
      { type: 'tableRow', attrs: { height: 32 }, content: [
        { type: 'tableHeader', attrs: { colspan: 2, colwidth: [200, 400] }, content: [{ type: 'paragraph', content: [{ type: 'text', text: 'H' }] }] },
      ] },
      { type: 'tableRow', content: [
        { type: 'tableCell', attrs: { colwidth: [200] }, content: [{ type: 'paragraph', content: [{ type: 'text', text: 'A' }] }] },
        { type: 'tableCell', attrs: { colwidth: [400] }, content: [{ type: 'paragraph', content: [{ type: 'text', text: 'B' }] }] },
      ] },
    ] }] })
    expect((document.match(/<w:gridCol/g) ?? []).length).toBe(2)
    expect(document).toContain('<w:tcW')
    expect(document).toContain('<w:gridSpan w:val="2"')
    expect(document).toContain('<w:trHeight')
    expect(document).toContain('w:fill="F2F3F5"')
  })

  it('represents callout variants, code blocks, and details with distinct round-trip styles', async () => {
    const { document, styles } = await parts({ type: 'doc', content: [
      { type: 'callout', attrs: { variant: 'info' }, content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Info' }] }] },
      { type: 'callout', attrs: { variant: 'warn' }, content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Warning' }] }] },
      { type: 'codeBlock', content: [{ type: 'text', text: 'const x = 1\nreturn x' }] },
      { type: 'details', content: [
        { type: 'detailsSummary', content: [{ type: 'text', text: 'Open me' }] },
        { type: 'detailsContent', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Hidden' }] }] },
      ] },
    ] })
    for (const style of ['CalloutInfo', 'CalloutWarn', 'CodeBlock', 'CodeBlockEnd', 'DetailsStart', 'DetailsSummary', 'DetailsEnd']) {
      expect(document).toContain(`w:pStyle w:val="${style}"`)
      expect(styles).toContain(`w:styleId="${style}"`)
    }
    expect(document).toContain('w:fill="E8F1FD"')
    expect(document).toContain('w:fill="FFF4E5"')
    expect(document).toContain('w:fill="F5F5F5"')
  })
})

import { buildDocFromParts, importDocx } from '../src/import/docx/index.js'

function canonicalTextInventory(node: PmNode): Array<{ text: string; marks: Array<{ type: string; attrs: Record<string, unknown> }> }> {
  const walk = (n: PmNode): PmNode[] => [...(n.type === 'text' ? [n] : []), ...(n.content ?? []).flatMap(walk)]
  return walk(node).map((n) => ({
    text: n.text ?? '',
    marks: [...(n.marks ?? [])].map((m) => ({ type: m.type, attrs: (m.attrs ?? {}) as Record<string, unknown> }))
      .sort((a, b) => a.type.localeCompare(b.type)),
  }))
}

describe('actual shared export → DOCX parser semantic inventory', () => {
  it('preserves source default font and point size through PM and DOCX', async () => {
    const document = Buffer.from('<?xml version="1.0"?><w:document xmlns:w="x"><w:body><w:p><w:r><w:t>正文</w:t></w:r></w:p></w:body></w:document>')
    const styles = Buffer.from('<?xml version="1.0"?><w:styles xmlns:w="x"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="微软雅黑" w:eastAsia="微软雅黑"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>')
    const imported = buildDocFromParts({
      parts: new Map([
        ['word/document.xml', { name: 'word/document.xml', data: document }],
        ['word/styles.xml', { name: 'word/styles.xml', data: styles }],
      ]),
      media: [], warnings: [],
    }).doc
    expect(imported.content?.[0]?.content?.[0]?.marks).toEqual([
      { type: 'textStyle', attrs: { fontFamily: '微软雅黑', fontSize: '16px' } },
    ])
    const zip = await JSZip.loadAsync(await exportDocx(imported))
    const exported = await zip.file('word/document.xml')!.async('string')
    expect(exported).toContain('w:ascii="微软雅黑"')
    expect(exported).toContain('w:sz w:val="24"')
  })

  it('reimports fontFamily, fontSize, color, highlight, and every supported text mark', async () => {
    const source: PmNode = { type: 'doc', content: [{ type: 'paragraph', content: [{
      type: 'text', text: 'inventory', marks: [
        { type: 'bold' }, { type: 'italic' }, { type: 'underline' }, { type: 'strike' },
        { type: 'subscript' },
        { type: 'link', attrs: { href: 'https://example.com/style' } },
        { type: 'highlight', attrs: { color: '#fff3a3' } },
        { type: 'textStyle', attrs: { color: '#123456', fontSize: '18px', fontFamily: 'Arial' } },
      ],
    }, { type: 'text', text: 'super', marks: [{ type: 'superscript' }] }] }] }
    const restored = (await importDocx(Buffer.from(await exportDocx(source)))).doc
    expect(canonicalTextInventory(restored)).toEqual(canonicalTextInventory(source))
  })
})

import userFileMath from './fixtures/user-file-math-regressions.json'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { exportMarkdown } from '../src/export/markdown.js'
import { parseMarkdownToPmDoc } from '../src/import/markdown/markdown.js'

const realMathFixtureDir = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures/user-files')

function mathLatex(node: PmNode): string[] {
  return [
    ...(node.type === 'inlineMath' || node.type === 'blockMath' ? [String(node.attrs?.latex ?? '')] : []),
    ...(node.content ?? []).flatMap(mathLatex),
  ]
}

function semanticMath(latex: string): string {
  return latex
    .replace(/\\textrm\{\s*\}/g, '')
    .replace(/\\begin\{matrix\}|\\end\{matrix\}/g, '')
    .replace(/\\left|\\right/g, '')
    .replace(/\s+/g, '')
}

function assertRenderableMath(formulas: string[]): void {
  expect(formulas).not.toHaveLength(0)
  for (const formula of formulas) {
    expect(formula).not.toMatch(/^\$|\$$/)
    expect(formula).not.toContain('\\\\hat')
    expect(formula).not.toContain('\\\\begin')
  }
}

describe('real Office math semantic round trips', () => {
  for (const fixture of ['03', '04', '11']) {
    it(`${fixture} preserves visible formula semantics through Markdown and DOCX`, async () => {
      const source = await importDocx(fs.readFileSync(`${realMathFixtureDir}/${fixture}-docx-fixed.docx`))
      const sourceMath = mathLatex(source.doc)
      const markdown = exportMarkdown(source.doc)
      const mdMath = mathLatex(parseMarkdownToPmDoc(markdown).doc)
      const bytes = await exportDocx(source.doc)
      const zip = await JSZip.loadAsync(bytes)
      const documentXml = await zip.file('word/document.xml')!.async('string')
      const docxMath = mathLatex((await importDocx(Buffer.from(bytes))).doc)

      assertRenderableMath(sourceMath)
      assertRenderableMath(mdMath)
      assertRenderableMath(docxMath)
      expect(mdMath.map(semanticMath)).toEqual(sourceMath.map(semanticMath))
      expect(docxMath.map(semanticMath)).toEqual(sourceMath.map(semanticMath))
      expect(documentXml).not.toContain('w:ascii="Consolas"')
      expect(documentXml).not.toContain('$\\begin')
      if (fixture === '04') {
        expect(sourceMath[0]).toContain('\\widehat')
        expect(docxMath[0]).not.toBe('x^{2}')
      }
    })
  }
})

describe('fresh user-file math export regressions', () => {
  it('exports case 11 as OMML and reimports exactly one semantic math node', async () => {
    const source: PmNode = { type: 'doc', content: [{ type: 'blockMath', attrs: { latex: userFileMath.case11 } }] }
    const bytes = await exportDocx(source)
    const zip = await JSZip.loadAsync(bytes)
    const document = await zip.file('word/document.xml')!.async('string')
    expect((document.match(/<m:oMath[ >]/g) ?? [])).toHaveLength(1)
    expect(document).not.toContain('w:ascii="Consolas"')
    expect(document).not.toContain('\\sqrt')

    const restored = (await importDocx(Buffer.from(bytes))).doc
    const math = (restored.content ?? []).filter((node) => node.type === 'blockMath')
    expect(math).toHaveLength(1)
    const semantic = String(math[0]?.attrs?.latex)
    expect(semantic).toMatch(/log/)
    expect(semantic).toMatch(/ln/)
    expect(semantic).toContain('=')
  })
})
