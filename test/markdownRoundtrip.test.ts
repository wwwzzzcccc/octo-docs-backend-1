import { describe, expect, it } from 'vitest'
import { exportMarkdown, type PmNode } from '../src/export/markdown.js'
import { parseMarkdownToPmDoc } from '../src/import/markdown/markdown.js'
import acceptanceFixtures from './fixtures/acceptance-semantic-gaps.json'

const text = (value: string, marks?: NonNullable<PmNode['marks']>): PmNode => ({
  type: 'text', text: value, ...(marks ? { marks } : {}),
})
const paragraph = (...content: PmNode[]): PmNode => ({ type: 'paragraph', content })
const doc = (...content: PmNode[]): PmNode => ({ type: 'doc', content })

function allTextNodes(node: PmNode): PmNode[] {
  return [...(node.type === 'text' ? [node] : []), ...(node.content ?? []).flatMap(allTextNodes)]
}

describe('server Markdown export/import round-trip', () => {
  it('restores non-Markdown style marks through allowlisted inline HTML', () => {
    const source = doc(paragraph(
      text('under', [{ type: 'underline' }]), text(' '),
      text('highlight', [{ type: 'highlight', attrs: { color: '#fff3a3' } }]), text(' '),
      text('styled', [{ type: 'textStyle', attrs: { color: 'rgb(10, 20, 30)', fontSize: '18px' } }]), text(' '),
      text('sub', [{ type: 'subscript' }]), text(' '),
      text('super', [{ type: 'superscript' }]),
    ))

    const markdown = exportMarkdown(source)
    expect(markdown).toContain('<u>under</u>')
    expect(markdown).toContain('<mark style="background-color:#fff3a3">highlight</mark>')
    expect(markdown).toContain('<span style="color:rgb(10, 20, 30);font-size:18px">styled</span>')

    const restored = parseMarkdownToPmDoc(markdown).doc
    const byText = new Map(allTextNodes(restored).map((n) => [n.text, n.marks]))
    expect(byText.get('under')).toContainEqual({ type: 'underline' })
    expect(byText.get('highlight')).toContainEqual({ type: 'highlight', attrs: { color: '#fff3a3' } })
    expect(byText.get('styled')).toContainEqual({
      type: 'textStyle', attrs: { color: 'rgb(10, 20, 30)', fontSize: '18px' },
    })
    expect(byText.get('sub')).toContainEqual({ type: 'subscript' })
    expect(byText.get('super')).toContainEqual({ type: 'superscript' })
  })

  it('default-denies unsafe style values instead of emitting injectable HTML', () => {
    const markdown = exportMarkdown(doc(paragraph(
      text('color', [{ type: 'textStyle', attrs: { color: 'red"><img src=x onerror=alert(1)>' } }]),
      text('size', [{ type: 'textStyle', attrs: { fontSize: '18px;position:fixed' } }]),
      text('mark', [{ type: 'highlight', attrs: { color: 'url(javascript:alert(1))' } }]),
    )))
    expect(markdown).toBe('colorsize<mark>mark</mark>\n')
    expect(markdown).not.toContain('<img')
    expect(markdown).not.toContain('position')
    expect(markdown).not.toContain('javascript')
  })

  it.each([
    ['bold', '**1.** '],
    ['italic', '*1.* '],
    ['strike', '~~1.~~ '],
  ])('moves trailing whitespace outside %s delimiters', (mark, expected) => {
    const markdown = exportMarkdown(doc(paragraph(text('1. ', [{ type: mark }]))))
    expect(markdown).toBe(`${expected}\n`)
    const restored = allTextNodes(parseMarkdownToPmDoc(markdown).doc)
    expect(restored.some((n) => n.text === '1.' && n.marks?.some((m) => m.type === mark))).toBe(true)
    expect(markdown).not.toContain(`${expected[0]}1. ${expected[0]}`)
  })

  it('round-trips a 30-formula fixture with manually bold numbering', () => {
    const source = doc(...Array.from({ length: 30 }, (_, i) => paragraph(
      text(`${i + 1}. `, [{ type: 'bold' }]),
      { type: 'inlineMath', attrs: { latex: `x_${i + 1}=${i + 1}^2` } },
    )))
    const markdown = exportMarkdown(source)
    expect(markdown.match(/^\*\*\d+\.\*\* /gm)).toHaveLength(30)
    const restored = parseMarkdownToPmDoc(markdown).doc
    const formulas = (restored.content ?? []).flatMap((p) => p.content ?? []).filter((n) => n.type === 'inlineMath')
    expect(formulas).toHaveLength(30)
    expect(formulas[29]?.attrs?.latex).toBe('x_30=30^2')
    const numbers = allTextNodes(restored).filter((n) => /^\d+\.$/.test(n.text ?? ''))
    expect(numbers).toHaveLength(30)
    expect(numbers.every((n) => n.marks?.some((m) => m.type === 'bold'))).toBe(true)
  })

  it('preserves a semantic ordered list and its non-default start', () => {
    const source = doc({
      type: 'orderedList', attrs: { start: 5 },
      content: Array.from({ length: 3 }, (_, i) => ({
        type: 'listItem', content: [paragraph({ type: 'inlineMath', attrs: { latex: `f_${i + 5}` } })],
      })),
    })
    const markdown = exportMarkdown(source)
    expect(markdown).toMatch(/^5\. \$f_5\$/)
    const restored = parseMarkdownToPmDoc(markdown).doc.content?.[0]
    expect(restored).toMatchObject({ type: 'orderedList', attrs: { start: 5 } })
    expect(restored?.content).toHaveLength(3)
  })

  it('preserves the group-02 nested task-list shape', () => {
    const source = acceptanceFixtures.group02NestedTask as PmNode
    const markdown = exportMarkdown(source)
    expect(markdown).toBe('- [x]  嵌套任务父\n    - [ ]  子任务\n')
    const restored = parseMarkdownToPmDoc(markdown).doc
    expect(restored.content?.[0]).toMatchObject({
      type: 'taskList', content: [{
        type: 'taskItem', attrs: { checked: true }, content: [
          { type: 'paragraph', content: [{ type: 'text', text: '嵌套任务父' }] },
          { type: 'taskList', content: [{
            type: 'taskItem', attrs: { checked: false }, content: [
              { type: 'paragraph', content: [{ type: 'text', text: '子任务' }] },
            ],
          }] },
        ],
      }],
    })
  })

  it('preserves intentional empty paragraphs around math/table sections', () => {
    const source = acceptanceFixtures.groups03_04_11Spacing as PmNode
    const markdown = exportMarkdown(source)
    expect(markdown.match(/data-octo-empty-paragraph/g)).toHaveLength(1)
    const restored = parseMarkdownToPmDoc(markdown).doc
    expect(restored.content?.map((node) => node.type)).toEqual([
      'heading', 'paragraph', 'blockMath', 'paragraph', 'paragraph',
    ])
    expect(restored.content?.[1]?.content?.[0]).toMatchObject({ text: '1.', marks: [{ type: 'bold' }] })
    // Markdown emphasis moves delimiter-adjacent trailing whitespace outside
    // the marked run. The rendered label still contains the same visible
    // separator; mark-boundary normalization is harmless.
    expect(restored.content?.filter((node) => node.type === 'paragraph' && !node.content)).toHaveLength(1)
  })

  it('keeps a top-level GFM table empty cell empty without leaking the paragraph sentinel', () => {
    const source = doc({
      type: 'table',
      content: [
        { type: 'tableRow', content: [
          { type: 'tableHeader', content: [paragraph(text('Header'))] },
          { type: 'tableHeader', content: [paragraph()] },
        ] },
        { type: 'tableRow', content: [
          { type: 'tableCell', content: [paragraph(text('Value'))] },
          { type: 'tableCell', content: [paragraph()] },
        ] },
      ],
    })

    const markdown = exportMarkdown(source)
    expect(markdown).not.toContain('data-octo-empty-paragraph')
    expect(markdown).toContain('| Header |  |')

    const restored = parseMarkdownToPmDoc(markdown).doc
    expect(allTextNodes(restored).some((node) => node.text?.includes('data-octo-empty-paragraph'))).toBe(false)
    expect(restored.content?.[0]?.content?.[1]?.content?.[1]).toEqual({ type: 'tableCell', content: [{ type: 'paragraph' }] })
  })

  it('imports legacy sentinel-bearing GFM cells as empty paragraphs, not literal marker text', () => {
    const legacy = [
      '| Header | Empty |',
      '| --- | --- |',
      '| Value | <p data-octo-empty-paragraph></p> |',
    ].join('\n')

    const restored = parseMarkdownToPmDoc(legacy).doc
    expect(allTextNodes(restored).some((node) => node.text?.includes('data-octo-empty-paragraph'))).toBe(false)
    expect(restored.content?.[0]?.content?.[1]?.content?.[1]).toEqual({ type: 'tableCell', content: [{ type: 'paragraph' }] })
  })

  it('keeps nested-table empty cells and intentional spacer paragraphs without marker text', () => {
    const nestedTable: PmNode = {
      type: 'table',
      content: [{ type: 'tableRow', content: [
        { type: 'tableCell', content: [paragraph()] },
      ] }],
    }
    const source = doc({
      type: 'table',
      content: [{ type: 'tableRow', content: [{
        type: 'tableCell',
        content: [paragraph(), nestedTable, paragraph()],
      }] }],
    })

    const markdown = exportMarkdown(source)
    expect(markdown).toContain('<table>')
    expect(markdown).toContain('<p></p><table>')
    expect(markdown).not.toContain('data-octo-empty-paragraph')

    const restored = parseMarkdownToPmDoc(markdown).doc
    expect(allTextNodes(restored).some((node) => node.text?.includes('data-octo-empty-paragraph'))).toBe(false)
    const outerCell = restored.content?.[0]?.content?.[0]?.content?.[0]
    expect(outerCell?.content?.map((node) => node.type)).toEqual(['paragraph', 'table', 'paragraph'])
    expect(outerCell?.content?.[0]).toEqual({ type: 'paragraph' })
    expect(outerCell?.content?.[2]).toEqual({ type: 'paragraph' })
    expect(outerCell?.content?.[1]?.content?.[0]?.content?.[0]?.content).toEqual([{ type: 'paragraph' }])
  })

  it('preserves visibly literal named entities instead of decoding them as markup', () => {
    const source = acceptanceFixtures.group07SpecialText as PmNode
    const markdown = exportMarkdown(source)
    expect(markdown).toContain('&amp;amp;')
    expect(allTextNodes(parseMarkdownToPmDoc(markdown).doc)[0]?.text).toBe(
      'XML 实体：< > & 与 &amp; &lt; &gt; 字面。',
    )
  })
})

describe('actual shared export → Markdown parser semantic inventory', () => {
  const canonicalMarks = (node: PmNode) => allTextNodes(node).map((n) => ({
    text: n.text,
    marks: [...(n.marks ?? [])].map((m) => ({ type: m.type, attrs: m.attrs ?? {} }))
      .sort((a, b) => a.type.localeCompare(b.type)),
  }))

  it('reimports fontFamily, fontSize, color, highlight, and every supported text mark', () => {
    const marks = [
      { type: 'bold' }, { type: 'italic' }, { type: 'underline' }, { type: 'strike' },
      { type: 'subscript' }, { type: 'superscript' },
      { type: 'link', attrs: { href: 'https://example.com/style' } },
      { type: 'highlight', attrs: { color: '#fff3a3' } },
      { type: 'textStyle', attrs: { color: '#123456', fontSize: '18px', fontFamily: 'Arial' } },
    ]
    const source = doc(paragraph(text('inventory', marks)))
    const restored = parseMarkdownToPmDoc(exportMarkdown(source)).doc
    expect(canonicalMarks(restored)).toEqual(canonicalMarks(source))
  })
})
