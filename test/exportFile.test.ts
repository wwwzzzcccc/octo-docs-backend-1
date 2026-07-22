import { beforeEach, describe, expect, it, vi } from 'vitest'

vi.mock('../src/api/guard.js', () => ({ requireDocRole: vi.fn() }))
vi.mock('../src/collab/liveDocRead.js', () => ({ readLiveDocState: vi.fn(async () => new Uint8Array([1])) }))
vi.mock('../src/collab/liveSheetWrite.js', () => ({ readLiveSheet: vi.fn(async () => ({ state: new Uint8Array([2]), baseSV: new Uint8Array() })) }))
vi.mock('../src/agent/conversion.js', () => ({ yDocStateToProsemirrorJSON: vi.fn() }))
vi.mock('../src/db/repos/docAttachmentRepo.js', () => ({
  docAttachmentRepo: { listByDoc: vi.fn(async () => []) },
}))
vi.mock('../src/collab/versionRestore.js', async () => {
  const actual = await vi.importActual<typeof import('../src/collab/versionRestore.js')>('../src/collab/versionRestore.js')
  return {
    ...actual,
    decodeSheetSnapshot: vi.fn(),
    decodeSheetDimsSnapshot: vi.fn(() => ({})),
    decodeSheetDrawingsSnapshot: vi.fn(() => ({})),
    decodeSheetHyperLinksSnapshot: vi.fn(() => ({})),
    decodeSheetListSnapshot: vi.fn(() => ({})),
  }
})

import { exportFileHandler } from '../src/api/routes/export.js'
import { requireDocRole } from '../src/api/guard.js'
import { yDocStateToProsemirrorJSON } from '../src/agent/conversion.js'
import { decodeSheetSnapshot } from '../src/collab/versionRestore.js'
import { exportDocx } from '../src/export/docx.js'
import { exportXlsx } from '../src/export/xlsx.js'

function response() {
  return {
    statusCode: 0, body: undefined as unknown, headers: {} as Record<string, string>,
    status(code: number) { this.statusCode = code; return this },
    json(body: unknown) { this.body = body; return this },
    setHeader(key: string, value: string) { this.headers[key.toLowerCase()] = value; return this },
    end(body: unknown) { this.body = body; return this },
  }
}
const req = (format: string) => ({ params: { docId: 'd1' }, query: { format }, uid: 'u1', spaceId: 's1', headers: {} })

beforeEach(() => {
  vi.mocked(requireDocRole).mockResolvedValue({
    role: 'reader',
    meta: { doc_id: 'd1', document_name: 'octo:s1:f:d1', title: '测试 Doc', doc_type: 'doc' },
  } as never)
  vi.mocked(yDocStateToProsemirrorJSON).mockReturnValue({ type: 'doc', content: [{ type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Hello' }] }] })
})

describe('GET /:docId/export/file handler', () => {
  it('rejects an unknown format before permission/storage work', async () => {
    const res = response()
    await exportFileHandler(req('txt') as never, res as never)
    expect(res.statusCode).toBe(400)
    expect(res.body).toEqual({ error: 'invalid_format' })
    expect(requireDocRole).not.toHaveBeenCalled()
  })

  it('exports live rich text as markdown with download and nosniff headers', async () => {
    const res = response()
    await exportFileHandler(req('md') as never, res as never)
    expect(res.statusCode).toBe(200)
    expect(Buffer.isBuffer(res.body)).toBe(true)
    expect((res.body as Buffer).toString()).toBe('# Hello\n')
    expect(res.headers['content-type']).toBe('text/markdown; charset=utf-8')
    expect(res.headers['content-disposition']).toContain("filename*=UTF-8''")
    expect(res.headers['x-content-type-options']).toBe('nosniff')
  })

  it('enforces doc_type before decoding', async () => {
    vi.mocked(requireDocRole).mockResolvedValue({ role: 'reader', meta: { doc_type: 'sheet' } } as never)
    const res = response()
    await exportFileHandler(req('docx') as never, res as never)
    expect(res.statusCode).toBe(409)
    expect(res.body).toEqual({ error: 'unsupported_doc_type' })
  })

  it('exports live sheet values and formulas as a real XLSX zip', async () => {
    vi.mocked(requireDocRole).mockResolvedValue({ role: 'reader', meta: { doc_id: 'd1', document_name: 'n', title: 'Sheet', doc_type: 'sheet' } } as never)
    vi.mocked(decodeSheetSnapshot).mockReturnValue({ 'default!0:0': { v: 2 }, 'default!0:1': { v: 4, f: 'A1*2' } })
    const res = response()
    await exportFileHandler(req('xlsx') as never, res as never)
    const body = res.body as Buffer
    expect(body.subarray(0, 2).toString()).toBe('PK')
    expect(res.headers['content-type']).toContain('spreadsheetml.sheet')
    expect(res.headers['x-content-type-options']).toBe('nosniff')
  })
})

describe('Office exporters', () => {
  it('produces a genuine DOCX OOXML package', async () => {
    const bytes = await exportDocx({ type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'real docx', marks: [{ type: 'bold' }] }] }] })
    expect(bytes.subarray(0, 2).toString()).toBe('PK')
    expect(bytes.includes(Buffer.from('word/'))).toBe(true)
  })

  it('embeds SVG with a real PNG fallback instead of labelling XML as PNG', async () => {
    const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2"><rect width="2" height="2" fill="red"/></svg>')
    const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z4QAAAABJRU5ErkJggg==', 'base64')
    const bytes = await exportDocx(
      { type: 'doc', content: [{ type: 'image', attrs: { attachId: 'att_svg', width: 2, height: 2 } }] },
      new Map([['att_svg', { data: svg, type: 'svg', width: 2, height: 2, fallback: png }]]),
    )
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(bytes)
    const media = Object.keys(zip.files).filter((name) => name.startsWith('word/media/'))
    const svgName = media.find((name) => name.endsWith('.svg'))
    const pngName = media.find((name) => name.endsWith('.png'))
    expect(svgName).toBeTruthy()
    expect(pngName).toBeTruthy()
    expect(await zip.file(svgName!)!.async('string')).toContain('<svg')
    expect(Buffer.from(await zip.file(pngName!)!.async('uint8array')).subarray(0, 8)).toEqual(png.subarray(0, 8))
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
    expect(contentTypes).toContain('image/svg+xml')
    expect(contentTypes).toContain('image/png')
  })

  it('writes math nodes as real Office Math rather than plain text', async () => {
    const bytes = await exportDocx({
      type: 'doc',
      content: [{ type: 'blockMath', attrs: { latex: 'x^2 + y^2' } }],
    })
    const yauzl = await import('yauzl')
    const xml = await new Promise<string>((resolve, reject) => {
      yauzl.fromBuffer(bytes, { lazyEntries: true }, (error, zip) => {
        if (error || !zip) return reject(error ?? new Error('zip unavailable'))
        zip.readEntry()
        zip.on('entry', (entry) => {
          if (entry.fileName !== 'word/document.xml') return zip.readEntry()
          zip.openReadStream(entry, (streamError, stream) => {
            if (streamError || !stream) return reject(streamError ?? new Error('stream unavailable'))
            const chunks: Buffer[] = []
            stream.on('data', (chunk) => chunks.push(Buffer.from(chunk)))
            stream.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')))
            stream.on('error', reject)
          })
        })
        zip.on('error', reject)
      })
    })
    expect(xml).toContain('<m:oMath')
    expect(xml).toContain('<m:sSup>')
    expect(xml).toContain('<m:t xml:space="preserve">x</m:t>')
    expect(xml).toContain('<m:t xml:space="preserve">y</m:t>')
    expect(xml).not.toContain('x^2 + y^2')
  })

  it('round-trips an XLSX formula/result through exceljs', async () => {
    const bytes = await exportXlsx({ 'default!0:0': { v: 3 }, 'default!0:1': { v: 6, f: 'A1*2' } }, 'S')
    const ExcelJS = (await import('exceljs')).default
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(bytes as never)
    expect(workbook.worksheets[0]!.getCell('B1').value).toEqual({ formula: 'A1*2', result: 6 })
  })

  it('round-trips Univer styles, dimensions, hyperlinks and floating images', async () => {
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADElEQVR42mNk+M/wHwAF/gL+Av7GAAAAAElFTkSuQmCC'
    const bytes = await exportXlsx(
      { 'default!0:0': { v: 'Octo', s: { ff: 'Arial', fs: 14, bl: 1, bg: { rgb: '#2C53F1' } } } },
      'S',
      {
        dims: { c0: 75, r0: 20 },
        hyperlinks: { 'default!link_1': { id: 'link_1', row: 0, column: 0, payload: 'https://octo.example', display: 'Octo' } },
        drawings: {
          'default!img_1': {
            drawingId: 'img_1', source: `data:image/png;base64,${png}`,
            transform: { width: 10, height: 12 }, sheetTransform: { from: { row: 1, column: 2 } },
          },
        },
      },
    )
    const ExcelJS = (await import('exceljs')).default
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(bytes as never)
    const sheet = workbook.worksheets[0]!
    expect(sheet.getCell('A1').value).toEqual({ text: 'Octo', hyperlink: 'https://octo.example' })
    expect(sheet.getCell('A1').font).toMatchObject({ name: 'Arial', size: 14, bold: true })
    expect(sheet.getCell('A1').fill).toMatchObject({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C53F1' } })
    expect(sheet.getColumn(1).width).toBeCloseTo(10, 1)
    expect(sheet.getRow(1).height).toBeCloseTo(15, 1)
    expect(sheet.getImages()).toHaveLength(1)
  })

  it('applies canonical dimensions only to their matching worksheet', async () => {
    const bytes = await exportXlsx(
      { 'default!0:0': { v: 'A' }, 'sheet-2!0:0': { v: 'B' } },
      'book',
      { dims: { 'default:c0': 75, 'default:r0': 20, 'sheet-2:c0': 145, 'sheet-2:r0': 40, c1: 200 } },
    )
    const ExcelJS = (await import('exceljs')).default
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(bytes as never)
    const first = workbook.getWorksheet('default')!
    const second = workbook.getWorksheet('sheet-2')!
    expect(first.getColumn(1).width).toBeCloseTo(10, 1)
    expect(first.getRow(1).height).toBeCloseTo(15, 1)
    expect(second.getColumn(1).width).toBeCloseTo(20, 1)
    expect(second.getRow(1).height).toBeCloseTo(30, 1)
    expect(first.getColumn(2).width).toBeCloseTo((200 - 5) / 7, 1)
    expect(second.getColumn(2).width).toBeUndefined()
  })
})
