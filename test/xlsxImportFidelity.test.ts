import { describe, expect, it } from 'vitest'
import { parseXlsx } from '../src/import/xlsx/parse.js'
import { xlsxWorkbookToSheetBatch } from '../src/import/xlsx/toSheetBatch.js'
import { readFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import ExcelJS from 'exceljs'
import { exportXlsx } from '../src/export/xlsx.js'

const fixtures = {
  mixed: '/Users/cc/Downloads/test-7-13 (3).xlsx',
  styles: '/Users/cc/Downloads/未命名表格 (1) (1).xlsx',
  images: '/Users/cc/Downloads/testtest (3).xlsx',
}
async function batch(path: string) { return xlsxWorkbookToSheetBatch(await parseXlsx(await readFile(path))) }

describe('server XLSX import fidelity fixtures', () => {
  it.skipIf(!existsSync(fixtures.mixed))('retains values, resolved styles and the PNG drawing', async () => {
    const b = await batch(fixtures.mixed)
    expect(Object.values(b.cells).filter(c => c.v !== undefined)).toHaveLength(23)
    expect(Object.values(b.cells).filter(c => c.s !== undefined).length).toBeGreaterThanOrEqual(19)
    expect(Object.keys(b.drawings)).toHaveLength(1)
    expect(Object.values(b.drawings)[0]!.source).toMatch(/^data:image\/png;base64,/)
  })
  it.skipIf(!existsSync(fixtures.styles))('retains all styled cells, including style-only cells', async () => {
    const b = await batch(fixtures.styles)
    expect(Object.values(b.cells).filter(c => c.v !== undefined)).toHaveLength(8)
    expect(Object.values(b.cells).filter(c => c.s !== undefined)).toHaveLength(36)
    expect(b.cells['default!1:1']?.s).toMatchObject({ fs: 8, bg: { rgb: '#2C53F1' } })
    expect(b.cells['default!3:3']?.s).toMatchObject({ bl: 1, ul: { s: 1 }, st: { s: 1 } })
  })
  it.skipIf(!existsSync(fixtures.images))('accepts an image-only workbook and imports both drawings', async () => {
    const b = await batch(fixtures.images)
    expect(Object.values(b.cells).filter(c => c.v !== undefined)).toHaveLength(0)
    expect(Object.keys(b.drawings)).toHaveLength(2)
  })
  it('maps dimensions, hyperlinks and warns instead of fabricating merges', () => {
    const b = xlsxWorkbookToSheetBatch({ truncated:false, warnings:[], sheets:[{name:'S',dims:{c0:75,r0:20},drawings:[],merges:[{startRow:0,startCol:0,endRow:0,endCol:1}],cells:[{row:0,col:0,value:'Octo',hyperlink:'https://octo.example',style:{ff:'Arial'}}]}] })
    expect(b.dims).toEqual({c0:75,r0:20})
    expect(b.hyperlinks['default!xlsx_link_0_0']).toMatchObject({row:0,column:0,payload:'https://octo.example'})
    expect(b.warnings).toContain('merged ranges not supported (1 range(s) skipped)')
  })
  it('caps the number of imported drawings', async () => {
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('Images')
    const imageId = workbook.addImage({
      base64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADElEQVR42mNk+M/wHwAF/gL+Av7GAAAAAElFTkSuQmCC',
      extension: 'png',
    })
    for (let i = 0; i < 51; i++) sheet.addImage(imageId, { tl: { col: i % 10, row: Math.floor(i / 10) }, ext: { width: 1, height: 1 } })
    const parsed = await parseXlsx(Buffer.from(await workbook.xlsx.writeBuffer()))
    expect(parsed.sheets[0]!.drawings).toHaveLength(50)
    expect(parsed.warnings.some(w => w.includes('count limit'))).toBe(true)
  })

  it('imports, sanitizes, stores and exports actual SVG OOXML media', async () => {
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('SVG')
    const dirty = '<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script><rect width="20" height="10" fill="green" onload="x()"/></svg>'
    const imageId = workbook.addImage({ base64: Buffer.from(dirty).toString('base64'), extension: 'svg' as never })
    sheet.addImage(imageId, { tl: { col: 2, row: 3 }, ext: { width: 120, height: 60 } })
    const parsed = await parseXlsx(Buffer.from(await workbook.xlsx.writeBuffer()))
    expect(parsed.sheets[0]!.drawings).toHaveLength(1)
    const drawing = parsed.sheets[0]!.drawings[0]!
    expect(drawing.mime).toBe('image/svg+xml')
    expect(drawing.bytes.toString()).toContain('<rect')
    expect(drawing.bytes.toString()).not.toMatch(/script|onload/i)

    const imported = xlsxWorkbookToSheetBatch(parsed)
    expect(Object.values(imported.drawings)[0]!.source).toMatch(/^data:image\/svg\+xml;base64,/)
    const exported = await exportXlsx({}, 'SVG', { drawings: imported.drawings })
    const roundTrip = await parseXlsx(exported)
    expect(roundTrip.sheets[0]!.drawings[0]).toMatchObject({ mime: 'image/svg+xml', row: 3, col: 2 })
  })

  it('round-trips cell presentation, dimensions, hyperlink and exact drawing offsets', async () => {
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('Styled')
    const cell = sheet.getCell('B2')
    cell.value = { text: 'Octo', hyperlink: 'https://octo.example' }
    cell.font = { name: 'Arial', size: 14, bold: true, italic: true, underline: true, strike: true, color: { argb: 'FF123456' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFABCDEF' } }
    cell.border = { top: { style: 'mediumDashDot', color: { argb: 'FF654321' } }, bottom: { style: 'double' } }
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    cell.numFmt = '#,##0.00'
    sheet.getColumn(2).width = 21
    sheet.getRow(2).height = 27
    const imageId = workbook.addImage({
      base64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADElEQVR42mNk+M/wHwAF/gL+Av7GAAAAAElFTkSuQmCC',
      extension: 'png',
    })
    sheet.addImage(imageId, {
      tl: { nativeCol: 1, nativeRow: 1, nativeColOff: 19050, nativeRowOff: 28575 } as never,
      ext: { width: 41, height: 23 },
    })

    const parsed = await parseXlsx(Buffer.from(await workbook.xlsx.writeBuffer()))
    const imported = xlsxWorkbookToSheetBatch(parsed)
    const exported = await exportXlsx(imported.cells, 'Styled', imported)
    const reparsed = await parseXlsx(exported)
    const actual = reparsed.sheets[0]!
    expect(actual.cells.find(c => c.row === 1 && c.col === 1)).toMatchObject({
      value: 'Octo', hyperlink: 'https://octo.example',
      style: {
        ff: 'Arial', fs: 14, bl: 1, it: 1, ul: { s: 1 }, st: { s: 1 },
        cl: { rgb: '#123456' }, bg: { rgb: '#ABCDEF' }, ht: 2, vt: 2, tb: 1,
        n: { pattern: '#,##0.00' },
        bd: { t: { s: 10, cl: { rgb: '#654321' } }, b: { s: 7, cl: { rgb: '#000000' } } },
      },
    })
    expect(actual.dims).toMatchObject({ c1: 152, r1: 36 })
    expect(actual.drawings[0]).toMatchObject({
      row: 1, col: 1, rowOffset: 3, colOffset: 2, width: 41, height: 23,
    })
  })

  it('skips SVG that fails sanitation instead of putting it in the sheet batch', async () => {
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('Unsafe')
    const unsafe = '<!DOCTYPE svg [<!ENTITY x SYSTEM "file:///etc/passwd">]><svg xmlns="http://www.w3.org/2000/svg">&x;</svg>'
    const imageId = workbook.addImage({ base64: Buffer.from(unsafe).toString('base64'), extension: 'svg' as never })
    sheet.addImage(imageId, { tl: { col: 0, row: 0 }, ext: { width: 20, height: 20 } })
    await expect(parseXlsx(Buffer.from(await workbook.xlsx.writeBuffer()))).rejects.toMatchObject({ reason: 'empty' })
  })
})
