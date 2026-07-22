import { describe, it, expect } from 'vitest'
import { xlsxWorkbookToSheetBatch, DEFAULT_SHEET_ID } from '../src/import/xlsx/toSheetBatch.js'
import type { ParsedXlsxWorkbook } from '../src/import/xlsx/parse.js'

function wb(
  partial: Partial<ParsedXlsxWorkbook> & {
    sheets: ParsedXlsxWorkbook['sheets']
  },
): ParsedXlsxWorkbook {
  return { truncated: false, warnings: [], ...partial }
}

describe('xlsxWorkbookToSheetBatch', () => {
  it('maps values and formulas onto default-sheet cell keys', () => {
    const { cells, warnings } = xlsxWorkbookToSheetBatch(
      wb({
        sheets: [
          {
            name: 'Data',
            merges: [],
            dims: {}, drawings: [],
            cells: [
              { row: 0, col: 0, value: 'hi' },
              { row: 0, col: 1, value: 42 },
              { row: 2, col: 0, formula: 'A1' },
            ],
          },
        ],
      }),
    )
    expect(warnings).toEqual([])
    expect(cells[`${DEFAULT_SHEET_ID}!0:0`]).toEqual({ v: 'hi' })
    expect(cells[`${DEFAULT_SHEET_ID}!0:1`]).toEqual({ v: 42 })
    expect(cells[`${DEFAULT_SHEET_ID}!2:0`]).toEqual({ f: 'A1' })
  })

  it('imports only the first sheet and warns about the rest', () => {
    const { cells, warnings } = xlsxWorkbookToSheetBatch(
      wb({
        sheets: [
          { name: 'One', merges: [], dims: {}, drawings: [], cells: [{ row: 0, col: 0, value: 'a' }] },
          { name: 'Two', merges: [], dims: {}, drawings: [], cells: [{ row: 0, col: 0, value: 'b' }] },
          { name: 'Three', merges: [], dims: {}, drawings: [], cells: [] },
        ],
      }),
    )
    expect(cells[`${DEFAULT_SHEET_ID}!0:0`]).toEqual({ v: 'a' })
    expect(warnings).toHaveLength(1)
    expect(warnings[0]).toContain('One')
    expect(warnings[0]).toContain('Two')
    expect(warnings[0]).toContain('Three')
  })

  it('warns when the workbook was truncated', () => {
    const { warnings } = xlsxWorkbookToSheetBatch(
      wb({
        truncated: true,
        sheets: [{ name: 'S', merges: [], dims: {}, drawings: [], cells: [{ row: 0, col: 0, value: 'x' }] }],
      }),
    )
    expect(warnings.some((w) => w.includes('truncated'))).toBe(true)
  })

  it('skips cells with neither value nor formula', () => {
    const { cells } = xlsxWorkbookToSheetBatch(
      wb({
        sheets: [
          {
            name: 'S',
            merges: [],
            dims: {}, drawings: [],
            cells: [
              { row: 0, col: 0, value: null },
              { row: 1, col: 0, value: 'kept' },
            ],
          },
        ],
      }),
    )
    expect(cells[`${DEFAULT_SHEET_ID}!0:0`]).toBeUndefined()
    expect(cells[`${DEFAULT_SHEET_ID}!1:0`]).toEqual({ v: 'kept' })
  })

  it('returns an empty batch for a workbook with no sheets', () => {
    const { cells, warnings } = xlsxWorkbookToSheetBatch(wb({ sheets: [] }))
    expect(cells).toEqual({})
    expect(warnings).toEqual([])
  })
})
