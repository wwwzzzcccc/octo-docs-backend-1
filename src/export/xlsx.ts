import ExcelJS from 'exceljs'
import type { SheetCell, StoredDrawing, StoredHyperLink, StoredSheetMeta } from '../agent/sheetConversion.js'
import { sanitizeSvg } from '../util/sanitizeSvg.js'

export interface XlsxExportExtras {
  dims?: Record<string, number>
  drawings?: Record<string, StoredDrawing>
  hyperlinks?: Record<string, StoredHyperLink>
  sheets?: Record<string, StoredSheetMeta>
}

type UniverColor = { rgb?: string }
type UniverBorder = { s?: number; cl?: UniverColor }

const BORDER_STYLE: Record<number, ExcelJS.BorderStyle> = {
  1: 'thin', 2: 'hair', 3: 'dotted', 4: 'dashed', 5: 'dashDot', 6: 'dashDotDot',
  7: 'double', 8: 'medium', 9: 'mediumDashed', 10: 'mediumDashDot',
  11: 'mediumDashDotDot', 12: 'slantDashDot', 13: 'thick',
}

function excelColor(value: unknown): Partial<ExcelJS.Color> | undefined {
  const rgb = (value as UniverColor | undefined)?.rgb
  if (typeof rgb !== 'string' || !/^#[0-9a-f]{6}$/i.test(rgb)) return undefined
  return { argb: `FF${rgb.slice(1).toUpperCase()}` }
}

function excelBorder(value: unknown): Partial<ExcelJS.Border> | undefined {
  const edge = value as UniverBorder | undefined
  if (!edge?.s) return undefined
  return { style: BORDER_STYLE[edge.s] ?? 'thin', color: excelColor(edge.cl) }
}

function applyStyle(target: ExcelJS.Cell, raw: unknown): void {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return
  const s = raw as Record<string, unknown>
  target.font = {
    ...target.font,
    name: typeof s.ff === 'string' ? s.ff : target.font?.name,
    size: typeof s.fs === 'number' ? s.fs : target.font?.size,
    bold: s.bl === 1 || s.bl === true,
    italic: s.it === 1 || s.it === true,
    underline: s.ul ? true : undefined,
    strike: s.st ? true : undefined,
    color: excelColor(s.cl),
  }
  const bg = excelColor(s.bg)
  if (bg) target.fill = { type: 'pattern', pattern: 'solid', fgColor: bg }
  const horizontal = s.ht === 2 ? 'center' : s.ht === 3 ? 'right' : s.ht === 4 ? 'justify' : s.ht === 1 ? 'left' : undefined
  const vertical = s.vt === 1 ? 'top' : s.vt === 3 ? 'bottom' : s.vt === 2 ? 'middle' : undefined
  if (horizontal || vertical || s.tb) target.alignment = { horizontal, vertical, wrapText: s.tb === 1 || s.tb === true }
  const numFmt = (s.n as { pattern?: unknown } | undefined)?.pattern
  if (typeof numFmt === 'string') target.numFmt = numFmt
  const bd = s.bd as Record<string, unknown> | undefined
  if (bd) target.border = {
    top: excelBorder(bd.t), bottom: excelBorder(bd.b),
    left: excelBorder(bd.l), right: excelBorder(bd.r),
  }
}

function parseDataUrl(source: unknown): { extension: 'png' | 'jpeg' | 'gif' | 'svg'; base64: string } | null {
  if (typeof source !== 'string') return null
  const match = /^data:image\/(png|jpeg|gif|svg\+xml);base64,([A-Za-z0-9+/=]+)$/.exec(source)
  if (!match) return null
  if (match[1] === 'svg+xml') {
    try {
      // Drawings are collaborative opaque objects and cannot be trusted merely
      // because they are already in Y.Doc. Re-sanitize before packaging XML.
      return { extension: 'svg', base64: sanitizeSvg(Buffer.from(match[2]!, 'base64')).toString('base64') }
    } catch {
      return null
    }
  }
  return { extension: match[1] as 'png' | 'jpeg' | 'gif', base64: match[2]! }
}

function sheetIdFromKey(key: string): string | null {
  const bang = key.indexOf('!')
  return bang > 0 ? key.slice(0, bang) : null
}

function dimensionForSheet(key: string, sheetId: string, onlySheet: boolean): string | null {
  const colon = key.indexOf(':')
  if (colon > 0) return key.slice(0, colon) === sheetId ? key.slice(colon + 1) : null
  // Legacy V1 dimensions had no sheet prefix. They belong only to the sole
  // sheet (or the canonical default sheet), never to every worksheet.
  return onlySheet || sheetId === 'default' ? key : null
}

export async function exportXlsx(
  cells: Record<string, SheetCell>,
  title: string,
  extras: XlsxExportExtras = {},
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Octo Docs'
  const bySheet = new Map<string, Array<{ row: number; col: number; cell: SheetCell }>>()
  for (const [sheetId] of Object.entries(extras.sheets ?? {}).sort(([, a], [, b]) => a.order - b.order)) {
    bySheet.set(sheetId, [])
  }
  for (const [key, cell] of Object.entries(cells)) {
    const match = /^([^!]+)!([0-9]+):([0-9]+)$/.exec(key)
    if (!match) continue
    const entries = bySheet.get(match[1]!) ?? []
    entries.push({ row: Number(match[2]), col: Number(match[3]), cell })
    bySheet.set(match[1]!, entries)
  }
  for (const key of Object.keys(extras.drawings ?? {})) {
    const sheetId = sheetIdFromKey(key)
    if (sheetId && !bySheet.has(sheetId)) bySheet.set(sheetId, [])
  }
  for (const key of Object.keys(extras.hyperlinks ?? {})) {
    const sheetId = sheetIdFromKey(key)
    if (sheetId && !bySheet.has(sheetId)) bySheet.set(sheetId, [])
  }
  if (!bySheet.size) bySheet.set('default', [])

  const used = new Set<string>()
  for (const [sheetId, entries] of bySheet) {
    let name = (extras.sheets?.[sheetId]?.name ?? (bySheet.size === 1 ? title : sheetId)).replace(/[\\/*?:[\]]/g, ' ').trim().slice(0, 31) || 'Sheet1'
    let suffix = 2
    const base = name
    while (used.has(name.toLowerCase())) name = `${base.slice(0, 28)} ${suffix++}`
    used.add(name.toLowerCase())
    const sheet = workbook.addWorksheet(name)

    for (const { row, col, cell } of entries) {
      const target = sheet.getCell(row + 1, col + 1)
      if (cell.f !== undefined) target.value = { formula: cell.f, result: cell.v ?? undefined }
      else target.value = cell.v as ExcelJS.CellValue
      applyStyle(target, cell.s)
    }
    for (const [key, px] of Object.entries(extras.dims ?? {})) {
      const localKey = dimensionForSheet(key, sheetId, bySheet.size === 1)
      if (!localKey) continue
      const row = /^r(\d+)$/.exec(localKey)
      const col = /^c(\d+)$/.exec(localKey)
      if (row) sheet.getRow(Number(row[1]) + 1).height = px * 72 / 96
      else if (col) sheet.getColumn(Number(col[1]) + 1).width = Math.max(1, (px - 5) / 7)
    }
    for (const [key, link] of Object.entries(extras.hyperlinks ?? {})) {
      if (sheetIdFromKey(key) !== sheetId || !Number.isInteger(link.row) || !Number.isInteger(link.column)) continue
      const target = sheet.getCell(link.row + 1, link.column + 1)
      const text = link.display ?? (typeof target.value === 'string' ? target.value : link.payload)
      target.value = { text, hyperlink: link.payload }
    }
    for (const [key, raw] of Object.entries(extras.drawings ?? {})) {
      if (sheetIdFromKey(key) !== sheetId) continue
      const drawing = raw as Record<string, unknown>
      const image = parseDataUrl(drawing.source)
      const from = (drawing.sheetTransform as { from?: Record<string, unknown> } | undefined)?.from
      const transform = drawing.transform as Record<string, unknown> | undefined
      if (!image || !from || !transform) continue
      const row = Number(from.row)
      const col = Number(from.column)
      const rowOffset = Number(from.rowOffset ?? 0)
      const columnOffset = Number(from.columnOffset ?? 0)
      const width = Number(transform.width)
      const height = Number(transform.height)
      if (![row, col, rowOffset, columnOffset, width, height].every(Number.isFinite)
        || row < 0 || col < 0 || rowOffset < 0 || columnOffset < 0 || width <= 0 || height <= 0) continue
      // ExcelJS writes SVG media/relationships correctly at runtime, although
      // its public ImageExtension type still omits "svg".
      const imageId = workbook.addImage({ base64: image.base64, extension: image.extension as never })
      // Passing only fractional `col`/`row` makes ExcelJS recalculate offsets
      // against its default cell dimensions. Preserve the OOXML EMU offsets
      // explicitly instead; imported offsets are stored in CSS pixels.
      sheet.addImage(imageId, {
        tl: {
          nativeCol: col,
          nativeRow: row,
          nativeColOff: Math.round(columnOffset * 9525),
          nativeRowOff: Math.round(rowOffset * 9525),
        } as never,
        ext: { width, height },
      })
    }
  }
  return Buffer.from(await workbook.xlsx.writeBuffer())
}
