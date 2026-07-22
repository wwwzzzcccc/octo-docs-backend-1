import type ExcelJS from 'exceljs'
import { sanitizeSvg } from '../../util/sanitizeSvg.js'

export const MAX_IMPORT_ROWS = 1000
export const MAX_IMPORT_COLS = 100
export const MAX_IMPORT_IMAGES = 50
export const MAX_IMPORT_IMAGE_BYTES = 5 * 1024 * 1024
export const MAX_IMPORT_IMAGE_TOTAL_BYTES = 20 * 1024 * 1024

export interface ParsedXlsxCell {
  row: number
  col: number
  value?: string | number | boolean | null
  formula?: string
  style?: Record<string, unknown>
  hyperlink?: string
}
export interface ParsedXlsxMerge { startRow: number; startCol: number; endRow: number; endCol: number }
export interface ParsedXlsxDrawing {
  id: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif' | 'image/svg+xml'
  bytes: Buffer
  row: number
  col: number
  rowOffset: number
  colOffset: number
  width: number
  height: number
}
export interface ParsedXlsxSheet {
  name: string
  cells: ParsedXlsxCell[]
  merges: ParsedXlsxMerge[]
  dims: Record<string, number>
  drawings: ParsedXlsxDrawing[]
}
export interface ParsedXlsxWorkbook { sheets: ParsedXlsxSheet[]; truncated: boolean; warnings: string[] }
export class XlsxParseError extends Error {
  constructor(public readonly reason: 'empty' | 'unreadable') { super(`xlsx parse failed: ${reason}`); this.name = 'XlsxParseError' }
}

const BORDER_STYLE: Record<string, number> = { thin:1,hair:2,dotted:3,dashed:4,dashDot:5,dashDotDot:6,double:7,medium:8,mediumDashed:9,mediumDashDot:10,mediumDashDotDot:11,slantDashDot:12,thick:13 }
function color(argb?: string): string | undefined { return argb && /^[0-9a-f]{8}$/i.test(argb) ? `#${argb.slice(2)}` : argb && /^[0-9a-f]{6}$/i.test(argb) ? `#${argb}` : undefined }
function edge(e: { style?: string; color?: { argb?: string } } | undefined): unknown {
  if (!e?.style) return undefined
  return { s: BORDER_STYLE[e.style] ?? 1, cl: { rgb: color(e.color?.argb) ?? '#000000' } }
}
function resolvedStyle(cell: ExcelJS.Cell): Record<string, unknown> | undefined {
  const s: Record<string, unknown> = {}
  const f = cell.font
  if (f) {
    if (f.bold) s.bl=1; if (f.italic) s.it=1; if (f.underline) s.ul={s:1}; if (f.strike) s.st={s:1}
    if (typeof f.size === 'number') s.fs=f.size; if (f.name) s.ff=f.name
    const c=color(f.color?.argb); if(c) s.cl={rgb:c}
  }
  const fill=cell.fill
  if (fill?.type==='pattern' && fill.pattern==='solid') { const c=color(fill.fgColor?.argb); if(c) s.bg={rgb:c} }
  const a=cell.alignment
  if (a?.horizontal) s.ht=a.horizontal==='center'?2:a.horizontal==='right'?3:a.horizontal==='justify'?4:1
  if (a?.vertical) s.vt=a.vertical==='top'?1:a.vertical==='bottom'?3:2
  if (a?.wrapText) s.tb=1
  if (cell.numFmt && cell.numFmt !== 'General') s.n={pattern:cell.numFmt}
  const b=cell.border
  if (b) { const bd:Record<string,unknown>={}; for(const [to,from] of [['t','top'],['b','bottom'],['l','left'],['r','right']] as const){const v=edge(b[from]);if(v)bd[to]=v} if(Object.keys(bd).length)s.bd=bd }
  return Object.keys(s).length ? s : undefined
}
function isDate(v: unknown): v is Date { return Object.prototype.toString.call(v)==='[object Date]' }

/**
 * Excel formulas can initiate network/DDE/RTD activity when a collaborator
 * opens a later export. Preserve normal workbook formulas, but fail closed for
 * external-reference primitives and link formulas at the trust boundary.
 */
export function isSafeImportedFormula(formula: string): boolean {
  const normalized = formula.replace(/^=/, '').replace(/\s+/g, '').toUpperCase()
  if (!normalized || normalized.length > 8192) return false
  if (/\b(?:WEBSERVICE|FILTERXML|RTD|HYPERLINK|IMPORTDATA|IMPORTXML|IMPORTHTML|IMPORTRANGE)\s*\(/.test(normalized)) return false
  // Excel external workbook / DDE spellings: [book.xlsx]Sheet!A1 and app|topic!item.
  if (/\[[^\]]+\][^!]*!/.test(normalized) || /(?:^|[=,+\-*/(])[^'"]+\|[^!]+!/.test(normalized)) return false
  return true
}
function readCell(cell: ExcelJS.Cell,row:number,col:number):ParsedXlsxCell|null {
  if(cell.isMerged && cell.master?.address!==cell.address) return null
  const out:ParsedXlsxCell={row,col}; const v=cell.value
  if(typeof v==='string'||typeof v==='number'||typeof v==='boolean') out.value=v
  else if(isDate(v)) out.value=(v.getTime()-Date.UTC(1899,11,30))/86400000
  else if(v && typeof v==='object') {
    const o=v as {formula?:string;result?:unknown;text?:string;hyperlink?:string}
    if(typeof o.formula==='string' && !/DISPIMG/i.test(o.formula)) {
      const r=o.result
      if(isSafeImportedFormula(o.formula)) out.formula=o.formula.replace(/^=/,'')
      if(isDate(r)) out.value=(r.getTime()-Date.UTC(1899,11,30))/86400000
      else if(r!=null&&typeof r!=='object') out.value=r as string|number|boolean
    }
    else if(typeof o.text==='string') out.value=o.text
    if(typeof o.hyperlink==='string') out.hyperlink=o.hyperlink
  }
  if(typeof cell.hyperlink==='string') out.hyperlink=cell.hyperlink
  out.style=resolvedStyle(cell)
  return out.value===undefined&&out.formula===undefined&&out.style===undefined&&out.hyperlink===undefined?null:out
}
function mimeFor(ext?:string):ParsedXlsxDrawing['mime']|null { const e=ext?.toLowerCase();return e==='png'?'image/png':e==='jpg'||e==='jpeg'?'image/jpeg':e==='gif'?'image/gif':e==='svg'?'image/svg+xml':null }
function pxOffset(native:number|undefined):number { return Math.max(0,(native??0)/9525) }

export async function parseXlsx(data:Buffer|ArrayBuffer):Promise<ParsedXlsxWorkbook>{
  const Excel=(await import('exceljs')).default; const wb=new Excel.Workbook()
  try{await wb.xlsx.load((Buffer.isBuffer(data)?data:Buffer.from(data)) as unknown as Parameters<typeof wb.xlsx.load>[0])}catch{throw new XlsxParseError('unreadable')}
  const sheets:ParsedXlsxSheet[]=[]; const warnings:string[]=[]; let truncated=false,totalImageBytes=0,imageCount=0
  for(const ws of wb.worksheets){
    if(ws.state!=='visible')continue
    const rows=Math.min(ws.rowCount,MAX_IMPORT_ROWS),cols=Math.min(ws.columnCount,MAX_IMPORT_COLS)
    if(ws.rowCount>MAX_IMPORT_ROWS||ws.columnCount>MAX_IMPORT_COLS)truncated=true
    const cells:ParsedXlsxCell[]=[]
    for(let r=1;r<=rows;r++)for(let c=1;c<=cols;c++){const x=readCell(ws.getCell(r,c),r-1,c-1);if(x)cells.push(x)}
    const dims:Record<string,number>={}
    for(let r=1;r<=rows;r++){const h=ws.getRow(r).height;if(typeof h==='number'&&h>0)dims[`r${r-1}`]=h*96/72}
    for(let c=1;c<=cols;c++){const w=ws.getColumn(c).width;if(typeof w==='number'&&w>0)dims[`c${c-1}`]=Math.max(1,Math.floor(w*7+5))}
    const merges:ParsedXlsxMerge[]=[]
    for(const range of ws.model.merges??[]){const m=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(range);if(!m)continue;const ci=(x:string)=>[...x.toUpperCase()].reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0)-1;merges.push({startCol:ci(m[1]!),startRow:+m[2]!-1,endCol:ci(m[3]!),endRow:+m[4]!-1})}
    const drawings:ParsedXlsxDrawing[]=[]
    for(const img of ws.getImages()){
      if(imageCount>=MAX_IMPORT_IMAGES){warnings.push('some images exceeded the import count limit and were skipped');break}
      const media=wb.getImage(Number(img.imageId));const mime=mimeFor(media?.extension);const rawBytes:Buffer|null=media?.buffer?Buffer.from(media.buffer):media?.base64?Buffer.from(media.base64.replace(/^data:[^,]+,/,''),'base64'):null
      if(!mime||!rawBytes||rawBytes.length===0||rawBytes.length>MAX_IMPORT_IMAGE_BYTES||totalImageBytes+rawBytes.length>MAX_IMPORT_IMAGE_TOTAL_BYTES){warnings.push('an unsupported or oversized image was skipped');continue}
      let bytes:Buffer=rawBytes
      if(mime==='image/svg+xml') {
        try { bytes=sanitizeSvg(bytes) }
        catch { warnings.push('an invalid or unsafe SVG image was skipped');continue }
      }
      const range=img.range as unknown as {tl:{nativeRow?:number;nativeCol?:number;nativeRowOff?:number;nativeColOff?:number};ext?:{width?:number;height?:number}}
      const row=range.tl.nativeRow??0,col=range.tl.nativeCol??0
      if(row>=MAX_IMPORT_ROWS||col>=MAX_IMPORT_COLS){warnings.push('an image outside the import grid was skipped');continue}
      imageCount++;totalImageBytes+=bytes.length
      drawings.push({id:`xlsx_img_${imageCount}`,mime,bytes,row,col,rowOffset:pxOffset(range.tl.nativeRowOff),colOffset:pxOffset(range.tl.nativeColOff),width:Math.max(1,range.ext?.width??96),height:Math.max(1,range.ext?.height??96)})
    }
    if(cells.length||drawings.length||Object.keys(dims).length||merges.length)sheets.push({name:ws.name,cells,merges,dims,drawings})
  }
  if(!sheets.length)throw new XlsxParseError('empty')
  return {sheets,truncated,warnings}
}
