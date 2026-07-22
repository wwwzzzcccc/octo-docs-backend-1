import type { SheetCell, StoredDrawing, StoredHyperLink } from '../../agent/sheetConversion.js'
import { sheetCellKey } from '../../agent/sheetConversion.js'
import type { ParsedXlsxWorkbook } from './parse.js'

export const DEFAULT_SHEET_ID='default'
export interface XlsxToSheetBatchResult {
  cells:Record<string,SheetCell>
  dims:Record<string,number>
  drawings:Record<string,StoredDrawing>
  hyperlinks:Record<string,StoredHyperLink>
  warnings:string[]
}
export function xlsxWorkbookToSheetBatch(workbook:ParsedXlsxWorkbook,sheetId=DEFAULT_SHEET_ID):XlsxToSheetBatchResult{
  const cells:Record<string,SheetCell>={},dims:Record<string,number>={},drawings:Record<string,StoredDrawing>={},hyperlinks:Record<string,StoredHyperLink>={}
  const warnings=[...(workbook.warnings??[])];const first=workbook.sheets[0]
  if(!first)return {cells,dims,drawings,hyperlinks,warnings}
  if(workbook.sheets.length>1)warnings.push(`only the first worksheet ("${first.name}") was imported; ${workbook.sheets.length-1} additional sheet(s) skipped: ${workbook.sheets.slice(1).map(s=>s.name).join(', ')}`)
  if(workbook.truncated)warnings.push('some content exceeded the import size limit and was truncated')
  if(first.merges.length)warnings.push(`merged ranges not supported (${first.merges.length} range(s) skipped)`)
  for(const c of first.cells){const cell:SheetCell={};if(c.value!==undefined&&c.value!==null)cell.v=c.value;if(c.formula!==undefined)cell.f=c.formula;if(c.style)cell.s=c.style;if(Object.keys(cell).length)cells[sheetCellKey(sheetId,c.row,c.col)]=cell
    if(c.hyperlink&&(/^(https?:|mailto:)/i.test(c.hyperlink)||c.hyperlink.startsWith('#'))){const id=`xlsx_link_${c.row}_${c.col}`;hyperlinks[`${sheetId}!${id}`]={id,row:c.row,column:c.col,payload:c.hyperlink,display:typeof c.value==='string'?c.value:undefined}}
  }
  Object.assign(dims,first.dims)
  for(const d of first.drawings){drawings[`${sheetId}!${d.id}`]={drawingId:d.id,drawingType:0,imageSourceType:'BASE64',source:`data:${d.mime};base64,${d.bytes.toString('base64')}`,transform:{left:0,top:0,width:d.width,height:d.height,angle:0,flipX:false,flipY:false,skewX:0,skewY:0},sheetTransform:{from:{row:d.row,column:d.col,rowOffset:d.rowOffset,columnOffset:d.colOffset},to:{row:d.row,column:d.col,rowOffset:d.rowOffset+d.height,columnOffset:d.colOffset+d.width}}}}
  return {cells,dims,drawings,hyperlinks,warnings}
}
