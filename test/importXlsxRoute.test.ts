import { describe, it, expect, vi } from 'vitest'
import type { Request, Response } from 'express'
import ExcelJS from 'exceljs'

// Offline route test: mock the auth guard, the live-sheet read, and the sheet-write service so
// importXlsxHandler runs without live infra (mirrors importDocxRoute.test.ts). We assert the
// handler maps guard/body/parse/write outcomes to the right status codes and forwards warnings.
vi.mock('../src/api/guard.js', () => ({
  requireDocRole: vi.fn(),
}))
vi.mock('../src/collab/liveSheetWrite.js', () => ({
  readLiveSheet: vi.fn(async () => ({
    state: new Uint8Array(),
    baseSV: new Uint8Array([1, 2, 3]),
  })),
}))
vi.mock('../src/collab/versionRestore.js', () => ({
  decodeSheetSnapshot: vi.fn(() => ({})),
  decodeSheetDimsSnapshot: vi.fn(() => ({})),
  decodeSheetDrawingsSnapshot: vi.fn(() => ({})),
  decodeSheetHyperLinksSnapshot: vi.fn(() => ({})),
  decodeSheetMergesSnapshot: vi.fn(() => ({})),
  decodeSheetListSnapshot: vi.fn(() => ({})),
}))
vi.mock('../src/api/services/editDocSheet.js', () => ({
  editDocSheet: vi.fn(),
}))

import { importXlsxHandler } from '../src/api/routes/import.js'
import { requireDocRole } from '../src/api/guard.js'
import { editDocSheet } from '../src/api/services/editDocSheet.js'

const guard = vi.mocked(requireDocRole)
const edit = vi.mocked(editDocSheet)

interface MockRes {
  statusCode: number
  body: unknown
  status(c: number): MockRes
  json(b: unknown): MockRes
}
function mockRes(): MockRes {
  return {
    statusCode: 0,
    body: undefined,
    status(c: number) {
      this.statusCode = c
      return this
    },
    json(b: unknown) {
      this.body = b
      return this
    },
  }
}
function req(body: unknown): Request {
  return {
    uid: 'u1',
    spaceId: 's1',
    params: { docId: 'd1' },
    body,
  } as unknown as Request
}
function field(res: MockRes, key: string): unknown {
  return (res.body as Record<string, unknown> | undefined)?.[key]
}
async function run(res: MockRes, body: unknown): Promise<void> {
  await importXlsxHandler(req(body), res as unknown as Response)
}

/** A guard success for a sheet doc. */
function sheetGuard() {
  return {
    meta: {
      doc_id: 'd1',
      document_name: 'doc-d1',
      doc_type: 'sheet',
      permission_epoch: 1,
    },
  } as unknown as Awaited<ReturnType<typeof requireDocRole>>
}

/** Build a tiny .xlsx buffer with one cell. */
async function xlsxBuffer(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('S')
  ws.getCell('A1').value = 'hi'
  return Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer)
}

describe('importXlsxHandler — status mapping', () => {
  it('400 when the guard denies access', async () => {
    guard.mockResolvedValueOnce(undefined as unknown as Awaited<ReturnType<typeof requireDocRole>>)
    const res = mockRes()
    await run(res, await xlsxBuffer())
    // guard already wrote its own response; handler just returns. statusCode stays 0 here
    // because the mocked guard doesn't call res. Assert editDocSheet was never reached.
    expect(edit).not.toHaveBeenCalled()
  })

  it('409 unsupported_doc_type for a non-sheet target', async () => {
    guard.mockResolvedValueOnce({
      meta: {
        doc_id: 'd1',
        document_name: 'doc-d1',
        doc_type: 'doc',
        permission_epoch: 1,
      },
    } as unknown as Awaited<ReturnType<typeof requireDocRole>>)
    const res = mockRes()
    await run(res, await xlsxBuffer())
    expect(res.statusCode).toBe(409)
    expect(field(res, 'error')).toBe('unsupported_doc_type')
  })

  it('400 invalid_body when the body is not a Buffer', async () => {
    guard.mockResolvedValueOnce(sheetGuard())
    const res = mockRes()
    await run(res, {})
    expect(res.statusCode).toBe(400)
    expect(field(res, 'error')).toBe('invalid_body')
  })

  it('400 empty_upload on a zero-length buffer', async () => {
    guard.mockResolvedValueOnce(sheetGuard())
    const res = mockRes()
    await run(res, Buffer.alloc(0))
    expect(res.statusCode).toBe(400)
    expect(field(res, 'error')).toBe('empty_upload')
  })

  it('422 import_failed on unreadable bytes', async () => {
    guard.mockResolvedValueOnce(sheetGuard())
    const res = mockRes()
    await run(res, Buffer.from('not a workbook', 'utf8'))
    expect(res.statusCode).toBe(422)
    expect(field(res, 'error')).toBe('import_failed')
  })

  it('200 with baseVersion + warnings on a successful import', async () => {
    guard.mockResolvedValueOnce(sheetGuard())
    edit.mockResolvedValueOnce({
      ok: true,
      bytes: 10,
      baseVersion: 'bv2',
      newDocVersionSeq: 2,
    })
    const res = mockRes()
    await run(res, await xlsxBuffer())
    expect(res.statusCode).toBe(200)
    expect(field(res, 'baseVersion')).toBe('bv2')
    expect(Array.isArray(field(res, 'warnings'))).toBe(true)
    // the parsed cell was handed to editDocSheet under the read base version
    expect(edit).toHaveBeenCalledOnce()
    const arg = edit.mock.calls[0]![0]
    expect(arg.cells['default!0:0']).toEqual({ v: 'hi' })
  })

  it('forwards editDocSheet failure status/body', async () => {
    guard.mockResolvedValueOnce(sheetGuard())
    edit.mockResolvedValueOnce({
      ok: false,
      status: 412,
      error: 'version_conflict',
    })
    const res = mockRes()
    await run(res, await xlsxBuffer())
    expect(res.statusCode).toBe(412)
    expect(field(res, 'error')).toBe('version_conflict')
  })
})
