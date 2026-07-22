import { describe, it, expect } from 'vitest'
import * as Y from 'yjs'
import {
  validateBoardOps,
  applyBoardOpsToDoc,
  measureBoardAfterEdit,
  BoardElementInvalidError,
  BoardFileInvalidError,
} from '../src/whiteboard/boardEdit.js'
import { getElementsMap, getFilesMap, readEntry, readElements } from '../src/whiteboard/ydoc.js'

/** A minimal Excalidraw rectangle element. */
function rect(id: string, over: Record<string, unknown> = {}): Record<string, unknown> {
  return { id, type: 'rectangle', index: 'a0', x: 0, y: 0, width: 10, height: 10, version: 1, versionNonce: 100, ...over }
}

/** A live board Y.Doc holding the given elements (each a per-element Y.Map). */
function boardDoc(elements: Array<Record<string, unknown>> = []): Y.Doc {
  const doc = new Y.Doc()
  const elMap = getElementsMap(doc)
  doc.transact(() => {
    for (const el of elements) {
      const y = new Y.Map<unknown>()
      for (const [k, v] of Object.entries(el)) y.set(k, v)
      elMap.set(el.id as string, y)
    }
  })
  return doc
}

describe('validateBoardOps — fail-closed contract', () => {
  it('accepts and normalizes a whitelisted element', () => {
    const v = validateBoardOps({ elements: [rect('e1')] })
    expect(v.upserts).toHaveLength(1)
    expect(v.upserts[0]!.id).toBe('e1')
    expect(v.upserts[0]!.type).toBe('rectangle')
    // normalizeElement guarantees clean version/versionNonce.
    expect(v.upserts[0]!.version).toBe(1)
  })

  it('rejects an element with a non-whitelisted type (422 shape)', () => {
    expect(() => validateBoardOps({ elements: [{ id: 'x', type: 'not_a_real_type' }] })).toThrow(
      BoardElementInvalidError,
    )
  })

  it('rejects an element with a missing/blank id', () => {
    expect(() => validateBoardOps({ elements: [{ type: 'rectangle' }] })).toThrow(BoardElementInvalidError)
    expect(() => validateBoardOps({ elements: [{ id: '', type: 'rectangle' }] })).toThrow(BoardElementInvalidError)
  })

  it('rejects a non-array elements / deletedElementIds', () => {
    expect(() => validateBoardOps({ elements: 'nope' })).toThrow(BoardElementInvalidError)
    expect(() => validateBoardOps({ deletedElementIds: 'nope' })).toThrow(BoardElementInvalidError)
    expect(() => validateBoardOps({ deletedElementIds: [123] })).toThrow(BoardElementInvalidError)
  })

  it('rejects a file ref with no usable attachId (board_file_invalid)', () => {
    expect(() => validateBoardOps({ files: { f1: { mimeType: 'image/png' } } })).toThrow(BoardFileInvalidError)
  })

  it('accepts a usable file ref', () => {
    const v = validateBoardOps({ files: { f1: { attachId: 'a1', mimeType: 'image/png' } } })
    expect(v.fileUpserts).toEqual([['f1', { attachId: 'a1', mimeType: 'image/png' }]])
    expect(v.fileDeletes).toEqual([])
  })
})

describe('applyBoardOpsToDoc — element-level upsert/delete (CAS)', () => {
  it('inserts a fresh element (no current => incoming wins)', () => {
    const doc = boardDoc()
    const v = validateBoardOps({ elements: [rect('e1', { x: 5 })] })
    doc.transact(() => applyBoardOpsToDoc(doc, v))
    expect(readEntry(getElementsMap(doc).get('e1')).x).toBe(5)
  })

  it('overwrites only the named element, leaving others untouched', () => {
    const doc = boardDoc([rect('e1', { x: 1 }), rect('e2', { x: 2 })])
    const v = validateBoardOps({ elements: [rect('e1', { x: 99, version: 2 })] })
    doc.transact(() => applyBoardOpsToDoc(doc, v))
    expect(readEntry(getElementsMap(doc).get('e1')).x).toBe(99)
    // e2 is not in the batch — untouched (element-level, not a full reconcile).
    expect(readEntry(getElementsMap(doc).get('e2')).x).toBe(2)
  })

  it('CAS: a lower/equal-version upsert does NOT clobber a higher-version live element', () => {
    const doc = boardDoc([rect('e1', { x: 1, version: 5, versionNonce: 10 })])
    // Incoming version 3 < live 5 => skipped.
    const v = validateBoardOps({ elements: [rect('e1', { x: 77, version: 3, versionNonce: 1 })] })
    doc.transact(() => applyBoardOpsToDoc(doc, v))
    expect(readEntry(getElementsMap(doc).get('e1')).x).toBe(1)
    expect(readEntry(getElementsMap(doc).get('e1')).version).toBe(5)
  })

  it('delete tombstones the element (soft-delete, key retained) with a superseding version', () => {
    const doc = boardDoc([rect('e1', { version: 4 })])
    const v = validateBoardOps({ deletedElementIds: ['e1'] })
    doc.transact(() => applyBoardOpsToDoc(doc, v))
    const e = readEntry(getElementsMap(doc).get('e1'))
    expect(e.isDeleted).toBe(true)
    expect(e.version).toBe(5) // bumped so the tombstone converges under CAS
    // The key is retained (§1.1 never hard-deletes).
    expect(getElementsMap(doc).has('e1')).toBe(true)
  })

  it('delete of a missing element is a no-op', () => {
    const doc = boardDoc()
    const v = validateBoardOps({ deletedElementIds: ['ghost'] })
    doc.transact(() => applyBoardOpsToDoc(doc, v))
    expect(getElementsMap(doc).has('ghost')).toBe(false)
  })


  it('file delete removes a stale reference during full-scene replacement', () => {
    const doc = boardDoc()
    doc.transact(() => {
      const file = new Y.Map<unknown>()
      file.set('attachId', 'a1')
      getFilesMap(doc).set('stale', file)
    })
    const v = validateBoardOps({ deletedFileIds: ['stale'] })
    doc.transact(() => applyBoardOpsToDoc(doc, v))
    expect(getFilesMap(doc).has('stale')).toBe(false)
  })

  it('file upsert writes the canonical file ref', () => {
    const doc = boardDoc()
    const v = validateBoardOps({ files: { f1: { attachId: 'a1' } } })
    doc.transact(() => applyBoardOpsToDoc(doc, v))
    expect(readEntry(getFilesMap(doc).get('f1')).attachId).toBe('a1')
  })
})

describe('measureBoardAfterEdit', () => {
  it('measures a non-trivial post-edit doc size', () => {
    const pre = Y.encodeStateAsUpdate(boardDoc())
    const v = validateBoardOps({ elements: [rect('e1')] })
    const { docBytes } = measureBoardAfterEdit(pre, v)
    expect(docBytes).toBeGreaterThan(2) // more than an empty ~2-byte update
  })

  it('does not mutate the caller state (pure)', () => {
    const doc = boardDoc([rect('e1')])
    const pre = Y.encodeStateAsUpdate(doc)
    measureBoardAfterEdit(pre, validateBoardOps({ elements: [rect('e2')] }))
    // The original doc still has only e1.
    expect([...readElements(doc).keys()]).toEqual(['e1'])
  })
})
