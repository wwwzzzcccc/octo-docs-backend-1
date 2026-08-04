import { describe, expect, it } from 'vitest'
import * as Y from 'yjs'
import {
  reconcileBoardRestoreOntoDoc,
  BOARD_RESTORE_META_FIELD,
  BOARD_RESTORE_EPOCH_KEY,
} from '../src/collab/liveRestore.js'
import { decodeBoardSnapshot } from '../src/collab/versionRestore.js'

const rect = (id: string, x: number) => ({
  id, type: 'rectangle', index: `a${id}`, x, y: 0, width: 5, height: 5,
  version: 1, versionNonce: 1,
})
function stateFromBoard(elements: Array<Record<string, unknown>>): Uint8Array {
  const doc = new Y.Doc()
  const map = doc.getMap<Y.Map<unknown>>('elements')
  for (const element of elements) {
    const entry = new Y.Map<unknown>()
    for (const [key, value] of Object.entries(element)) entry.set(key, value)
    map.set(String(element.id), entry)
  }
  return Y.encodeStateAsUpdate(doc)
}

describe('authoritative board version restore marker', () => {
  it('is monotonic, atomic with exact replacement, and survives persisted reconnect', () => {
    const live = new Y.Doc()
    Y.applyUpdate(live, stateFromBoard([rect('newer', 99), rect('later', 2)]))
    const target = stateFromBoard([rect('newer', 1)])

    let epoch = 0
    live.transact(() => { epoch = reconcileBoardRestoreOntoDoc(live, target) }, 'version-restore')
    expect(epoch).toBe(1)
    expect(live.getMap(BOARD_RESTORE_META_FIELD).get(BOARD_RESTORE_EPOCH_KEY)).toBe(1)
    expect(decodeBoardSnapshot(Y.encodeStateAsUpdate(live)).elements.map((e) => e.id)).toEqual(['newer'])

    live.transact(() => { epoch = reconcileBoardRestoreOntoDoc(live, target) }, 'version-restore')
    expect(epoch).toBe(2)

    const reopened = new Y.Doc()
    Y.applyUpdate(reopened, Y.encodeStateAsUpdate(live))
    expect(decodeBoardSnapshot(Y.encodeStateAsUpdate(reopened)).elements.map((e) => e.id)).toEqual(['newer'])
    expect(reopened.getMap(BOARD_RESTORE_META_FIELD).get(BOARD_RESTORE_EPOCH_KEY)).toBe(2)
  })

  it('fails before mutating the scene when the restore epoch is exhausted', () => {
    const live = new Y.Doc()
    Y.applyUpdate(live, stateFromBoard([rect('current', 99)]))
    live.getMap(BOARD_RESTORE_META_FIELD).set(BOARD_RESTORE_EPOCH_KEY, Number.MAX_SAFE_INTEGER)

    expect(() => {
      live.transact(() => reconcileBoardRestoreOntoDoc(live, stateFromBoard([rect('target', 1)])))
    }).toThrow('board_restore_epoch_exhausted')

    expect(decodeBoardSnapshot(Y.encodeStateAsUpdate(live)).elements.map((e) => e.id)).toEqual(['current'])
    expect(live.getMap(BOARD_RESTORE_META_FIELD).get(BOARD_RESTORE_EPOCH_KEY)).toBe(Number.MAX_SAFE_INTEGER)
  })
})
