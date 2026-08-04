/**
 * Live-document side of version restore (§4 feature #4, design §7.3).
 *
 * The DB-bound half (safety snapshot, role/epoch recheck, doc_meta) runs in the
 * restoreVersion service. This module performs the SECOND, indispensable half:
 * applying the reconcile onto the LIVE Hocuspocus document so connected clients
 * see the restore in real time — and so the restore is the single authoritative
 * content write, never clobbered by a stale client's union.
 *
 * Why this is the only correct content-write path (the Bug B fix):
 *   - openDirectConnection returns the SAME in-memory Y.Doc the connected tabs
 *     are editing. reconcileFragment inside connection.transact issues REAL Yjs
 *     deletes/inserts on that doc's struct store, so deletions become causal
 *     tombstones every client converges to (Hocuspocus broadcasts the update +
 *     publishes it over extension-redis to other nodes).
 *   - Because the live doc itself carries the deletions forward, its next store
 *     is on the union-safe direction (incoming ⊇ existing in persistence.store),
 *     so the diffUpdate bypass — not the union fallback — runs and the deleted
 *     content is NOT resurrected.
 *   - A second, separately computed yjs_document write (a transient-doc encode)
 *     would carry a DIFFERENT clientId and therefore force the union fallback on
 *     the live doc's next store, doubling the restored content. So restore must
 *     NOT write yjs_document directly; this path is authoritative.
 *
 * disconnect() (unloadImmediately default) flushes onStoreDocument with debounce
 * 0 and is awaited, so the restored state is durably persisted before we return.
 */
import type { Document } from '@hocuspocus/server'
import type * as Y from 'yjs'
import { getCollabServer } from './server.js'
import {
  decodeTargetSnapshot,
  reconcileFragment,
  reconcileSheetMap,
  reconcileBoardMaps,
} from './versionRestore.js'
import { advanceEditVersion } from './liveDocWrite.js'
import { COLLAB_FIELD } from '../schema/index.js'
import type { Node as PMNode } from 'prosemirror-model'

/** Top-level marker written atomically with an authoritative board restore. */
export const BOARD_RESTORE_META_FIELD = 'boardRestoreMeta'
export const BOARD_RESTORE_EPOCH_KEY = 'epoch'

/**
 * Apply one authoritative board restore to the live Y.Doc. The monotonic marker
 * and scene replacement share a transaction, so clients can choose replace
 * semantics instead of ordinary version reconciliation. Exported for tests.
 */
export function reconcileBoardRestoreOntoDoc(doc: Y.Doc, targetState: Uint8Array): number {
  const meta = doc.getMap<unknown>(BOARD_RESTORE_META_FIELD)
  const current = meta.get(BOARD_RESTORE_EPOCH_KEY)
  const previous = typeof current === 'number' && Number.isSafeInteger(current) && current >= 0
    ? current
    : 0
  // Yjs transactions are not rollback transactions when their callback throws. Validate the
  // marker before touching the scene so exhaustion cannot leave a replaced scene without its
  // corresponding authoritative epoch.
  if (previous >= Number.MAX_SAFE_INTEGER) throw new Error('board_restore_epoch_exhausted')
  const epoch = previous + 1
  reconcileBoardMaps(doc, targetState)
  meta.set(BOARD_RESTORE_EPOCH_KEY, epoch)
  return epoch
}

/**
 * Reconcile a restore's content onto a live doc, in place. MUST run inside the
 * doc's transaction so the deletions land as causal tombstones on the live struct
 * store. Shared between the live apply (applyRestoreToLiveDoc) and the unit tests
 * so the exact production sequence is covered without a live Hocuspocus server.
 *
 * The trailing advanceEditVersion is the P1 fix (mirrors commitLiveEdit /
 * commitLiveSheetEdit): a restore is a content WRITE, so its base-version token
 * must move forward even when the restore only DELETES. A delete-only reconcile
 * (rolling a sheet/doc back to a state with fewer cells or nodes) records only
 * tombstones, and Y.encodeStateVector tracks each client's insert clock, not
 * deletions — so without the bump the state vector was byte-identical after the
 * restore. That silently broke the two consumers of the token:
 *   - an in-flight sheet paginated read: its cursor embeds the pre-restore
 *     baseVersion; with the SV unchanged the drift guard failed to fire, so pages
 *     served before and after the restore stitched together an inconsistent grid
 *     instead of returning 409 sheet_changed.
 *   - the sheet/doc write guard: a bot could reuse its pre-restore baseVersion to
 *     slip a stale write past the If-Match optimistic-concurrency check.
 * Advancing a doc-private counter integrates one new struct on the acting doc's
 * clock, so the SV moves forward on a delete-only restore just like on a set.
 */
export function reconcileRestoreOntoDoc(doc: Y.Doc, targetPMDoc: PMNode, targetState: Uint8Array): void {
  const fragment = doc.getXmlFragment(COLLAB_FIELD)
  reconcileFragment(targetPMDoc, fragment)
  // Spreadsheet cells + dims + drawings (floating images) live in their own
  // Y.Maps, not the fragment — restore all three onto the same live doc so
  // connected clients converge on the restored grid. No-op for a text document
  // (it has no 'sheet' map). reconcileSheetMap reconciles all three surfaces.
  reconcileSheetMap(doc, targetState)
  // Advance the base-version token so a delete-only restore is still detected by
  // the read cursor / write guards (see the doc comment above).
  advanceEditVersion(doc)
}

/**
 * Apply a target version's content onto the live document and broadcast it.
 *
 * @param documentName canonical persistence/connection key for the document
 * @param uid          acting user id (becomes doc_meta.updated_by on store)
 * @param targetState  the target snapshot's raw Yjs bytes
 *
 * Throws SchemaIncompatibleError if the target cannot load under the current
 * schema (the caller validates this before mutating any state).
 */
export async function applyRestoreToLiveDoc(
  documentName: string,
  uid: string,
  targetState: Uint8Array,
): Promise<void> {
  const targetPMDoc = decodeTargetSnapshot(targetState)
  const server = getCollabServer()

  // openDirectConnection attaches to the document's single live in-memory copy
  // (loading it from DB if no client is connected). It bypasses onAuthenticate,
  // so authorization MUST already have been enforced by the caller (the restore
  // service rechecks admin + permission_epoch under the row lock).
  const connection = await server.hocuspocus.openDirectConnection(documentName, { user: { id: uid } })
  try {
    await connection.transact((doc: Document) => {
      reconcileRestoreOntoDoc(doc, targetPMDoc, targetState)
    })
  } finally {
    // Flushes the final store (awaited) and releases the direct connection.
    await connection.disconnect()
  }
}

/**
 * Board (Excalidraw scene) counterpart of applyRestoreToLiveDoc (delta #2).
 *
 * Applies a board version's ELEMENTS_FIELD / FILES_FIELD scene onto the live
 * whiteboard document and broadcasts it, exactly like the ProseMirror path: the
 * reconcile issues real Yjs deletes/inserts on the SAME in-memory Y.Doc the
 * connected boards are editing (openDirectConnection), so deletions become
 * causal tombstones every client converges to and the live doc's next store stays
 * on the union-safe direction — no transient-doc write that would diverge by
 * clientId and force the union fallback. The server-authoritative repair observer
 * already attached to the live board (server.ts afterLoadDocument) normalizes the
 * reconciled elements under REPAIR_ORIGIN, so a restored scene lands canonical.
 *
 * openDirectConnection bypasses onAuthenticate, so the caller (restoreVersion)
 * MUST already have rechecked admin + permission_epoch under the row lock.
 */
export async function applyBoardRestoreToLiveDoc(
  documentName: string,
  uid: string,
  targetState: Uint8Array,
): Promise<void> {
  const server = getCollabServer()
  const connection = await server.hocuspocus.openDirectConnection(documentName, { user: { id: uid } })
  try {
    await connection.transact((doc: Document) => {
      reconcileBoardRestoreOntoDoc(doc, targetState)
    })
  } finally {
    await connection.disconnect()
  }
}
