/**
 * Board (Excalidraw) scene edit core (bot REST surface). Pure, DB-free, no live
 * infrastructure — the whiteboard counterpart of agent/sheetConversion.ts and
 * collab/docBodyEdit.ts, so the element-batch validation + apply logic is
 * unit-testable without a running collab server.
 *
 * A board's scene is NOT a ProseMirror fragment or a flat cell map: it is the
 * two top-level whiteboard Y.Maps (`elements` keyed by id, `files` keyed by
 * fileId; see ydoc.ts / whiteboard-schema). A bot edits it with an element-level
 * upsert/delete batch that is applied straight onto the live doc's maps — the
 * SAME field-level, CAS-arbitrated write the front-end binding and the
 * server-authoritative repair perform, so a REST write and a WebSocket write
 * converge on one path (schema whitelist + `elementSupersedes` arbitration).
 */
import * as Y from 'yjs'
import {
  normalizeElement,
  elementSupersedes,
  normalizeFileRef,
  deterministicNonce,
  isReservedEntryKey,
  type WhiteboardElement,
  type FileRef,
} from './schema/index.js'
import { getElementsMap, getFilesMap, readEntry, writeEntryFields } from './ydoc.js'

/**
 * Raised when an upsert element fails the frozen schema whitelist / normalize
 * rules (missing/blank id, non-string or non-whitelisted `type`). Fail-closed
 * like SheetSnapshotInvalidError / BoardSnapshotInvalidError so a malformed
 * element can never reach the live doc. The service maps this to 422.
 */
export class BoardElementInvalidError extends Error {
  readonly code = 'board_element_invalid'
  constructor(message: string) {
    super(`board_element_invalid: ${message}`)
    this.name = 'BoardElementInvalidError'
  }
}

/**
 * Raised when a `files` entry carries no usable `attachId` (the XIN-699
 * grey-placeholder shape) — a reference no peer could ever resolve to a binary.
 * The service maps this to 422.
 */
export class BoardFileInvalidError extends Error {
  readonly code = 'board_file_invalid'
  constructor(message: string) {
    super(`board_file_invalid: ${message}`)
    this.name = 'BoardFileInvalidError'
  }
}

/** The raw, request-shaped scene edit batch (validated by validateBoardOps). */
export interface BoardOps {
  /** Full elements to upsert (CAS arbitrated: higher version wins). */
  elements?: unknown
  /** Element ids to tombstone (soft-delete, §1.1 — the key is retained). */
  deletedElementIds?: unknown
  /** File reference entries to upsert, keyed by fileId. */
  files?: unknown
  /** File ids to hard-delete. Used only by authoritative full-scene replacement. */
  deletedFileIds?: unknown
}

/** The validated, contract-checked batch ready to apply to a live board doc. */
export interface ValidatedBoardOps {
  upserts: WhiteboardElement[]
  deletes: string[]
  fileUpserts: Array<[string, FileRef]>
  fileDeletes: string[]
}

/** Coerce a stored `version` to the same clean int normalizeElement guarantees. */
function coerceVersion(v: unknown): number {
  return Number.isInteger(v) && (v as number) >= 1 ? (v as number) : 1
}

/**
 * Validate a raw scene edit batch fail-closed, WITHOUT mutating anything (the
 * board counterpart of validateSheetCellBatch). Every upsert element is run
 * through the shared `normalizeElement` rule set — an element the whitelist
 * rejects throws BoardElementInvalidError (→422) rather than being silently
 * dropped, so a caller learns its element was refused. Delete ids must be
 * non-empty strings; file entries must normalize to a usable FileRef.
 *
 * Does NOT reject an empty batch — the route shape gate rejects that as 400
 * invalid_body before the service is reached; an empty ValidatedBoardOps simply
 * applies as a no-op.
 */
export function validateBoardOps(ops: BoardOps): ValidatedBoardOps {
  const upserts: WhiteboardElement[] = []
  if (ops.elements !== undefined) {
    if (!Array.isArray(ops.elements)) {
      throw new BoardElementInvalidError('elements must be an array')
    }
    for (const raw of ops.elements) {
      // No NormalizeContext: this is per-element whitelist + numeric clamp +
      // index/nonce cleanup. Dangling-reference pruning across the whole set is
      // the server repair observer's job once the write lands.
      const el = normalizeElement(raw)
      if (!el) {
        throw new BoardElementInvalidError('element has a missing id or a non-whitelisted type')
      }
      upserts.push(el)
    }
  }

  const deletes: string[] = []
  if (ops.deletedElementIds !== undefined) {
    if (!Array.isArray(ops.deletedElementIds)) {
      throw new BoardElementInvalidError('deletedElementIds must be an array')
    }
    for (const id of ops.deletedElementIds) {
      if (typeof id !== 'string' || id.length === 0) {
        throw new BoardElementInvalidError('deletedElementIds entry must be a non-empty string')
      }
      deletes.push(id)
    }
  }

  const fileDeletes: string[] = []
  if (ops.deletedFileIds !== undefined) {
    if (!Array.isArray(ops.deletedFileIds)) throw new BoardFileInvalidError('deletedFileIds must be an array')
    for (const fid of ops.deletedFileIds) {
      if (typeof fid !== 'string' || fid.length === 0 || isReservedEntryKey(fid)) {
        throw new BoardFileInvalidError('deletedFileIds entry must be a non-empty safe string')
      }
      fileDeletes.push(fid)
    }
  }

  const fileUpserts: Array<[string, FileRef]> = []
  if (ops.files !== undefined) {
    if (!ops.files || typeof ops.files !== 'object' || Array.isArray(ops.files)) {
      throw new BoardFileInvalidError('files must be an object keyed by fileId')
    }
    for (const [fid, raw] of Object.entries(ops.files as Record<string, unknown>)) {
      if (fid.length === 0) throw new BoardFileInvalidError('empty fileId')
      // The fid is the KEY under which the ref is stored in the `files` map and
      // later rebuilt into a plain object (decodeBoardSnapshot `files[fid] = …`).
      // A reserved prototype name as the key corrupts that read-back exactly as a
      // reserved FIELD key does inside an element / ref (XIN-743) — reject it
      // fail-closed here so it never reaches the live Y.Map, symmetric with
      // normalizeElement / normalizeFileRef.
      if (isReservedEntryKey(fid)) {
        throw new BoardFileInvalidError(`file ${fid}: reserved fileId`)
      }
      const ref = normalizeFileRef(raw)
      if (!ref) throw new BoardFileInvalidError(`file ${fid}: no usable attachId`)
      fileUpserts.push([fid, ref])
    }
  }

  return { upserts, deletes, fileUpserts, fileDeletes }
}

/**
 * Apply a validated batch onto a board doc's live maps, in place. MUST be called
 * inside a Yjs transaction (the caller owns it so the whole batch is one update /
 * one broadcast). Unlike reconcileBoardMaps (a full make-live-equal-target
 * reconcile used by version restore), this is an element-LEVEL upsert/delete: it
 * only touches the ids named in the batch, so elements the batch does not mention
 * are left untouched.
 *
 *  - upsert: CAS via `elementSupersedes` (higher `version` wins, then smaller
 *    `versionNonce`). A non-superseding element is skipped (no write) so a
 *    concurrent higher-version edit is never clobbered — the same arbitration the
 *    FE binding and repair use. A superseding element is written field-level.
 *  - delete: soft-delete tombstone (§1.1 keeps the key) with a superseding
 *    `version` so the tombstone itself converges under CAS. A missing or
 *    already-tombstoned element is a no-op.
 *  - file upsert: field-level write of the canonical FileRef.
 */
export function applyBoardOpsToDoc(doc: Y.Doc, ops: ValidatedBoardOps): void {
  const elements = getElementsMap(doc)
  const files = getFilesMap(doc)

  for (const el of ops.upserts) {
    const existing = elements.get(el.id)
    const cur = existing instanceof Y.Map ? readEntry(existing) : undefined
    const currentCas = cur
      ? { version: coerceVersion(cur.version), versionNonce: coerceVersion(cur.versionNonce) }
      : undefined
    // CAS arbitration: only write when the incoming element supersedes the live
    // one. A fresh id (currentCas === undefined) always wins.
    if (!elementSupersedes(currentCas, el)) continue
    let yEl = existing
    if (!(yEl instanceof Y.Map)) {
      yEl = new Y.Map()
      elements.set(el.id, yEl as Y.Map<unknown>)
    }
    writeEntryFields(yEl, el as Record<string, unknown>)
  }

  for (const id of ops.deletes) {
    const existing = elements.get(id)
    if (!(existing instanceof Y.Map)) continue // nothing to tombstone
    const cur = readEntry(existing)
    if (cur.isDeleted === true) continue // already tombstoned
    const nextVersion = coerceVersion(cur.version) + 1
    existing.set('version', nextVersion)
    existing.set('versionNonce', deterministicNonce(`${id}:${nextVersion}`))
    existing.set('isDeleted', true)
  }

  for (const fid of ops.fileDeletes) files.delete(fid)

  for (const [fid, ref] of ops.fileUpserts) {
    let yF = files.get(fid)
    if (!(yF instanceof Y.Map)) {
      yF = new Y.Map()
      files.set(fid, yF as Y.Map<unknown>)
    }
    writeEntryFields(yF, ref as Record<string, unknown>)
  }
}

/**
 * Measure the board doc AFTER applying a batch, hydrated from the LIVE
 * `preEditState` (so it carries the live clientId clocks + tombstones) exactly as
 * persistence.store caps it at config.maxDocBytes — the board counterpart of
 * measureSheetAfterEdit / sizeAfterEdit. Lets the write path reject an oversized
 * result in its no-lock pre-flight, before commitLiveBoardEdit applies the batch
 * to the shared live Y.Doc, broadcasts it, and only then fails at store. Pure.
 */
export function measureBoardAfterEdit(
  preEditState: Uint8Array,
  ops: ValidatedBoardOps,
): { docBytes: number } {
  const scratch = new Y.Doc()
  Y.applyUpdate(scratch, preEditState)
  scratch.transact(() => applyBoardOpsToDoc(scratch, ops))
  return { docBytes: Y.encodeStateAsUpdate(scratch).length }
}
