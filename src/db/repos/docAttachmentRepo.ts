/**
 * doc_attachment repository (§3.4 / §3.5).
 *
 * Records the doc -> object-storage reference created when the front-end
 * presigns an upload (§3.5 step 4). Lets reads re-issue signed URLs (step 5)
 * and provides the doc -> attachment edges for garbage collection / permission
 * inheritance. The binary itself lives in object storage, never here.
 *
 * Columns map snake_case -> camelCase in the typed return (see DocAttachment).
 */
import { query } from '../pool.js'

export interface DocAttachment {
  attachId: string
  docId: string
  objectKey: string
  mime: string
  sizeBytes: number
  fileName: string
  createdBy: string
  createdAt: Date
}

interface DocAttachmentRow {
  attach_id: string
  doc_id: string
  object_key: string
  mime: string
  size_bytes: number
  file_name: string
  created_by: string
  created_at: Date
}

function mapRow(row: DocAttachmentRow): DocAttachment {
  return {
    attachId: row.attach_id,
    docId: row.doc_id,
    objectKey: row.object_key,
    mime: row.mime,
    sizeBytes: Number(row.size_bytes),
    fileName: row.file_name ?? '',
    createdBy: row.created_by,
    createdAt: row.created_at,
  }
}

export interface RegisterAttachmentInput {
  attachId: string
  docId: string
  objectKey: string
  mime: string
  sizeBytes: number
  fileName: string
  createdBy: string
}

export const docAttachmentRepo = {
  /** Insert a new attachment reference (§3.5 step 4). */
  async register(input: RegisterAttachmentInput): Promise<void> {
    await query(
      `INSERT INTO doc_attachment
         (attach_id, doc_id, object_key, mime, size_bytes, file_name, created_by)
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
      [
        input.attachId,
        input.docId,
        input.objectKey,
        input.mime,
        input.sizeBytes,
        input.fileName,
        input.createdBy,
      ],
    )
  },

  async getById(attachId: string): Promise<DocAttachment | null> {
    const rows = await query<DocAttachmentRow>(
      'SELECT * FROM doc_attachment WHERE attach_id = ? LIMIT 1',
      [attachId],
    )
    return rows[0] ? mapRow(rows[0]) : null
  },

  async listByDoc(docId: string): Promise<DocAttachment[]> {
    const rows = await query<DocAttachmentRow>(
      'SELECT * FROM doc_attachment WHERE doc_id = ? ORDER BY created_at ASC',
      [docId],
    )
    return rows.map(mapRow)
  },

  async deleteById(attachId: string): Promise<void> {
    await query('DELETE FROM doc_attachment WHERE attach_id = ?', [attachId])
  },
}
