/**
 * Filesystem-backed byte store for the self-hosted local-hmac attachment blob
 * gateway (XIN-717).
 *
 * The `local-hmac` driver mints signed PUT/GET URLs pointed at the docs-backend
 * origin (see storage/objectStore.ts + config.attachments.publicBaseUrl). Unlike
 * the s3/minio drivers — where the browser uploads straight to real object
 * storage — the local-hmac path has no external backend, so the docs-backend
 * itself must persist and serve the bytes. This module is that persistence: it
 * writes each uploaded object under a configured directory keyed by its physical
 * object key, and reads it back with the content type recorded at upload time.
 *
 * It is deliberately minimal and intended for dev / single-node self-hosted
 * deployments (the same envelope local-hmac already targets), NOT for the
 * horizontally-scaled production API, which uses the s3/minio drivers.
 */
import { createHash, randomBytes } from 'node:crypto'
import { createWriteStream, promises as fs } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { Transform, type Readable } from 'node:stream'
import { pipeline } from 'node:stream/promises'
import { config } from '../config/env.js'

export interface StoredObject {
  bytes: Buffer
  contentType: string
}

/**
 * Error raised by `putStream` when the incoming body exceeds the byte cap. The
 * `code` lets the gateway map it to a 413 without buffering the rest of the body.
 */
export class PayloadTooLargeError extends Error {
  readonly code = 'payload_too_large'
  constructor() {
    super('payload_too_large')
    this.name = 'PayloadTooLargeError'
  }
}

/** Root directory for stored blobs; falls back to a temp dir when unset. */
function storageRoot(): string {
  const configured = config.attachments.localDir.trim()
  return configured !== '' ? resolve(configured) : join(tmpdir(), 'octo-docs-attachments')
}

/**
 * Map a physical object key to an on-disk path under the storage root. The key
 * (`[prefix/]docId/attachId/fileName`) is already sanitised at presign time, but
 * this store is defence-in-depth: the key is hashed to a flat, opaque file name
 * so no key content — however crafted — can traverse out of the root or collide
 * with another key. The content type is recorded in a `.ct` sidecar next to it.
 */
function pathsFor(objectKey: string): { root: string; blobPath: string; typePath: string } {
  const root = storageRoot()
  const digest = createHash('sha256').update(objectKey, 'utf8').digest('hex')
  // Shard by the first two hex chars to avoid one giant directory.
  const blobPath = join(root, digest.slice(0, 2), digest)
  return { root, blobPath, typePath: `${blobPath}.ct` }
}

export interface BlobStore {
  put(objectKey: string, contentType: string, bytes: Buffer): Promise<void>
  /**
   * Stream `source` straight to disk under `objectKey`, aborting the write once
   * more than `limit` bytes arrive (a `PayloadTooLargeError` is thrown and no
   * object is persisted). Memory stays bounded to a single chunk regardless of
   * the body size, so a flood of large uploads cannot exhaust the heap the way
   * a full Buffer.concat would. Resolves with the number of bytes written.
   */
  putStream(
    objectKey: string,
    contentType: string,
    source: Readable,
    limit: number,
  ): Promise<{ bytes: number }>
  get(objectKey: string): Promise<StoredObject | null>
  delete(objectKey: string): Promise<void>
}

class LocalFsBlobStore implements BlobStore {
  async put(objectKey: string, contentType: string, bytes: Buffer): Promise<void> {
    const { blobPath, typePath } = pathsFor(objectKey)
    await fs.mkdir(dirname(blobPath), { recursive: true })
    await fs.writeFile(blobPath, bytes)
    await fs.writeFile(typePath, contentType, 'utf8')
  }

  async putStream(
    objectKey: string,
    contentType: string,
    source: Readable,
    limit: number,
  ): Promise<{ bytes: number }> {
    const { blobPath, typePath } = pathsFor(objectKey)
    await fs.mkdir(dirname(blobPath), { recursive: true })
    // Write to a per-request temp file first, then atomically rename on success,
    // so a rejected (over-limit) or interrupted upload never leaves a partial
    // object at the real key.
    const tmpPath = `${blobPath}.${randomBytes(8).toString('hex')}.part`
    let bytes = 0
    const cap = new Transform({
      transform(chunk: Buffer, _enc, cb) {
        bytes += chunk.length
        if (bytes > limit) {
          // Abort before the chunk is written; nothing beyond `limit` is buffered
          // or flushed to disk.
          cb(new PayloadTooLargeError())
          return
        }
        cb(null, chunk)
      },
    })
    try {
      await pipeline(source, cap, createWriteStream(tmpPath))
    } catch (err) {
      await fs.rm(tmpPath, { force: true }).catch(() => {})
      throw err
    }
    await fs.rename(tmpPath, blobPath)
    await fs.writeFile(typePath, contentType, 'utf8')
    return { bytes }
  }

  async get(objectKey: string): Promise<StoredObject | null> {
    const { blobPath, typePath } = pathsFor(objectKey)
    try {
      const bytes = await fs.readFile(blobPath)
      let contentType = 'application/octet-stream'
      try {
        contentType = (await fs.readFile(typePath, 'utf8')).trim() || contentType
      } catch {
        /* missing sidecar: fall back to octet-stream */
      }
      return { bytes, contentType }
    } catch {
      return null
    }
  }

  async delete(objectKey: string): Promise<void> {
    const { blobPath, typePath } = pathsFor(objectKey)
    await Promise.all([fs.rm(blobPath, { force: true }), fs.rm(typePath, { force: true })])
  }
}

let store: BlobStore | null = null

/** Resolve the process-wide local blob store (lazily constructed). */
export function getLocalBlobStore(): BlobStore {
  if (!store) store = new LocalFsBlobStore()
  return store
}
