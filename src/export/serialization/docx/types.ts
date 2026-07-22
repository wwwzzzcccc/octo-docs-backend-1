// @ts-nocheck -- verbatim migration of the proven frontend exporter; covered by focused fidelity tests
/**
 * Type definitions for the DOCX export module.
 * Re-exports the MdNode interface from markdown.ts for consistency.
 */

import type { PmNode as MdNode } from '../policy.js'
import type { ResolvedAttachment } from '../markdown.js'

export type { MdNode }

/** Options for the DOCX export function. */
export interface DocxExportOptions {
  /** Batch size for the resolve endpoint (RES-1 cap). Default 200. */
  batchSize?: number
  /** Resolve fn injection point (tests pass a stub). Defaults to the real REST client. */
  imageAdapter?: DocxImageAdapter
  /** name → unicode glyph for emoji nodes; defaults to the editor's emoji map. */
  emojiGlyph?: (name: string | null | undefined) => string | undefined
}

/** Internal context passed through the node/mark converters. */
export interface DocxImageAdapter {
  resolve(docId: string, attachIds: string[]): Promise<{ items: ResolvedAttachment[]; notFound: string[] }>
  fetch(url: string, maxBytes: number): Promise<DocxFetchedImage | ArrayBuffer | undefined>
  sanitizeDirectUrl?(url: string): string | null
}

export interface DocxFetchedImage {
  data: ArrayBuffer
  type?: 'png' | 'jpg' | 'gif' | 'bmp' | 'svg'
  width?: number
  height?: number
  fallback?: ArrayBuffer
}

export interface DocxContext {
  /** Resolved attachment URLs keyed by attachId. */
  urls: Map<string, ResolvedAttachment>
  /** Fetched image buffers keyed by URL. */
  imageBuffers: Map<string, DocxFetchedImage>
  /** Emoji glyph resolver. */
  emojiGlyph?: (name: string | null | undefined) => string | undefined
  /**
   * Dynamic numbering references for ordered lists whose start > 1.
   * docx derives an instance's starting number from its abstract level[0].start,
   * so a non-default start needs its own reference with a customized level set.
   */
  dynamicNumbering: Array<{ reference: string; start: number; level: number }>
  /** Monotonic counter handing each ordered list its own numbering instance (independent counting). */
  orderedListInstance: number
  /**
   * Optional upper bound (px) on rendered image width for the current context.
   * Set when converting content inside a table cell so images shrink to fit the
   * cell (which can be much narrower than the page when tables are nested),
   * instead of overflowing at the page-wide MAX_IMAGE_WIDTH cap.
   */
  maxImageWidthPx?: number
}

/** Represents a collected image reference that needs to be fetched. */
export interface ImageRef {
  attachId?: string
  src?: string
  resolvedUrl?: string
}
