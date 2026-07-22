import { serializeDocx } from './serialization/docx/index.js'
import type { DocxImageAdapter } from './serialization/docx/types.js'
import type { PmNode } from './serialization/policy.js'

export interface DocxImage { data: Buffer; type: 'png'|'jpg'|'gif'|'svg'; width: number; height: number; fallback?: Buffer }

export async function exportDocx(doc: PmNode, images: ReadonlyMap<string, DocxImage> = new Map()): Promise<Buffer> {
  const adapter: DocxImageAdapter = {
    async resolve(_docId, ids) {
      return { items: ids.filter((id) => images.has(id)).map((attachId) => ({ attachId, url: `octo-image:${attachId}`, fileName: attachId })), notFound: ids.filter((id) => !images.has(id)) }
    },
    async fetch(url, maxBytes) {
      const image = images.get(url.replace(/^octo-image:/, ''))
      if (!image || image.data.byteLength > maxBytes) return undefined
      const data = image.data.buffer.slice(image.data.byteOffset, image.data.byteOffset + image.data.byteLength) as ArrayBuffer
      const fallback = image.fallback?.buffer.slice(
        image.fallback.byteOffset,
        image.fallback.byteOffset + image.fallback.byteLength,
      ) as ArrayBuffer | undefined
      return { data, type: image.type, width: image.width, height: image.height, fallback }
    },
  }
  return Buffer.from(await serializeDocx('', doc, { imageAdapter: adapter }))
}
