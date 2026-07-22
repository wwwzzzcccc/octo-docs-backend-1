import { describe, it, expect } from 'vitest'
import {
  collectReferencedAttachIds,
  collectFormulaLatex,
  probeFailingFormulas,
  isSupportedImage,
  sniffImageExt,
  prepareTypstImage,
} from '../src/api/routes/export.js'

// The PDF export must only download images the document actually references,
// never the full attachment list (a doc can carry orphaned/unreferenced
// uploads). collectReferencedAttachIds walks the ProseMirror JSON and returns
// exactly the attachIds embedded by `image` nodes — this is the correctness
// gate that bounds the image-prefetch DoS surface alongside the count/total
// byte caps in resolveInputs.
describe('collectReferencedAttachIds', () => {
  it('collects attachIds only from image nodes, nested at any depth', () => {
    const doc = {
      type: 'doc',
      content: [
        { type: 'image', attrs: { attachId: 'a1' } },
        {
          type: 'callout',
          content: [
            { type: 'paragraph', content: [{ type: 'image', attrs: { attachId: 'a2' } }] },
          ],
        },
        { type: 'paragraph', content: [{ type: 'text', text: 'no image here' }] },
      ],
    }
    const ids = collectReferencedAttachIds(doc)
    expect(ids).toEqual(new Set(['a1', 'a2']))
  })

  it('ignores non-image nodes and image nodes without an attachId', () => {
    const doc = {
      type: 'doc',
      content: [
        { type: 'image', attrs: { attachId: null } },
        { type: 'image', attrs: {} },
        { type: 'fileAttachment', attrs: { attachId: 'f1' } }, // not an image
        { type: 'image', attrs: { attachId: 'keep' } },
      ],
    }
    expect(collectReferencedAttachIds(doc)).toEqual(new Set(['keep']))
  })

  it('returns an empty set for empty / malformed input', () => {
    expect(collectReferencedAttachIds(null)).toEqual(new Set())
    expect(collectReferencedAttachIds({})).toEqual(new Set())
    expect(collectReferencedAttachIds({ type: 'doc' })).toEqual(new Set())
  })
})

// A corrupt or renamed non-image attachment would otherwise abort the entire
// typst compile (decode error -> HTTP 500). isSupportedImage sniffs the magic
// bytes so a bad upload is dropped like any other unresolved image.
describe('isSupportedImage', () => {
  const bytes = (...b: number[]) => Buffer.from([...b, ...new Array(Math.max(0, 12 - b.length)).fill(0)])
  it('accepts PNG/JPEG/GIF magic bytes and an SVG root', () => {
    expect(isSupportedImage(bytes(0x89, 0x50, 0x4e, 0x47))).toBe(true) // PNG
    expect(isSupportedImage(bytes(0xff, 0xd8, 0xff))).toBe(true) // JPEG
    expect(isSupportedImage(bytes(0x47, 0x49, 0x46, 0x38))).toBe(true) // GIF
    expect(isSupportedImage(Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"/>'))).toBe(true)
  })
  it('rejects non-image or truncated buffers', () => {
    expect(isSupportedImage(Buffer.from('not an image at all'))).toBe(false)
    expect(isSupportedImage(Buffer.from([0x89, 0x50]))).toBe(false) // too short
    expect(isSupportedImage(Buffer.alloc(0))).toBe(false)
  })

  it('rejects WebP and BMP (typst v0.13.1 cannot decode them)', () => {
    // A valid WebP/BMP would otherwise pass the sniff, get written into the
    // compile root, and abort the whole export with `unknown image format`.
    const webp = Buffer.from([0x52, 0x49, 0x46, 0x46, 0, 0, 0, 0, 0x57, 0x45, 0x42, 0x50])
    const bmp = Buffer.from([0x42, 0x4d, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
    expect(isSupportedImage(webp)).toBe(false)
    expect(isSupportedImage(bmp)).toBe(false)
  })
})

// typst picks its image decoder from the file EXTENSION. Uploads can be
// mislabeled (a JPEG saved as `111.png`, stored with mime image/png); naming
// the compile-root file `.png` over JPEG bytes fails with "Invalid PNG
// signature" and aborts the whole export (HTTP 500). sniffImageExt derives the
// extension from the actual magic bytes so typst decodes it correctly.
describe('sniffImageExt', () => {
  const bytes = (...b: number[]) => Buffer.from([...b, ...new Array(Math.max(0, 12 - b.length)).fill(0)])
  it('returns the extension from the real magic bytes, ignoring any declared mime', () => {
    expect(sniffImageExt(bytes(0x89, 0x50, 0x4e, 0x47))).toBe('png')
    expect(sniffImageExt(bytes(0xff, 0xd8, 0xff))).toBe('jpg')
    expect(sniffImageExt(bytes(0x47, 0x49, 0x46, 0x38))).toBe('gif')
    expect(sniffImageExt(Buffer.from('<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg"/>'))).toBe('svg')
  })
  it('picks jpg for JPEG bytes even when the upload claims png (the 500 repro)', () => {
    // Real regression: `111.png` (mime image/png) whose bytes are actually JPEG.
    const jpegBytes = bytes(0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46)
    expect(sniffImageExt(jpegBytes)).toBe('jpg')
  })
  it('returns null for unrecognised / truncated buffers', () => {
    expect(sniffImageExt(Buffer.from('not an image'))).toBeNull()
    expect(sniffImageExt(Buffer.from([0x89, 0x50]))).toBeNull()
    expect(sniffImageExt(Buffer.alloc(0))).toBeNull()
  })
})

describe('prepareTypstImage — SVG export', () => {
  it('keeps a real SVG extension and supplies sanitized SVG bytes to Typst', () => {
    const input = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="20" height="10"><script>alert(1)</script><rect width="20" height="10" fill="red"/></svg>')
    const prepared = prepareTypstImage(input)
    expect(prepared?.ext).toBe('svg')
    expect(prepared?.bytes.toString('utf8')).toContain('<rect')
    expect(prepared?.bytes.toString('utf8')).not.toContain('<script')
    // Regression guard: SVG XML must never be put in a fake .png input.
    expect(prepared?.bytes.subarray(0, 4).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47]))).toBe(false)
  })

  it('rejects malformed/active XML rather than passing it to Typst', () => {
    expect(prepareTypstImage(Buffer.from('<!DOCTYPE svg [<!ENTITY x SYSTEM "file:///etc/passwd">]><svg>&x;</svg>'))).toBeNull()
  })
})

// The per-formula fallback isolates a broken formula after a whole-document
// compile failure. collectFormulaLatex must surface exactly the unique formula
// strings (any nesting depth) so probing has the right candidate set.
describe('collectFormulaLatex', () => {
  it('collects unique inline/block math latex at any depth', () => {
    const doc = {
      type: 'doc',
      content: [
        { type: 'blockMath', attrs: { latex: 'a^2 + b^2' } },
        { type: 'paragraph', content: [{ type: 'inlineMath', attrs: { latex: '\\pi' } }] },
        {
          type: 'callout',
          content: [{ type: 'paragraph', content: [{ type: 'inlineMath', attrs: { latex: 'a^2 + b^2' } }] }],
        },
        { type: 'paragraph', content: [{ type: 'text', text: 'no math' }] },
      ],
    }
    expect(new Set(collectFormulaLatex(doc))).toEqual(new Set(['a^2 + b^2', '\\pi']))
  })

  it('ignores math nodes without a non-empty latex string', () => {
    const doc = {
      type: 'doc',
      content: [
        { type: 'blockMath', attrs: { latex: '' } },
        { type: 'blockMath', attrs: {} },
        { type: 'inlineMath', attrs: { latex: 'x' } },
      ],
    }
    expect(collectFormulaLatex(doc)).toEqual(['x'])
  })
})

// Attribution is batched and MUST stay bounded: it holds a scarce compile slot,
// so an unbounded search is a DoS. On exhaustion unresolved formulas are marked
// verbatim; already-proven formulas can remain math in the partial retry.
describe('probeFailingFormulas — bounding', () => {
  it('conservatively returns all formulas when no probe compile is allowed', async () => {
    const many = Array.from({ length: 100 }, (_, i) => `x_${i}`)
    const { failing, exhausted } = await probeFailingFormulas(many, 't', { maxProbes: 0, budgetMs: 15_000 })
    expect(exhausted).toBe(true)
    expect(failing).toEqual(new Set(many))
  })

  it('does not exhaust for an empty formula set', async () => {
    const { failing, exhausted } = await probeFailingFormulas([], 't', { maxProbes: 40, budgetMs: 15_000 })
    expect(exhausted).toBe(false)
    expect(failing.size).toBe(0)
  })

  it('treats a zero probe budget as immediately exhausted', async () => {
    const { failing, exhausted } = await probeFailingFormulas(['a', 'b'], 't', { maxProbes: 40, budgetMs: 0 })
    expect(exhausted).toBe(true)
    expect(failing).toEqual(new Set(['a', 'b']))
  })
})
