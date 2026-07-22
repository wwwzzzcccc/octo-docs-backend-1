import { describe, expect, it } from 'vitest'
import { parseXmlOrdered } from '../src/import/docx/xml.js'

function nestedXml(depth: number): Buffer {
  return Buffer.from('<w:r>'.repeat(depth) + '<w:t>x</w:t>' + '</w:r>'.repeat(depth), 'utf8')
}

describe('ordered OOXML parser nesting bound', () => {
  it('accepts a deeply nested but bounded valid OOXML tree', () => {
    expect(() => parseXmlOrdered(nestedXml(361))).not.toThrow()
  })

  it('still rejects nesting beyond the finite safety ceiling', () => {
    expect(() => parseXmlOrdered(nestedXml(513))).toThrow(/Maximum nested tags exceeded/)
  })
})
