import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { exportMarkdown } from '../src/export/markdown.js'
import { exportDocx } from '../src/export/docx.js'

describe('CLI export font alignment', () => {
  const doc = {
    type: 'doc',
    content: [{
      type: 'paragraph',
      content: [{
        type: 'text', text: '中 English',
        marks: [{ type: 'textStyle', attrs: { fontFamily: 'SimSun, "宋体", serif', fontSize: '14px', color: '#123456' } }],
      }],
    }],
  }

  it('preserves safe font family, size and color in Markdown HTML fallback', () => {
    expect(exportMarkdown(doc)).toContain("style=\"color:#123456;font-size:14px;font-family:SimSun\"")
  })

  it('writes script-aware defaults and explicit run fonts in DOCX OOXML', async () => {
    const zip = await JSZip.loadAsync(await exportDocx(doc))
    const styles = await zip.file('word/styles.xml')!.async('string')
    const body = await zip.file('word/document.xml')!.async('string')
    expect(styles).toContain('w:ascii="Times New Roman"')
    expect(styles).toContain('w:eastAsia="SimSun"')
    expect(styles).toContain('w:ascii="Arial"')
    expect(styles).toContain('w:eastAsia="Microsoft YaHei"')
    expect(body).toContain('w:ascii="SimSun"')
    // CSS 14px = 10.5pt = 21 OOXML half-points.
    expect(body).toContain('w:sz w:val="21"')
    expect(body).toContain('w:color w:val="123456"')
  })
})
