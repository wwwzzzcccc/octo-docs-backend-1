// @ts-nocheck -- verbatim migration of the proven frontend exporter; covered by focused fidelity tests
export interface PmMark { type: string; attrs?: Record<string, unknown> }
export interface PmNode { type: string; attrs?: Record<string, unknown>; content?: PmNode[]; text?: string; marks?: PmMark[] }

const MAX_PM_TREE_DEPTH = 200
const MAX_PM_TREE_NODES = 200_000

/** Iterative preflight keeps every recursive serializer/walker inside a safe envelope. */
export function assertPmTreeBounds(root: PmNode): void {
  const stack: Array<{ node: PmNode; depth: number }> = [{ node: root, depth: 0 }]
  let count = 0
  while (stack.length) {
    const current = stack.pop()!
    if (current.depth > MAX_PM_TREE_DEPTH || ++count > MAX_PM_TREE_NODES) {
      throw new Error('document_tree_too_deep')
    }
    for (const child of current.node.content ?? []) stack.push({ node: child, depth: current.depth + 1 })
  }
}

export function sanitizeLinkHref(raw: string | null | undefined): string | null {
  if (!raw) return null
  const cleaned = raw.replace(/[\u0000-\u0020]+/g, '').toLowerCase()
  if (/^(?:\\\\|\/\/)/.test(cleaned)) return null
  const scheme = /^([a-z][a-z0-9+.-]*):/.exec(cleaned)?.[1]
  return !scheme || ['http', 'https', 'mailto', 'tel'].includes(scheme) ? raw : null
}
export function sanitizeBookmarkUrl(raw: string | null | undefined): string | null {
  if (!raw) return null
  return /^https?:\/\//i.test(raw.trim()) ? raw : null
}
export function safeCssColor(value: unknown): string | null {
  if (typeof value !== 'string') return null
  const v = value.trim()
  return /^(?:#[0-9a-f]{3,8}|rgba?\([\d.,%\s/]+\)|hsla?\([\d.,%\s/deg]+\)|[a-z]+)$/i.test(v) ? v : null
}
export function safeFontSize(value: unknown): string | null {
  if (typeof value !== 'string') return null
  const v = value.trim(); return /^\d+(?:\.\d+)?(?:px|pt|em|rem|%)$/i.test(v) ? v : null
}
export function explicitFontFamily(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined
  const face = value.split(',')[0]?.replace(/["']/g, '').trim()
  const generic = new Set(['serif','sans-serif','monospace','cursive','fantasy','system-ui','ui-serif','ui-sans-serif','ui-monospace'])
  return face && !generic.has(face.toLowerCase()) && /^[A-Za-z\u3400-\u9fff][A-Za-z0-9 \u3400-\u9fff_-]*$/.test(face) ? face : undefined
}
export type TextScript = 'cjk' | 'latin' | 'emoji' | 'other'
export function textScript(ch: string): TextScript {
  const cp = ch.codePointAt(0) ?? 0
  if ((cp >= 0x3400 && cp <= 0x9fff) || (cp >= 0xf900 && cp <= 0xfaff)) return 'cjk'
  if ((cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a) || (cp >= 0xc0 && cp <= 0x24f)) return 'latin'
  if (cp >= 0x1f000 || (cp >= 0x2600 && cp <= 0x27bf)) return 'emoji'
  return 'other'
}
export function fallbackFont(script: TextScript, heading = false): string {
  if (script === 'emoji') return 'Segoe UI Emoji'
  if (script === 'cjk') return heading ? 'Microsoft YaHei' : 'SimSun'
  return heading ? 'Arial' : 'Times New Roman'
}
export function splitByScript(text: string): Array<{ text: string; script: TextScript }> {
  const out: Array<{ text: string; script: TextScript }> = []
  for (const ch of text) {
    let script = textScript(ch)
    if (script === 'other') script = out.at(-1)?.script ?? 'latin'
    const last = out.at(-1); if (last?.script === script) last.text += ch; else out.push({ text: ch, script })
  }
  return out
}
