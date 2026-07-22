import { gitHubEmojis } from '@tiptap/extension-emoji'

/**
 * Resolve the names accepted by the editor's GitHub emoji set without depending
 * on frontend state. Both canonical names and aliases are indexed once; keeping
 * unknown names unresolved lets the Markdown importer preserve them literally.
 */
const emojiByShortcode = new Map<string, string>()

for (const emoji of gitHubEmojis) {
  if (!emoji.emoji) continue
  emojiByShortcode.set(emoji.name, emoji.emoji)
  for (const shortcode of emoji.shortcodes) {
    if (!emojiByShortcode.has(shortcode)) emojiByShortcode.set(shortcode, emoji.emoji)
  }
}

export function resolveGitHubEmoji(shortcode: string): string | undefined {
  return emojiByShortcode.get(shortcode)
}
