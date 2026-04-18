export function decodeHtmlEntities(input: string): string {
  // Fast path
  if (!/[&]/.test(input)) return input
  const ta = document.createElement('textarea')
  ta.innerHTML = input
  return ta.value
}

export function stripHtmlTags(input: string): string {
  return input.replace(/<[^>]*>/g, '')
}

export type HtmlToken = { type: 'tag'; value: string } | { type: 'text'; value: string }

export function tokenizeHtml(input: string): HtmlToken[] {
  const tokens: HtmlToken[] = []
  const re = /(<[^>]*>)/g
  let last = 0
  let m: RegExpExecArray | null
  while ((m = re.exec(input))) {
    const idx = m.index
    if (idx > last) tokens.push({ type: 'text', value: input.slice(last, idx) })
    tokens.push({ type: 'tag', value: m[1] })
    last = idx + m[1].length
  }
  if (last < input.length) tokens.push({ type: 'text', value: input.slice(last) })
  return tokens
}

export function rebuildHtml(tokens: HtmlToken[]): string {
  return tokens.map((t) => t.value).join('')
}
