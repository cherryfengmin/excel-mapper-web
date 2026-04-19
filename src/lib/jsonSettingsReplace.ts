import { decodeHtmlEntities, rebuildHtml, tokenizeHtml, stripHtmlTags } from './html'
import { lookupBest, type BuiltDictionary, normalizeText } from './dictionary'

export type ReplaceStats = {
  replaced: number
  unchanged: number
  skipped: number
  unmatched: number
  review: number
  unmatchedSamples: string[]
  settingsNodes: number
  replacements: Array<{ before: string; after: string }>
}

export type ReplaceOptions = {
  enableSubstringMatch: boolean
  maxUnmatchedSamples: number
  skipCssIdentifierMaxLen: number
  skipAlignmentKeywords: string[]
  allowRootAsSettingsWhenMissing: boolean
  skipKeysExact: string[]
  skipKeysRegex: RegExp[]
}

const DEFAULT_OPTS: ReplaceOptions = {
  enableSubstringMatch: true,
  maxUnmatchedSamples: 25,
  skipCssIdentifierMaxLen: 36,
  skipAlignmentKeywords: ['left', 'right', 'center', 'justify', 'top', 'bottom'],
  allowRootAsSettingsWhenMissing: true,
  skipKeysExact: [
    'section_css',
    'section_css_html',
    'block_css',
    'block_order',
    'addition_content_color',
    'addition_content_color_mb',
    'addition_content_alignment',
  ],
  skipKeysRegex: [
    /_color(_|$)/i,
    /_alignment(_|$)/i,
    /(^|_)css(_|$)/i,
  ],
}

export function replaceJsonSettings(
  jsonText: string,
  dict: BuiltDictionary,
  options?: Partial<ReplaceOptions>
): { ok: true; output: string; stats: ReplaceStats } | { ok: false; error: string } {
  const opts: ReplaceOptions = { ...DEFAULT_OPTS, ...(options ?? {}) }

  let root: any
  try {
    root = JSON.parse(jsonText)
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e) }
  }

  const stats: ReplaceStats = {
    replaced: 0,
    unchanged: 0,
    skipped: 0,
    unmatched: 0,
    review: 0,
    unmatchedSamples: [],
    settingsNodes: 0,
    replacements: [],
  }

  if (!root || typeof root !== 'object') {
    return { ok: true, output: JSON.stringify(root, null, 2), stats }
  }

  // Replace any nested `settings` nodes in the whole JSON tree.
  const hadRootSettings = hasOwn(root, 'settings')
  root = replaceAllSettingsNodes(root, dict, stats, opts)

  // If no `settings` key exists anywhere, optionally treat root as settings for quick paste of fragments.
  if (!hadRootSettings && stats.settingsNodes === 0 && opts.allowRootAsSettingsWhenMissing) {
    root = walk(root, dict, stats, opts)
  }

  return { ok: true, output: JSON.stringify(root, null, 2), stats }
}

function walk(node: any, dict: BuiltDictionary, stats: ReplaceStats, opts: ReplaceOptions): any {
  if (node === null || node === undefined) return node
  if (typeof node === 'string') return replaceStringSmart(node, dict, stats, opts)
  if (Array.isArray(node)) return node.map((x) => walk(x, dict, stats, opts))
  if (typeof node === 'object') {
    const out: any = Array.isArray(node) ? [] : {}
    for (const [k, v] of Object.entries(node)) {
      if (shouldSkipByKey(k, v, opts)) {
        stats.skipped += countStrings(v)
        out[k] = v
      } else {
        out[k] = walk(v, dict, stats, opts)
      }
    }
    return out
  }
  return node
}

function replaceAllSettingsNodes(node: any, dict: BuiltDictionary, stats: ReplaceStats, opts: ReplaceOptions): any {
  if (node === null || node === undefined) return node
  if (Array.isArray(node)) return node.map((x) => replaceAllSettingsNodes(x, dict, stats, opts))
  if (typeof node !== 'object') return node

  const out: any = {}
  for (const [k, v] of Object.entries(node)) {
    if (k === 'settings') {
      stats.settingsNodes += 1
      out[k] = walk(v, dict, stats, opts)
    } else {
      out[k] = replaceAllSettingsNodes(v, dict, stats, opts)
    }
  }
  return out
}

function hasOwn(obj: any, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(obj, key)
}

function shouldSkipByKey(key: string, value: any, opts: ReplaceOptions): boolean {
  const k = key.trim()
  if (!k) return false
  if (opts.skipKeysExact.includes(k)) return true
  for (const re of opts.skipKeysRegex) {
    if (re.test(k)) return true
  }

  // Special: arrays of ids (e.g. block_order) should be skipped even if key name differs slightly
  if (Array.isArray(value) && value.every((x) => typeof x === 'string')) {
    if (/order$/i.test(k) || /_order$/i.test(k)) return true
  }

  return false
}

function countStrings(node: any): number {
  if (node === null || node === undefined) return 0
  if (typeof node === 'string') return 1
  if (Array.isArray(node)) return node.reduce((acc, x) => acc + countStrings(x), 0)
  if (typeof node === 'object') {
    let n = 0
    for (const v of Object.values(node)) n += countStrings(v)
    return n
  }
  return 0
}

function replaceStringSmart(input: string, dict: BuiltDictionary, stats: ReplaceStats, opts: ReplaceOptions): string {
  const raw = input
  const t = normalizeText(raw)

  if (!t) {
    stats.unchanged += 1
    return raw
  }

  if (shouldSkipTechnical(raw, opts)) {
    stats.skipped += 1
    return raw
  }

  // Spec split: "Label: Value" or "Label：Value"
  const spec = splitSpecLabelValue(raw)
  if (spec) {
    const labelMatch = lookupBest(dict, spec.label)
    if (labelMatch.found) {
      stats.replaced += 1
      if (labelMatch.kind !== 'exact') stats.review += 1
      const next = `${labelMatch.value}${spec.sep}${spec.value}`
      recordReplacement(stats, raw, next)
      return next
    }
    // label not found: continue with other strategies on full string
  }

  // HTML tag protection
  if (containsHtmlTag(raw)) {
    const replaced = replaceHtmlPreserveTags(raw, dict, stats, opts)
    if (replaced !== null) {
      recordReplacement(stats, raw, replaced)
      return replaced
    }
    // if cannot confidently replace, try matching stripped version
    const stripped = stripHtmlTags(raw)
    const mStrip = lookupBest(dict, stripped)
    if (mStrip.found) {
      stats.replaced += 1
      if (mStrip.kind !== 'exact') stats.review += 1
      recordReplacement(stats, raw, mStrip.value)
      return mStrip.value
    }
    // fall through to plain match on full text
  }

  // Direct lookup variants
  const m = lookupBest(dict, raw)
  if (m.found) {
    stats.replaced += 1
    if (m.kind !== 'exact') stats.review += 1
    recordReplacement(stats, raw, m.value)
    return m.value
  }

  // Fallback: substring matching (risk-controlled)
  if (opts.enableSubstringMatch) {
    const sub = replaceByUniqueSubstring(raw, dict)
    if (sub.replaced) {
      stats.replaced += 1
      stats.review += 1
      recordReplacement(stats, raw, sub.value)
      if (sub.matchedNeedle) recordReplacement(stats, sub.matchedNeedle, sub.value)
      return sub.value
    }
  }

  stats.unmatched += 1
  if (stats.unmatchedSamples.length < opts.maxUnmatchedSamples) stats.unmatchedSamples.push(raw)
  return raw
}

function recordReplacement(stats: ReplaceStats, before: string, after: string) {
  if (before === after) return
  // avoid collecting extremely long blobs
  if (before.length > 8000 || after.length > 8000) return
  stats.replacements.push({ before, after })
}

function containsHtmlTag(s: string): boolean {
  return /<[^>]+>/.test(s)
}

function replaceHtmlPreserveTags(
  input: string,
  dict: BuiltDictionary,
  stats: ReplaceStats,
  opts: ReplaceOptions
): string | null {
  const tokens = tokenizeHtml(input)
  let changed = false

  const out = tokens.map((tok) => {
    if (tok.type === 'tag') return tok
    const text = tok.value
    const m = lookupBest(dict, text)
    if (m.found) {
      changed = true
      stats.replaced += 1
      if (m.kind !== 'exact') stats.review += 1
      return { type: 'text' as const, value: m.value }
    }

    // Try replacing pure text portion (strip tags) for this token if it contains entities
    const decoded = decodeHtmlEntities(text)
    const m2 = lookupBest(dict, decoded)
    if (m2.found) {
      changed = true
      stats.replaced += 1
      stats.review += 1
      return { type: 'text' as const, value: m2.value }
    }

    // As a last resort within HTML tokens, allow controlled substring replacement
    if (opts.enableSubstringMatch) {
      const sub = replaceByUniqueSubstring(text, dict)
      if (sub.replaced) {
        changed = true
        stats.replaced += 1
        stats.review += 1
        return { type: 'text' as const, value: sub.value }
      }
    }

    return tok
  })

  if (!changed) return null
  return rebuildHtml(out)
}

function splitSpecLabelValue(input: string): { label: string; sep: string; value: string } | null {
  // Only split on first colon-like separator
  const m = input.match(/^(.{1,80}?)([:：])\s*(.+)$/)
  if (!m) return null
  const label = m[1].trim()
  const sep = `${m[2]} `
  const value = m[3]
  if (!label || !value) return null
  return { label, sep, value }
}

function shouldSkipTechnical(value: string, opts: ReplaceOptions): boolean {
  const v = value.trim()
  if (!v) return true

  // URL
  if (/^https?:\/\//i.test(v)) return true

  // asset path / filename
  if (/\.(png|jpg|jpeg|svg|webp|gif|mp4|mov|pdf)(\?.*)?$/i.test(v)) return true

  // color
  if (/^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(v)) return true
  if (/^(rgb|rgba)\(/i.test(v)) return true

  // boolean / numeric
  if (/^(true|false)$/i.test(v)) return true
  if (/^-?\d+(\.\d+)?$/.test(v)) return true

  // alignment keywords
  if (opts.skipAlignmentKeywords.includes(v.toLowerCase())) return true

  // CSS-like identifier (avoid skipping normal words like "Lieferumfang")
  // Only skip when it *looks* like a class/token (contains '-' or '_' or digits), e.g. "title-large", "common_tab".
  if (
    /^[a-zA-Z_][\w-]*$/.test(v) &&
    v.length <= opts.skipCssIdentifierMaxLen &&
    /[-_\d]/.test(v)
  )
    return true

  return false
}

function replaceByUniqueSubstring(input: string, dict: BuiltDictionary): { replaced: boolean; value: string; matchedNeedle?: string } {
  // Conservative substring replacement:
  // - Only when input has no tags/entities (so we can safely mutate raw string)
  // - Only when there is exactly one candidate needle with word-ish boundary
  if (/[<&]/.test(input)) return { replaced: false, value: input }

  const hay = normalizeText(input)
  if (hay.length < 6) return { replaced: false, value: input }

  let matchKey: string | null = null
  let matchValue: string | null = null
  let matchNeedle: string | null = null

  for (const [k, v] of dict.map.entries()) {
    if (!k.startsWith('norm::')) continue
    const needle = k.slice('norm::'.length)
    if (needle.length < 6) continue

    // whole-word-ish boundary: require non-letter around match or edges
    const re = new RegExp(`(^|[^\\p{L}])${escapeRegExp(needle)}([^\\p{L}]|$)`, 'iu')
    if (!re.test(hay)) continue

    if (matchKey && matchKey !== k) {
      // more than one candidate => not unique
      return { replaced: false, value: input }
    }
    matchKey = k
    matchValue = v
    matchNeedle = needle
  }

  if (!matchNeedle || matchValue === null) return { replaced: false, value: input }

  const re = new RegExp(`(^|[^\\p{L}])(${escapeRegExp(matchNeedle)})([^\\p{L}]|$)`, 'iu')
  if (!re.test(input)) return { replaced: false, value: input }

  const next = input.replace(re, `$1${matchValue}$3`)
  return { replaced: next !== input, value: next, matchedNeedle: matchNeedle }
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

