type BuildResult =
  | { ok: true; mappings: Record<string, string>; stats: { pairs: number; used: number; skippedKeys: number } }
  | { ok: false; error: string }

const SKIP_KEYS_EXACT = new Set([
  'section_css',
  'section_css_html',
  'block_css',
  'block_order',
  'addition_content_color',
  'addition_content_color_mb',
  'addition_content_alignment',
])

const SKIP_KEYS_RE = [/(_|^)css(_|$)/i, /_color(_|$)/i, /_alignment(_|$)/i]

function shouldSkipKey(key: string): boolean {
  if (SKIP_KEYS_EXACT.has(key)) return true
  return SKIP_KEYS_RE.some((re) => re.test(key))
}

export function buildSettingsStringMappings(srcJsonText: string, dstJsonText: string): BuildResult {
  let src: any
  let dst: any
  try {
    src = JSON.parse(srcJsonText)
  } catch (e) {
    return { ok: false, error: `原始 JSON 解析失败：${e instanceof Error ? e.message : String(e)}` }
  }
  try {
    dst = JSON.parse(dstJsonText)
  } catch (e) {
    return { ok: false, error: `替换后 JSON 解析失败：${e instanceof Error ? e.message : String(e)}` }
  }

  const mappings: Record<string, string> = {}
  const stats = { pairs: 0, used: 0, skippedKeys: 0 }

  walkTogether(src, dst, (a, b) => {
    if (typeof a !== 'string' || typeof b !== 'string') return
    const k = a
    const v = b
    if (!k || !v) return
    stats.pairs += 1
    if (k !== v) {
      mappings[k] = v
      stats.used += 1
    }
  }, stats)

  return { ok: true, mappings, stats }
}

function walkTogether(
  a: any,
  b: any,
  onStringPair: (aStr: string, bStr: string) => void,
  stats: { skippedKeys: number }
) {
  if (a === null || a === undefined || b === null || b === undefined) return

  // If both are strings, record mapping
  if (typeof a === 'string' && typeof b === 'string') {
    onStringPair(a, b)
    return
  }

  // Arrays: align by index
  if (Array.isArray(a) && Array.isArray(b)) {
    const n = Math.min(a.length, b.length)
    for (let i = 0; i < n; i++) walkTogether(a[i], b[i], onStringPair, stats)
    return
  }

  // Objects: align by shared keys
  if (typeof a === 'object' && typeof b === 'object' && !Array.isArray(a) && !Array.isArray(b)) {
    const aObj = a as Record<string, any>
    const bObj = b as Record<string, any>
    for (const key of Object.keys(aObj)) {
      if (!(key in bObj)) continue

      if (key === 'settings') {
        walkTogether(aObj[key], bObj[key], onStringPair, stats)
        continue
      }

      if (shouldSkipKey(key)) {
        stats.skippedKeys += 1
        continue
      }

      walkTogether(aObj[key], bObj[key], onStringPair, stats)
    }
  }
}

