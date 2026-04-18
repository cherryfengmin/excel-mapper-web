export type RowStatus = 'OK' | 'NEED_REVIEW' | 'ERROR'
export type SegmentStatus = 'OK' | 'UNMATCHED' | 'NEED_REVIEW' | 'STRIKED'
export type SplitMethod = 'none' | 'marker' | 'inline_asterisk' | 'sentence_fallback'

export type MappingRecord = {
  row_id: number
  block_index: number
  item_index: number
  src_text: string
  dst_text: string
  row_status: RowStatus
  row_notes: string
  split_method: SplitMethod
  segment_status: SegmentStatus
}

export type MappingConfig = {
  startRow: number
  srcCol: number
  dstCol: number
  enableSentenceFallback: boolean
  markerPrefixes: string[] // e.g. ["*", "•", "-", "[Icon]", "[Symbol]"]
  sentenceAbbrevProtect: string[] // e.g. ["z. B.", "ca.", "bzw."]
}

export function normalizeText(s: string): string {
  return s.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
}

function trimSegment(s: string): string {
  let t = s.replace(/^\s+|\s+$/g, '')
  // Some Excel exports wrap text with quotes or tabs; strip them at boundaries only.
  t = t.replace(/^[\t"]+/, '').replace(/[\t"]+$/, '')
  return t.replace(/^\s+|\s+$/g, '')
}

function splitBlocks(text: string): string[] {
  const t = normalizeText(text)
  // split on blank lines (e.g. "\n\n" or "\n  \n")
  return t
    .split(/\n\s*\n+/g)
    .map((x) => trimSegment(x))
    .filter(Boolean)
}

function isMarkerLine(line: string, markerPrefixes: string[]): boolean {
  const t = line.trim()
  if (!t) return false
  for (const m of markerPrefixes) {
    if (m.startsWith('[')) {
      // [Icon] style marker: must be at line start
      if (t.startsWith(m)) return true
      continue
    }
    if (t.startsWith(m)) return true
  }
  // numeric bullets: 1. 2) (1)
  if (/^(\(?\d+\)?[.)])\s+/.test(t)) return true
  return false
}

/** EN [Icon] / DE [Symbol] on their own line inside a merged paragraph */
function splitEmbeddedGraphicMarkerList(block: string, markerPrefixes: string[]): string[] | null {
  const t = normalizeText(block)
  if (!/\[(?:Icon|Symbol)\]/.test(t)) return null
  const chunks = t.split(/\n\s*\[(?:Icon|Symbol)\]\s*\n/)
  if (chunks.length < 2) return null
  const prefix = trimSegment(chunks[0]!)
  const tail = trimSegment(chunks.slice(1).join('\n'))
  if (!tail) return null
  const tailLines = tail
    .split('\n')
    .map((x) => trimSegment(x))
    .filter(Boolean)
  if (tailLines.length < 1) return null
  const tailHasMarkers = tailLines.some((ln) => isMarkerLine(ln, markerPrefixes))
  if (tailHasMarkers) return null
  const out: string[] = []
  if (prefix) out.push(prefix)
  out.push(...tailLines)
  return out.length >= 2 ? out : null
}

function splitItemsByMarkers(block: string, markerPrefixes: string[]): { items: string[]; used: boolean } {
  const lines = normalizeText(block).split('\n')

  // Special case: "[Icon]" / "[Symbol]" header line followed by a newline-separated list.
  // Example:
  // [Icon]
  // Fruits and vegetables
  // Meat and fish
  // ...
  // In this structure, the first line is a header marker but the remaining lines are the true items.
  const nonEmpty = lines.map((x) => x.trim()).filter(Boolean)
  if (nonEmpty.length >= 2) {
    const first = nonEmpty[0]!
    const bracketMarkers = markerPrefixes.filter((m) => m.startsWith('['))
    const isHeaderMarker = bracketMarkers.some((m) => first === m || first.startsWith(`${m} `))
    if (isHeaderMarker) {
      const rest = nonEmpty.slice(1)
      const restHasMarkers = rest.some((ln) => isMarkerLine(ln, markerPrefixes))
      if (!restHasMarkers) {
        return { items: rest.map((x) => trimSegment(x)).filter(Boolean), used: true }
      }
    }
  }

  // Fallback: pure newline list without explicit markers.
  // Example:
  // Reduced by up to 90%
  // Before
  // After
  // In such cases, treat each non-empty line as its own item instead of joining them.
  if (nonEmpty.length >= 2) {
    const anyMarker = nonEmpty.some((ln) => isMarkerLine(ln, markerPrefixes))
    if (!anyMarker) {
      return { items: nonEmpty.map((x) => trimSegment(x)).filter(Boolean), used: true }
    }
  }

  const items: string[] = []
  let buf = ''
  let used = false

  for (const rawLine of lines) {
    const line = rawLine.trim()
    if (!line) continue
    if (isMarkerLine(line, markerPrefixes)) {
      used = true
      if (buf) items.push(trimSegment(buf))
      buf = line
    } else {
      buf = buf ? `${buf} ${line}` : line
    }
  }
  if (buf) items.push(trimSegment(buf))

  // One merged cell sometimes contains "\n[Icon]\n" / "\n[Symbol]\n" before a newline list; split it.
  if (items.length === 1 && items[0]) {
    const emb = splitEmbeddedGraphicMarkerList(items[0], markerPrefixes)
    if (emb) return { items: emb.map((x) => trimSegment(x)).filter(Boolean), used: true }
  }

  return { items, used }
}

function splitInlineAsterisk(text: string): { items: string[]; used: boolean } {
  // Handle: "... *A ... *B ..."
  const t = normalizeText(text)
  const starCount = (t.match(/\*/g) ?? []).length
  if (starCount < 2) return { items: [trimSegment(t)].filter(Boolean), used: false }

  // If there are already line-start asterisks, marker split should handle it; this is for inline.
  // Split on '*' but keep star with segment.
  const parts = t.split('*')
  const items: string[] = []
  for (let i = 1; i < parts.length; i++) {
    const p = trimSegment(parts[i])
    if (!p) continue
    items.push(`*${p}`)
  }
  if (items.length >= 2) return { items, used: true }
  return { items: [trimSegment(t)].filter(Boolean), used: false }
}

function protectAbbreviations(text: string, abbrevs: string[]): { protectedText: string; restore: (s: string) => string } {
  let out = text
  const tokens: Array<{ token: string; original: string }> = []

  for (let i = 0; i < abbrevs.length; i++) {
    const abbr = abbrevs[i]
    const token = `__ABBR_${i}__`
    if (out.includes(abbr)) {
      out = out.split(abbr).join(token)
      tokens.push({ token, original: abbr })
    }
  }

  const restore = (s: string) => {
    let r = s
    for (const t of tokens) r = r.split(t.token).join(t.original)
    return r
  }
  return { protectedText: out, restore }
}

function splitSentencesToTarget(text: string, targetCount: number, abbrevs: string[]): string[] {
  const t0 = trimSegment(normalizeText(text))
  if (!t0) return []
  if (targetCount <= 1) return [t0]

  const { protectedText, restore } = protectAbbreviations(t0, abbrevs)
  const raw = protectedText
    .split(/(?<=[.!?])\s+/g)
    .map((x) => trimSegment(x))
    .filter(Boolean)
    .map(restore)

  if (raw.length === 0) return [t0]

  // If we have enough sentences, group them into targetCount buckets.
  if (raw.length >= targetCount) {
    const items: string[] = []
    let idx = 0
    for (let i = 0; i < targetCount; i++) {
      const remaining = raw.length - idx
      const remainingBuckets = targetCount - i
      const take = Math.ceil(remaining / remainingBuckets)
      items.push(raw.slice(idx, idx + take).join(' '))
      idx += take
    }
    return items.map(trimSegment).filter(Boolean)
  }

  // Not enough sentences: return what we have (caller will mark NEED_REVIEW/UNMATCHED)
  return raw.map(trimSegment).filter(Boolean)
}

function alignItems(
  rowId: number,
  blockIndex: number,
  srcItems: string[],
  dstItems: string[],
  baseRowStatus: RowStatus,
  recordNotes: string,
  splitMethod: SplitMethod,
  blockBalanced: boolean
): MappingRecord[] {
  const n = Math.max(srcItems.length, dstItems.length)
  const out: MappingRecord[] = []

  for (let i = 0; i < n; i++) {
    const src = srcItems[i] ?? ''
    const dst = dstItems[i] ?? ''
    let segment_status: SegmentStatus = 'OK'
    if (!src || !dst) segment_status = 'UNMATCHED'
    if (src.startsWith('*') && !dst && src) segment_status = 'UNMATCHED'

    let row_status: RowStatus
    if (baseRowStatus === 'ERROR') row_status = 'ERROR'
    else if (!src || !dst) row_status = 'NEED_REVIEW'
    else if (blockBalanced) row_status = 'OK'
    else row_status = 'NEED_REVIEW'

    out.push({
      row_id: rowId,
      block_index: blockIndex,
      item_index: i + 1,
      src_text: src,
      dst_text: dst,
      row_status,
      row_notes: recordNotes,
      split_method: splitMethod,
      segment_status,
    })
  }
  return out
}

type IndexedItem = { blockIndex: number; text: string }

function toIndexedItems(
  blocks: string[],
  cfg: Pick<MappingConfig, 'markerPrefixes' | 'enableSentenceFallback' | 'sentenceAbbrevProtect'>
): { items: IndexedItem[]; split_method: SplitMethod } {
  const out: IndexedItem[] = []
  let anyUsed: SplitMethod = 'none'

  for (let b = 0; b < Math.max(blocks.length, 1); b++) {
    const block = blocks[b] ?? ''
    if (!trimSegment(block)) continue

    // Marker split (includes newline-list fallback inside splitItemsByMarkers)
    const marker = splitItemsByMarkers(block, cfg.markerPrefixes)

    let items = marker.items
    let split_method: SplitMethod = marker.used ? 'marker' : 'none'

    // 2) Inline asterisk split
    const inline = splitInlineAsterisk(block)
    if (inline.used && inline.items.length > items.length) {
      items = inline.items
      split_method = 'inline_asterisk'
    }

    // (We intentionally do NOT run sentence_fallback here; flatten alignment is already a fallback strategy.)

    if (split_method !== 'none') anyUsed = split_method === 'marker' ? 'marker' : anyUsed
    if (split_method === 'inline_asterisk') anyUsed = 'inline_asterisk'

    for (const it of items) out.push({ blockIndex: b + 1, text: trimSegment(it) })
  }

  return { items: out, split_method: anyUsed }
}

function alignIndexedItems(
  rowId: number,
  srcItems: IndexedItem[],
  dstItems: IndexedItem[],
  baseRowStatus: RowStatus,
  flatNotesForGaps: string,
  globalNotes: string,
  splitMethod: SplitMethod
): MappingRecord[] {
  const n = Math.max(srcItems.length, dstItems.length)
  const out: MappingRecord[] = []
  for (let i = 0; i < n; i++) {
    const src = srcItems[i]?.text ?? ''
    const dst = dstItems[i]?.text ?? ''
    const block_index = srcItems[i]?.blockIndex ?? dstItems[i]?.blockIndex ?? 1

    let segment_status: SegmentStatus = 'OK'
    if (!src || !dst) segment_status = 'UNMATCHED'
    if (src.startsWith('*') && !dst && src) segment_status = 'UNMATCHED'

    // Index-wise pairs with both sides filled are OK (flatten fixes cross-block drift).
    let row_status: RowStatus
    if (baseRowStatus === 'ERROR') row_status = 'ERROR'
    else if (!src || !dst) row_status = 'NEED_REVIEW'
    else row_status = 'OK'

    const gapNote = !src || !dst ? flatNotesForGaps : ''
    const row_notes = [globalNotes, gapNote].filter(Boolean).join(' | ')

    out.push({
      row_id: rowId,
      block_index,
      item_index: i + 1,
      src_text: src,
      dst_text: dst,
      row_status,
      row_notes,
      split_method: splitMethod,
      segment_status,
    })
  }
  return out
}

function scoreRecords(records: MappingRecord[]): number {
  // lower is better
  let s = 0
  for (const r of records) {
    if (r.segment_status === 'UNMATCHED') s += 10
    if (r.row_status === 'NEED_REVIEW') s += 2
    if (r.row_status === 'ERROR') s += 100
  }
  return s
}

export function buildMappingsFromRow(
  rowId: number,
  srcRaw: string,
  dstRaw: string,
  cfg: Pick<MappingConfig, 'markerPrefixes' | 'enableSentenceFallback' | 'sentenceAbbrevProtect'>
): MappingRecord[] {
  const srcCell = trimSegment(normalizeText(srcRaw))
  const dstCell = trimSegment(normalizeText(dstRaw))

  if (!srcCell && !dstCell) return []

  let row_status: RowStatus = 'OK'
  let globalRowNotes = ''

  if ((!!srcCell) !== (!!dstCell)) {
    row_status = 'ERROR'
    globalRowNotes = 'One side empty'
  }

  const srcBlocks = splitBlocks(srcCell)
  const dstBlocks = splitBlocks(dstCell)
  const blockCount = Math.max(srcBlocks.length, dstBlocks.length, 1)

  const byBlocks: MappingRecord[] = []

  for (let b = 0; b < blockCount; b++) {
    const srcBlock = srcBlocks[b] ?? ''
    const dstBlock = dstBlocks[b] ?? ''

    const srcMarker = splitItemsByMarkers(srcBlock, cfg.markerPrefixes)
    const dstMarker = splitItemsByMarkers(dstBlock, cfg.markerPrefixes)

    let srcItems = srcMarker.items
    let dstItems = dstMarker.items
    let split_method: SplitMethod = srcMarker.used || dstMarker.used ? 'marker' : 'none'

    // 2) Inline asterisk split (if mismatch and helpful)
    if (srcItems.length !== dstItems.length) {
      const srcInline = splitInlineAsterisk(srcBlock)
      const dstInline = splitInlineAsterisk(dstBlock)
      const improvedSrc = srcInline.used && srcInline.items.length > srcItems.length
      const improvedDst = dstInline.used && dstInline.items.length > dstItems.length
      if (improvedSrc || improvedDst) {
        if (improvedSrc) srcItems = srcInline.items
        if (improvedDst) dstItems = dstInline.items
        split_method = 'inline_asterisk'
      }
    }

    let blockNotes = ''

    // 3) Sentence fallback (strict)
    if (cfg.enableSentenceFallback && srcItems.length !== dstItems.length) {
      const srcHasMarker = srcMarker.used || splitInlineAsterisk(srcBlock).used
      const dstHasMarker = dstMarker.used || splitInlineAsterisk(dstBlock).used

      // Only split the side that has no markers, to match the other side count
      if (srcItems.length > dstItems.length && !dstHasMarker && dstBlock) {
        const next = splitSentencesToTarget(dstBlock, srcItems.length, cfg.sentenceAbbrevProtect)
        if (next.length !== dstItems.length) {
          dstItems = next
          split_method = 'sentence_fallback'
          blockNotes = appendNote(blockNotes, 'Sentence split fallback (dst)')
        }
      } else if (dstItems.length > srcItems.length && !srcHasMarker && srcBlock) {
        const next = splitSentencesToTarget(srcBlock, dstItems.length, cfg.sentenceAbbrevProtect)
        if (next.length !== srcItems.length) {
          srcItems = next
          split_method = 'sentence_fallback'
          blockNotes = appendNote(blockNotes, 'Sentence split fallback (src)')
        }
      }
    }

    if (srcItems.length !== dstItems.length) {
      blockNotes = appendNote(
        blockNotes,
        `Item count mismatch in block ${b + 1} (src=${srcItems.length}, dst=${dstItems.length})`
      )
    }

    const blockBalanced = srcItems.length === dstItems.length
    const recordNotes = [globalRowNotes, blockNotes].filter(Boolean).join(' | ')

    byBlocks.push(...alignItems(rowId, b + 1, srcItems, dstItems, row_status, recordNotes, split_method, blockBalanced))
  }

  // Fallback: if block-level alignment produced mismatches, try flattening all blocks into one ordered list and align by index.
  const needsFallback =
    row_status !== 'OK' ||
    byBlocks.some((r) => r.segment_status === 'UNMATCHED') ||
    byBlocks.some((r) => r.row_status === 'NEED_REVIEW')

  if (!needsFallback) return byBlocks

  const srcFlat = toIndexedItems(srcBlocks, cfg)
  const dstFlat = toIndexedItems(dstBlocks, cfg)
  const flatNotesForGaps = 'Flatten blocks for alignment'
  const flatSplit: SplitMethod =
    srcFlat.split_method === 'none' && dstFlat.split_method === 'none' ? 'none' : srcFlat.split_method === 'inline_asterisk' || dstFlat.split_method === 'inline_asterisk' ? 'inline_asterisk' : 'marker'

  const byFlat = alignIndexedItems(
    rowId,
    srcFlat.items,
    dstFlat.items,
    row_status,
    flatNotesForGaps,
    globalRowNotes,
    flatSplit
  )

  return scoreRecords(byFlat) <= scoreRecords(byBlocks) ? byFlat : byBlocks
}

function appendNote(notes: string, add: string): string {
  return notes ? `${notes} | ${add}` : add
}

export function buildQuotePairs(records: MappingRecord[]): string {
  // Output format: "<src>"：“<dst>”
  return records
    .filter((r) => r.src_text && r.dst_text)
    .map((r) => `"${escapeQuotes(r.src_text)}"：“${escapeQuotes(r.dst_text)}”`)
    .join('\n')
}

function escapeQuotes(s: string): string {
  return s.replace(/"/g, '\\"')
}
