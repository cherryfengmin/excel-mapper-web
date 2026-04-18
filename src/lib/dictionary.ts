import type { MappingRecord } from './mapping'
import { decodeHtmlEntities, stripHtmlTags } from './html'

export type DictKeyType =
  | 'raw'
  | 'normalized'
  | 'casefold'
  | 'decoded_normalized'
  | 'decoded_casefold'
  | 'strip_normalized'
  | 'strip_casefold'
  | 'strip_decoded_normalized'
  | 'strip_decoded_casefold'
  | 'line_normalized'
  | 'line_casefold'

export type DictionaryEntry = {
  key: string
  keyType: DictKeyType
  source: string
  target: string
  fromRowId: number
}

export type DictionaryBuildOptions = {
  includeOnlyOkRows: boolean
  includeMultilineLines: boolean
  allowDuplicatesSameTarget: boolean
}

export type BuiltDictionary = {
  map: Map<string, string>
  collisions: Array<{ key: string; existing: string; incoming: string }>
  size: number
}

export function normalizeText(s: string): string {
  // Normalize newlines + trim + Unicode normalization + common invisible chars + dash variants.
  const t = s.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
  const nfkc = typeof t.normalize === 'function' ? t.normalize('NFKC') : t
  return nfkc
    .replace(/[\u200B-\u200D\uFEFF]/g, '') // zero-width chars
    .replace(/[\u00A0]/g, ' ') // NBSP
    .replace(/[\u2010\u2011\u2012\u2013\u2014\u2212]/g, '-') // hyphen/dash variants to '-'
    .trim()
}

export function normalizeLoose(s: string): string {
  // For short labels / HTML tokens: collapse whitespace and strip trailing colon.
  return normalizeText(s)
    .replace(/\s+/g, ' ')
    .replace(/[:：]\s*$/g, '')
    .trim()
}

export function normalizeCollapseWhitespace(s: string): string {
  return normalizeText(s).replace(/\s+/g, ' ').trim()
}

export function casefold(s: string): string {
  return normalizeText(s).toLowerCase()
}

export function casefoldLoose(s: string): string {
  return normalizeLoose(s).toLowerCase()
}

export function casefoldCollapseWhitespace(s: string): string {
  return normalizeCollapseWhitespace(s).toLowerCase()
}

function addKey(
  map: Map<string, string>,
  collisions: Array<{ key: string; existing: string; incoming: string }>,
  key: string,
  target: string,
  allowDuplicatesSameTarget: boolean
) {
  if (!key) return
  const prev = map.get(key)
  if (prev === undefined) {
    map.set(key, target)
    return
  }
  if (allowDuplicatesSameTarget && prev === target) return
  collisions.push({ key, existing: prev, incoming: target })
}

export function buildDictionaryFromMappings(
  records: MappingRecord[],
  direction: 'srcToDst' | 'dstToSrc',
  options: DictionaryBuildOptions
): BuiltDictionary {
  const map = new Map<string, string>()
  const collisions: Array<{ key: string; existing: string; incoming: string }> = []

  const usable = options.includeOnlyOkRows
    ? records.filter((r) => r.row_status === 'OK' && r.segment_status === 'OK' && r.src_text && r.dst_text)
    : records.filter((r) => r.src_text && r.dst_text)

  for (const r of usable) {
    const source = direction === 'srcToDst' ? r.src_text : r.dst_text
    const target = direction === 'srcToDst' ? r.dst_text : r.src_text
    const rowId = r.row_id

    const raw = source
    const normalized = normalizeText(source)
    const cf = casefold(source)
    const loose = normalizeLoose(source)
    const looseCf = casefoldLoose(source)
    const ws = normalizeCollapseWhitespace(source)
    const wsCf = casefoldCollapseWhitespace(source)

    const decoded = decodeHtmlEntities(source)
    const decodedNorm = normalizeText(decoded)
    const decodedCf = casefold(decoded)

    const stripped = stripHtmlTags(source)
    const stripNorm = normalizeText(stripped)
    const stripCf = casefold(stripped)

    const stripDecoded = stripHtmlTags(decoded)
    const stripDecodedNorm = normalizeText(stripDecoded)
    const stripDecodedCf = casefold(stripDecoded)

    addKey(map, collisions, `raw::${raw}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `norm::${normalized}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `cf::${cf}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `loose::${loose}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `loose_cf::${looseCf}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `ws::${ws}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `ws_cf::${wsCf}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `d_norm::${decodedNorm}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `d_cf::${decodedCf}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `s_norm::${stripNorm}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `s_cf::${stripCf}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `sd_norm::${stripDecodedNorm}`, target, options.allowDuplicatesSameTarget)
    addKey(map, collisions, `sd_cf::${stripDecodedCf}`, target, options.allowDuplicatesSameTarget)

    // Also store loose variants for decoded/stripped (helps HTML label tokens like "Filterinformationen:")
    addKey(
      map,
      collisions,
      `d_loose::${normalizeLoose(decoded)}`,
      target,
      options.allowDuplicatesSameTarget
    )
    addKey(
      map,
      collisions,
      `s_loose::${normalizeLoose(stripped)}`,
      target,
      options.allowDuplicatesSameTarget
    )
    addKey(
      map,
      collisions,
      `sd_loose::${normalizeLoose(stripDecoded)}`,
      target,
      options.allowDuplicatesSameTarget
    )

    if (options.includeMultilineLines && normalized.includes('\n')) {
      const lines = normalized.split('\n').map((x) => x.trim()).filter(Boolean)
      for (const ln of lines) {
        addKey(map, collisions, `line_norm::${ln}`, target, options.allowDuplicatesSameTarget)
        addKey(map, collisions, `line_cf::${ln.toLowerCase()}`, target, options.allowDuplicatesSameTarget)
      }
    }

    // Keep metadata for debugging later if needed (currently unused)
    void rowId
  }

  return { map, collisions, size: map.size }
}

export type MatchKind = 'exact' | 'decoded' | 'stripped' | 'casefold' | 'line' | 'none'

export function lookupBest(dict: BuiltDictionary, input: string): { found: boolean; value: string; kind: MatchKind } {
  const raw = input
  const norm = normalizeText(input)
  const cf = casefold(input)
  const loose = normalizeLoose(input)
  const looseCf = casefoldLoose(input)
  const ws = normalizeCollapseWhitespace(input)
  const wsCf = casefoldCollapseWhitespace(input)

  const v0 = dict.map.get(`raw::${raw}`)
  if (v0 !== undefined) return { found: true, value: v0, kind: 'exact' }

  const v1 = dict.map.get(`norm::${norm}`)
  if (v1 !== undefined) return { found: true, value: v1, kind: 'exact' }

  const vl = dict.map.get(`loose::${loose}`)
  if (vl !== undefined) return { found: true, value: vl, kind: 'exact' }

  const vws = dict.map.get(`ws::${ws}`)
  if (vws !== undefined) return { found: true, value: vws, kind: 'exact' }

  const decoded = decodeHtmlEntities(input)
  const decodedNorm = normalizeText(decoded)
  const vd = dict.map.get(`d_norm::${decodedNorm}`)
  if (vd !== undefined) return { found: true, value: vd, kind: 'decoded' }

  const vdLoose = dict.map.get(`d_loose::${normalizeLoose(decoded)}`)
  if (vdLoose !== undefined) return { found: true, value: vdLoose, kind: 'decoded' }

  const stripped = stripHtmlTags(input)
  const stripNorm = normalizeText(stripped)
  const vs = dict.map.get(`s_norm::${stripNorm}`)
  if (vs !== undefined) return { found: true, value: vs, kind: 'stripped' }

  const vsLoose = dict.map.get(`s_loose::${normalizeLoose(stripped)}`)
  if (vsLoose !== undefined) return { found: true, value: vsLoose, kind: 'stripped' }

  const vcf =
    dict.map.get(`cf::${cf}`) ??
    dict.map.get(`d_cf::${casefold(decoded)}`) ??
    dict.map.get(`s_cf::${casefold(stripped)}`) ??
    dict.map.get(`loose_cf::${looseCf}`) ??
    dict.map.get(`ws_cf::${wsCf}`)
  if (vcf !== undefined) return { found: true, value: vcf, kind: 'casefold' }

  if (norm.includes('\n')) {
    const lines = norm.split('\n').map((x) => x.trim()).filter(Boolean)
    for (const ln of lines) {
      const vl = dict.map.get(`line_norm::${ln}`) ?? dict.map.get(`line_cf::${ln.toLowerCase()}`)
      if (vl !== undefined) return { found: true, value: vl, kind: 'line' }
    }
  }

  return { found: false, value: '', kind: 'none' }
}

