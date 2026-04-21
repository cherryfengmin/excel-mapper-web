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
  /** internalKey -> original source text (direction-dependent) */
  sourceByKey: Map<string, string>
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
  sourceByKey: Map<string, string>,
  collisions: Array<{ key: string; existing: string; incoming: string }>,
  key: string,
  source: string,
  target: string,
  allowDuplicatesSameTarget: boolean
) {
  if (!key) return
  const prev = map.get(key)
  if (prev === undefined) {
    map.set(key, target)
    sourceByKey.set(key, source)
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
  const sourceByKey = new Map<string, string>()
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

    addKey(map, sourceByKey, collisions, `raw::${raw}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `norm::${normalized}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `cf::${cf}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `loose::${loose}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `loose_cf::${looseCf}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `ws::${ws}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `ws_cf::${wsCf}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `d_norm::${decodedNorm}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `d_cf::${decodedCf}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `s_norm::${stripNorm}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `s_cf::${stripCf}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `sd_norm::${stripDecodedNorm}`, source, target, options.allowDuplicatesSameTarget)
    addKey(map, sourceByKey, collisions, `sd_cf::${stripDecodedCf}`, source, target, options.allowDuplicatesSameTarget)

    // Also store loose variants for decoded/stripped (helps HTML label tokens like "Filterinformationen:")
    addKey(
      map,
      sourceByKey,
      collisions,
      `d_loose::${normalizeLoose(decoded)}`,
      source,
      target,
      options.allowDuplicatesSameTarget
    )
    addKey(
      map,
      sourceByKey,
      collisions,
      `s_loose::${normalizeLoose(stripped)}`,
      source,
      target,
      options.allowDuplicatesSameTarget
    )
    addKey(
      map,
      sourceByKey,
      collisions,
      `sd_loose::${normalizeLoose(stripDecoded)}`,
      source,
      target,
      options.allowDuplicatesSameTarget
    )

    if (options.includeMultilineLines && normalized.includes('\n')) {
      const lines = normalized.split('\n').map((x) => x.trim()).filter(Boolean)
      for (const ln of lines) {
        addKey(map, sourceByKey, collisions, `line_norm::${ln}`, source, target, options.allowDuplicatesSameTarget)
        addKey(map, sourceByKey, collisions, `line_cf::${ln.toLowerCase()}`, source, target, options.allowDuplicatesSameTarget)
      }
    }

    // Keep metadata for debugging later if needed (currently unused)
    void rowId
  }

  return { map, sourceByKey, collisions, size: map.size }
}

export type MatchKind = 'exact' | 'decoded' | 'stripped' | 'casefold' | 'line' | 'none'

export function lookupBest(dict: BuiltDictionary, input: string): {
  found: boolean
  value: string
  kind: MatchKind
  /** matched internal dictionary key (e.g. "raw::...") */
  dictKey?: string
} {
  const raw = input
  const norm = normalizeText(input)
  const cf = casefold(input)
  const loose = normalizeLoose(input)
  const looseCf = casefoldLoose(input)
  const ws = normalizeCollapseWhitespace(input)
  const wsCf = casefoldCollapseWhitespace(input)

  const k0 = `raw::${raw}`
  const v0 = dict.map.get(k0)
  if (v0 !== undefined) return { found: true, value: v0, kind: 'exact', dictKey: k0 }

  const k1 = `norm::${norm}`
  const v1 = dict.map.get(k1)
  if (v1 !== undefined) return { found: true, value: v1, kind: 'exact', dictKey: k1 }

  const kl = `loose::${loose}`
  const vl = dict.map.get(kl)
  if (vl !== undefined) return { found: true, value: vl, kind: 'exact', dictKey: kl }

  const kws = `ws::${ws}`
  const vws = dict.map.get(kws)
  if (vws !== undefined) return { found: true, value: vws, kind: 'exact', dictKey: kws }

  const decoded = decodeHtmlEntities(input)
  const decodedNorm = normalizeText(decoded)
  const kd = `d_norm::${decodedNorm}`
  const vd = dict.map.get(kd)
  if (vd !== undefined) return { found: true, value: vd, kind: 'decoded', dictKey: kd }

  const kd2 = `d_loose::${normalizeLoose(decoded)}`
  const vdLoose = dict.map.get(kd2)
  if (vdLoose !== undefined) return { found: true, value: vdLoose, kind: 'decoded', dictKey: kd2 }

  const stripped = stripHtmlTags(input)
  const stripNorm = normalizeText(stripped)
  const ks = `s_norm::${stripNorm}`
  const vs = dict.map.get(ks)
  if (vs !== undefined) return { found: true, value: vs, kind: 'stripped', dictKey: ks }

  const ks2 = `s_loose::${normalizeLoose(stripped)}`
  const vsLoose = dict.map.get(ks2)
  if (vsLoose !== undefined) return { found: true, value: vsLoose, kind: 'stripped', dictKey: ks2 }

  const vcfKeyCandidates = [
    `cf::${cf}`,
    `d_cf::${casefold(decoded)}`,
    `s_cf::${casefold(stripped)}`,
    `loose_cf::${looseCf}`,
    `ws_cf::${wsCf}`,
  ] as const
  for (const k of vcfKeyCandidates) {
    const v = dict.map.get(k)
    if (v !== undefined) return { found: true, value: v, kind: 'casefold', dictKey: k }
  }

  if (norm.includes('\n')) {
    const lines = norm.split('\n').map((x) => x.trim()).filter(Boolean)
    for (const ln of lines) {
      const kln = `line_norm::${ln}`
      const kcf = `line_cf::${ln.toLowerCase()}`
      const vl = dict.map.get(kln)
      if (vl !== undefined) return { found: true, value: vl, kind: 'line', dictKey: kln }
      const vcf = dict.map.get(kcf)
      if (vcf !== undefined) return { found: true, value: vcf, kind: 'line', dictKey: kcf }
    }
  }

  return { found: false, value: '', kind: 'none' }
}

