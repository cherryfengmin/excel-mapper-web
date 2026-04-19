import type { RefObject } from 'react'
import { useLayoutEffect, useMemo, useRef, useState } from 'react'
import './App.css'
import {
  cellToString,
  columnIndexToLetter,
  getMaxColumnCount,
  readXlsxFile,
  sheetToMatrix,
  type ExcelWorkbook,
  type SheetMatrix,
} from './lib/excel'
import { buildMappingsFromRow, buildQuotePairs, type MappingRecord } from './lib/mapping'
import {
  buildDictionaryFromMappings,
  normalizeText,
  normalizeLoose,
  normalizeCollapseWhitespace,
  casefold,
  casefoldLoose,
  casefoldCollapseWhitespace,
} from './lib/dictionary'
import { replaceJsonSettings } from './lib/jsonSettingsReplace'
import { buildSettingsStringMappings } from './lib/settingsStringMappings'
import { decodeHtmlEntities, stripHtmlTags } from './lib/html'

/** JSON 编辑区：textarea（无行号） */
function JsonGutterTextarea(props: {
  value: string
  onChange: (next: string) => void
  placeholder?: string
  rows?: number
  textareaClassName?: string
  textareaRef?: RefObject<HTMLTextAreaElement | null>
  onPaste?: () => void
}) {
  const rows = props.rows ?? 10

  const setTextareaRef = (node: HTMLTextAreaElement | null) => {
    const r = props.textareaRef
    if (r) (r as { current: HTMLTextAreaElement | null }).current = node
  }

  return (
    <textarea
      ref={setTextareaRef}
      className={`textarea textareaMono${props.textareaClassName ? ` ${props.textareaClassName}` : ''}`}
      rows={rows}
      value={props.value}
      onChange={(e) => props.onChange(e.target.value)}
      placeholder={props.placeholder}
      onPaste={props.onPaste ? () => props.onPaste?.() : undefined}
    />
  )
}

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [workbook, setWorkbook] = useState<ExcelWorkbook | null>(null)
  const [sheetName, setSheetName] = useState<string>('')
  const [matrix, setMatrix] = useState<SheetMatrix | null>(null)

  const [startRow, setStartRow] = useState<number>(2)
  const [enableSentenceFallback, setEnableSentenceFallback] = useState(true)
  const [srcCol, setSrcCol] = useState<number | null>(null)
  const [dstCol, setDstCol] = useState<number | null>(null)
  const [markerPrefixesText, setMarkerPrefixesText] = useState('*\n•\n-\n[Icon]\n[Symbol]')
  const [query, setQuery] = useState('')
  const [statusFilter, setStatusFilter] = useState<'ALL' | 'OK' | 'NEED_REVIEW' | 'ERROR' | 'UNMATCHED'>('ALL')
  const [unmatchedQuery, setUnmatchedQuery] = useState('')
  const [copiedCells, setCopiedCells] = useState<Set<string>>(new Set())

  const [loadError, setLoadError] = useState<string>('')
  const [mappingRecords, setMappingRecords] = useState<MappingRecord[]>([])
  const [isDirty, setIsDirty] = useState(false)
  const [copyToast, setCopyToast] = useState<string>('')
  const [fieldsJson, setFieldsJson] = useState<string>('')
  const [fieldsJsonToast, setFieldsJsonToast] = useState<string>('')

  // JSON smart replace
  const [jsonInput, setJsonInput] = useState<string>('')
  const [jsonOutputEdited, setJsonOutputEdited] = useState<string>('') // editable output text
  const [jsonError, setJsonError] = useState<string>('')
  const [jsonToast, setJsonToast] = useState<string>('')
  const [jsonStats, setJsonStats] = useState<{
    replaced: number
    unchanged: number
    skipped: number
    unmatched: number
    review: number
    settingsNodes: number
  } | null>(null)
  const [jsonUnmatchedSamples, setJsonUnmatchedSamples] = useState<string[]>([])
  const [jsonDirection, setJsonDirection] = useState<'auto' | 'srcToDst' | 'dstToSrc'>('auto')
  const [jsonDictInfo, setJsonDictInfo] = useState<{ size: number; collisions: number } | null>(null)
  const [jsonAppliedDirection, setJsonAppliedDirection] = useState<'srcToDst' | 'dstToSrc' | ''>('')
  const [jsonReplacements, setJsonReplacements] = useState<Array<{ before: string; after: string }>>([])

  const headerRow = useMemo(() => (matrix && matrix.length ? matrix[0] ?? [] : []), [matrix])
  const srcHeaderName = useMemo(() => {
    if (srcCol === null) return ''
    const v = cellToString(headerRow[srcCol])
    return v.trim()
  }, [headerRow, srcCol])
  const dstHeaderName = useMemo(() => {
    if (dstCol === null) return ''
    const v = cellToString(headerRow[dstCol])
    return v.trim()
  }, [headerRow, dstCol])

  const colOptions = useMemo(() => {
    if (!matrix) return []
    const maxCols = getMaxColumnCount(matrix)
    return Array.from({ length: maxCols }, (_, i) => ({
      idx: i,
      label: `${columnIndexToLetter(i)}（第${i + 1}列）${cellToString((headerRow as any[])[i]).trim() ? ` · ${cellToString((headerRow as any[])[i]).trim()}` : ''}`,
    }))
  }, [matrix, headerRow])

  const fileName = file?.name ?? ''

  async function onPickFile(next: File | null) {
    setLoadError('')
    setFile(next)
    setWorkbook(null)
    setSheetName('')
    setMatrix(null)
    setSrcCol(null)
    setDstCol(null)
    setMappingRecords([])
    setIsDirty(false)

    if (!next) return

    try {
      const wb = await readXlsxFile(next)
      setWorkbook(wb)
      const first = wb.sheetNames[0] ?? ''
      setSheetName(first)
      if (first && wb.sheets[first]) {
        const m = sheetToMatrix(wb.sheets[first])
        setMatrix(m)
      }
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : String(e))
    }
  }

  function onChangeSheet(nextSheet: string) {
    if (!workbook) return
    setSheetName(nextSheet)
    const ws = workbook.sheets[nextSheet]
    if (!ws) {
      setMatrix(null)
      return
    }
    setMatrix(sheetToMatrix(ws))
    setSrcCol(null)
    setDstCol(null)
    setMappingRecords([])
    setIsDirty(false)
  }

  const markerPrefixes = useMemo(() => {
    return markerPrefixesText
      .split('\n')
      .map((x) => x.trim())
      .filter(Boolean)
  }, [markerPrefixesText])

  const sentenceAbbrevProtect = useMemo(
    () => ['z. B.', 'ca.', 'bzw.', 'u. a.', 'd. h.', 'z. T.', 'Nr.', 'Dr.'],
    []
  )

  const canUpdate = !!matrix && srcCol !== null && dstCol !== null

  async function onUpdateDocument() {
    setLoadError('')
    if (!canUpdate) return

    // Re-parse from selected file (useful if user re-selected a modified file)
    if (file) {
      try {
        const wb = await readXlsxFile(file)
        setWorkbook(wb)
        const ws = wb.sheets[sheetName] ?? wb.sheets[wb.sheetNames[0] ?? '']
        if (ws) setMatrix(sheetToMatrix(ws))
      } catch (e) {
        setLoadError(e instanceof Error ? e.message : String(e))
        return
      }
    }

    const m = matrix
    if (!m || srcCol === null || dstCol === null) return

    const startIdx0 = Math.max(0, startRow - 1)
    const out: MappingRecord[] = []
    for (let i = startIdx0; i < m.length; i++) {
      const row = m[i] ?? []
      const src = cellToString(row[srcCol])
      const dst = cellToString(row[dstCol])
      const rowId = i + 1
      out.push(
        ...buildMappingsFromRow(rowId, src, dst, {
          markerPrefixes,
          enableSentenceFallback,
          sentenceAbbrevProtect,
        })
      )
    }
    setMappingRecords(out)
    setIsDirty(false)
  }

  const filteredRecords = useMemo(() => {
    const q = query.trim().toLowerCase()
    return mappingRecords.filter((r) => {
      const status =
        r.row_status === 'ERROR'
          ? 'ERROR'
          : r.segment_status === 'UNMATCHED'
            ? 'UNMATCHED'
            : r.row_status === 'NEED_REVIEW'
              ? 'NEED_REVIEW'
              : 'OK'

      if (statusFilter !== 'ALL' && status !== statusFilter) return false

      if (!q) return true
      const hay = `${r.src_text}\n${r.dst_text}\n${r.row_status}\n${r.segment_status}\n${r.row_notes}`.toLowerCase()
      return hay.includes(q)
    })
  }, [mappingRecords, query, statusFilter])

  const canCopy = filteredRecords.some((r) => r.src_text && r.dst_text)
  const canExportFieldsJson = filteredRecords.some(
    (r) => r.row_status === 'OK' && r.segment_status === 'OK' && r.src_text && r.dst_text
  )

  const jsonInputLineCount = useMemo(() => Math.max(1, (jsonInput.match(/\n/g) ?? []).length + 1), [jsonInput])
  const jsonOutputLineCount = useMemo(
    () => Math.max(1, (jsonOutputEdited.match(/\n/g) ?? []).length + 1),
    [jsonOutputEdited]
  )

  const jsonUnmatchedMappings = useMemo(() => {
    if (!jsonStats || !jsonAppliedDirection) return []

    const jsonInputNorm = normalizeText(jsonInput)

    const buildVariants = (value: string) => {
      const decoded = decodeHtmlEntities(value)
      const stripped = stripHtmlTags(value)
      const stripDecoded = stripHtmlTags(decoded)
      return new Set([
        normalizeText(value),
        casefold(value),
        normalizeLoose(value),
        casefoldLoose(value),
        normalizeCollapseWhitespace(value),
        casefoldCollapseWhitespace(value),
        normalizeText(decoded),
        casefold(decoded),
        normalizeLoose(decoded),
        casefoldLoose(decoded),
        normalizeCollapseWhitespace(decoded),
        casefoldCollapseWhitespace(decoded),
        normalizeText(stripped),
        casefold(stripped),
        normalizeLoose(stripped),
        casefoldLoose(stripped),
        normalizeCollapseWhitespace(stripped),
        casefoldCollapseWhitespace(stripped),
        normalizeText(stripDecoded),
        casefold(stripDecoded),
        normalizeLoose(stripDecoded),
        casefoldLoose(stripDecoded),
        normalizeCollapseWhitespace(stripDecoded),
        casefoldCollapseWhitespace(stripDecoded),
      ].filter(Boolean))
    }

    const replacementVariants = Array.from(
      jsonReplacements.reduce((set: Set<string>, replacement) => {
        for (const variant of buildVariants(replacement.before)) {
          set.add(variant)
        }
        return set
      }, new Set<string>())
    )

    const beforeField = jsonAppliedDirection === 'srcToDst' ? 'src_text' : 'dst_text'
    const usableRows = mappingRecords.filter(
      (r) => r.row_status === 'OK' && r.segment_status === 'OK' && r.src_text && r.dst_text
    )

    const rowIsMatched = (before: string) => {
      const beforeVariants = Array.from(buildVariants(before))
      for (const bVar of beforeVariants) {
        if (!bVar) continue
        if (!jsonInputNorm.includes(bVar)) continue
        for (const repVar of replacementVariants) {
          if (!repVar) continue
          if (bVar === repVar) return true
          if (bVar.includes(repVar) || repVar.includes(bVar)) return true
        }
      }
      return false
    }

    return usableRows.filter((r) => {
      const before = r[beforeField as 'src_text' | 'dst_text']
      if (!before) return false
      if (!normalizeText(before)) return false
      if (!jsonInputNorm.includes(normalizeText(before))) return false
      return !rowIsMatched(before)
    })
  }, [jsonStats, jsonReplacements, jsonAppliedDirection, mappingRecords, jsonInput])

  const filteredUnmatchedMappings = useMemo(() => {
    const q = unmatchedQuery.trim().toLowerCase()
    if (!q) return jsonUnmatchedMappings
    return jsonUnmatchedMappings.filter((r) => {
      const hay = `${r.src_text}\n${r.dst_text}`.toLowerCase()
      return hay.includes(q)
    })
  }, [jsonUnmatchedMappings, unmatchedQuery])

  const jsonInputTextareaRef = useRef<HTMLTextAreaElement | null>(null)
  const jsonInputMoveCaretToStartAfterInput = useRef(false)

  useLayoutEffect(() => {
    if (!jsonInputMoveCaretToStartAfterInput.current) return
    jsonInputMoveCaretToStartAfterInput.current = false
    const el = jsonInputTextareaRef.current
    if (!el) return
    el.setSelectionRange(0, 0)
    el.scrollTop = 0
  }, [jsonInput])

  return (
    <div className="app">
      <header className="header">
        <div className="headerInner">
          <div>
            <h1>Excel 翻译映射</h1>
            <p className="subtle">
              上传 xlsx，选择两列做映射，生成 <code>"EN"：“DE”</code> 可复制列表
            </p>
          </div>
        </div>
      </header>

      <main className="main">
        <section className="card">
          <div className="cardHeader">
            <h2>1) 上传与配置</h2>
            <div className="row" />
          </div>
          <div className="grid2">
            <div className="field">
              <label htmlFor="file" className="fieldLabelTitle">
                <span className="labelRequired" aria-hidden="true">
                  *
                </span>
                <strong>Excel 文件（.xlsx）</strong>
              </label>
              <input
                id="file"
                type="file"
                accept=".xlsx"
                onChange={(e) => onPickFile(e.target.files?.[0] ?? null)}
              />
              {fileName ? <p className="help">已选择：{fileName}</p> : null}
              <p className="help">
                纯前端解析：不上传服务器；目前仅基于文本、真实换行与条目标记做切分。
              </p>
              {loadError ? <p className="help">解析失败：{loadError}</p> : null}
            </div>
            <div className="field">
              <label>参数</label>
              <div className="row">
                <label className="inline">
                  起始行
                  <input
                    type="number"
                    min={1}
                    value={startRow}
                    onChange={(e) => {
                      setStartRow(Number(e.target.value || 1))
                      setIsDirty(true)
                    }}
                  />
                </label>
                <label className="inline">
                  启用句子兜底
                  <input
                    type="checkbox"
                    checked={enableSentenceFallback}
                    onChange={(e) => {
                      setEnableSentenceFallback(e.target.checked)
                      setIsDirty(true)
                    }}
                  />
                </label>
              </div>
              <p className="help">
                兜底拆分会把结果标为 <code>NEED_REVIEW</code>，方便人工筛选复核。
              </p>
            </div>
          </div>

          <div className="grid2">
            <div className="field">
              <label className="fieldLabelTitle">
                <span className="labelRequired" aria-hidden="true">
                  *
                </span>
                <strong>工作表</strong>
              </label>
              <select
                disabled={!workbook}
                value={sheetName}
                onChange={(e) => onChangeSheet(e.target.value)}
              >
                {!workbook ? <option>（上传后可选）</option> : null}
                {workbook?.sheetNames.map((n) => (
                  <option key={n} value={n}>
                    {n}
                  </option>
                ))}
              </select>
            </div>
            <div className="field">
              <label className="fieldLabelTitle">
                <span className="labelRequired" aria-hidden="true">
                  *
                </span>
                <strong>条目标记（每行一个；用于切分 item）</strong>
              </label>
              <textarea
                className="textarea"
                rows={6}
                value={markerPrefixesText}
                onChange={(e) => {
                  setMarkerPrefixesText(e.target.value)
                  setIsDirty(true)
                }}
                placeholder="例如：*\n•\n-\n[Icon]"
              />
              <p className="help">
                规则：如果某段文本里出现这些“行首标记”，会按条目拆分；否则默认按换行或者空行分块。
              </p>
            </div>
          </div>

          <div className="grid2" style={{ marginTop: 12 }}>
            <div className="field">
              <label className="fieldLabelTitle">
                <span className="labelRequired" aria-hidden="true">
                  *
                </span>
                <strong>列映射</strong>
              </label>
              <div className="row">
                <select
                  disabled={!matrix}
                  value={srcCol ?? ''}
                  onChange={(e) =>
                    (setSrcCol(e.target.value === '' ? null : Number(e.target.value)), setIsDirty(true))
                  }
                >
                  <option value="">来源列（{srcHeaderName || '未选择'}）</option>
                  {colOptions.map((c) => (
                    <option key={c.idx} value={c.idx}>
                      {c.label}
                    </option>
                  ))}
                </select>
                <span className="arrow">→</span>
                <select
                  disabled={!matrix}
                  value={dstCol ?? ''}
                  onChange={(e) =>
                    (setDstCol(e.target.value === '' ? null : Number(e.target.value)), setIsDirty(true))
                  }
                >
                  <option value="">目标列（{dstHeaderName || '未选择'}）</option>
                  {colOptions.map((c) => (
                    <option key={c.idx} value={c.idx}>
                      {c.label}
                    </option>
                  ))}
                </select>
              </div>
            </div>
            <div className="field">
              <label>当前选择</label>
              <div className="kvs">
                <div>
                  <span className="k">起始行</span>
                  <span className="v">{startRow}</span>
                </div>
                <div>
                  <span className="k">来源列</span>
                  <span className="v">
                    {srcCol === null
                      ? '未选择'
                      : `${columnIndexToLetter(srcCol)}${srcHeaderName ? ` · ${srcHeaderName}` : ''}`}
                  </span>
                </div>
                <div>
                  <span className="k">目标列</span>
                  <span className="v">
                    {dstCol === null
                      ? '未选择'
                      : `${columnIndexToLetter(dstCol)}${dstHeaderName ? ` · ${dstHeaderName}` : ''}`}
                  </span>
                </div>
                <div>
                  <span className="k">记录数</span>
                  <span className="v">{mappingRecords.length}</span>
                </div>
              </div>
              <p className="help">
                复制输出会生成形如 <code>"Clear View, Smart Control"：“Klare Sicht, intelligente Steuerung”</code>
              </p>
              {isDirty ? (
                <p className="help">
                  配置已变更，点击 <span className="pill">输出映射</span> 重新生成结果。
                </p>
              ) : null}
            </div>
          </div>

          <div className="field" style={{ marginTop: 12 }}>
            <p className="help" style={{ marginTop: 0 }}>
              选择好两列与条目标记（如 <code>*</code>、<code>•</code>、<code>[Icon]</code>）后，点击下方生成映射表。
            </p>
            <div className="row" style={{ marginTop: 10, justifyContent: 'center' }}>
              <button className="btn btnPrimary" disabled={!canUpdate} onClick={onUpdateDocument}>
                输出映射
              </button>
            </div>
          </div>
        </section>

        <section className="card">
          <div className="cardHeader">
            <h2>2) 映射结果</h2>
            <div className="row">
              <input
                className="search"
                placeholder="搜索（EN/DE/notes）"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
              />
              <select
                className="selectSlim"
                value={statusFilter}
                onChange={(e) => setStatusFilter(e.target.value as typeof statusFilter)}
              >
                <option value="ALL">全部</option>
                <option value="OK">OK</option>
                <option value="NEED_REVIEW">NEED_REVIEW</option>
                <option value="UNMATCHED">UNMATCHED</option>
                <option value="ERROR">ERROR</option>
              </select>
              <button
                className="btn"
                disabled={!canCopy}
                onClick={async () => {
                  const text = buildQuotePairs(filteredRecords)
                  await navigator.clipboard.writeText(text)
                  setCopyToast('复制成功')
                  window.setTimeout(() => setCopyToast(''), 1500)
                }}
              >
                复制为 “EN：DE”
              </button>
              <button
                className="btn"
                disabled={!canExportFieldsJson}
                onClick={() => {
                  const pairs: Record<string, string> = {}
                  for (const r of filteredRecords) {
                    if (r.row_status !== 'OK' || r.segment_status !== 'OK') continue
                    if (!r.src_text || !r.dst_text) continue
                    // last write wins if duplicates exist
                    pairs[r.src_text] = r.dst_text
                  }
                  const payload = {
                    sourceColumn: srcHeaderName || '源列',
                    targetColumn: dstHeaderName || '目标列',
                    generatedAt: new Date().toISOString(),
                    count: Object.keys(pairs).length,
                    mappings: pairs,
                  }
                  setFieldsJson(JSON.stringify(payload, null, 2))
                  setFieldsJsonToast('已生成字段json')
                  window.setTimeout(() => setFieldsJsonToast(''), 1500)
                }}
              >
                输出字段json
              </button>
              <button
                className="btn"
                disabled={!jsonInput.trim() || !jsonOutputEdited.trim()}
                onClick={() => {
                  const res = buildSettingsStringMappings(jsonInput, jsonOutputEdited)
                  if (!res.ok) {
                    setFieldsJsonToast(res.error)
                    window.setTimeout(() => setFieldsJsonToast(''), 2500)
                    return
                  }

                  const payload = {
                    sourceColumn: srcHeaderName || '源列',
                    targetColumn: dstHeaderName || '目标列',
                    generatedAt: new Date().toISOString(),
                    count: Object.keys(res.mappings).length,
                    mappings: res.mappings,
                    alignStats: res.stats,
                  }

                  setFieldsJson(JSON.stringify(payload, null, 2))
                  setFieldsJsonToast('字段json已更新')
                  window.setTimeout(() => setFieldsJsonToast(''), 1500)
                }}
              >
                更新字段json
              </button>
            </div>
          </div>
          {copyToast ? <div className="toast">{copyToast}</div> : null}
          {fieldsJsonToast ? <div className="toast">{fieldsJsonToast}</div> : null}
          {mappingRecords.length ? (
            <>
              <div className="help" style={{ marginTop: 6 }}>
                总计 {mappingRecords.length} 条，当前显示 {filteredRecords.length} 条（复制按钮只复制当前筛选且 EN/DE 都非空的记录）。
              </div>
              <div className="tableWrap" role="region" aria-label="Mapping table">
                <table className="table">
                  <thead>
                    <tr>
                      <th>序号</th>
                      <th>row</th>
                      <th>block</th>
                      <th>item</th>
                      <th>{srcHeaderName || '源列'}</th>
                      <th>{dstHeaderName || '目标列'}</th>
                      <th>status</th>
                      <th>notes</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRecords.slice(0, 2000).map((r, idx) => {
                      const status =
                        r.row_status === 'ERROR'
                          ? 'ERROR'
                          : r.segment_status === 'UNMATCHED'
                            ? 'UNMATCHED'
                            : r.row_status === 'NEED_REVIEW'
                              ? 'NEED_REVIEW'
                              : 'OK'
                      return (
                        <tr key={`${r.row_id}-${r.block_index}-${r.item_index}-${idx}`}>
                          <td className="mono">{idx + 1}</td>
                          <td className="mono">{r.row_id}</td>
                          <td className="mono">{r.block_index}</td>
                          <td className="mono">{r.item_index}</td>
                          <td className="cell">
                            <div className={`cellContent ${copiedCells.has(`${r.row_id}-${r.block_index}-${r.item_index}-src`) ? 'copied' : ''}`}>
                              <span>{r.src_text}</span>
                              <button
                                className="copyIcon"
                                onClick={async () => {
                                  await navigator.clipboard.writeText(r.src_text)
                                  setCopiedCells(prev => new Set([...prev, `${r.row_id}-${r.block_index}-${r.item_index}-src`]))
                                  setCopyToast('源文本已复制')
                                  window.setTimeout(() => setCopyToast(''), 1500)
                                }}
                                title="复制源文本"
                              >
                                📋
                              </button>
                            </div>
                          </td>
                          <td className="cell">
                            <div className={`cellContent ${copiedCells.has(`${r.row_id}-${r.block_index}-${r.item_index}-dst`) ? 'copied' : ''}`}>
                              <span>{r.dst_text}</span>
                              <button
                                className="copyIcon"
                                onClick={async () => {
                                  await navigator.clipboard.writeText(r.dst_text)
                                  setCopiedCells(prev => new Set([...prev, `${r.row_id}-${r.block_index}-${r.item_index}-dst`]))
                                  setCopyToast('目标文本已复制')
                                  window.setTimeout(() => setCopyToast(''), 1500)
                                }}
                                title="复制目标文本"
                              >
                                📋
                              </button>
                            </div>
                          </td>
                          <td className={`badge badge_${status}`}>{status}</td>
                          <td className="notes">{r.row_notes}</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
              {filteredRecords.length > 2000 ? (
                <div className="help">为保证性能，仅展示前 2000 条。可通过筛选/搜索缩小范围。</div>
              ) : null}

              {fieldsJson ? (
                <div style={{ marginTop: 12 }}>
                  <div className="row" style={{ justifyContent: 'space-between' }}>
                    <div className="help" style={{ margin: 0 }}>
                      字段json（已格式化）
                    </div>
                    <button
                      className="btn"
                      onClick={async () => {
                        await navigator.clipboard.writeText(fieldsJson)
                        setFieldsJsonToast('复制成功')
                        window.setTimeout(() => setFieldsJsonToast(''), 1500)
                      }}
                    >
                      复制字段json
                    </button>
                  </div>
                  <textarea className="textarea textareaMono" rows={10} value={fieldsJson} readOnly />
                </div>
              ) : null}
            </>
          ) : canUpdate ? (
            <div className="placeholder mappingAreaHint">
              已选择列映射，但尚未生成数据。请点击下方的 <span className="pill">输出映射</span> 来生成映射结果。
            </div>
          ) : (
            <div className="placeholder mappingAreaHint">
              上传文件并选择列后，点击 <span className="pill">输出映射</span> 生成映射表。
            </div>
          )}
        </section>

        <section className="card">
          <div className="cardHeader">
            <h2>3) JSON 智能替换</h2>
            <div className="row">
              <select
                className="selectSlim"
                value={jsonDirection}
                onChange={(e) => setJsonDirection(e.target.value as typeof jsonDirection)}
                title="选择 JSON 当前语言与目标语言的替换方向"
              >
                <option value="auto">自动判断方向</option>
                <option value="srcToDst">{srcHeaderName || '源列'} → {dstHeaderName || '目标列'}</option>
                <option value="dstToSrc">{dstHeaderName || '目标列'} → {srcHeaderName || '源列'}</option>
              </select>
              <button
                className="btn btnPrimary"
                disabled={!mappingRecords.length || !jsonInput.trim()}
                onClick={() => {
                  setJsonError('')
                  setJsonToast('')
                  setJsonUnmatchedSamples([])
                  setJsonReplacements([])
                  setJsonDictInfo(null)
                  setJsonAppliedDirection('')

                  const commonDictOpts = {
                    includeOnlyOkRows: true,
                    includeMultilineLines: true,
                    allowDuplicatesSameTarget: true,
                  } as const

                  const runReplace = (dir: 'srcToDst' | 'dstToSrc') => {
                    const dict = buildDictionaryFromMappings(mappingRecords, dir, commonDictOpts)
                    const res = replaceJsonSettings(jsonInput, dict, { enableSubstringMatch: true })
                    return { dir, dict, res }
                  }

                  const candidates =
                    jsonDirection === 'auto'
                      ? [runReplace('srcToDst'), runReplace('dstToSrc')]
                      : [runReplace(jsonDirection)]

                  // pick best (most replaced, then least unmatched)
                  let best = candidates[0]!
                  for (const c of candidates.slice(1)) {
                    if (!c.res.ok && best.res.ok) continue
                    if (!best.res.ok && c.res.ok) {
                      best = c
                      continue
                    }
                    if (!c.res.ok || !best.res.ok) continue
                    if (c.res.stats.replaced > best.res.stats.replaced) best = c
                    else if (
                      c.res.stats.replaced === best.res.stats.replaced &&
                      c.res.stats.unmatched < best.res.stats.unmatched
                    )
                      best = c
                  }

                  setJsonDictInfo({ size: best.dict.size, collisions: best.dict.collisions.length })
                  setJsonAppliedDirection(best.dir)

                  const res = best.res

                  if (!res.ok) {
                    setJsonError(res.error)
                    setJsonOutputEdited('')
                    setJsonStats(null)
                    setJsonReplacements([])
                    return
                  }

                  setJsonOutputEdited(res.output)
                  setJsonStats({
                    replaced: res.stats.replaced,
                    unchanged: res.stats.unchanged,
                    skipped: res.stats.skipped,
                    unmatched: res.stats.unmatched,
                    review: res.stats.review,
                    settingsNodes: res.stats.settingsNodes,
                  })
                  setJsonUnmatchedSamples(res.stats.unmatchedSamples)
                  setJsonReplacements(res.stats.replacements)

                  if (res.stats.replaced === 0) {
                    setJsonToast(
                      '未替换任何字段：请检查替换方向是否选对，且 JSON.settings 的值需在映射表中可匹配（可展开查看未匹配示例）。'
                    )
                    window.setTimeout(() => setJsonToast(''), 5000)
                  }
                }}
              >
                应用映射
              </button>
              <button
                className="btn"
                disabled={!jsonOutputEdited.trim()}
                onClick={async () => {
                  await navigator.clipboard.writeText(jsonOutputEdited)
                  setJsonToast('复制成功')
                  window.setTimeout(() => setJsonToast(''), 1500)
                }}
              >
                复制替换后 JSON
              </button>
              <button
                className="btn"
                onClick={() => {
                  setJsonInput('')
                  setJsonOutputEdited('')
                  setJsonError('')
                  setJsonStats(null)
                  setJsonReplacements([])
                }}
              >
                清空
              </button>
            </div>
          </div>

          <div className="grid2">
            <div className="field">
              <label>输入 JSON（原始）</label>
              <JsonGutterTextarea
                value={jsonInput}
                onChange={setJsonInput}
                placeholder='例如：{"settings":{"title":"Highlights"}}'
                rows={Math.max(6, jsonInputLineCount)}
                textareaClassName="jsonTextareaAutoHeight"
                textareaRef={jsonInputTextareaRef}
                onPaste={() => {
                  jsonInputMoveCaretToStartAfterInput.current = true
                }}
              />
              <p className="help">只会替换 JSON 顶层字段 <code>settings</code> 内部的字符串值（递归）。</p>
              {jsonError ? <p className="help">解析失败：{jsonError}</p> : null}
            </div>
            <div className="field">
              <label>输出 JSON（替换后）</label>
              <JsonGutterTextarea
                value={jsonOutputEdited}
                onChange={setJsonOutputEdited}
                rows={Math.max(6, jsonOutputLineCount)}
                textareaClassName="jsonTextareaAutoHeight"
              />
              {jsonStats ? (
                <p className="help">
                  统计：replaced={jsonStats.replaced} unchanged={jsonStats.unchanged} skipped={jsonStats.skipped}{' '}
                  unmatched={jsonStats.unmatched} review={jsonStats.review} settingsNodes={jsonStats.settingsNodes}
                </p>
              ) : (
                <p className="help">点击 <code>应用映射</code> 生成替换后的 JSON。</p>
              )}
              {jsonDictInfo ? (
                <p className="help">
                  字典：keys={jsonDictInfo.size} collisions={jsonDictInfo.collisions}（仅使用 OK 映射生成）
                </p>
              ) : null}
              {jsonAppliedDirection ? (
                <p className="help">
                  本次方向：{jsonAppliedDirection === 'srcToDst' ? `${srcHeaderName || '源列'} → ${dstHeaderName || '目标列'}` : `${dstHeaderName || '目标列'} → ${srcHeaderName || '源列'}`}
                </p>
              ) : null}
              {jsonToast ? <div className="toast">{jsonToast}</div> : null}
              {jsonUnmatchedSamples.length ? (
                <details style={{ marginTop: 8 }}>
                  <summary className="help">未匹配示例（前 {jsonUnmatchedSamples.length} 条）</summary>
                  <div className="unmatched">
                    {jsonUnmatchedSamples.map((s, i) => (
                      <div key={i} className="unmatchedItem">
                        {s}
                      </div>
                    ))}
                  </div>
                </details>
              ) : null}
            </div>
          </div>

          {jsonStats && mappingRecords.length ? (
            <div style={{ marginTop: 12 }}>
              <div className="help">
                映射表总数：{mappingRecords.length} 条，已替换：{jsonStats.replaced} 条，未替换：{jsonUnmatchedMappings.length} 条
              </div>
              {jsonUnmatchedMappings.length ? (
                <div style={{ marginTop: 10 }}>
                  <div className="unmatchedTableTitle">
                    未应用的映射表
                  </div>
                  <input
                    className="search"
                    placeholder="搜索（EN/DE）"
                    value={unmatchedQuery}
                    onChange={(e) => setUnmatchedQuery(e.target.value)}
                    style={{ marginTop: 8, marginBottom: 8 }}
                  />
                  <div className="tableWrap" role="region" aria-label="未应用映射列表">
                    <table className="table unmatchedTable">
                      <thead>
                        <tr>
                          <th>序号</th>
                          <th>row</th>
                          <th>{srcHeaderName || '源列'}</th>
                          <th>{dstHeaderName || '目标列'}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredUnmatchedMappings.slice(0, 100).map((r, idx) => (
                          <tr key={`${r.row_id}-${r.block_index}-${r.item_index}-${idx}`}>
                            <td className="mono">{idx + 1}</td>
                            <td className="mono">{r.row_id}</td>
                            <td className="cell">
                              <div className={`cellContent ${copiedCells.has(`${r.row_id}-${r.block_index}-${r.item_index}-src`) ? 'copied' : ''}`}>
                                <span>{r.src_text}</span>
                                <button
                                  className="copyIcon"
                                  onClick={async () => {
                                    await navigator.clipboard.writeText(r.src_text)
                                    setCopiedCells(prev => new Set([...prev, `${r.row_id}-${r.block_index}-${r.item_index}-src`]))
                                    setCopyToast('源文本已复制')
                                    window.setTimeout(() => setCopyToast(''), 1500)
                                  }}
                                  title="复制源文本"
                                >
                                  📋
                                </button>
                              </div>
                            </td>
                            <td className="cell">
                              <div className={`cellContent ${copiedCells.has(`${r.row_id}-${r.block_index}-${r.item_index}-dst`) ? 'copied' : ''}`}>
                                <span>{r.dst_text}</span>
                                <button
                                  className="copyIcon"
                                  onClick={async () => {
                                    await navigator.clipboard.writeText(r.dst_text)
                                    setCopiedCells(prev => new Set([...prev, `${r.row_id}-${r.block_index}-${r.item_index}-dst`]))
                                    setCopyToast('目标文本已复制')
                                    window.setTimeout(() => setCopyToast(''), 1500)
                                  }}
                                  title="复制目标文本"
                                >
                                  📋
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {filteredUnmatchedMappings.length > 100 ? (
                      <div className="help">仅展示前 100 条未替换映射。</div>
                    ) : null}
                    {filteredUnmatchedMappings.length === 0 && unmatchedQuery.trim() ? (
                      <div className="help">未找到匹配的映射。</div>
                    ) : null}
                  </div>
                </div>
              ) : null}
            </div>
          ) : null}

          <div className="row" style={{ marginTop: 10 }}>
            <span className="help" style={{ margin: 0 }}>
              示例：
            </span>
            <button
              className="btn"
              type="button"
              onClick={() =>
                setJsonInput(
                  JSON.stringify(
                    {
                      settings: {
                        htmlText: 'Hello<br>World',
                        htmlSpan: '<span class="x">Highlights</span><br>More',
                      },
                    },
                    null,
                    2
                  )
                )
              }
            >
              HTML 标签
            </button>
            <button
              className="btn"
              type="button"
              onClick={() =>
                setJsonInput(
                  JSON.stringify(
                    {
                      settings: {
                        spec: 'Product Name: Dreame Rotafly Steamer P7',
                        url: 'https://example.com/a.png',
                        cssClass: 'title-large',
                        align: 'center',
                        enabled: true,
                      },
                    },
                    null,
                    2
                  )
                )
              }
            >
              规格/跳过
            </button>
          </div>

          {!mappingRecords.length ? (
            <div className="placeholder" style={{ marginTop: 12 }}>
              请先在上方生成映射表（点击 <span className="pill">输出映射</span>），再进行 JSON 替换。
            </div>
          ) : null}
        </section>
      </main>
    </div>
  )
}

export default App
