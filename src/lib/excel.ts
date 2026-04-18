import * as XLSX from 'xlsx'

export type ExcelWorkbook = {
  sheetNames: string[]
  sheets: Record<string, XLSX.WorkSheet>
}

export async function readXlsxFile(file: File): Promise<ExcelWorkbook> {
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type: 'array' })
  const sheets: Record<string, XLSX.WorkSheet> = {}
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name]
    if (ws) sheets[name] = ws
  }
  return { sheetNames: wb.SheetNames.slice(), sheets }
}

export type SheetMatrix = (string | number | boolean | null)[][]

export function sheetToMatrix(ws: XLSX.WorkSheet): SheetMatrix {
  // header: 1 -> array-of-arrays, raw: false -> formatted text where possible
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }) as SheetMatrix
}

export function getMaxColumnCount(matrix: SheetMatrix): number {
  let max = 0
  for (const row of matrix) max = Math.max(max, row.length)
  return max
}

export function columnIndexToLetter(idx0: number): string {
  // 0 -> A
  let n = idx0 + 1
  let s = ''
  while (n > 0) {
    const r = (n - 1) % 26
    s = String.fromCharCode(65 + r) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

export function columnLetterToIndex(letter: string): number | null {
  const t = letter.trim().toUpperCase()
  if (!/^[A-Z]+$/.test(t)) return null
  let n = 0
  for (let i = 0; i < t.length; i++) {
    n = n * 26 + (t.charCodeAt(i) - 64)
  }
  return n - 1
}

export function cellToString(v: string | number | boolean | null | undefined): string {
  if (v === null || v === undefined) return ''
  return String(v)
}
