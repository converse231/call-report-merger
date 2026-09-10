import Papa from 'papaparse'
import ExcelJS from 'exceljs'
import * as XLSX from 'xlsx'
import { cellText } from './merge.js'

// Route by file signature, not extension -- HubSpot/Excel exports get renamed
// or mislabeled often enough (a real .xls next to a zip-based one) that trusting
// the name alone produces a confusing "is this a zip file?" crash instead.
const ZIP_SIG = [0x50, 0x4b] // .xlsx (and any modern format that's secretly a zip)
const CFBF_SIG = [0xd0, 0xcf, 0x11, 0xe0] // legacy binary .xls (BIFF8 / OLE2)
const startsWith = (bytes, sig) => sig.every((b, i) => bytes[i] === b)

/** Reads a .csv, .xlsx, or legacy .xls into { name, headers, rows }. */
export async function readTable(file) {
  const table = /\.csv$/i.test(file.name) ? await readCsv(file) : await readSpreadsheet(file)
  if (!table.headers.length) throw new Error(`No column headers found in the first row of "${file.name}".`)
  if (!table.rows.length) throw new Error(`"${file.name}" has headers but no data rows.`)
  return table
}

function readCsv(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: 'greedy',
      complete: (res) =>
        resolve({ name: file.name, headers: (res.meta.fields ?? []).filter(Boolean), rows: res.data }),
      error: reject,
    })
  })
}

async function readSpreadsheet(file) {
  const bytes = new Uint8Array(await file.arrayBuffer())
  if (startsWith(bytes, ZIP_SIG)) return readXlsx(bytes, file.name)
  if (startsWith(bytes, CFBF_SIG)) return readLegacyXls(bytes, file.name)
  throw new Error(
    `"${file.name}" doesn't look like an Excel file (no .xlsx or .xls signature found). ` +
      'If this is really a CSV, make sure it ends in .csv.',
  )
}

/** .xlsx (zip/OOXML). Keeps the live workbook so the base file's own cells can be preserved on write. */
async function readXlsx(bytes, name) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(bytes)
  // HubSpot exports tack on an empty "Sheet1"; take the first one with data.
  const sheet = wb.worksheets.find((w) => w.actualRowCount > 1) ?? wb.worksheets[0]
  if (!sheet) throw new Error(`"${name}" has no worksheets.`)

  const head = sheet.getRow(1)
  const headers = []
  for (let c = 1; c <= sheet.columnCount; c++) headers.push(cellText(head.getCell(c).value))
  while (headers.length && !headers[headers.length - 1]) headers.pop()

  const rows = []
  const sheetRowNumbers = []
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r)
    const obj = {}
    let blank = true
    headers.forEach((h, i) => {
      if (!h) return
      const v = row.getCell(i + 1).value
      obj[h] = v
      if (cellText(v)) blank = false
    })
    if (blank) continue
    rows.push(obj)
    sheetRowNumbers.push(r)
  }

  return { name, headers: headers.filter(Boolean), rows, workbook: wb, sheet, sheetRowNumbers }
}

/**
 * Legacy binary .xls (BIFF8). ExcelJS can't read this format at all, so we lean
 * on SheetJS for parsing only. There's no live workbook to write back into, so
 * a merge on this base rebuilds a fresh .xlsx on output (same as a CSV base) --
 * values come through correctly, the original file's own formatting just isn't
 * preserved cell-for-cell because it never gets re-opened.
 */
function readLegacyXls(bytes, name) {
  const wb = XLSX.read(bytes, { type: 'array', cellDates: true })
  let aoa = null
  for (const sheetName of wb.SheetNames) {
    const candidate = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null })
    if (candidate.length > 1) { aoa = candidate; break }
  }
  if (!aoa) throw new Error(`"${name}" has no worksheets with data.`)

  const headers = (aoa[0] ?? []).map(cellText)
  while (headers.length && !headers[headers.length - 1]) headers.pop()

  const rows = []
  for (let r = 1; r < aoa.length; r++) {
    const line = aoa[r]
    const obj = {}
    let blank = true
    headers.forEach((h, i) => {
      if (!h) return
      const v = line[i]
      obj[h] = v
      if (cellText(v)) blank = false
    })
    if (blank) continue
    rows.push(obj)
  }

  return { name, headers: headers.filter(Boolean), rows }
}
