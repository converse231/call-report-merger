import Papa from 'papaparse'
import ExcelJS from 'exceljs'
import { cellText } from './merge.js'

/** Reads a .csv or .xlsx into { name, headers, rows }. XLSX rows keep raw cell values. */
export async function readTable(file) {
  const table = /\.xlsx?$/i.test(file.name) ? await readXlsx(file) : await readCsv(file)
  if (!table.headers.length) throw new Error('No column headers found in the first row of this file.')
  if (!table.rows.length) throw new Error('This file has headers but no data rows.')
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

async function readXlsx(file) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(await file.arrayBuffer())
  // HubSpot exports tack on an empty "Sheet1"; take the first one with data.
  const sheet = wb.worksheets.find((w) => w.actualRowCount > 1) ?? wb.worksheets[0]
  if (!sheet) throw new Error('This workbook has no worksheets.')

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

  return { name: file.name, headers: headers.filter(Boolean), rows, workbook: wb, sheet, sheetRowNumbers }
}
