import Papa from 'papaparse'
import ExcelJS from 'exceljs'
import { cellText } from './merge.js'

const FILL = {
  new: 'FFD8F0D8',      // green  - value added to an empty cell
  updated: 'FFFDE9B8',  // amber  - value replaced an existing one
  header: 'FFD6E4F7',   // blue   - column added by this merge
}

const paint = (cell, key) => {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: FILL[key] } }
}

function save(blob, filename) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

export function downloadCsv({ headers, rows }, filename) {
  const csv = Papa.unparse({ fields: headers, data: rows.map((r) => headers.map((h) => cellText(r[h]))) })
  save(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' }), filename)
}

/**
 * Base came from XLSX: write the new columns into the original workbook so every
 * existing cell keeps its own value, number format and column width. Nothing in
 * the base is read back and rewritten, so nothing in the base can be mangled.
 */
export function annotateWorkbook({ headers, rows, highlights, addedHeaders, targetHeaders, source }) {
  const ws = source.sheet
  const cols = new Map(
    targetHeaders.filter((h) => headers.includes(h)).map((h) => [h, headers.indexOf(h) + 1]),
  )

  const headerRow = ws.getRow(1)
  for (const h of addedHeaders) {
    const cell = headerRow.getCell(headers.indexOf(h) + 1)
    cell.value = h
    cell.style = { ...headerRow.getCell(1).style }
    paint(cell, 'header')
  }
  headerRow.commit()

  rows.forEach((row, rowIndex) => {
    const xlRow = ws.getRow(source.sheetRowNumbers[rowIndex])
    for (const [h, col] of cols) {
      const text = cellText(row[h])
      const cell = xlRow.getCell(col)
      cell.value = text || null
      cell.numFmt = '@' // text, so 12-digit call IDs never become 1.16574E+11
      // Always clear first: these columns are ours, and the workbook object is
      // reused across downloads, so a stale fill would otherwise linger.
      const kind = highlights.get(`${rowIndex}:${h}`)
      if (kind) paint(cell, kind)
      else cell.fill = undefined
    }
    xlRow.commit()
  })

  return source.workbook
}

/** Base came from CSV: no original workbook to preserve, so build a plain one. */
export function buildWorkbook({ headers, rows, highlights, addedHeaders }) {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Contacts')
  const added = new Set(addedHeaders)

  const headerRow = ws.addRow(headers)
  headers.forEach((h, i) => { if (added.has(h)) paint(headerRow.getCell(i + 1), 'header') })

  rows.forEach((row, rowIndex) => {
    // Everything as text: Excel otherwise turns 12-digit Record IDs into 1.14178E+11.
    const xlRow = ws.addRow(headers.map((h) => cellText(row[h])))
    headers.forEach((h, i) => {
      const kind = highlights.get(`${rowIndex}:${h}`)
      if (kind) paint(xlRow.getCell(i + 1), kind)
    })
  })

  ws.views = [{ state: 'frozen', ySplit: 1 }]
  return wb
}

export async function downloadXlsx(result, filename) {
  const wb = result.source?.workbook ? annotateWorkbook(result) : buildWorkbook(result)
  const buf = await wb.xlsx.writeBuffer()
  save(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename)
}
