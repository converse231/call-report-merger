import assert from 'node:assert/strict'
import fs from 'node:fs'
import ExcelJS from 'exceljs'
import {
  merge, parseWhen, cellText, targetColumn, defaultCarryColumns,
  mergeNoteHistory, parseNoteBlocks,
  detectCallKeyColumn, detectDateColumn, detectNotesColumn,
} from './merge.js'
import { readTable } from './readTable.js'
import { annotateWorkbook } from './download.js'

/* ---------- cell values ---------- */
assert.equal(cellText(null), '')
assert.equal(cellText(116573637177), '116573637177')            // no scientific notation
assert.equal(cellText(new Date('2026-09-08T15:01:00Z')), '2026-09-08 15:01')
assert.equal(cellText(new Date('1899-12-30T00:00:07Z')), '0:00:07')  // Excel time-serial duration
assert.equal(cellText(new Date('1899-12-30T01:02:03Z')), '1:02:03')
assert.equal(cellText({ richText: [{ text: 'a' }, { text: 'b ' }] }), 'ab')
assert.equal(cellText({ formula: 'A1', result: 42 }), '42')
assert.equal(cellText('  padded\n'), 'padded')

/* ---------- dates ---------- */
assert.equal(parseWhen('8/3/2026 17:06').day, '2026-08-03')
assert.equal(parseWhen('2026-09-04 16:32').day, '2026-09-04')
assert.equal(parseWhen('8/3/2026 5:06 PM').ms, parseWhen('8/3/2026 17:06').ms)
assert.equal(parseWhen(new Date('2026-09-08T15:01:00Z')).day, '2026-09-08')  // xlsx Date, read as UTC
assert.equal(parseWhen('').day, '(no date)')

/* ---------- column detection on the real 94-column export ---------- */
const NEAR_MISSES = [
  'Record ID', 'Activity date', 'Create date', 'Last modified date',
  'Associated Contact create attribution IDs', 'Associated Company IDs',
  'Associated Call IDs', 'Associated Contact IDs', 'Call notes', 'Call summary',
]
assert.equal(detectCallKeyColumn(NEAR_MISSES), 'Associated Contact IDs')
assert.equal(detectCallKeyColumn(['Record ID', 'Activity date']), 'Record ID')
assert.equal(detectDateColumn(NEAR_MISSES), 'Activity date')
assert.equal(detectNotesColumn(NEAR_MISSES), 'Call notes')

/* ---------- column naming is idempotent, so a re-run reuses the same columns ---------- */
assert.equal(targetColumn('Call notes'), 'Call notes')
assert.equal(targetColumn('Activity date'), 'Call Activity date')
assert.equal(targetColumn(targetColumn('Activity date')), 'Call Activity date')

/* ---------- the whitelist: 94 columns in, 8 out ---------- */
const WIDE = [...NEAR_MISSES, 'Activity assigned to', 'Call duration (HH:mm:ss)', 'Call Title',
  'Strategic Status Code', 'To Number', 'Recording URL', 'Owner talk speed']
assert.deepEqual(defaultCarryColumns(WIDE, 'Associated Contact IDs'), [
  'Record ID', 'Activity date', 'Call notes', 'Activity assigned to',
  'Call duration (HH:mm:ss)', 'Call Title', 'Strategic Status Code', 'To Number',
])
// an export missing one standard column just drops it -- no silent fallback to all 94
const NO_STATUS = WIDE.filter((h) => h !== 'Strategic Status Code')
assert.deepEqual(defaultCarryColumns(NO_STATUS, 'Associated Contact IDs'), [
  'Record ID', 'Activity date', 'Call notes', 'Activity assigned to',
  'Call duration (HH:mm:ss)', 'Call Title', 'To Number',
])
// an export with none of the known names falls back to carrying everything
assert.deepEqual(defaultCarryColumns(['id', 'when', 'memo'], 'id'), ['when', 'memo'])

/* ---------- merge behaviour ---------- */
const baseHeaders = ['Record ID', 'First Name']
const baseRows = [
  { 'Record ID': '111', 'First Name': 'Abby' },
  { 'Record ID': 222, 'First Name': 'Sam' },       // numeric id, as XLSX gives it
  { 'Record ID': '333', 'First Name': 'Nobody' },
]
const callHeaders = ['Associated Contact IDs', 'Activity date', 'Call notes', 'Strategic Status Code']
const callRows = [
  { 'Associated Contact IDs': '111', 'Activity date': '8/3/2026 17:06', 'Call notes': 'old day', 'Strategic Status Code': 'Refusal' },
  { 'Associated Contact IDs': '111', 'Activity date': '8/24/2026 12:30', 'Call notes': 'first', 'Strategic Status Code': 'No Answer' },
  { 'Associated Contact IDs': '111', 'Activity date': '8/24/2026 12:34', 'Call notes': 'second', 'Strategic Status Code': '' },
  { 'Associated Contact IDs': '111', 'Activity date': '8/24/2026 13:00', 'Call notes': 'second', 'Strategic Status Code': 'Do Not Call' },
  { 'Associated Contact IDs': '222;999', 'Activity date': '8/5/2026 09:00', 'Call notes': '', 'Strategic Status Code': 'Fax' },
  { 'Associated Contact IDs': '', 'Activity date': '8/5/2026 09:00', 'Call notes': 'orphan', 'Strategic Status Code': '' },
]
const opts = { baseRows, baseHeaders, callRows, callHeaders, baseKey: 'Record ID', callKey: 'Associated Contact IDs', dateCol: 'Activity date', notesCol: 'Call notes' }
const r = merge(opts)

assert.deepEqual(r.addedHeaders, ['Call Activity date', 'Call notes', 'Call Strategic Status Code'])
assert.equal(r.rows[0]['Call notes'], '[2026-08-24]\nfirst\n\nsecond')  // latest day only, deduped, oldest first
assert.equal(r.rows[0]['Call Strategic Status Code'], 'Do Not Call')
assert.equal(r.rows[0]['Call Activity date'], '8/24/2026 13:00')
assert.equal(r.rows[1]['Call Strategic Status Code'], 'Fax')      // semicolon-split ids reach the contact
assert.equal(r.rows[2]['Call notes'], '')                         // unmatched contact untouched
assert.deepEqual(r.stats.unmatched, [{ id: '999', count: 1 }])    // unknown id reported, not invented
assert.equal(r.stats.noContactId, 1)
assert.equal(r.stats.skippedOlderDay, 1)
assert.equal(r.stats.contactsUpdated, 2)
assert.equal(r.rows[0]['First Name'], 'Abby')                     // base columns never touched
assert.equal(r.highlights.get('0:Call notes'), 'new')
assert.equal(r.highlights.get('2:Call notes'), undefined)

// a blank latest call must not wipe a value the same day recorded earlier
const blanked = merge({
  ...opts,
  callRows: [
    { 'Associated Contact IDs': '111', 'Activity date': '8/24/2026 09:00', 'Call notes': 'a', 'Strategic Status Code': 'No Answer' },
    { 'Associated Contact IDs': '111', 'Activity date': '8/24/2026 10:00', 'Call notes': '', 'Strategic Status Code': '' },
  ],
})
assert.equal(blanked.rows[0]['Call Strategic Status Code'], 'No Answer')

// re-running on the merged output changes nothing
const again = merge({ ...opts, baseRows: r.rows, baseHeaders: r.headers })
assert.deepEqual(again.addedHeaders, [])
assert.equal(again.rows[0]['Call notes'], '[2026-08-24]\nfirst\n\nsecond')
assert.equal(again.highlights.size, 0)


/* ---------- notes accumulate across runs, newest day on top ---------- */
const day1 = mergeNoteHistory('', '2026-09-08', 'MON: voicemail\n\nMON: interested')
assert.equal(day1, '[2026-09-08]\nMON: voicemail\n\nMON: interested')

const day2 = mergeNoteHistory(day1, '2026-09-09', 'TUE: booked')
assert.equal(day2, '[2026-09-09]\nTUE: booked\n\n[2026-09-08]\nMON: voicemail\n\nMON: interested')

// an older day back-fills underneath rather than jumping to the top
const back = mergeNoteHistory(day2, '2026-09-01', 'LAST WEEK')
assert.deepEqual(parseNoteBlocks(back).map((b) => b.day), ['2026-09-09', '2026-09-08', '2026-09-01'])

// re-merging a day replaces that day only -- never duplicates it
const redo = mergeNoteHistory(day2, '2026-09-09', 'TUE: booked\n\nTUE: confirmed')
assert.deepEqual(parseNoteBlocks(redo).map((b) => b.day), ['2026-09-09', '2026-09-08'])
assert.match(redo, /TUE: confirmed/)
assert.equal(mergeNoteHistory(day2, '2026-09-09', 'TUE: booked'), day2)  // identical re-run is a no-op

// an empty day never wipes what is already on file
assert.equal(mergeNoteHistory(day2, '2026-09-10', ''), day2)

// notes written before day headers existed are kept, undated, at the bottom
const legacy = mergeNoteHistory('old plain note', '2026-09-09', 'TUE: booked')
assert.equal(legacy, '[2026-09-09]\nTUE: booked\n\n[(no date)]\nold plain note')
assert.deepEqual(parseNoteBlocks(legacy).map((b) => b.day), ['2026-09-09', '(no date)'])
// and it stays a separate block on the next run instead of folding into a day
assert.deepEqual(parseNoteBlocks(mergeNoteHistory(legacy, '2026-09-10', 'WED')).map((b) => b.day),
  ['2026-09-10', '2026-09-09', '(no date)'])

// end to end: two runs, and a contact with no new calls keeps what it had
const twoDayBase = [{ 'Record ID': '111', 'First Name': 'Abby' }, { 'Record ID': '222', 'First Name': 'Sam' }]
const runA = merge({ ...opts, baseRows: twoDayBase, callRows: [
  { 'Associated Contact IDs': '111', 'Activity date': '9/8/2026 09:00', 'Call notes': 'MON: abby' },
  { 'Associated Contact IDs': '222', 'Activity date': '9/8/2026 10:00', 'Call notes': 'MON: sam' },
] })
const runB = merge({ ...opts, baseRows: runA.rows, baseHeaders: runA.headers, callRows: [
  { 'Associated Contact IDs': '111', 'Activity date': '9/9/2026 14:00', 'Call notes': 'TUE: abby' },
] })
assert.equal(runB.rows[0]['Call notes'], '[2026-09-09]\nTUE: abby\n\n[2026-09-08]\nMON: abby')
assert.equal(runB.rows[1]['Call notes'], '[2026-09-08]\nMON: sam')  // no new calls, untouched
assert.equal(runB.highlights.get('0:Call notes'), 'new')  // appended, so green not amber
assert.equal(runB.highlights.get('1:Call notes'), undefined)

/* ---------- XLSX round trip against the real exports ---------- */
const BASE = 'docs/hubspot-CONTACTS.xlsx'
const CALLS = 'docs/hubspot-CALLLS.xlsx'
if (fs.existsSync(BASE) && fs.existsSync(CALLS)) {
  const asFile = (p) => ({ name: p.split('/').pop(), arrayBuffer: async () => fs.readFileSync(p) })
  const base = await readTable(asFile(BASE))
  const calls = await readTable(asFile(CALLS))
  const callKey = detectCallKeyColumn(calls.headers)
  const carry = defaultCarryColumns(calls.headers, callKey)
  assert.equal(carry.length, 8, 'exactly 8 of the 94 call columns are carried')

  const real = merge({
    baseRows: base.rows, baseHeaders: base.headers,
    callRows: calls.rows, callHeaders: calls.headers,
    carryColumns: carry, callKey, baseKey: 'Record ID',
    dateCol: detectDateColumn(calls.headers), notesCol: detectNotesColumn(calls.headers),
  })
  assert.equal(real.addedHeaders.length, 8)
  assert.equal(real.headers.length, base.headers.length + 8)
  assert.ok(real.stats.contactsUpdated > 0, 'the real files must actually match on something')

  const out = 'src/.merge-roundtrip.tmp.xlsx'
  await annotateWorkbook({ ...real, source: base }).xlsx.writeFile(out)

  const orig = new ExcelJS.Workbook(); await orig.xlsx.readFile(BASE)
  const back = new ExcelJS.Workbook(); await back.xlsx.readFile(out)
  const ows = orig.worksheets.find((w) => w.actualRowCount > 1)
  const nws = back.worksheets.find((w) => w.actualRowCount > 1)

  // every original cell keeps its value, its number format and its lack of colour
  for (let rn = 1; rn <= ows.rowCount; rn++) {
    for (let c = 1; c <= base.headers.length; c++) {
      const a = ows.getRow(rn).getCell(c), b = nws.getRow(rn).getCell(c)
      assert.deepEqual(b.value, a.value, `base value changed at r${rn}c${c}`)
      assert.equal(b.numFmt, a.numFmt, `base number format changed at r${rn}c${c}`)
      assert.equal(b.fill?.fgColor?.argb, undefined, `base cell coloured at r${rn}c${c}`)
    }
  }

  const hdr = []
  for (let c = 1; c <= nws.columnCount; c++) hdr.push(cellText(nws.getRow(1).getCell(c).value))
  assert.deepEqual(hdr.slice(base.headers.length), real.addedHeaders)
  const idCol = hdr.indexOf('Call Record ID') + 1
  const anyId = real.preview.map((p) => nws.getRow(base.sheetRowNumbers[p.rowIndex]).getCell(idCol).value)
  assert.ok(anyId.every((v) => /^\d{12}$/.test(String(v))), 'call Record IDs keep all 12 digits')

  const merged = await readTable(asFile(out))
  const rerun = merge({
    baseRows: merged.rows, baseHeaders: merged.headers,
    callRows: calls.rows, callHeaders: calls.headers,
    carryColumns: carry, callKey, baseKey: 'Record ID',
    dateCol: detectDateColumn(calls.headers), notesCol: detectNotesColumn(calls.headers),
  })
  assert.equal(rerun.addedHeaders.length, 0, 'a second merge must not duplicate columns')
  assert.equal(rerun.highlights.size, 0, 'a second merge must not change anything')
  fs.unlinkSync(out)
  console.log(`xlsx round trip: ${base.rows.length} base rows preserved, ${real.stats.contactsUpdated} contacts updated`)
} else {
  console.log('xlsx round trip: skipped (sample exports not in docs/)')
}

console.log('all checks passed')
