import assert from 'node:assert/strict'
import fs from 'node:fs'
import ExcelJS from 'exceljs'
import * as XLSX from 'xlsx'
import {
  merge, parseWhen, cellText, targetColumn, defaultCarryColumns,
  mergeNoteHistory, parseNoteBlocks, wasTrimmed,
  detectBaseKeyColumn, detectCallKeyColumn, detectDateColumn, detectNotesColumn,
} from './merge.js'
import { readTable } from './readTable.js'
import { annotateWorkbook, buildWorkbook } from './download.js'

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
// every day present in this export is folded into notes, not just the latest --
// this is the bug the user reported: 8/3's notes must not be dropped
assert.equal(r.rows[0]['Call notes'], '[2026-08-24]\nfirst\n\nsecond\n\n[2026-08-03]\nold day')
assert.equal(r.rows[0]['Call Strategic Status Code'], 'Do Not Call')  // non-notes columns: latest day only
assert.equal(r.rows[0]['Call Activity date'], '8/24/2026 13:00')
assert.equal(r.rows[1]['Call Strategic Status Code'], 'Fax')      // semicolon-split ids reach the contact
assert.equal(r.rows[2]['Call notes'], '')                         // unmatched contact untouched
assert.deepEqual(r.stats.unmatched, [{ id: '999', count: 1 }])    // unknown id reported, not invented
assert.equal(r.stats.noContactId, 1)
assert.equal(r.stats.historyDaysMerged, 2)                        // Abby's 8/3 + 8/24 (Sam's 8/5 call has no notes)
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
assert.equal(again.rows[0]['Call notes'], '[2026-08-24]\nfirst\n\nsecond\n\n[2026-08-03]\nold day')
assert.equal(again.highlights.size, 0)

// the exact bug report: one contact, five distinct call days in a single export --
// all five must survive, only the latest feeds the non-notes columns
const fiveDays = merge({
  baseHeaders, baseKey: 'Record ID', callHeaders, callKey: 'Associated Contact IDs', dateCol: 'Activity date', notesCol: 'Call notes',
  baseRows: [{ 'Record ID': '111', 'First Name': 'Abby' }],
  callRows: [
    { 'Associated Contact IDs': '111', 'Activity date': '8/3/2026 17:06', 'Call notes': 'Voicemail', 'Strategic Status Code': '' },
    { 'Associated Contact IDs': '111', 'Activity date': '8/4/2026 16:38', 'Call notes': 'Call back requested', 'Strategic Status Code': '' },
    { 'Associated Contact IDs': '111', 'Activity date': '8/5/2026 12:37', 'Call notes': 'Left message again', 'Strategic Status Code': '' },
    { 'Associated Contact IDs': '111', 'Activity date': '8/19/2026 13:01', 'Call notes': 'test test', 'Strategic Status Code': '' },
    { 'Associated Contact IDs': '111', 'Activity date': '8/24/2026 12:34', 'Call notes': 'just a test', 'Strategic Status Code': 'Scheduled Callback - AC' },
  ],
})
assert.deepEqual(
  fiveDays.rows[0]['Call notes'].match(/^\[[\d-]+\]/gm),
  ['[2026-08-24]', '[2026-08-19]', '[2026-08-05]', '[2026-08-04]', '[2026-08-03]'],
  'all five call days must appear, newest first',
)
assert.match(fiveDays.rows[0]['Call notes'], /Voicemail/)         // 8/3 -- the note the bug report said vanished
assert.equal(fiveDays.rows[0]['Call Strategic Status Code'], 'Scheduled Callback - AC')  // still latest-day only
assert.equal(fiveDays.stats.historyDaysMerged, 5)


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

/* ---------- legacy .xls input (BIFF8/OLE2, not zip-based) ---------- */
{
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([]), 'Sheet1') // HubSpot-style empty leading sheet
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Record ID', 'First Name'],
    [111, 'Abby'],
    [222, 'Sam'],
  ])
  XLSX.utils.book_append_sheet(wb, sheet, 'Contacts')
  const xlsBuf = XLSX.write(wb, { type: 'buffer', bookType: 'biff8' })
  assert.deepEqual([...xlsBuf.subarray(0, 4)], [0xd0, 0xcf, 0x11, 0xe0], 'fixture really is BIFF8, not zip')

  const asFile = (name, buf) => ({ name, arrayBuffer: async () => buf })
  const legacyBase = await readTable(asFile('legacy.xls', xlsBuf))
  assert.deepEqual(legacyBase.headers, ['Record ID', 'First Name'])
  assert.equal(legacyBase.rows.length, 2)
  assert.equal(legacyBase.workbook, undefined, 'no live workbook to preserve -- output rebuilds fresh')

  const legacyMerge = merge({
    baseRows: legacyBase.rows, baseHeaders: legacyBase.headers,
    callRows: [{ 'Associated Contact IDs': '111', 'Activity date': '9/8/2026 09:00', 'Call notes': 'hi' }],
    callHeaders: ['Associated Contact IDs', 'Activity date', 'Call notes'],
    baseKey: 'Record ID', callKey: 'Associated Contact IDs', dateCol: 'Activity date', notesCol: 'Call notes',
  })
  assert.equal(legacyMerge.rows[0]['Call notes'], '[2026-09-08]\nhi')
  const outWb = buildWorkbook(legacyMerge)
  assert.ok((await outWb.xlsx.writeBuffer()).byteLength > 0, 'a fresh xlsx must still be produced from a legacy base')

  // a file's real signature wins over a misleading extension in either direction
  const xlsxBuf = await (async () => {
    const w = new ExcelJS.Workbook()
    const s = w.addWorksheet('S')
    s.addRow(['a']).commit()
    s.addRow(['b']).commit()
    return w.xlsx.writeBuffer()
  })()
  assert.ok(new Uint8Array(xlsxBuf).subarray(0, 2).join() === [0x50, 0x4b].join(), 'fixture really is zip-based')
  const mislabeled = await readTable(asFile('really-xlsx.xls', xlsxBuf))
  assert.ok(mislabeled.workbook, 'zip content routes through ExcelJS even with a .xls name')

  await assert.rejects(
    readTable(asFile('not-a-spreadsheet.xls', new TextEncoder().encode('hello,world'))),
    /doesn't look like an Excel file/,
  )
  console.log('legacy .xls input: parses, merges, and signature-sniffing works')
}


/* ---------- QA regressions ---------- */

// unticking the notes chip must leave notes out entirely, not write to a
// column that isn't in the output
const noNotes = merge({ ...opts, carryColumns: ['Strategic Status Code'] })
assert.deepEqual(noNotes.addedHeaders, ['Call Strategic Status Code'])
assert.equal(noNotes.rows[0]['Call notes'], undefined)
assert.equal(noNotes.preview[0].newNotes, '')

// pointing the notes dropdown at the join key must not invent a stray column
const notesIsKey = merge({ ...opts, notesCol: 'Associated Contact IDs' })
assert.deepEqual(
  Object.keys(notesIsKey.rows[0]).filter((k) => !notesIsKey.headers.includes(k)), [],
  'no value may be written to a column missing from headers',
)

// chaining a second calls file onto the first merge keeps the first file's history
const chainA = merge({
  ...opts, baseRows: [{ 'Record ID': '111', 'First Name': 'Abby' }],
  callRows: [{ 'Associated Contact IDs': '111', 'Activity date': '9/8/2026 09:00', 'Call notes': 'MON' }],
})
const chainB = merge({
  ...opts, baseRows: chainA.rows, baseHeaders: chainA.headers,
  callRows: [{ 'Associated Contact IDs': '111', 'Activity date': '9/9/2026 09:00', 'Call notes': 'TUE' }],
})
assert.equal(chainB.rows[0]['Call notes'], '[2026-09-09]\nTUE\n\n[2026-09-08]\nMON')

// stats must not claim updates when nothing changed
assert.equal(again.stats.contactsUpdated, 0)
assert.equal(again.stats.historyDaysMerged, 0)
assert.equal(again.stats.contactsAlreadyCurrent, 2)
assert.ok(again.preview.every((p) => !p.changed))
assert.equal(r.preview.filter((p) => p.changed).length, r.stats.contactsUpdated)

// Excel writes 2-digit years when a column is formatted m/d/yy
assert.equal(parseWhen('8/3/26 17:06').day, '2026-08-03')
assert.equal(parseWhen('8/3/99').day, '1999-08-03')

// duplicate contact IDs in the base are reported rather than silently half-applied
const dup = merge({
  ...opts, baseRows: [{ 'Record ID': '111', 'First Name': 'A' }, { 'Record ID': '111', 'First Name': 'B' }],
  callRows: [{ 'Associated Contact IDs': '111', 'Activity date': '9/8/2026', 'Call notes': 'x' }],
})
assert.deepEqual(dup.stats.duplicateBaseIds, [{ id: '111', count: 2 }])

// Excel-mangled IDs get named, not just reported as zero matches
const mangled = merge({
  ...opts, baseRows: [{ 'Record ID': '1.14178E+11', 'First Name': 'A' }],
  callRows: [{ 'Associated Contact IDs': '1.14203E+11', 'Activity date': '9/8/2026', 'Call notes': 'x' }],
})
assert.equal(mangled.stats.mangledCallIds, 1)
assert.equal(mangled.stats.mangledBaseIds, 1)
assert.equal(mangled.stats.contactsUpdated, 0)
assert.equal(r.stats.mangledCallIds, 0, 'a normal 12-digit id is not mistaken for a mangled one')

// Excel rejects a workbook with a cell over 32,767 characters
{
  let cell = ''
  const d = new Date(Date.UTC(2025, 0, 1))
  for (let i = 0; i < 400; i++) {
    cell = mergeNoteHistory(cell, d.toISOString().slice(0, 10), ('Outbound answered call, Call ID: 402' + i + ' ').repeat(4))
    d.setUTCDate(d.getUTCDate() + 1)
  }
  assert.ok(cell.length <= 32767, 'notes cell must fit Excel limit, got ' + cell.length)
  assert.ok(wasTrimmed(cell))
  assert.ok(cell.startsWith('[2026-'), 'the newest days survive trimming')
  // the notice must not stack up run after run
  const more = mergeNoteHistory(cell, '2026-02-05', 'another call')
  assert.equal((more.match(/older notes trimmed/g) ?? []).length, 1)
  assert.ok(more.length <= 32767)
  // a single day too big on its own gets its text cut instead of vanishing
  const oneHuge = mergeNoteHistory('', '2026-01-01', 'y'.repeat(40000))
  assert.ok(oneHuge.length <= 32767 && wasTrimmed(oneHuge))
  // and merge() surfaces it so the UI can warn
  const trimmedMerge = merge({
    ...opts, baseRows: [{ 'Record ID': '111', 'First Name': 'A', 'Call notes': cell }],
    baseHeaders: ['Record ID', 'First Name', 'Call notes'],
    callRows: [{ 'Associated Contact IDs': '111', 'Activity date': '3/1/2026', 'Call notes': 'z' }],
  })
  assert.equal(trimmedMerge.stats.notesTrimmed, 1)
}

// preview ordering is stable for contacts sharing a day
{
  const sameDay = merge({
    ...opts,
    baseRows: [{ 'Record ID': '1' }, { 'Record ID': '2' }, { 'Record ID': '3' }],
    baseHeaders: ['Record ID'],
    callRows: ['1', '2', '3'].map((id) => ({ 'Associated Contact IDs': id, 'Activity date': '9/8/2026', 'Call notes': 'n' })),
  })
  assert.deepEqual(sameDay.preview.map((p) => p.id), ['1', '2', '3'])
}

/* ---------- XLSX round trip (synthetic, so it runs anywhere) ---------- */
{
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Contacts')
  ws.addRow(['Record ID', 'First Name', 'Last Activity Date']).commit()
  const dated = ws.addRow([246317725467, 'Randy', new Date(Date.UTC(2026, 8, 4, 16, 32))])
  dated.getCell(3).numFmt = 'm/d/yy "h":mm'
  dated.commit()
  ws.addRow([238297728268, 'Abby', null]).commit()
  const baseBuf = await wb.xlsx.writeBuffer()

  const rtBase = await readTable({ name: 'base.xlsx', arrayBuffer: async () => baseBuf })
  assert.ok(rtBase.workbook, '.xlsx keeps a live workbook so base cells can be preserved')

  const rt = merge({
    baseRows: rtBase.rows, baseHeaders: rtBase.headers,
    callHeaders: ['Associated Contact IDs', 'Activity date', 'Call notes'],
    callRows: [{ 'Associated Contact IDs': '238297728268', 'Activity date': '9/8/2026 15:01', 'Call notes': 'hello' }],
    baseKey: 'Record ID', callKey: 'Associated Contact IDs', dateCol: 'Activity date', notesCol: 'Call notes',
  })
  const out = 'src/.roundtrip.tmp.xlsx'
  await annotateWorkbook({ ...rt, source: rtBase }).xlsx.writeFile(out)

  const orig = new ExcelJS.Workbook(); await orig.xlsx.readFile(out)
  const ows = orig.worksheets[0]
  // untouched base cells keep their value, number format and lack of fill
  assert.equal(ows.getRow(2).getCell(1).value, 246317725467)
  assert.equal(ows.getRow(2).getCell(3).numFmt, 'm/d/yy "h":mm')
  assert.deepEqual(ows.getRow(2).getCell(3).value, new Date(Date.UTC(2026, 8, 4, 16, 32)))
  assert.equal(ows.getRow(2).getCell(1).fill?.fgColor?.argb, undefined)
  // the merged row is written as text and highlighted
  const notesCol = rt.headers.indexOf('Call notes') + 1
  assert.equal(ows.getRow(3).getCell(notesCol).value, '[2026-09-08]\nhello')
  assert.equal(ows.getRow(3).getCell(notesCol).fill?.fgColor?.argb, 'FFD8F0D8')
  assert.equal(ows.getRow(2).getCell(notesCol).value, null, 'a contact with no calls stays empty')

  // a stale fill from an earlier download must not survive a later one
  await annotateWorkbook({ ...rt, highlights: new Map(), source: rtBase }).xlsx.writeFile(out)
  const repaint = new ExcelJS.Workbook(); await repaint.xlsx.readFile(out)
  assert.equal(
    repaint.worksheets[0].getRow(3).getCell(notesCol).fill?.fgColor?.argb, undefined,
    'download must repaint from scratch, not accumulate stale colours',
  )

  fs.unlinkSync(out)
  console.log('xlsx round trip: base cells preserved, highlights repaint cleanly')
}

console.log('all checks passed')
