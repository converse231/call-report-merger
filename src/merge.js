// Pure merge logic. No DOM, no React -- so it stays testable with `npm test`.

export const UNKNOWN_DAY = '(no date)'

/**
 * The only call columns we carry over, in the order HubSpot exports them.
 * The calls export has ~94 columns; these 8 are the report.
 */
export const DEFAULT_CALL_COLUMNS = [
  'Record ID',
  'Activity assigned to',
  'Activity date',
  'Call duration (HH:mm:ss)',
  'Call notes',
  'Call Title',
  'Strategic Status Code',
  'To Number',
]

const pad = (n) => String(n).padStart(2, '0')

/**
 * Cells arrive as strings from CSV but as numbers, Dates and rich-text objects
 * from XLSX. Everything downstream compares and writes text, so normalise here.
 */
export function cellText(v) {
  if (v == null) return ''
  if (v instanceof Date) {
    // Excel stores a time-only value (a call duration) as a Date on day zero.
    if (v.getUTCFullYear() < 1900) {
      const s = Math.round(v.getTime() / 1000 - Date.UTC(1899, 11, 30) / 1000)
      return `${Math.floor(s / 3600)}:${pad(Math.floor(s / 60) % 60)}:${pad(s % 60)}`
    }
    return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())} ${pad(v.getUTCHours())}:${pad(v.getUTCMinutes())}`
  }
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('').trim()
    if ('result' in v) return cellText(v.result)
    if ('text' in v) return cellText(v.text)
    if ('error' in v) return ''
    return ''
  }
  return String(v).trim()
}

const norm = cellText

/** A 12-digit ID that Excel has rewritten as "1.14178E+11". */
const MANGLED_ID = /^\d+(\.\d+)?[eE][+-]?\d+$/

/** HubSpot IDs sometimes arrive as "238297728268.0" or with stray spaces. */
const normId = (v) => norm(v).replace(/\.0+$/, '')

/**
 * Parse the date formats HubSpot exports: "8/3/2026 17:06", "8/3/2026 5:06 PM",
 * "2026-09-04 16:32". Returns { ms, day } where day is a sortable YYYY-MM-DD.
 */
export function parseWhen(value) {
  // XLSX gives a real Date; exceljs decodes the sheet's wall clock into UTC.
  if (value instanceof Date && value.getUTCFullYear() >= 1900) {
    return {
      ms: value.getTime(),
      day: `${value.getUTCFullYear()}-${pad(value.getUTCMonth() + 1)}-${pad(value.getUTCDate())}`,
    }
  }
  const t = norm(value)
  if (!t) return { ms: 0, day: UNKNOWN_DAY }

  let y, mo, d, h = 0, mi = 0, s = 0, ampm = null
  let m = t.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/)
  if (m) {
    ;[, y, mo, d, h = 0, mi = 0, s = 0] = m
  } else {
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:[ ,]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?)?/)
    if (!m) return { ms: 0, day: UNKNOWN_DAY }
    ;[, mo, d, y, h = 0, mi = 0, s = 0, ampm] = m
  }

  y = +y; mo = +mo; d = +d; h = +h; mi = +mi; s = +s
  // Excel writes "8/3/26" when the column is formatted m/d/yy -- same window Excel uses.
  if (y < 100) y += y < 70 ? 2000 : 1900
  if (ampm) {
    const pm = ampm.toLowerCase() === 'pm'
    if (pm && h < 12) h += 12
    if (!pm && h === 12) h = 0
  }
  return {
    ms: new Date(y, mo - 1, d, h, mi, s).getTime(),
    day: `${y}-${pad(mo)}-${pad(d)}`,
  }
}

/**
 * Column in the calls file that holds the contact's Record ID.
 * The 94-column export has several near-misses ("Associated Contact create
 * attribution IDs"), so prefer an exact name, then the shortest plausible one.
 */
export function detectCallKeyColumn(callHeaders) {
  const score = (h) => {
    const l = h.trim().toLowerCase()
    if (l === 'associated contact ids') return 100
    if (l === 'associated contact id' || l === 'contact id' || l === 'contact ids') return 90
    if (l.includes('contact') && /ids?$/.test(l)) return 50 - l.length
    if (l === 'record id') return 10
    if (/ids?$/.test(l)) return 1
    return 0
  }
  return [...callHeaders].sort((a, b) => score(b) - score(a))[0] ?? ''
}

export function detectBaseKeyColumn(baseHeaders) {
  return baseHeaders.find((h) => h.toLowerCase() === 'record id') ?? baseHeaders[0] ?? ''
}

/** Where a call column lands in the base file. Stable, so re-running overwrites in place. */
export function targetColumn(callHeader) {
  return /^call[\s:]/i.test(callHeader) ? callHeader : `Call ${callHeader}`
}

export function detectDateColumn(callHeaders) {
  return (
    callHeaders.find((h) => /activity date|call date|^date$|timestamp/i.test(h)) ??
    callHeaders.find((h) => /date/i.test(h)) ??
    ''
  )
}

export function detectNotesColumn(callHeaders) {
  return callHeaders.find((h) => /note/i.test(h)) ?? ''
}

/** The default 8, minus anything this particular export doesn't have. */
export function defaultCarryColumns(callHeaders, callKey) {
  const picked = callHeaders.filter((h) => h !== callKey && DEFAULT_CALL_COLUMNS.includes(h))
  // Unrecognised export (renamed headers): fall back to carrying everything.
  return picked.length ? picked : callHeaders.filter((h) => h !== callKey)
}

const DAY_HEADER = /^\[(\d{4}-\d{2}-\d{2}|\(no date\))\]$/

/** Split an accumulated notes cell back into its per-day blocks. */
export function parseNoteBlocks(text) {
  const t = cellText(text)
  if (!t) return []
  const blocks = []
  let cur = null
  for (const line of t.split(/\r?\n/)) {
    if (DAY_HEADER.test(line.trim())) {
      cur = { day: line.trim().slice(1, -1), lines: [] }
      blocks.push(cur)
      continue
    }
    // ponytail: notes written before day headers existed keep their text but no
    // date -- they sink to the bottom. Date them from Call Activity date if needed.
    if (!cur) { cur = { day: null, lines: [] }; blocks.push(cur) }
    cur.lines.push(line)
  }
  return blocks
    .map((b) => ({ day: b.day, text: b.lines.join('\n').trim() }))
    .filter((b) => b.text)
}

export function formatNoteBlocks(blocks) {
  return blocks.map((b) => (b.day ? `[${b.day}]\n${b.text}` : b.text)).join('\n\n')
}

/**
 * Fold one day's notes into a cell's existing history, newest day first.
 * Re-merging a day replaces that day's block instead of duplicating it, so
 * running the same export twice is a no-op.
 */
export function mergeNoteHistory(existing, day, dayText) {
  // Undated legacy text gets its own marker, otherwise sitting below a dated
  // header would fold it into that day the next time the cell is parsed.
  const blocks = parseNoteBlocks(existing)
    .filter((b) => !b.text.startsWith(TRIM_PREFIX)) // drop a previous run's trim notice
    .map((b) => ({ ...b, day: b.day ?? UNKNOWN_DAY }))
  if (!dayText) return fitToCell(blocks) // an empty day must never wipe history
  const out = blocks.filter((b) => b.day !== day)
  out.push({ day, text: dayText })
  // Descending, which also drops "(no date)" to the bottom: '(' sorts below '0'.
  out.sort((a, b) => (a.day < b.day ? 1 : a.day > b.day ? -1 : 0))
  return fitToCell(out)
}

/** Excel refuses to open a file with a cell over 32,767 characters. */
export const MAX_NOTE_CHARS = 32000
const TRIM_PREFIX = '[… '
// No count in the notice: trimming happens a bit at a time across many merges,
// so any number here would describe one run, not the total actually dropped.
export const TRIM_NOTICE = `${TRIM_PREFIX}older notes trimmed to fit Excel's 32,767-character cell limit]`

/** True if a notes cell has had history dropped to fit Excel's cell limit. */
export const wasTrimmed = (text) => cellText(text).includes(TRIM_NOTICE)

/** Drop the oldest days until the cell fits, leaving a note saying so. */
function fitToCell(blocks) {
  let out = formatNoteBlocks(blocks)
  if (out.length <= MAX_NOTE_CHARS) return out

  const kept = [...blocks]
  while (kept.length > 1) {
    kept.pop() // blocks are newest-first, so the oldest goes
    out = formatNoteBlocks([...kept, { day: null, text: TRIM_NOTICE }])
    if (out.length <= MAX_NOTE_CHARS) return out
  }

  // A single day's notes are over the limit on their own: cut the text itself.
  const only = kept[0]
  const marker = `\n${TRIM_NOTICE}`
  const room = MAX_NOTE_CHARS - marker.length - (only.day ? only.day.length + 4 : 0)
  return formatNoteBlocks([{ day: only.day, text: only.text.slice(0, Math.max(0, room)) + marker }])
}

/**
 * For each contact: take only their most recent call day, merge that day's notes
 * into the notes history, and fill every other call column with the latest
 * non-empty value from that day.
 */
export function merge({ baseRows, baseHeaders, callRows, callHeaders, baseKey, callKey, dateCol, notesCol, carryColumns }) {
  const wanted = new Set(carryColumns ?? defaultCarryColumns(callHeaders, callKey))
  const carried = callHeaders.filter((h) => h !== callKey && wanted.has(h))
  const newHeaders = [...baseHeaders]
  for (const h of carried) {
    const t = targetColumn(h)
    if (!newHeaders.includes(t)) newHeaders.push(t)
  }
  const addedHeaders = newHeaders.filter((h) => !baseHeaders.includes(h))

  // The notes column only accumulates if it is actually being carried -- untick
  // its chip and it must be left out entirely, not written to a column that
  // isn't in the output.
  const notesActive = Boolean(notesCol) && carried.includes(notesCol)

  // contact id -> row index (last wins if the base has duplicates)
  const index = new Map()
  const duplicateIds = new Map()
  let mangledBaseIds = 0
  baseRows.forEach((r, i) => {
    const id = normId(r[baseKey])
    if (!id) return
    if (index.has(id)) duplicateIds.set(id, (duplicateIds.get(id) ?? 1) + 1)
    if (MANGLED_ID.test(id)) mangledBaseIds++
    index.set(id, i)
  })

  // Bucket calls by contact, then by day. A calls export is typically a full
  // history export, not just "today" -- every day it contains gets folded into
  // the notes history; only the contact's single latest day feeds the other
  // (non-history) columns like status/duration.
  const byContact = new Map() // id -> Map<day, {call, when}[]>
  const unmatched = new Map()
  let noContactId = 0
  let usedCalls = 0

  for (const call of callRows) {
    const ids = normId(call[callKey]).split(';').map(normId).filter(Boolean)
    if (!ids.length) { noContactId++; continue }
    const when = parseWhen(call[dateCol])
    for (const id of ids) {
      if (!index.has(id)) {
        unmatched.set(id, (unmatched.get(id) ?? 0) + 1)
        continue
      }
      let days = byContact.get(id)
      if (!days) { days = new Map(); byContact.set(id, days) }
      let dayCalls = days.get(when.day)
      if (!dayCalls) { dayCalls = []; days.set(when.day, dayCalls) }
      dayCalls.push({ call, when })
      usedCalls++
    }
  }

  const unmatchedList = [...unmatched.entries()].map(([id, count]) => ({ id, count }))
  const rows = baseRows.map((r) => ({ ...r }))
  const highlights = new Map() // `${rowIndex}:${header}` -> 'new' | 'updated'
  const preview = []
  let historyDaysMerged = 0
  let notesTrimmed = 0

  for (const [id, days] of byContact) {
    const rowIndex = index.get(id)
    const row = rows[rowIndex]
    const allDays = [...days.keys()].sort() // ascending
    const latestDay = allDays[allDays.length - 1]
    const latestCalls = days.get(latestDay).sort((a, b) => a.when.ms - b.when.ms)

    const values = {}
    const kindHint = {}
    const addedThisRun = [] // {day, text} for the preview -- only what actually changed

    if (notesActive) {
      const target = targetColumn(notesCol)
      let existing = row[target]
      let anyNew = false
      let anyUpdated = false

      for (const day of allDays) {
        const dayCalls = days.get(day).sort((a, b) => a.when.ms - b.when.ms)
        const seen = new Set()
        const notes = []
        for (const { call } of dayCalls) {
          const n = norm(call[notesCol])
          if (n && !seen.has(n)) { seen.add(n); notes.push(n) }
        }
        const dayText = notes.join('\n\n')
        if (!dayText) continue
        const alreadyHadDay = parseNoteBlocks(existing).some((b) => b.day === day)
        const before = norm(existing)
        existing = mergeNoteHistory(existing, day, dayText)
        if (existing === before) continue // already on file verbatim -- not a change
        if (alreadyHadDay) anyUpdated = true; else anyNew = true
        addedThisRun.push({ day, text: dayText })
        historyDaysMerged++
      }

      values[target] = existing
      if (wasTrimmed(existing)) notesTrimmed++
      // Appending a day is new data, not a replacement -- only re-merging a day
      // that is already on file overwrites anything.
      if (anyNew) kindHint[target] = 'new'
      else if (anyUpdated) kindHint[target] = 'updated'
    }

    for (const h of carried) {
      if (notesActive && h === notesCol) continue
      // latest non-empty: "latest wins", without blanking a field the last call left empty
      let v = ''
      for (let i = latestCalls.length - 1; i >= 0; i--) {
        const c = norm(latestCalls[i].call[h])
        if (c) { v = c; break }
      }
      values[targetColumn(h)] = v
    }

    let changed = false
    for (const [header, value] of Object.entries(values)) {
      if (!value) continue
      const before = norm(row[header])
      if (before === value) continue
      highlights.set(`${rowIndex}:${header}`, kindHint[header] ?? (before ? 'updated' : 'new'))
      row[header] = value
      changed = true
    }

    const newNotes = formatNoteBlocks(
      [...addedThisRun].sort((a, b) => (a.day < b.day ? 1 : a.day > b.day ? -1 : 0)),
    )
    preview.push({ rowIndex, id, day: latestDay, callCount: latestCalls.length, daysMerged: allDays.length, values, newNotes, changed })
  }

  for (const row of rows) for (const h of newHeaders) if (row[h] == null) row[h] = ''

  return {
    headers: newHeaders,
    addedHeaders,
    targetHeaders: carried.map(targetColumn),
    rows,
    highlights,
    preview: preview.sort((a, b) => (a.day < b.day ? 1 : a.day > b.day ? -1 : 0)),
    stats: {
      baseRows: baseRows.length,
      callRows: callRows.length,
      contactsUpdated: preview.filter((p) => p.changed).length,
      contactsAlreadyCurrent: preview.filter((p) => !p.changed).length,
      usedCalls,
      historyDaysMerged,
      noContactId,
      unmatched: unmatchedList,
      // Excel turns a 12-digit ID into 1.14178E+11 the moment a CSV is opened and
      // re-saved. That silently matches nothing, so name it rather than report zero.
      mangledCallIds: unmatchedList.filter((u) => MANGLED_ID.test(u.id)).length,
      mangledBaseIds,
      notesTrimmed,
      duplicateBaseIds: [...duplicateIds.entries()].map(([id, count]) => ({ id, count })),
    },
  }
}
