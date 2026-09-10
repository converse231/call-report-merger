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
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ ,]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?)?/)
    if (!m) return { ms: 0, day: UNKNOWN_DAY }
    ;[, mo, d, y, h = 0, mi = 0, s = 0, ampm] = m
  }

  y = +y; mo = +mo; d = +d; h = +h; mi = +mi; s = +s
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
  const blocks = parseNoteBlocks(existing).map((b) => ({ ...b, day: b.day ?? UNKNOWN_DAY }))
  if (!dayText) return formatNoteBlocks(blocks) // an empty day must never wipe history
  const out = blocks.filter((b) => b.day !== day)
  out.push({ day, text: dayText })
  // Descending, which also drops "(no date)" to the bottom: '(' sorts below '0'.
  out.sort((a, b) => (a.day < b.day ? 1 : a.day > b.day ? -1 : 0))
  return formatNoteBlocks(out)
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

  // contact id -> row index (last wins if the base has duplicates)
  const index = new Map()
  baseRows.forEach((r, i) => {
    const id = normId(r[baseKey])
    if (id) index.set(id, i)
  })

  // Bucket calls by contact, keeping only that contact's latest day.
  const byContact = new Map()
  const unmatched = new Map()
  let noContactId = 0
  let usedCalls = 0
  let skippedOlderDay = 0

  for (const call of callRows) {
    const ids = normId(call[callKey]).split(';').map(normId).filter(Boolean)
    if (!ids.length) { noContactId++; continue }
    const when = parseWhen(call[dateCol])
    for (const id of ids) {
      if (!index.has(id)) {
        unmatched.set(id, (unmatched.get(id) ?? 0) + 1)
        continue
      }
      let b = byContact.get(id)
      if (b && when.day > b.day) { skippedOlderDay += b.calls.length; b = null }
      if (!b) b = { day: when.day, calls: [] }
      if (when.day < b.day) { skippedOlderDay++; continue }
      b.calls.push({ call, when })
      byContact.set(id, b)
    }
  }

  const rows = baseRows.map((r) => ({ ...r }))
  const highlights = new Map() // `${rowIndex}:${header}` -> 'new' | 'updated'
  const preview = []

  for (const [id, bucket] of byContact) {
    const rowIndex = index.get(id)
    const row = rows[rowIndex]
    const calls = bucket.calls.sort((a, b) => a.when.ms - b.when.ms) // oldest -> newest
    usedCalls += calls.length

    const values = {}
    const kindHint = {}
    let newNotes = ''

    if (notesCol) {
      const seen = new Set()
      const notes = []
      for (const { call } of calls) {
        const n = norm(call[notesCol])
        if (n && !seen.has(n)) { seen.add(n); notes.push(n) }
      }
      newNotes = notes.join('\n\n')
      const target = targetColumn(notesCol)
      // Appending a day is new data, not a replacement -- only re-merging a day
      // that is already on file overwrites anything.
      kindHint[target] = parseNoteBlocks(row[target]).some((b) => b.day === bucket.day) ? 'updated' : 'new'
      values[target] = mergeNoteHistory(row[target], bucket.day, newNotes)
    }

    for (const h of carried) {
      if (h === notesCol) continue
      // latest non-empty: "latest wins", without blanking a field the last call left empty
      let v = ''
      for (let i = calls.length - 1; i >= 0; i--) {
        const c = norm(calls[i].call[h])
        if (c) { v = c; break }
      }
      values[targetColumn(h)] = v
    }

    for (const [header, value] of Object.entries(values)) {
      if (!value) continue
      const before = norm(row[header])
      if (before === value) continue
      highlights.set(`${rowIndex}:${header}`, kindHint[header] ?? (before ? 'updated' : 'new'))
      row[header] = value
    }

    preview.push({ rowIndex, id, day: bucket.day, callCount: calls.length, values, newNotes })
  }

  for (const row of rows) for (const h of newHeaders) if (row[h] == null) row[h] = ''

  return {
    headers: newHeaders,
    addedHeaders,
    targetHeaders: carried.map(targetColumn),
    rows,
    highlights,
    preview: preview.sort((a, b) => (a.day < b.day ? 1 : -1)),
    stats: {
      baseRows: baseRows.length,
      callRows: callRows.length,
      contactsUpdated: preview.length,
      usedCalls,
      skippedOlderDay,
      noContactId,
      unmatched: [...unmatched.entries()].map(([id, count]) => ({ id, count })),
    },
  }
}
