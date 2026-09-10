import React, { useEffect, useMemo, useRef, useState } from 'react'
import {
  merge, targetColumn, defaultCarryColumns, DEFAULT_CALL_COLUMNS,
  detectBaseKeyColumn, detectCallKeyColumn, detectDateColumn, detectNotesColumn,
} from './merge.js'
import { readTable } from './readTable.js'
import { downloadCsv, downloadXlsx } from './download.js'

function Drop({ label, hint, file, onFile, error, busy }) {
  const input = useRef(null)
  const [over, setOver] = useState(false)
  const open = () => input.current.click()

  return (
    <div
      className={`drop ${file ? 'ok' : ''} ${over ? 'over' : ''} ${error ? 'bad' : ''}`}
      role="button"
      tabIndex={0}
      aria-label={label}
      onClick={open}
      onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); open() } }}
      onDragOver={(e) => { e.preventDefault(); setOver(true) }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => { e.preventDefault(); setOver(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]) }}
    >
      <input
        ref={input} type="file" accept=".csv,.xlsx,.xls" hidden
        // reset the value so picking the same file again still fires a change
        onChange={(e) => { const f = e.target.files[0]; e.target.value = ''; if (f) onFile(f) }}
      />
      <div className="drop-label">{label}</div>
      {busy ? (
        <div className="drop-meta">Reading…</div>
      ) : file ? (
        <>
          <div className="drop-file">{file.name}</div>
          <div className="drop-meta">{file.rows.length.toLocaleString()} rows &middot; {file.headers.length} columns</div>
          <div className="drop-swap">Click to replace</div>
        </>
      ) : (
        <>
          <div className="drop-hint">{hint}</div>
          <div className="drop-swap">Drop a .xlsx, .xls, or .csv here, or click to browse</div>
        </>
      )}
      {error && <div className="drop-error">{error}</div>}
    </div>
  )
}

function Select({ label, value, options, onChange, note }) {
  return (
    <label className="field">
      <span className="field-label">{label}</span>
      <select value={value} onChange={(e) => onChange(e.target.value)}>
        <option value="">&mdash; none &mdash;</option>
        {options.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
      {note && <span className="field-note">{note}</span>}
    </label>
  )
}

function Stat({ n, label, tone = '', muted }) {
  return (
    <div className={`stat ${tone} ${muted ? 'muted' : ''}`}>
      <div className="stat-n">{n.toLocaleString()}</div>
      <div className="stat-l">{label}</div>
    </div>
  )
}

export default function App() {
  const [base, setBase] = useState(null)
  const [baseFile, setBaseFile] = useState(null)
  const [calls, setCalls] = useState(null)
  const [errors, setErrors] = useState({})
  const [loading, setLoading] = useState({})
  const [map, setMap] = useState({ baseKey: '', callKey: '', dateCol: '', notesCol: '' })
  const [carry, setCarry] = useState([])
  // Everything merged so far this session, so a second calls file builds on the
  // first instead of starting over from the file as uploaded.
  const [work, setWork] = useState(null)
  const [result, setResult] = useState(null)
  const [busy, setBusy] = useState(false)
  const [saving, setSaving] = useState('')

  // A file dropped outside a drop zone would otherwise navigate away and lose everything.
  useEffect(() => {
    const stop = (e) => e.preventDefault()
    window.addEventListener('dragover', stop)
    window.addEventListener('drop', stop)
    return () => { window.removeEventListener('dragover', stop); window.removeEventListener('drop', stop) }
  }, [])

  async function load(kind, file) {
    setResult(null)
    setLoading((l) => ({ ...l, [kind]: true }))
    try {
      const parsed = await readTable(file)
      setErrors((e) => ({ ...e, [kind]: null, merge: null }))
      if (kind === 'base') {
        setBase(parsed)
        setBaseFile(file)
        setWork(null) // a fresh base resets the session's accumulated merges
        setMap((m) => ({ ...m, baseKey: detectBaseKeyColumn(parsed.headers) }))
      } else {
        const callKey = detectCallKeyColumn(parsed.headers)
        setCalls(parsed)
        setMap((m) => ({
          ...m,
          callKey,
          dateCol: detectDateColumn(parsed.headers),
          notesCol: detectNotesColumn(parsed.headers),
        }))
        setCarry(defaultCarryColumns(parsed.headers, callKey))
      }
    } catch (err) {
      setErrors((e) => ({ ...e, [kind]: err.message }))
      if (kind === 'base') { setBase(null); setBaseFile(null); setWork(null) } else setCalls(null)
    } finally {
      setLoading((l) => ({ ...l, [kind]: false }))
    }
  }

  const blockers = []
  if (!base) blockers.push('a base contacts file')
  if (!calls) blockers.push('a calls export')
  if (base && calls && !map.baseKey) blockers.push('the base contact ID column')
  if (base && calls && !map.callKey) blockers.push('the calls contact ID column')
  if (base && calls && !map.dateCol) blockers.push('the calls date column')
  if (base && calls && !carry.length) blockers.push('at least one call column to add')
  const ready = blockers.length === 0

  function run() {
    setBusy(true)
    // yield once so the button repaints before the synchronous merge
    setTimeout(() => {
      try {
        const src = work ?? { rows: base.rows, headers: base.headers }
        const r = merge({
          baseRows: src.rows, baseHeaders: src.headers,
          callRows: calls.rows, callHeaders: calls.headers,
          carryColumns: carry, ...map,
        })
        // Colours and columns are cumulative across the session so a download
        // always reflects everything changed since the base was loaded.
        const highlights = new Map(work?.highlights ?? [])
        for (const [k, v] of r.highlights) highlights.set(k, v)
        const union = (a = [], b = []) => [...new Set([...a, ...b])]
        const addedHeaders = union(work?.addedHeaders, r.addedHeaders)
        const targetHeaders = union(work?.targetHeaders, r.targetHeaders)

        setWork({ rows: r.rows, headers: r.headers, highlights, addedHeaders, targetHeaders })
        setResult({ ...r, highlights, addedHeaders, targetHeaders, source: base.workbook ? base : null })
        setErrors((e) => ({ ...e, merge: null }))
      } catch (err) {
        setErrors((e) => ({ ...e, merge: err.message }))
      } finally {
        setBusy(false)
      }
    }, 0)
  }

  async function save(kind) {
    setSaving(kind)
    try {
      if (kind === 'xlsx') await downloadXlsx(result, `${stem}-merged.xlsx`)
      else downloadCsv(result, `${stem}-merged.csv`)
      setErrors((e) => ({ ...e, save: null }))
    } catch (err) {
      setErrors((e) => ({ ...e, save: `Could not build the ${kind.toUpperCase()}: ${err.message}` }))
    } finally {
      setSaving('')
    }
  }

  const nameCols = useMemo(
    () => (base?.headers ?? []).filter((h) => /name|email/i.test(h)).slice(0, 2),
    [base],
  )
  const missingStandard = useMemo(
    () => (calls ? DEFAULT_CALL_COLUMNS.filter((h) => !calls.headers.includes(h)) : []),
    [calls],
  )
  const sameFile = base && calls && base.name === calls.name
  const stem = base?.name.replace(/\.(csv|xlsx?)$/i, '') ?? 'merged'
  const st = result?.stats
  const changedRows = result ? result.preview.filter((p) => p.changed) : []

  return (
    <div className="app">
      <header>
        <h1>Call Report Merger</h1>
        <p>Folds a HubSpot calls export into your base contacts file, keeping a dated notes history per contact.</p>
      </header>

      <section>
        <h2><span className="step">1</span> Load the two exports</h2>
        <div className="drops">
          <Drop
            label="Base contacts file"
            hint="The file that gets updated. One row per contact, keyed by Record ID."
            file={base} error={errors.base} busy={loading.base} onFile={(f) => load('base', f)}
          />
          <Drop
            label="Calls export"
            hint="The call activity to merge in. Extra columns are ignored."
            file={calls} error={errors.calls} busy={loading.calls} onFile={(f) => load('calls', f)}
          />
        </div>
        {sameFile && (
          <div className="warn-box plain">
            Both slots have a file called <strong>{base.name}</strong>. That is almost certainly a
            mistake &mdash; the base is your contacts list, the other is the calls export.
          </div>
        )}
      </section>

      {base && calls && (
        <section>
          <h2><span className="step">2</span> Check the column mapping</h2>
          <p className="sub">Auto-detected from the headers. Change these if a future export names things differently.</p>
          <div className="fields">
            <Select label="Base &mdash; contact ID" value={map.baseKey} options={base.headers}
              onChange={(v) => setMap((m) => ({ ...m, baseKey: v }))} />
            <Select label="Calls &mdash; contact ID" value={map.callKey} options={calls.headers}
              note="Not the call&rsquo;s own Record ID"
              onChange={(v) => setMap((m) => ({ ...m, callKey: v }))} />
            <Select label="Calls &mdash; date" value={map.dateCol} options={calls.headers}
              note="Groups the calls into days"
              onChange={(v) => setMap((m) => ({ ...m, dateCol: v }))} />
            <Select label="Calls &mdash; notes" value={map.notesCol} options={calls.headers}
              note="The one column that accumulates"
              onChange={(v) => setMap((m) => ({ ...m, notesCol: v }))} />
          </div>

          {missingStandard.length > 0 && (
            <div className="warn-box plain">
              <strong>{missingStandard.join(', ')}</strong>{' '}
              {missingStandard.length === 1 ? 'is' : 'are'} not in this calls export, so{' '}
              {missingStandard.length === 1 ? 'that column' : 'those columns'} will be missing from the
              report. Re-export from HubSpot with {missingStandard.length === 1 ? 'it' : 'them'} included,
              or tick a replacement below.
            </div>
          )}

          {map.notesCol && !carry.includes(map.notesCol) && (
            <div className="warn-box plain">
              <strong>{targetColumn(map.notesCol)}</strong> is unticked below, so no notes history will be
              written this run. Tick it to keep the notes column.
            </div>
          )}

          <div className="picker">
            <div className="picker-head">
              <span className="field-label">Call columns to add &mdash; {carry.length} selected</span>
              <button className="link" onClick={() => setCarry(defaultCarryColumns(calls.headers, map.callKey))}>
                Reset to the standard {DEFAULT_CALL_COLUMNS.length}
              </button>
            </div>
            <div className="chips">
              {calls.headers.filter((h) => h !== map.callKey).map((h) => {
                const on = carry.includes(h)
                return (
                  <label key={h} className={`chip ${on ? 'on' : ''}`}>
                    <input
                      type="checkbox" checked={on}
                      onChange={() => setCarry((c) => (on ? c.filter((x) => x !== h) : [...c, h]))}
                    />
                    {targetColumn(h)}
                  </label>
                )
              })}
            </div>
          </div>

          <ul className="rules">
            <li><strong>{map.notesCol || 'Notes'}</strong> from every day in this export are merged in &mdash; a calls export is usually full history, not just today, so nothing gets dropped.</li>
            <li>Each day&rsquo;s notes are joined together (oldest call first, duplicates dropped) and <strong>added on top of the notes already in the base</strong>, under a <code>[date]</code> heading, newest day first.</li>
            <li>Re-merging a day replaces just that day&rsquo;s block &mdash; running the same export twice changes nothing.</li>
            <li>Every other call column takes the <strong>latest non-blank</strong> value from the contact&rsquo;s single most recent day &mdash; those columns hold the latest only, not a history.</li>
            <li>Existing base columns are never touched &mdash; their values, date formats and widths are left exactly as they are.</li>
          </ul>

          <div className="runbar">
            <button className="primary" disabled={!ready || busy} onClick={run}>
              {busy ? 'Merging…' : work ? 'Merge this file too' : 'Merge'}
            </button>
            {!ready && <span className="field-note">Still need {blockers.join(', ')}.</span>}
            {work && ready && !busy && (
              <span className="field-note">Builds on what you have already merged this session.</span>
            )}
          </div>
          {errors.merge && <div className="alert">{errors.merge}</div>}
        </section>
      )}

      {result && (
        <section>
          <h2><span className="step">3</span> Review and download</h2>

          <div className="stats">
            <Stat n={st.contactsUpdated} label="contacts updated" tone={st.contactsUpdated ? 'good' : 'warn'} />
            <Stat n={st.historyDaysMerged} label="call-days added to notes" />
            <Stat n={st.contactsAlreadyCurrent} label="already up to date" muted />
            <Stat n={st.unmatched.length} label="IDs not in base" tone={st.unmatched.length ? 'warn' : ''} />
            <Stat n={st.noContactId} label="calls with no contact ID" muted />
          </div>

          {st.contactsUpdated === 0 && (
            <div className="warn-box plain">
              <strong>Nothing changed.</strong>{' '}
              {st.contactsAlreadyCurrent > 0
                ? 'Every matching contact already has this call data — merging the same export twice is a no-op, so this is expected if you have already run it.'
                : 'No call in this export matched a contact in the base. Check that the two contact ID columns in step 2 point at real HubSpot Record IDs.'}
            </div>
          )}

          {(st.mangledCallIds > 0 || st.mangledBaseIds > 0) && (
            <div className="warn-box plain">
              <strong>Excel has damaged some Record IDs.</strong>{' '}
              {st.mangledCallIds > 0 && `${st.mangledCallIds} ID${st.mangledCallIds === 1 ? '' : 's'} in the calls export `}
              {st.mangledCallIds > 0 && st.mangledBaseIds > 0 && 'and '}
              {st.mangledBaseIds > 0 && `${st.mangledBaseIds} in the base `}
              look like <code>1.14178E+11</code> instead of a 12-digit number. That happens when a CSV is
              opened and re-saved in Excel, and those rows can never match. Re-export from HubSpot as
              .xlsx, or open the CSV without saving it.
            </div>
          )}

          {st.notesTrimmed > 0 && (
            <div className="warn-box plain">
              <strong>{st.notesTrimmed} contact{st.notesTrimmed === 1 ? "'s" : "s'"} notes hit Excel&rsquo;s cell limit.</strong>{' '}
              Excel refuses to open a file with a cell over 32,767 characters, so the oldest days were
              dropped from {st.notesTrimmed === 1 ? 'that cell' : 'those cells'} to keep the file valid.
              The full history is still in HubSpot.
            </div>
          )}

          {st.duplicateBaseIds.length > 0 && (
            <details className="warn-box">
              <summary>
                {st.duplicateBaseIds.length} contact ID{st.duplicateBaseIds.length === 1 ? ' appears' : 's appear'} more
                than once in the base &mdash; only the last row of each got the call data
              </summary>
              <div className="idlist">{st.duplicateBaseIds.map((u) => `${u.id} (${u.count}×)`).join('   ')}</div>
            </details>
          )}

          <p className="sub">
            {result.addedHeaders.length
              ? <>Appended {result.addedHeaders.length} new columns: <strong>{result.addedHeaders.join(', ')}</strong></>
              : <>No new columns &mdash; every call column already existed and was updated in place.</>}
          </p>

          {st.unmatched.length > 0 && (
            <details className="warn-box">
              <summary>
                {st.unmatched.length} contact IDs in the calls file have no row in the base &mdash; nothing was written for them
              </summary>
              <div className="idlist">{st.unmatched.map((u) => `${u.id} (${u.count})`).join('   ')}</div>
            </details>
          )}

          <div className="legend">
            <span><i className="sw new" /> new value or a day appended</span>
            <span><i className="sw updated" /> replaced an existing value</span>
            <span><i className="sw header" /> new column</span>
            <span className="legend-note">highlights appear in the .xlsx only &mdash; CSV has no cell colours</span>
          </div>

          {changedRows.length > 0 && (
            <>
              <div className="tablewrap">
                <table>
                  <thead>
                    <tr>
                      <th>{map.baseKey}</th>
                      {nameCols.map((h) => <th key={h}>{h}</th>)}
                      <th>Latest day</th>
                      <th className="num">Days merged</th>
                      <th>Notes added this run</th>
                    </tr>
                  </thead>
                  <tbody>
                    {changedRows.slice(0, 100).map((p) => (
                      <tr key={p.id}>
                        <td className="mono">{p.id}</td>
                        {nameCols.map((h) => <td key={h}>{result.rows[p.rowIndex][h]}</td>)}
                        <td className="mono">{p.day}</td>
                        <td className="num">{p.daysMerged}</td>
                        <td className="notes">{p.newNotes || <em>no notes</em>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {changedRows.length > 100 && (
                <p className="sub">Showing the 100 most recent of {changedRows.length.toLocaleString()} updated contacts.</p>
              )}
            </>
          )}

          <div className="downloads">
            <button className="primary" disabled={!!saving} onClick={() => save('xlsx')}>
              {saving === 'xlsx' ? 'Building…' : <>Download .xlsx<small>with highlights</small></>}
            </button>
            <button disabled={!!saving} onClick={() => save('csv')}>
              {saving === 'csv' ? 'Building…' : <>Download .csv<small>plain, import-ready</small></>}
            </button>
            {baseFile && (
              <button className="link reset" onClick={() => { setResult(null); setWork(null); load('base', baseFile) }}>
                Start over from the uploaded base
              </button>
            )}
          </div>
          {errors.save && <div className="alert">{errors.save}</div>}
        </section>
      )}
    </div>
  )
}
