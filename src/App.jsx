import React, { useMemo, useRef, useState } from 'react'
import {
  merge, targetColumn, defaultCarryColumns, DEFAULT_CALL_COLUMNS,
  detectBaseKeyColumn, detectCallKeyColumn, detectDateColumn, detectNotesColumn,
} from './merge.js'
import { readTable } from './readTable.js'
import { downloadCsv, downloadXlsx } from './download.js'

function Drop({ label, hint, file, onFile, error }) {
  const input = useRef(null)
  const [over, setOver] = useState(false)
  const pick = (f) => f && onFile(f)

  return (
    <div
      className={`drop ${file ? 'ok' : ''} ${over ? 'over' : ''} ${error ? 'bad' : ''}`}
      onClick={() => input.current.click()}
      onDragOver={(e) => { e.preventDefault(); setOver(true) }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => { e.preventDefault(); setOver(false); pick(e.dataTransfer.files[0]) }}
    >
      <input ref={input} type="file" accept=".csv,.xlsx,.xls" hidden onChange={(e) => pick(e.target.files[0])} />
      <div className="drop-label">{label}</div>
      {file ? (
        <>
          <div className="drop-file">{file.name}</div>
          <div className="drop-meta">{file.rows.length.toLocaleString()} rows &middot; {file.headers.length} columns</div>
          <div className="drop-swap">Click to replace</div>
        </>
      ) : (
        <>
          <div className="drop-hint">{hint}</div>
          <div className="drop-swap">Drop a .xlsx or .csv here, or click to browse</div>
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
  const [calls, setCalls] = useState(null)
  const [errors, setErrors] = useState({})
  const [map, setMap] = useState({ baseKey: '', callKey: '', dateCol: '', notesCol: '' })
  const [carry, setCarry] = useState([])
  const [result, setResult] = useState(null)
  const [busy, setBusy] = useState(false)

  async function load(kind, file) {
    setResult(null)
    try {
      const parsed = await readTable(file)
      setErrors((e) => ({ ...e, [kind]: null, merge: null }))
      if (kind === 'base') {
        setBase(parsed)
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
    }
  }

  const ready = base && calls && map.baseKey && map.callKey && map.dateCol && carry.length

  function run() {
    setBusy(true)
    // yield once so the button repaints before the synchronous merge
    setTimeout(() => {
      try {
        setResult({
          ...merge({
            baseRows: base.rows, baseHeaders: base.headers,
            callRows: calls.rows, callHeaders: calls.headers,
            carryColumns: carry, ...map,
          }),
          source: base.workbook ? base : null,
        })
        setErrors((e) => ({ ...e, merge: null }))
      } catch (err) {
        setErrors((e) => ({ ...e, merge: err.message }))
      } finally {
        setBusy(false)
      }
    }, 0)
  }

  const nameCols = useMemo(
    () => (base?.headers ?? []).filter((h) => /name|email/i.test(h)).slice(0, 2),
    [base],
  )
  const missingStandard = useMemo(
    () => (calls ? DEFAULT_CALL_COLUMNS.filter((h) => !calls.headers.includes(h)) : []),
    [calls],
  )
  const stem = base?.name.replace(/\.(csv|xlsx?)$/i, '') ?? 'merged'

  return (
    <div className="app">
      <header>
        <h1>Call Report Merger</h1>
        <p>Takes each contact&rsquo;s most recent call day out of a HubSpot calls export and writes it onto your base contacts file.</p>
      </header>

      <section>
        <h2><span className="step">1</span> Load the two exports</h2>
        <div className="drops">
          <Drop
            label="Base contacts file"
            hint="The file that gets updated. One row per contact, keyed by Record ID."
            file={base} error={errors.base} onFile={(f) => load('base', f)}
          />
          <Drop
            label="Calls export"
            hint="The new call activity to merge in. Extra columns are ignored."
            file={calls} error={errors.calls} onFile={(f) => load('calls', f)}
          />
        </div>
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
              note="Decides which day is the latest"
              onChange={(v) => setMap((m) => ({ ...m, dateCol: v }))} />
            <Select label="Calls &mdash; notes" value={map.notesCol} options={calls.headers}
              note="The one column that gets merged, not overwritten"
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
            <li>Only each contact&rsquo;s <strong>latest call day</strong> in this export is used &mdash; earlier days in the same file are ignored.</li>
            <li><strong>{map.notesCol || 'Notes'}</strong> from every call on that day are joined together, oldest first, duplicates dropped.</li>
            <li>That day is then <strong>added on top of the notes already in the base</strong>, under a <code>[date]</code> heading. Previous days are kept; re-merging a day replaces just that day.</li>
            <li>Every other call column takes the <strong>latest non-blank</strong> value from that day &mdash; those columns hold the latest only, not a history.</li>
            <li>Existing base columns are never touched &mdash; their values, date formats and widths are left exactly as they are.</li>
          </ul>

          <button className="primary" disabled={!ready || busy} onClick={run}>
            {busy ? 'Merging…' : 'Merge'}
          </button>
          {errors.merge && <div className="alert">{errors.merge}</div>}
        </section>
      )}

      {result && (
        <section>
          <h2><span className="step">3</span> Review and download</h2>

          <div className="stats">
            <Stat n={result.stats.contactsUpdated} label="contacts updated" tone="good" />
            <Stat n={result.stats.usedCalls} label="calls merged in" />
            <Stat n={result.stats.skippedOlderDay} label="calls on an older day" muted />
            <Stat n={result.stats.unmatched.length} label="IDs not in base" tone={result.stats.unmatched.length ? 'warn' : ''} />
            <Stat n={result.stats.noContactId} label="calls with no contact ID" muted />
          </div>

          <p className="sub">
            {result.addedHeaders.length
              ? <>Appended {result.addedHeaders.length} new columns: <strong>{result.addedHeaders.join(', ')}</strong></>
              : <>No new columns &mdash; every call column already existed and was updated in place.</>}
          </p>

          {result.stats.unmatched.length > 0 && (
            <details className="warn-box">
              <summary>
                {result.stats.unmatched.length} contact IDs in the calls file have no row in the base &mdash; nothing was written for them
              </summary>
              <div className="idlist">
                {result.stats.unmatched.map((u) => `${u.id} (${u.count})`).join('   ')}
              </div>
            </details>
          )}

          <div className="legend">
            <span><i className="sw new" /> new value or a day appended</span>
            <span><i className="sw updated" /> replaced an existing value</span>
            <span><i className="sw header" /> new column</span>
            <span className="legend-note">highlights appear in the .xlsx only &mdash; CSV has no cell colours</span>
          </div>

          <div className="tablewrap">
            <table>
              <thead>
                <tr>
                  <th>{map.baseKey}</th>
                  {nameCols.map((h) => <th key={h}>{h}</th>)}
                  <th>Day used</th>
                  <th className="num">Calls</th>
                  <th>Notes added this run</th>
                </tr>
              </thead>
              <tbody>
                {result.preview.slice(0, 100).map((p) => (
                  <tr key={p.id}>
                    <td className="mono">{p.id}</td>
                    {nameCols.map((h) => <td key={h}>{result.rows[p.rowIndex][h]}</td>)}
                    <td className="mono">{p.day}</td>
                    <td className="num">{p.callCount}</td>
                    <td className="notes">{p.newNotes || <em>no notes</em>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {result.preview.length > 100 && (
            <p className="sub">Showing the 100 most recent of {result.preview.length.toLocaleString()} updated contacts.</p>
          )}

          <div className="downloads">
            <button className="primary" onClick={() => downloadXlsx(result, `${stem}-merged.xlsx`)}>
              Download .xlsx<small>with highlights</small>
            </button>
            <button onClick={() => downloadCsv(result, `${stem}-merged.csv`)}>
              Download .csv<small>plain, import-ready</small>
            </button>
          </div>
        </section>
      )}
    </div>
  )
}
