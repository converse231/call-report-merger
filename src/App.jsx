import React, { useEffect, useMemo, useRef, useState } from 'react'
import {
  merge, targetColumn, defaultCarryColumns, DEFAULT_CALL_COLUMNS, parseNoteBlocks,
  detectBaseKeyColumn, detectCallKeyColumn, detectDateColumn, detectNotesColumn,
  detectSplitDateColumn, detectProjectColumn, groupByProject, sliceResult, splitTargets, NO_PROJECT,
} from './merge.js'
import { readTable } from './readTable.js'
import { downloadCsv, downloadXlsx } from './download.js'

const n = (x) => x.toLocaleString()
const scopeTag = (s) => (s ? `-${s.replace(/[^\w.-]+/g, '_').replace(/^_|_$/g, '')}` : '')

/* ------------------------------------------------------------------ files */

function DropZone({ step, title, hint, children, onFile, error, busy, filled }) {
  const input = useRef(null)
  const open = () => input.current.click()
  const [over, setOver] = useState(false)

  return (
    <div
      className={`zone ${filled ? 'filled' : ''} ${over ? 'over' : ''} ${error ? 'bad' : ''}`}
      onDragOver={(e) => { e.preventDefault(); setOver(true) }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => { e.preventDefault(); setOver(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]) }}
    >
      <input
        ref={input} type="file" accept=".csv,.xlsx,.xls" hidden
        // reset the value so picking the same file again still fires a change
        onChange={(e) => { const f = e.target.files[0]; e.target.value = ''; if (f) onFile(f) }}
      />
      <div className="zone-head">
        <span className={`zone-step ${filled ? 'done' : ''}`}>{filled ? '✓' : step}</span>
        <span className="zone-title">{title}</span>
      </div>

      {busy ? <div className="zone-hint">Reading the file…</div> : children}

      {!filled && !busy && <div className="zone-hint">{hint}</div>}
      <button className="zone-btn" onClick={open} disabled={busy}>
        {filled ? 'Choose a different file' : 'Choose file'}
        <span className="zone-drop">or drop it here</span>
      </button>
      {error && <div className="zone-error">{error}</div>}
    </div>
  )
}

/* ------------------------------------------------------- settings summary */

function Chip({ on, onToggle, children }) {
  return (
    <label className={`chip ${on ? 'on' : ''}`}>
      <input type="checkbox" checked={on} onChange={onToggle} />
      {children}
    </label>
  )
}

function Field({ label, value, options, onChange, note, noneLabel = '— none —' }) {
  return (
    <label className="field">
      <span className="field-label">{label}</span>
      <select value={value} onChange={(e) => onChange(e.target.value)}>
        <option value="">{noneLabel}</option>
        {options.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
      {note && <span className="field-note">{note}</span>}
    </label>
  )
}

/* ------------------------------------------------------------- note cell */

function Notes({ text }) {
  const blocks = parseNoteBlocks(text)
  if (!blocks.length) return <em className="dim">no notes</em>
  return (
    <div className="entries">
      {blocks.map((b, i) => (
        <div className="entry" key={i}>
          {b.day && <span className="entry-day">{b.day}</span>}
          <span className="entry-text">{b.text}</span>
        </div>
      ))}
    </div>
  )
}

/* ------------------------------------------------------------------- app */

/** The whole result half of the page. Exported so `npm test` can render it with real data. */
export function ResultView({ result, map, carry, projects, scope, onScope, saving, onSave, stem, nameCols, missingStandard, saveError, onStartOver }) {
  const st = result.stats
  const changed = result.preview.filter((p) => p.changed)
  const tag = scopeTag(scope)
  const scopeRows = scope ? projects.get(scope)?.length ?? 0 : result.rows.length
  return (
    <>
          <section className="out">
            <div className={`verdict ${st.contactsUpdated ? 'good' : 'warn'}`}>
              <span className="verdict-icon">{st.contactsUpdated ? '✓' : '!'}</span>
              <div>
                <div className="verdict-main">
                  {st.contactsUpdated
                    ? <>{n(st.contactsUpdated)} {st.contactsUpdated === 1 ? 'contact' : 'contacts'} updated</>
                    : <>Nothing changed</>}
                </div>
                <div className="verdict-sub">
                  {st.contactsUpdated
                    ? <>{n(st.historyDaysMerged)} call-{st.historyDaysMerged === 1 ? 'day' : 'days'} of notes added across {n(st.usedCalls)} calls</>
                    : st.contactsAlreadyCurrent > 0
                      ? <>All {n(st.contactsAlreadyCurrent)} matching contacts already have this call data.</>
                      : <>No call matched a contact. Check the contact ID columns under “Change” above.</>}
                </div>
              </div>
            </div>

            <div className="export">
              {projects.size > 1 ? (
                <div className="export-scope">
                  <span className="field-label">Export</span>
                  <div className="pills">
                    <button className={`pill ${!scope ? 'on' : ''}`} onClick={() => onScope('')}>
                      All contacts <b>{n(result.rows.length)}</b>
                    </button>
                    {[...projects].map(([name, idx]) => (
                      <button key={name} className={`pill ${scope === name ? 'on' : ''}`} onClick={() => onScope(name)}>
                        {name} <b>{n(idx.length)}</b>
                      </button>
                    ))}
                  </div>
                </div>
              ) : (
                <p className="export-note">
                  {!map.projectCol
                    ? <>Want one file per project? Pick your project column under <strong>Change</strong> above.</>
                    : <>No contact has a value in <strong>{map.projectCol}</strong> yet, so there are no projects to
                       split by. Fill it in on HubSpot and the options appear here.</>}
                </p>
              )}

              <div className="export-go">
                <button className="primary" disabled={!!saving} onClick={() => onSave('xlsx')}>
                  {saving === 'xlsx' ? 'Building…' : 'Download Excel'}
                </button>
                <button disabled={!!saving} onClick={() => onSave('csv')}>
                  {saving === 'csv' ? 'Building…' : 'CSV'}
                </button>
                <span className="export-file">
                  <code>{stem}-merged{tag}.xlsx</code> · {n(scopeRows)} rows
                  {scope && <> · built fresh, so the original file’s formatting isn’t carried over</>}
                </span>
              </div>
              {saveError && <div className="alert">{saveError}</div>}
            </div>
          </section>

          {(missingStandard.length > 0 || st.mangledCallIds > 0 || st.mangledBaseIds > 0 || st.notesTrimmed > 0
            || (map.notesCol && !carry.includes(map.notesCol))) && (
            <div className="notes-stack">
              {map.notesCol && !carry.includes(map.notesCol) && (
                <div className="note warn"><strong>{targetColumn(map.notesCol)}</strong> is unticked, so no call notes are being written.</div>
              )}
              {missingStandard.length > 0 && (
                <div className="note warn">
                  <strong>{missingStandard.join(', ')}</strong> {missingStandard.length === 1 ? 'is' : 'are'} missing
                  from the calls export, so {missingStandard.length === 1 ? 'it' : 'they'} can’t be filled in.
                  Re-export from HubSpot with {missingStandard.length === 1 ? 'it' : 'them'} included.
                </div>
              )}
              {(st.mangledCallIds > 0 || st.mangledBaseIds > 0) && (
                <div className="note warn">
                  <strong>Excel has damaged some Record IDs.</strong>{' '}
                  {st.mangledCallIds > 0 && `${st.mangledCallIds} in the calls file `}
                  {st.mangledCallIds > 0 && st.mangledBaseIds > 0 && 'and '}
                  {st.mangledBaseIds > 0 && `${st.mangledBaseIds} in the contacts file `}
                  read as <code>1.14178E+11</code> instead of a 12-digit number, so those rows can never match.
                  That happens when a CSV is opened and re-saved in Excel — re-export as .xlsx.
                </div>
              )}
              {st.notesTrimmed > 0 && (
                <div className="note warn">
                  <strong>{st.notesTrimmed} contact{st.notesTrimmed === 1 ? '' : 's'}</strong> hit Excel’s
                  32,767-character cell limit, so the oldest notes were dropped to keep the file openable.
                  The full history is still in HubSpot.
                </div>
              )}
            </div>
          )}

          {changed.length > 0 && (
            <section>
              <div className="out-head">
                <h2>What changed</h2>
                <span className="legend">
                  <i className="sw new" /> added
                  <i className="sw updated" /> replaced
                  <i className="sw header" /> new column
                  <em>colours show in the Excel file</em>
                </span>
              </div>
              <div className="tablewrap">
                <table>
                  <thead>
                    <tr>
                      <th>Contact</th>
                      <th className="num">Days</th>
                      <th>Notes added</th>
                    </tr>
                  </thead>
                  <tbody>
                    {changed.slice(0, 100).map((p) => (
                      <tr key={p.id}>
                        <td>
                          <div className="who">{nameCols.map((h) => result.rows[p.rowIndex][h]).join(' ').trim() || '—'}</div>
                          <div className="mono dim">{p.id}</div>
                        </td>
                        <td className="num">{p.daysMerged}</td>
                        <td><Notes text={p.newNotes} /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {changed.length > 100 && <p className="sub">Showing 100 of {n(changed.length)}.</p>}
            </section>
          )}

          <div className="tally">
            <span>{n(st.baseRows)} contacts in file</span>
            <span>{n(st.usedCalls)} calls read</span>
            {st.contactsAlreadyCurrent > 0 && <span>{n(st.contactsAlreadyCurrent)} already up to date</span>}
            {st.noContactId > 0 && <span>{n(st.noContactId)} calls with no contact</span>}
            {st.unmatched.length > 0 && (
              <details>
                <summary>{n(st.unmatched.length)} IDs not in contacts</summary>
                <div className="idlist">{st.unmatched.map((u) => `${u.id} (${u.count})`).join('   ')}</div>
              </details>
            )}
            {st.duplicateBaseIds.length > 0 && (
              <details>
                <summary>{n(st.duplicateBaseIds.length)} duplicate contact IDs</summary>
                <div className="idlist">
                  Only the last row of each was updated.{' '}
                  {st.duplicateBaseIds.map((u) => `${u.id} (${u.count}×)`).join('   ')}
                </div>
              </details>
            )}
            {onStartOver && (
              <button className="link" onClick={onStartOver}>
                Start over
              </button>
            )}
          </div>
    </>
  )
}

export default function App() {
  const [base, setBase] = useState(null)
  const [baseFile, setBaseFile] = useState(null)
  const [callFiles, setCallFiles] = useState([])
  const [errors, setErrors] = useState({})
  const [loading, setLoading] = useState({})
  const [map, setMap] = useState({ baseKey: '', callKey: '', dateCol: '', notesCol: '', splitCol: '', projectCol: '' })
  const [carry, setCarry] = useState([])
  const [scope, setScope] = useState('')
  const [tuning, setTuning] = useState(false)
  const [saving, setSaving] = useState('')

  // A file dropped outside a zone would otherwise navigate away and lose everything.
  useEffect(() => {
    const stop = (e) => e.preventDefault()
    window.addEventListener('dragover', stop)
    window.addEventListener('drop', stop)
    return () => { window.removeEventListener('dragover', stop); window.removeEventListener('drop', stop) }
  }, [])

  async function load(kind, file) {
    setLoading((l) => ({ ...l, [kind]: true }))
    try {
      const t = await readTable(file)
      setErrors((e) => ({ ...e, [kind]: null }))
      if (kind === 'base') {
        setBase(t)
        setBaseFile(file)
        setScope('')
        setMap((m) => ({
          ...m,
          baseKey: detectBaseKeyColumn(t.headers),
          splitCol: detectSplitDateColumn(t.headers),
          projectCol: detectProjectColumn(t.headers),
        }))
      } else {
        const callKey = detectCallKeyColumn(t.headers)
        setCallFiles((fs) => (fs.some((f) => f.name === t.name) ? fs : [...fs, t]))
        setMap((m) => ({
          ...m,
          callKey,
          dateCol: detectDateColumn(t.headers),
          notesCol: detectNotesColumn(t.headers),
        }))
        setCarry((c) => [...new Set([...c, ...defaultCarryColumns(t.headers, callKey)])])
      }
    } catch (err) {
      setErrors((e) => ({ ...e, [kind]: err.message }))
      if (kind === 'base') { setBase(null); setBaseFile(null) }
    } finally {
      setLoading((l) => ({ ...l, [kind]: false }))
    }
  }

  /**
   * The merge is pure and takes milliseconds, so there is no Merge button --
   * the result recomputes from the untouched contacts file whenever anything
   * changes. Recomputing from scratch is also what keeps several calls files
   * from being applied twice.
   * ponytail: synchronous; if a file ever gets big enough to jank, move to a worker.
   */
  const result = useMemo(() => {
    if (!base || !callFiles.length || !map.baseKey) return null
    const has = (t, col, fallback) => (t.headers.includes(col) ? col : fallback(t.headers))

    let rows = base.rows
    let headers = base.headers
    const highlights = new Map()
    let addedHeaders = []
    let targetHeaders = []
    let last = null

    for (const t of callFiles) {
      const callKey = has(t, map.callKey, detectCallKeyColumn)
      const picked = carry.filter((h) => t.headers.includes(h))
      const r = merge({
        baseRows: rows, baseHeaders: headers,
        callRows: t.rows, callHeaders: t.headers,
        baseKey: map.baseKey,
        callKey,
        dateCol: has(t, map.dateCol, detectDateColumn),
        notesCol: has(t, map.notesCol, detectNotesColumn),
        carryColumns: picked.length ? picked : defaultCarryColumns(t.headers, callKey),
        splitDateColumns: [map.splitCol],
      })
      rows = r.rows
      headers = r.headers
      for (const [k, v] of r.highlights) highlights.set(k, v)
      addedHeaders = [...new Set([...addedHeaders, ...r.addedHeaders])]
      targetHeaders = [...new Set([...targetHeaders, ...r.targetHeaders])]
      last = r
    }
    return {
      ...last, rows, headers, highlights, addedHeaders, targetHeaders,
      source: base.workbook ? base : null,
    }
  }, [base, callFiles, map, carry])

  const projects = useMemo(
    () => (result && map.projectCol ? groupByProject(result.rows, map.projectCol) : new Map()),
    [result, map.projectCol],
  )
  const missingStandard = useMemo(
    () => {
      const all = new Set(callFiles.flatMap((t) => t.headers))
      return callFiles.length ? DEFAULT_CALL_COLUMNS.filter((h) => !all.has(h)) : []
    },
    [callFiles],
  )

  const stem = base?.name.replace(/\.(csv|xlsx?)$/i, '') ?? 'contacts'
  const tag = scopeTag(scope)
  const scopeRows = scope ? projects.get(scope)?.length ?? 0 : result?.rows.length ?? 0
  const st = result?.stats
  const changed = result ? result.preview.filter((p) => p.changed) : []
  const nameCols = useMemo(
    () => (base?.headers ?? []).filter((h) => /name/i.test(h)).slice(0, 2),
    [base],
  )
  const callCols = [...new Set(callFiles.flatMap((t) => t.headers))].filter((h) => h !== map.callKey)

  async function save(kind) {
    setSaving(kind)
    try {
      const idx = scope ? projects.get(scope) : null
      const payload = idx ? sliceResult(result, idx) : result
      if (kind === 'xlsx') await downloadXlsx(payload, `${stem}-merged${tag}.xlsx`)
      else downloadCsv(payload, `${stem}-merged${tag}.csv`)
      setErrors((e) => ({ ...e, save: null }))
    } catch (err) {
      setErrors((e) => ({ ...e, save: `Could not build the file: ${err.message}` }))
    } finally {
      setSaving('')
    }
  }

  return (
    <div className="app">
      <header>
        <h1>Call Report Merger</h1>
        <p>Drop in your HubSpot contacts and calls exports. The merged file is ready straight away.</p>
      </header>

      {/* ---------------------------------------------------------- files */}
      <div className="zones">
        <DropZone
          step="1" title="Contacts" filled={!!base} busy={loading.base} error={errors.base}
          hint="The file that gets updated — one row per contact."
          onFile={(f) => load('base', f)}
        >
          {base && (
            <div className="file">
              <span className="file-name">{base.name}</span>
              <span className="file-meta">{n(base.rows.length)} contacts · {base.headers.length} columns</span>
            </div>
          )}
        </DropZone>

        <div className="zones-plus">+</div>

        <DropZone
          step="2" title="Calls" filled={callFiles.length > 0} busy={loading.calls} error={errors.calls}
          hint="The call activity to add in."
          onFile={(f) => load('calls', f)}
        >
          {callFiles.map((t) => (
            <div className="file" key={t.name}>
              <span className="file-name">{t.name}</span>
              <span className="file-meta">{n(t.rows.length)} calls</span>
              <button
                className="file-x" title="Remove this file"
                onClick={() => setCallFiles((fs) => fs.filter((f) => f.name !== t.name))}
              >×</button>
            </div>
          ))}
          {callFiles.length > 0 && (
            <div className="zone-hint">Add another export to combine several days.</div>
          )}
        </DropZone>
      </div>

      {/* ------------------------------------------------------- settings */}
      {base && callFiles.length > 0 && (
        <div className={`tune ${tuning ? 'open' : ''}`}>
          <div className="tune-bar">
            <p className="tune-say">
              Adding <strong>{carry.length}</strong> call {carry.length === 1 ? 'column' : 'columns'}
              {map.splitCol && <>, splitting <strong>{map.splitCol}</strong> into date + time</>}
              , matching on <strong>{map.callKey || '?'}</strong> → <strong>{map.baseKey || '?'}</strong>.
            </p>
            <button className="ghost" onClick={() => setTuning((v) => !v)}>
              {tuning ? 'Done' : 'Change'}
            </button>
          </div>

          {tuning && (
            <div className="tune-body">
              <div className="tune-group">
                <div className="tune-head">
                  <h3>Call columns to add</h3>
                  <button className="link" onClick={() => setCarry(defaultCarryColumns(callFiles.at(-1).headers, map.callKey))}>
                    Reset to the standard {DEFAULT_CALL_COLUMNS.length}
                  </button>
                </div>
                <div className="chips">
                  {callCols.map((h) => (
                    <Chip
                      key={h} on={carry.includes(h)}
                      onToggle={() => setCarry((c) => (c.includes(h) ? c.filter((x) => x !== h) : [...c, h]))}
                    >{targetColumn(h)}</Chip>
                  ))}
                </div>
              </div>

              <div className="tune-group">
                <h3>Split a date into two columns</h3>
                <div className="fields">
                  <Field
                    label="Date column" value={map.splitCol} options={base.headers}
                    noneLabel="— don’t split anything —"
                    note={map.splitCol
                      ? `Adds ${splitTargets(map.splitCol).join(' and ')}, leaving the original alone`
                      : 'A timestamp like 2026-09-09 07:49 is almost unfilterable in Excel'}
                    onChange={(v) => setMap((m) => ({ ...m, splitCol: v }))}
                  />
                </div>
              </div>

              <div className="tune-group">
                <h3>Column matching <span className="tune-auto">detected automatically</span></h3>
                <div className="fields">
                  <Field label="Contacts — contact ID" value={map.baseKey} options={base.headers}
                    onChange={(v) => setMap((m) => ({ ...m, baseKey: v }))} />
                  <Field label="Calls — contact ID" value={map.callKey} options={callCols.concat(map.callKey || [])}
                    note="Not the call’s own Record ID"
                    onChange={(v) => setMap((m) => ({ ...m, callKey: v }))} />
                  <Field label="Calls — date" value={map.dateCol} options={callCols}
                    note="Groups calls into days"
                    onChange={(v) => setMap((m) => ({ ...m, dateCol: v }))} />
                  <Field label="Calls — notes" value={map.notesCol} options={callCols}
                    note="The column that builds up history"
                    onChange={(v) => setMap((m) => ({ ...m, notesCol: v }))} />
                  <Field label="Contacts — project" value={map.projectCol} options={base.headers}
                    note="Powers the per-project export below"
                    onChange={(v) => { setMap((m) => ({ ...m, projectCol: v })); setScope('') }} />
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* --------------------------------------------------------- result */}
      {!result ? (
        <div className="waiting">
          {base || callFiles.length
            ? <>Add the {base ? 'calls' : 'contacts'} file and your merged result appears here.</>
            : <>Your merged file will appear here.</>}
        </div>
      ) : (
        <ResultView
          result={result} map={map} carry={carry} projects={projects}
          scope={scope} onScope={setScope} saving={saving} onSave={save}
          stem={stem} nameCols={nameCols} missingStandard={missingStandard}
          saveError={errors.save}
          onStartOver={baseFile ? () => { setCallFiles([]); load('base', baseFile) } : null}
        />
      )}
    </div>
  )
}
