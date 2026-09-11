import React from 'react'
import { renderToString } from 'react-dom/server'
import App, { ResultView } from './App.jsx'
import { merge, groupByProject, detectSplitDateColumn } from './merge.js'

// React SSR separates adjacent text nodes with <!-- --> markers
const flat = (h) => h.replace(/<!--\s*-->/g, '')
const must = (html, needle, what) => {
  if (!flat(html).includes(needle)) throw new Error(`${what}: expected ${JSON.stringify(needle)}`)
}

export function run() {
  const empty = renderToString(React.createElement(App))
  must(empty, 'Call Report Merger', 'empty state')
  must(empty, 'Your merged file will appear here', 'empty state')

  const baseHeaders = ['Record ID', 'First Name', 'Last Name', 'Create Date', 'Associated Project']
  const result = merge({
    baseHeaders,
    baseRows: [
      { 'Record ID': '111', 'First Name': 'Abby', 'Last Name': 'Miller', 'Create Date': new Date(Date.UTC(2026, 8, 9, 7, 49)), 'Associated Project': 'Blue Coats' },
      { 'Record ID': '222', 'First Name': 'Sam', 'Last Name': 'Ortega', 'Create Date': '8/3/2026 17:06', 'Associated Project': '' },
    ],
    callHeaders: ['Associated Contact IDs', 'Activity date', 'Call notes', 'Strategic Status Code'],
    callRows: [
      { 'Associated Contact IDs': '111', 'Activity date': '9/8/2026 15:01', 'Call notes': 'left a voicemail', 'Strategic Status Code': 'No Answer' },
      { 'Associated Contact IDs': '111', 'Activity date': '8/24/2026 12:30', 'Call notes': 'she is interested', 'Strategic Status Code': '' },
      { 'Associated Contact IDs': '999', 'Activity date': '9/8/2026 15:01', 'Call notes': 'orphan', 'Strategic Status Code': '' },
    ],
    baseKey: 'Record ID', callKey: 'Associated Contact IDs', dateCol: 'Activity date', notesCol: 'Call notes',
    splitDateColumns: ['Create Date'],
  })

  const map = { baseKey: 'Record ID', callKey: 'Associated Contact IDs', dateCol: 'Activity date', notesCol: 'Call notes', splitCol: 'Create Date', projectCol: 'Associated Project' }
  const projects = groupByProject(result.rows, 'Associated Project')

  const view = (over = {}) => renderToString(React.createElement(ResultView, {
    result, map, carry: ['Call notes', 'Strategic Status Code'], projects,
    scope: '', onScope: () => {}, saving: '', onSave: () => {},
    stem: 'contacts-2026-09-10', nameCols: ['First Name', 'Last Name'],
    missingStandard: [], saveError: null, onStartOver: () => {}, ...over,
  }))

  const html = view()
  must(html, '1 contact updated', 'verdict')
  must(html, 'contacts-2026-09-10-merged.xlsx', 'filename preview')
  must(html, 'Blue Coats', 'project pill')
  must(html, 'All contacts', 'all-rows pill')
  must(html, 'Abby Miller', 'changed row name')
  must(html, '2026-09-08', 'dated note entry')
  must(html, 'left a voicemail', 'note text')
  must(html, '1 IDs not in contacts', 'unmatched tally')
  if (flat(html).includes('[object Object]')) throw new Error('a raw object leaked into the markup')
  if (flat(html).includes('undefined')) throw new Error('"undefined" leaked into the markup')

  // scoped export renames the file and narrows the row count
  const scoped = view({ scope: 'Blue Coats' })
  must(scoped, 'contacts-2026-09-10-merged-Blue_Coats.xlsx', 'scoped filename')
  must(scoped, '1 rows', 'scoped row count')

  // the nothing-changed path
  const none = merge({
    baseHeaders, baseRows: [{ 'Record ID': '111' }],
    callHeaders: ['Associated Contact IDs', 'Activity date', 'Call notes'], callRows: [],
    baseKey: 'Record ID', callKey: 'Associated Contact IDs', dateCol: 'Activity date', notesCol: 'Call notes',
  })
  const noneHtml = renderToString(React.createElement(ResultView, {
    result: none, map, carry: [], projects: new Map(), scope: '', onScope: () => {}, saving: '',
    onSave: () => {}, stem: 'x', nameCols: [], missingStandard: ['Strategic Status Code'],
    saveError: 'boom', onStartOver: null,
  }))
  must(noneHtml, 'Nothing changed', 'empty verdict')
  must(noneHtml, 'boom', 'save error')
  must(noneHtml, 'Strategic Status Code', 'missing-column warning')

  console.log('rendered: empty, result, scoped export, and nothing-changed — all clean')
}
