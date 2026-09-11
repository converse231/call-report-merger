// Renders the real components headlessly. `npm test` runs it after merge.test.mjs.
import fs from 'node:fs'
import esbuild from 'esbuild'
import { createRequire } from 'node:module'

const out = 'src/.ui-probe.tmp.cjs'
await esbuild.build({
  entryPoints: ['src/ui.probe.jsx'],
  bundle: true, format: 'cjs', outfile: out,
  jsx: 'automatic', loader: { '.css': 'empty' }, logLevel: 'error',
})
try {
  createRequire(import.meta.url)(`./${out.split('/').pop()}`).run()
} finally {
  fs.unlinkSync(out)
}
