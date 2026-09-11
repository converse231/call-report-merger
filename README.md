# Call Report Merger

Merges a HubSpot **calls** export into a HubSpot **contacts** export.
Takes `.xlsx`, legacy `.xls`, or `.csv` on either side — routed by the file's
own signature, not its extension, so a misnamed file still loads correctly.

```bash
npm install
npm run dev     # open the printed localhost URL
npm test        # merge logic self-check
```

## What it does

Join key: `Associated Contact IDs` (calls) → `Record ID` (contacts).
The calls file's own `Record ID` is the *call's* ID, not the contact's. Note the
94-column export also has `Associated Contact create attribution IDs`, which is
a different thing — detection prefers the exact name.

Only 8 of the calls export's ~94 columns are carried, appended with a `Call ` prefix:

| From the calls export | Added to the base as |
|---|---|
| `Record ID` | `Call Record ID` |
| `Activity assigned to` | `Call Activity assigned to` |
| `Activity date` | `Call Activity date` |
| `Call duration (HH:mm:ss)` | `Call duration (HH:mm:ss)` |
| `Call notes` | `Call notes` |
| `Call Title` | `Call Title` |
| `Strategic Status Code` | `Call Strategic Status Code` |
| `To Number` | `Call To Number` |

The UI shows these as chips so you can verify — or change — the selection per run.

If the calls export is missing one of the eight, step 2 says so by name. That means
the export view in HubSpot doesn't include it — re-export with the column added
rather than substituting something else. `Strategic Status Code` is a custom
property and is easy to leave out; `Call outcome` is HubSpot's built-in field and
is **not** the same thing (Left voicemail / Connected, vs Do Not Call / Initial
Refusal / Scheduled Callback - AC).

For each contact:

1. Group their calls in this export **by day** — a calls export is normally a full
   history, not just today, so every day it contains is used, not only the latest.
2. For each day: join that day's notes, oldest call first, exact duplicates dropped,
   then **add that day on top of the notes already in the base** under a
   `[YYYY-MM-DD]` heading, newest day first. Earlier days already on file stay.
   Re-merging a day replaces only that day's block, so running the same export
   twice is a no-op.
3. Every other call column (status, duration, etc.) takes the **latest non-blank**
   value from the contact's single most recent day. Only notes accumulate — the
   rest hold the latest value, not a history.
4. Results are appended as new `Call …` columns.

## Splitting a date column

A HubSpot timestamp like `2026-09-09 07:49` is unfilterable in Excel — 3,102 contacts
produce ~2,900 distinct values. Pick a base column in step 2 (defaults to `Create Date`)
and two text columns are appended beside it:

| | |
|---|---|
| `Create Date (Date)` | `2026-09-09` — collapses ~2,900 values to ~120 days |
| `Create Date (Time)` | `07:49`, blank when the source had no clock |

The original column is never modified, and the split covers every row, not just
contacts that had calls. Both are written as text in `YYYY-MM-DD` / `HH:MM`, which
sorts and filters correctly without Excel reinterpreting anything.

## Downloading one project

Step 3 has a **Rows to download** picker built from the base's project column
(`Associated Project` by default), listing each project with its row count plus a
`(no project)` bucket. Choosing one narrows both downloads to those rows and tags
the filename. A contact associated with several projects — HubSpot joins them with
`;` — appears under each.

A project subset is written as a **fresh** workbook rather than patched into your
original file, so the base's own date formats and column widths aren't carried
over on that export. The full download is unaffected.

When the base is `.xlsx`, the new columns are written **into the original workbook**:
existing cells keep their own values, date formats and column widths, because they
are never read back and rewritten. `npm test` asserts this cell by cell.

A base file that's legacy `.xls` (or CSV) has no such workbook to write back into,
so the output is a freshly built `.xlsx` instead — values come through correctly,
the original file's own formatting just isn't preserved cell-for-cell since it's
parsed into plain values, not re-opened.

Re-running on an already-merged file writes into the same `Call …` columns rather
than adding duplicates. A contact with no calls in the new export is left alone.

A `Call notes` cell after three days of merging:

```
[2026-09-10]
Appt confirmed

[2026-09-09]
Left voicemail

Called back, interested

[2026-09-08]
Wrong number
```

## Output

- **`.xlsx`** — highlighted: green = new value, amber = replaced a value, blue = new column header.
  New cells are written as text, so 12-digit call IDs can't turn into `1.16574E+11`.
- **`.csv`** — plain, for re-importing to HubSpot. CSV has no cell colours; that's a format limit, not a missing feature.

Everything runs in the browser. No data leaves the machine.
