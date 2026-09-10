# Call Report Merger

Merges a HubSpot **calls** export into a HubSpot **contacts** export.
Takes `.xlsx` or `.csv` on either side.

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

1. Find their **most recent call day**. Earlier days are ignored.
2. Join every note from that day, oldest first, exact duplicates dropped, then
   **add that day on top of the notes already in the base** under a `[YYYY-MM-DD]`
   heading. Earlier days stay. Re-merging a day replaces only that day's block, so
   running the same export twice is a no-op.
3. Every other call column takes the **latest non-blank** value from that day.
   Only notes accumulate — the rest hold the latest value, not a history.
4. Results are appended as new `Call …` columns.

When the base is `.xlsx`, the new columns are written **into the original workbook**:
existing cells keep their own values, date formats and column widths, because they
are never read back and rewritten. `npm test` asserts this cell by cell over all
3,083 rows of the sample export.

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
