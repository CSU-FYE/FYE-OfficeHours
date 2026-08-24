# Office Hours Display

A static website that shows students when and where they can get help. The whole
schedule lives in one Excel workbook — `data/office-hours.xlsx` — which the page reads
directly in the browser. There is no database, no build step, and no server code.

The default view answers one question with no clicking: **when can I get help?** Each
half-hour of the week is coloured by *what kind* of help is there — green for faculty or
GTFs, gold for Learning Assistants, split for both — and carries no text at all. Names,
rooms, and courses appear when a student clicks a block. Filters are all opt-in.

The banner matches the First-Year Engineering Calendar (the CSU Calendar App) so the two
sites read as one family: same teal bar, same Poppins, same multicoloured rule. The font
files in `assets/fonts/` are copied from that project.

---

## Updating the schedule

1. Open the shared workbook on OneDrive and make the changes (or have collaborators make them).
2. Download it: **File → Save a Copy / Download a Copy**.
3. Save it into this folder as `data/office-hours.xlsx`, replacing the file that's there.
   The name must match exactly.
4. Commit and push:

```bash
git add data/office-hours.xlsx && git commit -m "Update office hours" && git push
```

GitHub Pages redeploys in about a minute. Then open **`<your-site-url>/?check=1`** — that
page lists anything in the workbook the site couldn't read. It's worth ten seconds after
every update, because bad rows are hidden from students rather than shown as errors.

---

## How to fill in the workbook

**One row per shift.** If someone holds three office hours a week, they get three rows.
Never put two names in one cell — that's what made the old grid schedules impossible to
filter. The workbook has dropdowns on the columns that matter, so most cells are a pick,
not a typo waiting to happen.

### `people` — everyone who holds office hours, listed once

| Column | Required | What goes in it |
|---|---|---|
| `name` | yes | The exact name used on the `shifts` sheet. Must be unique. |
| `display_name` | | What students see, if different (e.g. `Dr. Torres`). Blank = same as `name`. |
| `role` | yes | `faculty`, `gtf`, or `la`. |
| `courses` | | **Blank means the default** (ENGR 111 and ENGR 114). If you fill it in, list *all* courses they cover, separated by `;`. |
| `email` | | Shown to students in the detail panel. |
| `notes` | | Shown to students. Keep it short. |

### `shifts` — one row per weekly office hour

| Column | Required | What goes in it |
|---|---|---|
| `name` | yes | Pick from the dropdown (it reads the `people` sheet). |
| `day` | yes | Monday–Friday. |
| `start`, `end` | yes | `9:00 AM`, `1:30 PM`. Must land on the hour or half hour. |
| `mode` | | `in-person` or `online`. Blank = in-person. |
| `location` | | Blank = the default room for their role. For online, put the meeting link here and it becomes a clickable link. |
| `courses` | | Overrides the person's courses **for this shift only**. Rarely needed. |
| `active` | | `no` hides the row without deleting it. |
| `notes` | | e.g. "first half hour is drop-in only". |

### `exceptions` — one-off changes to the normal week

| To do this | Fill in |
|---|---|
| Cancel one shift | `name`, `date`, `start`, `end`, `type` = `cancelled` |
| Cancel someone's whole day | `name`, `date`, `type` = `cancelled` — leave `start`/`end` blank |
| Add one-off extra hours | `name`, `date`, `start`, `end`, `type` = `added` |

Dates go in as `2026-09-15` or `9/15/2026`. Cancelled hours disappear from the count for
that week; added hours show up tagged "One-off".

### `settings` — term dates, default rooms, and the announcement banner

Change `default_location_la` once and every LA's room updates. Put text in `announcement`
to show a banner across the top of the site (e.g. "No office hours Nov 26–28"); clear it
to remove the banner.

---

## Adding a new person

On the `people` sheet, add a row with their name and role. That's it — their name is then
available in the `shifts` dropdown. They cover ENGR 111 and ENGR 114 unless you say
otherwise in `courses`.

## Adding a new course

Nothing to configure. Type it in a person's `courses` cell and it appears as a filter
option on the site automatically.

---

## Previewing locally

Browsers won't let a page opened straight from a folder read the workbook, so double-clicking
`index.html` shows a "could not load" message with a file picker — you can use that picker to
preview any `.xlsx` without committing it. To see the real thing, serve the folder:

```bash
python3 -m http.server 8777
```

Then open <http://localhost:8777>.

## Publishing to GitHub Pages

Once, at setup: push this folder to a GitHub repository, then in the repo go to
**Settings → Pages** and set **Source** to *Deploy from a branch*, branch `main`, folder
`/ (root)`. The site appears at `https://<username>.github.io/<repo>/` a minute later.
After that, every `git push` republishes it.

---

## What's in this folder

| Path | What it is |
|---|---|
| `index.html` | The page. Also serves the check report at `?check=1`. |
| `assets/styles.css` | All styling, light and dark. |
| `assets/fonts/` | Poppins, copied from the First-Year Engineering Calendar so the banner matches. |
| `assets/xlsx.js` | Reads `.xlsx` files in the browser. No libraries — the browser can already unzip and parse XML. |
| `assets/data.js` | Turns the workbook into the schedule model, and collects anything wrong with it. |
| `assets/app.js` | Draws the grid, the filters, and the detail panel. |
| `data/office-hours.xlsx` | **The schedule.** The only file you normally touch. |
| `source-data/` | The two original grid workbooks, kept for reference. Nothing reads them. |
| `tools/make_workbook.py` | One-time script that converted those originals into the workbook. You don't need to run it. |
| `MIGRATION-NOTES.md` | What that conversion had to guess. **Worth reading once.** |

## Notes on how it behaves

- The site shows **the current week**, with real dates in the column headers. On weekends
  it shows the week ahead.
- Consecutive hours by the same person merge into one block, so a 4–6 PM shift reads as
  one block rather than two.
- Block colour follows whatever is *currently filtered*. Filter to Learning Assistants and
  every block turns gold, because that is all that is left.
- Blocks carry no visible text, so the colour is also written into each block's hover
  tooltip and its screen-reader label — colour is never the only way to read the grid.
- A filter with only one possible value stays hidden — the **Format** filter appears once
  somebody actually has online hours.
- Filter choices are stored in the URL, so a filtered view can be linked to directly:
  `.../#course=ENGR+114&mode=online` links straight to online ENGR 114 help.
- Rows the site can't read are hidden from students, never shown as errors. `?check=1` is
  where you find them.
