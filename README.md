# Office Hours Display

A static website that shows students when and where they can get help. It covers two
things: **ENGR office hours** in AV C144, and **tutoring** in AV C141 — chemistry,
physics, maths, computer science, civil engineering and biology. The whole schedule
lives in one Excel workbook — `data/office-hours.xlsx` — which the page reads directly
in the browser. There is no database, no build step, and no server code.

The default view answers one question with no clicking: **when can I get help?** Each
half-hour of the week is coloured by **which room** to walk to, and carries no text at
all. Names, courses, and who is on duty appear when a student clicks an hour. Filters are
all opt-in.

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

### If a fresh Faculty & GTF grid arrives instead

The department maintains those hours as a grid, not as this workbook. When a new copy
lands, save it into `source-data/` as `FA26_Faculty_GTF Office Hours.xlsx` (replacing what's
there) and run:

```bash
python3 tools/update_faculty_hours.py
```

It prints exactly which hours it added and removed, then rewrites only the rows of the
people named on that grid — LA hours, tutoring, roles and notes are left alone. Add
`--dry-run` to see the diff without touching anything. Then commit `data/office-hours.xlsx`
and check `?check=1` as above.

One thing to watch: this edits the copy in this folder, not the workbook on OneDrive. Upload
the result back over the OneDrive copy afterwards, or the next download from OneDrive will
put the old faculty hours back.

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
| `role` | yes | `faculty`, `gtf`, or `la`. Everyone who tutors is an `la`, whether or not they also hold ENGR office hours — there is no separate tutor role. |
| `courses` | | What they help with **at office hours**. **Blank means the default** (ENGR 111 and ENGR 114). If you fill it in, list *all* courses they cover, separated by `;`. Anyone covering ENGR 111 is counted for ENGR 123 automatically — see `course_implies`. |
| `courses_tutoring` | | What they help with **at tutoring**, separated by `;`. Blank means they do not tutor. This is the only place to edit it — every one of their tutoring shifts reads from this one cell. |
| `email` | | Shown to students in the detail panel. |
| `notes` | | Shown to students. Keep it short. |

### `shifts` — one row per weekly hour

| Column | Required | What goes in it |
|---|---|---|
| `name` | yes | Pick from the dropdown (it reads the `people` sheet). |
| `day` | yes | Monday–Sunday. A day with no rows shows as a narrow, crossed-out **Closed** column. |
| `start`, `end` | yes | `9:00 AM`, `1:30 PM`. Must land on the hour or half hour. |
| `program` | | `office hours` or `tutoring`. Blank = office hours. Decides both the room and which of the person's two course lists the shift advertises. **Students never see it** — no filter, no badge, no panel line; it shows on `?check=1` only. |
| `mode` | | `in-person` or `online`. Blank = in-person. |
| `location` | | Blank = the default room for the programme (`default_location_tutoring`), falling back to the default for their role. For online, put the meeting link here and it becomes a clickable link. |
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

Change `default_location_la` once and every LA's room updates; `default_location_tutoring`
does the same for the whole tutoring programme. Put text in `announcement` to show a
banner across the top of the site (e.g. "No office hours Nov 26–28"); clear it to remove
the banner.

Two settings shape the **Get help with** menu, which is now two dozen courses long:
`subjects` groups them under headings (`CHEM -> Chemistry; …`) and `course_order` sets
the order they are listed in. A course named in neither still appears — alphabetically,
at the end, under its own bare subject code.

---

## Adding a new person

On the `people` sheet, add a row with their name and role. That's it — their name is then
available in the `shifts` dropdown. They cover ENGR 111 and ENGR 114 unless you say
otherwise in `courses`. If they also tutor, fill in `courses_tutoring` and give their
tutoring shifts `program` = `tutoring`.

## Adding a new course

Nothing to configure. Type it in a person's `courses` cell and it appears as a filter
option on the site automatically.

If a course always comes along with another one, say so once in `settings` under
`course_implies` rather than editing 44 rows. It reads
`ENGR 111 -> ENGR 123`, meaning anyone who covers ENGR 111 is also listed under
ENGR 123 — including everyone on the blank-cell default. Several rules are separated by
`;`, and a rule can imply more than one course: `ENGR 111 -> ENGR 123, ENGR 124`.
`?check=1` prints the rules that are in effect.

---

## Previewing locally

Browsers won't let a page opened straight from a folder read the workbook, so double-clicking
`index.html` shows a "could not load" message with a file picker — you can use that picker to
preview any `.xlsx` without committing it. To see the real thing, serve the folder:

```bash
python3 tools/serve.py
```

Then open <http://localhost:8777>. Use that rather than `python3 -m http.server`: the stock
server sends no cache headers, so a browser will serve a stale `app.js` after you edit it and
you end up debugging a change that never loaded. GitHub Pages sends a real `max-age`, so this
only affects local previews.

## Embedding in Canvas

Add `?embed=1` to the URL and the page drops the header, the "right now" box, and the footer,
leaving the filters and the grid. Paste this into a Canvas page with the HTML editor (the
`</>` button):

```html
<p><iframe src="https://bengrier.github.io/FYE-OfficeHours/?embed=1"
   title="ENGR 111/114/123 office hours" width="100%" height="900"
   style="border: 1px solid #ddd; border-radius: 8px;"></iframe></p>
```

Canvas iframes don't grow to fit their contents, so pick a `height` and check it — 900 shows
the whole 9am–9pm grid on a laptop. Filters still write to the URL inside the frame, so you
can link a pre-filtered view: `?embed=1#course=ENGR%20123`. Anything after `#` is filters;
`?embed=1` has to come before it.

Students on phones get the day-by-day list instead of the grid, same as the full page.

## Publishing to GitHub Pages

Once, at setup: push this folder to a GitHub repository, then in the repo go to
**Settings → Pages** and set **Source** to *Deploy from a branch*, branch `main`, folder
`/ (root)`. The site appears at `https://<username>.github.io/<repo>/` a minute later.
After that, every `git push` republishes it.

---

## What's in this folder

| Path | What it is |
|---|---|
| `index.html` | The page. Also serves the check report at `?check=1` and the bare embed at `?embed=1`. |
| `assets/styles.css` | All styling, light and dark. |
| `assets/fonts/` | Poppins, copied from the First-Year Engineering Calendar so the banner matches. |
| `assets/xlsx.js` | Reads `.xlsx` files in the browser. No libraries — the browser can already unzip and parse XML. |
| `assets/data.js` | Turns the workbook into the schedule model, and collects anything wrong with it. |
| `assets/app.js` | Draws the grid, the filters, and the detail panel. |
| `data/office-hours.xlsx` | **The schedule.** The only file you normally touch. |
| `source-data/` | The two original grid workbooks, kept for reference. Nothing reads them. |
| `tools/make_workbook.py` | One-time script that converted those originals into the workbook. You don't need to run it. **If you ever do, run `add_tutoring.py` straight after** — `make_workbook.py` predates tutoring and writes a workbook without it. |
| `tools/add_tutoring.py` | Rebuilds the tutoring rows from the two tutoring grids in `source-data/`. Re-runnable, and it never touches an office-hours row. |
| `tools/update_faculty_hours.py` | Folds a fresh Faculty & GTF grid into the workbook. Run it when a new copy of that grid arrives — see above. |
| `MIGRATION-NOTES.md` | What that conversion had to guess. **Worth reading once.** |
| `HANDOFF.md` | Orientation for a developer (or a new AI chat) picking this up cold. |

## Notes on how it behaves

- The site shows **the Monday-to-Sunday week containing today**. It used to roll the
  weekend forward to the following week; Sunday tutoring ended that, because on a
  Saturday the next thing that happens is tomorrow evening.
- **On a phone the day tabs open on today**, whichever day that is — including
  Saturday, where the list says the day is closed.
- **A day with nothing on it all term keeps its column**, narrowed and crossed out and
  labelled *Closed*. A gap where a day should be is a question; an explicit "closed" is
  an answer. This is read from the whole schedule, never the filtered one, so choosing
  a course cannot make Friday collapse.
- Consecutive hours by the same person merge into one block, so a 4–6 PM shift reads as
  one block rather than two.
- **Rooms colour themselves.** Type a new room into a `location` cell and it gets its own
  colour, a legend entry, and a **Where** option with no code change. Every other room in the palette also
  carries a texture — stripes, dots — so two rooms stay distinguishable in greyscale or
  to a colour-blind reader; hue is never the only difference. `room_order` in `settings`
  pins which room gets which colour, so "green is C144" stays true as the schedule grows.
  A room named there **keeps its colour even while it has no hours**, so a room going quiet
  for a term does not reshuffle the colours of the ones that remain.
- A half hour split across two rooms shows both colours side by side.
- Block colour follows whatever is *currently filtered*, and the legend narrows with it.
- A block is one outlined shape covering an unbroken stretch of availability, with the
  colour **banded per half hour** inside it. A band boundary means the room actually
  changed; one person handing over to another in the same room draws no seam, so
  individual shift times still are not readable from the grid. The Person filter is what
  reveals those.
- Even so, each block is clickable **hour by hour**. Clicking outlines that hour and the
  panel lists only who is there then. On a desktop the panel is a side rail and the grid
  stays live, so you can step along the hours; on a phone it is a sheet.
- Blocks carry no visible text, so the colour is also written into each block's hover
  tooltip and its screen-reader label — colour is never the only way to read the grid.
- **Every filter option names something that really has hours behind it.** Where is built
  from the buildings actually in use, so it never offers a place nobody is sitting in. The
  `buildings` setting groups rooms that share a building (`AV -> Academic Village`); a room
  with no rule is its own building, so somewhere new needs no rule to show up. Filters with
  only one possible value hide entirely. There is no separate in-person/online filter —
  Where covers it.
- **Picking two courses narrows, it does not widen.** *Get help with* ANDs: choose ENGR 111
  and CHEM 111 and you get only the hours where one person covers both, which is the
  question a student with two hard classes is actually asking. Role, Where, and Person
  stay unions — a shift has one room and one person, so ANDing those would always empty
  the grid.
- **Online hours are a "room"** as far as colour goes: they get their own swatch and legend
  entry. Put the meeting URL in the `location` cell and the panel turns it into a
  *join the meeting* link.
- Filter choices are stored in the URL, so a filtered view can be linked to directly:
  `.../#course=ENGR+114&mode=online` links straight to online ENGR 114 help.
- Rows the site can't read are hidden from students, never shown as errors. `?check=1` is
  where you find them.
