# Handoff — FYE Office Hours

A static site showing students when and where they can get help — ENGR office hours in
AV C144, and tutoring in AV C141 across chemistry, physics, maths, CS, CIVE and biology.
Live at **<https://bengrier.github.io/FYE-OfficeHours/>**, source in
`bengrier/FYE-OfficeHours`, working copy at
`~/Desktop/CSU Work/AI App Creation/Office Hours Display`.

Everything the site knows comes from one Excel workbook, `data/office-hours.xlsx`, which the
page reads **in the browser**. No build step, no server code, no dependencies, no npm.

---

## The one-paragraph version

Ben maintains a shared workbook on OneDrive. He downloads it over
`data/office-hours.xlsx`, commits, pushes; GitHub Pages redeploys. The page fetches that
`.xlsx`, unzips and parses it in JavaScript, and draws a Monday–Sunday 9am–9pm grid. The
default view shows **when help exists and in which room**, with no names and no counts.
Clicking any hour reveals who is there. Filters (Get help with / Role / Where / Person) are
all opt-in and mirror into the URL.

---

## Architecture

| File | Does |
|---|---|
| `index.html` | Markup shell. Also serves the check report at `?check=1`. An inline pre-paint script adds `.embed` to `<html>` for `?embed=1` (Canvas iframe: no header, no "right now", no footer — CSS only, at the bottom of `styles.css`). |
| `assets/xlsx.js` | Reads `.xlsx` with no libraries: parses the zip central directory, inflates with `DecompressionStream('deflate-raw')`, reads sheet XML with `DOMParser`. |
| `assets/data.js` | Workbook → model. All the forgiving parsing lives here (times, days, roles, dates, courses). Anything unparseable becomes a *problem*, never an exception. |
| `assets/app.js` | Grid, filters, detail panel, "right now" strip, check report. |
| `assets/styles.css` | Everything visual. Light/dark. Room palette + textures. |
| `tools/make_workbook.py` | **One-time** migration from the original grid schedules. Ben never runs it — and it predates tutoring, so `add_tutoring.py` must follow it if he ever does. |
| `tools/add_tutoring.py` | Folds the two tutoring grids into the workbook **in place**, so the yellow role flags and every hand edit survive. Re-runnable; rebuilds only the tutoring rows. |
| `tools/serve.py` | Local preview server. **Use this, not `python3 -m http.server`** — see Gotchas. |
| `source-data/` | Original grid workbooks. **Gitignored** (see Privacy). |

Data flows one way: `xlsx.js` → `data.js` (`loadModel`) → `app.js` renders. `data.js` never
touches the DOM; `app.js` never parses spreadsheet values.

---

## The data contract

Four sheets, plus a "how to use this" tab written for collaborators.

- **`people`** — `name` (the key), `display_name`, `role` (`faculty`/`gtf`/`la`), `courses`,
  `email`, `notes`. Blank `courses` means the default.
- **`shifts`** — one row per weekly shift: `name`, `day`, `start`, `end`, `mode`,
  `location`, `courses`, `active`, `notes`, `program`. **Three shifts = three rows.** Never
  two names in a cell — that is the whole reason the old format had to go.
- **`exceptions`** — `name`, `date`, `start`/`end`, `type` (`cancelled`/`added`), `mode`,
  `location`, `note`. Blank times on a cancellation kills that person's whole day.
- **`settings`** — key/value. This is the control panel; prefer adding a setting over
  hardcoding.

**Programmes.** `shifts.program` is blank (office hours) or `tutoring`. It decides two
things that a person cannot: which room the shift defaults to
(`default_location_<program>`, which beats `default_location_<role>`), and which of the
person's course lists it advertises (`people.courses_<program>`, falling back to
`people.courses`). This exists because 26 of the 30 tutors are *also* ENGR Learning
Assistants — the same person, in a different room, helping with different courses. Role
stays what they are; the programme says which hat this row is. Nothing about it is
hardcoded to tutoring: any `courses_<name>` column on `people` is picked up.

Current settings worth knowing:

```
default_location_faculty|gtf|la   AV C144        every in-person office hour
default_location_tutoring         AV C141        the whole tutoring programme
default_program                   office hours   what a blank program cell means
program_order                     office hours; tutoring
room_order        AV C144; AV C141; Scott Engineering; Online
course_implies    ENGR 111 -> ENGR 123
buildings         AV -> Academic Village
subjects          CHEM -> Chemistry; PH -> Physics; ...   groups the course menu
course_order      ENGR first, then the tutoring grid's subjects  orders the course menu
day_start/day_end 9:00 AM / 9:00 PM
```

`course_implies` and `buildings` share one arrow syntax parsed by
`parseMappingRules` in `data.js`: `A -> B, C; D -> E`.

**No shift row carries an explicit room, and no tutoring row carries explicit courses.**
They are all blank and resolve through `default_location_*` and `courses_tutoring`, so a
room change or a tutor's new subject is a one-cell edit. The `(AV 147)` text still
littering the *source* grids is stale — ignore it.

---

## Design decisions worth not re-litigating

Each of these was asked for deliberately, usually after seeing the alternative.

1. **The grid shows availability and room, never names or head counts.** Reducing cognitive
   load was the founding requirement.
2. **A block is one outlined shape spanning an unbroken stretch of availability**, so
   students cannot read individual shift times off the grid. The Person filter is what
   reveals an individual's hours.
3. **Colour bands at 30-minute resolution inside that shape.** Painting the whole stretch
   with the union of its rooms was a lie. A band boundary means the room mix genuinely
   changed; one person handing over to another in the same room draws no seam.
4. **Colour = room, not role.** Roles live in the Role filter and the panel badges, and
   every role badge is the same neutral — faculty used to be green, which quietly ranked
   the people in a panel. The only coloured badge left is `One-off`, which marks an hour,
   not a person.
5. **Every room past the first also carries a texture** — hairline diagonals, dots — so the
   grid survives colour blindness and greyscale. Hue carries identity, texture confirms it.
   Textures are hairlines (1px on a 6px pitch), not bands; the chunky version was rejected.
6. **Clicking is hour by hour**, not by shift boundary — snapping the outline to shifts
   would let someone map the rota by clicking around. The clicked hour gets a double-ring
   outline.
7. **Desktop panel is a non-modal side rail**; the grid stays lit so you can step along the
   hours. On a phone it is a modal sheet and the grid becomes a day-at-a-time list.
8. **Every filter option names something that really has hours behind it.** Where is derived
   from buildings in use, so it cannot offer an empty place. Filters with one option hide.
   *Get help with* ANDs across courses — two courses means hours covering both — while the
   other three filters are unions, since a shift has only one room and one person.
9. **No dates anywhere** — the schedule repeats weekly. Exceptions still apply to the real
   current week, just undated on screen.
10. **A day nothing ever happens on keeps its column**, narrowed to 40px, hatched, and
    labelled *Closed* — asked for explicitly when Sunday tutoring arrived. A gap where a
    day should be is a question; "closed" is an answer. It is computed from the *whole*
    schedule, never the filtered one: a course filter that emptied Friday and collapsed
    its column would move every other column under the reader's cursor.
11. **The programme split is a maintenance concept, never a student-facing one.** It has
    to exist in the workbook, because someone can be a Learning Assistant in C144 on
    Tuesday and a tutor in C141 on Wednesday, and the room and course list follow from
    which. But a student looking for help does not care who staffs the hour: the room
    says where to walk and the courses say whether it is the right help. There was a
    "Kind of help" filter and a panel line for a few hours; both were cut on sight. The
    split surfaces on `?check=1` and nowhere else.
12. **The narrow view opens on today**, read off `week`'s own `isToday` rather than
    recomputed from the clock, so the tab and the column marked TODAY cannot disagree.
    On a Saturday that means opening on a closed day saying so.
13. **Bad rows are hidden from students, never shown as errors.** `?check=1` is where they
    surface, with the exact sheet and row number.

---

## Gotchas that cost real time

- **Local caching.** `python3 -m http.server` sends *no* `Cache-Control`, so browsers serve
  a stale `assets/data.js` after you edit it and you debug a change that never loaded. This
  happened. Use `python3 tools/serve.py`, which sends `no-store`. If a page still looks
  stale, check `performance.getEntriesByType('resource')` for `transferSize: 0`, and force
  with `fetch(url, {cache:'reload'})` then reload. GitHub Pages sends a real `max-age=600`,
  so production is fine — code changes reach students within ten minutes.
- **Palette slots are pinned by `room_order`, including rooms with no current hours.**
  Assigning slots from rooms *in use* meant deleting a room shifted every other room's
  colour. Do not "simplify" that back.
- **`?check=1` after every workbook change.** It is the only gate; nothing validates at
  commit time.
- **Excel coerces typed times to numbers.** `data.js` handles the numeric forms; the time
  columns are also formatted as text at the column level (not per cell — that left 400
  styled empty cells and sent Ctrl+End to row 400).
- **Screenshots via the browser tool go stale.** Trust DOM measurements over a screenshot
  when they disagree; navigate with `force: true` to refresh the capture.

---

## Privacy

The repo is **public**. `source-data/` is gitignored *and was purged from git history*
before the first push, because the original LA workbook's `Sheet2` held raw availability
survey responses with individual students' constraints and free-text comments. Do not
commit that folder. The published workbook holds names, roles, rooms and Teams links only —
which is the point of the site.

---

## Current state (25 Aug 2026)

48 people (33 LA, 6 faculty, 5 GTF, 4 tutor-only), 138 shifts — 65 office hours, 73
tutoring — 4 of them online with Teams links. Rooms in use: AV C144, AV C141, Online.
25 courses across 7 subjects. The week runs Monday–Sunday; Saturday is the one closed
column. `?check=1` reports zero errors and one warning, which is real (see below).

## Open items

0. **Three things the tutoring sources leave hanging**, all reported by `add_tutoring.py`
   or `?check=1` rather than papered over:
   - **Sunday has one named tutor.** River Wysock covers 5–8 PM; the second seat in all
     three hours is `*2nd Tutor TBD*` on the rota and is simply absent from the site
     until a name is filled in.
   - **Christine Parkin is on the Thursday 5–6 PM rota twice** (FINAL rows 11 and 13).
     The site counts her once and `?check=1` warns. One of those two cells probably meant
     to be someone else.
   - **Tom Brown and Izobel M** are on the Classes Grid with no subjects ticked and no
     shifts, so they are in the workbook as `tutor` and invisible to students. Harmless,
     but they are placeholders.

1. **Ten roles still guessed.** `Chloe`, `Dr. Dan`, `Dr. Harvey`, `Dr. Yume`, `DrT`, `Ojo`,
   `Prof Scheller`, `Sumaiya`, `Taylor`, `Winnie` — the source grids only ever gave a first
   name or a title. Their `role` cells are highlighted yellow in the workbook with a comment
   explaining why. Full names would also let students search by surname. Likely identities
   (unconfirmed, never merged automatically): `DrT`→Dr. Torres, `Chloe`→Chloe Brekhus,
   `Taylor`→Nell Taylor, `Prof Scheller`→Dylan Scheller *(who is flagged as a GTA in the
   source, so this role may be wrong)*. See `MIGRATION-NOTES.md`.
2. **23 light-blue cells** (`#C0E6F5`) on the source `Professor & GTFs` sheet, with no key
   anywhere and no hyperlinks. Never acted on. Ask what they mean.
3. **`Master Schedule`** in the source is hidden and its names mostly match nothing else —
   treated as a previous term. Confirm it can be ignored for good.
4. Ben confirmed himself as **faculty**; `CONFIRMED_ROLES` in `make_workbook.py` records
   that so a re-run cannot reintroduce the guess. Add to it as names are confirmed.

## Conventions

- Vanilla JS, ES modules, no dependencies. Keep it that way.
- Prefer a `settings` row over a constant in code.
- Comments explain *why*, especially where a simpler-looking approach was tried and failed.
- Commit messages say what changed and what it fixes.
- Verify in the browser and quote real numbers; do not ask Ben to check manually.
