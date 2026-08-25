# Handoff — FYE Office Hours

A static site showing students when and where they can get help with ENGR 111/114/123.
Live at **<https://bengrier.github.io/FYE-OfficeHours/>**, source in
`bengrier/FYE-OfficeHours`, working copy at
`~/Desktop/CSU Work/AI App Creation/Office Hours Display`.

Everything the site knows comes from one Excel workbook, `data/office-hours.xlsx`, which the
page reads **in the browser**. No build step, no server code, no dependencies, no npm.

---

## The one-paragraph version

Ben maintains a shared workbook on OneDrive. He downloads it over
`data/office-hours.xlsx`, commits, pushes; GitHub Pages redeploys. The page fetches that
`.xlsx`, unzips and parses it in JavaScript, and draws a Monday–Friday 9am–9pm grid. The
default view shows **when help exists and in which room**, with no names and no counts.
Clicking any hour reveals who is there. Filters (Get help with / Role / Where / Person) are
all opt-in and mirror into the URL.

---

## Architecture

| File | Does |
|---|---|
| `index.html` | Markup shell. Also serves the check report at `?check=1`. |
| `assets/xlsx.js` | Reads `.xlsx` with no libraries: parses the zip central directory, inflates with `DecompressionStream('deflate-raw')`, reads sheet XML with `DOMParser`. |
| `assets/data.js` | Workbook → model. All the forgiving parsing lives here (times, days, roles, dates, courses). Anything unparseable becomes a *problem*, never an exception. |
| `assets/app.js` | Grid, filters, detail panel, "right now" strip, check report. |
| `assets/styles.css` | Everything visual. Light/dark. Room palette + textures. |
| `tools/make_workbook.py` | **One-time** migration from the original grid schedules. Ben never runs it. |
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
  `location`, `courses`, `active`, `notes`. **Three shifts = three rows.** Never two names
  in a cell — that is the whole reason the old format had to go.
- **`exceptions`** — `name`, `date`, `start`/`end`, `type` (`cancelled`/`added`), `mode`,
  `location`, `note`. Blank times on a cancellation kills that person's whole day.
- **`settings`** — key/value. This is the control panel; prefer adding a setting over
  hardcoding.

Current settings worth knowing:

```
default_location_faculty|gtf|la   AV C144        every in-person hour, confirmed 25 Aug
room_order    AV C144; AV C141; Scott Engineering; Online
course_implies    ENGR 111 -> ENGR 123
buildings         AV -> Academic Village
day_start/day_end 9:00 AM / 9:00 PM
```

`course_implies` and `buildings` share one arrow syntax parsed by
`parseMappingRules` in `data.js`: `A -> B, C; D -> E`.

**No shift row carries an explicit room.** They are all blank and resolve to
`default_location_*`, so a room change is a one-cell edit. The `(AV 147)` text still
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
4. **Colour = room, not role.** Roles live in the Role filter and the panel badges.
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
9. **No dates anywhere** — the schedule repeats weekly. Exceptions still apply to the real
   current week, just undated on screen.
10. **Bad rows are hidden from students, never shown as errors.** `?check=1` is where they
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

44 people (33 LA, 6 faculty, 5 GTF), 65 shifts, 4 of them online with Teams links.
Rooms in use: AV C144 and Online. `?check=1` reports zero problems. Demo rows removed.

## Open items

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
