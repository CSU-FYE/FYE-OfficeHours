# Migration notes

How the two original grid workbooks in `source-data/` were converted into the
standardized `data/office-hours.xlsx`. **Everything the conversion had to guess is
listed here.** Fix anything wrong directly in the workbook — this file is a record,
not an input.

## Sources used

| Source | Sheet | Used |
|---|---|---|
| `FA26_Faculty_GTF Office Hours.xlsx` | `Office Hours` | yes |
| `Fall 2026 LA Office Hours.xlsx` | `Learning Assistants` | yes |
| `Fall 2026 LA Office Hours.xlsx` | `Professor & GTFs` | no — identical to the faculty workbook |
| `Fall 2026 LA Office Hours.xlsx` | `Master Schedule` | **no — see below** |
| `Fall 2026 LA Office Hours.xlsx` | `Sheet2` | no — raw availability survey, an input not a schedule |
| `Fall 2026 LA Office Hours.xlsx` | `Cancellations` | no — header row only, empty |

## Please confirm

1. **Chloe** — role guessed as `gtf`. Appears only as `Chloe` on the Faculty & GTF grid (first seen FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B4). Confirm the full name and whether they are faculty or a GTF.
2. **DrT** — role guessed as `faculty`. Appears only as `DrT` on the Faculty & GTF grid (first seen FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B4). Confirm the full name and whether they are faculty or a GTF.
3. **Taylor** — role guessed as `gtf`. Appears only as `Taylor` on the Faculty & GTF grid (first seen FA26_Faculty_GTF Office Hours.xlsx · Office Hours · C4). Confirm the full name and whether they are faculty or a GTF.
4. **Winnie** — role guessed as `gtf`. Appears only as `Winnie` on the Faculty & GTF grid (first seen FA26_Faculty_GTF Office Hours.xlsx · Office Hours · F4). Confirm the full name and whether they are faculty or a GTF.
5. **Ben** — role guessed as `gtf`. Appears only as `Ben` on the Faculty & GTF grid (first seen FA26_Faculty_GTF Office Hours.xlsx · Office Hours · D5). Confirm the full name and whether they are faculty or a GTF.
6. **Dr. Dan** — role guessed as `faculty`. Appears only as `Dr. Dan` on the Faculty & GTF grid (first seen FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B6). Confirm the full name and whether they are faculty or a GTF.
7. **Ojo** — role guessed as `gtf`. Appears only as `Ojo` on the Faculty & GTF grid (first seen FA26_Faculty_GTF Office Hours.xlsx · Office Hours · E6). Confirm the full name and whether they are faculty or a GTF.
8. **Dr. Harvey** — role guessed as `faculty`. Appears only as `Dr. Harvey` on the Faculty & GTF grid (first seen FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B7). Confirm the full name and whether they are faculty or a GTF.
9. **Dr. Yume** — role guessed as `faculty`. Appears only as `Dr. Yume` on the Faculty & GTF grid (first seen FA26_Faculty_GTF Office Hours.xlsx · Office Hours · C7). Confirm the full name and whether they are faculty or a GTF.
10. **Prof Scheller** — role guessed as `faculty`. Appears only as `Prof Scheller` on the Faculty & GTF grid (first seen FA26_Faculty_GTF Office Hours.xlsx · Office Hours · C8). Confirm the full name and whether they are faculty or a GTF.
11. **Sumaiya** — role guessed as `gtf`. Appears only as `Sumaiya` on the Faculty & GTF grid (first seen FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B10). Confirm the full name and whether they are faculty or a GTF.

### The `Master Schedule` tab was left out

Its names mostly do not appear on the `Learning Assistants` tab — AJ Tripp, Max Goodrich,
Maddie Potter, Trevor Ridge, Ethan Kramer and others appear nowhere else — and it stops at
5pm. It reads like a previous term. **If it is current, say so and it can be merged in.**

It also contains two notes to you: *"Please put GTA beside your name if you are one"* and
*"Chloe Brekhus - LA or GTA?"* — that second question is still open, and matters for the
role filter.

### Names that are probably the same person

The conversion did **not** merge these, because a wrong merge invents a schedule that
nobody holds. If you confirm a pair, just rename in the workbook.

| On the schedule | Possibly | Seen in |
|---|---|---|
| `DrT` | `Dr. Torres` | `Master Schedule` |
| `Chloe` | `Chloe Brekhus` | `Master Schedule` (flagged there as "LA or GTA?") |
| `Prof Scheller` | `Dylan Scheller (GTA)` | `Master Schedule` — if so the role is **gtf**, not faculty |
| `Taylor` | `Nell Taylor` | `Master Schedule` |

`Dr. Dan`, `Dr. Harvey`, `Dr. Yume`, `Winnie`, `Ojo`, `Sumaiya` and `Ben` appear only as
these short forms. Full names would let students search for them by surname.

## What was normalized automatically

- `3pm to 4pm` and `11 am to 12 pm ` (trailing space) → consistent `3:00 PM` / `11:00 AM` times
- `Chloe (AV147) and DrT` → two separate rows; the room moved into its own column
- `Prof Scheller (AV 147) & Dr. Yume`, and names split across two lines in one cell → separate rows
- `Sumaiya - AV147` → name and room separated
- `AV147`, `AV 147` → `AV C147`
- `(GTA)` suffixes stripped from names and turned into the `role` column
- Rooms matching the role default are left blank, so a room change is a one-line edit in `settings`
- Everyone was set to **in-person**. The `Online Office hours` section of the LA sheet was empty.

## Result

- **44 people** — 33 LAs, 6 GTFs, 5 faculty
- **66 shifts**, 0 with a room other than their role default

### Every shift, and the cell it came from

| Person | Day | Start | End | From |
|---|---|---|---|---|
| Chloe | Monday | 10:00 AM | 11:00 AM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B4` |
| DrT | Monday | 10:00 AM | 11:00 AM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B4` |
| Logan Rosebrock | Monday | 10:00 AM | 11:00 AM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · C6` |
| Chloe | Monday | 11:00 AM | 12:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B5` |
| Safwan Ahmad | Monday | 11:00 AM | 12:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · C7` |
| Dr. Dan | Monday | 12:00 PM | 1:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B6` |
| Lydia Iliev | Monday | 12:00 PM | 1:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · C8` |
| Ariana Wright | Monday | 1:00 PM | 2:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · C9` |
| Dr. Harvey | Monday | 1:00 PM | 2:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B7` |
| Kael Spry | Monday | 2:00 PM | 3:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · C10` |
| Sofia Hiller | Monday | 3:00 PM | 4:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · C11` |
| Sumaiya | Monday | 4:00 PM | 5:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B10` |
| Sumaiya | Monday | 5:00 PM | 6:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · B11` |
| Angel Tinoco | Tuesday | 10:00 AM | 11:00 AM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · D6` |
| Taylor | Tuesday | 10:00 AM | 11:00 AM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · C4` |
| Maisy Shull | Tuesday | 11:00 AM | 12:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · D7` |
| Jackson Keating | Tuesday | 12:00 PM | 1:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · D8` |
| Dr. Yume | Tuesday | 1:00 PM | 2:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · C7` |
| Joseph Ramirez | Tuesday | 1:00 PM | 2:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · D9` |
| Dr. Yume | Tuesday | 2:00 PM | 3:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · C8` |
| Mia Schulze | Tuesday | 2:00 PM | 3:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · D10` |
| Prof Scheller | Tuesday | 2:00 PM | 3:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · C8` |
| Prof Scheller | Tuesday | 3:00 PM | 4:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · C9` |
| Sophie DeMatteo | Tuesday | 3:00 PM | 4:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · D11` |
| Ben | Tuesday | 4:00 PM | 5:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · C10` |
| Ben | Tuesday | 5:00 PM | 6:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · C11` |
| Owen Myers | Wednesday | 10:00 AM | 11:00 AM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · E6` |
| Ben | Wednesday | 11:00 AM | 12:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · D5` |
| Kacey Hoang | Wednesday | 11:00 AM | 12:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · E7` |
| Dr. Dan | Wednesday | 12:00 PM | 1:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · D6` |
| Lauren Watts | Wednesday | 12:00 PM | 1:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · E8` |
| Winnie | Wednesday | 12:00 PM | 1:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · D6` |
| Dr. Yume | Wednesday | 1:00 PM | 2:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · D7` |
| River Wysock | Wednesday | 1:00 PM | 2:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · E9` |
| DrT | Wednesday | 2:00 PM | 3:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · D8` |
| Joseph Quiroz Hernandez | Wednesday | 2:00 PM | 3:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · E10` |
| Chloe | Wednesday | 3:00 PM | 4:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · D9` |
| DrT | Wednesday | 3:00 PM | 4:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · D9` |
| Hannah Gruber | Wednesday | 3:00 PM | 4:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · E11` |
| Sumaiya | Wednesday | 5:00 PM | 6:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · D11` |
| Jordy Medina Valverde | Wednesday | 9:00 AM | 10:00 AM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · E5` |
| Parker Bjick | Thursday | 10:00 AM | 11:00 AM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · F6` |
| Taylor | Thursday | 10:00 AM | 11:00 AM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · E4` |
| Austin Monroe | Thursday | 11:00 AM | 12:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · F7` |
| Taylor | Thursday | 11:00 AM | 12:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · E5` |
| Ojo | Thursday | 12:00 PM | 1:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · E6` |
| Tia Fountain | Thursday | 12:00 PM | 1:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · F8` |
| Dr. Harvey | Thursday | 1:00 PM | 2:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · E7` |
| Sarah Goudjil | Thursday | 1:00 PM | 2:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · F9` |
| Dr. Harvey | Thursday | 2:00 PM | 3:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · E8` |
| Prof Scheller | Thursday | 2:00 PM | 3:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · E8` |
| Tucker Cullen | Thursday | 2:00 PM | 3:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · F10` |
| Julie Rickerd | Thursday | 3:00 PM | 4:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · F11` |
| Prof Scheller | Thursday | 3:00 PM | 4:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · E9` |
| Ojo | Thursday | 4:00 PM | 5:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · E10` |
| Ojo | Thursday | 5:00 PM | 6:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · E11` |
| Sylvia Ingegneri | Thursday | 9:00 AM | 10:00 AM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · F5` |
| Mia White | Friday | 10:00 AM | 11:00 AM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · G6` |
| Winnie | Friday | 10:00 AM | 11:00 AM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · F4` |
| Jady Sharp | Friday | 11:00 AM | 12:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · G7` |
| Caleb Adams | Friday | 12:00 PM | 1:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · G8` |
| Winnie | Friday | 12:00 PM | 1:00 PM | `FA26_Faculty_GTF Office Hours.xlsx · Office Hours · F6` |
| Landon Gagliostro | Friday | 1:00 PM | 2:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · G9` |
| Olivia Kalinowski | Friday | 2:00 PM | 3:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · G10` |
| Tito Salcido Rascon | Friday | 3:00 PM | 4:00 PM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · G11` |
| Thomas Miller | Friday | 9:00 AM | 10:00 AM | `Fall 2026 LA Office Hours.xlsx · Learning Assistants · G5` |
