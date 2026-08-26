/*
 * Office Hours — reads data/office-hours.xlsx and draws the week.
 *
 * The default view answers one question with no interaction: when can I get help?
 * So the grid shows counts only. Names, rooms and courses live one click deep,
 * and the filters are all opt-in.
 */

import {
  loadModel, availabilityAt, parseMappingRules, formatTime, formatRange,
  DAYS, ROLE_LABELS, SLOT_MINUTES, dateKey, subjectOf,
} from "./data.js";

const WORKBOOK = "data/office-hours.xlsx";
const SLOT_H = 28;
// Breathing room above the first hour line so the 9 AM label is not clipped.
// Every absolutely positioned thing in the grid measures from here.
const GRID_PAD = 12;
const slotTop = (slot) => GRID_PAD + slot * SLOT_H;

const $ = (id) => document.getElementById(id);
const el = (tag, className, text) => {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
};

let model = null;
let week = [];
let activeDay = 0;
// Which hour the reader has picked, so the grid can outline it. Kept out of the
// block geometry on purpose: the block stays one shape, the outline moves.
let selection = null;
let roomIndex = new Map();
let buildingRules = [];
// Deliberately no `program`. Office hours and tutoring are one thing to a student
// looking for help: the room says where to walk, the courses say whether it is the
// right help, and which staffing programme pays for the hour is ours, not theirs.
const filters = { course: new Set(), role: new Set(), building: new Set(), person: new Set() };

/* ------------------------------------------------------------------ boot */

init();

async function init() {
  try {
    const response = await fetch(`${WORKBOOK}?_=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
    await start(await response.arrayBuffer(), response.headers.get("last-modified"));
  } catch (error) {
    offerFilePicker(error);
  }
}

async function start(buffer, lastModified) {
  model = await loadModel(buffer);
  week = currentWeek();
  // Read off the week rather than recomputed from the clock, so the tab that opens
  // and the column marked TODAY can never disagree.
  activeDay = week.find((day) => day.isToday)?.index ?? 0;

  document.title = model.settings.term_name
    ? `Office Hours · ${model.settings.term_name}`
    : "Office Hours";
  $("term").textContent = model.settings.term_name;
  $("updated").textContent = lastModified
    ? `Schedule updated ${new Date(lastModified).toLocaleDateString(undefined, { month: "long", day: "numeric" })}`
    : "";

  if (model.settings.announcement) {
    $("announcement").textContent = model.settings.announcement;
    $("announcement").hidden = false;
  }

  buildingRules = parseMappingRules(model.settings.buildings);
  roomIndex = buildRoomIndex();

  if (new URLSearchParams(location.search).has("check")) return renderReport();

  readHash();
  buildFilters();
  renderAll();
  window.addEventListener("hashchange", () => { readHash(); syncFilterButtons(); renderAll(); });
  // Keep "right now" honest without redrawing constantly.
  setInterval(renderNow, 60000);
}

/** Without a server, fetch() of a local file fails — let the file be opened by hand. */
function offerFilePicker(error) {
  const wrap = $("grid");
  wrap.innerHTML = "";
  const box = el("div", "message");
  box.append(el("strong", null, "Could not load the schedule"));

  const local = location.protocol === "file:";
  box.append(
    el("p", null,
      local
        ? "Opening this page straight from a folder stops the browser reading the workbook. Serve the folder (python3 -m http.server) — or choose the file below to preview it."
        : `Tried to read ${WORKBOOK} but got: ${error.message}. If the workbook was just added, give the site a minute and reload.`)
  );

  const input = el("input");
  input.type = "file";
  input.accept = ".xlsx";
  input.style.display = "none";
  input.addEventListener("change", async () => {
    const file = input.files[0];
    if (!file) return;
    wrap.innerHTML = "";
    try {
      await start(await file.arrayBuffer(), new Date(file.lastModified).toUTCString());
    } catch (e) {
      wrap.append(el("p", "message", `That file could not be read: ${e.message}`));
    }
  });

  const button = el("button", null, "Choose office-hours.xlsx…");
  button.type = "button";
  button.addEventListener("click", () => input.click());
  box.append(button, input);
  wrap.append(box);
}

/* ------------------------------------------------------------------ week */

/**
 * The Monday-to-Sunday week containing today.
 *
 * The weekend used to roll forward to the next week, because neither day held
 * anything. Sunday tutoring ended that: on a Saturday the very next thing that
 * happens is tomorrow evening, so the week has to keep today inside it.
 */
function currentWeek() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const monday = new Date(today);
  monday.setDate(today.getDate() - ((today.getDay() + 6) % 7));
  return DAYS.map((name, i) => {
    const date = new Date(monday);
    date.setDate(monday.getDate() + i);
    return {
      name, index: i, date,
      iso: dateKey(date),
      isToday: dateKey(date) === dateKey(new Date()),
      // A day nothing is ever scheduled on. Read from the whole schedule rather
      // than the filtered one, so picking a course cannot close a day.
      closed: !model.openDays.has(i),
    };
  });
}

/* --------------------------------------------------------------- filters */

const matches = (shift) =>
  (!filters.course.size || shift.courses.some((c) => filters.course.has(c))) &&
  (!filters.role.size || filters.role.has(shift.person.role)) &&
  (!filters.building.size || filters.building.has(buildingOf(shift))) &&
  (!filters.person.size || filters.person.has(shift.person.key));

// Learning Assistants first: they are the largest group and the one a student
// is most likely to be looking for by name.
const PERSON_GROUPS = ["la", "gtf", "faculty", "tutor"];
const GROUP_LABELS = {
  la: "Learning Assistants", gtf: "GTFs", faculty: "Faculty", tutor: "Tutors",
};

/** The heading a course sits under in the Get help with menu. */
const subjectLabel = (course) =>
  model.subjectNames.get(subjectOf(course)) || subjectOf(course);

function filterDefinitions() {
  const has = (predicate) => model.shifts.some(predicate);
  const people = [...model.people.values()].filter((p) => p.shifts.length);

  return [
    {
      key: "course", label: "Get help with",
      // Two dozen courses across seven subjects, so they are grouped the way the
      // Person list is. Order comes from course_order on the settings sheet.
      options: model.courses.map((c) => ({ value: c, label: c, group: subjectLabel(c) })),
    },
    {
      key: "role", label: "Role",
      options: model.roles.map((r) => ({ value: r, label: ROLE_LABELS[r] })),
    },
    {
      key: "building", label: "Where",
      // Built from the buildings actually in use, so nothing is offered that
      // has no hours behind it.
      options: orderedBuildings()
        .filter((b) => has((s) => buildingOf(s) === b))
        .map((b) => ({ value: b, label: b })),
    },
    {
      key: "person", label: "Person",
      options: PERSON_GROUPS.flatMap((role) =>
        people
          .filter((p) => p.role === role)
          .sort((a, b) => a.displayName.localeCompare(b.displayName))
          .map((p) => ({ value: p.key, label: p.displayName, group: GROUP_LABELS[role] }))
      ),
    },
  ].filter((f) => f.options.length > 1);
}

function buildFilters() {
  const row = $("filter-row");
  row.innerHTML = "";

  for (const def of filterDefinitions()) {
    const wrap = el("div", "filter");
    const button = el("button");
    button.type = "button";
    button.dataset.key = def.key;
    button.setAttribute("aria-expanded", "false");
    button.append(el("span", "label", def.label), el("span", "caret", "▾"));

    const menu = el("div", "menu");
    menu.hidden = true;
    let currentGroup = null;
    for (const option of def.options) {
      if (option.group && option.group !== currentGroup) {
        currentGroup = option.group;
        menu.append(el("div", "group", currentGroup));
      }
      const label = el("label");
      const input = el("input");
      input.type = "checkbox";
      input.value = option.value;
      input.checked = filters[def.key].has(option.value);
      input.addEventListener("change", () => {
        input.checked ? filters[def.key].add(option.value) : filters[def.key].delete(option.value);
        writeHash();
        syncFilterButtons();
        renderAll();
      });
      label.append(input, el("span", null, option.label));
      menu.append(label);
    }

    button.addEventListener("click", (event) => {
      event.stopPropagation();
      const open = menu.hidden;
      closeMenus();
      menu.hidden = !open;
      button.setAttribute("aria-expanded", String(open));
    });
    menu.addEventListener("click", (event) => event.stopPropagation());

    wrap.append(button, menu);
    row.append(wrap);
  }

  document.addEventListener("click", closeMenus);
  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    if (!$("panel").hidden) closePanel();
    else closeMenus();
  });

  $("clear").addEventListener("click", () => {
    for (const set of Object.values(filters)) set.clear();
    buildFilters();
    writeHash();
    syncFilterButtons();
    renderAll();
  });

  syncFilterButtons();
}

function closeMenus() {
  for (const menu of document.querySelectorAll(".menu")) menu.hidden = true;
  for (const button of document.querySelectorAll(".filter > button")) {
    button.setAttribute("aria-expanded", "false");
  }
}

function syncFilterButtons() {
  const labels = Object.fromEntries(filterDefinitions().map((d) => [d.key, d]));
  for (const button of document.querySelectorAll(".filter > button")) {
    const def = labels[button.dataset.key];
    if (!def) continue;
    const chosen = [...filters[def.key]];
    const text = button.querySelector(".label");
    if (!chosen.length) {
      text.textContent = def.label;
      button.classList.remove("active");
    } else {
      const names = chosen.map((v) => def.options.find((o) => o.value === v)?.label || v);
      text.textContent = names.length <= 2 ? names.join(", ") : `${def.label}: ${names.length}`;
      button.classList.add("active");
    }
  }
  $("clear").hidden = !Object.values(filters).some((set) => set.size);
}

function readHash() {
  const params = new URLSearchParams(location.hash.slice(1));
  for (const key of Object.keys(filters)) {
    filters[key] = new Set((params.get(key) || "").split(",").filter(Boolean));
  }
}

function writeHash() {
  const params = new URLSearchParams();
  for (const [key, set] of Object.entries(filters)) {
    if (set.size) params.set(key, [...set].join(","));
  }
  const hash = params.toString();
  history.replaceState(null, "", hash ? `#${hash}` : location.pathname + location.search);
}

/* ---------------------------------------------------------------- render */

function renderAll() {
  renderNow();
  renderSummary();
  renderGrid();
  renderDayTabs();
  renderDayList();
  renderLegend();
  $("footnote").textContent =
    "This schedule runs every week of the term. Hours can change — check back before you head over.";
}

/**
 * One block per unbroken stretch of availability.
 *
 * Splitting on every change of staffing drew a seam whenever one person handed
 * over to another, which told a student nothing they could act on — help is
 * either there or it isn't. So adjacent occupied slots merge into a single
 * block however much the roster churns inside it.
 */
function blocksFor(day) {
  if (day.closed) return [];
  const rows = [];
  for (let i = 0; i < model.slotCount; i++) {
    const minute = model.dayStart + i * SLOT_MINUTES;
    rows.push(availabilityAt(model, day.iso, day.index, minute, matches));
  }

  const blocks = [];
  let i = 0;
  while (i < model.slotCount) {
    if (!rows[i].length) { i++; continue; }
    let j = i;
    while (j + 1 < model.slotCount && rows[j + 1].length) j++;
    blocks.push({
      startSlot: i, endSlot: j,
      start: model.dayStart + i * SLOT_MINUTES,
      end: model.dayStart + (j + 1) * SLOT_MINUTES,
      entries: peopleAcross(rows, i, j),
    });
    i = j + 1;
  }
  return blocks;
}

/**
 * Who is present across a merged block, and for how long each.
 * Back-to-back shifts in the same place read as one stretch, so a 4–5 and a
 * 5–6 by the same person becomes "here 4 PM–6 PM" rather than two entries.
 */
function peopleAcross(rows, from, to) {
  const sameWhere = (a, b) => a.mode === b.mode && a.location === b.location;
  const byPerson = new Map();

  for (let i = from; i <= to; i++) {
    for (const { person, shift } of rows[i]) {
      let record = byPerson.get(person.key);
      if (!record) byPerson.set(person.key, (record = { person, runs: [] }));
      const last = record.runs[record.runs.length - 1];
      if (last && last.endSlot === i - 1 && sameWhere(last.shift, shift)) last.endSlot = i;
      else record.runs.push({ startSlot: i, endSlot: i, shift });
    }
  }

  const entries = [];
  for (const { person, runs } of byPerson.values()) {
    for (const run of runs) {
      entries.push({
        person,
        shift: run.shift,
        start: model.dayStart + run.startSlot * SLOT_MINUTES,
        end: model.dayStart + (run.endSlot + 1) * SLOT_MINUTES,
      });
    }
  }
  return entries;
}

/**
 * An unbroken block is still clickable hour by hour. The block is drawn as one
 * shape so a student cannot read the shift rota off the grid, but each hour is
 * its own target, so clicking asks "who is here at 2 PM" rather than "who is
 * here sometime this afternoon".
 */
function sectionsFor(block) {
  const sections = [];
  for (let start = block.start; start < block.end; ) {
    const end = Math.min(block.end, Math.floor(start / 60) * 60 + 60);
    sections.push({ start, end, entries: entriesBetween(block, start, end) });
    start = end;
  }
  return sections;
}

const entriesBetween = (block, start, end) =>
  block.entries.filter((entry) => entry.start < end && start < entry.end);

/**
 * Colour bands within a block, at the model's own half-hour resolution.
 *
 * Painting a whole stretch with the union of its roles was a lie: an afternoon
 * that ends with a lone LA was still showing faculty green. Bands are merged
 * only where the composition is actually unchanged, so a seam now means the
 * mix really did change — not that somebody's shift happened to tick over.
 */
function fillsFor(block) {
  const fills = [];
  for (let start = block.start; start < block.end; start += SLOT_MINUTES) {
    const entries = entriesBetween(block, start, start + SLOT_MINUTES);
    if (!entries.length) continue;
    const rooms = roomsIn(entries);
    const key = rooms.join("|");
    const last = fills[fills.length - 1];
    if (last && last.key === key && last.end === start) last.end = start + SLOT_MINUTES;
    else fills.push({ start, end: start + SLOT_MINUTES, key, rooms });
  }
  return fills;
}

const isSelected = (day, section) =>
  selection &&
  selection.dayIndex === day.index &&
  selection.start === section.start &&
  selection.end === section.end;

/** Where a shift meets. An online shift's "room" is the word Online, never its URL. */
const roomOf = (shift) => (shift.mode === "online" ? "Online" : shift.location || "Room TBC");

/**
 * Which building a shift meets in — the Where filter's unit.
 *
 * `buildings` in the settings sheet groups rooms that share a building
 * ("AV -> Academic Village"); anything unmatched is its own building, so a room
 * somewhere new still gets a Where option without anyone writing a rule. Every
 * option therefore names a place that genuinely has hours in it.
 */
function buildingOf(shift) {
  if (shift.mode === "online") return "Online";
  const room = roomOf(shift);
  for (const { source, targets } of buildingRules) {
    if (new RegExp(`^${source.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "i").test(room)) {
      return targets[0];
    }
  }
  return room;
}

/** Buildings in play, in the order their rooms are coloured. */
function orderedBuildings() {
  const seen = [];
  for (const room of orderedRooms()) {
    const shift = model.shifts.find((s) => roomOf(s) === room);
    const building = shift ? buildingOf(shift) : room;
    if (!seen.includes(building)) seen.push(building);
  }
  return seen;
}

/**
 * Room -> palette slot.
 *
 * Every room named in `room_order` reserves its slot whether or not it is
 * currently in use, so a room going quiet for a term does not shuffle the
 * colours of the ones that remain. Rooms not named there take the slots after.
 */
function buildRoomIndex() {
  const pinned = String(model.settings.room_order || "")
    .split(/[;,]/)
    .map((r) => r.trim())
    .filter(Boolean);

  const index = new Map();
  pinned.forEach((room, i) => index.set(room.toLowerCase(), i));

  let next = pinned.length;
  for (const room of roomsInUse()) {
    if (!index.has(room.toLowerCase())) index.set(room.toLowerCase(), next++);
  }
  return index;
}

/** How often each room is used, most-used first, then alphabetical. */
function roomsInUse() {
  const counts = new Map();
  const bump = (room) => counts.set(room, (counts.get(room) || 0) + 1);
  for (const shift of model.shifts) bump(roomOf(shift));
  for (const exception of model.exceptions) {
    if (exception.type === "added") {
      bump(exception.mode === "online" ? "Online" : exception.location);
    }
  }
  return [...counts.keys()].sort((a, b) => counts.get(b) - counts.get(a) || a.localeCompare(b));
}

/** The rooms with hours in them, in palette order — what the legend lists. */
const orderedRooms = () =>
  roomsInUse().sort((a, b) => paletteSlot(a) - paletteSlot(b));

const paletteSlot = (room) => roomIndex.get(String(room).toLowerCase()) ?? 0;

/** The distinct rooms in play across a set of entries, in palette order. */
function roomsIn(entries) {
  const rooms = [...new Set(entries.map((e) => roomOf(e.shift)))];
  return rooms.sort((a, b) => paletteSlot(a) - paletteSlot(b));
}

const roomsLabel = (rooms) =>
  rooms.length === 1 ? rooms[0] : `${rooms.slice(0, -1).join(", ")} and ${rooms[rooms.length - 1]}`;

function renderGrid() {
  const wrap = $("grid");
  wrap.innerHTML = "";
  const grid = el("div", "grid");
  const height = model.slotCount * SLOT_H + GRID_PAD * 2;

  // A closed day still gets a column, just a narrow one — the week reads as a
  // week, and Saturday says "closed" instead of leaving a hole to interpret.
  grid.style.gridTemplateColumns =
    `var(--time-col) ${week.map((d) => (d.closed ? "var(--closed-col)" : "1fr")).join(" ")}`;

  grid.append(el("div", "head corner"));
  for (const day of week) {
    const head = el("div", `head${day.isToday ? " today" : ""}${day.closed ? " closed" : ""}`);
    head.append(el("div", "dow", day.closed ? day.name.slice(0, 3) : day.name));
    grid.append(head);
  }

  const times = el("div", "times");
  times.style.height = `${height}px`;
  // Label sits *on* the hour line a block starts at, the way a calendar reads —
  // centring it in the band made blocks look half an hour out of place.
  for (let i = 0; i <= model.slotCount; i++) {
    const minute = model.dayStart + i * SLOT_MINUTES;
    if (minute % 60) continue;
    const label = el("div", "label", formatTime(minute));
    label.style.top = `${slotTop(i)}px`;
    times.append(label);
    const tick = el("div", `tick${i === model.slotCount ? " last" : ""}`);
    tick.style.top = `${slotTop(i)}px`;
    times.append(tick);
  }
  grid.append(times);

  for (const day of week) {
    const col = el("div", `col${day.isToday ? " today" : ""}${day.closed ? " closed" : ""}`);
    col.style.height = `${height}px`;

    if (day.closed) {
      const shut = el("div", "shut");
      shut.append(el("span", null, "Closed"));
      shut.setAttribute("aria-label", `${day.name}: closed, no hours.`);
      col.append(shut);
      grid.append(col);
      continue;
    }

    for (let i = 0; i <= model.slotCount; i++) {
      const rule = el("div", `rule${(model.dayStart + i * SLOT_MINUTES) % 60 === 0 ? " hour" : ""}`);
      rule.style.top = `${slotTop(i)}px`;
      col.append(rule);
    }

    for (const block of blocksFor(day)) {
      const slots = block.endSlot - block.startSlot + 1;
      const shape = el("div", "block");
      // Flush to the hour lines: any inset here reads as a misalignment.
      shape.style.top = `${slotTop(block.startSlot)}px`;
      shape.style.height = `${slots * SLOT_H}px`;

      const offset = (minute) => ((minute - block.start) / SLOT_MINUTES) * SLOT_H;
      for (const fill of fillsFor(block)) {
        const band = el("div", "fill");
        band.style.top = `${offset(fill.start)}px`;
        band.style.height = `${offset(fill.end) - offset(fill.start)}px`;
        // One stripe per room, so a half hour split across two rooms shows both.
        for (const room of fill.rooms) band.append(el("div", `room r${paletteSlot(room) % 6}`));
        shape.append(band);
      }

      for (const section of sectionsFor(block)) {
        const button = el("button", "section");
        button.type = "button";
        button.style.top = `${offset(section.start)}px`;
        button.style.height = `${offset(section.end) - offset(section.start)}px`;
        if (isSelected(day, section)) button.classList.add("selected");

        const count = section.entries.length;
        const description =
          `${day.name} ${formatRange(section.start, section.end)}: ` +
          `${roomsLabel(roomsIn(section.entries))} · ` +
          `${count} ${count === 1 ? "person" : "people"}.`;
        // The grid carries no text, so the label has to say what the colour says.
        button.setAttribute("aria-label", `${description} Open details.`);
        button.title = description;
        button.addEventListener("click", () => selectSection(day, section));
        shape.append(button);
      }

      col.append(shape);
    }

    if (day.isToday) {
      const minutes = new Date().getHours() * 60 + new Date().getMinutes();
      if (minutes >= model.dayStart && minutes <= model.dayEnd) {
        const line = el("div", "nowline");
        line.style.top = `${slotTop((minutes - model.dayStart) / SLOT_MINUTES)}px`;
        col.append(line);
      }
    }

    grid.append(col);
  }

  wrap.append(grid);
}

function renderDayTabs() {
  const tabs = $("daytabs");
  tabs.hidden = false;
  tabs.innerHTML = "";
  for (const day of week) {
    const button = el("button");
    button.type = "button";
    button.setAttribute("role", "tab");
    button.setAttribute("aria-selected", String(day.index === activeDay));
    if (day.closed) button.classList.add("closed");
    button.append(el("span", "dow", day.name.slice(0, 3)));
    button.addEventListener("click", () => {
      activeDay = day.index;
      renderDayTabs();
      renderDayList();
    });
    tabs.append(button);
  }
}

function renderDayList() {
  const host = $("grid");
  host.querySelector(".daylist")?.remove();

  const day = week[activeDay];
  const list = el("div", "daylist");
  const blocks = day.closed ? [] : blocksFor(day);

  if (day.closed) {
    list.append(el("p", "empty", `Closed on ${day.name}. Nothing is scheduled.`));
  } else if (!blocks.length) {
    list.append(el("p", "empty", "Nothing on this day with the filters you have chosen."));
  }
  for (const block of blocks) {
    const head = el("div", "daygroup");
    head.append(el("span", "range", formatRange(block.start, block.end)));
    list.append(head);

    for (const section of sectionsFor(block)) {
      const rooms = roomsIn(section.entries);
      const row = el("button", "row");
      row.type = "button";
      if (isSelected(day, section)) row.classList.add("selected");
      const edge = el("span", "edge");
      for (const room of rooms) edge.append(el("i", `room r${paletteSlot(room) % 6}`));
      row.append(
        edge,
        el("span", "when", formatRange(section.start, section.end)),
        el("span", "what", roomsLabel(rooms))
      );
      row.setAttribute(
        "aria-label",
        `${day.name} ${formatRange(section.start, section.end)}: ${roomsLabel(rooms)}.`
      );
      row.addEventListener("click", () => selectSection(day, section));
      list.append(row);
    }
  }
  host.append(list);
}

function renderLegend() {
  const legend = $("legend");
  legend.innerHTML = "";
  const rooms = orderedRooms().filter((room) =>
    model.shifts.some((s) => roomOf(s) === room && matches(s))
  );
  for (const room of rooms) {
    const key = el("span", "key");
    key.append(el("i", `swatch room r${paletteSlot(room) % 6}`), document.createTextNode(room));
    legend.append(key);
  }
  legend.append(el("span", "key", "Click any block to see who is there."));
  legend.hidden = !rooms.length;
}

/**
 * Nothing is reported when there is something to show — the grid speaks for
 * itself. This exists only so an empty result is never a silent one.
 */
function renderSummary() {
  const anything = week.some((day) => blocksFor(day).length);
  $("summary").textContent = anything ? "" : "Nothing matches those filters. Try clearing one.";
  $("summary").hidden = anything;
}

function renderNow() {
  const box = $("now");
  if (!model) return;
  const today = week.find((d) => d.isToday);
  const now = new Date();
  const minutes = now.getHours() * 60 + now.getMinutes();
  box.innerHTML = "";
  box.hidden = false;

  const heading = el("h2", null, "Right now");
  box.append(heading);

  // Read from the same merged blocks the grid draws, so "until" means until
  // they actually leave, not until their current row happens to end.
  const open =
    today && minutes >= model.dayStart && minutes < model.dayEnd
      ? blocksFor(today).find((b) => minutes >= b.start && minutes < b.end)
      : null;
  const entries = (open ? open.entries : []).filter((e) => minutes >= e.start && minutes < e.end);

  if (entries.length) {
    const list = el("ul");
    for (const { person, shift, end } of entries) {
      const item = el("li");
      item.append(el("b", null, person.displayName));
      item.append(el("span", null, ` · ${shift.mode === "online" ? "Online" : shift.location}`));
      item.append(el("span", null, ` · until ${formatTime(end)}`));
      list.append(item);
    }
    box.append(list);
  } else {
    const next = nextOpening(minutes);
    box.append(el("p", "none", next
      ? `Nothing open right now. Next: ${next.dayName} at ${formatTime(next.start)}.`
      : "Nothing open right now."));
  }
}

/**
 * The next block that opens, wrapping into the following week if the rest of
 * this one is spent — which is what a Sunday night reader is looking at, now
 * that Sunday is the last column rather than the first.
 */
function nextOpening(minutesNow) {
  const today = week.find((d) => d.isToday);
  const from = today ? today.index : 0;
  for (let step = 0; step < week.length; step++) {
    const day = week[(from + step) % week.length];
    for (const block of blocksFor(day)) {
      if (step === 0 && today && block.start <= minutesNow) continue;
      return { dayName: step === 0 ? "today" : day.name, start: block.start };
    }
  }
  return null;
}

/* ----------------------------------------------------------------- panel */

let lastFocused = null;

function selectSection(day, section) {
  selection = { dayIndex: day.index, start: section.start, end: section.end };
  renderGrid();
  renderDayList();
  openPanel(day, section);
}

function openPanel(day, section) {
  lastFocused = document.activeElement;
  $("panel-title").textContent = day.name;
  $("panel-sub").textContent = formatRange(section.start, section.end);

  const body = $("panel-body");
  body.innerHTML = "";

  const order = ["faculty", "gtf", "la", "tutor"];
  const entries = [...section.entries].sort(
    (a, b) =>
      order.indexOf(a.person.role) - order.indexOf(b.person.role) ||
      a.person.displayName.localeCompare(b.person.displayName)
  );

  for (const { person, shift, start, end } of entries) {
    const card = el("div", "person");
    const top = el("div", "person-top");
    top.append(el("span", "person-name", person.displayName));
    top.append(el("span", `badge ${person.role}`, ROLE_LABELS[person.role]));
    if (shift.oneOff) top.append(el("span", "badge oneoff", "One-off"));
    card.append(top);

    const meta = el("div", "person-meta");
    if (shift.mode === "online") {
      const line = el("div");
      line.append(document.createTextNode("Online · "));
      if (/^https?:\/\//i.test(shift.location)) {
        const link = el("a", null, "join the meeting");
        link.href = shift.location;
        link.rel = "noopener noreferrer";
        link.target = "_blank";
        line.append(link);
      } else {
        line.append(document.createTextNode(shift.location || "link to come"));
      }
      meta.append(line);
    } else {
      meta.append(el("div", null, `In person · ${shift.location}`));
    }
    meta.append(el("div", null, `Here ${formatRange(start, end)}`));
    meta.append(el("div", null, `Helps with ${shift.courses.join(", ")}`));
    if (person.email) {
      const line = el("div");
      const link = el("a", null, person.email);
      link.href = `mailto:${person.email}`;
      line.append(link);
      meta.append(line);
    }
    if (shift.notes) meta.append(el("div", null, shift.notes));
    if (person.notes) meta.append(el("div", null, person.notes));
    card.append(meta);
    body.append(card);
  }

  // On a phone the panel is a sheet over the page, so it traps focus and dims
  // what is behind it. On a desktop it is a side rail: the grid stays lit and
  // clickable, so a reader can step along the hours and watch the panel follow.
  const sheet = window.matchMedia("(max-width: 760px)").matches;
  $("scrim").hidden = !sheet;
  $("panel").hidden = false;
  $("panel").toggleAttribute("aria-modal", sheet);
  document.body.classList.add("panel-open");
  if (sheet) $("panel-close").focus();
}

function closePanel() {
  $("panel").hidden = true;
  $("scrim").hidden = true;
  document.body.classList.remove("panel-open");
  selection = null;
  renderGrid();
  renderDayList();
  if (lastFocused && lastFocused.isConnected) lastFocused.focus();
}

$("panel-close").addEventListener("click", closePanel);
$("scrim").addEventListener("click", closePanel);

/* ---------------------------------------------------------------- report */

function renderReport() {
  document.querySelector(".filters").hidden = true;
  $("now").hidden = true;
  $("legend").hidden = true;
  $("daytabs").hidden = true;
  $("summary").textContent = "";

  const wrap = $("grid");
  wrap.innerHTML = "";
  const report = el("div", "report");
  report.append(el("h2", null, "Workbook check"));

  const errors = model.problems.filter((p) => p.level === "error");
  const warnings = model.problems.filter((p) => p.level === "warning");

  report.append(el("p", null,
    `${model.people.size} people · ${model.shifts.length} shifts · ${model.exceptions.length} exceptions loaded.`));

  for (const program of model.programs) {
    const rows = model.shifts.filter((s) => s.program === program);
    report.append(el("p", null,
      `${rows.length} ${rows.length === 1 ? "shift" : "shifts"} on the ${program} programme, ` +
      `in ${[...new Set(rows.map(roomOf))].join(", ")}.`));
  }

  for (const { source, targets } of model.courseRules) {
    report.append(el("p", null,
      `Course rule in effect: anyone who covers ${source} is also listed under ${targets.join(", ")}.`));
  }

  if (!model.problems.length) {
    report.append(el("p", "ok", "No problems found. Everything in the workbook is being displayed."));
  } else {
    report.append(el("p", null,
      `${errors.length} ${errors.length === 1 ? "error" : "errors"} (rows hidden from students) · ` +
      `${warnings.length} ${warnings.length === 1 ? "warning" : "warnings"} (shown, but worth a look).`));

    const table = el("table");
    const head = el("tr");
    for (const heading of ["", "Where", "What to fix"]) head.append(el("th", null, heading));
    table.append(head);

    for (const p of [...errors, ...warnings]) {
      const tr = el("tr");
      tr.append(el("td", `lvl ${p.level}`, p.level === "error" ? "Error" : "Warning"));
      tr.append(el("td", "where", p.row ? `${p.sheet} row ${p.row}` : p.sheet));
      tr.append(el("td", null, p.message));
      table.append(tr);
    }
    report.append(table);
  }

  const back = el("p");
  const link = el("a", null, "← Back to the schedule");
  link.href = location.pathname;
  back.append(link);
  report.append(back);
  wrap.append(report);
}
