/*
 * Office Hours — reads data/office-hours.xlsx and draws the week.
 *
 * The default view answers one question with no interaction: when can I get help?
 * So the grid shows counts only. Names, rooms and courses live one click deep,
 * and the filters are all opt-in.
 */

import {
  loadModel, availabilityAt, formatTime, formatRange,
  DAYS, ROLE_LABELS, SLOT_MINUTES, dateKey,
} from "./data.js";

const WORKBOOK = "data/office-hours.xlsx";
const SLOT_H = 26;

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
const filters = { course: new Set(), role: new Set(), mode: new Set(), person: new Set() };

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
  activeDay = Math.min(Math.max(new Date().getDay() - 1, 0), 4);

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

function currentWeek() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  // Saturday and Sunday look ahead to the coming week rather than back at a spent one.
  const offset = today.getDay() === 0 ? 1 : today.getDay() === 6 ? 2 : 1 - today.getDay();
  const monday = new Date(today);
  monday.setDate(today.getDate() + offset);
  return DAYS.map((name, i) => {
    const date = new Date(monday);
    date.setDate(monday.getDate() + i);
    return { name, index: i, date, iso: dateKey(date), isToday: dateKey(date) === dateKey(new Date()) };
  });
}

/* --------------------------------------------------------------- filters */

const matches = (shift) =>
  (!filters.course.size || shift.courses.some((c) => filters.course.has(c))) &&
  (!filters.role.size || filters.role.has(shift.person.role)) &&
  (!filters.mode.size || filters.mode.has(shift.mode)) &&
  (!filters.person.size || filters.person.has(shift.person.key));

function filterDefinitions() {
  const countBy = (predicate) => model.shifts.filter(predicate).length;
  const people = [...model.people.values()]
    .filter((p) => p.shifts.length)
    .sort((a, b) => a.displayName.localeCompare(b.displayName));

  return [
    {
      key: "course", label: "Course",
      options: model.courses.map((c) => ({
        value: c, label: c, count: countBy((s) => s.courses.includes(c)),
      })),
    },
    {
      key: "role", label: "Who",
      options: model.roles.map((r) => ({
        value: r, label: ROLE_LABELS[r], count: countBy((s) => s.person.role === r),
      })),
    },
    {
      key: "mode", label: "Format",
      options: [
        { value: "in-person", label: "In person", count: countBy((s) => s.mode === "in-person") },
        { value: "online", label: "Online", count: countBy((s) => s.mode === "online") },
      ].filter((o) => o.count),
    },
    {
      key: "person", label: "Person",
      options: people.map((p) => ({ value: p.key, label: p.displayName, count: p.shifts.length })),
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
    for (const option of def.options) {
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
      label.append(input, el("span", null, option.label), el("span", "count", String(option.count)));
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
  $("legend").hidden = false;
  $("footnote").textContent =
    "Availability shown for this week. Schedules can change — check back before you head over.";
}

/** Slot-by-slot availability for one day, collapsed into blocks of identical people. */
function blocksFor(day) {
  const rows = [];
  for (let i = 0; i < model.slotCount; i++) {
    const minute = model.dayStart + i * SLOT_MINUTES;
    rows.push(availabilityAt(model, day.iso, day.index, minute, matches));
  }

  const blocks = [];
  let i = 0;
  while (i < model.slotCount) {
    if (!rows[i].length) { i++; continue; }
    const signature = (entries) => entries.map((e) => e.person.key).sort().join("|");
    const key = signature(rows[i]);
    let j = i;
    while (j + 1 < model.slotCount && signature(rows[j + 1]) === key) j++;
    blocks.push({
      startSlot: i, endSlot: j,
      start: model.dayStart + i * SLOT_MINUTES,
      end: model.dayStart + (j + 1) * SLOT_MINUTES,
      entries: rows[i],
    });
    i = j + 1;
  }
  return blocks;
}

const densityClass = (n) => (n >= 3 ? "s3" : n === 2 ? "s2" : "s1");

function renderGrid() {
  const wrap = $("grid");
  wrap.innerHTML = "";
  const grid = el("div", "grid");
  const height = model.slotCount * SLOT_H;

  grid.append(el("div", "head corner"));
  for (const day of week) {
    const head = el("div", `head${day.isToday ? " today" : ""}`);
    head.append(
      el("div", "dow", day.name),
      el("div", "date", day.date.toLocaleDateString(undefined, { month: "short", day: "numeric" }))
    );
    grid.append(head);
  }

  const times = el("div", "times");
  times.style.height = `${height}px`;
  for (let i = 0; i < model.slotCount; i++) {
    const minute = model.dayStart + i * SLOT_MINUTES;
    if (minute % 60 === 0) {
      const label = el("div", "label", formatTime(minute));
      // Centered in the hour band rather than on its edge, so nothing is clipped.
      label.style.top = `${(i + 1) * SLOT_H}px`;
      times.append(label);
    }
  }
  grid.append(times);

  for (const day of week) {
    const col = el("div", `col${day.isToday ? " today" : ""}`);
    col.style.height = `${height}px`;

    for (let i = 1; i < model.slotCount; i++) {
      const rule = el("div", `rule${(model.dayStart + i * SLOT_MINUTES) % 60 === 0 ? " hour" : ""}`);
      rule.style.top = `${i * SLOT_H}px`;
      col.append(rule);
    }

    for (const block of blocksFor(day)) {
      const slots = block.endSlot - block.startSlot + 1;
      const button = el("button", `block ${densityClass(block.entries.length)}${slots < 2 ? " short" : ""}`);
      button.type = "button";
      button.style.top = `${block.startSlot * SLOT_H + 2}px`;
      button.style.height = `${slots * SLOT_H - 4}px`;

      const count = block.entries.length;
      button.append(el("span", "n", String(count)));
      if (slots >= 2 || count > 1) {
        button.append(el("span", "who", "available"));
      }
      button.setAttribute(
        "aria-label",
        `${day.name} ${formatRange(block.start, block.end)}: ${count} ${count === 1 ? "person" : "people"} available. Open details.`
      );
      button.addEventListener("click", () => openPanel(day, block));
      col.append(button);
    }

    if (day.isToday) {
      const minutes = new Date().getHours() * 60 + new Date().getMinutes();
      if (minutes >= model.dayStart && minutes <= model.dayEnd) {
        const line = el("div", "nowline");
        line.style.top = `${((minutes - model.dayStart) / SLOT_MINUTES) * SLOT_H}px`;
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
    button.append(
      el("span", "dow", day.name.slice(0, 3)),
      el("span", "date", day.date.toLocaleDateString(undefined, { day: "numeric" }))
    );
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
  const blocks = blocksFor(day);

  if (!blocks.length) {
    list.append(el("p", "empty", "No office hours on this day with the filters you have chosen."));
  }
  for (const block of blocks) {
    const row = el("button", "row");
    row.type = "button";
    const count = block.entries.length;
    row.append(
      el("span", "when", formatRange(block.start, block.end)),
      el("span", "what", `${count} ${count === 1 ? "person" : "people"} available`)
    );
    row.addEventListener("click", () => openPanel(day, block));
    list.append(row);
  }
  host.append(list);
}

function renderSummary() {
  const shown = new Set();
  let hours = 0;
  for (const day of week) {
    for (const block of blocksFor(day)) {
      hours += ((block.end - block.start) / 60) * block.entries.length;
      for (const entry of block.entries) shown.add(entry.person.key);
    }
  }
  const filtered = Object.values(filters).some((set) => set.size);
  $("summary").textContent = shown.size
    ? `${shown.size} ${shown.size === 1 ? "person" : "people"} · ${Math.round(hours)} hours of help this week${filtered ? " matching your filters" : ""}.`
    : "Nothing matches those filters. Try clearing one.";
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

  const entries =
    today && minutes >= model.dayStart && minutes < model.dayEnd
      ? availabilityAt(model, today.iso, today.index, Math.floor(minutes / SLOT_MINUTES) * SLOT_MINUTES, matches)
      : [];

  if (entries.length) {
    const list = el("ul");
    for (const { person, shift } of entries) {
      const item = el("li");
      item.append(el("b", null, person.displayName));
      item.append(el("span", null, ` · ${shift.mode === "online" ? "Online" : shift.location}`));
      item.append(el("span", null, ` · until ${formatTime(shift.end)}`));
      list.append(item);
    }
    box.append(list);
  } else {
    const next = nextOpening(minutes);
    box.append(el("p", "none", next
      ? `No office hours right now. Next: ${next.dayName} at ${formatTime(next.start)}.`
      : "No office hours right now."));
  }
}

function nextOpening(minutesNow) {
  const today = week.find((d) => d.isToday);
  for (const day of week) {
    if (today && day.index < today.index) continue;
    for (const block of blocksFor(day)) {
      const isLater = !today || day.index > today.index || block.start > minutesNow;
      if (isLater) return { dayName: day.isToday ? "today" : day.name, start: block.start };
    }
  }
  return null;
}

/* ----------------------------------------------------------------- panel */

let lastFocused = null;

function openPanel(day, block) {
  lastFocused = document.activeElement;
  $("panel-title").textContent = `${day.name}, ${day.date.toLocaleDateString(undefined, { month: "long", day: "numeric" })}`;
  $("panel-sub").textContent = formatRange(block.start, block.end);

  const body = $("panel-body");
  body.innerHTML = "";

  const entries = [...block.entries].sort(
    (a, b) =>
      ["faculty", "gtf", "la"].indexOf(a.person.role) - ["faculty", "gtf", "la"].indexOf(b.person.role) ||
      a.person.displayName.localeCompare(b.person.displayName)
  );

  for (const { person, shift } of entries) {
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
    meta.append(el("div", null, `Here ${formatRange(shift.start, shift.end)}`));
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

  $("scrim").hidden = false;
  $("panel").hidden = false;
  $("panel-close").focus();
}

function closePanel() {
  $("panel").hidden = true;
  $("scrim").hidden = true;
  if (lastFocused) lastFocused.focus();
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
