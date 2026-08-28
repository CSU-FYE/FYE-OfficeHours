/*
 * Turns the workbook into the model the page renders from.
 *
 * Everything forgiving happens here: the spreadsheet is edited by a dozen people,
 * so times, days, roles and course codes are all normalized rather than trusted.
 * Anything that cannot be normalized is recorded as a problem instead of throwing,
 * so one bad row never blanks the page. Problems are shown on ?check=1.
 */

import { readWorkbook, toRecords } from "./xlsx.js";

// The whole week, Monday first. Saturday earns a column even with nothing in it:
// a gap where a day should be is a question, an explicit "closed" is an answer.
export const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
export const ROLE_LABELS = {
  faculty: "Faculty", gtf: "GTF", la: "Learning Assistant",
};
export const SLOT_MINUTES = 30;

const DEFAULTS = {
  day_start: "9:00 AM",
  day_end: "9:00 PM",
  default_courses: "ENGR 111; ENGR 114",
  course_implies: "",
  course_order: "",
  subjects: "",
  // Only reached when a location cell is left blank; every row on the shifts
  // sheet is expected to name its own room.
  default_location: "AV C144",
  tutoring_room: "",
  term_name: "",
  announcement: "",
};

/* ------------------------------------------------------------ normalizing */

const str = (v) => (v === null || v === undefined ? "" : String(v).trim());

/** Match names loosely so "Dr. Harvey" and "Dr Harvey" are the same person. */
export const nameKey = (v) =>
  str(v)
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();

const ROLE_ALIASES = {
  faculty: "faculty", professor: "faculty", prof: "faculty", instructor: "faculty",
  gtf: "gtf", gta: "gtf", ta: "gtf", "graduate ta": "gtf", "grad ta": "gtf",
  // Tutor folds into Learning Assistant. Most people who tutor are ENGR LAs
  // anyway, and to a student the two words named the same kind of help, so the
  // distinction only ever made the role filter longer. The old spellings stay
  // here so a workbook that still says "tutor" keeps working.
  la: "la", "learning assistant": "la",
  tutor: "la", "peer tutor": "la", tutoring: "la",
};

export function normalizeRole(v) {
  return ROLE_ALIASES[str(v).toLowerCase().replace(/[^a-z\s]/g, "").replace(/\s+/g, " ")] || null;
}

const TIME_RE = /^(\d{1,2})(?::(\d{2}))?\s*(am|pm|a|p)?$/i;

/**
 * "1:30 PM" | "13:30" | 0.5625 (an Excel time cell) -> minutes since midnight.
 * Returns null if it cannot be read.
 */
export function parseTime(value) {
  if (value === "" || value === null || value === undefined) return null;

  // Excel stores a time as a fraction of a day once the cell is formatted as time.
  if (typeof value === "number") {
    if (value < 0 || value >= 1) return null;
    return Math.round(value * 24 * 60);
  }

  const m = TIME_RE.exec(str(value).replace(/\./g, "").replace(/\s+/g, " ").trim());
  if (!m) return null;
  let hour = Number(m[1]);
  const minute = Number(m[2] || 0);
  const mer = (m[3] || "").toLowerCase()[0];
  if (minute > 59) return null;
  if (mer === "p" && hour !== 12) hour += 12;
  if (mer === "a" && hour === 12) hour = 0;
  if (hour > 24) return null;
  return hour * 60 + minute;
}

export function formatTime(minutes) {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  const mer = h < 12 || h === 24 ? "AM" : "PM";
  const h12 = h % 12 || 12;
  return m ? `${h12}:${String(m).padStart(2, "0")} ${mer}` : `${h12} ${mer}`;
}

export function formatRange(start, end) {
  return `${formatTime(start)}–${formatTime(end)}`;
}

/** "2026-09-01" | "9/1/2026" | 46266 (an Excel date cell) -> "2026-09-01". */
export function parseDate(value) {
  if (value === "" || value === null || value === undefined) return null;
  if (typeof value === "number") {
    if (value < 20000) return null;
    // Excel's epoch is 1899-12-30 (its 1900 leap-year bug is baked into the offset).
    const ms = Math.round(value) * 86400000 + Date.UTC(1899, 11, 30);
    return new Date(ms).toISOString().slice(0, 10);
  }
  const s = str(value);
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s);
  if (m) {
    const year = m[3].length === 2 ? `20${m[3]}` : m[3];
    return `${year}-${pad(m[1])}-${pad(m[2])}`;
  }
  return null;
}

const pad = (n) => String(n).padStart(2, "0");

export const dateKey = (d) =>
  `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

/** "engr111; ENGR 114" -> ["ENGR 111", "ENGR 114"] */
export function parseCourses(value) {
  return str(value)
    .split(/[;,]/)
    .map((c) => c.trim().replace(/^([A-Za-z]+)\s*(\d.*)$/, (_, a, b) => `${a.toUpperCase()} ${b}`))
    .filter(Boolean);
}

/**
 * "spanish; Bengali" -> ["Spanish", "Bengali"].
 *
 * Title-cased so one spelling wins the filter: whoever types "spanish" in one row
 * and "Spanish" in the next gets one option, not two. These are the languages a
 * person can help in *besides* English, which is why the list is usually empty.
 */
export function parseLanguages(value) {
  return str(value)
    .split(/[;,]/)
    .map((l) => l.trim().replace(/\S+/g, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase()))
    .filter(Boolean);
}

/**
 * "A -> B; C -> D, E" — the little arrow syntax the settings sheet uses for
 * anything that maps one name onto others: courses that come free with another
 * course, rooms that share a building. Kept in the sheet rather than in code so
 * a new rule is a cell edit, and so it is visible to whoever maintains it.
 */
export function parseMappingRules(value) {
  const rules = [];
  for (const part of str(value).split(";")) {
    const [from, to] = part.split(/->|→/);
    if (!from || !to) continue;
    const source = parseCourses(from)[0];
    const targets = parseCourses(to);
    if (source && targets.length) rules.push({ source, targets });
  }
  return rules;
}

/** Add every course implied by the ones already listed. */
export function expandCourses(courses, rules) {
  if (!rules.length) return courses;
  const out = [...courses];
  // Bounded rather than recursive: a chained rule still resolves, and a rule
  // that points back at itself cannot spin.
  for (let pass = 0; pass < 5; pass++) {
    let added = false;
    for (const { source, targets } of rules) {
      if (!out.includes(source)) continue;
      for (const target of targets) {
        if (!out.includes(target)) {
          out.push(target);
          added = true;
        }
      }
    }
    if (!added) break;
  }
  return out;
}

function normalizeDay(value) {
  const s = str(value).toLowerCase();
  if (!s) return null;
  // Matched on a prefix so "Mon" and "Monday" both work, but only when the prefix
  // picks out one day — a bare "S" is Saturday or Sunday and must not be guessed at.
  const hits = DAYS.filter((d) => d.toLowerCase().startsWith(s.slice(0, 3)));
  return hits.length === 1 ? hits[0] : null;
}

/**
 * A location is a link, not a room, when it starts like one. That is the whole
 * of what `mode` used to say, and it cannot get out of step with the location
 * the way a separate column could.
 */
const isLink = (location) => /^(https?:)?\/\//i.test(str(location));

/** Room names compared the way a person would: case and spacing don't count. */
const roomKey = (v) => str(v).toLowerCase().replace(/\s+/g, " ");

/**
 * The subject a course belongs to: the letters before its number, or the whole
 * label when it has no number ("Precalculus"). `subjects` in settings turns that
 * into something readable and sets the order the Get help with menu groups by.
 */
export const subjectOf = (course) => (/^[A-Za-z]+/.exec(str(course)) || [""])[0].toUpperCase()
  || str(course).toUpperCase();

/* ----------------------------------------------------------------- build  */

export async function loadModel(buffer) {
  const sheets = await readWorkbook(buffer);
  const find = (name) => {
    const key = Object.keys(sheets).find((k) => k.trim().toLowerCase() === name);
    return key ? toRecords(sheets[key]) : [];
  };
  return buildModel({
    people: find("people"),
    shifts: find("shifts"),
    exceptions: find("exceptions"),
    settings: find("settings"),
    sheetNames: Object.keys(sheets),
  });
}

export function buildModel({ people: peopleRows, shifts: shiftRows, exceptions: exceptionRows, settings: settingRows, sheetNames = [] }) {
  const problems = [];
  const problem = (level, sheet, row, message) => problems.push({ level, sheet, row, message });

  /* settings */
  const settings = { ...DEFAULTS };
  for (const r of settingRows) {
    const key = str(r.key).toLowerCase();
    if (key) settings[key] = str(r.value);
  }
  for (const [k, v] of Object.entries(DEFAULTS)) {
    if (!str(settings[k])) settings[k] = v;
  }

  const dayStart = parseTime(settings.day_start) ?? 9 * 60;
  const dayEnd = parseTime(settings.day_end) ?? 21 * 60;
  const courseRules = parseMappingRules(settings.course_implies);
  const expand = (courses) => expandCourses(courses, courseRules);
  const defaultCourses = expand(parseCourses(settings.default_courses));
  const defaultLocation = str(settings.default_location) || DEFAULTS.default_location;
  // The one room that runs on the people sheet's course lists. Everywhere else
  // is ENGR office hours and advertises `default_courses`, so where an hour is
  // held is the only thing that decides what it offers — no second column to
  // keep in step, and nothing a blank cell can quietly imply.
  const tutoringRoom = roomKey(settings.tutoring_room);

  if (!peopleRows.length) {
    problem("error", "people", null,
      sheetNames.length
        ? `No rows found. The workbook has these sheets: ${sheetNames.join(", ")}.`
        : "No rows found.");
  }

  /* people */
  const people = new Map();
  for (const r of peopleRows) {
    const name = str(r.name);
    if (!name) continue;
    const key = nameKey(name);
    if (people.has(key)) {
      problem("error", "people", r.__row, `"${name}" is listed twice. Names must be unique.`);
      continue;
    }
    const role = normalizeRole(r.role);
    if (!role) {
      problem("error", "people", r.__row,
        `"${name}" has role "${str(r.role) || "(blank)"}". Use faculty, gtf, or la. This person and their shifts are hidden.`);
      continue;
    }
    // Someone can be an ENGR Learning Assistant on Tuesday and tutor CHEM on
    // Wednesday. This column is the second of those: what they can help with in
    // the tutoring room. Their ENGR office hours advertise `default_courses`.
    const courses = expand(parseCourses(r.courses));
    people.set(key, {
      key,
      // `name` is both the key the shifts sheet joins on and what students read,
      // so the sheet asks for it written the way they should read it. There was a
      // second `display_name` column for the difference; in a term of use nobody
      // ever filled it in.
      name,
      role,
      courses,
      // A property of the person, not of the shift: someone who can explain a
      // derivative in Spanish can do it at office hours and at tutoring alike,
      // even though the rota only records it beside the tutoring roster.
      languages: parseLanguages(r.languages),
      shifts: [],
    });
  }

  /* shifts */
  const shifts = [];
  for (const r of shiftRows) {
    const rawName = str(r.name);
    if (!rawName) continue;

    const person = people.get(nameKey(rawName));
    if (!person) {
      problem("error", "shifts", r.__row,
        `"${rawName}" is not on the people sheet (or their row there is invalid). This shift is hidden.`);
      continue;
    }
    const day = normalizeDay(r.day);
    if (!day) {
      problem("error", "shifts", r.__row,
        `${rawName}: "${str(r.day) || "(blank)"}" is not a weekday. This shift is hidden.`);
      continue;
    }
    const start = parseTime(r.start);
    const end = parseTime(r.end);
    if (start === null || end === null) {
      problem("error", "shifts", r.__row,
        `${rawName} ${day}: could not read the time "${str(r.start)}" to "${str(r.end)}". Try 9:00 AM. This shift is hidden.`);
      continue;
    }
    if (end <= start) {
      problem("error", "shifts", r.__row,
        `${rawName} ${day}: ends (${formatTime(end)}) at or before it starts (${formatTime(start)}). This shift is hidden.`);
      continue;
    }
    if (start % SLOT_MINUTES || end % SLOT_MINUTES) {
      problem("warning", "shifts", r.__row,
        `${rawName} ${day} ${formatRange(start, end)}: times should land on the hour or half hour. Shown rounded.`);
    }
    if (start < dayStart || end > dayEnd) {
      problem("warning", "shifts", r.__row,
        `${rawName} ${day} ${formatRange(start, end)} falls partly outside the ${formatTime(dayStart)}–${formatTime(dayEnd)} grid and is clipped.`);
    }

    const location = str(r.location) || defaultLocation;
    if (!str(r.location)) {
      problem("warning", "shifts", r.__row,
        `${rawName} ${day}: no location, so this hour is advertised as ${defaultLocation}. Fill the location column in.`);
    }
    const atTutoring = tutoringRoom && roomKey(location) === tutoringRoom;
    if (atTutoring && !person.courses.length) {
      problem("warning", "shifts", r.__row,
        `${rawName} ${day}: this shift is in ${location}, but their courses cell on the people sheet is blank, so it advertises ${defaultCourses.join(", ")}.`);
    }

    shifts.push({
      id: `s${shifts.length}`,
      person,
      day,
      dayIndex: DAYS.indexOf(day),
      start: round(Math.max(start, dayStart)),
      end: round(Math.min(end, dayEnd)),
      mode: isLink(location) ? "online" : "in-person",
      location,
      courses: atTutoring && person.courses.length ? person.courses : defaultCourses,
      row: r.__row,
    });
  }

  for (const shift of shifts) shift.person.shifts.push(shift);

  /* overlapping duplicates for one person on one day */
  const seen = new Map();
  for (const s of shifts) {
    const key = `${s.person.key}|${s.dayIndex}`;
    for (const other of seen.get(key) || []) {
      if (s.start < other.end && other.start < s.end) {
        problem("warning", "shifts", s.row,
          `${s.person.name} is listed twice on ${s.day} over the same time (rows ${other.row} and ${s.row}). They will be counted once.`);
      }
    }
    seen.set(key, [...(seen.get(key) || []), s]);
  }

  /* exceptions */
  const exceptions = [];
  for (const r of exceptionRows) {
    const rawName = str(r.name);
    if (!rawName) continue;
    const person = people.get(nameKey(rawName));
    if (!person) {
      problem("error", "exceptions", r.__row, `"${rawName}" is not on the people sheet. This row is ignored.`);
      continue;
    }
    const date = parseDate(r.date);
    if (!date) {
      problem("error", "exceptions", r.__row,
        `${rawName}: could not read the date "${str(r.date)}". Try 2026-09-15. This row is ignored.`);
      continue;
    }
    const type = /^add(ed)?$/i.test(str(r.type)) ? "added" : "cancelled";
    if (str(r.type) && !/^(cancelled|canceled|cancel|added|add)$/i.test(str(r.type))) {
      problem("warning", "exceptions", r.__row,
        `${rawName} ${date}: type "${str(r.type)}" is not recognized, treated as cancelled.`);
    }
    const start = parseTime(r.start);
    const end = parseTime(r.end);
    if (type === "added" && (start === null || end === null)) {
      problem("error", "exceptions", r.__row,
        `${rawName} ${date}: added hours need both a start and an end time. This row is ignored.`);
      continue;
    }
    if (start !== null && end !== null && end <= start) {
      problem("error", "exceptions", r.__row,
        `${rawName} ${date}: ends at or before it starts. This row is ignored.`);
      continue;
    }

    const dayIndex = DAYS.indexOf(weekdayOf(date));
    if (type === "cancelled") {
      const matches = person.shifts.filter(
        (s) => s.dayIndex === dayIndex && (start === null || (s.start < end && start < s.end))
      );
      if (!matches.length) {
        problem("warning", "exceptions", r.__row,
          `${rawName} has no office hours on ${date}${start === null ? "" : ` at ${formatRange(start, end)}`}, so this cancellation changes nothing.`);
      }
    }

    exceptions.push({
      person, date, type, dayIndex,
      start: start === null ? null : round(start),
      end: end === null ? null : round(end),
      mode: isLink(r.location) ? "online" : "in-person",
      location: str(r.location) || defaultLocation,
      note: str(r.note),
      row: r.__row,
    });
  }

  const courses = order(
    [...new Set([...defaultCourses, ...shifts.flatMap((s) => s.courses)])],
    parseCourses(settings.course_order)
  );
  const roles = ["faculty", "gtf", "la"].filter((role) =>
    shifts.some((s) => s.person.role === role)
  );
  // Alphabetical, and only languages with hours actually behind them — offering
  // one that matches nothing is worse than not offering it. No `language_order`
  // setting: unlike courses and rooms, no language here outranks another.
  const languages = [...new Set(shifts.flatMap((s) => s.person.languages))].sort();
  // Which days the week is open at all — read from every shift, not the filtered
  // ones, so choosing a course cannot make Friday collapse into a closed column.
  const openDays = new Set(shifts.map((s) => s.dayIndex));
  const subjectNames = new Map(
    parseMappingRules(settings.subjects).map(({ source, targets }) => [
      subjectOf(source), targets[0],
    ])
  );

  return {
    settings, dayStart, dayEnd, people, shifts, exceptions, courses, roles, problems, courseRules,
    openDays, subjectNames, languages, defaultCourses, tutoringRoom,
    slotCount: Math.ceil((dayEnd - dayStart) / SLOT_MINUTES),
  };
}

const round = (m) => Math.round(m / SLOT_MINUTES) * SLOT_MINUTES;

/** Whether a location is the room whose hours run on the people sheet's courses. */
export const atTutoringRoom = (model, location) =>
  Boolean(model.tutoringRoom) && roomKey(location) === model.tutoringRoom;

/**
 * `values` sorted by a preferred list from the settings sheet, with anything the
 * list does not name falling in alphabetically behind it. Same reasoning as
 * `room_order`: a course or a room dropping out must not reshuffle the rest.
 */
function order(values, preferred, key = (v) => v) {
  const rank = new Map(preferred.map((v, i) => [key(v), i]));
  const at = (v) => rank.get(key(v)) ?? preferred.length;
  return [...values].sort((a, b) => at(a) - at(b) || String(a).localeCompare(String(b)));
}

/** Monday is 0 here, where JavaScript makes it 1 and puts Sunday at the front. */
function weekdayOf(isoDate) {
  const [y, m, d] = isoDate.split("-").map(Number);
  return DAYS[(new Date(y, m - 1, d).getDay() + 6) % 7];
}

/**
 * Everyone available on a given date and time, after filters, with exceptions applied.
 * Returns entries, not people, so one person appearing twice is still one row.
 */
export function availabilityAt(model, isoDate, dayIndex, minute, matches) {
  const out = [];
  const seen = new Set();

  for (const shift of model.shifts) {
    if (shift.dayIndex !== dayIndex) continue;
    if (minute < shift.start || minute >= shift.end) continue;
    if (matches && !matches(shift)) continue;
    if (seen.has(shift.person.key)) continue;
    const cancelled = model.exceptions.find(
      (e) =>
        e.type === "cancelled" &&
        e.person.key === shift.person.key &&
        e.date === isoDate &&
        (e.start === null || (minute >= e.start && minute < e.end))
    );
    seen.add(shift.person.key);
    out.push({ shift, person: shift.person, cancelled: cancelled || null });
  }

  for (const e of model.exceptions) {
    if (e.type !== "added" || e.date !== isoDate) continue;
    if (minute < e.start || minute >= e.end) continue;
    const pseudo = {
      id: `x${e.row}`, person: e.person, day: DAYS[e.dayIndex], dayIndex: e.dayIndex,
      start: e.start, end: e.end, mode: e.mode, location: e.location,
      courses: atTutoringRoom(model, e.location) && e.person.courses.length
        ? e.person.courses
        : model.defaultCourses,
      notes: e.note, oneOff: true,
    };
    if (matches && !matches(pseudo)) continue;
    if (seen.has(e.person.key)) continue;
    seen.add(e.person.key);
    out.push({ shift: pseudo, person: e.person, cancelled: null });
  }

  return out.filter((entry) => !entry.cancelled);
}
