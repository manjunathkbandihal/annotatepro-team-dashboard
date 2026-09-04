import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { createRoot } from "react-dom/client";
import {
  LayoutDashboard, Users, FolderKanban, ClipboardCheck,
  AlertTriangle, BarChart3, Settings, Search, Plus, Bell,
  Download, Upload, Menu, X, CheckCircle2, Target,
  Image as ImageIcon, ChevronRight, Trash2, Activity, ShieldCheck,
  LogOut, Mail, LockKeyhole, Pencil, Save, ExternalLink, FileText, Printer,
  Calendar, Sun, Clock, Archive
} from "lucide-react";
import { supabase, isSupabaseConfigured } from "./supabaseClient";
import "./styles.css";

/* =========================================================
   PROJECT TARGETS
========================================================= */
const projectTargets = {
  momah_phase2_april: 1000,
  MBS_Street_Detections_Phase2: 1000,
  MBS_frames_june_phase2: 1000,
  combined_aug_data_1: 800,
  iltwy_73026_53front_1: 350
};

/* =========================================================
   STATUS HELPER
   Pending     = completed 0 and remaining = total
   In Progress = completed > 0 and remaining > 0
   Completed   = remaining 0
========================================================= */
function getProjectStatus(total, completed, remaining) {
  const t = Math.max(0, Number(total) || 0);
  const c = Math.max(0, Number(completed) || 0);
  const r = Math.max(0, Number(remaining) || 0);

  if (!t) return "No Target";
  if (r === 0 || c >= t) return "Completed";
  if (c === 0 || r >= t) return "Pending";
  return "In Progress";
}

function getProjectStats(project) {
  // "total" is scope of work (Total Images). "target" is a separate,
  // unrelated number (Daily Target) and must never be used as a total.
  const total = Math.max(
    0,
    Number(project.totalImages ?? project.total) || 0
  );

  let completed;

  if (project.completed != null) {
    completed = Math.max(0, Number(project.completed) || 0);
  } else {
    completed = Math.max(
      0,
      total - (Number(project.remaining) || 0)
    );
  }

  completed = Math.min(completed, total || completed);

  const remaining = Math.max(0, total - completed);

  const progress = total
    ? Math.min(100, Math.round((completed / total) * 100))
    : 0;

  return {
    total,
    completed,
    remaining,
    progress,
    status: getProjectStatus(total, completed, remaining)
  };
}

/* =========================================================
   DEFAULT DATA
========================================================= */
const seed = {
  team: [
    { id: 1, name: "Manjunath", role: "Team Lead", target: 1000, completed: 862, reviewed: 540, errors: 12, status: "Active" },
    { id: 2, name: "Nandini Keni", role: "Annotator", target: 1000, completed: 918, reviewed: 420, errors: 18, status: "Active" },
    { id: 3, name: "Shantha Mantri", role: "Annotator", target: 1000, completed: 744, reviewed: 390, errors: 9, status: "Active" },
    { id: 4, name: "Shweta Kannagar", role: "Reviewer", target: 400, completed: 372, reviewed: 372, errors: 7, status: "Active" },
    { id: 5, name: "Kishore Devaragudi", role: "Annotator", target: 1000, completed: 1000, reviewed: 610, errors: 6, status: "Active" },
    { id: 6, name: "K Ganesh", role: "Annotator", target: 1000, completed: 581, reviewed: 240, errors: 21, status: "Away" }
  ],

  projects: [
    { id: 1, name: "momah_phase2_april", target: 1000, total: 1000, completed: 0, remaining: 1000, status: "Pending", deadline: "" },
    { id: 2, name: "MBS_Street_Detections_Phase2", target: 1000, total: 1000, completed: 0, remaining: 1000, status: "Pending", deadline: "" },
    { id: 3, name: "MBS_frames_june_phase2", target: 1000, total: 1000, completed: 0, remaining: 1000, status: "Pending", deadline: "" },
    { id: 4, name: "combined_aug_data_1", target: 800, total: 800, completed: 0, remaining: 800, status: "Pending", deadline: "" },
    { id: 5, name: "iltwy_73026_53front_1", target: 350, total: 350, completed: 0, remaining: 350, status: "Pending", deadline: "" }
  ],

  issues: [
    { id: 1, type: "Missed labels", project: "PCI_Annotations", owner: "Priya", severity: "High", status: "Open", date: "2026-08-11" },
    { id: 2, type: "Wrong prediction", project: "hase2_july_data_1", owner: "Kiran", severity: "Medium", status: "Open", date: "2026-08-11" },
    { id: 3, type: "Label inconsistency", project: "PCI_Annotations", owner: "Rahul", severity: "Low", status: "Resolved", date: "2026-08-10" }
  ],

  holidays: [],

  attendanceOverrides: []
};


/* =========================================================
   LOCAL STORAGE
========================================================= */

function loadData() {
  try {
    const saved = JSON.parse(localStorage.getItem("annotatepro-data"));

    if (!saved) return seed;

    return {
      ...seed,

      ...saved,

      team: Array.isArray(saved.team)
        ? saved.team
        : seed.team,

      projects: Array.isArray(saved.projects)
        ? saved.projects.map(p => {
            const s = getProjectStats(p);

            // Keep p.target (Daily Target) exactly as saved — it is a
            // separate number from Total Images and must not be touched here.
            return {
              ...p,
              ...s
            };
          })
        : seed.projects,

      issues: Array.isArray(saved.issues)
        ? saved.issues
        : seed.issues,

      sheetRecords: Array.isArray(saved.sheetRecords)
        ? saved.sheetRecords
        : [],

      accuracyRecords: Array.isArray(saved.accuracyRecords)
        ? saved.accuracyRecords
        : [],

      holidays: Array.isArray(saved.holidays)
        ? saved.holidays
        : [],

      attendanceOverrides: Array.isArray(saved.attendanceOverrides)
        ? saved.attendanceOverrides
        : [],

      accuracyFile: saved.accuracyFile || "",

      accuracyLastSync: saved.accuracyLastSync || ""
    };
  } catch {
    return seed;
  }
}


function saveData(data) {
  localStorage.setItem(
    "annotatepro-data",
    JSON.stringify(data)
  );
}


/* =========================================================
   PREFERENCES (device-local, not synced with the team)
   General Settings live here — how THIS browser likes to view
   the dashboard, not shared facts about the team.
========================================================= */

const PREFS_KEY = "annotatepro-prefs";

const defaultPrefs = {
  landingPage: "dashboard",
  reportRangeMode: "lastImportedWeek" // or "thisWeek"
};

function loadPrefs() {
  try {
    const raw = localStorage.getItem(PREFS_KEY);
    return raw ? { ...defaultPrefs, ...JSON.parse(raw) } : { ...defaultPrefs };
  } catch {
    return { ...defaultPrefs };
  }
}

function savePrefs(prefs) {
  try {
    localStorage.setItem(PREFS_KEY, JSON.stringify(prefs));
  } catch {
    // localStorage unavailable — preferences just won't persist.
  }
}


/* =========================================================
   ACTIVITY LOG (device-local only)
   A simple on-device record of notable actions, for the
   Security panel. Not synced or shared across devices/users.
========================================================= */

const ACTIVITY_KEY = "annotatepro-activity";
const ACTIVITY_LIMIT = 50;

function logActivity(message) {
  try {
    const raw = localStorage.getItem(ACTIVITY_KEY);
    const list = raw ? JSON.parse(raw) : [];
    list.unshift({ id: Date.now(), message, at: new Date().toISOString() });
    localStorage.setItem(ACTIVITY_KEY, JSON.stringify(list.slice(0, ACTIVITY_LIMIT)));
  } catch {
    // localStorage unavailable — activity just won't be recorded.
  }
}

function getActivityLog() {
  try {
    const raw = localStorage.getItem(ACTIVITY_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function clearActivityLog() {
  try {
    localStorage.removeItem(ACTIVITY_KEY);
  } catch {
    // no-op
  }
}


/* =========================================================
   PROJECT NAME
========================================================= */

function getConfiguredProjectName(project) {
  const clean = String(project || "").trim();

  const found = Object.keys(projectTargets).find(
    name => name.toLowerCase() === clean.toLowerCase()
  );

  return found || clean;
}


/* =========================================================
   DATE HELPERS
   IMPORTANT:
   All dates are treated as CALENDAR DATES.
   No timezone conversion is used.
========================================================= */

function pad2(value) {
  return String(value).padStart(2, "0");
}


/*
  Creates a validated YYYY-MM-DD date string.

  Example:
  toISODate(2026, 8, 28)
  => "2026-08-28"
*/
function toISODate(year, month, day) {
  const y = Number(year);
  const m = Number(month);
  const d = Number(day);

  if (
    !Number.isInteger(y) ||
    !Number.isInteger(m) ||
    !Number.isInteger(d)
  ) {
    return "";
  }

  if (
    y < 1900 ||
    m < 1 ||
    m > 12 ||
    d < 1 ||
    d > 31
  ) {
    return "";
  }

  /*
    IMPORTANT:
    Use LOCAL date construction only for validation.
    We never call toISOString() here.
  */
  const date = new Date(y, m - 1, d);

  if (Number.isNaN(date.getTime())) {
    return "";
  }

  if (
    date.getFullYear() !== y ||
    date.getMonth() !== m - 1 ||
    date.getDate() !== d
  ) {
    return "";
  }

  return `${y}-${pad2(m)}-${pad2(d)}`;
}


/*
  Converts a YYYY-MM-DD calendar date into a LOCAL Date object.

  DO NOT use:
      new Date("2026-08-28")

  for calendar calculations.

  This function prevents timezone-related weekday errors.
*/
function dateKeyToLocalDate(dateKey) {
  const match = String(dateKey || "")
    .trim()
    .match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (!match) return null;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);

  const date = new Date(
    year,
    month - 1,
    day
  );

  if (
    Number.isNaN(date.getTime()) ||
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }

  return date;
}


/*
  Returns weekday number safely.

  Sunday = 0
  Monday = 1
  Tuesday = 2
  Wednesday = 3
  Thursday = 4
  Friday = 5
  Saturday = 6

  Example:
  getCalendarDay("2026-07-31") => 5 (Friday)
  getCalendarDay("2026-08-01") => 6 (Saturday)
*/
function getCalendarDay(dateValue) {
  const normalized = normalizeDateValue(dateValue);

  if (!normalized) return null;

  const date = dateKeyToLocalDate(normalized);

  if (!date) return null;

  return date.getDay();
}


/* =========================================================
   DATE NORMALIZATION
========================================================= */

function normalizeDateValue(value) {
  if (value == null || value === "") {
    return "";
  }


  /*
    CASE 1:
    JavaScript Date object.

    IMPORTANT:
    Read the LOCAL calendar components directly.
    Do NOT use toISOString().
  */
  if (
    value instanceof Date &&
    !Number.isNaN(value.getTime())
  ) {
    /*
      IMPORTANT:
      XLSX (with cellDates: true) builds this Date object
      anchored to UTC midnight for the calendar date shown
      in the sheet. Reading it back with LOCAL getters
      (getFullYear/getMonth/getDate) can shift the day
      depending on the browser's timezone offset.
      Always use the UTC getters here.
    */
    return toISODate(
      value.getUTCFullYear(),
      value.getUTCMonth() + 1,
      value.getUTCDate()
    );
  }


  /*
    CASE 2:
    Excel serial date.

    Example:
    46262 -> calendar date

    Use UTC only for the Excel serial calculation,
    then extract UTC components.
  */
  if (
    typeof value === "number" &&
    Number.isFinite(value) &&
    value > 20000 &&
    value < 80000
  ) {
    const excelEpoch = Date.UTC(
      1899,
      11,
      30
    );

    const milliseconds =
      Math.round(value * 86400000);

    const d = new Date(
      excelEpoch + milliseconds
    );

    return toISODate(
      d.getUTCFullYear(),
      d.getUTCMonth() + 1,
      d.getUTCDate()
    );
  }


  const text = String(value).trim();

  if (!text) {
    return "";
  }


  /*
    CASE 3:
    YYYY-MM-DD
    YYYY/MM/DD
    YYYY.MM.DD
  */
  let m = text.match(
    /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/
  );

  if (m) {
    return toISODate(
      m[1],
      m[2],
      m[3]
    );
  }


  /*
    CASE 4:
    DD-MM-YYYY
    DD/MM/YYYY
    DD.MM.YYYY

    This is important for your Google Sheet format:

    28/08/2026
    => 2026-08-28
  */
  m = text.match(
    /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/
  );

  if (m) {
    return toISODate(
      m[3],
      m[2],
      m[1]
    );
  }


  /*
    CASE 5:
    Sometimes imported dates contain a time.

    Example:
    2026-08-28T00:00:00
    2026-08-28T00:00:00.000
  */
  m = text.match(
    /^(\d{4})-(\d{1,2})-(\d{1,2})(?:T|\s)/
  );

  if (m) {
    return toISODate(
      m[1],
      m[2],
      m[3]
    );
  }


  /*
    CASE 6:
    DD/MM/YYYY with time.
  */
  m = text.match(
    /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})(?:T|\s)/
  );

  if (m) {
    return toISODate(
      m[3],
      m[2],
      m[1]
    );
  }


  /*
    IMPORTANT:
    Do NOT fall back to:
        new Date(text)

    because browser timezone/date parsing can change
    the calendar day.

    Unknown date formats are rejected instead.
  */

  return "";
}


/* =========================================================
   DATE RANGE
========================================================= */

function getDateRange(filter) {
  if (!filter) {
    return {
      start: "",
      end: "",
      label: "Select a date"
    };
  }

  const start = normalizeDateValue(
    filter.start || ""
  );

  const end =
    filter.mode === "range"
      ? normalizeDateValue(filter.end || "")
      : start;

  if (!start) {
    return {
      start: "",
      end: "",
      label:
        filter.mode === "range"
          ? "Select a date range"
          : "Select a date"
    };
  }

  if (
    filter.mode === "range" &&
    !end
  ) {
    return {
      start,
      end: "",
      label: `${start} → Select end date`
    };
  }

  if (
    filter.mode === "range" &&
    end < start
  ) {
    return {
      start: end,
      end: start,
      label: `${end} → ${start}`
    };
  }

  return {
    start,
    end,
    label:
      start === end
        ? start
        : `${start} → ${end}`
  };
}


/* =========================================================
   DATE RANGE CHECK
========================================================= */

function isDateInRange(date, range) {
  if (
    !range?.start ||
    !range?.end ||
    !date
  ) {
    return false;
  }

  const normalized =
    normalizeDateValue(date);

  if (!normalized) {
    return false;
  }

  return (
    normalized >= range.start &&
    normalized <= range.end
  );
}


/* =========================================================
   LATEST IMPORTED DATE
========================================================= */

// The Daily Effort Sheet and Accuracy Report are two independent files on
// their own schedules — their dates must never be blended into one "latest
// date", or Attendance/productivity can end up anchored to a date that only
// exists in the OTHER file and has no real work data at all.
function getLatestImportedDate(data) {
  // Only rows with real work count — a pre-filled "Saturday"/"Sunday"
  // placeholder for a future date that hasn't happened yet must not
  // get picked as the latest date.
  const dates = (Array.isArray(data?.sheetRecords) ? data.sheetRecords : [])
    .filter(isWorkRow)
    .map(x => normalizeDateValue(x.date))
    .filter(Boolean)
    .sort();

  return dates.length
    ? dates[dates.length - 1]
    : "";
}

function getLatestAccuracyDate(data) {
  const dates = (Array.isArray(data?.accuracyRecords) ? data.accuracyRecords : [])
    .map(x => normalizeDateValue(x.date))
    .filter(Boolean)
    .sort();

  return dates.length ? dates[dates.length - 1] : "";
}


/* =========================================================
   NOTIFICATIONS
   Two kinds today:
   - Project deadlines that are approaching or already past
   - High severity issues that have been open too long
========================================================= */

const DEADLINE_WARNING_DAYS = 7;
const STALE_HIGH_SEVERITY_DAYS = 2;

function getTodayISO() {
  const d = new Date();
  return toISODate(d.getFullYear(), d.getMonth() + 1, d.getDate());
}

// Whole-day difference between two YYYY-MM-DD calendar dates (toISO - fromISO).
function daysBetweenDates(fromISO, toISO) {
  const a = dateKeyToLocalDate(fromISO);
  const b = dateKeyToLocalDate(toISO);
  if (!a || !b) return null;
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

function getNotifications(data) {
  const today = getTodayISO();
  const items = [];

  (Array.isArray(data?.projects) ? data.projects : []).forEach(p => {
    if (p.status === "Completed") return;

    const deadline = normalizeDateValue(p.deadline);
    if (!deadline) return;

    const daysLeft = daysBetweenDates(today, deadline);
    if (daysLeft == null) return;

    if (daysLeft < 0) {
      const overdue = Math.abs(daysLeft);
      items.push({
        id: `deadline-${p.id}`,
        kind: "overdue",
        title: `${p.name} is overdue`,
        detail: `Deadline was ${overdue} day${overdue === 1 ? "" : "s"} ago (${deadline})`,
        date: deadline
      });
    } else if (daysLeft <= DEADLINE_WARNING_DAYS) {
      items.push({
        id: `deadline-${p.id}`,
        kind: "deadline",
        title:
          daysLeft === 0
            ? `${p.name} is due today`
            : `${p.name} deadline in ${daysLeft} day${daysLeft === 1 ? "" : "s"}`,
        detail: `Due ${deadline}`,
        date: deadline
      });
    }
  });

  (Array.isArray(data?.issues) ? data.issues : []).forEach(issue => {
    if (issue.severity !== "High" || issue.status !== "Open") return;

    const opened = normalizeDateValue(issue.date);
    if (!opened) return;

    const age = daysBetweenDates(opened, today);
    if (age == null || age < STALE_HIGH_SEVERITY_DAYS) return;

    items.push({
      id: `issue-${issue.id}`,
      kind: "issue",
      title: `High severity issue open ${age} day${age === 1 ? "" : "s"}`,
      detail: `${issue.type || "Issue"} — ${issue.project || "Unknown project"}${
        issue.owner ? ` — ${issue.owner}` : ""
      }`,
      date: opened
    });
  });

  // Overdue and long-open issues first, then soonest deadlines.
  const rank = { overdue: 0, issue: 1, deadline: 2 };
  return items.sort((a, b) => {
    const r = (rank[a.kind] ?? 9) - (rank[b.kind] ?? 9);
    if (r !== 0) return r;
    return a.date < b.date ? -1 : a.date > b.date ? 1 : 0;
  });
}


/* =========================================================
   DATE FILTER
========================================================= */

function getPresetRange(key) {
  const today = getTodayISO();
  if (key === "today") return { start: today, end: today };

  if (key === "yesterday") {
    const d = dateKeyToLocalDate(today);
    d.setDate(d.getDate() - 1);
    const iso = toISODate(d.getFullYear(), d.getMonth() + 1, d.getDate());
    return { start: iso, end: iso };
  }

  if (key === "thisWeek") return getWeekRange(today);

  if (key === "lastWeek") {
    const d = dateKeyToLocalDate(today);
    d.setDate(d.getDate() - 7);
    return getWeekRange(toISODate(d.getFullYear(), d.getMonth() + 1, d.getDate()));
  }

  if (key === "thisMonth") return getMonthRange(today.slice(0, 7));

  if (key === "lastMonth") {
    const [y, m] = today.slice(0, 7).split("-").map(Number);
    const prevMonth = m === 1 ? 12 : m - 1;
    const prevYear = m === 1 ? y - 1 : y;
    return getMonthRange(`${prevYear}-${String(prevMonth).padStart(2, "0")}`);
  }

  return { start: "", end: "" };
}

function DateFilter({
  value,
  onChange,
  data
}) {
  const allDates = [
    ...(Array.isArray(data?.sheetRecords)
      ? data.sheetRecords.map(x =>
          normalizeDateValue(x.date)
        )
      : []),

    ...(Array.isArray(data?.accuracyRecords)
      ? data.accuracyRecords.map(x =>
          normalizeDateValue(x.date)
        )
      : [])
  ]
    .filter(Boolean)
    .sort();

  const firstDate =
    allDates[0] || "";

  const lastDate =
    allDates[allDates.length - 1] || "";

  const mode =
    value?.mode || "single";


  const updateFilter = patch => {
    const next = {
      ...value,
      ...patch
    };

    if (next.mode === "single") {
      next.start =
        normalizeDateValue(next.start || "");

      next.end =
        next.start || "";
    } else {
      next.start =
        normalizeDateValue(next.start || "");

      next.end =
        normalizeDateValue(next.end || "");
    }

    onChange(next);
  };


  const switchMode = nextMode => {
    if (nextMode === "range") {
      onChange({
        mode: "range",

        start:
          normalizeDateValue(
            value?.start || firstDate
          ),

        end:
          normalizeDateValue(
            value?.end ||
            lastDate ||
            value?.start ||
            firstDate
          )
      });
    } else {
      const singleDate =
        normalizeDateValue(
          value?.start ||
          lastDate
        );

      onChange({
        mode: "single",
        start: singleDate,
        end: singleDate
      });
    }
  };


  return (
    <div className="date-filter">

      <div className="date-filter-info">
        <b>Report date</b>

        <small>
          {mode === "range"
            ? (
                value?.start &&
                value?.end
                  ? `${value.start} → ${value.end}`
                  : "Select a start and end date"
              )
            : (
                value?.start ||
                "Select a date"
              )}
        </small>
      </div>


      <div className="date-filter-presets">
        {[
          ["today", "Today"],
          ["yesterday", "Yesterday"],
          ["thisWeek", "This week"],
          ["lastWeek", "Last week"],
          ["thisMonth", "This month"],
          ["lastMonth", "Last month"]
        ].map(([key, label]) => (
          <button
            key={key}
            type="button"
            className="date-preset-btn"
            onClick={() => {
              const preset = getPresetRange(key);
              onChange({ mode: "range", start: preset.start, end: preset.end });
            }}
          >
            {label}
          </button>
        ))}
      </div>

      <div className="date-filter-controls">

        <div className="date-filter-mode">

          <button
            type="button"
            className={
              mode === "single"
                ? "active"
                : ""
            }
            onClick={() =>
              switchMode("single")
            }
          >
            Single date
          </button>


          <button
            type="button"
            className={
              mode === "range"
                ? "active"
                : ""
            }
            onClick={() =>
              switchMode("range")
            }
          >
            From → To
          </button>

        </div>


        {mode === "single" ? (

          <input
            type="date"
            value={
              normalizeDateValue(
                value?.start || ""
              )
            }
            min={
              firstDate || undefined
            }
            max={
              lastDate || undefined
            }
            onChange={e =>
              updateFilter({
                start: e.target.value
              })
            }
            aria-label="Select report date"
          />

        ) : (

          <div className="date-range-inputs">

            <label>
              <span>From</span>

              <input
                type="date"
                value={
                  normalizeDateValue(
                    value?.start || ""
                  )
                }
                min={
                  firstDate || undefined
                }
                max={
                  value?.end ||
                  lastDate ||
                  undefined
                }
                onChange={e =>
                  updateFilter({
                    start: e.target.value
                  })
                }
                aria-label="Report start date"
              />
            </label>


            <span className="date-range-arrow">
              →
            </span>


            <label>
              <span>To</span>

              <input
                type="date"
                value={
                  normalizeDateValue(
                    value?.end || ""
                  )
                }
                min={
                  value?.start ||
                  firstDate ||
                  undefined
                }
                max={
                  lastDate || undefined
                }
                onChange={e =>
                  updateFilter({
                    end: e.target.value
                  })
                }
                aria-label="Report end date"
              />
            </label>

          </div>

        )}

      </div>
    </div>
  );
}


/* =========================================================
   APP
========================================================= */
function App() {
  const [session, setSession] = useState(null);
  const [profile, setProfile] = useState(null);
  const [authLoading, setAuthLoading] = useState(isSupabaseConfigured);

  useEffect(() => {
    if (!isSupabaseConfigured || !supabase) {
      setAuthLoading(false);
      return undefined;
    }

    let active = true;

    const loadSession = async () => {
      const { data: sessionData } = await supabase.auth.getSession();
      if (!active) return;
      setSession(sessionData.session || null);
      setAuthLoading(false);
    };

    loadSession();

    const { data: listener } = supabase.auth.onAuthStateChange(
      (_event, nextSession) => {
        setSession(nextSession || null);
        if (!nextSession) setProfile(null);
      }
    );

    return () => {
      active = false;
      listener.subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    if (!session || !supabase) return;

    let active = true;

    const loadProfile = async () => {
      const user = session.user;
      const fallbackName =
        user.user_metadata?.full_name ||
        user.user_metadata?.name ||
        user.email?.split("@")[0] ||
        "User";

      // Read the complete profile row instead of selecting only full_name.
      // This keeps the login compatible with profiles tables that use either
      // full_name or name for the display-name column.
      const { data: row, error } = await supabase
        .from("profiles")
        .select("*")
        .eq("id", user.id)
        .limit(1)
        .maybeSingle();

      if (!active) return;

      if (error) {
        console.error("Could not load profile row", error);
      }

      let resolvedRole = row?.role;

      // The database schema also exposes get_my_role() as a safe fallback.
      // This is useful if a profile SELECT is temporarily blocked by RLS.
      if (!resolvedRole) {
        const { data: roleData, error: roleError } = await supabase.rpc("get_my_role");
        if (!roleError && roleData) {
          resolvedRole = roleData;
        } else if (roleError) {
          console.warn("Could not load role through get_my_role()", roleError);
        }
      }

      const normalizedRole = ["admin", "team_lead", "member"].includes(
        String(resolvedRole || "").toLowerCase()
      )
        ? String(resolvedRole).toLowerCase()
        : "member";

      setProfile({
        id: user.id,
        full_name: row?.full_name || row?.name || fallbackName,
        role: normalizedRole
      });

      // Self-heal missing emails: since every user logs in themselves,
      // each login is a safe chance to record their own email on their
      // own profile row (RLS already has to allow this — it's the same
      // permission as updating your own name/role would need). This is
      // what lets User & access show a real email instead of a raw ID,
      // without needing any admin-only backend access.
      if (user.email && row?.email !== user.email) {
        const { error: emailError } = await supabase
          .from("profiles")
          .update({ email: user.email })
          .eq("id", user.id);

        if (emailError) {
          console.warn("Could not backfill email on profile", emailError);
        }
      }
    };

    loadProfile();
    return () => { active = false; };
  }, [session]);

  if (authLoading) {
    return <AuthLoading />;
  }

  if (isSupabaseConfigured && !session) {
    return <LoginScreen />;
  }

  if (isSupabaseConfigured && !profile) {
    return <AuthLoading text="Loading your account..." />;
  }

  return (
    <DashboardApp
      session={session}
      profile={profile || { full_name: "Manjunath", role: "team_lead" }}
      onSignOut={() => supabase?.auth.signOut()}
    />
  );
}

function DashboardApp({ session, profile, onSignOut }) {
  const [data, setData] = useState(loadData);
  const [dataReady, setDataReady] = useState(!isSupabaseConfigured);
  const [page, setPage] = useState(() => loadPrefs().landingPage || "dashboard");
  const [query, setQuery] = useState("");
  const [sidebar, setSidebar] = useState(true);
  const [showAdd, setShowAdd] = useState(false);
  const [addType, setAddType] = useState("team");
  const [editingMember, setEditingMember] = useState(null);
  const [viewingProfile, setViewingProfile] = useState(null);
  const [editingProject, setEditingProject] = useState(null);
  const [toast, setToast] = useState("");
  const [showNotif, setShowNotif] = useState(false);
  const notifRef = useRef(null);

  const notifications = useMemo(() => getNotifications(data), [data.projects, data.issues]);

  useEffect(() => {
    if (!showNotif) return;
    function handleClickAway(e) {
      if (notifRef.current && !notifRef.current.contains(e.target)) {
        setShowNotif(false);
      }
    }
    document.addEventListener("mousedown", handleClickAway);
    return () => document.removeEventListener("mousedown", handleClickAway);
  }, [showNotif]);

  const role = profile?.role || "member";
  const canManageTeam = ["admin", "team_lead"].includes(role);
  const canManageProjects = ["admin", "team_lead"].includes(role);
  const canManageIssues = ["admin", "team_lead"].includes(role);
  const canImport = ["admin", "team_lead"].includes(role);

  useEffect(() => {
    if (!isSupabaseConfigured || !supabase || !profile) return;

    let active = true;

    const loadCloudData = async () => {
      const { data: row, error } = await supabase
        .from("dashboard_state")
        .select("data")
        .eq("id", 1)
        .maybeSingle();

      if (!active) return;

      if (error) {
        console.error("Could not load online dashboard data", error);
        notify("Could not load online data. Using local backup.");
        setDataReady(true);
        return;
      }

      if (row?.data && Object.keys(row.data).length) {
        const cloud = row.data;
        const normalized = {
          ...seed,
          ...cloud,
          team: Array.isArray(cloud.team) ? cloud.team : seed.team,
          projects: Array.isArray(cloud.projects)
            ? cloud.projects.map(p => {
                const stats = getProjectStats(p);
                return { ...p, ...stats, target: stats.total, total: stats.total };
              })
            : seed.projects,
          issues: Array.isArray(cloud.issues) ? cloud.issues : seed.issues,
          sheetRecords: Array.isArray(cloud.sheetRecords) ? cloud.sheetRecords : [],
          accuracyRecords: Array.isArray(cloud.accuracyRecords) ? cloud.accuracyRecords : [],
          accuracyFile: cloud.accuracyFile || "",
          accuracyLastSync: cloud.accuracyLastSync || "",
          sheetFile: cloud.sheetFile || "",
          sheetLastSync: cloud.sheetLastSync || ""
        };
        setData(normalized);
        saveData(normalized);
      } else if (canManageTeam) {
        const local = loadData();
        setData(local);
        const { error: seedError } = await supabase
          .from("dashboard_state")
          .upsert({ id: 1, data: local, updated_at: new Date().toISOString() });
        if (seedError) console.error("Could not seed online data", seedError);
      }

      setDataReady(true);
    };

    loadCloudData();
    return () => { active = false; };
  }, [profile, canManageTeam]);

  const update = async next => {
    if (isSupabaseConfigured && supabase && !canManageTeam) {
      notify("Your role is view-only for this action.");
      return false;
    }

    setData(next);
    saveData(next);

    if (!isSupabaseConfigured || !supabase || !dataReady) return true;

    const { error } = await supabase
      .from("dashboard_state")
      .upsert({ id: 1, data: next, updated_at: new Date().toISOString() });

    if (error) {
      console.error("Online save failed", error);
      notify("Online save failed. Your local copy was kept.");
      return false;
    }

    return true;
  };

  const notify = msg => {
    setToast(msg);
    setTimeout(() => setToast(""), 2500);
  };

  const nav = [
    ["dashboard", "Dashboard", LayoutDashboard],
    ["team", "Team", Users],
    ["attendance", "Attendance", Calendar],
    ["projects", "Projects", FolderKanban],
    ["qa", "QA & Reviews", ClipboardCheck],
    ["issues", "Issues", AlertTriangle],
    ["analytics", "Analytics", BarChart3],
    ["reports", "Reports", FileText],
    ["settings", "Settings", Settings],
    ["sheet", "Sheet Import", Upload]
  ];

  const filteredTeam = useMemo(
    () =>
      data.team.filter(x =>
        String(x.name || "").toLowerCase().includes(query.toLowerCase())
      ),
    [data.team, query]
  );

  const totals = useMemo(() => {
    // Computed fresh from sheetRecords + the live "latest work date" —
    // never from the team.completed/target snapshot, which is only as
    // current as the last import and can drift out of sync with
    // Attendance (which always recalculates from scratch).
    const latest = getLatestImportedDate(data);
    const todayRows = (Array.isArray(data.sheetRecords) ? data.sheetRecords : []).filter(
      r => r.date === latest && isWorkRow(r)
    );

    const done = todayRows.reduce((s, x) => s + (Number(x.worked) || 0), 0);
    const reviewed = todayRows
      .filter(x => /review/i.test(String(x.type || "")))
      .reduce((s, x) => s + (Number(x.worked) || 0), 0);

    // Same "target = sum of each worked project's daily target" logic
    // parseWorkbook uses, just recalculated live instead of stored.
    const byPerson = new Map();
    todayRows.forEach(r => {
      const key = r.name;
      if (!byPerson.has(key)) byPerson.set(key, new Set());
      byPerson.get(key).add(getConfiguredProjectName(r.project));
    });
    let total = 0;
    byPerson.forEach(projectSet => {
      projectSet.forEach(pname => {
        total += Number(projectTargets[pname]) || 0;
      });
    });

    const remaining = data.projects.reduce(
      (s, x) => s + getProjectStats(x).remaining,
      0
    );

    return {
      total,
      done,
      reviewed,
      remaining,
      rate: total ? Math.round((done / total) * 100) : 0
    };
  }, [data]);

  function addRecord(e) {
    e.preventDefault();
    if (addType === "team" && !canManageTeam) return notify("Only Admin or Team Lead can manage team members.");
    if (addType === "project" && !canManageProjects) return notify("Only Admin or Team Lead can manage projects.");
    if (addType === "issue" && !canManageIssues) return notify("Only Admin or Team Lead can manage issues.");
    const f = new FormData(e.currentTarget);
    let next = { ...data };

    if (addType === "team") {
      const target = Number(f.get("target") || 1000);
      next.team = [
        ...data.team,
        {
          id: Date.now(),
          name: f.get("name"),
          role: f.get("role"),
          target,
          completed: Number(f.get("completed") || 0),
          reviewed: 0,
          errors: 0,
          status: "Active"
        }
      ];
    } else if (addType === "project") {
      const name = String(f.get("name") || "").trim();
      const totalImages = Math.max(0, Number(f.get("totalImages") || 0));
      const target = Math.max(0, Number(f.get("target") || 0));
      const completed = Math.max(0, Math.min(totalImages, Number(f.get("completed") || 0)));
      const remaining = Math.max(0, totalImages - completed);

      next.projects = [
        {
          id: Date.now(),
          name,
          totalImages,
          total: totalImages,
          target,
          completed,
          remaining,
          status: getProjectStatus(totalImages, completed, remaining),
          deadline: f.get("deadline") || "",
          archived: false,
          assignedEmployees: f.getAll("assignedEmployees"),
          assignedReviewers: f.getAll("assignedReviewers")
        },
        ...data.projects
      ];
    } else {
      next.issues = [
        ...data.issues,
        {
          id: Date.now(),
          type: f.get("type"),
          description: f.get("description") || "",
          project: f.get("project"),
          owner: f.get("owner"),
          assignedTo: f.get("assignedTo") || "",
          severity: f.get("severity"),
          status: "Open",
          date: new Date().toISOString().slice(0, 10),
          dueDate: f.get("dueDate") || "",
          resolution: "",
          resolvedDate: "",
          comments: []
        }
      ];
    }

    update(next);
    setShowAdd(false);
    notify("Saved successfully");
    logActivity(
      addType === "team"
        ? `Added team member "${f.get("name")}"`
        : addType === "project"
        ? `Added project "${f.get("name")}"`
        : `Logged issue "${f.get("type")}" on ${f.get("project")}`
    );
  }

  function editProject(e) {
    e.preventDefault();
    if (!canManageProjects) return notify("Only Admin or Team Lead can manage projects.");
    const f = new FormData(e.currentTarget);
    if (!editingProject) return;

    const name = String(f.get("name") || "").trim();
    const totalImages = Math.max(0, Number(f.get("totalImages") || 0));
    const target = Math.max(0, Number(f.get("target") || 0));
    const completed = Math.max(0, Math.min(totalImages, Number(f.get("completed") || 0)));
    const remaining = Math.max(0, totalImages - completed);

    const projects = data.projects.map(project =>
      project.id === editingProject.id
        ? {
            ...project,
            name,
            totalImages,
            target,
            total: totalImages,
            completed,
            remaining,
            status: getProjectStatus(totalImages, completed, remaining),
            deadline: f.get("deadline") || "",
            assignedEmployees: f.getAll("assignedEmployees"),
            assignedReviewers: f.getAll("assignedReviewers")
          }
        : project
    );

    update({ ...data, projects });
    setEditingProject(null);
    notify("Project updated successfully");
    logActivity(`Updated project "${name}"`);
  }

  function archiveProject(id) {
    if (!canManageProjects) return notify("Only Admin or Team Lead can archive projects.");

    const project = data.projects.find(p => p.id === id);
    if (!project) return;
    const archived = !project.archived;

    update({
      ...data,
      projects: data.projects.map(p => (p.id === id ? { ...p, archived } : p))
    });
    notify(archived ? `${project.name} archived` : `${project.name} unarchived`);
    logActivity(`${archived ? "Archived" : "Unarchived"} project "${project.name}"`);
  }

  function updateIssue(id, changes) {
    if (!canManageIssues) return notify("Only Admin or Team Lead can update issues.");

    const issues = data.issues.map(issue => {
      if (issue.id !== id) return issue;
      const next = { ...issue, ...changes };
      // Stamp a resolved date automatically the moment status becomes
      // Resolved/Closed, and clear it if it's reopened.
      if (changes.status) {
        if (["Resolved", "Closed"].includes(changes.status) && !issue.resolvedDate) {
          next.resolvedDate = new Date().toISOString().slice(0, 10);
        } else if (!["Resolved", "Closed"].includes(changes.status)) {
          next.resolvedDate = "";
        }
      }
      return next;
    });

    update({ ...data, issues });
    logActivity(`Updated issue "${data.issues.find(i => i.id === id)?.type || id}"`);
  }

  function addIssueComment(id, text) {
    if (!text || !text.trim()) return;
    const issue = data.issues.find(i => i.id === id);
    if (!issue) return;
    const comments = [
      ...(Array.isArray(issue.comments) ? issue.comments : []),
      { id: Date.now(), text: text.trim(), at: new Date().toISOString() }
    ];
    updateIssue(id, { comments });
  }

  function remove(kind, id) {
    if (kind === "projects" && !canManageProjects) return notify("Only Admin or Team Lead can delete projects.");
    if (kind === "issues" && !canManageIssues) return notify("Only Admin or Team Lead can delete issues.");
    if (!confirm("Delete this record?")) return;

    const removedName = data[kind].find(x => x.id === id)?.name || data[kind].find(x => x.id === id)?.type;

    update({
      ...data,
      [kind]: data[kind].filter(x => x.id !== id)
    });

    notify("Deleted");
    logActivity(`Deleted ${kind.slice(0, -1)}${removedName ? ` "${removedName}"` : ""}`);
  }

  function saveTeamMember(formData, editing) {
    if (!canManageTeam) return notify("Only Admin or Team Lead can manage team members.");

    const name = String(formData.get("name") || "").trim();
    if (!name) return;

    const member = {
      id: editing?.id || Date.now(),
      name,
      role: String(formData.get("role") || "Annotator"),
      target: Number(formData.get("target") || 0),
      completed: Number(formData.get("completed") || 0),
      reviewed: Number(formData.get("reviewed") || 0),
      errors: Number(formData.get("errors") || 0),
      status: String(formData.get("status") || "Active")
    };

    const nextTeam = editing
      ? data.team.map(x => x.id === editing.id ? member : x)
      : [...data.team, member];

    update({ ...data, team: nextTeam });
    setEditingMember(null);
    setShowAdd(false);
    notify(editing ? "Team member updated" : "Team member added");
  }

  function removeTeamMember(name) {
    if (!canManageTeam) return notify("Only Admin or Team Lead can delete team members.");
    if (!confirm(`Delete ${name} from the team?`)) return;

    const next = {
      ...data,
      team: data.team.filter(x => String(x.name).toLowerCase() !== String(name).toLowerCase()),
      sheetRecords: Array.isArray(data.sheetRecords)
        ? data.sheetRecords.filter(x => String(x.name).toLowerCase() !== String(name).toLowerCase())
        : data.sheetRecords
    };

    update(next);
    notify("Team member deleted");
    logActivity(`Deleted team member "${name}"`);
  }

  function toggleMemberStatus(name) {
    if (!canManageTeam) return notify("Only Admin or Team Lead can deactivate team members.");

    const member = data.team.find(x => String(x.name).toLowerCase() === String(name).toLowerCase());
    if (!member) return;

    const nextStatus = member.status === "Inactive" ? "Active" : "Inactive";

    const next = {
      ...data,
      team: data.team.map(x =>
        String(x.name).toLowerCase() === String(name).toLowerCase()
          ? { ...x, status: nextStatus }
          : x
      )
    };

    update(next);
    notify(nextStatus === "Inactive" ? `${name} deactivated` : `${name} reactivated`);
    logActivity(`${nextStatus === "Inactive" ? "Deactivated" : "Reactivated"} team member "${name}"`);
  }

  function clearImportedData() {
    if (!canManageTeam) return notify("Only Admin or Team Lead can clear imported data.");
    if (!confirm("Clear all imported sheet data and reset project completed counts? This cannot be undone.")) return;

    const projects = data.projects.map(p => {
      const totalImages = Math.max(0, Number(p.totalImages ?? p.total) || 0);
      return {
        ...p,
        completed: 0,
        remaining: totalImages,
        status: getProjectStatus(totalImages, 0, totalImages)
      };
    });

    update({
      ...data,
      sheetRecords: [],
      accuracyRecords: [],
      sheetFile: "",
      accuracyFile: "",
      sheetLastSync: "",
      projects
    });

    notify("Imported data cleared");
    logActivity("Cleared imported sheet data and reset project completed counts");
  }

  function exportData() {
    const blob = new Blob(
      [JSON.stringify(data, null, 2)],
      { type: "application/json" }
    );

    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "annotatepro-backup.json";
    a.click();
    URL.revokeObjectURL(a.href);
    notify("Backup exported");
  }

  function importData(e) {
    const file = e.target.files?.[0];
    if (!file) return;

    const r = new FileReader();

    r.onload = () => {
      try {
        const parsed = JSON.parse(r.result);

        if (parsed.team && parsed.projects && parsed.issues) {
          const projects = parsed.projects.map(p => {
            const s = getProjectStats(p);
            // Keep p.target (Daily Target) exactly as saved in the backup —
            // it's a separate number from Total Images and must not be
            // overwritten here.
            return {
              ...p,
              ...s
            };
          });

          update({ ...parsed, projects });
          notify("Backup imported");
        } else {
          alert("Invalid dashboard backup");
        }
      } catch {
        alert("Invalid JSON file");
      }
    };

    r.readAsText(file);
    e.target.value = "";
  }

  return (
    <div className="app">
      <aside className={"sidebar " + (!sidebar ? "collapsed" : "")}>
        <div className="brand">
          <div className="brand-mark">A</div>
          {sidebar && (
            <div>
              <b>AnnotatePro</b>
              <span>Team Operations</span>
            </div>
          )}
        </div>

        <nav>
          {nav.map(([key, label, Icon]) => (
            <button
              key={key}
              className={page === key ? "active" : ""}
              onClick={() => setPage(key)}
            >
              <Icon size={19} />
              {sidebar && <span>{label}</span>}
            </button>
          ))}
        </nav>

        {sidebar && (
          <div className="side-card">
            <ShieldCheck size={20} />
            <b>Quality first</b>
            <p>Track annotation accuracy and resolve issues early.</p>
          </div>
        )}

        <button className="collapse" onClick={() => setSidebar(!sidebar)}>
          {sidebar ? <ChevronRight size={18} /> : <Menu size={18} />}
          <span>{sidebar ? "Collapse" : "Expand"}</span>
        </button>
      </aside>

      <main className="main">
        <header className="topbar">
          <div className="mobile-menu" onClick={() => setSidebar(!sidebar)}>
            <Menu />
          </div>

          <div className="search">
            <Search size={18} />
            <input
              placeholder="Search team members..."
              value={query}
              onChange={e => setQuery(e.target.value)}
            />
          </div>

          <div className="top-actions">
            <div className="notif-wrap" ref={notifRef}>
              <button
                className="icon-btn"
                onClick={() => setShowNotif(v => !v)}
                aria-label="Notifications"
              >
                <Bell size={19} />
                {notifications.length > 0 && <i />}
              </button>

              {showNotif && (
                <div className="notif-panel">
                  <div className="notif-panel-head">
                    <b>Notifications</b>
                    <span className="muted">{notifications.length}</span>
                  </div>

                  {notifications.length === 0 ? (
                    <p className="notif-empty">You're all caught up.</p>
                  ) : (
                    <ul className="notif-list">
                      {notifications.map(n => (
                        <li key={n.id} className={`notif-item notif-${n.kind}`}>
                          <span className="notif-dot" />
                          <div>
                            <p className="notif-title">{n.title}</p>
                            <p className="notif-detail">{n.detail}</p>
                          </div>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              )}
            </div>

            <div className="avatar">{String(profile?.full_name || "U").slice(0, 2).toUpperCase()}</div>
            <div className="user">
              <b>{profile?.full_name || session?.user?.email || "User"}</b>
              <span>{roleLabel(role)}</span>
            </div>
            <button className="icon-btn" title="Sign out" onClick={onSignOut}>
              <LogOut size={18} />
            </button>
          </div>
        </header>

        <section className="content">
          {page === "dashboard" && (
            <Dashboard totals={totals} data={data} setPage={setPage} />
          )}

          {page === "team" && (
            <Team
              rows={filteredTeam}
              data={data}
              canManage={canManageTeam}
              openAdd={() => {
                if (!canManageTeam) return notify("Only Admin or Team Lead can add members.");
                setEditingMember(null);
                setAddType("team");
                setShowAdd(true);
              }}
              onEdit={member => { setEditingMember(member); setAddType("team"); setShowAdd(true); }}
              onDelete={removeTeamMember}
              onToggleStatus={toggleMemberStatus}
              onView={name => setViewingProfile(name)}
            />
          )}

          {page === "attendance" && (
            <Attendance
              data={data}
              update={update}
              canManage={canManageTeam}
              notify={notify}
            />
          )}

          {page === "projects" && (
            <Projects
              rows={data.projects}
              data={data}
              remove={remove}
              canManage={canManageProjects}
              openAdd={() => {
                setAddType("project");
                setShowAdd(true);
              }}
              openEdit={setEditingProject}
              onArchive={archiveProject}
            />
          )}

          {page === "qa" && <QA data={data} />}

          {page === "issues" && (
            <Issues
              rows={data.issues}
              data={data}
              remove={remove}
              canManage={canManageIssues}
              onUpdate={updateIssue}
              onComment={addIssueComment}
              openAdd={() => {
                setAddType("issue");
                setShowAdd(true);
              }}
            />
          )}

          {page === "analytics" && <Analytics data={data} />}

          {page === "reports" && <Reports data={data} />}

          {page === "settings" && (
            <SettingsPage
              exportData={exportData}
              importData={importData}
              clearImportedData={clearImportedData}
              role={role}
              storageOnline={isSupabaseConfigured}
              email={session?.user?.email}
              onSignOut={onSignOut}
              canManage={canManageTeam}
              notify={notify}
              myId={profile?.id}
              isAdmin={role === "admin"}
            />
          )}

          {page === "sheet" && (
            canImport ? (
              <SheetImport data={data} update={update} notify={notify} />
            ) : (
              <Page title="Sheet Import" subtitle="Daily imports are restricted to Admin and Team Lead accounts.">
                <Panel title="Permission required">
                  <p className="muted">Your current role can view dashboard data but cannot import or overwrite daily sheet data.</p>
                </Panel>
              </Page>
            )
          )}
        </section>
      </main>

      {showAdd && addType === "team" && (
        <TeamMemberModal
          member={editingMember}
          onClose={() => { setShowAdd(false); setEditingMember(null); }}
          onSubmit={saveTeamMember}
        />
      )}

      {showAdd && addType !== "team" && (
        <Modal
          type={addType}
          team={data.team}
          onClose={() => setShowAdd(false)}
          onSubmit={addRecord}
        />
      )}

      {editingProject && (
        <ProjectEditModal
          project={editingProject}
          team={data.team}
          onClose={() => setEditingProject(null)}
          onSubmit={editProject}
        />
      )}

      {viewingProfile && (
        <TeamProfileModal
          name={viewingProfile}
          data={data}
          onClose={() => setViewingProfile(null)}
        />
      )}

      {toast && (
        <div className="toast">
          <CheckCircle2 size={18} />
          {toast}
        </div>
      )}
    </div>
  );
}

function roleLabel(role) {
  return ({ admin: "Admin", team_lead: "Team Lead", member: "Member" }[role] || "Member");
}

function AuthLoading({ text = "Loading AnnotatePro..." }) {
  return (
    <div className="auth-shell">
      <div className="auth-card auth-loading">
        <div className="brand-mark">A</div>
        <h1>AnnotatePro</h1>
        <p>{text}</p>
      </div>
    </div>
  );
}

function LoginScreen() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function handleLogin(e) {
    e.preventDefault();

    if (!supabase) {
      setError(
        "Supabase is not configured."
      );

      return;
    }

    setBusy(true);
    setError("");

    const { data, error: loginError } =
      await supabase.auth.signInWithPassword({
        email,
        password
      });

    if (loginError) {
      setError(loginError.message);
      setBusy(false);
      return;
    }

    setBusy(false);
  }

  return (
    <div
      className="app"
      style={{
        minHeight: "100vh",
        width: "100%",
        background:
          "radial-gradient(circle at 15% 15%, rgba(99,102,241,0.16), transparent 34%), radial-gradient(circle at 85% 85%, rgba(14,165,233,0.13), transparent 32%), #f7f9fc",
        position: "relative",
        overflow: "hidden"
      }}
    >
      <div
        style={{
          position: "absolute",
          inset: 0,
          pointerEvents: "none",
          background:
            "linear-gradient(135deg, rgba(255,255,255,0.78), rgba(255,255,255,0.18))"
        }}
      />

      <div
        style={{
          minHeight: "100vh",
          width: "100%",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: "32px 20px",
          position: "relative",
          zIndex: 1
        }}
      >
        <div
          className="annotatepro-login-grid"
          style={{
            width: "100%",
            maxWidth: "980px",
            display: "grid",
            gridTemplateColumns: "minmax(0, 1.05fr) minmax(360px, 0.95fr)",
            borderRadius: "28px",
            overflow: "hidden",
            background: "rgba(255,255,255,0.94)",
            border: "1px solid rgba(148,163,184,0.22)",
            boxShadow: "0 30px 80px rgba(15,23,42,0.14)"
          }}
        >
          <div
            style={{
              padding: "52px",
              minHeight: "590px",
              display: "flex",
              flexDirection: "column",
              justifyContent: "space-between",
              background:
                "linear-gradient(145deg, #111827 0%, #1e293b 58%, #312e81 100%)",
              color: "#ffffff",
              position: "relative",
              overflow: "hidden"
            }}
          >
            <div
              style={{
                position: "absolute",
                width: "280px",
                height: "280px",
                borderRadius: "50%",
                background: "rgba(129,140,248,0.18)",
                top: "-120px",
                right: "-90px"
              }}
            />

            <div
              style={{
                position: "absolute",
                width: "220px",
                height: "220px",
                borderRadius: "50%",
                background: "rgba(56,189,248,0.12)",
                bottom: "-100px",
                left: "-80px"
              }}
            />

            <div style={{ position: "relative", zIndex: 1 }}>
              <div
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: "12px",
                  marginBottom: "42px"
                }}
              >
                <div
                  style={{
                    width: "48px",
                    height: "48px",
                    borderRadius: "15px",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: "22px",
                    fontWeight: 800,
                    background: "rgba(255,255,255,0.12)",
                    border: "1px solid rgba(255,255,255,0.18)",
                    boxShadow: "0 10px 25px rgba(0,0,0,0.16)"
                  }}
                >
                  A
                </div>

                <div>
                  <div
                    style={{
                      fontSize: "20px",
                      fontWeight: 800,
                      letterSpacing: "-0.02em"
                    }}
                  >
                    AnnotatePro
                  </div>
                  <div
                    style={{
                      fontSize: "12px",
                      color: "rgba(255,255,255,0.62)",
                      marginTop: "2px"
                    }}
                  >
                    Team Operations
                  </div>
                </div>
              </div>

              <div style={{ maxWidth: "470px" }}>
                <div
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    padding: "7px 11px",
                    borderRadius: "999px",
                    background: "rgba(255,255,255,0.09)",
                    border: "1px solid rgba(255,255,255,0.12)",
                    fontSize: "12px",
                    fontWeight: 700,
                    color: "#c7d2fe",
                    marginBottom: "18px"
                  }}
                >
                  Operations • Quality • Performance
                </div>

                <h1
                  style={{
                    margin: 0,
                    fontSize: "clamp(34px, 4vw, 48px)",
                    lineHeight: 1.05,
                    letterSpacing: "-0.045em",
                    color: "#ffffff"
                  }}
                >
                  Manage your team’s work with confidence.
                </h1>

                <p
                  style={{
                    margin: "20px 0 0",
                    fontSize: "15px",
                    lineHeight: 1.75,
                    color: "rgba(255,255,255,0.70)",
                    maxWidth: "430px"
                  }}
                >
                  A focused workspace for tracking daily production,
                  project progress, QA reviews, issues, and team performance
                  in one place.
                </p>
              </div>
            </div>

            <div
              style={{
                position: "relative",
                zIndex: 1,
                display: "grid",
                gridTemplateColumns: "repeat(3, 1fr)",
                gap: "10px",
                marginTop: "40px"
              }}
            >
              {[
                ["01", "Daily tracking"],
                ["02", "QA visibility"],
                ["03", "Team insights"]
              ].map(([number, label]) => (
                <div
                  key={number}
                  style={{
                    padding: "13px 12px",
                    borderRadius: "13px",
                    background: "rgba(255,255,255,0.07)",
                    border: "1px solid rgba(255,255,255,0.09)"
                  }}
                >
                  <div
                    style={{
                      fontSize: "10px",
                      fontWeight: 800,
                      color: "#a5b4fc",
                      marginBottom: "5px"
                    }}
                  >
                    {number}
                  </div>
                  <div
                    style={{
                      fontSize: "11px",
                      color: "rgba(255,255,255,0.72)",
                      fontWeight: 600
                    }}
                  >
                    {label}
                  </div>
                </div>
              ))}
            </div>
          </div>

          <div
            className="annotatepro-login-form-panel"
            style={{
              padding: "52px 46px",
              display: "flex",
              alignItems: "center",
              background: "#ffffff"
            }}
          >
            <div style={{ width: "100%", maxWidth: "390px", margin: "0 auto" }}>
              <div style={{ marginBottom: "30px" }}>
                <div
                  style={{
                    width: "44px",
                    height: "44px",
                    borderRadius: "13px",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    background: "#eef2ff",
                    color: "#4f46e5",
                    marginBottom: "18px"
                  }}
                >
                  <LockKeyhole size={20} />
                </div>

                <h2
                  style={{
                    margin: 0,
                    fontSize: "30px",
                    lineHeight: 1.15,
                    letterSpacing: "-0.035em",
                    color: "#0f172a"
                  }}
                >
                  Welcome back
                </h2>

                <p
                  style={{
                    margin: "9px 0 0",
                    color: "#64748b",
                    fontSize: "14px",
                    lineHeight: 1.6
                  }}
                >
                  Sign in to continue to your dashboard.
                </p>
              </div>

              <form onSubmit={handleLogin}>
                <div style={{ marginBottom: "20px" }}>
                  <label
                    htmlFor="login-email"
                    style={{
                      display: "block",
                      marginBottom: "8px",
                      fontSize: "13px",
                      fontWeight: 700,
                      color: "#334155"
                    }}
                  >
                    Email address
                  </label>
                  <input
                    id="login-email"
                    type="email"
                    value={email}
                    onChange={e => setEmail(e.target.value)}
                    placeholder="you@example.com"
                    autoComplete="email"
                    required
                    style={{
                      width: "100%",
                      boxSizing: "border-box",
                      height: "50px",
                      padding: "0 15px",
                      borderRadius: "12px",
                      border: "1px solid #dbe2ea",
                      background: "#f8fafc",
                      color: "#0f172a",
                      fontSize: "14px",
                      outline: "none"
                    }}
                  />
                </div>

                <div style={{ marginBottom: "18px" }}>
                  <label
                    htmlFor="login-password"
                    style={{
                      display: "block",
                      marginBottom: "8px",
                      fontSize: "13px",
                      fontWeight: 700,
                      color: "#334155"
                    }}
                  >
                    Password
                  </label>

                  <div style={{ position: "relative" }}>
                    <input
                      id="login-password"
                      type={showPassword ? "text" : "password"}
                      value={password}
                      onChange={e => setPassword(e.target.value)}
                      placeholder="Enter your password"
                      autoComplete="current-password"
                      required
                      style={{
                        width: "100%",
                        boxSizing: "border-box",
                        height: "50px",
                        padding: "0 78px 0 15px",
                        borderRadius: "12px",
                        border: "1px solid #dbe2ea",
                        background: "#f8fafc",
                        color: "#0f172a",
                        fontSize: "14px",
                        outline: "none"
                      }}
                    />

                    <button
                      type="button"
                      onClick={() => setShowPassword(value => !value)}
                      style={{
                        position: "absolute",
                        top: "50%",
                        right: "10px",
                        transform: "translateY(-50%)",
                        border: 0,
                        background: "transparent",
                        color: "#64748b",
                        fontSize: "12px",
                        fontWeight: 700,
                        cursor: "pointer",
                        padding: "7px 8px",
                        borderRadius: "8px"
                      }}
                    >
                      {showPassword ? "Hide" : "Show"}
                    </button>
                  </div>
                </div>

                {error && (
                  <div
                    role="alert"
                    style={{
                      marginBottom: "18px",
                      padding: "12px 13px",
                      borderRadius: "11px",
                      background: "#fff1f2",
                      border: "1px solid #fecdd3",
                      color: "#be123c",
                      fontSize: "13px",
                      lineHeight: 1.5
                    }}
                  >
                    {error}
                  </div>
                )}

                <button
                  className="primary"
                  type="submit"
                  disabled={busy}
                  style={{
                    width: "100%",
                    minHeight: "50px",
                    justifyContent: "center",
                    borderRadius: "12px",
                    fontSize: "14px",
                    fontWeight: 700,
                    boxShadow: "0 10px 24px rgba(79,70,229,0.18)",
                    opacity: busy ? 0.72 : 1,
                    cursor: busy ? "not-allowed" : "pointer"
                  }}
                >
                  <LockKeyhole size={17} />
                  {busy ? "Signing in..." : "Sign in to dashboard"}
                </button>
              </form>

              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "10px",
                  marginTop: "26px",
                  color: "#94a3b8",
                  fontSize: "11px"
                }}
              >
                <div style={{ height: "1px", flex: 1, background: "#e2e8f0" }} />
                <span>Secure team access</span>
                <div style={{ height: "1px", flex: 1, background: "#e2e8f0" }} />
              </div>
            </div>
          </div>
        </div>
      </div>

      <style>{`
        @media (max-width: 820px) {
          .annotatepro-login-grid {
            grid-template-columns: 1fr !important;
          }
        }

        @media (max-width: 820px) {
          .annotatepro-login-grid > div:first-child {
            display: none !important;
          }
        }

        @media (max-width: 480px) {
          .annotatepro-login-form-panel {
            padding: 34px 24px !important;
          }
        }
      `}</style>
    </div>
  );
}

/* =========================================================
   DASHBOARD
========================================================= */
function Dashboard({ totals, data, setPage }) {
  const latestDate = getLatestImportedDate(data);

  // Today's attendance snapshot — reuses the exact same logic as the
  // Attendance page (buildAttendanceMatrix) for a single day, so the
  // numbers here always agree with what Attendance shows.
  const attendanceToday = useMemo(() => {
    if (!latestDate) return { rows: [], summary: {} };
    return buildAttendanceMatrix(data, { start: latestDate, end: latestDate });
  }, [data, latestDate]);

  const presentToday = attendanceToday.summary["Present"] || 0;
  const absentToday = (attendanceToday.summary["No data"] || 0) + (attendanceToday.summary["Absent"] || 0);
  const leaveToday = attendanceToday.summary["Leave"] || 0;
  const absentNames = attendanceToday.rows
    .filter(r => ["No data", "Absent"].includes(r.cells[0]?.status))
    .map(r => r.name);

  const openIssues = data.issues.filter(x => x.status === "Open").length;
  const activeProjects = data.projects.filter(x => x.status === "In Progress").length;

  // Overall (all-time) project progress — Completed is cumulative across
  // the whole imported history, so this reflects true overall standing.
  const overallTarget = data.projects.reduce((s, p) => s + (Number(p.totalImages ?? p.total) || 0), 0);
  const overallCompleted = data.projects.reduce((s, p) => s + (Number(p.completed) || 0), 0);
  const overallCompletion = overallTarget ? Math.round((overallCompleted / overallTarget) * 100) : 0;

  // QA accuracy / errors — all-time totals from imported Accuracy Reports.
  const qaTotals = data.accuracyRecords.reduce(
    (acc, x) => {
      acc.tp += Number(x.tp) || 0;
      acc.fp += Number(x.fp) || 0;
      acc.fn += Number(x.fn) || 0;
      return acc;
    },
    { tp: 0, fp: 0, fn: 0 }
  );
  const qaDenom = qaTotals.tp + qaTotals.fp + qaTotals.fn;
  const qaAccuracy = qaDenom ? Math.round((qaTotals.tp / qaDenom) * 100) : 0;
  const totalErrors = qaTotals.fp + qaTotals.fn;

  const completion = totals.total
    ? Math.round((totals.done / totals.total) * 100)
    : 0;

  // Productivity trend — total images worked per day, last 7 work-days.
  const trend = useMemo(() => {
    const byDate = new Map();
    data.sheetRecords.filter(isWorkRow).forEach(r => {
      byDate.set(r.date, (byDate.get(r.date) || 0) + (Number(r.worked) || 0));
    });
    return [...byDate.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1)).slice(-7);
  }, [data.sheetRecords]);
  const trendMax = Math.max(1, ...trend.map(([, v]) => v));

  // Top performers today, from each team member's daily completed count.
  const topPerformers = [...data.team]
    .filter(x => (Number(x.completed) || 0) > 0)
    .sort((a, b) => (Number(b.completed) || 0) - (Number(a.completed) || 0))
    .slice(0, 5);
  const topMax = Math.max(1, ...topPerformers.map(x => Number(x.completed) || 0));

  // Needing attention — reuses the same logic as the notification bell
  // (overdue/near-deadline projects, stale high-severity issues), plus
  // anyone absent today.
  const notifications = getNotifications(data);

  // Overview panel: surface what's currently being worked on rather than
  // every project. In Progress first, then Pending, Completed/No Target last.
  const statusRank = { "In Progress": 0, "Pending": 1, "No Target": 2, "Completed": 3 };
  const highlightedProjects = [...data.projects]
    .sort((a, b) => {
      const r = (statusRank[a.status] ?? 9) - (statusRank[b.status] ?? 9);
      if (r !== 0) return r;
      return getProjectStats(b).progress - getProjectStats(a).progress;
    })
    .slice(0, 5);

  return (
    <div>
      <div className="page-head">
        <div>
          <p className="eyebrow">TEAM OPERATIONS</p>
          <h1>
  {new Date().getHours() < 12
    ? "Good morning"
    : new Date().getHours() < 17
    ? "Good afternoon"
    : new Date().getHours() < 21
    ? "Good evening"
    : "Good night"}
  , Manjunath 👋
</h1>
          <p className="sub">
            Here’s your team's operational overview for today.
          </p>
        </div>

        <button className="primary" onClick={() => setPage("team")}>
          <Users size={18} />
          Manage team
        </button>
      </div>

      <div className="cards">
        <Metric
          icon={Users}
          label="Team members"
          value={data.team.length}
          note={`${presentToday} present today`}
        />
        <Metric
          icon={AlertTriangle}
          label="Absent today"
          value={absentToday}
          note="No work logged today"
        />
        <Metric
          icon={Calendar}
          label="On leave today"
          value={leaveToday}
          note="Marked On Leave in the sheet"
        />
        <Metric
          icon={FolderKanban}
          label="Active projects"
          value={activeProjects}
          note={`${data.projects.length} total projects`}
        />

        <Metric
          icon={Target}
          label="Productivity today"
          value={totals.done.toLocaleString()}
          note={`${completion}% of today's target`}
        />
        <Metric
          icon={ImageIcon}
          label="Total images worked"
          value={overallCompleted.toLocaleString()}
          note={`${overallCompletion}% of overall target (${overallTarget.toLocaleString()})`}
        />
        <Metric
          icon={ShieldCheck}
          label="QA accuracy"
          value={qaDenom ? `${qaAccuracy}%` : "—"}
          note={qaDenom ? `${totalErrors.toLocaleString()} total errors` : "No Accuracy Report imported yet"}
        />
        <Metric
          icon={AlertTriangle}
          label="Open issues"
          value={openIssues}
          note="Needs attention"
        />
      </div>

      <Panel title="Today's operational summary">
        <p className="op-summary">
          {latestDate ? (
            <>
              As of <b>{latestDate}</b>, <b>{presentToday}</b> of <b>{data.team.length}</b> team members are present
              {absentToday ? <> ({absentToday} absent, no work logged)</> : null}
              {leaveToday ? <>, {leaveToday} on leave</> : null}. The team completed{" "}
              <b>{totals.done.toLocaleString()}</b> images today ({completion}% of today's target), out of{" "}
              <b>{overallCompleted.toLocaleString()}</b> completed overall ({overallCompletion}% of the total target).{" "}
              {qaDenom ? <>QA accuracy stands at <b>{qaAccuracy}%</b> across {qaDenom.toLocaleString()} reviewed images. </> : null}
              There {openIssues === 1 ? "is" : "are"} <b>{openIssues}</b> open issue{openIssues === 1 ? "" : "s"} and{" "}
              <b>{notifications.length}</b> item{notifications.length === 1 ? "" : "s"} needing attention.
            </>
          ) : (
            "Import a Daily Effort Sheet to see today's operational summary."
          )}
        </p>
      </Panel>

      <div className="grid two">
        <Panel title="Team performance" action="View all" onAction={() => setPage("team")}>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Member</th>
                  <th>Progress</th>
                  <th>Completed</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {data.team.slice(0, 5).map(x => (
                  <tr key={x.id}>
                    <td>
                      <div className="person">
                        <div className="mini-avatar">
                          {x.name.slice(0, 1)}
                        </div>
                        <div>
                          <b>{x.name}</b>
                          <small>{x.role}</small>
                        </div>
                      </div>
                    </td>
                    <td>
                      <Progress
                        value={
                          x.target
                            ? Math.min(
                                100,
                                Math.round((x.completed / x.target) * 100)
                              )
                            : 0
                        }
                      />
                    </td>
                    <td>
                      <b>{Number(x.completed).toLocaleString()}</b> /{" "}
                      {Number(x.target).toLocaleString()}
                    </td>
                    <td>
                      <Status text={x.status} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Panel>

        <Panel
          title="Project progress"
          action="View projects"
          onAction={() => setPage("projects")}
        >
          <div className="project-list">
            {highlightedProjects.map(p => {
              const s = getProjectStats(p);

              return (
                <div className="project-row" key={p.id}>
                  <div className="project-icon">
                    <FolderKanban size={18} />
                  </div>

                  <div className="grow">
                    <div className="row-title">
                      <b>{p.name}</b>
                      <span>{s.total ? `${s.progress}%` : "No target"}</span>
                    </div>

                    <Progress value={s.progress} />

                    <small>
                      {s.total
                        ? `${s.remaining.toLocaleString()} images remaining`
                        : "Set a daily target"}
                    </small>
                  </div>
                </div>
              );
            })}
          </div>

          {data.projects.length > highlightedProjects.length && (
            <p className="muted" style={{ marginTop: 14 }}>
              Showing {highlightedProjects.length} of {data.projects.length} projects, most active first.{" "}
              <button className="link-btn" style={{ display: "inline" }} onClick={() => setPage("projects")}>
                View all
              </button>
            </p>
          )}
        </Panel>
      </div>

      <div className="grid two">
        <Panel title="Productivity trend" action="View reports" onAction={() => setPage("reports")}>
          {!trend.length ? (
            <p className="muted">No Daily Effort records imported yet.</p>
          ) : (
            <div className="bars">
              {trend.map(([date, value]) => (
                <div className="bar-row" key={date}>
                  <span>{date.slice(5)}</span>
                  <div>
                    <i style={{ width: `${(value / trendMax) * 100}%` }} />
                  </div>
                  <b>{value.toLocaleString()}</b>
                </div>
              ))}
            </div>
          )}
        </Panel>

        <Panel title="Top performers today" action="View team" onAction={() => setPage("team")}>
          {!topPerformers.length ? (
            <p className="muted">No completed work logged today yet.</p>
          ) : (
            <div className="bars">
              {topPerformers.map(x => (
                <div className="bar-row" key={x.id}>
                  <span>{x.name}</span>
                  <div>
                    <i style={{ width: `${(Number(x.completed) / topMax) * 100}%` }} />
                  </div>
                  <b>{Number(x.completed).toLocaleString()}</b>
                </div>
              ))}
            </div>
          )}
        </Panel>
      </div>

      <div className="grid three">
        <Panel title="Needing attention" action={notifications.length ? "View all" : undefined} onAction={() => setPage("projects")}>
          {!notifications.length && !absentNames.length ? (
            <p className="muted">Nothing needs attention right now.</p>
          ) : (
            <ul className="attention-list">
              {absentNames.slice(0, 3).map(name => (
                <li key={`absent-${name}`}>
                  <span className="attention-dot attention-absent" />
                  <span>{name} — no work logged today</span>
                </li>
              ))}
              {notifications.slice(0, 5).map(n => (
                <li key={n.id}>
                  <span className={`attention-dot attention-${n.kind}`} />
                  <span>{n.title}</span>
                </li>
              ))}
            </ul>
          )}
        </Panel>

        <Panel title="Today's activity">
          <div className="activity">
            <Activity />
            <div>
              <b>{totals.done.toLocaleString()}</b>
              <span>images completed</span>
            </div>
          </div>

          <div className="activity">
            <CheckCircle2 />
            <div>
              <b>{totals.reviewed.toLocaleString()}</b>
              <span>reviews completed</span>
            </div>
          </div>

          <div className="activity">
            <AlertTriangle />
            <div>
              <b>{openIssues}</b>
              <span>issues open</span>
            </div>
          </div>
        </Panel>

        <Panel title="Quick actions">
          <button className="quick" onClick={() => setPage("team")}>
            <Users />
            Update team progress
            <ChevronRight />
          </button>

          <button className="quick" onClick={() => setPage("issues")}>
            <AlertTriangle />
            Review open issues
            <ChevronRight />
          </button>

          <button className="quick" onClick={() => setPage("settings")}>
            <Download />
            Backup dashboard
            <ChevronRight />
          </button>
        </Panel>
      </div>
    </div>
  );
}

/* =========================================================
   COMMON UI
========================================================= */
function Metric({ icon: Icon, label, value, note, trend }) {
  return (
    <div className="metric">
      <div className="metric-top">
        <div className="metric-icon">
          <Icon size={20} />
        </div>
        <span className="trend">{trend}</span>
      </div>
      <h2>{value}</h2>
      <b>{label}</b>
      <small>{note}</small>
    </div>
  );
}

function Progress({ value }) {
  return (
    <div className="progress">
      <span
        style={{
          width: `${Math.max(0, Math.min(100, Number(value) || 0))}%`
        }}
      />
    </div>
  );
}

function Status({ text }) {
  return (
    <span
      className={
        "status " +
        String(text || "")
          .toLowerCase()
          .replaceAll(" ", "-")
    }
    >
      <i />
      {text}
    </span>
  );
}

function Panel({ title, action, onAction, children }) {
  return (
    <div className="panel">
      <div className="panel-head">
        <h3>{title}</h3>
        {action && (
          <button className="link-btn" onClick={onAction}>
            {action}
            <ChevronRight size={15} />
          </button>
        )}
      </div>
      {children}
    </div>
  );
}

/* =========================================================
   TEAM
========================================================= */
function Team({ rows, data, openAdd, canManage, onEdit, onDelete, onToggleStatus, onView }) {
  const [statusFilter, setStatusFilter] = useState("all");
  const [roleFilter, setRoleFilter] = useState("all");
  const [projectFilter, setProjectFilter] = useState("all");
  const [sortBy, setSortBy] = useState("name");
  const projectRows = useMemo(() => {
    const records = Array.isArray(data.sheetRecords)
      ? data.sheetRecords
      : [];

    if (!records.length) {
      return rows.map(x => ({
        id: `member-${x.id}`,
        name: x.name,
        project: "—",
        role: x.role,
        target: Number(x.target) || 0,
        completed: Number(x.completed) || 0,
        reviewed: Number(x.reviewed) || 0,
        errors: Number(x.errors) || 0,
        status: x.status || "Active"
      }));
    }

    const dates = [
      ...new Set(records.map(x => x.date).filter(Boolean))
    ].sort();

    const latestDate = dates[dates.length - 1];

    const latestRecords = records.filter(
      x =>
        x.date === latestDate &&
        x.project &&
        !["Saturday", "Sunday", "On Leave"].includes(x.project)
    );

    const grouped = {};

    latestRecords.forEach(record => {
      const name = String(record.name || "").trim();
      const project = getConfiguredProjectName(record.project);

      if (!name || !project) return;

      const type = String(record.type || "").trim();
      const derivedRole = /review/i.test(type) ? "Reviewer" : "Annotator";
      const key = `${name}|||${project}|||${derivedRole}`;

      if (!grouped[key]) {
        // Status, Errors and a manually-set Role belong to the roster
        // (data.team), not to today's sheet rows — pull the real values
        // instead of assuming everyone is Active with 0 errors.
        const member = rows.find(m => m.name.toLowerCase() === name.toLowerCase());

        grouped[key] = {
          id: key,
          name,
          project,
          role: member?.role || derivedRole,
          target: Number(projectTargets[project]) || 0,
          completed: 0,
          reviewed: 0,
          errors: Number(member?.errors) || 0,
          status: member?.status || "Active"
        };
      }

      const worked = Number(record.worked) || 0;
      grouped[key].completed += worked;

      if (/review/i.test(type)) {
        grouped[key].reviewed += worked;
      }
    });

    const result = Object.values(grouped);

    rows.forEach(member => {
      const exists = result.some(
        x => x.name.toLowerCase() === member.name.toLowerCase()
      );

      if (!exists) {
        result.push({
          id: `member-${member.id}`,
          name: member.name,
          project: "—",
          role: member.role,
          target: Number(member.target) || 0,
          completed: Number(member.completed) || 0,
          reviewed: Number(member.reviewed) || 0,
          errors: Number(member.errors) || 0,
          status: member.status || "Active"
        });
      }
    });

    return result;
  }, [data.sheetRecords, rows]);

  const roleOptions = [...new Set(projectRows.map(x => x.role).filter(Boolean))].sort();
  const projectOptions = [...new Set(projectRows.map(x => x.project).filter(p => p && p !== "—"))].sort();

  const visibleRows = useMemo(() => {
    let list = projectRows;

    if (statusFilter !== "all") {
      list = list.filter(x => (x.status || "Active") === statusFilter);
    }
    if (roleFilter !== "all") {
      list = list.filter(x => x.role === roleFilter);
    }
    if (projectFilter !== "all") {
      list = list.filter(x => x.project === projectFilter);
    }

    const sorted = [...list];
    if (sortBy === "name") {
      sorted.sort((a, b) => a.name.localeCompare(b.name));
    } else if (sortBy === "completed") {
      sorted.sort((a, b) => (Number(b.completed) || 0) - (Number(a.completed) || 0));
    } else if (sortBy === "target") {
      sorted.sort((a, b) => (Number(b.target) || 0) - (Number(a.target) || 0));
    } else if (sortBy === "progress") {
      const pct = x => (Number(x.target) ? (Number(x.completed) / Number(x.target)) : 0);
      sorted.sort((a, b) => pct(b) - pct(a));
    }
    return sorted;
  }, [projectRows, statusFilter, roleFilter, projectFilter, sortBy]);

  const imported =
    Array.isArray(data.sheetRecords) &&
    data.sheetRecords.length > 0;

  const latestDate = imported
    ? [...new Set(data.sheetRecords.filter(isWorkRow).map(x => x.date).filter(Boolean))]
        .sort()
        .pop()
    : null;

  return (
    <Page
      title="Team members"
      subtitle={
        imported
          ? `Daily project-wise productivity from imported sheet${
              latestDate ? ` • ${latestDate}` : ""
            }.`
          : "Monitor individual productivity, targets and quality."
      }
      action={canManage ? "+ Add member" : undefined}
      onAction={openAdd}
    >
      <Panel title="Ranking">
        <div className="grid two">
          <div>
            <p className="muted settings-note">Productivity — by today's completed images</p>
            <ol className="rank-list">
              {[...data.team]
                .sort((a, b) => (Number(b.completed) || 0) - (Number(a.completed) || 0))
                .slice(0, 5)
                .map((m, i) => (
                  <li key={m.id}>
                    <span className="rank-num">{i + 1}</span>
                    <button className="name-link" onClick={() => onView(m.name)}>{m.name}</button>
                    <b>{Number(m.completed).toLocaleString()}</b>
                  </li>
                ))}
            </ol>
          </div>

          <div>
            <p className="muted settings-note">Quality — by accuracy across imported Accuracy Reports</p>
            {(() => {
              const byName = new Map();
              (data.accuracyRecords || []).forEach(r => {
                const key = r.name || "Unknown";
                if (!byName.has(key)) byName.set(key, { tp: 0, fp: 0, fn: 0 });
                const e = byName.get(key);
                e.tp += Number(r.tp) || 0;
                e.fp += Number(r.fp) || 0;
                e.fn += Number(r.fn) || 0;
              });
              const qualityRanked = [...byName.entries()]
                .map(([name, v]) => {
                  const denom = v.tp + v.fp + v.fn;
                  return { name, accuracy: denom ? Math.round((v.tp / denom) * 100) : 0, denom };
                })
                .filter(x => x.denom > 0)
                .sort((a, b) => b.accuracy - a.accuracy)
                .slice(0, 5);

              return qualityRanked.length ? (
                <ol className="rank-list">
                  {qualityRanked.map((x, i) => (
                    <li key={x.name}>
                      <span className="rank-num">{i + 1}</span>
                      <button className="name-link" onClick={() => onView(x.name)}>{x.name}</button>
                      <b>{x.accuracy}%</b>
                    </li>
                  ))}
                </ol>
              ) : (
                <p className="muted">No Accuracy Report imported yet.</p>
              );
            })()}
          </div>
        </div>
      </Panel>

      <Panel title={`${visibleRows.length} of ${projectRows.length} records`}>
        <div className="report-controls">
          <label>
            Status
            <select value={statusFilter} onChange={e => setStatusFilter(e.target.value)}>
              <option value="all">All</option>
              <option value="Active">Active</option>
              <option value="Away">Away</option>
              <option value="Inactive">Inactive</option>
            </select>
          </label>

          <label>
            Role
            <select value={roleFilter} onChange={e => setRoleFilter(e.target.value)}>
              <option value="all">All</option>
              {roleOptions.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </label>

          <label>
            Project
            <select value={projectFilter} onChange={e => setProjectFilter(e.target.value)}>
              <option value="all">All</option>
              {projectOptions.map(p => <option key={p} value={p}>{p}</option>)}
            </select>
          </label>

          <label>
            Sort by
            <select value={sortBy} onChange={e => setSortBy(e.target.value)}>
              <option value="name">Name</option>
              <option value="completed">Completed</option>
              <option value="target">Target</option>
              <option value="progress">Progress</option>
            </select>
          </label>
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Member</th>
                <th>Project</th>
                <th>Role</th>
                <th>Target</th>
                <th>Completed</th>
                <th>Progress</th>
                <th>Reviewed</th>
                <th>Errors</th>
                <th>Status</th>
                {canManage && <th>Actions</th>}
              </tr>
            </thead>

            <tbody>
              {visibleRows.map(x => {
                const target = Number(x.target) || 0;
                const completed = Number(x.completed) || 0;
                const progress = target
                  ? Math.min(100, Math.round((completed / target) * 100))
                  : 0;

                return (
                  <tr key={x.id}>
                    <td>
                      <div className="person">
                        <div className="mini-avatar">
                          {x.name.slice(0, 1).toUpperCase()}
                        </div>
                        <button className="name-link" onClick={() => onView(x.name)}>
                          <b>{x.name}</b>
                        </button>
                      </div>
                    </td>

                    <td><b>{x.project}</b></td>
                    <td>{x.role}</td>
                    <td>{target ? target.toLocaleString() : "—"}</td>
                    <td><b>{completed.toLocaleString()}</b></td>

                    <td>
                      <div style={{ minWidth: "100px" }}>
                        <Progress value={progress} />
                        <small style={{ display: "block", marginTop: "4px" }}>
                          {progress}%
                        </small>
                      </div>
                    </td>

                    <td>{Number(x.reviewed).toLocaleString()}</td>

                    <td>
                      <span className={x.errors > 15 ? "danger-text" : "good-text"}>
                        {x.errors}
                      </span>
                    </td>

                    <td>
                      <Status
                        text={
                          target && progress >= 100
                            ? "Completed"
                            : x.status
                        }
                      />
                    </td>

                    {canManage && (
                      <td>
                        <div className="row-actions">
                          <button
                            className="secondary compact-btn"
                            onClick={() => {
                              const member = data.team.find(m => String(m.name).toLowerCase() === String(x.name).toLowerCase());
                              onEdit(member || { id: x.id, name: x.name, role: x.role, target: target, completed, reviewed: Number(x.reviewed) || 0, errors: Number(x.errors) || 0, status: x.status || "Active" });
                            }}
                          >Edit</button>
                          <button
                            className="secondary compact-btn"
                            onClick={() => onToggleStatus(x.name)}
                            title={x.status === "Inactive" ? `Reactivate ${x.name}` : `Deactivate ${x.name}`}
                          >
                            {x.status === "Inactive" ? "Reactivate" : "Deactivate"}
                          </button>
                          <button className="delete" onClick={() => onDelete(x.name)} title={`Delete ${x.name}`}>
                            <Trash2 size={16} />
                          </button>
                        </div>
                      </td>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Panel>
    </Page>
  );
}

/* =========================================================
   ATTENDANCE
   Derived from the imported sheet where possible:
   - Week Off: the calendar date is a Saturday/Sunday
   - Holiday: a date added to the manual holiday list
   - Leave: a sheet row whose Project cell reads "On Leave" / "Leave"
     (per the team's convention, absences are also logged this way)
   - Present: a real work row exists for that person/date
   - No data: nothing on file for that person/date
========================================================= */

function listDatesInRange(range) {
  const dates = [];
  let cur = dateKeyToLocalDate(range.start);
  const end = dateKeyToLocalDate(range.end);
  if (!cur || !end) return dates;

  while (cur.getTime() <= end.getTime()) {
    dates.push(toISODate(cur.getFullYear(), cur.getMonth() + 1, cur.getDate()));
    cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 1);
  }
  return dates;
}

function isWeekendDate(dateISO) {
  const d = dateKeyToLocalDate(dateISO);
  if (!d) return false;
  const day = d.getDay();
  return day === 0 || day === 6;
}

function getAttendanceStatus(dateISO, rowsForCell, holidaySet, override) {
  // A manual correction always wins over anything auto-derived.
  if (override) return override;

  if (holidaySet.has(dateISO)) return "Holiday";
  if (isWeekendDate(dateISO)) return "Week Off";

  if (!rowsForCell || !rowsForCell.length) return "No data";

  const onLeave = rowsForCell.some(r => ["On Leave", "Leave"].includes(r.project));
  if (onLeave) return "Leave";

  const worked = rowsForCell.some(isWorkRow);
  if (worked) return "Present";

  // Rows exist but only carry a weekend-style placeholder on a date the
  // calendar doesn't consider a weekend (rare) — treat conservatively as leave.
  return "Leave";
}

const ATTENDANCE_STATUSES = ["Present", "Absent", "Leave", "Half Day", "Week Off", "Holiday"];

function buildAttendanceMatrix(data, range) {
  const dates = listDatesInRange(range);

  const names = (Array.isArray(data.team) && data.team.length)
    ? data.team.map(t => t.name)
    : [...new Set((Array.isArray(data.sheetRecords) ? data.sheetRecords : []).map(r => r.name))];

  const byKey = new Map();
  (Array.isArray(data.sheetRecords) ? data.sheetRecords : []).forEach(r => {
    const key = `${r.name}__${r.date}`;
    if (!byKey.has(key)) byKey.set(key, []);
    byKey.get(key).push(r);
  });

  const holidaySet = new Set(
    (Array.isArray(data.holidays) ? data.holidays : []).map(h => h.date)
  );

  const overrideMap = new Map(
    (Array.isArray(data.attendanceOverrides) ? data.attendanceOverrides : [])
      .map(o => [`${o.name}__${o.date}`, o.status])
  );

  const rows = names.map(name => ({
    name,
    cells: dates.map(date => ({
      date,
      status: getAttendanceStatus(
        date,
        byKey.get(`${name}__${date}`),
        holidaySet,
        overrideMap.get(`${name}__${date}`)
      )
    }))
  }));

  const summary = { Present: 0, Absent: 0, Leave: 0, "Half Day": 0, "Week Off": 0, Holiday: 0, "No data": 0 };
  rows.forEach(row => {
    row.cells.forEach(c => {
      summary[c.status] = (summary[c.status] || 0) + 1;
    });
  });

  // "Working day" slots are every cell that isn't a Week Off or Holiday —
  // the meaningful denominator for attendance percentages.
  const workingDaySlots = Object.entries(summary).reduce(
    (s, [status, count]) => (["Week Off", "Holiday"].includes(status) ? s : s + count),
    0
  );
  const presentPct = workingDaySlots
    ? Math.round(((summary.Present + summary["Half Day"] * 0.5) / workingDaySlots) * 100)
    : null;

  return { dates, rows, summary, workingDaySlots, presentPct };
}

function Attendance({ data, update, canManage, notify }) {
  const latestDate = getLatestImportedDate(data);
  const anchorDate = loadPrefs().reportRangeMode === "thisWeek" ? getTodayISO() : (latestDate || getTodayISO());
  const defaultRange = getWeekRange(anchorDate);

  const [filter, setFilter] = useState({
    mode: "range",
    start: defaultRange.start,
    end: defaultRange.end
  });
  const [holidayDate, setHolidayDate] = useState("");
  const [holidayName, setHolidayName] = useState("");

  const range = useMemo(() => getDateRange(filter), [filter]);
  const matrix = useMemo(() => buildAttendanceMatrix(data, range), [data, range]);

  // Daily trend — present % for each date in the range, across everyone.
  const trend = useMemo(() => {
    return matrix.dates.map(date => {
      let present = 0;
      let slots = 0;
      matrix.rows.forEach(row => {
        const cell = row.cells.find(c => c.date === date);
        if (!cell) return;
        if (["Week Off", "Holiday"].includes(cell.status)) return;
        slots += 1;
        if (cell.status === "Present") present += 1;
        if (cell.status === "Half Day") present += 0.5;
      });
      return { date, pct: slots ? Math.round((present / slots) * 100) : null };
    });
  }, [matrix]);

  function addHoliday(e) {
    e.preventDefault();
    if (!canManage) return notify && notify("Only Admin or Team Lead can manage holidays.");

    const d = normalizeDateValue(holidayDate);
    if (!d) return;

    const holidays = [
      ...(Array.isArray(data.holidays) ? data.holidays : []),
      { id: Date.now(), date: d, name: holidayName.trim() || "Holiday" }
    ];
    update({ ...data, holidays });
    setHolidayDate("");
    setHolidayName("");
  }

  function removeHoliday(id) {
    if (!canManage) return notify && notify("Only Admin or Team Lead can manage holidays.");
    update({
      ...data,
      holidays: (Array.isArray(data.holidays) ? data.holidays : []).filter(h => h.id !== id)
    });
  }

  function setAttendanceOverride(name, date, status) {
    if (!canManage) return notify && notify("Only Admin or Team Lead can mark attendance.");

    const existing = Array.isArray(data.attendanceOverrides) ? data.attendanceOverrides : [];
    const withoutThis = existing.filter(o => !(o.name === name && o.date === date));

    // Empty selection means "back to auto" — just drop the override.
    const next = status ? [...withoutThis, { id: `${name}__${date}`, name, date, status }] : withoutThis;

    update({ ...data, attendanceOverrides: next });
    logActivity(`Marked ${name}'s attendance on ${date} as ${status || "auto"}`);
  }

  const pct = (status) => matrix.workingDaySlots ? Math.round(((matrix.summary[status] || 0) / matrix.workingDaySlots) * 100) : 0;

  return (
    <Page
      title="Attendance"
      subtitle="Present, absent, on leave, half day, week off, and holidays."
    >
      <Panel title="Date range">
        <DateFilter value={filter} onChange={setFilter} data={data} />
      </Panel>

      <div className="cards">
        <Metric icon={CheckCircle2} label="Present" value={matrix.summary["Present"] || 0} note={matrix.workingDaySlots ? `${pct("Present")}% of working days` : undefined} />
        <Metric icon={AlertTriangle} label="Absent" value={matrix.summary["Absent"] || 0} note={matrix.workingDaySlots ? `${pct("Absent")}% of working days` : undefined} />
        <Metric icon={Users} label="On leave" value={matrix.summary["Leave"] || 0} note={matrix.workingDaySlots ? `${pct("Leave")}% of working days` : undefined} />
        <Metric icon={Clock} label="Half day" value={matrix.summary["Half Day"] || 0} />
        <Metric icon={Calendar} label="Week off" value={matrix.summary["Week Off"] || 0} />
        <Metric icon={Sun} label="Holiday" value={matrix.summary["Holiday"] || 0} />
        <Metric icon={FileText} label="Working days" value={matrix.workingDaySlots} note="Present + Absent + Leave + Half day slots" />
        <Metric icon={BarChart3} label="Attendance rate" value={matrix.presentPct != null ? `${matrix.presentPct}%` : "—"} note="Present + ½ Half day, of working days" />
      </div>

      {canManage && (
        <Panel title="Holidays">
          <form className="report-controls" onSubmit={addHoliday}>
            <label>
              Date
              <input type="date" value={holidayDate} onChange={e => setHolidayDate(e.target.value)} />
            </label>
            <label>
              Name
              <input
                type="text"
                placeholder="e.g. Independence Day"
                value={holidayName}
                onChange={e => setHolidayName(e.target.value)}
              />
            </label>
            <button className="secondary" type="submit">
              <Plus size={16} /> Add holiday
            </button>
          </form>

          {(data.holidays || []).length > 0 && (
            <ul className="holiday-list">
              {[...data.holidays]
                .sort((a, b) => (a.date < b.date ? -1 : 1))
                .map(h => (
                  <li key={h.id}>
                    <span><b>{h.date}</b> — {h.name}</span>
                    <button className="delete" onClick={() => removeHoliday(h.id)}>
                      <Trash2 size={14} />
                    </button>
                  </li>
                ))}
            </ul>
          )}
        </Panel>
      )}

      <Panel title="Attendance trend">
        {!trend.length ? (
          <p className="muted">Select a date range to see the trend.</p>
        ) : (
          <div className="bars">
            {trend.map(t => (
              <div className="bar-row" key={t.date}>
                <span>{t.date.slice(5)}</span>
                <div><i style={{ width: `${t.pct ?? 0}%` }} /></div>
                <b>{t.pct != null ? `${t.pct}%` : "—"}</b>
              </div>
            ))}
          </div>
        )}
      </Panel>

      <Panel title={`Attendance grid — ${matrix.rows.length} team member${matrix.rows.length === 1 ? "" : "s"}`}>
        {canManage && (
          <p className="muted settings-note">Click any cell to correct it manually. Choosing "Auto" removes the correction.</p>
        )}

        {!matrix.dates.length ? (
          <p className="muted">Select a date range to see attendance.</p>
        ) : !matrix.rows.length ? (
          <p className="muted">No team members or sheet records to show yet.</p>
        ) : (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Name</th>
                  {matrix.dates.map(d => <th key={d}>{d.slice(5)}</th>)}
                </tr>
              </thead>
              <tbody>
                {matrix.rows.map(row => (
                  <tr key={row.name}>
                    <td><b>{row.name}</b></td>
                    {row.cells.map(c =>
                      canManage ? (
                        <td key={c.date}>
                          <select
                            className={`status-select status-${c.status.toLowerCase().replace(/\s+/g, "-")}`}
                            value={c.status}
                            onChange={e => setAttendanceOverride(row.name, c.date, e.target.value === "Auto" ? "" : e.target.value)}
                          >
                            <option value="Auto">Auto ({c.status})</option>
                            {ATTENDANCE_STATUSES.map(s => (
                              <option key={s} value={s}>{s}</option>
                            ))}
                          </select>
                        </td>
                      ) : (
                        <td key={c.date}><Status text={c.status} /></td>
                      )
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Panel>
    </Page>
  );
}

/* =========================================================
   PROJECTS
========================================================= */
function Projects({ rows, data, remove, openAdd, openEdit, canManage, onArchive }) {
  const [showArchived, setShowArchived] = useState(false);
  const [viewingProject, setViewingProject] = useState(null);

  const visibleRows = rows.filter(p => (showArchived ? true : !p.archived));
  const archivedCount = rows.filter(p => p.archived).length;

  return (
    <Page
      title="Projects"
      subtitle="Track workload, daily targets and completion across projects."
      action={canManage ? "+ Add project" : undefined}
      onAction={openAdd}
    >
      {archivedCount > 0 && (
        <label className="report-toggle" style={{ marginBottom: 16, display: "inline-flex" }}>
          <input type="checkbox" checked={showArchived} onChange={e => setShowArchived(e.target.checked)} />
          Show archived ({archivedCount})
        </label>
      )}

      <div className="project-cards">
        {visibleRows.map(p => {
          const s = getProjectStats(p);

          return (
            <div className={`project-card ${p.archived ? "project-card-archived" : ""}`} key={p.id}>
              <div className="project-card-top">
                <div className="project-icon">
                  <FolderKanban />
                </div>

                {canManage && (
                  <div className="project-card-actions">
                    <button
                      className="icon-btn project-edit-btn"
                      type="button"
                      title="Edit project"
                      aria-label={`Edit ${p.name}`}
                      onClick={() => openEdit(p)}
                    >
                      <Pencil size={16} />
                    </button>

                    <button
                      className="icon-btn project-edit-btn"
                      type="button"
                      title={p.archived ? "Unarchive project" : "Archive project"}
                      aria-label={p.archived ? `Unarchive ${p.name}` : `Archive ${p.name}`}
                      onClick={() => onArchive(p.id)}
                    >
                      <Archive size={16} />
                    </button>

                    <button
                      className="delete"
                      type="button"
                      title="Delete project"
                      aria-label={`Delete ${p.name}`}
                      onClick={() => remove("projects", p.id)}
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                )}
              </div>

              <button className="name-link" onClick={() => setViewingProject(p)}>
                <h3>{p.name}</h3>
              </button>

              {/* Status and remaining images are calculated automatically from total images and completed values */}
              <Status text={p.archived ? "Archived" : s.status} />

              <div className="pc-stat">
                <span>Total images</span>
                <b>{s.total ? s.total.toLocaleString() : "—"}</b>
              </div>

              <div className="pc-stat">
                <span>Daily target</span>
                <b>{p.target ? Number(p.target).toLocaleString() : "—"}</b>
              </div>

              <div className="pc-stat">
                <span>Completed</span>
                <b>{s.completed.toLocaleString()}</b>
              </div>

              <div className="pc-stat">
                <span>Completion</span>
                <b>{s.total ? `${s.progress}%` : "—"}</b>
              </div>

              <Progress value={s.progress} />

              <div className="pc-foot">
                <span>
                  {s.total
                    ? `${s.remaining.toLocaleString()} remaining`
                    : "Set target"}
                </span>

                <span>Due {p.deadline || "—"}</span>
              </div>
            </div>
          );
        })}
      </div>

      {viewingProject && (
        <ProjectProfileModal
          project={viewingProject}
          data={data}
          onClose={() => setViewingProject(null)}
        />
      )}
    </Page>
  );
}

function ProjectProfileModal({ project, data, onClose }) {
  const s = getProjectStats(project);

  const myRecords = (data.sheetRecords || []).filter(
    r => getConfiguredProjectName(r.project) === project.name && isWorkRow(r)
  );

  // Daily production — last 7 work-days for this project.
  const byDate = new Map();
  myRecords.forEach(r => byDate.set(r.date, (byDate.get(r.date) || 0) + (Number(r.worked) || 0)));
  const daily = [...byDate.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1)).slice(-7);
  const dailyMax = Math.max(1, ...daily.map(([, v]) => v));

  // Weekly production — last 6 weeks.
  const byWeek = new Map();
  myRecords.forEach(r => {
    const w = getWeekRange(r.date);
    const key = w.start;
    byWeek.set(key, (byWeek.get(key) || 0) + (Number(r.worked) || 0));
  });
  const weekly = [...byWeek.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1)).slice(-6);
  const weeklyMax = Math.max(1, ...weekly.map(([, v]) => v));

  // Monthly production — last 6 months.
  const byMonth = new Map();
  myRecords.forEach(r => {
    const key = r.date.slice(0, 7);
    byMonth.set(key, (byMonth.get(key) || 0) + (Number(r.worked) || 0));
  });
  const monthly = [...byMonth.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1)).slice(-6);
  const monthlyMax = Math.max(1, ...monthly.map(([, v]) => v));

  // Completion forecast — simple linear projection from recent pace.
  // Calendar-day based, doesn't account for weekends/leave, so treat this
  // as a rough estimate, not a commitment.
  const recentDates = [...byDate.keys()].sort().slice(-14);
  const recentTotal = recentDates.reduce((sum, d) => sum + (byDate.get(d) || 0), 0);
  const avgPace = recentDates.length ? recentTotal / recentDates.length : 0;
  let forecastText = "Not enough recent data to forecast.";
  if (s.total && s.remaining <= 0) {
    forecastText = "Already complete.";
  } else if (avgPace > 0 && s.remaining > 0) {
    const daysNeeded = Math.ceil(s.remaining / avgPace);
    const lastDate = dateKeyToLocalDate(recentDates[recentDates.length - 1]) || new Date();
    const est = new Date(lastDate);
    est.setDate(est.getDate() + daysNeeded);
    const estISO = toISODate(est.getFullYear(), est.getMonth() + 1, est.getDate());
    forecastText = `At the recent pace of ~${Math.round(avgPace).toLocaleString()} images/day, completion is estimated around ${estISO} (${daysNeeded} day${daysNeeded === 1 ? "" : "s"} away).`;
  }

  return (
    <div className="modal-bg" onClick={onClose}>
      <div className="modal profile-modal" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <div className="person">
            <div className="project-icon"><FolderKanban size={18} /></div>
            <div>
              <b>{project.name}</b>
              <small>Due {project.deadline || "—"} • <Status text={project.archived ? "Archived" : s.status} /></small>
            </div>
          </div>

          <button className="icon-btn" type="button" onClick={onClose} aria-label="Close">
            <X />
          </button>
        </div>

        <div className="profile-stats">
          <div><span>Total</span><b>{s.total.toLocaleString()}</b><small>images</small></div>
          <div><span>Completed</span><b>{s.completed.toLocaleString()}</b><small>{s.progress}%</small></div>
          <div><span>Remaining</span><b>{s.remaining.toLocaleString()}</b><small>images</small></div>
          <div><span>Daily target</span><b>{project.target ? Number(project.target).toLocaleString() : "—"}</b><small>per day</small></div>
        </div>

        <div className="modal-section">
          <b>Assigned employees</b>
          <p className="muted">{(project.assignedEmployees || []).length ? project.assignedEmployees.join(", ") : "None assigned"}</p>
        </div>

        <div className="modal-section">
          <b>Assigned reviewers</b>
          <p className="muted">{(project.assignedReviewers || []).length ? project.assignedReviewers.join(", ") : "None assigned"}</p>
        </div>

        <div className="modal-section">
          <b>Daily production — last 7 work-days</b>
          {daily.length ? (
            <div className="bars">
              {daily.map(([date, value]) => (
                <div className="bar-row" key={date}>
                  <span>{date.slice(5)}</span>
                  <div><i style={{ width: `${(value / dailyMax) * 100}%` }} /></div>
                  <b>{value.toLocaleString()}</b>
                </div>
              ))}
            </div>
          ) : <p className="muted">No work logged yet.</p>}
        </div>

        <div className="modal-section">
          <b>Weekly production — last 6 weeks</b>
          {weekly.length ? (
            <div className="bars">
              {weekly.map(([week, value]) => (
                <div className="bar-row" key={week}>
                  <span>{week.slice(5)}</span>
                  <div><i style={{ width: `${(value / weeklyMax) * 100}%` }} /></div>
                  <b>{value.toLocaleString()}</b>
                </div>
              ))}
            </div>
          ) : <p className="muted">No work logged yet.</p>}
        </div>

        <div className="modal-section">
          <b>Monthly production — last 6 months</b>
          {monthly.length ? (
            <div className="bars">
              {monthly.map(([month, value]) => (
                <div className="bar-row" key={month}>
                  <span>{month}</span>
                  <div><i style={{ width: `${(value / monthlyMax) * 100}%` }} /></div>
                  <b>{value.toLocaleString()}</b>
                </div>
              ))}
            </div>
          ) : <p className="muted">No work logged yet.</p>}
        </div>

        <div className="modal-section">
          <b>Completion forecast</b>
          <p className="muted">{forecastText}</p>
        </div>
      </div>
    </div>
  );
}

/* =========================================================
   QA
========================================================= */
function QA({ data }) {
  const latestDate = getLatestAccuracyDate(data);
  const [period, setPeriod] = useState(() => ({
    mode: "single",
    start: latestDate,
    end: latestDate
  }));
  const range = getDateRange(period);

  const allAccuracyRows = Array.isArray(data.accuracyRecords) ? data.accuracyRecords : [];
  const accuracyRows = allAccuracyRows.filter(x => isDateInRange(x.date, range));

  const totalDaily = accuracyRows.reduce((s, x) => s + (Number(x.dailyCount) || 0), 0);
  const totalTP = accuracyRows.reduce((s, x) => s + (Number(x.tp) || 0), 0);
  const totalFP = accuracyRows.reduce((s, x) => s + (Number(x.fp) || 0), 0);
  const totalFN = accuracyRows.reduce((s, x) => s + (Number(x.fn) || 0), 0);
  const totalErrors = totalFP + totalFN;
  const denominator = totalTP + totalFP + totalFN;

  const calculatedAccuracy = denominator > 0
    ? (totalTP / denominator) * 100
    : accuracyRows.length
      ? accuracyRows.reduce((s, x) => s + (Number(x.accuracy) || 0), 0) / accuracyRows.length
      : 0;

  const averageScore = accuracyRows.length
    ? accuracyRows.reduce((s, x) => s + (Number(x.score) || 0), 0) / accuracyRows.length
    : 0;

  // QA trend — accuracy % over time, computed from the FULL history
  // (independent of the period selector above), grouped 3 ways.
  function accuracyByBucket(keyFn, limit) {
    const buckets = new Map();
    allAccuracyRows.forEach(x => {
      if (!x.date) return;
      const key = keyFn(x.date);
      if (!buckets.has(key)) buckets.set(key, { tp: 0, fp: 0, fn: 0 });
      const b = buckets.get(key);
      b.tp += Number(x.tp) || 0;
      b.fp += Number(x.fp) || 0;
      b.fn += Number(x.fn) || 0;
    });
    return [...buckets.entries()]
      .sort((a, b) => (a[0] < b[0] ? -1 : 1))
      .slice(-limit)
      .map(([key, b]) => {
        const denom = b.tp + b.fp + b.fn;
        return { key, pct: denom ? Math.round((b.tp / denom) * 100) : null };
      });
  }

  const dailyTrend = accuracyByBucket(d => d, 7);
  const weeklyTrend = accuracyByBucket(d => getWeekRange(d).start, 6);
  const monthlyTrend = accuracyByBucket(d => d.slice(0, 7), 6);

  // Employee quality ranking with trend context, for this page specifically
  // (Team page has a simpler top-5 version of this same idea).
  const byName = new Map();
  allAccuracyRows.forEach(x => {
    const key = x.name || "Unknown";
    if (!byName.has(key)) byName.set(key, { tp: 0, fp: 0, fn: 0, count: 0 });
    const e = byName.get(key);
    e.tp += Number(x.tp) || 0;
    e.fp += Number(x.fp) || 0;
    e.fn += Number(x.fn) || 0;
    e.count += 1;
  });
  const employeeQuality = [...byName.entries()]
    .map(([name, v]) => {
      const denom = v.tp + v.fp + v.fn;
      return { name, accuracy: denom ? Math.round((v.tp / denom) * 100) : 0, denom, count: v.count };
    })
    .filter(x => x.denom > 0)
    .sort((a, b) => b.accuracy - a.accuracy);

  // Reviewer performance — only meaningful if the Accuracy Report sheet
  // actually has a Reviewer column. Most don't yet, so this is honest
  // about being empty rather than guessing who reviewed what.
  const hasReviewerData = allAccuracyRows.some(x => x.reviewer);
  const byReviewer = new Map();
  if (hasReviewerData) {
    allAccuracyRows.forEach(x => {
      if (!x.reviewer) return;
      const key = x.reviewer;
      if (!byReviewer.has(key)) byReviewer.set(key, { reviewed: 0, errorsFound: 0, days: new Set() });
      const r = byReviewer.get(key);
      r.reviewed += Number(x.dailyCount) || 0;
      r.errorsFound += (Number(x.fp) || 0) + (Number(x.fn) || 0);
      if (x.date) r.days.add(x.date);
    });
  }
  const reviewerRows = [...byReviewer.entries()]
    .map(([name, v]) => ({
      name,
      reviewed: v.reviewed,
      errorsFound: v.errorsFound,
      days: v.days.size,
      perDay: v.days.size ? Math.round(v.reviewed / v.days.size) : 0
    }))
    .sort((a, b) => b.reviewed - a.reviewed);

  return (
    <Page
      title="QA & Reviews"
      subtitle={
        range.start && range.end
          ? `${range.label}: ${accuracyRows.length} Accuracy Report records${data.accuracyFile ? ` • ${data.accuracyFile}` : ""}.`
          : "Select a specific date or date range to view the Accuracy Report."
      }
    >
      <DateFilter value={period} onChange={setPeriod} data={data} />

      <div className="cards">
        <Metric icon={ClipboardCheck} label="Total reviewed" value={totalDaily.toLocaleString()} note="Daily Count in selected period" trend={accuracyRows.length ? "Live" : "Pending"} />
        <Metric icon={CheckCircle2} label="Accuracy" value={`${calculatedAccuracy.toFixed(1)}%`} note={accuracyRows.length ? "From Accuracy Report" : "No report in selected period"} trend={accuracyRows.length ? "Good" : "Pending"} />
        <Metric icon={Target} label="Daily count" value={totalDaily.toLocaleString()} note="Images reported" trend={accuracyRows.length ? "Imported" : "Pending"} />
        <Metric icon={AlertTriangle} label="Total errors" value={totalErrors.toLocaleString()} note="FP + FN in selected period" trend={totalErrors ? "Monitor" : "Clear"} />
      </div>

      <Panel title={`Review readiness — ${range.label}`}>
        <div className="qa-grid">
          <div><h2>{calculatedAccuracy.toFixed(1)}%</h2><p className="muted">Overall accuracy</p><Progress value={calculatedAccuracy} /></div>
          <div><h2>{totalTP.toLocaleString()}</h2><p className="muted">True positives (TP)</p></div>
          <div><h2>{totalFP.toLocaleString()}</h2><p className="muted">False positives (FP)</p></div>
          <div><h2>{totalFN.toLocaleString()}</h2><p className="muted">False negatives (FN)</p></div>
          <div><h2>{averageScore.toFixed(1)}</h2><p className="muted">Average score</p></div>
          <div><h2>{data.issues.filter(x => x.status === "Open").length}</h2><p className="muted">Open QA issues</p></div>
        </div>
      </Panel>

      <div className="grid two">
        <Panel title="Error breakdown — all-time">
          {totalTP + totalFP + totalFN === 0 && !allAccuracyRows.length ? (
            <p className="muted">No Accuracy Report imported yet.</p>
          ) : (
            (() => {
              const allFP = allAccuracyRows.reduce((s, x) => s + (Number(x.fp) || 0), 0);
              const allFN = allAccuracyRows.reduce((s, x) => s + (Number(x.fn) || 0), 0);
              const maxErr = Math.max(1, allFP, allFN);
              return (
                <div className="bars">
                  <div className="bar-row">
                    <span>False positives</span>
                    <div><i style={{ width: `${(allFP / maxErr) * 100}%` }} /></div>
                    <b>{allFP.toLocaleString()}</b>
                  </div>
                  <div className="bar-row">
                    <span>False negatives</span>
                    <div><i style={{ width: `${(allFN / maxErr) * 100}%` }} /></div>
                    <b>{allFN.toLocaleString()}</b>
                  </div>
                </div>
              );
            })()
          )}
          <p className="muted settings-note">Your Accuracy Report doesn't include a finer error-category breakdown beyond False Positive / False Negative — add one to your sheet if you want more detail here.</p>
        </Panel>

        <Panel title="Employee quality ranking — all-time">
          {!employeeQuality.length ? (
            <p className="muted">No Accuracy Report imported yet.</p>
          ) : (
            <ol className="rank-list">
              {employeeQuality.slice(0, 8).map((x, i) => (
                <li key={x.name}>
                  <span className="rank-num">{i + 1}</span>
                  <span>{x.name}</span>
                  <b>{x.accuracy}%</b>
                </li>
              ))}
            </ol>
          )}
        </Panel>
      </div>

      <div className="grid three">
        <Panel title="Daily QA trend">
          {!dailyTrend.length ? <p className="muted">No data yet.</p> : (
            <div className="bars">
              {dailyTrend.map(t => (
                <div className="bar-row" key={t.key}>
                  <span>{t.key.slice(5)}</span>
                  <div><i style={{ width: `${t.pct ?? 0}%` }} /></div>
                  <b>{t.pct != null ? `${t.pct}%` : "—"}</b>
                </div>
              ))}
            </div>
          )}
        </Panel>

        <Panel title="Weekly QA trend">
          {!weeklyTrend.length ? <p className="muted">No data yet.</p> : (
            <div className="bars">
              {weeklyTrend.map(t => (
                <div className="bar-row" key={t.key}>
                  <span>{t.key.slice(5)}</span>
                  <div><i style={{ width: `${t.pct ?? 0}%` }} /></div>
                  <b>{t.pct != null ? `${t.pct}%` : "—"}</b>
                </div>
              ))}
            </div>
          )}
        </Panel>

        <Panel title="Monthly QA trend">
          {!monthlyTrend.length ? <p className="muted">No data yet.</p> : (
            <div className="bars">
              {monthlyTrend.map(t => (
                <div className="bar-row" key={t.key}>
                  <span>{t.key}</span>
                  <div><i style={{ width: `${t.pct ?? 0}%` }} /></div>
                  <b>{t.pct != null ? `${t.pct}%` : "—"}</b>
                </div>
              ))}
            </div>
          )}
        </Panel>
      </div>

      <Panel title="Reviewer performance">
        {!hasReviewerData ? (
          <p className="muted">
            Your Accuracy Report sheet doesn't have a "Reviewer" column, so there's no way to know who performed each review — this section can't show real data yet.
            Add a Reviewer column to your sheet (any of these header names work: Reviewer, Reviewed By, QA Reviewer) and re-import to populate this automatically.
          </p>
        ) : (
          <div className="table-wrap">
            <table>
              <thead>
                <tr><th>Reviewer</th><th>Images reviewed</th><th>Errors found</th><th>Active days</th><th>Avg / day</th></tr>
              </thead>
              <tbody>
                {reviewerRows.map(r => (
                  <tr key={r.name}>
                    <td><b>{r.name}</b></td>
                    <td>{r.reviewed.toLocaleString()}</td>
                    <td>{r.errorsFound.toLocaleString()}</td>
                    <td>{r.days}</td>
                    <td>{r.perDay.toLocaleString()}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Panel>

      <Panel title={`Accuracy Report — ${accuracyRows.length} records`}>
        {!accuracyRows.length ? (
          <p className="muted">No Accuracy Report data exists for <b>{range.label}</b>.</p>
        ) : (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Date</th><th>Name</th><th>Daily Count</th><th>TP</th><th>FP</th><th>FN</th>
                  <th>Accuracy</th><th>Score</th><th>Detailed Report</th><th>Images Used</th><th>Comment</th>
                </tr>
              </thead>
              <tbody>
                {accuracyRows.map((x, i) => (
                  <tr key={x.id || i}>
                    <td>{x.date || "—"}</td>
                    <td><b>{x.name}</b></td>
                    <td>{Number(x.dailyCount || 0).toLocaleString()}</td>
                    <td>{Number(x.tp || 0).toLocaleString()}</td>
                    <td>{Number(x.fp || 0).toLocaleString()}</td>
                    <td>{Number(x.fn || 0).toLocaleString()}</td>
                    <td><b>{Number(x.accuracy || 0).toFixed(1)}%</b></td>
                    <td><b>{Number(x.score || 0).toFixed(1)}</b></td>
                    <td>
                      {x.link ? (
                        <a className="link-btn" href={x.link} target="_blank" rel="noreferrer">
                          Open <ExternalLink size={14} />
                        </a>
                      ) : "—"}
                    </td>
                    <td>{x.imagesUsed || "—"}</td>
                    <td>{x.comment || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Panel>
    </Page>
  );
}

/* =========================================================
   ISSUES
========================================================= */
function Issues({ rows, data, remove, openAdd, canManage, onUpdate, onComment }) {
  const [statusFilter, setStatusFilter] = useState("all");
  const [projectFilter, setProjectFilter] = useState("all");
  const [ownerFilter, setOwnerFilter] = useState("all");
  const [severityFilter, setSeverityFilter] = useState("all");
  const [dateFilter, setDateFilter] = useState({ mode: "range", start: "", end: "" });
  const [viewingIssue, setViewingIssue] = useState(null);

  const projectOptions = [...new Set(rows.map(x => x.project).filter(Boolean))].sort();
  const ownerOptions = [...new Set(rows.map(x => x.owner).filter(Boolean))].sort();
  const range = getDateRange(dateFilter);

  const filteredRows = rows.filter(x => {
    if (statusFilter !== "all" && x.status !== statusFilter) return false;
    if (projectFilter !== "all" && x.project !== projectFilter) return false;
    if (ownerFilter !== "all" && x.owner !== ownerFilter) return false;
    if (severityFilter !== "all" && x.severity !== severityFilter) return false;
    if (range.start && range.end && !isDateInRange(x.date, range)) return false;
    return true;
  });

  const today = getTodayISO();
  const openCount = rows.filter(x => x.status === "Open").length;
  const inProgressCount = rows.filter(x => x.status === "In Progress").length;
  const resolvedCount = rows.filter(x => ["Resolved", "Closed"].includes(x.status)).length;
  const overdueCount = rows.filter(
    x => x.dueDate && x.dueDate < today && !["Resolved", "Closed"].includes(x.status)
  ).length;

  const resolvedWithDates = rows.filter(x => x.resolvedDate && x.date);
  const avgResolutionDays = resolvedWithDates.length
    ? Math.round(
        resolvedWithDates.reduce((s, x) => s + (daysBetweenDates(x.date, x.resolvedDate) || 0), 0) /
          resolvedWithDates.length
      )
    : null;

  // Issue trend — created per day, last 14 days that have any issue.
  const byDate = new Map();
  rows.forEach(x => { if (x.date) byDate.set(x.date, (byDate.get(x.date) || 0) + 1); });
  const trend = [...byDate.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1)).slice(-14);
  const trendMax = Math.max(1, ...trend.map(([, v]) => v));

  // Project-wise issue counts.
  const byProject = new Map();
  rows.forEach(x => { if (x.project) byProject.set(x.project, (byProject.get(x.project) || 0) + 1); });
  const projectCounts = [...byProject.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8);
  const projectMax = Math.max(1, ...projectCounts.map(([, v]) => v));

  return (
    <Page
      title="Issues tracker"
      subtitle="Capture annotation problems and close them before submission."
      action={canManage ? "+ Log issue" : undefined}
      onAction={openAdd}
    >
      <div className="cards">
        <Metric icon={AlertTriangle} label="Open" value={openCount} />
        <Metric icon={Clock} label="In progress" value={inProgressCount} />
        <Metric icon={CheckCircle2} label="Resolved / closed" value={resolvedCount} />
        <Metric icon={AlertTriangle} label="Overdue" value={overdueCount} note="Past due date, still open" />
      </div>

      <div className="grid two">
        <Panel title="Issue trend — last 14 days with activity">
          {!trend.length ? <p className="muted">No issues logged yet.</p> : (
            <div className="bars">
              {trend.map(([date, count]) => (
                <div className="bar-row" key={date}>
                  <span>{date.slice(5)}</span>
                  <div><i style={{ width: `${(count / trendMax) * 100}%` }} /></div>
                  <b>{count}</b>
                </div>
              ))}
            </div>
          )}
        </Panel>

        <Panel title="Project-wise issues">
          {!projectCounts.length ? <p className="muted">No issues logged yet.</p> : (
            <div className="bars">
              {projectCounts.map(([project, count]) => (
                <div className="bar-row" key={project}>
                  <span>{project}</span>
                  <div><i style={{ width: `${(count / projectMax) * 100}%` }} /></div>
                  <b>{count}</b>
                </div>
              ))}
            </div>
          )}
          {avgResolutionDays != null && (
            <p className="muted settings-note">Average resolution time: <b>{avgResolutionDays} day{avgResolutionDays === 1 ? "" : "s"}</b>, based on {resolvedWithDates.length} resolved issue{resolvedWithDates.length === 1 ? "" : "s"}.</p>
          )}
        </Panel>
      </div>

      <Panel title={`${filteredRows.length} of ${rows.length} issues`}>
        <div className="report-controls">
          <label>
            Status
            <select value={statusFilter} onChange={e => setStatusFilter(e.target.value)}>
              <option value="all">All</option>
              <option value="Open">Open</option>
              <option value="In Progress">In Progress</option>
              <option value="Resolved">Resolved</option>
              <option value="Closed">Closed</option>
            </select>
          </label>
          <label>
            Project
            <select value={projectFilter} onChange={e => setProjectFilter(e.target.value)}>
              <option value="all">All</option>
              {projectOptions.map(p => <option key={p} value={p}>{p}</option>)}
            </select>
          </label>
          <label>
            Employee
            <select value={ownerFilter} onChange={e => setOwnerFilter(e.target.value)}>
              <option value="all">All</option>
              {ownerOptions.map(o => <option key={o} value={o}>{o}</option>)}
            </select>
          </label>
          <label>
            Priority
            <select value={severityFilter} onChange={e => setSeverityFilter(e.target.value)}>
              <option value="all">All</option>
              <option value="High">High</option>
              <option value="Medium">Medium</option>
              <option value="Low">Low</option>
            </select>
          </label>
          <DateFilter value={dateFilter} onChange={setDateFilter} data={data} />
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Issue</th>
                <th>Project</th>
                <th>Employee</th>
                <th>Assigned to</th>
                <th>Severity</th>
                <th>Status</th>
                <th>Due</th>
                <th>Date</th>
                <th></th>
              </tr>
            </thead>

            <tbody>
              {filteredRows.map(x => {
                const overdue = x.dueDate && x.dueDate < today && !["Resolved", "Closed"].includes(x.status);
                return (
                  <tr key={x.id}>
                    <td>
                      <button className="name-link" onClick={() => setViewingIssue(x)}>
                        <b>{x.type}</b>
                      </button>
                    </td>
                    <td>{x.project}</td>
                    <td>{x.owner}</td>
                    <td>{x.assignedTo || "—"}</td>
                    <td>
                      <span className={"severity " + x.severity.toLowerCase()}>
                        {x.severity}
                      </span>
                    </td>
                    <td><Status text={x.status} /></td>
                    <td className={overdue ? "danger-text" : undefined}>{x.dueDate || "—"}{overdue ? " (overdue)" : ""}</td>
                    <td>{x.date}</td>
                    <td>
                      {canManage && (
                        <button
                          className="delete"
                          onClick={() => remove("issues", x.id)}
                        >
                          <Trash2 size={16} />
                        </button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Panel>

      {viewingIssue && (
        <IssueDetailModal
          issue={rows.find(r => r.id === viewingIssue.id) || viewingIssue}
          canManage={canManage}
          onUpdate={onUpdate}
          onComment={onComment}
          onClose={() => setViewingIssue(null)}
        />
      )}
    </Page>
  );
}

function IssueDetailModal({ issue, canManage, onUpdate, onComment, onClose }) {
  const [resolution, setResolution] = useState(issue.resolution || "");
  const [commentText, setCommentText] = useState("");

  return (
    <div className="modal-bg" onClick={onClose}>
      <div className="modal profile-modal" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <div>
            <p className="eyebrow">ISSUE</p>
            <h2>{issue.type}</h2>
          </div>
          <button className="icon-btn" type="button" onClick={onClose} aria-label="Close">
            <X />
          </button>
        </div>

        <div className="modal-section">
          <b>Description</b>
          <p className="muted">{issue.description || "No description added."}</p>
        </div>

        <div className="modal-section">
          <div className="report-controls">
            <label>
              Project
              <input value={issue.project || ""} disabled />
            </label>
            <label>
              Employee
              <input value={issue.owner || ""} disabled />
            </label>
            <label>
              Severity
              <input value={issue.severity || ""} disabled />
            </label>
          </div>
        </div>

        {canManage ? (
          <div className="modal-section">
            <div className="report-controls">
              <label>
                Status
                <select value={issue.status} onChange={e => onUpdate(issue.id, { status: e.target.value })}>
                  <option value="Open">Open</option>
                  <option value="In Progress">In Progress</option>
                  <option value="Resolved">Resolved</option>
                  <option value="Closed">Closed</option>
                </select>
              </label>
              <label>
                Due date
                <input type="date" value={issue.dueDate || ""} onChange={e => onUpdate(issue.id, { dueDate: e.target.value })} />
              </label>
            </div>
          </div>
        ) : (
          <div className="modal-section">
            <b>Status</b>
            <Status text={issue.status} />
          </div>
        )}

        <div className="modal-section">
          <b>Resolution</b>
          {canManage ? (
            <>
              <textarea
                className="issue-resolution"
                rows={3}
                value={resolution}
                onChange={e => setResolution(e.target.value)}
                placeholder="What fixed this?"
              />
              <button className="secondary compact-btn" onClick={() => onUpdate(issue.id, { resolution })}>Save resolution</button>
            </>
          ) : (
            <p className="muted">{issue.resolution || "Not resolved yet."}</p>
          )}
        </div>

        <div className="modal-section">
          <b>Comments</b>
          {(issue.comments || []).length ? (
            <ul className="comment-list">
              {issue.comments.map(c => (
                <li key={c.id}>
                  <p>{c.text}</p>
                  <small>{new Date(c.at).toLocaleString()}</small>
                </li>
              ))}
            </ul>
          ) : (
            <p className="muted">No comments yet.</p>
          )}

          {canManage && (
            <div className="comment-form">
              <input
                type="text"
                placeholder="Add a comment…"
                value={commentText}
                onChange={e => setCommentText(e.target.value)}
                onKeyDown={e => {
                  if (e.key === "Enter" && commentText.trim()) {
                    onComment(issue.id, commentText);
                    setCommentText("");
                  }
                }}
              />
              <button
                className="secondary compact-btn"
                onClick={() => {
                  if (!commentText.trim()) return;
                  onComment(issue.id, commentText);
                  setCommentText("");
                }}
              >
                Post
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* =========================================================
   ANALYTICS
========================================================= */
function Analytics({ data }) {
  const latestDate = getLatestImportedDate(data);

  const [period, setPeriod] = useState(() => ({
    mode: "single",
    start: latestDate,
    end: latestDate
  }));

  const range = getDateRange(period);

  /*
    IMPORTANT:
    Always normalize imported Daily Effort dates before
    comparing them with the selected dashboard date.
    This prevents timezone/date-format shifts.
  */
  const normalizedStart = normalizeDateValue(range.start);
  const normalizedEnd = normalizeDateValue(range.end);

  const normalizedRange = {
    ...range,
    start: normalizedStart,
    end: normalizedEnd
  };

  const sheetRows = Array.isArray(data.sheetRecords)
    ? data.sheetRecords.filter(x => {
        const recordDate = normalizeDateValue(x.date);

        if (!recordDate || !normalizedRange.start || !normalizedRange.end) {
          return false;
        }

        return (
          recordDate >= normalizedRange.start &&
          recordDate <= normalizedRange.end
        );
      })
    : [];

  const accuracyRows = Array.isArray(data.accuracyRecords)
    ? data.accuracyRecords.filter(x =>
        isDateInRange(x.date, normalizedRange)
      )
    : [];

  const memberMap = {};

  sheetRows.forEach(x => {
    const name = String(x.name || "").trim();
    if (!name) return;

    if (!memberMap[name]) {
      memberMap[name] = {
        name,
        completed: 0,
        reviewed: 0
      };
    }

    const worked = Number(x.worked) || 0;

    memberMap[name].completed += worked;

    if (/review/i.test(String(x.type || ""))) {
      memberMap[name].reviewed += worked;
    }
  });

  const productivityRows = Object.values(memberMap).sort(
    (a, b) => b.completed - a.completed
  );

  const max = Math.max(
    ...productivityRows.map(x => x.completed),
    1
  );

  const maxDaily = Math.max(
    ...accuracyRows.map(x => Number(x.dailyCount) || 0),
    1
  );

  const averageAccuracy = accuracyRows.length
    ? accuracyRows.reduce(
        (s, x) => s + (Number(x.accuracy) || 0),
        0
      ) / accuracyRows.length
    : 0;

  const averageScore = accuracyRows.length
    ? accuracyRows.reduce(
        (s, x) => s + (Number(x.score) || 0),
        0
      ) / accuracyRows.length
    : 0;

  const totalWorked = productivityRows.reduce(
    (s, x) => s + x.completed,
    0
  );

  const totalTarget = data.team.reduce(
    (s, x) => s + (Number(x.target) || 0),
    0
  );

  // Productivity trend — daily/weekly/monthly, from the FULL history
  // (independent of the period selector above).
  const allWorkRecords = (Array.isArray(data.sheetRecords) ? data.sheetRecords : []).filter(isWorkRow);

  function sumByBucket(keyFn, limit) {
    const buckets = new Map();
    allWorkRecords.forEach(r => {
      const key = keyFn(r.date);
      buckets.set(key, (buckets.get(key) || 0) + (Number(r.worked) || 0));
    });
    return [...buckets.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1)).slice(-limit);
  }

  const dailyProductivity = sumByBucket(d => d, 14);
  const weeklyProductivity = sumByBucket(d => getWeekRange(d).start, 8);
  const monthlyProductivity = sumByBucket(d => d.slice(0, 7), 6);

  // Project productivity comparison, for the selected period.
  const byProject = new Map();
  sheetRows.filter(isWorkRow).forEach(r => {
    const project = getConfiguredProjectName(r.project);
    byProject.set(project, (byProject.get(project) || 0) + (Number(r.worked) || 0));
  });
  const projectComparison = [...byProject.entries()].sort((a, b) => b[1] - a[1]);
  const projectMax = Math.max(1, ...projectComparison.map(([, v]) => v));

  // Target vs actual, per employee, for the selected period.
  const targetVsActual = data.team
    .map(t => ({ name: t.name, target: Number(t.target) || 0, actual: memberMap[t.name]?.completed || 0 }))
    .filter(x => x.target || x.actual)
    .sort((a, b) => b.actual - a.actual);

  // Attendance vs productivity, for the selected period (only meaningful
  // for range mode with a real span — a single day makes attendance % trivial).
  const attendanceForRange = buildAttendanceMatrix(data, normalizedRange);
  const attendanceVsProductivity = attendanceForRange.rows
    .map(row => {
      const workingSlots = row.cells.filter(c => !["Week Off", "Holiday"].includes(c.status)).length;
      const presentSlots = row.cells.filter(c => c.status === "Present").length + row.cells.filter(c => c.status === "Half Day").length * 0.5;
      const attendancePct = workingSlots ? Math.round((presentSlots / workingSlots) * 100) : null;
      return { name: row.name, attendancePct, productivity: memberMap[row.name]?.completed || 0 };
    })
    .filter(x => x.attendancePct != null);

  // QA trend — condensed daily/weekly view, from the full accuracy history.
  const allAccuracyAll = Array.isArray(data.accuracyRecords) ? data.accuracyRecords : [];
  function accuracyByBucket(keyFn, limit) {
    const buckets = new Map();
    allAccuracyAll.forEach(x => {
      if (!x.date) return;
      const key = keyFn(x.date);
      if (!buckets.has(key)) buckets.set(key, { tp: 0, fp: 0, fn: 0 });
      const b = buckets.get(key);
      b.tp += Number(x.tp) || 0;
      b.fp += Number(x.fp) || 0;
      b.fn += Number(x.fn) || 0;
    });
    return [...buckets.entries()]
      .sort((a, b) => (a[0] < b[0] ? -1 : 1))
      .slice(-limit)
      .map(([key, b]) => {
        const denom = b.tp + b.fp + b.fn;
        return { key, pct: denom ? Math.round((b.tp / denom) * 100) : null };
      });
  }
  const qaDailyTrend = accuracyByBucket(d => d, 10);
  const qaWeeklyTrend = accuracyByBucket(d => getWeekRange(d).start, 6);

  return (
    <Page
      title="Analytics"
      subtitle={`Productivity and accuracy analytics for ${range.label.toLowerCase()}.`}
    >
      <DateFilter
        value={period}
        onChange={setPeriod}
        data={data}
      />

      <Panel title={`Completed images by team member — ${range.label}`}>
        {!productivityRows.length ? (
          <p className="muted">
            No Daily Effort records exist for <b>{range.label}</b>.
          </p>
        ) : (
          <div className="bars">
            {productivityRows.map(x => (
              <div className="bar-row" key={x.name}>
                <span>{x.name}</span>

                <div>
                  <i
                    style={{
                      width: `${(x.completed / max) * 100}%`
                    }}
                  />
                </div>

                <b>{x.completed.toLocaleString()}</b>
              </div>
            ))}
          </div>
        )}
      </Panel>

      <div className="grid three">
        <Panel title="Capacity">
          <div className="big-number">
            {totalTarget
              ? Math.round((totalWorked / totalTarget) * 100)
              : 0}
            %
          </div>

          <p className="muted">
            Selected-period work vs current team targets
          </p>
        </Panel>

        <Panel title="Accuracy health">
          <div className="big-number">
            {accuracyRows.length
              ? `${averageAccuracy.toFixed(1)}%`
              : "—"}
          </div>

          <p className="muted">
            Average accuracy from Accuracy Report
          </p>

          <Progress value={averageAccuracy} />
        </Panel>

        <Panel title="Average score">
          <div className="big-number">
            {accuracyRows.length
              ? averageScore.toFixed(1)
              : "—"}
          </div>

          <p className="muted">
            Score from Accuracy Report
          </p>
        </Panel>
      </div>

      <Panel title={`Accuracy by team member — ${range.label}`}>
        {!accuracyRows.length ? (
          <p className="muted">
            No Accuracy Report data exists for <b>{range.label}</b>.
          </p>
        ) : (
          <div className="bars">
            {accuracyRows.map((x, i) => {
              const accuracy = Math.max(
                0,
                Math.min(100, Number(x.accuracy) || 0)
              );

              return (
                <div className="bar-row" key={x.id || i}>
                  <span>{x.name}</span>

                  <div>
                    <i
                      style={{
                        width: `${accuracy}%`
                      }}
                    />
                  </div>

                  <b>{accuracy.toFixed(1)}%</b>
                </div>
              );
            })}
          </div>
        )}
      </Panel>

      <Panel title={`Daily count from Accuracy Report — ${range.label}`}>
        {!accuracyRows.length ? (
          <p className="muted">
            No Accuracy Report data available.
          </p>
        ) : (
          <div className="bars">
            {accuracyRows.map((x, i) => (
              <div className="bar-row" key={x.id || i}>
                <span>{x.name}</span>

                <div>
                  <i
                    style={{
                      width: `${
                        ((Number(x.dailyCount) || 0) / maxDaily) * 100
                      }%`
                    }}
                  />
                </div>

                <b>
                  {Number(x.dailyCount || 0).toLocaleString()}
                </b>
              </div>
            ))}
          </div>
        )}
      </Panel>

      <div className="grid three">
        <Panel title="Daily productivity — last 14 days">
          {!dailyProductivity.length ? <p className="muted">No data yet.</p> : (
            <div className="bars">
              {dailyProductivity.map(([date, v]) => (
                <div className="bar-row" key={date}>
                  <span>{date.slice(5)}</span>
                  <div><i style={{ width: `${(v / Math.max(1, ...dailyProductivity.map(([, x]) => x))) * 100}%` }} /></div>
                  <b>{v.toLocaleString()}</b>
                </div>
              ))}
            </div>
          )}
        </Panel>

        <Panel title="Weekly productivity — last 8 weeks">
          {!weeklyProductivity.length ? <p className="muted">No data yet.</p> : (
            <div className="bars">
              {weeklyProductivity.map(([week, v]) => (
                <div className="bar-row" key={week}>
                  <span>{week.slice(5)}</span>
                  <div><i style={{ width: `${(v / Math.max(1, ...weeklyProductivity.map(([, x]) => x))) * 100}%` }} /></div>
                  <b>{v.toLocaleString()}</b>
                </div>
              ))}
            </div>
          )}
        </Panel>

        <Panel title="Monthly productivity — last 6 months">
          {!monthlyProductivity.length ? <p className="muted">No data yet.</p> : (
            <div className="bars">
              {monthlyProductivity.map(([month, v]) => (
                <div className="bar-row" key={month}>
                  <span>{month}</span>
                  <div><i style={{ width: `${(v / Math.max(1, ...monthlyProductivity.map(([, x]) => x))) * 100}%` }} /></div>
                  <b>{v.toLocaleString()}</b>
                </div>
              ))}
            </div>
          )}
        </Panel>
      </div>

      <div className="grid two">
        <Panel title={`Project comparison — ${range.label}`}>
          {!projectComparison.length ? (
            <p className="muted">No Daily Effort records exist for <b>{range.label}</b>.</p>
          ) : (
            <div className="bars">
              {projectComparison.map(([project, v]) => (
                <div className="bar-row" key={project}>
                  <span>{project}</span>
                  <div><i style={{ width: `${(v / projectMax) * 100}%` }} /></div>
                  <b>{v.toLocaleString()}</b>
                </div>
              ))}
            </div>
          )}
        </Panel>

        <Panel title={`Target vs actual — ${range.label}`}>
          {!targetVsActual.length ? (
            <p className="muted">No team targets or work logged for <b>{range.label}</b>.</p>
          ) : (
            <div className="table-wrap">
              <table>
                <thead><tr><th>Name</th><th>Target</th><th>Actual</th><th>Achievement</th></tr></thead>
                <tbody>
                  {targetVsActual.map(x => (
                    <tr key={x.name}>
                      <td><b>{x.name}</b></td>
                      <td>{x.target.toLocaleString()}</td>
                      <td>{x.actual.toLocaleString()}</td>
                      <td>{x.target ? `${Math.round((x.actual / x.target) * 100)}%` : "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </Panel>
      </div>

      <div className="grid two">
        <Panel title={`Attendance vs productivity — ${range.label}`}>
          {!attendanceVsProductivity.length ? (
            <p className="muted">Not enough attendance data for this period.</p>
          ) : (
            <div className="table-wrap">
              <table>
                <thead><tr><th>Name</th><th>Attendance</th><th>Images worked</th></tr></thead>
                <tbody>
                  {attendanceVsProductivity.map(x => (
                    <tr key={x.name}>
                      <td><b>{x.name}</b></td>
                      <td>{x.attendancePct}%</td>
                      <td>{x.productivity.toLocaleString()}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </Panel>

        <Panel title="QA accuracy trend">
          {!qaDailyTrend.length ? (
            <p className="muted">No Accuracy Report imported yet.</p>
          ) : (
            <>
              <p className="muted settings-note">Daily (last 10)</p>
              <div className="bars">
                {qaDailyTrend.map(t => (
                  <div className="bar-row" key={t.key}>
                    <span>{t.key.slice(5)}</span>
                    <div><i style={{ width: `${t.pct ?? 0}%` }} /></div>
                    <b>{t.pct != null ? `${t.pct}%` : "—"}</b>
                  </div>
                ))}
              </div>
              <p className="muted settings-note" style={{ marginTop: 14 }}>
                Project-wise and reviewer-wise QA trends aren't available — your Accuracy Report doesn't currently include a project or reviewer column. See QA & Reviews for the weekly/monthly breakdown.
              </p>
            </>
          )}
        </Panel>
      </div>
    </Page>
  );
}

/* =========================================================
   REPORTS
   Report types: Daily, Weekly, Monthly, Custom range,
   Project-wise, Employee-wise, QA report, Productivity report.
   Export: Excel (.xlsx), CSV, Print/PDF (browser print).
========================================================= */

const WORK_TYPE_LABELS = {
  daily: "Daily Report",
  weekly: "Weekly Report",
  monthly: "Monthly Report",
  custom: "Custom Date-range Report",
  project: "Project-wise Report",
  employee: "Employee-wise Report",
  qa: "QA Report",
  productivity: "Productivity Report",
  attendance: "Attendance Report",
  issues: "Issues Report"
};

// Monday-Sunday week containing the given YYYY-MM-DD date.
function getWeekRange(dateStr) {
  const d = dateKeyToLocalDate(dateStr);
  if (!d) return { start: "", end: "" };

  const day = d.getDay(); // 0 = Sunday
  const diffToMonday = day === 0 ? -6 : 1 - day;

  const monday = new Date(d);
  monday.setDate(d.getDate() + diffToMonday);

  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);

  return {
    start: toISODate(monday.getFullYear(), monday.getMonth() + 1, monday.getDate()),
    end: toISODate(sunday.getFullYear(), sunday.getMonth() + 1, sunday.getDate())
  };
}

// Full calendar month for an "YYYY-MM" value (from <input type="month">).
function getMonthRange(monthStr) {
  const m = /^(\d{4})-(\d{2})$/.exec(String(monthStr || ""));
  if (!m) return { start: "", end: "" };

  const year = Number(m[1]);
  const month = Number(m[2]);
  const lastDay = new Date(year, month, 0).getDate();

  return {
    start: toISODate(year, month, 1),
    end: toISODate(year, month, lastDay)
  };
}

function downloadBlob(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function exportRowsToExcel(filename, sheetName, rows) {
  if (!rows.length) return;
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31) || "Report");
  XLSX.writeFile(wb, filename);
}

function exportRowsToCSV(filename, rows) {
  if (!rows.length) return;
  const ws = XLSX.utils.json_to_sheet(rows);
  const csv = XLSX.utils.sheet_to_csv(ws);
  downloadBlob(filename, csv, "text/csv;charset=utf-8;");
}

// Excludes weekend/leave placeholder rows so "work" totals aren't skewed.
function isWorkRow(row) {
  return !["Saturday", "Sunday", "On Leave", "Leave", "Holiday"].includes(row.project);
}

function buildTimeReport(data, range) {
  const rows = (Array.isArray(data.sheetRecords) ? data.sheetRecords : [])
    .filter(x => isDateInRange(x.date, range) && isWorkRow(x));

  const byName = new Map();
  rows.forEach(x => {
    const key = x.name || "Unknown";
    if (!byName.has(key)) {
      byName.set(key, { name: key, imagesWorked: 0, workingDays: new Set(), projects: new Set() });
    }
    const entry = byName.get(key);
    entry.imagesWorked += Number(x.worked) || 0;
    if (x.date) entry.workingDays.add(x.date);
    if (x.project) entry.projects.add(x.project);
  });

  const summary = [...byName.values()]
    .map(x => ({
      Name: x.name,
      "Images Worked": x.imagesWorked,
      "Working Days": x.workingDays.size,
      "Avg / Day": x.workingDays.size ? Math.round(x.imagesWorked / x.workingDays.size) : 0,
      Projects: [...x.projects].join(", ")
    }))
    .sort((a, b) => b["Images Worked"] - a["Images Worked"]);

  const detail = rows
    .map(x => ({
      Date: x.date,
      Name: x.name,
      Project: x.project,
      Type: x.type,
      "Images Worked": Number(x.worked) || 0,
      Link: x.link || ""
    }))
    .sort((a, b) => (a.Date < b.Date ? -1 : a.Date > b.Date ? 1 : a.Name < b.Name ? -1 : 1));

  return { summary, detail };
}

function buildProjectReport(data, range, projectFilter) {
  const rows = (Array.isArray(data.sheetRecords) ? data.sheetRecords : [])
    .filter(x => isDateInRange(x.date, range) && isWorkRow(x))
    .filter(x => !projectFilter || x.project === projectFilter);

  const byProject = new Map();
  rows.forEach(x => {
    const key = x.project || "Unknown";
    if (!byProject.has(key)) {
      byProject.set(key, { project: key, imagesWorked: 0, people: new Set(), days: new Set() });
    }
    const entry = byProject.get(key);
    entry.imagesWorked += Number(x.worked) || 0;
    if (x.name) entry.people.add(x.name);
    if (x.date) entry.days.add(x.date);
  });

  const projectMeta = new Map(
    (Array.isArray(data.projects) ? data.projects : []).map(p => [p.name, p])
  );

  const summary = [...byProject.values()]
    .map(x => {
      const meta = projectMeta.get(x.project);
      const stats = meta ? getProjectStats(meta) : null;
      return {
        Project: x.project,
        "Images Worked (in range)": x.imagesWorked,
        "Team Members": [...x.people].join(", "),
        "Active Days": x.days.size,
        Target: stats ? stats.total : "",
        "Completed (overall)": stats ? stats.completed : "",
        "Remaining (overall)": stats ? stats.remaining : "",
        "Progress %": stats ? `${stats.progress}%` : "",
        Status: meta ? meta.status : "",
        Deadline: meta ? (meta.deadline || "") : ""
      };
    })
    .sort((a, b) => b["Images Worked (in range)"] - a["Images Worked (in range)"]);

  const detail = rows
    .map(x => ({
      Date: x.date,
      Project: x.project,
      Name: x.name,
      Type: x.type,
      "Images Worked": Number(x.worked) || 0,
      Link: x.link || ""
    }))
    .sort((a, b) => (a.Date < b.Date ? -1 : a.Date > b.Date ? 1 : 1));

  return { summary, detail };
}

function buildEmployeeReport(data, range, employeeFilter) {
  const rows = (Array.isArray(data.sheetRecords) ? data.sheetRecords : [])
    .filter(x => isDateInRange(x.date, range))
    .filter(x => !employeeFilter || x.name === employeeFilter);

  const workRows = rows.filter(isWorkRow);

  const detail = workRows
    .map(x => ({
      Date: x.date,
      Name: x.name,
      Project: x.project,
      Type: x.type,
      "Images Worked": Number(x.worked) || 0,
      Link: x.link || ""
    }))
    .sort((a, b) => (a.Date < b.Date ? -1 : a.Date > b.Date ? 1 : 1));

  const totalImages = workRows.reduce((s, x) => s + (Number(x.worked) || 0), 0);
  const workingDays = new Set(workRows.map(x => x.date)).size;
  const leaveDays = rows.filter(x => ["On Leave", "Leave"].includes(x.project)).length;
  const projects = [...new Set(workRows.map(x => x.project).filter(Boolean))];

  const summary = employeeFilter
    ? [{
        Name: employeeFilter,
        "Total Images Worked": totalImages,
        "Working Days": workingDays,
        "Avg / Day": workingDays ? Math.round(totalImages / workingDays) : 0,
        "Leave Days": leaveDays,
        Projects: projects.join(", ")
      }]
    : buildTimeReport(data, range).summary;

  return { summary, detail };
}

function buildQAReportRows(data, range) {
  const rows = (Array.isArray(data.accuracyRecords) ? data.accuracyRecords : [])
    .filter(x => isDateInRange(x.date, range));

  const byName = new Map();
  rows.forEach(x => {
    const key = x.name || "Unknown";
    if (!byName.has(key)) {
      byName.set(key, { name: key, dailyCount: 0, tp: 0, fp: 0, fn: 0, scoreSum: 0, scoreCount: 0 });
    }
    const entry = byName.get(key);
    entry.dailyCount += Number(x.dailyCount) || 0;
    entry.tp += Number(x.tp) || 0;
    entry.fp += Number(x.fp) || 0;
    entry.fn += Number(x.fn) || 0;
    entry.scoreSum += Number(x.score) || 0;
    entry.scoreCount += 1;
  });

  const summary = [...byName.values()]
    .map(x => {
      const denom = x.tp + x.fp + x.fn;
      const accuracy = denom > 0 ? (x.tp / denom) * 100 : 0;
      return {
        Name: x.name,
        "Daily Count": x.dailyCount,
        TP: x.tp,
        FP: x.fp,
        FN: x.fn,
        "Accuracy %": Number(accuracy.toFixed(1)),
        "Avg Score": x.scoreCount ? Number((x.scoreSum / x.scoreCount).toFixed(1)) : 0
      };
    })
    .sort((a, b) => b["Accuracy %"] - a["Accuracy %"]);

  const detail = rows
    .map(x => ({
      Date: x.date,
      Name: x.name,
      "Daily Count": x.dailyCount,
      TP: x.tp,
      FP: x.fp,
      FN: x.fn,
      "Accuracy %": Number((x.accuracy || 0).toFixed(1)),
      Score: x.score,
      "Images Used": x.imagesUsed,
      Comment: x.comment
    }))
    .sort((a, b) => (a.Date < b.Date ? -1 : a.Date > b.Date ? 1 : 1));

  return { summary, detail };
}

function buildProductivityReport(data, range) {
  const timeReport = buildTimeReport(data, range);
  const targetByName = new Map(
    (Array.isArray(data.team) ? data.team : []).map(t => [t.name, Number(t.target) || 0])
  );

  const summary = timeReport.summary
    .map(x => {
      const target = targetByName.get(x.Name) || 0;
      const achievement = target ? Math.min(999, Math.round((x["Images Worked"] / target) * 100)) : "";
      return {
        Name: x.Name,
        "Images Worked": x["Images Worked"],
        Target: target || "",
        "Achievement %": target ? `${achievement}%` : "",
        "Working Days": x["Working Days"],
        "Avg / Day": x["Avg / Day"],
        Projects: x.Projects
      };
    })
    .sort((a, b) => b["Images Worked"] - a["Images Worked"]);

  return { summary, detail: timeReport.detail };
}

function buildAttendanceReport(data, range) {
  const matrix = buildAttendanceMatrix(data, range);

  const summary = matrix.rows
    .map(row => {
      const counts = { Present: 0, Absent: 0, Leave: 0, "Half Day": 0, "Week Off": 0, Holiday: 0 };
      row.cells.forEach(c => { counts[c.status] = (counts[c.status] || 0) + 1; });
      const workingSlots = row.cells.length - counts["Week Off"] - counts.Holiday;
      const presentPct = workingSlots
        ? Math.round(((counts.Present + counts["Half Day"] * 0.5) / workingSlots) * 100)
        : 0;
      return {
        Name: row.name,
        Present: counts.Present,
        Absent: counts.Absent,
        Leave: counts.Leave,
        "Half Day": counts["Half Day"],
        "Week Off": counts["Week Off"],
        Holiday: counts.Holiday,
        "Attendance %": presentPct
      };
    })
    .sort((a, b) => b["Attendance %"] - a["Attendance %"]);

  const detail = [];
  matrix.rows.forEach(row => {
    row.cells.forEach(c => {
      detail.push({ Date: c.date, Name: row.name, Status: c.status });
    });
  });
  detail.sort((a, b) => (a.Date < b.Date ? -1 : a.Date > b.Date ? 1 : a.Name < b.Name ? -1 : 1));

  return { summary, detail };
}

function buildIssuesReport(data, range) {
  const rows = (Array.isArray(data.issues) ? data.issues : [])
    .filter(x => isDateInRange(x.date, range));

  const byProject = new Map();
  rows.forEach(x => {
    const key = x.project || "Unknown";
    if (!byProject.has(key)) byProject.set(key, { project: key, open: 0, inProgress: 0, resolved: 0, closed: 0, total: 0 });
    const entry = byProject.get(key);
    entry.total += 1;
    if (x.status === "Open") entry.open += 1;
    else if (x.status === "In Progress") entry.inProgress += 1;
    else if (x.status === "Resolved") entry.resolved += 1;
    else if (x.status === "Closed") entry.closed += 1;
  });

  const summary = [...byProject.values()]
    .map(x => ({
      Project: x.project,
      Open: x.open,
      "In Progress": x.inProgress,
      Resolved: x.resolved,
      Closed: x.closed,
      Total: x.total
    }))
    .sort((a, b) => b.Total - a.Total);

  const detail = rows
    .map(x => ({
      Date: x.date,
      Issue: x.type,
      Project: x.project,
      Employee: x.owner,
      "Assigned To": x.assignedTo || "",
      Priority: x.severity,
      Status: x.status,
      "Due Date": x.dueDate || "",
      "Resolved Date": x.resolvedDate || "",
      Resolution: x.resolution || ""
    }))
    .sort((a, b) => (a.Date < b.Date ? -1 : a.Date > b.Date ? 1 : 1));

  return { summary, detail };
}

function Reports({ data }) {
  const latestDate = getLatestImportedDate(data);
  const anchorDate = loadPrefs().reportRangeMode === "thisWeek" ? getTodayISO() : latestDate;

  const [reportType, setReportType] = useState("daily");
  const [singleDate, setSingleDate] = useState(anchorDate);
  const [weekDate, setWeekDate] = useState(anchorDate);
  const [monthValue, setMonthValue] = useState(anchorDate ? anchorDate.slice(0, 7) : "");
  const [customFilter, setCustomFilter] = useState(() => ({
    mode: "range",
    start: anchorDate,
    end: anchorDate
  }));
  const [projectFilter, setProjectFilter] = useState("");
  const [employeeFilter, setEmployeeFilter] = useState("");
  const [showDetail, setShowDetail] = useState(false);

  const range = useMemo(() => {
    if (reportType === "daily") {
      const d = normalizeDateValue(singleDate);
      return { start: d, end: d, label: d || "Select a date" };
    }
    if (reportType === "weekly") {
      const w = getWeekRange(normalizeDateValue(weekDate));
      return { ...w, label: w.start ? `${w.start} → ${w.end}` : "Select a date" };
    }
    if (reportType === "monthly") {
      const m = getMonthRange(monthValue);
      return { ...m, label: m.start ? `${m.start} → ${m.end}` : "Select a month" };
    }
    // custom, project, employee, qa, productivity all use the custom range picker
    return getDateRange(customFilter);
  }, [reportType, singleDate, weekDate, monthValue, customFilter]);

  const report = useMemo(() => {
    if (!range.start || !range.end) return { summary: [], detail: [] };
    switch (reportType) {
      case "project":
        return buildProjectReport(data, range, projectFilter);
      case "employee":
        return buildEmployeeReport(data, range, employeeFilter);
      case "qa":
        return buildQAReportRows(data, range);
      case "productivity":
        return buildProductivityReport(data, range);
      case "attendance":
        return buildAttendanceReport(data, range);
      case "issues":
        return buildIssuesReport(data, range);
      default:
        return buildTimeReport(data, range);
    }
  }, [data, range, reportType, projectFilter, employeeFilter]);

  const rowsForExport = showDetail && report.detail.length ? report.detail : report.summary;
  const reportLabel = WORK_TYPE_LABELS[reportType] || "Report";
  const fileBase = `AnnotatePro_${reportType}_${range.start || "report"}_${range.end || ""}`.replace(/[^a-zA-Z0-9_-]/g, "");

  function handlePrint() {
    window.print();
  }

  return (
    <Page
      title="Reports"
      subtitle="Build a report, preview it, then export to Excel, CSV, or print/PDF."
    >
      <Panel title="Report builder">
        <div className="report-controls">
          <label>
            Report type
            <select value={reportType} onChange={e => setReportType(e.target.value)}>
              <option value="daily">Daily report</option>
              <option value="weekly">Weekly report</option>
              <option value="monthly">Monthly report</option>
              <option value="custom">Custom date-range report</option>
              <option value="project">Project-wise report</option>
              <option value="employee">Employee-wise report</option>
              <option value="qa">QA report</option>
              <option value="productivity">Productivity report</option>
              <option value="attendance">Attendance report</option>
              <option value="issues">Issues report</option>
            </select>
          </label>

          {reportType === "daily" && (
            <label>
              Date
              <input type="date" value={singleDate || ""} onChange={e => setSingleDate(e.target.value)} />
            </label>
          )}

          {reportType === "weekly" && (
            <label>
              Any date in the week
              <input type="date" value={weekDate || ""} onChange={e => setWeekDate(e.target.value)} />
            </label>
          )}

          {reportType === "monthly" && (
            <label>
              Month
              <input type="month" value={monthValue || ""} onChange={e => setMonthValue(e.target.value)} />
            </label>
          )}

          {["custom", "project", "employee", "qa", "productivity", "attendance", "issues"].includes(reportType) && (
            <DateFilter value={customFilter} onChange={setCustomFilter} data={data} />
          )}

          {reportType === "project" && (
            <label>
              Project
              <select value={projectFilter} onChange={e => setProjectFilter(e.target.value)}>
                <option value="">All projects</option>
                {(data.projects || []).map(p => (
                  <option key={p.id} value={p.name}>{p.name}</option>
                ))}
              </select>
            </label>
          )}

          {reportType === "employee" && (
            <label>
              Employee
              <select value={employeeFilter} onChange={e => setEmployeeFilter(e.target.value)}>
                <option value="">All team members</option>
                {(data.team || []).map(t => (
                  <option key={t.id} value={t.name}>{t.name}</option>
                ))}
              </select>
            </label>
          )}
        </div>

        <div className="report-actions">
          <label className="report-toggle">
            <input type="checkbox" checked={showDetail} onChange={e => setShowDetail(e.target.checked)} />
            Export/print detailed rows instead of summary
          </label>

          <div className="report-export-btns">
            <button
              className="secondary"
              disabled={!rowsForExport.length}
              onClick={() => exportRowsToExcel(`${fileBase}.xlsx`, reportLabel, rowsForExport)}
            >
              <Download size={16} /> Excel
            </button>
            <button
              className="secondary"
              disabled={!rowsForExport.length}
              onClick={() => exportRowsToCSV(`${fileBase}.csv`, rowsForExport)}
            >
              <Download size={16} /> CSV
            </button>
            <button className="secondary" disabled={!rowsForExport.length} onClick={handlePrint}>
              <Printer size={16} /> Print / PDF
            </button>
          </div>
        </div>
      </Panel>

      <div className="report-printable">
        <div className="report-print-head">
          <h2>{reportLabel}</h2>
          <p className="muted">{range.label}</p>
        </div>

        <Panel title={`${reportLabel} — ${report.summary.length} row${report.summary.length === 1 ? "" : "s"}`}>
          {!range.start || !range.end ? (
            <p className="muted">Select a date{["custom", "project", "employee", "qa", "productivity", "attendance", "issues"].includes(reportType) ? " range" : ""} to build the report.</p>
          ) : !report.summary.length ? (
            <p className="muted">No data found for <b>{range.label}</b>.</p>
          ) : (
            <ReportTable rows={report.summary} />
          )}
        </Panel>

        {showDetail && report.detail.length > 0 && (
          <Panel title={`Detailed rows — ${report.detail.length}`}>
            <ReportTable rows={report.detail} />
          </Panel>
        )}
      </div>
    </Page>
  );
}

function ReportTable({ rows }) {
  if (!rows.length) return null;
  const columns = Object.keys(rows[0]);

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {columns.map(col => <th key={col}>{col}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i}>
              {columns.map(col => <td key={col}>{row[col] === "" || row[col] == null ? "—" : String(row[col])}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* =========================================================
   SHEET IMPORT
========================================================= */
function SheetImport({ data, update, notify }) {
  const [preview, setPreview] = useState([]);
  const [accuracyPreview, setAccuracyPreview] = useState([]);
  const [fileName, setFileName] = useState(data.sheetFile || "");
  const [accuracyFileName, setAccuracyFileName] = useState(data.accuracyFile || "");
  const [busy, setBusy] = useState(false);
  const [accuracyBusy, setAccuracyBusy] = useState(false);

  function normalizeHeader(value) {
    return String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/[%()]/g, "")
      .replace(/[_-]+/g, " ")
      .replace(/\s+/g, " ");
  }

  function findHeaderIndex(headers, candidates) {
    const normalized = headers.map(normalizeHeader);
    const wanted = candidates.map(normalizeHeader);
    const exact = normalized.findIndex(h => wanted.includes(h));
    if (exact >= 0) return exact;
    return normalized.findIndex(h => wanted.some(c => h.includes(c) || c.includes(h)));
  }

  function parseNumber(value) {
    if (value == null || value === "") return 0;
    const cleaned = String(value)
      .replace(/,/g, "")
      .replace(/%/g, "")
      .trim();
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  }

  function parseAccuracyValue(value) {
    if (value == null || value === "") return 0;
    const text = String(value).trim();
    const n = parseNumber(text);
    if (text.includes("%")) return n;
    if (n > 0 && n <= 1) return n * 100;
    return n;
  }

  function parseWorkbook(file) {
    setBusy(true);
    setFileName(file.name);

    const reader = new FileReader();

    reader.onload = e => {
      try {
        /*
          IMPORTANT: do NOT use cellDates: true here.
          SheetJS builds those Date objects using the
          browser's LOCAL clock, so reading them back
          (with either local or UTC getters) is ambiguous
          and shifts by a day depending on the user's
          timezone offset. Instead we keep raw Excel
          serial numbers and convert them ourselves with
          pure UTC-epoch math (see normalizeDateValue),
          which is timezone-independent.
        */
        const wb = XLSX.read(e.target.result, {
          type: "array"
        });

        // A project code in this sheet never contains a space (they're all
        // underscore_separated) and is never a long sentence. This filters
        // out stray notes typed into a project cell by mistake, without
        // needing to hardcode any specific note text.
        function looksLikeProjectName(text) {
          if (!text) return false;
          if (text.length > 60) return false;
          if (/\s/.test(text)) return false;
          return true;
        }

        const records = [];
        const names = [];

        // Import every tab in the workbook (e.g. one tab per month) and
        // combine them into one dataset, instead of only reading the first.
        wb.SheetNames.forEach(sheetName => {
          const ws = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, {
            header: 1,
            defval: null,
            raw: true
          });

          const dateStarts = [];
          const firstRow = rows[0] || [];

          for (let c = 1; c < firstRow.length; c++) {
            const rawValue = firstRow[c];
            let date = "";

            /*
              IMPORTANT:
              Spreadsheet date headers must be treated as calendar dates.
              Do not allow timezone conversion to move the date backward.
            */
            if (rawValue instanceof Date && !Number.isNaN(rawValue.getTime())) {
              // Same fix as normalizeDateValue(): XLSX gives us a Date
              // anchored to UTC midnight, so read it back with UTC getters.
              date = toISODate(
                rawValue.getUTCFullYear(),
                rawValue.getUTCMonth() + 1,
                rawValue.getUTCDate()
              );
            } else {
              date = normalizeDateValue(rawValue);
            }

            if (date) dateStarts.push({ c, date });
          }

          for (let r = 1; r < rows.length; r++) {
            const name = rows[r]?.[0];
            if (!name) continue;

            const cleanName = String(name).trim();
            if (!names.includes(cleanName)) names.push(cleanName);

            for (const d of dateStarts) {
              const project = rows[r]?.[d.c];
              const type = rows[r]?.[d.c + 1];
              const worked = rows[r]?.[d.c + 2];
              const link = rows[r]?.[d.c + 3];

              if (project == null && type == null && worked == null && link == null) continue;

              const ps = String(project ?? "").split("\n");
              const ts = String(type ?? "").split("\n");
              const nums = String(worked ?? "").split("\n");
              const ls = String(link ?? "").split("\n");

              ps.forEach((p, i) => {
                const text = p.trim();
                if (!text) return;

                const isPlaceholder = ["Saturday", "Sunday", "On Leave", "Leave", "Holiday"].includes(text);
                if (!isPlaceholder && !looksLikeProjectName(text)) return;

                const raw = nums[i] ?? nums[nums.length - 1] ?? "";
                const n = parseNumber(raw);

                records.push({
                  id: `${d.date}-${cleanName}-${records.length}`,
                  date: d.date,
                  name: cleanName,
                  project: text,
                  type: (ts[i] ?? ts[ts.length - 1] ?? "").trim(),
                  worked: n,
                  link: (ls[i] ?? ls[ls.length - 1] ?? "").trim()
                });
              });
            }
          }
        });

        // Only dates with real work count toward "today" — a pre-filled
        // Saturday/Sunday placeholder for a future date that hasn't
        // happened yet must never be picked as the latest date, or every
        // team member's daily completed/target would compute against an
        // empty day and show as 0.
        const workDates = [...new Set(records.filter(isWorkRow).map(x => x.date).filter(Boolean))].sort();
        const today = workDates[workDates.length - 1] || new Date().toISOString().slice(0, 10);
        const todayRows = records.filter(
          x => x.date === today && !["Saturday", "Sunday", "On Leave"].includes(x.project)
        );

        const team = names.map((name, i) => {
          const person = todayRows.filter(x => x.name === name);
          const projectNames = [...new Set(person.map(x => getConfiguredProjectName(x.project)).filter(Boolean))];
          const target = projectNames.reduce((sum, project) => sum + (projectTargets[project] || 0), 0);
          const completed = person.reduce((s, x) => s + (Number(x.worked) || 0), 0);
          const reviewed = person.filter(x => /review/i.test(x.type)).reduce((s, x) => s + (Number(x.worked) || 0), 0);
          const leave = records.some(x => x.date === today && x.name === name && x.project === "On Leave");

          // Role, Errors, a manually-set Inactive status, and the member's
          // id are yours to set — import must never silently reset them.
          // Only Target/Completed/Reviewed/Active-vs-Away are recalculated
          // from the sheet.
          const existing = (data.team || []).find(
            m => String(m.name).toLowerCase() === String(name).toLowerCase()
          );

          return {
            id: existing?.id ?? (1000 + i),
            name,
            role: existing?.role || "Annotator",
            target,
            completed,
            reviewed,
            errors: existing?.errors ?? 0,
            status: existing?.status === "Inactive" ? "Inactive" : (leave ? "Away" : "Active")
          };
        });

        // Existing team members this import's sheet has no row for at all
        // (e.g. someone deactivated and removed from the sheet) are kept
        // as-is instead of disappearing from the roster.
        const importedNames = new Set(names.map(n => n.toLowerCase()));
        const untouchedMembers = (data.team || []).filter(
          m => !importedNames.has(String(m.name).toLowerCase())
        );
        const fullTeam = [...team, ...untouchedMembers];

        // Completed is summed across the WHOLE imported history (all dates),
        // not just the latest day, so status reflects true overall progress.
        const workRecords = records.filter(
          x => !["Saturday", "Sunday", "On Leave"].includes(x.project)
        );

        // Only projects you've already created get their stats updated from
        // the sheet — nothing new is auto-created here. Daily Target, Total
        // Images and Deadline are never touched by import; only Completed
        // and Status get recalculated from the sheet data.
        const projectMap = {};
        (data.projects || []).forEach(p => {
          projectMap[String(p.name).toLowerCase()] = {
            name: p.name,
            id: p.id,
            completed: 0,
            target: Number(p.target) || 0,
            totalImages: Math.max(0, Number(p.totalImages ?? p.total) || 0),
            deadline: p.deadline || "",
            archived: !!p.archived,
            assignedEmployees: Array.isArray(p.assignedEmployees) ? p.assignedEmployees : [],
            assignedReviewers: Array.isArray(p.assignedReviewers) ? p.assignedReviewers : []
          };
        });

        const unmatchedProjectNames = new Set();

        workRecords.forEach(x => {
          const configuredName = getConfiguredProjectName(x.project);
          if (!configuredName) return;

          const key = configuredName.toLowerCase();
          if (!projectMap[key]) {
            unmatchedProjectNames.add(configuredName);
            return; // no matching project — skip, don't auto-create
          }
          projectMap[key].completed += Number(x.worked) || 0;
        });

        const projects = Object.values(projectMap).map(v => {
          const totalImages = Math.max(0, Number(v.totalImages) || 0);
          const completed = Math.max(0, Math.min(totalImages, Number(v.completed) || 0));
          const remaining = Math.max(0, totalImages - completed);
          return {
            id: v.id,
            name: v.name,
            target: v.target,
            totalImages,
            total: totalImages,
            completed,
            remaining,
            status: getProjectStatus(totalImages, completed, remaining),
            deadline: v.deadline,
            archived: v.archived,
            assignedEmployees: v.assignedEmployees,
            assignedReviewers: v.assignedReviewers
          };
        });

        update({
          ...data,
          team: fullTeam,
          projects,
          sheetRecords: records,
          sheetFile: file.name,
          sheetLastSync: new Date().toISOString()
        });

        setPreview(records.slice(0, 25));
        logActivity(`Imported sheet "${file.name}" (${records.length} work records)`);

        if (unmatchedProjectNames.size) {
          const names = [...unmatchedProjectNames].slice(0, 5).join(", ");
          const more = unmatchedProjectNames.size > 5 ? ` and ${unmatchedProjectNames.size - 5} more` : "";
          notify(
            `Imported ${records.length} work records. Not linked to a project yet: ${names}${more}. Add them on the Projects page if you want their targets tracked.`
          );
        } else {
          notify(`Imported ${records.length} work records`);
        }
      } catch (err) {
        console.error(err);
        alert("Could not read this Excel file. Please use .xlsx format.");
      } finally {
        setBusy(false);
      }
    };

    reader.readAsArrayBuffer(file);
  }

  function parseAccuracyWorkbook(file) {
    setAccuracyBusy(true);
    setAccuracyFileName(file.name);

    const reader = new FileReader();

    reader.onload = e => {
      try {
        // Same reasoning as parseWorkbook: skip cellDates
        // and let normalizeDateValue convert raw serial
        // numbers with timezone-independent UTC math.
        const wb = XLSX.read(e.target.result, {
          type: "array"
        });

        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {
          header: 1,
          defval: null,
          raw: true
        });

        if (!rows.length) {
          alert("The Accuracy Report is empty.");
          return;
        }

        // Find the header row even if the sheet has a title above it.
        let headerRowIndex = -1;
        let headers = [];

        for (let r = 0; r < Math.min(rows.length, 10); r++) {
          const candidate = rows[r] || [];
          const normalized = candidate.map(normalizeHeader);
          const hasName = normalized.some(h => ["name", "member", "team member", "employee name"].includes(h));
          const hasAccuracy = normalized.some(h => h.includes("accuracy"));
          const hasTP = normalized.some(h => h === "tp" || h.includes("true positive"));
          if (hasName && (hasAccuracy || hasTP)) {
            headerRowIndex = r;
            headers = candidate;
            break;
          }
        }

        if (headerRowIndex === -1) {
          headerRowIndex = 0;
          headers = rows[0] || [];
        }

        const dateIndex = findHeaderIndex(headers, [
          "date", "report date", "accuracy date", "day"
        ]);
        const nameIndex = findHeaderIndex(headers, [
          "name", "member", "team member", "employee name", "employee"
        ]);
        const dailyCountIndex = findHeaderIndex(headers, [
          "daily count", "dailycount", "count", "daily output", "images count", "daily images"
        ]);
        const tpIndex = findHeaderIndex(headers, [
          "tp", "true positive", "true positives"
        ]);
        const fpIndex = findHeaderIndex(headers, [
          "fp", "false positive", "false positives"
        ]);
        const fnIndex = findHeaderIndex(headers, [
          "fn", "false negative", "false negatives"
        ]);
        const accuracyIndex = findHeaderIndex(headers, [
          "accuracy", "accuracy ", "accuracy percent", "accuracy percentage"
        ]);
        const scoreIndex = findHeaderIndex(headers, [
          "score", "qa score", "quality score"
        ]);
        const linkIndex = findHeaderIndex(headers, [
          "link to the detailed report", "detailed report", "report link", "link", "detailed report link"
        ]);
        const imagesUsedIndex = findHeaderIndex(headers, [
          "images used for calculating accuracy", "images used", "images used for accuracy", "accuracy images", "images"
        ]);
        const commentIndex = findHeaderIndex(headers, [
          "comment", "comments", "remark", "remarks", "note", "notes"
        ]);
        // Optional — most Accuracy Report sheets don't have this column yet,
        // but if one does, Reviewer performance can use it automatically.
        const reviewerIndex = findHeaderIndex(headers, [
          "reviewer", "reviewed by", "qa reviewer", "review by"
        ]);

        if (nameIndex === -1) {
          alert(`Could not find the Name column. Detected headers: ${headers.filter(Boolean).join(" | ")}`);
          return;
        }

        if (dailyCountIndex === -1 && tpIndex === -1 && accuracyIndex === -1) {
          alert(`Could not find Accuracy Report data columns. Detected headers: ${headers.filter(Boolean).join(" | ")}`);
          return;
        }

        const records = [];

        for (let r = headerRowIndex + 1; r < rows.length; r++) {
          const row = rows[r] || [];
          const name = String(row[nameIndex] ?? "").trim();
          if (!name) continue;

          const date = dateIndex >= 0 ? normalizeDateValue(row[dateIndex]) : "";
          const dailyCount = dailyCountIndex >= 0 ? parseNumber(row[dailyCountIndex]) : 0;
          const tp = tpIndex >= 0 ? parseNumber(row[tpIndex]) : 0;
          const fp = fpIndex >= 0 ? parseNumber(row[fpIndex]) : 0;
          const fn = fnIndex >= 0 ? parseNumber(row[fnIndex]) : 0;

          let accuracy = accuracyIndex >= 0 ? parseAccuracyValue(row[accuracyIndex]) : 0;
          if (!accuracy && tp + fp + fn > 0) {
            accuracy = (tp / (tp + fp + fn)) * 100;
          }

          const score = scoreIndex >= 0 ? parseNumber(row[scoreIndex]) : 0;
          const link = linkIndex >= 0 ? String(row[linkIndex] ?? "").trim() : "";
          const imagesUsed = imagesUsedIndex >= 0 ? String(row[imagesUsedIndex] ?? "").trim() : "";
          const comment = commentIndex >= 0 ? String(row[commentIndex] ?? "").trim() : "";
          const reviewer = reviewerIndex >= 0 ? String(row[reviewerIndex] ?? "").trim() : "";

          records.push({
            id: `accuracy-${Date.now()}-${r}`,
            date,
            name,
            dailyCount,
            tp,
            fp,
            fn,
            accuracy: Math.max(0, Math.min(100, accuracy)),
            score,
            link,
            imagesUsed,
            comment,
            reviewer
          });
        }

        if (!records.length) {
          alert("No team member records were found in the Accuracy Report.");
          return;
        }

        update({
          ...data,
          accuracyRecords: records,
          accuracyFile: file.name,
          accuracyLastSync: new Date().toISOString()
        });

        setAccuracyPreview(records.slice(0, 25));
        notify(`Imported Accuracy Report for ${records.length} team members`);
        logActivity(`Imported Accuracy Report "${file.name}" (${records.length} team members)`);
      } catch (err) {
        console.error(err);
        alert("Could not read this Accuracy Report. Please use .xlsx format and check the column names.");
      } finally {
        setAccuracyBusy(false);
      }
    };

    reader.readAsArrayBuffer(file);
  }

  const storedAccuracy = accuracyPreview.length
    ? accuracyPreview
    : data.accuracyRecords || [];

  return (
    <Page
      title="Sheet Import"
      subtitle="Import your Daily Effort Sheet and Accuracy Report. Both sources will update the dashboard automatically."
    >
      <div className="grid two">
        <Panel title="Daily Effort Sheet">
          <div className="import-box">
            <Upload size={28} />
            <h3>{busy ? "Importing..." : "Upload your .xlsx file"}</h3>
            <p>Use Google Sheets → File → Download → Microsoft Excel (.xlsx).</p>
            <label className="primary upload-label">
              <Upload size={17} />
              Choose Excel file
              <input
                type="file"
                accept=".xlsx,.xls"
                hidden
                onChange={e => e.target.files?.[0] && parseWorkbook(e.target.files[0])}
              />
            </label>
            {fileName && (
              <div className="import-success">
                <CheckCircle2 size={17} />
                <span>
                  <b>{fileName}</b>
                  <small>
                    Last imported: {data.sheetLastSync ? new Date(data.sheetLastSync).toLocaleString() : "just now"}
                  </small>
                </span>
              </div>
            )}
          </div>
        </Panel>

        <Panel title="Accuracy Report">
          <div className="import-box">
            <CheckCircle2 size={28} />
            <h3>{accuracyBusy ? "Importing Accuracy Report..." : "Upload Accuracy Report"}</h3>
            <p>
              Required fields: Name, Daily Count, TP, FP, FN, Accuracy(%), Score, Detailed Report Link, Images Used and Comment.
            </p>
            <label className="primary upload-label">
              <Upload size={17} />
              Choose Accuracy Excel
              <input
                type="file"
                accept=".xlsx,.xls"
                hidden
                onChange={e => e.target.files?.[0] && parseAccuracyWorkbook(e.target.files[0])}
              />
            </label>
            {accuracyFileName && (
              <div className="import-success">
                <CheckCircle2 size={17} />
                <span>
                  <b>{accuracyFileName}</b>
                  <small>
                    Last imported: {data.accuracyLastSync ? new Date(data.accuracyLastSync).toLocaleString() : "just now"}
                  </small>
                </span>
              </div>
            )}
          </div>
        </Panel>
      </div>

      <div className="grid two">
        <Panel title="Daily Effort sheet structure">
          <div className="structure-list">
            <div><b>Name</b><span>Team member</span></div>
            <div><b>Project</b><span>Project worked on</span></div>
            <div><b>Annotation/Review</b><span>Work type</span></div>
            <div><b>Total images worked</b><span>Daily output</span></div>
            <div><b>Link to the range</b><span>Work/range link</span></div>
          </div>
        </Panel>

        <Panel title="Accuracy Report structure">
          <div className="structure-list">
            <div><b>Name</b><span>Team member</span></div>
            <div><b>Daily Count</b><span>Daily images completed</span></div>
            <div><b>TP / FP / FN</b><span>Accuracy counts</span></div>
            <div><b>Accuracy(%) / Score</b><span>Quality result</span></div>
            <div><b>Detailed Report / Images / Comment</b><span>Supporting QA details</span></div>
          </div>
        </Panel>
      </div>

      <Panel title={`Imported daily records ${data.sheetRecords?.length ? `(${data.sheetRecords.length})` : ""}`}>
        {preview.length === 0 && !data.sheetRecords?.length ? (
          <p className="muted">Upload the Daily Effort Excel file to preview and sync daily work.</p>
        ) : (
          <div className="table-wrap">
            <table>
              <thead>
                <tr><th>Date</th><th>Name</th><th>Project</th><th>Type</th><th>Images worked</th></tr>
              </thead>
              <tbody>
                {(preview.length ? preview : data.sheetRecords.slice(0, 25)).map((x, i) => (
                  <tr key={x.id || i}>
                    <td>{x.date}</td>
                    <td><b>{x.name}</b></td>
                    <td>{x.project}</td>
                    <td>{x.type}</td>
                    <td><b>{Number(x.worked || 0).toLocaleString()}</b></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Panel>

      <Panel title={`Imported Accuracy records ${storedAccuracy.length ? `(${storedAccuracy.length})` : ""}`}>
        {!storedAccuracy.length ? (
          <p className="muted">Upload the Accuracy Report to preview the team accuracy data.</p>
        ) : (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Date</th><th>Name</th><th>Daily Count</th><th>TP</th><th>FP</th><th>FN</th>
                  <th>Accuracy</th><th>Score</th><th>Detailed Report</th><th>Images Used</th><th>Comment</th>
                </tr>
              </thead>
              <tbody>
                {storedAccuracy.map((x, i) => (
                  <tr key={x.id || i}>
                    <td>{x.date || "—"}</td>
                    <td><b>{x.name}</b></td>
                    <td>{Number(x.dailyCount || 0).toLocaleString()}</td>
                    <td>{Number(x.tp || 0).toLocaleString()}</td>
                    <td>{Number(x.fp || 0).toLocaleString()}</td>
                    <td>{Number(x.fn || 0).toLocaleString()}</td>
                    <td><b>{Number(x.accuracy || 0).toFixed(1)}%</b></td>
                    <td><b>{Number(x.score || 0).toFixed(1)}</b></td>
                    <td>
                      {x.link ? (
                        <a className="link-btn" href={x.link} target="_blank" rel="noreferrer">
                          Open <ExternalLink size={14} />
                        </a>
                      ) : "—"}
                    </td>
                    <td>{x.imagesUsed || "—"}</td>
                    <td>{x.comment || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Panel>
    </Page>
  );
}

/* =========================================================
   SETTINGS / PAGE / MODAL
========================================================= */
function SettingsPage({
  exportData,
  importData,
  clearImportedData,
  role,
  storageOnline,
  email,
  onSignOut,
  canManage,
  notify,
  myId,
  isAdmin
}) {
  const [prefs, setPrefs] = useState(loadPrefs);
  const [activity, setActivity] = useState(getActivityLog);
  const [showPasswordForm, setShowPasswordForm] = useState(false);

  function updatePref(key, value) {
    const next = { ...prefs, [key]: value };
    setPrefs(next);
    savePrefs(next);
  }

  function handleClearActivity() {
    if (!confirm("Clear this device's activity history? This cannot be undone.")) return;
    clearActivityLog();
    setActivity([]);
  }

  return (
    <Page
      title="Settings"
      subtitle="Manage how the dashboard looks, your data, and account security."
    >
      <Panel title="General">
        <div className="settings-row">
          <div>
            <b>Landing page</b>
            <p>Which page opens first when you log in. Applies next time you open the dashboard.</p>
          </div>

          <select
            value={prefs.landingPage}
            onChange={e => updatePref("landingPage", e.target.value)}
          >
            <option value="dashboard">Dashboard</option>
            <option value="team">Team</option>
            <option value="attendance">Attendance</option>
            <option value="projects">Projects</option>
            <option value="qa">QA & Reviews</option>
            <option value="issues">Issues</option>
            <option value="analytics">Analytics</option>
            <option value="reports">Reports</option>
          </select>
        </div>

        <div className="settings-row">
          <div>
            <b>Default report range</b>
            <p>What Reports and Attendance use as the starting date range before you pick one yourself.</p>
          </div>

          <select
            value={prefs.reportRangeMode}
            onChange={e => updatePref("reportRangeMode", e.target.value)}
          >
            <option value="lastImportedWeek">Latest imported week</option>
            <option value="thisWeek">This calendar week</option>
          </select>
        </div>

        <p className="muted settings-note">
          These are personal display preferences saved on this browser only — they don't change what your team sees.
        </p>
      </Panel>

      {isAdmin && <UserAccessPanel myId={myId} notify={notify} />}

      <Panel title="Data management">
        <div className="settings-row">
          <div>
            <b>Export backup</b>
            <p>Download all team, project and issue data as JSON.</p>
          </div>

          <button className="secondary" onClick={exportData}>
            <Download size={17} />
            Export
          </button>
        </div>

        <div className="settings-row">
          <div>
            <b>Import backup</b>
            <p>Restore a previously exported dashboard backup.</p>
          </div>

          <label className="secondary">
            <Upload size={17} />
            Import
            <input
              type="file"
              accept=".json"
              onChange={importData}
              hidden
            />
          </label>
        </div>

        <div className="settings-row">
          <div>
            <b>Clear imported data</b>
            <p>Removes all imported sheet and accuracy records, and resets every project's Completed count to 0. Daily Target, Total Images and Deadline are kept as-is.</p>
          </div>

          <button
            className="secondary danger"
            onClick={() => (canManage ? clearImportedData() : notify("Only Admin or Team Lead can clear imported data."))}
          >
            <Trash2 size={17} />
            Clear
          </button>
        </div>

        <div className="settings-row">
          <div>
            <b>Storage</b>
            <p>{storageOnline ? "Dashboard data is connected to Supabase online storage." : "Supabase is not configured, so the dashboard is using localStorage."}</p>
          </div>

          <span className={`status ${storageOnline ? "active" : "pending"}`}>
            <i />
            {storageOnline ? "Online" : "Local"}
          </span>
        </div>
      </Panel>

      <Panel title="Security">
        <div className="settings-row">
          <div>
            <b>Signed-in account</b>
            <p>{email || "Local demo"} • {roleLabel(role)}</p>
          </div>
          {storageOnline && (
            <button className="secondary" onClick={onSignOut}>
              <LogOut size={17} /> Sign out
            </button>
          )}
        </div>

        {storageOnline && (
          <div className="settings-row">
            <div>
              <b>Change password</b>
              <p>Updates the password for your own account only.</p>
            </div>
            <button className="secondary" onClick={() => setShowPasswordForm(v => !v)}>
              <LockKeyhole size={17} /> {showPasswordForm ? "Cancel" : "Change"}
            </button>
          </div>
        )}

        {showPasswordForm && (
          <ChangePasswordForm notify={notify} onDone={() => setShowPasswordForm(false)} />
        )}

        <div className="settings-row settings-row-top">
          <div>
            <b>Activity history</b>
            <p>A record of notable actions on this device only — not synced or shared with your team.</p>
          </div>

          {activity.length > 0 && (
            <button className="link-btn" onClick={handleClearActivity}>
              Clear
            </button>
          )}
        </div>

        {activity.length === 0 ? (
          <p className="muted">No activity recorded on this device yet.</p>
        ) : (
          <ul className="activity-list">
            {activity.map(a => (
              <li key={a.id}>
                <span>{a.message}</span>
                <small>{new Date(a.at).toLocaleString()}</small>
              </li>
            ))}
          </ul>
        )}

        <div className="settings-row-top">
          <b style={{ fontSize: 13 }}>What each role can do</b>
          <div className="role-reference">
            <div><b>Admin</b><span>Everything, including managing users and roles.</span></div>
            <div><b>Team Lead</b><span>Manage team, projects, issues, imports and holidays.</span></div>
            <div><b>Member</b><span>View-only access across the dashboard.</span></div>
          </div>
        </div>
      </Panel>
    </Page>
  );
}

function UserAccessPanel({ myId, notify }) {
  const [users, setUsers] = useState(null);
  const [loadError, setLoadError] = useState("");
  const [savingId, setSavingId] = useState(null);

  useEffect(() => {
    if (!supabase) {
      setUsers([]);
      return undefined;
    }

    let active = true;

    async function load() {
      const { data: rows, error } = await supabase.from("profiles").select("*");
      if (!active) return;

      if (error) {
        console.error("Could not load user list", error);
        setLoadError("Could not load the user list. Your Supabase policies may not allow Admins to read all profiles.");
        setUsers([]);
        return;
      }

      setUsers(rows || []);
    }

    load();
    return () => { active = false; };
  }, []);

  async function changeRole(userId, newRole, oldRole) {
    if (newRole === oldRole) return;
    setSavingId(userId);

    const { error } = await supabase.from("profiles").update({ role: newRole }).eq("id", userId);
    setSavingId(null);

    if (error) {
      console.error("Could not update role", error);
      notify("Could not update that user's role. Check your Supabase policies allow this.");
      return;
    }

    setUsers(prev => prev.map(u => (u.id === userId ? { ...u, role: newRole } : u)));
    notify("Role updated");
    logActivity(`Changed a team member's role to ${roleLabel(newRole)}`);
  }

  return (
    <Panel title="User & access">
      <p className="muted settings-note">
        New people sign up on the login screen — set their role here afterward. Admin can manage everything; Team Lead can manage team, projects, issues and imports; Member is view-only.
      </p>

      {users === null ? (
        <p className="muted">Loading users…</p>
      ) : loadError ? (
        <p className="muted">{loadError}</p>
      ) : users.length === 0 ? (
        <p className="muted">No users found yet.</p>
      ) : (
        <ul className="user-list">
          {users.map(u => {
            const name = u.full_name || u.name || u.email || "Unnamed";
            const contact = u.email || u.user_email || u.contact_email || `ID ${String(u.id).slice(0, 8)}…`;
            return (
              <li key={u.id}>
                <div className="user-list-identity">
                  <div className="mini-avatar">{String(name).slice(0, 2).toUpperCase()}</div>
                  <div>
                    <b>{name}</b>
                    <span>{contact}</span>
                  </div>
                </div>

                <select
                  value={u.role || "member"}
                  disabled={u.id === myId || savingId === u.id}
                  onChange={e => changeRole(u.id, e.target.value, u.role)}
                >
                  <option value="admin">Admin</option>
                  <option value="team_lead">Team Lead</option>
                  <option value="member">Member</option>
                </select>
              </li>
            );
          })}
        </ul>
      )}
    </Panel>
  );
}

function ChangePasswordForm({ notify, onDone }) {
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [busy, setBusy] = useState(false);

  async function submit(e) {
    e.preventDefault();
    if (password.length < 6) return notify("Password must be at least 6 characters.");
    if (password !== confirm) return notify("Passwords don't match.");

    setBusy(true);
    const { error } = await supabase.auth.updateUser({ password });
    setBusy(false);

    if (error) {
      notify(`Could not change password: ${error.message}`);
      return;
    }

    notify("Password updated");
    logActivity("Changed account password");
    setPassword("");
    setConfirm("");
    onDone();
  }

  return (
    <form className="password-form" onSubmit={submit}>
      <label>
        New password
        <input
          type="password"
          value={password}
          onChange={e => setPassword(e.target.value)}
          minLength={6}
          required
        />
      </label>
      <label>
        Confirm new password
        <input
          type="password"
          value={confirm}
          onChange={e => setConfirm(e.target.value)}
          minLength={6}
          required
        />
      </label>
      <button className="primary" type="submit" disabled={busy}>
        {busy ? "Saving…" : "Save new password"}
      </button>
    </form>
  );
}

function Page({
  title,
  subtitle,
  action,
  onAction,
  children
}) {
  return (
    <div>
      <div className="page-head">
        <div>
          <p className="eyebrow">TEAM OPERATIONS</p>
          <h1>{title}</h1>
          <p className="sub">{subtitle}</p>
        </div>

        {action && (
          <button className="primary" onClick={onAction}>
            <Plus size={18} />
            {action.replace("+ ", "")}
          </button>
        )}
      </div>

      {children}
    </div>
  );
}

function TeamProfileModal({ name, data, onClose }) {
  const member = data.team.find(m => m.name === name) || {
    name, role: "—", status: "—", target: 0, completed: 0, reviewed: 0, errors: 0
  };

  const myRecords = (data.sheetRecords || []).filter(r => r.name === name);
  const workRecords = myRecords.filter(isWorkRow);

  const assignedProjects = [...new Set(
    workRecords.map(r => getConfiguredProjectName(r.project)).filter(Boolean)
  )];

  const allTimeCompleted = workRecords.reduce((s, r) => s + (Number(r.worked) || 0), 0);

  const qa = (data.accuracyRecords || [])
    .filter(r => r.name === name)
    .reduce(
      (acc, r) => {
        acc.tp += Number(r.tp) || 0;
        acc.fp += Number(r.fp) || 0;
        acc.fn += Number(r.fn) || 0;
        return acc;
      },
      { tp: 0, fp: 0, fn: 0 }
    );
  const qaDenom = qa.tp + qa.fp + qa.fn;
  const qaAccuracy = qaDenom ? Math.round((qa.tp / qaDenom) * 100) : null;

  // Attendance summary — up to the last 30 tracked days ending at the
  // latest imported date, reusing the same logic as the Attendance page.
  const latestDate = getLatestImportedDate(data);
  const rangeEnd = latestDate || getTodayISO();
  const endDate = dateKeyToLocalDate(rangeEnd);
  let rangeStart = rangeEnd;
  if (endDate) {
    const d = new Date(endDate);
    d.setDate(d.getDate() - 29);
    rangeStart = toISODate(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }
  const attendance = buildAttendanceMatrix(data, { start: rangeStart, end: rangeEnd });
  const myRow = attendance.rows.find(r => r.name === name);
  const attCounts = { Present: 0, Leave: 0, "Week Off": 0, Holiday: 0, "No data": 0 };
  (myRow?.cells || []).forEach(c => {
    attCounts[c.status] = (attCounts[c.status] || 0) + 1;
  });
  const trackedDays = attCounts.Present + attCounts.Leave + attCounts["No data"];
  const attendancePct = trackedDays ? Math.round((attCounts.Present / trackedDays) * 100) : null;

  // Performance history — this person's last 7 work-days.
  const byDate = new Map();
  workRecords.forEach(r => {
    byDate.set(r.date, (byDate.get(r.date) || 0) + (Number(r.worked) || 0));
  });
  const history = [...byDate.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1)).slice(-7);
  const historyMax = Math.max(1, ...history.map(([, v]) => v));

  // Rank today among the whole team, by today's completed count.
  const ranked = [...data.team].sort((a, b) => (Number(b.completed) || 0) - (Number(a.completed) || 0));
  const rankIndex = ranked.findIndex(m => m.name === name);
  const rank = rankIndex >= 0 ? rankIndex + 1 : null;

  return (
    <div className="modal-bg" onClick={onClose}>
      <div className="modal profile-modal" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <div className="person">
            <div className="mini-avatar">{name.slice(0, 1).toUpperCase()}</div>
            <div>
              <b>{name}</b>
              <small>{member.role || "—"} • <Status text={member.status || "Active"} /></small>
            </div>
          </div>

          <button className="icon-btn" type="button" onClick={onClose} aria-label="Close">
            <X />
          </button>
        </div>

        <div className="profile-stats">
          <div>
            <span>Today</span>
            <b>{Number(member.completed || 0).toLocaleString()}</b>
            <small>of {Number(member.target || 0).toLocaleString()} target</small>
          </div>
          <div>
            <span>All-time</span>
            <b>{allTimeCompleted.toLocaleString()}</b>
            <small>images worked</small>
          </div>
          <div>
            <span>QA accuracy</span>
            <b>{qaAccuracy != null ? `${qaAccuracy}%` : "—"}</b>
            <small>{qaDenom ? `${qaDenom.toLocaleString()} reviewed` : "No data"}</small>
          </div>
          <div>
            <span>Rank today</span>
            <b>{rank ? `#${rank}` : "—"}</b>
            <small>of {data.team.length}</small>
          </div>
          <div>
            <span>Errors</span>
            <b className={Number(member.errors) > 15 ? "danger-text" : "good-text"}>
              {Number(member.errors) || 0}
            </b>
            <small>logged on roster</small>
          </div>
        </div>

        <div className="modal-section">
          <b>Assigned projects</b>
          <p className="muted">{assignedProjects.length ? assignedProjects.join(", ") : "None on record"}</p>
        </div>

        <div className="modal-section">
          <b>Attendance — last {trackedDays || 0} tracked day{trackedDays === 1 ? "" : "s"}</b>
          <p className="muted">
            {attendancePct != null ? `${attendancePct}% present` : "No attendance data yet"}
            {attCounts.Leave ? ` • ${attCounts.Leave} on leave` : ""}
            {attCounts["No data"] ? ` • ${attCounts["No data"]} unaccounted` : ""}
          </p>
        </div>

        <div className="modal-section">
          <b>Performance history</b>
          {history.length ? (
            <div className="bars">
              {history.map(([date, value]) => (
                <div className="bar-row" key={date}>
                  <span>{date.slice(5)}</span>
                  <div><i style={{ width: `${(value / historyMax) * 100}%` }} /></div>
                  <b>{value.toLocaleString()}</b>
                </div>
              ))}
            </div>
          ) : (
            <p className="muted">No work history yet.</p>
          )}
        </div>
      </div>
    </div>
  );
}

function ProjectEditModal({ project, team, onClose, onSubmit }) {
  const stats = getProjectStats(project);
  const names = (team || []).map(t => t.name);
  const assignedEmployees = project.assignedEmployees || [];
  const assignedReviewers = project.assignedReviewers || [];

  return (
    <div className="modal-bg">
      <div className="modal">
        <div className="modal-head">
          <div>
            <p className="eyebrow">EDIT PROJECT</p>
            <h2>Edit project</h2>
          </div>

          <button className="icon-btn" type="button" onClick={onClose} aria-label="Close">
            <X />
          </button>
        </div>

        <form onSubmit={onSubmit}>
          <label>
            Project name
            <input
              name="name"
              defaultValue={project.name || ""}
              required
            />
          </label>

          <label>
            Total number of images
            <input
              name="totalImages"
              type="number"
              min="0"
              defaultValue={stats.total}
              required
            />
          </label>

          <label>
            Daily target
            <input
              name="target"
              type="number"
              min="0"
              defaultValue={project.target || 0}
              required
            />
          </label>

          <label>
            Completed
            <input
              name="completed"
              type="number"
              min="0"
              max={stats.total}
              defaultValue={stats.completed}
              required
            />
          </label>

          <label>
            Deadline
            <input
              name="deadline"
              type="date"
              defaultValue={project.deadline || ""}
            />
          </label>

          {names.length > 0 && (
            <>
              <label>Assigned employees</label>
              <div className="checkbox-list">
                {names.map(n => (
                  <label key={`emp-${n}`} className="checkbox-item">
                    <input type="checkbox" name="assignedEmployees" value={n} defaultChecked={assignedEmployees.includes(n)} />
                    {n}
                  </label>
                ))}
              </div>

              <label>Assigned reviewers</label>
              <div className="checkbox-list">
                {names.map(n => (
                  <label key={`rev-${n}`} className="checkbox-item">
                    <input type="checkbox" name="assignedReviewers" value={n} defaultChecked={assignedReviewers.includes(n)} />
                    {n}
                  </label>
                ))}
              </div>
            </>
          )}

          <p className="muted project-edit-note">
            Remaining images and status are calculated automatically from total images and completed values.
          </p>

          <div className="modal-actions">
            <button type="button" className="secondary" onClick={onClose}>
              Cancel
            </button>
            <button className="primary" type="submit">
              <Save size={16} />
              Save changes
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}


function TeamMemberModal({ member, onClose, onSubmit }) {
  const editing = Boolean(member);
  return (
    <div className="modal-bg">
      <div className="modal">
        <div className="modal-head">
          <div><p className="eyebrow">TEAM MEMBER</p><h2>{editing ? "Edit team member" : "Add team member"}</h2></div>
          <button className="icon-btn" onClick={onClose}><X /></button>
        </div>
        <form onSubmit={e => { e.preventDefault(); onSubmit(new FormData(e.currentTarget), member); }}>
          <label>Name<input name="name" defaultValue={member?.name || ""} required /></label>
          <label>Role<select name="role" defaultValue={member?.role || "Annotator"}><option>Member</option><option>Annotator</option><option>Reviewer</option><option>Team Lead</option></select></label>
          <label>Target<input name="target" type="number" min="0" defaultValue={member?.target ?? 1000} /></label>
          <label>Completed<input name="completed" type="number" min="0" defaultValue={member?.completed ?? 0} /></label>
          <label>Reviewed<input name="reviewed" type="number" min="0" defaultValue={member?.reviewed ?? 0} /></label>
          <label>Errors<input name="errors" type="number" min="0" defaultValue={member?.errors ?? 0} /></label>
          <label>Status<select name="status" defaultValue={member?.status || "Active"}><option>Active</option><option>Away</option><option>Inactive</option></select></label>
          <div className="modal-actions"><button type="button" className="secondary" onClick={onClose}>Cancel</button><button className="primary" type="submit">{editing ? "Save changes" : "Add member"}</button></div>
        </form>
      </div>
    </div>
  );
}

function Modal({ type, team, onClose, onSubmit }) {
  const labels = {
    team: ["Add team member", "Name", "Role", "Target", "Completed"],
    project: ["Add project", "Project name", "Total number of images", "Daily target", "Completed", "Status"],
    issue: ["Log issue", "Issue type", "Project", "Owner", "Severity"]
  };

  const l = labels[type];
  const names = (team || []).map(t => t.name);

  return (
    <div className="modal-bg">
      <div className="modal">
        <div className="modal-head">
          <div>
            <p className="eyebrow">NEW RECORD</p>
            <h2>{l[0]}</h2>
          </div>

          <button className="icon-btn" onClick={onClose}>
            <X />
          </button>
        </div>

        <form onSubmit={onSubmit}>
          <label>
            {l[1]}
            <input
              name={
                type === "team"
                  ? "name"
                  : type === "project"
                  ? "name"
                  : "type"
              }
              required
            />
          </label>

          {type === "team" && (
            <>
              <label>
                {l[2]}
                <select name="role">
                  <option>Annotator</option>
                  <option>Reviewer</option>
                  <option>Team Lead</option>
                </select>
              </label>

              <label>
                {l[3]}
                <input
                  name="target"
                  type="number"
                  defaultValue="1000"
                />
              </label>

              <label>
                {l[4]}
                <input
                  name="completed"
                  type="number"
                  defaultValue="0"
                />
              </label>
            </>
          )}

          {type === "project" && (
            <>
              <label>
                Total number of images
                <input
                  name="totalImages"
                  type="number"
                  min="0"
                  required
                  placeholder="e.g. 5000, 10000, 25000"
                />
              </label>

              <label>
                Daily target
                <input
                  name="target"
                  type="number"
                  min="0"
                  required
                  placeholder="e.g. 100, 350, 600, 800, 1000"
                />
              </label>

              <label>
                Completed
                <input
                  name="completed"
                  type="number"
                  min="0"
                  defaultValue="0"
                />
              </label>

              <p className="muted project-edit-note">Remaining images and status are calculated automatically from total images and completed images.</p>

              <label>
                Deadline
                <input name="deadline" type="date" />
              </label>

              {names.length > 0 && (
                <>
                  <label>Assigned employees</label>
                  <div className="checkbox-list">
                    {names.map(n => (
                      <label key={`emp-${n}`} className="checkbox-item">
                        <input type="checkbox" name="assignedEmployees" value={n} />
                        {n}
                      </label>
                    ))}
                  </div>

                  <label>Assigned reviewers</label>
                  <div className="checkbox-list">
                    {names.map(n => (
                      <label key={`rev-${n}`} className="checkbox-item">
                        <input type="checkbox" name="assignedReviewers" value={n} />
                        {n}
                      </label>
                    ))}
                  </div>
                </>
              )}
            </>
          )}

          {type === "issue" && (
            <>
              <label>
                {l[2]}
                <input name="project" required />
              </label>

              <label>
                {l[3]}
                <input name="owner" required />
              </label>

              <label>
                {l[4]}
                <select name="severity">
                  <option>High</option>
                  <option>Medium</option>
                  <option>Low</option>
                </select>
              </label>

              <label>
                Description
                <textarea name="description" rows={3} placeholder="What went wrong?" />
              </label>

              {names.length > 0 && (
                <label>
                  Assigned to
                  <select name="assignedTo">
                    <option value="">Unassigned</option>
                    {names.map(n => <option key={n} value={n}>{n}</option>)}
                  </select>
                </label>
              )}

              <label>
                Due date
                <input name="dueDate" type="date" />
              </label>
            </>
          )}

          <div className="modal-actions">
            <button
              type="button"
              className="secondary"
              onClick={onClose}
            >
              Cancel
            </button>

            <button className="primary" type="submit">
              Save record
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

/* =========================================================
   START APP
========================================================= */
createRoot(document.getElementById("root")).render(<App />);
