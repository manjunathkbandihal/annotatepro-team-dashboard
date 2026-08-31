import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { createRoot } from "react-dom/client";
import {
  LayoutDashboard, Users, FolderKanban, ClipboardCheck,
  AlertTriangle, BarChart3, Settings, Search, Plus, Bell,
  Download, Upload, Menu, X, CheckCircle2, Target,
  Image as ImageIcon, ChevronRight, Trash2, Activity, ShieldCheck,
  LogOut, Mail, LockKeyhole, Pencil, Save, ExternalLink
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
  const total = Math.max(
    0,
    Number(project.target ?? project.total) || 0
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
    { id: 2, name: "Rahul", role: "Annotator", target: 1000, completed: 918, reviewed: 420, errors: 18, status: "Active" },
    { id: 3, name: "Priya", role: "Annotator", target: 1000, completed: 744, reviewed: 390, errors: 9, status: "Active" },
    { id: 4, name: "Arun", role: "Reviewer", target: 400, completed: 372, reviewed: 372, errors: 7, status: "Active" },
    { id: 5, name: "Sneha", role: "Annotator", target: 1000, completed: 1000, reviewed: 610, errors: 6, status: "Active" },
    { id: 6, name: "Kiran", role: "Annotator", target: 1000, completed: 581, reviewed: 240, errors: 21, status: "Away" }
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
  ]
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

            return {
              ...p,
              ...s,
              target: s.total,
              total: s.total
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
    return toISODate(
      value.getFullYear(),
      value.getMonth() + 1,
      value.getDate()
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

function getLatestImportedDate(data) {
  const dates = [
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

  return dates.length
    ? dates[dates.length - 1]
    : "";
}


/* =========================================================
   DATE FILTER
========================================================= */

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
  const [page, setPage] = useState("dashboard");
  const [query, setQuery] = useState("");
  const [sidebar, setSidebar] = useState(true);
  const [showAdd, setShowAdd] = useState(false);
  const [addType, setAddType] = useState("team");
  const [editingMember, setEditingMember] = useState(null);
  const [editingProject, setEditingProject] = useState(null);
  const [toast, setToast] = useState("");

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
    ["projects", "Projects", FolderKanban],
    ["qa", "QA & Reviews", ClipboardCheck],
    ["issues", "Issues", AlertTriangle],
    ["analytics", "Analytics", BarChart3],
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
    const total = data.team.reduce((s, x) => s + (Number(x.target) || 0), 0);
    const done = data.team.reduce((s, x) => s + (Number(x.completed) || 0), 0);
    const reviewed = data.team.reduce((s, x) => s + (Number(x.reviewed) || 0), 0);
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
        ...data.projects,
        {
          id: Date.now(),
          name,
          totalImages,
          total: totalImages,
          target,
          completed,
          remaining,
          status: getProjectStatus(totalImages, completed, remaining),
          deadline: f.get("deadline") || ""
        }
      ];
    } else {
      next.issues = [
        ...data.issues,
        {
          id: Date.now(),
          type: f.get("type"),
          project: f.get("project"),
          owner: f.get("owner"),
          severity: f.get("severity"),
          status: "Open",
          date: new Date().toISOString().slice(0, 10)
        }
      ];
    }

    update(next);
    setShowAdd(false);
    notify("Saved successfully");
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
            deadline: f.get("deadline") || ""
          }
        : project
    );

    update({ ...data, projects });
    setEditingProject(null);
    notify("Project updated successfully");
  }

  function remove(kind, id) {
    if (kind === "projects" && !canManageProjects) return notify("Only Admin or Team Lead can delete projects.");
    if (kind === "issues" && !canManageIssues) return notify("Only Admin or Team Lead can delete issues.");
    if (!confirm("Delete this record?")) return;

    update({
      ...data,
      [kind]: data[kind].filter(x => x.id !== id)
    });

    notify("Deleted");
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
            return {
              ...p,
              ...s,
              target: s.total,
              total: s.total
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
            <button className="icon-btn">
              <Bell size={19} />
              <i />
            </button>
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
            />
          )}

          {page === "projects" && (
            <Projects
              rows={data.projects}
              remove={remove}
              canManage={canManageProjects}
              openAdd={() => {
                setAddType("project");
                setShowAdd(true);
              }}
              openEdit={setEditingProject}
            />
          )}

          {page === "qa" && <QA data={data} />}

          {page === "issues" && (
            <Issues
              rows={data.issues}
              remove={remove}
              canManage={canManageIssues}
              openAdd={() => {
                setAddType("issue");
                setShowAdd(true);
              }}
            />
          )}

          {page === "analytics" && <Analytics data={data} />}

          {page === "settings" && (
            <SettingsPage
              exportData={exportData}
              importData={importData}
              role={role}
              storageOnline={isSupabaseConfigured}
              email={session?.user?.email}
              onSignOut={onSignOut}
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
          onClose={() => setShowAdd(false)}
          onSubmit={addRecord}
        />
      )}

      {editingProject && (
        <ProjectEditModal
          project={editingProject}
          onClose={() => setEditingProject(null)}
          onSubmit={editProject}
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
  const active = data.team.filter(x => x.status === "Active").length;
  const openIssues = data.issues.filter(x => x.status === "Open").length;
  const completion = totals.total
    ? Math.round((totals.done / totals.total) * 100)
    : 0;

  return (
    <div>
      <div className="page-head">
        <div>
          <p className="eyebrow">TEAM OPERATIONS</p>
          <h1>Good evening, Manjunath 👋</h1>
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
          icon={Target}
          label="Daily completion"
          value={totals.done.toLocaleString()}
          note={`${completion}% of target`}
          trend="+8.4%"
        />
        <Metric
          icon={ImageIcon}
          label="Images remaining"
          value={totals.remaining.toLocaleString()}
          note="Across active projects"
          trend="-12.2%"
        />
        <Metric
          icon={Users}
          label="Active team"
          value={active}
          note={`${data.team.length} total members`}
          trend="Live"
        />
        <Metric
          icon={AlertTriangle}
          label="Open issues"
          value={openIssues}
          note="Needs attention"
          trend={openIssues ? "Action" : "Clear"}
        />
      </div>

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
            {data.projects.map(p => {
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
        </Panel>
      </div>

      <div className="grid three">
        <Panel title="QA snapshot">
          <div className="big-number">
            {data.team
              .reduce((s, x) => s + (Number(x.reviewed) || 0), 0)
              .toLocaleString()}
          </div>
          <p className="muted">Images reviewed</p>
          <div className="qa-line">
            <span>Accuracy health</span>
            <b>96.8%</b>
          </div>
          <Progress value={97} />
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
function Team({ rows, data, openAdd, canManage, onEdit, onDelete }) {
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
      const role = /review/i.test(type) ? "Reviewer" : "Annotator";
      const key = `${name}|||${project}|||${role}`;

      if (!grouped[key]) {
        grouped[key] = {
          id: key,
          name,
          project,
          role,
          target: Number(projectTargets[project]) || 0,
          completed: 0,
          reviewed: 0,
          errors: 0,
          status: "Active"
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

  const imported =
    Array.isArray(data.sheetRecords) &&
    data.sheetRecords.length > 0;

  const latestDate = imported
    ? [...new Set(data.sheetRecords.map(x => x.date).filter(Boolean))]
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
      <Panel title={`${projectRows.length} records`}>
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
              {projectRows.map(x => {
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
                        <b>{x.name}</b>
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
   PROJECTS
========================================================= */
function Projects({ rows, remove, openAdd, openEdit, canManage }) {
  return (
    <Page
      title="Projects"
      subtitle="Track workload, daily targets and completion across projects."
      action={canManage ? "+ Add project" : undefined}
      onAction={openAdd}
    >
      <div className="project-cards">
        {rows.map(p => {
          const s = getProjectStats(p);

          return (
            <div className="project-card" key={p.id}>
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

              <h3>{p.name}</h3>

              {/* Status and remaining images are calculated automatically from total images and completed values */}
              <Status text={s.status} />

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
    </Page>
  );
}

/* =========================================================
   QA
========================================================= */
function QA({ data }) {
  const latestDate = getLatestImportedDate(data);
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
function Issues({ rows, remove, openAdd, canManage }) {
  return (
    <Page
      title="Issues tracker"
      subtitle="Capture annotation problems and close them before submission."
      action={canManage ? "+ Log issue" : undefined}
      onAction={openAdd}
    >
      <Panel
        title={`${rows.filter(x => x.status === "Open").length} open issues`}
      >
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Issue</th>
                <th>Project</th>
                <th>Owner</th>
                <th>Severity</th>
                <th>Status</th>
                <th>Date</th>
                <th></th>
              </tr>
            </thead>

            <tbody>
              {rows.map(x => (
                <tr key={x.id}>
                  <td><b>{x.type}</b></td>
                  <td>{x.project}</td>
                  <td>{x.owner}</td>
                  <td>
                    <span className={"severity " + x.severity.toLowerCase()}>
                      {x.severity}
                    </span>
                  </td>
                  <td><Status text={x.status} /></td>
                  <td>{x.date}</td>
                  <td>
                    <button
                      className="delete"
                      onClick={() => remove("issues", x.id)}
                    >
                      <Trash2 size={16} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Panel>
    </Page>
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
    </Page>
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
        const wb = XLSX.read(e.target.result, {
          type: "array",
          cellDates: true
        });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {
          header: 1,
          defval: null,
          raw: true
        });

        const dateStarts = [];

const firstRow = rows[0] || [];

for (let c = 1; c < firstRow.length; c++) {
  const date = normalizeDateValue(firstRow[c]);

  if (date) {
    dateStarts.push({
      c,
      date
    });
  }
}

console.log(
  "DATE STARTS:",
  dateStarts.map(x => ({
    column: x.c,
    date: x.date
  }))
);

console.log("DATE STARTS:", dateStarts);

        const records = [];
        const names = [];

        for (let r = 2; r < rows.length; r++) {
          const name = rows[r]?.[0];
          if (!name) continue;

          const cleanName = String(name).trim();
          if (!names.includes(cleanName)) names.push(cleanName);

          for (const d of dateStarts) {
            const project = rows[r]?.[d.c];
            const type = rows[r]?.[d.c + 1];
            const worked = rows[r]?.[d.c + 2];
            const link = rows[r]?.[d.c + 3];

            console.log("IMPORT CHECK:", {
  date: d.date,
  name: cleanName,
  project,
  type,
  worked,
  link
});

            if (project == null && type == null && worked == null && link == null) continue;

            const ps = String(project ?? "").split("\n");
            const ts = String(type ?? "").split("\n");
            const nums = String(worked ?? "").split("\n");
            const ls = String(link ?? "").split("\n");

            ps.forEach((p, i) => {
              const text = p.trim();
              if (!text) return;
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

        const workDates = [...new Set(records.map(x => x.date).filter(Boolean))].sort();
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

          return {
            id: 1000 + i,
            name,
            role: "Annotator",
            target,
            completed,
            reviewed,
            errors: 0,
            status: leave ? "Away" : "Active"
          };
        });

        const projectMap = {};
        Object.entries(projectTargets).forEach(([name, target]) => {
          const existing = data.projects?.find(p => String(p.name).toLowerCase() === String(name).toLowerCase());
          projectMap[name] = {
            completed: 0,
            target,
            totalImages: Math.max(0, Number(existing?.totalImages ?? existing?.total ?? existing?.target) || target),
            deadline: existing?.deadline || ""
          };
        });

        todayRows.forEach(x => {
          const configuredName = getConfiguredProjectName(x.project);
          if (!configuredName) return;
          if (!projectMap[configuredName]) {
            projectMap[configuredName] = {
              completed: 0,
              target: projectTargets[configuredName] || 0,
              totalImages: Math.max(0, Number(projectTargets[configuredName]) || 0),
              deadline: ""
            };
          }
          projectMap[configuredName].completed += Number(x.worked) || 0;
        });

        const projects = Object.entries(projectMap).map(([name, v], i) => {
          const target = Number(v.target) || 0;
          const totalImages = Math.max(0, Number(v.totalImages ?? target) || 0);
          const completed = Math.max(0, Math.min(totalImages, Number(v.completed) || 0));
          const remaining = Math.max(0, totalImages - completed);
          return {
            id: 2000 + i,
            name,
            target,
            totalImages,
            total: totalImages,
            completed,
            remaining,
            status: getProjectStatus(totalImages, completed, remaining),
            deadline: v.deadline || ""
          };
        });

        update({
          ...data,
          team,
          projects,
          sheetRecords: records,
          sheetFile: file.name,
          sheetLastSync: new Date().toISOString()
        });

        setPreview(records.slice(0, 25));
        notify(`Imported ${records.length} work records`);
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
        const wb = XLSX.read(e.target.result, {
          type: "array",
          cellDates: true
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
            comment
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
function SettingsPage({ exportData, importData, role, storageOnline, email, onSignOut }) {
  return (
    <Page
      title="Settings"
      subtitle="Manage dashboard data and backups."
    >
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
            <b>Storage</b>
            <p>{storageOnline ? "Dashboard data is connected to Supabase online storage." : "Supabase is not configured, so the dashboard is using localStorage."}</p>
          </div>

          <span className={`status ${storageOnline ? "active" : "pending"}`}>
            <i />
            {storageOnline ? "Online" : "Local"}
          </span>
        </div>

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
      </Panel>
    </Page>
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

function ProjectEditModal({ project, onClose, onSubmit }) {
  const stats = getProjectStats(project);

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

function Modal({ type, onClose, onSubmit }) {
  const labels = {
    team: ["Add team member", "Name", "Role", "Target", "Completed"],
    project: ["Add project", "Project name", "Total number of images", "Daily target", "Completed", "Status"],
    issue: ["Log issue", "Issue type", "Project", "Owner", "Severity"]
  };

  const l = labels[type];

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
