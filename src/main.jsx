import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { createRoot } from "react-dom/client";
import {
  LayoutDashboard, Users, FolderKanban, ClipboardCheck,
  AlertTriangle, BarChart3, Settings, Search, Plus, Bell,
  Download, Upload, Menu, X, CheckCircle2, Target, Pencil, Save,
  ExternalLink, Image as ImageIcon, ChevronRight, Trash2, Activity, ShieldCheck
} from "lucide-react";
import "./styles.css";
import { supabase, isSupabaseConfigured } from "./supabaseClient";

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
    Number(project.totalImages ?? project.total ?? project.target) || 0
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
    { id: 1, name: "momah_phase2_april", target: 1000, totalImages: 1000, total: 1000, completed: 0, remaining: 1000, status: "Pending", deadline: "" },
    { id: 2, name: "MBS_Street_Detections_Phase2", target: 1000, totalImages: 1000, total: 1000, completed: 0, remaining: 1000, status: "Pending", deadline: "" },
    { id: 3, name: "MBS_frames_june_phase2", target: 1000, totalImages: 1000, total: 1000, completed: 0, remaining: 1000, status: "Pending", deadline: "" },
    { id: 4, name: "combined_aug_data_1", target: 800, totalImages: 800, total: 800, completed: 0, remaining: 800, status: "Pending", deadline: "" },
    { id: 5, name: "iltwy_73026_53front_1", target: 350, totalImages: 350, total: 350, completed: 0, remaining: 350, status: "Pending", deadline: "" }
  ],
  accuracyRecords: [],
  accuracyFile: "",
  accuracyLastSync: "",
  teamDeleted: [],
  issues: [    { id: 1, type: "Missed labels", project: "PCI_Annotations", owner: "Priya", severity: "High", status: "Open", date: "2026-08-11" },
    { id: 2, type: "Wrong prediction", project: "hase2_july_data_1", owner: "Kiran", severity: "Medium", status: "Open", date: "2026-08-11" },
    { id: 3, type: "Label inconsistency", project: "PCI_Annotations", owner: "Rahul", severity: "Low", status: "Resolved", date: "2026-08-10" }
  ]
};

function loadData() {
  try {
    const saved = JSON.parse(localStorage.getItem("annotatepro-data"));
    if (!saved) return seed;

    return {
      ...seed,
      ...saved,
      team: Array.isArray(saved.team) ? saved.team : seed.team,
      projects: Array.isArray(saved.projects)
        ? saved.projects.map(p => {
            const totalImages = Math.max(0, Number(p.totalImages ?? p.total ?? p.target) || 0);
            const normalized = { ...p, totalImages };
            const s = getProjectStats(normalized);
            return { ...normalized, ...s, totalImages, target: Math.max(0, Number(p.target) || 0), total: totalImages };
          })
        : seed.projects,
      sheetRecords: Array.isArray(saved.sheetRecords)
        ? saved.sheetRecords.map(r => ({ ...r, date: normalizeDateValue(r.date) }))
        : [],
      accuracyRecords: Array.isArray(saved.accuracyRecords)
        ? saved.accuracyRecords.map(r => ({ ...r, date: normalizeDateValue(r.date) }))
        : [],
      accuracyFile: saved.accuracyFile || "",
      accuracyLastSync: saved.accuracyLastSync || "",
      sheetFile: saved.sheetFile || "",
      sheetLastSync: saved.sheetLastSync || "",
      issues: Array.isArray(saved.issues) ? saved.issues : seed.issues,
      teamDeleted: Array.isArray(saved.teamDeleted) ? saved.teamDeleted : []
    };
  } catch {
    return seed;
  }
}

function saveData(data) {
  localStorage.setItem("annotatepro-data", JSON.stringify(data));
}

async function loadOnlineData() {
  if (!isSupabaseConfigured || !supabase) return null;

  const { data: row, error } = await supabase
    .from("dashboard_state")
    .select("data")
    .eq("id", 1)
    .maybeSingle();

  if (error) {
    console.error("Supabase load failed:", error);
    return null;
  }

  if (!row?.data) return null;

  const saved = row.data;
  return {
    ...seed,
    ...saved,
    team: Array.isArray(saved.team) ? saved.team : seed.team,
    projects: Array.isArray(saved.projects)
      ? saved.projects.map(p => {
          const totalImages = Math.max(0, Number(p.totalImages ?? p.total ?? p.target) || 0);
          const normalized = { ...p, totalImages };
          const s = getProjectStats(normalized);
          return { ...normalized, ...s, totalImages, target: Math.max(0, Number(p.target) || 0), total: totalImages };
        })
      : seed.projects,
    sheetRecords: Array.isArray(saved.sheetRecords)
      ? saved.sheetRecords.map(r => ({ ...r, date: normalizeDateValue(r.date) }))
      : [],
    accuracyRecords: Array.isArray(saved.accuracyRecords)
      ? saved.accuracyRecords.map(r => ({ ...r, date: normalizeDateValue(r.date) }))
      : [],
    accuracyFile: saved.accuracyFile || "",
    accuracyLastSync: saved.accuracyLastSync || "",
    sheetFile: saved.sheetFile || "",
    sheetLastSync: saved.sheetLastSync || "",
    issues: Array.isArray(saved.issues) ? saved.issues : seed.issues,
    teamDeleted: Array.isArray(saved.teamDeleted) ? saved.teamDeleted : []
  };
}

async function saveOnlineData(data) {
  if (!isSupabaseConfigured || !supabase) return;

  const { error } = await supabase
    .from("dashboard_state")
    .upsert({
      id: 1,
      data,
      updated_at: new Date().toISOString()
    }, { onConflict: "id" });

  if (error) {
    console.error("Supabase save failed:", error);
    throw error;
  }
}

function getConfiguredProjectName(project) {
  const clean = String(project || "").trim();
  const found = Object.keys(projectTargets).find(
    name => name.toLowerCase() === clean.toLowerCase()
  );
  return found || clean;
}

/* =========================================================
   DATE HELPERS
   Handles Excel serial dates, Date values, and common text dates.
========================================================= */
function pad2(n) {
  return String(n).padStart(2, "0");
}

function toISODate(year, month, day) {
  const y = Number(year), m = Number(month), d = Number(day);
  if (!y || !m || !d) return "";
  const date = new Date(y, m - 1, d);
  if (Number.isNaN(date.getTime())) return "";
  if (date.getFullYear() !== y || date.getMonth() !== m - 1 || date.getDate() !== d) return "";
  return `${y}-${pad2(m)}-${pad2(d)}`;
}

function normalizeDateValue(value) {
  if (value == null || value === "") return "";

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return toISODate(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }

  if (typeof value === "number" && Number.isFinite(value) && value > 20000 && value < 80000) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(excelEpoch.getTime() + Math.round(value * 86400000));
    return toISODate(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
  }

  const text = String(value).trim();
  let m = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) return toISODate(m[1], m[2], m[3]);

  m = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (m) return toISODate(m[3], m[2], m[1]);

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime())
    ? ""
    : toISODate(parsed.getFullYear(), parsed.getMonth() + 1, parsed.getDate());
}

function getDateRange(filter) {
  if (!filter) {
    return { start: "", end: "", label: "Select a date" };
  }

  const start = filter.start || "";
  const end = filter.mode === "range" ? (filter.end || "") : start;

  if (!start) {
    return { start: "", end: "", label: filter.mode === "range" ? "Select a date range" : "Select a date" };
  }

  if (filter.mode === "range" && !end) {
    return { start, end: "", label: `${start} → Select end date` };
  }

  if (filter.mode === "range" && end < start) {
    return { start: end, end: start, label: `${end} → ${start}` };
  }

  return {
    start,
    end,
    label: start === end ? start : `${start} → ${end}`
  };
}

function isDateInRange(date, range) {
  if (!range.start || !range.end || !date) return false;

  // Always compare normalized YYYY-MM-DD values. This prevents
  // stored ISO timestamps (e.g. 2026-08-12T00:00:00.000Z) or
  // Excel/text date formats from being excluded accidentally.
  const normalized = normalizeDateValue(date);
  return normalized >= range.start && normalized <= range.end;
}

function getLatestImportedDate(data) {
  const dates = [
    ...(Array.isArray(data?.sheetRecords) ? data.sheetRecords.map(x => normalizeDateValue(x.date)) : []),
    ...(Array.isArray(data?.accuracyRecords) ? data.accuracyRecords.map(x => normalizeDateValue(x.date)) : [])
  ].filter(Boolean).sort();

  return dates.length ? dates[dates.length - 1] : "";
}

function DateFilter({ value, onChange, data }) {
  const allDates = [
    ...(Array.isArray(data.sheetRecords) ? data.sheetRecords.map(x => normalizeDateValue(x.date)) : []),
    ...(Array.isArray(data.accuracyRecords) ? data.accuracyRecords.map(x => normalizeDateValue(x.date)) : [])
  ].filter(Boolean).sort();

  const firstDate = allDates[0] || "";
  const lastDate = allDates[allDates.length - 1] || "";
  const mode = value?.mode || "single";

  const updateFilter = patch => {
    const next = { ...value, ...patch };

    if (next.mode === "single") {
      next.end = next.start || "";
    }

    onChange(next);
  };

  const switchMode = nextMode => {
    if (nextMode === "range") {
      onChange({
        mode: "range",
        start: value?.start || firstDate,
        end: value?.end || lastDate || value?.start || firstDate
      });
    } else {
      onChange({
        mode: "single",
        start: value?.start || lastDate,
        end: value?.start || lastDate
      });
    }
  };

  return (
    <div className="date-filter">
      <div className="date-filter-info">
        <b>Report date</b>
        <small>
          {mode === "range"
            ? (value?.start && value?.end
                ? `${value.start} → ${value.end}`
                : "Select a start and end date")
            : (value?.start || "Select a date")}
        </small>
      </div>

      <div className="date-filter-controls">
        <div className="date-filter-mode">
          <button
            type="button"
            className={mode === "single" ? "active" : ""}
            onClick={() => switchMode("single")}
          >
            Single date
          </button>
          <button
            type="button"
            className={mode === "range" ? "active" : ""}
            onClick={() => switchMode("range")}
          >
            From → To
          </button>
        </div>

        {mode === "single" ? (
          <input
            type="date"
            value={value?.start || ""}
            min={firstDate || undefined}
            max={lastDate || undefined}
            onChange={e => updateFilter({ start: e.target.value })}
            aria-label="Select report date"
          />
        ) : (
          <div className="date-range-inputs">
            <label>
              <span>From</span>
              <input
                type="date"
                value={value?.start || ""}
                min={firstDate || undefined}
                max={value?.end || lastDate || undefined}
                onChange={e => updateFilter({ start: e.target.value })}
                aria-label="Report start date"
              />
            </label>
            <span className="date-range-arrow">→</span>
            <label>
              <span>To</span>
              <input
                type="date"
                value={value?.end || ""}
                min={value?.start || firstDate || undefined}
                max={lastDate || undefined}
                onChange={e => updateFilter({ end: e.target.value })}
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
  const [data, setData] = useState(loadData);
  const [page, setPage] = useState("dashboard");
  const [storageStatus, setStorageStatus] = useState(
    isSupabaseConfigured ? "connecting" : "local"
  );
  const [query, setQuery] = useState("");
  const [sidebar, setSidebar] = useState(true);
  const [showAdd, setShowAdd] = useState(false);
  const [addType, setAddType] = useState("team");
  const [editingProject, setEditingProject] = useState(null);
  const [toast, setToast] = useState("");

  useEffect(() => {
    let active = true;

    async function syncFromCloud() {
      if (!isSupabaseConfigured) {
        setStorageStatus("local");
        return;
      }

      const cloudData = await loadOnlineData();

      if (!active) return;

      if (cloudData) {
        setData(cloudData);
        saveData(cloudData);
        setStorageStatus("online");
      } else {
        // First online run: publish the existing local dashboard data.
        try {
          await saveOnlineData(data);
          if (active) setStorageStatus("online");
        } catch {
          if (active) setStorageStatus("error");
        }
      }
    }

    syncFromCloud();

    return () => {
      active = false;
    };
  }, []);

  const update = next => {
    setData(next);
    saveData(next);

    if (isSupabaseConfigured) {
      saveOnlineData(next)
        .then(() => setStorageStatus("online"))
        .catch(() => setStorageStatus("error"));
    }
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
    const f = new FormData(e.currentTarget);
    let next = { ...data };

    if (addType === "team") {
      const name = String(f.get("name") || "").trim();

      const duplicate = data.team.some(
        (member) =>
          String(member.name || "").trim().toLowerCase() ===
          name.toLowerCase()
      );

      if (duplicate) {
        notify("This member already exists. Edit the existing row instead.");
        return;
      }

      const target = Number(f.get("target") || 1000);
      next.team = [
        ...data.team,
        {
          id: Date.now(),
          name,
          role: f.get("role"),
          target,
          completed: Number(f.get("completed") || 0),
          reviewed: 0,
          errors: 0,
          status: "Active"
        }
      ];

      next.teamDeleted = (data.teamDeleted || []).filter(
        deletedName =>
          String(deletedName).toLowerCase() !== name.toLowerCase()
      );
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
          deadline: f.get("deadline")
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
    if (!confirm("Delete this record?")) return;

    update({
      ...data,
      [kind]: data[kind].filter(x => x.id !== id)
    });

    notify("Deleted");
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
            const totalImages = Math.max(0, Number(p.totalImages ?? p.total ?? p.target) || 0);
            const normalized = { ...p, totalImages };
            const s = getProjectStats(normalized);
            return {
              ...normalized,
              ...s,
              totalImages,
              target: Math.max(0, Number(p.target) || 0),
              total: totalImages
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
            <div className="avatar">MB</div>
            <div className="user">
              <b>Manjunath</b>
              <span>Team Lead</span>
            </div>
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
              update={update}
              notify={notify}
              openAdd={() => {
                setAddType("team");
                setShowAdd(true);
              }}
            />
          )}

          {page === "projects" && (
            <Projects
              rows={data.projects}
              remove={remove}
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
            />
          )}

          {page === "sheet" && (
            <SheetImport data={data} update={update} notify={notify} />
          )}
        </section>
      </main>

      {showAdd && (
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
function Team({ rows, data, update, notify, openAdd }) {
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});

  /*
    Edits made in the Team page are stored in data.teamEdits.
    This keeps Excel-imported data intact while allowing you to
    correct/update the visible member/project/target/etc. directly.
  */
  const getEditKey = (row) =>
    `${String(row.name || "").trim().toLowerCase()}|||${String(
      row.project || "—"
    ).trim().toLowerCase()}|||${String(row.role || "").trim().toLowerCase()}`;

  const projectRows = useMemo(() => {
    const records = Array.isArray(data.sheetRecords)
      ? data.sheetRecords
      : [];

    const edits = data.teamEdits || {};
    const deleted = new Set(
      Array.isArray(data.teamDeleted)
        ? data.teamDeleted.map(x => String(x).trim().toLowerCase())
        : []
    );

    if (!records.length) {
      return rows
        .filter(x => !deleted.has(String(x.name || "").trim().toLowerCase()))
        .map((x) => {
        const base = {
          id: `member-${x.id}`,
          name: x.name,
          project: "—",
          role: x.role,
          target: Number(x.target) || 0,
          completed: Number(x.completed) || 0,
          reviewed: Number(x.reviewed) || 0,
          errors: Number(x.errors) || 0,
          status: x.status || "Active"
        };

        return {
          ...base,
          ...(edits[getEditKey(base)] || {})
        };
      });
    }

    const dates = [
      ...new Set(records.map((x) => x.date).filter(Boolean))
    ].sort();

    const latestDate = dates[dates.length - 1];

    const latestRecords = records.filter(
      (x) =>
        x.date === latestDate &&
        x.project &&
        !["Saturday", "Sunday", "On Leave"].includes(x.project)
    );

    const grouped = {};

    latestRecords.forEach((record) => {
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

    rows.forEach((member) => {
      const exists = result.some(
        (x) =>
          x.name.toLowerCase() ===
          String(member.name || "").toLowerCase()
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

    return result
      .filter(row => !deleted.has(String(row.name || "").trim().toLowerCase()))
      .map((row) => ({
        ...row,
        ...(edits[getEditKey(row)] || {})
      }));
  }, [data.sheetRecords, data.teamEdits, data.teamDeleted, rows]);

  const imported =
    Array.isArray(data.sheetRecords) &&
    data.sheetRecords.length > 0;

  const latestDate = imported
    ? [...new Set(data.sheetRecords.map((x) => x.date).filter(Boolean))]
        .sort()
        .pop()
    : null;

  function startEdit(row) {
    setEditingId(row.id);
    setEditForm({
      name: row.name || "",
      project: row.project === "—" ? "" : row.project || "",
      role: row.role || "Annotator",
      target: Number(row.target) || 0,
      completed: Number(row.completed) || 0,
      reviewed: Number(row.reviewed) || 0,
      errors: Number(row.errors) || 0,
      status: row.status || "Active"
    });
  }

  function cancelEdit() {
    setEditingId(null);
    setEditForm({});
  }

  function updateEditField(field, value) {
    setEditForm((prev) => ({
      ...prev,
      [field]: value
    }));
  }

  function saveEdit(row) {
    const name = String(editForm.name || "").trim();

    if (!name) {
      notify("Member name is required.");
      return;
    }

    const target = Math.max(0, Number(editForm.target) || 0);
    const completed = Math.max(0, Number(editForm.completed) || 0);
    const reviewed = Math.max(0, Number(editForm.reviewed) || 0);
    const errors = Math.max(0, Number(editForm.errors) || 0);
    const project = String(editForm.project || "").trim() || "—";

    const progress = target
      ? Math.min(100, Math.round((completed / target) * 100))
      : 0;

    const status =
      target && progress >= 100
        ? "Completed"
        : completed > 0
        ? "In Progress"
        : "Pending";

    const cleaned = {
      name,
      project,
      role: editForm.role || "Annotator",
      target,
      completed,
      reviewed,
      errors,
      status
    };

    /*
      Save the edited row as an override. The imported Excel data
      remains available and will continue to refresh the base values.
    */
    const nextEdits = {
      ...(data.teamEdits || {})
    };

    const oldKey = getEditKey(row);

    // Keep the original row key so the edit continues to apply
    // even after the member/project name is changed.
    nextEdits[oldKey] = cleaned;

    const nextTeam = Array.isArray(data.team)
      ? data.team.map((member) => {
          if (
            String(member.name || "").trim().toLowerCase() ===
            String(row.name || "").trim().toLowerCase()
          ) {
            return {
              ...member,
              name
            };
          }

          return member;
        })
      : data.team;

    update({
      ...data,
      team: nextTeam,
      teamEdits: nextEdits
    });

    setEditingId(null);
    setEditForm({});
    notify("Team record updated");
  }

  function deleteMember(row) {
    const name = String(row.name || "").trim();
    if (!name) return;

    if (!window.confirm(`Delete ${name} from the team? This will remove all of this member's team records from the dashboard.`)) {
      return;
    }

    const key = name.toLowerCase();
    const nextDeleted = [
      ...(Array.isArray(data.teamDeleted) ? data.teamDeleted : []),
      key
    ].filter((value, index, arr) => arr.indexOf(value) === index);

    const nextTeam = (Array.isArray(data.team) ? data.team : []).filter(
      member => String(member.name || "").trim().toLowerCase() !== key
    );

    const nextEdits = Object.fromEntries(
      Object.entries(data.teamEdits || {}).filter(([editKey, value]) => {
        const editName = String(value?.name || "").trim().toLowerCase();
        return editName !== key && !editKey.startsWith(`${name.toLowerCase()}|||`);
      })
    );

    update({
      ...data,
      team: nextTeam,
      teamDeleted: nextDeleted,
      teamEdits: nextEdits
    });

    if (editingId === row.id) cancelEdit();
    notify(`${name} deleted`);
  }

  function inputStyle() {
    return {
      width: "100%",
      minWidth: "85px",
      padding: "7px 8px",
      border: "1px solid #d9deea",
      borderRadius: "7px",
      background: "#fff",
      fontSize: "13px",
      outline: "none"
    };
  }

  return (
    <Page
      title="Team members"
      subtitle={
        imported
          ? `Daily project-wise productivity from imported sheet${
              latestDate ? ` • ${latestDate}` : ""
            }. You can edit any row directly.`
          : "Monitor and edit individual productivity, targets and quality."
      }
      action="+ Add member"
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
                <th>Action</th>
              </tr>
            </thead>

            <tbody>
              {projectRows.map((x) => {
                const isEditing = editingId === x.id;

                const target = isEditing
                  ? Math.max(0, Number(editForm.target) || 0)
                  : Number(x.target) || 0;

                const completed = isEditing
                  ? Math.max(0, Number(editForm.completed) || 0)
                  : Number(x.completed) || 0;

                const reviewed = isEditing
                  ? Math.max(0, Number(editForm.reviewed) || 0)
                  : Number(x.reviewed) || 0;

                const errors = isEditing
                  ? Math.max(0, Number(editForm.errors) || 0)
                  : Number(x.errors) || 0;

                const progress = target
                  ? Math.min(100, Math.round((completed / target) * 100))
                  : 0;

                const calculatedStatus =
                  target && progress >= 100
                    ? "Completed"
                    : completed > 0
                    ? "In Progress"
                    : "Pending";

                return (
                  <tr key={x.id}>
                    <td>
                      {isEditing ? (
                        <input
                          style={inputStyle()}
                          value={editForm.name}
                          onChange={(e) =>
                            updateEditField("name", e.target.value)
                          }
                        />
                      ) : (
                        <div className="person">
                          <div className="mini-avatar">
                            {String(x.name || "")
                              .slice(0, 1)
                              .toUpperCase()}
                          </div>
                          <b>{x.name}</b>
                        </div>
                      )}
                    </td>

                    <td>
                      {isEditing ? (
                        <input
                          style={inputStyle()}
                          value={editForm.project}
                          placeholder="Project name"
                          onChange={(e) =>
                            updateEditField("project", e.target.value)
                          }
                        />
                      ) : (
                        <b>{x.project}</b>
                      )}
                    </td>

                    <td>
                      {isEditing ? (
                        <select
                          style={inputStyle()}
                          value={editForm.role}
                          onChange={(e) =>
                            updateEditField("role", e.target.value)
                          }
                        >
                          <option>Annotator</option>
                          <option>Reviewer</option>
                          <option>Team Lead</option>
                        </select>
                      ) : (
                        x.role
                      )}
                    </td>

                    <td>
                      {isEditing ? (
                        <input
                          style={inputStyle()}
                          type="number"
                          min="0"
                          value={editForm.target}
                          onChange={(e) =>
                            updateEditField("target", e.target.value)
                          }
                        />
                      ) : (
                        target ? target.toLocaleString() : "—"
                      )}
                    </td>

                    <td>
                      {isEditing ? (
                        <input
                          style={inputStyle()}
                          type="number"
                          min="0"
                          value={editForm.completed}
                          onChange={(e) =>
                            updateEditField("completed", e.target.value)
                          }
                        />
                      ) : (
                        <b>{completed.toLocaleString()}</b>
                      )}
                    </td>

                    <td>
                      <div style={{ minWidth: "100px" }}>
                        <Progress value={progress} />
                        <small style={{ display: "block", marginTop: "4px" }}>
                          {progress}%
                        </small>
                      </div>
                    </td>

                    <td>
                      {isEditing ? (
                        <input
                          style={inputStyle()}
                          type="number"
                          min="0"
                          value={editForm.reviewed}
                          onChange={(e) =>
                            updateEditField("reviewed", e.target.value)
                          }
                        />
                      ) : (
                        reviewed.toLocaleString()
                      )}
                    </td>

                    <td>
                      {isEditing ? (
                        <input
                          style={inputStyle()}
                          type="number"
                          min="0"
                          value={editForm.errors}
                          onChange={(e) =>
                            updateEditField("errors", e.target.value)
                          }
                        />
                      ) : (
                        <span
                          className={
                            errors > 15 ? "danger-text" : "good-text"
                          }
                        >
                          {errors}
                        </span>
                      )}
                    </td>

                    <td>
                      {isEditing ? (
                        <Status text={calculatedStatus} />
                      ) : (
                        <Status
                          text={
                            target && progress >= 100
                              ? "Completed"
                              : x.status
                          }
                        />
                      )}
                    </td>

                    <td>
                      {isEditing ? (
                        <div
                          style={{
                            display: "flex",
                            gap: "6px",
                            alignItems: "center"
                          }}
                        >
                          <button
                            type="button"
                            className="primary"
                            style={{
                              padding: "7px 9px",
                              minWidth: "auto"
                            }}
                            title="Save"
                            onClick={() => saveEdit(x)}
                          >
                            <Save size={15} />
                          </button>

                          <button
                            type="button"
                            className="secondary"
                            style={{
                              padding: "7px 9px",
                              minWidth: "auto"
                            }}
                            title="Cancel"
                            onClick={cancelEdit}
                          >
                            <X size={15} />
                          </button>
                        </div>
                      ) : (
                        <button
                          type="button"
                          className="secondary"
                          style={{
                            padding: "7px 10px",
                            minWidth: "auto"
                          }}
                          title="Edit this member"
                          onClick={() => startEdit(x)}
                        >
                          <Pencil size={15} />
                          Edit
                        </button>

                        <button
                          type="button"
                          className="delete"
                          style={{
                            padding: "7px 9px",
                            minWidth: "auto"
                          }}
                          title="Delete this member"
                          aria-label={`Delete ${x.name}`}
                          onClick={() => deleteMember(x)}
                        >
                          <Trash2 size={15} />
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
    </Page>
  );
}

/* =========================================================
   PROJECTS
========================================================= */
function Projects({ rows, remove, openAdd, openEdit }) {
  return (
    <Page
      title="Projects"
      subtitle="Track workload, daily targets and completion across projects."
      action="+ Add project"
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
   QA & REVIEWS
   Accuracy Report data is shown here after import.
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
function Issues({ rows, remove, openAdd }) {
  return (
    <Page
      title="Issues tracker"
      subtitle="Capture annotation problems and close them before submission."
      action="+ Log issue"
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

  const sheetRows = Array.isArray(data.sheetRecords)
    ? data.sheetRecords.filter(x => isDateInRange(x.date, range))
    : [];

  const accuracyRows = Array.isArray(data.accuracyRecords)
    ? data.accuracyRecords.filter(x => isDateInRange(x.date, range))
    : [];

  const memberMap = {};
  sheetRows.forEach(x => {
    const name = String(x.name || "").trim();
    if (!name) return;

    if (!memberMap[name]) {
      memberMap[name] = { name, completed: 0, reviewed: 0 };
    }

    const worked = Number(x.worked) || 0;
    memberMap[name].completed += worked;

    if (/review/i.test(String(x.type || ""))) {
      memberMap[name].reviewed += worked;
    }
  });

  const productivityRows = Object.values(memberMap).sort((a, b) => b.completed - a.completed);
  const max = Math.max(...productivityRows.map(x => x.completed), 1);
  const maxDaily = Math.max(...accuracyRows.map(x => Number(x.dailyCount) || 0), 1);

  const averageAccuracy = accuracyRows.length
    ? accuracyRows.reduce((s, x) => s + (Number(x.accuracy) || 0), 0) / accuracyRows.length
    : 0;

  const averageScore = accuracyRows.length
    ? accuracyRows.reduce((s, x) => s + (Number(x.score) || 0), 0) / accuracyRows.length
    : 0;

  const totalWorked = productivityRows.reduce((s, x) => s + x.completed, 0);
  const totalTarget = data.team.reduce((s, x) => s + (Number(x.target) || 0), 0);

  return (
    <Page
      title="Analytics"
      subtitle={`Productivity and accuracy analytics for ${range.label.toLowerCase()}.`}
    >
      <DateFilter value={period} onChange={setPeriod} data={data} />

      <Panel title={`Completed images by team member — ${range.label}`}>
        {!productivityRows.length ? (
          <p className="muted">No Daily Effort records exist for <b>{range.label}</b>.</p>
        ) : (
          <div className="bars">
            {productivityRows.map(x => (
              <div className="bar-row" key={x.name}>
                <span>{x.name}</span>
                <div><i style={{ width: `${(x.completed / max) * 100}%` }} /></div>
                <b>{x.completed.toLocaleString()}</b>
              </div>
            ))}
          </div>
        )}
      </Panel>

      <div className="grid three">
        <Panel title="Capacity">
          <div className="big-number">{totalTarget ? Math.round((totalWorked / totalTarget) * 100) : 0}%</div>
          <p className="muted">Selected-period work vs current team targets</p>
        </Panel>

        <Panel title="Accuracy health">
          <div className="big-number">{accuracyRows.length ? `${averageAccuracy.toFixed(1)}%` : "—"}</div>
          <p className="muted">Average accuracy from Accuracy Report</p>
          <Progress value={averageAccuracy} />
        </Panel>

        <Panel title="Average score">
          <div className="big-number">{accuracyRows.length ? averageScore.toFixed(1) : "—"}</div>
          <p className="muted">Score from Accuracy Report</p>
        </Panel>
      </div>

      <Panel title={`Accuracy by team member — ${range.label}`}>
        {!accuracyRows.length ? (
          <p className="muted">No Accuracy Report data exists for <b>{range.label}</b>.</p>
        ) : (
          <div className="bars">
            {accuracyRows.map((x, i) => {
              const accuracy = Math.max(0, Math.min(100, Number(x.accuracy) || 0));
              return (
                <div className="bar-row" key={x.id || i}>
                  <span>{x.name}</span>
                  <div><i style={{ width: `${accuracy}%` }} /></div>
                  <b>{accuracy.toFixed(1)}%</b>
                </div>
              );
            })}
          </div>
        )}
      </Panel>

      <Panel title={`Daily count from Accuracy Report — ${range.label}`}>
        {!accuracyRows.length ? (
          <p className="muted">No Accuracy Report data available.</p>
        ) : (
          <div className="bars">
            {accuracyRows.map((x, i) => (
              <div className="bar-row" key={x.id || i}>
                <span>{x.name}</span>
                <div><i style={{ width: `${((Number(x.dailyCount) || 0) / maxDaily) * 100}%` }} /></div>
                <b>{Number(x.dailyCount || 0).toLocaleString()}</b>
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

        for (let c = 1; c < (rows[0]?.length || 0); c++) {
          const v = rows[0][c];
          if (v instanceof Date) {
            const pad = n => String(n).padStart(2, "0");
            dateStarts.push({
              c,
              date: `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`
            });
          }
        }

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
function SettingsPage({ exportData, importData }) {
  return (
    <Page
      title="Settings"
      subtitle="Manage dashboard data and backups."
    >
      <Panel title="Data management">
        <div className="settings-row">
          <div>
            <b>Export backup</b>
            <p>Download all team, project, Accuracy Report and issue data as JSON.</p>
          </div>

          <button className="secondary" onClick={exportData}>
            <Download size={17} />
            Export
          </button>
        </div>

        <div className="settings-row">
          <div>
            <b>Import backup</b>
            <p>Restore a previously exported dashboard backup, including Accuracy Report data.</p>
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
            <p>
              {storageStatus === "online"
                ? "Dashboard data is synchronized with the online Supabase database."
                : storageStatus === "connecting"
                ? "Connecting to the online database..."
                : storageStatus === "error"
                ? "Online storage is configured but could not be reached. Local backup remains active."
                : "Online storage is not configured yet, so this browser is using localStorage."}
            </p>
          </div>

          <span className={"status " + (storageStatus === "online" ? "active" : storageStatus === "error" ? "away" : "pending")}>
            <i />
            {storageStatus === "online" ? "Online" : storageStatus === "connecting" ? "Connecting" : storageStatus === "error" ? "Error" : "Local"}
          </span>
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
                {l[2]}
                <input
                  name="totalImages"
                  type="number"
                  min="0"
                  required
                  placeholder="e.g. 5000, 10000, 25000"
                />
              </label>

              <label>
                {l[3]}
                <input
                  name="target"
                  type="number"
                  min="0"
                  required
                  placeholder="e.g. 350, 800, 1000"
                />
              </label>

              <label>
                {l[4]}
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
