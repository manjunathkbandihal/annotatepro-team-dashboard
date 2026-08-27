import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { createRoot } from "react-dom/client";
import {
  LayoutDashboard,
  Users,
  FolderKanban,
  ClipboardCheck,
  AlertTriangle,
  BarChart3,
  Settings,
  Search,
  Plus,
  Bell,
  Download,
  Upload,
  Menu,
  X,
  CheckCircle2,
  Target,
  Image as ImageIcon,
  ChevronRight,
  Trash2,
  Activity,
  ShieldCheck,
  LogOut,
  Lock,
  UserPlus,
  Save,
  RefreshCw
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
   PROJECT STATUS
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

  completed = Math.min(
    completed,
    total || completed
  );

  const remaining = Math.max(
    0,
    total - completed
  );

  const progress = total
    ? Math.min(
        100,
        Math.round((completed / total) * 100)
      )
    : 0;

  return {
    total,
    completed,
    remaining,
    progress,
    status: getProjectStatus(
      total,
      completed,
      remaining
    )
  };
}

/* =========================================================
   DEFAULT DATA
========================================================= */

const seed = {
  team: [
    {
      id: 1,
      name: "Manjunath",
      role: "Team Lead",
      target: 1000,
      completed: 862,
      reviewed: 540,
      errors: 12,
      status: "Active"
    },
    {
      id: 2,
      name: "Rahul",
      role: "Annotator",
      target: 1000,
      completed: 918,
      reviewed: 420,
      errors: 18,
      status: "Active"
    },
    {
      id: 3,
      name: "Priya",
      role: "Annotator",
      target: 1000,
      completed: 744,
      reviewed: 390,
      errors: 9,
      status: "Active"
    },
    {
      id: 4,
      name: "Arun",
      role: "Reviewer",
      target: 400,
      completed: 372,
      reviewed: 372,
      errors: 7,
      status: "Active"
    },
    {
      id: 5,
      name: "Sneha",
      role: "Annotator",
      target: 1000,
      completed: 1000,
      reviewed: 610,
      errors: 6,
      status: "Active"
    },
    {
      id: 6,
      name: "Kiran",
      role: "Annotator",
      target: 1000,
      completed: 581,
      reviewed: 240,
      errors: 21,
      status: "Away"
    }
  ],

  projects: [
    {
      id: 1,
      name: "momah_phase2_april",
      target: 1000,
      total: 1000,
      completed: 0,
      remaining: 1000,
      status: "Pending",
      deadline: ""
    },
    {
      id: 2,
      name: "MBS_Street_Detections_Phase2",
      target: 1000,
      total: 1000,
      completed: 0,
      remaining: 1000,
      status: "Pending",
      deadline: ""
    },
    {
      id: 3,
      name: "MBS_frames_june_phase2",
      target: 1000,
      total: 1000,
      completed: 0,
      remaining: 1000,
      status: "Pending",
      deadline: ""
    },
    {
      id: 4,
      name: "combined_aug_data_1",
      target: 800,
      total: 800,
      completed: 0,
      remaining: 800,
      status: "Pending",
      deadline: ""
    },
    {
      id: 5,
      name: "iltwy_73026_53front_1",
      target: 350,
      total: 350,
      completed: 0,
      remaining: 350,
      status: "Pending",
      deadline: ""
    }
  ],

  issues: [
    {
      id: 1,
      type: "Missed labels",
      project: "PCI_Annotations",
      owner: "Priya",
      severity: "High",
      status: "Open",
      date: "2026-08-11"
    },
    {
      id: 2,
      type: "Wrong prediction",
      project: "hase2_july_data_1",
      owner: "Kiran",
      severity: "Medium",
      status: "Open",
      date: "2026-08-11"
    },
    {
      id: 3,
      type: "Label inconsistency",
      project: "PCI_Annotations",
      owner: "Rahul",
      severity: "Low",
      status: "Resolved",
      date: "2026-08-10"
    }
  ],

  sheetRecords: []
};

/* =========================================================
   LOCAL FALLBACK
========================================================= */

function loadLocalData() {
  try {
    const saved = JSON.parse(
      localStorage.getItem("annotatepro-data")
    );

    if (!saved) {
      return seed;
    }

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
        : []
    };
  } catch {
    return seed;
  }
}

/* =========================================================
   ONLINE DATA HELPERS
========================================================= */

async function loadOnlineData() {
  if (!supabase) {
    return loadLocalData();
  }

  const { data, error } = await supabase
    .from("dashboard_state")
    .select("data")
    .eq("id", 1)
    .maybeSingle();

  if (error) {
    console.error("Supabase load error:", error);
    throw error;
  }

  if (!data?.data || Object.keys(data.data).length === 0) {
    return seed;
  }

  const saved = data.data;

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
      : []
  };
}

async function saveOnlineData(data) {
  if (!supabase) {
    localStorage.setItem(
      "annotatepro-data",
      JSON.stringify(data)
    );

    return;
  }

  const { error } = await supabase
    .from("dashboard_state")
    .upsert(
      {
        id: 1,
        data,
        updated_at: new Date().toISOString()
      },
      {
        onConflict: "id"
      }
    );

  if (error) {
    console.error(
      "Supabase save error:",
      error
    );

    throw error;
  }
}

/* =========================================================
   PROJECT NAME
========================================================= */

function getConfiguredProjectName(project) {
  const clean = String(project || "").trim();

  const found = Object.keys(projectTargets).find(
    name =>
      name.toLowerCase() === clean.toLowerCase()
  );

  return found || clean;
}

/* =========================================================
   PERMISSIONS
========================================================= */

function canEdit(role) {
  return (
    role === "admin" ||
    role === "team_lead"
  );
}

function canDelete(role) {
  return role === "admin";
}

function canManageProjects(role) {
  return (
    role === "admin" ||
    role === "team_lead"
  );
}

function canManageIssues(role) {
  return (
    role === "admin" ||
    role === "team_lead"
  );
}

/* =========================================================
   LOGIN
========================================================= */

function Login({ onLoggedIn }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
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

    onLoggedIn(data.user);

    setBusy(false);
  }

  return (
    <div className="app">
      <div
        style={{
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: "24px",
          width: "100%"
        }}
      >
        <div
          className="panel"
          style={{
            width: "100%",
            maxWidth: "430px"
          }}
        >
          <div
            style={{
              textAlign: "center",
              marginBottom: "28px"
            }}
          >
            <div
              className="brand-mark"
              style={{
                margin: "0 auto 14px"
              }}
            >
              A
            </div>

            <h1
              style={{
                marginBottom: "6px"
              }}
            >
              AnnotatePro
            </h1>

            <p className="muted">
              Team Operations Dashboard
            </p>
          </div>

          <form onSubmit={handleLogin}>
            <label>
              Email
              <input
                type="email"
                value={email}
                onChange={e =>
                  setEmail(e.target.value)
                }
                placeholder="Enter your email"
                required
              />
            </label>

            <label>
              Password
              <input
                type="password"
                value={password}
                onChange={e =>
                  setPassword(e.target.value)
                }
                placeholder="Enter your password"
                required
              />
            </label>

            {error && (
              <div
                style={{
                  marginBottom: "16px",
                  padding: "12px",
                  borderRadius: "10px",
                  background: "#fff1f2",
                  color: "#be123c"
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
                justifyContent: "center"
              }}
            >
              <Lock size={17} />

              {busy
                ? "Signing in..."
                : "Sign in"}
            </button>
          </form>
        </div>
      </div>
    </div>
  );
}

/* =========================================================
   PROFILE
========================================================= */

async function getProfile(user) {
  if (!supabase || !user) {
    return null;
  }

  const { data, error } = await supabase
    .from("profiles")
    .select("id, full_name, role")
    .eq("id", user.id)
    .single();

  if (error) {
    console.error(
      "Profile load error:",
      error
    );

    return null;
  }

  return data;
}

/* =========================================================
   APP
========================================================= */

function App() {
  const [session, setSession] =
    useState(null);

  const [profile, setProfile] =
    useState(null);

  const [data, setData] =
    useState(null);

  const [loading, setLoading] =
    useState(true);

  const [dataLoading, setDataLoading] =
    useState(false);

  const [page, setPage] =
    useState("dashboard");

  const [query, setQuery] =
    useState("");

  const [sidebar, setSidebar] =
    useState(true);

  const [showAdd, setShowAdd] =
    useState(false);

  const [addType, setAddType] =
    useState("team");

  const [toast, setToast] =
    useState("");

  /* =======================================================
     AUTH SESSION
  ======================================================= */

  useEffect(() => {
    if (!supabase) {
      setLoading(false);
      return;
    }

    let mounted = true;

    async function init() {
      const {
        data: sessionData
      } = await supabase.auth.getSession();

      if (!mounted) return;

      const currentSession =
        sessionData?.session || null;

      setSession(currentSession);

      if (currentSession?.user) {
        const p = await getProfile(
          currentSession.user
        );

        if (!mounted) return;

        setProfile(p);
      }

      setLoading(false);
    }

    init();

    const {
      data: listener
    } = supabase.auth.onAuthStateChange(
      async (_event, newSession) => {
        if (!mounted) return;

        setSession(newSession);

        if (newSession?.user) {
          const p = await getProfile(
            newSession.user
          );

          if (!mounted) return;

          setProfile(p);
        } else {
          setProfile(null);
          setData(null);
        }
      }
    );

    return () => {
      mounted = false;
      listener?.subscription?.unsubscribe();
    };
  }, []);

  /* =======================================================
     LOAD DASHBOARD DATA
  ======================================================= */

  useEffect(() => {
    if (!session?.user) {
      return;
    }

    let mounted = true;

    async function load() {
      setDataLoading(true);

      try {
        const result =
          await loadOnlineData();

        if (mounted) {
          setData(result);
        }
      } catch (error) {
        console.error(error);

        if (mounted) {
          setToast(
            "Could not load online data"
          );

          setData(
            loadLocalData()
          );
        }
      } finally {
        if (mounted) {
          setDataLoading(false);
        }
      }
    }

    load();

    return () => {
      mounted = false;
    };
  }, [session]);

  /* =======================================================
     UPDATE DATA
  ======================================================= */

  const update = async next => {
    setData(next);

    try {
      await saveOnlineData(next);

      notify("Saved online");
    } catch (error) {
      console.error(error);

      notify(
        "Save failed. Check Supabase connection."
      );
    }
  };

  function notify(msg) {
    setToast(msg);

    setTimeout(
      () => setToast(""),
      2500
    );
  }

  /* =======================================================
     LOGOUT
  ======================================================= */

  async function logout() {
    if (supabase) {
      await supabase.auth.signOut();
    }

    setSession(null);
    setProfile(null);
    setData(null);
  }

  /* =======================================================
     NAVIGATION
  ======================================================= */

  const role =
    profile?.role || "member";

  const nav = [
    [
      "dashboard",
      "Dashboard",
      LayoutDashboard,
      true
    ],
    [
      "team",
      "Team",
      Users,
      true
    ],
    [
      "projects",
      "Projects",
      FolderKanban,
      true
    ],
    [
      "qa",
      "QA & Reviews",
      ClipboardCheck,
      true
    ],
    [
      "issues",
      "Issues",
      AlertTriangle,
      true
    ],
    [
      "analytics",
      "Analytics",
      BarChart3,
      true
    ],
    [
      "settings",
      "Settings",
      Settings,
      true
    ],
    [
      "sheet",
      "Sheet Import",
      Upload,
      canEdit(role)
    ]
  ];

  /* =======================================================
     FILTERED TEAM
  ======================================================= */

  const filteredTeam =
    useMemo(
      () =>
        (data?.team || []).filter(
          x =>
            String(
              x.name || ""
            )
              .toLowerCase()
              .includes(
                query.toLowerCase()
              )
        ),
      [data?.team, query]
    );

  /* =======================================================
     TOTALS
  ======================================================= */

  const totals =
    useMemo(() => {
      const team =
        data?.team || [];

      const projects =
        data?.projects || [];

      const total =
        team.reduce(
          (s, x) =>
            s +
            (Number(x.target) ||
              0),
          0
        );

      const done =
        team.reduce(
          (s, x) =>
            s +
            (Number(
              x.completed
            ) || 0),
          0
        );

      const reviewed =
        team.reduce(
          (s, x) =>
            s +
            (Number(
              x.reviewed
            ) || 0),
          0
        );

      const remaining =
        projects.reduce(
          (s, x) =>
            s +
            getProjectStats(
              x
            ).remaining,
          0
        );

      return {
        total,
        done,
        reviewed,
        remaining,
        rate: total
          ? Math.round(
              (done / total) *
                100
            )
          : 0
      };
    }, [data]);

  /* =======================================================
     ADD RECORD
  ======================================================= */

  async function addRecord(e) {
    e.preventDefault();

    if (!canEdit(role)) {
      notify(
        "You do not have permission."
      );

      return;
    }

    const f =
      new FormData(
        e.currentTarget
      );

    let next = {
      ...data
    };

    if (
      addType === "team"
    ) {
      const target =
        Number(
          f.get("target") ||
            1000
        );

      next.team = [
        ...(data.team || []),
        {
          id: Date.now(),
          name: f.get("name"),
          role: f.get("role"),
          target,
          completed: Number(
            f.get(
              "completed"
            ) || 0
          ),
          reviewed: 0,
          errors: 0,
          status: "Active"
        }
      ];
    } else if (
      addType === "project"
    ) {
      const name =
        String(
          f.get("name") ||
            ""
        ).trim();

      const target =
        Number(
          f.get("target") ||
            0
        );

      const completed =
        Number(
          f.get(
            "completed"
          ) || 0
        );

      const remaining =
        Math.max(
          0,
          target -
            completed
        );

      next.projects = [
        ...(data.projects ||
          []),
        {
          id: Date.now(),
          name,
          total: target,
          target,
          completed,
          remaining,
          status:
            getProjectStatus(
              target,
              completed,
              remaining
            ),
          deadline:
            f.get("deadline") ||
            ""
        }
      ];
    } else {
      next.issues = [
        ...(data.issues ||
          []),
        {
          id: Date.now(),
          type: f.get("type"),
          project:
            f.get("project"),
          owner:
            f.get("owner"),
          severity:
            f.get(
              "severity"
            ),
          status: "Open",
          date: new Date()
            .toISOString()
            .slice(0, 10)
        }
      ];
    }

    setShowAdd(false);

    await update(next);
  }

  /* =======================================================
     DELETE
  ======================================================= */

  async function remove(
    kind,
    id
  ) {
    if (!canDelete(role)) {
      notify(
        "Only Admin can delete records."
      );

      return;
    }

    if (
      !confirm(
        "Delete this record?"
      )
    ) {
      return;
    }

    const next = {
      ...data,
      [kind]: (
        data[kind] || []
      ).filter(
        x => x.id !== id
      )
    };

    await update(next);
  }

  /* =======================================================
     EXPORT
  ======================================================= */

  function exportData() {
    const blob =
      new Blob(
        [
          JSON.stringify(
            data,
            null,
            2
          )
        ],
        {
          type:
            "application/json"
        }
      );

    const a =
      document.createElement(
        "a"
      );

    a.href =
      URL.createObjectURL(
        blob
      );

    a.download =
      "annotatepro-backup.json";

    a.click();

    URL.revokeObjectURL(
      a.href
    );

    notify(
      "Backup exported"
    );
  }

  /* =======================================================
     IMPORT BACKUP
  ======================================================= */

  async function importData(e) {
    if (!canEdit(role)) {
      notify(
        "You do not have permission."
      );

      return;
    }

    const file =
      e.target.files?.[0];

    if (!file) return;

    const r =
      new FileReader();

    r.onload = async () => {
      try {
        const parsed =
          JSON.parse(
            r.result
          );

        if (
          parsed.team &&
          parsed.projects &&
          parsed.issues
        ) {
          const projects =
            parsed.projects.map(
              p => {
                const s =
                  getProjectStats(
                    p
                  );

                return {
                  ...p,
                  ...s,
                  target:
                    s.total,
                  total:
                    s.total
                };
              }
            );

          await update({
            ...parsed,
            projects
          });
        } else {
          alert(
            "Invalid dashboard backup"
          );
        }
      } catch {
        alert(
          "Invalid JSON file"
        );
      }
    };

    r.readAsText(file);

    e.target.value = "";
  }

  /* =======================================================
     LOADING
  ======================================================= */

  if (loading) {
    return (
      <div className="app">
        <div
          style={{
            minHeight:
              "100vh",
            display: "flex",
            alignItems:
              "center",
            justifyContent:
              "center"
          }}
        >
          <div
            style={{
              textAlign:
                "center"
            }}
          >
            <RefreshCw
              size={28}
              className="spin"
            />

            <p>
              Loading AnnotatePro...
            </p>
          </div>
        </div>
      </div>
    );
  }

  /* =======================================================
     SUPABASE CONFIG ERROR
  ======================================================= */

  if (!isSupabaseConfigured) {
    return (
      <div className="app">
        <div
          style={{
            minHeight:
              "100vh",
            display: "flex",
            alignItems:
              "center",
            justifyContent:
              "center",
            padding: "24px"
          }}
        >
          <div
            className="panel"
            style={{
              maxWidth:
                "600px"
            }}
          >
            <h2>
              Supabase is not configured
            </h2>

            <p className="muted">
              Add the VITE_SUPABASE_URL
              and
              VITE_SUPABASE_ANON_KEY
              environment variables
              in Vercel, then redeploy.
            </p>
          </div>
        </div>
      </div>
    );
  }

  /* =======================================================
     LOGIN
  ======================================================= */

  if (!session) {
    return (
      <Login
        onLoggedIn={user => {
          setSession({
            user
          });
        }}
      />
    );
  }

  /* =======================================================
     PROFILE ERROR
  ======================================================= */

  if (!profile) {
    return (
      <div className="app">
        <div
          style={{
            minHeight:
              "100vh",
            display: "flex",
            alignItems:
              "center",
            justifyContent:
              "center",
            padding: "24px"
          }}
        >
          <div
            className="panel"
            style={{
              maxWidth:
                "600px"
            }}
          >
            <ShieldCheck
              size={32}
            />

            <h2>
              Profile not found
            </h2>

            <p className="muted">
              Your Supabase user exists,
              but no dashboard profile
              was found.
            </p>

            <button
              className="secondary"
              onClick={logout}
            >
              <LogOut
                size={17}
              />
              Logout
            </button>
          </div>
        </div>
      </div>
    );
  }

  /* =======================================================
     DATA LOADING
  ======================================================= */

  if (
    dataLoading ||
    !data
  ) {
    return (
      <div className="app">
        <div
          style={{
            minHeight:
              "100vh",
            display: "flex",
            alignItems:
              "center",
            justifyContent:
              "center"
          }}
        >
          <div
            style={{
              textAlign:
                "center"
            }}
          >
            <RefreshCw
              size={28}
            />

            <p>
              Loading online dashboard...
            </p>
          </div>
        </div>
      </div>
    );
  }

  /* =======================================================
     MAIN UI
  ======================================================= */

  return (
    <div className="app">
      <aside
        className={
          "sidebar " +
          (!sidebar
            ? "collapsed"
            : "")
        }
      >
        <div className="brand">
          <div className="brand-mark">
            A
          </div>

          {sidebar && (
            <div>
              <b>
                AnnotatePro
              </b>

              <span>
                Team Operations
              </span>
            </div>
          )}
        </div>

        <nav>
          {nav
            .filter(
              item => item[3]
            )
            .map(
              ([
                key,
                label,
                Icon
              ]) => (
                <button
                  key={key}
                  className={
                    page === key
                      ? "active"
                      : ""
                  }
                  onClick={() =>
                    setPage(
                      key
                    )
                  }
                >
                  <Icon
                    size={19}
                  />

                  {sidebar && (
                    <span>
                      {label}
                    </span>
                  )}
                </button>
              )
            )}
        </nav>

        {sidebar && (
          <div className="side-card">
            <ShieldCheck
              size={20}
            />

            <b>
              {role === "admin"
                ? "Administrator"
                : role ===
                  "team_lead"
                ? "Team Lead"
                : "Team Member"}
            </b>

            <p>
              {role === "admin"
                ? "Full dashboard control."
                : role ===
                  "team_lead"
                ? "Manage team operations."
                : "Dashboard read-only access."}
            </p>
          </div>
        )}

        <button
          className="collapse"
          onClick={() =>
            setSidebar(
              !sidebar
            )
          }
        >
          {sidebar ? (
            <ChevronRight
              size={18}
            />
          ) : (
            <Menu
              size={18}
            />
          )}

          <span>
            {sidebar
              ? "Collapse"
              : "Expand"}
          </span>
        </button>
      </aside>

      <main className="main">
        <header className="topbar">
          <div
            className="mobile-menu"
            onClick={() =>
              setSidebar(
                !sidebar
              )
            }
          >
            <Menu />
          </div>

          <div className="search">
            <Search
              size={18}
            />

            <input
              placeholder="Search team members..."
              value={query}
              onChange={e =>
                setQuery(
                  e.target.value
                )
              }
            />
          </div>

          <div className="top-actions">
            <button className="icon-btn">
              <Bell
                size={19}
              />

              <i />
            </button>

            <div className="avatar">
              {(
                profile.full_name ||
                session.user.email ||
                "U"
              )
                .slice(0, 2)
                .toUpperCase()}
            </div>

            <div className="user">
              <b>
                {profile.full_name ||
                  session.user.email}
              </b>

              <span>
                {role ===
                "team_lead"
                  ? "Team Lead"
                  : role ===
                    "admin"
                  ? "Admin"
                  : "Member"}
              </span>
            </div>

            <button
              className="icon-btn"
              title="Logout"
              onClick={logout}
            >
              <LogOut
                size={19}
              />
            </button>
          </div>
        </header>

        <section className="content">
          {page ===
            "dashboard" && (
            <Dashboard
              totals={totals}
              data={data}
              setPage={setPage}
            />
          )}

          {page === "team" && (
            <Team
              rows={
                filteredTeam
              }
              data={data}
              role={role}
              openAdd={() => {
                if (
                  !canEdit(
                    role
                  )
                ) {
                  notify(
                    "You do not have permission."
                  );

                  return;
                }

                setAddType(
                  "team"
                );

                setShowAdd(
                  true
                );
              }}
              remove={
                canDelete(role)
                  ? remove
                  : null
              }
            />
          )}

          {page ===
            "projects" && (
            <Projects
              rows={
                data.projects
              }
              remove={
                canDelete(role)
                  ? remove
                  : null
              }
              openAdd={() => {
                if (
                  !canManageProjects(
                    role
                  )
                ) {
                  notify(
                    "You do not have permission."
                  );

                  return;
                }

                setAddType(
                  "project"
                );

                setShowAdd(
                  true
                );
              }}
              role={role}
            />
          )}

          {page === "qa" && (
            <QA
              data={data}
            />
          )}

          {page ===
            "issues" && (
            <Issues
              rows={
                data.issues
              }
              remove={
                canDelete(role)
                  ? remove
                  : null
              }
              openAdd={() => {
                if (
                  !canManageIssues(
                    role
                  )
                ) {
                  notify(
                    "You do not have permission."
                  );

                  return;
                }

                setAddType(
                  "issue"
                );

                setShowAdd(
                  true
                );
              }}
            />
          )}

          {page ===
            "analytics" && (
            <Analytics
              data={data}
            />
          )}

          {page ===
            "settings" && (
            <SettingsPage
              exportData={
                exportData
              }
              importData={
                importData
              }
              role={role}
              online
            />
          )}

          {page === "sheet" &&
            canEdit(role) && (
              <SheetImport
                data={data}
                update={update}
                notify={notify}
              />
            )}
        </section>
      </main>

      {showAdd && (
        <Modal
          type={addType}
          onClose={() =>
            setShowAdd(false)
          }
          onSubmit={
            addRecord
          }
        />
      )}

      {toast && (
        <div className="toast">
          <CheckCircle2
            size={18}
          />

          {toast}
        </div>
      )}
    </div>
  );
}

/* =========================================================
   DASHBOARD
========================================================= */

function Dashboard({
  totals,
  data,
  setPage
}) {
  const active =
    data.team.filter(
      x =>
        x.status ===
        "Active"
    ).length;

  const openIssues =
    data.issues.filter(
      x =>
        x.status ===
        "Open"
    ).length;

  const completion =
    totals.total
      ? Math.round(
          (totals.done /
            totals.total) *
            100
        )
      : 0;

  return (
    <div>
      <div className="page-head">
        <div>
          <p className="eyebrow">
            TEAM OPERATIONS
          </p>

          <h1>
            Good evening,
            Manjunath 👋
          </h1>

          <p className="sub">
            Here’s your team's
            operational overview
            for today.
          </p>
        </div>

        <button
          className="primary"
          onClick={() =>
            setPage("team")
          }
        >
          <Users
            size={18}
          />
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
          trend={
            openIssues
              ? "Action"
              : "Clear"
          }
        />
      </div>

      <div className="grid two">
        <Panel
          title="Team performance"
          action="View all"
          onAction={() =>
            setPage("team")
          }
        >
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>
                    Member
                  </th>
                  <th>
                    Progress
                  </th>
                  <th>
                    Completed
                  </th>
                  <th>
                    Status
                  </th>
                </tr>
              </thead>

              <tbody>
                {data.team
                  .slice(
                    0,
                    5
                  )
                  .map(
                    x => (
                      <tr
                        key={
                          x.id
                        }
                      >
                        <td>
                          <div className="person">
                            <div className="mini-avatar">
                              {x.name.slice(
                                0,
                                1
                              )}
                            </div>

                            <div>
                              <b>
                                {
                                  x.name
                                }
                              </b>

                              <small>
                                {
                                  x.role
                                }
                              </small>
                            </div>
                          </div>
                        </td>

                        <td>
                          <Progress
                            value={
                              x.target
                                ? Math.min(
                                    100,
                                    Math.round(
                                      (x.completed /
                                        x.target) *
                                        100
                                    )
                                  )
                                : 0
                            }
                          />
                        </td>

                        <td>
                          <b>
                            {Number(
                              x.completed
                            ).toLocaleString()}
                          </b>{" "}
                          /{" "}
                          {Number(
                            x.target
                          ).toLocaleString()}
                        </td>

                        <td>
                          <Status
                            text={
                              x.status
                            }
                          />
                        </td>
                      </tr>
                    )
                  )}
              </tbody>
            </table>
          </div>
        </Panel>

        <Panel
          title="Project progress"
          action="View projects"
          onAction={() =>
            setPage(
              "projects"
            )
          }
        >
          <div className="project-list">
            {data.projects.map(
              p => {
                const s =
                  getProjectStats(
                    p
                  );

                return (
                  <div
                    className="project-row"
                    key={
                      p.id
                    }
                  >
                    <div className="project-icon">
                      <FolderKanban
                        size={18}
                      />
                    </div>

                    <div className="grow">
                      <div className="row-title">
                        <b>
                          {
                            p.name
                          }
                        </b>

                        <span>
                          {s.total
                            ? `${s.progress}%`
                            : "No target"}
                        </span>
                      </div>

                      <Progress
                        value={
                          s.progress
                        }
                      />

                      <small>
                        {s.total
                          ? `${s.remaining.toLocaleString()} images remaining`
                          : "Set a daily target"}
                      </small>
                    </div>
                  </div>
                );
              }
            )}
          </div>
        </Panel>
      </div>

      <div className="grid three">
        <Panel title="QA snapshot">
          <div className="big-number">
            {data.team
              .reduce(
                (
                  s,
                  x
                ) =>
                  s +
                  (Number(
                    x.reviewed
                  ) ||
                    0),
                0
              )
              .toLocaleString()}
          </div>

          <p className="muted">
            Images reviewed
          </p>

          <div className="qa-line">
            <span>
              Accuracy health
            </span>

            <b>
              96.8%
            </b>
          </div>

          <Progress
            value={97}
          />
        </Panel>

        <Panel title="Today's activity">
          <div className="activity">
            <Activity />

            <div>
              <b>
                {totals.done.toLocaleString()}
              </b>

              <span>
                images completed
              </span>
            </div>
          </div>

          <div className="activity">
            <CheckCircle2 />

            <div>
              <b>
                {totals.reviewed.toLocaleString()}
              </b>

              <span>
                reviews completed
              </span>
            </div>
          </div>

          <div className="activity">
            <AlertTriangle />

            <div>
              <b>
                {openIssues}
              </b>

              <span>
                issues open
              </span>
            </div>
          </div>
        </Panel>

        <Panel title="Quick actions">
          <button
            className="quick"
            onClick={() =>
              setPage("team")
            }
          >
            <Users />
            Update team progress
            <ChevronRight />
          </button>

          <button
            className="quick"
            onClick={() =>
              setPage(
                "issues"
              )
            }
          >
            <AlertTriangle />
            Review open issues
            <ChevronRight />
          </button>

          <button
            className="quick"
            onClick={() =>
              setPage(
                "settings"
              )
            }
          >
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

function Metric({
  icon: Icon,
  label,
  value,
  note,
  trend
}) {
  return (
    <div className="metric">
      <div className="metric-top">
        <div className="metric-icon">
          <Icon
            size={20}
          />
        </div>

        <span className="trend">
          {trend}
        </span>
      </div>

      <h2>
        {value}
      </h2>

      <b>
        {label}
      </b>

      <small>
        {note}
      </small>
    </div>
  );
}

function Progress({
  value
}) {
  return (
    <div className="progress">
      <span
        style={{
          width: `${Math.max(
            0,
            Math.min(
              100,
              Number(
                value
              ) || 0
            )
          )}%`
        }}
      />
    </div>
  );
}

function Status({
  text
}) {
  return (
    <span
      className={
        "status " +
        String(
          text || ""
        )
          .toLowerCase()
          .replaceAll(
            " ",
            "-"
          )
      }
    >
      <i />
      {text}
    </span>
  );
}

function Panel({
  title,
  action,
  onAction,
  children
}) {
  return (
    <div className="panel">
      <div className="panel-head">
        <h3>
          {title}
        </h3>

        {action && (
          <button
            className="link-btn"
            onClick={
              onAction
            }
          >
            {action}

            <ChevronRight
              size={15}
            />
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

function Team({
  rows,
  data,
  role,
  openAdd,
  remove
}) {
  const projectRows =
    useMemo(() => {
      const records =
        Array.isArray(
          data.sheetRecords
        )
          ? data.sheetRecords
          : [];

      if (!records.length) {
        return rows.map(
          x => ({
            id: `member-${x.id}`,
            name: x.name,
            project: "—",
            role: x.role,
            target:
              Number(
                x.target
              ) || 0,
            completed:
              Number(
                x.completed
              ) || 0,
            reviewed:
              Number(
                x.reviewed
              ) || 0,
            errors:
              Number(
                x.errors
              ) || 0,
            status:
              x.status ||
              "Active"
          })
        );
      }

      const dates = [
        ...new Set(
          records
            .map(
              x =>
                x.date
            )
            .filter(
              Boolean
            )
        )
      ].sort();

      const latestDate =
        dates[
          dates.length -
            1
        ];

      const latestRecords =
        records.filter(
          x =>
            x.date ===
              latestDate &&
            x.project &&
            ![
              "Saturday",
              "Sunday",
              "On Leave"
            ].includes(
              x.project
            )
        );

      const grouped =
        {};

      latestRecords.forEach(
        record => {
          const name =
            String(
              record.name ||
                ""
            ).trim();

          const project =
            getConfiguredProjectName(
              record.project
            );

          if (
            !name ||
            !project
          ) {
            return;
          }

          const type =
            String(
              record.type ||
                ""
            ).trim();

          const recordRole =
            /review/i.test(
              type
            )
              ? "Reviewer"
              : "Annotator";

          const key = `${name}|||${project}|||${recordRole}`;

          if (
            !grouped[
              key
            ]
          ) {
            grouped[
              key
            ] = {
              id: key,
              name,
              project,
              role: recordRole,
              target:
                Number(
                  projectTargets[
                    project
                  ]
                ) || 0,
              completed: 0,
              reviewed: 0,
              errors: 0,
              status:
                "Active"
            };
          }

          const worked =
            Number(
              record.worked
            ) || 0;

          grouped[
            key
          ].completed +=
            worked;

          if (
            /review/i.test(
              type
            )
          ) {
            grouped[
              key
            ].reviewed +=
              worked;
          }
        }
      );

      const result =
        Object.values(
          grouped
        );

      rows.forEach(
        member => {
          const exists =
            result.some(
              x =>
                x.name.toLowerCase() ===
                member.name.toLowerCase()
            );

          if (!exists) {
            result.push({
              id: `member-${member.id}`,
              name: member.name,
              project: "—",
              role: member.role,
              target:
                Number(
                  member.target
                ) || 0,
              completed:
                Number(
                  member.completed
                ) || 0,
              reviewed:
                Number(
                  member.reviewed
                ) || 0,
              errors:
                Number(
                  member.errors
                ) || 0,
              status:
                member.status ||
                "Active"
            });
          }
        }
      );

      return result;
    }, [
      data.sheetRecords,
      rows
    ]);

  const imported =
    Array.isArray(
      data.sheetRecords
    ) &&
    data.sheetRecords
      .length > 0;

  const latestDate =
    imported
      ? [
          ...new Set(
            data.sheetRecords
              .map(
                x =>
                  x.date
              )
              .filter(
                Boolean
              )
          )
        ]
          .sort()
          .pop()
      : null;

  return (
    <Page
      title="Team members"
      subtitle={
        imported
          ? `Daily project-wise productivity from imported sheet${
              latestDate
                ? ` • ${latestDate}`
                : ""
            }.`
          : "Monitor individual productivity, targets and quality."
      }
      action={
        role === "member"
          ? null
          : "+ Add member"
      }
      onAction={
        openAdd
      }
    >
      <Panel
        title={`${projectRows.length} records`}
      >
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>
                  Member
                </th>

                <th>
                  Project
                </th>

                <th>
                  Role
                </th>

                <th>
                  Target
                </th>

                <th>
                  Completed
                </th>

                <th>
                  Progress
                </th>

                <th>
                  Reviewed
                </th>

                <th>
                  Errors
                </th>

                <th>
                  Status
                </th>

                {remove && (
                  <th>
                    Action
                  </th>
                )}
              </tr>
            </thead>

            <tbody>
              {projectRows.map(
                x => {
                  const target =
                    Number(
                      x.target
                    ) || 0;

                  const completed =
                    Number(
                      x.completed
                    ) || 0;

                  const progress =
                    target
                      ? Math.min(
                          100,
                          Math.round(
                            (completed /
                              target) *
                              100
                          )
                        )
                      : 0;

                  return (
                    <tr
                      key={
                        x.id
                      }
                    >
                      <td>
                        <div className="person">
                          <div className="mini-avatar">
                            {x.name
                              .slice(
                                0,
                                1
                              )
                              .toUpperCase()}
                          </div>

                          <b>
                            {
                              x.name
                            }
                          </b>
                        </div>
                      </td>

                      <td>
                        <b>
                          {
                            x.project
                          }
                        </b>
                      </td>

                      <td>
                        {
                          x.role
                        }
                      </td>

                      <td>
                        {target
                          ? target.toLocaleString()
                          : "—"}
                      </td>

                      <td>
                        <b>
                          {completed.toLocaleString()}
                        </b>
                      </td>

                      <td>
                        <div
                          style={{
                            minWidth:
                              "100px"
                          }}
                        >
                          <Progress
                            value={
                              progress
                            }
                          />

                          <small
                            style={{
                              display:
                                "block",
                              marginTop:
                                "4px"
                            }}
                          >
                            {
                              progress
                            }
                            %
                          </small>
                        </div>
                      </td>

                      <td>
                        {Number(
                          x.reviewed
                        ).toLocaleString()}
                      </td>

                      <td>
                        <span
                          className={
                            x.errors >
                            15
                              ? "danger-text"
                              : "good-text"
                          }
                        >
                          {
                            x.errors
                          }
                        </span>
                      </td>

                      <td>
                        <Status
                          text={
                            target &&
                            progress >=
                              100
                              ? "Completed"
                              : x.status
                          }
                        />
                      </td>

                      {remove && (
                        <td>
                          <button
                            className="delete"
                            onClick={() =>
                              remove(
                                "team",
                                x.id
                              )
                            }
                          >
                            <Trash2
                              size={
                                16
                              }
                            />
                          </button>
                        </td>
                      )}
                    </tr>
                  );
                }
              )}
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

function Projects({
  rows,
  remove,
  openAdd,
  role
}) {
  return (
    <Page
      title="Projects"
      subtitle="Track workload, daily targets and completion across projects."
      action={
        canManageProjects(
          role
        )
          ? "+ Add project"
          : null
      }
      onAction={openAdd}
    >
      <div className="project-cards">
        {rows.map(
          p => {
            const s =
              getProjectStats(
                p
              );

            return (
              <div
                className="project-card"
                key={
                  p.id
                }
              >
                <div className="project-card-top">
                  <div className="project-icon">
                    <FolderKanban />
                  </div>

                  {remove && (
                    <button
                      className="delete"
                      onClick={() =>
                        remove(
                          "projects",
                          p.id
                        )
                      }
                    >
                      <Trash2
                        size={
                          16
                        }
                      />
                    </button>
                  )}
                </div>

                <h3>
                  {
                    p.name
                  }
                </h3>

                <Status
                  text={
                    s.status
                  }
                />

                <div className="pc-stat">
                  <span>
                    Daily target
                  </span>

                  <b>
                    {s.total
                      ? s.total.toLocaleString()
                      : "—"}
                  </b>
                </div>

                <div className="pc-stat">
                  <span>
                    Completion
                  </span>

                  <b>
                    {s.total
                      ? `${s.progress}%`
                      : "—"}
                  </b>
                </div>

                <Progress
                  value={
                    s.progress
                  }
                />

                <div className="pc-foot">
                  <span>
                    {s.total
                      ? `${s.remaining.toLocaleString()} remaining`
                      : "Set target"}
                  </span>

                  <span>
                    Due{" "}
                    {p.deadline ||
                      "—"}
                  </span>
                </div>
              </div>
            );
          }
        )}
      </div>
    </Page>
  );
}

/* =========================================================
   QA
========================================================= */

function QA({
  data
}) {
  const reviewed =
    data.team.reduce(
      (s, x) =>
        s +
        (Number(
          x.reviewed
        ) || 0),
      0
    );

  const errors =
    data.team.reduce(
      (s, x) =>
        s +
        (Number(
          x.errors
        ) || 0),
      0
    );

  return (
    <Page
      title="QA & Reviews"
      subtitle="Keep annotation quality visible and actionable."
    >
      <div className="cards">
        <Metric
          icon={
            ClipboardCheck
          }
          label="Total reviewed"
          value={reviewed.toLocaleString()}
          note="Team review volume"
          trend="+6.2%"
        />

        <Metric
          icon={
            CheckCircle2
          }
          label="Quality health"
          value="96.8%"
          note="Based on current issues"
          trend="Good"
        />

        <Metric
          icon={
            AlertTriangle
          }
          label="Total errors"
          value={errors}
          note="Across active members"
          trend="Monitor"
        />
      </div>

      <Panel title="Review readiness">
        <div className="qa-grid">
          <div>
            <h2>
              96.8%
            </h2>

            <p className="muted">
              Overall quality health
            </p>

            <Progress
              value={
                96.8
              }
            />
          </div>

          <div>
            <h2>
              {reviewed.toLocaleString()}
            </h2>

            <p className="muted">
              Reviews completed
            </p>
          </div>

          <div>
            <h2>
              {
                data.issues.filter(
                  x =>
                    x.status ===
                    "Open"
                ).length
              }
            </h2>

            <p className="muted">
              Open QA issues
            </p>
          </div>
        </div>
      </Panel>
    </Page>
  );
}

/* =========================================================
   ISSUES
========================================================= */

function Issues({
  rows,
  remove,
  openAdd
}) {
  return (
    <Page
      title="Issues tracker"
      subtitle="Capture annotation problems and close them before submission."
      action={
        openAdd
          ? "+ Log issue"
          : null
      }
      onAction={
        openAdd
      }
    >
      <Panel
        title={`${rows.filter(
          x =>
            x.status ===
            "Open"
        ).length} open issues`}
      >
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>
                  Issue
                </th>

                <th>
                  Project
                </th>

                <th>
                  Owner
                </th>

                <th>
                  Severity
                </th>

                <th>
                  Status
                </th>

                <th>
                  Date
                </th>

                {remove && (
                  <th />
                )}
              </tr>
            </thead>

            <tbody>
              {rows.map(
                x => (
                  <tr
                    key={
                      x.id
                    }
                  >
                    <td>
                      <b>
                        {
                          x.type
                        }
                      </b>
                    </td>

                    <td>
                      {
                        x.project
                      }
                    </td>

                    <td>
                      {
                        x.owner
                      }
                    </td>

                    <td>
                      <span
                        className={
                          "severity " +
                          x.severity.toLowerCase()
                        }
                      >
                        {
                          x.severity
                        }
                      </span>
                    </td>

                    <td>
                      <Status
                        text={
                          x.status
                        }
                      />
                    </td>

                    <td>
                      {
                        x.date
                      }
                    </td>

                    {remove && (
                      <td>
                        <button
                          className="delete"
                          onClick={() =>
                            remove(
                              "issues",
                              x.id
                            )
                          }
                        >
                          <Trash2
                            size={
                              16
                            }
                          />
                        </button>
                      </td>
                    )}
                  </tr>
                )
              )}
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

function Analytics({
  data
}) {
  const max =
    Math.max(
      ...data.team.map(
        x =>
          Number(
            x.completed
          ) || 0
      ),
      1
    );

  return (
    <Page
      title="Analytics"
      subtitle="Understand productivity trends and team capacity."
    >
      <Panel title="Completed images by team member">
        <div className="bars">
          {data.team.map(
            x => (
              <div
                className="bar-row"
                key={
                  x.id
                }
              >
                <span>
                  {
                    x.name
                  }
                </span>

                <div>
                  <i
                    style={{
                      width: `${
                        ((Number(
                          x.completed
                        ) || 0) /
                          max) *
                        100
                      }%`
                    }}
                  />
                </div>

                <b>
                  {
                    x.completed
                  }
                </b>
              </div>
            )
          )}
        </div>
      </Panel>

      <div className="grid two">
        <Panel title="Capacity">
          <div className="big-number">
            {Math.round(
              (data.team.reduce(
                (
                  s,
                  x
                ) =>
                  s +
                  (Number(
                    x.completed
                  ) ||
                    0),
                0
              ) /
                Math.max(
                  1,
                  data.team.reduce(
                    (
                      s,
                      x
                    ) =>
                      s +
                      (Number(
                        x.target
                      ) ||
                        0),
                      0
                  )
                )) *
                100
            )}
            %
          </div>

          <p className="muted">
            Team target utilization today
          </p>
        </Panel>

        <Panel title="Operational health">
          <div className="health">
            <CheckCircle2 />

            <b>
              Healthy
            </b>

            <span>
              Most active work
              is progressing
              within target.
            </span>
          </div>
        </Panel>
      </div>
    </Page>
  );
}

/* =========================================================
   SHEET IMPORT
========================================================= */

function SheetImport({
  data,
  update,
  notify
}) {
  const [
    preview,
    setPreview
  ] = useState([]);

  const [
    fileName,
    setFileName
  ] = useState("");

  const [
    busy,
    setBusy
  ] = useState(false);

  function parseWorkbook(
    file
  ) {
    setBusy(true);
    setFileName(
      file.name
    );

    const reader =
      new FileReader();

    reader.onload =
      async e => {
        try {
          const wb =
            XLSX.read(
              e.target.result,
              {
                type: "array",
                cellDates:
                  true
              }
            );

          const ws =
            wb.Sheets[
              wb
                .SheetNames[0]
            ];

          const rows =
            XLSX.utils.sheet_to_json(
              ws,
              {
                header: 1,
                defval:
                  null,
                raw: true
              }
            );

          const dateStarts =
            [];

          for (
            let c = 1;
            c <
            (rows[0]
              ?.length ||
              0);
            c++
          ) {
            const v =
              rows[0][
                c
              ];

            if (
              v instanceof
              Date
            ) {
              const pad =
                n =>
                  String(
                    n
                  ).padStart(
                    2,
                    "0"
                  );

              dateStarts.push(
                {
                  c,
                  date: `${v.getFullYear()}-${pad(
                    v.getMonth() +
                      1
                  )}-${pad(
                    v.getDate()
                  )}`
                }
              );
            }
          }

          const records =
            [];

          const names =
            [];

          for (
            let r = 2;
            r <
            rows.length;
            r++
          ) {
            const name =
              rows[r]?.[0];

            if (!name)
              continue;

            const cleanName =
              String(
                name
              ).trim();

            if (
              !names.includes(
                cleanName
              )
            ) {
              names.push(
                cleanName
              );
            }

            for (
              const d of dateStarts
            ) {
              const project =
                rows[r]?.[
                  d.c
                ];

              const type =
                rows[r]?.[
                  d.c + 1
                ];

              const worked =
                rows[r]?.[
                  d.c + 2
                ];

              const link =
                rows[r]?.[
                  d.c + 3
                ];

              if (
                project ==
                  null &&
                type ==
                  null &&
                worked ==
                  null &&
                link ==
                  null
              ) {
                continue;
              }

              const ps =
                String(
                  project ??
                    ""
                ).split(
                  "\n"
                );

              const ts =
                String(
                  type ??
                    ""
                ).split(
                  "\n"
                );

              const nums =
                String(
                  worked ??
                    ""
                ).split(
                  "\n"
                );

              const ls =
                String(
                  link ??
                    ""
                ).split(
                  "\n"
                );

              ps.forEach(
                (
                  p,
                  i
                ) => {
                  const text =
                    p.trim();

                  if (!text)
                    return;

                  const raw =
                    nums[
                      i
                    ] ??
                    nums[
                      nums.length -
                        1
                    ] ??
                    "";

                  const n =
                    Number(
                      String(
                        raw
                      )
                        .replace(
                          /,/g,
                          ""
                        )
                        .replace(
                          /--|---/g,
                          ""
                        )
                    ) || 0;

                  records.push(
                    {
                      id: `${d.date}-${cleanName}-${records.length}`,
                      date:
                        d.date,
                      name:
                        cleanName,
                      project:
                        text,
                      type:
                        (
                          ts[
                            i
                          ] ??
                          ts[
                            ts.length -
                              1
                          ] ??
                          ""
                        ).trim(),
                      worked:
                        n,
                      link:
                        (
                          ls[
                            i
                          ] ??
                          ls[
                            ls.length -
                              1
                          ] ??
                          ""
                        ).trim()
                    }
                  );
                }
              );
            }
          }

          const workDates =
            [
              ...new Set(
                records
                  .map(
                    x =>
                      x.date
                  )
                  .filter(
                    Boolean
                  )
              )
            ].sort();

          const today =
            workDates[
              workDates.length -
                1
            ] ||
            new Date()
              .toISOString()
              .slice(
                0,
                10
              );

          const todayRows =
            records.filter(
              x =>
                x.date ===
                  today &&
                ![
                  "Saturday",
                  "Sunday",
                  "On Leave"
                ].includes(
                  x.project
                )
            );

          const team =
            names.map(
              (
                name,
                i
              ) => {
                const person =
                  todayRows.filter(
                    x =>
                      x.name ===
                      name
                  );

                const projectNames =
                  [
                    ...new Set(
                      person
                        .map(
                          x =>
                            getConfiguredProjectName(
                              x.project
                            )
                        )
                        .filter(
                          Boolean
                        )
                    )
                  ];

                const target =
                  projectNames.reduce(
                    (
                      sum,
                      project
                    ) =>
                      sum +
                      (projectTargets[
                        project
                      ] ||
                        0),
                    0
                  );

                const completed =
                  person.reduce(
                    (
                      s,
                      x
                    ) =>
                      s +
                      (Number(
                        x.worked
                      ) ||
                        0),
                    0
                  );

                const reviewed =
                  person
                    .filter(
                      x =>
                        /review/i.test(
                          x.type
                        )
                    )
                    .reduce(
                      (
                        s,
                        x
                      ) =>
                        s +
                        (Number(
                          x.worked
                        ) ||
                          0),
                      0
                    );

                const leave =
                  records.some(
                    x =>
                      x.date ===
                        today &&
                      x.name ===
                        name &&
                      x.project ===
                        "On Leave"
                  );

                return {
                  id:
                    1000 +
                    i,
                  name,
                  role:
                    "Annotator",
                  target,
                  completed,
                  reviewed,
                  errors: 0,
                  status:
                    leave
                      ? "Away"
                      : "Active"
                };
              }
            );

          const projectMap =
            {};

          Object.entries(
            projectTargets
          ).forEach(
            ([
              name,
              target
            ]) => {
              projectMap[
                name
              ] = {
                completed: 0,
                target
              };
            }
          );

          todayRows.forEach(
            x => {
              const configuredName =
                getConfiguredProjectName(
                  x.project
                );

              if (
                !configuredName
              )
                return;

              if (
                !projectMap[
                  configuredName
                ]
              ) {
                projectMap[
                  configuredName
                ] = {
                  completed: 0,
                  target:
                    projectTargets[
                      configuredName
                    ] ||
                    0
                };
              }

              projectMap[
                configuredName
              ].completed +=
                Number(
                  x.worked
                ) || 0;
            }
          );

          const projects =
            Object.entries(
              projectMap
            ).map(
              (
                [
                  name,
                  v
                ],
                i
              ) => {
                const target =
                  Number(
                    v.target
                  ) || 0;

                const completed =
                  Math.max(
                    0,
                    Number(
                      v.completed
                    ) || 0
                  );

                const remaining =
                  Math.max(
                    0,
                    target -
                      completed
                  );

                return {
                  id:
                    2000 +
                    i,
                  name,
                  target,
                  total:
                    target,
                  completed,
                  remaining,
                  status:
                    getProjectStatus(
                      target,
                      completed,
                      remaining
                    ),
                  deadline:
                    ""
                };
              }
            );

          await update({
            ...data,
            team,
            projects,
            sheetRecords:
              records,
            sheetFile:
              file.name,
            sheetLastSync:
              new Date().toISOString()
          });

          setPreview(
            records.slice(
              0,
              25
            )
          );

          notify(
            `Imported ${records.length} work records`
          );
        } catch (err) {
          console.error(
            err
          );

          alert(
            "Could not read this Excel file. Please use .xlsx format."
          );
        } finally {
          setBusy(
            false
          );
        }
      };

    reader.readAsArrayBuffer(
      file
    );
  }

  return (
    <Page
      title="Sheet Import"
      subtitle="Import your existing Daily Effort Sheet and use it as the dashboard data source."
    >
      <div className="grid two">
        <Panel title="Import Excel / Google Sheets export">
          <div className="import-box">
            <Upload
              size={28}
            />

            <h3>
              {busy
                ? "Importing..."
                : "Upload your .xlsx file"}
            </h3>

            <p>
              Use Google Sheets →
              File → Download →
              Microsoft Excel
              (.xlsx).
            </p>

            <label className="primary upload-label">
              <Upload
                size={17}
              />

              Choose Excel file

              <input
                type="file"
                accept=".xlsx,.xls"
                hidden
                onChange={e =>
                  e.target.files?.[0] &&
                  parseWorkbook(
                    e.target.files[0]
                  )
                }
              />
            </label>

            {fileName && (
              <div className="import-success">
                <CheckCircle2
                  size={17}
                />

                <span>
                  <b>
                    {
                      fileName
                    }
                  </b>

                  <small>
                    Last imported:{" "}
                    {data.sheetLastSync
                      ? new Date(
                          data.sheetLastSync
                        ).toLocaleString()
                      : "just now"}
                  </small>
                </span>
              </div>
            )}
          </div>
        </Panel>

        <Panel title="Detected sheet structure">
          <div className="structure-list">
            <div>
              <b>
                Name
              </b>

              <span>
                Team member
              </span>
            </div>

            <div>
              <b>
                Project
              </b>

              <span>
                Project worked on
              </span>
            </div>

            <div>
              <b>
                Annotation/Review
              </b>

              <span>
                Work type
              </span>
            </div>

            <div>
              <b>
                Total images worked
              </b>

              <span>
                Daily output
              </span>
            </div>

            <div>
              <b>
                Link to the range
              </b>

              <span>
                Work/range link
              </span>
            </div>
          </div>
        </Panel>
      </div>

      <Panel
        title={`Imported records ${
          data.sheetRecords?.length
            ? `(${data.sheetRecords.length})`
            : ""
        }`}
      >
        {preview.length ===
          0 &&
        !data.sheetRecords?.length ? (
          <p className="muted">
            Upload the Excel file
            to preview and sync
            your daily work.
          </p>
        ) : (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>
                    Date
                  </th>

                  <th>
                    Name
                  </th>

                  <th>
                    Project
                  </th>

                  <th>
                    Type
                  </th>

                  <th>
                    Images worked
                  </th>
                </tr>
              </thead>

              <tbody>
                {(preview.length
                  ? preview
                  : data.sheetRecords.slice(
                      0,
                      25
                    )
                ).map(
                  (
                    x,
                    i
                  ) => (
                    <tr
                      key={
                        x.id ||
                        i
                      }
                    >
                      <td>
                        {
                          x.date
                        }
                      </td>

                      <td>
                        <b>
                          {
                            x.name
                          }
                        </b>
                      </td>

                      <td>
                        {
                          x.project
                        }
                      </td>

                      <td>
                        {
                          x.type
                        }
                      </td>

                      <td>
                        <b>
                          {Number(
                            x.worked ||
                              0
                          ).toLocaleString()}
                        </b>
                      </td>
                    </tr>
                  )
                )}
              </tbody>
            </table>
          </div>
        )}
      </Panel>
    </Page>
  );
}

/* =========================================================
   SETTINGS
========================================================= */

function SettingsPage({
  exportData,
  importData,
  role,
  online
}) {
  return (
    <Page
      title="Settings"
      subtitle="Manage dashboard data and backups."
    >
      <Panel title="Data management">
        <div className="settings-row">
          <div>
            <b>
              Storage
            </b>

            <p>
              Dashboard data is
              stored online in
              Supabase.
            </p>
          </div>

          <span className="status active">
            <i />
            {online
              ? "Online"
              : "Local"}
          </span>
        </div>

        <div className="settings-row">
          <div>
            <b>
              Export backup
            </b>

            <p>
              Download all team,
              project and issue
              data as JSON.
            </p>
          </div>

          <button
            className="secondary"
            onClick={
              exportData
            }
          >
            <Download
              size={17}
            />

            Export
          </button>
        </div>

        {canEdit(role) && (
          <div className="settings-row">
            <div>
              <b>
                Import backup
              </b>

              <p>
                Restore a previously
                exported dashboard
                backup.
              </p>
            </div>

            <label className="secondary">
              <Upload
                size={17}
              />

              Import

              <input
                type="file"
                accept=".json"
                onChange={
                  importData
                }
                hidden
              />
            </label>
          </div>
        )}
      </Panel>
    </Page>
  );
}

/* =========================================================
   PAGE
========================================================= */

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
          <p className="eyebrow">
            TEAM OPERATIONS
          </p>

          <h1>
            {title}
          </h1>

          <p className="sub">
            {subtitle}
          </p>
        </div>

        {action && (
          <button
            className="primary"
            onClick={
              onAction
            }
          >
            <Plus
              size={18}
            />

            {action.replace(
              "+ ",
              ""
            )}
          </button>
        )}
      </div>

      {children}
    </div>
  );
}

/* =========================================================
   MODAL
========================================================= */

function Modal({
  type,
  onClose,
  onSubmit
}) {
  const labels = {
    team: [
      "Add team member",
      "Name",
      "Role",
      "Target",
      "Completed"
    ],

    project: [
      "Add project",
      "Project name",
      "Daily target",
      "Completed today",
      "Status"
    ],

    issue: [
      "Log issue",
      "Issue type",
      "Project",
      "Owner",
      "Severity"
    ]
  };

  const l =
    labels[type];

  return (
    <div className="modal-bg">
      <div className="modal">
        <div className="modal-head">
          <div>
            <p className="eyebrow">
              NEW RECORD
            </p>

            <h2>
              {l[0]}
            </h2>
          </div>

          <button
            className="icon-btn"
            onClick={
              onClose
            }
          >
            <X />
          </button>
        </div>

        <form
          onSubmit={
            onSubmit
          }
        >
          <label>
            {l[1]}

            <input
              name={
                type ===
                "team"
                  ? "name"
                  : type ===
                    "project"
                  ? "name"
                  : "type"
              }
              required
            />
          </label>

          {type ===
            "team" && (
            <>
              <label>
                {l[2]}

                <select name="role">
                  <option>
                    Annotator
                  </option>

                  <option>
                    Reviewer
                  </option>

                  <option>
                    Team Lead
                  </option>
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

          {type ===
            "project" && (
            <>
              <label>
                {l[2]}

                <input
                  name="target"
                  type="number"
                  min="0"
                  required
                  placeholder="e.g. 100, 350, 600, 800, 1000"
                />
              </label>

              <label>
                {l[3]}

                <input
                  name="completed"
                  type="number"
                  min="0"
                  defaultValue="0"
                />
              </label>

              <label>
                {l[4]}

                <select name="status">
                  <option>
                    Pending
                  </option>

                  <option>
                    In Progress
                  </option>

                  <option>
                    Completed
                  </option>
                </select>
              </label>

              <label>
                Deadline

                <input
                  name="deadline"
                  type="date"
                />
              </label>
            </>
          )}

          {type ===
            "issue" && (
            <>
              <label>
                {l[2]}

                <input
                  name="project"
                  required
                />
              </label>

              <label>
                {l[3]}

                <input
                  name="owner"
                  required
                />
              </label>

              <label>
                {l[4]}

                <select name="severity">
                  <option>
                    High
                  </option>

                  <option>
                    Medium
                  </option>

                  <option>
                    Low
                  </option>
                </select>
              </label>
            </>
          )}

          <div className="modal-actions">
            <button
              type="button"
              className="secondary"
              onClick={
                onClose
              }
            >
              Cancel
            </button>

            <button
              className="primary"
              type="submit"
            >
              <Save
                size={17}
              />

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

createRoot(
  document.getElementById(
    "root"
  )
).render(
  <App />
);
