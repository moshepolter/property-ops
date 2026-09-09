
import React, { useState, useEffect, useMemo, useRef } from "react";
import Papa from "papaparse";
// PDF upload support for RIS reports. Requires: npm install pdfjs-dist
// The worker is loaded from a CDN so it works with any bundler — no local worker file needed.
import * as pdfjsLib from "pdfjs-dist";
pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.worker.min.js";
// Firebase for real persistence + login. Requires: npm install firebase
// Fill in firebaseConfig below with the values from your Firebase project settings.
import { initializeApp } from "firebase/app";
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { getFirestore, doc, getDoc, setDoc } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "YOUR_API_KEY",
  authDomain: "YOUR_PROJECT.firebaseapp.com",
  projectId: "YOUR_PROJECT_ID",
  storageBucket: "YOUR_PROJECT.appspot.com",
  messagingSenderId: "YOUR_SENDER_ID",
  appId: "YOUR_APP_ID",
};
const firebaseApp = initializeApp(firebaseConfig);
const auth = getAuth(firebaseApp);
const db = getFirestore(firebaseApp);
import {
  Search, Building2, Users, Wrench, AlertTriangle, Gavel, HardHat,
  CalendarClock, ScrollText, MessageSquare, Archive as ArchiveIcon,
  Plus, X, Camera, Download, LayoutDashboard, ChevronDown, ChevronRight,
  Trash2, Pencil, Upload, Menu, Printer, CheckCircle2
} from "lucide-react";

/* ============================== constants ============================== */

const STORAGE_KEY = "pm-ops-data-v1";
const uid = () => Math.random().toString(36).slice(2, 10);
const todayISO = () => new Date().toISOString().slice(0, 10);

function filesToDataUrls(fileList) {
  const files = Array.from(fileList);
  return Promise.all(files.map(file => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve({ id: uid(), name: file.name, dataUrl: reader.result });
    reader.onerror = reject;
    reader.readAsDataURL(file);
  })));
}

const AGENCIES = ["HPD", "DSNY", "Other"];
const WO_STATUSES = ["Open", "In Progress", "Done"];
const WO_PRIORITIES = ["Routine", "Urgent", "Emergency"];
const RENT_STATUSES = ["Current", "Late", "In Arrears"];
const HPD_STATUSES = ["Open", "In Progress", "Certified", "Dismissed"];
const DSNY_STATUSES = ["Unpaid", "Disputing online", "Paid"];
const OTHER_STATUSES = ["Open", "In Progress", "Resolved", "Dismissed"];
const COURT_RESULTS = [
  "Pending", "Stipulation (payment plan)", "Case dismissed",
  "Adjourned / next date set", "Judgment for landlord", "Judgment for tenant",
  "Settled / withdrawn"
];
const DEFAULT_ATTORNEY_CHECKLIST = ["Lease", "Ledger", "Pre-suite notices"];
const CASE_STAGES = ["Filed", "Served", "Awaiting court date", "In court", "Awaiting decision", "Post-decision"];
const INSPECTION_TYPES = ["Boiler", "Elevator", "Fire Alarm", "Other recurring"];
const APPOINTMENT_TYPES = ["DOB Inspection", "Section 8 Inspection", "Other"];
const OTHER_AGENCY_PRESETS = ["DOB", "FDNY", "ECB", "DEP", "Con Edison"];
const ADD_NEW = "__add_new__";
const LOCAL_LAWS = [
  { key: "LL97", name: "Local Law 97 — Building Emissions" },
  { key: "LL87", name: "Local Law 87 — Energy Audit" },
  { key: "LL84", name: "Local Law 84 — Benchmarking" },
  { key: "LL11", name: "Local Law 11 / FISP — Facade" },
  { key: "LL126", name: "Local Law 126 — Water Tank" },
  { key: "LL55", name: "Local Law 55 — Indoor Allergen Hazards" },
  { key: "LL152", name: "Local Law 152 — Gas Piping Inspection" },
];

const emptyData = () => ({
  buildings: [], units: [], tenants: [], vendors: [], workOrders: [],
  violations: [], courtCases: [], inspections: [], appointments: [],
  localLaws: [], bossReminders: [], customAppointmentTypes: [], customInspectionTypes: [],
  customOtherAgencies: [], quickNotes: [], importHistory: [],
});

function violationClosedStatuses(agency) {
  return agency === "HPD" ? ["Certified", "Dismissed"] : agency === "DSNY" ? ["Paid"] : ["Resolved", "Dismissed"];
}
function isViolationClosed(v) {
  return violationClosedStatuses(v.agency).includes(v.status);
}

/* ============================== date helpers ============================== */

function daysUntil(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr + "T00:00:00");
  const now = new Date(); now.setHours(0, 0, 0, 0);
  return Math.round((d - now) / 86400000);
}
function flagFor(dateStr) {
  const d = daysUntil(dateStr);
  if (d === null) return null;
  if (d < 0) return "overdue";
  if (d <= 7) return "soon";
  return null;
}
function fmtDate(dateStr) {
  if (!dateStr) return "—";
  const d = new Date(dateStr + "T00:00:00");
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}
function addMonths(dateStr, n) {
  const d = new Date(dateStr + "T00:00:00");
  d.setMonth(d.getMonth() + n);
  return d.toISOString().slice(0, 10);
}
function isInFollowUpWindow(dateStr) {
  const d = daysUntil(dateStr);
  if (d === null) return false;
  return d >= -7 && d <= 1;
}
function icsFor(title, dateStr, notes, timeFrom, timeTo) {
  const dt = (dateStr || todayISO()).replace(/-/g, "");
  let lines;
  if (timeFrom) {
    const startStr = dt + "T" + timeFrom.replace(":", "") + "00";
    const endStr = dt + "T" + (timeTo || timeFrom).replace(":", "") + "00";
    lines = [
      "BEGIN:VCALENDAR", "VERSION:2.0", "BEGIN:VEVENT",
      `UID:${uid()}@propertyops`, `DTSTART:${startStr}`, `DTEND:${endStr}`,
      `SUMMARY:${title.replace(/\n/g, " ")}`, `DESCRIPTION:${(notes || "").replace(/\n/g, " ")}`,
      "END:VEVENT", "END:VCALENDAR"
    ];
  } else {
    lines = [
      "BEGIN:VCALENDAR", "VERSION:2.0", "BEGIN:VEVENT",
      `UID:${uid()}@propertyops`, `DTSTART;VALUE=DATE:${dt}`, `DTEND;VALUE=DATE:${dt}`,
      `SUMMARY:${title.replace(/\n/g, " ")}`, `DESCRIPTION:${(notes || "").replace(/\n/g, " ")}`,
      "END:VEVENT", "END:VCALENDAR"
    ];
  }
  return "data:text/calendar;charset=utf8," + encodeURIComponent(lines.join("\r\n"));
}

function fmtTime(t) {
  if (!t) return "";
  const [h, m] = t.split(":").map(Number);
  const period = h >= 12 ? "PM" : "AM";
  const h12 = h % 12 === 0 ? 12 : h % 12;
  return `${h12}:${String(m).padStart(2, "0")} ${period}`;
}
const TIME_OPTIONS = (() => {
  const out = [];
  for (let h = 0; h < 24; h++) {
    for (let m = 0; m < 60; m += 15) {
      out.push(`${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`);
    }
  }
  return out;
})();

/* ============================== RIS report parsing ============================== */

const APT_RE = "[A-Z]{1,3}\\d{1,4}";
const SKIP_LINE_RE = /^\d{2}\/\d{2}\/\d{4}|FISCAL PERIOD|^Page:|PROP #|TELEPHONE\/EMAIL LIST|^APT:|BUILDING DIRECTORY|AGED ARREARS|^LEGAL:|^\*\s*-\s*MOVED OUT|^TOTALS:/i;

function cleanLines(text) {
  return text.split("\n").map(l => l.trim()).filter(l => l && !SKIP_LINE_RE.test(l));
}

function money(n) {
  return "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Aged Arrears: "A1 AUDREY LYNN MELENDEZ UNKNO 1698.80 1698.80 3794.00 7191.60"
function parseArrearsText(text) {
  const lineRe = new RegExp(`^(${APT_RE})(\\*)?\\s+(.+?)\\s+UNKNO\\s+([\\d,]+\\.\\d{2})\\s+([\\d,]+\\.\\d{2})\\s+([\\d,]+\\.\\d{2})\\s+([\\d,]+\\.\\d{2})\\s*$`);
  const out = [];
  for (const line of cleanLines(text)) {
    const m = line.match(lineRe);
    if (!m) continue;
    const [, apt, moved, rawName, d1, d2, d3, total] = m;
    const name = rawName.replace(/^N-\d+\s+/, "").trim();
    const totalNum = parseFloat(total.replace(/,/g, ""));
    const bucket61 = parseFloat(d3.replace(/,/g, ""));
    let status = "Current";
    if (totalNum > 0) status = bucket61 > 0 ? "In Arrears" : "Late";
    out.push({ apt, movedOut: !!moved, name, balance: money(totalNum), status });
  }
  return out;
}

// Telephone/Email list: header line "A1 AUDREY LYNN MELENDEZ" then indented "CELL - ...", "EMAIL ADDRESS - ..."
function parseContactsText(text) {
  const headerRe = new RegExp(`^(${APT_RE})\\s+(.+)$`);
  const detailRe = /^(CELL|EMAIL ADDRESS|HOME|WORK|OTHER|FAX)\s*-\s*(.+)$/i;
  const out = [];
  let current = null;
  for (const line of cleanLines(text)) {
    const detailMatch = line.match(detailRe);
    if (detailMatch && current) {
      const [, label, value] = detailMatch;
      if (/EMAIL/i.test(label) && !current.email) current.email = value.trim().split(/\s+/)[0];
      if (/CELL/i.test(label) && !current.phone) current.phone = value.trim().match(/[\d()+\-.\s]{7,}/)?.[0]?.trim() || value.trim();
      continue;
    }
    const headerMatch = line.match(headerRe);
    if (headerMatch) {
      if (current) out.push(current);
      current = { apt: headerMatch[1], name: headerMatch[2].trim(), phone: "", email: "" };
    }
  }
  if (current) out.push(current);
  return out.filter(c => c.phone || c.email);
}

// Building Directory: two apt/name pairs per line, best-effort split
function parseDirectoryText(text) {
  const pairRe = new RegExp(`(${APT_RE})\\s+(.*?)(?=(?:${APT_RE})\\s|$)`, "g");
  const out = [];
  for (const line of cleanLines(text)) {
    let m;
    pairRe.lastIndex = 0;
    while ((m = pairRe.exec(line))) {
      const apt = m[1];
      let name = m[2].trim().replace(/,$/, "");
      if (!name) continue;
      let needsReview = false;
      if (name.includes("&")) {
        // Multi-occupant unit — "LAST, FIRST & LAST2, FIRST2" is too ambiguous to
        // safely reorder without risking mixing up who belongs to which name.
        // Leave as-extracted and flag for manual review rather than guess wrong.
        needsReview = true;
      } else if (name.includes(",")) {
        // "LAST, FIRST MIDDLE" -> "First Middle Last"
        const [last, rest] = name.split(",").map(s => s.trim());
        if (last && rest) name = `${rest} ${last}`;
      }
      out.push({ apt, name, needsReview });
    }
  }
  return out;
}

// Extracts text from an uploaded PDF, reconstructing line breaks from item
// y-position since pdf.js returns a flat list of positioned text fragments.
async function extractPdfText(file) {
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  let fullText = "";
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const content = await page.getTextContent();
    const items = content.items.map(it => ({
      str: it.str, x: it.transform[4], y: Math.round(it.transform[5]),
    }));
    items.sort((a, b) => (b.y - a.y) || (a.x - b.x));
    let lines = [];
    let currentY = null, currentLine = [];
    for (const it of items) {
      if (currentY === null || Math.abs(it.y - currentY) > 3) {
        if (currentLine.length) lines.push(currentLine);
        currentLine = [it];
        currentY = it.y;
      } else {
        currentLine.push(it);
      }
    }
    if (currentLine.length) lines.push(currentLine);
    fullText += lines.map(line => line.sort((a, b) => a.x - b.x).map(i => i.str).join(" ")).join("\n") + "\n";
  }
  return fullText;
}

/* ============================== small UI atoms ============================== */

function Flag({ date, label }) {
  const f = flagFor(date);
  if (!f) return <span className="pill pill-muted">{fmtDate(date)}</span>;
  return <span className={`pill ${f === "overdue" ? "pill-danger" : "pill-warn"}`}>{fmtDate(date)}{label ? ` · ${label}` : ""}</span>;
}

function Field({ label, children }) {
  return (
    <label className="field">
      <span className="field-label">{label}</span>
      {children}
    </label>
  );
}

function IconBtn({ onClick, title, children, danger }) {
  return (
    <button className={`icon-btn ${danger ? "icon-btn-danger" : ""}`} onClick={onClick} title={title} type="button">
      {children}
    </button>
  );
}

function Section({ icon, title, count, action, children }) {
  return (
    <div className="section">
      <div className="section-head">
        <div className="section-title">
          {icon}
          <h2>{title}</h2>
          {count != null && <span className="count-badge">{count}</span>}
        </div>
        {action}
      </div>
      {children}
    </div>
  );
}

function EmptyState({ text }) {
  return <div className="empty-state">{text}</div>;
}

function PrintButton({ label }) {
  return (
    <button className="btn-ghost no-print" onClick={() => window.print()} title={`Print ${label}`}>
      <Printer size={14} /> Print
    </button>
  );
}

function TypeSelectWithAdd({ value, options, onChange, onAddType }) {
  const [adding, setAdding] = useState(false);
  const [newType, setNewType] = useState("");

  const save = () => {
    const trimmed = newType.trim();
    if (!trimmed) return;
    onAddType(trimmed);
    onChange(trimmed);
    setNewType(""); setAdding(false);
  };

  if (adding) {
    return (
      <div className="inline-form" style={{ margin: 0 }}>
        <input
          autoFocus placeholder="New type name" value={newType}
          onChange={e => setNewType(e.target.value)}
          onKeyDown={e => e.key === "Enter" && save()}
        />
        <button className="btn-primary" type="button" onClick={save}>Save</button>
        <button className="btn-ghost" type="button" onClick={() => { setAdding(false); setNewType(""); }}>Cancel</button>
      </div>
    );
  }

  return (
    <select value={value} onChange={e => {
      if (e.target.value === ADD_NEW) setAdding(true);
      else onChange(e.target.value);
    }}>
      {options.map(t => <option key={t}>{t}</option>)}
      <option value={ADD_NEW}>+ Add new type…</option>
    </select>
  );
}

function PhotoUploader({ photos, onAdd, onRemove }) {
  const ref = useRef(null);
  return (
    <div className="photo-uploader">
      <div className="photo-grid">
        {(photos || []).map(p => (
          <div className="photo-thumb" key={p.id}>
            <img src={p.dataUrl} alt={p.name} />
            <button className="photo-remove" onClick={() => onRemove(p.id)} title="Remove"><X size={12} /></button>
          </div>
        ))}
        <button className="photo-add" onClick={() => ref.current.click()} type="button">
          <Camera size={16} />
        </button>
      </div>
      <input
        ref={ref} type="file" accept="image/*" multiple hidden
        onChange={async (e) => {
          if (!e.target.files.length) return;
          const uploaded = await filesToDataUrls(e.target.files);
          onAdd(uploaded);
          e.target.value = "";
        }}
      />
    </div>
  );
}

function DocumentUploader({ documents, onAdd, onRemove }) {
  const ref = useRef(null);
  return (
    <div className="doc-uploader">
      {(documents || []).map(d => (
        <div className="doc-chip" key={d.id}>
          <a href={d.dataUrl} download={d.name} className="doc-chip-name" title={d.name}>{d.name}</a>
          <button className="doc-remove" onClick={() => onRemove(d.id)} title="Remove"><X size={12} /></button>
        </div>
      ))}
      <button className="btn-ghost" type="button" onClick={() => ref.current.click()}>
        <Upload size={14} /> Upload document
      </button>
      <input
        ref={ref} type="file" multiple hidden
        onChange={async (e) => {
          if (!e.target.files.length) return;
          const uploaded = await filesToDataUrls(e.target.files);
          onAdd(uploaded);
          e.target.value = "";
        }}
      />
    </div>
  );
}

/* ============================== login ============================== */

function LoginScreen({ onSignedIn }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    setBusy(true);
    try {
      await signInWithEmailAndPassword(auth, email, password);
    } catch (err) {
      setError("Couldn't sign in — check your email and password.");
    }
    setBusy(false);
  };

  return (
    <div className="login-shell">
      <form className="login-card" onSubmit={submit}>
        <div className="brand-mark" style={{ marginBottom: 14 }}>PO</div>
        <h1 className="page-title" style={{ marginBottom: 4 }}>Property Ops</h1>
        <p className="hint" style={{ marginBottom: 16 }}>Sign in to continue.</p>
        <Field label="Email"><input type="email" value={email} onChange={e => setEmail(e.target.value)} required /></Field>
        <Field label="Password"><input type="password" value={password} onChange={e => setPassword(e.target.value)} required /></Field>
        {error && <div className="hint" style={{ color: "var(--danger)" }}>{error}</div>}
        <button className="btn-primary" type="submit" disabled={busy} style={{ marginTop: 10, width: "100%", justifyContent: "center" }}>
          {busy ? "Signing in…" : "Sign in"}
        </button>
      </form>
      <Styles />
    </div>
  );
}

/* ============================== app ============================== */

export default function PropertyOpsApp() {
  const [user, setUser] = useState(undefined); // undefined = checking, null = signed out
  const [data, setData] = useState(emptyData());
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [query, setQuery] = useState("");
  const [navOpen, setNavOpen] = useState(false);
  const saveTimer = useRef(null);

  useEffect(() => onAuthStateChanged(auth, u => setUser(u || null)), []);

  useEffect(() => {
    if (!user) return;
    setLoaded(false);
    (async () => {
      try {
        const snap = await getDoc(doc(db, "users", user.uid, "appData", "main"));
        if (snap.exists()) setData({ ...emptyData(), ...snap.data() });
        else setData(emptyData());
      } catch (e) { console.error("load failed", e); }
      setLoaded(true);
    })();
  }, [user]);

  useEffect(() => {
    if (!loaded || !user) return;
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try { await setDoc(doc(db, "users", user.uid, "appData", "main"), data); }
      catch (e) { console.error("save failed", e); }
    }, 400);
    return () => clearTimeout(saveTimer.current);
  }, [data, loaded, user]);

  // generic collection ops
  const add = (col, item) => setData(d => ({ ...d, [col]: [...d[col], { id: uid(), ...item }] }));
  const update = (col, id, patch) => setData(d => ({ ...d, [col]: d[col].map(x => x.id === id ? { ...x, ...patch } : x) }));
  const remove = (col, id) => setData(d => ({ ...d, [col]: d[col].filter(x => x.id !== id) }));

  const buildingName = (id) => data.buildings.find(b => b.id === id)?.address || "—";
  const unitLabel = (id) => { const u = data.units.find(x => x.id === id); return u ? u.unitNumber : "—"; };
  const vendorName = (id) => data.vendors.find(v => v.id === id)?.name || "—";
  const tenantName = (id) => data.tenants.find(t => t.id === id)?.name || "—";

  const tabs = [
    { key: "dashboard", label: "Dashboard", icon: <LayoutDashboard size={16} /> },
    { key: "buildings", label: "Buildings", icon: <Building2 size={16} /> },
    { key: "rent", label: "Rent Collection", icon: <Users size={16} /> },
    { key: "workorders", label: "Work Orders", icon: <Wrench size={16} /> },
    { key: "violations", label: "Violations", icon: <AlertTriangle size={16} /> },
    { key: "vendors", label: "Vendors", icon: <HardHat size={16} /> },
    { key: "court", label: "Court Cases", icon: <Gavel size={16} /> },
    { key: "inspections", label: "Appointments", icon: <CalendarClock size={16} /> },
    { key: "laws", label: "NYC Local Laws", icon: <ScrollText size={16} /> },
    { key: "reminders", label: "Boss Reminders", icon: <MessageSquare size={16} /> },
    { key: "quicknotes", label: "Quick Notes", icon: <Pencil size={16} /> },
  ];

  const searchResults = useMemo(() => {
    if (!query.trim()) return null;
    const q = query.toLowerCase();
    return {
      tenants: data.tenants.filter(t => t.name?.toLowerCase().includes(q) || buildingName(t.buildingId).toLowerCase().includes(q)),
      violations: data.violations.filter(v => (v.violationNumber || "").toLowerCase().includes(q) || (v.description || "").toLowerCase().includes(q) || buildingName(v.buildingId).toLowerCase().includes(q)),
      courtCases: data.courtCases.filter(c => (c.caseNumber || "").toLowerCase().includes(q) || tenantName(c.tenantId).toLowerCase().includes(q)),
      buildings: data.buildings.filter(b => (b.address || "").toLowerCase().includes(q)),
    };
  }, [query, data]);

  if (user === undefined) return <div className="app-shell"><div className="loading">Loading…</div><Styles /></div>;
  if (user === null) return <LoginScreen />;
  if (!loaded) return <div className="app-shell"><div className="loading">Loading your ops board…</div><Styles /></div>;

  return (
    <div className="app-shell">
      <header className="topbar no-print">
        <button className="menu-btn" onClick={() => setNavOpen(o => !o)} title="Menu">
          <Menu size={18} />
        </button>
        <div className="brand">
          <div className="brand-mark">PO</div>
          <div>
            <div className="brand-title">Property Ops</div>
            <div className="brand-sub">{data.buildings.length} buildings tracked</div>
          </div>
        </div>
        <div className="search-wrap">
          <Search size={16} className="search-icon" />
          <input
            className="search-input"
            placeholder="Search tenants, violations, cases, addresses…"
            value={query}
            onChange={e => setQuery(e.target.value)}
          />
          {query && <button className="search-clear" onClick={() => setQuery("")}><X size={14} /></button>}
        </div>
        <button className="btn-ghost no-print" title="Sign out" onClick={() => signOut(auth)} style={{ color: "#fff", borderColor: "rgba(255,255,255,0.3)" }}>
          Sign out
        </button>
      </header>

      {searchResults ? (
        <SearchResults results={searchResults} buildingName={buildingName} onClose={() => setQuery("")} />
      ) : (
        <div className="layout">
          {navOpen && <div className="nav-scrim no-print" onClick={() => setNavOpen(false)} />}
          <nav className={`sidenav no-print ${navOpen ? "sidenav-open" : ""}`}>
            {tabs.map(t => (
              <button key={t.key} className={`nav-item ${tab === t.key ? "nav-item-active" : ""}`} onClick={() => { setTab(t.key); setNavOpen(false); }}>
                {t.icon}<span>{t.label}</span>
              </button>
            ))}
          </nav>
          <main className="content">
            {tab === "dashboard" && <Dashboard data={data} buildingName={buildingName} tenantName={tenantName} setTab={setTab} />}
            {tab === "buildings" && <BuildingsTab data={data} add={add} update={update} remove={remove} setData={setData} buildingName={buildingName} />}
            {tab === "rent" && <RentTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} unitLabel={unitLabel} setData={setData} />}
            {tab === "workorders" && <WorkOrdersTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} vendorName={vendorName} />}
            {tab === "violations" && <ViolationsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} vendorName={vendorName} setData={setData} />}
            {tab === "vendors" && <VendorsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} />}
            {tab === "court" && <CourtTab data={data} add={add} update={update} remove={remove} tenantName={tenantName} buildingName={buildingName} />}
            {tab === "inspections" && <AppointmentsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} setData={setData} />}
            {tab === "laws" && <LocalLawsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} />}
            {tab === "reminders" && <RemindersTab data={data} add={add} update={update} remove={remove} />}
            {tab === "quicknotes" && <QuickNotesTab data={data} add={add} update={update} remove={remove} />}
          </main>
        </div>
      )}
      <Styles />
    </div>
  );
}

/* ============================== search ============================== */

function SearchResults({ results, buildingName, onClose }) {
  const total = results.tenants.length + results.violations.length + results.courtCases.length + results.buildings.length;
  return (
    <div className="content" style={{ maxWidth: 900, margin: "0 auto" }}>
      <div className="search-results-head">
        <h2>{total} result{total === 1 ? "" : "s"}</h2>
        <button className="btn-ghost" onClick={onClose}>Clear search</button>
      </div>
      {results.buildings.length > 0 && (
        <Section icon={<Building2 size={16} />} title="Buildings" count={results.buildings.length}>
          {results.buildings.map(b => <div className="row" key={b.id}>{b.address}</div>)}
        </Section>
      )}
      {results.tenants.length > 0 && (
        <Section icon={<Users size={16} />} title="Tenants" count={results.tenants.length}>
          {results.tenants.map(t => (
            <div className="row" key={t.id}>
              <strong>{t.name}</strong> — {buildingName(t.buildingId)} · {t.status}
            </div>
          ))}
        </Section>
      )}
      {results.violations.length > 0 && (
        <Section icon={<AlertTriangle size={16} />} title="Violations" count={results.violations.length}>
          {results.violations.map(v => (
            <div className="row" key={v.id}>
              <strong>{v.agency}</strong> #{v.violationNumber} — {buildingName(v.buildingId)} · {v.status}
            </div>
          ))}
        </Section>
      )}
      {results.courtCases.length > 0 && (
        <Section icon={<Gavel size={16} />} title="Court Cases" count={results.courtCases.length}>
          {results.courtCases.map(c => (
            <div className="row" key={c.id}>
              #{c.caseNumber || "—"} — {c.result}
            </div>
          ))}
        </Section>
      )}
      {total === 0 && <EmptyState text="Nothing matches that search." />}
    </div>
  );
}

/* ============================== dashboard ============================== */

const TAB_LABELS = {
  rent: "Rent Collection", violations: "Violations", court: "Court Cases",
  inspections: "Appointments", quicknotes: "Quick Notes",
};

function AttentionPanel({ icon, label, items, tab, setTab, renderItem, itemKey }) {
  const [open, setOpen] = useState(false);
  if (items.length === 0) return null;
  return (
    <div className="followup-panel">
      <button className="followup-panel-head" onClick={() => setOpen(o => !o)}>
        {icon}
        <span className="attention-count">{items.length}</span>
        <span className="attention-label">{label}</span>
        {open ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
      </button>
      {open && (
        <div className="followup-panel-body">
          {items.slice(0, 8).map((it, i) => (
            <div className="followup-item" key={itemKey ? itemKey(it) : i}>{renderItem(it)}</div>
          ))}
          {items.length > 8 && <div className="hint" style={{ margin: "4px 0" }}>+ {items.length - 8} more</div>}
          <button className="btn-ghost" style={{ marginTop: 8 }} onClick={() => setTab(tab)}>
            View in {TAB_LABELS[tab] || tab}
          </button>
        </div>
      )}
    </div>
  );
}

function Dashboard({ data, buildingName, tenantName, setTab }) {
  const [rentPanelOpen, setRentPanelOpen] = useState(false);
  const [showAllOverdue, setShowAllOverdue] = useState(false);
  const [monthsToShow, setMonthsToShow] = useState(1);

  const overdueTenants = data.tenants.filter(t => t.status !== "Current");
  const violationItems = data.violations.filter(v => !isViolationClosed(v) && flagFor(v.cureDeadline));
  const courtItems = data.courtCases.filter(c => !c.archived && flagFor(c.nextCourtDate));
  const stipItems = data.courtCases.filter(c => !c.archived && c.result === "Stipulation (payment plan)" && flagFor(c.nextPaymentDue));
  const recurringItems = data.appointments.filter(a => !a.completed && a.recurring && flagFor(a.date));
  const appointmentItems = data.appointments.filter(a => !a.completed && !a.recurring && flagFor(a.date));
  const quickNoteItems = data.quickNotes || [];
  const followUpCutoff = addMonths(todayISO(), monthsToShow);
  const allFollowUps = data.tenants.filter(t => t.followUpDate).length;
  const followUpTenants = data.tenants
    .filter(t => t.followUpDate && t.followUpDate <= followUpCutoff)
    .slice()
    .sort((a, b) => a.followUpDate.localeCompare(b.followUpDate));
  const openViolations = data.violations.filter(v => !isViolationClosed(v)).length;
  const overdueShown = showAllOverdue ? overdueTenants : overdueTenants.slice(0, 5);

  const rentPanelCount = overdueTenants.length + allFollowUps;
  const totalAttention = rentPanelCount + violationItems.length + courtItems.length + stipItems.length + recurringItems.length + appointmentItems.length + quickNoteItems.length;

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Dashboard</h1>
        <PrintButton label="Dashboard" />
      </div>

      {totalAttention === 0 ? (
        <div className="all-clear"><CheckCircle2 size={18} /> Nothing needs attention right now.</div>
      ) : (
        <>
          {rentPanelCount > 0 && (
            <div className="followup-panel">
              <button className="followup-panel-head" onClick={() => setRentPanelOpen(o => !o)}>
                <Users size={18} className="attention-icon" style={{ color: "var(--danger)" }} />
                <span className="attention-count">{rentPanelCount}</span>
                <span className="attention-label">Tenants not current on rent / Tenants to follow up</span>
                {rentPanelOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
              </button>
              {rentPanelOpen && (
                <div className="followup-panel-body">
                  {overdueTenants.length > 0 && (
                    <div className="followup-subsection">
                      <div className="followup-subsection-title">Not current on rent ({overdueTenants.length})</div>
                      {overdueShown.map(t => (
                        <div className="followup-item" key={t.id}>
                          <span className={`pill ${t.status === "Late" ? "pill-warn" : "pill-danger"}`}>{t.status}</span>
                          <div className="followup-item-main">
                            <div className="followup-item-name">{t.name} <span className="row-muted">— {buildingName(t.buildingId)}</span></div>
                            {t.balance && <div className="followup-item-note">Balance: {t.balance}</div>}
                          </div>
                        </div>
                      ))}
                      {overdueTenants.length > 5 && (
                        <button className="btn-ghost" style={{ marginTop: 6 }} onClick={() => setShowAllOverdue(s => !s)}>
                          {showAllOverdue ? "Show fewer" : `Show all ${overdueTenants.length}`}
                        </button>
                      )}
                    </div>
                  )}
                  {allFollowUps > 0 && (
                    <div className="followup-subsection">
                      <div className="followup-subsection-title">Tenants to follow up ({allFollowUps})</div>
                      {followUpTenants.map(t => (
                        <div className={`followup-item ${isInFollowUpWindow(t.followUpDate) ? "followup-item-due" : ""}`} key={t.id}>
                          <span className="pill pill-muted">{fmtDate(t.followUpDate)}</span>
                          <div className="followup-item-main">
                            <div className="followup-item-name">{t.name} <span className="row-muted">— {buildingName(t.buildingId)}</span></div>
                            {t.followUpNote && <div className="followup-item-note">{t.followUpNote}</div>}
                          </div>
                        </div>
                      ))}
                      {followUpTenants.length < allFollowUps && (
                        <button className="btn-ghost" style={{ marginTop: 6 }} onClick={() => setMonthsToShow(m => m + 1)}>
                          Show next month
                        </button>
                      )}
                    </div>
                  )}
                  <button className="btn-ghost" style={{ marginTop: 6 }} onClick={() => setTab("rent")}>View in Rent Collection</button>
                </div>
              )}
            </div>
          )}

          <AttentionPanel
            icon={<AlertTriangle size={18} className="attention-icon" style={{ color: "var(--danger)" }} />}
            label="Violation deadlines due or overdue" items={violationItems} tab="violations" setTab={setTab}
            itemKey={v => v.id}
            renderItem={v => (
              <>
                <Flag date={v.cureDeadline} />
                <div className="followup-item-main">
                  <div className="followup-item-name">#{v.violationNumber} · {v.agency} <span className="row-muted">— {buildingName(v.buildingId)}</span></div>
                </div>
              </>
            )}
          />

          <AttentionPanel
            icon={<Gavel size={18} className="attention-icon" style={{ color: "var(--danger)" }} />}
            label="Court dates due or overdue" items={courtItems} tab="court" setTab={setTab}
            itemKey={c => c.id}
            renderItem={c => (
              <>
                <Flag date={c.nextCourtDate} />
                <div className="followup-item-main">
                  <div className="followup-item-name">{tenantName(c.tenantId)} {c.caseNumber && `· Docket #${c.caseNumber}`}</div>
                </div>
              </>
            )}
          />

          <AttentionPanel
            icon={<Gavel size={18} className="attention-icon" style={{ color: "var(--danger)" }} />}
            label="Stipulation payments due" items={stipItems} tab="court" setTab={setTab}
            itemKey={c => c.id}
            renderItem={c => (
              <>
                <Flag date={c.nextPaymentDue} />
                <div className="followup-item-main">
                  <div className="followup-item-name">{tenantName(c.tenantId)}</div>
                  {c.stipulationTerms && <div className="followup-item-note">{c.stipulationTerms}</div>}
                </div>
              </>
            )}
          />

          <AttentionPanel
            icon={<CalendarClock size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label="Recurring inspections due or overdue" items={recurringItems} tab="inspections" setTab={setTab}
            itemKey={a => a.id}
            renderItem={a => (
              <>
                <Flag date={a.date} />
                <div className="followup-item-main">
                  <div className="followup-item-name">{a.type} <span className="row-muted">— {buildingName(a.buildingId)}</span></div>
                </div>
              </>
            )}
          />

          <AttentionPanel
            icon={<CalendarClock size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label="Appointments coming up" items={appointmentItems} tab="inspections" setTab={setTab}
            itemKey={a => a.id}
            renderItem={a => (
              <>
                <Flag date={a.date} />
                <div className="followup-item-main">
                  <div className="followup-item-name">{a.type} <span className="row-muted">— {buildingName(a.buildingId)}</span></div>
                </div>
              </>
            )}
          />

          <AttentionPanel
            icon={<Pencil size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label="Quick notes to organize" items={quickNoteItems} tab="quicknotes" setTab={setTab}
            itemKey={n => n.id}
            renderItem={n => (
              <div className="followup-item-main">
                <div className="followup-item-name">{n.text}</div>
                <div className="followup-item-note">{fmtDate(n.date)}</div>
              </div>
            )}
          />
        </>
      )}

      <h2 className="section-heading">By building</h2>
      {data.buildings.length === 0 ? (
        <EmptyState text="Add your buildings to see a per-building breakdown." />
      ) : (
        <div className="dash-grid">
          {data.buildings.map(b => {
            const bViolations = data.violations.filter(v => v.buildingId === b.id && !isViolationClosed(v)).length;
            const bTenantsLate = data.tenants.filter(t => t.buildingId === b.id && t.status !== "Current").length;
            const bWO = data.workOrders.filter(w => w.buildingId === b.id && w.status !== "Done").length;
            const bCourt = data.courtCases.filter(c => c.buildingId === b.id && !c.archived).length;
            const chips = [
              bViolations > 0 && { text: `${bViolations} open violations`, tone: "warn" },
              bTenantsLate > 0 && { text: `${bTenantsLate} tenants behind`, tone: "danger" },
              bWO > 0 && { text: `${bWO} open work orders`, tone: "muted" },
              bCourt > 0 && { text: `${bCourt} open cases`, tone: "danger" },
            ].filter(Boolean);
            return (
              <div className="dash-building-card" key={b.id}>
                <div className="dash-building-name">{b.address}</div>
                {chips.length === 0 ? (
                  <span className="dash-building-clear">All clear</span>
                ) : (
                  <div className="dash-building-chips">
                    {chips.map((c, i) => <span key={i} className={`pill pill-${c.tone}`}>{c.text}</span>)}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ============================== buildings ============================== */

function BuildingsTab({ data, add, update, remove, setData, buildingName }) {
  const [form, setForm] = useState(null);
  const [expanded, setExpanded] = useState(null);
  const [section, setSection] = useState("buildings");
  const fileRef = useRef(null);

  const submit = () => {
    if (!form.address) return;
    if (form.id) update("buildings", form.id, form);
    else add("buildings", form);
    setForm(null);
  };

  const handleCSV = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: (results) => {
        setData(d => {
          const next = { ...d, buildings: [...d.buildings], units: [...d.units], tenants: [...d.tenants] };
          results.data.forEach(row => {
            const address = row.address || row.Address || row.building || row.Building;
            if (!address) return;
            let building = next.buildings.find(b => b.address.toLowerCase() === address.toLowerCase());
            if (!building) {
              building = { id: uid(), address, notes: "" };
              next.buildings.push(building);
            }
            const unitNumber = row.unit || row.Unit || row.unitNumber;
            let unit = null;
            if (unitNumber) {
              unit = next.units.find(u => u.buildingId === building.id && u.unitNumber === unitNumber);
              if (!unit) {
                unit = { id: uid(), buildingId: building.id, unitNumber };
                next.units.push(unit);
              }
            }
            const tenantName = row.tenant || row.Tenant || row.tenantName;
            if (tenantName && unit) {
              const exists = next.tenants.find(t => t.unitId === unit.id && t.name === tenantName);
              if (!exists) {
                next.tenants.push({
                  id: uid(), buildingId: building.id, unitId: unit.id, name: tenantName,
                  phone: row.phone || "", email: row.email || "",
                  balance: row.balance || "", status: row.status || "Current",
                  notes: "", messageLog: [],
                });
              }
            }
          });
          return next;
        });
      }
    });
    e.target.value = "";
  };

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Buildings</h1>
        <div className="page-actions">
          <PrintButton label="Buildings" />
          <button className="btn-ghost" onClick={() => fileRef.current.click()}><Upload size={14} /> Import CSV</button>
          <input ref={fileRef} type="file" accept=".csv" hidden onChange={handleCSV} />
          <button className="btn-primary" onClick={() => setForm({ address: "", notes: "" })}><Plus size={14} /> Add building</button>
        </div>
      </div>
      <p className="hint">CSV columns recognized: address, unit, tenant, phone, email, balance, status.</p>

      <div className="filter-row">
        <button className={`chip ${section === "buildings" ? "chip-active" : ""}`} onClick={() => setSection("buildings")}>Buildings</button>
        <button className={`chip ${section === "import" ? "chip-active" : ""}`} onClick={() => setSection("import")}>Import Directory & Contacts</button>
      </div>

      {section === "import" ? (
        <ImportSection data={data} setData={setData} buildingName={buildingName} allowedTypes={["directory", "contacts"]} />
      ) : (
      <>
      {form && (
        <div className="form-panel">
          <Field label="Address"><input value={form.address} onChange={e => setForm({ ...form, address: e.target.value })} /></Field>
          <Field label="Notes"><textarea value={form.notes || ""} onChange={e => setForm({ ...form, notes: e.target.value })} /></Field>
          <div className="form-actions">
            <button className="btn-primary" onClick={submit}>Save</button>
            <button className="btn-ghost" onClick={() => setForm(null)}>Cancel</button>
          </div>
        </div>
      )}

      {data.buildings.length === 0 && <EmptyState text="No buildings yet — import a CSV or add one manually." />}
      {data.buildings.map(b => {
        const units = data.units.filter(u => u.buildingId === b.id);
        return (
          <div className="list-card" key={b.id}>
            <div className="list-card-head" onClick={() => setExpanded(expanded === b.id ? null : b.id)}>
              {expanded === b.id ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
              <div className="list-card-title">{b.address}</div>
              <span className="pill pill-muted">{units.length} units</span>
              <div className="spacer" />
              <IconBtn title="Edit" onClick={(e) => { e.stopPropagation(); setForm(b); }}><Pencil size={14} /></IconBtn>
              <IconBtn title="Delete" danger onClick={(e) => { e.stopPropagation(); remove("buildings", b.id); }}><Trash2 size={14} /></IconBtn>
            </div>
            {expanded === b.id && (
              <div className="list-card-body">
                {units.length === 0 && <div className="hint">No units added yet.</div>}
                {units.map(u => {
                  const tenants = data.tenants.filter(t => t.unitId === u.id);
                  return (
                    <div key={u.id} className="row">
                      Unit {u.unitNumber} — {tenants.map(t => t.name).join(", ") || "no tenant on file"}
                    </div>
                  );
                })}
                <button className="btn-ghost" style={{ marginTop: 8 }} onClick={() => add("units", { buildingId: b.id, unitNumber: "" })}>
                  <Plus size={14} /> Add unit
                </button>
              </div>
            )}
          </div>
        );
      })}
      </>
      )}
    </div>
  );
}

/* ============================== rent collection ============================== */

function RentTab({ data, add, update, remove, buildingName, unitLabel, setData }) {
  const [msgFor, setMsgFor] = useState(null);
  const [msgText, setMsgText] = useState("");
  const [noteFor, setNoteFor] = useState(null);
  const [noteText, setNoteText] = useState("");
  const [followFor, setFollowFor] = useState(null);
  const [followDate, setFollowDate] = useState("");
  const [followNote, setFollowNote] = useState("");
  const [statusFilter, setStatusFilter] = useState("All");
  const [buildingFilter, setBuildingFilter] = useState("All");
  const [section, setSection] = useState("sheet");

  const addRow = () => add("tenants", {
    buildingId: data.buildings[0]?.id || "", unitId: "", name: "", phone: "",
    email: "", balance: "", status: "Current", notes: [], messageLog: [],
  });


  const sendMsg = (tenantId) => {
    if (!msgText.trim()) return;
    const t = data.tenants.find(x => x.id === tenantId);
    update("tenants", tenantId, { messageLog: [...(t.messageLog || []), { date: todayISO(), text: msgText }] });
    setMsgText(""); setMsgFor(null);
  };

  const notesArr = (t) => Array.isArray(t.notes) ? t.notes : (t.notes ? [{ date: todayISO(), text: t.notes }] : []);

  const addNote = (tenantId) => {
    if (!noteText.trim()) return;
    const t = data.tenants.find(x => x.id === tenantId);
    update("tenants", tenantId, { notes: [...notesArr(t), { date: todayISO(), text: noteText }] });
    setNoteText(""); setNoteFor(null);
  };

  const openFollow = (t) => {
    setFollowDate(t.followUpDate || "");
    setFollowNote(t.followUpNote || "");
    setFollowFor(followFor === t.id ? null : t.id);
  };
  const saveFollow = (tenantId) => {
    update("tenants", tenantId, { followUpDate: followDate, followUpNote: followNote });
    setFollowFor(null);
  };

  const byBuilding = t => buildingFilter === "All" || t.buildingId === buildingFilter;
  const visibleTenants = (statusFilter === "Follow-ups"
    ? data.tenants.filter(t => t.followUpDate).slice().sort((a, b) => a.followUpDate.localeCompare(b.followUpDate))
    : data.tenants.filter(t => statusFilter === "All" || t.status === statusFilter)
  ).filter(byBuilding);

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Rent Collection</h1>
        <div className="page-actions">
          <PrintButton label="Rent Collection" />
          <button className="btn-primary" onClick={addRow}><Plus size={14} /> Add row</button>
        </div>
      </div>

      <div className="filter-row">
        <button className={`chip ${section === "sheet" ? "chip-active" : ""}`} onClick={() => setSection("sheet")}>Sheet</button>
        <button className={`chip ${section === "import" ? "chip-active" : ""}`} onClick={() => setSection("import")}>Import Aged Arrears</button>
      </div>

      {section === "import" ? (
        <ImportSection data={data} setData={setData} buildingName={buildingName} allowedTypes={["arrears"]} />
      ) : (
      <>
      <div className="filter-row">
        {["All", ...RENT_STATUSES].map(s => (
          <button key={s} className={`chip ${statusFilter === s ? "chip-active" : ""}`} onClick={() => setStatusFilter(s)}>{s}</button>
        ))}
        <button className={`chip ${statusFilter === "Follow-ups" ? "chip-active" : ""}`} onClick={() => setStatusFilter("Follow-ups")}>Follow-ups</button>
      </div>
      <div className="filter-row">
        <button className={`chip ${buildingFilter === "All" ? "chip-active" : ""}`} onClick={() => setBuildingFilter("All")}>All buildings</button>
        {data.buildings.map(b => (
          <button key={b.id} className={`chip ${buildingFilter === b.id ? "chip-active" : ""}`} onClick={() => setBuildingFilter(b.id)}>{b.address}</button>
        ))}
      </div>

      {data.tenants.length === 0 ? (
        <EmptyState text="No tenants yet — import RIS data or add a row." />
      ) : (
        <div className="sheet-wrap">
          <table className="sheet">
            <thead>
              <tr>
                <th>Building</th><th>Unit</th><th>Tenant</th><th>Phone</th>
                <th className="sheet-col-balance">Balance</th><th>Status</th>
                <th className="sheet-col-notes">Notes</th><th>Follow-up</th><th></th>
              </tr>
            </thead>
            <tbody>
              {visibleTenants.length === 0 && (
                <tr><td colSpan={9}><div className="hint" style={{ padding: "10px 4px" }}>{statusFilter === "Follow-ups" ? "No tenants with a follow-up set." : "No tenants with this status."}</div></td></tr>
              )}
              {visibleTenants.map(t => (
                <React.Fragment key={t.id}>
                  <tr className={t.status !== "Current" ? "sheet-row-flag" : ""}>
                    <td>
                      <select className="sheet-input" value={t.buildingId} onChange={e => update("tenants", t.id, { buildingId: e.target.value, unitId: "" })}>
                        <option value="">—</option>
                        {data.buildings.map(b => <option key={b.id} value={b.id}>{b.address}</option>)}
                      </select>
                    </td>
                    <td>
                      <select className="sheet-input" value={t.unitId} onChange={e => update("tenants", t.id, { unitId: e.target.value })}>
                        <option value="">—</option>
                        {data.units.filter(u => u.buildingId === t.buildingId).map(u => <option key={u.id} value={u.id}>{u.unitNumber}</option>)}
                      </select>
                    </td>
                    <td><input className="sheet-input" value={t.name} onChange={e => update("tenants", t.id, { name: e.target.value })} /></td>
                    <td><input className="sheet-input" value={t.phone} onChange={e => update("tenants", t.id, { phone: e.target.value })} /></td>
                    <td className="sheet-col-balance"><input className="sheet-input" value={t.balance} onChange={e => update("tenants", t.id, { balance: e.target.value })} /></td>
                    <td>
                      <select className={`sheet-input sheet-status-${t.status === "Current" ? "ok" : t.status === "Late" ? "warn" : "danger"}`} value={t.status} onChange={e => update("tenants", t.id, { status: e.target.value })}>
                        {RENT_STATUSES.map(s => <option key={s}>{s}</option>)}
                      </select>
                    </td>
                    <td className="sheet-col-notes">
                      <button className="sheet-note-preview" onClick={() => setNoteFor(noteFor === t.id ? null : t.id)} title="Click to view/add notes">
                        {notesArr(t).length === 0 ? <span className="row-muted">— add note —</span> : (
                          notesArr(t).slice(-2).reverse().map((n, i) => (
                            <span key={i} className="sheet-note-line">
                              <span className="sheet-note-date">{fmtDate(n.date)}</span>
                              <span className="sheet-note-text">{n.text}</span>
                            </span>
                          ))
                        )}
                      </button>
                    </td>
                    <td>
                      <button className="sheet-follow-btn" onClick={() => openFollow(t)} title="Set a follow-up date">
                        <CalendarClock size={14} />
                        {t.followUpDate && <span className="sheet-follow-date">{fmtDate(t.followUpDate)}</span>}
                      </button>
                    </td>
                    <td className="sheet-actions">
                      <IconBtn title="Notes" onClick={() => setNoteFor(noteFor === t.id ? null : t.id)}><Pencil size={14} /></IconBtn>
                      <IconBtn title="Message log" onClick={() => setMsgFor(msgFor === t.id ? null : t.id)}><MessageSquare size={14} /></IconBtn>
                      <IconBtn title="Delete" danger onClick={() => remove("tenants", t.id)}><Trash2 size={14} /></IconBtn>
                    </td>
                  </tr>
                  {followFor === t.id && (
                    <tr className="sheet-expand-row">
                      <td colSpan={9}>
                        <strong>Follow-up</strong>
                        <div className="inline-form">
                          <input type="date" value={followDate} onChange={e => setFollowDate(e.target.value)} />
                          <input placeholder="What's this follow-up about?" value={followNote} onChange={e => setFollowNote(e.target.value)} style={{ flex: 1 }} />
                          <button className="btn-primary" onClick={() => saveFollow(t.id)}>Save</button>
                          {t.followUpDate && <button className="btn-ghost" onClick={() => { setFollowDate(""); setFollowNote(""); update("tenants", t.id, { followUpDate: "", followUpNote: "" }); setFollowFor(null); }}>Clear</button>}
                        </div>
                      </td>
                    </tr>
                  )}
                  {noteFor === t.id && (
                    <tr className="sheet-expand-row">
                      <td colSpan={9}>
                        <strong>Notes</strong>
                        <div className="inline-form">
                          <input placeholder="Payment plans, disputes, promises to pay…" value={noteText} onChange={e => setNoteText(e.target.value)} onKeyDown={e => e.key === "Enter" && addNote(t.id)} />
                          <button className="btn-primary" onClick={() => addNote(t.id)}>Add note</button>
                        </div>
                        {notesArr(t).length === 0 && <div className="hint">No notes yet.</div>}
                        {notesArr(t).slice().reverse().map((n, i) => (
                          <div key={i} className="row row-muted">{fmtDate(n.date)} — {n.text}</div>
                        ))}
                      </td>
                    </tr>
                  )}
                  {msgFor === t.id && (
                    <tr className="sheet-expand-row">
                      <td colSpan={9}>
                        <strong>Message log</strong>
                        <div className="inline-form">
                          <input placeholder="What did you send / say?" value={msgText} onChange={e => setMsgText(e.target.value)} onKeyDown={e => e.key === "Enter" && sendMsg(t.id)} />
                          <button className="btn-primary" onClick={() => sendMsg(t.id)}>Log it</button>
                        </div>
                        {(t.messageLog || []).length === 0 && <div className="hint">No messages logged yet.</div>}
                        {(t.messageLog || []).slice().reverse().map((m, i) => (
                          <div key={i} className="row row-muted">{fmtDate(m.date)} — {m.text}</div>
                        ))}
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
      )}
      </>
      )}
    </div>
  );
}

/* ============================== RIS import ============================== */

const IMPORT_TYPES = [
  { key: "arrears", label: "Aged Arrears", hint: "Updates balance and status. Paste the full text of the Aged Arrears report." },
  { key: "directory", label: "Building Directory", hint: "Updates tenant names. Two-column layout — review carefully before confirming." },
  { key: "contacts", label: "Telephone / Email List", hint: "Updates phone and email." },
];

function buildImportDiff(type, parsedEntries, data, buildingId) {
  const unitsForBuilding = data.units.filter(u => u.buildingId === buildingId);
  const tenantByUnit = unitId => data.tenants.find(t => t.unitId === unitId);
  const changes = [];

  parsedEntries.forEach(entry => {
    const unit = unitsForBuilding.find(u => (u.unitNumber || "").toUpperCase() === entry.apt.toUpperCase());
    let fields = {};
    if (type === "arrears") fields = { balance: entry.balance, status: entry.status };
    if (type === "directory") fields = { name: entry.name };
    if (type === "contacts") {
      if (entry.phone) fields.phone = entry.phone;
      if (entry.email) fields.email = entry.email;
    }
    if (unit) {
      const tenant = tenantByUnit(unit.id);
      if (tenant) {
        const diffFields = {};
        Object.entries(fields).forEach(([k, v]) => { if (v && tenant[k] !== v) diffFields[k] = v; });
        if (Object.keys(diffFields).length > 0) {
          changes.push({
            apt: entry.apt, unitId: unit.id, tenantId: tenant.id, isNew: false,
            fields: diffFields, before: Object.fromEntries(Object.keys(diffFields).map(k => [k, tenant[k] || "—"])),
            name: tenant.name || entry.name, movedOut: entry.movedOut, needsReview: entry.needsReview,
          });
        }
      } else {
        changes.push({ apt: entry.apt, unitId: unit.id, tenantId: null, isNew: true, fields: { ...fields, name: fields.name || entry.name || "" }, name: entry.name, movedOut: entry.movedOut, needsReview: entry.needsReview });
      }
    } else {
      changes.push({ apt: entry.apt, unitId: null, tenantId: null, isNew: true, newUnit: true, fields: { ...fields, name: fields.name || entry.name || "" }, name: entry.name, movedOut: entry.movedOut, needsReview: entry.needsReview });
    }
  });

  let missing = [];
  if (type === "arrears") {
    const parsedApts = new Set(parsedEntries.map(e => e.apt.toUpperCase()));
    missing = unitsForBuilding
      .filter(u => !parsedApts.has((u.unitNumber || "").toUpperCase()))
      .map(u => ({ apt: u.unitNumber, name: tenantByUnit(u.id)?.name || "(no tenant on file)" }));
  }
  return { changes, missing };
}

function ImportSection({ data, setData, buildingName, allowedTypes }) {
  const availableTypes = IMPORT_TYPES.filter(t => !allowedTypes || allowedTypes.includes(t.key));
  const [buildingId, setBuildingId] = useState(data.buildings[0]?.id || "");
  const [type, setType] = useState(availableTypes[0]?.key || "arrears");
  const [rawText, setRawText] = useState("");
  const [preview, setPreview] = useState(null);
  const [expandedEntry, setExpandedEntry] = useState(null);
  const [pdfStatus, setPdfStatus] = useState(null); // null | "reading" | "error"
  const [pdfError, setPdfError] = useState("");
  const pdfInputRef = useRef(null);

  const runParse = (text, forType, bId) => {
    if (!text.trim() || !bId) return;
    const parser = forType === "arrears" ? parseArrearsText : forType === "directory" ? parseDirectoryText : parseContactsText;
    const entries = parser(text);
    const diff = buildImportDiff(forType, entries, data, bId);
    setPreview({ ...diff, parsedCount: entries.length });
  };

  const parse = () => runParse(rawText, type, buildingId);

  const handlePdfUpload = async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file || !buildingId) return;
    setPdfStatus("reading");
    setPdfError("");
    setPreview(null);
    try {
      const text = await extractPdfText(file);
      setRawText(text);
      setPdfStatus(null);
      runParse(text, type, buildingId);
    } catch (err) {
      setPdfStatus("error");
      setPdfError("Couldn't read that PDF automatically — paste the text below instead (open the PDF, Ctrl/Cmd+A, Ctrl/Cmd+C, then paste).");
    }
  };

  const confirm = () => {
    if (!preview) return;
    setData(d => {
      const next = { ...d, units: [...d.units], tenants: [...d.tenants], importHistory: [...d.importHistory] };
      preview.changes.forEach(ch => {
        let unitId = ch.unitId;
        if (ch.newUnit) {
          const unit = { id: uid(), buildingId, unitNumber: ch.apt };
          next.units.push(unit);
          unitId = unit.id;
        }
        if (ch.isNew) {
          next.tenants.push({
            id: uid(), buildingId, unitId, name: ch.fields.name || ch.name || "",
            phone: ch.fields.phone || "", email: ch.fields.email || "",
            balance: ch.fields.balance || "", status: ch.fields.status || "Current",
            notes: [], messageLog: [],
          });
        } else {
          next.tenants = next.tenants.map(t => t.id === ch.tenantId ? { ...t, ...ch.fields } : t);
        }
      });
      next.importHistory = [
        {
          id: uid(), date: todayISO(), buildingId, type,
          updated: preview.changes.filter(c => !c.isNew).length,
          added: preview.changes.filter(c => c.isNew).length,
          missing: preview.missing.length,
          details: preview.changes.map(c => ({ apt: c.apt, name: c.name, isNew: c.isNew, fields: c.fields })),
          missingDetails: preview.missing,
        },
        ...next.importHistory,
      ];
      return next;
    });
    setPreview(null);
    setRawText("");
  };

  const typeInfo = IMPORT_TYPES.find(t => t.key === type);
  const history = (data.importHistory || []).filter(h => !allowedTypes || allowedTypes.includes(h.type));

  return (
    <div>
      <div className="form-panel" style={{ marginTop: 0 }}>
        <Field label="Building">
          <select value={buildingId} onChange={e => { setBuildingId(e.target.value); setPreview(null); }}>
            <option value="">—</option>
            {data.buildings.map(b => <option key={b.id} value={b.id}>{b.address}</option>)}
          </select>
        </Field>
        {availableTypes.length > 1 && (
          <Field label="Report type">
            <select value={type} onChange={e => { setType(e.target.value); setPreview(null); }}>
              {availableTypes.map(t => <option key={t.key} value={t.key}>{t.label}</option>)}
            </select>
          </Field>
        )}
        <div className="field" style={{ gridColumn: "1 / -1" }}>
          <span className="field-label">Upload the PDF — it parses automatically</span>
          <div className="inline-form" style={{ margin: 0 }}>
            <button className="btn-primary" type="button" onClick={() => pdfInputRef.current.click()} disabled={!buildingId || pdfStatus === "reading"}>
              <Upload size={14} /> {pdfStatus === "reading" ? "Reading PDF…" : "Upload PDF"}
            </button>
            <input ref={pdfInputRef} type="file" accept="application/pdf" hidden onChange={handlePdfUpload} />
          </div>
          {pdfStatus === "error" && <div className="hint" style={{ color: "var(--danger)" }}>{pdfError}</div>}
        </div>
        <div className="field" style={{ gridColumn: "1 / -1" }}>
          <span className="field-label">Or paste the report text (open the PDF, select all, copy, paste here)</span>
          <textarea rows={8} value={rawText} onChange={e => { setRawText(e.target.value); setPreview(null); }} placeholder="Paste RIS report text here…" />
        </div>
        <div className="form-actions">
          <button className="btn-primary" onClick={parse} disabled={!rawText.trim() || !buildingId}>Parse & preview</button>
        </div>
      </div>
      <p className="hint">{typeInfo.hint}</p>

      {preview && (
        <div className="import-preview">
          <div className="import-preview-summary">
            {preview.changes.filter(c => c.isNew).length} new · {preview.changes.filter(c => !c.isNew).length} to update
            {type === "arrears" && ` · ${preview.missing.length} not in this file`}
            {" "}({preview.parsedCount} rows read)
          </div>
          {preview.changes.length === 0 && preview.missing.length === 0 && (
            <div className="hint">No changes found — everything already matches.</div>
          )}
          {preview.changes.map((c, i) => (
            <div className="import-row" key={i}>
              <span className="pill pill-muted">{c.apt}</span>
              <span className="import-row-name">{c.name}</span>
              {c.isNew && <span className="pill pill-warn">New</span>}
              {c.movedOut && <span className="pill pill-danger">Moved out (flagged)</span>}
              {c.needsReview && <span className="pill pill-warn">⚠ Review — shared unit, verify names</span>}
              <div className="import-row-fields">
                {Object.entries(c.fields).map(([k, v]) => (
                  <span key={k} className="import-field">{k}: {c.before?.[k] ? `${c.before[k]} → ` : ""}{String(v)}</span>
                ))}
              </div>
            </div>
          ))}
          {preview.missing.length > 0 && (
            <>
              <div className="row" style={{ marginTop: 10 }}><strong>Not found in this file (review — may have moved out)</strong></div>
              {preview.missing.map((m, i) => (
                <div className="import-row" key={i}>
                  <span className="pill pill-muted">{m.apt}</span>
                  <span className="import-row-name">{m.name}</span>
                </div>
              ))}
            </>
          )}
          <div className="form-actions" style={{ marginTop: 10 }}>
            <button className="btn-primary" onClick={confirm} disabled={preview.changes.length === 0}>Confirm import</button>
            <button className="btn-ghost" onClick={() => setPreview(null)}>Cancel</button>
          </div>
        </div>
      )}

      <h2 className="section-heading">Last Imported</h2>
      {history.length === 0 && <EmptyState text="No imports yet." />}
      {history.slice(0, 20).map(h => (
        <div className="list-card" key={h.id}>
          <div className="list-card-head" onClick={() => setExpandedEntry(expandedEntry === h.id ? null : h.id)}>
            {expandedEntry === h.id ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
            <div className="list-card-title">{IMPORT_TYPES.find(t => t.key === h.type)?.label || h.type}</div>
            <span className="pill pill-muted">{buildingName(h.buildingId)}</span>
            <span className="pill pill-muted">{fmtDate(h.date)}</span>
            <span className="pill pill-ok">{h.added} new</span>
            <span className="pill pill-warn">{h.updated} updated</span>
            {h.missing > 0 && <span className="pill pill-danger">{h.missing} missing</span>}
          </div>
          {expandedEntry === h.id && (
            <div className="list-card-body">
              {(h.details || []).map((d, i) => (
                <div key={i} className="row">
                  {d.apt} — {d.name} {d.isNew && "(new)"} {Object.entries(d.fields || {}).map(([k, v]) => `${k}: ${v}`).join(", ")}
                </div>
              ))}
              {(h.missingDetails || []).map((m, i) => (
                <div key={"m" + i} className="row row-muted">{m.apt} — {m.name} (not in this import)</div>
              ))}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

/* ============================== work orders ============================== */

function WorkOrdersTab({ data, add, update, remove, buildingName, vendorName }) {
  const [form, setForm] = useState(null);
  const [filter, setFilter] = useState("All");

  const submit = () => {
    if (!form.description) return;
    if (form.id) update("workOrders", form.id, form);
    else add("workOrders", form);
    setForm(null);
  };

  const list = data.workOrders.filter(w => filter === "All" || w.status === filter);

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Work Orders</h1>
        <div className="page-actions">
          <PrintButton label="Work Orders" />
          <button className="btn-primary" onClick={() => setForm({ buildingId: data.buildings[0]?.id || "", vendorId: "", description: "", status: "Open", priority: "Routine", dateOpened: todayISO() })}>
            <Plus size={14} /> Add work order
          </button>
        </div>
      </div>
      <div className="filter-row">
        {["All", ...WO_STATUSES].map(s => (
          <button key={s} className={`chip ${filter === s ? "chip-active" : ""}`} onClick={() => setFilter(s)}>{s}</button>
        ))}
      </div>

      {form && (
        <div className="form-panel">
          <Field label="Building">
            <select value={form.buildingId} onChange={e => setForm({ ...form, buildingId: e.target.value })}>
              <option value="">—</option>
              {data.buildings.map(b => <option key={b.id} value={b.id}>{b.address}</option>)}
            </select>
          </Field>
          <Field label="Vendor">
            <select value={form.vendorId} onChange={e => setForm({ ...form, vendorId: e.target.value })}>
              <option value="">Unassigned</option>
              {data.vendors.map(v => <option key={v.id} value={v.id}>{v.name}</option>)}
            </select>
          </Field>
          <Field label="Description"><textarea value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} /></Field>
          <Field label="Priority">
            <select value={form.priority} onChange={e => setForm({ ...form, priority: e.target.value })}>
              {WO_PRIORITIES.map(p => <option key={p}>{p}</option>)}
            </select>
          </Field>
          <Field label="Status">
            <select value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}>
              {WO_STATUSES.map(s => <option key={s}>{s}</option>)}
            </select>
          </Field>
          <div className="form-actions">
            <button className="btn-primary" onClick={submit}>Save</button>
            <button className="btn-ghost" onClick={() => setForm(null)}>Cancel</button>
          </div>
        </div>
      )}

      {list.length === 0 && <EmptyState text="No work orders here." />}
      {list.map(w => (
        <div className="list-card" key={w.id}>
          <div className="list-card-head">
            <div className="list-card-title">{w.description}</div>
            {w.priority !== "Routine" && <span className={`pill ${w.priority === "Emergency" ? "pill-danger" : "pill-warn"}`}>{w.priority}</span>}
            <span className={`pill ${w.status === "Done" ? "pill-ok" : "pill-muted"}`}>{w.status}</span>
            <span className="pill pill-muted">{buildingName(w.buildingId)}</span>
            {w.vendorId && <span className="pill pill-muted">{vendorName(w.vendorId)}</span>}
            <div className="spacer" />
            <IconBtn title="Edit" onClick={() => setForm(w)}><Pencil size={14} /></IconBtn>
            <IconBtn title="Delete" danger onClick={() => remove("workOrders", w.id)}><Trash2 size={14} /></IconBtn>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ============================== violations ============================== */

function ViolationsTab({ data, add, update, remove, buildingName, vendorName, setData }) {
  const [agency, setAgency] = useState("HPD");
  const [form, setForm] = useState(null);
  const [view, setView] = useState("active");
  const [noteFor, setNoteFor] = useState(null);
  const [noteText, setNoteText] = useState("");
  const fileRef = useRef(null);

  const statusesFor = (a) => a === "HPD" ? HPD_STATUSES : a === "DSNY" ? DSNY_STATUSES : OTHER_STATUSES;
  const isClosed = isViolationClosed;
  const otherAgencyOptions = [...OTHER_AGENCY_PRESETS, ...(data.customOtherAgencies || [])];
  const addCustomOtherAgency = (trimmed) => {
    if (!otherAgencyOptions.includes(trimmed)) {
      setData(d => ({ ...d, customOtherAgencies: [...(d.customOtherAgencies || []), trimmed] }));
    }
  };

  const blankForm = (a) => ({
    buildingId: data.buildings[0]?.id || "", violationNumber: "",
    class: "", description: "", dateIssued: todayISO(), cureDeadline: "",
    fineAmount: "", company: "", otherAgency: otherAgencyOptions[0] || "",
    status: statusesFor(a)[0], vendorId: "",
  });

  const submit = () => {
    if (!form.violationNumber) return;
    if (form.id) update("violations", form.id, form);
    else add("violations", { ...form, agency, photos: [], notes: [] });
    setForm(null);
  };

  const addNote = (v) => {
    if (!noteText.trim()) return;
    update("violations", v.id, { notes: [...(v.notes || []), { date: todayISO(), text: noteText }] });
    setNoteText(""); setNoteFor(null);
  };

  const handleCSV = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: (results) => {
        setData(d => {
          const next = { ...d, violations: [...d.violations] };
          results.data.forEach(row => {
            const address = row.address || row.building || row.Building;
            const building = d.buildings.find(b => (b.address || "").toLowerCase() === (address || "").toLowerCase());
            const a = row.agency || agency;
            next.violations.push({
              id: uid(), agency: a,
              violationNumber: row.violationNumber || row.number || row.id || "",
              buildingId: building ? building.id : "",
              class: row.class || "",
              description: row.description || "",
              dateIssued: row.dateIssued || "",
              cureDeadline: row.cureDeadline || row.deadline || "",
              fineAmount: row.fineAmount || row.fine || "",
              company: row.company || "",
              status: row.status || statusesFor(a)[0],
              vendorId: "", photos: [], notes: [],
            });
          });
          return next;
        });
      }
    });
    e.target.value = "";
  };

  let list = data.violations.filter(v => v.agency === agency && (view === "active" ? !isClosed(v) : isClosed(v)));
  if (view === "active") {
    list = [...list].sort((a, b) => {
      const da = daysUntil(a.cureDeadline); const db = daysUntil(b.cureDeadline);
      if (da === null && db === null) return 0;
      if (da === null) return 1;
      if (db === null) return -1;
      return da - db;
    });
  }

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Violations</h1>
        <div className="page-actions">
          <PrintButton label="Violations" />
          <button className="btn-ghost" onClick={() => fileRef.current.click()}><Upload size={14} /> Import CSV</button>
          <input ref={fileRef} type="file" accept=".csv" hidden onChange={handleCSV} />
          <button className="btn-primary" onClick={() => setForm(blankForm(agency))}>
            <Plus size={14} /> Add violation
          </button>
        </div>
      </div>
      <p className="hint">CSV columns recognized: agency, violationNumber, address, class, description, dateIssued, cureDeadline, fineAmount, company, status.</p>

      <div className="filter-row">
        {AGENCIES.map(a => (
          <button key={a} className={`chip ${agency === a ? "chip-active" : ""}`} onClick={() => setAgency(a)}>{a}</button>
        ))}
        <div className="spacer" />
        <button className={`chip ${view === "active" ? "chip-active" : ""}`} onClick={() => setView("active")}>Active / Due</button>
        <button className={`chip ${view === "closed" ? "chip-active" : ""}`} onClick={() => setView("closed")}>{agency === "DSNY" ? "Paid" : "Closed"}</button>
      </div>

      {form && (
        <div className="form-panel">
          <Field label="Building">
            <select value={form.buildingId} onChange={e => setForm({ ...form, buildingId: e.target.value })}>
              <option value="">—</option>
              {data.buildings.map(b => <option key={b.id} value={b.id}>{b.address}</option>)}
            </select>
          </Field>
          <Field label="Violation #"><input value={form.violationNumber} onChange={e => setForm({ ...form, violationNumber: e.target.value })} /></Field>

          {agency === "HPD" && (
            <>
              <Field label="Class"><input value={form.class} onChange={e => setForm({ ...form, class: e.target.value })} placeholder="A / B / C" /></Field>
              <Field label="Description"><textarea value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} /></Field>
              <Field label="Date issued"><input type="date" value={form.dateIssued} onChange={e => setForm({ ...form, dateIssued: e.target.value })} /></Field>
              <Field label="Certify / cure deadline"><input type="date" value={form.cureDeadline} onChange={e => setForm({ ...form, cureDeadline: e.target.value })} /></Field>
              <Field label="Vendor assigned">
                <select value={form.vendorId} onChange={e => setForm({ ...form, vendorId: e.target.value })}>
                  <option value="">Unassigned</option>
                  {data.vendors.map(v => <option key={v.id} value={v.id}>{v.name}</option>)}
                </select>
              </Field>
            </>
          )}

          {agency === "DSNY" && (
            <>
              <Field label="Fine amount"><input value={form.fineAmount} onChange={e => setForm({ ...form, fineAmount: e.target.value })} placeholder="$" /></Field>
              <Field label="Date issued"><input type="date" value={form.dateIssued} onChange={e => setForm({ ...form, dateIssued: e.target.value })} /></Field>
            </>
          )}

          {agency === "Other" && (
            <>
              <Field label="Violation agency">
                <TypeSelectWithAdd value={form.otherAgency} options={otherAgencyOptions} onChange={v => setForm({ ...form, otherAgency: v })} onAddType={addCustomOtherAgency} />
              </Field>
              <Field label="Violation company handling"><input value={form.company} onChange={e => setForm({ ...form, company: e.target.value })} /></Field>
              <Field label="Description"><textarea value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} /></Field>
              <Field label="Date issued"><input type="date" value={form.dateIssued} onChange={e => setForm({ ...form, dateIssued: e.target.value })} /></Field>
              <Field label="Correction deadline"><input type="date" value={form.cureDeadline} onChange={e => setForm({ ...form, cureDeadline: e.target.value })} /></Field>
            </>
          )}

          <Field label="Status">
            <select value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}>
              {statusesFor(agency).map(s => <option key={s}>{s}</option>)}
            </select>
          </Field>
          <div className="form-actions">
            <button className="btn-primary" onClick={submit}>Save</button>
            <button className="btn-ghost" onClick={() => setForm(null)}>Cancel</button>
          </div>
        </div>
      )}

      {list.length === 0 && <EmptyState text={`No ${agency} violations here.`} />}
      {list.map(v => {
        const flag = view === "active" ? flagFor(v.cureDeadline) : null;
        return (
          <div className={`list-card ${flag === "overdue" ? "list-card-danger" : flag === "soon" ? "list-card-warn" : ""}`} key={v.id}>
            <div className="list-card-head">
              <div className="list-card-title">#{v.violationNumber} {v.class && `· Class ${v.class}`}</div>
              <span className={`pill ${isClosed(v) ? "pill-ok" : "pill-muted"}`}>{v.status}</span>
              <span className="pill pill-muted">{buildingName(v.buildingId)}</span>
              {agency === "HPD" && v.vendorId && <span className="pill pill-muted">{vendorName(v.vendorId)}</span>}
              {agency === "DSNY" && v.fineAmount && <span className="pill pill-muted">{v.fineAmount}</span>}
              {agency === "Other" && v.otherAgency && <span className="pill pill-muted">{v.otherAgency}</span>}
              {agency === "Other" && v.company && <span className="pill pill-muted">{v.company}</span>}
              {agency !== "DSNY" && view === "active" && <Flag date={v.cureDeadline} />}
              <div className="spacer" />
              <IconBtn title="Edit" onClick={() => setForm(v)}><Pencil size={14} /></IconBtn>
              <IconBtn title="Delete" danger onClick={() => remove("violations", v.id)}><Trash2 size={14} /></IconBtn>
            </div>
            <div className="list-card-body">
              {v.description && <div className="row">{v.description}</div>}
              <PhotoUploader
                photos={v.photos}
                onAdd={(newPhotos) => update("violations", v.id, { photos: [...(v.photos || []), ...newPhotos] })}
                onRemove={(id) => update("violations", v.id, { photos: (v.photos || []).filter(p => p.id !== id) })}
              />
              <div className="row" style={{ marginTop: 8 }}>
                <strong>Notes</strong>
                <button className="btn-ghost" style={{ marginLeft: 8 }} onClick={() => setNoteFor(noteFor === v.id ? null : v.id)}><Plus size={14} /> Add note</button>
              </div>
              {noteFor === v.id && (
                <div className="inline-form">
                  <input placeholder="Update…" value={noteText} onChange={e => setNoteText(e.target.value)} />
                  <button className="btn-primary" onClick={() => addNote(v)}>Save</button>
                </div>
              )}
              {(v.notes || []).slice().reverse().map((n, i) => (
                <div key={i} className="row row-muted">{fmtDate(n.date)} — {n.text}</div>
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ============================== vendors ============================== */

function VendorsTab({ data, add, update, remove, buildingName }) {
  const [form, setForm] = useState(null);
  const [selected, setSelected] = useState(null);

  const submit = () => {
    if (!form.name) return;
    if (form.id) update("vendors", form.id, form);
    else add("vendors", form);
    setForm(null);
  };

  const openWO = (vid) => data.workOrders.filter(w => w.vendorId === vid && w.status !== "Done");
  const openViol = (vid) => data.violations.filter(v => v.vendorId === vid && !isViolationClosed(v));

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Vendors</h1>
        <div className="page-actions">
          <PrintButton label="Vendors" />
          <button className="btn-primary" onClick={() => setForm({ name: "", phone: "", email: "", specialty: "" })}><Plus size={14} /> Add vendor</button>
        </div>
      </div>

      {form && (
        <div className="form-panel">
          <Field label="Name"><input value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></Field>
          <Field label="Specialty"><input value={form.specialty} onChange={e => setForm({ ...form, specialty: e.target.value })} placeholder="Plumbing, electrical, super, etc." /></Field>
          <Field label="Phone"><input value={form.phone} onChange={e => setForm({ ...form, phone: e.target.value })} /></Field>
          <Field label="Email"><input value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} /></Field>
          <div className="form-actions">
            <button className="btn-primary" onClick={submit}>Save</button>
            <button className="btn-ghost" onClick={() => setForm(null)}>Cancel</button>
          </div>
        </div>
      )}

      {data.vendors.length === 0 && <EmptyState text="No vendors yet." />}
      {data.vendors.map(v => (
        <div className="list-card" key={v.id}>
          <div className="list-card-head" onClick={() => setSelected(selected === v.id ? null : v.id)}>
            {selected === v.id ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
            <div className="list-card-title">{v.name}</div>
            <span className="pill pill-muted">{v.specialty || "—"}</span>
            <span className="pill pill-muted">{openWO(v.id).length} open work orders</span>
            <span className="pill pill-muted">{openViol(v.id).length} open violations</span>
            <div className="spacer" />
            <IconBtn title="Edit" onClick={(e) => { e.stopPropagation(); setForm(v); }}><Pencil size={14} /></IconBtn>
            <IconBtn title="Delete" danger onClick={(e) => { e.stopPropagation(); remove("vendors", v.id); }}><Trash2 size={14} /></IconBtn>
          </div>
          {selected === v.id && (
            <div className="list-card-body">
              <strong>Open work orders</strong>
              {openWO(v.id).length === 0 && <div className="hint">None.</div>}
              {openWO(v.id).map(w => <div key={w.id} className="row">{w.description} — {buildingName(w.buildingId)}</div>)}
              <strong style={{ marginTop: 8, display: "block" }}>Open violations</strong>
              {openViol(v.id).length === 0 && <div className="hint">None.</div>}
              {openViol(v.id).map(vi => <div key={vi.id} className="row">#{vi.violationNumber} — {buildingName(vi.buildingId)} — <Flag date={vi.cureDeadline} /></div>)}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

/* ============================== court cases ============================== */

function CourtTab({ data, add, update, remove, tenantName, buildingName }) {
  const [form, setForm] = useState(null);
  const [view, setView] = useState("active");
  const [checklistText, setChecklistText] = useState({});
  const [detailsFor, setDetailsFor] = useState(null);

  const submit = () => {
    if (form.id) update("courtCases", form.id, form);
    else add("courtCases", {
      ...form, archived: false, documents: [],
      checklist: DEFAULT_ATTORNEY_CHECKLIST.map(label => ({ id: uid(), label, checked: false })),
    });
    setForm(null);
  };

  const list = data.courtCases.filter(c => view === "closed" ? c.archived : !c.archived);

  const toggleChecklistItem = (c, itemId) => {
    update("courtCases", c.id, {
      checklist: (c.checklist || []).map(i => i.id === itemId ? { ...i, checked: !i.checked } : i)
    });
  };
  const addChecklistItem = (c) => {
    const text = (checklistText[c.id] || "").trim();
    if (!text) return;
    update("courtCases", c.id, { checklist: [...(c.checklist || []), { id: uid(), label: text, checked: false }] });
    setChecklistText({ ...checklistText, [c.id]: "" });
  };
  const removeChecklistItem = (c, itemId) => {
    update("courtCases", c.id, { checklist: (c.checklist || []).filter(i => i.id !== itemId) });
  };

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Court Cases</h1>
        <div className="page-actions">
          <PrintButton label="Court Cases" />
          <button className="btn-primary" onClick={() => setForm({ tenantId: "", buildingId: data.buildings[0]?.id || "", caseNumber: "", nextCourtDate: "", result: "Pending", stage: CASE_STAGES[0], stipulationTerms: "", nextPaymentDue: "" })}>
            <Plus size={14} /> Add case
          </button>
        </div>
      </div>
      <div className="filter-row">
        <button className={`chip ${view === "active" ? "chip-active" : ""}`} onClick={() => setView("active")}>Active</button>
        <button className={`chip ${view === "closed" ? "chip-active" : ""}`} onClick={() => setView("closed")}>Closed / Archived</button>
      </div>

      {form && (
        <div className="form-panel">
          <Field label="Tenant">
            <select value={form.tenantId} onChange={e => setForm({ ...form, tenantId: e.target.value })}>
              <option value="">—</option>
              {data.tenants.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
            </select>
          </Field>
          <Field label="Building">
            <select value={form.buildingId} onChange={e => setForm({ ...form, buildingId: e.target.value })}>
              <option value="">—</option>
              {data.buildings.map(b => <option key={b.id} value={b.id}>{b.address}</option>)}
            </select>
          </Field>
          <Field label="Docket #"><input value={form.caseNumber} onChange={e => setForm({ ...form, caseNumber: e.target.value })} /></Field>
          <Field label="Where the case stands">
            <select value={form.stage || CASE_STAGES[0]} onChange={e => setForm({ ...form, stage: e.target.value })}>
              {CASE_STAGES.map(s => <option key={s}>{s}</option>)}
            </select>
          </Field>
          <Field label="Court date"><input type="date" value={form.nextCourtDate} onChange={e => setForm({ ...form, nextCourtDate: e.target.value })} /></Field>
          <Field label="Result">
            <select value={form.result} onChange={e => setForm({ ...form, result: e.target.value })}>
              {COURT_RESULTS.map(r => <option key={r}>{r}</option>)}
            </select>
          </Field>
          {form.result === "Adjourned / next date set" && (
            <Field label="New court date"><input type="date" value={form.nextCourtDate} onChange={e => setForm({ ...form, nextCourtDate: e.target.value })} /></Field>
          )}
          {form.result === "Stipulation (payment plan)" && (
            <>
              <Field label="Payment plan terms"><textarea value={form.stipulationTerms} onChange={e => setForm({ ...form, stipulationTerms: e.target.value })} placeholder="e.g. $200/month starting 10/1 on top of current rent, for 6 months" /></Field>
              <Field label="Next payment due date"><input type="date" value={form.nextPaymentDue} onChange={e => setForm({ ...form, nextPaymentDue: e.target.value })} /></Field>
            </>
          )}
          <div className="form-actions">
            <button className="btn-primary" onClick={submit}>Save</button>
            <button className="btn-ghost" onClick={() => setForm(null)}>Cancel</button>
          </div>
        </div>
      )}

      {list.length === 0 && <EmptyState text={view === "active" ? "No open court cases." : "Nothing closed yet."} />}
      {list.map(c => {
        const checkedCount = (c.checklist || []).filter(i => i.checked).length;
        return (
        <div className="list-card" key={c.id}>
          <div className="list-card-head">
            <div className="list-card-title">{tenantName(c.tenantId)} {c.caseNumber && `· Docket #${c.caseNumber}`}</div>
            {c.stage && <span className="pill pill-muted">{c.stage}</span>}
            <span className="pill pill-muted">{c.result}</span>
            <span className="pill pill-muted">{buildingName(c.buildingId)}</span>
            {c.nextCourtDate && !c.archived && <Flag date={c.nextCourtDate} label="court date" />}
            <div className="spacer" />
            <IconBtn title="Edit" onClick={() => setForm(c)}><Pencil size={14} /></IconBtn>
            {view === "closed"
              ? <IconBtn title="Restore to active" onClick={() => update("courtCases", c.id, { archived: false })}><ArchiveIcon size={14} /></IconBtn>
              : <IconBtn title="Move to closed" onClick={() => update("courtCases", c.id, { archived: true })}><ArchiveIcon size={14} /></IconBtn>}
            <IconBtn title="Delete" danger onClick={() => remove("courtCases", c.id)}><Trash2 size={14} /></IconBtn>
          </div>
          {c.result === "Stipulation (payment plan)" && (
            <div className="list-card-body">
              {c.stipulationTerms && <div className="row"><em>Terms:</em> {c.stipulationTerms}</div>}
              {c.nextPaymentDue && <div className="row">Next payment due: <Flag date={c.nextPaymentDue} /></div>}
            </div>
          )}
          <div className="list-card-body" style={{ paddingTop: c.result === "Stipulation (payment plan)" ? 0 : undefined }}>
            {detailsFor === c.id ? (
              <>
                <div className="row"><strong>To send attorney</strong></div>
                <div className="checklist">
                  {(c.checklist || []).map(item => (
                    <label className="checklist-item" key={item.id}>
                      <input type="checkbox" checked={item.checked} onChange={() => toggleChecklistItem(c, item.id)} />
                      <span className={item.checked ? "strike" : ""}>{item.label}</span>
                      <button className="checklist-remove" onClick={() => removeChecklistItem(c, item.id)} title="Remove"><X size={12} /></button>
                    </label>
                  ))}
                </div>
                <div className="inline-form">
                  <input placeholder="Add item…" value={checklistText[c.id] || ""} onChange={e => setChecklistText({ ...checklistText, [c.id]: e.target.value })} onKeyDown={e => e.key === "Enter" && addChecklistItem(c)} />
                  <button className="btn-ghost" onClick={() => addChecklistItem(c)}>Add</button>
                </div>
                <div className="row" style={{ marginTop: 6 }}><strong>Documents</strong></div>
                <DocumentUploader
                  documents={c.documents}
                  onAdd={(newDocs) => update("courtCases", c.id, { documents: [...(c.documents || []), ...newDocs] })}
                  onRemove={(id) => update("courtCases", c.id, { documents: (c.documents || []).filter(d => d.id !== id) })}
                />
                <button className="btn-primary" style={{ marginTop: 10 }} onClick={() => setDetailsFor(null)}>Save</button>
              </>
            ) : (
              <button className="btn-ghost" onClick={() => setDetailsFor(c.id)}>
                Checklist & documents {checkedCount > 0 && `(${checkedCount}/${(c.checklist || []).length} sent)`} {(c.documents || []).length > 0 && `· ${c.documents.length} doc${c.documents.length === 1 ? "" : "s"}`}
              </button>
            )}
          </div>
        </div>
      );})}
    </div>
  );
}

/* ============================== appointments (incl. recurring) ============================== */

function AppointmentsTab({ data, add, update, remove, buildingName, setData }) {
  const [form, setForm] = useState(null);
  const [view, setView] = useState("upcoming");
  const allTypes = [...APPOINTMENT_TYPES, ...(data.customAppointmentTypes || [])];

  const submit = () => {
    if (!form.type) return;
    if (form.id) update("appointments", form.id, form);
    else add("appointments", { ...form, completed: false });
    setForm(null);
  };

  const addCustomType = (trimmed) => {
    if (!allTypes.includes(trimmed)) {
      setData(d => ({ ...d, customAppointmentTypes: [...(d.customAppointmentTypes || []), trimmed] }));
    }
  };

  const blank = () => ({ buildingId: data.buildings[0]?.id || "", unitId: "", type: allTypes[0] || "Other", date: "", timeFrom: "", timeTo: "", notes: "", recurring: false, completed: false });

  const list = data.appointments.filter(a => view === "upcoming" ? !a.completed : a.completed);

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Appointments</h1>
        <div className="page-actions">
          <PrintButton label="Appointments" />
          <button className="btn-primary" onClick={() => setForm(blank())}><Plus size={14} /> Add appointment</button>
        </div>
      </div>
      <p className="hint">Inspections, DOB/Section 8 visits, and anything else on the calendar. Mark "Recurring" for things like boiler or fire alarm checks. Reminders flag at 7, 3, and 1 day before.</p>

      <div className="filter-row">
        <button className={`chip ${view === "upcoming" ? "chip-active" : ""}`} onClick={() => setView("upcoming")}>Upcoming</button>
        <button className={`chip ${view === "completed" ? "chip-active" : ""}`} onClick={() => setView("completed")}>Completed Inspections</button>
      </div>

      {form && (
        <div className="form-panel">
          <Field label="Building">
            <select value={form.buildingId} onChange={e => setForm({ ...form, buildingId: e.target.value, unitId: "" })}>
              <option value="">—</option>
              {data.buildings.map(b => <option key={b.id} value={b.id}>{b.address}</option>)}
            </select>
          </Field>
          <Field label="Unit (optional)">
            <select value={form.unitId} onChange={e => setForm({ ...form, unitId: e.target.value })}>
              <option value="">General / whole building</option>
              {data.units.filter(u => u.buildingId === form.buildingId).map(u => <option key={u.id} value={u.id}>Unit {u.unitNumber}</option>)}
            </select>
          </Field>
          <Field label="Type">
            <TypeSelectWithAdd value={form.type} options={allTypes} onChange={v => setForm({ ...form, type: v })} onAddType={addCustomType} />
          </Field>
          <Field label="Date"><input type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} /></Field>
          <Field label="Time from">
            <select value={form.timeFrom} onChange={e => setForm({ ...form, timeFrom: e.target.value })}>
              <option value="">—</option>
              {TIME_OPTIONS.map(t => <option key={t} value={t}>{fmtTime(t)}</option>)}
            </select>
          </Field>
          <Field label="Time to">
            <select value={form.timeTo} onChange={e => setForm({ ...form, timeTo: e.target.value })}>
              <option value="">—</option>
              {TIME_OPTIONS.map(t => <option key={t} value={t}>{fmtTime(t)}</option>)}
            </select>
          </Field>
          <Field label="Recurring?">
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "var(--ink)" }}>
              <input type="checkbox" checked={!!form.recurring} onChange={e => setForm({ ...form, recurring: e.target.checked })} />
              Repeats (boiler, elevator, fire alarm, etc.)
            </label>
          </Field>
          <Field label="Notes"><textarea value={form.notes} onChange={e => setForm({ ...form, notes: e.target.value })} /></Field>
          <div className="form-actions">
            <button className="btn-primary" onClick={submit}>Save</button>
            <button className="btn-ghost" onClick={() => setForm(null)}>Cancel</button>
          </div>
        </div>
      )}

      {list.length === 0 && <EmptyState text={view === "upcoming" ? "No appointments scheduled." : "Nothing completed yet."} />}
      {list.map(a => {
        const d = daysUntil(a.date);
        const reminderHit = view === "upcoming" && d !== null && [0, 1, 3, 7].includes(d);
        return (
          <div className="list-card" key={a.id}>
            <div className="list-card-head">
              <div className="list-card-title">{a.type}</div>
              {a.recurring && <span className="pill pill-muted">Recurring</span>}
              <span className="pill pill-muted">{buildingName(a.buildingId)}</span>
              {a.unitId && <span className="pill pill-muted">Unit {data.units.find(u => u.id === a.unitId)?.unitNumber || "—"}</span>}
              {view === "upcoming" && <Flag date={a.date} />}
              {view === "completed" && <span className="pill pill-muted">{fmtDate(a.date)}</span>}
              {(a.timeFrom || a.timeTo) && (
                <span className="pill pill-muted">{fmtTime(a.timeFrom)}{a.timeTo ? ` – ${fmtTime(a.timeTo)}` : ""}</span>
              )}
              {reminderHit && <span className="pill pill-warn">Reminder</span>}
              <div className="spacer" />
              {view === "upcoming" && (
                <a className="btn-ghost" href={icsFor(`${a.type} — ${buildingName(a.buildingId)}`, a.date, a.notes, a.timeFrom, a.timeTo)} download={`${a.type.replace(/\s/g, "-")}.ics`}>
                  <Download size={14} /> Add to calendar
                </a>
              )}
              <IconBtn title={view === "upcoming" ? "Mark completed" : "Move back to upcoming"} onClick={() => update("appointments", a.id, { completed: !a.completed })}>
                <CheckCircle2 size={14} />
              </IconBtn>
              <IconBtn title="Edit" onClick={() => setForm(a)}><Pencil size={14} /></IconBtn>
              <IconBtn title="Delete" danger onClick={() => remove("appointments", a.id)}><Trash2 size={14} /></IconBtn>
            </div>
            {a.notes && <div className="list-card-body"><div className="row">{a.notes}</div></div>}
          </div>
        );
      })}
    </div>
  );
}

/* ============================== NYC local laws ============================== */

function LocalLawsTab({ data, add, update, remove, buildingName }) {
  const [buildingId, setBuildingId] = useState(data.buildings[0]?.id || "");

  useEffect(() => { if (!buildingId && data.buildings[0]) setBuildingId(data.buildings[0].id); }, [data.buildings]);

  const rowsFor = (bid) => LOCAL_LAWS.map(law => {
    const existing = data.localLaws.find(l => l.buildingId === bid && l.lawKey === law.key);
    return existing || { buildingId: bid, lawKey: law.key, deadline: "", status: "Not started" };
  });

  const saveRow = (row) => {
    const existing = data.localLaws.find(l => l.buildingId === row.buildingId && l.lawKey === row.lawKey);
    if (existing) update("localLaws", existing.id, row);
    else add("localLaws", row);
  };

  if (data.buildings.length === 0) return <div><h1 className="page-title">NYC Local Laws</h1><EmptyState text="Add a building first." /></div>;

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">NYC Local Laws</h1>
        <PrintButton label="NYC Local Laws" />
      </div>
      <div className="filter-row">
        {data.buildings.map(b => (
          <button key={b.id} className={`chip ${buildingId === b.id ? "chip-active" : ""}`} onClick={() => setBuildingId(b.id)}>{b.address}</button>
        ))}
      </div>
      <div className="law-table">
        {rowsFor(buildingId).map(row => (
          <div className="law-row" key={row.lawKey}>
            <div className="law-name">{LOCAL_LAWS.find(l => l.key === row.lawKey).name}</div>
            <input type="date" value={row.deadline} onChange={e => saveRow({ ...row, deadline: e.target.value })} />
            <select value={row.status} onChange={e => saveRow({ ...row, status: e.target.value })}>
              <option>Not started</option>
              <option>In progress</option>
              <option>Filed / compliant</option>
              <option>Not applicable</option>
            </select>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================== boss reminders ============================== */

function RemindersTab({ data, add, update, remove }) {
  const [text, setText] = useState("");

  const submit = () => {
    if (!text.trim()) return;
    add("bossReminders", { text, dateRaised: todayISO(), status: "Open" });
    setText("");
  };

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Boss Reminders</h1>
        <PrintButton label="Boss Reminders" />
      </div>
      <div className="inline-form">
        <input placeholder="Something to bring up with your boss…" value={text} onChange={e => setText(e.target.value)} onKeyDown={e => e.key === "Enter" && submit()} />
        <button className="btn-primary" onClick={submit}>Add</button>
      </div>
      {data.bossReminders.length === 0 && <EmptyState text="Nothing on the list right now." />}
      {data.bossReminders.map(r => (
        <div className="list-card" key={r.id}>
          <div className="list-card-head">
            <input type="checkbox" checked={r.status === "Done"} onChange={e => update("bossReminders", r.id, { status: e.target.checked ? "Done" : "Open" })} />
            <div className={`list-card-title ${r.status === "Done" ? "strike" : ""}`}>{r.text}</div>
            <span className="pill pill-muted">{fmtDate(r.dateRaised)}</span>
            <div className="spacer" />
            <IconBtn title="Delete" danger onClick={() => remove("bossReminders", r.id)}><Trash2 size={14} /></IconBtn>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ============================== quick notes ============================== */

function QuickNotesTab({ data, add, update, remove }) {
  const [text, setText] = useState("");

  const submit = () => {
    if (!text.trim()) return;
    add("quickNotes", { text, date: todayISO() });
    setText("");
  };

  const notes = (data.quickNotes || []).slice().reverse();

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Quick Notes</h1>
        <PrintButton label="Quick Notes" />
      </div>
      <p className="hint">A scratchpad for anything you need to jot down fast — sort it into the right tab later.</p>
      <div className="inline-form">
        <input placeholder="Jot something down…" value={text} onChange={e => setText(e.target.value)} onKeyDown={e => e.key === "Enter" && submit()} />
        <button className="btn-primary" onClick={submit}>Add</button>
      </div>
      {notes.length === 0 && <EmptyState text="Nothing jotted down yet." />}
      {notes.map(n => (
        <div className="list-card" key={n.id}>
          <div className="list-card-head">
            <div className="list-card-title">{n.text}</div>
            <span className="pill pill-muted">{fmtDate(n.date)}</span>
            <div className="spacer" />
            <IconBtn title="Delete" danger onClick={() => remove("quickNotes", n.id)}><Trash2 size={14} /></IconBtn>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ============================== styles ============================== */

function Styles() {
  return (
    <style>{`
      :root {
        --bg: #F5F3EE;
        --panel: #FFFFFF;
        --ink: #23262B;
        --ink-soft: #6B6558;
        --border: #DAD5C8;
        --navy: #1E2A38;
        --accent: #C1622D;
        --ok: #4C7A5E;
        --ok-bg: #E7EFE9;
        --warn: #C1690C;
        --warn-bg: #F7E8D6;
        --danger: #A63D33;
        --danger-bg: #F5E1DE;
      }
      * { box-sizing: border-box; }
      .app-shell {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Inter, sans-serif;
        color: var(--ink);
        background: var(--bg);
        min-height: 100%;
        border-radius: 12px;
        overflow: hidden;
      }
      .loading { padding: 40px; text-align: center; color: var(--ink-soft); }
      .topbar {
        display: flex; align-items: center; justify-content: space-between;
        gap: 14px; padding: 14px 20px; background: var(--navy); color: #fff;
      }
      .menu-btn {
        background: rgba(255,255,255,0.1); border: none; color: #fff; width: 34px; height: 34px;
        border-radius: 6px; display: flex; align-items: center; justify-content: center; cursor: pointer; flex-shrink: 0;
      }
      .menu-btn:hover { background: rgba(255,255,255,0.2); }
      .brand { display: flex; align-items: center; gap: 10px; }
      .brand-mark {
        width: 34px; height: 34px; border-radius: 6px; background: var(--accent);
        display: flex; align-items: center; justify-content: center;
        font-family: Georgia, "Times New Roman", serif; font-weight: 700; font-size: 14px;
      }
      .brand-title { font-family: Georgia, "Times New Roman", serif; font-size: 16px; line-height: 1.2; }
      .brand-sub { font-size: 11px; color: #C9CFD6; }
      .search-wrap { position: relative; flex: 1; max-width: 420px; }
      .search-icon { position: absolute; left: 10px; top: 9px; color: #9AA3AD; }
      .search-input {
        width: 100%; padding: 8px 32px; border-radius: 6px; border: 1px solid transparent;
        background: rgba(255,255,255,0.12); color: #fff; font-size: 13px;
      }
      .search-input::placeholder { color: #B7BEC6; }
      .search-clear { position: absolute; right: 8px; top: 8px; background: none; border: none; color: #C9CFD6; cursor: pointer; }
      .layout { display: flex; align-items: flex-start; position: relative; }
      .nav-scrim { position: fixed; inset: 0; background: rgba(20,20,20,0.35); z-index: 20; }
      .sidenav {
        width: 220px; flex-shrink: 0; background: var(--panel);
        border-right: 1px solid var(--border); padding: 12px 8px;
        position: fixed; top: 0; bottom: 0; left: 0; z-index: 21;
        transform: translateX(-100%); transition: transform 0.18s ease;
        box-shadow: 2px 0 12px rgba(0,0,0,0.15); padding-top: 20px; overflow-y: auto;
      }
      .sidenav-open { transform: translateX(0); }
      .content { width: 100%; }
      .nav-item {
        display: flex; align-items: center; gap: 8px; width: 100%; text-align: left;
        padding: 8px 10px; border-radius: 6px; border: none; background: none;
        color: var(--ink); font-size: 13px; cursor: pointer; margin-bottom: 2px;
      }
      .nav-item:hover { background: #F0EEE7; }
      .nav-item-active { background: var(--navy); color: #fff; }
      .content { flex: 1; padding: 24px 28px; min-width: 0; }
      .page-title { font-family: Georgia, "Times New Roman", serif; font-size: 24px; margin: 0 0 14px; }
      .section-heading { font-family: Georgia, "Times New Roman", serif; font-size: 18px; margin: 28px 0 10px; }
      .page-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px; flex-wrap: wrap; gap: 8px; }
      .page-actions { display: flex; gap: 8px; }
      .hint { color: var(--ink-soft); font-size: 12px; margin: 4px 0 16px; }
      .btn-primary {
        display: inline-flex; align-items: center; gap: 6px; background: var(--navy); color: #fff;
        border: none; padding: 8px 14px; border-radius: 6px; font-size: 13px; cursor: pointer;
      }
      .btn-primary:hover { background: #16202B; }
      .btn-ghost {
        display: inline-flex; align-items: center; gap: 6px; background: transparent; color: var(--ink);
        border: 1px solid var(--border); padding: 7px 12px; border-radius: 6px; font-size: 13px; cursor: pointer; text-decoration: none;
      }
      .btn-ghost:hover { background: #EFEBE2; }
      .icon-btn { background: none; border: none; color: var(--ink-soft); cursor: pointer; padding: 4px; border-radius: 4px; display: inline-flex; }
      .icon-btn:hover { background: #EFEBE2; color: var(--ink); }
      .icon-btn-danger:hover { color: var(--danger); }
      .form-panel {
        background: var(--panel); border: 1px solid var(--border); border-radius: 8px;
        padding: 16px; margin: 12px 0 18px; display: grid; grid-template-columns: 1fr 1fr; gap: 12px;
      }
      .form-panel .field:has(textarea) { grid-column: 1 / -1; }
      .form-actions { grid-column: 1 / -1; display: flex; gap: 8px; }
      .field { display: flex; flex-direction: column; gap: 4px; font-size: 12px; color: var(--ink-soft); }
      .field input, .field select, .field textarea {
        font-size: 13px; padding: 7px 9px; border: 1px solid var(--border); border-radius: 5px;
        background: #fff; color: var(--ink); font-family: inherit;
      }
      .field textarea { min-height: 60px; resize: vertical; }
      .list-card { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; margin-bottom: 10px; overflow: hidden; }
      .list-card-danger { border-left: 4px solid var(--danger); }
      .list-card-warn { border-left: 4px solid var(--warn); }
      .list-card-head { display: flex; align-items: center; gap: 8px; padding: 12px 14px; cursor: default; flex-wrap: wrap; }
      .list-card-title { font-weight: 600; font-size: 14px; margin-right: 4px; }
      .list-card-body { padding: 0 14px 14px 14px; border-top: 1px solid var(--border); padding-top: 10px; font-size: 13px; }
      .spacer { flex: 1; }
      .pill { font-size: 11px; padding: 3px 8px; border-radius: 20px; background: #EEEAE0; color: var(--ink-soft); white-space: nowrap; }
      .pill-muted { background: #EEEAE0; color: var(--ink-soft); }
      .pill-ok { background: var(--ok-bg); color: var(--ok); }
      .pill-warn { background: var(--warn-bg); color: var(--warn); }
      .pill-danger { background: var(--danger-bg); color: var(--danger); }
      .count-badge { background: var(--accent); color: #fff; font-size: 11px; padding: 1px 7px; border-radius: 10px; }
      .row { padding: 5px 0; font-size: 13px; }
      .row-muted { color: var(--ink-soft); }
      .strike { text-decoration: line-through; color: var(--ink-soft); }
      .inline-form { display: flex; gap: 8px; margin: 8px 0; }
      .inline-form input { flex: 1; padding: 7px 9px; border: 1px solid var(--border); border-radius: 5px; font-size: 13px; }
      .filter-row { display: flex; gap: 6px; align-items: center; margin-bottom: 14px; flex-wrap: wrap; }
      .chip { border: 1px solid var(--border); background: #fff; padding: 6px 12px; border-radius: 20px; font-size: 12px; cursor: pointer; color: var(--ink-soft); }
      .chip-active { background: var(--navy); border-color: var(--navy); color: #fff; }
      .card-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 8px; }
      .stat-card { text-align: left; border: 1px solid var(--border); background: var(--panel); border-radius: 8px; padding: 16px; cursor: pointer; }
      .stat-card:hover { border-color: var(--navy); }
      .stat-value { font-family: Georgia, serif; font-size: 30px; line-height: 1; margin-bottom: 6px; }
      .stat-label { font-size: 12px; color: var(--ink-soft); }
      .stat-danger .stat-value { color: var(--danger); }
      .stat-warn .stat-value { color: var(--warn); }
      .stat-ok .stat-value { color: var(--ok); }
      .building-overview-list { display: flex; flex-direction: column; gap: 8px; }
      .building-row { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px; }
      .building-row-name { font-weight: 600; font-size: 14px; margin-bottom: 4px; }
      .building-row-stats { display: flex; gap: 14px; font-size: 12px; color: var(--ink-soft); flex-wrap: wrap; }
      .section { margin-bottom: 20px; }
      .section-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }
      .section-title { display: flex; align-items: center; gap: 8px; }
      .section-title h2 { font-size: 15px; margin: 0; font-family: Georgia, serif; }
      .empty-state { color: var(--ink-soft); font-size: 13px; padding: 18px; text-align: center; border: 1px dashed var(--border); border-radius: 8px; }
      .law-table { display: flex; flex-direction: column; gap: 6px; }
      .law-row { display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 10px; align-items: center; background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 10px 14px; }
      .law-name { font-size: 13px; }
      .law-row input, .law-row select { padding: 6px 8px; border: 1px solid var(--border); border-radius: 5px; font-size: 12px; }
      .search-results-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
      .photo-uploader { margin-top: 4px; }
      .photo-grid { display: flex; gap: 8px; flex-wrap: wrap; }
      .photo-thumb { position: relative; width: 56px; height: 56px; border-radius: 6px; overflow: hidden; border: 1px solid var(--border); }
      .photo-thumb img { width: 100%; height: 100%; object-fit: cover; display: block; }
      .photo-remove {
        position: absolute; top: 2px; right: 2px; background: rgba(0,0,0,0.55); color: #fff;
        border: none; border-radius: 50%; width: 16px; height: 16px; display: flex; align-items: center; justify-content: center; cursor: pointer;
      }
      .photo-add {
        width: 56px; height: 56px; border-radius: 6px; border: 1px dashed var(--border); background: #fff;
        display: flex; align-items: center; justify-content: center; color: var(--ink-soft); cursor: pointer;
      }
      .photo-add:hover { border-color: var(--navy); color: var(--navy); }

      .checklist { display: flex; flex-direction: column; gap: 4px; margin: 4px 0; }
      .checklist-item { display: flex; align-items: center; gap: 8px; font-size: 13px; }
      .checklist-remove { background: none; border: none; color: var(--ink-soft); cursor: pointer; margin-left: auto; padding: 2px; }
      .checklist-remove:hover { color: var(--danger); }
      .doc-uploader { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; margin: 4px 0; }
      .doc-chip { display: flex; align-items: center; gap: 4px; background: #EEEAE0; border-radius: 20px; padding: 4px 6px 4px 10px; font-size: 12px; }
      .doc-chip-name { color: var(--ink); text-decoration: none; max-width: 140px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .doc-chip-name:hover { text-decoration: underline; }
      .doc-remove { background: none; border: none; color: var(--ink-soft); cursor: pointer; border-radius: 50%; display: flex; padding: 2px; }
      .doc-remove:hover { color: var(--danger); }
      .import-preview { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 14px; margin-bottom: 16px; }
      .import-preview-summary { font-weight: 600; font-size: 13px; margin-bottom: 8px; }
      .import-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; padding: 6px 0; border-bottom: 1px solid var(--border); font-size: 13px; }
      .import-row:last-of-type { border-bottom: none; }
      .import-row-name { font-weight: 600; }
      .import-row-fields { display: flex; flex-wrap: wrap; gap: 8px; color: var(--ink-soft); font-size: 12px; }

      .attention-icon { flex-shrink: 0; }
      .attention-count { font-family: Georgia, serif; font-size: 17px; margin-right: 2px; }
      .attention-label { font-size: 13px; color: var(--ink-soft); flex: 1; }
      .all-clear {
        display: flex; align-items: center; gap: 10px; background: var(--ok-bg); color: var(--ok);
        border-radius: 8px; padding: 16px; font-size: 13px; margin-bottom: 24px;
      }
      .dash-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 10px; }
      .dash-building-card { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 14px; }
      .dash-building-name { font-weight: 600; font-size: 13px; margin-bottom: 8px; }
      .dash-building-chips { display: flex; flex-wrap: wrap; gap: 6px; }
      .dash-building-clear { font-size: 12px; color: var(--ok); }

      .sheet-wrap { overflow-x: auto; border: 1px solid var(--border); border-radius: 8px; background: var(--panel); }
      .sheet { border-collapse: collapse; width: 100%; font-size: 13px; }
      .sheet th {
        text-align: left; font-weight: 600; font-size: 11px; color: var(--ink-soft);
        background: #EFEBE2; padding: 8px 10px; border-bottom: 1px solid var(--border); border-right: 1px solid var(--border);
        position: sticky; top: 0;
      }
      .sheet td { border-bottom: 1px solid var(--border); border-right: 1px solid var(--border); padding: 0; }
      .sheet tr:last-child td { border-bottom: none; }
      .sheet td:last-child, .sheet th:last-child { border-right: none; }
      .sheet-row-flag { background: #FBF6EF; }
      .sheet-input {
        width: 100%; border: none; background: transparent; padding: 8px 10px; font-size: 13px;
        font-family: inherit; color: var(--ink); border-radius: 0;
      }
      .sheet-input:focus { outline: 2px solid var(--navy); outline-offset: -2px; background: #fff; }
      .sheet-status-ok { color: var(--ok); font-weight: 600; }
      .sheet-status-warn { color: var(--warn); font-weight: 600; }
      .sheet-status-danger { color: var(--danger); font-weight: 600; }
      .sheet-actions { display: flex; gap: 2px; padding: 4px 6px !important; white-space: nowrap; }
      .sheet-expand-row td { background: #FAF9F6; padding: 10px 14px !important; }
      .sheet-note-preview {
        display: flex; flex-direction: column; align-items: flex-start; gap: 3px; width: 100%; text-align: left;
        background: none; border: none; padding: 8px 10px; cursor: pointer; font-family: inherit;
      }
      .sheet-note-preview:hover { background: #F0EEE7; }
      .sheet-note-line { display: flex; flex-direction: column; gap: 0; }
      .sheet-note-date { font-size: 10px; color: var(--ink-soft); }
      .sheet-note-text { font-size: 12px; color: var(--ink); white-space: normal; overflow-wrap: break-word; }
      .sheet-col-notes { min-width: 220px; max-width: 320px; }
      .sheet-col-balance { width: 76px; max-width: 76px; }
      .sheet-follow-btn {
        display: flex; flex-direction: column; align-items: center; gap: 2px; background: none; border: none;
        color: var(--ink-soft); cursor: pointer; padding: 8px 10px; width: 100%;
      }
      .sheet-follow-btn:hover { background: #F0EEE7; color: var(--navy); }
      .sheet-follow-date { font-size: 10px; }
      .followup-scroll { max-height: 520px; overflow-y: auto; padding-right: 4px; }
      .followup-panel { background: var(--panel); border: 1px solid var(--border); border-left: 4px solid var(--warn); border-radius: 8px; margin-bottom: 24px; overflow: hidden; }
      .followup-panel-head {
        display: flex; align-items: center; gap: 12px; width: 100%; background: none; border: none;
        padding: 12px 14px; cursor: pointer; text-align: left;
      }
      .followup-panel-head:hover { background: #FAF6EE; }
      .followup-panel-body { border-top: 1px solid var(--border); padding: 6px 14px 10px; }
      .followup-subsection { margin-bottom: 12px; }
      .followup-subsection:last-child { margin-bottom: 0; }
      .followup-subsection-title { font-size: 12px; font-weight: 600; color: var(--ink-soft); margin: 8px 0 4px; }
      .followup-item { display: flex; align-items: flex-start; gap: 10px; padding: 8px 0; border-bottom: 1px solid var(--border); }
      .followup-item:last-child { border-bottom: none; }
      .followup-item-due { background: #FBF6EF; margin: 0 -8px; padding: 8px; border-radius: 6px; border-bottom-color: transparent; }
      .followup-item-main { flex: 1; }
      .followup-item-name { font-size: 13px; font-weight: 600; }
      .followup-item-note { font-size: 12px; color: var(--ink-soft); margin-top: 2px; }

      @media (max-width: 720px) {
        .form-panel { grid-template-columns: 1fr; }
        .law-row { grid-template-columns: 1fr; }
        .sidenav { width: 78%; }
      }
      .login-shell {
        min-height: 100vh; display: flex; align-items: center; justify-content: center;
        background: var(--bg); font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Inter, sans-serif;
      }
      .login-card {
        background: var(--panel); border: 1px solid var(--border); border-radius: 10px;
        padding: 32px; width: 320px; display: flex; flex-direction: column;
      }
      @media print {
        .no-print { display: none !important; }
        .app-shell { border-radius: 0; background: #fff; }
        .content { padding: 0; }
        .list-card, .stat-card, .building-row, .dash-building-card { break-inside: avoid; border: 1px solid #ccc; }
        body { background: #fff; }
      }
    `}</style>
  );
}
