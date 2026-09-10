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
  apiKey: "AIzaSyA5Op33F-BQSfKVbe3zx3jlbfZdsCSWT2c",
  authDomain: "management-board-cb4cc.firebaseapp.com",
  projectId: "management-board-cb4cc",
  storageBucket: "management-board-cb4cc.firebasestorage.app",
  messagingSenderId: "179973991586",
  appId: "1:179973991586:web:7a88edf20297671eb0bbfb",
  measurementId: "G-ZNC3V3JXE2",
};
const firebaseApp = initializeApp(firebaseConfig);
const auth = getAuth(firebaseApp);
const db = getFirestore(firebaseApp);
import {
  Search, Building2, Users, Wrench, AlertTriangle, Gavel, HardHat, Home,
  CalendarClock, ScrollText, MessageSquare, Archive as ArchiveIcon,
  Plus, X, Camera, Download, LayoutDashboard, ChevronDown, ChevronRight, ChevronLeft,
  Trash2, Pencil, Upload, Menu, Printer, CheckCircle2
} from "lucide-react";

/* ============================== constants ============================== */

const STORAGE_KEY = "pm-ops-data-v1";
const uid = () => Math.random().toString(36).slice(2, 10);
// IMPORTANT: never use `.toISOString()` for local dates — that returns the UTC
// date, not the local one. For anyone west of UTC (all of the US), once evening
// hits, UTC has already rolled over to tomorrow — the app would think "today"
// is tomorrow, misclassifying everything due today as overdue and everything
// due tomorrow as due today. Always build the date string from local components.
const toLocalISO = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const todayISO = () => toLocalISO(new Date());

// "B4" must sort before "B39", not after it — plain string sort gets this wrong
// because it compares character by character. Split into a letter prefix and a
// numeric suffix and compare those separately.
function unitSortKey(u) {
  const m = String(u || "").match(/^([A-Za-z]*)(\d*)/);
  return [m[1] || "", parseInt(m[2] || "0", 10) || 0];
}
function compareUnits(a, b) {
  const [al, an] = unitSortKey(a);
  const [bl, bn] = unitSortKey(b);
  if (al !== bl) return al.localeCompare(bl);
  return an - bn;
}

// Used anywhere a unit picker needs to show who lives there, not just the bare
// apt number — makes it obvious at a glance which apt you're actually picking.
function unitOptionLabel(unit, allTenants) {
  const names = allTenants.filter(t => t.unitId === unit.id).map(t => t.name).filter(Boolean);
  return names.length ? `${unit.unitNumber} — ${names.join(" & ")}` : `${unit.unitNumber} (no tenant on file)`;
}

// Full addresses ("333 OVINGTON AVENUE BROOKLYN, NEW YORK 11209") are useful in
// the Buildings tab itself, but everywhere else a short form ("333 OVINGTON
// AVENUE") is plenty and keeps rows/pills from getting cluttered.
function shortAddress(address) {
  if (!address) return "";
  let s = address.replace(/,\s*[A-Za-z .]+\s+\d{5}(-\d{4})?\s*$/, "");
  s = s.replace(/\s+(BROOKLYN|MANHATTAN|QUEENS|BRONX|STATEN ISLAND|NEW YORK)\s*$/i, "");
  return s.trim() || address;
}

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
  return toLocalISO(d);
}
function addDays(dateStr, n) {
  const d = new Date(dateStr + "T00:00:00");
  d.setDate(d.getDate() + n);
  return toLocalISO(d);
}
function isInFollowUpWindow(dateStr) {
  const d = daysUntil(dateStr);
  if (d === null) return false;
  return d >= -7 && d <= 1;
}
// Tenants can now have several follow-ups instead of just one. Old records
// only had followUpDate/followUpNote — read those in as a single-item list
// so nothing already saved gets lost.
function tenantFollowUps(t) {
  if (Array.isArray(t.followUps)) return t.followUps;
  // Every tenant on the old single-follow-up format used to get the exact same
  // literal id "legacy" here — harmless for one tenant, but the instant a SECOND
  // tenant also had an old-format follow-up, every list keyed by that id (Due
  // Today, the follow-up panel, the calendar) would collide and silently drop
  // or misrender one of them. Tie the id to the tenant so it's always unique.
  if (t.followUpDate) return [{ id: `legacy-${t.id}`, date: t.followUpDate, note: t.followUpNote || "" }];
  return [];
}
function earliestFollowUpDate(t) {
  const dates = tenantFollowUps(t).map(f => f.date).filter(Boolean).sort();
  return dates[0] || null;
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
const SKIP_LINE_RE = /^\d{2}\/\d{2}\/\d{4}|FISCAL PERIOD|^Page:|PROP #|TELEPHONE\/EMAIL LIST|BUILDING DIRECTORY|AGED ARREARS|^LEGAL:|^\*\s*-\s*MOVED OUT|^TOTALS:|^TENANT NAME:/i;

function cleanLines(text) {
  return text.split("\n").map(l => l.trim()).map(l => {
    // Some reports print "APT: A1" as one line (the real code merged with the
    // column label) instead of the usual separate "APT: TENANT NAME:" header
    // row — keep the code, drop just the label, so that apartment isn't lost.
    const m = l.match(/^APT:\s*(.+)$/i);
    return m ? m[1].trim() : l;
  }).filter(l => l && !SKIP_LINE_RE.test(l) && !/^APT:?$/i.test(l));
}

// Every RIS report repeats a header like:
// "PROP # 333OVI: SG & SONS REALTY LLC - 333 OVINGTON AVENUE BROOKLYN, NEW YORK 11209"
// Extract a stable property code + the address so uploads can auto-detect/create the building.
// Stops at the ZIP code so trailing text merged in by PDF extraction (fiscal period, page
// numbers, etc.) never gets pulled into the address.
function parseBuildingHeader(text) {
  const m = text.match(/PROP\s*#\s*([^\s:]+):\s*[\s\S]+?-\s*([\s\S]+?\b\d{5}\b)/);
  if (!m) return { propCode: "", address: "" };
  const norm = s => s.replace(/\s+/g, " ").trim();
  return { propCode: norm(m[1]), address: norm(m[2]) };
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
  // Different PDF viewers copy multi-line contact cells in unpredictable order —
  // sometimes labels and values stay paired, sometimes all labels get grouped
  // together with all their values afterward. Rather than trying to track which
  // line goes with which, treat the whole report as one continuous block: find
  // every apartment header, then pull the first phone number and first email
  // found anywhere between that header and the next one, wherever it landed.
  const cleaned = cleanLines(text).join(" ");
  const LABELS = "(?:CELL|EMAIL ADDRESS|HOME|WORK|OTHER|FAX)";
  // Some buildings also have commercial/storefront units identified by a bare
  // number (no letter prefix, e.g. "9516 JH ORGANIC INC."), and others use
  // digit-then-letter codes ("1B", "2BB") instead of letter-then-digit. Support
  // all three, but require the numeric/digit-first forms to have a real name
  // after them — otherwise a phone number written with spaces instead of dashes
  // ("718 833 3607 FAX") can look just like a unit code.
  const ALT_APT_RE = "\\d{1,2}[A-Z]{1,2}";
  const NEXT_HEADER = `(?:${APT_RE}|\\d{4}|${ALT_APT_RE})\\s+(?:MR\\.|MRS\\.|MS\\.|[A-Z])`;
  const headerRe = new RegExp(
    `(?:^|\\s)(?:` +
      `(${APT_RE})\\s+(?!${LABELS}\\b)((?:MR\\.|MRS\\.|MS\\.|[A-Z])[A-Za-z.,'\\-\\s]*?)` +
      `|` +
      `(\\d{4})\\s+(?!${LABELS}\\b)((?:MR\\.|MRS\\.|MS\\.|[A-Z])[A-Za-z.,'\\-\\s]+?)` +
      `|` +
      `(${ALT_APT_RE})\\s+(?!${LABELS}\\b)((?:MR\\.|MRS\\.|MS\\.|[A-Z])[A-Za-z.,'\\-\\s]*?)` +
    `)(?=\\s+${LABELS}\\b|\\s+${NEXT_HEADER}|$)`,
    "g"
  );
  const headers = [];
  let m;
  while ((m = headerRe.exec(cleaned))) {
    headers.push({ apt: m[1] || m[3] || m[5], name: (m[2] || m[4] || m[6] || "").trim(), start: m.index, end: m.index + m[0].length });
  }

  const phoneRe = /\(?\d{3}\)?[-.\s]*\d{3}[-.\s]*\d{4}/;
  const emailRe = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;

  return headers.map((h, i) => {
    const blockEnd = i + 1 < headers.length ? headers[i + 1].start : cleaned.length;
    const block = cleaned.slice(h.end, blockEnd);
    const phone = block.match(phoneRe);
    const email = block.match(emailRe);
    return { apt: h.apt, name: h.name, phone: phone ? phone[0].trim() : "", email: email ? email[0].trim() : "" };
  }); // keep every detected apartment, even ones with no phone/email on file
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

function PinSetupPanel({ onClose }) {
  const [pin1, setPin1] = useState("");
  const [pin2, setPin2] = useState("");
  const [error, setError] = useState("");
  const hasExisting = !!localStorage.getItem(PIN_STORAGE_KEY);

  const save = () => {
    if (pin1.length < 4) { setError("At least 4 digits."); return; }
    if (pin1 !== pin2) { setError("PINs don't match."); return; }
    localStorage.setItem(PIN_STORAGE_KEY, pin1);
    onClose();
  };
  const clear = () => {
    localStorage.removeItem(PIN_STORAGE_KEY);
    onClose();
  };

  return (
    <div className="form-panel no-print" style={{ margin: "0 20px 16px" }}>
      <Field label="New PIN (4–6 digits)">
        <input type="password" inputMode="numeric" pattern="[0-9]*" maxLength={6} value={pin1} onChange={e => setPin1(e.target.value.replace(/\D/g, ""))} />
      </Field>
      <Field label="Confirm PIN">
        <input type="password" inputMode="numeric" pattern="[0-9]*" maxLength={6} value={pin2} onChange={e => setPin2(e.target.value.replace(/\D/g, ""))} />
      </Field>
      {error && <div className="hint" style={{ color: "var(--danger)", gridColumn: "1 / -1" }}>{error}</div>}
      <div className="form-actions">
        <button className="btn-primary" onClick={save}>Save PIN</button>
        {hasExisting && <button className="btn-ghost" onClick={clear}>Remove PIN</button>}
        <button className="btn-ghost" onClick={onClose}>Cancel</button>
      </div>
    </div>
  );
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

const PIN_STORAGE_KEY = "pm-ops-device-pin";

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

function PinLockScreen({ onUnlock, onForgot }) {
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");

  const submit = (e) => {
    e.preventDefault();
    const saved = localStorage.getItem(PIN_STORAGE_KEY);
    if (pin && pin === saved) {
      onUnlock();
    } else {
      setError("Wrong PIN — try again.");
      setPin("");
    }
  };

  return (
    <div className="login-shell">
      <form className="login-card" onSubmit={submit}>
        <div className="brand-mark" style={{ marginBottom: 14 }}>PO</div>
        <h1 className="page-title" style={{ marginBottom: 4 }}>Enter PIN</h1>
        <p className="hint" style={{ marginBottom: 16 }}>Quick unlock for this device — you're still signed in.</p>
        <Field label="PIN">
          <input
            type="password" inputMode="numeric" pattern="[0-9]*" maxLength={6} autoFocus
            value={pin} onChange={e => setPin(e.target.value.replace(/\D/g, ""))}
          />
        </Field>
        {error && <div className="hint" style={{ color: "var(--danger)" }}>{error}</div>}
        <button className="btn-primary" type="submit" disabled={!pin} style={{ marginTop: 10, width: "100%", justifyContent: "center" }}>
          Unlock
        </button>
        <button type="button" className="btn-ghost" style={{ marginTop: 8, width: "100%", justifyContent: "center" }} onClick={onForgot}>
          Forgot PIN — sign out fully
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
  const [pinUnlocked, setPinUnlocked] = useState(() => !localStorage.getItem(PIN_STORAGE_KEY));
  const [showPinSetup, setShowPinSetup] = useState(false);
  const [saveError, setSaveError] = useState(false);
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
      try {
        // Firestore rejects any field that's literally `undefined` (as opposed to
        // just missing) and throws instead of saving anything. A JSON round-trip
        // strips those out automatically so a stray undefined somewhere can never
        // silently break autosave.
        const safe = JSON.parse(JSON.stringify(data));
        await setDoc(doc(db, "users", user.uid, "appData", "main"), safe);
        setSaveError(false);
      } catch (e) {
        console.error("save failed", e);
        setSaveError(true);
      }
    }, 400);
    return () => clearTimeout(saveTimer.current);
  }, [data, loaded, user]);

  // generic collection ops
  const add = (col, item) => setData(d => ({ ...d, [col]: [...d[col], { id: uid(), ...item }] }));
  const update = (col, id, patch) => setData(d => ({ ...d, [col]: d[col].map(x => x.id === id ? { ...x, ...patch } : x) }));
  const remove = (col, id) => setData(d => ({ ...d, [col]: d[col].filter(x => x.id !== id) }));

  const buildingName = (id) => shortAddress(data.buildings.find(b => b.id === id)?.address) || "—";
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
  if (!pinUnlocked) return (
    <PinLockScreen
      onUnlock={() => setPinUnlocked(true)}
      onForgot={() => { localStorage.removeItem(PIN_STORAGE_KEY); signOut(auth); }}
    />
  );
  if (!loaded) return <div className="app-shell"><div className="loading">Loading your ops board…</div><Styles /></div>;

  return (
    <div className="app-shell">
      <header className="topbar no-print">
        <div className="topbar-left">
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
        <div className="topbar-actions">
          <button className="btn-ghost no-print" title="Set a quick-unlock PIN for this device" onClick={() => setShowPinSetup(s => !s)} style={{ color: "#fff", borderColor: "rgba(255,255,255,0.3)" }}>
            {localStorage.getItem(PIN_STORAGE_KEY) ? "Change PIN" : "Set PIN"}
          </button>
          <button className="btn-ghost no-print" title="Sign out" onClick={() => { localStorage.removeItem(PIN_STORAGE_KEY); signOut(auth); }} style={{ color: "#fff", borderColor: "rgba(255,255,255,0.3)" }}>
            Sign out
          </button>
        </div>
      </header>

      {saveError && (
        <div className="save-error-banner no-print">
          <AlertTriangle size={16} /> Couldn't save your last change — check your internet connection. Your edits are still here on screen, but won't be there if you close the tab until this clears.
        </div>
      )}

      {showPinSetup && (
        <PinSetupPanel onClose={() => setShowPinSetup(false)} />
      )}

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
            {tab === "dashboard" && <Dashboard data={data} buildingName={buildingName} tenantName={tenantName} setTab={setTab} setData={setData} />}
            {tab === "buildings" && <BuildingsTab data={data} add={add} update={update} remove={remove} setData={setData} buildingName={buildingName} />}
            {tab === "rent" && <RentTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} unitLabel={unitLabel} setData={setData} />}
            {tab === "workorders" && <WorkOrdersTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} vendorName={vendorName} />}
            {tab === "violations" && <ViolationsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} vendorName={vendorName} setData={setData} />}
            {tab === "vendors" && <VendorsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} />}
            {tab === "court" && <CourtTab data={data} add={add} update={update} remove={remove} tenantName={tenantName} buildingName={buildingName} />}
            {tab === "inspections" && <AppointmentsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} setData={setData} />}
            {tab === "laws" && <LocalLawsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} />}
            {tab === "reminders" && <RemindersTab data={data} add={add} update={update} remove={remove} />}
            {tab === "quicknotes" && <QuickNotesTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} />}
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
          {results.buildings.map(b => <div className="row" key={b.id}>{shortAddress(b.address)}</div>)}
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

function AttentionPanel({ icon, label, items, tab, setTab, renderItem, itemKey, extraAction }) {
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
          <div className="row" style={{ marginTop: 8, gap: 8 }}>
            <button className="btn-ghost" onClick={() => setTab(tab)}>
              View in {TAB_LABELS[tab] || tab}
            </button>
            {extraAction}
          </div>
        </div>
      )}
    </div>
  );
}

// Pulls together everything due on one specific date, across every part of the
// app, so the dashboard calendar can show a single day view of it all.
// Pulls together every dated item across the whole app — no single-date filter,
// so nothing (overdue included) can silently fall outside a visible window.
function allDatedItems(data, tenantName, buildingName) {
  const items = [];
  data.violations.forEach(v => {
    if (v.cureDeadline && !isViolationClosed(v)) {
      items.push({ key: `v-${v.id}`, date: v.cureDeadline, type: `${v.agency} cure deadline`, label: `#${v.violationNumber}`, sub: buildingName(v.buildingId), tab: "violations" });
    }
  });
  data.courtCases.forEach(c => {
    if (c.archived) return;
    if (c.nextCourtDate) items.push({ key: `c-${c.id}`, date: c.nextCourtDate, type: "Court date", label: tenantName(c.tenantId), sub: buildingName(c.buildingId), tab: "court" });
    if (c.result === "Stipulation (payment plan)" && c.nextPaymentDue) items.push({ key: `p-${c.id}`, date: c.nextPaymentDue, type: "Payment due", label: tenantName(c.tenantId), sub: buildingName(c.buildingId), tab: "court" });
  });
  data.appointments.forEach(a => {
    if (!a.completed && a.date) {
      items.push({ key: `a-${a.id}`, date: a.date, type: a.recurring ? "Recurring appointment" : "Appointment", label: a.type, sub: buildingName(a.buildingId), tab: "inspections" });
    }
  });
  data.tenants.forEach(t => {
    tenantFollowUps(t).forEach(f => {
      if (f.date) items.push({ key: `f-${f.id}`, date: f.date, type: "Follow-up", label: t.name, sub: buildingName(t.buildingId), note: f.note, tab: "rent" });
    });
  });
  (data.quickNotes || []).forEach(n => {
    if (n.reminderDate && !n.done) {
      items.push({ key: `n-${n.id}`, date: n.reminderDate, type: "Reminder", label: n.text, sub: n.buildingId ? buildingName(n.buildingId) : "", tab: "quicknotes" });
    }
  });
  return items;
}

function monthGridDays(refDate) {
  const d = new Date(refDate + "T00:00:00");
  const firstOfMonth = new Date(d.getFullYear(), d.getMonth(), 1);
  const gridStart = new Date(firstOfMonth);
  gridStart.setDate(gridStart.getDate() - firstOfMonth.getDay());
  return Array.from({ length: 42 }, (_, i) => {
    const cell = new Date(gridStart);
    cell.setDate(cell.getDate() + i);
    return toLocalISO(cell);
  });
}

function DashboardCalendar({ data, buildingName, tenantName, setTab }) {
  const [viewMode, setViewMode] = useState("day"); // day | week | month | year
  const [refDate, setRefDate] = useState(todayISO());
  const [showingOverdue, setShowingOverdue] = useState(false);
  const today = todayISO();

  const all = allDatedItems(data, tenantName, buildingName);
  const itemsByDate = {};
  all.forEach(i => { (itemsByDate[i.date] = itemsByDate[i.date] || []).push(i); });
  const overdueItems = all.filter(i => i.date < today).sort((a, b) => a.date.localeCompare(b.date));

  const shift = (n) => {
    setShowingOverdue(false);
    if (viewMode === "day") setRefDate(d => addDays(d, n));
    else if (viewMode === "week") setRefDate(d => addDays(d, n * 7));
    else if (viewMode === "month") setRefDate(d => addMonths(d, n));
    else setRefDate(d => addMonths(d, n * 12));
  };
  const pickDate = (d) => { setRefDate(d); setShowingOverdue(false); };

  const weekDays = Array.from({ length: 7 }, (_, i) => addDays(addDays(refDate, -new Date(refDate + "T00:00:00").getDay()), i));
  const monthDays = monthGridDays(refDate);
  const monthLabel = new Date(refDate + "T00:00:00").toLocaleDateString("en-US", { month: "long", year: "numeric" });
  const yearNum = new Date(refDate + "T00:00:00").getFullYear();
  const monthCounts = Array.from({ length: 12 }, (_, m) => {
    const prefix = `${yearNum}-${String(m + 1).padStart(2, "0")}`;
    return all.filter(i => i.date.startsWith(prefix)).length;
  });

  const selectedItems = showingOverdue ? overdueItems : (itemsByDate[refDate] || []);
  const selectedLabel = showingOverdue
    ? `Overdue (${overdueItems.length})`
    : (refDate === today ? "Today" : new Date(refDate + "T00:00:00").toLocaleDateString("en-US", { weekday: "long", month: "short", day: "numeric" }));

  const Row = (item) => (
    <button key={item.key} className="dash-detail-item" onClick={() => setTab(item.tab)}>
      <span className="pill pill-danger">{item.type}</span>
      <div className="followup-item-main">
        <div className="followup-item-name">{item.label}{item.sub ? <span className="row-muted"> — {item.sub}</span> : null}</div>
        {item.note && <div className="followup-item-note">{item.note}</div>}
      </div>
    </button>
  );

  return (
    <div className="dash-calendar dash-calendar-compact">
      <div className="dash-calendar-head">
        <button className="icon-btn" onClick={() => shift(-1)}><ChevronLeft size={15} /></button>
        <span className="dash-calendar-title">
          {viewMode === "month" ? monthLabel : viewMode === "year" ? yearNum : viewMode === "week" ? `Week of ${fmtDate(weekDays[0])}` : new Date(refDate + "T00:00:00").toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" })}
        </span>
        <button className="icon-btn" onClick={() => shift(1)}><ChevronRight size={15} /></button>
        <div className="spacer" />
        {overdueItems.length > 0 && (
          <button className={`chip dash-cal-overdue-chip ${showingOverdue ? "chip-active" : ""}`} onClick={() => setShowingOverdue(s => !s)}>
            {overdueItems.length} overdue
          </button>
        )}
      </div>

      <div className="filter-row dash-cal-modes">
        {["day", "week", "month", "year"].map(m => (
          <button key={m} className={`chip ${viewMode === m ? "chip-active" : ""}`} onClick={() => { setViewMode(m); setShowingOverdue(false); }}>{m[0].toUpperCase() + m.slice(1)}</button>
        ))}
        {refDate !== today && <button className="btn-ghost" onClick={() => pickDate(today)}>Today</button>}
      </div>

      {viewMode === "week" && (
        <div className="dash-cal-week">
          {weekDays.map(d => (
            <button key={d} className={`dash-cal-cell ${d === refDate ? "dash-cal-cell-selected" : ""} ${d === today ? "dash-cal-cell-today" : ""}`} onClick={() => pickDate(d)}>
              <div className="dash-cal-cell-label">{new Date(d + "T00:00:00").toLocaleDateString("en-US", { weekday: "short" })}</div>
              <div className="dash-cal-cell-num">{new Date(d + "T00:00:00").getDate()}</div>
              {(itemsByDate[d] || []).length > 0 && <span className="dash-cal-dot" />}
            </button>
          ))}
        </div>
      )}

      {viewMode === "month" && (
        <div className="dash-cal-month">
          {["S", "M", "T", "W", "T", "F", "S"].map((d, i) => <div key={i} className="dash-cal-month-dow">{d}</div>)}
          {monthDays.map(d => {
            const inMonth = d.slice(0, 7) === refDate.slice(0, 7);
            return (
              <button key={d} className={`dash-cal-cell dash-cal-cell-sm ${d === refDate ? "dash-cal-cell-selected" : ""} ${d === today ? "dash-cal-cell-today" : ""} ${!inMonth ? "dash-cal-cell-dim" : ""}`} onClick={() => pickDate(d)}>
                <div className="dash-cal-cell-num">{new Date(d + "T00:00:00").getDate()}</div>
                {(itemsByDate[d] || []).length > 0 && <span className="dash-cal-dot" />}
              </button>
            );
          })}
        </div>
      )}

      {viewMode === "year" && (
        <div className="dash-cal-year">
          {Array.from({ length: 12 }, (_, m) => (
            <button key={m} className="dash-cal-month-cell" onClick={() => { setViewMode("month"); setRefDate(`${yearNum}-${String(m + 1).padStart(2, "0")}-01`); }}>
              <div>{new Date(yearNum, m, 1).toLocaleDateString("en-US", { month: "short" })}</div>
              {monthCounts[m] > 0 && <span className="dash-cal-dot" />}
            </button>
          ))}
        </div>
      )}

      <div className="dash-cal-section">
        <div className={`dash-cal-section-title ${showingOverdue ? "dash-cal-overdue" : ""}`}>{selectedLabel}</div>
        {selectedItems.length === 0 ? <div className="hint">Nothing here.</div> : selectedItems.map(Row)}
      </div>
    </div>
  );
}

function Dashboard({ data, buildingName, tenantName, setTab, setData }) {
  const [rentPanelOpen, setRentPanelOpen] = useState(false);
  const [showAllOverdue, setShowAllOverdue] = useState(false);
  const [confirmingCleanup, setConfirmingCleanup] = useState(false);
  const [expandedBuilding, setExpandedBuilding] = useState(null);

  const cleanupAllEmptyUnits = () => {
    setData(d => ({
      ...d,
      units: d.units.filter(u => d.tenants.some(t => t.unitId === u.id)),
    }));
    setConfirmingCleanup(false);
  };
  const [monthsToShow, setMonthsToShow] = useState(1);

  // Tenants with a follow-up already scheduled show up under "to follow up" —
  // no need to also flag them under "not current on rent", that's just noise.
  const overdueTenants = data.tenants.filter(t => t.status !== "Current" && tenantFollowUps(t).length === 0);
  // Each open violation goes into exactly ONE of these buckets, by agency, so it
  // never shows up twice on the dashboard. Window covers overdue + due within 10 days.
  const violationDue = (v) => { const d = daysUntil(v.cureDeadline); return d !== null && d <= 10 && !isViolationClosed(v); };
  const hpdDue = data.violations.filter(v => v.agency === "HPD" && violationDue(v));
  const dobDue = data.violations.filter(v => v.agency === "Other" && v.otherAgency === "DOB" && violationDue(v));
  const fdnyDue = data.violations.filter(v => v.agency === "Other" && v.otherAgency === "FDNY" && violationDue(v));
  const otherViolationsDue = data.violations.filter(v =>
    (v.agency === "DSNY" || (v.agency === "Other" && v.otherAgency !== "DOB" && v.otherAgency !== "FDNY")) && violationDue(v)
  );
  const courtItems = data.courtCases.filter(c => !c.archived && flagFor(c.nextCourtDate));
  const stipItems = data.courtCases.filter(c => !c.archived && c.result === "Stipulation (payment plan)" && flagFor(c.nextPaymentDue));
  const recurringItems = data.appointments.filter(a => !a.completed && a.recurring && flagFor(a.date));
  const appointmentItems = data.appointments.filter(a => !a.completed && !a.recurring && flagFor(a.date));
  const quickNoteItems = (data.quickNotes || []).filter(n => !n.done);
  const bossReminderItems = data.bossReminders.filter(r => r.status !== "Done");
  const followUpCutoff = addMonths(todayISO(), monthsToShow);
  const tenantsWithFollowUps = data.tenants.filter(t => tenantFollowUps(t).length > 0);
  const allFollowUps = tenantsWithFollowUps.length;
  // Flatten to one row per (tenant, follow-up) pair, since a tenant can have several.
  const followUpEntries = tenantsWithFollowUps
    .flatMap(t => tenantFollowUps(t).map(f => ({ tenant: t, followUp: f })))
    .filter(e => e.followUp.date && e.followUp.date <= followUpCutoff)
    .sort((a, b) => (a.followUp.date || "").localeCompare(b.followUp.date || ""));
  const overdueShown = showAllOverdue ? overdueTenants : overdueTenants.slice(0, 5);
  const vacantUnits = data.units.filter(u => !data.tenants.some(t => t.unitId === u.id));

  const rentPanelCount = overdueTenants.length + allFollowUps;
  const totalAttention = rentPanelCount + courtItems.length + stipItems.length + recurringItems.length + appointmentItems.length + quickNoteItems.length + vacantUnits.length + hpdDue.length + dobDue.length + fdnyDue.length + otherViolationsDue.length + bossReminderItems.length;

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Dashboard</h1>
        <PrintButton label="Dashboard" />
      </div>

      <DashboardCalendar data={data} buildingName={buildingName} tenantName={tenantName} setTab={setTab} />

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
                            {t.balance && <div className="followup-item-note">Balance: ${t.balance}</div>}
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
                      {followUpEntries.map(({ tenant: t, followUp: f }) => (
                        <div className={`followup-item ${isInFollowUpWindow(f.date) ? "followup-item-due" : ""}`} key={f.id}>
                          <span className="pill pill-muted">{fmtDate(f.date)}</span>
                          <div className="followup-item-main">
                            <div className="followup-item-name">{t.name} <span className="row-muted">— {buildingName(t.buildingId)}</span></div>
                            {f.note && <div className="followup-item-note">{f.note}</div>}
                          </div>
                        </div>
                      ))}
                      {followUpEntries.length < tenantsWithFollowUps.flatMap(t => tenantFollowUps(t)).length && (
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
            label="HPD violations due or overdue" items={hpdDue} tab="violations" setTab={setTab}
            itemKey={v => v.id}
            renderItem={v => (
              <>
                <Flag date={v.cureDeadline} />
                <div className="followup-item-main">
                  <div className="followup-item-name">#{v.violationNumber} <span className="row-muted">— {buildingName(v.buildingId)}</span></div>
                </div>
              </>
            )}
          />

          <AttentionPanel
            icon={<AlertTriangle size={18} className="attention-icon" style={{ color: "var(--danger)" }} />}
            label="DOB violations due or overdue" items={dobDue} tab="violations" setTab={setTab}
            itemKey={v => v.id}
            renderItem={v => (
              <>
                <Flag date={v.cureDeadline} />
                <div className="followup-item-main">
                  <div className="followup-item-name">#{v.violationNumber} <span className="row-muted">— {buildingName(v.buildingId)}</span></div>
                </div>
              </>
            )}
          />

          <AttentionPanel
            icon={<AlertTriangle size={18} className="attention-icon" style={{ color: "var(--danger)" }} />}
            label="FDNY violations due or overdue" items={fdnyDue} tab="violations" setTab={setTab}
            itemKey={v => v.id}
            renderItem={v => (
              <>
                <Flag date={v.cureDeadline} />
                <div className="followup-item-main">
                  <div className="followup-item-name">#{v.violationNumber} <span className="row-muted">— {buildingName(v.buildingId)}</span></div>
                </div>
              </>
            )}
          />

          <AttentionPanel
            icon={<AlertTriangle size={18} className="attention-icon" style={{ color: "var(--danger)" }} />}
            label="Other violations due or overdue (DSNY, ECB, DEP, etc.)" items={otherViolationsDue} tab="violations" setTab={setTab}
            itemKey={v => v.id}
            renderItem={v => (
              <>
                <Flag date={v.cureDeadline} />
                <div className="followup-item-main">
                  <div className="followup-item-name">#{v.violationNumber} · {v.agency === "Other" ? v.otherAgency : v.agency} <span className="row-muted">— {buildingName(v.buildingId)}</span></div>
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
                <div className="followup-item-name">{n.text}{n.buildingId ? ` — ${buildingName(n.buildingId)}` : ""}</div>
                <div className="followup-item-note">{n.reminderDate ? `Reminder: ${fmtDate(n.reminderDate)}` : fmtDate(n.date)}</div>
              </div>
            )}
          />

          <AttentionPanel
            icon={<MessageSquare size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label="Boss reminders" items={bossReminderItems} tab="reminders" setTab={setTab}
            itemKey={r => r.id}
            renderItem={r => (
              <div className="followup-item-main">
                <div className="followup-item-name">{r.text}</div>
                <div className="followup-item-note">{fmtDate(r.dateRaised)}</div>
              </div>
            )}
          />

          <AttentionPanel
            icon={<Home size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label="Units with no tenant on file (vacant or a data gap)" items={vacantUnits} tab="buildings" setTab={setTab}
            itemKey={u => u.id}
            renderItem={u => (
              <div className="followup-item-main">
                <div className="followup-item-name">Unit {u.unitNumber || "—"} <span className="row-muted">— {buildingName(u.buildingId)}</span></div>
              </div>
            )}
            extraAction={
              confirmingCleanup ? (
                <>
                  <span className="row-muted" style={{ fontSize: 12 }}>Remove all {vacantUnits.length} empty units, across every building?</span>
                  <button className="btn-primary" style={{ background: "var(--danger)", borderColor: "var(--danger)" }} onClick={cleanupAllEmptyUnits}>Yes, clean up</button>
                  <button className="btn-ghost" onClick={() => setConfirmingCleanup(false)}>Cancel</button>
                </>
              ) : (
                <button className="btn-primary" onClick={() => setConfirmingCleanup(true)}>Clean up all empty units</button>
              )
            }
          />
        </>
      )}

      <h2 className="section-heading">By building</h2>
      {data.buildings.length === 0 ? (
        <EmptyState text="Add your buildings to see a per-building breakdown." />
      ) : (
        <div className="dash-grid">
          {data.buildings.map(b => {
            const bViolationsList = data.violations.filter(v => v.buildingId === b.id && !isViolationClosed(v));
            const bTenantsLateList = data.tenants.filter(t => t.buildingId === b.id && t.status !== "Current");
            const bWOList = data.workOrders.filter(w => w.buildingId === b.id && w.status !== "Done");
            const bCourtList = data.courtCases.filter(c => c.buildingId === b.id && !c.archived);
            const chips = [
              bViolationsList.length > 0 && { text: `${bViolationsList.length} open violations`, tone: "warn" },
              bTenantsLateList.length > 0 && { text: `${bTenantsLateList.length} tenants behind`, tone: "danger" },
              bWOList.length > 0 && { text: `${bWOList.length} open work orders`, tone: "muted" },
              bCourtList.length > 0 && { text: `${bCourtList.length} open cases`, tone: "danger" },
            ].filter(Boolean);
            const isExpanded = expandedBuilding === b.id;
            return (
              <div className={`dash-building-card ${chips.length > 0 ? "dash-building-clickable" : ""}`} key={b.id}
                onClick={() => chips.length > 0 && setExpandedBuilding(isExpanded ? null : b.id)}>
                <div className="dash-building-name">
                  {shortAddress(b.address)}
                  {chips.length > 0 && (isExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />)}
                </div>
                {chips.length === 0 ? (
                  <span className="dash-building-clear">All clear</span>
                ) : (
                  <div className="dash-building-chips">
                    {chips.map((c, i) => <span key={i} className={`pill pill-${c.tone}`}>{c.text}</span>)}
                  </div>
                )}
                {isExpanded && (
                  <div className="dash-building-detail">
                    {bTenantsLateList.map(t => (
                      <button key={t.id} className="dash-detail-item" onClick={(e) => { e.stopPropagation(); setTab("rent"); }}>
                        <span className="pill pill-danger">{t.status}</span>
                        <div className="followup-item-main"><div className="followup-item-name">{t.name}{t.balance ? ` — $${t.balance}` : ""}</div></div>
                      </button>
                    ))}
                    {bViolationsList.map(v => (
                      <button key={v.id} className="dash-detail-item" onClick={(e) => { e.stopPropagation(); setTab("violations"); }}>
                        <span className="pill pill-warn">{v.agency}</span>
                        <div className="followup-item-main"><div className="followup-item-name">#{v.violationNumber}{v.cureDeadline ? ` — cure by ${fmtDate(v.cureDeadline)}` : ""}</div></div>
                      </button>
                    ))}
                    {bCourtList.map(c => (
                      <button key={c.id} className="dash-detail-item" onClick={(e) => { e.stopPropagation(); setTab("court"); }}>
                        <span className="pill pill-danger">Court</span>
                        <div className="followup-item-main"><div className="followup-item-name">{tenantName(c.tenantId)}{c.nextCourtDate ? ` — ${fmtDate(c.nextCourtDate)}` : ""}</div></div>
                      </button>
                    ))}
                    {bWOList.map(w => (
                      <button key={w.id} className="dash-detail-item" onClick={(e) => { e.stopPropagation(); setTab("workorders"); }}>
                        <span className="pill pill-muted">{w.status}</span>
                        <div className="followup-item-main"><div className="followup-item-name">{w.description}</div></div>
                      </button>
                    ))}
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
  const [expandedUnit, setExpandedUnit] = useState(null);
  const [section, setSection] = useState("buildings");
  const [pendingDelete, setPendingDelete] = useState(null);
  const fileRef = useRef(null);

  const deleteBuilding = (buildingId) => {
    setData(d => ({
      ...d,
      buildings: d.buildings.filter(b => b.id !== buildingId),
      units: d.units.filter(u => u.buildingId !== buildingId),
      tenants: d.tenants.filter(t => t.buildingId !== buildingId),
    }));
    setPendingDelete(null);
  };
  const [pendingDeleteUnit, setPendingDeleteUnit] = useState(null);
  const deleteUnit = (unitId) => {
    setData(d => ({
      ...d,
      units: d.units.filter(u => u.id !== unitId),
      tenants: d.tenants.filter(t => t.unitId !== unitId),
    }));
    setPendingDeleteUnit(null);
    setExpandedUnit(null);
  };
  const [pendingCleanup, setPendingCleanup] = useState(null);
  const cleanupEmptyUnits = (buildingId) => {
    setData(d => ({
      ...d,
      units: d.units.filter(u => u.buildingId !== buildingId || d.tenants.some(t => t.unitId === u.id)),
    }));
    setPendingCleanup(null);
  };

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
        <button className={`chip ${section === "import" ? "chip-active" : ""}`} onClick={() => setSection("import")}>Import Contacts</button>
      </div>

      {section === "import" ? (
        <ImportSection data={data} setData={setData} buildingName={buildingName} allowedTypes={["contacts"]} />
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
              {(() => {
                const emptyCount = units.filter(u => !data.tenants.some(t => t.unitId === u.id)).length;
                return emptyCount > 0 && <span className="pill pill-warn">{emptyCount} empty</span>;
              })()}
              <div className="spacer" />
              {pendingDelete === b.id ? (
                <>
                  <span className="row-muted" style={{ fontSize: 12, color: data.tenants.some(t => t.buildingId === b.id && (data.courtCases || []).some(c => c.tenantId === t.id && !c.archived)) ? "var(--danger)" : undefined }}>
                    Delete building + {units.length} units + {data.tenants.filter(t => t.buildingId === b.id).length} tenants
                    {(() => {
                      const n = data.tenants.filter(t => t.buildingId === b.id && (data.courtCases || []).some(c => c.tenantId === t.id && !c.archived)).length;
                      return n > 0 ? ` — ${n} of them ${n === 1 ? "has" : "have"} an OPEN COURT CASE` : "";
                    })()}?
                  </span>
                  <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); deleteBuilding(b.id); }} style={{ color: "var(--danger)" }}>Yes, delete</button>
                  <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); setPendingDelete(null); }}>Cancel</button>
                </>
              ) : pendingCleanup === b.id ? (
                <>
                  <span className="row-muted" style={{ fontSize: 12 }}>Remove {units.filter(u => !data.tenants.some(t => t.unitId === u.id)).length} empty units from this building?</span>
                  <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); cleanupEmptyUnits(b.id); }} style={{ color: "var(--danger)" }}>Yes, clean up</button>
                  <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); setPendingCleanup(null); }}>Cancel</button>
                </>
              ) : (
                <>
                  {units.some(u => !data.tenants.some(t => t.unitId === u.id)) && (
                    <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); setPendingCleanup(b.id); }}>Clean up empty units</button>
                  )}
                  <IconBtn title="Edit" onClick={(e) => { e.stopPropagation(); setForm(b); }}><Pencil size={14} /></IconBtn>
                  <IconBtn title="Delete" danger onClick={(e) => { e.stopPropagation(); setPendingDelete(b.id); }}><Trash2 size={14} /></IconBtn>
                </>
              )}
            </div>
            {expanded === b.id && (
              <div className="list-card-body">
                {units.length === 0 && <div className="hint">No units added yet.</div>}
                {units.slice().sort((a, b2) => compareUnits(a.unitNumber, b2.unitNumber)).map(u => {
                  const tenants = data.tenants.filter(t => t.unitId === u.id);
                  const isOpen = expandedUnit === u.id;
                  return (
                    <div className="unit-block" key={u.id}>
                      <button className="unit-row" onClick={() => setExpandedUnit(isOpen ? null : u.id)}>
                        {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                        <span className="unit-row-number">Unit {u.unitNumber}</span>
                        <span className="unit-row-name">{tenants.map(t => t.name).filter(Boolean).join(", ") || "no tenant on file"}</span>
                        {tenants[0]?.status && tenants[0].status !== "Current" && (
                          <span className={`pill ${tenants[0].status === "Late" ? "pill-warn" : "pill-danger"}`}>{tenants[0].status}</span>
                        )}
                        {units.filter(u2 => u2.id !== u.id && u.unitNumber && (u2.unitNumber || "").toUpperCase() === u.unitNumber.toUpperCase()).length > 0 && (
                          <span className="pill pill-danger">⚠ Duplicate #</span>
                        )}
                      </button>
                      {isOpen && (
                        <div className="unit-detail">
                          <div className="unit-detail-row" style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 2 }}>
                            <span className="row-muted" style={{ fontSize: 12 }}>Unit number:</span>
                            <input
                              className="sheet-input" style={{ width: 90, border: "1px solid var(--border)", borderRadius: 4 }}
                              value={u.unitNumber} onChange={e => update("units", u.id, { unitNumber: e.target.value })}
                            />
                            {pendingDeleteUnit === u.id ? (
                              <>
                                <span className="row-muted" style={{ fontSize: 12, color: tenants.some(t => (data.courtCases || []).some(c => c.tenantId === t.id && !c.archived)) ? "var(--danger)" : undefined }}>
                                  Delete unit{tenants.length > 0 ? ` + ${tenants.length} tenant(s)` : ""}
                                  {tenants.some(t => (data.courtCases || []).some(c => c.tenantId === t.id && !c.archived)) ? " — has an OPEN COURT CASE" : ""}?
                                </span>
                                <button className="btn-ghost" onClick={() => deleteUnit(u.id)} style={{ color: "var(--danger)" }}>Yes, delete</button>
                                <button className="btn-ghost" onClick={() => setPendingDeleteUnit(null)}>Cancel</button>
                              </>
                            ) : (
                              <IconBtn title="Delete unit" danger onClick={() => setPendingDeleteUnit(u.id)}><Trash2 size={14} /></IconBtn>
                            )}
                          </div>
                          {units.filter(u2 => u2.id !== u.id && u.unitNumber && (u2.unitNumber || "").toUpperCase() === u.unitNumber.toUpperCase()).length > 0 && (
                            <div className="unit-detail-row" style={{ color: "var(--danger)", fontSize: 12, marginBottom: 6 }}>
                              ⚠ Another unit in this building already has number "{u.unitNumber}" — check for a duplicate.
                            </div>
                          )}
                          {tenants.length === 0 && <div className="hint">No tenant on file for this unit yet.</div>}
                          {tenants.map(t => (
                            <div key={t.id} className="unit-detail-tenant">
                              <div className="unit-detail-row"><strong>{t.name || "(no name on file)"}</strong></div>
                              {t.phone && <div className="unit-detail-row">Phone: {t.phone}</div>}
                              {t.email && <div className="unit-detail-row">Email: {t.email}</div>}
                              <div className="unit-detail-row">Balance: {t.balance ? `$${t.balance}` : "—"} · Status: <span className={`pill ${t.status === "Current" ? "pill-ok" : t.status === "Late" ? "pill-warn" : "pill-danger"}`}>{t.status || "Current"}</span></div>
                              {tenantFollowUps(t).length > 0 && (
                                <div className="unit-detail-row">
                                  Follow-up{tenantFollowUps(t).length > 1 ? "s" : ""}: {tenantFollowUps(t).map(f => `${fmtDate(f.date)}${f.note ? ` — ${f.note}` : ""}`).join("; ")}
                                </div>
                              )}
                              {(Array.isArray(t.notes) ? t.notes : []).length > 0 && (
                                <div className="unit-detail-row row-muted">Latest note: {t.notes[t.notes.length - 1].text}</div>
                              )}
                            </div>
                          ))}
                        </div>
                      )}
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
  const [noteFor, setNoteFor] = useState(null);
  const [noteText, setNoteText] = useState("");
  const [followFor, setFollowFor] = useState(null);
  const [newFollowDate, setNewFollowDate] = useState("");
  const [newFollowNote, setNewFollowNote] = useState("");
  const [statusFilter, setStatusFilter] = useState("All");
  const [buildingFilter, setBuildingFilter] = useState("All");
  const [section, setSection] = useState("sheet");

  const inCourt = (tenantId) => (data.courtCases || []).some(c => c.tenantId === tenantId && !c.archived);

  const [newTenant, setNewTenant] = useState(null);

  const openNewTenant = () => setNewTenant({
    buildingId: data.buildings[0]?.id || "", unitId: "", name: "", phone: "", balance: "", status: "Current",
  });
  const saveNewTenant = () => {
    if (!newTenant.buildingId || !newTenant.unitId) return;
    add("tenants", { ...newTenant, email: "", notes: [], followUps: [] });
    setNewTenant(null);
  };

  const exportCSV = () => {
    const rows = visibleTenants.map(t => ({
      Building: buildingName(t.buildingId),
      Unit: unitOf(t),
      Tenant: t.name,
      Phone: t.phone,
      Balance: t.balance,
      Status: t.status,
      "Latest Note": notesArr(t).length ? notesArr(t)[notesArr(t).length - 1].text : "",
      "Follow-ups": tenantFollowUps(t).map(f => f.date).join("; "),
    }));
    const csv = Papa.unparse(rows);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "rent-collection.csv";
    a.click();
    URL.revokeObjectURL(url);
  };

  // One combined log per tenant now — old data may still have a separate
  // messageLog array from before; fold it in so nothing gets lost, but new
  // entries only ever get added to "notes" going forward.
  const notesArr = (t) => {
    const n = Array.isArray(t.notes) ? t.notes : (t.notes ? [{ date: todayISO(), text: t.notes }] : []);
    const m = Array.isArray(t.messageLog) ? t.messageLog : [];
    return [...n, ...m].sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  };

  const addNote = (tenantId) => {
    if (!noteText.trim()) return;
    const t = data.tenants.find(x => x.id === tenantId);
    const existingNotes = Array.isArray(t.notes) ? t.notes : (t.notes ? [{ date: todayISO(), text: t.notes }] : []);
    update("tenants", tenantId, { notes: [...existingNotes, { date: todayISO(), text: noteText }] });
    setNoteText(""); setNoteFor(null);
  };

  const addFollowUp = (tenantId) => {
    if (!newFollowDate) return;
    const t = data.tenants.find(x => x.id === tenantId);
    update("tenants", tenantId, { followUps: [...tenantFollowUps(t), { id: uid(), date: newFollowDate, note: newFollowNote }] });
    setNewFollowDate(""); setNewFollowNote("");
  };
  const removeFollowUp = (tenantId, followUpId) => {
    const t = data.tenants.find(x => x.id === tenantId);
    update("tenants", tenantId, { followUps: tenantFollowUps(t).filter(f => f.id !== followUpId) });
  };

  const byBuilding = t => buildingFilter === "All" || t.buildingId === buildingFilter;
  const unitOf = t => data.units.find(u => u.id === t.unitId)?.unitNumber || "";
  const visibleTenants = (statusFilter === "Follow-ups"
    ? data.tenants.filter(t => tenantFollowUps(t).length > 0).slice().sort((a, b) => (earliestFollowUpDate(a) || "").localeCompare(earliestFollowUpDate(b) || ""))
    : data.tenants.filter(t => statusFilter === "All" || t.status === statusFilter)
        .slice()
        .sort((a, b) => buildingName(a.buildingId).localeCompare(buildingName(b.buildingId)) || compareUnits(unitOf(a), unitOf(b)))
  ).filter(byBuilding);

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Rent Collection</h1>
        <div className="page-actions">
          <PrintButton label="Rent Collection" />
          <button className="btn-ghost" onClick={exportCSV}><Download size={14} /> Export CSV</button>
          <button className="btn-primary" onClick={openNewTenant}><Plus size={14} /> Add tenant</button>
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
      {newTenant && (
        <div className="form-panel">
          <Field label="Building">
            <select value={newTenant.buildingId} onChange={e => setNewTenant({ ...newTenant, buildingId: e.target.value, unitId: "" })}>
              <option value="">—</option>
              {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
            </select>
          </Field>
          <Field label="Unit">
            <select value={newTenant.unitId} onChange={e => setNewTenant({ ...newTenant, unitId: e.target.value })}>
              <option value="">—</option>
              {data.units.filter(u => u.buildingId === newTenant.buildingId).sort((a, b) => compareUnits(a.unitNumber, b.unitNumber)).map(u => <option key={u.id} value={u.id}>{u.unitNumber}</option>)}
            </select>
          </Field>
          <Field label="Tenant name"><input value={newTenant.name} onChange={e => setNewTenant({ ...newTenant, name: e.target.value })} /></Field>
          <Field label="Phone"><input value={newTenant.phone} onChange={e => setNewTenant({ ...newTenant, phone: e.target.value })} /></Field>
          <Field label="Balance">
            <div className="balance-input-wrap">
              <span className="balance-dollar">$</span>
              <input className="balance-input" value={newTenant.balance} onChange={e => setNewTenant({ ...newTenant, balance: e.target.value })} />
            </div>
          </Field>
          <Field label="Status">
            <select value={newTenant.status} onChange={e => setNewTenant({ ...newTenant, status: e.target.value })}>
              {RENT_STATUSES.map(s => <option key={s}>{s}</option>)}
            </select>
          </Field>
          <p className="hint" style={{ gridColumn: "1 / -1" }}>Building and unit can only be set here or changed later from the Buildings tab — not editable in the sheet, so tenants never accidentally end up on the wrong building.</p>
          <div className="form-actions">
            <button className="btn-primary" onClick={saveNewTenant} disabled={!newTenant.buildingId || !newTenant.unitId}>Save</button>
            <button className="btn-ghost" onClick={() => setNewTenant(null)}>Cancel</button>
          </div>
        </div>
      )}
      <div className="filter-row">
        {["All", ...RENT_STATUSES].map(s => (
          <button key={s} className={`chip ${statusFilter === s ? "chip-active" : ""}`} onClick={() => setStatusFilter(s)}>{s}</button>
        ))}
        <button className={`chip ${statusFilter === "Follow-ups" ? "chip-active" : ""}`} onClick={() => setStatusFilter("Follow-ups")}>Follow-ups</button>
      </div>
      <div className="filter-row">
        <button className={`chip ${buildingFilter === "All" ? "chip-active" : ""}`} onClick={() => setBuildingFilter("All")}>All buildings</button>
        {data.buildings.map(b => (
          <button key={b.id} className={`chip ${buildingFilter === b.id ? "chip-active" : ""}`} onClick={() => setBuildingFilter(b.id)}>{shortAddress(b.address)}</button>
        ))}
      </div>

      {data.tenants.length === 0 ? (
        <EmptyState text="No tenants yet — import RIS data or add a tenant." />
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
                    <td className="sheet-readonly">{buildingName(t.buildingId)}</td>
                    <td className="sheet-readonly">{unitOf(t) || "—"}</td>
                    <td>
                      <div className="sheet-name-cell">
                        <input className="sheet-input" value={t.name} onChange={e => update("tenants", t.id, { name: e.target.value })} />
                        {inCourt(t.id) && <span className="pill pill-danger sheet-court-pill" title="Active court case — no need to independently follow up">In Court</span>}
                      </div>
                    </td>
                    <td><input className="sheet-input" value={t.phone} onChange={e => update("tenants", t.id, { phone: e.target.value })} /></td>
                    <td className="sheet-col-balance">
                      <div className="balance-input-wrap">
                        <span className="balance-dollar">$</span>
                        <input className="sheet-input balance-input" value={t.balance} onChange={e => update("tenants", t.id, { balance: e.target.value })} />
                      </div>
                    </td>
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
                      <button className="sheet-follow-btn" onClick={() => setFollowFor(followFor === t.id ? null : t.id)} title="Follow-ups">
                        <CalendarClock size={14} />
                        {tenantFollowUps(t).length > 0 && (
                          <span className="sheet-follow-date">
                            {fmtDate(earliestFollowUpDate(t))}{tenantFollowUps(t).length > 1 ? ` +${tenantFollowUps(t).length - 1}` : ""}
                          </span>
                        )}
                      </button>
                    </td>
                    <td className="sheet-actions">
                      <IconBtn title="Notes" onClick={() => setNoteFor(noteFor === t.id ? null : t.id)}><Pencil size={14} /></IconBtn>
                      <IconBtn title="Delete" danger onClick={() => {
                        if (inCourt(t.id) && !window.confirm(`${t.name || "This tenant"} has an open court case. Delete anyway? The case will stay but lose its link to this tenant.`)) return;
                        remove("tenants", t.id);
                      }}><Trash2 size={14} /></IconBtn>
                    </td>
                  </tr>
                  {followFor === t.id && (
                    <tr className="sheet-expand-row">
                      <td colSpan={9}>
                        <strong>Follow-ups</strong>
                        <div className="inline-form">
                          <input type="date" value={newFollowDate} onChange={e => setNewFollowDate(e.target.value)} />
                          <input placeholder="What's this follow-up about?" value={newFollowNote} onChange={e => setNewFollowNote(e.target.value)} style={{ flex: 1 }} />
                          <button className="btn-primary" onClick={() => addFollowUp(t.id)}>Add follow-up</button>
                        </div>
                        {tenantFollowUps(t).length === 0 && <div className="hint">No follow-ups yet.</div>}
                        {tenantFollowUps(t).slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).map((f) => (
                          <div key={f.id} className="row row-muted" style={{ display: "flex", alignItems: "center", gap: 8 }}>
                            <span>{fmtDate(f.date)}{f.note ? ` — ${f.note}` : ""}</span>
                            <button className="checklist-remove" onClick={() => removeFollowUp(t.id, f.id)} title="Remove"><X size={12} /></button>
                          </div>
                        ))}
                      </td>
                    </tr>
                  )}
                  {noteFor === t.id && (
                    <tr className="sheet-expand-row">
                      <td colSpan={9}>
                        <strong>Notes</strong>
                        <div className="inline-form">
                          <input placeholder="Payment plans, disputes, calls, texts, anything…" value={noteText} onChange={e => setNoteText(e.target.value)} onKeyDown={e => e.key === "Enter" && addNote(t.id)} />
                          <button className="btn-primary" onClick={() => addNote(t.id)}>Add note</button>
                        </div>
                        {notesArr(t).length === 0 && <div className="hint">No notes yet.</div>}
                        {notesArr(t).slice().reverse().map((n, i) => (
                          <div key={i} className="row row-muted">{fmtDate(n.date)} — {n.text}</div>
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
  { key: "arrears", label: "Aged Arrears", hint: "Updates balance, status, and tenant names." },
  { key: "contacts", label: "Telephone / Email List", hint: "Updates phone and email." },
];

function buildImportDiff(type, parsedEntries, data, buildingId) {
  const unitsForBuilding = data.units.filter(u => u.buildingId === buildingId);
  const tenantByUnit = unitId => data.tenants.find(t => t.unitId === unitId);
  const hasActiveCourtCase = tenantId => (data.courtCases || []).some(c => c.tenantId === tenantId && !c.archived);
  const changes = [];

  parsedEntries.forEach(entry => {
    const unit = unitsForBuilding.find(u => (u.unitNumber || "").toUpperCase() === entry.apt.toUpperCase());
    let fields = {};
    if (type === "arrears") fields = { balance: entry.balance, status: entry.status, name: entry.name };
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
          // If this tenant was already flagged Late/In Arrears before this import, don't
          // silently overwrite whatever's being tracked for them — require an explicit
          // approve. Tenants currently "Current" (or brand-new changes) apply automatically.
          const priorStatus = tenant.status || "Current";
          const needsApproval = type === "arrears" && priorStatus !== "Current";
          changes.push({
            apt: entry.apt, unitId: unit.id, tenantId: tenant.id, isNew: false,
            fields: diffFields, before: Object.fromEntries(Object.keys(diffFields).map(k => [k, tenant[k] || "—"])),
            name: tenant.name || entry.name, movedOut: entry.movedOut, needsReview: entry.needsReview,
            priorStatus, needsApproval, approved: !needsApproval,
            inCourt: hasActiveCourtCase(tenant.id),
          });
        }
      } else {
        changes.push({ apt: entry.apt, unitId: unit.id, tenantId: null, isNew: true, fields: { ...fields, name: fields.name || entry.name || "" }, name: entry.name, movedOut: entry.movedOut, needsReview: entry.needsReview, needsApproval: false, approved: true, inCourt: false });
      }
    } else {
      changes.push({ apt: entry.apt, unitId: null, tenantId: null, isNew: true, newUnit: true, fields: { ...fields, name: fields.name || entry.name || "" }, name: entry.name, movedOut: entry.movedOut, needsReview: entry.needsReview, needsApproval: false, approved: true, inCourt: false });
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

function ImportBlock({ type, data, setData, buildingName }) {
  const typeInfo = IMPORT_TYPES.find(t => t.key === type);
  const [rawText, setRawText] = useState("");
  const [showPaste, setShowPaste] = useState(false);
  const [preview, setPreview] = useState(null);
  const [pdfStatus, setPdfStatus] = useState(null); // null | "reading" | "error"
  const [pdfError, setPdfError] = useState("");
  const pdfInputRef = useRef(null);

  const runParse = (text) => {
    if (!text.trim()) return;
    const parser = type === "arrears" ? parseArrearsText : type === "directory" ? parseDirectoryText : parseContactsText;
    const entries = parser(text);
    const header = parseBuildingHeader(text);
    const existing = data.buildings.find(b =>
      (header.propCode && b.risPropCode === header.propCode) ||
      (header.address && (b.address || "").trim().toLowerCase() === header.address.trim().toLowerCase())
    );
    // If no matching building exists yet, diff against a placeholder id so every
    // row in the report shows as "new" — the real building gets created on confirm.
    const diff = buildImportDiff(type, entries, data, existing ? existing.id : "__pending__");
    setPreview({ ...diff, parsedCount: entries.length, header, matchedBuildingId: existing ? existing.id : null });
  };

  const handlePdfUpload = async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setPdfStatus("reading");
    setPdfError("");
    setPreview(null);
    try {
      const text = await extractPdfText(file);
      setRawText(text);
      setPdfStatus(null);
      runParse(text);
    } catch (err) {
      setPdfStatus("error");
      setPdfError("Couldn't read that PDF automatically — use \"paste the text instead\" below (open the PDF, Ctrl/Cmd+A, Ctrl/Cmd+C, then paste).");
      setShowPaste(true);
    }
  };

  const toggleApprove = (idx) => {
    setPreview(p => ({
      ...p,
      changes: p.changes.map((c, i) => i === idx ? { ...c, approved: !c.approved } : c),
    }));
  };

  const confirm = () => {
    if (!preview) return;
    setData(d => {
      const next = { ...d, buildings: [...d.buildings], units: [...d.units], tenants: [...d.tenants], importHistory: [...d.importHistory] };
      let buildingId = preview.matchedBuildingId;
      if (!buildingId) {
        const newBuilding = {
          id: uid(),
          address: preview.header.address || "Unknown address (from import)",
          risPropCode: preview.header.propCode || "",
          notes: "",
        };
        next.buildings.push(newBuilding);
        buildingId = newBuilding.id;
      }
      const applied = preview.changes.filter(ch => !ch.needsApproval || ch.approved);
      applied.forEach(ch => {
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
          updated: applied.filter(c => !c.isNew).length,
          added: applied.filter(c => c.isNew).length,
          missing: preview.missing.length,
          skipped: preview.changes.length - applied.length,
          details: applied.map(c => ({ apt: c.apt, name: c.name, isNew: c.isNew, fields: c.fields })),
          missingDetails: preview.missing,
        },
        ...next.importHistory,
      ];
      return next;
    });
    setPreview(null);
    setRawText("");
    setShowPaste(false);
  };

  return (
    <div className="import-block">
      <div className="import-block-head">
        <div className="import-block-title">{typeInfo.label}</div>
        <button className="btn-primary" type="button" onClick={() => pdfInputRef.current.click()} disabled={pdfStatus === "reading"}>
          <Upload size={14} /> {pdfStatus === "reading" ? "Reading PDF…" : `Upload ${typeInfo.label} PDF`}
        </button>
        <input ref={pdfInputRef} type="file" accept="application/pdf" hidden onChange={handlePdfUpload} />
      </div>
      <p className="hint">{typeInfo.hint} The building is detected automatically from the report and created if it doesn't exist yet.</p>
      {pdfStatus === "error" && <div className="hint" style={{ color: "var(--danger)" }}>{pdfError}</div>}

      <button className="btn-ghost" type="button" onClick={() => setShowPaste(s => !s)}>
        {showPaste ? "Hide paste option" : "Or paste the text instead"}
      </button>
      {showPaste && (
        <div style={{ marginTop: 8 }}>
          <textarea rows={6} value={rawText} onChange={e => { setRawText(e.target.value); setPreview(null); }} placeholder="Open the PDF, select all, copy, paste here…" />
          <div className="form-actions" style={{ marginTop: 8 }}>
            <button className="btn-primary" onClick={() => runParse(rawText)} disabled={!rawText.trim()}>Parse & preview</button>
          </div>
        </div>
      )}

      {preview && (
        <div className="import-preview">
          {preview.header.address && (
            <div className="hint" style={{ marginBottom: 8 }}>
              Detected: <strong>{preview.header.address}</strong> — {preview.matchedBuildingId ? "matched to an existing building" : "will create this as a new building"}
            </div>
          )}
          <div className="import-preview-summary">
            {preview.changes.filter(c => c.isNew).length} new · {preview.changes.filter(c => !c.isNew && !c.needsApproval).length} to update
            {preview.changes.some(c => c.needsApproval) && ` · ${preview.changes.filter(c => c.needsApproval).length} already flagged behind — needs your approval`}
            {type === "arrears" && ` · ${preview.missing.length} not in this file`}
            {" "}({preview.parsedCount} rows read)
          </div>
          {preview.changes.length === 0 && preview.missing.length === 0 && (
            <div className="hint">No changes found — everything already matches.</div>
          )}
          {preview.changes.map((c, i) => (
            <div className={`import-row ${c.needsApproval ? "import-row-approval" : ""}`} key={i}>
              <span className="pill pill-muted">{c.apt}</span>
              <span className="import-row-name">{c.name}</span>
              {c.isNew && <span className="pill pill-warn">New</span>}
              {c.movedOut && <span className="pill pill-danger">Moved out (flagged)</span>}
              {c.needsReview && <span className="pill pill-warn">⚠ Review — shared unit, verify names</span>}
              {c.inCourt && <span className="pill pill-danger">In Court</span>}
              <div className="import-row-fields">
                {Object.entries(c.fields).map(([k, v]) => (
                  <span key={k} className="import-field">{k}: {c.before?.[k] ? `${c.before[k]} → ` : ""}{String(v)}</span>
                ))}
              </div>
              {c.needsApproval && (
                <label className="import-approve">
                  <input type="checkbox" checked={c.approved} onChange={() => toggleApprove(i)} />
                  Already marked "{c.priorStatus}" — update to this?
                </label>
              )}
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
    </div>
  );
}

function ImportSection({ data, setData, buildingName, allowedTypes }) {
  const availableTypes = IMPORT_TYPES.filter(t => !allowedTypes || allowedTypes.includes(t.key));
  const [expandedEntry, setExpandedEntry] = useState(null);
  const history = (data.importHistory || []).filter(h => !allowedTypes || allowedTypes.includes(h.type));

  return (
    <div>
      {availableTypes.map(t => (
        <ImportBlock key={t.key} type={t.key} data={data} setData={setData} buildingName={buildingName} />
      ))}

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
            {h.skipped > 0 && <span className="pill pill-muted">{h.skipped} not approved</span>}
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
          <button className="btn-primary" onClick={() => setForm({ buildingId: data.buildings[0]?.id || "", unitId: "", vendorId: "", description: "", status: "Open", priority: "Routine", dateOpened: todayISO() })}>
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
            <select value={form.buildingId} onChange={e => setForm({ ...form, buildingId: e.target.value, unitId: "" })}>
              <option value="">—</option>
              {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
            </select>
          </Field>
          <Field label="Apt # (optional)">
            <select value={form.unitId || ""} onChange={e => setForm({ ...form, unitId: e.target.value })}>
              <option value="">Whole building</option>
              {data.units.filter(u => u.buildingId === form.buildingId).sort((a, b) => compareUnits(a.unitNumber, b.unitNumber)).map(u => <option key={u.id} value={u.id}>{unitOptionLabel(u, data.tenants)}</option>)}
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
            {w.unitId && <span className="pill pill-accent">Apt {data.units.find(u => u.id === w.unitId)?.unitNumber || "—"}</span>}
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
  const [dueFilter, setDueFilter] = useState("all");
  const [noteFor, setNoteFor] = useState(null);
  const [noteText, setNoteText] = useState("");
  const [expandedRow, setExpandedRow] = useState(null);
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
    buildingId: data.buildings[0]?.id || "", unitId: "", violationNumber: "",
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
    if (dueFilter !== "all") {
      const maxDays = dueFilter === "24h" ? 1 : dueFilter === "1w" ? 7 : 10;
      list = list.filter(v => { const d = daysUntil(v.cureDeadline); return d !== null && d <= maxDays; });
    }
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
      {view === "active" && (
        <div className="filter-row">
          <button className={`chip ${dueFilter === "all" ? "chip-active" : ""}`} onClick={() => setDueFilter("all")}>All due dates</button>
          <button className={`chip ${dueFilter === "24h" ? "chip-active" : ""}`} onClick={() => setDueFilter("24h")}>Cure due ≤ 24 hrs</button>
          <button className={`chip ${dueFilter === "1w" ? "chip-active" : ""}`} onClick={() => setDueFilter("1w")}>Cure due ≤ 1 week</button>
          <button className={`chip ${dueFilter === "10d" ? "chip-active" : ""}`} onClick={() => setDueFilter("10d")}>Cure due ≤ 10 days</button>
        </div>
      )}

      {form && (
        <div className="form-panel">
          <Field label="Building">
            <select value={form.buildingId} onChange={e => setForm({ ...form, buildingId: e.target.value, unitId: "" })}>
              <option value="">—</option>
              {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
            </select>
          </Field>
          <Field label="Apt # (optional)">
            <select value={form.unitId || ""} onChange={e => setForm({ ...form, unitId: e.target.value })}>
              <option value="">Whole building</option>
              {data.units.filter(u => u.buildingId === form.buildingId).sort((a, b) => compareUnits(a.unitNumber, b.unitNumber)).map(u => <option key={u.id} value={u.id}>{unitOptionLabel(u, data.tenants)}</option>)}
            </select>
          </Field>
          <Field label="Violation #"><input value={form.violationNumber} onChange={e => setForm({ ...form, violationNumber: e.target.value })} /></Field>

          {agency === "HPD" && (
            <>
              <Field label="Class"><input value={form.class} onChange={e => setForm({ ...form, class: e.target.value })} placeholder="A / B / C" /></Field>
              <Field label="Description"><textarea value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} /></Field>
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
        const isOpen = expandedRow === v.id;
        return (
          <div className={`list-card ${flag === "overdue" ? "list-card-danger" : flag === "soon" ? "list-card-warn" : ""}`} key={v.id}>
            <div className="list-card-head" onClick={() => setExpandedRow(isOpen ? null : v.id)} style={{ cursor: "pointer", alignItems: "flex-start" }}>
              {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
              {v.unitId && <span className="pill pill-accent">Apt {data.units.find(u => u.id === v.unitId)?.unitNumber || "—"}</span>}
              <div className="violation-title-group">
                <div className="list-card-title">#{v.violationNumber}</div>
                {v.description && <div className="violation-desc-preview">{v.description}</div>}
              </div>
              <span className={`pill ${isClosed(v) ? "pill-ok" : "pill-muted"}`}>{v.status}</span>
              <span className="pill pill-muted">{buildingName(v.buildingId)}</span>
              {agency === "HPD" && v.vendorId && <span className="pill pill-muted">{vendorName(v.vendorId)}</span>}
              {agency === "DSNY" && v.fineAmount && <span className="pill pill-muted">{v.fineAmount}</span>}
              {agency === "Other" && v.otherAgency && <span className="pill pill-muted">{v.otherAgency}</span>}
              {agency === "Other" && v.company && <span className="pill pill-muted">{v.company}</span>}
              {agency !== "DSNY" && view === "active" && <Flag date={v.cureDeadline} />}
              <div className="spacer" />
              {agency === "HPD" && view === "active" && (
                <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); update("violations", v.id, { status: "Certified" }); }}>
                  Mark Certified
                </button>
              )}
              <IconBtn title="Edit" onClick={(e) => { e.stopPropagation(); setForm(v); }}><Pencil size={14} /></IconBtn>
              <IconBtn title="Delete" danger onClick={(e) => { e.stopPropagation(); remove("violations", v.id); }}><Trash2 size={14} /></IconBtn>
            </div>
            {isOpen && (
              <div className="list-card-body">
                {v.class && <div className="row"><strong>Class:</strong> {v.class}</div>}
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
            )}
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
    if (!form.tenantId) return; // a case with no tenant is a silent data-quality trap
    // Only patch the fields this form actually manages — checklist, documents, and
    // archived get edited from their own controls elsewhere on the card, so if we
    // sent the whole `form` object here, saving this form after toggling a checklist
    // item (both are on-screen at once) would silently revert that checklist change
    // back to whatever it was when this form was opened.
    const fields = {
      tenantId: form.tenantId, buildingId: form.buildingId, unitId: form.unitId, caseNumber: form.caseNumber,
      stage: form.stage, nextCourtDate: form.nextCourtDate, result: form.result,
      stipulationTerms: form.stipulationTerms || "", nextPaymentDue: form.nextPaymentDue || "",
    };
    if (form.id) update("courtCases", form.id, fields);
    else add("courtCases", {
      ...fields, archived: false, documents: [],
      checklist: DEFAULT_ATTORNEY_CHECKLIST.map(label => ({ id: uid(), label, checked: false })),
    });
    setForm(null);
  };

  // Old cases were saved before unitId existed — recover it from the tenant on
  // file so editing an old case still starts from the right apt.
  const openEdit = (c) => setForm({ ...c, unitId: c.unitId || data.tenants.find(t => t.id === c.tenantId)?.unitId || "" });

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
          <button className="btn-primary" onClick={() => setForm({ tenantId: "", buildingId: "", unitId: "", caseNumber: "", nextCourtDate: "", result: "Pending", stage: CASE_STAGES[0], stipulationTerms: "", nextPaymentDue: "" })}>
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
          <Field label="Building">
            <select value={form.buildingId} onChange={e => setForm({ ...form, buildingId: e.target.value, unitId: "", tenantId: "" })}>
              <option value="">—</option>
              {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
            </select>
          </Field>
          <Field label="Apt # — Tenant">
            <select
              value={form.tenantId} disabled={!form.buildingId}
              onChange={e => {
                const t = data.tenants.find(x => x.id === e.target.value);
                setForm({ ...form, tenantId: e.target.value, unitId: t ? t.unitId : "" });
              }}
            >
              <option value="">{form.buildingId ? "—" : "Pick a building first"}</option>
              {data.tenants
                .filter(t => t.buildingId === form.buildingId)
                .slice()
                .sort((a, b) => compareUnits(data.units.find(u => u.id === a.unitId)?.unitNumber, data.units.find(u => u.id === b.unitId)?.unitNumber))
                .map(t => (
                  <option key={t.id} value={t.id}>
                    {data.units.find(u => u.id === t.unitId)?.unitNumber || "—"} — {t.name || "(no name on file)"}
                  </option>
                ))}
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
            <p className="hint" style={{ gridColumn: "1 / -1", marginTop: -6 }}>Update the "Court date" field above to the new date.</p>
          )}
          {form.result === "Stipulation (payment plan)" && (
            <>
              <Field label="Payment plan terms"><textarea value={form.stipulationTerms} onChange={e => setForm({ ...form, stipulationTerms: e.target.value })} placeholder="e.g. $200/month starting 10/1 on top of current rent, for 6 months" /></Field>
              <Field label="Next payment due date"><input type="date" value={form.nextPaymentDue} onChange={e => setForm({ ...form, nextPaymentDue: e.target.value })} /></Field>
            </>
          )}
          <div className="form-actions">
            <button className="btn-primary" onClick={submit} disabled={!form.tenantId}>Save</button>
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
            {c.unitId && <span className="pill pill-accent">Apt {data.units.find(u => u.id === c.unitId)?.unitNumber || "—"}</span>}
            <div className="list-card-title">{tenantName(c.tenantId)} {c.caseNumber && `· Docket #${c.caseNumber}`}</div>
            {c.stage && <span className="pill pill-muted">{c.stage}</span>}
            <span className="pill pill-muted">{c.result}</span>
            <span className="pill pill-muted">{buildingName(c.buildingId)}</span>
            {c.nextCourtDate && !c.archived && <Flag date={c.nextCourtDate} label="court date" />}
            <div className="spacer" />
            <IconBtn title="Edit" onClick={() => openEdit(c)}><Pencil size={14} /></IconBtn>
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
                <button className="btn-primary" style={{ marginTop: 10 }} onClick={() => setDetailsFor(null)}>Done</button>
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
              {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
            </select>
          </Field>
          <Field label="Unit (optional)">
            <select value={form.unitId} onChange={e => setForm({ ...form, unitId: e.target.value })}>
              <option value="">General / whole building</option>
              {data.units.filter(u => u.buildingId === form.buildingId).sort((a, b) => compareUnits(a.unitNumber, b.unitNumber)).map(u => <option key={u.id} value={u.id}>{unitOptionLabel(u, data.tenants)}</option>)}
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
                <a className="btn-ghost" href={icsFor(`${a.type} — ${buildingName(a.buildingId)}`, a.date, a.notes, a.timeFrom, a.timeTo)} download={`${(a.type || "appointment").replace(/\s/g, "-")}.ics`}>
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

  useEffect(() => {
    if (data.buildings.length === 0) return;
    if (!data.buildings.some(b => b.id === buildingId)) setBuildingId(data.buildings[0].id);
  }, [data.buildings, buildingId]);

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
          <button key={b.id} className={`chip ${buildingId === b.id ? "chip-active" : ""}`} onClick={() => setBuildingId(b.id)}>{shortAddress(b.address)}</button>
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
  const [editingId, setEditingId] = useState(null);
  const [editText, setEditText] = useState("");

  const submit = () => {
    if (!text.trim()) return;
    add("bossReminders", { text, dateRaised: todayISO(), status: "Open" });
    setText("");
  };
  const startEdit = (r) => { setEditingId(r.id); setEditText(r.text); };
  const saveEdit = (id) => {
    if (!editText.trim()) return;
    update("bossReminders", id, { text: editText });
    setEditingId(null);
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
            {editingId === r.id ? (
              <>
                <input className="sheet-input" style={{ flex: 1, border: "1px solid var(--border)", borderRadius: 4 }} value={editText} onChange={e => setEditText(e.target.value)} onKeyDown={e => e.key === "Enter" && saveEdit(r.id)} autoFocus />
                <button className="btn-primary" onClick={() => saveEdit(r.id)}>Save</button>
                <button className="btn-ghost" onClick={() => setEditingId(null)}>Cancel</button>
              </>
            ) : (
              <>
                <div className={`list-card-title ${r.status === "Done" ? "strike" : ""}`}>{r.text}</div>
                <span className="pill pill-muted">{fmtDate(r.dateRaised)}</span>
                <div className="spacer" />
                <IconBtn title="Edit" onClick={() => startEdit(r)}><Pencil size={14} /></IconBtn>
                <IconBtn title="Delete" danger onClick={() => remove("bossReminders", r.id)}><Trash2 size={14} /></IconBtn>
              </>
            )}
          </div>
        </div>
      ))}
    </div>
  );
}

/* ============================== quick notes ============================== */

function QuickNotesTab({ data, add, update, remove, buildingName }) {
  const [text, setText] = useState("");
  const [reminderDate, setReminderDate] = useState("");
  const [buildingId, setBuildingId] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editText, setEditText] = useState("");
  const [showDone, setShowDone] = useState(false);

  const submit = () => {
    if (!text.trim()) return;
    add("quickNotes", { text, date: todayISO(), done: false, reminderDate, buildingId });
    setText(""); setReminderDate(""); setBuildingId("");
  };
  const startEdit = (n) => { setEditingId(n.id); setEditText(n.text); };
  const saveEdit = (id) => {
    if (!editText.trim()) return;
    update("quickNotes", id, { text: editText });
    setEditingId(null);
  };

  const allNotes = (data.quickNotes || []).slice().reverse();
  const notes = showDone ? allNotes : allNotes.filter(n => !n.done);

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Quick Notes</h1>
        <PrintButton label="Quick Notes" />
      </div>
      <p className="hint">A checklist for anything you need to jot down fast — check it off when it's handled, optionally set a reminder date, and tie it to a building if it's related to one.</p>
      <div className="form-panel">
        <Field label="Note"><input placeholder="Jot something down…" value={text} onChange={e => setText(e.target.value)} onKeyDown={e => e.key === "Enter" && submit()} /></Field>
        <Field label="Reminder date (optional)"><input type="date" value={reminderDate} onChange={e => setReminderDate(e.target.value)} /></Field>
        <Field label="Building (optional)">
          <select value={buildingId} onChange={e => setBuildingId(e.target.value)}>
            <option value="">—</option>
            {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
          </select>
        </Field>
        <div className="form-actions">
          <button className="btn-primary" onClick={submit}>Add</button>
        </div>
      </div>
      <div className="filter-row">
        <button className={`chip ${!showDone ? "chip-active" : ""}`} onClick={() => setShowDone(false)}>Open</button>
        <button className={`chip ${showDone ? "chip-active" : ""}`} onClick={() => setShowDone(true)}>All (incl. checked off)</button>
      </div>
      {notes.length === 0 && <EmptyState text="Nothing here." />}
      {notes.map(n => (
        <div className="list-card" key={n.id}>
          <div className="list-card-head">
            <input type="checkbox" checked={!!n.done} onChange={e => update("quickNotes", n.id, { done: e.target.checked })} />
            {editingId === n.id ? (
              <>
                <input className="sheet-input" style={{ flex: 1, border: "1px solid var(--border)", borderRadius: 4 }} value={editText} onChange={e => setEditText(e.target.value)} onKeyDown={e => e.key === "Enter" && saveEdit(n.id)} autoFocus />
                <button className="btn-primary" onClick={() => saveEdit(n.id)}>Save</button>
                <button className="btn-ghost" onClick={() => setEditingId(null)}>Cancel</button>
              </>
            ) : (
              <>
                <div className={`list-card-title ${n.done ? "strike" : ""}`}>{n.text}</div>
                {n.reminderDate && <span className="pill pill-warn">Reminder: {fmtDate(n.reminderDate)}</span>}
                {n.buildingId && <span className="pill pill-muted">{buildingName(n.buildingId)}</span>}
                <span className="pill pill-muted">{fmtDate(n.date)}</span>
                <div className="spacer" />
                <IconBtn title="Edit" onClick={() => startEdit(n)}><Pencil size={14} /></IconBtn>
                <IconBtn title="Delete" danger onClick={() => remove("quickNotes", n.id)}><Trash2 size={14} /></IconBtn>
              </>
            )}
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
        gap: 16px; padding: 14px 20px; background: var(--navy); color: #fff;
        padding-top: max(14px, env(safe-area-inset-top));
        padding-left: max(20px, env(safe-area-inset-left));
        padding-right: max(20px, env(safe-area-inset-right));
      }
      .topbar-left { display: flex; align-items: center; gap: 12px; flex-shrink: 0; }
      .topbar-actions { display: flex; align-items: center; gap: 8px; flex-shrink: 0; }
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
      .search-wrap { position: relative; flex: 1; max-width: 440px; margin: 0 auto; }
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
        box-shadow: 2px 0 12px rgba(0,0,0,0.15); overflow-y: auto;
        padding-top: max(20px, env(safe-area-inset-top));
        padding-bottom: max(12px, env(safe-area-inset-bottom));
        padding-left: max(8px, env(safe-area-inset-left));
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
      .content { flex: 1; padding: 24px 28px; padding-bottom: max(24px, env(safe-area-inset-bottom)); min-width: 0; }
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
      .pill-accent { background: var(--navy); color: #fff; font-weight: 700; }
      .violation-title-group { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
      .violation-desc-preview {
        font-size: 12px; color: var(--ink-soft); max-width: 420px;
        overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
      }
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
      .unit-block { border-bottom: 1px solid var(--border); }
      .unit-block:last-of-type { border-bottom: none; }
      .unit-row {
        display: flex; align-items: center; gap: 10px; width: 100%; text-align: left;
        background: none; border: none; padding: 8px 4px; cursor: pointer; font-family: inherit; font-size: 13px;
      }
      .unit-row:hover { background: #F0EEE7; }
      .unit-row-number { font-weight: 600; flex-shrink: 0; }
      .unit-row-name { color: var(--ink-soft); flex: 1; }
      .unit-detail { padding: 4px 4px 12px 26px; }
      .unit-detail-tenant { margin-bottom: 8px; }
      .unit-detail-tenant:last-child { margin-bottom: 0; }
      .unit-detail-row { font-size: 13px; padding: 2px 0; }
      .import-block { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 16px; margin-bottom: 18px; }
      .import-block-head { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; margin-bottom: 4px; }
      .import-block-title { font-family: Georgia, serif; font-size: 16px; font-weight: 600; margin-right: auto; }
      .import-preview { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 14px; margin-bottom: 16px; }
      .import-preview-summary { font-weight: 600; font-size: 13px; margin-bottom: 8px; }
      .import-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; padding: 6px 0; border-bottom: 1px solid var(--border); font-size: 13px; }
      .import-row:last-of-type { border-bottom: none; }
      .import-row-name { font-weight: 600; }
      .import-row-fields { display: flex; flex-wrap: wrap; gap: 8px; color: var(--ink-soft); font-size: 12px; }
      .import-row-approval { background: #FBF6EF; margin: 0 -10px; padding: 6px 10px; border-radius: 6px; }
      .import-approve { display: flex; align-items: center; gap: 6px; font-size: 12px; color: var(--ink); margin-top: 4px; width: 100%; cursor: pointer; }
      .sheet-name-cell { display: flex; align-items: center; gap: 4px; }
      .sheet-name-cell .sheet-input { flex: 1; }
      .sheet-court-pill { flex-shrink: 0; margin-right: 4px; font-size: 10px; padding: 2px 6px; }

      .attention-icon { flex-shrink: 0; }
      .attention-count { font-family: Georgia, serif; font-size: 17px; margin-right: 2px; }
      .attention-label { font-size: 13px; color: var(--ink-soft); flex: 1; }
      .all-clear {
        display: flex; align-items: center; gap: 10px; background: var(--ok-bg); color: var(--ok);
        border-radius: 8px; padding: 16px; font-size: 13px; margin-bottom: 24px;
      }
      .dash-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 10px; }
      .dash-building-card { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 14px; }
      .dash-building-clickable { cursor: pointer; }
      .dash-building-clickable:hover { border-color: var(--navy); }
      .dash-building-name { display: flex; align-items: center; gap: 6px; justify-content: space-between; font-weight: 600; font-size: 13px; margin-bottom: 8px; }
      .dash-building-detail { margin-top: 10px; padding-top: 10px; border-top: 1px solid var(--border); display: flex; flex-direction: column; gap: 6px; }
      .dash-detail-item {
        display: flex; align-items: center; gap: 10px; width: 100%; text-align: left;
        background: #fff; border: 1px solid var(--border); border-radius: 6px;
        padding: 7px 10px; cursor: pointer; font: inherit;
      }
      .dash-detail-item:hover { background: var(--bg); border-color: var(--navy); }
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
      .sheet-readonly { padding: 8px 10px; color: var(--ink-soft); font-size: 13px; }
      .balance-input-wrap { position: relative; display: flex; align-items: center; }
      .balance-dollar { position: absolute; left: 8px; color: var(--ink-soft); font-size: 13px; pointer-events: none; }
      .balance-input { padding-left: 18px !important; width: 100%; }
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
      .dash-calendar {
        background: var(--panel); border: 1px solid var(--border); border-radius: 8px;
        padding: 12px 14px; margin-bottom: 18px;
      }
      .dash-calendar-compact { max-width: 480px; }
      .dash-calendar-head { display: flex; align-items: center; gap: 6px; margin-bottom: 8px; color: var(--navy); }
      .dash-calendar-title { font-weight: 700; font-size: 13px; }
      .dash-cal-overdue-chip { border-color: var(--danger); color: var(--danger); }
      .dash-cal-modes { margin-bottom: 8px; gap: 4px; }
      .dash-cal-modes .chip { font-size: 11px; padding: 3px 9px; }
      .dash-cal-week { display: grid; grid-template-columns: repeat(7, 1fr); gap: 4px; margin-bottom: 8px; }
      .dash-cal-month { display: grid; grid-template-columns: repeat(7, 1fr); gap: 3px; margin-bottom: 8px; }
      .dash-cal-month-dow { text-align: center; font-size: 9px; color: var(--ink-soft); font-weight: 700; padding-bottom: 2px; }
      .dash-cal-cell {
        display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 1px;
        background: #fff; border: 1px solid var(--border); border-radius: 5px;
        padding: 5px 2px; cursor: pointer; position: relative; font: inherit;
      }
      .dash-cal-cell-sm { padding: 3px 2px; aspect-ratio: 1; }
      .dash-cal-cell:hover { border-color: var(--navy); }
      .dash-cal-cell-today { border-color: var(--navy); border-width: 2px; }
      .dash-cal-cell-selected { background: var(--navy); }
      .dash-cal-cell-selected .dash-cal-cell-label, .dash-cal-cell-selected .dash-cal-cell-num { color: #fff; }
      .dash-cal-cell-dim { opacity: 0.35; }
      .dash-cal-cell-label { font-size: 9px; text-transform: uppercase; color: var(--ink-soft); letter-spacing: 0.02em; }
      .dash-cal-cell-num { font-size: 13px; font-weight: 700; }
      .dash-cal-dot { width: 5px; height: 5px; border-radius: 999px; background: var(--danger); }
      .dash-cal-year { display: grid; grid-template-columns: repeat(4, 1fr); gap: 6px; margin-bottom: 8px; }
      .dash-cal-month-cell {
        display: flex; flex-direction: column; align-items: center; gap: 3px;
        background: #fff; border: 1px solid var(--border); border-radius: 6px;
        padding: 10px 4px; cursor: pointer; font: inherit; font-size: 12px; font-weight: 600;
      }
      .dash-cal-month-cell:hover { border-color: var(--navy); }
      .dash-cal-section { margin-top: 4px; }
      .dash-cal-section-title { font-size: 11px; font-weight: 700; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.03em; margin-bottom: 6px; }
      .dash-cal-overdue { color: var(--danger); }
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
        .topbar { flex-wrap: wrap; }
        .search-wrap { order: 3; max-width: 100%; width: 100%; margin: 0; flex: 1 1 100%; }
        .brand-sub { display: none; }
        .content { padding: 16px; padding-bottom: max(16px, env(safe-area-inset-bottom)); }
        /* iOS Safari auto-zooms the whole page when you tap an input with a font
           smaller than 16px — jarring on every single field in an app this
           form-heavy. Force 16px on mobile only, so desktop stays compact. */
        input, select, textarea { font-size: 16px !important; }
        /* Icon-only buttons (edit/delete/etc, used on nearly every row) were
           only ~22px of actual tap area — below the ~44px minimum comfortable
           touch target, easy to mis-tap Delete instead of Edit on a real phone. */
        .icon-btn { min-width: 40px; min-height: 40px; justify-content: center; }
        .btn-primary, .btn-ghost { min-height: 40px; }
        .page-actions, .form-actions { gap: 8px; }
        .stat-value { font-size: 24px; }
        .sheet-input, select.sheet-input { padding: 10px 8px; }
      }
      .save-error-banner {
        display: flex; align-items: center; gap: 8px; background: var(--danger-bg); color: var(--danger);
        padding: 10px 20px; font-size: 13px; border-bottom: 1px solid var(--danger);
      }
      .login-shell {
        min-height: 100vh; display: flex; align-items: center; justify-content: center;
        background: var(--bg); font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Inter, sans-serif;
        padding: 20px; box-sizing: border-box;
      }
      .login-card {
        background: var(--panel); border: 1px solid var(--border); border-radius: 10px;
        padding: 32px; width: 100%; max-width: 320px; display: flex; flex-direction: column; box-sizing: border-box;
      }
      @media print {
        .no-print, .icon-btn, .btn-ghost, .btn-primary, .page-actions,
        .filter-row, .form-panel, .import-preview, .import-block,
        .sheet-actions, .sheet-follow-btn, .chevron-icon,
        .lucide-chevron-right, .lucide-chevron-down, .search-wrap, .menu-btn {
          display: none !important;
        }
        .app-shell { border-radius: 0; background: #fff; }
        .content { padding: 0; }
        .list-card, .stat-card, .building-row, .dash-building-card { break-inside: avoid; border: 1px solid #ccc; }
        body { background: #fff; }
        .topbar { background: #fff !important; color: #000 !important; border-bottom: 2px solid #000; }
        .brand-title { color: #000; }
        .brand-sub { color: #444; }
        .brand-mark { background: #ddd !important; color: #000 !important; }
        .list-card-head { cursor: default !important; }
        .list-card-body { display: block !important; }
        .sheet-wrap { overflow: visible; border: none; }
        .sheet { width: 100%; border-collapse: collapse; }
        .sheet th, .sheet td { border: 1px solid #999 !important; padding: 6px 8px !important; }
        .sheet th { background: #eee !important; color: #000 !important; }
        .sheet-input, select.sheet-input {
          border: none !important; background: none !important; -webkit-appearance: none;
          appearance: none; padding: 0 !important; color: #000 !important; pointer-events: none;
        }
        .sheet-status-ok, .sheet-status-warn, .sheet-status-danger { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        .sheet-note-preview { pointer-events: none; }
        .sheet-note-text { white-space: normal !important; max-width: none !important; }
        .pill { -webkit-print-color-adjust: exact; print-color-adjust: exact; border: 1px solid currentColor; }
        .sheet-court-pill { border: 1px solid var(--danger); }
      }
    `}</style>
  );
}
