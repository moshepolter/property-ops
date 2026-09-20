import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
// PDF upload support for RIS reports. Requires: npm install pdfjs-dist
// The worker is loaded from a CDN so it works with any bundler — no local worker file needed.
import * as pdfjsLib from "pdfjs-dist";
// pdf.js 4.x only ships an ES-module worker build (.mjs) — the classic .js
// build some older examples reference doesn't exist at this version, and
// pointing at it 404s silently, breaking every single PDF upload with no
// clue why. cdnjs is primary since it's already the CSP-allowed host used
// elsewhere in this app; unpkg is a fallback extractPdfText retries with if
// the primary ever fails to load, so one CDN having a bad day doesn't take
// PDF reading down entirely.
const PDF_WORKER_PRIMARY = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.worker.min.mjs";
const PDF_WORKER_FALLBACK = "https://unpkg.com/pdfjs-dist@4.0.379/build/pdf.worker.min.mjs";
pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_PRIMARY;
// Firebase for real persistence + login. Requires: npm install firebase
// Fill in firebaseConfig below with the values from your Firebase project settings.
import { initializeApp } from "firebase/app";
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { getFirestore, doc, setDoc, onSnapshot } from "firebase/firestore";
import { getStorage, ref as storageRef, uploadBytes, getDownloadURL, deleteObject } from "firebase/storage";
import { getFunctions, httpsCallable } from "firebase/functions";

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
const storage = getStorage(firebaseApp);
const functions = getFunctions(firebaseApp);
import {
  Search, Building2, Users, Wrench, AlertTriangle, Gavel, HardHat, Home, Phone, Mail,
  CalendarClock, ScrollText, MessageSquare, Archive as ArchiveIcon, DollarSign, StickyNote,
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
// numeric suffix and compare those separately. Two different real-world
// conventions exist though: letter-then-number (A1, B39) AND number-then-
// letter (1H, 4A, 4B) — the old version only recognized the first, which
// silently dropped the trailing letter on the second, so "1G" and "1H" (or
// "4A" and "4B") all collapsed to the same sort key and ended up in
// whatever arbitrary order they happened to already be in.
function unitSortKey(u) {
  const s = String(u || "");
  let m = s.match(/^([A-Za-z]+)(\d+)/);
  if (m) return [m[1], parseInt(m[2], 10), ""];
  m = s.match(/^(\d+)([A-Za-z]*)/);
  if (m) return ["", parseInt(m[1], 10), m[2] || ""];
  m = s.match(/^([A-Za-z]*)/);
  return [m[1] || s, 0, ""];
}
function compareUnits(a, b) {
  const [al, an, as] = unitSortKey(a);
  const [bl, bn, bs] = unitSortKey(b);
  if (al !== bl) return al.localeCompare(bl);
  if (an !== bn) return an - bn;
  return as.localeCompare(bs);
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
// Turns a full HPD/DSNY-style legal violation description into a short,
// readable label for the collapsed view — e.g. "self-closing doors that is
// missing or defective latch in the entrance located at apt b10..." becomes
// "Self-closing door". The full text is still shown once expanded; this is
// only for quickly scanning a list.
const VIOLATION_KEYWORDS = [
  [/self[\s-]?closing\s+doors?/i, "Self-closing door"],
  [/window\s+guards?/i, "Window guard"],
  [/lead[\s-]?based\s+paint|\blead\b.*\bpaint\b/i, "Lead paint"],
  [/\bmold\b/i, "Mold"],
  [/infestation.*roach|roach.*infestation/i, "Roach infestation"],
  [/infestation.*mice|mice.*infestation|\brodent/i, "Mice/rodent infestation"],
  [/infestation.*bed\s?bug|bed\s?bug/i, "Bed bugs"],
  [/water\s+leak/i, "Water leak"],
  [/smoke\s+detector/i, "Smoke detector"],
  [/carbon\s+monoxide/i, "CO detector"],
  [/\bheat\b.*(insufficient|inadequate|lack)|lack.*\bheat\b/i, "Heat"],
  [/hot\s+water/i, "Hot water"],
  [/plaster|paint\b.*ceiling|paint\b.*wall/i, "Plaster/paint"],
  [/masonry/i, "Masonry"],
  [/electrical|wiring/i, "Electrical"],
  [/gas\s+(leak|piping|meter)/i, "Gas"],
  [/elevator/i, "Elevator"],
  [/fire\s+escape/i, "Fire escape"],
  [/floor.*defective|defective.*floor/i, "Flooring"],
  [/ceiling.*collapse|collapse.*ceiling/i, "Ceiling collapse"],
  [/sink|faucet|plumbing/i, "Plumbing"],
  [/toilet/i, "Toilet"],
  [/bathtub|shower/i, "Bathtub/shower"],
  [/window\b.*(broken|defective|missing)/i, "Window"],
  [/lock|latch/i, "Lock/latch"],
  [/lighting|light\s+fixture/i, "Lighting"],
  [/garbage|refuse|trash/i, "Refuse/garbage"],
  [/pest\s+control|extermination/i, "Pest control"],
];
function summarizeViolationDescription(description) {
  if (!description) return "";
  // Strip the leading legal-citation clause (everything up to and
  // including the first colon), which is boilerplate, not the issue.
  const afterCitation = description.replace(/^[^:]*:\s*/, "");
  for (const [pattern, label] of VIOLATION_KEYWORDS) {
    if (pattern.test(afterCitation)) return label;
  }
  // No known keyword matched — fall back to a trimmed excerpt, cut at the
  // "located at" / "in the" clause that usually starts the location detail.
  const trimmed = afterCitation.split(/\s+located at\s+|\s+in the \d/i)[0];
  return trimmed.length > 60 ? trimmed.slice(0, 60).trim() + "…" : trimmed.trim();
}

function shortAddress(address) {
  if (!address) return "";
  let s = address.replace(/,\s*[A-Za-z .]+\s+\d{5}(-\d{4})?\s*$/, "");
  s = s.replace(/\s+(BROOKLYN|MANHATTAN|QUEENS|BRONX|STATEN ISLAND|NEW YORK)\s*$/i, "");
  // The two steps above can strip the zip/state and the city name
  // separately, leaving the comma that used to sit between them orphaned
  // at the end ("333 Ovington Avenue," instead of "333 Ovington Avenue") —
  // clean that up before trimming.
  s = s.replace(/,\s*$/, "");
  return s.trim() || address;
}

// Shared by Violations and Work Orders for the "copy" feature — groups
// selected items by building+unit so the same address and tenant contact
// line isn't repeated for multiple issues at the same apartment; each
// group lists its address once, then every issue in that group, then the
// tenant's contact line. Different apartments get their own group,
// separated by a divider line, so a batch spanning several units pastes as
// one clean, readable block instead of a jumble.
function formatItemsForCopy(items, data) {
  const groups = [];
  for (const item of items) {
    let group = groups.find(g => g.buildingId === item.buildingId && g.unitId === item.unitId);
    if (!group) {
      group = { buildingId: item.buildingId, unitId: item.unitId, issues: [] };
      groups.push(group);
    }
    group.issues.push(item.description || "(no description)");
  }
  const blocks = groups.map(g => {
    const building = data.buildings.find(b => b.id === g.buildingId);
    const addr = building ? shortAddress(building.address) : "(building not found)";
    const unit = data.units.find(u => u.id === g.unitId);
    const addressLine = unit && unit.unitNumber ? `${addr}, Apt ${unit.unitNumber}` : addr;
    const issueLines = g.issues.join("\n");
    const tenant = data.tenants.find(t => t.unitId === g.unitId);
    const tenantLine = tenant
      ? `${tenant.name || "(no name on file)"}: ${tenant.phone || "(no phone on file)"}, please schedule`
      : `(no tenant on file), please schedule`;
    return `${addressLine}\n${issueLines}\n${tenantLine}`;
  });
  return blocks.join("\n-------\n");
}

// Uploads go to Firebase Storage now, not straight into Firestore as base64 —
// a single un-resized phone photo can be several MB, and Firestore caps an
// entire document at 1MB total. Storage has no meaningful size limit, and we
// only ever keep a small URL string in the actual data record.
// Resizes and re-compresses a photo in the browser before it ever leaves the
// device — a straight-from-camera phone photo can be several MB, which is
// wasteful once it's headed to storage. This keeps most photos in the tens-
// to-low-hundreds of KB range instead, with no visible quality loss at the
// sizes anyone actually views a violation/work-order photo at.
function compressImage(file, maxDimension = 1600, quality = 0.75) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDimension || height > maxDimension) {
          if (width > height) { height = Math.round(height * (maxDimension / width)); width = maxDimension; }
          else { width = Math.round(width * (maxDimension / height)); height = maxDimension; }
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        canvas.toBlob(
          (blob) => blob ? resolve(new File([blob], file.name.replace(/\.\w+$/, ".jpg"), { type: "image/jpeg" })) : reject(new Error("compression produced no blob")),
          "image/jpeg", quality
        );
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function uploadFilesToStorage(fileList, pathPrefix, { compressImages: shouldCompress = false } = {}) {
  const uidPart = auth.currentUser?.uid || "unknown";
  const files = Array.from(fileList);
  return Promise.all(files.map(async (file) => {
    let uploadFile = file;
    if (shouldCompress && file.type.startsWith("image/")) {
      try { uploadFile = await compressImage(file); }
      catch (err) { console.error("photo compression failed, uploading original", err); }
    }
    const id = uid();
    const storagePath = `users/${uidPart}/${pathPrefix}/${id}-${uploadFile.name}`;
    const fileRef = storageRef(storage, storagePath);
    await uploadBytes(fileRef, uploadFile);
    const url = await getDownloadURL(fileRef);
    return { id, name: file.name, url, storagePath };
  }));
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

// Converts the HPD report's MM/DD/YYYY dates to the ISO (YYYY-MM-DD) format
// every other date field in the app uses — "-" or empty means no date, same
// convention as the rest of the app's blank date fields.
function isoFromMDY(mdy) {
  if (!mdy || mdy === "-") return "";
  const [mm, dd, yyyy] = mdy.split("/");
  if (!mm || !dd || !yyyy) return "";
  return `${yyyy}-${mm}-${dd}`;
}
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
// A missing date counts as "due" too — no date set means go set one, not a
// free pass. dateDueStrict is for cases where an unset date genuinely isn't
// urgent (a stip payment with no scheduled date isn't overdue the same way
// a missing court date is).
function dateDue(d) { const days = daysUntil(d); return days === null || days <= 7; }
function dateDueStrict(d) { const days = daysUntil(d); return days !== null && days <= 7; }
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
// Global so every embedded header fragment gets stripped, not just the
// first — a flattened single-line paste (no real line breaks preserved by
// whatever copied it) can have several of these mashed together with the
// actual tenant data all on one line, with nothing to separate them.
const SKIP_LINE_RE = /^\d{2}\/\d{2}\/\d{4}.*Page:\s*\d+\s*$|^\d{2}\/\d{2}\/\d{4}|FISCAL PERIOD[^A-Z]*TO\s+\d{2}\/\d{2}\/\d{4}|Page:\s*\d+|PROP #\s*\S+:[^\n]*?\d{5}\b|TELEPHONE\/EMAIL LIST|BUILDING DIRECTORY|AGED ARREARS FOR [A-Za-z]+\s*,\s*\d{4}|CODE:\s*0-30 DAYS:\s*31-60 DAYS:\s*61\+ DAYS:\s*TOTAL DUE:|LEGAL:|\*\s*-\s*MOVED OUT|TOTALS:|TENANT NAME:/ig;

function cleanLines(text) {
  return text.split("\n").map(l => l.trim()).map(l => {
    // Some reports print "APT: A1" as one line (the real code merged with the
    // column label) instead of the usual separate "APT: TENANT NAME:" header
    // row — keep the code, drop just the label, so that apartment isn't lost.
    const m = l.match(/^APT:\s*(.+)$/i);
    return m ? m[1].trim() : l;
  }).map(l => l.replace(SKIP_LINE_RE, " ").replace(/\s+/g, " ").trim())
    .filter(l => l && !/^APT:?$/i.test(l));
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

// CSV column headers are never guaranteed to match a specific case or
// spacing — "Address", "ADDRESS", "address ", and "address" are all the
// same column to a human, but a plain row.address lookup only catches one
// of them, silently treating the rest as blank. This checks every actual
// column name in the row (trimmed, lowercased) against each candidate name
// given, so the exact capitalization/whitespace the CSV happens to use
// never matters.
function csvField(row, ...names) {
  const normalized = {};
  Object.keys(row).forEach(k => { normalized[k.trim().toLowerCase()] = row[k]; });
  for (const name of names) {
    const v = normalized[name.toLowerCase()];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  return "";
}

// Balances can be a plain number ("1200"), empty, or — from arrears imports
// done before a past fix — a display-formatted string like "$1,234.56".
// Plain parseFloat chokes on the $ and returns NaN. This strips anything
// that isn't a digit/decimal/minus first, so totals, sorting, and the
// import-delta math are correct regardless of which format a given tenant's
// balance happens to be in. Shared at module level since buildImportDiff
// runs outside RentTab and needs the same safety.
function parseBalance(b) {
  return parseFloat(String(b || "0").replace(/[^0-9.-]/g, "")) || 0;
}

// Aged Arrears: "A1 AUDREY LYNN MELENDEZ UNKNO 1698.80 1698.80 3794.00 7191.60"
function parseArrearsTextLineFormat(text) {
  // Locally scoped, not the shared APT_RE — some buildings number units
  // letter-first (A1, B62), others digit-first with a letter suffix (1A,
  // 4D, 6K), others a bare 2-letter code with no digit at all ("GA", "GB"
  // for a garden-level unit), others a bare number alone for a commercial
  // unit ("8101"), others a range covering several combined units
  // ("211-15" for units 211 through 215), others two units combined under
  // one listing ("A_&_B"), others a floor identified by number instead of a
  // unit code — a single floor ("3_FL"), a merged pair of floors
  // ("4&5_FL"), or a floor with a direction/section suffix ("6_FL_N") — and
  // others a single bare letter ("C", "D", "E"). Recognize all of these
  // without changing APT_RE's behavior for the contacts parser, which
  // intentionally stays letter-first-only there.
  const COMBO_APT_RE = "[A-Z]_&_[A-Z]";
  const FLOOR_APT_RE = "\\d{1,2}(?:&\\d{1,2})?_FL(?:_[A-Z])?";
  // A bare letter-only unit code can be 1 to 4 letters long with no digit at
  // all — "C"/"D"/"E" for a single letter, "GA"/"GB" for a 2-letter garden
  // unit, or "HERC" (a commercial/super unit, 4 letters). Wide but safe
  // here specifically because each span below is already isolated between
  // two unambiguous dollar-amount boundaries — there's no risk of a
  // coincidental word elsewhere in the document winning out the way there
  // would be in a continuous whole-document scan.
  const LETTER_ONLY_APT_RE = "[A-Z]{1,4}";
  const LINE_APT_RE = `(?:${APT_RE}|\\d{1,4}-\\d{1,4}|\\d{1,4}[A-Z]{1,4}|${COMBO_APT_RE}|${FLOOR_APT_RE}|${LETTER_ONLY_APT_RE}|\\d{1,4})`;
  // Joined into one continuous string rather than matched line-by-line —
  // some paste sources flatten the whole report onto a single line with no
  // real breaks at all, and a per-line match would only ever find one
  // record in that case no matter how many tenants are actually in the
  // text. Working through the whole text instead finds every one
  // regardless of whether real line breaks survived the paste.
  const cleaned = cleanLines(text).join(" ");

  // Rather than one giant pattern trying to tell "a single-letter unit code"
  // apart from "an ordinary single-letter word inside someone's name" using
  // a lookbehind (a regex feature older mobile Safari doesn't support at
  // all — it throws immediately, which is enough to crash the whole app
  // before anything renders), this finds the unambiguous part first: every
  // "4 consecutive dollar amounts" sequence, which needs no lookaround of
  // any kind. Each one marks where one entry ENDS. The text between one
  // boundary and the next is then that entry's own self-contained "unit
  // code + name" span, processed on its own — its own start IS the
  // context, so there's nothing left to look behind for.
  const AMOUNTS_RE = /([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})(?=\s|$)/g;
  const boundaries = [];
  let am;
  while ((am = AMOUNTS_RE.exec(cleaned))) {
    boundaries.push({ start: am.index, end: am.index + am[0].length, values: [am[1], am[2], am[3], am[4]] });
  }

  const out = [];
  let spanStart = 0;
  for (const b of boundaries) {
    const span = cleaned.slice(spanStart, b.start).trim();
    spanStart = b.end;
    // The FIRST valid unit-code occurrence within the span, not necessarily
    // right at its start — a building's own company name sometimes repeats
    // itself right before the very first entry ("Neptune Group LLC APT:
    // A_&_B Neptune Plumbing...", where none of "Neptune", "Group", or
    // "LLC" individually satisfy any unit-code pattern as a complete,
    // isolated word) and generic header-stripping can't know every
    // building's name in advance to strip it specifically. The first VALID
    // token in the span is always the real unit code — using the first
    // (not last) occurrence also correctly steps past a coincidental
    // matching word inside the tenant's own name later in the same span (a
    // business named "...Steps To Success" has "To" as a standalone
    // 2-letter word, but it always comes after the real unit code, never
    // before it).
    const startMatch = span.match(new RegExp(`(?:^|\\s)(${LINE_APT_RE})(\\*)?\\s+`));
    if (!startMatch) continue; // no recognizable unit code anywhere in this span — skip, don't guess
    const apt = startMatch[1];
    const moved = startMatch[2];
    const rawName = span.slice(startMatch.index + startMatch[0].length).replace(/\s*UNKNO\s*$/, "").trim();
    const name = rawName.replace(/^N-\d+\s+/, "").trim();
    const [d1, d2, d3, total] = b.values;
    const totalNum = parseFloat(total.replace(/,/g, ""));
    const bucket1 = parseFloat(d1.replace(/,/g, ""));
    const bucket2 = parseFloat(d2.replace(/,/g, ""));
    const bucket61 = parseFloat(d3.replace(/,/g, ""));
    let status = "Current";
    if (totalNum > 0) status = bucket61 > 0 ? "In Arrears" : "Late";
    // Store balance as a plain number string, same as every manually-typed
    // balance in the sheet — NOT money()'s "$1,234.56" display format, which
    // parseFloat can't read (returns NaN), silently breaking any total or
    // sort built on top of it. money() is for display only, never storage.
    out.push({ apt, movedOut: !!moved, name, balance: totalNum.toFixed(2), status, aging: { bucket1, bucket2, bucket61 } });
  }
  return out;
}

// A second RIS layout prints the report in columns instead of one line per
// tenant: a block of every unit number, then a block of every tenant name,
// then a block of codes, then four blocks of dollar amounts (one row per
// aging bucket) — each block internally space/newline-separated rather than
// aligned per tenant. It's split by page (each starting with "PROP #"), since
// a trailing number in each dollar row is a page subtotal, not a per-unit
// value, and has to be dropped using that page's own unit count.
//
// Names in this layout can't be reliably split back apart — two short names
// often land on the same extracted line with no delimiter between them, and
// there's no way to know where one ends and the next begins. Rather than
// guess and risk assigning someone else's name to a unit, entries from a
// page where the name-line count doesn't match the unit count get an empty
// name (which — since a blank value never overwrites an existing tenant's
// name — just leaves whatever's already on file untouched) and are flagged
// needsReview so the person importing knows to double-check that page.
// Every section label this report style can use — a section's content is
// found between its label and whichever of these comes next in the text,
// rather than assuming a fixed order. Different buildings' exports print
// TOTAL DUE either first (right after the header) or last (after the aging
// buckets), so a fixed start-label/end-label pairing breaks on whichever
// order it didn't expect.
const ARREARS_SECTION_MARKERS = ["TOTAL DUE:", "APT:\\s*LEGAL:", "TENANT NAME:", "CODE:", "0-30 DAYS:", "31-60 DAYS:", "61\\+ DAYS:", "AGED ARREARS FOR"];
function extractArrearsSection(page, labelPattern) {
  const startMatch = page.match(new RegExp(labelPattern));
  if (!startMatch) return null;
  const startIdx = startMatch.index + startMatch[0].length;
  let endIdx = page.length;
  for (const marker of ARREARS_SECTION_MARKERS) {
    const re = new RegExp(marker, "g");
    re.lastIndex = startIdx;
    const nextMatch = re.exec(page);
    if (nextMatch && nextMatch.index < endIdx) endIdx = nextMatch.index;
  }
  return page.slice(startIdx, endIdx);
}

function parseArrearsTextColumnar(text) {
  // Matches BOTH apartment-numbering conventions — letter-first (A1, B62)
  // and number-first (0B, 1H, 2BB) — since which one a building uses varies,
  // and a pattern that only recognized one left the other's units silently
  // unmatched (nothing to do with a broken report — just a different
  // building's normal numbering style).
  const APT_TOKEN_RE = /\b(?:[A-Za-z]{1,4}\d{1,4}|\d{1,4}[A-Za-z]{1,4}|[A-Za-z]{1,4})\*?/g;
  const NUM_RE = /[\d,]*\d\.\d{2}/g;
  const pages = text.split(/(?=PROP #)/).filter(p => p.trim());
  const out = [];

  pages.forEach(page => {
    const aptSection = extractArrearsSection(page, "APT:\\s*LEGAL:");
    if (!aptSection) return;
    const cleanedAptSection = aptSection.replace(/\bN-\d+\b/g, " ");
    const apts = (cleanedAptSection.match(APT_TOKEN_RE) || []);
    const n = apts.length;
    if (n === 0) return;

    const bucket = (label) => {
      const section = extractArrearsSection(page, label);
      if (!section) return [];
      // Only the first n values are per-unit — anything after that on a
      // page is a page subtotal, never a tenant's own balance.
      return (section.match(NUM_RE) || []).slice(0, n).map(s => parseFloat(s.replace(/,/g, "")));
    };
    const b1 = bucket("0-30 DAYS:");
    const b2 = bucket("31-60 DAYS:");
    const b3 = bucket("61\\+ DAYS:");
    const b4 = bucket("TOTAL DUE:");

    const nameSection = extractArrearsSection(page, "TENANT NAME:");
    const nameLines = nameSection ? nameSection.split("\n").map(l => l.trim()).filter(Boolean) : [];
    const namesReliable = nameLines.length === n;

    for (let i = 0; i < n; i++) {
      const rawApt = apts[i];
      const movedOut = rawApt.endsWith("*");
      const apt = rawApt.replace("*", "").toUpperCase();
      const totalNum = b4[i] !== undefined ? b4[i] : 0;
      const bucket61 = b3[i] || 0;
      let status = "Current";
      if (totalNum > 0) status = bucket61 > 0 ? "In Arrears" : "Late";
      out.push({
        apt, movedOut,
        name: namesReliable ? nameLines[i] : "",
        needsReview: !namesReliable,
        balance: totalNum.toFixed(2), status,
        aging: { bucket1: b1[i] || 0, bucket2: b2[i] || 0, bucket61 },
      });
    }
  });
  return out;
}

function parseArrearsText(text) {
  // A columnar-format report (TOTAL DUE:, TENANT NAME:, CODE:, and the
  // three aging-bucket headers each as their own standalone section,
  // followed by a block of numbers) can occasionally trip the line-format
  // parser's own "4 consecutive dollar amounts" boundary detection — a run
  // of aging-bucket numbers can coincidentally look like that pattern,
  // producing SOME output even though it's actually garbage, which
  // previously meant the line parser's result got trusted blindly and the
  // columnar parser never even ran. Counting how many of these markers
  // appear gives a reliable structural signal for which format this
  // actually is, checked before trusting either parser's raw output.
  const columnarMarkerCount = ARREARS_SECTION_MARKERS.filter(marker => new RegExp(marker).test(text)).length;
  if (columnarMarkerCount >= 4) {
    const columnar = parseArrearsTextColumnar(text);
    if (columnar.length > 0) return columnar;
  }
  const lineFormat = parseArrearsTextLineFormat(text);
  if (lineFormat.length > 0) return lineFormat;
  return parseArrearsTextColumnar(text);
}

// Telephone/Email list: header line "A1 AUDREY LYNN MELENDEZ" then indented "CELL - ...", "EMAIL ADDRESS - ..."
function parseContactsText(text) {
  // Some reports number each entry ("183. GA", "1. 1A") — strip that leading
  // sequence number first, or it reads as a bare numeric unit code with what
  // looks like the start of a name right after it (the real unit code),
  // misidentifying the sequence number as the unit instead.
  const seqStripped = text.replace(/(?:^|\n)\s*\d{1,4}\.\s+/g, "\n");
  // Different PDF viewers copy multi-line contact cells in unpredictable order —
  // sometimes labels and values stay paired, sometimes all labels get grouped
  // together with all their values afterward. Rather than trying to track which
  // line goes with which, treat the whole report as one continuous block: find
  // every apartment header, then pull the first phone number and first email
  // found anywhere between that header and the next one, wherever it landed.
  const lines = cleanLines(seqStripped);
  // Track which character offsets in the joined text are genuine line
  // starts from the source document — needed to tell a real letter-only
  // unit ("D Beauty Options Inc.", starting its own new line) apart from
  // an annotation word that just happens to share a line with a phone
  // number ("CELL - 973-851-1585 KATY" — KATY is whose cell that is, not
  // a new unit; "103 Bertha Trujillo" starting the actual next line is the
  // real header). A text-pattern-only check can't reliably tell these
  // apart since both follow a phone number — but only one of them starts
  // a fresh line in the original document.
  const lineStarts = new Set();
  let cleaned = "";
  for (const line of lines) {
    if (cleaned) cleaned += " ";
    lineStarts.add(cleaned.length);
    cleaned += line;
  }
  const LABELS = "(?:CELL|EMAIL ADDRESS|HOME|WORK|OTHER|FAX)";
  // & included so a business literally named with one ("Nixon & Son Meat
  // Market Corp.") doesn't break the match partway through — anywhere else
  // in these reports an ampersand only shows up inside a name, never as
  // meaningful punctuation of its own.
  const NAME_CHARS = "A-Za-z.,'&\\-\\s";
  // Some buildings also have commercial/storefront units identified by a bare
  // number (no letter prefix — anywhere from a single digit like "1" or "3"
  // up to a longer code like "9516 JH ORGANIC INC." or "319"), others use
  // digit-then-letter codes ("1B", "2BB"), and others a bare letter-only code
  // with no digit at all — anywhere from one letter ("C", "D", "E") to a
  // 2-letter garden-unit code ("GA", "GB") up to a 4-letter commercial code
  // ("HERC" for Hercules Corp). Support all of these, but require the
  // numeric/digit-first forms to have a real name after them — otherwise a
  // phone number written with spaces instead of dashes ("718 833 3607 FAX")
  // can look just like a unit code.
  const ALT_APT_RE = "\\d{1,2}[A-Z]{1,2}";
  // A longer bare number can have its own single-letter sub-unit suffix too
  // ("2106A" — a sub-division of unit 2106) — same idea as ALT_APT_RE but
  // for numbers longer than the 2-digit cap that one allows.
  const LONG_ALT_APT_RE = "\\d{3,4}[A-Z]";
  // A commercial/mixed-use building can also have two units combined under
  // one listing ("A_&_B"), or a floor identified by number instead of a
  // unit code ("6_FL").
  const COMBO_APT_RE = "[A-Z]_&_[A-Z]";
  const FLOOR_APT_RE = "\\d{1,2}_FL";
  // X excluded specifically because "WORK - 555-1234 X 119" (a phone
  // extension) is a far more common way to see a lone "X" in these reports
  // than an actual unit named X.
  const LETTER_ONLY_APT_RE = "[A-WYZ][A-Za-z]{0,3}";
  // A name normally starts with a capital letter (or Mr./Mrs./Ms.) — but a
  // business can be named starting with digits ("718 Bistro Inc."), so a
  // digit-led name is accepted too, as long as a real letter shows up
  // somewhere in it. That second clause is what separates a genuine name
  // like that from a bare number (a phone extension, an apartment number
  // munged into the wrong spot) that isn't a name at all.
  const LETTER_ONLY_NAME_RE = `(?:(?:MR\\.|MRS\\.|MS\\.|[A-Z])[${NAME_CHARS}]*?|\\d[\\d\\s]*(?!(?:MR\\.|MRS\\.|MS\\.)\\s)[A-Za-z][A-Za-z0-9${NAME_CHARS.replace("A-Za-z", "")}]*?)`;
  // A letter-only unit — whether 1, 2, 3, or 4 letters — only counts as a
  // header where it starts a genuine new line in the source document (or is
  // the very first thing in it). Checking the text immediately before a
  // candidate match (a phone number, an email domain) isn't reliable on its
  // own — an annotation word sharing a line with a phone number ("CELL -
  // 973-851-1585 KATY" — KATY is whose cell that is, not a new unit) is
  // textually indistinguishable from a genuine letter-only unit that
  // legitimately follows a phone number too ("CELL - 917-686-3777 D Beauty
  // Options Inc.", where D starts its own real new line). Line boundaries
  // from the original document are the one signal that actually tells them
  // apart, since only the genuine unit starts a fresh line.
  // Letter-only left out of this lookahead (lookahead itself is fine
  // everywhere — it's only lookbehind that's the compatibility problem) —
  // since the real context check happens after matching against line
  // boundaries, including it here just meant another entry's own
  // name-capture would stop early the moment ANY coincidental letter-only
  // word appeared later in it, even though that word would go on to fail
  // the real check anyway.
  const NEXT_HEADER = `(?:${APT_RE}|\\d{1,4}|${ALT_APT_RE}|${LONG_ALT_APT_RE}|${COMBO_APT_RE}|${FLOOR_APT_RE})\\s+(?:MR\\.|MRS\\.|MS\\.|[A-Z])`;
  const headerRe = new RegExp(
    `(?:^|\\s)(?:` +
      `(${APT_RE})\\s+(?!${LABELS}\\b)((?:MR\\.|MRS\\.|MS\\.|[A-Z])[${NAME_CHARS}]*?)` +
      `|` +
      `(\\d{1,4})\\s+(?!${LABELS}\\b)((?:MR\\.|MRS\\.|MS\\.|[A-Z])[${NAME_CHARS}]+?)` +
      `|` +
      `(${LONG_ALT_APT_RE})\\s+(?!${LABELS}\\b)((?:MR\\.|MRS\\.|MS\\.|[A-Z])[${NAME_CHARS}]*?)` +
      `|` +
      `(${ALT_APT_RE})\\s+(?!${LABELS}\\b)((?:MR\\.|MRS\\.|MS\\.|[A-Z])[${NAME_CHARS}]*?)` +
      `|` +
      `(${COMBO_APT_RE})\\s+(?!${LABELS}\\b)((?:MR\\.|MRS\\.|MS\\.|[A-Z])[${NAME_CHARS}]*?)` +
      `|` +
      `(${FLOOR_APT_RE})\\s+(?!${LABELS}\\b)((?:MR\\.|MRS\\.|MS\\.|[A-Z])[${NAME_CHARS}]*?)` +
      `|` +
      `(${LETTER_ONLY_APT_RE})\\s+(?!${LABELS}\\b)(${LETTER_ONLY_NAME_RE})` +
    `)(?=\\s+${LABELS}\\b|\\s+${NEXT_HEADER}|$)`,
    "g"
  );
  const rawHeaders = [];
  let m;
  while ((m = headerRe.exec(cleaned))) {
    const apt = m[1] || m[3] || m[5] || m[7] || m[9] || m[11] || m[13];
    const name = (m[2] || m[4] || m[6] || m[8] || m[10] || m[12] || m[14] || "").trim();
    // Where the actual unit-code text starts, skipping the leading
    // separator captured by the outer (?:^|\s) — needed so the "text right
    // before this match" check below looks at the real preceding content,
    // not a boundary that includes the separator itself.
    const matchStart = m.index + (m[0].match(/^\s*/)[0].length);
    rawHeaders.push({ apt, name, start: m.index, end: m.index + m[0].length, matchStart });
  }
  // Letter-only matches (of any length, 1-4) only survive if they start a
  // genuine new line in the source document (or are the very first thing
  // in it) — otherwise it's almost always just an ordinary word sharing a
  // line with a phone number ("CELL - 555-1234 WIFE") rather than a real
  // unit code starting its own entry.
  const headers = rawHeaders.filter(h => {
    if (!/^[A-Z]{1,4}$/.test(h.apt)) return true;
    return lineStarts.has(h.matchStart);
  });

  const phoneRe = /\(?\d{3}\)?[-.\s]*\d{3}[-.\s]*\d{4}/;
  const phoneReGlobal = new RegExp(phoneRe.source, "g");
  const emailRe = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;

  return headers.map((h, i) => {
    const blockEnd = i + 1 < headers.length ? headers[i + 1].start : cleaned.length;
    const block = cleaned.slice(h.end, blockEnd);
    // A tenant's block can list more than one number (their own cell, a
    // spouse's, a work line) — capture all of them, not just the first,
    // so PhoneCycleCell has something to actually cycle through. Deduped
    // and joined with "; ", the same format that cell already expects.
    const phoneMatches = block.match(phoneReGlobal) || [];
    const uniquePhones = [...new Set(phoneMatches.map(p => p.trim()))];
    const email = block.match(emailRe);
    return { apt: h.apt, name: h.name, phone: uniquePhones.join("; "), email: email ? email[0].trim() : "" };
  }); // keep every detected apartment, even ones with no phone/email on file
}

// Attorney "Complete Client Status" export (e.g. Azoulay Weiss, LLP) —
// a narrative case log, not a table: each case block repeats its own
// header on every page it spans (with Name/Index/Addr/Landlord blank on
// continuation pages), followed by a chronological action list with the
// MOST RECENT action listed first. This pulls out one record per unique
// case number, using whichever page first supplied each field, and takes
// the very first dated action line as "the latest on the case."
// Parses an HPD "Open Violations" building report (the PDF export from HPD's
// online violation lookup) into individual violation records. Each
// violation is a 3-line block: a header line (violation ID, class, order #,
// apt, story, reported date, NOV issued date), a second line (NOV ID, NOV
// type, correction-by date, certification-by date, status, status date,
// actual certification date), and a wrapped description.
function parseHpdViolationsText(text) {
  const addrMatch = text.match(/^(.+?,\s*(?:Brooklyn|Queens|Bronx|Manhattan|Staten Island),\s*\d{5})\s*$/m);
  const buildingAddress = addrMatch ? addrMatch[1].trim() : "";

  // The "story" field isn't always a number — it can be "Not Applicable",
  // "Fire Escape", "Yards / Courts", "All Stories", etc., so it's captured
  // non-greedy up to the first recognizable date (or "-") that follows.
  const line1Re = /^(\d{6,9})\s+([A-Z])\s+(\S+)\s+(\S+)\s+(.+?)\s+(\d{2}\/\d{2}\/\d{4}|-)\s+(\d{2}\/\d{2}\/\d{4}|-)\s*$/;
  const line2Re = /^(\S+)\s+(\S+(?:\s\S+)?)\s+(\d{2}\/\d{2}\/\d{4}|-)\s+(\d{2}\/\d{2}\/\d{4}|-)\s+(.+?)\s+(\d{2}\/\d{2}\/\d{4})\s+(-|\d{2}\/\d{2}\/\d{4})\s*$/;
  // Lines that are page headers/footers/column labels repeated on every
  // page — stripped out of the description rather than left to contaminate
  // the wrapped text, the same problem solved for the court cases parser.
  const skipLineRe = new RegExp(
    "^(Page \\d+ of \\d+|Generated on|VIOLATION ID|NOV ID|VIOLATION DESCRIPTION" +
    (buildingAddress ? "|" + buildingAddress.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") : "") + ")",
    "i"
  );

  const blocks = text.split(new RegExp("(?=^" + line1Re.source.slice(1) + ")", "m")).filter(b => /^\d{6,9}\s+[A-Z]/.test(b.trim()));

  const violations = [];
  for (const block of blocks) {
    const lines = block.split("\n").map(l => l.trim()).filter(Boolean);
    if (lines.length < 2) continue;
    const m1 = lines[0].match(line1Re);
    if (!m1) continue;
    const m2 = lines[1].match(line2Re);
    const descLines = lines.slice(2).filter(l => !skipLineRe.test(l));
    const description = descLines.join(" ").replace(/\s+/g, " ").trim();

    const v = {
      violationId: m1[1], class: m1[2], orderNum: m1[3].replace(/\*$/, ""),
      apt: m1[4] === "-" ? "" : m1[4], story: m1[5],
      reportedDate: m1[6], novIssuedDate: m1[7], description,
      novId: "", novType: "", correctionByDate: "", certByDate: "", status: "", statusDate: "", actualCertDate: "",
    };
    if (m2) {
      v.novId = m2[1]; v.novType = m2[2]; v.correctionByDate = m2[3];
      v.certByDate = m2[4]; v.status = m2[5]; v.statusDate = m2[6]; v.actualCertDate = m2[7];
    }
    // Lead: word-boundary match so "lead-based paint" is caught without
    // "leading to" or "leads" false-triggering.
    v.isLead = /\blead\b/i.test(description);
    // Mold at or above the 10 sq ft threshold HPD uses under Local Law 55
    // to require licensed remediation (Class B/C) — flagged whenever mold
    // is mentioned and the description doesn't explicitly say "less than
    // 10", since an unspecified size is safer to flag than to miss.
    v.isMoldOver10 = /\bmold\b/i.test(description) && !/less than 10/i.test(description);
    violations.push(v);
  }
  return { buildingAddress, violations };
}

function parseCourtCasesText(text) {
  const lines = text.split("\n").filter(line => {
    const t = line.trim();
    if (!t) return false;
    if (/^Page:\s*\d+$/.test(t)) return false;
    if (/^Date of Report:/.test(t)) return false;
    if (/^AZOULAY WEISS, LLP$/.test(t)) return false;
    if (/^Complete Client Status$/.test(t)) return false;
    if (/^From Client = /.test(t)) return false;
    if (/^All actions$/.test(t)) return false;
    if (/^Client: MSHPEL\s+Mitchell Shelfogel$/.test(t)) return false;
    return true;
  });
  const cleaned = lines.join("\n");
  const blocks = cleaned.split(/(?=Case#:\s*\d+)/).filter(b => b.includes("Case#:"));

  const casesByNumber = {};
  for (const block of blocks) {
    const caseNumMatch = block.match(/Case#:\s*(\d+)/);
    if (!caseNumMatch) continue;
    const caseNum = caseNumMatch[1];

    const buildingMatch = block.match(/Building:\s*(\S+)/);
    const aptMatch = block.match(/Apt:\s*(.+?)\s*$/m);
    const nameMatch = block.match(/Name:\s*(.+?)\s{2,}Index:/);
    const indexMatch = block.match(/Index:\s*(\S+)/);
    // Match up to whichever comes first — the Assg field on the same line
    // if present, or end of line if not (not every case has one).
    const addrMatch = block.match(/Addr:\s*(.+?)(?:\s{2,}Assg\.:|\s*$)/m);
    const assgMatch = block.match(/Assg\.:\s*(\S+)/);
    const landlordMatch = block.match(/Landlord:\s*(.+?)\s*$/m);
    // Every dated action line in this block, not just the first — a case
    // typically has several actions (reminders, appearances, filings) and
    // the report lists them most-recent-first within each page. Long
    // descriptions that wrap onto indented continuation lines or across a
    // page break may only capture their first line here — reliably
    // stitching wrapped text across a page boundary risks garbling it, so
    // this stays honest about capturing the dated line itself rather than
    // guessing at where a wrapped sentence continues.
    // A "Stipulation:" line lists a payment schedule (date + dollar amount
    // pairs, e.g. "09/01/2026 1696.25 N"), sometimes wrapping onto a
    // following line with no "Stipulation:" prefix — these aren't case
    // actions and were being misread as ones with a garbled "description"
    // that's actually just a dollar figure. A genuine action's description
    // is words ("Court Appearance", "Reminder"); a bare amount like
    // "1696.25 N" or "1696.25" never is, so that shape is what's filtered.
    // Split into segments, each starting at a dated line (the report uses
    // both MM/DD/YYYY and the shorter M/D/YY here) and running until the
    // next dated line, a Stipulation:/Repairs:/Comment:/Landlord: section
    // (those aren't part of the action itself), or the end of this page's
    // block — giving the FULL wrapped description (e.g. what a reminder
    // note actually said), not just its first line. This only joins text
    // that's still within the same page; it never reaches across a page
    // break, since reliably knowing where a wrapped sentence continues on
    // the next page isn't something that can be verified without risking
    // garbled results.
    const segments = block.split(/(?=^\d{1,2}\/\d{1,2}\/\d{2,4}\s)|(?<=\.\s)(?=\d{1,2}\/\d{1,2}\/\d{2,4}\s)/m);
    const actionMatches = [];
    for (const seg of segments) {
      const dateMatch = seg.match(/^(\d{1,2}\/\d{1,2}\/(\d{2}|\d{4}))\s+([\s\S]*)/);
      if (!dateMatch) continue;
      let rest = dateMatch[3];
      const markerIdx = rest.search(/^(Stipulation:|Repairs:|Comment:|Landlord:)/m);
      if (markerIdx !== -1) rest = rest.slice(0, markerIdx);
      const fullText = rest.split("\n").map(l => l.trim()).filter(Boolean).join(" ").replace(/\s+/g, " ").trim().slice(0, 500);
      // A "Stipulation:" line lists a payment schedule (date + dollar
      // amount pairs, e.g. "09/01/2026 1696.25 N"), sometimes wrapping
      // onto a following line with no "Stipulation:" prefix — these
      // aren't case actions and were being misread as ones with a
      // garbled "description" that's actually just a dollar figure. A
      // genuine action's description is words ("Court Appearance",
      // "Reminder"); a bare amount like "1696.25 N" never is, so that
      // shape is what's filtered.
      if (/^\$?[\d,]+\.\d{2}\s*[A-Z]?$/.test(fullText)) continue;
      if (fullText) actionMatches.push([null, dateMatch[1], fullText]);
    }

    if (!casesByNumber[caseNum]) {
      casesByNumber[caseNum] = {
        caseNumber: caseNum, building: buildingMatch ? buildingMatch[1] : "",
        apt: aptMatch ? aptMatch[1].trim() : "", name: "", index: "", address: "",
        assigned: "", landlord: "", latestActionDate: "", latestActionDesc: "", actions: [],
      };
    }
    const c = casesByNumber[caseNum];
    // Continuation pages leave Name/Index/Addr/Assg/Landlord blank — only
    // overwrite with a genuinely non-empty value found on a later page,
    // never blank out what an earlier page already supplied.
    if (nameMatch && nameMatch[1].trim()) c.name = nameMatch[1].trim();
    if (indexMatch && indexMatch[1].trim() && indexMatch[1] !== "/") c.index = indexMatch[1].trim();
    if (addrMatch && addrMatch[1].trim()) c.address = addrMatch[1].trim();
    if (assgMatch && assgMatch[1].trim()) c.assigned = assgMatch[1].trim();
    if (landlordMatch && landlordMatch[1].trim()) c.landlord = landlordMatch[1].trim();
    for (const m of actionMatches) {
      // Normalize both MM/DD/YYYY and the shorter M/D/YY into the same
      // ISO date — YY is always 20YY here, every date in this report
      // falls within this system's operating years.
      const parts = m[1].split("/");
      const mm = parts[0].padStart(2, "0"), dd = parts[1].padStart(2, "0");
      const yyyy = parts[2].length === 2 ? "20" + parts[2] : parts[2];
      const isoDate = `${yyyy}-${mm}-${dd}`;
      const desc = m[2].trim();
      // Skip an exact date+description repeat — the same action line
      // shouldn't be logged twice if it somehow appears on more than one
      // page for this case.
      if (!c.actions.some(a => a.date === isoDate && a.desc === desc)) {
        c.actions.push({ date: isoDate, desc });
      }
    }
  }
  // Newest action first, matching how every other log in the app displays.
  // latestActionDate/Desc are derived from this actual sort, not from
  // whichever block happened to be processed first — a case's actions
  // aren't reliably newest-first within a single page, so trusting "the
  // first match found" could point at an older action instead of the
  // genuinely most recent one.
  Object.values(casesByNumber).forEach(c => {
    c.actions.sort((a, b) => b.date.localeCompare(a.date));
    if (c.actions.length > 0) {
      c.latestActionDate = c.actions[0].date;
      c.latestActionDesc = c.actions[0].desc;
    }
  });
  return Object.values(casesByNumber);
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
  let pdf;
  try {
    pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  } catch (e) {
    // If the worker itself failed to load (a CDN outage, a URL that
    // stopped resolving) rather than the PDF being genuinely unreadable,
    // one retry against a different CDN can recover from it instead of
    // failing every single upload on what's really a single point of
    // failure. Only worth trying once — if this also fails, the error is
    // almost certainly the PDF itself, not the worker.
    if (pdfjsLib.GlobalWorkerOptions.workerSrc !== PDF_WORKER_FALLBACK) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_FALLBACK;
      pdf = await pdfjsLib.getDocument({ data: buffer.slice(0) }).promise;
    } else {
      throw e;
    }
  }
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

function IconBtn({ onClick, title, children, danger, active }) {
  return (
    <button className={`icon-btn ${danger ? "icon-btn-danger" : ""} ${active ? "icon-btn-active" : ""}`} onClick={onClick} title={title} type="button">
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

function PhotoUploader({ photos, onAdd, onRemove, pathPrefix }) {
  const ref = useRef(null);
  const [uploading, setUploading] = useState(false);
  return (
    <div className="photo-uploader">
      <div className="photo-grid">
        {(photos || []).map(p => (
          <div className="photo-thumb" key={p.id}>
            <img src={p.url || p.dataUrl} alt={p.name} />
            <button className="photo-remove" onClick={() => onRemove(p)} title="Remove"><X size={12} /></button>
          </div>
        ))}
        <button className="photo-add" onClick={() => ref.current.click()} type="button" disabled={uploading}>
          <Camera size={16} />
        </button>
      </div>
      {uploading && <div className="hint">Uploading…</div>}
      <input
        ref={ref} type="file" accept="image/*" multiple hidden
        onChange={async (e) => {
          if (!e.target.files.length) return;
          setUploading(true);
          try {
            const uploaded = await uploadFilesToStorage(e.target.files, pathPrefix, { compressImages: true });
            onAdd(uploaded);
          } catch (err) { console.error("photo upload failed", err); }
          setUploading(false);
          e.target.value = "";
        }}
      />
    </div>
  );
}

function DocumentUploader({ documents, onAdd, onRemove, pathPrefix }) {
  const ref = useRef(null);
  const [uploading, setUploading] = useState(false);
  return (
    <div className="doc-uploader">
      {(documents || []).map(d => (
        <div className="doc-chip" key={d.id}>
          <a href={d.url || d.dataUrl} download={d.name} target={d.url ? "_blank" : undefined} rel="noreferrer" className="doc-chip-name" title={d.name}>{d.name}</a>
          <button className="doc-remove" onClick={() => onRemove(d)} title="Remove"><X size={12} /></button>
        </div>
      ))}
      <button className="btn-ghost" type="button" onClick={() => ref.current.click()} disabled={uploading}>
        <Upload size={14} /> {uploading ? "Uploading…" : "Upload document"}
      </button>
      <input
        ref={ref} type="file" multiple hidden
        onChange={async (e) => {
          if (!e.target.files.length) return;
          setUploading(true);
          try {
            const uploaded = await uploadFilesToStorage(e.target.files, pathPrefix);
            onAdd(uploaded);
          } catch (err) { console.error("document upload failed", err); }
          setUploading(false);
          e.target.value = "";
        }}
      />
    </div>
  );
}

/* ============================== login ============================== */

const PIN_STORAGE_KEY = "pm-ops-device-pin";

function LoginScreen() {
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
        <div className="brand-mark" style={{ marginBottom: 14 }}>O</div>
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
        <div className="brand-mark" style={{ marginBottom: 14 }}>O</div>
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
  // Always mirrors the latest `data` — needed so a save that gets queued
  // while another is still in flight can read the truly current data when
  // it actually runs, not a stale value captured back when it was queued.
  const dataRef = useRef(data);
  useEffect(() => { dataRef.current = data; }, [data]);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [query, setQuery] = useState("");
  const [navOpen, setNavOpen] = useState(false);
  const [pinUnlocked, setPinUnlocked] = useState(() => !localStorage.getItem(PIN_STORAGE_KEY));
  const [showPinSetup, setShowPinSetup] = useState(false);
  const [saveError, setSaveError] = useState(false);
  const [loadError, setLoadError] = useState(false);
  const saveTimer = useRef(null);
  // Saves are debounced 400ms so typing doesn't hammer Firestore on every
  // keystroke — but that also means a change sits unsaved for up to 400ms
  // (plus however long the actual write takes). Closing the tab or hitting
  // back in that window would lose it silently with zero warning. This
  // tracks whether a save is still owed and, if so, asks the browser to
  // confirm before actually leaving.
  const hasPendingSave = useRef(false);
  // Tracks whether a real (existing) document has been loaded at least
  // once this session — used to tell a genuinely brand-new account (no
  // document yet, safe to start empty) apart from a later snapshot that
  // suspiciously claims the document vanished after real data was already
  // showing (almost certainly a transient glitch, not an actual deletion).
  const hasLoadedRealDataOnce = useRef(false);
  // The save-effect below depends on `loaded`, so it re-fires the instant
  // `loaded` flips from false to true after the initial snapshot settles —
  // even though nothing was actually edited. Left unguarded, that fires a
  // real save of whatever `data` happens to be at that exact moment,
  // including — in the worst case — a misdiagnosed "document missing, so
  // start empty" state for an existing account with real data on the
  // server, silently overwriting it. This ref skips exactly that one,
  // load-driven save; only a genuine subsequent edit schedules one.
  const justLoaded = useRef(true);
  // A write to Firestore fully replaces the document rather than merging —
  // so if two saves ever end up in flight at once and the earlier-started
  // one happens to take longer (a slow connection, a brief hiccup — exactly
  // what this app has to tolerate, being used on mobile), it could land
  // AFTER the faster, newer one and silently overwrite it with stale data.
  // These two refs keep saves strictly sequential: never more than one in
  // flight, and if a newer save becomes owed while one is still running,
  // it's queued rather than started concurrently, then fires immediately
  // (using whatever is truly latest at that moment, via dataRef) the
  // instant the in-flight one finishes.
  const saveInFlight = useRef(false);
  const saveQueued = useRef(false);
  // When the current unsaved-changes streak started — used to tell a
  // normal, brief in-flight save (a few hundred ms, completely routine)
  // apart from one that's been stuck for a while (a real connection
  // problem worth surfacing). null whenever nothing is currently owed.
  const pendingSince = useRef(null);
  // hasPendingSave is a ref on purpose (refs don't trigger re-renders,
  // which is fine for beforeunload/sign-out since those only ever read it
  // at the moment of leaving) — but the dashboard banner below needs
  // something reactive to actually show and hide with. This mirrors the
  // same "has it been stuck too long" signal in real React state instead.
  const [saveStuck, setSaveStuck] = useState(false);
  const [docSizeWarning, setDocSizeWarning] = useState(false);
  const SAVE_STUCK_THRESHOLD_MS = 8000;
  useEffect(() => {
    const interval = setInterval(() => {
      if (hasPendingSave.current && pendingSince.current && Date.now() - pendingSince.current > SAVE_STUCK_THRESHOLD_MS) {
        setSaveStuck(true);
      } else if (!hasPendingSave.current) {
        setSaveStuck(false);
      }
    }, 2000);
    return () => clearInterval(interval);
  }, []);

  useEffect(() => onAuthStateChanged(auth, u => setUser(u || null)), []);

  useEffect(() => {
    // Reset unconditionally on ANY user change, including the transition
    // to signed-out — not just when a new user is present. Both loaded and
    // data need to flip immediately: loaded so the save effect's own guard
    // (`if (!loaded...) return`) correctly blocks it without depending on
    // effect-ordering timing, and data so that even in the worst case
    // where something slips past that guard anyway, it would save empty
    // data rather than the previous account's private data into whatever
    // account signs in next in the same browser session.
    setLoaded(false);
    setData(emptyData());
    hasPendingSave.current = false;
    pendingSince.current = null;
    setSaveStuck(false);
    justLoaded.current = true;
    if (!user) return;
    setLoadError(false);
    hasLoadedRealDataOnce.current = false;
    // A live listener instead of a one-time fetch — using this on both a
    // phone and a laptop in the same sitting means each device needs to see
    // the other's changes, not just whatever was saved the last time this
    // device happened to load the page.
    const unsub = onSnapshot(
      doc(db, "users", user.uid, "appData", "main"),
      (snap) => {
        // This fires again the moment OUR OWN save lands too (Firestore
        // echoes every write back through the listener) — hasPendingWrites
        // is true for that echo, since it reflects our own write before the
        // server has fully confirmed it. Our local state is already ahead
        // of that echo, so there's nothing new to apply — skip it.
        if (snap.metadata.hasPendingWrites) { setLoaded(true); return; }
        // If this device has a local edit queued but not saved yet, don't
        // let an update arriving from the other device stomp on it —
        // that's the one case where applying a remote update immediately
        // would actively discard something the user just typed here.
        if (hasPendingSave.current) { setLoaded(true); return; }
        if (snap.exists()) {
          // A plain {...emptyData(), ...snap.data()} spread would let an
          // explicit null in any field (a stray manual edit in the Firestore
          // console, or some future bug) silently override the safe []
          // default — and the very next .map()/.filter() anywhere on that
          // field would crash the entire app with a blank screen, not just
          // one feature. Keep the empty-array default for any field that's
          // null or undefined instead of trusting whatever's in the doc.
          const loaded = snap.data();
          const merged = { ...emptyData() };
          Object.keys(merged).forEach(k => {
            if (loaded[k] !== null && loaded[k] !== undefined) merged[k] = loaded[k];
          });
          setData(merged);
          hasLoadedRealDataOnce.current = true;
        } else if (!hasLoadedRealDataOnce.current && !snap.metadata.fromCache) {
          // Genuinely a brand-new account with nothing saved yet — this is
          // the only situation where treating "no document" as "start
          // empty" is actually safe. Requires a server-confirmed read, not
          // a cached one: a snapshot served from local cache before the
          // network round-trip completes could plausibly report "not
          // found" for a document that actually exists on the server, and
          // this determination is irreversible enough that an unconfirmed
          // signal isn't good enough grounds for it.
          setData(emptyData());
        } else if (!hasLoadedRealDataOnce.current) {
          // First read said "missing" but came from cache, not the server —
          // wait for the server-confirmed snapshot instead of acting on this
          // one. Leave data as-is (still emptyData() from the reset above)
          // and don't mark loaded yet, so nothing saves in the meantime.
          return;
        } else {
          // Real data was already loaded once this session, so a later
          // snapshot claiming the document is gone doesn't add up — that's
          // far more likely a transient hiccup than an actual deletion.
          // Trusting it would wipe local state, and autosave would then
          // write that emptiness back to Firestore moments later, turning
          // a momentary blip into permanent data loss. Leave local data
          // exactly as it is and surface the problem instead of acting on
          // a signal this suspicious.
          console.error("Snapshot reported the data document missing after data had already loaded — ignoring rather than clearing local state.");
          setLoadError(true);
        }
        setLoaded(true);
      },
      (e) => {
        console.error("load failed", e);
        setLoadError(true);
        setLoaded(true);
      }
    );
    return () => unsub();
  }, [user]);

  const runSave = useCallback(async () => {
    if (saveInFlight.current) {
      // Something is already saving — don't start a second write
      // concurrently. Mark that another one is owed; the in-flight save's
      // own completion below will fire it immediately once this one is
      // done, reading dataRef fresh at that point rather than whatever was
      // current back when this call was made.
      saveQueued.current = true;
      return;
    }
    saveInFlight.current = true;
    try {
      // Firestore rejects any field that's literally `undefined` (as opposed to
      // just missing) and throws instead of saving anything. A JSON round-trip
      // strips those out automatically so a stray undefined somewhere can never
      // silently break autosave.
      const safe = JSON.parse(JSON.stringify(dataRef.current));
      // Firestore has a hard 1MB (1,048,576 byte) limit per document, and this
      // app keeps everything — every building, tenant, court case log, payment
      // history — in one single document. There's real room to grow into (this
      // is a rough character-count estimate, not Firestore's exact byte
      // accounting, but close enough to warn meaningfully early), but as more
      // gets logged over time this is a genuine wall the app would eventually
      // hit. Warning here, well before the hard limit, means it surfaces as an
      // actionable heads-up instead of every future save silently failing with
      // no explanation the moment the document tips over 1MB.
      const approxSize = JSON.stringify(safe).length;
      setDocSizeWarning(approxSize > 800000);
      await setDoc(doc(db, "users", user.uid, "appData", "main"), safe);
      setSaveError(false);
      // Only a confirmed success with nothing further queued means there's
      // truly nothing left owed — a failed save, or one where a newer
      // change already arrived while this one was running, needs to keep
      // the flag (and the leave-page warning it drives) reflecting that.
      if (!saveQueued.current) { hasPendingSave.current = false; pendingSince.current = null; }
    } catch (e) {
      console.error("save failed", e);
      setSaveError(true);
    } finally {
      saveInFlight.current = false;
      if (saveQueued.current) {
        saveQueued.current = false;
        runSave();
      }
    }
  }, [user]);

  useEffect(() => {
    if (!loaded || !user) return;
    if (justLoaded.current) { justLoaded.current = false; return; }
    if (!hasPendingSave.current) pendingSince.current = Date.now();
    hasPendingSave.current = true;
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(runSave, 400);
    return () => clearTimeout(saveTimer.current);
  }, [data, loaded, user, runSave]);

  useEffect(() => {
    const handler = (e) => {
      if (hasPendingSave.current) {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", handler);
    return () => window.removeEventListener("beforeunload", handler);
  }, []);

  // generic collection ops
  const add = (col, item) => setData(d => ({ ...d, [col]: [...d[col], { id: uid(), ...item }] }));
  const update = (col, id, patch) => setData(d => ({ ...d, [col]: d[col].map(x => x.id === id ? { ...x, ...patch } : x) }));
  const remove = (col, id) => setData(d => ({ ...d, [col]: d[col].filter(x => x.id !== id) }));

  const buildingName = (id) => shortAddress(data.buildings.find(b => b.id === id)?.address) || "—";
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
    { key: "quicknotes", label: "Quick Notes / Reminder", icon: <Pencil size={16} /> },
  ];

  const searchResults = useMemo(() => {
    if (!query.trim()) return null;
    const q = query.toLowerCase();
    return {
      tenants: data.tenants.filter(t => t.name?.toLowerCase().includes(q) || buildingName(t.buildingId).toLowerCase().includes(q)),
      violations: data.violations.filter(v => (v.violationNumber || "").toLowerCase().includes(q) || (v.description || "").toLowerCase().includes(q) || buildingName(v.buildingId).toLowerCase().includes(q)),
      courtCases: data.courtCases.filter(c => (c.caseNumber || "").toLowerCase().includes(q) || tenantName(c.tenantId).toLowerCase().includes(q)),
      buildings: data.buildings.filter(b => (b.address || "").toLowerCase().includes(q)),
      appointments: data.appointments.filter(a => (a.type || "").toLowerCase().includes(q) || (a.notes || "").toLowerCase().includes(q) || buildingName(a.buildingId).toLowerCase().includes(q)),
      workOrders: data.workOrders.filter(w => (w.description || "").toLowerCase().includes(q) || buildingName(w.buildingId).toLowerCase().includes(q) || (w.status || "").toLowerCase().includes(q)),
      vendors: data.vendors.filter(v => (v.name || "").toLowerCase().includes(q) || (v.specialty || "").toLowerCase().includes(q) || (v.phone || "").toLowerCase().includes(q)),
      quickNotes: data.quickNotes.filter(n => (n.text || "").toLowerCase().includes(q) || buildingName(n.buildingId).toLowerCase().includes(q)),
    };
  }, [query, data]);

  // Sign-out is a button click, not a browser navigation — the existing
  // beforeunload warning only covers closing the tab or navigating away,
  // so without this, clicking Sign Out while an edit is still sitting in
  // its debounce window (or a save is actively failing) would discard it
  // silently, with none of the protection closing the tab already has.
  const handleSignOut = async () => {
    if (hasPendingSave.current) {
      clearTimeout(saveTimer.current);
      await runSave();
      const deadline = Date.now() + 5000;
      while (hasPendingSave.current && saveInFlight.current && Date.now() < deadline) {
        await new Promise(r => setTimeout(r, 100));
      }
      if (hasPendingSave.current) {
        const proceed = window.confirm(
          "Your last change hasn't saved yet (possibly a connection problem) — signing out now will lose it. Sign out anyway?"
        );
        if (!proceed) return;
      }
    }
    localStorage.removeItem(PIN_STORAGE_KEY);
    signOut(auth);
  };

  if (user === undefined) return <div className="app-shell"><div className="loading">Loading…</div><Styles /></div>;
  if (user === null) return <LoginScreen />;
  if (!pinUnlocked) return (
    <PinLockScreen
      onUnlock={() => setPinUnlocked(true)}
      onForgot={handleSignOut}
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
          <button className="brand" onClick={() => setTab("dashboard")} title="Go to Dashboard">
            <div className="brand-mark">O</div>
            <div>
              <div className="brand-title">Property Ops</div>
              <div className="brand-sub">{data.buildings.length} buildings tracked</div>
            </div>
          </button>
        </div>
        <div className="search-wrap">
          <Search size={16} className="search-icon" />
          <input
            className="search-input"
            placeholder="Search tenants, violations, work orders, cases, vendors, notes, addresses…"
            value={query}
            onChange={e => setQuery(e.target.value)}
          />
          {query && <button className="search-clear" onClick={() => setQuery("")}><X size={14} /></button>}
        </div>
        <div className="topbar-actions">
          <button className="btn-ghost no-print" title="Set a quick-unlock PIN for this device" onClick={() => setShowPinSetup(s => !s)} style={{ color: "#fff", borderColor: "rgba(255,255,255,0.3)" }}>
            {localStorage.getItem(PIN_STORAGE_KEY) ? "Change PIN" : "Set PIN"}
          </button>
          <button className="btn-ghost no-print" title="Sign out" onClick={handleSignOut} style={{ color: "#fff", borderColor: "rgba(255,255,255,0.3)" }}>
            Sign out
          </button>
        </div>
      </header>

      {loadError && (
        <div className="save-error-banner no-print">
          <AlertTriangle size={16} /> Couldn't load your data — check your internet connection, or a Firestore permissions issue. What's showing below may be empty or incomplete until this is fixed and you refresh.
        </div>
      )}

      {saveError && (
        <div className="save-error-banner no-print">
          <AlertTriangle size={16} /> Couldn't save your last change — check your internet connection. Your edits are still here on screen, but won't be there if you close the tab until this clears.
        </div>
      )}

      {docSizeWarning && (
        <div className="save-error-banner no-print">
          <AlertTriangle size={16} /> Your data is getting close to Firestore's 1MB storage limit for a single document. Saves are still working, but once that limit is hit, they'd start silently failing. Worth flagging to me soon so we can look at archiving older records or restructuring storage before it becomes a real problem.
        </div>
      )}

      {showPinSetup && (
        <PinSetupPanel onClose={() => setShowPinSetup(false)} />
      )}

      {searchResults ? (
        <SearchResults results={searchResults} buildingName={buildingName} onClose={() => setQuery("")} setTab={setTab} />
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
            {tab === "dashboard" && <Dashboard data={data} buildingName={buildingName} tenantName={tenantName} setTab={setTab} setData={setData} saveStuck={saveStuck} />}
            {tab === "buildings" && <BuildingsTab data={data} add={add} update={update} remove={remove} setData={setData} buildingName={buildingName} />}
            {tab === "rent" && <RentTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} setData={setData} />}
            {tab === "workorders" && <WorkOrdersTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} vendorName={vendorName} />}
            {tab === "violations" && <ViolationsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} vendorName={vendorName} setData={setData} />}
            {tab === "vendors" && <VendorsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} />}
            {tab === "court" && <CourtTab data={data} add={add} update={update} remove={remove} setData={setData} tenantName={tenantName} buildingName={buildingName} />}
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

function SearchResults({ results, buildingName, onClose, setTab }) {
  const total = results.tenants.length + results.violations.length + results.courtCases.length + results.buildings.length + results.appointments.length + results.workOrders.length + results.vendors.length + results.quickNotes.length;
  const goTo = (t) => { setTab(t); onClose(); };
  return (
    <div className="content" style={{ maxWidth: 900, margin: "0 auto" }}>
      <div className="search-results-head">
        <h2>{total} result{total === 1 ? "" : "s"}</h2>
        <button className="btn-ghost" onClick={onClose}>Clear search</button>
      </div>
      {results.buildings.length > 0 && (
        <Section icon={<Building2 size={16} />} title="Buildings" count={results.buildings.length}>
          {results.buildings.map(b => <button className="row search-result-row" key={b.id} onClick={() => goTo("buildings")}>{shortAddress(b.address)}</button>)}
        </Section>
      )}
      {results.tenants.length > 0 && (
        <Section icon={<Users size={16} />} title="Tenants" count={results.tenants.length}>
          {results.tenants.map(t => (
            <button className="row search-result-row" key={t.id} onClick={() => goTo("rent")}>
              <strong>{t.name}</strong> — {buildingName(t.buildingId)} · {t.status}
            </button>
          ))}
        </Section>
      )}
      {results.violations.length > 0 && (
        <Section icon={<AlertTriangle size={16} />} title="Violations" count={results.violations.length}>
          {results.violations.map(v => (
            <button className="row search-result-row" key={v.id} onClick={() => goTo("violations")}>
              <strong>{v.agency}</strong> #{v.violationNumber} — {buildingName(v.buildingId)} · {v.status}
            </button>
          ))}
        </Section>
      )}
      {results.workOrders.length > 0 && (
        <Section icon={<Wrench size={16} />} title="Work Orders" count={results.workOrders.length}>
          {results.workOrders.map(w => (
            <button className="row search-result-row" key={w.id} onClick={() => goTo("workorders")}>
              {w.description} — {buildingName(w.buildingId)} · {w.status}
            </button>
          ))}
        </Section>
      )}
      {results.courtCases.length > 0 && (
        <Section icon={<Gavel size={16} />} title="Court Cases" count={results.courtCases.length}>
          {results.courtCases.map(c => (
            <button className="row search-result-row" key={c.id} onClick={() => goTo("court")}>
              #{c.caseNumber || "—"} — {c.result}
            </button>
          ))}
        </Section>
      )}
      {results.appointments.length > 0 && (
        <Section icon={<CalendarClock size={16} />} title="Appointments" count={results.appointments.length}>
          {results.appointments.map(a => (
            <button className="row search-result-row" key={a.id} onClick={() => goTo("inspections")}>
              <strong>{a.type}</strong> — {buildingName(a.buildingId)}{a.date ? ` · ${fmtDate(a.date)}` : ""}{a.completed ? " · Completed" : ""}
            </button>
          ))}
        </Section>
      )}
      {results.vendors.length > 0 && (
        <Section icon={<Wrench size={16} />} title="Vendors" count={results.vendors.length}>
          {results.vendors.map(v => (
            <button className="row search-result-row" key={v.id} onClick={() => goTo("vendors")}>
              <strong>{v.name}</strong>{v.specialty ? ` — ${v.specialty}` : ""}
            </button>
          ))}
        </Section>
      )}
      {results.quickNotes.length > 0 && (
        <Section icon={<StickyNote size={16} />} title="Quick Notes" count={results.quickNotes.length}>
          {results.quickNotes.map(n => (
            <button className="row search-result-row" key={n.id} onClick={() => goTo("quicknotes")}>
              {n.text}{n.buildingId ? ` — ${buildingName(n.buildingId)}` : ""}
            </button>
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
  inspections: "Appointments", quicknotes: "Quick Notes / Reminder",
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

// Pulls together every dated item across the whole app — violations,
// court dates, payments due, appointments, tenant follow-ups, and reminders
// — with no single-date filter, so nothing (overdue included) can silently
// fall outside a visible window.
function allDatedItems(data, tenantName, buildingName) {
  const items = [];
  data.violations.forEach(v => {
    if (v.cureDeadline && !isViolationClosed(v)) {
      items.push({ key: `v-${v.id}`, date: v.cureDeadline, type: `${v.agency} cure deadline`, label: `#${v.violationNumber}`, sub: buildingName(v.buildingId), tab: "violations" });
    }
    if (v.hasHearing && v.hearingDate && !isViolationClosed(v)) {
      items.push({ key: `vh-${v.id}`, date: v.hearingDate, type: `${v.agency} hearing`, label: `#${v.violationNumber}`, sub: buildingName(v.buildingId), tab: "violations" });
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
  const today = todayISO();

  const all = allDatedItems(data, tenantName, buildingName);
  const itemsByDate = {};
  all.forEach(i => { (itemsByDate[i.date] = itemsByDate[i.date] || []).push(i); });

  const shift = (n) => {
    if (viewMode === "day") setRefDate(d => addDays(d, n));
    else if (viewMode === "week") setRefDate(d => addDays(d, n * 7));
    // Normalized to the 1st before adding — landing on, say, the 31st and
    // stepping forward a month would otherwise roll past a shorter month
    // entirely (Jan 31 + 1 month lands on Mar 3, silently skipping February).
    else if (viewMode === "month") setRefDate(d => addMonths(d.slice(0, 8) + "01", n));
    else setRefDate(d => addMonths(d.slice(0, 8) + "01", n * 12));
  };
  const pickDate = (d) => setRefDate(d);

  const weekDays = Array.from({ length: 7 }, (_, i) => addDays(addDays(refDate, -new Date(refDate + "T00:00:00").getDay()), i));
  const monthDays = monthGridDays(refDate);
  const monthLabel = new Date(refDate + "T00:00:00").toLocaleDateString("en-US", { month: "long", year: "numeric" });
  const yearNum = new Date(refDate + "T00:00:00").getFullYear();
  const monthCounts = Array.from({ length: 12 }, (_, m) => {
    const prefix = `${yearNum}-${String(m + 1).padStart(2, "0")}`;
    return all.filter(i => i.date.startsWith(prefix)).length;
  });

  const selectedItems = itemsByDate[refDate] || [];
  const selectedLabel = refDate === today ? "Today" : new Date(refDate + "T00:00:00").toLocaleDateString("en-US", { weekday: "long", month: "short", day: "numeric" });

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
    <div className="dash-calendar">
      <div className="dash-calendar-head">
        <button className="icon-btn" onClick={() => shift(-1)}><ChevronLeft size={15} /></button>
        <span className="dash-calendar-title">
          {viewMode === "month" ? monthLabel : viewMode === "year" ? yearNum : viewMode === "week" ? `Week of ${fmtDate(weekDays[0])}` : new Date(refDate + "T00:00:00").toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" })}
        </span>
        <button className="icon-btn" onClick={() => shift(1)}><ChevronRight size={15} /></button>
      </div>

      <div className="filter-row dash-cal-modes">
        {["day", "week", "month", "year"].map(m => (
          <button key={m} className={`chip ${viewMode === m ? "chip-active" : ""}`} onClick={() => setViewMode(m)}>{m[0].toUpperCase() + m.slice(1)}</button>
        ))}
        {refDate !== today && <button className="btn-ghost" onClick={() => pickDate(today)}>Today</button>}
      </div>

      {viewMode === "week" && (
        <div className="dash-cal-week">
          {weekDays.map(d => {
            const dayItems = itemsByDate[d] || [];
            return (
              <button key={d} className={`dash-cal-cell dash-cal-cell-week ${d === refDate ? "dash-cal-cell-selected" : ""} ${d === today ? "dash-cal-cell-today" : ""}`} onClick={() => pickDate(d)}>
                <div className="dash-cal-cell-label">{new Date(d + "T00:00:00").toLocaleDateString("en-US", { weekday: "short" })}</div>
                <div className="dash-cal-cell-num">{new Date(d + "T00:00:00").getDate()}</div>
                {dayItems.length > 0 && <span className="dash-cal-count">{dayItems.length}</span>}
                <div className="dash-cal-cell-items">
                  {dayItems.slice(0, 3).map(item => <div key={item.key} className="dash-cal-cell-item">{item.label}</div>)}
                  {dayItems.length > 3 && <div className="dash-cal-cell-item row-muted">+{dayItems.length - 3} more</div>}
                </div>
              </button>
            );
          })}
        </div>
      )}

      {viewMode === "month" && (
        <div className="dash-cal-month">
          {["S", "M", "T", "W", "T", "F", "S"].map((d, i) => <div key={i} className="dash-cal-month-dow">{d}</div>)}
          {monthDays.map(d => {
            const inMonth = d.slice(0, 7) === refDate.slice(0, 7);
            const count = (itemsByDate[d] || []).length;
            return (
              <button key={d} className={`dash-cal-cell dash-cal-cell-sm ${d === refDate ? "dash-cal-cell-selected" : ""} ${d === today ? "dash-cal-cell-today" : ""} ${!inMonth ? "dash-cal-cell-dim" : ""}`} onClick={() => pickDate(d)}>
                <div className="dash-cal-cell-num">{new Date(d + "T00:00:00").getDate()}</div>
                {count > 0 && <span className="dash-cal-count">{count}</span>}
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
        <div className="dash-cal-section-title">{selectedLabel}</div>
        {selectedItems.length === 0 ? <div className="hint">Nothing here.</div> : selectedItems.map(Row)}
      </div>
    </div>
  );
}

function Dashboard({ data: rawData, buildingName, tenantName, setTab, setData, saveStuck }) {
  const [violationsPanelOpen, setViolationsPanelOpen] = useState(false);
  const [excludedSectionOpen, setExcludedSectionOpen] = useState(false);
  // Buildings marked "Mitch's father" are excluded from every main
  // computation below — every reference to `data.X` throughout this
  // component already reads from this scoped version, so shadowing the
  // prop here is enough to keep them out of every stat, panel, and total
  // without touching each computation individually. They still get their
  // own small section further down, built straight from rawData.
  const mainBuildingIds = new Set(rawData.buildings.filter(b => !b.excludedOwner).map(b => b.id));
  const excludedBuildingsList = rawData.buildings.filter(b => b.excludedOwner);
  const data = {
    ...rawData,
    buildings: rawData.buildings.filter(b => mainBuildingIds.has(b.id)),
    units: rawData.units.filter(u => mainBuildingIds.has(u.buildingId)),
    tenants: rawData.tenants.filter(t => mainBuildingIds.has(t.buildingId)),
    violations: rawData.violations.filter(v => mainBuildingIds.has(v.buildingId)),
    workOrders: rawData.workOrders.filter(w => mainBuildingIds.has(w.buildingId)),
    // Court cases are never excluded by building owner — for court
    // purposes, every building is treated as a regular one.
    courtCases: rawData.courtCases,
    appointments: rawData.appointments.filter(a => mainBuildingIds.has(a.buildingId)),
    localLaws: (rawData.localLaws || []).filter(l => mainBuildingIds.has(l.buildingId)),
  };
  const [testEmailStatus, setTestEmailStatus] = useState(null); // null | "sending" | "sent" | "error"
  const [expandedBuilding, setExpandedBuilding] = useState(null);
  const [statOpen, setStatOpen] = useState(null);
  const [quickNoteText, setQuickNoteText] = useState("");
  const submitQuickNote = () => {
    if (!quickNoteText.trim()) return;
    setData(d => ({ ...d, quickNotes: [...(d.quickNotes || []), { id: uid(), text: quickNoteText.trim(), date: todayISO(), done: false, reminderDate: "", buildingId: "" }] }));
    setQuickNoteText("");
  };

  // Nudges to re-upload Aged Arrears once it's been a week — driven by
  // when the last import actually happened, not a fixed calendar day, so it
  // never nags right after you've just done it and never goes quiet if
  // you're actually behind.
  const arrearsImports = (data.importHistory || []).filter(h => h.type === "arrears");
  const lastArrearsImport = arrearsImports.length
    ? arrearsImports.reduce((latest, h) => h.date > latest.date ? h : latest)
    : null;
  const daysSinceArrearsImport = lastArrearsImport ? daysUntil(lastArrearsImport.date) !== null ? -daysUntil(lastArrearsImport.date) : null : null;
  const showArrearsNudge = data.tenants.length > 0 && (!lastArrearsImport || (daysSinceArrearsImport !== null && daysSinceArrearsImport >= 7));

  // Tenants with a follow-up already scheduled show up under "to follow up" —
  // no need to also flag them under "not current on rent", that's just noise.
  const overdueTenants = data.tenants.filter(t => t.status !== "Current" && tenantFollowUps(t).length === 0);
  const today = todayISO();
  // Each open violation goes into exactly ONE of these buckets, by agency, so it
  // never shows up twice on the dashboard. Window covers overdue + due within 10 days.
  // Every OPEN violation shows here now, not just ones due soon — this used
  // to only include violations due within 10 days, which matched the
  // building card's "open violations" count except when a violation had a
  // cure date further out, creating a confusing "it's open, why isn't it
  // here" gap. Showing everything open means this always matches that
  // count exactly, and the panel is still sorted soonest-due-first so
  // urgency isn't lost. Lead violations are excluded entirely — they're
  // always kept even past their deadline (see the HPD import), which would
  // otherwise clog this panel with mostly-overdue lead items that are
  // already being worked on separately; they're still fully visible on the
  // Violations page itself, just not competing for space here.
  const violationDue = (v) => !isViolationClosed(v) && !v.isLead;
  // Soonest cure deadline first; missing deadlines sort to the end — same
  // convention the Violations page itself already uses, so the order here
  // matches what you'd see over there too.
  const byCureDeadline = (a, b) => {
    const da = daysUntil(a.cureDeadline), db = daysUntil(b.cureDeadline);
    if (da === null && db === null) return 0;
    if (da === null) return 1;
    if (db === null) return -1;
    return da - db;
  };
  // Every distinct agency actually in use gets its own group — HPD, DSNY,
  // and whatever specific agency names ("Other" violations are filed
  // under — DOB, FDNY, ECB, custom ones, whatever) show up on real
  // violations. No agency is hardcoded as more important than another;
  // this mirrors the Violations page's own dynamic tabs exactly, so what
  // you see here always matches what's over there.
  const violationAgencyName = (v) => v.agency === "Other" ? (v.otherAgency || "Other") : v.agency;
  const openViolationsByAgency = {};
  data.violations.filter(violationDue).forEach(v => {
    const name = violationAgencyName(v);
    (openViolationsByAgency[name] = openViolationsByAgency[name] || []).push(v);
  });
  const violationAgencyGroups = Object.keys(openViolationsByAgency)
    .sort((a, b) => a.localeCompare(b))
    .map(name => ({ name, items: openViolationsByAgency[name].sort(byCureDeadline) }));
  // A stipulation case's normal resting state is having NO next court date —
  // proceedings are over, all that's left is the payment schedule (tracked
  // separately below). Treating that empty field as "needs attention" would
  // make every settled stipulation permanently show up here for no reason.
  // A case still actively in litigation with no date set is a real gap
  // though, so that one still counts.
  const courtItems = data.courtCases.filter(c => !c.archived && (c.result === "Stipulation (payment plan)" ? dateDueStrict(c.nextCourtDate) : dateDue(c.nextCourtDate)));
  const byHearingDate = (a, b) => {
    const da = daysUntil(a.hearingDate), db = daysUntil(b.hearingDate);
    if (da === null && db === null) return 0;
    if (da === null) return 1;
    if (db === null) return -1;
    return da - db;
  };
  const hearingItems = data.violations.filter(v => v.hasHearing && !isViolationClosed(v) && dateDue(v.hearingDate)).sort(byHearingDate);
  const recurringItems = data.appointments.filter(a => !a.completed && a.recurring && dateDue(a.date));
  const appointmentItems = data.appointments.filter(a => !a.completed && !a.recurring && dateDue(a.date));
  // A reminder set for later stays off the dashboard until that date actually
  // arrives — no need to see it every day until then, it'll show up on its own.
  const quickNoteItems = (data.quickNotes || []).filter(n => !n.done && (!n.reminderDate || n.reminderDate <= today));
  const bossReminderItems = data.bossReminders.filter(r => r.status !== "Done");
  // Same snooze rule for tenant follow-ups — only show once due (or overdue),
  // not the whole month in advance.
  const tenantsWithFollowUps = data.tenants.filter(t => tenantFollowUps(t).some(f => f.date && f.date <= today));
  const allFollowUps = tenantsWithFollowUps.length;
  // Flatten to one row per (tenant, follow-up) pair, since a tenant can have several.
  const followUpEntries = tenantsWithFollowUps
    .flatMap(t => tenantFollowUps(t).filter(f => f.date && f.date <= today).map(f => ({ tenant: t, followUp: f })))
    .sort((a, b) => (a.followUp.date || "").localeCompare(b.followUp.date || ""));
  const vacantUnits = data.units.filter(u => !data.tenants.some(t => t.unitId === u.id && !t.movedOut));
  // A new tenant created from an arrears import only ever has a name and
  // balance — never phone or email, since the report doesn't carry that.
  // Flagging specifically the ones sharing a unit with a moved-out tenant
  // (a real turnover, not just a data gap) is what's actually actionable:
  // go get this person's contact info.
  const newTenantsNeedingContact = data.tenants.filter(t =>
    !t.movedOut && !t.phone && !t.email &&
    data.tenants.some(other => other.unitId === t.unitId && other.movedOut)
  );

  const rentPanelCount = overdueTenants.length + allFollowUps;
  const totalViolationsDue = violationAgencyGroups.reduce((sum, g) => sum + g.items.length, 0);
  const totalAttention = rentPanelCount + recurringItems.length + appointmentItems.length + quickNoteItems.length + totalViolationsDue + bossReminderItems.length + hearingItems.length;

  // Top stat row + follow-up roster
  const allDated = allDatedItems(data, tenantName, buildingName);
  const overdueCount = allDated.filter(i => i.date < today).length;
  const openWorkOrders = data.workOrders.filter(w => w.status !== "Done");
  const openCourtCases = data.courtCases.filter(c => !c.archived);
  const clearBuildingIds = new Set(data.buildings.filter(b => {
    const hasViolation = data.violations.some(v => v.buildingId === b.id && !isViolationClosed(v) && !v.isLead);
    const hasLateTenant = data.tenants.some(t => t.buildingId === b.id && t.status !== "Current");
    const hasOpenWO = data.workOrders.some(w => w.buildingId === b.id && w.status !== "Done");
    const hasOpenCourt = data.courtCases.some(c => c.buildingId === b.id && !c.archived);
    return !hasViolation && !hasLateTenant && !hasOpenWO && !hasOpenCourt;
  }).map(b => b.id));
  const rosterEntries = followUpEntries.slice(0, 8);

  const exportAllData = () => {
    const buildingMap = Object.fromEntries(data.buildings.map(b => [b.id, shortAddress(b.address)]));
    const unitMap = Object.fromEntries(data.units.map(u => [u.id, u.unitNumber]));
    const tenantMap = Object.fromEntries(data.tenants.map(t => [t.id, t.name]));
    const vendorMap = Object.fromEntries(data.vendors.map(v => [v.id, v.name]));
    const lawNameMap = Object.fromEntries(LOCAL_LAWS.map(l => [l.key, l.name]));

    const wb = XLSX.utils.book_new();
    const addSheet = (name, rows) => {
      const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{}]);
      XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    };

    // Building + unit + tenant + the numbers all live on one sheet — no separate
    // Buildings/Units tabs, and no internal ids anywhere, just what's on screen.
    addSheet("Tenants", data.tenants.map(t => ({
      Building: buildingMap[t.buildingId] || "",
      Unit: unitMap[t.unitId] || "",
      Tenant: t.name || "",
      Phone: t.phone || "",
      Email: t.email || "",
      Balance: t.balance || "",
      "Monthly Rent": t.rentAmount || "",
      "Months Behind": (parseBalance(t.rentAmount) > 0 && parseBalance(t.balance) > 0) ? (parseBalance(t.balance) / parseBalance(t.rentAmount)).toFixed(1) : "",
      Status: t.status || "",
      "Latest Note": Array.isArray(t.notes) && t.notes.length ? t.notes[t.notes.length - 1].text : "",
      "Next Follow-up": (() => { const d = earliestFollowUpDate(t); return d ? fmtDate(d) : ""; })(),
    })));

    addSheet("Violations", data.violations.map(v => ({
      Building: buildingMap[v.buildingId] || "",
      Unit: unitMap[v.unitId] || "",
      Agency: v.agency === "Other" ? (v.otherAgency || "Other") : v.agency,
      "Violation #": v.violationNumber || "",
      Class: v.class || "",
      Description: v.description || "",
      "Cure Deadline": v.cureDeadline ? fmtDate(v.cureDeadline) : "",
      "Fine Amount": v.fineAmount || "",
      Company: v.company || "",
      Status: v.status || "",
      Vendor: vendorMap[v.vendorId] || "",
    })));

    addSheet("Work Orders", data.workOrders.map(w => ({
      Building: buildingMap[w.buildingId] || "",
      Unit: unitMap[w.unitId] || "",
      Description: w.description || "",
      Priority: w.priority || "",
      Status: w.status || "",
      Vendor: vendorMap[w.vendorId] || "",
      "Date Opened": w.dateOpened ? fmtDate(w.dateOpened) : "",
    })));

    addSheet("Court Cases", data.courtCases.map(c => ({
      Building: buildingMap[c.buildingId] || "",
      Unit: unitMap[c.unitId] || "",
      Tenant: tenantMap[c.tenantId] || "",
      "Docket #": c.caseNumber || "",
      Stage: c.stage || "",
      "Court Date": c.nextCourtDate ? fmtDate(c.nextCourtDate) : "",
      Result: c.result || "",
      "Stipulation Terms": c.stipulationTerms || "",
      "Next Payment Due": c.nextPaymentDue ? fmtDate(c.nextPaymentDue) : "",
      Status: c.archived ? "Closed" : "Active",
    })));

    addSheet("Appointments", data.appointments.map(a => ({
      Building: buildingMap[a.buildingId] || "",
      Unit: unitMap[a.unitId] || "",
      Type: a.type || "",
      Date: a.date ? fmtDate(a.date) : "",
      "Time From": a.timeFrom || "",
      "Time To": a.timeTo || "",
      Recurring: a.recurring ? "Yes" : "No",
      Completed: a.completed ? "Yes" : "No",
      Notes: a.notes || "",
    })));

    addSheet("Vendors", data.vendors.map(v => ({
      Name: v.name || "",
      Specialty: v.specialty || "",
      Phone: v.phone || "",
      Email: v.email || "",
    })));

    addSheet("Local Laws", data.localLaws.map(l => ({
      Building: buildingMap[l.buildingId] || "",
      Law: lawNameMap[l.lawKey] || l.lawKey || "",
      Deadline: l.deadline ? fmtDate(l.deadline) : "",
      Status: l.status || "",
    })));

    addSheet("Boss Reminders", data.bossReminders.map(r => ({
      Text: r.text || "",
      "Date Raised": r.dateRaised ? fmtDate(r.dateRaised) : "",
      Status: r.status || "",
    })));

    addSheet("Quick Notes", (data.quickNotes || []).map(n => ({
      Text: n.text || "",
      Date: n.date ? fmtDate(n.date) : "",
      "Reminder Date": n.reminderDate ? fmtDate(n.reminderDate) : "",
      Building: n.buildingId ? (buildingMap[n.buildingId] || "") : "",
      Done: n.done ? "Yes" : "No",
    })));

    XLSX.writeFile(wb, `property-ops-backup-${today}.xlsx`);
  };

  const sendTestEmail = async () => {
    setTestEmailStatus("sending");
    try {
      const fn = httpsCallable(functions, "sendTestDigestEmail");
      await fn();
      setTestEmailStatus("sent");
    } catch (e) {
      console.error("Test email failed", e);
      setTestEmailStatus("error");
    }
    setTimeout(() => setTestEmailStatus(null), 4000);
  };

  return (
    <div className="dashboard-page">
      {saveStuck && (
        <div className="unsaved-alert-banner no-print">
          <AlertTriangle size={16} />
          Not saved — a change has been waiting to save for a while. Check your connection; your latest changes are still only on this device.
        </div>
      )}
      <div className="page-head">
        <h1 className="page-title">Dashboard</h1>
        <div className="page-actions">
          <button className="btn-ghost" onClick={sendTestEmail} disabled={testEmailStatus === "sending"}>
            <Mail size={14} />
            {testEmailStatus === "sending" ? "Sending…" : testEmailStatus === "sent" ? "Sent!" : testEmailStatus === "error" ? "Failed — try again" : "Send test email"}
          </button>
          <button className="btn-ghost" onClick={exportAllData}><Download size={14} /> Export backup</button>
          <PrintButton label="Dashboard" />
        </div>
      </div>

      <div className="print-only">
        <div className="print-header">
          <div className="print-mark">O</div>
          <div className="print-header-text">
            <h1 className="print-title">Property Overview</h1>
            <div className="print-subtitle">As of {fmtDate(todayISO())}</div>
          </div>
        </div>

        <div className="print-stats-row">
          <div className="print-stat"><div className="print-stat-num">{overdueCount}</div><div>Overdue</div></div>
          <div className="print-stat"><div className="print-stat-num">{openWorkOrders.length}</div><div>Open work orders</div></div>
          <div className="print-stat"><div className="print-stat-num">{openCourtCases.length}</div><div>Open court cases</div></div>
          <div className="print-stat"><div className="print-stat-num">{clearBuildingIds.size}</div><div>Buildings clear</div></div>
        </div>

        {violationAgencyGroups.length > 0 && (
          <div className="print-section">
            <div className="print-section-head"><span>Open violations ({totalViolationsDue})</span></div>
            <table className="print-table">
              <thead><tr><th>Agency</th><th>Building</th><th>Violation #</th><th>Cure deadline</th></tr></thead>
              <tbody>
                {violationAgencyGroups.flatMap(g => g.items.map(v => (
                  <tr key={v.id}><td>{g.name}</td><td>{buildingName(v.buildingId)}</td><td>#{v.violationNumber}</td><td>{v.cureDeadline ? fmtDate(v.cureDeadline) : "not set"}</td></tr>
                )))}
              </tbody>
            </table>
          </div>
        )}

        {courtItems.length > 0 && (
          <div className="print-section">
            <div className="print-section-head"><span>Court dates due or overdue</span></div>
            <table className="print-table">
              <thead><tr><th>Building</th><th>Tenant</th><th>Next court date</th></tr></thead>
              <tbody>
                {courtItems.map(c => (
                  <tr key={c.id}><td>{buildingName(c.buildingId)}</td><td>{tenantName(c.tenantId)}</td><td>{c.nextCourtDate ? fmtDate(c.nextCourtDate) : "not set"}</td></tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {(appointmentItems.length > 0 || recurringItems.length > 0) && (
          <div className="print-section">
            <div className="print-section-head"><span>Appointments coming up</span></div>
            <table className="print-table">
              <thead><tr><th>Building</th><th>Type</th><th>Date</th></tr></thead>
              <tbody>
                {[...appointmentItems, ...recurringItems].map(a => (
                  <tr key={a.id}><td>{buildingName(a.buildingId)}</td><td>{a.type}</td><td>{a.date ? fmtDate(a.date) : "not set"}</td></tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        <div className="print-section">
          <div className="print-section-head"><span>By building</span></div>
          <table className="print-table">
            <thead><tr><th>Building</th><th>Open violations</th><th>Tenants behind</th><th>Open work orders</th><th>Open court cases</th></tr></thead>
            <tbody>
              {data.buildings.map(b => (
                <tr key={b.id}>
                  <td>{b.address}</td>
                  <td>{data.violations.filter(v => v.buildingId === b.id && !isViolationClosed(v)).length}</td>
                  <td>{data.tenants.filter(t => t.buildingId === b.id && t.status !== "Current").length}</td>
                  <td>{data.workOrders.filter(w => w.buildingId === b.id && w.status !== "Done").length}</td>
                  <td>{data.courtCases.filter(c => c.buildingId === b.id && !c.archived).length}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="print-footer">
          <span>Property Ops — Property Overview</span>
          <span>Generated {fmtDate(todayISO())}</span>
        </div>
      </div>

      <div className="dash-quick-note">
        <Pencil size={14} className="dash-quick-note-icon" />
        <input
          autoFocus
          placeholder="Quick note — type and hit Enter…"
          value={quickNoteText}
          onChange={e => setQuickNoteText(e.target.value)}
          onKeyDown={e => { if (e.key === "Enter") submitQuickNote(); }}
        />
      </div>

      {showArrearsNudge && (
        <button className="dash-arrears-nudge" onClick={() => setTab("rent")}>
          <DollarSign size={14} />
          {lastArrearsImport
            ? `It's been ${daysSinceArrearsImport} days since your last Aged Arrears upload — probably time for a new one.`
            : "No Aged Arrears report uploaded yet — head to Rent Collection to bring balances up to date."}
        </button>
      )}

      <div className="dash-stats-row">
        <button className="dash-stat-card" onClick={() => setStatOpen(s => s === "overdue" ? null : "overdue")}>
          <div className="dash-stat-num" style={{ color: "var(--danger)" }}>{overdueCount}</div>
          <div className="dash-stat-label">Overdue</div>
          <div className="dash-stat-sub">Past due date, any type</div>
        </button>
        <button className="dash-stat-card" onClick={() => setStatOpen(s => s === "workorders" ? null : "workorders")}>
          <div className="dash-stat-num" style={{ color: "var(--warn)" }}>{openWorkOrders.length}</div>
          <div className="dash-stat-label">Open work orders</div>
          <div className="dash-stat-sub">Not marked done</div>
        </button>
        <button className="dash-stat-card" onClick={() => setStatOpen(s => s === "courtcases" ? null : "courtcases")}>
          <div className="dash-stat-num">{openCourtCases.length}</div>
          <div className="dash-stat-label">Open court cases</div>
          <div className="dash-stat-sub">Active, not archived</div>
        </button>
        <button className="dash-stat-card" onClick={() => setStatOpen(s => s === "tenantsdue" ? null : "tenantsdue")}>
          <div className="dash-stat-num" style={{ color: "var(--danger)" }}>{overdueTenants.length}</div>
          <div className="dash-stat-label">Tenants due</div>
          <div className="dash-stat-sub">Not current on rent</div>
        </button>
      </div>

      {statOpen && (
        <div className="dash-stat-detail">
          {statOpen === "overdue" && (
            allDated.filter(i => i.date < today).length === 0
              ? <div className="hint">Nothing overdue.</div>
              : allDated.filter(i => i.date < today).sort((a, b) => a.date.localeCompare(b.date)).map(item => (
                <button key={item.key} className="dash-detail-item" onClick={() => setTab(item.tab)}>
                  <span className="pill pill-danger">{item.type}</span>
                  <div className="followup-item-main">
                    <div className="followup-item-name">{item.label}{item.sub ? <span className="row-muted"> — {item.sub}</span> : null}</div>
                  </div>
                  <span className="pill pill-muted">{fmtDate(item.date)}</span>
                </button>
              ))
          )}
          {statOpen === "workorders" && (
            openWorkOrders.length === 0
              ? <div className="hint">No open work orders.</div>
              : openWorkOrders.map(w => (
                <button key={w.id} className="dash-detail-item" onClick={() => setTab("workorders")}>
                  <span className={`pill ${w.priority === "Emergency" ? "pill-danger" : w.priority === "Urgent" ? "pill-warn" : "pill-muted"}`}>{w.priority}</span>
                  <div className="followup-item-main">
                    <div className="followup-item-name">{w.description}<span className="row-muted"> — {buildingName(w.buildingId)}</span></div>
                  </div>
                  <span className="pill pill-muted">{w.status}</span>
                </button>
              ))
          )}
          {statOpen === "courtcases" && (
            openCourtCases.length === 0
              ? <div className="hint">No open court cases.</div>
              : openCourtCases.map(c => (
                <button key={c.id} className="dash-detail-item" onClick={() => setTab("court")}>
                  <span className="pill pill-muted">{c.stage || "—"}</span>
                  <div className="followup-item-main">
                    <div className="followup-item-name">{tenantName(c.tenantId)}<span className="row-muted"> — {buildingName(c.buildingId)}</span></div>
                  </div>
                  {c.nextCourtDate && <span className="pill pill-muted">{fmtDate(c.nextCourtDate)}</span>}
                </button>
              ))
          )}
          {statOpen === "tenantsdue" && (
            overdueTenants.length === 0
              ? <div className="hint">No tenants behind on rent.</div>
              : overdueTenants.map(t => (
                <button key={t.id} className="dash-detail-item" onClick={() => setTab("rent")}>
                  <span className={`pill ${t.status === "Late" ? "pill-warn" : "pill-danger"}`}>{t.status}</span>
                  <div className="followup-item-main">
                    <div className="followup-item-name">{t.name}<span className="row-muted"> — {buildingName(t.buildingId)}</span></div>
                    {t.balance && <div className="followup-item-note">Balance: ${t.balance}</div>}
                  </div>
                </button>
              ))
          )}
        </div>
      )}

      <div className="dash-calendar-hero">
        <DashboardCalendar data={data} buildingName={buildingName} tenantName={tenantName} setTab={setTab} />
      </div>

      <div className="dash-col-main">
        {rosterEntries.length > 0 && (
          <div className="dash-roster">
            <div className="dash-roster-title">Follow up with ({followUpEntries.length})</div>
              {rosterEntries.map(({ tenant: t, followUp: f }) => (
                <div className="dash-roster-row" key={f.id}>
                  <div className="dash-roster-info">
                    <div className="dash-roster-name">{t.name}</div>
                    <div className="dash-roster-sub">{buildingName(t.buildingId)}{f.note ? ` — ${f.note}` : ""}{!f.note && t.balance ? ` — $${t.balance} behind` : ""}</div>
                  </div>
                  <div className="dash-roster-meta">
                    {t.phone && <a href={`tel:${t.phone}`} className="dash-roster-phone" onClick={(e) => e.stopPropagation()} title={t.phone}><Phone size={14} /></a>}
                    <span className="pill pill-muted">{fmtDate(f.date)}</span>
                  </div>
                </div>
              ))}
              {followUpEntries.length > rosterEntries.length && (
                <div className="row-muted" style={{ fontSize: 12, marginTop: 4 }}>
                  + {followUpEntries.length - rosterEntries.length} more not shown here
                </div>
              )}
              <button className="btn-ghost" style={{ marginTop: 6 }} onClick={() => setTab("rent")}>View in Rent Collection</button>
            </div>
          )}

      {totalAttention === 0 ? (
        <div className="all-clear"><CheckCircle2 size={18} /> Nothing needs attention right now.</div>
      ) : (
        <>

          {violationAgencyGroups.length > 0 && (
            <div className="followup-panel">
              <button className="followup-panel-head" onClick={() => setViolationsPanelOpen(o => !o)}>
                <AlertTriangle size={18} className="attention-icon" style={{ color: "var(--danger)" }} />
                <span className="attention-count">{violationAgencyGroups.reduce((sum, g) => sum + g.items.length, 0)}</span>
                <span className="attention-label">Violations — all open <span className="dash-panel-sub">(grouped by agency, sorted by cure deadline)</span></span>
                {violationsPanelOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
              </button>
              {violationsPanelOpen && (
                <div className="followup-panel-body">
                  {violationAgencyGroups.map((g, gi) => (
                    <div key={g.name} style={{ marginTop: gi === 0 ? 0 : 14 }}>
                      <div className="violations-group-heading">{g.name} <span className="dash-panel-sub">({g.items.length})</span></div>
                      {g.items.map(v => (
                        <button key={v.id} className="dash-detail-item" onClick={() => setTab("violations")}>
                          <Flag date={v.cureDeadline} />
                          <div className="followup-item-main">
                            <div className="followup-item-name">#{v.violationNumber} <span className="row-muted">— {buildingName(v.buildingId)}</span></div>
                          </div>
                        </button>
                      ))}
                    </div>
                  ))}
                  <button className="btn-ghost" style={{ marginTop: 10 }} onClick={() => setTab("violations")}>View in Violations</button>
                </div>
              )}
            </div>
          )}

          <AttentionPanel
            icon={<Gavel size={18} className="attention-icon" style={{ color: "var(--danger)" }} />}
            label={<>Hearings due or overdue <span className="dash-panel-sub">(within 7 days, or no date set)</span></>} items={hearingItems} tab="violations" setTab={setTab}
            itemKey={v => v.id}
            renderItem={v => (
              <>
                <Flag date={v.hearingDate} />
                <div className="followup-item-main">
                  <div className="followup-item-name">#{v.violationNumber} <span className="row-muted">— {buildingName(v.buildingId)}</span></div>
                  {v.hearingCompany && <div className="followup-item-note">{v.hearingCompany}</div>}
                </div>
              </>
            )}
          />

          <AttentionPanel
            icon={<CalendarClock size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label={<>Recurring inspections due or overdue <span className="dash-panel-sub">(within 7 days, or no date set)</span></>} items={recurringItems} tab="inspections" setTab={setTab}
            itemKey={a => a.id}
            renderItem={a => (
              <>
                <Flag date={a.date} />
                <div className="followup-item-main">
                  <div className="followup-item-name">
                    {a.type} <span className="row-muted">— {buildingName(a.buildingId)}{a.unitId ? ` (Unit ${data.units.find(u => u.id === a.unitId)?.unitNumber || "—"})` : ""}</span>
                  </div>
                  {(a.timeFrom || a.timeTo) && <div className="followup-item-note">{fmtTime(a.timeFrom)}{a.timeTo ? ` – ${fmtTime(a.timeTo)}` : ""}</div>}
                </div>
              </>
            )}
          />

          <AttentionPanel
            icon={<CalendarClock size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label={<>Appointments coming up <span className="dash-panel-sub">(within 7 days, or no date set)</span></>} items={appointmentItems} tab="inspections" setTab={setTab}
            itemKey={a => a.id}
            renderItem={a => (
              <>
                <Flag date={a.date} />
                <div className="followup-item-main">
                  <div className="followup-item-name">
                    {a.type} <span className="row-muted">— {buildingName(a.buildingId)}{a.unitId ? ` (Unit ${data.units.find(u => u.id === a.unitId)?.unitNumber || "—"})` : ""}</span>
                  </div>
                  {(a.timeFrom || a.timeTo) && <div className="followup-item-note">{fmtTime(a.timeFrom)}{a.timeTo ? ` – ${fmtTime(a.timeTo)}` : ""}</div>}
                </div>
              </>
            )}
          />

          <AttentionPanel
            icon={<Pencil size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label={<>Quick notes to organize <span className="dash-panel-sub">(shows once due, or always if no reminder set)</span></>} items={quickNoteItems} tab="quicknotes" setTab={setTab}
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
            label={<>Boss reminders <span className="dash-panel-sub">(shows until checked off)</span></>} items={bossReminderItems} tab="reminders" setTab={setTab}
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
            label={<>New tenants missing contact info <span className="dash-panel-sub">(replaced a moved-out tenant)</span></>} items={newTenantsNeedingContact} tab="rent" setTab={setTab}
            itemKey={t => t.id}
            renderItem={t => (
              <div className="followup-item-main">
                <div className="followup-item-name">{t.name || "(no name on file)"} <span className="row-muted">— Unit {data.units.find(u => u.id === t.unitId)?.unitNumber || "—"} {buildingName(t.buildingId)}</span></div>
              </div>
            )}
          />
        </>
      )}
        </div>

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

      {excludedBuildingsList.length > 0 && (
        <div style={{ marginTop: 24 }}>
          <button
            className="list-card-head"
            style={{ width: "100%", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 8, cursor: "pointer" }}
            onClick={() => setExcludedSectionOpen(o => !o)}
          >
            {excludedSectionOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
            <div className="list-card-title">Mitch's father ({excludedBuildingsList.length} building{excludedBuildingsList.length === 1 ? "" : "s"})</div>
            <span className="pill pill-muted">excluded from totals above</span>
          </button>
          {excludedSectionOpen && (
            <div className="dash-grid" style={{ marginTop: 10 }}>
              {excludedBuildingsList.map(b => {
                const bTenants = rawData.tenants.filter(t => t.buildingId === b.id);
                const bViolations = rawData.violations.filter(v => v.buildingId === b.id && !isViolationClosed(v));
                const bLate = bTenants.filter(t => t.status !== "Current");
                const bOwed = bTenants.reduce((sum, t) => sum + parseBalance(t.balance), 0);
                return (
                  <div className="dash-building-card" key={b.id}>
                    <div className="dash-building-name">{shortAddress(b.address)}</div>
                    {bViolations.length === 0 && bLate.length === 0 ? (
                      <span className="dash-building-clear">All clear</span>
                    ) : (
                      <div className="dash-building-chips">
                        {bViolations.length > 0 && <span className="pill pill-warn">{bViolations.length} open violations</span>}
                        {bLate.length > 0 && <span className="pill pill-danger">{bLate.length} tenants behind (${bOwed.toFixed(2)})</span>}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ============================== buildings ============================== */

function BuildingsTab({ data, add, update, remove, setData, buildingName }) {
  const [form, setForm] = useState(null);
  // Buildings default to collapsed — track which ones have been explicitly
  // expanded instead of which are closed, so opening the tab always starts
  // clean with nothing open until you choose to look inside one.
  const [expandedIds, setExpandedIds] = useState(new Set());
  const toggleExpanded = (id) => {
    setExpandedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  const [expandedUnit, setExpandedUnit] = useState(null);
  const [unitSearch, setUnitSearch] = useState("");
  const [section, setSection] = useState("buildings");
  const [pendingDelete, setPendingDelete] = useState(null);
  const fileRef = useRef(null);
  const [confirmingMerge, setConfirmingMerge] = useState(false);

  // Units sharing the same building + unit number are duplicates — this can
  // happen when the same apartment appeared as two separate report entries
  // (a departing tenant's line and a new arrival's line) before matching
  // logic recognized they belonged to one unit, so each independently
  // created its own unit record.
  const duplicateUnitGroups = useMemo(() => {
    const byKey = new Map();
    data.units.forEach(u => {
      const key = u.buildingId + "|" + (u.unitNumber || "").trim().toUpperCase();
      if (!byKey.has(key)) byKey.set(key, []);
      byKey.get(key).push(u);
    });
    return [...byKey.values()].filter(group => group.length > 1);
  }, [data.units]);
  const duplicateUnitCount = duplicateUnitGroups.reduce((sum, g) => sum + (g.length - 1), 0);

  const mergeDuplicateUnits = () => {
    setData(d => {
      let units = [...d.units], tenants = [...d.tenants], violations = [...d.violations],
        workOrders = [...d.workOrders], courtCases = [...d.courtCases];
      duplicateUnitGroups.forEach(group => {
        // Keep whichever unit in the group has the most tenants attached —
        // the one most likely to be the "real" one everything should have
        // been pointing at all along. Ties keep the first.
        const keeper = group.slice().sort((a, b) =>
          tenants.filter(t => t.unitId === b.id).length - tenants.filter(t => t.unitId === a.id).length
        )[0];
        const duplicateIds = new Set(group.filter(u => u.id !== keeper.id).map(u => u.id));
        tenants = tenants.map(t => duplicateIds.has(t.unitId) ? { ...t, unitId: keeper.id } : t);
        violations = violations.map(v => duplicateIds.has(v.unitId) ? { ...v, unitId: keeper.id } : v);
        workOrders = workOrders.map(w => duplicateIds.has(w.unitId) ? { ...w, unitId: keeper.id } : w);
        courtCases = courtCases.map(c => duplicateIds.has(c.unitId) ? { ...c, unitId: keeper.id } : c);
        units = units.filter(u => !duplicateIds.has(u.id));
      });
      return { ...d, units, tenants, violations, workOrders, courtCases };
    });
    setConfirmingMerge(false);
  };

  const deleteBuilding = (buildingId) => {
    setData(d => ({
      ...d,
      buildings: d.buildings.filter(b => b.id !== buildingId),
      units: d.units.filter(u => u.buildingId !== buildingId),
      tenants: d.tenants.filter(t => t.buildingId !== buildingId),
      violations: d.violations.filter(v => v.buildingId !== buildingId),
      workOrders: d.workOrders.filter(w => w.buildingId !== buildingId),
      appointments: d.appointments.filter(a => a.buildingId !== buildingId),
      localLaws: d.localLaws.filter(l => l.buildingId !== buildingId),
      // Court cases intentionally stay — same as deleting a tenant, a case is
      // a legal record that shouldn't silently disappear, it just loses its
      // building link.
    }));
    setPendingDelete(null);
  };
  const [pendingDeleteUnit, setPendingDeleteUnit] = useState(null);
  const deleteUnit = (unitId) => {
    setData(d => ({
      ...d,
      units: d.units.filter(u => u.id !== unitId),
      tenants: d.tenants.filter(t => t.unitId !== unitId),
      // Anything still pointing at this specific apartment (a violation, work
      // order, appointment) drops back to "whole building" instead of being
      // left pointing at a unit that no longer exists — keeps the record,
      // just loses the apartment-specific tag.
      violations: d.violations.map(v => v.unitId === unitId ? { ...v, unitId: "" } : v),
      workOrders: d.workOrders.map(w => w.unitId === unitId ? { ...w, unitId: "" } : w),
      appointments: d.appointments.map(a => a.unitId === unitId ? { ...a, unitId: "" } : a),
    }));
    setPendingDeleteUnit(null);
    setExpandedUnit(null);
  };
  const [pendingCleanup, setPendingCleanup] = useState(null);
  const cleanupEmptyUnits = (buildingId) => {
    setData(d => {
      const removedIds = new Set(d.units.filter(u => u.buildingId === buildingId && !d.tenants.some(t => t.unitId === u.id)).map(u => u.id));
      return {
        ...d,
        units: d.units.filter(u => !removedIds.has(u.id)),
        violations: d.violations.map(v => removedIds.has(v.unitId) ? { ...v, unitId: "" } : v),
        workOrders: d.workOrders.map(w => removedIds.has(w.unitId) ? { ...w, unitId: "" } : w),
        appointments: d.appointments.map(a => removedIds.has(a.unitId) ? { ...a, unitId: "" } : a),
      };
    });
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
            const address = csvField(row, "address", "building");
            if (!address) return;
            let building = next.buildings.find(b => b.address.toLowerCase() === address.toLowerCase());
            if (!building) {
              building = { id: uid(), address, notes: "" };
              next.buildings.push(building);
            }
            const unitNumber = csvField(row, "unit", "unitNumber");
            let unit = null;
            if (unitNumber) {
              unit = next.units.find(u => u.buildingId === building.id && u.unitNumber === unitNumber);
              if (!unit) {
                unit = { id: uid(), buildingId: building.id, unitNumber };
                next.units.push(unit);
              }
            }
            const tenantName = csvField(row, "tenant", "tenantName");
            if (tenantName && unit) {
              const exists = next.tenants.find(t => t.unitId === unit.id && t.name === tenantName);
              if (!exists) {
                next.tenants.push({
                  id: uid(), buildingId: building.id, unitId: unit.id, name: tenantName,
                  phone: csvField(row, "phone"), email: csvField(row, "email"),
                  balance: csvField(row, "balance"), status: csvField(row, "status") || "Current",
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

      {duplicateUnitGroups.length > 0 && (
        <div className="form-panel" style={{ borderColor: "var(--danger)", marginBottom: 16 }}>
          <div style={{ fontWeight: 700, marginBottom: 4 }}>
            {duplicateUnitCount} duplicate unit{duplicateUnitCount === 1 ? "" : "s"} found across {duplicateUnitGroups.length} apartment{duplicateUnitGroups.length === 1 ? "" : "s"}
          </div>
          <p className="hint">
            Same unit number showing up as two separate units — usually from an import that ran before a matching fix went in.
            Merging keeps whichever copy has more tenants attached, moves every tenant, violation, work order, and court case from the other copy onto it, then removes the duplicate.
          </p>
          <div style={{ marginBottom: 8 }}>
            {duplicateUnitGroups.map((group, i) => (
              <div key={i} className="row row-muted">{buildingName(group[0].buildingId)} — Apt {group[0].unitNumber || "—"} ({group.length} copies)</div>
            ))}
          </div>
          {confirmingMerge ? (
            <div className="form-actions">
              <span className="row-muted" style={{ fontSize: 12 }}>Merge all {duplicateUnitCount} duplicate{duplicateUnitCount === 1 ? "" : "s"} now? This can't be undone.</span>
              <button className="btn-primary" style={{ background: "var(--danger)", borderColor: "var(--danger)" }} onClick={mergeDuplicateUnits}>Yes, merge</button>
              <button className="btn-ghost" onClick={() => setConfirmingMerge(false)}>Cancel</button>
            </div>
          ) : (
            <button className="btn-primary" style={{ background: "var(--danger)", borderColor: "var(--danger)" }} onClick={() => setConfirmingMerge(true)}>Merge duplicate units</button>
          )}
        </div>
      )}

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

      {data.buildings.length > 0 && (
        <div className="inline-form" style={{ marginBottom: 14 }}>
          <input placeholder="Find a unit or tenant across every building…" value={unitSearch} onChange={e => setUnitSearch(e.target.value)} />
          {unitSearch && <button className="btn-ghost" onClick={() => setUnitSearch("")}>Clear</button>}
        </div>
      )}

      {data.buildings.length === 0 && <EmptyState text="No buildings yet — import a CSV or add one manually." />}
      {data.buildings.map(b => {
        const q = unitSearch.trim().toLowerCase();
        const allUnits = data.units.filter(u => u.buildingId === b.id);
        const units = q
          ? allUnits.filter(u => (u.unitNumber || "").toLowerCase().includes(q) || data.tenants.some(t => t.unitId === u.id && (t.name || "").toLowerCase().includes(q)))
          : allUnits;
        if (q && units.length === 0) return null;
        return (
          <div className="list-card" key={b.id}>
            <div className="list-card-head" onClick={() => toggleExpanded(b.id)}>
              {(q || expandedIds.has(b.id)) ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
              <div className="list-card-title">{b.address}</div>
              {b.excludedOwner && <span className="pill pill-warn">Mitch's father</span>}
              <span className="pill pill-muted">{q ? `${units.length} match${units.length === 1 ? "" : "es"} of ${allUnits.length}` : `${units.length} units`}</span>
              {(() => {
                const emptyCount = allUnits.filter(u => !data.tenants.some(t => t.unitId === u.id)).length;
                return emptyCount > 0 && <span className="pill pill-warn">{emptyCount} empty</span>;
              })()}
              <div className="spacer" />
              {pendingDelete === b.id ? (
                <>
                  <span className="row-muted" style={{ fontSize: 12, color: data.tenants.some(t => t.buildingId === b.id && (data.courtCases || []).some(c => c.tenantId === t.id && !c.archived)) ? "var(--danger)" : undefined }}>
                    Delete building + {allUnits.length} units + {data.tenants.filter(t => t.buildingId === b.id).length} tenants
                    {(() => {
                      const vN = data.violations.filter(v => v.buildingId === b.id).length;
                      const wN = data.workOrders.filter(w => w.buildingId === b.id).length;
                      const aN = data.appointments.filter(a => a.buildingId === b.id).length;
                      const parts = [];
                      if (vN) parts.push(`${vN} violation${vN === 1 ? "" : "s"}`);
                      if (wN) parts.push(`${wN} work order${wN === 1 ? "" : "s"}`);
                      if (aN) parts.push(`${aN} appointment${aN === 1 ? "" : "s"}`);
                      return parts.length ? ` + ${parts.join(" + ")}` : "";
                    })()}
                    {(() => {
                      const n = data.tenants.filter(t => t.buildingId === b.id && (data.courtCases || []).some(c => c.tenantId === t.id && !c.archived)).length;
                      return n > 0 ? ` — ${n} of them ${n === 1 ? "has" : "have"} an OPEN COURT CASE (kept, just unlinked)` : "";
                    })()}?
                  </span>
                  <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); deleteBuilding(b.id); }} style={{ color: "var(--danger)" }}>Yes, delete</button>
                  <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); setPendingDelete(null); }}>Cancel</button>
                </>
              ) : pendingCleanup === b.id ? (
                <>
                  <span className="row-muted" style={{ fontSize: 12 }}>Remove {allUnits.filter(u => !data.tenants.some(t => t.unitId === u.id)).length} empty units from this building?</span>
                  <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); cleanupEmptyUnits(b.id); }} style={{ color: "var(--danger)" }}>Yes, clean up</button>
                  <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); setPendingCleanup(null); }}>Cancel</button>
                </>
              ) : (
                <>
                  {allUnits.some(u => !data.tenants.some(t => t.unitId === u.id)) && (
                    <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); setPendingCleanup(b.id); }}>Clean up empty units</button>
                  )}
                  <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); update("buildings", b.id, { excludedOwner: !b.excludedOwner }); }}>
                    {b.excludedOwner ? "Unmark Mitch's father" : "Mark: Mitch's father"}
                  </button>
                  <IconBtn title="Edit" onClick={(e) => { e.stopPropagation(); setForm(b); }}><Pencil size={14} /></IconBtn>
                  <IconBtn title="Delete" danger onClick={(e) => { e.stopPropagation(); setPendingDelete(b.id); }}><Trash2 size={14} /></IconBtn>
                </>
              )}
            </div>
            {(q || expandedIds.has(b.id)) && (
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
                        <span className="unit-row-name">
                          {tenants.map(t => t.name + (t.movedOut ? " (moved out)" : "")).filter(Boolean).join(", ") || "no tenant on file"}
                        </span>
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
                              <div className="unit-detail-row"><strong>{t.name || "(no name on file)"}</strong>{t.movedOut && <span className="pill pill-warn" style={{ marginLeft: 6 }}>Moved Out</span>}</div>
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

// Stores multiple numbers in the same t.phone field as a single "; "-joined
// string rather than adding a new array field — every other place in the
// app that reads t.phone (CSV export, print views, Dashboard, the RIS
// import) keeps working unchanged, just showing the joined string if it
// ever displays phone outside this cell. Click the arrow to cycle through
// saved numbers one at a time; click + to add another.
function PhoneCycleCell({ tenant, update }) {
  const [idx, setIdx] = useState(0);
  const numbers = (tenant.phone || "").split(";").map(s => s.trim());
  const safeIdx = Math.min(idx, numbers.length - 1);
  const current = numbers[safeIdx] || "";

  const setNumbers = (next) => update("tenants", tenant.id, { phone: next.join("; ") });
  const updateCurrent = (val) => {
    const next = [...numbers];
    next[safeIdx] = val;
    setNumbers(next);
  };
  const addNumber = () => {
    setNumbers([...numbers, ""]);
    setIdx(numbers.length);
  };
  const removeCurrent = () => {
    if (numbers.length <= 1) { updateCurrent(""); return; }
    const next = numbers.filter((_, i) => i !== safeIdx);
    setNumbers(next);
    setIdx(i => Math.min(i, next.length - 1));
  };
  const cycle = () => setIdx(i => (i + 1) % numbers.length);

  return (
    <div className="phone-cycle-cell">
      <input className="sheet-input" value={current} onChange={e => updateCurrent(e.target.value)} placeholder="phone" onClick={e => e.stopPropagation()} />
      {numbers.length > 1 && (
        <>
          <button type="button" className="phone-cycle-btn" onClick={e => { e.stopPropagation(); cycle(); }} title={`Number ${safeIdx + 1} of ${numbers.length} — click for next`}>
            {safeIdx + 1}/{numbers.length} <ChevronRight size={11} />
          </button>
          <button type="button" className="phone-cycle-btn" onClick={e => { e.stopPropagation(); removeCurrent(); }} title="Remove this number">
            <X size={11} />
          </button>
        </>
      )}
      <button type="button" className="phone-cycle-btn" onClick={e => { e.stopPropagation(); addNumber(); }} title="Add another number">
        <Plus size={11} />
      </button>
    </div>
  );
}

function RentTab({ data: rawData, add, update, remove, buildingName, setData }) {
  const [excludedSectionOpen, setExcludedSectionOpen] = useState(false);
  const mainBuildingIds = new Set(rawData.buildings.filter(b => !b.excludedOwner).map(b => b.id));
  const excludedBuildingsList = rawData.buildings.filter(b => b.excludedOwner);
  const data = {
    ...rawData,
    buildings: rawData.buildings.filter(b => mainBuildingIds.has(b.id)),
    units: rawData.units.filter(u => mainBuildingIds.has(u.buildingId)),
    tenants: rawData.tenants.filter(t => mainBuildingIds.has(t.buildingId)),
    courtCases: rawData.courtCases,
  };
  const [noteFor, setNoteFor] = useState(null);
  const [noteText, setNoteText] = useState("");
  const [followFor, setFollowFor] = useState(null);
  const [newFollowDate, setNewFollowDate] = useState("");
  const [newFollowNote, setNewFollowNote] = useState("");
  const [payFor, setPayFor] = useState(null);
  const [newPayAmount, setNewPayAmount] = useState("");
  const [newPayNote, setNewPayNote] = useState("");
  const [statusFilter, setStatusFilter] = useState("All");
  const [sortMode, setSortMode] = useState("building"); // building | balance | oldest
  const [expandedBuildings, setExpandedBuildings] = useState(new Set());
  const [section, setSection] = useState("sheet");

  const inCourt = (tenantId) => (data.courtCases || []).some(c => c.tenantId === tenantId && !c.archived);
  // Creates a linked court case for the tenant — the moment one exists,
  // inCourt(t.id) becomes true, which is what shows the red "In Court" pill
  // on their row and makes them show up under the In Court filter; there's
  // no separate flag to keep in sync with the case itself.
  const markInCourt = (tenantId) => {
    if (inCourt(tenantId)) return;
    const t = data.tenants.find(x => x.id === tenantId);
    if (!t) return;
    add("courtCases", {
      tenantId, buildingId: t.buildingId, unitId: t.unitId, caseNumber: "",
      stage: CASE_STAGES[0], nextCourtDate: "", result: "Pending",
      stipulationTerms: "", nextPaymentDue: "",
      archived: false, documents: [],
      checklist: DEFAULT_ATTORNEY_CHECKLIST.map(label => ({ id: uid(), label, checked: false })),
    });
  };
  const [selectedTenantId, setSelectedTenantId] = useState(null);

  const [newTenant, setNewTenant] = useState(null);

  const openNewTenant = () => setNewTenant({
    buildingId: data.buildings[0]?.id || "", unitId: "", name: "", phone: "", balance: "", rentAmount: "", status: "Current",
  });
  const saveNewTenant = () => {
    if (!newTenant.buildingId || !newTenant.unitId) return;
    add("tenants", { ...newTenant, email: "", notes: [], followUps: [], payments: [] });
    setNewTenant(null);
  };

  // 3 = money sitting in the oldest (61+) bucket, 2 = middle bucket, 1 = only
  // the newest bucket, 0 = nothing owed. Tenants without an aging breakdown
  // (added manually, not from an Aged Arrears import) fall back to their
  // status instead of defaulting to 0, so a manually-flagged "In Arrears"
  // tenant doesn't rank the same as someone who owes nothing.
  const agingSeverity = (t) => {
    if (t.aging) {
      if (t.aging.bucket61 > 0) return 3;
      if (t.aging.bucket2 > 0) return 2;
      if (t.aging.bucket1 > 0) return 1;
      return 0;
    }
    if (t.status === "In Arrears") return 3;
    if (t.status === "Late") return 2;
    return 0;
  };

  const unitOf = t => data.units.find(u => u.id === t.unitId)?.unitNumber || "";

  const passesFilters = (t) => {
    if (statusFilter === "Follow-ups") { if (tenantFollowUps(t).length === 0) return false; }
    else if (statusFilter === "61+") { if (agingSeverity(t) < 3) return false; }
    else if (statusFilter === "In Court") { if (!inCourt(t.id)) return false; }
    else if (statusFilter !== "All" && t.status !== statusFilter) return false;
    return true;
  };

  const sortTenants = (list) => {
    if (sortMode === "balance") return [...list].sort((a, b) => parseBalance(b.balance) - parseBalance(a.balance));
    if (sortMode === "balanceLow") return [...list].sort((a, b) => parseBalance(a.balance) - parseBalance(b.balance));
    if (sortMode === "oldest") return [...list].sort((a, b) => agingSeverity(b) - agingSeverity(a) || parseBalance(b.balance) - parseBalance(a.balance));
    return [...list].sort((a, b) => compareUnits(unitOf(a), unitOf(b)));
  };

  // One group per building that actually has at least one tenant, so the
  // list of buildings shown doesn't jump around as filters change — a
  // filter can empty a building's visible tenant list, but the building
  // itself stays put with an empty-state message inside.
  const buildingGroups = data.buildings
    .filter(b => data.tenants.some(t => t.buildingId === b.id && !t.movedOut))
    .map(b => {
      const allTenantsHere = data.tenants.filter(t => t.buildingId === b.id && !t.movedOut);
      const tenantsHere = sortTenants(allTenantsHere.filter(passesFilters));
      const totalOwed = tenantsHere.reduce((sum, t) => sum + parseBalance(t.balance), 0);
      const oldTenants = tenantsHere.filter(t => agingSeverity(t) === 3);
      const oldTotal = oldTenants.reduce((sum, t) => sum + parseBalance(t.balance), 0);
      return { building: b, tenants: tenantsHere, totalCount: allTenantsHere.length, totalOwed, oldCount: oldTenants.length, oldTotal };
    });

  const totalOwedAll = buildingGroups.reduce((sum, g) => sum + g.totalOwed, 0);
  const visibleTenants = buildingGroups.flatMap(g => g.tenants);
  const callBackTenants = rawData.tenants.filter(t => t.callBack);
  const movedOutTenants = data.tenants.filter(t => t.movedOut);
  const movedOutBuildingGroups = data.buildings
    .filter(b => movedOutTenants.some(t => t.buildingId === b.id))
    .map(b => {
      const tenantsHere = movedOutTenants.filter(t => t.buildingId === b.id).sort((a, b2) => compareUnits(unitOf(a), unitOf(b2)));
      const totalOwed = tenantsHere.reduce((sum, t) => sum + parseBalance(t.balance), 0);
      return { building: b, tenants: tenantsHere, totalOwed };
    });
  const [mainView, setMainView] = useState("current");

  const toggleBuildingExpanded = (id) => setExpandedBuildings(prev => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });

  const exportCSV = () => {
    const rows = visibleTenants.map(t => ({
      Building: buildingName(t.buildingId),
      Unit: unitOf(t),
      Tenant: t.name,
      Phone: t.phone,
      Balance: t.balance,
      "Monthly Rent": t.rentAmount || "",
      "Months Behind": (parseBalance(t.rentAmount) > 0 && parseBalance(t.balance) > 0) ? (parseBalance(t.balance) / parseBalance(t.rentAmount)).toFixed(1) : "",
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
  const paymentsArr = (t) => Array.isArray(t.payments) ? t.payments : [];

  // A ref, not state — state updates are batched/async, so a second click
  // firing before React re-renders would still see the old (empty) guard
  // and slip through. A ref updates immediately, so it actually blocks a
  // fast double-click on the same button from logging the same note,
  // follow-up, or payment twice.
  const noteSubmitting = useRef(false);
  const addNote = (tenantId) => {
    if (!noteText.trim() || noteSubmitting.current) return;
    noteSubmitting.current = true;
    const t = data.tenants.find(x => x.id === tenantId);
    const existingNotes = Array.isArray(t.notes) ? t.notes : (t.notes ? [{ date: todayISO(), text: t.notes }] : []);
    update("tenants", tenantId, { notes: [...existingNotes, { date: todayISO(), text: noteText }] });
    setNoteText(""); setNoteFor(null);
    setTimeout(() => { noteSubmitting.current = false; }, 400);
  };

  const followSubmitting = useRef(false);
  const addFollowUp = (tenantId) => {
    if (!newFollowDate || followSubmitting.current) return;
    followSubmitting.current = true;
    const t = data.tenants.find(x => x.id === tenantId);
    update("tenants", tenantId, { followUps: [...tenantFollowUps(t), { id: uid(), date: newFollowDate, note: newFollowNote }] });
    setNewFollowDate(""); setNewFollowNote(""); setFollowFor(null);
    setTimeout(() => { followSubmitting.current = false; }, 400);
  };
  const removeFollowUp = (tenantId, followUpId) => {
    const t = data.tenants.find(x => x.id === tenantId);
    update("tenants", tenantId, { followUps: tenantFollowUps(t).filter(f => f.id !== followUpId) });
  };

  // Logging a payment both records it (so there's a dated trail of what came
  // in and when) and moves the balance — no more overtyping one number with
  // nothing to show why it changed. Removing a logged payment puts the money
  // back on the balance, so undoing a mistaken entry doesn't leave the
  // balance wrong.
  const paySubmitting = useRef(false);
  const addPayment = (tenantId) => {
    const amt = parseBalance(newPayAmount);
    if (!amt || paySubmitting.current) return;
    paySubmitting.current = true;
    const t = data.tenants.find(x => x.id === tenantId);
    const newBalance = (parseBalance(t.balance) - amt).toFixed(2);
    update("tenants", tenantId, {
      balance: newBalance,
      payments: [...paymentsArr(t), { id: uid(), date: todayISO(), amount: amt.toFixed(2), note: newPayNote }],
    });
    setNewPayAmount(""); setNewPayNote(""); setPayFor(null);
    setTimeout(() => { paySubmitting.current = false; }, 400);
  };
  const removePayment = (tenantId, paymentId) => {
    const t = data.tenants.find(x => x.id === tenantId);
    const p = paymentsArr(t).find(x => x.id === paymentId);
    if (!p) return;
    const newBalance = (parseBalance(t.balance) + parseBalance(p.amount)).toFixed(2);
    update("tenants", tenantId, { balance: newBalance, payments: paymentsArr(t).filter(x => x.id !== paymentId) });
  };

  const renderTenantTable = (tenants) => (
    <div className="sheet-wrap">
      <table className="sheet">
        <thead>
          <tr>
            <th>Unit</th><th>Tenant</th><th>Phone</th>
            <th className="sheet-col-balance">Balance</th><th>Status</th>
            <th className="sheet-col-notes">Notes</th><th>Follow-up</th><th>Payments</th><th></th>
          </tr>
        </thead>
        <tbody>
          {tenants.length === 0 && (
            <tr><td colSpan={9}><div className="hint" style={{ padding: "10px 4px" }}>No tenants match this filter.</div></td></tr>
          )}
          {tenants.map(t => (
            <React.Fragment key={t.id}>
              <tr
                className={`${t.status !== "Current" ? "sheet-row-flag" : ""} ${selectedTenantId === t.id ? "sheet-row-selected" : ""}`}
                onClick={() => setSelectedTenantId(selectedTenantId === t.id ? null : t.id)}
              >
                <td className="sheet-readonly">{unitOf(t) || "—"}</td>
                <td>
                  <div className="sheet-name-cell">
                    <input className="sheet-input" value={t.name} onChange={e => update("tenants", t.id, { name: e.target.value })} />
                    {inCourt(t.id) && <span className="pill pill-danger sheet-court-pill" title="Active court case — no need to independently follow up">In Court</span>}
                  </div>
                </td>
                <td><PhoneCycleCell tenant={t} update={update} /></td>
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
                <td>
                  <button className="sheet-follow-btn" onClick={() => setPayFor(payFor === t.id ? null : t.id)} title="Payment history">
                    <DollarSign size={14} />
                    {paymentsArr(t).length > 0 && <span className="sheet-follow-date">{paymentsArr(t).length}</span>}
                  </button>
                </td>
                <td className="sheet-actions">
                  <IconBtn title="Notes" onClick={() => setNoteFor(noteFor === t.id ? null : t.id)}><Pencil size={14} /></IconBtn>
                  <IconBtn
                    title={t.callBack ? "Flagged to call back — click to clear" : "Flag to call back (didn't answer, try again — not a dated follow-up)"}
                    active={!!t.callBack}
                    onClick={() => update("tenants", t.id, { callBack: !t.callBack })}
                  ><Phone size={14} /></IconBtn>
                  <IconBtn
                    title={inCourt(t.id) ? "Already in court — a case is already linked to this tenant" : "Mark in court — creates a linked court case"}
                    active={inCourt(t.id)}
                    onClick={() => markInCourt(t.id)}
                  ><Gavel size={14} /></IconBtn>
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
              {payFor === t.id && (
                <tr className="sheet-expand-row">
                  <td colSpan={9}>
                    <strong>Payments</strong>
                    <div className="row" style={{ display: "flex", alignItems: "center", gap: 8, margin: "6px 0" }}>
                      <span style={{ fontSize: 13 }}>Monthly rent:</span>
                      <div className="balance-input-wrap" style={{ maxWidth: 110 }}>
                        <span className="balance-dollar">$</span>
                        <input className="balance-input" placeholder="—" value={t.rentAmount || ""} onChange={e => update("tenants", t.id, { rentAmount: e.target.value })} />
                      </div>
                      {parseBalance(t.rentAmount) > 0 && parseBalance(t.balance) > 0 && (
                        <span className="row-muted" style={{ fontSize: 12 }}>
                          ≈ {(parseBalance(t.balance) / parseBalance(t.rentAmount)).toFixed(1)} months behind
                        </span>
                      )}
                    </div>
                    {t.aging && (
                      <div className="hint" style={{ marginTop: 4 }}>
                        Aging breakdown from last import: ${t.aging.bucket1.toFixed(2)} current · ${t.aging.bucket2.toFixed(2)} 31–60 days · ${t.aging.bucket61.toFixed(2)} 61+ days
                      </div>
                    )}
                    <div className="inline-form">
                      <div className="balance-input-wrap" style={{ maxWidth: 120 }}>
                        <span className="balance-dollar">$</span>
                        <input className="balance-input" placeholder="Amount" value={newPayAmount} onChange={e => setNewPayAmount(e.target.value)} />
                      </div>
                      <input placeholder="Note (optional)" value={newPayNote} onChange={e => setNewPayNote(e.target.value)} style={{ flex: 1 }} onKeyDown={e => e.key === "Enter" && addPayment(t.id)} />
                      <button className="btn-primary" onClick={() => addPayment(t.id)}>Log payment</button>
                    </div>
                    <p className="hint" style={{ margin: "4px 0" }}>Logging a payment subtracts it from the balance above and keeps a dated record here. A weekly Aged Arrears re-import logs entries here automatically too — a lower balance than last time shows as a payment, a higher one as a charge.</p>
                    {paymentsArr(t).length === 0 && <div className="hint">No payments logged yet.</div>}
                    {paymentsArr(t).slice().reverse().map((p) => {
                      const amt = parseBalance(p.amount);
                      const isCharge = amt < 0;
                      return (
                        <div key={p.id} className="row row-muted" style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <span className={`pill ${isCharge ? "pill-warn" : "pill-ok"}`}>{isCharge ? "Charge" : "Payment"}: ${Math.abs(amt).toFixed(2)}</span>
                          <span>{fmtDate(p.date)}{p.auto ? " · from import" : ""}{p.note ? ` — ${p.note}` : ""}</span>
                          <button className="checklist-remove" onClick={() => removePayment(t.id, p.id)} title="Remove (reverses its effect on the balance)"><X size={12} /></button>
                        </div>
                      );
                    })}
                  </td>
                </tr>
              )}
            </React.Fragment>
          ))}
        </tbody>
      </table>
    </div>
  );

  return (
    <div className="rent-page">
      <div className="page-head">
        <h1 className="page-title">Rent Collection</h1>
        <div className="page-actions">
          <PrintButton label="Rent Collection" />
          <button className="btn-ghost" onClick={exportCSV}><Download size={14} /> Export CSV</button>
          <button className="btn-primary" onClick={openNewTenant}><Plus size={14} /> Add tenant</button>
        </div>
      </div>

      <div className="print-only">
        <div className="print-header">
          <div className="print-mark">O</div>
          <div className="print-header-text">
            <h1 className="print-title">Rent Ledger</h1>
            <div className="print-subtitle">As of {fmtDate(todayISO())}</div>
          </div>
        </div>
        <div className="print-stats-row">
          <div className="print-stat"><div className="print-stat-num">{buildingGroups.length}</div><div>Buildings</div></div>
          <div className="print-stat"><div className="print-stat-num">{buildingGroups.reduce((sum, g) => sum + g.tenants.length, 0)}</div><div>Tenants</div></div>
          <div className="print-stat"><div className="print-stat-num">${totalOwedAll.toLocaleString("en-US", { maximumFractionDigits: 0 })}</div><div>Total owed</div></div>
        </div>
        {buildingGroups.map(g => (
          <div className="print-section" key={g.building.id}>
            <div className="print-section-head">
              <span>{g.building.address}</span>
              <span>{g.tenants.length} tenant{g.tenants.length === 1 ? "" : "s"} — ${g.totalOwed.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} owed</span>
            </div>
            <table className="print-table">
              <thead>
                <tr><th>Unit</th><th>Tenant</th><th>Status</th><th>0–30 days</th><th>31–60 days</th><th>61+ days</th><th>Total owed</th></tr>
              </thead>
              <tbody>
                {g.tenants.map(t => (
                  <tr key={t.id}>
                    <td>{unitOf(t)}</td>
                    <td>{t.name || "(no name on file)"}</td>
                    <td>{t.status}</td>
                    <td>{t.aging ? `$${t.aging.bucket1.toFixed(2)}` : "—"}</td>
                    <td>{t.aging ? `$${t.aging.bucket2.toFixed(2)}` : "—"}</td>
                    <td>{t.aging ? `$${t.aging.bucket61.toFixed(2)}` : "—"}</td>
                    <td className="print-table-amount">${parseBalance(t.balance).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ))}
        <div className="print-grand-total">Total owed across all buildings: ${totalOwedAll.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</div>
        <div className="print-footer">
          <span>Property Ops — Rent Ledger</span>
          <span>Generated {fmtDate(todayISO())}</span>
        </div>
      </div>

      <div className="filter-row">
        <button className={`chip ${section === "sheet" ? "chip-active" : ""}`} onClick={() => setSection("sheet")}>Sheet</button>
        <button className={`chip ${section === "import" ? "chip-active" : ""}`} onClick={() => setSection("import")}>Import Aged Arrears</button>
      </div>

      {section === "sheet" && (
        <div className="filter-row">
          <button className={`chip ${mainView === "current" ? "chip-active" : ""}`} onClick={() => setMainView("current")}>Current</button>
          <button className={`chip ${mainView === "movedOut" ? "chip-active" : ""}`} onClick={() => setMainView("movedOut")}>
            Moved Out{movedOutTenants.length > 0 ? ` (${movedOutTenants.length})` : ""}
          </button>
        </div>
      )}

      {section === "import" ? (
        <ImportSection data={data} setData={setData} buildingName={buildingName} allowedTypes={["arrears"]} />
      ) : mainView === "current" ? (
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
          <Field label="Monthly rent (optional)">
            <div className="balance-input-wrap">
              <span className="balance-dollar">$</span>
              <input className="balance-input" value={newTenant.rentAmount} onChange={e => setNewTenant({ ...newTenant, rentAmount: e.target.value })} />
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

      {data.tenants.length === 0 ? (
        <EmptyState text="No tenants yet — import RIS data or add a tenant." />
      ) : (
        <>
          {statusFilter !== "Call Back" && (
            <div className="rent-total-banner">
              <div className="rent-total-num">${totalOwedAll.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</div>
              <div className="rent-total-label">total owed across {visibleTenants.length} tenant{visibleTenants.length === 1 ? "" : "s"}{statusFilter !== "All" ? " (matching current filters)" : ""}</div>
            </div>
          )}

          <div className="filter-row">
            {/* "In Arrears" intentionally left out here — that status is
                defined as "the 61+ day bucket has money", the exact same
                condition "61+ days only" below already filters on, so the
                two chips always showed identical results. Still a valid
                status value elsewhere (the dropdowns), just redundant as
                its own filter chip. */}
            {["All", ...RENT_STATUSES.filter(s => s !== "In Arrears")].map(s => (
              <button key={s} className={`chip ${statusFilter === s ? "chip-active" : ""}`} onClick={() => setStatusFilter(s)}>{s}</button>
            ))}
            <button className={`chip ${statusFilter === "61+" ? "chip-active" : ""}`} onClick={() => setStatusFilter("61+")}>61+ days only</button>
            <span className="row-muted" style={{ margin: "0 2px", fontSize: 14 }}>|</span>
            <button className={`chip ${statusFilter === "Follow-ups" ? "chip-active" : ""}`} onClick={() => setStatusFilter("Follow-ups")}>Follow-ups</button>
            <button className={`chip ${statusFilter === "Call Back" ? "chip-active" : ""}`} onClick={() => setStatusFilter("Call Back")}>
              Call Back{callBackTenants.length > 0 ? ` (${callBackTenants.length})` : ""}
            </button>
            <button className={`chip ${statusFilter === "In Court" ? "chip-active" : ""}`} onClick={() => setStatusFilter("In Court")}>In Court</button>
          </div>
          <div className="filter-row">
            <span className="row-muted" style={{ fontSize: 12, marginRight: 2 }}>Sort:</span>
            <button className={`chip ${sortMode === "building" ? "chip-active" : ""}`} onClick={() => setSortMode("building")}>By unit</button>
            <button className={`chip ${sortMode === "balance" ? "chip-active" : ""}`} onClick={() => setSortMode("balance")}>Highest balance</button>
            <button className={`chip ${sortMode === "balanceLow" ? "chip-active" : ""}`} onClick={() => setSortMode("balanceLow")}>Lowest balance</button>
            <button className={`chip ${sortMode === "oldest" ? "chip-active" : ""}`} onClick={() => setSortMode("oldest")}>Oldest debt</button>
          </div>

          {statusFilter === "Call Back" ? (
            // Call Back intentionally bypasses buildingGroups (which excludes
            // "Mitch's father" buildings) and shows every flagged tenant
            // regardless — it's meant to be a complete reminder list, not
            // scoped by any other business rule, matching how it behaved as
            // its own dedicated section before this became a filter chip.
            <div className="list-card">
              <div className="list-card-head">
                <div className="list-card-title">Call Back</div>
                <span className="pill pill-muted">tenants flagged to call back — not a dated follow-up, just a "didn't answer, try again" list</span>
              </div>
              <div className="list-card-body" style={{ padding: "10px 14px 14px" }}>
                {callBackTenants.length === 0
                  ? <div className="hint">No one flagged for a call back right now.</div>
                  : renderTenantTable(callBackTenants)}
              </div>
            </div>
          ) : buildingGroups.map(g => {
            const isOpen = expandedBuildings.has(g.building.id);
            return (
              <div className="list-card" key={g.building.id}>
                <div className="list-card-head" onClick={() => toggleBuildingExpanded(g.building.id)} style={{ cursor: "pointer" }}>
                  {isOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                  <div className="list-card-title">{shortAddress(g.building.address)}</div>
                  <span className="pill pill-muted">{g.tenants.length} tenant{g.tenants.length === 1 ? "" : "s"}</span>
                  <span className={`pill ${g.totalOwed > 0 ? "pill-warn" : "pill-ok"}`}>${g.totalOwed.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} owed</span>
                  {g.oldCount > 0 && (
                    <span className="pill pill-danger">
                      {g.oldCount} tenant{g.oldCount === 1 ? "" : "s"} 61+ days overdue for ${g.oldTotal.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </span>
                  )}
                </div>
                {isOpen && (
                  <div className="list-card-body" style={{ padding: "10px 14px 14px" }}>
                    {renderTenantTable(g.tenants)}
                  </div>
                )}
              </div>
            );
          })}
        </>
      )}
      </>
      ) : null}



      {section === "sheet" && mainView === "movedOut" && (
        movedOutBuildingGroups.length === 0 ? (
          <EmptyState text="No moved-out tenants yet." />
        ) : movedOutBuildingGroups.map(g => {
          const isOpen = expandedBuildings.has("movedOut_" + g.building.id);
          return (
            <div className="list-card" key={g.building.id}>
              <div className="list-card-head" onClick={() => toggleBuildingExpanded("movedOut_" + g.building.id)} style={{ cursor: "pointer" }}>
                {isOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                <div className="list-card-title">{shortAddress(g.building.address)}</div>
                <span className="pill pill-muted">{g.tenants.length} tenant{g.tenants.length === 1 ? "" : "s"}</span>
                <span className={`pill ${g.totalOwed > 0 ? "pill-warn" : "pill-ok"}`}>${g.totalOwed.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} owed</span>
              </div>
              {isOpen && (
                <div className="list-card-body" style={{ padding: "10px 14px 14px" }}>
                  {renderTenantTable(g.tenants)}
                </div>
              )}
            </div>
          );
        })
      )}

      {excludedBuildingsList.length > 0 && (
        <div style={{ marginTop: 24 }}>
          <button
            className="list-card-head"
            style={{ width: "100%", background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 8, cursor: "pointer" }}
            onClick={() => setExcludedSectionOpen(o => !o)}
          >
            {excludedSectionOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
            <div className="list-card-title">Mitch's father ({excludedBuildingsList.length} building{excludedBuildingsList.length === 1 ? "" : "s"})</div>
            <span className="pill pill-muted">excluded from the total above</span>
          </button>
          {excludedSectionOpen && excludedBuildingsList.map(b => {
            const bTenants = rawData.tenants.filter(t => t.buildingId === b.id);
            const bOwed = bTenants.reduce((sum, t) => sum + parseBalance(t.balance), 0);
            return (
              <div className="list-card" key={b.id} style={{ marginTop: 10 }}>
                <div className="list-card-head">
                  <div className="list-card-title">{shortAddress(b.address)}</div>
                  <span className="pill pill-muted">{bTenants.length} tenant{bTenants.length === 1 ? "" : "s"}</span>
                  <span className={`pill ${bOwed > 0 ? "pill-warn" : "pill-ok"}`}>${bOwed.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} owed</span>
                </div>
                <div className="list-card-body" style={{ padding: "10px 14px 14px" }}>
                  {renderTenantTable(bTenants)}
                </div>
              </div>
            );
          })}
        </div>
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
  const hasActiveCourtCase = tenantId => (data.courtCases || []).some(c => c.tenantId === tenantId && !c.archived);
  const changes = [];
  const normalizeName = n => (n || "").trim().toUpperCase();
  // Two report entries sharing one apt happen routinely — one line for who's
  // leaving, one for who's now there. Matching entries independently (as an
  // earlier version of this did) let the wrong entry claim an existing
  // tenant's record: an unrelated new arrival's name/balance could overwrite
  // a departing tenant's identity, and the departing tenant's own entry
  // would then flag that now-overwritten record as moved-out — silently
  // erasing the real departing tenant and mislabeling the real new arrival.
  // Grouping by apt and matching by exact name FIRST (claiming that tenant
  // so no other entry in the same batch can also claim them) fixes this: an
  // entry only falls back to "this is the existing active tenant, just with
  // a new name" when it's the one and only unmatched entry for that apt,
  // matched against the one and only remaining unclaimed active tenant —
  // the genuine single-name-change turnover case.
  const entriesByApt = new Map();
  parsedEntries.forEach(entry => {
    const key = normalizeName(entry.apt);
    if (!entriesByApt.has(key)) entriesByApt.set(key, []);
    entriesByApt.get(key).push(entry);
  });
  const claimedTenantIds = new Set();
  // Shared placeholder for a unit that doesn't exist yet, so multiple
  // entries for the same not-yet-existing apt attach to one unit once
  // created on confirm, not a separate duplicate each.
  const pendingNewUnits = new Map();

  entriesByApt.forEach((entriesForApt, aptKey) => {
    const unit = unitsForBuilding.find(u => normalizeName(u.unitNumber) === aptKey);
    const candidates = unit ? data.tenants.filter(t => t.unitId === unit.id && !claimedTenantIds.has(t.id)) : [];

    // Pass 1: exact name match per entry, regardless of active/moved-out
    // status — this is what lets Maheen's own asterisk-marked line find
    // and update her own existing record even while Dharmesh's separate,
    // non-matching line is present in the same batch.
    const matchedTenant = new Map();
    entriesForApt.forEach(entry => {
      const match = candidates.find(t => !claimedTenantIds.has(t.id) && t.name && entry.name && normalizeName(t.name) === normalizeName(entry.name));
      if (match) { claimedTenantIds.add(match.id); matchedTenant.set(entry, match); }
    });

    // Pass 2: exactly one leftover entry and exactly one leftover active
    // tenant means a genuine single turnover with a changed name — anything
    // more ambiguous than that falls through to "new" rather than guessing.
    const unmatchedEntries = entriesForApt.filter(e => !matchedTenant.has(e));
    if (unmatchedEntries.length === 1) {
      const remainingActive = candidates.filter(t => !t.movedOut && !claimedTenantIds.has(t.id));
      if (remainingActive.length === 1) {
        claimedTenantIds.add(remainingActive[0].id);
        matchedTenant.set(unmatchedEntries[0], remainingActive[0]);
      }
    }

    entriesForApt.forEach(entry => {
      let fields = {};
      if (type === "arrears") fields = { balance: entry.balance, status: entry.status, name: entry.name };
      if (type === "directory") fields = { name: entry.name };
      if (type === "contacts") {
        if (entry.phone) fields.phone = entry.phone;
        if (entry.email) fields.email = entry.email;
      }
      // aging (the 30/60/90+ bucket breakdown) rides along separately from
      // `fields` — fields feeds the diff/approval preview, which renders each
      // value with String(v), so an object there would show as "[object
      // Object]". aging gets applied straight through on confirm instead.
      const aging = type === "arrears" ? entry.aging : undefined;

      if (unit) {
        const tenant = matchedTenant.get(entry);
        if (tenant) {
          const diffFields = {};
          // Compare by actual meaning, not raw string equality — "1500" and
          // "1500.00" are the same balance, and "JANE DOE" and "JANE DOE "
          // are the same name; treating them as different would make a
          // re-import of literally the same report keep surfacing "changes"
          // that aren't real changes, just formatting drift from an older
          // import, a manual edit, or trailing whitespace.
          Object.entries(fields).forEach(([k, v]) => {
            if (!v) return;
            const current = tenant[k];
            const same = k === "balance"
              ? Math.abs(parseBalance(current) - parseBalance(v)) < 0.005
              : (current || "").toString().trim() === v.toString().trim();
            if (!same) diffFields[k] = v;
          });
          // A tenant whose only "change" is newly qualifying as moved-out
          // (report shows the asterisk, they're not flagged that way yet)
          // still needs to surface here even with zero field-level diffs —
          // the status flip itself is the meaningful change, not something
          // diffFields tracks.
          // The report's own legend says "* - MOVED OUT", and balances on
          // asterisk-marked units stay completely frozen across separate
          // report snapshots weeks apart — the signature of a departed
          // tenant's unpaid debt just sitting there, not someone still
          // accruing rent. Trust the asterisk directly for a tenant seen
          // for the first time, or when a new, different name shows up.
          // But a tenant already confirmed as the current occupant through
          // an earlier genuine turnover split is a different case — the
          // same "sticky" asterisk that correctly stays on a departed
          // tenant's frozen debt for months can also linger on the unit
          // itself even after someone new has moved in, and re-flipping
          // that confirmed-active person moved-out just because the same
          // asterisk showed up again (with no name change this time) would
          // permanently and incorrectly hide them from the rent sheet —
          // this only trusts the asterisk alone for someone not already
          // confirmed active that way.
          const nameActuallyChanged = entry.name && tenant.name && normalizeName(entry.name) !== normalizeName(tenant.name);
          const genuineMoveOutSignal = entry.movedOut && !tenant.movedOut && (!tenant.fromTurnover || nameActuallyChanged);
          if (Object.keys(diffFields).length > 0 || genuineMoveOutSignal) {
            const priorStatus = tenant.status || "Current";
            // Auto-apply by default now — only pause for approval if this
            // tenant is actively being worked: an open follow-up reminder, or
            // flagged to call back. Their prior Late/In Arrears status alone
            // no longer gates this, and neither does an active court case on
            // its own — a court-case tenant isn't shown on the spreadsheet at
            // all, so there's no risk of overwriting something visible
            // mid-conversation the way there is for a follow-up or call-back;
            // their balance still updates here, and the court tab's balance
            // pill (pulled live from this same tenant record) reflects it
            // automatically.
            const activelyWorking = tenantFollowUps(tenant).length > 0 || tenant.callBack === true;
            const needsApproval = type === "arrears" && activelyWorking;
            // If the balance changed, work out whether that's a payment (balance went
            // down) or a new charge (balance went up), so it can be logged as a dated
            // ledger entry on confirm instead of just silently becoming a new number.
            let balanceDelta = 0;
            if (type === "arrears" && "balance" in diffFields) {
              balanceDelta = parseBalance(tenant.balance) - parseBalance(diffFields.balance);
            }
            const existingFollowUps = tenantFollowUps(tenant);
            changes.push({
              apt: entry.apt, unitId: unit.id, tenantId: tenant.id, isNew: false,
              fields: diffFields, before: Object.fromEntries(Object.keys(diffFields).map(k => [k, tenant[k] || "—"])),
              name: tenant.name || entry.name, movedOut: genuineMoveOutSignal, nameActuallyChanged, needsReview: entry.needsReview,
              priorStatus, needsApproval, approved: !needsApproval, aging, balanceDelta,
              existingFollowUps, clearFollowUps: false, callBackFlag: tenant.callBack === true,
              inCourt: hasActiveCourtCase(tenant.id),
            });
          }
        } else {
          changes.push({ apt: entry.apt, unitId: unit.id, tenantId: null, isNew: true, fields: { ...fields, name: fields.name || entry.name || "" }, name: entry.name, movedOut: entry.movedOut, needsReview: entry.needsReview, needsApproval: false, approved: true, inCourt: false, aging });
        }
      } else {
        const existingToken = pendingNewUnits.get(aptKey);
        if (existingToken) {
          changes.push({ apt: entry.apt, unitId: existingToken, tenantId: null, isNew: true, newUnit: false, fields: { ...fields, name: fields.name || entry.name || "" }, name: entry.name, movedOut: entry.movedOut, needsReview: entry.needsReview, needsApproval: false, approved: true, inCourt: false, aging });
        } else {
          const token = "__pending_" + uid();
          pendingNewUnits.set(aptKey, token);
          changes.push({ apt: entry.apt, unitId: token, tenantId: null, isNew: true, newUnit: true, fields: { ...fields, name: fields.name || entry.name || "" }, name: entry.name, movedOut: entry.movedOut, needsReview: entry.needsReview, needsApproval: false, approved: true, inCourt: false, aging });
        }
      }
    });
  });

  let missing = [];
  if (type === "arrears") {
    const parsedApts = new Set(parsedEntries.map(e => normalizeName(e.apt)));
    const missingUnits = unitsForBuilding.filter(u => !parsedApts.has(normalizeName(u.unitNumber)));
    for (const u of missingUnits) {
      const tenant = data.tenants.find(t => t.unitId === u.id && !t.movedOut);
      missing.push({ apt: u.unitNumber, name: tenant?.name || "(no tenant on file)" });
      if (!tenant) continue; // nothing to mark paid if no tenant exists on this unit
      const currentBalance = parseBalance(tenant.balance);
      if (currentBalance <= 0 && (tenant.status || "Current") === "Current") continue; // already correct, no change needed
      const activelyWorking = tenantFollowUps(tenant).length > 0 || tenant.callBack === true;
      changes.push({
        apt: u.unitNumber, unitId: u.id, tenantId: tenant.id, isNew: false, derivedFromMissing: true,
        fields: { balance: "0.00", status: "Current" },
        before: { balance: tenant.balance || "0.00", status: tenant.status || "Current" },
        name: tenant.name, needsReview: false,
        priorStatus: tenant.status || "Current", needsApproval: activelyWorking, approved: !activelyWorking,
        aging: undefined, balanceDelta: currentBalance,
        existingFollowUps: tenantFollowUps(tenant), clearFollowUps: false, callBackFlag: tenant.callBack === true,
        inCourt: hasActiveCourtCase(tenant.id),
      });
    }
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
    // A parsed count far lower than what's already on file for this building is a
    // strong signal the parser only read part of the report (a page cut off, a
    // format quirk) rather than genuinely fewer tenants — worth surfacing loudly
    // here specifically because a low read now means the "missing from this
    // report" tenants auto-apply straight to $0/Current on confirm; silently
    // trusting a bad read would zero out real balances, not just miss new ones.
    const existingTenantCount = existing ? data.tenants.filter(t => t.buildingId === existing.id).length : 0;
    const suspiciouslyLowCount = type === "arrears" && existingTenantCount > 5 && entries.length < existingTenantCount * 0.5;
    setPreview({ ...diff, parsedCount: entries.length, header, matchedBuildingId: existing ? existing.id : null, suspiciouslyLowCount, existingTenantCount });
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
  const toggleClearFollowUps = (idx) => {
    setPreview(p => ({
      ...p,
      changes: p.changes.map((c, i) => i === idx ? { ...c, clearFollowUps: !c.clearFollowUps } : c),
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
      // Resolves a shared "__pending_" placeholder token (see buildImportDiff)
      // to the real unit id once that unit's own change has been processed —
      // so a second entry for the same not-yet-existing apt attaches to the
      // one real unit that gets created, not a second duplicate.
      const resolvedUnitTokens = new Map();
      applied.forEach(ch => {
        let unitId = ch.unitId;
        if (ch.newUnit) {
          const unit = { id: uid(), buildingId, unitNumber: ch.apt };
          next.units.push(unit);
          unitId = unit.id;
          resolvedUnitTokens.set(ch.unitId, unit.id);
        } else if (typeof unitId === "string" && unitId.startsWith("__pending_")) {
          unitId = resolvedUnitTokens.get(unitId) || unitId;
        }
        if (ch.isNew) {
          next.tenants.push({
            id: uid(), buildingId, unitId, name: ch.fields.name || ch.name || "",
            phone: ch.fields.phone || "", email: ch.fields.email || "",
            balance: ch.fields.balance || "", status: ch.fields.status || "Current",
            notes: [], messageLog: [], payments: [],
            ...(ch.movedOut ? { movedOut: true } : {}),
            ...(ch.aging ? { aging: ch.aging } : {}),
          });
        } else if (ch.movedOut && ch.nameActuallyChanged) {
          // A genuine detected turnover (the report's asterisk plus an
          // actual name change) — don't overwrite the departing tenant's
          // identity and history with the new person's. The old tenant
          // keeps their own name, balance, notes, and payment history
          // exactly as they were, just flagged moved-out so they drop out
          // of the main sheet; a brand new record gets created for the
          // incoming tenant with this import's balance, not anything
          // carried over from the old one. Same unit as before — the
          // apartment hasn't changed, only who lives there.
          next.tenants = next.tenants.map(t => t.id === ch.tenantId ? { ...t, movedOut: true } : t);
          next.tenants.push({
            id: uid(), buildingId, unitId: ch.unitId,
            name: ch.fields.name || ch.name || "", phone: "", email: "",
            balance: ch.fields.balance || "0.00", status: ch.fields.status || "Current",
            notes: [], messageLog: [], payments: [], followUps: [], fromTurnover: true,
            ...(ch.aging ? { aging: ch.aging } : {}),
          });
        } else if (ch.movedOut) {
          // Asterisk present, same name still listed — no one new has
          // moved in yet, this is just the departed tenant's unpaid debt
          // still being tracked. Flag moved-out (drops them off the main
          // sheet, unit reads as vacant) and apply whatever this report
          // shows for their balance, in case it genuinely changed — no
          // second record, since there's no new person to create one for.
          next.tenants = next.tenants.map(t => t.id === ch.tenantId ? { ...t, ...ch.fields, movedOut: true, ...(ch.aging ? { aging: ch.aging } : {}) } : t);
        } else {
          next.tenants = next.tenants.map(t => {
            if (t.id !== ch.tenantId) return t;
            let updated = { ...t, ...ch.fields, ...(ch.aging ? { aging: ch.aging } : {}) };
            // A balance that dropped since last import is a payment; one that
            // rose is a new charge. Either way it becomes a dated, signed
            // ledger entry instead of the balance just silently changing —
            // same array as manually-logged payments (positive = payment
            // received, negative = charge added), so "undo" works the same
            // way for both without needing separate logic.
            if (ch.balanceDelta) {
              updated.payments = [
                ...(Array.isArray(t.payments) ? t.payments : []),
                { id: uid(), date: todayISO(), amount: ch.balanceDelta.toFixed(2), note: "From RIS import", auto: true },
              ];
            }
            // Only clear follow-ups if the person running the import explicitly
            // opted into it for this tenant in the preview — never silent.
            if (ch.clearFollowUps) updated.followUps = [];
            return updated;
          });
        }
      });
      next.importHistory = [
        {
          id: uid(), date: todayISO(), buildingId, type,
          updated: applied.filter(c => !c.isNew).length,
          added: applied.filter(c => c.isNew).length,
          missing: preview.missing.length,
          skipped: preview.changes.length - applied.length,
          details: applied.map(c => ({ apt: c.apt, name: c.name, isNew: c.isNew, fields: c.fields, before: c.before })),
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
          {preview.suspiciouslyLowCount && (
            <div className="hint" style={{ color: "var(--danger)", fontWeight: 700, marginBottom: 8, padding: 8, border: "1px solid var(--danger)", borderRadius: 6 }}>
              ⚠ This file only read {preview.parsedCount} units, but this building already has {preview.existingTenantCount} on file — that's a big enough gap it's more likely something went wrong reading the report than that this many tenants genuinely paid off at once. Anyone missing from this file gets auto-marked paid ($0/Current) on confirm unless they have a follow-up or call-back flag — double-check this parsed correctly before confirming, since a partial read would zero out real balances along with the genuine ones.
            </div>
          )}
          <div className="import-preview-summary">
            {preview.changes.filter(c => c.isNew).length} new · {preview.changes.filter(c => !c.isNew && !c.needsApproval).length} to update
            {preview.changes.some(c => c.needsApproval) && ` · ${preview.changes.filter(c => c.needsApproval).length} already flagged behind — needs your approval`}
            {type === "arrears" && ` · ${preview.missing.length} not in this file`}
            {" "}({preview.parsedCount} rows read)
          </div>
          {preview.changes.length === 0 && (
            <div className="hint">
              Nothing to import — every unit in this file already matches what's on file exactly (same name, balance, and status). {preview.missing.length > 0 ? "The list below is unrelated: it's units on file that weren't found in this file at all." : ""}
            </div>
          )}
          {preview.changes.map((c, i) => (
            <div className={`import-row ${c.needsApproval ? "import-row-approval" : ""}`} key={i}>
              <span className="pill pill-muted">{c.apt}</span>
              <span className="import-row-name">{c.name}</span>
              {c.isNew && <span className="pill pill-warn">New</span>}
              {c.derivedFromMissing && <span className="pill pill-ok">Not in this report — marking paid up</span>}
              {c.movedOut && c.nameActuallyChanged && <span className="pill pill-danger" title="Report marks this unit moved-out and the name changed — the old tenant's history stays intact and gets flagged moved-out; a new record is created for the incoming name">Moved out — new tenant record created</span>}
              {c.movedOut && !c.nameActuallyChanged && <span className="pill pill-warn" title="Report marks this unit moved-out, same name still listed — no one new has moved in yet, this tenant just gets flagged moved-out">Moved out — no replacement yet</span>}
              {c.needsReview && <span className="pill pill-warn">⚠ Review — shared unit, verify names</span>}
              {c.inCourt && <span className="pill pill-danger">In Court</span>}
              <div className="import-row-fields">
                {Object.entries(c.fields).map(([k, v]) => (
                  <span key={k} className="import-field">{k}: {c.before?.[k] ? `${c.before[k]} → ` : ""}{String(v)}</span>
                ))}
              </div>
              {!!c.balanceDelta && (
                <span className={`pill ${c.balanceDelta > 0 ? "pill-ok" : "pill-warn"}`}>
                  {c.balanceDelta > 0 ? `Payment detected: $${c.balanceDelta.toFixed(2)}` : `Charge added: $${Math.abs(c.balanceDelta).toFixed(2)}`}
                </span>
              )}
              {c.needsApproval && (
                <label className="import-approve">
                  <input type="checkbox" checked={c.approved} onChange={() => toggleApprove(i)} />
                  {[
                    c.existingFollowUps && c.existingFollowUps.length > 0
                      ? `You have a follow-up scheduled ${c.existingFollowUps.map(f => fmtDate(f.date) + (f.note ? ` (${f.note})` : "")).join(", ")}`
                      : null,
                    c.callBackFlag ? "flagged to call back" : null,
                  ].filter(Boolean).join(" and ")} — update to this anyway?
                </label>
              )}
              {c.balanceDelta > 0 && c.existingFollowUps && c.existingFollowUps.length > 0 && (c.fields.status === "Current" || parseBalance(c.fields.balance) <= 0) && (
                <label className="import-approve">
                  <input type="checkbox" checked={!!c.clearFollowUps} onChange={() => toggleClearFollowUps(i)} />
                  Paid off and back to Current — clear their active follow-up ({c.existingFollowUps.map(f => fmtDate(f.date)).join(", ")})?
                </label>
              )}
            </div>
          ))}
          {preview.missing.filter(m => !preview.changes.some(c => c.derivedFromMissing && c.apt === m.apt)).length > 0 && (
            <>
              <div className="row" style={{ marginTop: 10 }}><strong>Not found in this file (already $0 / Current — no change needed)</strong></div>
              {preview.missing.filter(m => !preview.changes.some(c => c.derivedFromMissing && c.apt === m.apt)).map((m, i) => (
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
                  {d.apt} — {d.name} {d.isNew && "(new)"} {Object.entries(d.fields || {}).map(([k, v]) => `${k}: ${d.before?.[k] ? `${d.before[k]} → ` : ""}${v}`).join(", ")}
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
  const [selected, setSelected] = useState(new Set());
  const [copiedId, setCopiedId] = useState(null);
  const [expandedRow, setExpandedRow] = useState(null);
  const [noteFor, setNoteFor] = useState(null);
  const [noteText, setNoteText] = useState("");

  const submit = () => {
    if (!form.description) return;
    const exists = data.workOrders.some(w => w.id === form.id);
    if (exists) update("workOrders", form.id, form);
    else {
      add("workOrders", form); // form.id was pre-generated when the form opened
      setExpandedRow(form.id); // so photos/notes can be added immediately, no re-opening needed
    }
    setForm(null);
  };

  const list = data.workOrders
    .filter(w => filter === "All" || w.status === filter)
    .slice()
    .sort((a, b) => (b.dateOpened || "").localeCompare(a.dateOpened || ""));

  const toggleSelected = (id) => {
    setSelected(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  const copyOne = (w) => {
    navigator.clipboard.writeText(formatItemsForCopy([w], data));
    setCopiedId(w.id);
    setTimeout(() => setCopiedId(null), 1500);
  };
  const copySelected = () => {
    const items = list.filter(w => selected.has(w.id));
    navigator.clipboard.writeText(formatItemsForCopy(items, data));
    setCopiedId("__batch__");
    setTimeout(() => setCopiedId(null), 1500);
  };
  const addNote = (w) => {
    if (!noteText.trim()) return;
    // timestamp (a plain millisecond number) is what sorting relies on, not
    // date alone — several notes logged the same day would otherwise have
    // no reliable order between them. date stays for display (fmtDate
    // reads that), timestamp never needs to be shown.
    update("workOrders", w.id, { notes: [...(w.notes || []), { date: todayISO(), timestamp: Date.now(), text: noteText }] });
    setNoteText(""); setNoteFor(null);
  };

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Work Orders</h1>
        <div className="page-actions">
          {selected.size > 0 && (
            <button className="btn-ghost" onClick={copySelected}>
              <ScrollText size={14} /> {copiedId === "__batch__" ? "Copied!" : `Copy selected (${selected.size})`}
            </button>
          )}
          <PrintButton label="Work Orders" />
          <button className="btn-primary" onClick={() => setForm({ id: uid(), buildingId: data.buildings[0]?.id || "", unitId: "", vendorId: "", description: "", status: "Open", priority: "Routine", dateOpened: todayISO() })}>
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
      {list.map(w => {
        const isOpen = expandedRow === w.id;
        const sortedNotes = [...(w.notes || [])].sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
        return (
        <div className="list-card" key={w.id}>
          <div className="list-card-head" onClick={() => setExpandedRow(isOpen ? null : w.id)} style={{ cursor: "pointer" }}>
            {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
            <input type="checkbox" checked={selected.has(w.id)} onChange={(e) => { e.stopPropagation(); toggleSelected(w.id); }} onClick={(e) => e.stopPropagation()} title="Select for batch copy" />
            {w.unitId && <span className="pill pill-accent">Apt {data.units.find(u => u.id === w.unitId)?.unitNumber || "—"}</span>}
            <div className="list-card-title">{w.description}</div>
            {w.priority !== "Routine" && <span className={`pill ${w.priority === "Emergency" ? "pill-danger" : "pill-warn"}`}>{w.priority}</span>}
            <span className={`pill ${w.status === "Done" ? "pill-ok" : "pill-muted"}`}>{w.status}</span>
            <span className="pill pill-muted">{buildingName(w.buildingId)}</span>
            {w.vendorId && <span className="pill pill-muted">{vendorName(w.vendorId)}</span>}
            <div className="spacer" />
            <IconBtn title={copiedId === w.id ? "Copied!" : "Copy for texting/emailing"} onClick={(e) => { e.stopPropagation(); copyOne(w); }}><ScrollText size={14} /></IconBtn>
            <IconBtn title="Edit" onClick={(e) => { e.stopPropagation(); setForm(w); }}><Pencil size={14} /></IconBtn>
            <IconBtn title="Delete" danger onClick={(e) => { e.stopPropagation(); remove("workOrders", w.id); }}><Trash2 size={14} /></IconBtn>
          </div>
          {isOpen && (
            <div className="list-card-body">
              <PhotoUploader
                photos={w.photos}
                pathPrefix={`workOrders/${w.id}/photos`}
                onAdd={(newPhotos) => update("workOrders", w.id, { photos: [...(w.photos || []), ...newPhotos] })}
                onRemove={(p) => {
                  update("workOrders", w.id, { photos: (w.photos || []).filter(x => x.id !== p.id) });
                  if (p.storagePath) deleteObject(storageRef(storage, p.storagePath)).catch(() => {});
                }}
              />
              <div className="row" style={{ marginTop: 8 }}>
                <strong>Notes</strong>
                <button className="btn-ghost" style={{ marginLeft: 8 }} onClick={() => setNoteFor(noteFor === w.id ? null : w.id)}><Plus size={14} /> Add note</button>
              </div>
              {noteFor === w.id && (
                <div className="inline-form">
                  <input placeholder="Update…" value={noteText} onChange={e => setNoteText(e.target.value)} onKeyDown={e => e.key === "Enter" && addNote(w)} />
                  <button className="btn-primary" onClick={() => addNote(w)}>Save</button>
                </div>
              )}
              {sortedNotes.map((n, i) => (
                <div key={i} className="row row-muted">{fmtDate(n.date)} — {n.text}</div>
              ))}
            </div>
          )}
        </div>
      );})}
    </div>
  );
}

/* ============================== violations ============================== */

function ViolationsTab({ data, add, update, remove, buildingName, vendorName, setData }) {
  const [agency, setAgency] = useState("All");
  const [form, setForm] = useState(null);
  const [view, setView] = useState("active");
  const [dueFilter, setDueFilter] = useState("all");
  const [noteFor, setNoteFor] = useState(null);
  const [noteText, setNoteText] = useState("");
  const [expandedRow, setExpandedRow] = useState(null);
  const [selected, setSelected] = useState(new Set());
  const [copiedId, setCopiedId] = useState(null);
  const [showHpdImport, setShowHpdImport] = useState(false);
  const [buildingFilter, setBuildingFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [confirmingDeleteAllHpd, setConfirmingDeleteAllHpd] = useState(false);
  const [expandedGroups, setExpandedGroups] = useState(new Set());
  const fileRef = useRef(null);

  // Scoped to HPD only — a generic "delete all violations" would also wipe
  // DSNY and other agencies' data, which has nothing to do with re-testing
  // an HPD import.
  const hpdViolationCount = data.violations.filter(v => v.agency === "HPD").length;
  const deleteAllHpdViolations = () => {
    setData(d => ({ ...d, violations: d.violations.filter(v => v.agency !== "HPD") }));
    setConfirmingDeleteAllHpd(false);
  };

  const toggleSelected = (id) => {
    setSelected(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  const copyOne = (v) => {
    navigator.clipboard.writeText(formatItemsForCopy([v], data));
    setCopiedId(v.id);
    setTimeout(() => setCopiedId(null), 1500);
  };
  const copySelected = () => {
    const items = data.violations.filter(v => selected.has(v.id));
    navigator.clipboard.writeText(formatItemsForCopy(items, data));
    setCopiedId("__batch__");
    setTimeout(() => setCopiedId(null), 1500);
  };

  const statusesFor = (a) => a === "HPD" ? HPD_STATUSES : a === "DSNY" ? DSNY_STATUSES : OTHER_STATUSES;
  const isClosed = isViolationClosed;
  const otherAgencyOptions = [...OTHER_AGENCY_PRESETS, ...(data.customOtherAgencies || [])];
  const addCustomOtherAgency = (trimmed) => {
    if (!otherAgencyOptions.includes(trimmed)) {
      setData(d => ({ ...d, customOtherAgencies: [...(d.customOtherAgencies || []), trimmed] }));
    }
  };

  // A tab can be "HPD", "DSNY", the generic "Other" catch-all, or a dynamic
  // agency like "DOB"/"FDNY" — this resolves any of those down to the real
  // stored shape: agency is always HPD/DSNY/Other, otherAgency carries the
  // specific one when it applies. Used for both new violations (so the form
  // opens pre-set to whatever tab you were on) and CSV import.
  const agencyFieldsFor = (a) => {
    if (a === "HPD" || a === "DSNY") return { agency: a, otherAgency: "" };
    if (a === "Other" || a === "All" || !a) return { agency: "Other", otherAgency: "" };
    return { agency: "Other", otherAgency: a };
  };

  const blankForm = (a) => {
    const { agency: ag, otherAgency: oa } = agencyFieldsFor(a);
    return {
      id: uid(), agency: ag, buildingId: data.buildings[0]?.id || "", unitId: "", violationNumber: "",
      class: "", description: "", dateIssued: todayISO(), cureDeadline: "",
      fineAmount: "", company: "", otherAgency: oa || otherAgencyOptions[0] || "",
      status: statusesFor(ag)[0], vendorId: "",
    };
  };

  const submit = () => {
    if (!form.violationNumber) return;
    const exists = data.violations.some(x => x.id === form.id);
    if (exists) update("violations", form.id, form);
    else {
      add("violations", { ...form, photos: [], notes: [] }); // form.id was pre-generated when the form opened
      setExpandedRow(form.id); // so photos/notes can be added immediately, no re-opening needed
    }
    setForm(null);
  };

  const addNote = (v) => {
    if (!noteText.trim()) return;
    // timestamp (a plain millisecond number) is what sorting relies on, not
    // date alone — several notes logged the same day would otherwise have
    // no reliable order between them. date stays for display (fmtDate
    // reads that), timestamp never needs to be shown.
    update("violations", v.id, { notes: [...(v.notes || []), { date: todayISO(), timestamp: Date.now(), text: noteText }] });
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
            const address = csvField(row, "address", "building");
            const building = d.buildings.find(b => (b.address || "").toLowerCase() === address.toLowerCase());
            const rawAgency = csvField(row, "agency");
            const resolved = agencyFieldsFor(rawAgency || agency);
            const finalOtherAgency = csvField(row, "otherAgency", "agencyDetail") || resolved.otherAgency;
            next.violations.push({
              id: uid(), agency: resolved.agency, otherAgency: finalOtherAgency,
              violationNumber: csvField(row, "violationNumber", "number", "id"),
              buildingId: building ? building.id : "",
              class: csvField(row, "class"),
              description: csvField(row, "description"),
              dateIssued: csvField(row, "dateIssued"),
              cureDeadline: csvField(row, "cureDeadline", "deadline"),
              fineAmount: csvField(row, "fineAmount", "fine"),
              company: csvField(row, "company"),
              status: csvField(row, "status") || statusesFor(resolved.agency)[0],
              vendorId: "", photos: [], notes: [],
            });
          });
          return next;
        });
      }
    });
    e.target.value = "";
  };

  // Any distinct "Other Agency" value actually in use (DOB, FDNY, DEP, etc.)
  // gets its own top-level tab automatically, instead of being buried inside
  // a generic "Other" bucket — "Other" stays as the catch-all for anything
  // without a specific agency set.
  const dynamicOtherAgencies = [...new Set(data.violations.filter(v => v.agency === "Other" && v.otherAgency).map(v => v.otherAgency))].sort();
  const agencyTabs = ["All", "HPD", "DSNY", ...dynamicOtherAgencies];

  const matchesAgency = (v, a) => {
    if (a === "HPD") return v.agency === "HPD";
    if (a === "DSNY") return v.agency === "DSNY";
    if (a === "Other") {
      // A true catch-all: anything not HPD, not DSNY, and not matching a
      // known dynamic agency lands here — no matter what garbage value its
      // own agency/otherAgency fields actually hold. A violation with a
      // corrupted agency (e.g. literally "All" from an old bug) used to
      // match nothing at all here, making it invisible on this whole page
      // while still being counted everywhere else that doesn't require a
      // bucket match, like the Dashboard.
      if (v.agency === "HPD" || v.agency === "DSNY") return false;
      if (v.agency === "Other" && dynamicOtherAgencies.includes(v.otherAgency)) return false;
      return true;
    }
    return v.agency === "Other" && v.otherAgency === a;
  };
  const filterAndSort = (a) => {
    let l = data.violations.filter(v => matchesAgency(v, a) && (view === "active" ? !isClosed(v) : isClosed(v)));
    if (buildingFilter !== "All") l = l.filter(v => v.buildingId === buildingFilter);
    if (search.trim()) {
      const q = search.trim().toLowerCase();
      l = l.filter(v => {
        const unitNum = v.unitId ? (data.units.find(u => u.id === v.unitId)?.unitNumber || "") : "";
        return [v.violationNumber, unitNum, v.class, v.description, v.status, v.company, buildingName(v.buildingId)]
          .some(f => (f || "").toString().toLowerCase().includes(q));
      });
    }
    if (view === "active") {
      if (dueFilter !== "all") {
        const maxDays = dueFilter === "24h" ? 1 : dueFilter === "1w" ? 7 : 10;
        l = l.filter(v => { const d = daysUntil(v.cureDeadline); return d !== null && d <= maxDays; });
      }
      l = [...l].sort((a2, b2) => {
        if (!!a2.isLead !== !!b2.isLead) return a2.isLead ? 1 : -1;
        const da = daysUntil(a2.cureDeadline); const db = daysUntil(b2.cureDeadline);
        if (da === null && db === null) return 0;
        if (da === null) return 1;
        if (db === null) return -1;
        return da - db;
      });
    } else {
      // Closed/paid — no longer has a meaningful future date, so newest
      // (most recently issued) first instead, like everything else that's
      // a log of what happened rather than a schedule of what's coming.
      l = [...l].sort((a2, b2) => (b2.dateIssued || "").localeCompare(a2.dateIssued || ""));
    }
    return l;
  };

  const renderRow = (v, rowAgency) => {
    const flag = view === "active" ? flagFor(v.cureDeadline) : null;
    const isOpen = expandedRow === v.id;
    return (
      <div className={`list-card list-card-compact ${flag === "overdue" ? "list-card-danger" : flag === "soon" ? "list-card-warn" : ""}`} key={v.id}>
        <div className="list-card-head" onClick={() => setExpandedRow(isOpen ? null : v.id)} style={{ cursor: "pointer", alignItems: "flex-start" }}>
          {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
          <input type="checkbox" checked={selected.has(v.id)} onChange={(e) => { e.stopPropagation(); toggleSelected(v.id); }} onClick={(e) => e.stopPropagation()} title="Select for batch copy" />
          {v.unitId && <span className="pill pill-accent">Apt {data.units.find(u => u.id === v.unitId)?.unitNumber || "—"}</span>}
          <div className="violation-title-group">
            <div className="list-card-title">#{v.violationNumber}</div>
            {v.description && <div className="violation-desc-preview">{summarizeViolationDescription(v.description)}</div>}
          </div>
          <span className={`pill ${isClosed(v) ? "pill-ok" : "pill-muted"}`}>{v.status}</span>
          <span className="pill pill-muted">{buildingName(v.buildingId)}</span>
          {rowAgency === "HPD" && v.vendorId && <span className="pill pill-muted">{vendorName(v.vendorId)}</span>}
          {rowAgency === "DSNY" && v.fineAmount && <span className="pill pill-muted">{v.fineAmount}</span>}
          {rowAgency === "Other" && v.otherAgency && <span className="pill pill-muted">{v.otherAgency}</span>}
          {v.hasHearing && <span className="pill pill-warn"><Gavel size={11} /> Hearing{v.hearingDate ? ` ${fmtDate(v.hearingDate)}` : ""}</span>}
          {v.isLead && <span className="pill pill-danger">Lead</span>}
          {v.isMoldOver10 && <span className="pill pill-warn">Mold ≥10 sq ft</span>}
          {rowAgency !== "HPD" && rowAgency !== "DSNY" && v.company && <span className="pill pill-muted">{v.company}</span>}
          {rowAgency !== "DSNY" && view === "active" && <Flag date={v.cureDeadline} />}
          <div className="spacer" />
          {rowAgency === "HPD" && view === "active" && (
            <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); update("violations", v.id, { status: "Certified" }); }}>
              Mark Certified
            </button>
          )}
          {rowAgency === "DSNY" && view === "active" && (
            <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); update("violations", v.id, { status: "Paid" }); }}>
              Mark paid
            </button>
          )}
          <IconBtn title={copiedId === v.id ? "Copied!" : "Copy for texting/emailing"} onClick={(e) => { e.stopPropagation(); copyOne(v); }}><ScrollText size={14} /></IconBtn>
          <IconBtn title="Edit" onClick={(e) => { e.stopPropagation(); setForm(v); }}><Pencil size={14} /></IconBtn>
          <IconBtn title="Delete" danger onClick={(e) => { e.stopPropagation(); remove("violations", v.id); }}><Trash2 size={14} /></IconBtn>
        </div>
        {isOpen && (
          <div className="list-card-body">
            {rowAgency === "DSNY" && view === "active" && (
              <div className="row" style={{ display: "flex", gap: 8, marginBottom: 4 }}>
                <button className={`chip ${v.status === "Disputing online" ? "chip-active" : ""}`} onClick={() => update("violations", v.id, { status: "Disputing online" })}>
                  Fighting it
                </button>
                <button className={`chip ${v.status === "Paid" ? "chip-active" : ""}`} onClick={() => update("violations", v.id, { status: "Paid" })}>
                  Paying it
                </button>
              </div>
            )}
            {v.description && <div className="row">{v.description}</div>}
            <PhotoUploader
              photos={v.photos}
              pathPrefix={`violations/${v.id}/photos`}
              onAdd={(newPhotos) => update("violations", v.id, { photos: [...(v.photos || []), ...newPhotos] })}
              onRemove={(p) => {
                update("violations", v.id, { photos: (v.photos || []).filter(x => x.id !== p.id) });
                if (p.storagePath) deleteObject(storageRef(storage, p.storagePath)).catch(() => {});
              }}
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
            {[...(v.notes || [])].sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0)).map((n, i) => (
              <div key={i} className="row row-muted">{fmtDate(n.date)} — {n.text}</div>
            ))}
          </div>
        )}
      </div>
    );
  };

  let list = agency === "All" ? [] : filterAndSort(agency);

  return (
    <div className="violations-page">
      <div className="page-head">
        <h1 className="page-title">Violations</h1>
        <div className="page-actions">
          {selected.size > 0 && (
            <button className="btn-ghost" onClick={copySelected}>
              <ScrollText size={14} /> {copiedId === "__batch__" ? "Copied!" : `Copy selected (${selected.size})`}
            </button>
          )}
          <PrintButton label="Violations" />
          <button className="btn-ghost" onClick={() => setShowHpdImport(s => !s)}><Upload size={14} /> Import HPD violations</button>
          {hpdViolationCount > 0 && (
            confirmingDeleteAllHpd ? (
              <>
                <span className="row-muted" style={{ fontSize: 12 }}>Delete all {hpdViolationCount} HPD violation{hpdViolationCount === 1 ? "" : "s"}? This can't be undone.</span>
                <button className="btn-ghost" style={{ color: "var(--danger)" }} onClick={deleteAllHpdViolations}>Yes, delete all</button>
                <button className="btn-ghost" onClick={() => setConfirmingDeleteAllHpd(false)}>Cancel</button>
              </>
            ) : (
              <button className="btn-ghost" style={{ color: "var(--danger)" }} onClick={() => setConfirmingDeleteAllHpd(true)}><Trash2 size={14} /> Delete all HPD</button>
            )
          )}
          <button className="btn-primary" onClick={() => setForm(blankForm(agency === "All" ? "HPD" : agency))}>
            <Plus size={14} /> Add violation
          </button>
        </div>
      </div>

      {showHpdImport && <HpdViolationsImportSection data={data} add={add} update={update} onImported={() => setShowHpdImport(false)} />}

      <div className="row" style={{ marginBottom: 10 }}>
        <input
          type="text" placeholder="Search violation #, apt, keyword…" value={search}
          onChange={e => setSearch(e.target.value)}
          style={{ width: "100%", maxWidth: 360 }}
        />
      </div>

      {data.buildings.length > 1 && (
        <div className="filter-row">
          <button className={`chip ${buildingFilter === "All" ? "chip-active" : ""}`} onClick={() => setBuildingFilter("All")}>All buildings</button>
          {data.buildings.filter(b => data.violations.some(v => v.buildingId === b.id)).map(b => (
            <button key={b.id} className={`chip ${buildingFilter === b.id ? "chip-active" : ""}`} onClick={() => setBuildingFilter(b.id)}>{shortAddress(b.address)}</button>
          ))}
        </div>
      )}

      <div className="print-only">
        <div className="print-header">
          <div className="print-mark">O</div>
          <div className="print-header-text">
            <h1 className="print-title">Violations Report — {view === "closed" ? (agency === "DSNY" ? "Paid" : "Closed") : "Active / Due"}</h1>
            <div className="print-subtitle">As of {fmtDate(todayISO())}</div>
          </div>
        </div>
        <div className="print-stats-row">
          <div className="print-stat"><div className="print-stat-num">{agencyTabs.filter(a => a !== "All").reduce((sum, a) => sum + filterAndSort(a).length, 0)}</div><div>Total {view === "closed" ? "resolved" : "open"}</div></div>
          {agencyTabs.filter(a => a !== "All").map(a => (
            <div className="print-stat" key={a}><div className="print-stat-num">{filterAndSort(a).length}</div><div>{a}</div></div>
          ))}
        </div>
        {agencyTabs.filter(a => a !== "All").map(a => {
          const items = filterAndSort(a);
          if (items.length === 0) return null;
          return (
            <div className="print-section" key={a}>
              <div className="print-section-head">
                <span>{a}</span>
                <span>{items.length} violation{items.length === 1 ? "" : "s"}</span>
              </div>
              <table className="print-table">
                <thead>
                  <tr><th>Building</th><th>Unit</th><th>Violation #</th><th>Class</th><th>Issued</th><th>{view === "closed" ? "Resolved" : "Cure deadline"}</th><th>Status</th></tr>
                </thead>
                <tbody>
                  {items.map(v => (
                    <tr key={v.id}>
                      <td>{buildingName(v.buildingId)}</td>
                      <td>{data.units.find(u => u.id === v.unitId)?.unitNumber || "—"}</td>
                      <td>#{v.violationNumber || "—"}</td>
                      <td>{v.class || "—"}</td>
                      <td>{v.dateIssued ? fmtDate(v.dateIssued) : "—"}</td>
                      <td>{v.cureDeadline ? fmtDate(v.cureDeadline) : "not set"}</td>
                      <td>{v.status}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          );
        })}
        <div className="print-footer">
          <span>Property Ops — Violations Report</span>
          <span>Generated {fmtDate(todayISO())}</span>
        </div>
      </div>
      <p className="hint">CSV columns recognized: agency, otherAgency, violationNumber, address, class, description, dateIssued, cureDeadline, fineAmount, company, status.</p>

      <div className="filter-row">
        {agencyTabs.map(a => (
          <button key={a} className={`chip ${agency === a ? "chip-active" : ""}`} onClick={() => { setAgency(a); setForm(null); }}>{a}</button>
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
          <Field label="Agency">
            <TypeSelectWithAdd
              value={form.agency === "Other" ? form.otherAgency : form.agency}
              options={["HPD", "DSNY", ...otherAgencyOptions]}
              onChange={v => {
                if (v === "HPD" || v === "DSNY") {
                  setForm({ ...form, agency: v, otherAgency: "", status: statusesFor(v)[0] });
                } else {
                  setForm({ ...form, agency: "Other", otherAgency: v, status: statusesFor("Other")[0] });
                }
              }}
              onAddType={addCustomOtherAgency}
            />
          </Field>

          <Field label="Hearing scheduled?">
            <label className="hearing-checkbox-row">
              <input type="checkbox" checked={!!form.hasHearing} onChange={e => setForm({ ...form, hasHearing: e.target.checked })} />
              <span>This violation has a hearing</span>
            </label>
          </Field>
          {form.hasHearing && (
            <>
              <Field label="Hearing date"><input type="date" value={form.hearingDate || ""} onChange={e => setForm({ ...form, hearingDate: e.target.value })} /></Field>
              <Field label="Company handling hearing"><input value={form.hearingCompany || ""} onChange={e => setForm({ ...form, hearingCompany: e.target.value })} /></Field>
            </>
          )}

          {form.agency === "HPD" && (
            <>
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

          {form.agency === "DSNY" && (
            <>
              <Field label="Fine amount"><input value={form.fineAmount} onChange={e => setForm({ ...form, fineAmount: e.target.value })} placeholder="$" /></Field>
              <Field label="Date issued"><input type="date" value={form.dateIssued} onChange={e => setForm({ ...form, dateIssued: e.target.value })} /></Field>
            </>
          )}

          {form.agency === "Other" && (
            <>
              <Field label="Violation company handling"><input value={form.company} onChange={e => setForm({ ...form, company: e.target.value })} /></Field>
              <Field label="Description"><textarea value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} /></Field>
              <Field label="Date issued"><input type="date" value={form.dateIssued} onChange={e => setForm({ ...form, dateIssued: e.target.value })} /></Field>
              <Field label="Correction deadline"><input type="date" value={form.cureDeadline} onChange={e => setForm({ ...form, cureDeadline: e.target.value })} /></Field>
            </>
          )}

          <Field label="Status">
            <select value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}>
              {statusesFor(form.agency).map(s => <option key={s}>{s}</option>)}
            </select>
          </Field>
          <div className="form-actions">
            <button className="btn-primary" onClick={submit}>Save</button>
            <button className="btn-ghost" onClick={() => setForm(null)}>Cancel</button>
          </div>
        </div>
      )}

      {agency === "All" ? (
        (() => {
          const groups = ["HPD", "DSNY", ...dynamicOtherAgencies, "Other"].map(a => ({ agency: a, items: filterAndSort(a) })).filter(g => g.items.length > 0);
          if (groups.length === 0) return <EmptyState text="No violations here." />;
          // Collapsed by default on page load (nothing in expandedGroups
          // yet) — clicking a group heading, or "Expand all", opts it in.
          const allExpanded = groups.every(g => expandedGroups.has(g.agency));
          const toggleGroup = (a) => setExpandedGroups(prev => {
            const next = new Set(prev);
            if (next.has(a)) next.delete(a); else next.add(a);
            return next;
          });
          return (
            <>
              <div className="row" style={{ marginBottom: 8 }}>
                <button className="btn-ghost" onClick={() => setExpandedGroups(allExpanded ? new Set() : new Set(groups.map(g => g.agency)))}>
                  {allExpanded ? "Collapse all" : "Expand all"}
                </button>
              </div>
              {groups.map((g, gi) => {
                const isCollapsed = !expandedGroups.has(g.agency);
                return (
                  <div key={g.agency} style={{ marginTop: gi === 0 ? 0 : 20 }}>
                    <div className="violations-group-heading" style={{ cursor: "pointer" }} onClick={() => toggleGroup(g.agency)}>
                      {isCollapsed ? <ChevronRight size={14} /> : <ChevronDown size={14} />} {g.agency} <span className="dash-panel-sub">({g.items.length})</span>
                    </div>
                    {!isCollapsed && g.items.map(v => renderRow(v, g.agency))}
                  </div>
                );
              })}
            </>
          );
        })()
      ) : (
        <>
          {list.length === 0 && <EmptyState text={`No ${agency} violations here.`} />}
          {list.map(v => renderRow(v, agency))}
        </>
      )}
    </div>
  );
}

/* ============================== vendors ============================== */

function VendorsTab({ data, add, update, remove, buildingName }) {
  const [form, setForm] = useState(null);
  const [selected, setSelected] = useState(null);
  const [pendingDelete, setPendingDelete] = useState(null);

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
            {pendingDelete === v.id ? (
              <>
                <span className="row-muted" style={{ fontSize: 12, color: (openWO(v.id).length + openViol(v.id).length) > 0 ? "var(--danger)" : undefined }}>
                  Delete {v.name}?{(openWO(v.id).length + openViol(v.id).length) > 0 ? ` Still has ${openWO(v.id).length + openViol(v.id).length} open item${(openWO(v.id).length + openViol(v.id).length) === 1 ? "" : "s"} assigned.` : ""}
                </span>
                <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); remove("vendors", v.id); setPendingDelete(null); }} style={{ color: "var(--danger)" }}>Yes, delete</button>
                <button className="btn-ghost" onClick={(e) => { e.stopPropagation(); setPendingDelete(null); }}>Cancel</button>
              </>
            ) : (
              <>
                <IconBtn title="Edit" onClick={(e) => { e.stopPropagation(); setForm(v); }}><Pencil size={14} /></IconBtn>
                <IconBtn title="Delete" danger onClick={(e) => { e.stopPropagation(); setPendingDelete(v.id); }}><Trash2 size={14} /></IconBtn>
              </>
            )}
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

// Matches a law firm's "Complete Client Status" export against the app's
// own tenants — court report building codes are the firm's own internal
// numbering and don't reliably match this app's building records (a
// commercial unit can even show a completely different "Building:" code
// than its real address, e.g. an AKA street), so building matching goes by
// address text instead, and unit matching falls back to a name-overlap
// check when the report's apt value doesn't cleanly correspond to a real
// unit number (a storefront listed as "STORE FRONT" rather than its
// actual unit number, for instance).
function normalizeAddrForMatch(s) { return (s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); }
function normalizeAptForMatch(s) { return (s || "").toUpperCase().replace(/[^A-Z0-9]/g, ""); }
function normalizeNameForMatch(s) { return (s || "").toLowerCase().replace(/[^a-z\s]/g, "").split(/\s+/).filter(w => w.length > 2); }

const GENERIC_ADDRESS_FRAGMENTS = new Set(["brooklyn", "queens", "bronx", "manhattan", "statenisland", "newyork", "ny", "nyc"]);
function isGenericAddressFragment(normFrag) {
  return /^\d{5}(-?\d{4})?$/.test(normFrag) || GENERIC_ADDRESS_FRAGMENTS.has(normFrag);
}
// Street name alone, stripping the leading house number — needed because a
// report can show a different number within the same building's known
// range (e.g. "321 Ovington Avenue" vs. the building's stored "333
// Ovington Avenue"), and only the street name is reliably shared between
// them.
function streetNameOnly(addr) {
  let raw = (addr || "").split(",")[0].replace(/^[\d-]+\s*/, "");
  // Word-level abbreviations within the name itself (not a trailing street
  // type) — expanded here, before spaces are stripped, since they need
  // real word boundaries to match safely ("Ft" as a whole word, not as a
  // substring buried inside "Fort" once everything's run together).
  raw = raw.replace(/\bft\b/gi, "fort").replace(/\bmt\b/gi, "mount");
  let s = normalizeAddrForMatch(raw);
  // Common street-suffix abbreviations vary between sources ("Avenue" vs
  // "Ave", "Street" vs "St") — normalized to the same short form so those
  // don't cause an otherwise-identical street name to miss.
  s = s.replace(/avenue$/, "ave").replace(/street$/, "st").replace(/boulevard$/, "blvd")
       .replace(/place$/, "pl").replace(/road$/, "rd").replace(/parkway$/, "pkwy");
  return s;
}
function findCourtBuildingMatch(courtAddress, buildings) {
  const rawFragments = (courtAddress || "").split(/,|\bAKA\b/i).map(f => f.trim()).filter(Boolean);
  // A bare city name or zip code matches nearly every building in the
  // portfolio equally — keeping those in the fragment list let a genuinely
  // different address (like a house number outside any known range) fall
  // through to a false match on "Brooklyn" alone, landing on whichever
  // building happened to be first, rather than correctly matching nothing.
  const realFragments = rawFragments.filter(f => !isGenericAddressFragment(normalizeAddrForMatch(f)));
  if (realFragments.length === 0) return null;
  // Pass 1: the full number+street fragment, most specific — either
  // direction of substring containment, since the stored address and the
  // report's address aren't always the same length or order.
  for (const frag of realFragments) {
    const nf = normalizeAddrForMatch(frag);
    if (!nf) continue;
    const match = buildings.find(b => {
      const nb = normalizeAddrForMatch(b.address);
      return nb.includes(nf) || nf.includes(nb);
    });
    if (match) return match;
  }
  // Pass 2: street name only, ignoring the leading number — containment
  // rather than strict equality, since one side may include more of the
  // name than the other even after suffix normalization.
  for (const frag of realFragments) {
    const streetOnly = streetNameOnly(frag);
    if (!streetOnly) continue;
    const match = buildings.find(b => {
      const bStreet = streetNameOnly(b.address);
      return bStreet && (bStreet.includes(streetOnly) || streetOnly.includes(bStreet));
    });
    if (match) return match;
  }
  return null;
}
function findCourtTenantMatch(courtCase, building, tenants, units) {
  if (!building) return { tenant: null, reason: "no building match" };
  const buildingTenants = tenants.filter(t => t.buildingId === building.id);
  const naApt = normalizeAptForMatch(courtCase.apt);
  const unitMatches = buildingTenants.filter(t => {
    const unit = units.find(u => u.id === t.unitId);
    return unit && normalizeAptForMatch(unit.unitNumber) === naApt;
  });
  // A unit can have both an active tenant and an already-moved-out one on
  // file — prefer whoever's actually there now, since a court case is far
  // more often against the current occupant than someone who's already
  // left. Only fall back to a moved-out tenant when there's no active one
  // on that unit, since an ongoing case can genuinely still be against
  // someone who's since moved out.
  const byApt = unitMatches.find(t => !t.movedOut) || unitMatches[0] || null;
  if (byApt) return { tenant: byApt, reason: "matched by unit number" };
  const courtWords = normalizeNameForMatch(courtCase.name);
  if (courtWords.length === 0) return { tenant: null, reason: "no unit match, no usable name" };
  let best = null, bestScore = 0;
  for (const t of buildingTenants) {
    const overlap = courtWords.filter(w => normalizeNameForMatch(t.name).includes(w)).length;
    if (overlap > bestScore) { bestScore = overlap; best = t; }
  }
  if (best && bestScore >= 1) return { tenant: best, reason: `matched by name (${bestScore} word${bestScore === 1 ? "" : "s"} in common)` };
  return { tenant: null, reason: "no unit or name match found" };
}

function HpdViolationsImportSection({ data, add, update, onImported }) {
  const [rawText, setRawText] = useState("");
  const [preview, setPreview] = useState(null);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);
  const [pdfStatus, setPdfStatus] = useState(null);
  const [pdfError, setPdfError] = useState("");
  const [showPaste, setShowPaste] = useState(false);
  const pdfInputRef = useRef(null);

  const runPreview = (text = undefined) => {
    const source = text !== undefined ? text : rawText;
    setError(""); setResult(null);
    if (!source.trim()) { setError("Paste the report text first."); return; }
    const { buildingAddress, violations } = parseHpdViolationsText(source);
    if (violations.length === 0) { setError("No violations found — make sure this is the full HPD \"Open Violations\" report text."); return; }
    const building = findCourtBuildingMatch(buildingAddress, data.buildings);
    const today = todayISO();
    const rows = violations.map(v => {
      const unit = (building && v.apt)
        ? data.units.find(u => u.buildingId === building.id && normalizeAptForMatch(u.unitNumber) === normalizeAptForMatch(v.apt))
        : null;
      // Already certified (HPD's "NOV CERT" status), or a certification was
      // already submitted and HPD is in its verification process (CIV14
      // MAILED — tenant is being asked to confirm the correction; CIV10
      // MAILED — the tenant disputed it) — confirmed via HPD's own support
      // materials that both CIV14 and CIV10 only ever occur after an owner
      // has already certified. None of these are "still awaiting a first
      // certification", so all three map to Certified rather than Open.
      // Lead violations are always included regardless of deadline —
      // they're notoriously slow to resolve (special testing requirements,
      // can't be certified through the fast online process), so an overdue
      // lead hazard is exactly the kind of thing that needs to stay
      // visible, not disappear because the paperwork deadline passed.
      // Everything else: not yet certified but the certify-by deadline has
      // already gone by — skip it, since there's nothing actionable left
      // to track for a deadline that's already passed. Otherwise it's
      // still open and needs attention before its deadline.
      let include, mappedStatus;
      if ((/CERT/i.test(v.status) && !/INVALID/i.test(v.status)) || /^CIV1[04] MAILED$/i.test(v.status)) { include = true; mappedStatus = "Certified"; }
      else if (v.isLead) { include = true; mappedStatus = "Open"; }
      else if (v.certByDate !== "-" && isoFromMDY(v.certByDate) < today) { include = false; mappedStatus = null; }
      else { include = true; mappedStatus = "Open"; }
      const existing = data.violations.find(ev => ev.agency === "HPD" && ev.violationNumber === v.violationId);
      const pastDeadline = v.certByDate !== "-" && isoFromMDY(v.certByDate) < today;
      return { ...v, unit, include, mappedStatus, existing, pastDeadline };
    });
    setPreview({ buildingAddress, building, rows });
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
      runPreview(text);
    } catch (err) {
      setPdfStatus("error");
      setPdfError("Couldn't read that PDF automatically — use \"paste the text instead\" below (open the PDF, Ctrl/Cmd+A, Ctrl/Cmd+C, then paste).");
      setShowPaste(true);
    }
  };

  const runImport = () => {
    if (!preview) return;
    let created = 0, alreadyOnFile = 0, skipped = 0;
    for (const row of preview.rows) {
      if (!row.include) { skipped++; continue; }
      // Already on file — left completely untouched. This runs weekly, and
      // an existing violation is likely already being worked (a vendor
      // called, a note added, status moved along) — re-importing shouldn't
      // reset any of that. Only a genuinely new violation gets added.
      if (row.existing) { alreadyOnFile++; continue; }
      const fields = {
        agency: "HPD", buildingId: preview.building ? preview.building.id : "",
        unitId: row.unit ? row.unit.id : "",
        violationNumber: row.violationId, class: row.class, description: row.description,
        dateIssued: isoFromMDY(row.novIssuedDate) || isoFromMDY(row.reportedDate) || "",
        cureDeadline: isoFromMDY(row.certByDate) || "",
        status: row.mappedStatus, isLead: row.isLead, isMoldOver10: row.isMoldOver10,
      };
      add("violations", { ...fields, id: uid(), fineAmount: "", company: "", otherAgency: "", vendorId: "", notes: [] });
      created++;
    }
    setResult({ created, alreadyOnFile, skipped });
    if (onImported) setTimeout(onImported, 2500);
    setPreview(null); setRawText("");
  };

  return (
    <div className="form-panel" style={{ marginBottom: 16 }}>
      <div style={{ fontWeight: 700, marginBottom: 4 }}>Import HPD violations</div>
      <p className="hint">
        Upload the HPD "Open Violations" building report PDF directly, or paste its text. Matches to a building by address and to units by apartment number.
        Violations past their certification deadline without being certified are skipped; already-certified ones are imported marked as closed.
      </p>
      <div className="form-actions" style={{ marginTop: 4, marginBottom: 8 }}>
        <button className="btn-primary" type="button" onClick={() => pdfInputRef.current.click()} disabled={pdfStatus === "reading"}>
          <Upload size={14} /> {pdfStatus === "reading" ? "Reading PDF…" : "Upload HPD violations PDF"}
        </button>
        <input ref={pdfInputRef} type="file" accept="application/pdf" hidden onChange={handlePdfUpload} />
        <button className="btn-ghost" type="button" onClick={() => setShowPaste(s => !s)}>{showPaste ? "Hide" : "Or paste the text instead"}</button>
      </div>
      {pdfStatus === "error" && <div className="hint" style={{ color: "var(--danger)" }}>{pdfError}</div>}
      {showPaste && (
        <textarea rows={8} value={rawText} onChange={e => { setRawText(e.target.value); setPreview(null); setResult(null); }} placeholder="Paste the full report text here…" />
      )}
      {error && <div className="hint" style={{ color: "var(--danger)" }}>{error}</div>}
      {showPaste && (
        <div className="form-actions" style={{ marginTop: 8 }}>
          <button className="btn-primary" onClick={() => runPreview()} disabled={!rawText.trim()}>Preview</button>
        </div>
      )}
      {preview && (
        <div style={{ marginTop: 12 }}>
          <div className="row">
            <strong>{preview.buildingAddress || "Address not found"}</strong>
            {preview.building
              ? <span className="pill pill-ok" style={{ marginLeft: 8 }}>Matched to {shortAddress(preview.building.address)}</span>
              : <span className="pill pill-danger" style={{ marginLeft: 8 }}>No matching building — add this building first</span>}
          </div>
          <div className="print-stats-row no-print" style={{ margin: "10px 0" }}>
            <div className="print-stat"><div className="print-stat-num">{preview.rows.filter(r => r.mappedStatus === "Open" && !r.pastDeadline).length}</div><div>Open — within deadline</div></div>
            <div className="print-stat"><div className="print-stat-num">{preview.rows.filter(r => r.mappedStatus === "Open" && r.pastDeadline).length}</div><div>Lead, past deadline (still tracked)</div></div>
            <div className="print-stat"><div className="print-stat-num">{preview.rows.filter(r => r.mappedStatus === "Certified").length}</div><div>Already certified</div></div>
            <div className="print-stat"><div className="print-stat-num">{preview.rows.filter(r => r.include && r.existing).length}</div><div>Already on file</div></div>
            <div className="print-stat"><div className="print-stat-num">{preview.rows.filter(r => !r.include).length}</div><div>Skipped — deadline passed</div></div>
            <div className="print-stat"><div className="print-stat-num">{preview.rows.filter(r => r.isLead).length}</div><div>Lead flagged</div></div>
            <div className="print-stat"><div className="print-stat-num">{preview.rows.filter(r => r.isMoldOver10).length}</div><div>Mold ≥10 sq ft</div></div>
          </div>
          <div style={{ maxHeight: 400, overflowY: "auto" }}>
            {preview.rows.filter(r => r.include).sort((a, b) => (isoFromMDY(a.certByDate) || "9999").localeCompare(isoFromMDY(b.certByDate) || "9999")).map(r => (
              <div className="row row-muted" key={r.violationId}>
                <span className="pill pill-accent">{r.class}</span>{" "}
                {r.apt ? `Apt ${r.apt}` : "Building-wide"} — #{r.violationId}
                {r.mappedStatus === "Certified" && <span className="pill pill-ok" style={{ marginLeft: 6 }}>Certified</span>}
                {r.certByDate !== "-" && <span className="pill pill-muted" style={{ marginLeft: 6 }}>Cert by {fmtDate(isoFromMDY(r.certByDate))}</span>}
                {r.isLead && <span className="pill pill-danger" style={{ marginLeft: 6 }}>Lead</span>}
                {r.isMoldOver10 && <span className="pill pill-warn" style={{ marginLeft: 6 }}>Mold ≥10 sq ft</span>}
                {!r.unit && r.apt && <span className="pill pill-muted" style={{ marginLeft: 6 }}>Unit not on file</span>}
                {r.existing && <span className="pill pill-muted" style={{ marginLeft: 6 }}>Already on file — won't be touched</span>}
              </div>
            ))}
          </div>
          <div className="form-actions" style={{ marginTop: 10 }}>
            <button className="btn-primary" onClick={runImport} disabled={!preview.building}>Confirm import</button>
            <button className="btn-ghost" onClick={() => setPreview(null)}>Cancel</button>
          </div>
        </div>
      )}
      {result && (
        <div className="hint" style={{ marginTop: 8 }}>
          Done — {result.created} new violation{result.created === 1 ? "" : "s"} added, {result.alreadyOnFile} already on file (left untouched), {result.skipped} skipped (past deadline, not certified).
        </div>
      )}
    </div>
  );
}

function CourtCaseImportSection({ data, add, update }) {
  const [rawText, setRawText] = useState("");
  const [preview, setPreview] = useState(null);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);
  const [pdfStatus, setPdfStatus] = useState(null); // null | "reading" | "error"
  const [pdfError, setPdfError] = useState("");
  const [showPaste, setShowPaste] = useState(false);
  const pdfInputRef = useRef(null);

  const runPreview = (text = undefined) => {
    const source = text !== undefined ? text : rawText;
    setError(""); setResult(null);
    if (!source.trim()) { setError("Paste the report text first."); return; }
    const cases = parseCourtCasesText(source);
    if (cases.length === 0) { setError("No cases found — make sure this is the full report text, including the \"Case#:\" lines."); return; }
    const rows = cases.map(c => {
      const building = findCourtBuildingMatch(c.address, data.buildings);
      const { tenant, reason } = findCourtTenantMatch(c, building, data.tenants, data.units);
      const yearMatch = c.latestActionDate.match(/^(\d{4})-/);
      const suspiciousYear = yearMatch && Math.abs(parseInt(yearMatch[1]) - new Date().getFullYear()) > 3;
      const existingCase = tenant
        ? data.courtCases.find(cc => cc.tenantId === tenant.id && !cc.archived)
        : data.courtCases.find(cc => !cc.tenantId && cc.caseNumber === c.caseNumber && !cc.archived);
      return { ...c, building, tenant, matchReason: reason, suspiciousYear, existingCase };
    });
    setPreview(rows);
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
      runPreview(text);
    } catch (err) {
      setPdfStatus("error");
      setPdfError("Couldn't read that PDF automatically — use \"paste the text instead\" below (open the PDF, Ctrl/Cmd+A, Ctrl/Cmd+C, then paste).");
      setShowPaste(true);
    }
  };

  const runImport = () => {
    if (!preview) return;
    let created = 0, updated = 0;
    const today = todayISO();
    for (const row of preview) {
      // Even without a matched tenant (a housing-department case like one
      // filed against "DHPD" rather than a person, or a name that doesn't
      // line up with anyone on file), still save it as long as the
      // building matched — visible with its building and details instead
      // of silently vanishing. Only skip when there's truly nothing to
      // link it to.
      if (!row.tenant && !row.building) continue;
      const existingLog = row.existingCase?.log || [];
      // A future-dated action (like a scheduled "Court Appearance") hasn't
      // happened yet — logging it under its own future date would put it
      // at the top of the log once sorted newest-first, making it look
      // like the most recent thing that already occurred. Instead, it's
      // logged under TODAY's date (when this was actually learned from
      // the report) with the real date spelled out in the note itself —
      // "Court Appearance — scheduled for 09/27/2026" rather than dating
      // the entry itself as 09/27. A genuinely past action still gets
      // logged under its own real date as usual.
      const newEntries = (row.actions || [])
        .map(a => a.date > today
          ? { ...a, note: `${a.desc} — scheduled for ${fmtDate(a.date)}`, isRescheduleNote: true }
          : { ...a, note: a.desc, isRescheduleNote: false })
        .filter(a => a.isRescheduleNote
          // Rescheduled notes get re-dated to today on every import while
          // still upcoming, so dedup by the note text itself (which
          // embeds the real date) rather than by date.
          ? !existingLog.some(l => l.note === a.note)
          : !existingLog.some(l => l.date === a.date && l.note === a.note))
        .map(a => ({ id: uid(), date: a.isRescheduleNote ? today : a.date, note: a.note, source: "attorney report" }));
      const newLog = [...existingLog, ...newEntries];
      // The report's own action history often already names a scheduled
      // future court date ("Court Appearance" dated ahead of today) — pull
      // the earliest one as the case's next court date instead of leaving
      // it blank and making the person re-enter a date the report already
      // gave. Picking the EARLIEST future one, not just any, since that's
      // the next thing actually coming up.
      const futureAppearances = (row.actions || [])
        .filter(a => a.date > today && /Court Appearance/i.test(a.desc))
        .sort((a, b) => a.date.localeCompare(b.date));
      const parsedNextCourtDate = futureAppearances.length > 0 ? futureAppearances[0].date : "";
      const fields = {
        tenantId: row.tenant ? row.tenant.id : "",
        buildingId: row.tenant ? row.tenant.buildingId : row.building.id,
        unitId: row.tenant ? row.tenant.unitId : "",
        // The report's own name (e.g. "DHPD") — the only label available
        // when there's no matched tenant to pull a name from.
        rawName: row.tenant ? "" : row.name,
        caseNumber: row.caseNumber, log: newLog,
      };
      if (row.existingCase) {
        // Never overwrite a date already on file — it may have been set or
        // corrected by hand since the last import, and this report's own
        // "future" date could by now be in the past relative to it.
        if (!row.existingCase.nextCourtDate && parsedNextCourtDate) fields.nextCourtDate = parsedNextCourtDate;
        update("courtCases", row.existingCase.id, fields);
        updated++;
      } else {
        add("courtCases", {
          ...fields, stage: CASE_STAGES[0], nextCourtDate: parsedNextCourtDate, result: "Pending",
          stipulationTerms: "", nextPaymentDue: "", archived: false, documents: [],
          checklist: DEFAULT_ATTORNEY_CHECKLIST.map(label => ({ id: uid(), label, checked: false })),
        });
        created++;
      }
    }
    setResult({ created, updated, unmatched: preview.filter(r => !r.tenant && !r.building).length });
    setPreview(null); setRawText("");
  };

  return (
    <div className="form-panel" style={{ marginBottom: 16 }}>
      <div style={{ fontWeight: 700, marginBottom: 4 }}>Import from attorney report</div>
      <p className="hint">Upload the "Complete Client Status" PDF directly, or paste its text. Each case gets matched to a tenant by building address and unit number (falling back to name matching when the report's apt value doesn't line up with a real unit, like a storefront listed as "STORE FRONT"). Review the matches below before confirming — nothing is saved until you click Confirm import.</p>
      <div className="form-actions" style={{ marginTop: 4, marginBottom: 8 }}>
        <button className="btn-primary" type="button" onClick={() => pdfInputRef.current.click()} disabled={pdfStatus === "reading"}>
          <Upload size={14} /> {pdfStatus === "reading" ? "Reading PDF…" : "Upload attorney report PDF"}
        </button>
        <input ref={pdfInputRef} type="file" accept="application/pdf" hidden onChange={handlePdfUpload} />
        <button className="btn-ghost" type="button" onClick={() => setShowPaste(s => !s)}>{showPaste ? "Hide" : "Or paste the text instead"}</button>
      </div>
      {pdfStatus === "error" && <div className="hint" style={{ color: "var(--danger)" }}>{pdfError}</div>}
      {showPaste && (
        <textarea rows={8} value={rawText} onChange={e => { setRawText(e.target.value); setPreview(null); setResult(null); }} placeholder="Paste the full report text here…" />
      )}
      {error && <div className="hint" style={{ color: "var(--danger)" }}>{error}</div>}
      {showPaste && (
        <div className="form-actions" style={{ marginTop: 8 }}>
          <button className="btn-primary" onClick={() => runPreview()} disabled={!rawText.trim()}>Preview</button>
        </div>
      )}
      {preview && (
        <div style={{ marginTop: 12 }}>
          {preview.map(row => (
            <div key={row.caseNumber} className="hint" style={{ marginBottom: 6, paddingBottom: 6, borderBottom: "1px solid var(--border)" }}>
              <strong>Case #{row.caseNumber}</strong> — {row.name || "(no name on file)"}, apt {row.apt || "?"}
              <br />
              {row.tenant
                ? <span style={{ color: "var(--ok)" }}>✓ {row.existingCase ? "will update" : "will create"} — matched to {row.tenant.name} ({row.matchReason})</span>
                : <span style={{ color: "var(--warn)" }}>⚠ not matched — {row.matchReason}. Skipped; add manually if needed.</span>}
              <br />
              Latest: {fmtDate(row.latestActionDate)} — {row.latestActionDesc}
              {row.suspiciousYear && <span style={{ color: "var(--danger)" }}> ⚠ this date's year looks off — check the source report</span>}
            </div>
          ))}
          <div className="form-actions" style={{ marginTop: 8 }}>
            <button className="btn-primary" onClick={runImport}>Confirm import</button>
            <button className="btn-ghost" onClick={() => setPreview(null)}>Cancel</button>
          </div>
        </div>
      )}
      {result && (
        <div className="hint" style={{ marginTop: 12, fontWeight: 700, color: "var(--ok)" }}>
          Done — {result.created} case{result.created === 1 ? "" : "s"} created, {result.updated} updated{result.unmatched > 0 ? `, ${result.unmatched} skipped (no match)` : ""}.
        </div>
      )}
    </div>
  );
}

function CourtTab({ data, add, update, remove, setData, tenantName, buildingName }) {
  const [form, setForm] = useState(null);
  const [view, setView] = useState("active");
  const [checklistText, setChecklistText] = useState({});
  const [detailsFor, setDetailsFor] = useState(null);
  const [section, setSection] = useState("cases"); // cases | import
  const [buildingFilter, setBuildingFilter] = useState("All");
  const [dueOnlyFilter, setDueOnlyFilter] = useState(null); // null | "court" | "stip"
  const [logForm, setLogForm] = useState({});
  const [confirmingDeleteAll, setConfirmingDeleteAll] = useState(false);

  const deleteAllCases = () => {
    setData(d => ({ ...d, courtCases: [] }));
    setConfirmingDeleteAll(false);
  };

  const addLogEntry = (c) => {
    const entry = logForm[c.id];
    if (!entry || !entry.note || !entry.note.trim()) return;
    const newEntry = { id: uid(), date: entry.date || todayISO(), note: entry.note.trim(), source: "manual" };
    update("courtCases", c.id, { log: [...(c.log || []), newEntry] });
    setLogForm({ ...logForm, [c.id]: { date: todayISO(), note: "" } });
  };

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

  const list = data.courtCases
    .filter(c => view === "closed" ? c.archived : !c.archived)
    .filter(c => buildingFilter === "All" || c.buildingId === buildingFilter)
    .filter(c => {
      if (dueOnlyFilter === "court") return c.result === "Stipulation (payment plan)" ? dateDueStrict(c.nextCourtDate) : dateDue(c.nextCourtDate);
      if (dueOnlyFilter === "stip") return c.result === "Stipulation (payment plan)" && dateDue(c.nextPaymentDue);
      return true;
    })
    .slice()
    .sort((a, b) => (a.nextCourtDate || "9999-99-99").localeCompare(b.nextCourtDate || "9999-99-99"));

  // Buildings that actually have at least one active case, so the filter
  // row only ever shows relevant options, not every building regardless
  // of whether it has any cases.
  const buildingsWithCases = data.buildings.filter(b => data.courtCases.some(c => c.buildingId === b.id && !c.archived));
  const courtDatesDueCount = data.courtCases.filter(c => !c.archived && (c.result === "Stipulation (payment plan)" ? dateDueStrict(c.nextCourtDate) : dateDue(c.nextCourtDate))).length;
  const stipDueCount = data.courtCases.filter(c => !c.archived && c.result === "Stipulation (payment plan)" && dateDue(c.nextPaymentDue)).length;

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
    <div className="court-page">
      <div className="page-head">
        <h1 className="page-title">Court Cases</h1>
        <div className="page-actions">
          <PrintButton label="Court Cases" />
          {data.courtCases.length > 0 && (
            confirmingDeleteAll ? (
              <>
                <span className="row-muted" style={{ fontSize: 12 }}>Delete all {data.courtCases.length} case{data.courtCases.length === 1 ? "" : "s"}? This can't be undone.</span>
                <button className="btn-ghost" style={{ color: "var(--danger)" }} onClick={deleteAllCases}>Yes, delete all</button>
                <button className="btn-ghost" onClick={() => setConfirmingDeleteAll(false)}>Cancel</button>
              </>
            ) : (
              <button className="btn-ghost" style={{ color: "var(--danger)" }} onClick={() => setConfirmingDeleteAll(true)}><Trash2 size={14} /> Delete all</button>
            )
          )}
          <button className="btn-primary" onClick={() => setForm({ tenantId: "", buildingId: "", unitId: "", caseNumber: "", nextCourtDate: "", result: "Pending", stage: CASE_STAGES[0], stipulationTerms: "", nextPaymentDue: "" })}>
            <Plus size={14} /> Add case
          </button>
        </div>
      </div>

      <div className="no-print" style={{ display: "flex", gap: 10, marginBottom: 14 }}>
        <div className="dash-stat-card" style={{ flex: 1, cursor: "default" }}>
          <div className="dash-stat-num">{data.courtCases.filter(c => !c.archived).length}</div>
          <div className="dash-stat-label">Active cases</div>
        </div>
        <button
          className="dash-stat-card"
          style={{ flex: 1, borderColor: dueOnlyFilter === "court" ? "var(--navy)" : undefined }}
          onClick={() => { setSection("cases"); setDueOnlyFilter(f => f === "court" ? null : "court"); }}
        >
          <div className="dash-stat-num" style={{ color: courtDatesDueCount > 0 ? "var(--danger)" : undefined }}>{courtDatesDueCount}</div>
          <div className="dash-stat-label">Court dates coming up</div>
        </button>
        <button
          className="dash-stat-card"
          style={{ flex: 1, borderColor: dueOnlyFilter === "stip" ? "var(--navy)" : undefined }}
          onClick={() => { setSection("cases"); setDueOnlyFilter(f => f === "stip" ? null : "stip"); }}
        >
          <div className="dash-stat-num" style={{ color: stipDueCount > 0 ? "var(--warn)" : undefined }}>{stipDueCount}</div>
          <div className="dash-stat-label">Payments coming up</div>
        </button>
      </div>

      <div className="filter-row">
        <button className={`chip ${section === "cases" ? "chip-active" : ""}`} onClick={() => setSection("cases")}>Cases</button>
        <button className={`chip ${section === "import" ? "chip-active" : ""}`} onClick={() => setSection("import")}>Import from attorney report</button>
      </div>

      {section === "cases" && buildingsWithCases.length > 1 && (
        <div className="filter-row">
          <button className={`chip ${buildingFilter === "All" ? "chip-active" : ""}`} onClick={() => setBuildingFilter("All")}>All buildings</button>
          {buildingsWithCases.map(b => (
            <button key={b.id} className={`chip ${buildingFilter === b.id ? "chip-active" : ""}`} onClick={() => setBuildingFilter(b.id)}>{shortAddress(b.address)}</button>
          ))}
        </div>
      )}

      {section === "import" && <CourtCaseImportSection data={data} add={add} update={update} />}

      <div className="print-only">
        <div className="print-header">
          <div className="print-mark">O</div>
          <div className="print-header-text">
            <h1 className="print-title">Court Cases — {view === "closed" ? "Closed / Archived" : "Active"}</h1>
            <div className="print-subtitle">As of {fmtDate(todayISO())}</div>
          </div>
        </div>
        <div className="print-stats-row">
          <div className="print-stat"><div className="print-stat-num">{list.length}</div><div>Total cases</div></div>
          <div className="print-stat"><div className="print-stat-num">{list.filter(c => c.nextCourtDate).length}</div><div>With a court date</div></div>
          <div className="print-stat"><div className="print-stat-num">{list.filter(c => c.result === "Stipulation (payment plan)").length}</div><div>Payment plans</div></div>
        </div>
        {[...new Set(list.map(c => c.buildingId))].map(bId => {
          const items = list.filter(c => c.buildingId === bId);
          return (
            <div className="print-section" key={bId || "unknown"}>
              <div className="print-section-head">
                <span>{buildingName(bId)}</span>
                <span>{items.length} case{items.length === 1 ? "" : "s"}</span>
              </div>
              <table className="print-table">
                <thead><tr><th>Tenant</th><th>Docket #</th><th>Next court date</th><th>Stage</th><th>Result</th></tr></thead>
                <tbody>
                  {items.map(c => (
                    <tr key={c.id}>
                      <td>{c.tenantId ? tenantName(c.tenantId) : (c.rawName || "(no tenant matched)")}</td>
                      <td>{c.caseNumber || "—"}</td>
                      <td>{c.nextCourtDate ? fmtDate(c.nextCourtDate) : "not set"}</td>
                      <td>{c.stage || "—"}</td>
                      <td>{c.result || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          );
        })}
        <div className="print-footer">
          <span>Property Ops — Court Cases</span>
          <span>Generated {fmtDate(todayISO())}</span>
        </div>
      </div>
      {section === "cases" && dueOnlyFilter && (
        <div className="hint" style={{ marginBottom: 10 }}>
          Showing only {dueOnlyFilter === "court" ? "cases with a court date coming up or overdue" : "cases with a payment coming up or overdue"}.{" "}
          <button className="btn-ghost" style={{ padding: "2px 8px" }} onClick={() => setDueOnlyFilter(null)}>Clear</button>
        </div>
      )}

      {section === "cases" && (
      <div className="filter-row">
        <button className={`chip ${view === "active" ? "chip-active" : ""}`} onClick={() => setView("active")}>Active</button>
        <button className={`chip ${view === "closed" ? "chip-active" : ""}`} onClick={() => setView("closed")}>Closed / Archived</button>
      </div>
      )}

      {section === "cases" && form && (
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
                    {data.units.find(u => u.id === t.unitId)?.unitNumber || "—"} — {t.name || "(no name on file)"}{t.movedOut ? " (moved out)" : ""}
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

      {section === "cases" && list.length === 0 && <EmptyState text={view === "active" ? "No open court cases." : "Nothing closed yet."} />}
      {section === "cases" && list.map(c => {
        const checkedCount = (c.checklist || []).filter(i => i.checked).length;
        const sortedLog = [...(c.log || [])].sort((a, b) => (b.date || "").localeCompare(a.date || ""));
        const latestLog = sortedLog[0];
        const isOpen = detailsFor === c.id;
        return (
        <div className="list-card" key={c.id}>
          <div className="list-card-head" onClick={() => setDetailsFor(isOpen ? null : c.id)} style={{ cursor: "pointer" }}>
            {isOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
            {c.unitId && <span className="pill pill-accent">Apt {data.units.find(u => u.id === c.unitId)?.unitNumber || "—"}</span>}
            <div className="list-card-title">{c.tenantId ? tenantName(c.tenantId) : (c.rawName || "(no tenant matched)")} {c.caseNumber && `· Docket #${c.caseNumber}`}</div>
            <span className="pill pill-muted">{buildingName(c.buildingId)}</span>
            {c.nextCourtDate && !c.archived && <Flag date={c.nextCourtDate} label="court date" />}
          </div>
          {/* Always visible even collapsed — the whole point is not having to
              open every card just to see what's currently going on with it. */}
          <div className="list-card-body" style={{ padding: "8px 14px" }}>
            {latestLog
              ? <div className="row"><em>{latestLog.date > todayISO() ? "Scheduled" : "Latest"} ({fmtDate(latestLog.date)}):</em> {latestLog.note}</div>
              : <div className="hint">No log entries yet.</div>}
          </div>
          {isOpen && (
            <>
              <div className="list-card-head" style={{ paddingTop: 0 }}>
                {c.stage && <span className="pill pill-muted">{c.stage}</span>}
                <span className="pill pill-muted">{c.result}</span>
                {c.tenantId && (() => {
                  const t = data.tenants.find(x => x.id === c.tenantId);
                  if (!t) return null;
                  const bal = parseBalance(t.balance);
                  return <span className={`pill ${bal > 0 ? "pill-warn" : "pill-ok"}`}>${bal.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} owed</span>;
                })()}
                <div className="spacer" />
                <IconBtn title="Edit" onClick={(e) => { e.stopPropagation(); openEdit(c); }}><Pencil size={14} /></IconBtn>
                {view === "closed"
                  ? <IconBtn title="Restore to active" onClick={(e) => { e.stopPropagation(); update("courtCases", c.id, { archived: false }); }}><ArchiveIcon size={14} /></IconBtn>
                  : <IconBtn title="Move to closed" onClick={(e) => { e.stopPropagation(); update("courtCases", c.id, { archived: true }); }}><ArchiveIcon size={14} /></IconBtn>}
                <IconBtn title="Delete" danger onClick={(e) => { e.stopPropagation(); remove("courtCases", c.id); }}><Trash2 size={14} /></IconBtn>
              </div>
              <div className="list-card-body">
                <div className="row"><strong>Log</strong></div>
                {sortedLog.length === 0
                  ? <div className="hint" style={{ marginBottom: 6 }}>Nothing logged yet — add the first entry below, or import from an attorney report.</div>
                  : sortedLog.map(l => {
                      const isFuture = l.date > todayISO();
                      return (
                      <div className="row" key={l.id} style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
                        <span style={{ minWidth: 90, color: "var(--ink-soft)", fontSize: 12 }}>{fmtDate(l.date)}</span>
                        <span style={{ flex: 1 }}>
                          {isFuture && <span className="pill pill-warn" style={{ marginRight: 6, fontSize: 10 }}>Scheduled — hasn't happened yet</span>}
                          {l.note}{l.source === "attorney report" && <span className="pill pill-muted" style={{ marginLeft: 6, fontSize: 10 }}>from report</span>}
                        </span>
                        <button className="checklist-remove" title="Remove this entry" onClick={() => update("courtCases", c.id, { log: (c.log || []).filter(x => x.id !== l.id) })}><X size={12} /></button>
                      </div>
                      );
                    })}
                <div className="inline-form" style={{ marginTop: 8 }}>
                  <input type="date" value={logForm[c.id]?.date || todayISO()} onChange={e => setLogForm({ ...logForm, [c.id]: { ...logForm[c.id], date: e.target.value } })} />
                  <input placeholder="What happened…" value={logForm[c.id]?.note || ""} onChange={e => setLogForm({ ...logForm, [c.id]: { ...logForm[c.id], note: e.target.value } })} onKeyDown={e => e.key === "Enter" && addLogEntry(c)} />
                  <button className="btn-ghost" onClick={() => addLogEntry(c)}>Add</button>
                </div>
              </div>
              {c.result === "Stipulation (payment plan)" && (
                <div className="list-card-body">
                  {c.stipulationTerms && <div className="row"><em>Terms:</em> {c.stipulationTerms}</div>}
                  {c.nextPaymentDue && <div className="row">Next payment due: <Flag date={c.nextPaymentDue} /></div>}
                </div>
              )}
              <div className="list-card-body" style={{ paddingTop: c.result === "Stipulation (payment plan)" ? 0 : undefined }}>
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
                  pathPrefix={`courtCases/${c.id}/documents`}
                  onAdd={(newDocs) => update("courtCases", c.id, { documents: [...(c.documents || []), ...newDocs] })}
                  onRemove={(d) => {
                    update("courtCases", c.id, { documents: (c.documents || []).filter(x => x.id !== d.id) });
                    if (d.storagePath) deleteObject(storageRef(storage, d.storagePath)).catch(() => {});
                  }}
                />
              </>
              </div>
            </>
          )}
        </div>
      );})}
    </div>
  );
}

/* ============================== appointments (incl. recurring) ============================== */

function AppointmentsTab({ data, add, update, remove, buildingName, setData }) {
  const [form, setForm] = useState(null);
  const [view, setView] = useState("upcoming");
  const [expandedRow, setExpandedRow] = useState(null);
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

  // Defaults to 9am-11am so the time picker opens near a normal workday start
  // instead of midnight — still fully editable to anything earlier or later.
  const blank = () => ({ buildingId: data.buildings[0]?.id || "", unitId: "", type: allTypes[0] || "Other", date: "", timeFrom: "09:00", timeTo: "11:00", notes: "", recurring: false, completed: false });

  const list = data.appointments
    .filter(a => view === "upcoming" ? !a.completed : a.completed)
    .slice()
    .sort((a, b) => {
      const cmp = (a.date || "").localeCompare(b.date || "") || (a.timeFrom || "").localeCompare(b.timeFrom || "");
      return view === "upcoming" ? cmp : -cmp;
    });

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
        const isOpen = expandedRow === a.id;
        return (
          <div className="list-card" key={a.id}>
            <div className="list-card-head" onClick={() => setExpandedRow(isOpen ? null : a.id)} style={{ cursor: "pointer" }}>
              {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
              <div className="list-card-title">{a.type}</div>
              <span className="pill pill-muted">{buildingName(a.buildingId)}</span>
              {a.unitId && <span className="pill pill-muted">Unit {data.units.find(u => u.id === a.unitId)?.unitNumber || "—"}</span>}
              {view === "upcoming" && <Flag date={a.date} />}
              {view === "completed" && <span className="pill pill-muted">{fmtDate(a.date)}</span>}
              {(a.timeFrom || a.timeTo) && (
                <span className="pill pill-muted">{fmtTime(a.timeFrom)}{a.timeTo ? ` – ${fmtTime(a.timeTo)}` : ""}</span>
              )}
              <div className="spacer" />
              <IconBtn title={view === "upcoming" ? "Mark completed" : "Move back to upcoming"} onClick={(e) => { e.stopPropagation(); update("appointments", a.id, { completed: !a.completed }); }}>
                <CheckCircle2 size={14} />
              </IconBtn>
              <IconBtn title="Edit" onClick={(e) => { e.stopPropagation(); setForm(a); }}><Pencil size={14} /></IconBtn>
              <IconBtn title="Delete" danger onClick={(e) => { e.stopPropagation(); remove("appointments", a.id); }}><Trash2 size={14} /></IconBtn>
            </div>
            {isOpen && (
              <div className="list-card-body">
                <div className="row" style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {a.recurring && <span className="pill pill-muted">Recurring</span>}
                  {reminderHit && <span className="pill pill-warn">Reminder</span>}
                </div>
                {a.notes && <div className="row">{a.notes}</div>}
                {view === "upcoming" && (
                  <a className="btn-ghost" style={{ marginTop: 6, display: "inline-flex" }} href={icsFor(`${a.type} — ${buildingName(a.buildingId)}`, a.date, a.notes, a.timeFrom, a.timeTo)} download={`${(a.type || "appointment").replace(/\s/g, "-")}.ics`}>
                    <Download size={14} /> Add to calendar
                  </a>
                )}
              </div>
            )}
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
            <div className="law-name">{LOCAL_LAWS.find(l => l.key === row.lawKey)?.name || row.lawKey}</div>
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
      {data.bossReminders.slice().sort((a, b) => (b.dateRaised || "").localeCompare(a.dateRaised || "")).map(r => (
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
  const [showAllReminders, setShowAllReminders] = useState(false);

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

  const today = todayISO();
  const allNotes = (data.quickNotes || []).slice().reverse();
  // Notes = the plain checklist (no reminder date). Reminders = notes that DO have
  // a reminder date, shown separately, and only once due unless "All upcoming" is picked.
  const plainNotes = (showDone ? allNotes : allNotes.filter(n => !n.done)).filter(n => !n.reminderDate);
  const allReminders = (data.quickNotes || []).filter(n => n.reminderDate).sort((a, b) => a.reminderDate.localeCompare(b.reminderDate));
  const remindersToShow = (showAllReminders ? allReminders : allReminders.filter(n => n.reminderDate <= today)).filter(n => showDone || !n.done);

  const NoteRow = (n) => (
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
  );

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Quick Notes / Reminder</h1>
        <PrintButton label="Quick Notes" />
      </div>

      <h2 className="section-heading">Notes</h2>
      <p className="hint">A checklist for anything you need to jot down fast — check it off when it's handled, and tie it to a building if it's related to one.</p>
      <div className="form-panel">
        <Field label="Note"><input placeholder="Jot something down…" value={text} onChange={e => setText(e.target.value)} onKeyDown={e => e.key === "Enter" && submit()} /></Field>
        <Field label="Reminder date (optional — makes this a Reminder instead)"><input type="date" value={reminderDate} onChange={e => setReminderDate(e.target.value)} /></Field>
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
      {plainNotes.length === 0 && <EmptyState text="Nothing here." />}
      {plainNotes.map(NoteRow)}

      <h2 className="section-heading" style={{ marginTop: 28 }}>Reminders <span className="dash-panel-sub">— notes with a date, shown once due</span></h2>
      <p className="hint">Same checklist as above, just for anything you gave a reminder date. Stays off this list and the Dashboard until that date arrives.</p>
      <div className="filter-row">
        <button className={`chip ${!showAllReminders ? "chip-active" : ""}`} onClick={() => setShowAllReminders(false)}>Due now</button>
        <button className={`chip ${showAllReminders ? "chip-active" : ""}`} onClick={() => setShowAllReminders(true)}>All upcoming</button>
      </div>
      {remindersToShow.length === 0 && <EmptyState text={showAllReminders ? "No reminders set." : "Nothing due yet."} />}
      {remindersToShow.map(NoteRow)}
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
        --brand-blue: #2F6FE0;
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
        overflow-wrap: break-word;
      }
      .loading { padding: 40px; text-align: center; color: var(--ink-soft); }
      .topbar {
        display: flex; align-items: center; justify-content: flex-start;
        gap: 16px; padding: 14px 20px; background: var(--navy); color: #fff;
        padding-top: max(14px, env(safe-area-inset-top));
        padding-left: max(20px, env(safe-area-inset-left));
        padding-right: max(20px, env(safe-area-inset-right));
      }
      .topbar-left { display: flex; align-items: center; gap: 12px; flex-shrink: 0; }
      .topbar-actions { display: flex; align-items: center; gap: 8px; flex-shrink: 0; margin-left: auto; }
      .menu-btn {
        background: rgba(255,255,255,0.1); border: none; color: #fff; width: 34px; height: 34px;
        border-radius: 6px; display: flex; align-items: center; justify-content: center; cursor: pointer; flex-shrink: 0;
      }
      .menu-btn:hover { background: rgba(255,255,255,0.2); }
      .brand { display: flex; align-items: center; gap: 10px; background: none; border: none; padding: 0; cursor: pointer; text-align: left; font: inherit; color: inherit; }
      .brand-mark {
        width: 34px; height: 34px; border-radius: 50%; background: var(--brand-blue);
        display: flex; align-items: center; justify-content: center; color: #fff;
        font-family: Georgia, "Times New Roman", serif; font-weight: 700; font-size: 17px;
      }
      .brand-title { font-family: Georgia, "Times New Roman", serif; font-size: 16px; line-height: 1.2; }
      .brand-sub { font-size: 11px; color: #C9CFD6; }
      .search-wrap { position: relative; flex: 0 1 440px; margin-left: 16px; }
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
      .violations-group-heading { display: flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 700; color: var(--ink); text-transform: uppercase; letter-spacing: 0.03em; padding-bottom: 6px; margin-bottom: 8px; border-bottom: 2px solid var(--border); }
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
      .icon-btn-active { background: var(--navy); color: #fff; }
      .icon-btn-active:hover { background: var(--navy); color: #fff; }
      .form-panel {
        background: var(--panel); border: 1px solid var(--border); border-radius: 8px;
        padding: 16px; margin: 12px 0 18px; display: grid; grid-template-columns: 1fr 1fr; gap: 12px;
      }
      .form-panel .field:has(textarea) { grid-column: 1 / -1; }
      .form-actions { grid-column: 1 / -1; display: flex; gap: 8px; }
      .field { display: flex; flex-direction: column; gap: 4px; font-size: 12px; color: var(--ink-soft); }
      .hearing-checkbox-row { display: flex; align-items: center; gap: 8px; font-size: 13px; color: var(--ink); cursor: pointer; padding: 7px 0; }
      .hearing-checkbox-row input[type="checkbox"] { width: auto; flex-shrink: 0; }
      .field input, .field select, .field textarea {
        font-size: 13px; padding: 7px 9px; border: 1px solid var(--border); border-radius: 5px;
        background: #fff; color: var(--ink); font-family: inherit; width: 100%; box-sizing: border-box;
      }
      .field textarea { min-height: 60px; resize: vertical; }
      .list-card { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; margin-bottom: 10px; overflow: hidden; }
      .list-card-compact { margin-bottom: 4px; }
      .list-card-compact .list-card-head { padding: 6px 12px; gap: 6px; font-size: 13px; }
      .list-card-compact .violation-desc-preview { font-size: 12px; }
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
      .search-result-row { display: block; width: 100%; text-align: left; background: none; border: none; font: inherit; color: inherit; cursor: pointer; border-radius: 4px; }
      .search-result-row:hover { background: var(--panel); padding-left: 4px; }
      .row-muted { color: var(--ink-soft); }
      .strike { text-decoration: line-through; color: var(--ink-soft); }
      .inline-form { display: flex; gap: 8px; margin: 8px 0; flex-wrap: wrap; }
      .inline-form input { flex: 1; min-width: 120px; padding: 7px 9px; border: 1px solid var(--border); border-radius: 5px; font-size: 13px; }
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
      .law-row input, .law-row select { padding: 6px 8px; border: 1px solid var(--border); border-radius: 5px; font-size: 12px; width: 100%; box-sizing: border-box; }
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
      .dash-panel-sub { font-size: 11px; color: var(--ink-soft); font-weight: 400; text-transform: none; letter-spacing: normal; }
      .all-clear {
        display: flex; align-items: center; gap: 10px; background: var(--ok-bg); color: var(--ok);
        border-radius: 8px; padding: 16px; font-size: 13px; margin-bottom: 24px;
      }
      .dash-stats-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 16px; }
      .dash-quick-note {
        display: flex; align-items: center; gap: 8px; background: var(--panel); border: 1px solid var(--border);
        border-radius: 8px; padding: 6px 12px; margin-bottom: 14px;
      }
      .dash-quick-note-icon { color: var(--ink-soft); flex-shrink: 0; }
      .dash-quick-note input { border: none; background: none; outline: none; font-size: 13px; width: 100%; padding: 4px 0; }
      .dash-arrears-nudge {
        display: flex; align-items: center; gap: 8px; width: 100%; text-align: left;
        background: var(--panel); color: var(--ink); border: 1px solid var(--border); border-left: 3px solid var(--warn);
        border-radius: 8px; padding: 8px 12px; margin-bottom: 14px; font-size: 13px; cursor: pointer; font: inherit;
      }
      .dash-arrears-nudge svg { color: var(--warn); flex-shrink: 0; }
      .dash-arrears-nudge:hover { border-color: var(--navy); }
      .dash-stat-card {
        background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 12px; text-align: center;
        cursor: pointer; font: inherit; width: 100%;
      }
      .dash-stat-card:hover { border-color: var(--navy); }
      .dash-stat-detail {
        background: var(--panel); border: 1px solid var(--border); border-radius: 8px;
        padding: 10px; margin-bottom: 16px; display: flex; flex-direction: column; gap: 6px;
      }
      .dash-stat-num { font-size: 22px; font-weight: 700; color: var(--navy); line-height: 1.2; }
      .dash-stat-label { font-size: 11px; color: var(--ink-soft); margin-top: 2px; }
      .dash-stat-sub { font-size: 10px; color: var(--ink-soft); margin-top: 1px; }
      .dash-col-main { max-width: none; }
      .dash-roster { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px; margin-bottom: 16px; }
      .dash-roster-title { font-size: 12px; font-weight: 700; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.03em; margin-bottom: 8px; }
      .dash-roster-row { display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 8px 0; border-bottom: 1px solid var(--border); }
      .dash-roster-row:last-of-type { border-bottom: none; }
      .dash-roster-name { font-size: 13px; font-weight: 500; }
      .dash-roster-sub { font-size: 12px; color: var(--ink-soft); }
      .dash-roster-meta { display: flex; align-items: center; gap: 8px; flex-shrink: 0; }
      .dash-roster-phone { display: flex; color: var(--navy); }
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
      .rent-total-banner {
        background: var(--panel); border: 1px solid var(--border); border-left: 4px solid var(--navy);
        border-radius: 8px; padding: 12px 16px; margin-bottom: 14px;
      }
      .rent-total-num { font-size: 26px; font-weight: 700; color: var(--navy); line-height: 1.2; }
      .rent-total-label { font-size: 12px; color: var(--ink-soft); margin-top: 2px; }
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
      .sheet-row-selected { background: #E8F0FE !important; box-shadow: inset 3px 0 0 var(--navy); cursor: pointer; }
      .sheet-row-selected td { background: transparent; }
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
      .phone-cycle-cell { display: flex; align-items: center; gap: 2px; flex-wrap: wrap; }
      .phone-cycle-cell .sheet-input { flex: 1; min-width: 80px; }
      .phone-cycle-btn {
        display: flex; align-items: center; gap: 1px; background: none; border: 1px solid var(--border);
        border-radius: 4px; color: var(--ink-soft); cursor: pointer; padding: 2px 4px; font-size: 10px; white-space: nowrap;
      }
      .phone-cycle-btn:hover { background: #F0EEE7; color: var(--navy); }
      .followup-scroll { max-height: 520px; overflow-y: auto; padding-right: 4px; }
      .dash-calendar {
        background: var(--panel); border: 1px solid var(--border); border-radius: 8px;
        padding: 16px 18px; margin-bottom: 18px;
      }
      .dash-calendar-hero { margin-bottom: 4px; }
      .dash-calendar-head { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; color: var(--navy); }
      .dash-calendar-title { font-weight: 700; font-size: 16px; }
      .dash-cal-modes { margin-bottom: 10px; gap: 6px; }
      .dash-cal-modes .chip { font-size: 12px; padding: 4px 12px; }
      .dash-cal-week { display: grid; grid-template-columns: repeat(7, 1fr); gap: 6px; margin-bottom: 10px; }
      .dash-cal-month { display: grid; grid-template-columns: repeat(7, 1fr); gap: 4px; margin-bottom: 10px; }
      .dash-cal-month-dow { text-align: center; font-size: 11px; color: var(--ink-soft); font-weight: 700; padding-bottom: 3px; }
      .dash-cal-cell {
        display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 2px;
        background: #fff; border: 1px solid var(--border); border-radius: 6px;
        padding: 8px 4px; cursor: pointer; position: relative; font: inherit;
      }
      .dash-cal-cell-sm { padding: 4px 3px; aspect-ratio: 1; }
      .dash-cal-cell-week { align-items: stretch; justify-content: flex-start; min-height: 90px; padding: 8px 6px; gap: 3px; }
      .dash-cal-cell-week .dash-cal-cell-label, .dash-cal-cell-week .dash-cal-cell-num { align-self: center; }
      .dash-cal-cell:hover { border-color: var(--navy); }
      .dash-cal-cell-today { border-color: var(--navy); border-width: 2px; }
      .dash-cal-cell-selected { background: var(--navy); }
      .dash-cal-cell-selected .dash-cal-cell-label, .dash-cal-cell-selected .dash-cal-cell-num { color: #fff; }
      .dash-cal-cell-dim { opacity: 0.35; }
      .dash-cal-cell-label { font-size: 10px; text-transform: uppercase; color: var(--ink-soft); letter-spacing: 0.02em; }
      .dash-cal-cell-num { font-size: 15px; font-weight: 700; }
      .dash-cal-cell-items { display: flex; flex-direction: column; gap: 2px; margin-top: 2px; width: 100%; }
      .dash-cal-cell-item {
        font-size: 10px; line-height: 1.25; text-align: left; color: var(--ink);
        background: var(--warn-bg); border-radius: 3px; padding: 2px 4px;
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
      }
      .dash-cal-cell-selected .dash-cal-cell-item { background: rgba(255,255,255,0.15); color: #fff; }
      .dash-cal-dot { width: 6px; height: 6px; border-radius: 999px; background: var(--danger); }
      .dash-cal-count {
        position: absolute; top: 2px; right: 2px; min-width: 14px; height: 14px; padding: 0 3px;
        border-radius: 999px; background: var(--danger); color: #fff; font-size: 9px; font-weight: 700;
        display: flex; align-items: center; justify-content: center; line-height: 1;
      }
      .dash-cal-cell-selected .dash-cal-count { background: #fff; color: var(--navy); }
      .dash-cal-year { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin-bottom: 10px; }
      .dash-cal-month-cell {
        display: flex; flex-direction: column; align-items: center; gap: 3px;
        background: #fff; border: 1px solid var(--border); border-radius: 6px;
        padding: 14px 4px; cursor: pointer; font: inherit; font-size: 13px; font-weight: 600;
      }
      .dash-cal-month-cell:hover { border-color: var(--navy); }
      .dash-cal-section { margin-top: 4px; }
      .dash-cal-section-title { font-size: 11px; font-weight: 700; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.03em; margin-bottom: 6px; }
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
        .dash-cal-cell-week .dash-cal-cell-items { display: none; }
        .dash-cal-cell-week { min-height: 52px; }
        .dash-stats-row { grid-template-columns: repeat(2, 1fr); }
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
        .page-actions, .form-actions { gap: 8px; flex-wrap: wrap; }
        .stat-value { font-size: 24px; }
        .sheet-input, select.sheet-input { padding: 10px 8px; }
        /* .spacer uses flex:1 to push trailing buttons (Mark paid, edit,
           delete, view toggles, etc.) to the right — fine on desktop, but
           combined with flex-wrap on a narrow screen it fragments whatever
           comes after it onto its own oddly-indented line instead of
           grouping together. Forcing every .spacer to break onto its own
           full-width, zero-height line means everything after it wraps
           together as one clean group on the next line instead of
           scattering — applies everywhere .spacer is used (list rows,
           filter/tab rows, etc.), not just one page.
        */
        .spacer { flex: 1 1 100%; height: 0; }
        .list-card-head, .filter-row { row-gap: 6px; }
        /* The sheet table's Notes preview column is sized for desktop
           readability (220-320px) — on a phone that alone eats more than
           half the screen, forcing heavy horizontal scroll just to see
           Balance or Status. Shrinking it to a short preview on mobile
           still shows enough to know there's a note, without dominating
           the layout; tapping it still opens the full text either way. */
        .sheet-col-notes { min-width: 110px; max-width: 140px; }
        .sheet-note-text { font-size: 11px; }
        /* Same idea for the phone cell — the cycle/add buttons add real
           width on top of the input itself, so give it less room to work
           with before the row needs to scroll. */
        .phone-cycle-cell .sheet-input { min-width: 60px; }
        .phone-cycle-btn { padding: 2px 3px; font-size: 9px; }
      }
      .save-error-banner {
        display: flex; align-items: center; gap: 8px; background: var(--danger-bg); color: var(--danger);
        padding: 10px 20px; font-size: 13px; border-bottom: 1px solid var(--danger);
      }
      .unsaved-alert-banner {
        display: flex; align-items: center; gap: 8px; background: var(--danger-bg); color: var(--danger);
        padding: 12px 16px; font-size: 14px; font-weight: 700; border: 1px solid var(--danger);
        border-radius: 8px; margin-bottom: 16px;
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
      .print-only { display: none; }
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

        /* Dashboard, Rent Collection, and Court Cases have their own
           purpose-built print layout below (.print-only) instead of the
           generic "hide the buttons and print whatever's on screen"
           treatment every other page still uses — the calendar, the
           collapsible panels, and the interactive sheet don't translate to
           paper cleanly no matter how they're restyled, so on these three
           pages specifically the normal screen content is hidden entirely
           and replaced with a plain, purpose-built table report instead. */
        .dashboard-page > :not(.print-only):not(.page-head),
        .rent-page > :not(.print-only):not(.page-head),
        .court-page > :not(.print-only):not(.page-head),
        .violations-page > :not(.print-only):not(.page-head) {
          display: none !important;
        }
        .print-only { display: block !important; }
        @page { margin: 0.6in 0.55in; }
        .print-header {
          display: flex; align-items: center; gap: 10px; margin-bottom: 4px;
          padding-bottom: 10px; border-bottom: 3px solid var(--brand-blue);
        }
        .print-mark {
          width: 30px; height: 30px; border-radius: 50%; background: var(--brand-blue); color: #fff;
          display: flex; align-items: center; justify-content: center; flex-shrink: 0;
          font-family: Georgia, "Times New Roman", serif; font-weight: 700; font-size: 15px;
        }
        .print-header-text { flex: 1; }
        .print-title { font-size: 20px; font-weight: 700; margin: 0; color: #000; line-height: 1.2; }
        .print-subtitle { font-size: 11px; color: #666; margin-top: 1px; }
        .print-stats-row { display: flex; gap: 12px; margin: 16px 0 20px; }
        .print-stat {
          flex: 1; border: 1px solid #ddd; border-top: 3px solid var(--brand-blue); border-radius: 4px;
          padding: 10px 14px; text-align: center; background: #fafafa;
        }
        .print-stat-num { font-size: 22px; font-weight: 700; color: #000; font-variant-numeric: tabular-nums; }
        .print-stat div:last-child { font-size: 10px; color: #666; text-transform: uppercase; letter-spacing: 0.03em; margin-top: 2px; }
        .print-section { margin-bottom: 20px; break-inside: avoid; }
        .print-section-head {
          display: flex; justify-content: space-between; align-items: baseline; font-weight: 700; font-size: 12.5px;
          color: #fff; background: var(--navy); padding: 6px 10px; border-radius: 3px 3px 0 0; margin-bottom: 0;
          -webkit-print-color-adjust: exact; print-color-adjust: exact;
        }
        .print-section-head span:last-child { font-weight: 500; font-size: 11px; opacity: 0.85; }
        .print-table { width: 100%; border-collapse: collapse; font-size: 11.5px; }
        .print-table th, .print-table td { border: 1px solid #ddd; padding: 6px 9px; text-align: left; }
        .print-table th {
          background: #eef2f8; color: #000; font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.02em;
          -webkit-print-color-adjust: exact; print-color-adjust: exact;
        }
        .print-table tbody tr:nth-child(even) td { background: #f7f8fa; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        .print-table-amount { text-align: right; font-variant-numeric: tabular-nums; font-weight: 600; }
        .print-grand-total {
          font-weight: 700; font-size: 13.5px; text-align: right; border-top: 3px solid var(--navy);
          padding-top: 10px; margin-top: 4px; color: #000;
        }
        .print-footer {
          position: fixed; bottom: 0.3in; left: 0.55in; right: 0.55in;
          display: flex; justify-content: space-between; font-size: 9px; color: #999;
          border-top: 1px solid #ddd; padding-top: 4px;
        }
      }
    `}</style>
  );
}
