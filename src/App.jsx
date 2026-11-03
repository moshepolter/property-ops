import React, { useState, useEffect, useMemo, useRef, useCallback, useLayoutEffect } from "react";
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
import { getFirestore, doc, setDoc, onSnapshot, runTransaction } from "firebase/firestore";
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
  Trash2, Pencil, Upload, Menu, Printer, CheckCircle2, Ticket, Siren, BellRing
} from "lucide-react";

/* ============================== constants ============================== */

const STORAGE_KEY = "pm-ops-data-v1";
// Shown at the bottom of the Dashboard so you can tell at a glance which copy of
// the app a given phone or browser tab is actually running.
const APP_BUILD = "sync-fix-34";
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

// What an edit form actually changed, compared with the record as it was when
// the form was opened. Saving only these — instead of writing the whole form
// back — means a form that has been open for a while can't put back fields
// that something else (another device, a note added meanwhile) changed in the
// meantime.
function changedFieldsOnly(original, edited) {
  if (!original) return edited;
  const patch = {};
  // Only fields the form actually has. A field the form doesn't manage (a case's
  // id, log, documents…) is simply not part of it — never treat that as "cleared".
  Object.keys(edited).forEach(k => {
    if (JSON.stringify(original[k]) !== JSON.stringify(edited[k])) patch[k] = edited[k];
  });
  return patch;
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
const OTHER_STATUSES = ["Open", "In Progress", "Certified", "Resolved", "Dismissed"];
const COURT_RESULTS = [
  "Pending", "Stipulation (payment plan)", "Case dismissed",
  "Adjourned / next date set", "Judgment for landlord", "Judgment for tenant",
  "Settled / withdrawn"
];
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
  customOtherAgencies: [], quickNotes: [], importHistory: [], caseStageOptions: [], emergencyRepairs: [], collections: [], smokeCoDetectors: [], unitTurns: [], insurancePolicies: [],
});

// Every field in emptyData() is an array of records with an id. When a
// remote update arrives on this device while a local edit is still
// pending save, blindly ignoring the remote update entirely would let this
// device's own save (a full document replace) later overwrite it — another
// tab's just-saved change would silently revert the moment this device's
// unrelated pending save fires. That's not hypothetical: with several tabs
// open on different devices, it's an everyday occurrence — check a box on
// one tab while a completely different tab has any pending edit of its
// own, and without this, that other tab's next save would silently
// resurrect the unchecked version, making the change look like it "came
// back" moments later.
//
// Blindly taking the remote version for everything would instead discard
// whatever's actively being edited locally right now. So this is
// record-aware: changedIds says which records THIS tab has actually
// touched since its last successful save (passed in from the ref that
// add/update/remove maintain). A record in that set keeps its local
// version, since this tab's own in-progress edit to it takes priority. A
// record NOT in that set — one this tab hasn't touched at all — safely
// takes the server's version, since there's nothing local to protect and
// the server reflects whatever another tab most recently did to it. A
// record that exists only remotely (added elsewhere, not here at all) is
// folded in either way.
// A few lists (custom appointment types, inspection types, other agencies) are
// just plain text — no ids — so none of the by-id logic below can apply to them.
const isPlainList = (arr) => Array.isArray(arr) && arr.some(x => x != null && typeof x !== "object");

function mergeRemoteWhilePending(localData, remoteData, changedIds) {
  const merged = { ...localData };
  Object.keys(emptyData()).forEach(k => {
    const local = Array.isArray(localData[k]) ? localData[k] : [];
    const remote = Array.isArray(remoteData[k]) ? remoteData[k] : [];
    // Plain-text lists only ever grow (options are added, never removed), so keep
    // everything either side has. Without this, the by-id matching below sees
    // text entries as having no id at all and silently throws the whole list away.
    if (isPlainList(local) || isPlainList(remote)) {
      const seen = new Set(local);
      merged[k] = [...local, ...remote.filter(x => !seen.has(x))];
      return;
    }
    const changed = (changedIds && changedIds[k]) || new Set();
    const remoteById = new Map(remote.filter(x => x && x.id).map(x => [x.id, x]));
    const localIds = new Set(local.map(x => x && x.id));
    const reconciled = local.map(x => {
      if (x && !changed.has(x.id) && remoteById.has(x.id)) return remoteById.get(x.id);
      return x;
    })
    // A local record that's untouched by this tab and no longer exists on
    // the server was very likely deleted by another tab — drop it here too,
    // rather than silently re-adding it the next time this tab saves. A
    // record this tab just created locally (pending, not saved yet) is
    // exempt, since it's correctly absent from the server for now.
    .filter(x => !x || changed.has(x.id) || remoteById.has(x.id));
    // Not one this tab just deleted and hasn't saved yet: the server still has
    // it, and re-adding it here would silently undo the delete (the next
    // save would then write it straight back).
    const remoteOnly = remote.filter(x => x && !localIds.has(x.id) && !changed.has(x.id));
    merged[k] = [...reconciled, ...remoteOnly];
  });
  return merged;
}

function violationClosedStatuses(agency) {
  return agency === "HPD" ? ["Certified", "Dismissed"] : agency === "DSNY" ? ["Paid"] : ["Certified", "Resolved", "Dismissed"];
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
// A date that lands on a Saturday or Sunday moves to the following Monday.
function nextWeekdayISO(d) {
  if (!d) return "";
  const dow = new Date(d + "T00:00:00").getDay();
  if (dow === 6) return addDays(d, 2);
  if (dow === 0) return addDays(d, 1);
  return d;
}
// The day after a payment is due: follow up with the attorney on whether the tenant paid.
function paymentFollowUpDate(due) { return due ? nextWeekdayISO(addDays(due, 1)) : ""; }
// Two weeks after the petition was served: has the tenant answered / start a default judgment.
function defaultJudgmentReminderDate(servedDate) { return servedDate ? nextWeekdayISO(addDays(servedDate, 14)) : ""; }
// Petition check-in date: 21 days after the lease/ledger went to the lawyer; if
// that lands on a Saturday or Sunday it moves to the following Monday.
function petitionReminderDate(sentDate) {
  if (!sentDate) return "";
  let d = addDays(sentDate, 21);
  const dow = new Date(d + "T00:00:00").getDay();
  if (dow === 6) d = addDays(d, 2);
  else if (dow === 0) d = addDays(d, 1);
  return d;
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
// Notes newest first: by date, then by the moment each was added, then by
// position in the list — so several notes added the same day always show the
// one added last on top, even for older notes saved without a timestamp.
function newestFirst(notes) {
  return (Array.isArray(notes) ? notes : [])
    .map((n, i) => ({ n, i }))
    .sort((a, b) => (b.n.date || "").localeCompare(a.n.date || "") || ((b.n.timestamp || 0) - (a.n.timestamp || 0)) || (b.i - a.i))
    .map(x => x.n);
}

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
// "Tue, Oct 14 · 9:00 AM" — a work order's scheduled visit. The year only shows
// when it isn't this year. Used on the work order row and in the Export backup.
function scheduleText(dateStr, timeStr = "") {
  if (!dateStr) return "";
  const d = new Date(dateStr + "T00:00:00");
  const opts = { weekday: "short", month: "short", day: "numeric" };
  if (d.getFullYear() !== new Date().getFullYear()) opts.year = "numeric";
  return d.toLocaleDateString("en-US", opts) + (timeStr ? ` · ${fmtTime(timeStr)}` : "");
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

// Quick "add the case number" box shown right on a court case that has none.
function CaseNumberQuickAdd({ onSave }) {
  const [open, setOpen] = useState(false);
  const [v, setV] = useState("");
  const stop = (e) => e.stopPropagation();
  const save = () => { const t = v.trim(); if (!t) return; onSave(t); setOpen(false); setV(""); };
  if (!open) return <button type="button" className="btn-ghost" onClick={(e) => { stop(e); setOpen(true); }}><Plus size={12} /> Case #</button>;
  return (
    <span style={{ display: "inline-flex", gap: 4, alignItems: "center" }} onClick={stop}>
      <input autoFocus placeholder="Case number" value={v} onChange={e => setV(e.target.value)} style={{ width: 150 }}
        onKeyDown={e => { if (e.key === "Enter") save(); if (e.key === "Escape") setOpen(false); }} />
      <button type="button" className="btn-primary" onClick={save} disabled={!v.trim()}>Save</button>
      <button type="button" className="btn-ghost" onClick={() => setOpen(false)}>Cancel</button>
    </span>
  );
}

// Stage buttons for a court case. Tap a button, pick the date, Save — that
// becomes "where it's up to" (the last one tapped) and sets its reminder.
// "Sent to lawyer" (lease/ledger) reminds 21 days later, "Petition served"
// 14 days later; buttons you add yourself remind after the number of days you
// give them. Weekend reminders move to Monday.
function CaseStageBar({ options, current, onApply, onAddOption, onRemoveOption, onClear }) {
  const [picking, setPicking] = useState(null); // "lawyer" | "petition" | option id
  const [date, setDate] = useState(todayISO());
  const [lease, setLease] = useState(true);
  const [ledger, setLedger] = useState(false);
  const [editing, setEditing] = useState(false);
  const [adding, setAdding] = useState(false);
  const [newLabel, setNewLabel] = useState("");
  const [newDays, setNewDays] = useState("");
  const custom = options || [];
  const open = (key) => { if (picking === key) { setPicking(null); return; } setPicking(key); setDate(todayISO()); setLease(true); setLedger(false); };
  const hint = () => {
    if (!date) return "";
    if (picking === "lawyer") return `Reminder on ${fmtDate(petitionReminderDate(date))} (21 days) to confirm the petition is being started.`;
    if (picking === "petition") return `Reminder on ${fmtDate(defaultJudgmentReminderDate(date))} (2 weeks) to follow up with the attorney / start a default judgment.`;
    const o = custom.find(x => x.id === picking);
    return o && Number(o.days) > 0 ? `Reminder on ${fmtDate(nextWeekdayISO(addDays(date, Number(o.days))))} (${o.days} days).` : "";
  };
  const save = () => {
    if (!date) return;
    if (picking === "lawyer") { if (!lease && !ledger) return; onApply({ kind: "lawyer", label: "Sent to lawyer", date, lease, ledger }); }
    else if (picking === "petition") onApply({ kind: "petition", label: "Petition served", date });
    else { const o = custom.find(x => x.id === picking); if (!o) return; onApply({ kind: "custom", label: o.label, date, days: Number(o.days) || 0 }); }
    setPicking(null);
  };
  const addButton = () => {
    const label = newLabel.trim();
    if (!label) return;
    onAddOption({ label, days: Number(newDays) > 0 ? Number(newDays) : 0 });
    setNewLabel(""); setNewDays(""); setAdding(false);
  };
  return (
    <div>
      <div className="row" style={{ marginBottom: 6 }}>
        <strong>Where it's up to:</strong>{" "}
        {current && current.label
          ? <span>{current.label}{current.date ? ` — ${fmtDate(current.date)}` : ""} {onClear && <button type="button" className="checklist-remove" title="Clear" onClick={onClear}><X size={12} /></button>}</span>
          : <span className="row-muted">not set</span>}
      </div>
      <div className="inline-form" style={{ flexWrap: "wrap", margin: 0 }}>
        <button type="button" className={`chip ${picking === "lawyer" || current?.label === "Sent to lawyer" ? "chip-active" : ""}`} onClick={() => open("lawyer")}>Sent to lawyer</button>
        <button type="button" className={`chip ${picking === "petition" || current?.label === "Petition served" ? "chip-active" : ""}`} onClick={() => open("petition")}>Petition served</button>
        {custom.map(o => (
          <span key={o.id} style={{ display: "inline-flex", alignItems: "center", gap: 2 }}>
            <button type="button" className={`chip ${picking === o.id || current?.label === o.label ? "chip-active" : ""}`} onClick={() => open(o.id)}>{o.label}{Number(o.days) > 0 ? ` · ${o.days}d` : ""}</button>
            {editing && (
              <button type="button" className="checklist-remove" title="Remove this button"
                onClick={() => { if (window.confirm(`Remove the "${o.label}" button? Cases already set to it keep it.`)) onRemoveOption(o); }}><X size={12} /></button>
            )}
          </span>
        ))}
        <button type="button" className="btn-ghost" onClick={() => setAdding(a => !a)}><Plus size={14} /> Add button</button>
        {custom.length > 0 && <button type="button" className="btn-ghost" onClick={() => setEditing(e => !e)}>{editing ? "Done" : "Edit buttons"}</button>}
      </div>
      {picking && (
        <div className="inline-form" style={{ flexWrap: "wrap", alignItems: "center", marginTop: 6 }}>
          {picking === "lawyer" && (
            <>
              <label style={{ display: "flex", gap: 4, alignItems: "center" }}><input type="checkbox" checked={lease} onChange={e => setLease(e.target.checked)} /> Lease</label>
              <label style={{ display: "flex", gap: 4, alignItems: "center" }}><input type="checkbox" checked={ledger} onChange={e => setLedger(e.target.checked)} /> Ledger</label>
            </>
          )}
          <input type="date" value={date} onChange={e => setDate(e.target.value)} style={{ flex: "0 0 auto", width: 170 }} />
          <button type="button" className="btn-primary" onClick={save} disabled={!date || (picking === "lawyer" && !lease && !ledger)}>Save</button>
          <button type="button" className="btn-ghost" onClick={() => setPicking(null)}>Cancel</button>
          {hint() && <div className="hint" style={{ flexBasis: "100%" }}>{hint()}</div>}
        </div>
      )}
      {adding && (
        <div className="inline-form" style={{ flexWrap: "wrap", alignItems: "center", marginTop: 6 }}>
          <input placeholder="Button name, e.g. Judgment entered" value={newLabel} onChange={e => setNewLabel(e.target.value)} onKeyDown={e => e.key === "Enter" && addButton()} />
          <input placeholder="Remind me after how many days (blank = none)" inputMode="numeric" value={newDays} onChange={e => setNewDays(e.target.value.replace(/\D/g, ""))} style={{ maxWidth: 260 }} />
          <button type="button" className="btn-primary" onClick={addButton} disabled={!newLabel.trim()}>Add</button>
          <button type="button" className="btn-ghost" onClick={() => setAdding(false)}>Cancel</button>
        </div>
      )}
    </div>
  );
}

// "Where it's up to" for a court case — a row of options YOU define (e.g. "14 day
// rent notice served"). Tap one to set it, tap it again to clear it. New options
// are typed in right here and are shared by every case. Removing an option never
// changes a case that already uses it.
function CaseStagePicker({ options, value, onPick, onAdd, onRemove }) {
  const [text, setText] = useState("");
  const [editing, setEditing] = useState(false);
  const list = options || [];
  const labels = list.map(o => o.label);
  const addNow = () => {
    const label = text.trim();
    if (!label) return;
    const existing = list.find(o => (o.label || "").toLowerCase() === label.toLowerCase());
    if (existing) onPick(existing.label); // already in your list — just select it
    else { onAdd(label); onPick(label); }
    setText("");
  };
  return (
    <div>
      <div className="inline-form" style={{ flexWrap: "wrap", margin: 0 }}>
        {list.map(o => (
          <span key={o.id} style={{ display: "inline-flex", alignItems: "center", gap: 2 }}>
            <button type="button" className={`chip ${value === o.label ? "chip-active" : ""}`} onClick={() => onPick(value === o.label ? "" : o.label)}>{o.label}</button>
            {editing && (
              <button type="button" className="checklist-remove" title="Remove this option"
                onClick={() => { if (window.confirm(`Remove "${o.label}" from your options? Cases already set to it keep it.`)) onRemove(o); }}>
                <X size={12} />
              </button>
            )}
          </span>
        ))}
        {value && !labels.includes(value) && (
          <button type="button" className="chip chip-active" title="Not in your list — tap to clear" onClick={() => onPick("")}>{value}</button>
        )}
        {list.length === 0 && !value && <span className="hint">No options yet — add your first one below.</span>}
      </div>
      <div className="inline-form" style={{ marginTop: 6 }}>
        <input placeholder="Add an option, e.g. 14 day rent notice served…" value={text} onChange={e => setText(e.target.value)} onKeyDown={e => e.key === "Enter" && addNow()} />
        <button type="button" className="btn-ghost" onClick={addNow}>Add</button>
        {list.length > 0 && <button type="button" className="btn-ghost" onClick={() => setEditing(!editing)}>{editing ? "Done" : "Edit options"}</button>}
      </div>
    </div>
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

function PropertyOpsAppInner() {
  const [user, setUser] = useState(undefined); // undefined = checking, null = signed out
  const [data, setDataRaw] = useState(emptyData());
  // Always mirrors the latest `data` — needed so a save that gets queued
  // while another is still in flight can read the truly current data when
  // it actually runs, not a stale value captured back when it was queued.
  const dataRef = useRef(data);
  useEffect(() => { dataRef.current = data; }, [data]);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("dashboard");
  // { tab, id } — set alongside setTab(...) when navigating to a specific
  // record, not just its general tab. The destination tab's own effect
  // consumes this (scrolling to and expanding the matching record) and
  // clears it immediately after, so it's a one-shot instruction rather
  // than something that keeps re-triggering.
  const [deepLink, setDeepLink] = useState(null);
  const goTo = (targetTab, id) => { setTab(targetTab); setDeepLink({ tab: targetTab, id }); };
  const [query, setQuery] = useState("");
  const [navOpen, setNavOpen] = useState(false);
  const [pinUnlocked, setPinUnlocked] = useState(() => !localStorage.getItem(PIN_STORAGE_KEY));
  const [showPinSetup, setShowPinSetup] = useState(false);
  const [saveError, setSaveError] = useState(false);
  // Why the last save failed (Firestore's own error code/message), so the red
  // notice can say something more useful than "check your connection".
  const [saveErrorDetail, setSaveErrorDetail] = useState("");
  // "idle" | "saving" | "saved" | "error" — drives the small save indicator that
  // shows on every tab, so you can see an edit has really reached the server
  // before you refresh or close the page.
  const [saveStatus, setSaveStatus] = useState("idle");
  const [loadError, setLoadError] = useState(false);
  // True once this copy has seen the data document being written by an OLDER
  // copy of the app (one without the revision counter) — i.e. another phone,
  // tablet or browser tab is still running the old version, which can overwrite
  // changes made here.
  const [olderVersionOpen, setOlderVersionOpen] = useState(false);
  const saveTimer = useRef(null);
  // Saves are debounced 400ms so typing doesn't hammer Firestore on every
  // keystroke — but that also means a change sits unsaved for up to 400ms
  // (plus however long the actual write takes). Closing the tab or hitting
  // back in that window would lose it silently with zero warning. This
  // tracks whether a save is still owed and, if so, asks the browser to
  // confirm before actually leaving.
  const hasPendingSave = useRef(false);
  // { collectionName: Set of record ids changed locally since the last
  // successful save } — lets a remote update arriving mid-save correctly
  // distinguish "a record this tab is actively changing, don't clobber it"
  // from "a record this tab hasn't touched, safe to take the server's
  // version of" when merging (see mergeRemoteWhilePending below).
  const pendingChangedIds = useRef({});
  const markChanged = (col, id) => {
    const set = pendingChangedIds.current[col] || new Set();
    set.add(id);
    pendingChangedIds.current[col] = set;
  };
  // Every actual user-initiated change — whether through add/update/remove
  // or a direct setData call anywhere else in the app (imports, bulk
  // deletes, the quick-note box, etc.) — goes through this wrapper rather
  // than the raw setter. It diffs old vs new state per collection right
  // here and marks every added, changed, or removed record id as pending,
  // automatically. That means the multi-tab protection this exists for
  // can never silently miss a call site the way manually adding
  // markChanged() to every individual place that touches data could —
  // there's nothing left to remember to instrument.
  const setData = (updater) => {
    setDataRaw(prev => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      Object.keys(emptyData()).forEach(col => {
        const prevArr = Array.isArray(prev[col]) ? prev[col] : [];
        const nextArr = Array.isArray(next[col]) ? next[col] : [];
        if (prevArr === nextArr) return; // untouched collection, nothing to diff
        // Plain-text lists have no ids to track, so a change to one is recorded
        // against the list as a whole — otherwise adding a custom option would
        // never count as an edit and would never be saved.
        if (isPlainList(prevArr) || isPlainList(nextArr)) {
          if (JSON.stringify(prevArr) !== JSON.stringify(nextArr)) markChanged(col, "_list");
          return;
        }
        const prevById = new Map(prevArr.filter(x => x && x.id).map(x => [x.id, x]));
        nextArr.forEach(x => {
          if (!x || !x.id) return;
          if (!prevById.has(x.id) || prevById.get(x.id) !== x) markChanged(col, x.id);
        });
        const nextIds = new Set(nextArr.filter(x => x && x.id).map(x => x.id));
        prevArr.forEach(x => { if (x && x.id && !nextIds.has(x.id)) markChanged(col, x.id); });
      });
      // Settings that aren't lists (the last-used appointment type) have no
      // ids to track, so record them by name — otherwise, now that saving only
      // happens for tracked edits, changing one would never be saved at all.
      const listKeys = Object.keys(emptyData());
      new Set([...Object.keys(prev), ...Object.keys(next)]).forEach(k => {
        if (!listKeys.includes(k) && prev[k] !== next[k]) markChanged("_fields", k);
      });
      return next;
    });
  };
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
  // The newest revision of the data document this device knows about. Every
  // save stamps the document with a revision number that goes up by one; an
  // update arriving with a LOWER number than this is a late arrival from before
  // something this device has already seen or saved, and is ignored.
  const knownRev = useRef(0);
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
    setDataRaw(emptyData());
    hasPendingSave.current = false;
    pendingSince.current = null;
    pendingChangedIds.current = {};
    setSaveStuck(false);
    justLoaded.current = true;
    knownRev.current = 0;
    setOlderVersionOpen(false);
    clearTimeout(saveRetryTimer.current);
    saveRetryCount.current = 0;
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
        // A late, older picture of the data must never be applied. Without this,
        // an update that was already on its way when this device saved could
        // arrive afterwards and roll the screen back to before the edit you just
        // made — and the next edit would then be built on the old version.
        if (snap.exists()) {
          const incomingRev = snap.data()._rev;
          if (incomingRev === undefined) {
            // No revision number at all: this write came from a copy of the app
            // that predates it. That's real, current data — apply it like any
            // other update, don't treat it as "older" and go deaf to it — and
            // say so, because that older copy can overwrite changes made here.
            if (knownRev.current > 0) setOlderVersionOpen(true);
          } else {
            if (incomingRev < knownRev.current) { setLoaded(true); return; }
            knownRev.current = incomingRev;
          }
        }
        // If this device has a local edit queued but not saved yet, don't
        // apply the remote update wholesale — that would discard whatever
        // is being edited here right now. But don't ignore it outright
        // either: this device's own save is a full document replace, and
        // would otherwise silently erase anything the other device just
        // saved. Fold in only the records that exist solely in the remote
        // version, leaving every local record exactly as it is.
        if (hasPendingSave.current) {
          if (snap.exists()) {
            const loaded = snap.data();
            setDataRaw(d => mergeRemoteWhilePending(d, loaded, pendingChangedIds.current));
          }
          setLoaded(true);
          return;
        }
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
          // The one saved setting that isn't a list: without this the last-used
          // appointment type was saved but never read back, so it reset to the
          // default every time the app was opened.
          if (loaded.lastUsedAppointmentType !== null && loaded.lastUsedAppointmentType !== undefined) merged.lastUsedAppointmentType = loaded.lastUsedAppointmentType;
          setDataRaw(merged);
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
          setDataRaw(emptyData());
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

  const saveRetryTimer = useRef(null);
  const saveRetryCount = useRef(0);
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
      // Save by folding THIS device's changes into whatever the server holds right
      // now, inside a transaction (the server re-runs it if someone else wrote in
      // between) — instead of replacing the server's copy with this device's view
      // of everything. A plain overwrite lets a device that hasn't yet heard about
      // another device's edit silently erase it. Here, only the records this device
      // actually changed win; everything else is whatever the server has.
      const changedAtStart = {};
      Object.entries(pendingChangedIds.current).forEach(([col, ids]) => { changedAtStart[col] = new Set(ids); });
      const dataRef_ = doc(db, "users", user.uid, "appData", "main");
      const savedRev = await runTransaction(db, async (tx) => {
        const snap = await tx.get(dataRef_);
        const remote = snap.exists() ? snap.data() : null;
        const merged = remote ? { ...remote, ...mergeRemoteWhilePending(safe, remote, changedAtStart) } : safe;
        // A setting that isn't a list: this device's value only wins if THIS
        // device just changed it; otherwise it's whatever the server has.
        if (remote) {
          const listKeys = Object.keys(emptyData());
          Object.keys(safe).forEach(k => {
            if (listKeys.includes(k)) return;
            merged[k] = ((changedAtStart._fields && changedAtStart._fields.has(k)) || !(k in remote)) ? safe[k] : remote[k];
          });
        }
        const nextRev = Math.max((remote && remote._rev) || 0, knownRev.current) + 1;
        tx.set(dataRef_, { ...JSON.parse(JSON.stringify(merged)), _rev: nextRev });
        return nextRev;
      });
      knownRev.current = Math.max(knownRev.current, savedRev);
      setSaveError(false);
      setSaveErrorDetail("");
      saveRetryCount.current = 0;
      clearTimeout(saveRetryTimer.current);
      // Only a confirmed success with nothing further owed means there's
      // truly nothing left pending — a failed save needs to keep the flag
      // (and the leave-page warning it drives) reflecting that. saveQueued
      // alone isn't a reliable signal here: it only catches a save being
      // directly re-attempted while one was in flight, not a new edit that
      // arrived via setData/markChanged and is still quietly waiting out
      // its own debounce timer. Comparing what was actually just written
      // to the current data catches both cases precisely — if they differ,
      // something changed locally since this save started, and that
      // change's own tracking must survive until it, too, is confirmed
      // saved, however it happens to get there.
      const unchangedSinceSave = JSON.stringify(dataRef.current) === JSON.stringify(safe);
      if (!saveQueued.current && unchangedSinceSave) {
        hasPendingSave.current = false; pendingSince.current = null; pendingChangedIds.current = {};
        setSaveStatus("saved");
        setTimeout(() => setSaveStatus(st => (st === "saved" ? "idle" : st)), 3000);
      }
    } catch (e) {
      console.error("save failed", e);
      setSaveError(true);
      setSaveErrorDetail(String((e && (e.code || e.message)) || e));
      setSaveStatus("error");
      // A failed save otherwise just sits there — nothing re-triggers
      // runSave unless the user happens to make another edit, which a
      // transient network blip gives no reason to expect. Retry on its own
      // instead, with backoff (2s, 4s, 8s... capped at 30s) so a real
      // outage doesn't hammer Firestore, but a brief hiccup recovers
      // without depending on the user noticing anything or acting on it.
      clearTimeout(saveRetryTimer.current);
      const delay = Math.min(30000, 2000 * Math.pow(2, saveRetryCount.current));
      saveRetryCount.current += 1;
      saveRetryTimer.current = setTimeout(() => { if (hasPendingSave.current) runSave(); }, delay);
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
    // Only a change made on THIS device needs saving. This effect also fires
    // when data merely ARRIVES from another device; writing that straight back
    // is a full-document write of whatever this device has at that instant,
    // and it can land after your next edit and overwrite it with an older
    // version — which is exactly how a note or work-order edit used to "revert".
    // Every real edit is recorded in pendingChangedIds the moment it's made.
    if (!Object.values(pendingChangedIds.current).some(s => s && s.size > 0)) return;
    if (!hasPendingSave.current) pendingSince.current = Date.now();
    hasPendingSave.current = true;
    setSaveStatus(st => (st === "error" ? st : "saving"));
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(runSave, 400);
    return () => clearTimeout(saveTimer.current);
  }, [data, loaded, user, runSave]);

  // On a phone, switching apps or locking the screen can freeze the page before
  // the short wait before saving is up, and the browser's "are you sure you want
  // to leave" warning never fires there. Save right away the moment the app is
  // hidden or being closed.
  useEffect(() => {
    const flushNow = () => {
      if (!hasPendingSave.current) return;
      clearTimeout(saveTimer.current);
      runSave();
    };
    const onVisibility = () => { if (document.visibilityState === "hidden") flushNow(); };
    document.addEventListener("visibilitychange", onVisibility);
    window.addEventListener("pagehide", flushNow);
    return () => {
      document.removeEventListener("visibilitychange", onVisibility);
      window.removeEventListener("pagehide", flushNow);
    };
  }, [runSave]);

  const prunedHistoryOnce = useRef(false);
  useEffect(() => {
    if (!loaded || !user || prunedHistoryOnce.current) return;
    prunedHistoryOnce.current = true;
    const byKey = {};
    (data.importHistory || []).forEach(h => {
      const key = `${h.buildingId}::${h.type}`;
      (byKey[key] = byKey[key] || []).push(h);
    });
    const keepIds = new Set();
    // Explicitly sorted by date, most recent first — not relying on the
    // array's stored order, so this is correct regardless of how these
    // entries happen to be arranged. Getting this backwards would mean
    // keeping old records and discarding recent ones, which is exactly
    // the opposite of what this is supposed to do.
    Object.values(byKey).forEach(entries => {
      entries.slice().sort((a, b) => (b.date || "").localeCompare(a.date || "")).slice(0, 10).forEach(h => keepIds.add(h.id));
    });
    const needsPruning = (data.importHistory || []).length > keepIds.size;
    if (needsPruning) {
      setData(d => ({ ...d, importHistory: (d.importHistory || []).filter(h => keepIds.has(h.id)) }));
    }
  }, [loaded, user]);

  const trimmedCertifiedHpdOnce = useRef(false);
  useEffect(() => {
    if (!loaded || !user || trimmedCertifiedHpdOnce.current) return;
    trimmedCertifiedHpdOnce.current = true;
    const needsTrim = (v) => v.agency === "HPD" && v.status === "Certified" && Object.keys(v).length > 6;
    if ((data.violations || []).some(needsTrim)) {
      setData(d => ({
        ...d,
        violations: (d.violations || []).map(v => needsTrim(v)
          ? { id: v.id, agency: "HPD", buildingId: v.buildingId, unitId: v.unitId, violationNumber: v.violationNumber, status: "Certified" }
          : v),
      }));
    }
  }, [loaded, user]);

  // One-time: merge duplicate HPD violations (same violation number) into one.
  // Keeps the fullest record, carries over notes, drops the rest.
  const dedupedHpdOnce = useRef(false);
  useEffect(() => {
    if (!loaded || !user || dedupedHpdOnce.current) return;
    dedupedHpdOnce.current = true;
    const groups = new Map();
    (data.violations || []).forEach(v => {
      if (v.agency !== "HPD" || !v.violationNumber) return;
      const k = String(v.violationNumber);
      groups.set(k, [...(groups.get(k) || []), v]);
    });
    const drop = new Set(), replace = new Map();
    groups.forEach(list => {
      if (list.length < 2) return;
      const keep = list.slice().sort((a, b) => (Array.isArray(b.notes) ? b.notes.length : 0) - (Array.isArray(a.notes) ? a.notes.length : 0) || Object.keys(b).length - Object.keys(a).length)[0];
      const notes = [];
      list.forEach(v => (Array.isArray(v.notes) ? v.notes : []).forEach(n => { if (!notes.some(x => JSON.stringify(x) === JSON.stringify(n))) notes.push(n); }));
      list.forEach(v => { if (v.id !== keep.id) drop.add(v.id); });
      if (Array.isArray(keep.notes) && notes.length !== keep.notes.length) replace.set(keep.id, { ...keep, notes });
    });
    if (drop.size || replace.size) {
      setData(d => ({ ...d, violations: (d.violations || []).filter(v => !drop.has(v.id)).map(v => replace.get(v.id) || v) }));
    }
  }, [loaded, user]);

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
  const add = (col, item) => {
    const id = item.id || uid();
    markChanged(col, id);
    setData(d => ({ ...d, [col]: [...d[col], { ...item, id }] }));
    return id;
  };
  const update = (col, id, patch) => {
    markChanged(col, id);
    setData(d => ({
      ...d,
      [col]: d[col].map(x => {
        if (x.id !== id) return x;
        const merged = { ...x, ...patch };
        // HPD only, and only once it actually becomes Certified — every
        // other agency (DOB, FDNY, DSNY, ECB, DEP, Con Edison) keeps its
        // full record forever, even after marked done. Keeping the
        // violation number, building, unit, and agency means a future HPD
        // re-import still correctly recognizes this one as already on
        // file and leaves it alone, and it still shows on the page.
        if (col === "violations" && merged.agency === "HPD" && merged.status === "Certified" && x.status !== "Certified") {
          return { id: x.id, agency: "HPD", buildingId: x.buildingId, unitId: x.unitId, violationNumber: x.violationNumber, status: "Certified" };
        }
        return merged;
      }),
    }));
  };
  const remove = (col, id) => { markChanged(col, id); setData(d => ({ ...d, [col]: d[col].filter(x => x.id !== id) })); };

  const buildingName = (id) => shortAddress(data.buildings.find(b => b.id === id)?.address) || "—";
  const vendorName = (id) => data.vendors.find(v => v.id === id)?.name || "—";
  const tenantName = (id) => data.tenants.find(t => t.id === id)?.name || "—";

  const tabs = [
    { key: "dashboard", label: "Dashboard", icon: <LayoutDashboard size={16} /> },
    { key: "buildings", label: "Buildings", icon: <Building2 size={16} /> },
    { key: "rent", label: "Rent Collection", icon: <Users size={16} /> },
    { key: "workorders", label: "Work Orders", icon: <Wrench size={16} /> },
    { key: "violations", label: "Violations", icon: <AlertTriangle size={16} /> },
    { key: "ero", label: "Emergency Repairs", icon: <Siren size={16} /> },
    { key: "vendors", label: "Vendors", icon: <HardHat size={16} /> },
    { key: "court", label: "Court Cases", icon: <Gavel size={16} /> },
    { key: "collections", label: "Collections", icon: <DollarSign size={16} /> },
    { key: "inspections", label: "Appointments", icon: <CalendarClock size={16} /> },
    { key: "smokeco", label: "Smoke / CO", icon: <BellRing size={16} /> },
    { key: "unitturn", label: "Unit Turn", icon: <Home size={16} /> },
    { key: "insurance", label: "Insurance", icon: <ScrollText size={16} /> },
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
          <AlertTriangle size={16} /> Couldn't save your last change — check your internet connection. Your edits are still here on screen, but won't be there if you close the tab until this clears.{saveErrorDetail ? ` (Reason: ${saveErrorDetail})` : ""}
        </div>
      )}

      {saveStatus !== "idle" && (
        <div className="no-print" style={{ position: "fixed", left: 10, bottom: "calc(env(safe-area-inset-bottom, 0px) + 10px)", zIndex: 60, padding: "6px 12px", borderRadius: 999, fontSize: 12, fontWeight: 600, color: "#fff", pointerEvents: "none",
          background: saveStatus === "saved" ? "rgba(22,128,60,.92)" : saveStatus === "error" ? "rgba(185,28,28,.95)" : "rgba(30,30,30,.9)" }}>
          {saveStatus === "saving" ? "Saving… don't refresh yet" : saveStatus === "saved" ? "✓ Saved" : "⚠ Not saved yet"}
        </div>
      )}

      {olderVersionOpen && (
        <div className="save-error-banner no-print">
          <AlertTriangle size={16} /> Another phone, tablet, or browser tab still has an OLD version of this app open — it can overwrite changes you make here. Close it, or fully quit and reopen the app there, then reload this page.
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
            {tab === "dashboard" && <Dashboard data={data} buildingName={buildingName} tenantName={tenantName} setTab={setTab} goTo={goTo} setData={setData} update={update} saveStuck={saveStuck} />}
            {tab === "buildings" && <BuildingsTab data={data} add={add} update={update} remove={remove} setData={setData} buildingName={buildingName} />}
            {tab === "rent" && <RentTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} setData={setData} deepLink={deepLink} setDeepLink={setDeepLink} />}
            {tab === "workorders" && <WorkOrdersTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} vendorName={vendorName} tenantName={tenantName} deepLink={deepLink} setDeepLink={setDeepLink} />}
            {tab === "violations" && <ViolationsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} vendorName={vendorName} setData={setData} deepLink={deepLink} setDeepLink={setDeepLink} />}
            {tab === "ero" && <EroTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} vendorName={vendorName} deepLink={deepLink} setDeepLink={setDeepLink} />}
            {tab === "vendors" && <VendorsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} />}
            {tab === "court" && <CourtTab data={data} add={add} update={update} remove={remove} setData={setData} tenantName={tenantName} buildingName={buildingName} deepLink={deepLink} setDeepLink={setDeepLink} />}
            {tab === "collections" && <CollectionsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} deepLink={deepLink} setDeepLink={setDeepLink} />}
            {tab === "inspections" && <AppointmentsTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} setData={setData} deepLink={deepLink} setDeepLink={setDeepLink} />}
            {tab === "smokeco" && <SmokeCoTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} />}
            {tab === "unitturn" && <UnitTurnTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} />}
            {tab === "insurance" && <InsuranceTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} />}
            {tab === "quicknotes" && <QuickNotesTab data={data} add={add} update={update} remove={remove} buildingName={buildingName} deepLink={deepLink} setDeepLink={setDeepLink} />}
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
  ero: "Emergency Repairs", collections: "Collections", inspections: "Appointments", quicknotes: "Quick Notes / Reminder", dashboard: "Dashboard", insurance: "Insurance",
};

function AttentionPanel({ icon, label, items, tab, setTab, renderItem, itemKey, extraAction, limit = 8 }) {
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
          {items.slice(0, limit).map((it, i) => (
            <div className="followup-item" key={itemKey ? itemKey(it) : i}>{renderItem(it)}</div>
          ))}
          {items.length > limit && <div className="hint" style={{ margin: "4px 0" }}>+ {items.length - limit} more</div>}
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
    if (v.cureDeadline && !isViolationClosed(v) && !v.isLead) {
      items.push({ key: `v-${v.id}`, date: v.cureDeadline, type: `${v.agency} cure deadline`, label: `#${v.violationNumber}`, sub: buildingName(v.buildingId), tab: "violations" });
    }
    if (v.hasHearing && v.hearingDate && !isViolationClosed(v)) {
      items.push({ key: `vh-${v.id}`, date: v.hearingDate, type: `${v.agency} hearing`, label: `#${v.violationNumber}`, sub: buildingName(v.buildingId), tab: "violations" });
    }
  });
  data.courtCases.forEach(c => {
    if (c.archived) return;
    if (c.nextCourtDate) items.push({ key: `c-${c.id}`, date: c.nextCourtDate, type: "Court date", label: tenantName(c.tenantId), sub: buildingName(c.buildingId), tab: "court" });
    // Notes dated in the future (e.g. "Eviction" on Oct 28) are scheduled items too.
    // Skip one that just repeats the case's own next court date.
    (c.log || []).forEach(l => {
      if (!l || !l.date || l.date <= todayISO() || l.date === c.nextCourtDate) return;
      items.push({ key: `cl-${c.id}-${l.id}`, targetId: c.id, date: l.date, type: "Court note", label: tenantName(c.tenantId), note: l.note, sub: buildingName(c.buildingId), tab: "court" });
    });
    if (c.result === "Stipulation (payment plan)" && c.nextPaymentDue) items.push({ key: `p-${c.id}`, date: c.nextPaymentDue, type: "Payment due", label: tenantName(c.tenantId), sub: buildingName(c.buildingId), tab: "court" });
    // The day after a payment is due: follow up with the attorney on whether the tenant paid.
    if (c.result === "Stipulation (payment plan)" && c.nextPaymentDue && c.payFollowDismissed !== c.nextPaymentDue) {
      items.push({ key: `pf-${c.id}`, targetId: c.id, date: paymentFollowUpDate(c.nextPaymentDue), type: "Payment follow-up", label: tenantName(c.tenantId), note: `Follow up with the attorney — did the tenant pay (due ${fmtDate(c.nextPaymentDue)})?`, sub: buildingName(c.buildingId), tab: "court" });
    }
    // Kept after the due date itself was checked off the dashboard.
    if (c.payFollowUpDate) {
      items.push({ key: `pfs-${c.id}`, targetId: c.id, date: c.payFollowUpDate, type: "Payment follow-up", label: tenantName(c.tenantId), note: "Follow up with the attorney — did the tenant pay?", sub: buildingName(c.buildingId), tab: "court" });
    }
  });
  (data.emergencyRepairs || []).forEach(e => {
    if (e.status !== "Completed" && e.dueByDate) {
      items.push({ key: `ero-${e.id}`, targetId: e.id, date: e.dueByDate, type: "ERO due", label: `ERO ${e.eroNumber || ""}`.trim(), note: e.description, sub: buildingName(e.buildingId), tab: "ero" });
    }
  });
  data.workOrders.forEach(w => {
    if (w.courtCaseId && w.dueByDate && w.status !== "Done") {
      items.push({ key: `wo-${w.id}`, date: w.dueByDate, type: "Work order due", label: w.description, sub: buildingName(w.buildingId), tab: "workorders" });
    }
  });
  data.appointments.forEach(a => {
    if (!a.completed && a.date) {
      items.push({ key: `a-${a.id}`, date: a.date, type: a.recurring ? "Recurring appointment" : "Appointment", label: a.type, sub: buildingName(a.buildingId), tab: "inspections" });
    }
  });
  data.tenants.forEach(t => {
    tenantFollowUps(t).forEach(f => {
      if (f.date) items.push({ key: `f-${f.id}`, targetId: t.id, date: f.date, type: "Follow-up", label: t.name, sub: buildingName(t.buildingId), note: f.note, tab: "rent" });
    });
  });
  (data.quickNotes || []).forEach(n => {
    if (n.reminderDate && !n.done) {
      items.push({ key: `n-${n.id}`, date: n.reminderDate, type: "Reminder", label: n.text, sub: n.buildingId ? buildingName(n.buildingId) : "", tab: "quicknotes" });
    }
  });
  // Insurance renewals: 2 months before expiry (until the three steps are done),
  // 1 month before (until chosen & bound), and the expiry date itself.
  (data.insurancePolicies || []).forEach(p => {
    if (!p.expirationDate) return;
    const bn = (Array.isArray(p.buildingIds) ? p.buildingIds : []).map(id => buildingName(id)).join(", ") || p.insured || "";
    const lbl = `${bn ? `${bn} — ` : ""}${p.type || "Insurance"}${p.carrier ? ` (${p.carrier})` : ""}`;
    const rem = insuranceRemaining(p);
    if (rem.length) items.push({ key: `ins-${p.id}`, targetId: p.id, date: addMonthsISO(p.expirationDate, -2), type: "Insurance renewal", label: lbl, note: `Expires ${fmtDate(p.expirationDate)} — to do: ${rem.map(r => r.todo).join("; ")}`, tab: "insurance" });
    if (!p.stepBound) {
      items.push({ key: `insd-${p.id}`, targetId: p.id, date: addMonthsISO(p.expirationDate, -1), type: "Insurance renewal", label: lbl, note: `Expires ${fmtDate(p.expirationDate)} — compare coverage and price, choose, and bind the renewal`, tab: "insurance" });
      items.push({ key: `inse-${p.id}`, targetId: p.id, date: p.expirationDate, type: "Insurance renewal", label: lbl, note: "Policy expires today — confirm the renewal is bound, then press Renewed", tab: "insurance" });
    }
  });
  // Anything scheduled for a Saturday gets a second reminder on the Monday after.
  const weekend = [];
  items.forEach(i => {
    if (!i.date) return;
    const dt0 = new Date(i.date + "T00:00:00");
    if (isNaN(dt0) || dt0.getDay() !== 6) return;
    weekend.push({
      ...i, key: `sat-${i.key}`, targetId: i.targetId || i.key.slice(i.key.indexOf("-") + 1),
      date: addDays(i.date, 2), type: "Weekend reminder",
      note: `Was scheduled for Saturday ${fmtDate(i.date)} — ${i.type}${i.note ? `: ${i.note}` : ""}`,
    });
  });
  return items.concat(weekend);
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

function DashboardCalendar({ data, buildingName, tenantName, setTab, goTo }) {
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
    <button key={item.key} className="dash-detail-item" onClick={() => goTo(item.tab, item.targetId || item.key.slice(item.key.indexOf("-") + 1))}>
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

function Dashboard({ data: rawData, buildingName, tenantName, setTab, goTo, setData, update, saveStuck }) {
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
    workOrders: rawData.workOrders.filter(w => mainBuildingIds.has(w.buildingId)),
    // Court cases and violations are never excluded by building owner —
    // for those two, every building is treated as a regular one, since
    // both are actively managed regardless of which owner's building it is.
    courtCases: rawData.courtCases,
    violations: rawData.violations,
    appointments: rawData.appointments.filter(a => mainBuildingIds.has(a.buildingId)),
    localLaws: (rawData.localLaws || []).filter(l => mainBuildingIds.has(l.buildingId)),
  };
  const [testEmailStatus, setTestEmailStatus] = useState(null); // null | "sending" | "sent" | "error"
  const [expandedBuilding, setExpandedBuilding] = useState(null);
  const [statOpen, setStatOpen] = useState(null);
  const [showAllFollowUps, setShowAllFollowUps] = useState(false);
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
  const collectionItems = (data.collections || []).filter(c => !c.archived).sort((a, b) => parseBalance(b.amountOwed) - parseBalance(a.amountOwed));
  const collectionsTotal = collectionItems.reduce((sum, c) => sum + parseBalance(c.amountOwed), 0);
  const eroItems = (data.emergencyRepairs || []).filter(e => e.status !== "Completed").sort((a, b) => (a.dueByDate || "9999").localeCompare(b.dueByDate || "9999"));
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

  const allDated = allDatedItems(data, tenantName, buildingName);
  // Saturday items reminded on the Monday itself, and payment follow-ups with the attorney.
  const weekendItems = allDated.filter(i => i.type === "Weekend reminder" && i.date === today);
  const insuranceItems = allDated.filter(i => i.type === "Insurance renewal" && i.date <= today).sort((a, b) => a.date.localeCompare(b.date));
  const payFollowItems = allDated.filter(i => i.type === "Payment follow-up" && i.date <= today).sort((a, b) => a.date.localeCompare(b.date));
  const rentPanelCount = overdueTenants.length + allFollowUps;
  const totalViolationsDue = violationAgencyGroups.reduce((sum, g) => sum + g.items.length, 0);
  const totalAttention = rentPanelCount + recurringItems.length + appointmentItems.length + quickNoteItems.length + totalViolationsDue + hearingItems.length + eroItems.length + collectionItems.length + weekendItems.length + payFollowItems.length + insuranceItems.length;

  // Top stat row + follow-up roster
  const overdueCount = allDated.filter(i => i.date < today && i.type !== "Follow-up" && i.type !== "Weekend reminder" && i.type !== "Insurance renewal").length;
  // Each item type in the Overdue list is a different kind of record with
  // its own notion of "done" — a violation's cure deadline is cleared by
  // marking it certified/resolved, a hearing by clearing the hearing flag,
  // a court date or stipulation payment by clearing that date (the next
  // one gets set once it's actually scheduled), an appointment or a quick
  // note reminder by their own existing completed/done field. The item's
  // key prefix (set in allDatedItems) says which kind it is.
  const completeDatedItem = (item) => {
    const id = item.key.slice(item.key.indexOf("-") + 1);
    if (item.key.startsWith("v-")) {
      const v = rawData.violations.find(x => x.id === id);
      if (v) update("violations", id, { status: violationClosedStatuses(v.agency)[0] });
    } else if (item.key.startsWith("vh-")) {
      update("violations", id, { hasHearing: false });
    } else if (item.key.startsWith("c-")) {
      update("courtCases", id, { nextCourtDate: "" });
    } else if (item.key.startsWith("pfs-")) {
      update("courtCases", id, { payFollowUpDate: "" });
    } else if (item.key.startsWith("pf-")) {
      const cc = rawData.courtCases.find(x => x.id === id);
      if (cc) update("courtCases", id, { payFollowDismissed: cc.nextPaymentDue || "" });
    } else if (item.key.startsWith("p-")) {
      // Clearing the due date keeps the "follow up with the attorney" reminder alive.
      const cc = rawData.courtCases.find(x => x.id === id);
      const keep = cc && cc.nextPaymentDue && cc.payFollowDismissed !== cc.nextPaymentDue ? paymentFollowUpDate(cc.nextPaymentDue) : "";
      update("courtCases", id, { nextPaymentDue: "", ...(keep ? { payFollowUpDate: keep } : {}) });
    } else if (item.key.startsWith("a-")) {
      update("appointments", id, { completed: true });
    } else if (item.key.startsWith("n-")) {
      update("quickNotes", id, { done: true });
    } else if (item.key.startsWith("wo-")) {
      update("workOrders", id, { status: "Done" });
    } else if (item.key.startsWith("ero-")) {
      update("emergencyRepairs", id, { status: "Completed", completedDate: todayISO() });
    }
  };
  const openWorkOrders = data.workOrders.filter(w => w.status !== "Done");
  const openCourtCases = data.courtCases.filter(c => !c.archived);
  const clearBuildingIds = new Set(data.buildings.filter(b => {
    const hasViolation = data.violations.some(v => v.buildingId === b.id && !isViolationClosed(v) && !v.isLead);
    const hasLateTenant = data.tenants.some(t => t.buildingId === b.id && t.status !== "Current");
    const hasOpenWO = data.workOrders.some(w => w.buildingId === b.id && w.status !== "Done");
    const hasOpenCourt = data.courtCases.some(c => c.buildingId === b.id && !c.archived);
    return !hasViolation && !hasLateTenant && !hasOpenWO && !hasOpenCourt;
  }).map(b => b.id));
  const rosterEntries = showAllFollowUps ? followUpEntries : followUpEntries.slice(0, 8);

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
      "Moved In": t.moveInDate ? fmtDate(t.moveInDate) : "",
      "Lease Ends": t.leaseEnd ? fmtDate(t.leaseEnd) : "",
      "Moved Out": t.movedOut ? "Yes" : "",
      "Other People in Apt": (Array.isArray(t.occupants) ? t.occupants : []).map(o => `${o.name || ""}${o.relationship ? ` (${o.relationship})` : ""}${o.phone ? ` ${o.phone}` : ""}`.trim()).join(" | "),
      Notes: (Array.isArray(t.notes) ? t.notes : []).map(n => `${fmtDate(n.date)}: ${n.text}`).join(" | "),
      "Payment History": (Array.isArray(t.payments) ? t.payments : []).slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).map(p => `${fmtDate(p.date)}: $${p.amount}${p.note ? ` (${p.note})` : ""}`).join(" | "),
      "Follow-ups": tenantFollowUps(t).map(f => `${fmtDate(f.date)}${f.note ? `: ${f.note}` : ""}`).join(" | "),
    })));

    // The Rent Collection page as it appears on screen: one row per tenant, in
    // building / unit order, with the aging columns, flags and a grand total.
    const rentRows = data.tenants
      .slice()
      .sort((a, b) => {
        const bi = data.buildings.findIndex(x => x.id === a.buildingId) - data.buildings.findIndex(x => x.id === b.buildingId);
        return bi || compareUnits(unitMap[a.unitId] || "", unitMap[b.unitId] || "");
      })
      .map(t => ({
        Building: buildingMap[t.buildingId] || "",
        Unit: unitMap[t.unitId] || "",
        Tenant: t.name || "",
        Phone: t.phone || "",
        Status: t.status || "",
        "0-30 Days": t.aging ? t.aging.bucket1 : "",
        "31-60 Days": t.aging ? t.aging.bucket2 : "",
        "61+ Days": t.aging ? t.aging.bucket61 : "",
        "Total Owed": parseBalance(t.balance),
        "In Court": (data.courtCases || []).some(c => c.tenantId === t.id && !c.archived) ? "Yes" : "",
        "Call Back": t.callBack ? "Yes" : "",
        "Voucher Program": t.voucherProgram ? "Yes" : "",
        Notes: (Array.isArray(t.notes) ? t.notes : []).map(n => `${fmtDate(n.date)}: ${n.text}`).join(" | "),
        "Follow-ups": tenantFollowUps(t).map(f => `${fmtDate(f.date)}${f.note ? `: ${f.note}` : ""}`).join(" | "),
        "Payment History": (Array.isArray(t.payments) ? t.payments : []).slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).map(pm => `${fmtDate(pm.date)}: $${pm.amount}${pm.note ? ` (${pm.note})` : ""}`).join(" | "),
      }));
    rentRows.push({ Building: "TOTAL OWED", "Total Owed": rentRows.reduce((sum, r) => sum + (Number(r["Total Owed"]) || 0), 0) });
    addSheet("Rent Collection", rentRows);

    addSheet("Violations", data.violations.map(v => ({
      Building: buildingMap[v.buildingId] || "",
      Unit: unitMap[v.unitId] || "",
      Agency: v.agency === "Other" ? (v.otherAgency || "Other") : v.agency,
      "Violation #": v.violationNumber || "",
      Class: v.class || "",
      Description: v.description || "",
      "Cure Deadline": v.cureDeadline ? fmtDate(v.cureDeadline) : "",
      "Fine Amount": v.fineAmount || "",
      Hearing: v.hasHearing ? (v.hearingDate ? fmtDate(v.hearingDate) : "Yes (no date)") : "",
      Company: v.company || "",
      Status: v.status || "",
      Vendor: vendorMap[v.vendorId] || "",
      Notes: (Array.isArray(v.notes) ? v.notes : []).map(n => `${fmtDate(n.date)}: ${n.text}`).join(" | "),
    })));

    addSheet("Work Orders", data.workOrders.map(w => ({
      Building: buildingMap[w.buildingId] || "",
      Unit: unitMap[w.unitId] || "",
      Description: (w.items && w.items.length > 0) ? w.items.map(it => it.description).join(" | ") : (w.description || ""),
      Priority: w.priority || "",
      Status: w.status || "",
      Vendor: (w.items && w.items.length > 0)
        ? w.items.map(it => vendorMap[it.vendorId]).filter(Boolean).join(", ")
        : (vendorMap[w.vendorId] || ""),
      "Date Opened": w.dateOpened ? fmtDate(w.dateOpened) : "",
      "Court Linked": w.isCourtConnected || w.courtCaseId ? "Yes" : "No",
      "Due By": w.dueByDate ? fmtDate(w.dueByDate) : "",
      "Scheduled": scheduleText(w.scheduledDate, w.scheduledTime),
      "Access Date": w.accessDate ? fmtDate(w.accessDate) : "",
      "Access Granted": w.accessGranted === true ? "Yes" : w.accessGranted === false ? "No" : "",
      Notes: (w.notes || []).slice().sort((a, b) => (a.timestamp || 0) - (b.timestamp || 0)).map(n => `${fmtDate(n.date)}: ${n.text}`).join(" | "),
    })));

    addSheet("Court Cases", data.courtCases.map(c => ({
      Building: buildingMap[c.buildingId] || "",
      Unit: unitMap[c.unitId] || "",
      Tenant: tenantMap[c.tenantId] || c.rawName || "",
      "Case #": c.caseNumber || "",
      Stage: c.stage || "",
      "Court Date": c.nextCourtDate ? fmtDate(c.nextCourtDate) : "",
      Result: c.result || "",
      "Stipulation Terms": c.stipulationTerms || "",
      "Next Payment Due": c.nextPaymentDue ? fmtDate(c.nextPaymentDue) : "",
      Status: c.archived ? "Closed" : "Active",
      "Sent to Lawyer": [c.sentLease && "Lease", c.sentLedger && "Ledger"].filter(Boolean).join(" + "),
      "Sent to Lawyer Date": c.sentToLawyerDate ? fmtDate(c.sentToLawyerDate) : "",
      "Stage Date": c.stageDate ? fmtDate(c.stageDate) : "",
      "Petition Served": c.petitionServed ? (c.petitionServedDate ? fmtDate(c.petitionServedDate) : "Yes") : "",
      "Owes (tenant balance)": c.tenantId ? ((data.tenants.find(t => t.id === c.tenantId) || {}).balance || "") : "",
      Log: (c.log || []).slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).map(l => `${fmtDate(l.date)}: ${l.note}`).join(" | "),
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

    addSheet("Emergency Repairs", (data.emergencyRepairs || []).map(e => ({
      "ERO #": e.eroNumber || "",
      Building: buildingMap[e.buildingId] || "",
      Unit: unitMap[e.unitId] || "",
      "Connected Violations": (e.violationIds || []).map(id => (data.violations.find(v => v.id === id) || {}).violationNumber).filter(Boolean).join(", "),
      Work: e.description || "",
      "Handled By": e.vendorId ? (vendorMap[e.vendorId] || "") : (e.handledBy || ""),
      Issued: e.dateIssued ? fmtDate(e.dateIssued) : "",
      "Due By": e.dueByDate ? fmtDate(e.dueByDate) : "",
      "Reached Out to HPD": e.hpdContacted ? (e.hpdContactedDate ? `Yes (${fmtDate(e.hpdContactedDate)})` : "Yes") : "No",
      Status: e.status || "",
      Completed: e.completedDate ? fmtDate(e.completedDate) : "",
      Notes: (e.notes || []).slice().sort((a, b) => (a.timestamp || 0) - (b.timestamp || 0)).map(n => `${fmtDate(n.date)}: ${n.text}`).join(" | "),
    })));

    addSheet("Collections", (data.collections || []).map(c => ({
      Tenant: c.tenantId ? (tenantMap[c.tenantId] || c.tenantName || "") : (c.tenantName || ""),
      Building: buildingMap[c.buildingId] || "",
      Unit: unitMap[c.unitId] || "",
      "Amount Owed": c.amountOwed || "",
      "Moved Out": c.movedOutDate ? fmtDate(c.movedOutDate) : "",
      Lawyer: c.lawyer || "",
      "Where It's Up To": c.stage || "",
      Status: c.archived ? "Closed" : "Active",
      Log: (c.log || []).slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).map(l => `${fmtDate(l.date)}: ${l.note}`).join(" | "),
    })));

    addSheet("Insurance", (data.insurancePolicies || []).map(p => ({
      Policy: p.type || "", Insured: p.insured || "", Buildings: (Array.isArray(p.buildingIds) ? p.buildingIds : []).map(id => buildingMap[id] || "").filter(Boolean).join(", ") || p.covers || "", Starts: p.effectiveDate ? fmtDate(p.effectiveDate) : "", Broker: p.broker || "", Carrier: p.carrier || "", "Policy #": p.policyNumber || "", Premium: p.premium || "",
      Expires: p.expirationDate ? fmtDate(p.expirationDate) : "",
      "Quotes requested": p.stepQuotes ? fmtDate(p.stepQuotes) : "", "Dec pages & loss runs": p.stepDecs ? fmtDate(p.stepDecs) : "",
      "Sent to other brokers": p.stepOther ? fmtDate(p.stepOther) : "", "Chosen & bound": p.stepBound ? fmtDate(p.stepBound) : "", Notes: p.notes || "",
    })));
    addSheet("Unit Turns", (data.unitTurns || []).map(t => ({
      Building: buildingMap[t.buildingId] || "", Apt: unitMap[t.unitId] || "", "Previous Tenant": t.prevTenant || "",
      "Occupancy Ended": t.endedDate ? fmtDate(t.endedDate) : "", "Inspection Date": t.inspectionDate ? fmtDate(t.inspectionDate) : "",
      "Inspected By": t.inspectedBy === "third" ? "Third party" : "Owner/agent", Inspector: t.inspectorName || "",
      "Exempt": t.exempt === "hpd" ? "HPD Lead Free exemption" : t.exempt === "tested" ? "Tested, no lead paint" : "No",
      "Lead Hazards Found": t.hazardsFound === "yes" ? "Yes" : t.hazardsFound === "no" ? "No" : "",
      "Checklist Done": `${TURN_CHECKLIST_FLAT.filter(i => t.checklist?.[i.key]?.done).length}/${TURN_CHECKLIST_FLAT.length}`,
      "New Tenant": t.newTenant || "", "Move-in": t.moveInDate ? fmtDate(t.moveInDate) : "",
      Rooms: (t.rooms || []).map(r => `${r.name}: ${TURN_ROOM_COLS.map(c => `${c.label.split(" ")[0]} ${r[c.key] || "-"}`).join(", ")}`).join(" | "),
    })));

    addSheet("Smoke-CO", (data.smokeCoDetectors || []).map(r => ({
      Building: buildingMap[r.buildingId] || "",
      Apt: unitMap[r.unitId] || "",
      "Date Installed": r.installedDate ? fmtDate(r.installedDate) : "",
    })));

    addSheet("Case Stage Options", (data.caseStageOptions || []).map(o => ({ Option: o.label || "", "Reminder after (days)": o.days || "" })));

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
            allDated.filter(i => i.date < today && i.type !== "Follow-up" && i.type !== "Weekend reminder" && i.type !== "Insurance renewal").length === 0
              ? <div className="hint">Nothing overdue.</div>
              : allDated.filter(i => i.date < today && i.type !== "Follow-up" && i.type !== "Weekend reminder" && i.type !== "Insurance renewal").sort((a, b) => a.date.localeCompare(b.date)).map(item => (
                <div key={item.key} className="dash-detail-item" onClick={() => goTo(item.tab, item.targetId || item.key.slice(item.key.indexOf("-") + 1))}>
                  <span className="pill pill-danger">{item.type}</span>
                  <div className="followup-item-main">
                    <div className="followup-item-name">{item.label}{item.sub ? <span className="row-muted"> — {item.sub}</span> : null}</div>
                  </div>
                  <span className="pill pill-muted">{fmtDate(item.date)}</span>
                  <IconBtn title="Mark completed" onClick={(e) => { e.stopPropagation(); completeDatedItem(item); }}><CheckCircle2 size={14} /></IconBtn>
                </div>
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
                <button key={t.id} className="dash-detail-item" onClick={() => goTo("rent", t.id)}>
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
        <DashboardCalendar data={data} buildingName={buildingName} tenantName={tenantName} setTab={setTab} goTo={goTo} />
      </div>

      <div className="dash-col-main">
        {rosterEntries.length > 0 && (
          <div className="dash-roster">
            <div className="dash-roster-title">Follow up with ({followUpEntries.length})</div>
              {rosterEntries.map(({ tenant: t, followUp: f }) => (
                <div className="dash-roster-row" key={f.id} onClick={() => goTo("rent", t.id)} style={{ cursor: "pointer" }}>
                  <div className="dash-roster-info">
                    <div className="dash-roster-name">{t.name}</div>
                    <div className="dash-roster-sub">{buildingName(t.buildingId)}{f.note ? ` — ${f.note}` : ""}{!f.note && t.balance ? ` — $${t.balance} behind` : ""}</div>
                  </div>
                  <div className="dash-roster-meta">
                    {t.phone && <a href={`tel:${t.phone}`} className="dash-roster-phone" onClick={(e) => e.stopPropagation()} title={t.phone}><Phone size={14} /></a>}
                    <span className={`pill ${f.date < today ? "pill-danger" : "pill-muted"}`}>{fmtDate(f.date)}</span>
                    <IconBtn
                      title="Done — remove this follow-up"
                      onClick={() => setData(d => ({
                        ...d,
                        tenants: d.tenants.map(t2 => t2.id === t.id ? { ...t2, followUps: tenantFollowUps(t2).filter(f2 => f2.id !== f.id) } : t2),
                      }))}
                    >
                      <CheckCircle2 size={14} />
                    </IconBtn>
                  </div>
                </div>
              ))}
              {followUpEntries.length > 8 && (
                <button className="btn-ghost" style={{ marginTop: 4 }} onClick={() => setShowAllFollowUps(x => !x)}>
                  {showAllFollowUps ? "Show fewer" : `Show all ${followUpEntries.length}`}
                </button>
              )}
              <button className="btn-ghost" style={{ marginTop: 6 }} onClick={() => setTab("rent")}>View in Rent Collection</button>
            </div>
          )}

      {totalAttention === 0 ? (
        <div className="all-clear"><CheckCircle2 size={18} /> Nothing needs attention right now.</div>
      ) : (
        <>

          {weekendItems.length > 0 && (
            <AttentionPanel
              icon={<CalendarClock size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
              label={<>Scheduled for Saturday <span className="dash-panel-sub">(reminder for today)</span></>} items={weekendItems} tab="dashboard" setTab={setTab}
              itemKey={i => i.key}
              renderItem={i => (
                <div className="followup-item-main" style={{ cursor: "pointer" }} onClick={() => goTo(i.tab, i.targetId)}>
                  <div className="followup-item-name">{i.label}{i.sub ? <span className="row-muted"> — {i.sub}</span> : null}</div>
                  {i.note && <div className="followup-item-note">{i.note}</div>}
                </div>
              )}
            />
          )}

          <AttentionPanel
            icon={<Gavel size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label={<>Payment follow-ups with the attorney <span className="dash-panel-sub">(day after a payment was due)</span></>} items={payFollowItems} tab="court" setTab={setTab}
            itemKey={i => i.key}
            renderItem={i => (
              <>
                <Flag date={i.date} />
                <div className="followup-item-main" style={{ cursor: "pointer" }} onClick={() => goTo("court", i.targetId)}>
                  <div className="followup-item-name">{i.label}{i.sub ? <span className="row-muted"> — {i.sub}</span> : null}</div>
                  {i.note && <div className="followup-item-note">{i.note}</div>}
                </div>
                <IconBtn title="Done — followed up" onClick={() => completeDatedItem(i)}><CheckCircle2 size={14} /></IconBtn>
              </>
            )}
          />

          <AttentionPanel
            icon={<ScrollText size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label={<>Insurance renewals <span className="dash-panel-sub">(start 2 months before expiry)</span></>} items={insuranceItems} tab="insurance" setTab={setTab}
            itemKey={i => i.key}
            renderItem={i => (
              <>
                <Flag date={i.date} />
                <div className="followup-item-main" style={{ cursor: "pointer" }} onClick={() => setTab("insurance")}>
                  <div className="followup-item-name">{i.label}</div>
                  {i.note && <div className="followup-item-note">{i.note}</div>}
                </div>
              </>
            )}
          />

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
                  <div className="followup-item-name">{buildingName(v.buildingId)}{v.unitId ? ` · Apt ${data.units.find(u => u.id === v.unitId)?.unitNumber || ""}` : ""} <span className="row-muted">— #{v.violationNumber}</span></div>
                  {v.description && <div className="followup-item-note">{summarizeViolationDescription(v.description)}</div>}
                  {v.hearingCompany && <div className="followup-item-note">Handled by: {v.hearingCompany}</div>}
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
            icon={<DollarSign size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label={<>Collections <span className="dash-panel-sub">({money(collectionsTotal)} owed by former tenants)</span></>} items={collectionItems} tab="collections" setTab={setTab}
            itemKey={c => c.id} limit={200}
            renderItem={c => (
              <div className="followup-item-main" style={{ cursor: "pointer" }} onClick={() => goTo("collections", c.id)}>
                <div className="followup-item-name">{c.tenantName} <span className="row-muted">— {buildingName(c.buildingId)} · {money(parseBalance(c.amountOwed))}</span></div>
                <div className="followup-item-note">{c.movedOutDate ? `Moved out ${fmtDate(c.movedOutDate)} · ` : ""}{c.lawyer ? `Lawyer: ${c.lawyer}` : "No lawyer yet"} · {c.stage || "stage not set"}</div>
              </div>
            )}
          />

          <AttentionPanel
            icon={<Siren size={18} className="attention-icon" style={{ color: "var(--danger)" }} />}
            label={<>Open emergency repair orders <span className="dash-panel-sub">(shows until completed)</span></>} items={eroItems} tab="ero" setTab={setTab}
            itemKey={e => e.id} limit={200}
            renderItem={e => (
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <div className="followup-item-main" style={{ flex: 1, cursor: "pointer" }} onClick={() => goTo("ero", e.id)}>
                  <div className="followup-item-name">ERO {e.eroNumber || ""} <span className="row-muted">— {buildingName(e.buildingId)}</span></div>
                  <div className="followup-item-note">{e.dueByDate ? `Due ${fmtDate(e.dueByDate)}${e.dueByDate < today ? " — overdue" : ""}` : "No due date"}{e.hpdContacted ? "" : " · HPD not contacted yet"}</div>
                </div>
                <IconBtn title="Mark completed" onClick={() => update("emergencyRepairs", e.id, { status: "Completed", completedDate: todayISO() })}><CheckCircle2 size={16} /></IconBtn>
              </div>
            )}
          />

          <AttentionPanel
            icon={<Pencil size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label={<>Quick notes to organize <span className="dash-panel-sub">(shows once due, or always if no reminder set)</span></>} items={quickNoteItems} tab="quicknotes" setTab={setTab} limit={200}
            itemKey={n => n.id}
            renderItem={n => (
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <div className="followup-item-main" style={{ flex: 1 }}>
                  <div className="followup-item-name">{n.text}{n.buildingId ? ` — ${buildingName(n.buildingId)}` : ""}</div>
                  <div className="followup-item-note">{n.reminderDate ? `Reminder: ${fmtDate(n.reminderDate)}${n.reminderDate < today ? " — overdue, stays here until checked off" : ""}` : fmtDate(n.date)}</div>
                </div>
                <IconBtn title="Done — check this off" onClick={() => update("quickNotes", n.id, { done: true })}><CheckCircle2 size={16} /></IconBtn>
              </div>
            )}
          />

          
          <AttentionPanel
            icon={<Home size={18} className="attention-icon" style={{ color: "var(--warn)" }} />}
            label={<>New tenants missing contact info <span className="dash-panel-sub">(replaced a moved-out tenant)</span></>} items={newTenantsNeedingContact} tab="rent" setTab={setTab}
            itemKey={t => t.id}
            renderItem={t => (
              <div className="followup-item-main" style={{ cursor: "pointer" }} onClick={() => goTo("rent", t.id)}>
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
                      <button key={t.id} className="dash-detail-item" onClick={(e) => { e.stopPropagation(); goTo("rent", t.id); }}>
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
      <div className="dash-storage-footer no-print">
        {(() => {
          const sizeBytes = JSON.stringify(rawData).length;
          const sizeKB = (sizeBytes / 1024).toFixed(0);
          const pctUsed = (sizeBytes / 1048576 * 100).toFixed(1);
          return `Document size: ${sizeKB} KB of 1,024 KB (${pctUsed}%) · version ${APP_BUILD}`;
        })()}
      </div>
    </div>
  );
}

/* ============================== buildings ============================== */

// Every piece of tenant info in one place — used for an existing tenant (each
// change saves straight away, sending only that field) and for the new-tenant
// form on the Buildings tab (changes collect in a draft until Save).
// Defined outside the editor so typing in an input never remounts it (which
// would drop focus after every keystroke).
function TenantFieldBox({ label, children }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 2, fontSize: 12, minWidth: 140, flex: "1 1 150px" }}>
      <span className="row-muted">{label}</span>{children}
    </label>
  );
}

function TenantInfoEditor({ value: t, onChange }) {
  const [occName, setOccName] = useState("");
  const [occPhone, setOccPhone] = useState("");
  const [occRel, setOccRel] = useState("");
  const occupants = Array.isArray(t.occupants) ? t.occupants : [];
  const addOccupant = () => {
    if (!occName.trim()) return;
    onChange({ occupants: [...occupants, { id: uid(), name: occName.trim(), phone: occPhone.trim(), relationship: occRel.trim() }] });
    setOccName(""); setOccPhone(""); setOccRel("");
  };
  const box = { border: "1px solid var(--border)", borderRadius: 4 };
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8, margin: "4px 0" }}>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
        <TenantFieldBox label="Name"><input className="sheet-input" style={box} placeholder="Tenant name" value={t.name || ""} onChange={e => onChange({ name: e.target.value })} /></TenantFieldBox>
        <TenantFieldBox label="Phone"><input className="sheet-input" style={box} type="tel" value={t.phone || ""} onChange={e => onChange({ phone: e.target.value })} /></TenantFieldBox>
        <TenantFieldBox label="Email"><input className="sheet-input" style={box} type="email" value={t.email || ""} onChange={e => onChange({ email: e.target.value })} /></TenantFieldBox>
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
        <TenantFieldBox label="Monthly rent ($)"><input className="sheet-input" style={box} inputMode="decimal" value={t.rentAmount || ""} onChange={e => onChange({ rentAmount: e.target.value })} /></TenantFieldBox>
        <TenantFieldBox label="Balance owed ($)"><input className="sheet-input" style={box} inputMode="decimal" value={t.balance || ""} onChange={e => onChange({ balance: e.target.value })} /></TenantFieldBox>
        <TenantFieldBox label="Status">
          <select style={box} value={t.status || "Current"} onChange={e => onChange({ status: e.target.value })}>
            {RENT_STATUSES.map(st => <option key={st}>{st}</option>)}
          </select>
        </TenantFieldBox>
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
        <TenantFieldBox label="Moved in"><input className="sheet-input" style={box} type="date" value={t.moveInDate || ""} onChange={e => onChange({ moveInDate: e.target.value })} /></TenantFieldBox>
        <TenantFieldBox label="Lease ends"><input className="sheet-input" style={box} type="date" value={t.leaseEnd || ""} onChange={e => onChange({ leaseEnd: e.target.value })} /></TenantFieldBox>
        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, flex: "1 1 150px" }}>
          <input type="checkbox" checked={!!t.movedOut} onChange={e => onChange({ movedOut: e.target.checked })} /> Moved out
        </label>
      </div>
      <div>
        <div className="row-muted" style={{ fontSize: 12, marginBottom: 2 }}>Other people in the apartment (roommates, family, emergency contact)</div>
        {occupants.length === 0 && <div className="hint">None added.</div>}
        {occupants.map(o => (
          <div key={o.id} style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 6, marginBottom: 3 }}>
            <input className="sheet-input" style={{ ...box, width: 150 }} placeholder="Name" value={o.name || ""} onChange={e => onChange({ occupants: occupants.map(x => x.id === o.id ? { ...x, name: e.target.value } : x) })} />
            <input className="sheet-input" style={{ ...box, width: 130 }} placeholder="Phone" type="tel" value={o.phone || ""} onChange={e => onChange({ occupants: occupants.map(x => x.id === o.id ? { ...x, phone: e.target.value } : x) })} />
            <input className="sheet-input" style={{ ...box, width: 130 }} placeholder="Who (spouse, roommate…)" value={o.relationship || ""} onChange={e => onChange({ occupants: occupants.map(x => x.id === o.id ? { ...x, relationship: e.target.value } : x) })} />
            <button type="button" className="checklist-remove" title="Remove" onClick={() => onChange({ occupants: occupants.filter(x => x.id !== o.id) })}><X size={12} /></button>
          </div>
        ))}
        <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 6 }}>
          <input className="sheet-input" style={{ ...box, width: 150 }} placeholder="Name" value={occName} onChange={e => setOccName(e.target.value)} onKeyDown={e => e.key === "Enter" && addOccupant()} />
          <input className="sheet-input" style={{ ...box, width: 130 }} placeholder="Phone" type="tel" value={occPhone} onChange={e => setOccPhone(e.target.value)} onKeyDown={e => e.key === "Enter" && addOccupant()} />
          <input className="sheet-input" style={{ ...box, width: 130 }} placeholder="Who (spouse, roommate…)" value={occRel} onChange={e => setOccRel(e.target.value)} onKeyDown={e => e.key === "Enter" && addOccupant()} />
          <button type="button" className="btn-ghost" disabled={!occName.trim()} onClick={addOccupant}><Plus size={14} /> Add person</button>
        </div>
      </div>
    </div>
  );
}

function BuildingsTab({ data, add, update, remove, setData, buildingName }) {
  const [form, setForm] = useState(null);
  // The form panel renders once, near the top of the page — not inline with
  // whichever building card was clicked. With a full list of buildings,
  // clicking Edit on one further down opens the form off-screen above the
  // current scroll position, which looks exactly like the button doing
  // nothing at all. Scroll it into view the moment it opens instead.
  const formRef = useRef(null);
  useEffect(() => {
    if (form && formRef.current) formRef.current.scrollIntoView({ behavior: "smooth", block: "start" });
  }, [form]);
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
  // The new-tenant form's draft, for whichever unit it's open on.
  const [draftTenant, setDraftTenant] = useState(null); // { unitId, buildingId, ...fields }
  const blankDraft = (buildingId, unitId) => ({ unitId, buildingId, name: "", phone: "", email: "", rentAmount: "", balance: "", status: "Current", moveInDate: "", leaseEnd: "", movedOut: false, occupants: [] });
  // Same fields the Rent page's own "Add tenant" creates, so a tenant added here
  // behaves identically everywhere else (rent roll, follow-ups, dashboard).
  const saveDraftTenant = () => {
    if (!draftTenant || !(draftTenant.name || "").trim()) return;
    add("tenants", { ...draftTenant, name: draftTenant.name.trim(), notes: [], followUps: [], payments: [] });
    setDraftTenant(null);
  };
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
          <button className="btn-primary" onClick={() => setForm({ address: "", alsoKnownAs: "", notes: "" })}><Plus size={14} /> Add building</button>
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
        <div className="form-panel" ref={formRef}>
          <Field label="Address"><input value={form.address} onChange={e => setForm({ ...form, address: e.target.value })} /></Field>
          <Field label="Also known as (optional)">
            <input
              value={form.alsoKnownAs || ""} onChange={e => setForm({ ...form, alsoKnownAs: e.target.value })}
              placeholder="e.g. a cross-street address HPD lists this corner building under, comma-separated if more than one"
            />
          </Field>
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
      {(() => {
        const q = unitSearch.trim().toLowerCase();
        if (q && data.buildings.length > 0) {
          const anyMatch = data.buildings.some(b =>
            data.units.some(u => u.buildingId === b.id && ((u.unitNumber || "").toLowerCase().includes(q) || data.tenants.some(t => t.unitId === u.id && (t.name || "").toLowerCase().includes(q))))
          );
          if (!anyMatch) return <EmptyState text="No unit or tenant matches that search." />;
        }
        return null;
      })()}
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
                              {t.movedOut && <span className="pill pill-warn">Moved Out</span>}
                              <TenantInfoEditor value={t} onChange={(patch) => update("tenants", t.id, patch)} />
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
                          {!tenants.some(t => !t.movedOut) && (
                            draftTenant && draftTenant.unitId === u.id ? (
                              <div className="form-panel" style={{ marginTop: 6 }}>
                                <strong>New tenant</strong>
                                <TenantInfoEditor value={draftTenant} onChange={(patch) => setDraftTenant(d => ({ ...d, ...patch }))} />
                                <div className="form-actions">
                                  <button className="btn-primary" disabled={!(draftTenant.name || "").trim()} onClick={saveDraftTenant}>Save tenant</button>
                                  <button className="btn-ghost" onClick={() => setDraftTenant(null)}>Cancel</button>
                                </div>
                              </div>
                            ) : (
                              <button className="btn-ghost" style={{ marginTop: 6 }} onClick={() => setDraftTenant(blankDraft(b.id, u.id))}>
                                <Plus size={14} /> Add tenant
                              </button>
                            )
                          )}
                        </div>
                      )}
                    </div>
                  );
                })}
                <button className="btn-ghost" style={{ marginTop: 8 }} onClick={() => { const id = add("units", { buildingId: b.id, unitNumber: "" }); setExpandedUnit(id); }}>
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

function RentTab({ data: rawData, add, update, remove, buildingName, setData, deepLink, setDeepLink }) {
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
  const [highlightTenantId, setHighlightTenantId] = useState(null);

  useEffect(() => {
    if (!deepLink || deepLink.tab !== "rent") return;
    // Look in rawData, not the scoped copy, so a tenant in one of the "Mitch's
    // father" buildings (kept out of the main list) is found too.
    const target = rawData.tenants.find(t => t.id === deepLink.id);
    if (target) {
      setSection("sheet");
      setStatusFilter("All"); // a filter that excludes this tenant would otherwise hide them even with their building expanded
      if (!mainBuildingIds.has(target.buildingId)) {
        setExcludedSectionOpen(true); // they live in the "Mitch's father" section at the bottom
      } else if (target.movedOut) {
        setMainView("movedOut"); // moved-out tenants have their own view, grouped as "movedOut_<building>"
        setExpandedBuildings(prev => new Set(prev).add("movedOut_" + target.buildingId));
      } else {
        setMainView("current");
        setExpandedBuildings(prev => new Set(prev).add(target.buildingId));
      }
      setHighlightTenantId(target.id);
      setTimeout(() => {
        document.getElementById(`tenant-row-${target.id}`)?.scrollIntoView({ behavior: "smooth", block: "center" });
      }, 100);
      setTimeout(() => setHighlightTenantId(null), 2500);
    }
    setDeepLink(null);
  }, [deepLink]);

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
      stage: "", nextCourtDate: "", result: "Pending",
      stipulationTerms: "", nextPaymentDue: "",
      archived: false, documents: [],
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
    })
    // A building with tenants overall but none matching the active filter
    // shouldn't still get a card — that's an empty building header for
    // nothing, exactly the noise a filter like "Follow-ups" is meant to
    // cut through.
    .filter(g => g.tenants.length > 0);

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
                id={`tenant-row-${t.id}`}
                className={`${t.status !== "Current" ? "sheet-row-flag" : ""} ${selectedTenantId === t.id ? "sheet-row-selected" : ""} ${highlightTenantId === t.id ? "sheet-row-deeplinked" : ""}`}
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
                    title={t.voucherProgram ? "On a voucher program (e.g. Section 8) — click to clear" : "Mark as on a voucher program (e.g. Section 8)"}
                    active={!!t.voucherProgram}
                    onClick={() => update("tenants", t.id, { voucherProgram: !t.voucherProgram })}
                  ><Ticket size={14} /></IconBtn>
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
          ) : buildingGroups.length === 0 ? (
            <EmptyState text="Nothing matches this filter right now." />
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
    // A row needing approval (an active follow-up or call-back flag) is
    // silently excluded from what actually gets applied below unless its
    // checkbox was checked — easy to miss among many other rows, and the
    // exact scenario that looks like "the balance just didn't update" when
    // it's really "this specific change was never approved." Surfacing it
    // as an explicit, unmissable choice here instead of leaving it to a
    // small checkbox and a line of summary text further up the page.
    const stillPending = preview.changes.filter(c => c.needsApproval && !c.approved);
    if (stillPending.length > 0) {
      const names = stillPending.slice(0, 5).map(c => c.name).join(", ");
      const more = stillPending.length > 5 ? ` and ${stillPending.length - 5} more` : "";
      const proceed = window.confirm(
        `${stillPending.length} change${stillPending.length === 1 ? "" : "s"} still need${stillPending.length === 1 ? "s" : ""} approval and won't be applied: ${names}${more}.\n\nContinue and skip those, or Cancel to go back and check their approval boxes first?`
      );
      if (!proceed) return;
    }
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
      const newHistoryEntry = {
        id: uid(), date: todayISO(), buildingId, type,
        updated: applied.filter(c => !c.isNew).length,
        added: applied.filter(c => c.isNew).length,
        missing: preview.missing.length,
        skipped: preview.changes.length - applied.length,
        details: applied.map(c => ({ apt: c.apt, name: c.name, isNew: c.isNew, fields: c.fields, before: c.before })),
        missingDetails: preview.missing,
      };
      // Keep only the 10 most recent entries for this specific building and
      // import type — older ones for that combination are dropped
      // automatically here, rather than needing a manual cleanup click
      // each time storage creeps up. A different building or type keeps
      // its own separate history, untouched by this.
      const sameKey = next.importHistory.filter(h => h.buildingId === buildingId && h.type === type);
      const keepIds = new Set(sameKey.slice(0, 9).map(h => h.id));
      next.importHistory = [
        newHistoryEntry,
        ...next.importHistory.filter(h => h.buildingId !== buildingId || h.type !== type || keepIds.has(h.id)),
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
  // This only ever removes the historical log entry — a record of what an
  // import did at the time — never the actual tenants/buildings/units it
  // created or updated, which live in their own separate collections and
  // are completely untouched by any of this.
  const deleteHistoryEntry = (id) => setData(d => ({ ...d, importHistory: (d.importHistory || []).filter(h => h.id !== id) }));
  const oldCutoff = new Date(); oldCutoff.setDate(oldCutoff.getDate() - 30);
  const oldCount = history.filter(h => h.date && h.date < toLocalISO(oldCutoff)).length;
  const deleteOldHistory = () => {
    const cutoffIso = toLocalISO(oldCutoff);
    setData(d => ({ ...d, importHistory: (d.importHistory || []).filter(h => !(h.date && h.date < cutoffIso)) }));
  };

  return (
    <div>
      {availableTypes.map(t => (
        <ImportBlock key={t.key} type={t.key} data={data} setData={setData} buildingName={buildingName} />
      ))}

      <div className="row" style={{ marginTop: 20, marginBottom: 4 }}>
        <h2 className="section-heading" style={{ margin: 0 }}>Last Imported</h2>
        {oldCount > 0 && (
          <button className="btn-ghost" style={{ marginLeft: 8 }} onClick={deleteOldHistory}>
            <Trash2 size={14} /> Clear {oldCount} older than 30 days
          </button>
        )}
      </div>
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
            <div className="spacer" />
            <IconBtn title="Delete this import record (doesn't undo the import itself)" danger onClick={(e) => { e.stopPropagation(); deleteHistoryEntry(h.id); }}><Trash2 size={14} /></IconBtn>
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

// One cell of the Work Orders spreadsheet: you just type in it. The change is
// saved when you click away or press Enter (Shift+Enter = new line in a text
// cell). onCommit may return false to say "that didn't match anything" — the
// cell then flashes red and goes back to what it was.
function WoCell({ value, onCommit, type = "text", multiline = false, placeholder = "", align, narrow }) {
  const [v, setV] = useState(value ?? "");
  const [focus, setFocus] = useState(false);
  const [bad, setBad] = useState(false);
  const ref = useRef(null);
  useEffect(() => { if (!focus) setV(value ?? ""); }, [value, focus]);
  useLayoutEffect(() => {
    if (multiline && ref.current) { ref.current.style.height = "auto"; ref.current.style.height = ref.current.scrollHeight + "px"; }
  }, [v, multiline]);
  const commit = () => {
    setFocus(false);
    if (v === (value ?? "")) return;
    const ok = onCommit(v);
    if (ok === false) { setV(value ?? ""); setBad(true); setTimeout(() => setBad(false), 1400); }
  };
  const common = {
    ref, value: v, placeholder, className: `wo-cell${bad ? " wo-cell-bad" : ""}`,
    onChange: e => setV(e.target.value), onFocus: () => setFocus(true), onBlur: commit,
    style: { textAlign: align, minWidth: narrow },
  };
  return multiline
    ? <textarea rows={1} {...common} onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); e.currentTarget.blur(); } }} />
    : <input type={type} {...common} onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }} />;
}

function WorkOrdersTab({ data, add, update, remove, buildingName, vendorName, tenantName, deepLink, setDeepLink }) {
  const [form, setForm] = useState(null);
  const [view, setView] = useState("active");
  const [filter, setFilter] = useState("All");
  const [courtOnly, setCourtOnly] = useState(false);
  // The record exactly as it was when the edit form was opened — what the form
  // is compared against on Save, so only what was actually changed gets saved.
  const [formOriginal, setFormOriginal] = useState(null);
  const [search, setSearch] = useState("");
  const [vendorFilter, setVendorFilter] = useState("All");
  const [buildingFilter, setBuildingFilter] = useState("All");
  const [selected, setSelected] = useState(new Set());
  const [copiedId, setCopiedId] = useState(null);
  const [expandedRow, setExpandedRow] = useState(null);
  const [noteFor, setNoteFor] = useState(null);
  const [noteText, setNoteText] = useState("");
  const [highlightWoId, setHighlightWoId] = useState(null);
  // Scheduling: which work order's schedule box is open, and what's typed in it.
  const [schedulingFor, setSchedulingFor] = useState(null);
  const [schedDate, setSchedDate] = useState("");
  const [schedTime, setSchedTime] = useState("");
  const openSchedule = (w) => {
    if (schedulingFor === w.id) { setSchedulingFor(null); return; }
    setSchedDate(w.scheduledDate || ""); setSchedTime(w.scheduledTime || ""); setSchedulingFor(w.id);
  };
  // Only these two fields are sent — never the whole work order — so scheduling
  // can't put back anything that was changed on it somewhere else.
  const saveSchedule = (w) => {
    if (!schedDate) return;
    update("workOrders", w.id, { scheduledDate: schedDate, scheduledTime: schedTime });
    setSchedulingFor(null);
  };
  const clearSchedule = (w) => {
    update("workOrders", w.id, { scheduledDate: "", scheduledTime: "" });
    setSchedulingFor(null);
  };

  // Typing-based lookups for the spreadsheet cells.
  const norm = (x) => String(x || "").trim().toLowerCase();
  const findBuildingByText = (t) => {
    const q = norm(t); if (!q) return null;
    const exact = data.buildings.find(b => norm(buildingName(b.id)) === q || norm(shortAddress(b.address)) === q || norm(b.address) === q);
    if (exact) return exact;
    const hits = data.buildings.filter(b => norm(b.address).includes(q) || norm(buildingName(b.id)).includes(q));
    return hits.length === 1 ? hits[0] : (hits.length > 1 ? hits.find(b => norm(b.address).startsWith(q)) || null : null);
  };
  const findUnitByText = (buildingId, t) => {
    const q = norm(t).replace(/^(apt|apartment|unit|#)\s*/, "");
    return data.units.find(u => u.buildingId === buildingId && norm(u.unitNumber) === q) || null;
  };
  const findVendorByText = (t) => {
    const q = norm(t); if (!q) return null;
    const exact = data.vendors.find(v => norm(v.name) === q);
    if (exact) return exact;
    const hits = data.vendors.filter(v => norm(v.name).includes(q));
    return hits.length === 1 ? hits[0] : null;
  };
  const blankRow = () => ({ apt: "", building: buildingFilter !== "All" ? buildingName(buildingFilter) : "", description: "", vendor: "", priority: "", due: "", court: false });
  const [newRow, setNewRow] = useState(blankRow);
  const [newRowError, setNewRowError] = useState("");
  const addNewRow = () => {
    if (!newRow.description.trim()) { setNewRowError("Type a description"); return; }
    const b = findBuildingByText(newRow.building);
    if (!b) { setNewRowError("Building not found"); return; }
    let unitId = "";
    if (newRow.apt.trim()) { const u = findUnitByText(b.id, newRow.apt); if (!u) { setNewRowError("Apt not found in that building"); return; } unitId = u.id; }
    let vendorId = "";
    if (newRow.vendor.trim()) { const v = findVendorByText(newRow.vendor); if (!v) { setNewRowError("Vendor not found"); return; } vendorId = v.id; }
    let priority = "Routine";
    if (newRow.priority.trim()) { const p = WO_PRIORITIES.find(x => x.toLowerCase().startsWith(newRow.priority.trim().toLowerCase())); if (!p) { setNewRowError("Priority: Routine, Urgent or Emergency"); return; } priority = p; }
    add("workOrders", {
      id: uid(), buildingId: b.id, unitId, vendorId, description: newRow.description.trim(), status: "Open", priority,
      dateOpened: todayISO(), isCourtConnected: !!newRow.court, courtCaseId: "", dueByDate: newRow.due || "", accessDate: "", accessGranted: null, items: [],
    });
    setNewRow({ ...blankRow(), building: newRow.building });
    setNewRowError("");
  };
  const newRowKey = (e) => { if (e.key === "Enter") addNewRow(); };

  useEffect(() => {
    if (!deepLink || deepLink.tab !== "workorders") return;
    const target = data.workOrders.find(w => w.id === deepLink.id);
    if (target) {
      setView(target.status === "Done" ? "closed" : "active");
      setFilter("All");
      setCourtOnly(false);
      setVendorFilter("All");
      setBuildingFilter("All");
      setSearch("");
      setHighlightWoId(target.id);
      setTimeout(() => {
        document.getElementById(`wo-row-${target.id}`)?.scrollIntoView({ behavior: "smooth", block: "center" });
      }, 100);
      setTimeout(() => setHighlightWoId(null), 2500);
    }
    setDeepLink(null);
  }, [deepLink]);

  const submit = () => {
    if (!form.description) return;
    const exists = data.workOrders.some(w => w.id === form.id);
    if (exists) {
      const patch = changedFieldsOnly(formOriginal, form);
      if (Object.keys(patch).length > 0) update("workOrders", form.id, patch);
    }
    else {
      add("workOrders", form); // form.id was pre-generated when the form opened
    }
    setForm(null);
  };

  const list = data.workOrders
    .filter(w => view === "active" ? w.status !== "Done" : w.status === "Done")
    .filter(w => view === "closed" || filter === "All" || w.status === filter)
    .filter(w => !courtOnly || w.isCourtConnected || w.courtCaseId)
    .filter(w => buildingFilter === "All" || w.buildingId === buildingFilter)
    .filter(w => {
      if (vendorFilter === "All") return true;
      if (w.items && w.items.length > 0) return w.items.some(it => it.vendorId === vendorFilter);
      return w.vendorId === vendorFilter;
    })
    .filter(w => {
      if (!search.trim()) return true;
      const q = search.trim().toLowerCase();
      const unit = w.unitId ? data.units.find(u => u.id === w.unitId) : null;
      const unitTenant = w.unitId ? data.tenants.find(t => t.unitId === w.unitId && !t.movedOut) : null;
      const vendorNames = w.items && w.items.length > 0
        ? w.items.map(it => it.vendorId ? vendorName(it.vendorId) : "").join(" ")
        : (w.vendorId ? vendorName(w.vendorId) : "");
      const desc = w.items && w.items.length > 0 ? w.items.map(it => it.description).join(" ") : w.description;
      return [desc, unit?.unitNumber, buildingName(w.buildingId), vendorNames, w.status, w.priority, unitTenant?.name]
        .some(f => (f || "").toString().toLowerCase().includes(q));
    })
    .slice()
    .sort((a, b) => {
      // Most urgent first: anything tied to a court case, soonest due date on top;
      // then emergencies/urgent; then newest opened.
      const ac = (a.isCourtConnected || a.courtCaseId) ? 0 : 1, bc = (b.isCourtConnected || b.courtCaseId) ? 0 : 1;
      if (ac !== bc) return ac - bc;
      if (ac === 0) {
        const d = (a.dueByDate || "9999-99-99").localeCompare(b.dueByDate || "9999-99-99");
        if (d !== 0) return d;
      }
      const pr = (w) => (w.priority === "Emergency" ? 0 : w.priority === "Urgent" ? 1 : 2);
      if (pr(a) !== pr(b)) return pr(a) - pr(b);
      return (b.dateOpened || "").localeCompare(a.dateOpened || "");
    });

  const toggleSelected = (id) => {
    setSelected(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  // A work order's own description can go stale relative to its items once
  // they've been individually edited post-split (see the list-card-title
  // fix above) — same correction applied here before handing off to the
  // shared copy-formatting function, which just reads .description as-is.
  const withCurrentDescription = (w) => ({ ...w, description: w.items && w.items.length > 0 ? w.items.map(it => it.description).join(", ") : w.description });
  const copyOne = (w) => {
    navigator.clipboard.writeText(formatItemsForCopy([withCurrentDescription(w)], data));
    setCopiedId(w.id);
    setTimeout(() => setCopiedId(null), 1500);
  };
  const copySelected = () => {
    const items = list.filter(w => selected.has(w.id)).map(withCurrentDescription);
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
          <button className="btn-primary" onClick={() => setForm({ id: uid(), buildingId: data.buildings[0]?.id || "", unitId: "", vendorId: "", description: "", status: "Open", priority: "Routine", dateOpened: todayISO(), isCourtConnected: false, courtCaseId: "", dueByDate: "", accessDate: "", accessGranted: null, items: [] })}>
            <Plus size={14} /> Add work order
          </button>
        </div>
      </div>
      <div className="row" style={{ marginBottom: 10, flexWrap: "wrap" }}>
        <input
          type="text" placeholder="Search description, apt, building, vendor, status…" value={search}
          onChange={e => setSearch(e.target.value)}
          style={{ width: "100%", maxWidth: 360 }}
        />
        <select value={buildingFilter} onChange={e => { setBuildingFilter(e.target.value); setSelected(new Set()); }} style={{ maxWidth: 220 }}>
          <option value="All">All buildings</option>
          {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
        </select>
        <select value={vendorFilter} onChange={e => { setVendorFilter(e.target.value); setSelected(new Set()); }} style={{ maxWidth: 220 }}>
          <option value="All">All vendors</option>
          {data.vendors.map(v => <option key={v.id} value={v.id}>{v.name}</option>)}
        </select>
      </div>
      <div className="filter-row">
        <button className={`chip ${view === "active" ? "chip-active" : ""}`} onClick={() => { setView("active"); setSelected(new Set()); }}>Active</button>
        <button className={`chip ${view === "closed" ? "chip-active" : ""}`} onClick={() => { setView("closed"); setSelected(new Set()); }}>Done</button>
        <span className="wo-filter-divider" />
        <button className={`chip ${courtOnly ? "chip-active" : ""}`} onClick={() => { setCourtOnly(o => !o); setSelected(new Set()); }}><Gavel size={11} /> Court-linked only</button>
      </div>
      {view === "active" && (
        <div className="filter-row">
          {["All", ...WO_STATUSES.filter(s => s !== "Done")].map(s => (
            <button key={s} className={`chip ${filter === s ? "chip-active" : ""}`} onClick={() => { setFilter(s); setSelected(new Set()); }}>{s}</button>
          ))}
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
          <Field label="Vendor">
            <select value={form.vendorId} onChange={e => setForm({ ...form, vendorId: e.target.value })}>
              <option value="">Unassigned</option>
              {data.vendors.map(v => <option key={v.id} value={v.id}>{v.name}</option>)}
            </select>
          </Field>
          <Field label="Description">
            <textarea
              value={form.description}
              onChange={e => setForm({ ...form, description: e.target.value })}
              onBlur={() => {
                // Auto-split into separate items once the person is done
                // typing (on blur, not on every keystroke, so a comma mid-
                // sentence doesn't fragment things before they've finished
                // writing) — each item gets its own vendor assignment
                // below, carrying forward whatever vendor was already set
                // as a starting point for all of them.
                const parts = form.description.split(",").map(s => s.trim()).filter(Boolean);
                if (parts.length > 1) {
                  setForm({ ...form, items: parts.map(p => ({ id: uid(), description: p, vendorId: form.vendorId || "" })) });
                }
              }}
            />
          </Field>
          {form.items && form.items.length > 0 ? (
            <Field label="Vendor per item">
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {form.items.map((it, i) => (
                  <div key={it.id} style={{ display: "flex", gap: 8, alignItems: "center" }}>
                    <input
                      style={{ flex: 1 }} value={it.description}
                      onChange={e => setForm({ ...form, items: form.items.map((x, xi) => xi === i ? { ...x, description: e.target.value } : x) })}
                    />
                    <select
                      value={it.vendorId} style={{ flex: 1 }}
                      onChange={e => setForm({ ...form, items: form.items.map((x, xi) => xi === i ? { ...x, vendorId: e.target.value } : x) })}
                    >
                      <option value="">Unassigned</option>
                      {data.vendors.map(v => <option key={v.id} value={v.id}>{v.name}</option>)}
                    </select>
                    <button type="button" className="checklist-remove" title="Remove this item" onClick={() => setForm({ ...form, items: form.items.filter((_, xi) => xi !== i) })}><X size={12} /></button>
                  </div>
                ))}
                <button type="button" className="btn-ghost" onClick={() => setForm({ ...form, items: [] })}>Merge back into one item</button>
              </div>
            </Field>
          ) : (
            <Field label="Vendor">
              <select value={form.vendorId} onChange={e => setForm({ ...form, vendorId: e.target.value })}>
                <option value="">Unassigned</option>
                {data.vendors.map(v => <option key={v.id} value={v.id}>{v.name}</option>)}
              </select>
            </Field>
          )}
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
          <Field label="Connected to a court case?">
            <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}>
              <input
                type="checkbox" checked={!!form.isCourtConnected}
                onChange={e => setForm({
                  ...form, isCourtConnected: e.target.checked,
                  // Unchecking clears the link and anything that only makes sense
                  // alongside it, rather than leaving stale court data behind on
                  // a work order that's no longer marked as court-connected.
                  ...(e.target.checked ? {} : { courtCaseId: "", dueByDate: "", accessDate: "", accessGranted: null }),
                })}
              />
              Yes, this is connected to a court case
            </label>
          </Field>
          {form.isCourtConnected && (
            <>
              <Field label="Linked case">
                <select value={form.courtCaseId || ""} onChange={e => setForm({ ...form, courtCaseId: e.target.value })}>
                  <option value="">Select a case…</option>
                  {data.courtCases.filter(c => !c.archived && c.buildingId === form.buildingId).map(c => (
                    <option key={c.id} value={c.id}>{c.tenantId ? tenantName(c.tenantId) : (c.rawName || "Case")} {c.caseNumber ? `#${c.caseNumber}` : ""}</option>
                  ))}
                </select>
              </Field>
              <Field label="Needs to be completed by"><input type="date" value={form.dueByDate || ""} onChange={e => setForm({ ...form, dueByDate: e.target.value })} /></Field>
              <Field label="Access date (optional)"><input type="date" value={form.accessDate || ""} onChange={e => setForm({ ...form, accessDate: e.target.value })} /></Field>
              <Field label="Do we have access on that date?">
                <select
                  value={form.accessGranted === true ? "yes" : form.accessGranted === false ? "no" : ""}
                  onChange={e => setForm({ ...form, accessGranted: e.target.value === "yes" ? true : e.target.value === "no" ? false : null })}
                >
                  <option value="">Not yet confirmed</option>
                  <option value="yes">Yes — access confirmed</option>
                  <option value="no">No — access denied/unclear</option>
                </select>
              </Field>
            </>
          )}
          <div className="form-actions">
            <button className="btn-primary" onClick={submit}>Save</button>
            <button className="btn-ghost" onClick={() => setForm(null)}>Cancel</button>
          </div>
        </div>
      )}

      {list.length === 0 && view !== "active" && <EmptyState text="Nothing marked done yet." />}
      <div className="hint" style={{ marginBottom: 6 }}>Just type in any cell — it saves when you click away or press Enter. Building, apt and vendor can be typed as a few letters of the name.</div>
      <div className="sheet-wrap wo-sheet-wrap">
        <table className="sheet wo-sheet">
          <thead>
            <tr>
              <th>Apt</th>
              <th>Building</th>
              <th>Description</th>
              <th>Vendor</th>
              <th>Priority</th>
              <th>Status</th>
              <th>Court</th>
              <th>Due by</th>
              <th>Scheduled</th>
              <th>Notes</th>
              <th style={{ width: 170 }}></th>
            </tr>
          </thead>
          <tbody>
            {view === "active" && (
              <tr className="wo-sheet-row wo-new-row">
                <td className="wo-edit" data-label="Apt"><input className="wo-cell" placeholder="Apt" value={newRow.apt} onChange={e => setNewRow({ ...newRow, apt: e.target.value })} onKeyDown={newRowKey} /></td>
                <td className="wo-edit" data-label="Building"><input className="wo-cell" placeholder="Building" value={newRow.building} onChange={e => setNewRow({ ...newRow, building: e.target.value })} onKeyDown={newRowKey} /></td>
                <td className="wo-edit" data-label="Description"><input className="wo-cell" placeholder="New work order — type here" value={newRow.description} onChange={e => setNewRow({ ...newRow, description: e.target.value })} onKeyDown={newRowKey} /></td>
                <td className="wo-edit" data-label="Vendor"><input className="wo-cell" placeholder="Vendor" value={newRow.vendor} onChange={e => setNewRow({ ...newRow, vendor: e.target.value })} onKeyDown={newRowKey} /></td>
                <td className="wo-edit" data-label="Priority"><input className="wo-cell" placeholder="Routine" value={newRow.priority} onChange={e => setNewRow({ ...newRow, priority: e.target.value })} onKeyDown={newRowKey} /></td>
                <td className="wo-edit" data-label="Status"><span className="row-muted" style={{ padding: "8px 10px", display: "block" }}>Open</span></td>
                <td className="wo-edit" data-label="Court"><input type="checkbox" style={{ margin: 10 }} checked={newRow.court} onChange={e => setNewRow({ ...newRow, court: e.target.checked })} title="Connected to a court case" /></td>
                <td className="wo-edit" data-label="Due by"><input className="wo-cell" type="date" value={newRow.due} onChange={e => setNewRow({ ...newRow, due: e.target.value })} onKeyDown={newRowKey} /></td>
                <td className="wo-edit" data-label="Scheduled"></td>
                <td className="wo-edit" data-label="Notes">{newRowError && <span style={{ color: "var(--danger, #B3261E)", fontSize: 12, padding: "8px 10px", display: "block" }}>{newRowError}</span>}</td>
                <td className="wo-sheet-actions"><button className="btn-primary" onClick={addNewRow}><Plus size={14} /> Add</button></td>
              </tr>
            )}
            {list.map(w => {
              const isOpen = expandedRow === w.id;
              const sortedNotes = newestFirst(w.notes);
              const unit = w.unitId ? data.units.find(u => u.id === w.unitId) : null;
              const unitTenant = w.unitId ? data.tenants.find(t => t.unitId === w.unitId && !t.movedOut) : null;
              const hasItems = w.items && w.items.length > 0;
              const isCourt = !!(w.isCourtConnected || w.courtCaseId);
              const setItem = (idx, patch) => {
                const items = w.items.map((x, xi) => xi === idx ? { ...x, ...patch } : x);
                update("workOrders", w.id, { items, description: items.map(it => it.description).join(", ") });
              };
              return (
                <React.Fragment key={w.id}>
                  <tr className={`wo-sheet-row ${highlightWoId === w.id ? "sheet-row-deeplinked" : ""}`} id={`wo-row-${w.id}`}>
                    <td className="wo-edit" data-label="Apt">
                      <WoCell value={unit ? (unit.unitNumber || "") : ""} placeholder="—" narrow={50} onCommit={t => {
                        if (!t.trim()) { update("workOrders", w.id, { unitId: "" }); return true; }
                        const u = findUnitByText(w.buildingId, t);
                        if (!u) return false;
                        update("workOrders", w.id, { unitId: u.id }); return true;
                      }} />
                    </td>
                    <td className="wo-edit" data-label="Building">
                      <WoCell value={buildingName(w.buildingId)} narrow={110} onCommit={t => {
                        const b = findBuildingByText(t);
                        if (!b) return false;
                        let unitId = "";
                        if (unit) { const same = data.units.find(u => u.buildingId === b.id && String(u.unitNumber).toLowerCase() === String(unit.unitNumber).toLowerCase()); unitId = same ? same.id : ""; }
                        update("workOrders", w.id, { buildingId: b.id, unitId }); return true;
                      }} />
                    </td>
                    <td className="wo-edit wo-sheet-desc" data-label="Description">
                      {hasItems
                        ? w.items.map((it, idx) => <WoCell key={it.id} multiline value={it.description} onCommit={t => { setItem(idx, { description: t }); return true; }} />)
                        : <WoCell multiline value={w.description || ""} onCommit={t => { update("workOrders", w.id, { description: t }); return true; }} />}
                      {unitTenant && (
                        <div className="wo-sheet-tenant" style={{ padding: "0 10px 6px" }}>
                          {unitTenant.name}
                          {(unitTenant.phone || "").split(";").map(s2 => s2.trim()).filter(Boolean).map((num, i) => (
                            <a key={i} href={`tel:${num}`} className="wo-tenant-phone" style={{ marginLeft: 6 }}>{num}</a>
                          ))}
                        </div>
                      )}
                    </td>
                    <td className="wo-edit" data-label="Vendor">
                      {hasItems
                        ? w.items.map((it, idx) => <WoCell key={it.id} value={it.vendorId ? vendorName(it.vendorId) : ""} placeholder="—" narrow={100} onCommit={t => {
                            if (!t.trim()) { setItem(idx, { vendorId: "" }); return true; }
                            const v = findVendorByText(t); if (!v) return false; setItem(idx, { vendorId: v.id }); return true;
                          }} />)
                        : <WoCell value={w.vendorId ? vendorName(w.vendorId) : ""} placeholder="—" narrow={100} onCommit={t => {
                            if (!t.trim()) { update("workOrders", w.id, { vendorId: "" }); return true; }
                            const v = findVendorByText(t); if (!v) return false; update("workOrders", w.id, { vendorId: v.id }); return true;
                          }} />}
                    </td>
                    <td className="wo-edit" data-label="Priority">
                      <WoCell value={w.priority || "Routine"} narrow={80} onCommit={t => {
                        const p = WO_PRIORITIES.find(x => x.toLowerCase().startsWith(t.trim().toLowerCase()));
                        if (!p || !t.trim()) return false;
                        update("workOrders", w.id, { priority: p }); return true;
                      }} />
                    </td>
                    <td className="wo-edit" data-label="Status">
                      <WoCell value={w.status || ""} narrow={90} onCommit={t => {
                        const st = WO_STATUSES.find(x => x.toLowerCase().startsWith(t.trim().toLowerCase()));
                        if (!st || !t.trim()) return false;
                        update("workOrders", w.id, { status: st }); return true;
                      }} />
                    </td>
                    <td className="wo-edit" data-label="Court">
                      <input type="checkbox" style={{ margin: 10 }} checked={isCourt} title="Connected to a court case"
                        onChange={e => update("workOrders", w.id, e.target.checked ? { isCourtConnected: true } : { isCourtConnected: false, courtCaseId: "", dueByDate: "", accessDate: "", accessGranted: null })} />
                      {isCourt && w.dueByDate && w.status !== "Done" && <div style={{ padding: "0 8px 6px" }}><Flag date={w.dueByDate} label="due by" /></div>}
                    </td>
                    <td className="wo-edit" data-label="Due by">
                      <WoCell type="date" value={w.dueByDate || ""} onCommit={t => { update("workOrders", w.id, { dueByDate: t }); return true; }} />
                    </td>
                    <td className="wo-edit" data-label="Scheduled">
                      <WoCell type="date" value={w.scheduledDate || ""} onCommit={t => { update("workOrders", w.id, t ? { scheduledDate: t } : { scheduledDate: "", scheduledTime: "" }); return true; }} />
                      {w.scheduledDate && w.scheduledTime && <div className="row-muted" style={{ padding: "0 10px 6px", fontSize: 11 }}>{fmtTime(w.scheduledTime)}</div>}
                    </td>
                    <td className="wo-edit" data-label="Notes" style={{ minWidth: 170 }}>
                      <WoCell multiline value="" placeholder="Add note…" onCommit={t => {
                        if (!t.trim()) return true;
                        update("workOrders", w.id, { notes: [...(w.notes || []), { date: todayISO(), timestamp: Date.now(), text: t.trim() }] });
                        return true;
                      }} />
                      {sortedNotes.map((n, i) => (
                        <div key={i} style={{ fontSize: 12, padding: "0 10px 4px" }}><span className="row-muted">{fmtDate(n.date)}</span> {n.text}</div>
                      ))}
                      {(w.photos || []).length > 0 && <div className="row-muted" style={{ padding: "0 10px 6px", fontSize: 12 }}>{(w.photos || []).length} photo{(w.photos || []).length === 1 ? "" : "s"}</div>}
                    </td>
                    <td className="wo-sheet-actions" onClick={(e) => e.stopPropagation()}>
                      <input type="checkbox" checked={selected.has(w.id)} onChange={() => toggleSelected(w.id)} title="Select for batch copy" />
                      <IconBtn
                        title={w.status === "Done" ? "Done — click to reopen" : "Mark done"}
                        active={w.status === "Done"}
                        onClick={() => update("workOrders", w.id, { status: w.status === "Done" ? "Open" : "Done" })}
                      ><CheckCircle2 size={14} /></IconBtn>
                      <IconBtn title={w.scheduledDate ? "Set a time" : "Schedule"} onClick={() => openSchedule(w)}><CalendarClock size={14} /></IconBtn>
                      <IconBtn title="Photos" onClick={() => setExpandedRow(isOpen ? null : w.id)}><StickyNote size={14} /></IconBtn>
                      <IconBtn title={copiedId === w.id ? "Copied!" : "Copy for texting/emailing"} onClick={() => copyOne(w)}><ScrollText size={14} /></IconBtn>
                      <IconBtn title="Link to a court case / more" onClick={() => { const f = { ...w, isCourtConnected: w.isCourtConnected || !!w.courtCaseId }; setFormOriginal(f); setForm(f); window.scrollTo({ top: 0, behavior: "smooth" }); }}><Pencil size={14} /></IconBtn>
                      <IconBtn title="Delete" danger onClick={() => remove("workOrders", w.id)}><Trash2 size={14} /></IconBtn>
                    </td>
                  </tr>
                  {schedulingFor === w.id && (
                    <tr className="wo-sheet-detail-row">
                      <td colSpan={11}>
                        <div className="inline-form" style={{ flexWrap: "wrap", alignItems: "center", margin: 0 }}>
                          <strong>Scheduled for</strong>
                          <input type="date" value={schedDate} onChange={e => setSchedDate(e.target.value)} style={{ flex: "0 0 auto", width: 170 }} />
                          <select value={schedTime} onChange={e => setSchedTime(e.target.value)}>
                            <option value="">No specific time</option>
                            {TIME_OPTIONS.map(t => <option key={t} value={t}>{fmtTime(t)}</option>)}
                          </select>
                          <button className="btn-primary" onClick={() => saveSchedule(w)} disabled={!schedDate}>Save</button>
                          {w.scheduledDate && <button className="btn-ghost" onClick={() => clearSchedule(w)}>Clear</button>}
                          <button className="btn-ghost" onClick={() => setSchedulingFor(null)}>Cancel</button>
                        </div>
                      </td>
                    </tr>
                  )}
                  {isOpen && (
                    <tr className="wo-sheet-detail-row">
                      <td colSpan={11}>
                        <PhotoUploader
                          photos={w.photos}
                          pathPrefix={`workOrders/${w.id}/photos`}
                          onAdd={(newPhotos) => update("workOrders", w.id, { photos: [...(w.photos || []), ...newPhotos] })}
                          onRemove={(p) => {
                            update("workOrders", w.id, { photos: (w.photos || []).filter(x => x.id !== p.id) });
                            if (p.storagePath) deleteObject(storageRef(storage, p.storagePath)).catch(() => {});
                          }}
                        />
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================== violations ============================== */

function ViolationsTab({ data, add, update, remove, buildingName, vendorName, setData, deepLink, setDeepLink }) {
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
  const [highlightViolationId, setHighlightViolationId] = useState(null);
  const fileRef = useRef(null);

  useEffect(() => {
    if (!deepLink || deepLink.tab !== "violations") return;
    const target = data.violations.find(v => v.id === deepLink.id);
    if (target) {
      setAgency("All");
      setView(isViolationClosed(target) ? "closed" : "active");
      setDueFilter("all");
      setBuildingFilter("All");
      setSearch("");
      setExpandedGroups(prev => new Set(prev).add(target.agency));
      setExpandedRow(target.id);
      setHighlightViolationId(target.id);
      setTimeout(() => {
        document.getElementById(`violation-row-${target.id}`)?.scrollIntoView({ behavior: "smooth", block: "center" });
      }, 100);
      setTimeout(() => setHighlightViolationId(null), 2500);
    }
    setDeepLink(null);
  }, [deepLink]);

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
      if (dueFilter === "lead") {
        l = l.filter(v => v.isLead);
      } else if (dueFilter !== "all") {
        const maxDays = dueFilter === "24h" ? 1 : dueFilter === "1w" ? 7 : 10;
        l = l.filter(v => { const d = daysUntil(v.cureDeadline); return d !== null && d <= maxDays && !v.isLead; });
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
      <div className={`list-card list-card-compact ${flag === "overdue" ? "list-card-danger" : flag === "soon" ? "list-card-warn" : ""} ${highlightViolationId === v.id ? "sheet-row-deeplinked" : ""}`} key={v.id} id={`violation-row-${v.id}`}>
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
          <IconBtn
            title={isViolationClosed(v) ? "Reopen" : `Mark ${violationClosedStatuses(v.agency)[0]}`}
            active={isViolationClosed(v)}
            onClick={(e) => { e.stopPropagation(); update("violations", v.id, { status: isViolationClosed(v) ? statusesFor(v.agency)[0] : violationClosedStatuses(v.agency)[0] }); }}
          ><CheckCircle2 size={14} /></IconBtn>
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
            {newestFirst(v.notes).map((n, i) => (
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

      {(() => {
        // Every open violation with a hearing, soonest court date first (no date last),
        // in one clear list so you don't have to dig through each agency section.
        const hearings = data.violations
          .filter(v => v.hasHearing && !isViolationClosed(v) && (buildingFilter === "All" || v.buildingId === buildingFilter))
          .slice()
          .sort((a, b) => (a.hearingDate || "9999-99-99").localeCompare(b.hearingDate || "9999-99-99"));
        if (hearings.length === 0) return null;
        const jump = (v) => {
          setAgency("All"); setView("active"); setDueFilter("all"); setSearch("");
          setExpandedGroups(prev => new Set(prev).add(v.agency));
          setExpandedRow(v.id); setHighlightViolationId(v.id);
          setTimeout(() => document.getElementById(`violation-row-${v.id}`)?.scrollIntoView({ behavior: "smooth", block: "center" }), 150);
          setTimeout(() => setHighlightViolationId(null), 2500);
        };
        return (
          <div className="form-panel" style={{ marginBottom: 12 }}>
            <div style={{ fontWeight: 700, marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
              <Gavel size={15} /> Hearings ({hearings.length}) <span className="row-muted" style={{ fontWeight: 400, fontSize: 12 }}>— in order of court date</span>
            </div>
            {hearings.map(v => {
              const d = v.hearingDate ? daysUntil(v.hearingDate) : null;
              const dt = v.hearingDate ? new Date(v.hearingDate + "T00:00:00") : null;
              const tone = d === null ? "var(--ink-soft)" : d < 0 ? "var(--danger)" : d <= 7 ? "var(--warn)" : "var(--navy)";
              const unit = v.unitId ? data.units.find(u => u.id === v.unitId) : null;
              return (
                <div key={v.id} onClick={() => jump(v)} style={{ display: "flex", gap: 12, alignItems: "flex-start", padding: "8px 0", borderTop: "1px solid var(--border)", cursor: "pointer" }}>
                  <div style={{ minWidth: 64, textAlign: "center", color: tone }}>
                    {dt ? (
                      <>
                        <div style={{ fontSize: 11, textTransform: "uppercase" }}>{dt.toLocaleDateString("en-US", { month: "short" })}</div>
                        <div style={{ fontSize: 22, fontWeight: 700, lineHeight: 1 }}>{dt.getDate()}</div>
                        <div style={{ fontSize: 11 }}>{dt.toLocaleDateString("en-US", { weekday: "short" })}</div>
                      </>
                    ) : <div style={{ fontSize: 12, fontWeight: 600 }}>No date</div>}
                    {d !== null && <div style={{ fontSize: 10, marginTop: 2 }}>{d < 0 ? `${-d}d ago` : d === 0 ? "today" : `in ${d}d`}</div>}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600 }}>{buildingName(v.buildingId)}{unit ? ` · Apt ${unit.unitNumber}` : ""}</div>
                    <div style={{ fontSize: 13 }}><span className="pill pill-muted" style={{ marginRight: 6 }}>{v.agency}</span>#{v.violationNumber}{v.class ? ` · Class ${v.class}` : ""}</div>
                    {v.description && <div className="row-muted" style={{ fontSize: 12, marginTop: 2 }}>{summarizeViolationDescription(v.description)}</div>}
                    {v.hearingCompany && <div style={{ fontSize: 12, marginTop: 2 }}>Handled by: {v.hearingCompany}</div>}
                  </div>
                </div>
              );
            })}
          </div>
        );
      })()}

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
          {agency === "HPD" && (
            <button className={`chip ${dueFilter === "lead" ? "chip-active" : ""}`} onClick={() => setDueFilter("lead")}>Lead</button>
          )}
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
              <Field label="Cure / correction deadline"><input type="date" value={form.cureDeadline} onChange={e => setForm({ ...form, cureDeadline: e.target.value })} /></Field>
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

  const openWO = (vid) => data.workOrders.filter(w => (w.vendorId === vid || (w.items && w.items.some(it => it.vendorId === vid))) && w.status !== "Done");
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
// A building's primary address plus any alternate ones stored for it — a
// corner-lot building can genuinely have two valid addresses (one per
// intersecting street), and HPD or another agency may register it under
// whichever one isn't the building's usual/primary address. No amount of
// abbreviation normalization can bridge two genuinely different street
// names, so this is explicit, stored knowledge rather than inferred.
function buildingAddressCandidates(b) {
  const extra = (b.alsoKnownAs || "").split(",").map(s => s.trim()).filter(Boolean);
  return [b.address, ...extra];
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
    const match = buildings.find(b => buildingAddressCandidates(b).some(addr => {
      const nb = normalizeAddrForMatch(addr);
      return nb.includes(nf) || nf.includes(nb);
    }));
    if (match) return match;
  }
  // Pass 2: street name only, ignoring the leading number — containment
  // rather than strict equality, since one side may include more of the
  // name than the other even after suffix normalization.
  for (const frag of realFragments) {
    const streetOnly = streetNameOnly(frag);
    if (!streetOnly) continue;
    const match = buildings.find(b => buildingAddressCandidates(b).some(addr => {
      const bStreet = streetNameOnly(addr);
      return bStreet && (bStreet.includes(streetOnly) || streetOnly.includes(bStreet));
    }));
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
    // Violation numbers already on file (re-checked now, not at preview time)
    // plus ones added in this very run — so the same violation can never be
    // saved twice, even if it appears twice in one report.
    const seenNumbers = new Set((data.violations || []).filter(ev => ev.agency === "HPD").map(ev => String(ev.violationNumber)));
    for (const row of preview.rows) {
      if (!row.include) { skipped++; continue; }
      if (seenNumbers.has(String(row.violationId))) { alreadyOnFile++; continue; }
      seenNumbers.add(String(row.violationId));
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
      // Already-certified HPD violations are saved in the slim shape (just
      // enough to recognise them on the next report) to save storage.
      if (fields.status === "Certified") {
        add("violations", { id: uid(), agency: "HPD", buildingId: fields.buildingId, unitId: fields.unitId, violationNumber: fields.violationNumber, status: "Certified" });
      } else {
        add("violations", { ...fields, id: uid(), fineAmount: "", company: "", otherAgency: "", vendorId: "", notes: [] });
      }
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
          ...fields, stage: "", nextCourtDate: parsedNextCourtDate, result: "Pending",
          stipulationTerms: "", nextPaymentDue: "", archived: false, documents: [],
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

function CourtTab({ data, add, update, remove, setData, tenantName, buildingName, deepLink, setDeepLink }) {
  const [form, setForm] = useState(null);
  // The case's fields exactly as they were when the edit form was opened, so Save
  // can send only what was actually changed in the form (see submit below).
  const [formOriginal, setFormOriginal] = useState(null);
  const [view, setView] = useState("active");
  const [detailsFor, setDetailsFor] = useState(null);
  const [section, setSection] = useState("cases"); // cases | import
  const [buildingFilter, setBuildingFilter] = useState("All");
  const [dueOnlyFilter, setDueOnlyFilter] = useState(null); // null | "court" | "stip"
  const [logForm, setLogForm] = useState({});
  const [confirmingDeleteAll, setConfirmingDeleteAll] = useState(false);
  const [search, setSearch] = useState("");
  const [highlightCaseId, setHighlightCaseId] = useState(null);

  useEffect(() => {
    if (!deepLink || deepLink.tab !== "court") return;
    const target = data.courtCases.find(c => c.id === deepLink.id);
    if (target) {
      setSection("cases");
      setView(target.archived ? "closed" : "active");
      setBuildingFilter("All");
      setDueOnlyFilter(null);
      setSearch("");
      setDetailsFor(target.id);
      setHighlightCaseId(target.id);
      setTimeout(() => {
        document.getElementById(`case-row-${target.id}`)?.scrollIntoView({ behavior: "smooth", block: "center" });
      }, 100);
      setTimeout(() => setHighlightCaseId(null), 2500);
    }
    setDeepLink(null);
  }, [deepLink]);

  const deleteAllCases = () => {
    setData(d => ({ ...d, courtCases: [] }));
    setConfirmingDeleteAll(false);
  };

  const addLogEntry = (c) => {
    const entry = logForm[c.id];
    if (!entry) return;
    if (!entry.quickType) entry.quickType = "other"; // plain note, no chip picked
    // "Other note" is the one type with no built-in label of its own — it
    // needs actual text typed in, since there's nothing else to log.
    if (entry.quickType === "other" && !(entry.note || "").trim()) return;
    if (entry.quickType === "courtDateSet" && !entry.newCourtDate) return;
    if (entry.quickType === "reminder" && (!entry.remDate || !(entry.note || "").trim())) return;

    const extra = (entry.note || "").trim();
    const fields = {};
    let noteText;
    if (entry.quickType === "adjourned") {
      noteText = "Adjourned" + (entry.newCourtDate ? ` — adjourned to ${fmtDate(entry.newCourtDate)}` : "");
      if (entry.newCourtDate) { fields.nextCourtDate = entry.newCourtDate; fields.result = "Adjourned / next date set"; }
    } else if (entry.quickType === "courtDateSet") {
      // The first court date (not an adjournment) — goes straight onto the calendar.
      noteText = `Court date set for ${fmtDate(entry.newCourtDate)}`;
      fields.nextCourtDate = entry.newCourtDate;
    } else if (entry.quickType === "reminder") {
      // A reminder about this case — shows on the dashboard and calendar on that date.
      noteText = `Reminder set for ${fmtDate(entry.remDate)}`;
      add("quickNotes", { id: uid(), text: `Court — ${tenantName(c.tenantId)}, ${buildingName(c.buildingId)}: ${extra}`, date: todayISO(), done: false, reminderDate: entry.remDate, buildingId: c.buildingId });
    } else if (entry.quickType === "stipulation") {
      noteText = "Stipulation agreed" + (entry.stipTerms ? ` — ${entry.stipTerms}` : "");
      fields.result = "Stipulation (payment plan)";
      if (entry.stipTerms) fields.stipulationTerms = entry.stipTerms;
      if (entry.stipNextDue) fields.nextPaymentDue = entry.stipNextDue;
    } else if (entry.quickType === "eviction") {
      noteText = "Eviction";
    } else if (entry.quickType === "dismissed") {
      noteText = "Case dismissed"; fields.result = "Case dismissed";
    } else if (entry.quickType === "judgmentLandlord") {
      noteText = "Judgment for landlord"; fields.result = "Judgment for landlord";
    } else if (entry.quickType === "judgmentTenant") {
      noteText = "Judgment for tenant"; fields.result = "Judgment for tenant";
    } else if (entry.quickType === "settled") {
      noteText = "Settled / withdrawn"; fields.result = "Settled / withdrawn";
    } else {
      noteText = extra;
    }
    if (extra && entry.quickType !== "other") noteText += ` — ${extra}`;

    const newEntry = { id: uid(), date: entry.date || todayISO(), note: noteText, source: "manual" };
    fields.log = [...(c.log || []), newEntry];
    update("courtCases", c.id, fields);
    setLogForm({ ...logForm, [c.id]: { date: todayISO(), note: "", newCourtDate: "", remDate: "", quickType: null, stipTerms: "", stipNextDue: "" } });
  };

  // A stage button was tapped. buildStage works out what to save on the case and
  // which dashboard reminder to set; applyStage saves it on an existing case.
  const buildStage = (c, st) => {
    const who = `${tenantName(c.tenantId)}, ${buildingName(c.buildingId)}`;
    const fields = { stage: st.label, stageDate: st.date };
    let note = null;
    if (st.kind === "lawyer") {
      Object.assign(fields, { sentLease: !!st.lease, sentLedger: !!st.ledger, sentToLawyerDate: st.date });
      note = { key: "petitionReminderId", text: `Confirm petition is being started — ${who}`, reminderDate: petitionReminderDate(st.date) };
    } else if (st.kind === "petition") {
      Object.assign(fields, { petitionServed: true, petitionServedDate: st.date });
      note = { key: "defaultReminderId", text: `Follow up with the attorney — did ${tenantName(c.tenantId)} answer the petition (served ${fmtDate(st.date)})? If not, start a default judgment — ${buildingName(c.buildingId)}`, reminderDate: defaultJudgmentReminderDate(st.date) };
    } else if (st.days > 0) {
      note = { key: "", text: `${st.label} (${fmtDate(st.date)}) — follow up: ${who}`, reminderDate: nextWeekdayISO(addDays(st.date, st.days)) };
    }
    const logEntry = { id: uid(), date: st.date, note: `Stage: ${st.label}${st.kind === "lawyer" ? ` (${[st.lease && "lease", st.ledger && "ledger"].filter(Boolean).join(" + ")})` : ""}`, source: "stage" };
    return { fields, note, logEntry };
  };
  const applyStage = (c, st) => {
    const { fields, note, logEntry } = buildStage(c, st);
    if (note) {
      const existing = note.key && c[note.key] ? (data.quickNotes || []).find(n => n.id === c[note.key]) : null;
      if (existing) { update("quickNotes", existing.id, { reminderDate: note.reminderDate, text: note.text, done: false }); fields[note.key] = existing.id; }
      else { const qid = uid(); add("quickNotes", { id: qid, text: note.text, date: todayISO(), done: false, reminderDate: note.reminderDate, buildingId: c.buildingId }); if (note.key) fields[note.key] = qid; }
    }
    fields.log = [...(c.log || []), logEntry];
    update("courtCases", c.id, fields);
  };

  const submit = () => {
    if (!form.tenantId) return; // a case with no tenant is a silent data-quality trap
    // Only patch the fields this form actually manages — documents and archived
    // get edited from their own controls elsewhere on the card, so if we sent the
    // whole `form` object here, saving this form would silently revert any change
    // made there back to whatever it was when this form was opened.
    const fields = {
      tenantId: form.tenantId, buildingId: form.buildingId, unitId: form.unitId, caseNumber: form.caseNumber,
      nextCourtDate: form.nextCourtDate, result: form.result,
      stipulationTerms: form.stipulationTerms || "", nextPaymentDue: form.nextPaymentDue || "",
    };
    const newCaseId = form.id || uid();
    if (form.id) {
      // Only what this form actually changed — "where it's up to" can also be set
      // from the card itself while this form is open, and sending every field
      // back would put that change back to what it was when the form opened.
      const patch = changedFieldsOnly(formOriginal, fields);
      if (Object.keys(patch).length > 0) update("courtCases", form.id, patch);
    }
    else {
      // A starting stage picked on the new-case form is applied as the case is created.
      const base = { ...fields, id: newCaseId, archived: false, documents: [] };
      if (form.pendingStage) {
        const { fields: sf, note, logEntry } = buildStage(base, form.pendingStage);
        Object.assign(base, sf, { log: [logEntry] });
        if (note) { const qid = uid(); add("quickNotes", { id: qid, text: note.text, date: todayISO(), done: false, reminderDate: note.reminderDate, buildingId: base.buildingId }); if (note.key) base[note.key] = qid; }
      }
      add("courtCases", base);
    }
    setForm(null);
  };

  // Old cases were saved before unitId existed — recover it from the tenant on
  // file so editing an old case still starts from the right apt.
  const openEdit = (c) => { const f = { ...c, unitId: c.unitId || data.tenants.find(t => t.id === c.tenantId)?.unitId || "" }; setFormOriginal(f); setForm(f); };

  const list = data.courtCases
    .filter(c => view === "closed" ? c.archived : !c.archived)
    .filter(c => buildingFilter === "All" || c.buildingId === buildingFilter)
    .filter(c => {
      if (dueOnlyFilter === "court") return c.result === "Stipulation (payment plan)" ? dateDueStrict(c.nextCourtDate) : dateDue(c.nextCourtDate);
      if (dueOnlyFilter === "stip") return c.result === "Stipulation (payment plan)" && dateDue(c.nextPaymentDue);
      return true;
    })
    .filter(c => {
      if (!search.trim()) return true;
      const q = search.trim().toLowerCase();
      const name = c.tenantId ? tenantName(c.tenantId) : (c.rawName || "");
      return [name, c.caseNumber, buildingName(c.buildingId), c.stage, c.result, c.stipulationTerms]
        .some(f => (f || "").toString().toLowerCase().includes(q));
    })
    .slice()
    .sort((a, b) => (a.nextCourtDate || "9999-99-99").localeCompare(b.nextCourtDate || "9999-99-99"));

  // Buildings that actually have at least one active case, so the filter
  // row only ever shows relevant options, not every building regardless
  // of whether it has any cases.
  const buildingsWithCases = data.buildings.filter(b => data.courtCases.some(c => c.buildingId === b.id && !c.archived));
  const courtDatesDueCount = data.courtCases.filter(c => !c.archived && (c.result === "Stipulation (payment plan)" ? dateDueStrict(c.nextCourtDate) : dateDue(c.nextCourtDate))).length;
  const stipDueCount = data.courtCases.filter(c => !c.archived && c.result === "Stipulation (payment plan)" && dateDue(c.nextPaymentDue)).length;


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
          <button className="btn-primary" onClick={() => setForm({ tenantId: "", buildingId: "", unitId: "", caseNumber: "", nextCourtDate: "", result: "Pending", stage: "", stipulationTerms: "", nextPaymentDue: "", pendingStage: null })}>
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

      {section === "cases" && (
        <div className="row" style={{ marginBottom: 10 }}>
          <input
            type="text" placeholder="Search tenant, case #, building, stage, result…" value={search}
            onChange={e => setSearch(e.target.value)}
            style={{ width: "100%", maxWidth: 360 }}
          />
        </div>
      )}

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
        {(() => {
          // Balance comes from the tenant's record; the two newest log entries are the
          // last two updates entered on the case.
          const owedOf = (c) => { const t = c.tenantId ? data.tenants.find(x => x.id === c.tenantId) : null; return t ? parseBalance(t.balance) : 0; };
          const lastTwo = (c) => (c.log || []).map((l, i) => ({ ...l, _i: i })).sort((x, y) => (y.date || "").localeCompare(x.date || "") || y._i - x._i).slice(0, 2);
          const aptOf = (c) => {
            const t = c.tenantId ? data.tenants.find(x => x.id === c.tenantId) : null;
            const u = data.units.find(x => x.id === (c.unitId || (t && t.unitId)));
            return u ? u.unitNumber : "—";
          };
          const grand = list.reduce((sum, c) => sum + owedOf(c), 0);
          return (
            <>
              <div className="print-stats-row">
                <div className="print-stat"><div className="print-stat-num">{list.length}</div><div>Total cases</div></div>
                <div className="print-stat"><div className="print-stat-num">{money(grand)}</div><div>Total owed</div></div>
              </div>
              {[...new Set(list.map(c => c.buildingId))].map(bId => {
                const items = list.filter(c => c.buildingId === bId).slice().sort((x, y) => owedOf(y) - owedOf(x));
                const bTotal = items.reduce((sum, c) => sum + owedOf(c), 0);
                return (
                  <div className="print-section" key={bId || "unknown"}>
                    <div className="print-section-head">
                      <span>{buildingName(bId)}</span>
                      <span>{items.length} case{items.length === 1 ? "" : "s"} · {money(bTotal)} owed</span>
                    </div>
                    <table className="print-table">
                      <thead><tr><th style={{ width: "7%" }}>Apt</th><th style={{ width: "20%" }}>Tenant</th><th style={{ width: "14%", textAlign: "right" }}>Owes</th><th>Last 2 updates</th></tr></thead>
                      <tbody>
                        {items.map(c => {
                          const ups = lastTwo(c);
                          return (
                            <tr key={c.id} style={{ breakInside: "avoid" }}>
                              <td style={{ fontWeight: 600 }}>{aptOf(c)}</td>
                              <td style={{ fontWeight: 600 }}>{c.tenantId ? tenantName(c.tenantId) : (c.rawName || "(no tenant matched)")}</td>
                              <td className="print-table-amount">{money(owedOf(c))}</td>
                              <td>
                                {ups.length === 0
                                  ? <span style={{ color: "#888" }}>No updates yet</span>
                                  : ups.map(u => <div key={u.id || u._i} style={{ marginBottom: 3 }}><strong>{fmtDate(u.date)}</strong> — {u.note}</div>)}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                );
              })}
              <div className="print-grand-total">Total owed across all cases: {money(grand)}</div>
            </>
          );
        })()}
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
          <Field label="Case #"><input value={form.caseNumber} onChange={e => setForm({ ...form, caseNumber: e.target.value })} /></Field>
          <Field label="Court date"><input type="date" value={form.nextCourtDate} onChange={e => setForm({ ...form, nextCourtDate: e.target.value })} /></Field>
          {!form.id && (
            <Field label="Starting stage">
              <CaseStageBar
                options={data.caseStageOptions}
                current={form.pendingStage ? { label: form.pendingStage.label, date: form.pendingStage.date } : null}
                onApply={(st) => setForm({ ...form, pendingStage: st })}
                onAddOption={(o) => add("caseStageOptions", o)}
                onRemoveOption={(o) => remove("caseStageOptions", o.id)}
                onClear={() => setForm({ ...form, pendingStage: null })}
              />
            </Field>
          )}
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
        // Newest date first; entries sharing a date keep the order they were added in
      // (the most recently added on top), instead of falling out in arbitrary order.
      const sortedLog = (c.log || []).map((l, i) => ({ ...l, _i: i })).sort((a, b) => (b.date || "").localeCompare(a.date || "") || b._i - a._i);
        const latestLog = sortedLog[0];
        const isOpen = detailsFor === c.id;
        return (
        <div className={`list-card ${highlightCaseId === c.id ? "sheet-row-deeplinked" : ""}`} key={c.id} id={`case-row-${c.id}`}>
          <div className="list-card-head" onClick={() => setDetailsFor(isOpen ? null : c.id)} style={{ cursor: "pointer" }}>
            {isOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
            {c.unitId && <span className="pill pill-accent">Apt {data.units.find(u => u.id === c.unitId)?.unitNumber || "—"}</span>}
            <div className="list-card-title">{c.tenantId ? tenantName(c.tenantId) : (c.rawName || "(no tenant matched)")} {c.caseNumber && `· Case #${c.caseNumber}`}</div>
            {!c.caseNumber && <CaseNumberQuickAdd onSave={(t) => update("courtCases", c.id, { caseNumber: t })} />}
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
                <div className="inline-form" style={{ marginTop: 8, flexWrap: "wrap" }}>
                  {[
                    { key: "courtDateSet", label: "Court date set" },
                    { key: "adjourned", label: "Adjourned" },
                    { key: "stipulation", label: "Stipulation" },
                    { key: "eviction", label: "Eviction" },
                    { key: "dismissed", label: "Case dismissed" },
                    { key: "judgmentLandlord", label: "Judgment (landlord)" },
                    { key: "judgmentTenant", label: "Judgment (tenant)" },
                    { key: "settled", label: "Settled / withdrawn" },
                    { key: "reminder", label: "Set reminder" },
                    { key: "other", label: "Other note" },
                  ].map(opt => (
                    <button
                      key={opt.key} type="button"
                      className={`chip ${logForm[c.id]?.quickType === opt.key ? "chip-active" : ""}`}
                      onClick={() => setLogForm({ ...logForm, [c.id]: { ...logForm[c.id], quickType: opt.key } })}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
                {(
                  <div style={{ marginTop: 6, display: "flex", flexDirection: "column", gap: 6 }}>
                    {logForm[c.id]?.quickType === "courtDateSet" && (
                      <label className="hint" style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        Court date:
                        <input
                          type="date" value={logForm[c.id]?.newCourtDate || ""}
                          onChange={e => setLogForm({ ...logForm, [c.id]: { ...logForm[c.id], newCourtDate: e.target.value } })}
                        />
                      </label>
                    )}
                    {logForm[c.id]?.quickType === "reminder" && (
                      <label className="hint" style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        Remind me on:
                        <input
                          type="date" value={logForm[c.id]?.remDate || ""}
                          onChange={e => setLogForm({ ...logForm, [c.id]: { ...logForm[c.id], remDate: e.target.value } })}
                        />
                        <span>— type what for below</span>
                      </label>
                    )}
                    {logForm[c.id]?.quickType === "adjourned" && (
                      <label className="hint" style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        Adjourned to:
                        <input
                          type="date" value={logForm[c.id]?.newCourtDate || ""}
                          onChange={e => setLogForm({ ...logForm, [c.id]: { ...logForm[c.id], newCourtDate: e.target.value } })}
                        />
                      </label>
                    )}
                    {logForm[c.id]?.quickType === "stipulation" && (
                      <>
                        <textarea
                          placeholder="Terms — e.g. $200/month starting 10/1 on top of current rent, for 6 months"
                          value={logForm[c.id]?.stipTerms || ""}
                          onChange={e => setLogForm({ ...logForm, [c.id]: { ...logForm[c.id], stipTerms: e.target.value } })}
                        />
                        <label className="hint" style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          Next payment due:
                          <input
                            type="date" value={logForm[c.id]?.stipNextDue || ""}
                            onChange={e => setLogForm({ ...logForm, [c.id]: { ...logForm[c.id], stipNextDue: e.target.value } })}
                          />
                        </label>
                      </>
                    )}
                    <div className="inline-form" style={{ flexWrap: "wrap" }}>
                      <input type="date" value={logForm[c.id]?.date || todayISO()} onChange={e => setLogForm({ ...logForm, [c.id]: { ...logForm[c.id], date: e.target.value } })} />
                      <input placeholder={logForm[c.id]?.quickType === "reminder" ? "What should I remind you about?" : logForm[c.id]?.quickType && logForm[c.id]?.quickType !== "other" ? "Any extra detail (optional)…" : "Type a note / update…"} value={logForm[c.id]?.note || ""} onChange={e => setLogForm({ ...logForm, [c.id]: { ...logForm[c.id], note: e.target.value } })} onKeyDown={e => e.key === "Enter" && addLogEntry(c)} />
                      <button className="btn-ghost" onClick={() => addLogEntry(c)}>Add</button>
                    </div>
                  </div>
                )}
              </div>
              <div className="list-card-body" style={{ paddingTop: 0 }}>
                {(c.sentLease || c.sentLedger) && (
                  <div className="row row-muted" style={{ marginBottom: 6 }}>
                    Sent to lawyer: {[c.sentLease && "Lease", c.sentLedger && "Ledger"].filter(Boolean).join(" + ")}{c.sentToLawyerDate ? ` on ${fmtDate(c.sentToLawyerDate)} · petition check ${fmtDate(petitionReminderDate(c.sentToLawyerDate))}` : ""}
                  </div>
                )}
                {c.petitionServed && (
                  <div className="row row-muted" style={{ marginBottom: 6 }}>
                    Petition served{c.petitionServedDate ? ` on ${fmtDate(c.petitionServedDate)} · default judgment check ${fmtDate(defaultJudgmentReminderDate(c.petitionServedDate))}` : ""}
                  </div>
                )}
                <CaseStageBar
                  options={data.caseStageOptions}
                  current={{ label: c.stage, date: c.stageDate }}
                  onApply={(st) => applyStage(c, st)}
                  onAddOption={(o) => add("caseStageOptions", o)}
                  onRemoveOption={(o) => remove("caseStageOptions", o.id)}
                  onClear={() => update("courtCases", c.id, { stage: "", stageDate: "" })}
                />
              </div>
              {c.result === "Stipulation (payment plan)" && (
                <div className="list-card-body">
                  {c.stipulationTerms && <div className="row"><em>Terms:</em> {c.stipulationTerms}</div>}
                  {c.nextPaymentDue && <div className="row">Next payment due: <Flag date={c.nextPaymentDue} /></div>}
                </div>
              )}
              <div className="list-card-body" style={{ paddingTop: c.result === "Stipulation (payment plan)" ? 0 : undefined }}>
              <>
                <div className="row"><strong>Documents</strong></div>
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

/* ============================== emergency repair orders (ERO) ============================== */

function SmokeCoTab({ data, add, update, remove, buildingName }) {
  // Deliberately simple: building, apartment, date installed. Not linked to
  // tenants and not shown on the Dashboard.
  const [buildingId, setBuildingId] = useState("");
  const [unitId, setUnitId] = useState("");
  const [installed, setInstalled] = useState(todayISO());
  const [filter, setFilter] = useState("All");
  const unitsFor = (bid) => data.units.filter(u => u.buildingId === bid).sort((a, b) => compareUnits(a.unitNumber, b.unitNumber));
  const unitLabel = (id) => data.units.find(u => u.id === id)?.unitNumber || "—";
  const canAdd = buildingId && unitId && installed;
  const addOne = () => {
    if (!canAdd) return;
    add("smokeCoDetectors", { id: uid(), buildingId, unitId, installedDate: installed });
    setUnitId(""); // keep building + date so a whole building can be entered quickly
  };
  const all = data.smokeCoDetectors || [];
  const shown = all.filter(r => filter === "All" || r.buildingId === filter);
  // Spreadsheet order: by building (in the order buildings are listed), then apartment.
  const bIndex = (id) => data.buildings.findIndex(b => b.id === id);
  const sortedRows = shown.slice().sort((x, y) => (bIndex(x.buildingId) - bIndex(y.buildingId)) || compareUnits(unitLabel(x.unitId), unitLabel(y.unitId)));
  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Smoke / CO Detectors</h1>
        <div className="page-actions"><PrintButton label="Smoke / CO Detectors" /></div>
      </div>
      <div className="form-panel" style={{ marginBottom: 12 }}>
        <div className="inline-form" style={{ flexWrap: "wrap" }}>
          <select value={buildingId} onChange={e => { setBuildingId(e.target.value); setUnitId(""); }}>
            <option value="">Building…</option>
            {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
          </select>
          <select value={unitId} onChange={e => setUnitId(e.target.value)} disabled={!buildingId}>
            <option value="">Apt…</option>
            {unitsFor(buildingId).map(u => <option key={u.id} value={u.id}>{u.unitNumber}</option>)}
          </select>
          <label className="hint" style={{ display: "flex", alignItems: "center", gap: 6 }}>
            Installed:
            <input type="date" value={installed} onChange={e => setInstalled(e.target.value)} />
          </label>
          <button className="btn-primary" disabled={!canAdd} onClick={addOne}><Plus size={14} /> Add</button>
        </div>
      </div>
      <div className="row" style={{ marginBottom: 10 }}>
        <select value={filter} onChange={e => setFilter(e.target.value)} style={{ maxWidth: 220 }}>
          <option value="All">All buildings</option>
          {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
        </select>
      </div>
      {shown.length === 0 && <EmptyState text="No smoke / CO installs logged yet." />}
      {shown.length > 0 && (
        <div className="sheet-wrap">
          <table className="sheet">
            <thead>
              <tr>
                <th>Building</th>
                <th>Apt</th>
                <th>Date installed</th>
                <th style={{ width: 40 }}></th>
              </tr>
            </thead>
            <tbody>
              {sortedRows.map(r => (
                <tr key={r.id}>
                  <td data-label="Building" style={{ fontWeight: 600 }}>{buildingName(r.buildingId)}</td>
                  <td data-label="Apt">{unitLabel(r.unitId)}</td>
                  <td data-label="Date installed">{r.installedDate ? fmtDate(r.installedDate) : ""}</td>
                  <td style={{ textAlign: "center" }}>
                    <IconBtn title="Delete this row" danger onClick={() => { if (window.confirm("Delete this entry?")) remove("smokeCoDetectors", r.id); }}><Trash2 size={14} /></IconBtn>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function EroTab({ data, add, update, remove, buildingName, vendorName, deepLink, setDeepLink }) {
  const [form, setForm] = useState(null);
  const [formOriginal, setFormOriginal] = useState(null);
  const [view, setView] = useState("active");
  const [buildingFilter, setBuildingFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [highlightId, setHighlightId] = useState(null);
  const [noteDrafts, setNoteDrafts] = useState({});
  const list0 = data.emergencyRepairs || [];
  const addEroNote = (e) => {
    const text = (noteDrafts[e.id] || "").trim();
    if (!text) return;
    // Sends only the notes list, so adding a note can't put back anything else
    // that was changed on this ERO from another device.
    update("emergencyRepairs", e.id, { notes: [...(e.notes || []), { id: uid(), date: todayISO(), timestamp: Date.now(), text }] });
    setNoteDrafts({ ...noteDrafts, [e.id]: "" });
  };

  useEffect(() => {
    if (!deepLink || deepLink.tab !== "ero") return;
    const target = list0.find(e => e.id === deepLink.id);
    if (target) {
      setView(target.status === "Completed" ? "done" : "active");
      setBuildingFilter("All"); setSearch("");
      setHighlightId(target.id);
      setTimeout(() => document.getElementById(`ero-row-${target.id}`)?.scrollIntoView({ behavior: "smooth", block: "center" }), 100);
      setTimeout(() => setHighlightId(null), 2500);
    }
    setDeepLink(null);
  }, [deepLink]);

  const blank = () => ({
    id: uid(), eroNumber: "", buildingId: data.buildings[0]?.id || "", unitId: "", violationIds: [],
    description: "", vendorId: "", handledBy: "", dateIssued: todayISO(), dueByDate: "",
    hpdContacted: false, hpdContactedDate: "", status: "Open", completedDate: "",
  });
  const openEdit = (e) => { const f = { violationIds: [], ...e }; setForm(f); setFormOriginal(f); };

  const submit = () => {
    if (!form.eroNumber.trim() && !form.description.trim()) return;
    const exists = list0.some(e => e.id === form.id);
    const f = { ...form };
    if (f.hpdContacted && !f.hpdContactedDate) f.hpdContactedDate = todayISO();
    if (!f.hpdContacted) f.hpdContactedDate = "";
    if (exists) {
      const patch = changedFieldsOnly(formOriginal, f);
      if (Object.keys(patch).length > 0) update("emergencyRepairs", f.id, patch);
    } else add("emergencyRepairs", f);
    setForm(null);
  };

  const list = list0
    .filter(e => view === "active" ? e.status !== "Completed" : e.status === "Completed")
    .filter(e => buildingFilter === "All" || e.buildingId === buildingFilter)
    .filter(e => {
      if (!search.trim()) return true;
      const q = search.trim().toLowerCase();
      const vNums = (e.violationIds || []).map(id => data.violations.find(v => v.id === id)?.violationNumber || "").join(" ");
      return [e.eroNumber, e.description, buildingName(e.buildingId), e.vendorId ? vendorName(e.vendorId) : "", e.handledBy, vNums]
        .some(x => (x || "").toString().toLowerCase().includes(q));
    })
    .slice()
    .sort((a, b) => (a.dueByDate || "9999-99-99").localeCompare(b.dueByDate || "9999-99-99"));

  const formViolations = form ? data.violations.filter(v => v.buildingId === form.buildingId && (!isViolationClosed(v) || (form.violationIds || []).includes(v.id))) : [];
  const toggleViolation = (id) => {
    const cur = form.violationIds || [];
    setForm({ ...form, violationIds: cur.includes(id) ? cur.filter(x => x !== id) : [...cur, id] });
  };

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Emergency Repair Orders</h1>
        <div className="page-actions">
          <PrintButton label="Emergency Repair Orders" />
          <button className="btn-primary" onClick={() => { const f = blank(); setForm(f); setFormOriginal(null); }}>
            <Plus size={14} /> Add ERO
          </button>
        </div>
      </div>
      <div className="row" style={{ marginBottom: 10, flexWrap: "wrap" }}>
        <input type="text" placeholder="Search ERO #, violation #, work, building, who…" value={search} onChange={e => setSearch(e.target.value)} style={{ width: "100%", maxWidth: 360 }} />
        <select value={buildingFilter} onChange={e => setBuildingFilter(e.target.value)} style={{ maxWidth: 220 }}>
          <option value="All">All buildings</option>
          {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
        </select>
      </div>
      <div className="filter-row">
        <button className={`chip ${view === "active" ? "chip-active" : ""}`} onClick={() => setView("active")}>Open</button>
        <button className={`chip ${view === "done" ? "chip-active" : ""}`} onClick={() => setView("done")}>Completed</button>
      </div>

      {form && (
        <div className="form-panel">
          <Field label="ERO number"><input value={form.eroNumber} onChange={e => setForm({ ...form, eroNumber: e.target.value })} placeholder="ERO #" /></Field>
          <Field label="Building">
            <select value={form.buildingId} onChange={e => setForm({ ...form, buildingId: e.target.value, unitId: "", violationIds: [] })}>
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
          <Field label="Connected violation(s)">
            {formViolations.length === 0
              ? <div className="hint">No violations on file for this building.</div>
              : <div style={{ display: "flex", flexDirection: "column", gap: 4, maxHeight: 180, overflowY: "auto" }}>
                  {formViolations.map(v => (
                    <label key={v.id} style={{ display: "flex", gap: 8, alignItems: "flex-start", fontSize: 13 }}>
                      <input type="checkbox" checked={(form.violationIds || []).includes(v.id)} onChange={() => toggleViolation(v.id)} />
                      <span>#{v.violationNumber} <span className="row-muted">{v.agency}{v.description ? ` — ${v.description.slice(0, 70)}` : ""}</span></span>
                    </label>
                  ))}
                </div>}
          </Field>
          <Field label="What the work is"><textarea value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} /></Field>
          <Field label="Who's taking care of it (vendor)">
            <select value={form.vendorId || ""} onChange={e => setForm({ ...form, vendorId: e.target.value })}>
              <option value="">Not assigned</option>
              {data.vendors.map(v => <option key={v.id} value={v.id}>{v.name}</option>)}
            </select>
          </Field>
          <Field label="…or someone else (super, in-house, etc.)"><input value={form.handledBy || ""} onChange={e => setForm({ ...form, handledBy: e.target.value })} /></Field>
          <Field label="ERO issued"><input type="date" value={form.dateIssued || ""} onChange={e => setForm({ ...form, dateIssued: e.target.value })} /></Field>
          <Field label="Needs to be done by"><input type="date" value={form.dueByDate || ""} onChange={e => setForm({ ...form, dueByDate: e.target.value })} /></Field>
          <label style={{ display: "flex", gap: 8, alignItems: "center", margin: "6px 0" }}>
            <input type="checkbox" checked={!!form.hpdContacted} onChange={e => setForm({ ...form, hpdContacted: e.target.checked })} />
            Reached out to HPD
          </label>
          <div className="form-actions">
            <button className="btn-primary" onClick={submit}>Save</button>
            <button className="btn-ghost" onClick={() => setForm(null)}>Cancel</button>
          </div>
        </div>
      )}

      {list.length === 0 && <EmptyState text={view === "active" ? "No open emergency repair orders." : "Nothing marked completed yet."} />}
      {list.map(e => {
        const unit = e.unitId ? data.units.find(u => u.id === e.unitId) : null;
        const vNums = (e.violationIds || []).map(id => data.violations.find(v => v.id === id)).filter(Boolean);
        const who = e.vendorId ? vendorName(e.vendorId) : (e.handledBy || "");
        const done = e.status === "Completed";
        return (
          <div className={`list-card ${highlightId === e.id ? "sheet-row-deeplinked" : ""}`} key={e.id} id={`ero-row-${e.id}`}>
            <div className="list-card-head">
              <div className="list-card-title">ERO {e.eroNumber || "(no number)"} <span className="row-muted">— {buildingName(e.buildingId)}{unit ? ` · Apt ${unit.unitNumber}` : ""}</span></div>
              {done
                ? <span className="pill pill-muted">Completed {fmtDate(e.completedDate)}</span>
                : e.dueByDate && <Flag date={e.dueByDate} label="due" />}
              <IconBtn title="Edit" onClick={() => openEdit(e)}><Pencil size={14} /></IconBtn>
              <IconBtn title="Delete" danger onClick={() => { if (window.confirm("Delete this ERO?")) remove("emergencyRepairs", e.id); }}><Trash2 size={14} /></IconBtn>
            </div>
            <div className="list-card-body">
              {e.description && <div className="row">{e.description}</div>}
              {vNums.length > 0 && <div className="row"><em>Violations:</em> {vNums.map(v => `#${v.violationNumber}`).join(", ")}</div>}
              <div className="row"><em>Handled by:</em> {who || "not assigned"}</div>
              <div className="row"><em>Issued:</em> {e.dateIssued ? fmtDate(e.dateIssued) : "—"}{e.dueByDate ? <> · <em>Due:</em> {fmtDate(e.dueByDate)}</> : null}</div>
              <div className="row" style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
                <label style={{ display: "flex", gap: 6, alignItems: "center" }}>
                  <input type="checkbox" checked={!!e.hpdContacted}
                    onChange={ev => update("emergencyRepairs", e.id, { hpdContacted: ev.target.checked, hpdContactedDate: ev.target.checked ? todayISO() : "" })} />
                  Reached out to HPD{e.hpdContacted && e.hpdContactedDate ? ` (${fmtDate(e.hpdContactedDate)})` : ""}
                </label>
                {!done
                  ? <button className="btn-ghost" onClick={() => update("emergencyRepairs", e.id, { status: "Completed", completedDate: todayISO() })}><CheckCircle2 size={14} /> Mark completed</button>
                  : <button className="btn-ghost" onClick={() => update("emergencyRepairs", e.id, { status: "Open", completedDate: "" })}>Reopen</button>}
              </div>
              <div className="row" style={{ marginTop: 6 }}><strong>Notes</strong></div>
              {(e.notes || []).length === 0 && <div className="hint">No notes yet.</div>}
              {newestFirst(e.notes).map((n, i) => (
                <div className="row row-muted" key={n.id || i} style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
                  <span style={{ minWidth: 90, fontSize: 12 }}>{fmtDate(n.date)}</span>
                  <span style={{ flex: 1 }}>{n.text}</span>
                  <button className="checklist-remove" title="Remove this note" onClick={() => update("emergencyRepairs", e.id, { notes: (e.notes || []).filter(x => x !== n) })}><X size={12} /></button>
                </div>
              ))}
              <div className="inline-form" style={{ marginTop: 6 }}>
                <input placeholder="Add a note — calls, HPD contact, vendor updates…" value={noteDrafts[e.id] || ""}
                  onChange={ev => setNoteDrafts({ ...noteDrafts, [e.id]: ev.target.value })}
                  onKeyDown={ev => ev.key === "Enter" && addEroNote(e)} />
                <button className="btn-ghost" onClick={() => addEroNote(e)}>Add note</button>
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ============================== collections (former tenants, owed money) ============================== */

function CollectionsTab({ data, add, update, remove, buildingName, deepLink, setDeepLink }) {
  const [form, setForm] = useState(null);
  const [formOriginal, setFormOriginal] = useState(null);
  const [view, setView] = useState("active");
  const [buildingFilter, setBuildingFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [highlightId, setHighlightId] = useState(null);
  const [openId, setOpenId] = useState(null);
  const [logForm, setLogForm] = useState({});
  const all = data.collections || [];

  useEffect(() => {
    if (!deepLink || deepLink.tab !== "collections") return;
    const target = all.find(c => c.id === deepLink.id);
    if (target) {
      setView(target.archived ? "closed" : "active");
      setBuildingFilter("All"); setSearch("");
      setOpenId(target.id); setHighlightId(target.id);
      setTimeout(() => document.getElementById(`coll-row-${target.id}`)?.scrollIntoView({ behavior: "smooth", block: "center" }), 100);
      setTimeout(() => setHighlightId(null), 2500);
    }
    setDeepLink(null);
  }, [deepLink]);

  const blank = () => ({
    id: uid(), tenantId: "", tenantName: "", buildingId: data.buildings[0]?.id || "", unitId: "",
    amountOwed: "", movedOutDate: "", lawyer: "", stage: "", log: [], archived: false,
  });
  const openEdit = (c) => { setForm({ ...c }); setFormOriginal({ ...c }); };

  const pickTenant = (id) => {
    const t = data.tenants.find(x => x.id === id);
    if (!t) { setForm({ ...form, tenantId: "" }); return; }
    setForm({
      ...form, tenantId: t.id, tenantName: t.name || form.tenantName,
      buildingId: t.buildingId || form.buildingId, unitId: t.unitId || "",
      amountOwed: form.amountOwed || (parseBalance(t.balance) > 0 ? String(parseBalance(t.balance).toFixed(2)) : ""),
    });
  };

  const submit = () => {
    if (!form.tenantName.trim()) return;
    if (formOriginal) {
      const patch = changedFieldsOnly(formOriginal, form);
      if (Object.keys(patch).length > 0) update("collections", form.id, patch);
    } else add("collections", form);
    setForm(null); setFormOriginal(null);
  };

  const addLog = (c) => {
    const e = logForm[c.id];
    if (!e || !(e.note || "").trim()) return;
    update("collections", c.id, { log: [...(c.log || []), { id: uid(), date: e.date || todayISO(), note: e.note.trim() }] });
    setLogForm({ ...logForm, [c.id]: { date: todayISO(), note: "" } });
  };

  const list = all
    .filter(c => view === "active" ? !c.archived : c.archived)
    .filter(c => buildingFilter === "All" || c.buildingId === buildingFilter)
    .filter(c => {
      if (!search.trim()) return true;
      const q = search.trim().toLowerCase();
      return [c.tenantName, c.lawyer, c.stage, buildingName(c.buildingId)].some(x => (x || "").toLowerCase().includes(q));
    })
    .slice()
    .sort((a, b) => parseBalance(b.amountOwed) - parseBalance(a.amountOwed));
  const totalOwed = list.reduce((s, c) => s + parseBalance(c.amountOwed), 0);
  const unitsFor = form ? data.units.filter(u => u.buildingId === form.buildingId).sort((a, b) => compareUnits(a.unitNumber, b.unitNumber)) : [];
  const tenantsFor = form ? data.tenants.filter(t => t.buildingId === form.buildingId) : [];

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Collections</h1>
        <div className="page-actions">
          <PrintButton label="Collections" />
          <button className="btn-primary" onClick={() => { setForm(blank()); setFormOriginal(null); }}><Plus size={14} /> Add collection</button>
        </div>
      </div>
      <p className="hint">Former tenants who moved out owing money — how much, who's handling it, and where it's up to. Uses the same "where it's up to" options as Court Cases.</p>
      <div className="row" style={{ marginBottom: 10, flexWrap: "wrap" }}>
        <input type="text" placeholder="Search name, lawyer, stage, building…" value={search} onChange={e => setSearch(e.target.value)} style={{ width: "100%", maxWidth: 360 }} />
        <select value={buildingFilter} onChange={e => setBuildingFilter(e.target.value)} style={{ maxWidth: 220 }}>
          <option value="All">All buildings</option>
          {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
        </select>
      </div>
      <div className="filter-row">
        <button className={`chip ${view === "active" ? "chip-active" : ""}`} onClick={() => setView("active")}>Active</button>
        <button className={`chip ${view === "closed" ? "chip-active" : ""}`} onClick={() => setView("closed")}>Closed</button>
        <span className="hint" style={{ marginLeft: 8 }}>{list.length} · {money(totalOwed)} owed</span>
      </div>

      {form && (
        <div className="form-panel">
          <Field label="Building">
            <select value={form.buildingId} onChange={e => setForm({ ...form, buildingId: e.target.value, unitId: "", tenantId: "" })}>
              <option value="">—</option>
              {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
            </select>
          </Field>
          <Field label="Pick from tenants on file (optional — fills in name and balance)">
            <select value={form.tenantId || ""} onChange={e => pickTenant(e.target.value)}>
              <option value="">—</option>
              {tenantsFor.map(t => <option key={t.id} value={t.id}>{t.name || "(no name)"}{t.movedOut ? " (moved out)" : ""}</option>)}
            </select>
          </Field>
          <Field label="Name"><input value={form.tenantName} onChange={e => setForm({ ...form, tenantName: e.target.value })} /></Field>
          <Field label="Apt # (optional)">
            <select value={form.unitId || ""} onChange={e => setForm({ ...form, unitId: e.target.value })}>
              <option value="">—</option>
              {unitsFor.map(u => <option key={u.id} value={u.id}>{u.unitNumber}</option>)}
            </select>
          </Field>
          <Field label="Amount owed ($)"><input inputMode="decimal" value={form.amountOwed} onChange={e => setForm({ ...form, amountOwed: e.target.value })} /></Field>
          <Field label="Date moved out"><input type="date" value={form.movedOutDate || ""} onChange={e => setForm({ ...form, movedOutDate: e.target.value })} /></Field>
          <Field label="Lawyer handling it"><input value={form.lawyer || ""} onChange={e => setForm({ ...form, lawyer: e.target.value })} /></Field>
          <Field label="Where it's up to">
            <CaseStagePicker
              options={data.caseStageOptions} value={form.stage || ""}
              onPick={(label) => setForm({ ...form, stage: label })}
              onAdd={(label) => add("caseStageOptions", { label })}
              onRemove={(o) => remove("caseStageOptions", o.id)}
            />
          </Field>
          <div className="form-actions">
            <button className="btn-primary" onClick={submit}>Save</button>
            <button className="btn-ghost" onClick={() => { setForm(null); setFormOriginal(null); }}>Cancel</button>
          </div>
        </div>
      )}

      {list.length === 0 && <EmptyState text={view === "active" ? "No active collections." : "Nothing closed yet."} />}
      {list.map(c => {
        const unit = c.unitId ? data.units.find(u => u.id === c.unitId) : null;
        const open = openId === c.id;
        const log = (c.log || []).map((l, i) => ({ ...l, _i: i })).sort((a, b) => (b.date || "").localeCompare(a.date || "") || b._i - a._i);
        const lf = logForm[c.id] || { date: todayISO(), note: "" };
        return (
          <div className={`list-card ${highlightId === c.id ? "sheet-row-deeplinked" : ""}`} key={c.id} id={`coll-row-${c.id}`}>
            <div className="list-card-head" onClick={() => setOpenId(open ? null : c.id)} style={{ cursor: "pointer" }}>
              <div className="list-card-title">{c.tenantName} <span className="row-muted">— {buildingName(c.buildingId)}{unit ? ` · Apt ${unit.unitNumber}` : ""}</span></div>
              <span className="pill pill-danger">{money(parseBalance(c.amountOwed))}</span>
              <IconBtn title="Edit" onClick={(e) => { e.stopPropagation(); openEdit(c); }}><Pencil size={14} /></IconBtn>
              <IconBtn title={c.archived ? "Reopen" : "Close — paid, settled or written off"} onClick={(e) => { e.stopPropagation(); update("collections", c.id, { archived: !c.archived }); }}><CheckCircle2 size={14} /></IconBtn>
              <IconBtn title="Delete" danger onClick={(e) => { e.stopPropagation(); if (window.confirm("Delete this collection?")) remove("collections", c.id); }}><Trash2 size={14} /></IconBtn>
            </div>
            <div className="list-card-body">
              <div className="row"><em>Moved out:</em> {c.movedOutDate ? fmtDate(c.movedOutDate) : "—"} · <em>Lawyer:</em> {c.lawyer || "not assigned"}</div>
              <div className="row"><em>Up to:</em> {c.stage || "not set"}</div>
              {log.length === 0 && <div className="hint">No log entries yet.</div>}
              {log.map(l => (
                <div className="row" key={l.id} style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
                  <span style={{ minWidth: 90, color: "var(--ink-soft)", fontSize: 12 }}>{fmtDate(l.date)}</span>
                  <span style={{ flex: 1 }}>{l.date > todayISO() && <span className="pill pill-warn" style={{ marginRight: 6, fontSize: 10 }}>Scheduled</span>}{l.note}</span>
                  <button className="checklist-remove" title="Remove this entry" onClick={(ev) => { ev.stopPropagation(); update("collections", c.id, { log: (c.log || []).filter(x => x.id !== l.id) }); }}><X size={12} /></button>
                </div>
              ))}
              <div className="inline-form" style={{ marginTop: 6, flexWrap: "wrap" }} onClick={(e) => e.stopPropagation()}>
                <input type="date" value={lf.date} onChange={e => setLogForm({ ...logForm, [c.id]: { ...lf, date: e.target.value } })} />
                <input placeholder="Add a note / update…" value={lf.note} onChange={e => setLogForm({ ...logForm, [c.id]: { ...lf, note: e.target.value } })} onKeyDown={e => e.key === "Enter" && addLog(c)} />
                <button className="btn-ghost" onClick={() => addLog(c)}>Add</button>
              </div>
            </div>
            {open && (
              <div className="list-card-body" style={{ paddingTop: 0 }}>
                <div className="row" style={{ marginTop: 10 }}><strong>Where it's up to</strong></div>
                <CaseStagePicker
                  options={data.caseStageOptions} value={c.stage || ""}
                  onPick={(label) => update("collections", c.id, { stage: label })}
                  onAdd={(label) => add("caseStageOptions", { label })}
                  onRemove={(o) => remove("caseStageOptions", o.id)}
                />
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

/* ============================== appointments (incl. recurring) ============================== */

function AppointmentsTab({ data, add, update, remove, buildingName, setData, deepLink, setDeepLink }) {
  const [form, setForm] = useState(null);
  const [view, setView] = useState("upcoming");
  const [expandedRow, setExpandedRow] = useState(null);
  const [highlightApptId, setHighlightApptId] = useState(null);
  const allTypes = [...APPOINTMENT_TYPES, ...(data.customAppointmentTypes || [])];

  useEffect(() => {
    if (!deepLink || deepLink.tab !== "inspections") return;
    const target = (data.appointments || []).find(a => a.id === deepLink.id);
    if (target) {
      setView(target.completed ? "completed" : "upcoming");
      setExpandedRow(target.id);
      setHighlightApptId(target.id);
      setTimeout(() => {
        document.getElementById(`appt-row-${target.id}`)?.scrollIntoView({ behavior: "smooth", block: "center" });
      }, 100);
      setTimeout(() => setHighlightApptId(null), 2500);
    }
    setDeepLink(null);
  }, [deepLink]);

  const submit = () => {
    if (!form.type) return;
    if (form.id) update("appointments", form.id, form);
    else add("appointments", { ...form, completed: false });
    // Stored in the actual saved data, not component state — component
    // state resets to nothing every time this tab unmounts (switching
    // tabs and back, reloading the page), which silently undid the whole
    // point of remembering this the moment you left the page.
    setData(d => ({ ...d, lastUsedAppointmentType: form.type }));
    setForm(null);
  };

  const addCustomType = (trimmed) => {
    if (!allTypes.includes(trimmed)) {
      setData(d => ({ ...d, customAppointmentTypes: [...(d.customAppointmentTypes || []), trimmed] }));
    }
  };

  // Defaults to 9am-11am so the time picker opens near a normal workday start
  // instead of midnight — still fully editable to anything earlier or later.
  const blank = () => ({ buildingId: data.buildings[0]?.id || "", unitId: "", type: data.lastUsedAppointmentType || allTypes[0] || "Other", date: "", timeFrom: "09:00", timeTo: "11:00", notes: "", recurring: false, completed: false });

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
          <div className={`list-card ${highlightApptId === a.id ? "sheet-row-deeplinked" : ""}`} key={a.id} id={`appt-row-${a.id}`}>
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

/* ============================== insurance ============================== */

// Same calendar day N months away (clamped to the end of a shorter month).
function addMonthsISO(d, n) {
  if (!d) return "";
  const [y, m, dd] = d.split("-").map(Number);
  const first = new Date(y, m - 1 + n, 1);
  const last = new Date(first.getFullYear(), first.getMonth() + 1, 0).getDate();
  first.setDate(Math.min(dd, last));
  return toLocalISO(first);
}
// Renewal steps, in order. The first three get a reminder 2 months before expiry;
// "chosen & bound" gets one 1 month before expiry.
const INSURANCE_STEPS = [
  { key: "stepQuotes", short: "Quotes requested", todo: "ask the current broker for renewal quotes" },
  { key: "stepDecs", short: "Dec pages & loss runs", todo: "get the declaration pages and loss runs" },
  { key: "stepOther", short: "Sent to other brokers", todo: "send the information to other brokers for competing quotes" },
];
function insuranceRemaining(p) { return INSURANCE_STEPS.filter(st => !p[st.key]); }

// The policies from the Fairmont Insurance Brokers summaries (prepared June 2026).
const INSURANCE_SEED = [
  { type: "Package (Property + Liability + Flood)", insured: "Newkirk Group LLC", covers: "1401-1425 Newkirk Ave aka 582-588 Marlborough Rd, Brooklyn 11226", broker: "Fairmont Insurance Brokers LLC, 718-232-3300", carrier: "Utica Mutual Insurance Company", policyNumber: "CPP 5503348", premium: "$34,164.28", effectiveDate: "2026-03-15", expirationDate: "2027-03-15", notes: "Building $12,699,721 ($10,000 ded). Rental income 12 months. Flood $1,000,000 ($25,000 ded). Sewer backup $250,000. General liability $1M occurrence / $2M aggregate. Mortgagees: Hancock JV Bidco LLC & SIG CRE 2023 Venture LLC (c/o Rialto Capital Advisors); Signature Bank ISAOA ATIMA." },
  { type: "Umbrella", insured: "Newkirk Group LLC", covers: "1401-1425 Newkirk Ave aka 582-588 Marlborough Rd, Brooklyn 11226", broker: "Fairmont Insurance Brokers LLC, 718-232-3300", carrier: "Utica Mutual Insurance Company", policyNumber: "CULP 5510540", premium: "$3,069.00", effectiveDate: "2026-03-15", expirationDate: "2027-03-15", notes: "Limit $2,000,000 / $2,000,000." },
  { type: "Business Auto", insured: "Newkirk Group LLC", covers: "Business auto", broker: "Fairmont Insurance Brokers LLC, 718-232-3300", carrier: "Utica Mutual Insurance Company", policyNumber: "5505272", premium: "$204.00", effectiveDate: "2026-03-15", expirationDate: "2027-03-15", notes: "Combined single limit $1,000,000." },
  { type: "Liability", insured: "Shpelrub Realty LLC", covers: "2177 East 21st St, Brooklyn 11229", broker: "Fairmont Insurance Brokers LLC, 718-232-3300", carrier: "Hadron Specialty Insurance Company", policyNumber: "H0208GL000068-01", premium: "$33,382.45", effectiveDate: "2025-12-10", expirationDate: "2026-12-10", notes: "General liability $1M occurrence / $2M aggregate. Lead and mold liability included. Mortgagee: Wells Fargo Bank, NA as Trustee (Benchmark 2020)." },
  { type: "Property", insured: "Shpelrub Realty LLC", covers: "2177 East 21st St, Brooklyn 11229", broker: "Fairmont Insurance Brokers LLC, 718-232-3300", carrier: "Pennsylvania Manufacturers Association Insurance Co", policyNumber: "8026019317306", premium: "$23,468.98", effectiveDate: "2026-02-08", expirationDate: "2027-02-08", notes: "Building $11,084,160 replacement cost ($10,000 ded). Equipment breakdown $12,056,993. Sewer backup $25,000." },
  { type: "Liability", insured: "Rafyana Realty Corp", covers: "2822 Brighton 8th St aka 205/217 Neptune Ave, Brooklyn 11235", broker: "Fairmont Insurance Brokers LLC, 718-232-3300", carrier: "Clear Blue Specialty Insurance Company", policyNumber: "BY01-8000193-04", premium: "$21,589.68", effectiveDate: "2025-10-20", expirationDate: "2026-10-20", notes: "General liability $1M occurrence / $2M aggregate. This is the 2025-26 term — press Renewed once the 2026-27 summary arrives." },
  { type: "Property", insured: "Rafyana Realty Corp", covers: "2822 Brighton 8th St aka 205/217 Neptune Ave, Brooklyn 11235", broker: "Fairmont Insurance Brokers LLC, 718-232-3300", carrier: "National Casualty Company", policyNumber: "RYO0006154", premium: "$17,305.65", effectiveDate: "2025-10-20", expirationDate: "2026-10-20", notes: "Building $5,524,100 ($5,000 ded). Flood excluded. Mortgagee: Northfield Bank. This is the 2025-26 term — press Renewed once the 2026-27 summary arrives." },
  { type: "Package (Property + Liability)", insured: "Neptune Group LLC", covers: "350-362 Neptune Ave, Brooklyn 11235", broker: "Fairmont Insurance Brokers LLC, 718-232-3300", carrier: "Allianz Global Risks US Insurance Company", policyNumber: "AXMP25000516-00", premium: "$38,871.74", effectiveDate: "2025-12-30", expirationDate: "2026-12-30", notes: "Building $29,035,794 ($10,000 ded). Rental income $2,100,000. Ordinance & law $1,000,000. General liability $1M occurrence / $2M aggregate. Mortgagee: Benchmark 2021-B23 Mortgagee Trust." },
];

// Pieces used by the Insurance tab.
const INSURANCE_TYPES = ["Property", "Liability", "Flood", "Umbrella", "Package", "Workers' Comp", "Business Auto"];
const INSURANCE_CORE = ["Property", "Liability", "Flood", "Umbrella"];
const insBuildingIds = (p) => (Array.isArray(p.buildingIds) ? p.buildingIds : []);
// Which of the four core types a policy counts as (a package counts as several).
function insuranceCovers(type) {
  const t = String(type || "").toLowerCase();
  return INSURANCE_CORE.filter(c => t.includes(c.toLowerCase()));
}

function InsurancePolicyForm({ form, setForm, buildings, onSave, onSaveAnother, onCancel, saveLabel, existing, notice }) {
  const toggleB = (id) => setForm({ ...form, buildingIds: form.buildingIds.includes(id) ? form.buildingIds.filter(x => x !== id) : [...form.buildingIds, id] });
  const set = (k, v) => setForm({ ...form, [k]: v });
  const F = (label, node, span) => <label style={{ display: "flex", flexDirection: "column", gap: 3, fontSize: 12, gridColumn: span ? `span ${span}` : undefined }}><span className="row-muted">{label}</span>{node}</label>;
  const have = new Set();
  existing.filter(p => insBuildingIds(p).some(id => form.buildingIds.includes(id))).forEach(p => insuranceCovers(p.type).forEach(c => have.add(c)));
  const ok = form.type.trim() && form.buildingIds.length > 0;
  return (
    <div className="form-panel" style={{ marginBottom: 12 }}>
      <div style={{ fontWeight: 600, marginBottom: 6 }}>1. Which building(s) is this policy for?</div>
      <div className="inline-form" style={{ flexWrap: "wrap", margin: "0 0 4px" }}>
        {buildings.map(b => <button key={b.id} type="button" className={`chip ${form.buildingIds.includes(b.id) ? "chip-active" : ""}`} onClick={() => toggleB(b.id)}>{shortAddress(b.address)}</button>)}
      </div>
      {form.buildingIds.length > 0 && (
        <div className="hint" style={{ marginBottom: 8 }}>
          Already entered for {form.buildingIds.length === 1 ? "this building" : "these buildings"}: {INSURANCE_CORE.map(c => <span key={c} style={{ marginRight: 8, color: have.has(c) ? "var(--ok, #1E7B34)" : undefined }}>{have.has(c) ? "✓" : "○"} {c}</span>)}
        </div>
      )}
      <div style={{ fontWeight: 600, margin: "10px 0 6px" }}>2. What kind of policy?</div>
      <div className="inline-form" style={{ flexWrap: "wrap", margin: "0 0 6px" }}>
        {INSURANCE_TYPES.map(t => <button key={t} type="button" className={`chip ${form.type === t ? "chip-active" : ""}`} onClick={() => set("type", t)}>{t}</button>)}
        <input placeholder="or type another kind…" value={INSURANCE_TYPES.includes(form.type) ? "" : form.type} onChange={e => set("type", e.target.value)} style={{ maxWidth: 220 }} />
      </div>
      <div style={{ fontWeight: 600, margin: "10px 0 6px" }}>3. Details</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))", gap: 10 }}>
        {F("Insured company (named insured)", <input value={form.insured} onChange={e => set("insured", e.target.value)} />)}
        {F("Insurance carrier", <input value={form.carrier} onChange={e => set("carrier", e.target.value)} />)}
        {F("Policy number", <input value={form.policyNumber} onChange={e => set("policyNumber", e.target.value)} />)}
        {F("Broker (name, phone, email)", <input value={form.broker} onChange={e => set("broker", e.target.value)} />)}
        {F("Premium", <input value={form.premium} placeholder="$" onChange={e => set("premium", e.target.value)} />)}
        {F("Policy starts", <input type="date" value={form.effectiveDate} onChange={e => set("effectiveDate", e.target.value)} />)}
        {F("Policy expires", <input type="date" value={form.expirationDate} onChange={e => set("expirationDate", e.target.value)} />)}
      </div>
      <div style={{ marginTop: 10 }}>{F("Coverage limits, deductibles, mortgagees, notes", <textarea rows={2} value={form.notes} onChange={e => set("notes", e.target.value)} />)}</div>
      {existing && form.id && (
        <>
          <div style={{ fontWeight: 600, margin: "10px 0 6px" }}>Renewal steps (dates done)</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: 10 }}>
            {[["stepQuotes", "1. Renewal quotes requested"], ["stepDecs", "2. Dec pages & loss runs received"], ["stepOther", "3. Sent to other brokers"], ["stepBound", "4. Chosen & bound"]].map(([k, l]) => F(l, <input key={k} type="date" value={form[k] || ""} onChange={e => set(k, e.target.value)} />))}
          </div>
        </>
      )}
      <div className="form-actions" style={{ alignItems: "center" }}>
        <button className="btn-primary" onClick={onSave} disabled={!ok}>{saveLabel}</button>
        {onSaveAnother && <button className="btn-ghost" onClick={onSaveAnother} disabled={!ok} title="Saves this one and keeps the building, company and broker filled in">Save &amp; add another for the same building</button>}
        <button className="btn-ghost" onClick={onCancel}>Cancel</button>
        {!ok && <span className="hint">Pick a building and a policy type.</span>}
        {notice && <span className="hint" style={{ color: "var(--ok, #1E7B34)" }}>{notice}</span>}
      </div>
    </div>
  );
}

function InsuranceTab({ data, add, update, remove, buildingName }) {
  const blank = { type: "", buildingIds: [], insured: "", broker: "", carrier: "", policyNumber: "", premium: "", effectiveDate: "", expirationDate: "", notes: "" };
  const [form, setForm] = useState(null);       // add form
  const [notice, setNotice] = useState("");
  const [editing, setEditing] = useState(null); // policy id being edited
  const [ef, setEf] = useState(null);
  const [renewing, setRenewing] = useState(null);
  const [rf, setRf] = useState({});
  const [mailFor, setMailFor] = useState(null);
  const [copied, setCopied] = useState("");
  const [bFilter, setBFilter] = useState("All");
  const today = todayISO();
  const policies = (data.insurancePolicies || []).slice().sort((a, b) => (a.expirationDate || "9999-99-99").localeCompare(b.expirationDate || "9999-99-99"));
  const coversText = (p) => insBuildingIds(p).map(id => buildingName(id)).join(", ") || p.covers || "";
  const missingSeed = INSURANCE_SEED.filter(sd => !policies.some(p => (p.policyNumber || "").trim() === sd.policyNumber));
  const loadSeed = () => {
    if (!window.confirm(`Add ${missingSeed.length} policies from the Fairmont insurance summaries (Newkirk, Shpelrub, Rafyana, Neptune)? They're linked to a building when the address matches one of yours.`)) return;
    missingSeed.forEach(sd => {
      const key = (sd.covers || "").toLowerCase();
      const hit = data.buildings.find(b => { const a = (b.address || "").toLowerCase(); return (key.includes("newkirk") && a.includes("newkirk")) || (key.includes("east 21st") && a.includes("21st")) || (key.includes("brighton 8th") && a.includes("brighton 8th")) || (key.includes("350-362 neptune") && a.includes("neptune")); });
      add("insurancePolicies", { id: uid(), ...sd, buildingIds: hit ? [hit.id] : [] });
    });
  };
  const saveNew = (another) => {
    if (!form.type.trim() || form.buildingIds.length === 0) return;
    const rec = { id: uid(), ...form, type: form.type.trim() };
    delete rec.covers;
    add("insurancePolicies", rec);
    if (another) {
      setNotice(`Saved ${rec.type} for ${form.buildingIds.map(buildingName).join(", ")} — now the next one.`);
      setForm({ ...blank, buildingIds: form.buildingIds, insured: form.insured, broker: form.broker });
      setTimeout(() => setNotice(""), 4000);
    } else { setForm(null); setNotice(""); }
  };
  const startEdit = (p) => { setEditing(p.id); setEf({ ...blank, ...p, buildingIds: insBuildingIds(p) }); };
  const saveEdit = () => {
    const patch = {};
    Object.keys(ef).forEach(k => { if (k !== "id" && k !== "history" && JSON.stringify(ef[k]) !== JSON.stringify((policies.find(p => p.id === editing) || {})[k])) patch[k] = ef[k]; });
    if (Object.keys(patch).length) update("insurancePolicies", editing, patch);
    setEditing(null);
  };
  const startRenew = (p) => { setRenewing(p.id); setRf({ effectiveDate: p.expirationDate || "", expirationDate: "", carrier: p.carrier || "", policyNumber: "", broker: p.broker || "", premium: "" }); };
  const saveRenew = (p) => {
    if (!rf.expirationDate) return;
    const old = { expired: p.expirationDate || "", carrier: p.carrier || "", broker: p.broker || "", policyNumber: p.policyNumber || "", premium: p.premium || "", renewedOn: today };
    update("insurancePolicies", p.id, {
      effectiveDate: rf.effectiveDate, expirationDate: rf.expirationDate, carrier: rf.carrier, policyNumber: rf.policyNumber, broker: rf.broker, premium: rf.premium,
      stepQuotes: "", stepDecs: "", stepOther: "", stepBound: "", history: [...(p.history || []), old],
    });
    setRenewing(null);
  };
  const printAll = () => {
    if (policies.length === 0) { window.alert("No policies to print yet."); return; }
    const d = (x) => (x ? fmtDate(x) : "");
    const order = data.buildings.map(b => b.id);
    const byBuilding = (p) => { const ids = insBuildingIds(p); return ids.length ? Math.min(...ids.map(i => order.indexOf(i)).filter(i => i >= 0)) : 9999; };
    const rows = policies.slice().sort((a, b) => byBuilding(a) - byBuilding(b) || (a.expirationDate || "").localeCompare(b.expirationDate || "")).map(p => `<tr>
      <td class="l"><b>${escHtml(coversText(p))}</b></td><td class="l"><b>${escHtml(p.type)}</b></td><td class="l">${escHtml(p.insured)}</td><td class="l">${escHtml(p.broker)}</td><td class="l">${escHtml(p.carrier)}</td>
      <td>${escHtml(p.policyNumber)}</td><td>${escHtml(d(p.effectiveDate))}${p.effectiveDate || p.expirationDate ? " –<br>" : ""}<b>${escHtml(d(p.expirationDate))}</b></td><td>${escHtml(p.premium)}</td>
      <td>${escHtml(p.expirationDate ? d(addMonthsISO(p.expirationDate, -2)) : "")}</td>
      <td>${escHtml(d(p.stepQuotes))}</td><td>${escHtml(d(p.stepDecs))}</td><td>${escHtml(d(p.stepOther))}</td><td>${escHtml(d(p.stepBound))}</td><td class="l">${escHtml(p.notes)}</td></tr>`).join("");
    openPrintWindow("Insurance schedule", `<h1>Insurance schedule</h1><div class="sub">Printed ${escHtml(fmtDate(today))} — by building</div>
      <table style="font-size:9.5px"><tr><th class="l">Building</th><th class="l">Policy</th><th class="l">Insured</th><th class="l">Broker</th><th class="l">Carrier</th><th>Policy #</th><th>Policy period</th><th>Premium</th><th>Start renewal</th><th>Quotes requested</th><th>Dec pages &amp; loss runs</th><th>Sent to other brokers</th><th>Chosen &amp; bound</th><th class="l">Notes</th></tr>${rows}</table>
      <div class="note">Renewal process, starting two months before each policy expires: (1) ask the current broker for renewal quotes, (2) request declaration pages and loss runs, (3) send the information to other brokers for competing quotes, then compare coverage and price one month before expiration.</div>`);
  };
  const emailFor = (p) => {
    const cv = coversText(p);
    const what = `${p.type}${p.insured ? ` policy for ${p.insured}` : " policy"}`;
    const cov = cv ? ` (${cv})` : "";
    const cur = `${p.carrier ? ` with ${p.carrier}` : ""}${p.policyNumber ? `, policy number ${p.policyNumber}` : ""}`;
    const exp = p.expirationDate ? fmtDate(p.expirationDate) : "";
    const toCurrent = `Subject: Renewal quotes, declaration pages and loss runs — ${cv || p.insured || ""} ${p.type}${p.policyNumber ? ` (${p.policyNumber})` : ""}\n\nHello,\n\nThe ${what}${cov}${cur}${exp ? ` expires on ${exp}` : " is coming up for renewal"}. Please send us:\n\n1. Your renewal quotes\n2. The current declaration pages\n3. Currently valued loss runs for the past five years\n\nWe will also be getting competing quotes from other brokers, so please send these as soon as you can.\n\nThank you,\nMoshe`;
    const toOthers = `Subject: Quote request — ${cv || p.insured || ""} ${p.type}${exp ? `, effective ${exp}` : ""}\n\nHello,\n\nWe are reviewing the renewal of the ${what}${cov}${p.carrier ? `, currently with ${p.carrier}` : ""}${exp ? `, which expires on ${exp}` : ""}. Attached are the current declaration pages and loss runs.${p.notes ? `\n\nCurrent coverage details: ${p.notes}` : ""}\n\nPlease send a quote with the same or better coverage and tell us where your coverage differs from what we have now, so we can compare coverage and pricing side by side.\n\nThank you,\nMoshe`;
    return { toCurrent, toOthers };
  };
  const copyText = (key, text) => { navigator.clipboard.writeText(text); setCopied(key); setTimeout(() => setCopied(""), 1500); };

  // ----- one policy card -----
  const PolicyCard = (p) => {
    const days = p.expirationDate ? daysUntil(p.expirationDate) : null;
    const rem = insuranceRemaining(p);
    const inWindow = p.expirationDate && today >= addMonthsISO(p.expirationDate, -2);
    const tone = days === null ? "#9AA0A6" : days < 0 ? "#B3261E" : days <= 60 ? "#C77700" : "#1E7B34";
    const steps = [["stepQuotes", "Quotes requested"], ["stepDecs", "Dec pages & loss runs"], ["stepOther", "Sent to other brokers"], ["stepBound", "Chosen & bound"]];
    const mail = mailFor === p.id ? emailFor(p) : null;
    if (editing === p.id) {
      return <InsurancePolicyForm key={p.id} form={ef} setForm={setEf} buildings={data.buildings} existing={policies} saveLabel="Save changes" onSave={saveEdit} onCancel={() => setEditing(null)} />;
    }
    return (
      <div key={p.id} className="list-card" style={{ borderLeft: `4px solid ${tone}`, marginBottom: 8 }}>
        <div className="list-card-body">
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "baseline" }}>
            <strong style={{ fontSize: 15 }}>{p.type}</strong>
            {p.carrier && <span>{p.carrier}</span>}
            {p.policyNumber && <span className="row-muted">#{p.policyNumber}</span>}
            {p.premium && <span className="row-muted">{p.premium}</span>}
            <span style={{ marginLeft: "auto", color: tone, fontWeight: 600 }}>
              {p.expirationDate ? (days < 0 ? `Expired ${fmtDate(p.expirationDate)}` : `Expires ${fmtDate(p.expirationDate)} · ${days} days`) : "No expiration date"}
            </span>
          </div>
          <div className="row-muted" style={{ fontSize: 12, marginTop: 2 }}>
            {p.insured && <span>{p.insured} · </span>}{p.broker && <span>Broker: {p.broker}</span>}
          </div>
          {p.notes && <div style={{ fontSize: 12, marginTop: 4 }}>{p.notes}</div>}
          {(p.history || []).length > 0 && <div className="row-muted" style={{ fontSize: 11, marginTop: 2 }}>Past terms: {(p.history || []).map(h => `${h.expired ? fmtDate(h.expired) : "?"} ${h.carrier || ""}${h.premium ? ` ${h.premium}` : ""}`).join("; ")}</div>}
          {p.expirationDate && (
            <div style={{ marginTop: 8 }}>
              <div className="row-muted" style={{ fontSize: 11, marginBottom: 3 }}>
                Renewal steps {inWindow ? "— start now" : `— reminders start ${fmtDate(addMonthsISO(p.expirationDate, -2))}`} (tap to mark done / undo)
              </div>
              <div className="inline-form" style={{ flexWrap: "wrap", margin: 0 }}>
                {steps.map(([k, l]) => (
                  <button key={k} type="button" className={`chip ${p[k] ? "chip-active" : ""}`} title={p[k] ? `Done ${fmtDate(p[k])} — tap to undo` : "Tap when done"}
                    onClick={() => update("insurancePolicies", p.id, { [k]: p[k] ? "" : today })}>{p[k] ? "✓" : "○"} {l}{p[k] ? ` · ${fmtDate(p[k])}` : ""}</button>
                ))}
              </div>
              {rem.length === 0 && !p.stepBound && <div className="hint" style={{ marginTop: 4 }}>All three steps done — compare the quotes and choose by {fmtDate(addMonthsISO(p.expirationDate, -1))}.</div>}
            </div>
          )}
          <div className="inline-form" style={{ flexWrap: "wrap", margin: "8px 0 0" }}>
            <button className="btn-ghost" onClick={() => startEdit(p)}><Pencil size={14} /> Edit</button>
            <button className="btn-ghost" onClick={() => setMailFor(mailFor === p.id ? null : p.id)}><Mail size={14} /> Email brokers</button>
            <button className="btn-ghost" onClick={() => startRenew(p)}>Renewed</button>
            <button className="btn-ghost" onClick={() => { if (window.confirm("Delete this policy?")) remove("insurancePolicies", p.id); }}><Trash2 size={14} /> Delete</button>
          </div>
          {mail && (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 12, marginTop: 8 }}>
              {[["toCurrent", "To your current broker"], ["toOthers", "To other brokers (attach the dec pages and loss runs)"]].map(([k, label]) => (
                <div key={k}>
                  <div className="row"><strong>{label}</strong> <button className="btn-ghost" style={{ marginLeft: 8 }} onClick={() => copyText(p.id + k, mail[k])}>{copied === p.id + k ? "Copied!" : "Copy"}</button></div>
                  <textarea readOnly rows={11} value={mail[k]} style={{ width: "100%", boxSizing: "border-box" }} />
                </div>
              ))}
            </div>
          )}
          {renewing === p.id && (
            <div className="inline-form" style={{ flexWrap: "wrap", alignItems: "center", marginTop: 8 }}>
              <strong>New term</strong>
              <label className="hint">Starts <input type="date" value={rf.effectiveDate} onChange={e => setRf({ ...rf, effectiveDate: e.target.value })} /></label>
              <label className="hint">Expires <input type="date" value={rf.expirationDate} onChange={e => setRf({ ...rf, expirationDate: e.target.value })} /></label>
              <input placeholder="Carrier" value={rf.carrier} onChange={e => setRf({ ...rf, carrier: e.target.value })} />
              <input placeholder="Policy #" value={rf.policyNumber} onChange={e => setRf({ ...rf, policyNumber: e.target.value })} />
              <input placeholder="Broker" value={rf.broker} onChange={e => setRf({ ...rf, broker: e.target.value })} />
              <input placeholder="Premium" value={rf.premium} onChange={e => setRf({ ...rf, premium: e.target.value })} style={{ maxWidth: 110 }} />
              <button className="btn-primary" disabled={!rf.expirationDate} onClick={() => saveRenew(p)}>Save</button>
              <button className="btn-ghost" onClick={() => setRenewing(null)}>Cancel</button>
              <div className="hint" style={{ flexBasis: "100%" }}>The old term is kept under "Past terms" and the renewal steps reset for next year.</div>
            </div>
          )}
        </div>
      </div>
    );
  };

  const groups = data.buildings
    .filter(b => bFilter === "All" || b.id === bFilter)
    .map(b => ({ b, items: policies.filter(p => insBuildingIds(p).includes(b.id)) }))
    .filter(g => g.items.length > 0 || bFilter !== "All");
  const unlinked = bFilter === "All" ? policies.filter(p => insBuildingIds(p).length === 0) : [];
  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Insurance</h1>
        <div className="page-actions">
          {missingSeed.length > 0 && <button className="btn-ghost" onClick={loadSeed}>Load the {missingSeed.length} policies from the Fairmont summaries</button>}
          <button className="btn-ghost" onClick={printAll}><Printer size={14} /> Print all</button>
          <button className="btn-primary" onClick={() => { setForm(form ? null : { ...blank, buildingIds: bFilter !== "All" ? [bFilter] : [] }); setNotice(""); }}><Plus size={14} /> Add policy</button>
        </div>
      </div>
      <div className="hint" style={{ marginBottom: 8 }}>
        Two months before a policy expires you get a reminder on the Dashboard and calendar to ask your broker for renewal quotes, get the dec pages and loss runs, and send everything to other brokers. One month before, a reminder to compare and choose.
      </div>
      {form && <InsurancePolicyForm form={form} setForm={setForm} buildings={data.buildings} existing={policies} saveLabel="Save" onSave={() => saveNew(false)} onSaveAnother={() => saveNew(true)} onCancel={() => { setForm(null); setNotice(""); }} notice={notice} />}
      <div className="filter-row">
        <select value={bFilter} onChange={e => setBFilter(e.target.value)} style={{ maxWidth: 240 }}>
          <option value="All">All buildings</option>
          {data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}
        </select>
      </div>
      {policies.length === 0 && !form && <EmptyState text="No policies yet — press Add policy." />}
      {groups.map(({ b, items }) => {
        const have = new Set();
        items.forEach(p => insuranceCovers(p.type).forEach(c => have.add(c)));
        return (
          <div key={b.id} style={{ marginBottom: 16 }}>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center", margin: "6px 0" }}>
              <h2 style={{ fontSize: 16, margin: 0 }}>{shortAddress(b.address)}</h2>
              {INSURANCE_CORE.map(c => have.has(c)
                ? <span key={c} className="pill pill-ok">✓ {c}</span>
                : <button key={c} type="button" className="chip" title={`Add ${c} for this building`} onClick={() => { setForm({ ...blank, type: c, buildingIds: [b.id], insured: items[0]?.insured || "", broker: items[0]?.broker || "" }); setNotice(""); window.scrollTo({ top: 0, behavior: "smooth" }); }}>+ {c}</button>)}
            </div>
            {items.map(PolicyCard)}
          </div>
        );
      })}
      {unlinked.length > 0 && (
        <div style={{ marginBottom: 16 }}>
          <h2 style={{ fontSize: 16, margin: "6px 0" }}>Not linked to a building yet</h2>
          <div className="hint" style={{ marginBottom: 6 }}>Press Edit and pick the building(s) so these show up under them.</div>
          {unlinked.map(PolicyCard)}
        </div>
      )}
    </div>
  );
}

/* ============================== unit turn ============================== */

const TURN_ROOM_COLS = [
  { key: "peeling", label: "Peeling paint", rule: "Yes = remediate" },
  { key: "teeth", label: "Teeth marks on painted chewable surface", rule: "Yes = abate" },
  { key: "subsurface", label: "Deteriorated painted subsurface", rule: "Yes = remediate" },
  { key: "doorFrame", label: "Door & door frame friction surfaces painted", rule: "Yes = abate" },
  { key: "windowFrame", label: "Window frame friction surface painted", rule: "Yes = abate / window channels" },
  { key: "impact", label: "Painted area shows impact damage", rule: "Yes = remediate" },
  { key: "floors", label: "Window wells & bare floors smooth and cleanable", rule: "No = remediate" },
];
const TURN_CHECKLIST = [
  { group: "Lead paint (Local Law 1 turnover)", items: [
    ["lead_inspect", "Turnover lead-paint inspection done — every surface in every room, incl. inside closets & cabinets"],
    ["lead_fix", "Hazards fixed — wet scrape & repaint by EPA-certified (RRP) workers using safe work practices"],
    ["lead_abate", "Friction / impact surfaces or teeth marks abated, or window channels installed (if found)"],
    ["lead_clean", "Cleanup after lead work (and dust clearance if required)"],
    ["lead_records", "Records filed: SF LL1-06 summary, SF LL1-07 detail, contractor EPA certs & affidavits (keep 10 years)"],
  ] },
  { group: "Safety", items: [
    ["smoke", "Smoke detector installed and working"],
    ["co", "Carbon monoxide detector installed and working"],
    ["windowGuards", "Window guards installed / checked"],
    ["gas", "Stove, gas, heat and hot water checked"],
    ["locks", "Locks changed; keys / fobs ready; intercom and mailbox updated"],
  ] },
  { group: "Condition", items: [
    ["pests", "Pest inspection (roaches, mice) and bed-bug check"],
    ["mold", "Mold, leaks and water damage checked"],
    ["clean", "Cleaned; trash removed"],
    ["paint", "Painted; walls, ceilings and trim touched up"],
    ["floors", "Floors repaired / refinished"],
    ["fixtures", "Plumbing, fixtures, appliances and electric tested"],
    ["photos", "Photos taken before move-in"],
  ] },
  { group: "New tenant paperwork", items: [
    ["leadDisclosure", "Lead-paint disclosure and pamphlet given (building built before 1978)"],
    ["bedbug", "Bed-bug infestation history notice given"],
    ["windowNotice", "Window guard notice / form signed"],
    ["detectorNotice", "Smoke / CO detector notice signed"],
    ["floodNotice", "Flood and sprinkler disclosures in the lease"],
    ["legalRent", "Rent-stabilized: legal rent checked, vacancy rider given, DHCR registration and improvement records filed"],
  ] },
];
const TURN_CHECKLIST_FLAT = TURN_CHECKLIST.flatMap(g => g.items.map(([k, l]) => ({ key: k, label: l, group: g.group })));

function escHtml(x) { return String(x == null ? "" : x).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
function openPrintWindow(title, bodyHtml) {
  const w = window.open("", "_blank");
  if (!w) { window.alert("Your browser blocked the print window — allow pop-ups for this site and try again."); return; }
  w.document.write(`<!doctype html><html><head><title>${escHtml(title)}</title><style>
    @page { size: landscape; margin: 0.5in; }
    body { font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: #000; }
    h1 { font-size: 15px; margin: 0; } .sub { font-size: 11px; margin: 2px 0 10px; }
    .meta { display: flex; flex-wrap: wrap; gap: 6px 24px; margin: 8px 0 12px; }
    .meta div { flex: 1 1 30%; border-bottom: 1px solid #000; min-height: 18px; padding-top: 4px; }
    .meta span { font-size: 9px; color: #333; display: block; }
    table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid #000; padding: 4px 5px; text-align: center; vertical-align: middle; }
    th { font-size: 9.5px; } td.l, th.l { text-align: left; }
    .note { font-size: 9px; margin-top: 8px; } .pb { page-break-after: always; }
    .box { display: inline-block; width: 11px; height: 11px; border: 1px solid #000; margin-right: 6px; vertical-align: -2px; text-align: center; line-height: 11px; font-size: 10px; }
  </style></head><body>${bodyHtml}<script>window.onload=function(){setTimeout(function(){window.print();},250);};</script></body></html>`);
  w.document.close();
}

function UnitTurnTab({ data, add, update, remove, buildingName }) {
  const thisYear = String(new Date().getFullYear());
  const [bFilter, setBFilter] = useState("All");
  const [year, setYear] = useState("All");
  const [open, setOpen] = useState(null);
  const [nb, setNb] = useState(data.buildings[0]?.id || "");
  const [nu, setNu] = useState("");
  const [nTenant, setNTenant] = useState("");
  const [nEnded, setNEnded] = useState("");
  const [roomName, setRoomName] = useState({});
  const [customText, setCustomText] = useState({});
  const unitsFor = (bid) => data.units.filter(u => u.buildingId === bid).sort((a, b) => compareUnits(a.unitNumber, b.unitNumber));
  const unitNum = (id) => data.units.find(u => u.id === id)?.unitNumber || "—";
  const addr = (bid) => data.buildings.find(b => b.id === bid)?.address || "";
  const turnYear = (t) => (t.inspectionDate || t.endedDate || t.createdAt || "").slice(0, 4);
  const pickUnit = (id) => {
    setNu(id);
    const ts = data.tenants.filter(t => t.unitId === id);
    const cur = ts.find(t => !t.movedOut) || ts[ts.length - 1];
    setNTenant(cur?.name || "");
  };
  const startTurn = () => {
    if (!nb || !nu) return;
    const id = uid();
    add("unitTurns", { id, buildingId: nb, unitId: nu, prevTenant: nTenant, endedDate: nEnded, createdAt: todayISO(), exempt: "no", inspectionDate: "", inspectedBy: "owner", inspectorName: "", hazardsFound: "", rooms: [], checklist: {}, customItems: [], notes: [] });
    setOpen(id); setNu(""); setNTenant(""); setNEnded("");
  };
  const years = Array.from(new Set((data.unitTurns || []).map(turnYear).filter(Boolean))).sort().reverse();
  const shown = (data.unitTurns || [])
    .filter(t => bFilter === "All" || t.buildingId === bFilter)
    .filter(t => year === "All" || turnYear(t) === year)
    .slice().sort((a, b) => (b.endedDate || b.createdAt || "").localeCompare(a.endedDate || a.createdAt || ""));
  const progress = (t) => {
    const done = TURN_CHECKLIST_FLAT.filter(i => t.checklist?.[i.key]?.done).length + (t.customItems || []).filter(i => i.done).length;
    return `${done}/${TURN_CHECKLIST_FLAT.length + (t.customItems || []).length}`;
  };
  const ynLabel = { yes: "Yes", no: "No", "": "—" };
  const cycle = (v) => (v === "" || v == null ? "yes" : v === "yes" ? "no" : "");
  const hazardsSuggested = (t) => (t.rooms || []).some(r => ["peeling", "teeth", "subsurface", "doorFrame", "windowFrame", "impact"].some(k => r[k] === "yes") || r.floors === "no");
  const exemptLabel = { no: "No", hpd: "Yes — HPD Lead Free exemption", tested: "Yes — tested, no lead-based paint" };

  // ----- printing -----
  // The two HPD forms are printed on top of the official sample forms (SF LL1-06
  // and SF LL1-07), with everything filled in. Owner/agent name, borough and the
  // property registration # come from the "Form details" box for the building.
  const bInfo = (bid) => {
    const b = data.buildings.find(x => x.id === bid) || {};
    return { address: b.address || "", ownerAgent: b.hpdOwnerAgent || "", borough: b.hpdBorough || "Brooklyn", regId: b.hpdRegId || "" };
  };
  const warnMissing = (bids) => {
    const missing = [];
    bids.forEach(bid => { const i = bInfo(bid); const miss = [!i.ownerAgent && "owner/agent name", !i.regId && "property registration #"].filter(Boolean); if (miss.length) missing.push(`${shortAddress(i.address)}: ${miss.join(", ")}`); });
    return missing.length === 0 || window.confirm(`These are blank on the form (fill them in under "Form details"):\n\n${missing.join("\n")}\n\nPrint anyway?`);
  };
  const printLL107 = (t) => {
    if (!warnMissing([t.buildingId])) return;
    const i = bInfo(t.buildingId);
    openHpdPrint("SF LL1-07", buildLL107Pages({ year: turnYear(t) || thisYear, address: i.address, apt: unitNum(t.unitId), inspector: t.inspectorName || "", borough: i.borough, regId: i.regId }, t.rooms || []));
  };
  const printLL106 = () => {
    const byB = {};
    shown.forEach(t => { (byB[t.buildingId] = byB[t.buildingId] || []).push(t); });
    const bids = Object.keys(byB);
    if (bids.length === 0) { window.alert("No unit turns to print for this filter."); return; }
    if (!warnMissing(bids)) return;
    const yr = year === "All" ? thisYear : year;
    const pages = bids.map(bid => {
      const i = bInfo(bid);
      const rows = byB[bid].slice().sort((a, b) => compareUnits(unitNum(a.unitId), unitNum(b.unitId))).map(t => ({
        apt: unitNum(t.unitId), exempt: t.exempt || "no", ended: t.endedDate, inspected: t.inspectionDate, by: t.inspectedBy, inspector: t.inspectorName, hazards: t.hazardsFound,
      }));
      return buildLL106Pages({ year: yr, address: i.address, ownerAgent: i.ownerAgent, borough: i.borough, regId: i.regId }, rows);
    }).join("");
    openHpdPrint("SF LL1-06", pages);
  };
  const printChecklist = (t) => {
    const items = (g) => g.map(i => `<tr><td class="l" style="width:70%"><span class="box">${t.checklist?.[i.key]?.done ? "X" : ""}</span>${escHtml(i.label)}</td><td>${escHtml(t.checklist?.[i.key]?.date ? fmtDate(t.checklist[i.key].date) : "")}</td><td class="l">${escHtml(t.checklist?.[i.key]?.note || "")}</td></tr>`).join("");
    const body = TURN_CHECKLIST.map(g => `<h1 style="margin-top:10px;font-size:12px">${escHtml(g.group)}</h1><table>${items(g.items.map(([k, l]) => ({ key: k, label: l })))}</table>`).join("")
      + ((t.customItems || []).length ? `<h1 style="margin-top:10px;font-size:12px">Other</h1><table>${(t.customItems || []).map(i => `<tr><td class="l" style="width:70%"><span class="box">${i.done ? "X" : ""}</span>${escHtml(i.text)}</td><td>${escHtml(i.date ? fmtDate(i.date) : "")}</td><td class="l">${escHtml(i.note || "")}</td></tr>`).join("")}</table>` : "");
    openPrintWindow("Unit turn checklist", `<h1>Unit Turn Checklist — ${escHtml(addr(t.buildingId))}, Apt ${escHtml(unitNum(t.unitId))}</h1>
      <div class="sub">Previous tenant: ${escHtml(t.prevTenant || "")} &nbsp; Occupancy ended: ${escHtml(t.endedDate ? fmtDate(t.endedDate) : "")} &nbsp; New tenant: ${escHtml(t.newTenant || "")} &nbsp; Move-in: ${escHtml(t.moveInDate ? fmtDate(t.moveInDate) : "")}</div>${body}`);
  };

  const field = (label, node) => <label style={{ display: "flex", flexDirection: "column", gap: 2, fontSize: 12 }}><span className="row-muted">{label}</span>{node}</label>;

  return (
    <div>
      <div className="page-head">
        <h1 className="page-title">Unit Turn</h1>
        <div className="page-actions">
          <button className="btn-ghost" onClick={printLL106} title="HPD summary form for the buildings and year shown"><Printer size={14} /> Print summary (LL1-06)</button>
        </div>
      </div>
      <div className="hint" style={{ marginBottom: 8 }}>Start a turn, then answer the questions on its card — every answer saves when you press Enter or click away. Print the HPD forms (SF LL1-06 summary, SF LL1-07 per unit) or the checklist any time.</div>
      <div className="form-panel" style={{ marginBottom: 10 }}>
        <div className="inline-form" style={{ flexWrap: "wrap", alignItems: "flex-end", margin: 0 }}>
          {field("Building", <select value={nb} onChange={e => { setNb(e.target.value); setNu(""); }}>{data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}</select>)}
          {field("Apt", <select value={nu} onChange={e => pickUnit(e.target.value)}><option value="">Select apt…</option>{unitsFor(nb).map(u => <option key={u.id} value={u.id}>{u.unitNumber}</option>)}</select>)}
          {field("Previous tenant", <input value={nTenant} onChange={e => setNTenant(e.target.value)} />)}
          {field("Date occupancy ended", <input type="date" value={nEnded} onChange={e => setNEnded(e.target.value)} />)}
          <button className="btn-primary" onClick={startTurn} disabled={!nb || !nu}><Plus size={14} /> Start unit turn</button>
        </div>
      </div>
      <details className="form-panel" style={{ marginBottom: 10 }}>
        <summary style={{ cursor: "pointer", fontWeight: 600 }}>Form details printed on the HPD forms (owner/agent, borough, registration #)</summary>
        <div style={{ marginTop: 8 }}>
          {data.buildings.map(b => (
            <div key={b.id} style={{ display: "grid", gridTemplateColumns: "minmax(140px, 1fr) repeat(3, minmax(160px, 1fr))", gap: 8, alignItems: "center", marginBottom: 4 }}>
              <strong>{shortAddress(b.address)}</strong>
              <WoCell value={b.hpdOwnerAgent || ""} placeholder="Owner/agent name" onCommit={v => { update("buildings", b.id, { hpdOwnerAgent: v }); return true; }} />
              <WoCell value={b.hpdBorough || "Brooklyn"} placeholder="Borough" onCommit={v => { update("buildings", b.id, { hpdBorough: v }); return true; }} />
              <WoCell value={b.hpdRegId || ""} placeholder="Property registration #" onCommit={v => { update("buildings", b.id, { hpdRegId: v }); return true; }} />
            </div>
          ))}
        </div>
        <div className="hint">Typed once per building. The owner/agent signature is left blank on the printed forms for you to sign.</div>
      </details>
      <div className="filter-row">
        <select value={bFilter} onChange={e => setBFilter(e.target.value)} style={{ maxWidth: 220 }}><option value="All">All buildings</option>{data.buildings.map(b => <option key={b.id} value={b.id}>{shortAddress(b.address)}</option>)}</select>
        <select value={year} onChange={e => setYear(e.target.value)} style={{ maxWidth: 140 }}><option value="All">All years</option>{years.map(y => <option key={y} value={y}>{y}</option>)}</select>
      </div>
      {shown.length === 0 && <EmptyState text="No unit turns yet." />}
      {shown.map(t => {
        const isOpen = open === t.id;
        const set = (patch) => update("unitTurns", t.id, patch);
        const setCheck = (key, patch) => set({ checklist: { ...(t.checklist || {}), [key]: { ...(t.checklist?.[key] || {}), ...patch } } });
        const setRoom = (rid, patch) => set({ rooms: t.rooms.map(r => r.id === rid ? { ...r, ...patch } : r) });
        return (
          <div className="list-card" key={t.id}>
            <div className="list-card-head" onClick={() => setOpen(isOpen ? null : t.id)} style={{ cursor: "pointer" }}>
              {isOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
              <span className="pill pill-accent">Apt {unitNum(t.unitId)}</span>
              <div className="list-card-title">{buildingName(t.buildingId)}{t.prevTenant ? ` · ${t.prevTenant}` : ""}</div>
              {t.endedDate && <span className="pill pill-muted">Ended {fmtDate(t.endedDate)}</span>}
              {t.inspectionDate ? <span className="pill pill-ok">Inspected {fmtDate(t.inspectionDate)}</span> : <span className="pill pill-warn">Inspection not done</span>}
              <span className="pill pill-muted">{progress(t)} done</span>
            </div>
            {isOpen && (
              <div className="list-card-body">
                <div className="inline-form" style={{ flexWrap: "wrap", margin: "0 0 8px" }}>
                  <button className="btn-ghost" onClick={() => printLL107(t)}><Printer size={14} /> Print LL1-07 (this unit)</button>
                  <button className="btn-ghost" onClick={() => printChecklist(t)}><Printer size={14} /> Print checklist</button>
                  <button className="btn-ghost" onClick={() => { if (window.confirm("Delete this unit turn and everything on it?")) { remove("unitTurns", t.id); setOpen(null); } }}><Trash2 size={14} /> Delete</button>
                </div>

                <div className="row"><strong>1. Turnover inspection (HPD form SF LL1-06)</strong></div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 8, margin: "6px 0 12px" }}>
                  {field("Previous tenant", <WoCell value={t.prevTenant || ""} onCommit={v => { set({ prevTenant: v }); return true; }} />)}
                  {field("Date the previous tenant's occupancy ended", <WoCell type="date" value={t.endedDate || ""} onCommit={v => { set({ endedDate: v }); return true; }} />)}
                  {field("Is the apartment exempt from the turnover inspection?", (
                    <select value={t.exempt || "no"} onChange={e => set({ exempt: e.target.value })}>
                      {Object.entries(exemptLabel).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
                    </select>
                  ))}
                  {field("Date of the turnover inspection", <WoCell type="date" value={t.inspectionDate || ""} onCommit={v => { set({ inspectionDate: v }); return true; }} />)}
                  {field("Inspection performed by", (
                    <select value={t.inspectedBy || "owner"} onChange={e => set({ inspectedBy: e.target.value })}><option value="owner">Owner / agent</option><option value="third">Third party</option></select>
                  ))}
                  {field("Name of the person who performed it", <WoCell value={t.inspectorName || ""} onCommit={v => { set({ inspectorName: v }); return true; }} />)}
                  {field("Lead-based paint hazards found?", (
                    <select value={t.hazardsFound || ""} onChange={e => set({ hazardsFound: e.target.value })}><option value="">Not answered</option><option value="yes">Yes</option><option value="no">No</option></select>
                  ))}
                </div>
                {hazardsSuggested(t) && t.hazardsFound !== "yes" && <div className="hint" style={{ color: "var(--danger, #B3261E)" }}>Your room answers below include a hazard — "Lead-based paint hazards found?" should probably be Yes.</div>}

                <div className="row" style={{ marginTop: 6 }}><strong>2. Room-by-room (HPD form SF LL1-07)</strong> <span className="row-muted">— tap a box to cycle Yes / No / blank</span></div>
                <div className="sheet-wrap" style={{ margin: "6px 0 12px" }}>
                  <table className="sheet">
                    <thead><tr><th>Room</th>{TURN_ROOM_COLS.map(c => <th key={c.key} title={c.rule} style={{ minWidth: 92 }}>{c.label}<div className="row-muted" style={{ fontWeight: 400 }}>{c.rule}</div></th>)}<th style={{ width: 30 }}></th></tr></thead>
                    <tbody>
                      {(t.rooms || []).map(r => (
                        <tr key={r.id}>
                          <td className="wo-edit"><WoCell value={r.name || ""} onCommit={v => { setRoom(r.id, { name: v }); return true; }} narrow={110} /></td>
                          {TURN_ROOM_COLS.map(c => {
                            const v = r[c.key] || "";
                            const bad = (c.key === "floors" ? v === "no" : v === "yes");
                            return <td key={c.key} style={{ textAlign: "center" }}><button type="button" className="chip" style={{ minWidth: 44, background: bad ? "#FDE7E4" : v ? "#E6F4EA" : undefined }} onClick={() => setRoom(r.id, { [c.key]: cycle(v) })}>{ynLabel[v]}</button></td>;
                          })}
                          <td><IconBtn title="Remove room" danger onClick={() => set({ rooms: t.rooms.filter(x => x.id !== r.id) })}><X size={12} /></IconBtn></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="inline-form" style={{ flexWrap: "wrap", margin: "0 0 12px" }}>
                  <input placeholder="Add a room (e.g. Kitchen, closet) and press Enter" value={roomName[t.id] || ""} onChange={e => setRoomName({ ...roomName, [t.id]: e.target.value })}
                    onKeyDown={e => { if (e.key === "Enter" && (roomName[t.id] || "").trim()) { set({ rooms: [...(t.rooms || []), { id: uid(), name: roomName[t.id].trim() }] }); setRoomName({ ...roomName, [t.id]: "" }); } }} />
                  <button className="btn-ghost" onClick={() => { const have = (t.rooms || []).map(r => (r.name || "").toLowerCase()); const std = ["Living room", "Kitchen", "Bathroom", "Bedroom 1", "Bedroom 2", "Hallway", "Closets"].filter(n => !have.includes(n.toLowerCase())); set({ rooms: [...(t.rooms || []), ...std.map(n => ({ id: uid(), name: n }))] }); }}>Add standard rooms</button>
                </div>

                <div className="row"><strong>3. Turn checklist</strong></div>
                {TURN_CHECKLIST.map(g => (
                  <div key={g.group} style={{ margin: "8px 0" }}>
                    <div className="row-muted" style={{ fontSize: 12, fontWeight: 600 }}>{g.group}</div>
                    {g.items.map(([k, label]) => {
                      const c = t.checklist?.[k] || {};
                      return (
                        <div key={k} style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 8, padding: "3px 0" }}>
                          <label style={{ display: "flex", gap: 6, alignItems: "center", flex: "1 1 320px" }}>
                            <input type="checkbox" checked={!!c.done} onChange={e => setCheck(k, { done: e.target.checked, date: e.target.checked ? (c.date || todayISO()) : c.date })} />
                            <span style={{ textDecoration: c.done ? "line-through" : "none", color: c.done ? "var(--ink-soft)" : undefined }}>{label}</span>
                          </label>
                          {c.done && <input type="date" value={c.date || ""} onChange={e => setCheck(k, { date: e.target.value })} style={{ width: 150 }} />}
                          {c.done && <WoCell value={c.note || ""} placeholder="Note (who / what)" onCommit={v => { setCheck(k, { note: v }); return true; }} narrow={160} />}
                        </div>
                      );
                    })}
                  </div>
                ))}
                <div style={{ margin: "8px 0" }}>
                  <div className="row-muted" style={{ fontSize: 12, fontWeight: 600 }}>Other items for this unit</div>
                  {(t.customItems || []).map(i => (
                    <div key={i.id} style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 8, padding: "3px 0" }}>
                      <label style={{ display: "flex", gap: 6, alignItems: "center", flex: "1 1 320px" }}>
                        <input type="checkbox" checked={!!i.done} onChange={e => set({ customItems: t.customItems.map(x => x.id === i.id ? { ...x, done: e.target.checked, date: e.target.checked ? (x.date || todayISO()) : x.date } : x) })} />
                        <span style={{ textDecoration: i.done ? "line-through" : "none" }}>{i.text}</span>
                      </label>
                      <IconBtn title="Remove" danger onClick={() => set({ customItems: t.customItems.filter(x => x.id !== i.id) })}><X size={12} /></IconBtn>
                    </div>
                  ))}
                  <div className="inline-form">
                    <input placeholder="Add your own item and press Enter" value={customText[t.id] || ""} onChange={e => setCustomText({ ...customText, [t.id]: e.target.value })}
                      onKeyDown={e => { if (e.key === "Enter" && (customText[t.id] || "").trim()) { set({ customItems: [...(t.customItems || []), { id: uid(), text: customText[t.id].trim(), done: false }] }); setCustomText({ ...customText, [t.id]: "" }); } }} />
                  </div>
                </div>

                <div className="row"><strong>4. New tenant</strong></div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 8, margin: "6px 0 12px" }}>
                  {field("New tenant", <WoCell value={t.newTenant || ""} onCommit={v => { set({ newTenant: v }); return true; }} />)}
                  {field("Move-in date", <WoCell type="date" value={t.moveInDate || ""} onCommit={v => { set({ moveInDate: v }); return true; }} />)}
                </div>
                <div className="row"><strong>Notes</strong></div>
                <WoCell multiline value="" placeholder="Add a note and press Enter…" onCommit={v => { if (!v.trim()) return true; set({ notes: [...(t.notes || []), { date: todayISO(), timestamp: Date.now(), text: v.trim() }] }); return true; }} />
                {newestFirst(t.notes).map((n, i) => <div key={i} style={{ fontSize: 12, padding: "2px 10px" }}><span className="row-muted">{fmtDate(n.date)}</span> {n.text}</div>)}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

/* ============================== boss reminders ============================== */

function RemindersTab({ data, add, update, remove }) {
  const [view, setView] = useState("active");
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
      <div className="filter-row">
        <button className={`chip ${view === "active" ? "chip-active" : ""}`} onClick={() => setView("active")}>Active</button>
        <button className={`chip ${view === "done" ? "chip-active" : ""}`} onClick={() => setView("done")}>Done</button>
      </div>
      {(() => {
        const filtered = data.bossReminders.filter(r => view === "active" ? r.status !== "Done" : r.status === "Done");
        return filtered.length === 0 && <EmptyState text={view === "active" ? "Nothing on the list right now." : "Nothing marked done yet."} />;
      })()}
      {data.bossReminders.filter(r => view === "active" ? r.status !== "Done" : r.status === "Done").slice().sort((a, b) => (b.dateRaised || "").localeCompare(a.dateRaised || "")).map(r => (
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

function QuickNotesTab({ data, add, update, remove, buildingName, deepLink, setDeepLink }) {
  const [text, setText] = useState("");
  const [reminderDate, setReminderDate] = useState("");
  const [buildingId, setBuildingId] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editText, setEditText] = useState("");
  const [showDone, setShowDone] = useState(false);
  const [showDoneReminders, setShowDoneReminders] = useState(false);
  const [showAllReminders, setShowAllReminders] = useState(false);
  const [highlightNoteId, setHighlightNoteId] = useState(null);

  useEffect(() => {
    if (!deepLink || deepLink.tab !== "quicknotes") return;
    const target = (data.quickNotes || []).find(n => n.id === deepLink.id);
    if (target) {
      // A target that isn't due yet, or already checked off, would be
      // invisible under the default filters — widen whichever one applies
      // so the thing just navigated to is actually on screen.
      if (target.reminderDate && target.reminderDate > todayISO()) setShowAllReminders(true);
      if (target.done) setShowDoneReminders(true);
      setHighlightNoteId(target.id);
      setTimeout(() => {
        document.getElementById(`note-row-${target.id}`)?.scrollIntoView({ behavior: "smooth", block: "center" });
      }, 100);
      setTimeout(() => setHighlightNoteId(null), 2500);
    }
    setDeepLink(null);
  }, [deepLink]);

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
  const remindersToShow = (showAllReminders ? allReminders : allReminders.filter(n => n.reminderDate <= today)).filter(n => showDoneReminders || !n.done);

  const NoteRow = (n) => (
    <div className={`list-card ${highlightNoteId === n.id ? "sheet-row-deeplinked" : ""}`} key={n.id} id={`note-row-${n.id}`}>
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
        <button className={`chip ${showDoneReminders ? "chip-active" : ""}`} onClick={() => setShowDoneReminders(s => !s)}>{showDoneReminders ? "Hide checked off" : "Show checked off"}</button>
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
      .wo-card-top { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 12px 14px 6px; }
      .wo-card-identity { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; min-width: 0; }
      .wo-card-actions { display: flex; align-items: center; gap: 2px; flex-shrink: 0; }
      .wo-card-title-row { padding: 0 14px; }
      .wo-card-title-row .list-card-title {
        overflow: hidden; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical;
      }
      .wo-card-tags { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; padding: 6px 14px 12px; }
      .wo-full-description { white-space: pre-wrap; word-break: break-word; margin-bottom: 10px; line-height: 1.5; }
      .wo-tenant-info { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; margin-bottom: 10px; padding-bottom: 10px; border-bottom: 1px solid var(--border); }
      .wo-tenant-phone { display: inline-flex; align-items: center; gap: 4px; color: var(--navy); text-decoration: none; font-size: 13px; }
      .wo-tenant-phone:hover { text-decoration: underline; }
      .wo-filter-divider { width: 1px; height: 20px; background: var(--border); margin: 0 4px; }
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
      .dash-storage-footer { margin-top: 24px; padding-top: 10px; border-top: 1px solid var(--border); font-size: 11px; color: var(--ink-soft); opacity: 0.6; text-align: center; }

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
      .sheet-row-deeplinked { animation: deeplink-flash 2.5s ease-out; }
      @keyframes deeplink-flash {
        0% { background: #FFF3CD; } 70% { background: #FFF3CD; } 100% { background: transparent; }
      }
      .sheet-row-selected td { background: transparent; }

      .wo-sheet-row td.wo-edit, .wo-new-row td.wo-edit { padding: 0; }
      .wo-cell { display: block; width: 100%; box-sizing: border-box; border: 1px solid transparent; border-radius: 0; background: transparent; padding: 8px 10px; font: inherit; color: inherit; resize: none; overflow: hidden; }
      .wo-cell:hover { background: #F6F3EA; }
      .wo-cell:focus { outline: 2px solid var(--navy); outline-offset: -2px; background: #fff; }
      .wo-cell-bad { background: #FDE7E4 !important; outline: 2px solid #B3261E; outline-offset: -2px; }
      .wo-new-row { background: #FBF6EF; }
      .wo-sheet-row { cursor: default; }
      .wo-sheet-row:nth-child(even) { background: var(--bg); }
      .wo-sheet-row:hover { background: #E8F0FE; }
      .wo-sheet-row td { padding: 10px 12px; vertical-align: top; }
      .wo-sheet-expand { color: var(--ink-soft); text-align: center; }
      .wo-sheet-apt { font-weight: 600; white-space: nowrap; }
      .wo-sheet-building { color: var(--ink); font-weight: 600; white-space: nowrap; }
      .wo-sheet-desc { min-width: 220px; max-width: 420px; }
      .wo-sheet-desc-text { overflow: hidden; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
      .wo-sheet-tenant { font-size: 12px; color: var(--ink-soft); margin-top: 2px; }
      .wo-sheet-vendor { white-space: nowrap; font-size: 13px; }
      .wo-sheet-status { display: flex; flex-direction: column; align-items: flex-start; gap: 4px; }
      .wo-sheet-actions { display: flex; align-items: center; gap: 2px; white-space: nowrap; }
      .wo-sheet-detail-row td { background: var(--bg); padding: 14px 16px; border-top: none; }
      @media (max-width: 720px) {
        .wo-sheet-wrap { overflow: visible; border: none; background: none; }
        .wo-sheet thead { display: none; }
        .wo-sheet, .wo-sheet tbody, .wo-sheet tr { display: block; width: 100%; }
        .wo-sheet-detail-row td { display: block !important; width: 100%; }
        .wo-sheet-row {
          display: flex; flex-wrap: wrap; width: 100%;
          background: var(--panel); border: 1px solid var(--border); border-radius: 8px;
          margin-bottom: 10px; padding: 10px 12px;
        }
        .wo-sheet-row td { display: block; border: none !important; padding: 3px 0 !important; width: 100%; }
        .wo-sheet-expand { display: none; }
        .wo-sheet-apt, .wo-sheet-building { display: inline-block !important; width: auto !important; }
        .wo-sheet-apt::after { content: " · "; color: var(--ink-soft); }
        .wo-sheet-desc { order: 3; margin-top: 4px; max-width: 100%; }
        .wo-sheet-desc-text { -webkit-line-clamp: 3; }
        .wo-sheet-vendor, .wo-sheet-status { order: 4; margin-top: 4px; }
        .wo-sheet-vendor:empty, .wo-sheet-status:empty { display: none; }
        .wo-sheet-vendor:not(:empty)::before { content: attr(data-label) ": "; color: var(--ink-soft); font-size: 12px; }
        .wo-sheet-actions {
          order: 5; display: flex !important; justify-content: flex-end; width: 100% !important;
          margin-top: 8px; padding-top: 8px !important; border-top: 1px solid var(--border) !important;
        }
      }
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
        input:not([type="checkbox"]):not([type="radio"]), select, textarea { font-size: 16px !important; }
        input[type="checkbox"], input[type="radio"] { width: 18px; height: 18px; flex-shrink: 0; margin: 0; }
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

// A crash anywhere in the render tree (a malformed field, an unexpected
// data shape, any unhandled exception) would otherwise take down the
// entire app to a blank white screen with zero explanation — which looks
// exactly like "all my data is gone" even though nothing on the server
// was touched; the data this app saves lives in Firestore, entirely
// independent of whether this specific render happened to succeed. This
// boundary catches that instead of letting it go to a blank screen, and
// says plainly that the data itself is safe, since a panicked reaction to
// what looks like total data loss is itself a risk — someone might delete
// the account, wipe local storage, or take some other drastic step trying
// to "fix" what is actually just a display crash.
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false };
  }
  static getDerivedStateFromError() {
    return { hasError: true };
  }
  componentDidCatch(error, info) {
    console.error("Render crash caught by ErrorBoundary:", error, info);
  }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{
          display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
          minHeight: "100vh", padding: 24, textAlign: "center", fontFamily: "system-ui, sans-serif",
        }}>
          <h2 style={{ marginBottom: 8 }}>Something went wrong displaying this page</h2>
          <p style={{ maxWidth: 480, color: "#555", marginBottom: 16 }}>
            Your data is safe — it's saved on the server, not in this screen. This is just a display
            problem. Try reloading the page below.
          </p>
          <button
            onClick={() => window.location.reload()}
            style={{ padding: "10px 20px", borderRadius: 8, border: "none", background: "#2F6FE0", color: "white", fontSize: 15, cursor: "pointer" }}
          >
            Reload the page
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

export default function PropertyOpsApp() {
  return (
    <ErrorBoundary>
      <PropertyOpsAppInner />
    </ErrorBoundary>
  );
}

// Background images of the official HPD forms SF LL1-06 and SF LL1-07 (Revised 3/6/2020) used by the Unit Turn printouts.
const HPD_LL106_BG = "data:image/png;base64,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IoouEVF0iYgoukREFF0iougSEVF0iYgoukRE0SUiougSEVF0iYgoukRE0SUiougSEVF0iYiiS0RE0SUiougSEUWXiIiiS0RE0SUiii4REUWXiIiiS0RE0SUiii4REUWXiIiiS0QUXSIiii4REUWXiCi6REQUXSIiii4RUXSJiCi6REQUXSIiBW3PfLe82dm207mdoWjKO1E2DuVdtrcM1k6v0sveFdu4E+xtOChN+TUMiS3sg63sPvt5J8Pa3vhQ0S4iO+X+VV0wisjrdME4u9nZWu7sNq5kOQp2omy2dyP2ckk0dXbM7dmdYHpm99zm7wrW/rmd2H2mJ9x8dE1sdqlvfPj5Np5jZxsTO7HH3nUndtsx9P2nO7xK3/nFiT27E9q/u+WD8qVv/q7w7tsTO7H72r/7zOjSBaOI7MW6oIpARBRdIiKKLhERRZeIKLpERBRdIiKKLhFRdImIKLpERBRdIqLoEhFRdImIKLpERBRdIqLoEhFRdImIKLpERNElIqLoEhFRdImIoktERNElIqLoEhFFl4iIoktERNElIqLoEhFFl4iIoktERNElIoouERFFl4iIoktEFF0iIoouERFFl4goukREFF0iIoouERFFl4goukREFF0iIoouEVF0iYgoukREFF0iougSEVF0iYgoukRE0SUiougSEVF0iYgoukRE0SUiougSEVF0iYiiS0RE0SUisiVtz3y3uvnZ9nZtZ8BWd6Jsuqnusr1lsHd4lZyXvSu2cSfYWz8oHav6GoZEt6nuxO6zcZ59MqztjQ8V7SKyU+5f1QWjiLxOF4yzm77Cc2e3cSXLUbATZbO9G7GXS6Kps2Nuz+4E0zO75zZ/V7D2z+3E7jM94eaja2KzS33jw8+38Rw725jYiT32rjux246h7z/d4VX6zi9O7Nmd0P7dLR+UL33zd4V3357Yid3X/t1nRpcuGEVkL9YFVQQiougSEVF0iYgoukRE0SUiougSEVF0iYiiS0RE0SUiougSEUWXiIiiS0RE0SUiougSEUWXiIiiS0RE0SUiii4REUWXiIiiS0QUXSIiii4REUWXiCi6REQUXSIiii4REUWXiCi6REQUXSIiii4RUXSJiCi6REQUXSKi6BIRUXSJiCi6RETRJSKi6BIRUXSJiCi6RETRJSKi6BIRUXSJiKJLRETRJSKi6BIRRZeIiKJLRETR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Bt3r2ov6++H+/+ot0MALpr0zQc/TZVP89OQTahC4Z+7/BHhkjhws1de+LJoeOmMWStB/2nI+I+LI4Y8Ac6S/XAeCb77r3IfbR89Q+oIFfuebzp3ip5NXIjP287cLr87mvx1Ef3fyCTUavnfacj4DmLVOWx2tt7IGIZ6aHDoZ9NRvTnunraU3if2hM6XZ/SujqzZ02pTu88gceaf0VuGGVj08ctB03OiKffxv0QgP9lnO49atjvadKT0l30RuHz3zpGVrjA4v1brkBUkub2ynRiMo83ja9MS3yC/3QoBPjfMwgEmcMCakcduUotV6M06fXH4amMZVaASdzsOAqJY0S5h8Uo5vAo9DZ+EqPApKpcincRtnvtZyW+s2znytPt25/1ZL3Wgx+990bW6EnU7SL3Xd7+y43HqpCMTUb5ueeIr5oPOtK8C9sDR/dZW1vhd0lhZ9GrdMT8dMcZG3TMmZoVKHKcuNPw07nfllV8lP/VJ0mXwTeRQY51Fha0K1lBB5wUqAcYDeASf9v/SltrN0Jb3CdFXSkYS6BpxscrPs4yXAqjo4QG/SK01vckFoVV3g7JCVdpTTO9D8p5TNovlGr0spmWcJsM4OpSvqZGtjVXuyuk3VXe0GfLot+epYA/l9+uSf9KcCU003r7rsjplVdVu2qiddw+yl3/nQpB/pddPSSDa5lC7DODqwRHbsNsDZ7ZxbHkEbXQl395XM73yoo0MXjCLPPLTUCakoumTvqUQqA1F0yZ7jKrrkBbJVBK+iN+a2cWb//F/CzXzsK1Ph7isY81/+q44OERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERGR14Nxkn9LKgqRV5T9Km7UL/0//ncA8z//n+GLKbQT//k5Uxx+O9hc6B4eif7ruiYc+x8hYJ34H+GLLMm+E/v/dtlLB94JXsY+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+Bc6C4Q2czZSdV1peaYQQ+w3Ns+e0qEhYmjowyBt7St/os0fw0brbPrgsShJFZWae/Pt2iTco2P734I52GRoKF54FVYZG1V054NsUSwPVUqRsKMZATpBGPgXYR7U6Gjy8O5VxsaSp3hrQuO9+VhnXRV9cRZyg9LtxW5ElYdQmdrE7dZFNrORF/1qbIzPoSC63cUzCaweTbz4KBkgQLxc/IpHQJR+VfybPTwtTlQ33/SRHFONenMy0fJBX0jAuA9XLWzdhCWroptxmlXtyvhhQLyYpjhyv7lwyA0Lm7d3ZpZKvRTOQ/xJ6atFAw8MTxYgvpMfnx0Sz8BWelYL0S30pileYtXAapvOAjpX7IgY4tyljxd2X1kIgyi5vo6XgQ02wBRD3hggmidx3rtXTXZWAzcMbN9JfllNrF5+gnQveYRXiiHKQrGFK82j41mIFohN8kO4gH5gLGza7OqRjQyUm/tkHuLV28QzSTxCuJAO8vAmiq3kMKW7cWceqiQbUvPJK6vLbHqi+TTgIVxIPsQY4hmrdHwnK3Qz3cyaVCXrxUJE1W1+pjZLtqs9MGvTPaemJ91ge63H9uLN72YGuccm+5FYm6ccMo4kKS5nRNOX6Hr6D/D+OTXiAcmKf3ZoMQ3jjDSzIZIe1j1D928tOUth45hKyjEwovPzZ/98/6+pDecDTwMDw93PyT9p8fTEycV1XlENMNDsqUxfIQemBoa7VS4+mSA72ClWWenSQQPDeSFnGOCV7K1rfYW/qUTLQlE6q2P6SzVtwYCXlx5oJsc7I7rUoHJzsxpmXdAtR3UkWwWtXNbq9Xd39LnefbClxZy8n9PeSZ4/NjKSFXvl3VQE2YkGsp/6ur2ZNb9Xw9I+cNIiueh2F/V3l8r75dzwW2+Isfk98v65o7jq9//6uyan19/5gx0/7731/l887bO+8u6/AdF9U9yet4Gc8Z+KLydGSXianqu39Qdas3BF6cLNgCK636Pob5T//07wVOYiGW3C05xgj27P5AHEH63p9H2LIrp/DYjpEp4mzh/w+FsT0f3rUjXhUBzNfB1/18RUj4yohYhOEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEISc/w9TdGd0mkIaVwAAAABJRU5ErkJggg==";
const HPD_LL107_BG = "data:image/png;base64,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2dAr8i6U+or7AFniTiacyQH8vv+9OUEQDwLZO1eUeuaA7vCrpFOCIHZS6Jq5vg/mAql0B0EP1m0dTbI/eLBSjyQt7CSD0Vt9t0zbvXHx1Y9m3QlKggxIEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARxX8DW25l7ePVi9J0qKUeaud/IUbt/ABwtu97O/e5Dq7ff/NWdKunxF0gz91l3fZziwg7gSmX9/ZxURBDEzoNCF0EQBEEQBEEQBEEQBEEQBEEQBEEQBEEQxKY5JEkHxMMEP3T0ts5nB8V2Tut1197XUY7/zw48wKrf8IBisbDp0roFOfNtK/G2vX0tf+UFsCIZI6DUUwD44G2W0tvdcVs+z3q2dXpxnTFCWyKsZVa7XWYrsTbfd/8aMLvhQMoFFl4jT78tOl+4g0rkPQvV9Rtlqfn2zMB3+cCrZA1rkgGYC9VMSarbijzlmZ/onVXxTGkr8b17fOeOurA0r/MvkKvfFsKf1/kjm+pKNh7hG4SNhQ+t8stBAHWzRjsz6//a/qEaB/f7U8t9MGv/xndTix8cY3Mfck3aB7bmScZsclAzAECb+9eEG466cOM8H6Jpxm2yMsqObaq34xur2h+ph7GoKBpaHisAuHGy+Yk5A2bIFjbGt55SmT/Cyok1Q1fx1mbGYx9+Hg0ut+ZJKyc3u8gB4Np793Hvs4lj/Ar90vG2MWcz4g4VVd9mcxIGoNBlcTwF7/zGutxowrgDHXErVb6fMxxtKijVGQA4PRIAF7xbALxUCDt5LgA43YWcAMB7BLWL5h0ABBy0dQEodgurt5JdRbXfHbR1MYGcYAUAsLsKKBUAVuoOlh5yOYFiYAwmkMuBFVCSgXWCLRIAKwWFRhbJieVcDmQeG8aDzwIA3lPI5WJzWM3lwHKAA4BJhIYrxfehSt3C6lsgZUsnsmX4f9RwwkPgiKDpOAIAt0eFNraFI2XCbgFWkoE5o1FWYHKWE7FoTlhIJFqpKyFX0GYdwWwxTrcAHMFSjginWwTncxlLVgglvV8Gzhsd0LHrZ8g8+yOg80+KT/3yF9g75B/wFB/6zFMf+5/IvPxDYP8f/R3aXy491fPJ6udqv9/d9dNfU8to4Hc++Xdgn/35L75c/6NP/Qidg6Xuf9TYe+ifjz71iAI6B0tdP/01vlz/o08t/Hve47798g+B7Ms/BMCP/93L//pHOHyg9MQ00Hm01NPjvTv8c535D08+9fN/+k94qvuHmZdm/viJH/cOFp/6uX7kP7Ie90ePfPXHYIf7S4X5X+PL9UOhRf6d+696ujLP/JjMASDr/ggAMi//EHyo+6ke90d7h5aOPvnLX4Af+0xJTvM/+99+u+eXv/izX/4Cj3z1x7CG2/+Fkpy25+//Qqnrp79+5D+ynmd+DMS2zL9c6g5s+aPsnzzV/dNfxw3Hlg/g01/8O2vfLz52/bP/c+gpOQ3wY58p/eug8L1HS/96OmnCHjm9/wtP/Vzjy/Nf7H7yp7/G097PkP2TJ5/6uX72/8F7nvlxKFr2P5WeuPHJvwPwxL/nPe6Ps/sXjpYScgHY9eKP8EX/05958qe/RvvLpa4f4dOffPwLT/7yF5EjZv7DU93/+LF/z3vcH+/9dz8OJePHfzz8VFj9HTPqKvjBuuZBAJmyBnr3TKqOeJLNBhbHsKSAjt21jEsNo/ncQkPkWzUyA4tTvB9A5si87gF4/+LlTBkQ+Va9pE297jMJcN+ets8Au7omptoleP/cGJY0AHS0XLouUEe9DqDfINM/N8n6YOqmrsEAtLtz021lQOQjiyzppXm1pMgSALDEy+GMr2P3hDI1IHNkfvkA0L5ncjrvmhrqdR3NB0W+VSPbNzuWt6ft6lucyhyBqaNup1SBLdnRxcs3jwTHH3js8rsi0XBs+QAUBziTYFIDnS21vAt0yLlaUHjmyGItL1MmzJfK86wPEJ27p3jwwMLRxy5fH0Bdm3otEu2omW+XgbVNvQ7w/pRcIfln5nkZ6F+8fBNAvjy/fCB2xIOtUzfLwflISNZrJb1vep/NhLfyPJBvOaHbX5QK2RtvtPLytUt8uO+N6IjWEc2frADy1PVjXRVqGU07CQ2UL/wEHf5pff0FoZGfOdv+olT5lhP6gz8EUL7wE/PGn40AWioUgtD1/KlldN+4DKfvjUzLm7r4L1UOwL06Ob175bVvndIA3z3CzMxZ/MMf4dbrfz5iT+u7cRrsaQDlyCIXP/9e9YlPnSdLAICne1UQxbtmL+/6+hkgO3N27wuAf6HKCrJ66mvXKomlrPKFn6C0cgb8sxUA6F45iZvP4dbr3wj0Hdgy23JCv/sVezyTl6pTLYmGY8sHsMSBAgQYFPDp0XeH3Srru3aOHeuuAECbf1rvT5nw3eGDM+ee/RyA8inFn7YBbs8p9ZM/F9eW/3gEoWgZeaWy98i/AMCNN/5sBAC/Ndr+otCRXAHdb1U//zS4/L6eB1C6du7xrwodOGLGvVJp+3JwPhBL9tzou8NudQeNuvYe//Zj54GCp7GgXcCMGZ0x4/ArTlyKrzGdA7CgTJXW9JvS7WnAqxq4CxqzTAL+GBYgUVjQmOOwO22vXADkkv1/QdWZOw9UcnA8japdlBAafjR+mtV1/ywwn1gz5nIKGMukLSJAK13RyuM4H7IjCCan4XEB+GOYYwI3qjDVVU9yelUDdw6Y4gDA3ElgMnlvL7Bl0dNY5ALwqoYJDV/HDScoH4DPBYSW4L4GZpVfzYGLaZhxu3zTtaBxVSVN6KuWcUxzAAsKwVFtvsIKCwwaiMZNBTOplXUzhsWkXAErVUwzMKOxBMA7h0UWOSL3K1h8I3alSDIr6Q4KXSyHzBFAzgNGCcBTgONpoBafy4wde2MKqFG7aKLmnt7yHIBJ21Tgawl4GkZJyBpguAAmw6OVA4SPSk4BTFRtTw0gmMPoZNsK1h+SNw+5UYDPZMoiQqPxYYqHl5lpfihqygbCmiOI7au1NAkwqQDD7Bq4ArxkRxDacsnej8EkYCCBZMOJyjcQkKpgFwWmAc3AfQUsWePKGmDSJtSehh+4g2/jVc4DjJJB+LWiOT5gUuMiT8HXiOUKmAPqDIbZ6L1idwaO6PgpDcSSTUfut0MmjNcqbJ8rluUVAPpJ20KEQWhFa3oG5Ey0qEOsUvNBeOPWIxg0AB24kf445EKP1VnkLvMcXAShSwMMJdvHsEDzgHq+K14wVQDACsleiBmNqCEGFmFCQ7xDtghc9iK6jrwRjEZZoHur3Fxu9eBUAwxtPeAsmvwb3XCY/jjEBz12eKsBU+1VCixqOHHv4mt5XYw/jYKJCndsixLa2qnRhDrRsHzbAQnek1hHtaIJv2ldRSxXQgGArw7qwM2MEqEjivQDNKFkN5vH9Ps5dJk6LnW5f5PSnmjo5X0ulVun9rD2/GTaTOpkHNGPBl7LGPatOpqLf4rng2A4CMDDIhfLrh1CTbiHROonGvmjAgm/ZY3DMIAV4TjCoXFXFLtynwQAUy1U2/yEVnr7gd80PWVfpFT7vFNj6GJAR0dipOYeu1BNh52ws5F/27LEIOZTkcQOjNcxYToeZQ8CeG+VaM1JyhUz9uLQqIovaB2xIXSFkt2874YDm2p4Sv7N+gewoeuFt6g5rMnKpXV2zqqGZ/98yJ+m+s9LAHz4fGh5zyk7DzhztLemEnHpaMvEzfK6ImQGMQT0LlXJGkGTHP+q+ACA+ndDzmK8ub1/rpZ/sukZEzodH3RzWyIwzMKZA4fUcrOC1FPcNxzytp5VX6nYmV9T0daTK+bG2YHBV1c54g6ZNmU3d5hIj7R0tMAVNAvvg8IMtYnNtZhU9yrqBouTjYco+Ss/ecacBoC8x+XVsFmcfNlNeGK2dURnyw0XYam25U+50+iZpEFXhF1aYuVq1+LZeGvx1mk88WTSVvEKUTxCEXr1PQ9RB7ykLWf1S+7frJqiAKgxZnxWkFcaxslJgzUxYWpBRvuTTUVrjjfZfMXv7ECwqCreCx1RN72e3qmhS9mJvazFQ2MwE6mxzRsRNF/cbOgSQLToJN+DWr2yUv3Cr2qpMzQA5Gb/99bIg7zqZ5ILEl7atQwTujF0jT9ziXdfIgM0tsyWC1eTvivnV3fbCcPFFuGNoxj5HvTHU1sWqrKSajjRioCzZFTBxAXUGeyNesBoWW1uQitMsC25Yh4tia11161BrkTs0lKBAUxcCR2xYSk+lOy+e3BgUwJxqaGKAA9vey1lBNDlw0CCuwC4oci12dCluoCMUEBWgAsFVVg1JqvxZC4Wo6V1Wyc5rdSpxhT2juFSLpNAagXHdjX0E8ZwbAXA8TUAxmS9ySTDxgclgb54LBzP6F2AJ60R2rJtda+fbDhR6JIayvFTW4CiH58hmpkQKIZio85ju4eiLXEgeNY2beomciUmvVkgAx06Yj0DwI3OT0q2w0IXyzn7TBXTWYmOsJvwWR8y7jyM7sM+ASxljw8NSmoQmxvB5gV6fAWwAXQYhVpWgkV3i7jIAX6mNRm6VA/A+1BrPz50QAAAHwBz6wBkEKrqXLABAIbZLb7qAevz1phrEOjsB7f68ZePDQ92Jdqyg0wZAApMQJfQHvl1tVOA9wOAqXYIdiA50g1sOZ+RYMFtl0wfmFTphhN1Rq6CTsYDT/eDl+2zfCoj2SHZ3IR5N9y2lHGBDoE6FyIUzWd96E15UkBSrkS87QKXGuBH0ONHjuixI+jsi85PSrbDJox79wI3NBbZMVW4ocPZijvktLwJUy0fk5Xfw4rO5VA4Qesom2GyPHy9cA2A1348d0tj4eaQch6bDDuF4aXX4CPV2465w9r5rfEF5HLFrlcBoHN4Wb4JX+3fax959vjgzd0AfH2sfgIApo8M35TpGyeMct4kAlR5H8MVADCVFxyn1FqJ1qGeGWxTnwBUeehaddoddqY+Ha4n8WHlfHQMAKbcoeXCleRakrXlyvKxGt89Ybftz/PWarrhxONojRpLjPZMtTzMW+09mFl+7Lp8s5kJAXX4iWCFbEUd7mSF78Pnxz76vUA0T/W2y8rvJTwpICVX7BGHOrhfBby9x3PvI3RET+09nns/Oj8p2c4adWmgPnsW8M6yQryaeUkVWy9pYEIXrmqNbMvIyBkuobH5VGYPZYMBgJUxXlgYB+Cfy/lnAX/sZrF1MtjpTfNWwFepAfqtcafojyO/MjJyNiMB+ONOYUYBU6zVzgy88bY9ZwCYcdYCA2Bm2pEzVSQtkqMJY8zVGsNMBTBg5QsjJ1SctWH2emnlHIBJXdBYUM7CeGS4s7zYMmYXscadwmwl4eyBLc2Z5eKewJb+dFHO6IaGE0YgX8FPPb19VTmtwT0Yb4zJGZUyIUxwrdnrcqECaA2cXi7KWQ1P5eqRaJeX5VVlw6Q3zVutvbWO5Aq8QQeHXC/umQCwMp7zz8WOeFrnFs8FnpiWbKfiFAEg+OkBK9pPpwAGPPunAL70gp2QCGoYa03Ugk+nAABPfAOhEosi2mkV+/mvBmfkEmd8/gWAfc1FLioCjgj1XRT2n2K4xWbDT1iECbJNwiRWPwLZbwN4/JvWsXMAWNHqySnEX0LDFSMFBgYIvse2ZAlbhkaKG044hYu/5MJvLJFvPjgxNiET9lqvuPbtAizRBu2WQDTbHmNJRMPBiKsjwms++3WWcsQgSyEripRkubj1P0B8/gUbuogt8MQ31m5XXys32fo1N/xD3Dke+Wb45+7Y8o7yyt2w/rNf35GGu0O3PHOr76IQt2EV0UyXugBk6Gn4O4zhAuiimfSOI3tHSlHPHNAdfpXUeadCl98sdKnnDyyXPOog7iw+H5x2uq+QIh7O0DVzfR/MBdLmnaL5YzQT7j74F0k7dxZvvO8gZis7Q1i9g0rdIbBSnDKb2BzJV3027mquTF6id/TejY6ip3AXbXlHKd0N+7dR0yUIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgrjv4YN3sKxDg2LVxkOy2aHsoFhvd4JM/6av3+tuSsqD2EjIfF8s4JqSr/F927DeQXmfOMUaimymOetBneW7d/F961v10F1QWsdQX8IBj26/oHzftt30AQw1mzusKDcZoVhp6xEsO9B8e2d3aXUz7m7aslmPWG93MnR1bdpji5vyYtaDjYRs6woFXB2YIsnX+L5tsv0l9zaj8nbYgiKbac56kONu1oMaK7L68o0Xb5PrRo/uzfaWm49E/Ggx4XS8uH3ltnWt0zbvZLx/cEIXH2owUKa0frPeaH+KUqn5dvfGaX2v1NB9b7ou3nPPRkJFb1St3UDd+12Rm/Sghstt7vKZ0p2r0ibKypjR8Xtg8O5uCl3Njmo4zpj1g8pG+yOrDgKom+adspyau1ehy+j4Smzo7oUXg3tVI8iF+eqaO7W+B4oEgM4XtlvQ5jxI65S9zKYqZpo7XG95G76xRllJHG9+ersD1i2g682b14NKdnOhy+ci5RMrJ9Z3kY32R65QAHDjZPN9966Zw5xK2L149a5dxx+p36sqCbXOzgvmXigSgFPYZkGb9KALJmWvxsuvUXZzh3Me3YZvrFFWyhKbVHbxlrod1V9obd68Hu5RV2ERDT3CRh6ylRa6Zv9670IX6g/YdTYSQ9/3FaxvqyKbOsvU75wkmynL3BObGo2His2FLjmvgtBVBABHoADA6RYAl8FWVkqucBYARzC7XmGPEyjZPsCxx/EeCSaQy4FJezoAOODBOjXLIZcDUOwWABzBE1N5xy4ZOT0FAGClZOfCU/vCI8Nrimg2CnvBsHy7qQtAcF0AYPZsXioAYLlcTkT1dNDWBQQ1DC8TVSO8UiBFomdNaiI4NqnD4HAn1mWxWwCOAMBFeASTrEc4Vi/x8WExTLIekQPLxeLY8sJvgCPACmiLVGr3hMYF7ykEOrf9dmA4u0xcDGSzdeYCJQk4UcmxIgHkwHIiViRiH7GaY4XgiqyQ2tvgQaxUDGziWNtFfuKIpL0aLh9qkcngzEBMZssOjNtTsOeHsgamSyhrlW/E+gkGNqWEk/LI5tZVcsKWm3LUoKqBhPa4lMXC/1gh1G3CJwLPCb3OaomLxMG8p4CgeSUEeqDIbGpad2By5ZkfAwA7/stfgH35lzde/iGyf/JU909/vfeLfwc+/HPNj32mJKfBP/NDAMi8/EN8+ndbvvDUz3V43KF/PvrUIwrIv1zq/umvwYeffOqXS/+R9bg/euSrPwY/9pnSv54Gvlw/1N31018DyP4nPOX+CJ2Dpe5/1Pj0pz/d83MN4HNzv0D2T5586uca+cEnn3pEAfu/8OQ/l/7e9jmfm3+u+8mf/jrcl/kP9sjgmkPdXe/Yijzy5bk/Lj3+M7DDB0rdP/01vuz94pGvLBwtPaIy/y+U7KWe9nd/5smf/hr5P+l+6vGf4dmjPT3uj9mhA6Wun/4aX67/0ad+hM/3P/W/nvxheBkA+79QemIan6v9fndXLEV+74/ZZ/9eA+DDP9d7h5aOPvnLXwB7j5b+9TT77N/rUIdRTQJ5AaBzsNT1018/11sFDvdW2eH+px7/GR75yuIXHtvf/jPgi48poMEUj3xl8Qsth8Vv9/z9raHPlJ6YBjL/4anuf9Q4fKD0hF11+eInf7brqzMvlx7/mbXw4QOlJ6ZD4/4iM/zkk7/8Bb7oP9fzsz/5IZB9+YcAgioHBkHnnxSf+uUvsHdo6ehTP/+n/xQWFSsSAJ7/N5meT1ZDRWL/F0oyWPaxmsu+/EPgy7sUHv9//jjYyz/zQ/zOJ/8u6UHs2Gee6un+IQD82S9/gUe++uPIT774yZnIXo2Xt7V+2vvZR776Y+DL3i/Q/nLpY7d2/WzXiz/CF/1Pf+ZJ6xNP9bg/AvAV2dbjqdB0CWWt8o1AP5/+9Kd7fv7LF38E7P9Cqeunv8YX/ed6/v43Q91P/vIXAMAO95cK879+/t9kejxlneOpj/3PzL+JT2h/udT1o0At/M/+t9/u+eUvrAMByFjbZV668h+eCmwo3937YwCB56B3sPjUzzWe9nf3P/Vzjf2/X40PHn7yyV/+4pH/yHrcH/FIoIdw1MVblbIHGiUBLhQYcOCxy+/G/UD7nsnpvJtYqAKy5XnWHx2XOTKvewB2dPHyzSNAb+vUcr+pm7oGA9Ah52r5MiDyu2sZNxiJ1+vIDCxO8X4AWRmPh48+dvn6AHDUn1Q9QKZv6fqBeF149xTvi/Z1tFy6LqJrdsj5D4KlWMYH5lSHQKZrboqXAQEw1j+ve2DqqC/Za2WfmedloKdlUnUA9XpdA7vcuamMC4h8q8Yud+n6AMLLAOB9s2M5AB22FoEUjd3FkfnlA0DmyGItn5qHhzUJ5QXA+xcvZ8pQDgCp0P7MxHSnAGN9dagCwKxW0qZgrK+Omq4vYd/uyWpeAAdbp26WsatrYqpdJjqkJdVh/7V7IuMebLl0/SCA9j3aZxLgvl2zybfqyCC8359aPggg88I86zsaFRUr0k5iTF1Hisz2zY7lrQECzRkmwYUEcn5yb4MHte+ZUGY+3BjIkvSTpcg3kpe3tbYbbZEDfq0j/JK3MnXsnlCmZuee9XpsuoSyVvlGqJ/QJ3f1LU5ljliFIb97omrPa3fnptvKqFsdANm+xVpH6oT+xcs3Q7WYGup1HToQItvhoK/yEtjXOh+2L+s5mf65SdZn3ZT1B9UMDrYS2uYVC/QQLtNzH0tBjFOfAbivMwCTl6pTLfEC9IUqK8hq8rTS2+PP/n50XH7mbPuLUmVbTuh3vwLmzlzMfOXW638+Yp2q79o5dqy7ApRPXT/WVQHgnfz2KY0O/7S+/oLQKF47Hy5DZvacUj/5c6EXLutd3xC6zR9dPhbZRp5S/Gkg2OdenZzejeCa6Loxmn9pJaj4rdPZb8oq3h7Hzc+Gm0bbXxT6xF9cCJZLi29VP/804J3QmW9Jde0aPv8JmIvTcLoqQPnCT9DtneDDArCXAZDxzmARgLS1CKRo1PnM2b0vAG3+ab0/tSOsSSgvgHzLCf3BH6LGbVDpm7nMCuXzyK68Jj72NMCNwipTILvymtBfeq+CpVHFXbeSca9U2r6M7huX4fS9EQXK3/lu9ptuFUC4JzAud9+qVr8lFbqu/K2vpULBhi6UL/wkMki+5YRuf1EqZK+de/Z5/+Q/fVNGi8yBIgFcePYzI7EiSytnwD9bARBqzocAgwPIWnJv2oNQuHE5881LaSVGfuLH9kpfPqh1smsQb1WfLb9jv3RbmbpmL+/6+hkAOPX5R8/HpmtQVtI3Iv2EPtm9chI3nwPQdeVv/WdnLrM9AIC+G6fBnsbby//WOnmPd5LtF7+MT+Dy+3o+Usupr12rYFfgQAhtB/67Jz/4llTMnTnX/uKtwE9O6/0wM2fxD39ktfvEF4OWag8OJXz9z0eAQiTQQzjqcnz4zAYHxQHHBwAmNPy4Vd6owlTTc2pvHFM8Os4fwwIkip7GIheZ1nF4Z2MxxDTMeAbAgjLVhFDugsYsk4A/Hi1DtvkKK0zgjIYHgYKnzVh0woLCeAbhPqHhq/CaXE5hIQyuZgwr2sXKODDNok3p2xErVUwz4JKGpwUA7k7j1jRQZQC8qoFcgH8OCC4DgBlgPq5FIEXjbcYxzDGBrgWNq6lGF9YkkBcACgsacxw+F+BGcTkdaHkSep4D3NerTQFMBnc4ZhR8JcD9ChbfYO48UImXhXA5qBbCPYFxM34VnpKAV/GhCoBcCgxaNZFBCp7GgnYBfxxTbEIn67mWIt05YMoaN9CcURLccEDq5N60B0HW4C03qLHBT1Lr1fbyttYpP/aquJqWiclpeFysKrhRWckqRfoJfJK5k8AkDxQmNIyyK5hTwFhiRUbOwrydPIEZjaWUWiIHimwHTGhPSXAxjYXAW6zn+GeBeWZrMhdWwR4cSQjrnEY9rKFLLsHXVhFLHJC+vW0iN7gjuAL4LDrO0zBKQi4BPgT3dfJ47itgiQGYAmqJVTY5DXttLz445wUzV3vvRs4j6Khgz/eZCPfpQiA/fAhuFBA+HOgpQFnXNCzalAoAmAPqwT4tAp+N6jkJMKnspXUYs03gQnEtmtxd8jQMBGStcW9Yk0BeAPYgLnwmwH07yFIM8KqAzwWKTU3hVdNGcXwAmokqotEzAL8KE4SuYE9gXHu4sLIoBwjHU5MJg8h5wChha2MqYVHrKZJJFSo70pwqIOczMKESexs8CEInZn2xnWtruGtweVuN5I6cj6CLSRnXpMq2BTcqK1mlSD+BTzKhYLuoxfieRjAiDrv8sM4meYJhbrSdpR0o0aAqgBLgvoIJ7Bp7jglq4oceYA+OJERSoIdxwigVjCpUAMDnUskaAJhqr0oF81xONDlX6MRx+uMQH/QAEL9KN1rHADAsmFuxxIKHDozgJ4rkPfGNMgFhZwCsO2yzPoQO9qnnu6YRXvPnjbeP1VMAWCF1q0I0u68d3MfrXrD3p3Yx61YMQcXsZQD4/OgZnVissVI0vV3OVk0kw5qE8trRSQ8Y4GupCj6Y6QGyzIrlQ1bFUjNTJHTrCIQPFjGUVnVV+tHknsC4gvUA4lH7VMA8Bw8fEdOxQZi8AkA/GV+vsTarFcnQ1gNuQ1ekuc8hN98pPjCKx3sbPUiLZtpK+kmTs5o8TiVMVOVYbtHkAULWVFlhlSL9+GH3r2FDdx3A9DOBLzCjU2Exvkx4gq8OapVUS9qBHJEUKK5NpAtW4M0ev2CxhEgK9BCGLi7fzkE/ZVWrBZNv2j7YPXYh0b339gO/aXZ64jjNgI4ONG/LCLrX1c6p088KiuxBAO+BDxYShmQHgX+J5LH7MOEeEuNrXRMA0HFAwNug/ge7gPcA7l4B0NkvsNJwRHAZwBvvG0z8eimQYpPEzTOSl2Ff0MQLFTkf/GI5uLpRshqMhxpMEZd4qAt4JwpdB4GmdQ33hMbNWP0GbUn8k6+aGERvFDyasG/VMHSJQS5p+fd+071RF1PJbP2ppSZPgjbZUi1U23y9pkqa35ZP6MfejEgEtoXlw7mxSDcGWB26ohPGXhwaVamKxw4U2C4aGKyWPX9UhNHTNPQUSQkjgR7G0IVjCLsYo+X/CLxo4cyBQyqyRnv/XC3/ZLPT08dhVgFQH7kdoVcqAOZwSE4sx7cW/UtJVwv2eWeO9tZUcM2mdTvgXco8t/7lOtwJvQ9ApqUKZAcWL+Wfbpyi2MsAuCq6Bl5vlGLLJOSdVYDRUE9DXgH8McQD0OoXuLzSTMXxbd+uuVpHHGEuITV6TWD3hMb1xhHNwX3In/p3ytsmdGP48DnElY/If/Sb7g1DV3nI8bYculbH1WbjX/XvhpzFtVXSfEaa0M+q5RL/9MF9S9XmwVQ3nHDj7MDgq+mKRw4U2m7t6h1tmbB3UNeXcB2BHvzQZTSAnFV89Qs87KRm9UtuJXSQ4q3TeOLJ5g1Rv+TaG0eiDniTAPCx5sOrTQ3GtD8JANy9UmEH4kWWyURZwT5g4eTLrgquuauJK+dbTuhd64cu1jdzGR0Auj0AJe8kPvb0qvB88mVXAYC5oJ+PFRdKsblGFi8XeZNR/Rftv/PcTtsmE2fVOA/HQylTxPS9fxr5RwO1Gqz1k9BoT2BcP3EVo+Sv/GYGEcCWnxOfW9V9GFbY8+ZHPpNbaro3vFcz3eWPYqvX06v+06sHiaxc7Vo8u45KmpHUTzRcjuLi4uhwVzWUNTnUanLCzNkB+W6q4pEDBbYLqa+SPds6orPlVSO61RKGAj2Ey/SFlZGRkZPBQmCNO5EnLySeF5HzaxcQHSd1uHhYT1+4zhDeLmvwINEkdLEgolbiFfR0INDhPsCrFsIjDBeJZ3zslCK3cW/ORLB24U4BkAtJ6SQSlwkiaLzQGkvRLF5ouaqxCTTWSAX/+szxdcMqts95U1PEkgfPf2rWbFLRqOXIuOkWXi3K2mqD2DAra1vxtYQAkeZ8XTCo56RaRzzmvj1yQq0/eGoWulhj76DZqrNZy4WRk3pdWTcoODiWx6N6/7xjo7JoCF1i9QkzWjZeyjpQaLs4xjd6TnIgytOhq0HCQKCHMHTJJSC6xejz8FZ5oCUO5DfjU1kBLhRUW1gMUIgG9T6XCG6XpQ2muoBMwy+JE2GPAdDFWAKgG3D85DqMRnRNSLC+aAEPLGh5uY0nFkwA2ZZqw36jJNDR0MebzS4lK/ujoIQOw5oE8to1rjB0FX3AZ27SI1mfv5YpUpWpZwC4TWJlYxS1xl1KOUWNS9XEIKoIpPZs3MyNklilOeX4qHOpk3tXe2D0U0Ulgb5m9nJWGc3WOpiVIiusgrLpqzAm683svqayRKN+AB9uODVJbWYSSKyjGe0Gv15JnpCouF7lQEmBuAQrN3oOA4BS+HxfvILIV3nfg/dL7E2ELmYdNGhEPrMzI2T6wKTCEpf8CADtINN04h0cBzaADqMwn5Fg++CjG/kBGCYZAHi6H7y81KR95wV6/HQLWcq4QIcwzEUHAJUR/Ei8dumyPg/BPj4A5taja+pudIQeyQ4gL6oAB+9rbMiFhngj0SGAosdyAroNvJyQLvNCdBkAnRJtsQ+HEq4RujKS2TwnoQ7DmgTyAkAtK8F6AKO6lgBfdQNt4eV95daamSJZE7RLwGNH0NlnVA9WVTYIJMEea1yPH0GsKD/TqpoYZDor0bGJ54V4ov1VOwV4f1JzgCoa+Fyo1N5G8Q4PDwU7dAntspm9hlY5n621bfYuGwCwmHEb5+/+8rHhweAHl9pJDAubK8voAmvQD2CqHYIdCMdAR8Onh3zVA9aXWOmvtgs+KBInZLvApY4rXmAicqDQdnF87EKnSHpOnQs2AADZPvSnbyiEEhomWSBQdkg8YKFrE2tdwb1x9XQwwg/HQPvzvLUKjx1b/qgGas8MtqlPNDvfHgev/XjulsbK8rEa3z3hV8rDzi34+lj9BABTLQ/z1lOrT54sD18vXGtYpVeHO1nh+7461N6lgFk+fPOjUbioHN4rryDc1zm8LN8Mr4nKC8PO5DOBK2SHHU9hLjOYaWmMKc8cfyvRJv1q764uBchdx+F/Z748mInGOlPlIb6QDy8DgB+rOfHtx1CKNdbi+bHr9qxQh2FNQnkBLNwcUs5jk4AqXwEw/qfDuhjddlKyqSkSjeX5Q0UFeGrv8dz7GHOHtfNbzbLehXuscf2qO8QKb0VtJv20a2CQRXZMFW5sONuezxz/zQ+ipS4+rJyPjiU1hxpbgs98ndrbwPifOk6x9RwATLvDztSnV9vrunhvVeiytQbg6cNPfFIBnj68d0/l46lYVHnBcUqtFQColY+/f75RJc18I6UfAFPu0HLhSth3Dof/Tx8ZvimTh5WPL7cmT2CHOrhfjSquykPXfhU5UGC7yJcq5WOF8c8lPMfjgzd3A8BUb3tDEwkt6Otj9RNWIMYesMi1qVGXHfTMB4dWfQXAwJ8uyhkN/xxrOQNg9npp5Vw83DV2AGyi4+Cfy/lnAXNmubhnEphQjn8WZpy1wAC4qpzWq8oOmpPZ21bGeGFhPLnUCuD0clHOanOZd13VGt44bzkTTVBmrxcWKgj2+eNOYUZF15xRzmy4+uSPOeYS4E2Xdo8FsgI2B98Ua9Xh/wAwebN7VoUSLFRLj50zwbeVacc7axBcBsDscrHlbFSLUMLmeGNM2rNCHYY1CeUF4I/dLLZOAqhBAbgx7RQXq6GSa8Gop8EUsfomlrtvVAGc1rnFc7g17hR92xi1Tmk52mONe1EVC7MqqL+v/ORqd2gQ7ywrNFvfTioSgKdyOlLkyllebLGxKdAc4EPBKD+x18SlWA9C37WRE1WbF3lBOZE7GJO0l1p1eVtrrWHG2J5LGjBnmbyYOBEAK18YOaFstoUFnVNRwUllNfpGoJ/o2IVxpzBbCcoccwoL1slmph05U01qmbVe1YkTvOvFPRNxxSd1QQcOlLCdCa48oQthYdZzvPG2PWcA6IWfFBbHE0rTkQXNOGsJBDIGDyEi9WEfVReI0sgVBRgAVgRE8mCWOu6Jb6Bo9zH7GXwUg3OYTbHNElcK0pwUUpN2ljjXKdivgQR2NysisS+QMbxmISz8I9+EI8M9gbAiLN7uSVUgsWgQnxB9iTPq2Qs1nMlEtKCWCxSYi6qWS9Yg/AzkBQAn+DcowF7Jbtr1reiglCniZRAeSM6LycMaqpvYw0ITFeKdn/9qeuEkLMQpRmckimosmRWSRnWieoWaQy6+rN0bWyM4L/MXAtj17fC8lHOF9sr8hVh9eV5MWDT5GReQ/TaAx78Z9ONpWZsqy5EN+hGNTprIrlgsJJSKII188oQgr7wT+XPC8oHtRMKW0YTWnh56Yux0Im4hkUekXIbYOk98474S5yPfJJ1uDva18oe9qPFtxKFrDR7Zpj0f+eZtnEzc5xNG4n5FLt0LDxHqw64nE01vPyfpmtte0YYLoMuQL+08sqSCHQuT79yL0OV/2KHLZ4PTvH/9uu65sM2y+eC0032FnIlC107hQehn7814qOh/2PX0zw4cxML5dY+5qrdXtjfedxDRvRuCWE2be381+64HYNR1T1681yY//Jo6PXevqsWeArVOgiAIgiAIglhj0rHOvo98nfRDEMSHgvfq+vvp4QiCIHYgFLoIgtiBrP9wxMjDqpbMSw0vyLo/+pmX365++Jo5pT90RTzy1Q9JiP1Pv071vTfs/f3bCl31hzei1+/DunOY+n2gmQ8/dH3kQxPiwzHAw1ZfoEnGZJowEgTxAEChiyAICl0EQRAUugiCICh0EQRBoYsgCIJCF0EQBIUugiAodBEEQVDoIgiCoNBFEASFLoIgCApdBEEQFLoIgiAodBEEQaGLIAiCQhdBEASFrntIvo90QNxvZPopdBERxWJRrN7a1gV2UGzi9OwAqZC4cw7IB9cLXV0PWuWzZP/t624I8EdVsw6hZ6G68fmlEumQuJ1xxxBgLkSelilJRRNGYjOY+Vrzns5s/FIAAHUT/99OIzBiqzDM19ihaNxlTOx1mcEHf+RADnAboWsUnQPNejp/ZDMvUrlxMtGFSFInsWUu6vYX3UrwZeVEHLpYgUZdxLrMarfpgGpTkU+T/ojb5Ibq2prXUegibPhRBQDFbgEAvBT2dUWAC9ivrFTM5QAADlhJxgcyCTjgPQJgguUEaZPYOooBrNQtABQAR7CSBJhALgfACRyzJ3Ku4FAHbV0A7JfQf2nC+NDROQDvpAYfknh7HAD40Juq8w/+jyPmQhU41AXg+xrAoWuPu+aUQv6owLVz2HX0uxi62Cv7Tur9ZRy/9T3SJLFlNAMfklh4DZmhv0Tnx/+xz5xSu17Cce9V5F+Cf0KDD8m+vw6GKvZQDF081DKN/X1YeA3sUBf8EztuCkCjrjugw/7Fy5ky0LFnorov3pw5Mr98ANjVNTFt5u2m/DPzrB/oaZlUHQADIPK7axkXdW3qNdIksXWEQfueyem8ax0qW55n/TB1U9dgRxcv3zwCdMj5D8r26PBQkW/VyPbNjuXLyHTNTfEyjboeMpisId9yQn/wh2B9M5e561Yj1c6c3fsCULx1mXddtX1a91vVJ74IeCd05lvB8n751PVjXZVry388QsoktoE08C9UWUFaxyu9Pf7s7+PW638+AmRbTuh3vwJ03RjNv7QCAPGh5Qs/QWnlDPhnK3h7HDc/u/NGDGT726JNKBQWNOY4eOs4/KqMdvljmGMCch6+CrauVDHHBS5peDpYXVhQpkpGILZLRk7jRhWmGiyzeuOYCv2p6GkscsHlFMLnDKNDvaqBOwdMcayMA9OMRl0P1ZCrh/V4VcgaYLj4pa8B/WS019MwEBpA9JTXHODDdo9h6JoCaKZIbJMSOnwblEIXWwH8MAzJJcCH+CejgPGno5M0AEwCTL4FGHuwodD1cIWugzAXALnQAwZw9AAidYABoB8FxHupTYCTOo6RKoltcRB4WwPIpe9P2w4T4oMeAOLnqadwgkM1wNDWA84AsEJm59WdQtdtYKawWAWDXZsXmYMAftN4kP4cMvJKInQJ8INdwHukP+J2mdOL0wB6+5s4HoCOjlWbUofus31pxwEBj0LXQxW6LtnPWQUY/TF/DMBi40Hz2eGbnkp7lDuh95H6iNvmkgaA9v65Wv7JJrtnFYCU76UPndCAAT/gXco8R6HrIQxgWJwEAG0mm+3uni/mLiQnh5r1zVxGBymOuEMUb53GE81ClzcJALvWPHROAUC+5YTeRaHrYUTZdQa/+ZKVvHIpudbAoZl4i5RG3DnkfPPt+uO2b+VCh4uwyUNNsDHn6Z1Ya7ovfwdCl70v7XPRbK+Q9cgzSgA3CgBY47G0VE/cDs38SbVZx4QE61t9qAmf2QGQo9D1MFLLSrAe+LofaHNXBbZnh4cOhGPcPtbvAZDoaHC1OheCNElsC+0gU25YxWCSYT4jwfbB193okE0OrXYK8H6Ag+/AzL40Ybx9Fm4OKeexSTN+ZIgVrjTuHXMdp7j7DQDAfG974Rr8au+uLpU+yufHPkq/YSS213c+M9imPpH2J32sfmJl+ViN755A5YVhZ/KZ1YfO8WHlfHRsLjOYaaFR18OIP3az2DoJzNaKhcVq497+90dOjOft/ys/KSyOA5M3u2dV8NQXAGMAT+XufMqSzWTJzxwlA+50Zq+XVs4hcihjADPOWmDOLBf3TAIzypmtpA/VALBylhdbxuBNl3aPPVga+cgrD60zZF+RW1lfcGyKcFYswC5j5YLVrBz4KxLIWE1+6QVmD3QK9rfXwRKXAMCKm+hnXik3bmLFYrFYLDZPLPfENzYuc9e3t6oZcT945ockxOe/cV/VV0ReAIi0P6Fovco6ZiE6NjiUJR2XFQXE/VRfAHjiWzRhvItrDOE/S0t2gWEesBkE60EiwToYFMJn6MMjsLQU9ZBhOWZ+ezF2CADgvUrWeIgd0MwDOu1PmE/4m6lFxwaHmqTjBttowkgkMZBAm3/Xiq/XtanXV002Kdc98aDPjEgFdzl06aEx1rN4t4pfGcGu/9wkXw7luicecGjUdbdD11l+8IB3jhRBEDTq2lHcOFEw8xoAMHE3f+PKin4NAJzCogITLGc0gJLdCCZrjpwC75rTAJzCUg1w/JvuVLAzOh1AqXVeA46/XGz47SVBUOh6iAjW8AHM3c3IddjFtXNAdkiYUyrMdb/3SPCKUT74vSGx5/yQ7HsV6ByAuVBF5ycekUs/B7D/919lxySuVAAA+/vgndRhlnOyH0ETRuIu0v7MxHSnAI4+dvn6QJjrntv0+ACAQV91tO+uZVywgcXJ5QMAslIDAC/PIbtnYloCAHb1LU5ljiDMck4QNOoi7iJ9M5dZoXw+s+eU+smfiyDXfcamxwcA/smT//TNIxd/csytYvYM/vtXABSvnW8F0OGPo3jrMtsDAOheOYmbzyHMck4QNOoi7qId5TRMtYA2X2El+m13kB4fADChPXy0apQDcwZY5ALwx40GeN+shqgj+F24OwlMcpnOck4QFLqIu2JHowDFkPOSGQFsenz7fwVGzUWJ50ywG0C+ZRzQbcEpTCjAg0hnOScImjASdwNmeoAsg+A96R0m+UXD/gSE2ddx2wdl+2Y1MJ/tHwv6Mg2Y8KUfYZZzgqDQRdydYddBACsQ2YPYMO99ZlAmvonfAFhRveKsjXUawBo/aSMICl3EncUfs8OolQo2egqDDe6eiG48ApUvAjCjQ53lSnp8RhAUuoi7jk2Mr/3JjQdo8k3F49A1y90q4I8OdVcAgAlNgy5iB8wzSAUPBOErQPUmFta5r5JfPdUPAP45DgQJy3n6NTIEQaGLuEvzReYCAOqhQTcIYcnd41lhgxYA+HABbmjCSFDoIu5F6FLdQFsZSxkX6BDr5ro3TCL5GsgFVsYBgaIBAFPtEOzAznxHDPEwQWtd2+B+fGJg/E+HdfEdrKjDnazw/XVz3fvLh64/k/xe+X3kh3TxHQDAlDu0vDrDPkHQqGvnh/uXm2/Pb/G1KnzwDgp1Y9opLlaB08tFOatX57o38ac/5rgXNeLE+JMZMeUUFyt2DDbuFGYriLKcf1jk7/5LajLxvQrsc3ekL94HrxbIbOGHruyg2Niwd8DylJu+OY+vkbD7iS0m8s68AnSWtxLrXml6dDAzdGyC+jAdeTFKjx8dFCUudyRYuN4VHOPEufFtQSxRdEozYh3PLBYLYv06HJKbrOx62lyVq50Vi0WxdSsn8vJ//oXNnHGf5abf5KsFNq3zNerryMiveAHZhgS8j3xzKx4sN24m9gDKTX8XKMzfubK6f7ty22UE09dUgnyY+Sg9fnRQlLjc5sY30Y56MjVPmLA8UfRmeXwI8C9Nr+e73f9K3Y2uZggwE2Pkm3dF552f/R7Ah96qAtmjr5ZKNGHcqbh3sPHp+oOkGWHm6/xQYurV2TiWMXqNaJi5rdkzM/N11lsm32zGmjpPR4KhtZ2UAeCQAAo+6ub+qBaNurYe7cUdDF0XWh8o3Xij6DhwoNow90w2o1NrxGpWuK0L+6NoG9xXIe9sFrpObaZ/ZMU17z7VOACGAgBhcOMkjbp2aujy9R10K/2gqWd2PCPX2X33RpmLZzMueefd0LnPANhXh8rafeOyFLq2rjIDgJW6ARQRvHuHleygoRSNHZyegv0I2zHvkcHBRYAL3i3sEA6sAKfbnt2dy4mdH7uYDCueA8sJAKVuAbACShJOrBwmWY8IjmQCuRwAx6ol+NgaC9oFwAMbBJ+F0Cas1C0AOIJZk7BSHGED/a957v3ph9bD0Nad+OoIgAVOFjolHIBJlApBRZkMK1y0amalQmSApv0rk4CocQBCgUnAAe+x/tsTGKrULQL3Lli3DgbSjgj+KSZNaptP/F/cBjZveZowbpmcAdDbB/d1PvR9jc4vfA98SJq3K4lU8GH698yQCPK782FhLlQ7/+1fgQ+9qTr/4MflK/8ngN5P/YAPfm9IuK8DnQNAmCB+B+OrrkqYIt/F8fd+gP19WHgNfPAHg7e+N3SlivxLMFfHsWvg/3vwk+d7+2FOqV0v4bj3KvIvwT+hkRkSC//X1kew1c8CfEhi5mz0yY+9Oui9CoAd6oJ3UuPQtcdd/4QGOyb98D1N7S9C/hJrn3t/Rq4haS5UgfxL2HMOyLws8PY4ht6pYO8ffhd86M3r0esGhq5Udw38twFz4X8MCfd1ZI/+15fgvg52qAv+CQ122MXb/+fLOGb+cg2LQgBybg/Ahcauo9/F0MVe2XdSgw/Jvr8GwhcatB/9LjLH3lTo/dQPAH7s+xq9H/8BskdfRecAvJPReK23D+7rgE1i4r6OqA1swfI06toywgC8b3Ys0Ud17Jm73pNMBR+mf+9ouXTd9iK9rVPLiQdgMuXAjALAQV/lJXj/3BiWHoAJpGJhinyjTV0j2zc7li8D6Df2pfFH/cnlfQBjfXVk+ucmWR9M3dQ12NHFyzePAPta5/PbmPxpDvTumVQdMv7EPruqvMudm8q4APLPzPMy0C7nbwbPD7EBv9YhsPa59yX53RNVCeDAkuoAMNg6qXoFlAQEB7hR8esGBMBY//xy/z5f5QXADiypvESma26Kl4G8O1fbhzrqa00sjZJgYokLMKPAAIj87lrGBTrk/AdlRC80WOIAhwCkAuBDAtIBHAPev3g5U46iTtR89u2erObjNrAVy1Po2lboynhnriberdh14/TodSAzMzqeCWZNJy+d44B7dXJUAwBzZy6eTISl7I2RamSD3x09zSQyLZeuqhvVna8gzZDZc3ribFZcqN4aOY/SypmrlW6A7x55AwCyraOXXuMukF0ZqZiZ05fOt+HW62bkJLItpyfO5cHcmdHX5dYvXGPg5ZlLo8t90Sfw/KnXAMBcPH1RdQHovjBaLQFdK6MnnaAbERdP/Y3E2ufelxRmLl+sAvx3TpzhLjLywqXTrAzlAJIJMB/FW5cvxkP47K3R8exzo6eZBDK/c2KUucDbpy9WikD3rdOj1ZXXcOrEmp1RAay1Bgkevsa9PDqquoCuG6NnCgC6V05erORhmACDBJMqiHhcMED4yLecnjgfPVURN5+l0UsXmRu1ga1YnkLX1kNXHWAGiJ/u4nIc/plkKvgw/bvQsGkaMq3j8M4mOrKxxGrnhPaUhONpVAt4IEJXKkW+Oweb5X7WPgpS9BRWlAtgEto/C8yHPwYvehqLXHAxjQW19QsbIGPG4Vec6BNYUHUAuDUNVBmAlSqmGZichHc+WAryqriKdc69P91Q27cJXIanBdq8KrxqETUOiGUBx49fNxA4HObZrPa0BHAZfrWAlXFgmoG5UzCXNrIo840ScMLQtaBMlYPLKSxU4xca+EygoAvhz/dVAcxwQNZQWNCYi4JN3HxmFHwlojawFctT6NoWhicXE7lRgEmngrfPdOowFHFfpx7wTL2ctQIoEfrIg7EemEiRz6QKkvIEz6rKJQDKAbxqrKlojw/BfQWzveGn42mgxqNPYCrRAgFgDqgzm4I/6HxyPuCrDc69D4dddh2qCqMFcj6AKoPPwFu0hFyKXzcQOpzBtH03gV8FlG35htkMR+vPjGsc3EBJiPC4KaAWOP54/EIDX0uIGgOzoUszcJ8LSAVZSzaadPPRcRvYiuUpdG2nu4PPjyaUz0yy4w+2OQUAqrMraB5m1QAhjf1hjjAPgoYMBO/p6Ylq1tbT08UQvdJD1IMgbSvLitEMQXzQ09MNwbaphVCBhkWfif6CFQWLlM+gIyMIYw9b59z7kOn2rkTblwZAncFwwYySkArz2f5mPicS8YIVJcA2Ts7mczgG2oGspdQdhKj4hQZKQs7zcF5Z48h5WnKhITM9PV2J8uLmk/gFGsNWLE93GLeFN943eFo3CV3RbN6mf59wD4nxTZa5yMWyW3sQYjsaUuTvS8dq8V4yKuSPCoTTEHR0NIvzmw5dwZjXIPqM6ewXWGns61P9yNrn3o8sLB/OJX75VLdt34f8e199FuKdxOsG1pgDouOAgLeZqhomxBJqDOKdNQ7QVoHqKSYvcRT8MOLlar8SzFcslWUp0XzYoS4gLjRHoeuuLuV8HMBV0TXw+jqNKEj/7p052lvb5Nzd50PLe049CKHLNKTIn9Brtw92tGXiZjn8NqsAqI9s88JyHcfPDixeyj+dHKAlQ+0OxD99cN9StXFcZZT8X/UlxoVa9bqB1TOuA96lzHObuhbkxxR8zuS62ZA06oz5S0wK2wP7TEr1kU/90ASWNTqc5EXNp71rrtaxPQVQ6NrmnOiCfh5h5i6zaoUqSv++cPLl1E8e11nLyntcXlUPgG7kUkOK/DnV4OGIp8bZ1hGdjUKXNwkAj2xzxU+aYB4Uf0aUvJP42NPJOa1ODkDiOVSTc+9PFkeHu+LQlQucSz0p6oZzXydeN7BG/5JvOaF3PbdhQl0ARgnxDgx3zBr+Gb7QoM65b5SQdhzlayHe+dVnHR8Gi5NNmg/Q9/5p5B+Nx45bsDytdW191GWfSJmMlxqbhK4o/btnbxrWN1Z0bnZkZPwB0E9GVlO3G0zjsEbnEqHLSSZkDV7/aLYXurictxdmJvqMw9pC43Jc1GY1s0OvNc+9X8dd553U/A/SB1ROKp/nfCB63cBaYT5nVW8gN7yUKkgF394MSGqRC0AkXmjg85yBKkgdnrZHLXGxFN2GirHNh8nphpkpha67GbpYuEBiIMHKgGEu4KzqiZJrKj6XQAGaA/m1+kHHrhg4cmfrp8dXUaTWLPky7rAdOABLruqx8I9qC1qE1esW6WitYikjgC4/+kxrOBm6JBAsY2sOZCXWP/c+R2UAyDqwxKX2ddEPfbRpo3eD27xALnhytG11D5Py+TajYVTRXzWRBOtLvNDAR3EJuhgOzlTRwGdSAWrV7/Cbm2ULlqfQta3Q1SnRZrRRRXQKwNc9yAykew+JfQAfAHPr1sjdyA9giUt+ZI1ya+3Hhw6I4PdAO5Wc09u3qKMU+eAQqHYK8MTdrvmMi7wI5jp1LtiA1RjDfEaC7YOPLnSKrV5435EFBZ/1IePOR5+JxgeeaBOm2oPwse3FjMsOAGufe19y1I6ywtkj70PenQZ83qqg3Bri1w2s5qDIyyrAwfsAU90n+AEAcu0hT437AHShoTxfd6NDJl5oYJSrUMuEctUzPgyXCqhlJVj0YnbbfAAYXUC7TMbCzVue1rq2TJ0B/FjNWQGqLxzPjX8OZvzIMP9oyqZB+vfO4WX5JgD4lfKwcwseO7b8Ud283AXkcsWuV3f0OxCzxwH/LKIU+fOZ47/5wRwfVs5H47th3vLhzuJK0DF7fPDmbgC+PlY/sbJ8rMZ3T/iV8rHC+Oe2NEs9bt+i61XdIaflzegzjpflwUxyzDDtDvNZe1/e04f37ql8fO1z70vyw8lXCHiqN+94VcAwD9DiChKvG1g10/yHIe4rzGUGMy0Apt0h3gJf7d87subclBkAym0sr/LCsDP5TOKFBqqs4bModDEDX7dqYOHmkHIeC9e7guYDoPr8oWJi/WwrlqdR19aXGBgwu1xsOQvM6NxMFcCsclrOpI6x6d/9cacwYy0zoRz/LPxzLHFg8DyRAaA18isjI2cz8m7khUmnD98w32+YNJ8d3GIYrQP12VGNZIp8jZWzvNgyFudd9c8sF/0zwXdvvG3PGQBmnLXAnFku7pkEJnRhprqVCxugPjeqAFxSxdZLOv6MOoZq6bFzBgjT7t1Qzsr54OSzTF7EOuduk461Zz75of7bLX3MKSxUorcHmEvLxZZLAHxlgJrRQPy6AZ3IiKs1gMtO64SGN13aPQbgxrTTehGYYmvnjvOxBKAOFXstjAFmlDNbQeKFBjWjgoPtafOA8jXgj90stkYr9UHzATCx3H2jiqgNbN3yzaHc9M0b9isiygDPC0HW96JIp4IP0r8ncu0FZxQFWHBouFIs7ErP518A2NdcfK289pXLG4abpqTTh7/ibhTpXgmvJ5toZnPhLEyRb1P/FEW4psEAgBdFvMhRFAg1s0qva3qm2OjC4WcuuV0kUu6zYnyJwFhrnhuRzk0fpWu3RfQ26PXZr68p49de+dbWWmKT+lrXEtFaYSh7pOroYc/QyYK7kLu+Hb5+IFB9cKoj1q4vQ6yQlBZDQ4WOLlJ6y4UChT6Qs99Z9CYBXki2gYTlKTf9ncdoWQ0zwPu1oAeaRzoVfJD+PZHyPThj3i7w60RvqG2vI6/YX+G1bqvXyZSkuq+0FKTIryHSgo6WZ/15RBswH/wznzgt1Ov2Lxx+1pPbdSLlvpmPLxEYa81z14pAQ29Vgc4vfgdA76Oq+182a7lM68Xb13Cs1LTsOtRs5Huhk9XDY8PXD8wnvBdYWn9YGykkpcVaw8V0Sm91RNk07QF1+91EK4l+LdkGtmJ5mjBuvW0oeVfK1QUgIzRf2VabNYmfc2cGyUr3ZOlASwCCCQBSQ2/acNyfmiL13e7MiFSwZap/fFeKVc8fWC556tntOfXKibjhsAIZ6Z6gPg5AMqHBhcKFTT8KxgzpjkLXh8Di3RmqTrj74F/EzNL2Tq+TYe45+imACUgFZvQWLCBIdbcPTRi3McC5O69E8U6cvniiis1GLi6CzOlB0vUC4AhWkok87zYJe5Q+HFHm72BTIsF+UEJQVipFeLGbWtpaoy4GsNaqsM9jOgKsEKaKT6e7T6aMtwYSNlV/kKafC/tlq/nZt+dol2nU9dB2tnenWH9uK0d3/t6PjuDtcfAhiYXXkBn6S3R+/B/74jzvNj18nD4cQP4leEFG8ZMaONQF93U29FYVvZ/9HjJDfxmUZVOEBwlC00nFiSRLHGBGPwkwHzj0L+ezA//VporPDsP9e0Rp7hMp4xWwv4zj7/y3wR8M3vpe/qjA2+Po/IP/44g59X8PY885hAn67xr+g7HORqOuHUvmyLzuAdr3TE7nXXtrOVueZ/1BnvcgPXyUPhwAO+BPfVAGOmxi8V1dE1Pt0ihpH6TmflQWDvq14FnzdFJxIoXhAtxXHAifM2dBqvijZr5dIkpzn0gZD6CuTV3bVP1H/UnVAyDzwjzrO7qkOhAl6CcodD2g5GdGz2Yk/AuXLmppN5WujlacIM97kB4+Sh8OICNHL54BIG1i8e4bly+qPuicbVaOH5f1u6Oj2maGSycVJ9LjFwg4ZokBIpjnZ8JU8W+PnpNxuvtEyngA1857IxWbqn9h9NLZjACyM6OV9k+OnuEyTNBPUOh6YNvNGBYgcaMKE6a098ZtEngAYXr4MH04ALR5CguVMLE4c+eBSg4qB97KBWQ9LmtCm6r9PXk6qTiRNoGWEEs+D16BA4Sp4rmpYEbH6e4TKeNjZnUdZzQ8CMAfxxSb0B5EmKCftEuh60HF09EjZuHCyErwpmIgSg8fpg8Hgp/r6jCxOBNVYImjxsF8LWze5USmcFtQOqk4kUYJCGWYZGHoClLFOz5gqnG6+0TK+Jgg34sJjWkqMFqECfpJuRtAy/Q7Gv1xAA2vzA5y6AneAyBKHw4AItHnMzCUAA74HNxXUskrqbKWgtC10ANGil5L/09CvuNr8W7ybfb60UjTYZp7w8UHRn0KqV882P+d4HdQ0RPq4oMe0PMTFLoe9KbDgN5+4Derd6XTwyM1OguC10EAHnwmf2X0p7hQTcpqSCpONIy6PsfkFWj5/zO66X4RxKREyviGWc/g6seHOzpIsRS6HgLa++dq+Seb7Eilh2/KJQA+jBYfqanPMl83LcsmFSc9N2WJMaOgPsM2SEyYSBmf5pCcCF6Yntb4xm/podBFKtjJiDqKt07jiSahK0gPv2vVICBqTpjTAGC0/JVaYtxHqqwwF2w6qTiRDkmc+4D6rLNG6NJ2zp5KGZ8adLlXKqwxdHmk8c1Ay/Q7Gqkh59eeSyJOHw7YTMbJ0BVsVgWhfe4YxGUJhHnIVyUVJxL4rGCAOpcNPwLSiQTXNnQlUsYnm5+prDYcaZxC1wM+YBYIJyCNzs4QvcgjTB8O2DceMzcKXeHTYIrvUYaHN78EAJSjN4Mp+in3eqFLF33AXzURXOIAL8fp7tMp41cbK4lqI71S6HqgYQPoMAraQePT7oZJFqaHD9OHA8Bipg8dcehSPQDvA5YyPnztKsRldYowOXs6qTjRGGdkHcGbI9KjsT70Ik53n0gZ32ApF42L8kGCfmKjrptUsFPx2o/nbmnUnhlsU59oGAscq58I08OH6cMBeLq3vRDf4xpzh7XzW+M28bhyFaKy/HePL7fa5OzppOLEqtndFcCgVTfYRvW2y8rvxWnu45TxaUupQ+1dDWEvSNBPuqVR1z2is7yNkxpzAm9psnIu558FZq+XVs4hTnRvgjzvYXr4MH04AHNmubBwPkosfmvcKfrjgK+XAGVUoqzTy2LGtsZ0UvHtsk9uUWOsd1Bu8Rptm8z3zg5tuei1qUEBRvkagNZRqnhcXpZXlY7T3Mcp4xFoHwaAucy7rjakKAwT9G+LzIG7PdYZ2NgMW36nwR2HctNvhSAT+dYi2JdeQObo5vuZZG76J76BIP27TbIuEhnDU3neWSLPu1NA8rAgoThDmEE8KsvuyQFhUvFYM2u4Zb5vHcm/5q6tsabseuWVF9bzzCZCfOmVb29Ki4+/EqixoRVuxhDp3PSRigLlIZFj3SmslTI+1L5AfFy0N3XO+vVt7rjfvqPtYHV9P/LNjVz6lW/zrbee1VBu+ntN929XtthPbv+3zUH6d5thXCcyhqfyvJtaKgl78rAgobhB+OhWVFaUSTyZXn9dSt138tXdRe9syxbnD/Lq9U0d2LVQ8azEpTtgiDpi5SGRY31VunuNxCPB8Rd7XLQ3dc6OnMZt1gy3P7ygYHOHFz9+a4snmAck2a9O1KP3X1VuszS5sNXmy83YJot+J3hOtx5JzI+OPjiG+JBDlxm7R6tQFLruMBdat3jCXcq5es+5lvjNkfPo7Za26rHzDdnogfa46HDscyPSPCsK/cAY4sNl02a4/SBJyr6zbPk3M6b+kFb8w2pba0lsKLv/zhrfkQq2iNMtAEfw7jiTe5RL3BHgAqwQJxkvJLI6BIeVuoEocbxdRXcEC/PIF3K5zcmxMPZhKwLcCh2kV5dAAbxbADmwnAhT3wOlrkblhS8tYja7vtMjgxeYFgAueLfIgeXiAqziwm9rRCSBnICDtq5QIFZAWwFWooS+c2A5az8mAxFsMn8rkjVpkOV/p414ApmD+kuUCmG++9hVC0nTbItSssQCSoVEedYMSGrcfnLBu0VocIQ5+4ME/dsShyaMWyQzJPxR1fmJR+TSu0Emd7sJAB9673zvp37AB783JNzXgfYXAbz3g1DVQ8I/obH3CLpeQ28/zCkFANnB7+DQtcdd/4QGH5KA9+qmBFlc/JA1MXTqsPRPaHQOwFyoovML3+PHvjMsFl7DMYG9Vyo29T3QORDdWWt/Ed6rOPReBZnB7wDo7TMXqsgMCXNKdX7qB8gc+3+j8w9+XH57P+T+738wKOG+HmXUD4pbg8e/iuO/+auhi4dapoN879mB//oSrvztkFx4DWGifiAzjMP+dzo/8Yhc+sej3wUfkubt/+tlHDN/mR38DrC/D95JHWb532GR65j0zwEI6r9r4L8NmAv/Y0i4r4Md6oJ3UiM7JIDf/NUGulyf9kHsORe+EoEfe3XQezUu7/Gv4vhv/j8AQreI3OMPfly+UjnUhWvnEMh47RyyA389uF1xaNS1RQ62Tt0sA1mpo0zuwSYg31qxt7gP+iovwQYWx7AU+f+Bxy6/K5A5sljLy0z/3CRLPE2Qf2ael4GO3RPK1HaIJkTn7ineBzawOBklP+htqeVd1HW9Hqa+B+/3a+Gz4f2Ll28m+83y/PIBoKPl0nWR6BzKGjVdX8K+3ZPVvAgz6gfFrYVfN3UFkW/VUb53dmBe7+vYXcvLKFE/YOqo1639GICOPXPXe2C3AdjVtziVOYIwy//Ool3O3+xDlO+esf755f59vsoL7HLnpjIucPSxy3pJbaTL9SPGwJzqEPFrDPYZJMqzZgDA+/2p5YPxJzJljWzXxJQddvW0TKoOAGygjm2KQ6OuLQ663CuVti8DxWvnW3ddqLKCrIabwPpu2NUTk/t6hwABAABJREFU/rsnP/iWVLx1RPMnK2GfKC9Vp1rQ5p/W+2FmzuIf/igutvut6uefBrpmL+/6+pmdoovyKcWfBmbP4L9/Jdj03Oi7w2711OcfPQ9037gMp++NTOup64ddqxf5fZ28cZh5e/zxrwrtXp2c3p1wyRtvtOovvVfB0qjirlvJyDdV/nNhcWsJc+ON/zwCoHzhJ1i4rHd9Q2hkfue77X/6wql3h12V2XNK/eTPhQb8E3/xdtXaDwC6bpzmAyuvfeuUtVz3ykncfA5A6e3xZ39/h/lm18po5ls+ovpnb43ufeG5Ex98S1bNxWk4XRUu31QffOHERrpcP3TdOp39pqz61vmB508tJ8q78cZ/HgEHkG85odtflCr8RPbGG63OymXMAgC8EzrzLamQeexVuU1xaNS1RX35FSy+AfjjRgeZ3MNNyO85Fxw1oT0lwX2N6Vy8GKPha3QtaFxV/llgPvG725UqphmYnIa3c5IpLyiMZ2DOIE6lPqv8alDjMPV90VMmePEfMzr1mklvHItMQmj4ibmZGQvWz2cUfCWCjPphcRvgVU2U7x2XscgWlK9EmKg/MUgbNxoAl+Pw476CuZPAJJfpLP87Zb4oJ+GdB6L6mzHMs1ntaYlb00CVgfsK8/FbCbaFGcOKduPXGCyoerPyCp7GgnajT5gxo3MGsJOKSxqeFgAmoLYpDoWuLa5NB+ndPdu6oONN6A8GXTa1u7CvZ4/ik4EEglzvwYaYOaAeJsjaMaFrCvBtOIiqMh0l24lS38slYMUqxsRpK2zAtknYdTo5hZdcYtJhRv2wuA2YRKxbvwqDKUDnwkT9iYsE42OjkoZgQsE2+2SW/50SuoQKn1IO8t0rGExHWQt1tGuTulyjc1CAysUlTjUvT84DRonoE55qzLqkhU3gv01xaMK4xQWewM/t0yu5nIg3ZfeMpDwJhgG5qF2Yaq9SYOFTRazQYCsDAALsbr2g9q7gQ2iw1E1RHdXfpr4X8TMHvjqoVUOtlYB6vmu6URFBRyEQZtQPi9sAHZ4n4rJ0lKg/KXgwDmzoyDVgooRZQu+o0BX9oKKh/vYXRoVdDEFfukldro16CtFrDHTT8pi8AkA/GX5aURYzA2dTtoXR2xaHQtf2QheAMJN7uImxBlf3uVRu3HAn3WMXquGX/FGBVQ/vmWqh2ubrHaaSTPNfM4ep75OMvTg0qlYdOeEeEs1+RMQOdQHvhKlGmxS31kxidb73Zon6m4Su4FEvsUOdU69Zf3T2C6zAcLfa7W9Fl2uTeI1B0/ICq0WfdjrvvmRfqs4PdkX22KY4FLq23LdFhJncg00r/9yw0uizoeuFt+KloTMHDqnloJSjLRP2rmS6N/t3Q87iTlPI4O7V6dXtekY8ugm4cXZgcPWjH96Zo7211REN7V1ztY71iluDZvneN07U32TguEM9s0n9swOLl/JPw1eHO4tXtqLLNUm/xmCT5ZmLpuvIOQDocCf0vq2blkLXbXdsAKJM7uEmU/liw3DE+6AwU42/z+qX3L8J1N46orOrQhcrV7sWz+6w5sLlm4o3CV1h6vsUM2cHGrLyifeAhZMvu01CV9/7p5F/NHqtztwmI0qzfO9Bov7VQrLG9SK9YwddJpjgNqt/yTuJjz2NzO7rxYXKFnS51twj/RqDpuUFVos+bW9+kf0uALC+mctRjsVtikOha4uhK+HrQSb3aNMsd6vJY9u8EZH6cclCVVa0rCJ4tejqQNBy4eoO+jWKABg099W6LUk/mRyqzmipovw8DGDiCgCv+pkmWaqlHbHatV2z6ZWnpvne11hybwhdBkKD79C38RgIDda8/nLBqsY7Jepb0eXayybJVyI0Kc8uGMpa+Bltv/yM7SLeWvf0TdmZotGWqGcAuKnmG23yVH9jI6oDgJNcC1JFwD5KsLo1MWZfz8B2Rj74IuD4umlVRJz6XnEgL5LDVp0DuoN+k9tBq7YRqtQsPNqM+lEm/a1O663Z+FqhywWc6LauDxfgO/TVbQYS6G9e/7CnsR65NV02kJFgQSgSUZxaXZ4qAlyq6LOJjdY7nULXHcdjR9AZPAYfZHKPN41nUyOHpezxoUEJdA4AQKYPTCqojGSHZJ0LtjrdpL98bHiwC8gO7Qhd5F3W58Ewueo1s9qBCFPf1zKSBw0q2wUuNVRe5MsAwI+gx1d8AMyto54Vmb50x11Auwwz6keZ9Dduwk3yvQeJ+ldrXPcgMwBAMgAw1Q7BDng7NHRVe5B3m9dft4GXAT/zF0ODLjavy2YR5wDyopp6JUKz8qazEh1GRZ8AgH0C4WvfZGyObYpDE8Ythi6193jufft/kMk93rSwXD6fOHZF53IonAjvte/P89YqZvmx6/JNjw/e3L3a+SovOE6ptYKd8USROvyEvAJ/+dD1Zxr21MrH3z8fpL5fYceuf9JOpNmhDu5XMc+H+GwRwLxVXOfwsnwTi2z45kqqlOrzh4oqyqgfZtLfkKb53oNE/as0Pn5kmH8Uvtq/dwQAptyh5cKVHeqb0+4wny02rf98eTDjAytKFlFUetO6bMJKdph7Kv1KhCblLbJjqnBDR592xDasiu8DgF/t3RXLuD1xaNS1RU7r3OI5u1wfZnKPNvnVUpStVGtkW0ZGznAZZNL0p4tyRsMbY3JGeeNte84gSmceZCdn5QsjJ1T37d78Wc/ch47eucJmr8uFCvwxx73YME5Z0DkVpr73z7Ldl4K4f724ZwJYqTor5wHgxnjOPwd/3CnMKHjj/LFzqWcVJpa7b1SjjPphJv110YjzvZswaykQJeoHAmVrWDPNKqflDDDFWmEMsDDuFGYriLL8rz326L39BPd31BgAbihn5Xyi/tEEXWOhWnrsnEF2z8jI6eXy5nS5xtjOH3NaL8XODwDNyvPOssLi2fgTADB3s+jbkyZvds+qUMe3Ic4aUG76pg5XjNYSgkzu8SYmkgnKn/1TAF96AV8LBtZBJnj7URTharWIssVnvw3g8W/iiW+scelkbvpt8ewrm85yvqZmwgJecVkRwVpekF49Fy5h8EKywiIcRrKitB9B5YMk7Pa4cGu8EMgLUZ73xAfWyNUuIrsEeeEFEmn343zvLDousd0RqXT9LLmUkyDK1R7lzm/IaM8H74wx0sVuMjd96I+J+ufCChcFhHWsZ7+e0uW6rM5NDxGs3QYXy6HBNuFV4RSDXcV4ZSu2QkLG5uJQbvo7jj+P6Kc8NpN7vMnoZILy4PZiprUaLLgsJT7m4/FZ9NA3D35SJKfukuysPPOTO7iQE+RQj9Or2zVgAH4tWeHo5rhJZdM3iQISWzVCBfu1KM974mPdEVdCoGiD0UjmezfRcYntS0il64/T969BlDs/05jgvumS9JaNsa1E+aE/JupfDys8D+jghwkam37hQFM1p8xVR4NtwqtGm4LDNVJWSMi4TXFowngXyQFcKr6y2WhhuAC6DHumerfmixibI7PcCeTCvDVSPK1kQxIwm789ua4x7k6ifM0B5qoHwwQ06rp7qGcO6A6/+uymx1A+H5x2uq9kNh3rtt4vazLLHSHKnZ/IaF+8Cqyc0HfEGHcnUX5911HV1lp9MExAoevuMXN9H8wFzGx6MOyN9x3EbIXftQfq7+idS4qCwOqM9vU7Y4y7kyj/huroMFcfEMNR6Lp7+KeKYlFhC9P4txdySzX4d2hEX2qd13B8DSZrcPyb7rRAzmg4haUaADhyXgO84G8vL+ulxYfXtDmwXJ233HTnl2UNYEW/hhxyy3UUagCKYl4Djr9cDDP4sGLrFFCcB/huBWepbXc1ZQxW9GsAeJc1jX3q0xroTsbZU0WxVHtAbECh6272yFteWLqT7w7tHMDCa1Ei+M5PPCL/6QiO3/pekLocnQNYPAk+JPH2tm5MP8SrZuw45P7vf+wPfly+8sPB74AddvH23w7jMN78h6G/BDoH4I+qVIL73j64r/Oh72t0fuF7GLp4qOXvX8LxW98LM7j39mHmLDIvC2Ozi2QHvxNk8v+QPfL+hZbpH9jm1T93OWndrNSmjnqUup33L021uejYM1HdR9raIjVdX7Lp1gEg787V9pk66ksaDMgMLE7xfiQT3PO+2bFEUjORb7XGCBP7Z/sWax3AwZZL1w9GRzVk8ido1PVwwFu+nxrEFa+db339GyNR6vJM64hWv836Zi5zN/2zcWLjideX3qvgY9kbb7QCQPet0+ygf+IvLihkAHT4p/X1F4ROJLjPeGeQnF+XL/zEvP6NEbAgsX+Pd5LtFzfdt6rVb4VPVzRm8ido1PWQjLoMUstsQTb2KHV5cUVjpspbx+FXJalrWxHMZtFn7hTMpXizu6Axy2QywT0z6cUArxo8GRgk9pezMG/rjF+FF6WhbszkT1DoelhMK1Jfkz8qNgDkEgDNfY3G5PDEJgnW4DlUKsG9nAZ8LZMJ7k2DMZLZwwzApAJM8JaD8MjGTP4ETRgfkna1fPSMTo664n8dARHmpUPPzk1o/KGPuoJYlU7wZd8ukE5w7/O0MeL/WS4XngLBegDxaGiw1Zn8CQpdD0O7Gj/y8ulmrh+kLhfvAIDIHESUY5y4rRAWhS4A0I+m+pHxvsHTusmpNrF/+C6VTCqDfvNM/gSFrgedmbH+A6812Z5OXe6PAVgkbd1lroqugSavpo8S+9sxsDeeNEbzTP4Eha4HfixwdfmFJplzw9TldkajzSRp6nZhTcZgDZo3F/Tzq7eGif2j1O1pY6zO5E8kdEcqeIDHXUzGOXUiiwepy23o2mnvSb0/54sy/VU0CVKYbPZa8yCxf5BgelUG/ZnbycRMoYvY0USJ4BuHCboNYNJv1qCILeFDAm0Ib3cY1QVkhFoV4GAgwcpNjGG0C2SavAFak3YpdD107HLBoeJE8FETClKX1zIueqWv+4E2l9R1O6Ou6j7BD8Dogh01qbxAT/qHqJ0SbUYbVUSnSBvDJvavtgs+KDx+BOgIB1pBJn9iDWit60GFHe50PB0ngo/GCEHq8oXlw3sLb5nxI0PsdrKxs1ZqXdPuEG8B1DPH3/oHAJPl4euFa+kRwrGaswJUXzieG/9c0hhBYv+p8vHlVvhVd4hFLx0OMvmTumnUdSfoGHKz8Zt8OssbHB4esOGBd4GV68XWS4lE8Bo2gV2YutwfY4WZKmZrxcJi9Tb6vhfXr/p9EMPvQBr5dYZcBrgx7bReBKZYq4YBVsZ4YWEcyQT3s8vFlrPAjM7NVMN5oDGIEvvfGmetVzUuqmJhVgU7g0z+d1Hd+b67r/3e9Uf0mQN35aqUm36Vol955c8/8s3o67Nf38BtggPw7NfXTVq+Wb/bWm76IA84i7Pgx+nXEynfWbGwNc2kZjx44qvrV/0ueabYwsw5TCN/B4hytUc571MJ7uWaCe6DI3gBjcaQ9luYtT80hogy+afqe0fVHbwDITuwcX23zZfW133222sKR7np7yR5/0zL6q3dv13Z+NT1kpZvqoCtjwcaEsGHQ68wLXgix/htcP+nC47SyN9R5SLMcqqRTNK/VoL74Ai/tpYx5pE2hm5umbuh7lJpRzZGmjBuBbky32RupTeT0TJOWp4Z3FYB96n/yOrdKvrQHZrmRWnkH4T1nbuh7rq5O8L23t01Axp1bc3KzTZeaN3EmXHSclbYVgH3J8zXd220dOvOjDCEenD8766o+8bJuyOs8yiNuu4bRPMB1abcaZ2R1Q5+2QU35BU7Xd071P0odG2BHFj4xsyecDbjCHABVoBjH/wsdedyiQjHSuEYqwDA6RZgArlc8D/goK0rXYDTXciJHaMSxw8ryQpWG2HNwErRjK9otwQ7AFaylRcs0GN0MisArNQdK5uXCgC44N2bUgrvKQBMAnAAVkBJ2nKCK4LbC9pvsXHsVVgBJRl+2Qnq5ptSNyugrQAEGiwGqufdwRQUTAIOeI/YkvuFUjjdCM+1hksaUABgUVtpNHtCZJow3l2OCRx/838B6O1HkHScD713vvdTP+CD3xsS7utA5wCAK5V4vt8XJBvPDP0lskPCP/HPL+G49yryL8E/oTF08VDLdLKA9hcBvPeDHTMONQAOdeHauezRV4HeR88jMyT8UQX0honZ2aEu+Ce0rb4ONngnNQ5de9y1W/ix72v0fvwHyB59NTMo4b7Oj2P//reqQer8zj/4cTmh1bUj15A0F6qPHP0uMHSlmh3468FbuyD3f//mkMTCawAfFuZCFfv7sPBaIDcA5I8KXDuH7MBfD/7mr7afrf9eqpsP/SXQ+3v/ZRPqzg7815dw5W+H5MJr0c5D1/bKeQBs6Lfe2HX0uxi62Cv7TmpkhwTwm7/ahChWisHvDYk954dk36sAH5Loei36PCaw90oF7HBg41VmZ8ekf45GXfdmpUvbd0xl+ucmmX0mJt9asfPIg77KS/D+uTEsxQPwbHk+zNLAgAOPXX5XmLqpa7Cji5dvHrFZypMFsIHFMSztnNUZUQeyXRNTiXHKwdapm+VkYvZM19wULwfVB4Bd7txUxgWQf2ael4HgtzTSARyDfbsnq3lhaqjXdZQ6P0wDvwH53RNVGc/r2UDd1HR9Ce17JqfzLtDbOrXcj2zf7Fi+nJC7p2VSddjjcV9n69+WusEOzOt9HbtreRntRPseDQBZOQ4GQOR31zIucPSxy3pT7hdKMeirjnZ7boecq+XL0Wdd1+sAsqGNV5m9Xc7fvI0ny2jUtQVOff7jP8BHADNzFv/wRwDA+m7YJsV/9+QH35Iq0/KmLv5LNToj8/b4418Nf4nL5KXqVMut1/98BMi2nNDvfgVA+cJPkgXw1hHNn6zsGJVIDTgrlzEbV9m9Umn7MpCJErPj7XHc/GxQfQAwF6fhdFWA7reqn38aAIySVS5WAOFjaVRx162c+tq1CqLU+WEa+A0ozFxme+KvmcdelepL71XQfqHKCrLK3JmLma+gtHIG/LOVWG7vhM58SypkHntV3tfZ+relbmR+57vtf/rCqXeHXRXsBLqu/K3/MaA/vBdSPnX9WFeFyzfVB184sZmpq5WCf/LkP33zyMWfHHOrrO/aOXasuxJ+nvr8o+cBFEMbrzJ718po5ls+jbruJf5ZYJ4BQH5POOKd0J6ScDyNaqJL9MaxyKKpvtAI7xAVPY1FLuIs5UEB3NeYzu0gXWggZ4D45X7cr2DxjWRi9pVxYJolqn9rGqgyACtVTAe/+yuAGQ7IGmYUfBUMnKLU+UEa+A1HJRomOWaYCBKY3qjCVAvItI7DOwt3DpjiCbkvaXha2OPv72z921E3cBmLbEH5SoQ7Aa/iA8i64euKF5SpcnBfYX5TMSGUYkJ7+GjVKAdcTMOMZ6LPgNjGDWZnchLeeZow3nNswOm/ETaoChC0t2TqkhXAj+YvyeQocinYE2VosgUwgzv8kuq7PIMBoPmqlWSdTMwOAIY15IbRADAH1FmoNO5zgeC5XR01TG1T53ubnEOnF9hNNX1FW5zNBM8a5NbCHn9fZ+vfnrr9KgymbBYRuzNMaNgTPYAyFcXDzd3EDKWowKg5QAUJfJZY9BkQ27jB7Eyo23rxKIWubRH8RiO7J7XMyAAWLKUmmo+KQle1N/Im8UFPTzdEQ1oTBsNsh7aTWMwMNK4kJ5taqK5k9VlRsFRDqXHkPC250ACcYuyePT09YtPtCdPtXSndx9rNFUWoWYa2np4ulpLbXtLoxCUfIHUnn+wPXLcOALx8tsH9Nt1xJqXQABgcA8AwEX42dvONZmfx7wporeuekT8q4ANgrCGj3CIXy+6ab0afdI9dCMcBHR3N56JcKncHPVwvAHhV96XXmreloK4HBLxk9Tv7BVYa6o1c7VeC+QrsUBfwTlDYFlPnLywfzo0129HbnyxmHwCTkJsf7Ipywt/X2fq3p+4mO8NY1RCjDXer3ZtafkpLkYhQ4ldN08Q2M7u+LV1Q6NrOmOtoy8TNMoCVf+57o6EFDi3vObVmwzpz4JAK7DWrAKyeBfls6HqU9mRnLHXBXDRdR84lx54NI/sD3qXMc4nqZwcWL+Wfbqi3lOojn/qhAdq75mpRXN9i6nz/9MF9S6vbK9r752r5J+MlMA2YhNwd7oTet81L3v/qbrYziD/v9qeV5avDncVNpUBqkGJTsSZt9ttdFqHQtR2ltY7obBmAqXyxYTjmcXm1ISCJ6BUvmNUvuZXAa5rnhM94HxRmqjsodAkA/kX2u+t0pfmWE3rXc4nql7yT+FhD6NJCvPOrzzo+0Pf+aeSD35BsOXX+4uhwV3W1CMVbp/FEHLrmrIlCuVnfzGWE0fK+zta/LXU33WkZ/2p6cJTZfb24UNlcL5GUIhGL9EfYGoI1mN00G5jRWtfdxYlexzrL3dSe3OzIyHjD6lfyN3QL4a0rvcZqSps3MnJ2xynEXObWc0WTBOvIReoKqi8XVhehCnvUEhdL9g2s8WBsq6L4550mfbq0y8F1HrSZpNwAE7d1yQ9P3diUupvvtEewcjoceKdGXtuKFAnqDAD3dfi56oQGs5tVr02g0HVvJo127KT60yMsx64MOzIa1PJV/Q8DoNrWMoh96tWRO2jUZT2RCWQkUM8AcFc1mWarNqnQVTTwWZgYSATd9vZS5/scyIomV/S5BApGNVFutKx8X2frT6gbEtzdtLrX2OlX0kuuLHA/VtiO1iVQ9KPPpj/5FcnQJYF+Cl33lDoXLLi/Mp5uIrX240MHRPB7IAD8SCJLeaYPTCoYJhnmMxKsyTPbS9njQ4MyLuC+D105YJ+A48NnfewgAI8dQWdfg5Pxvqj6AHQbeLlRpxkfhksFowtolwBQYGKrqfOPAtKH4S5L6087yJQBH93ID6DaKcD7A7ntgAAdUei6n7P1x+o2ug/7xKbU3WRnxFQmFcb9zF8MDbrArsGNJEkoL1w40/3g5aXoE9pJB68Gs5tqD/K3oWda69oGHh+8uTsYcS+XzycH4Mjlil2vRiab33s89360d3+et1bh62P1EyvLx2p896r8vVjRuRwKJ/TGPf/9kaVcMyAzrIrvw1edQ3sqH4enUnUGMJcZzLRE1QcwXx7MNN7GqjMDX7dqoPr8oaICoMpD16pbTJ2fH14uXIGnDz/xyVSrrT0z2KY+Ab9SHnZuYY4PK+ejY1ZuAH61d1dXePxtZ+vf0tg9uD8YrPows2l1m2r5mKz83mbU3Wxn7G+pG00rShZRVHrjh6Ij5SVCUXmYt56KPlErH38/2TgazT7tDqdfm0CjrrvYUDVgAG+8bc8ZANDwq6Uo76XWyK+MjJzNyCjFzY3xnH8OwRH+dFHOaJhx1gJzZrm4ZzJczIwLyLaMjJzhct0cOZbH/7T59t7bHjB0lDd/bI0DczeL/jngMpMXAeC0zi2eSyZm96ZLu8cQVh/AQrX02Dlj1Rn195gHlK+BieXuG1UAk7qgt5o6f8wpLFRgxtieS/H1AcxeL62cAzChHP8sVv7/7P3rd1zHfecLf6t2Q3qckYDqRkNeTixydzdATURb3ECTVEa2JQAE6JnHk0S4Zp0nTmJLBLPWrDOzMr5Ef8MZX85r8yL5PO9EgKSeNefJWiJAAJn1xDJJXNo5E49FAI2iM8qxRaC7QMczsom99/Ni1770BbfuBkmQv88LNtG96/arqu+uql37V1d4unnazzeA+Xtdy9K/vl5v/Tupaw98zxXASc/Uxhnvt1Fz9+aeU6lbSu3G3K6LYNVc/xiY3p1phf+z6yJ2+Ny5iY1uYMcNEjoXbjD7A27JRMstGX6uqbiM1nF5td+V+tiExkO+6asucAkAob93Jvz1aTDgy68D7C8t/GW3vt7zSe5f4fksh/dd8Es0gpe/DuBPXvcjKLvPRH3Tf3kLgdnBGXgVyl3jb+vnvMw3feyvEbpm1z7veRqVjtkjxS/1ly8iKzDeghOP+s73vLUzUdYytxqV6iS0x3cRVBpLl3qRj1RBJOtC527LlZ7d+Wo3hrb58YVvACffeuvbJgC87N1nDH3ewZ+YW/VEUY+5/YbGRKnpmfB+DKvihW949f/y13csr5cLEU3bd6bvf/JU1YxFrt/KBR7IN32DCfYkh/7eXRXetQDzBuAChh6pu9q/uH+F57Pce/sh/CUagfBGW0bLTrd9ZjZs91c9rvEdZsqIa3YXAJw8Kh2zR4pf6i/fvysX4e9+d1YjvvO1mXY7OdZJ6Kyo0MT5Ci/yUefv0evr9dYPGDs4e+c9KxtdvW/jud9VnzryW8neeAf8a2/j5MdKnbytGm9u3bbgqlLTu8r7MawK4Xif1s4TZi8XKpq2bzj/01mtmrGS62tf9aAJY2NnlCnAEIpv1lwjcYCbcucIeItsWKbr2L3vSJNqvaJXR5asqh23496cmHyvFUj14Z8HTPBkt9ufAFJDqq1HPFRzKw4wSxrNuUffyjTqaijytVMbGVu+vFBzBMdOqXYnt3ME3GlYputyjZ/7V1Tr5WxGnL3zKuvQ7gywwk15y+zG0izsKwN4aRyY6rNwSz5UcxebhmRrS46vHQAr06irocxtHD91aApLNc/Alu4c7zs8uYsIEo2Trrp8k+cTVOsVBt3FMNaBgDMhnrkCYOlv2aQE1i+L+9MP19x3ZXufdUvZVw+AlWnU1VDs86mWvEKh9gnYpbRYl7uIwHQBsLSzCnZYAokCgISZVwBY17q+d7N0S16FP4RfgJmrXiB+SPJmxczVhLkAAPoDLN2yACQcpS/d6iia+29TrSPhbKSd1cBqzFxFwtnoWpdgAnFXAWnhV4QvbwD4sMLQZaDjNbcfObQOq6beHbRrn83tXkqLwir275gnGnU9sjgrC/VVu7syL3dznVgF2OjwaDeeGgIwYgHJsf4zAmCn+0a8FRPWP9x/RgBtY/3e03c+OtzvPf1qGgL4mAV0DOLE6+DDxkj/GwBiY/1v+Fe+CZx43buUjWy5xqKo0tFhnRgesQBjdLj/DSA2DPSnTveNCBz5unH2TaBjpP+MAIyRfv/hI4PCiednr7d3wxic2/z7PqDvmdz0CXOHpPbZ3O7K/OoBMTpJ1wFFKCB2eG4x9MTOTjkLv+qu5gy8d/36PQAIXLR778owmJ7PQwHfNT6G3HybicC/uEoAKV72IiZRQdI7g8Bzq6+/Opbn3Sgot1gE712/bnQDx1vy/v5xDoXVKTxzVcJZm2WTBWD1srg7Q3cCmjA+7tL1IZC+H/XEbpjvyuQXqzgD5+YPlD433nPRDu8pO0c8OGDV961v3pg9MvAb+P7F5RcBwUzJHepS25G5ffW5rwml3eoD8B2w333nP5zz/DX86t+CWUtX2964DwDI2goff/zyb24D7kUDztvATbwsbpItadT1eOPtJixGPbG32hJrs1WcgTPXX33TLtoBOMoEc+Pha27aNb47iyWFwL94gQOmMumk2B2wr2KdmSVu9cNqAFJrCiscXCxiTQJAbLR7xbsD+dNHgqTrCUKVuJ8wtZfyCmfgLrMqF0ukQGqVg/vaNwtI741az0G7BAoMDhdMSIG4Q9bejsgZBP7wNFoN5irgcsEdqd3lt5rLVwHgpjdCs/Vuiks0LSfpeuxHXVBAPhb1GSJCeSl1Bu7IPtP/Iu6/+6JSEAVeuj2MBXH4/sUdmKxZpmAWyOQ74J1BkIhu4wqqwTSy2c6St6s3L5dcocoCEDtDa10H+V4vT4grFeOpCNoZ+PQbI+Pe/Tx00S6/CPPmcfOftugtvn9xKbi7ymHScGBXN5SIW/2SW4Ln3ig8L0WRrWjU9aTe4yEAd1yWvzpdel8atKdyAO5e4Z7/pbbelamc91uBs8NFKVI7zASlyRyH0wPG3dHWuTJV1VLLU1NTk6RXJF2ExhmXXeEdXJT/nLEvzOcAYOmK51EufX/C3zXmsLirlGmubjUh9QYHKpEoOCzeuBcmH18E0PPxxLyqdp9Zn5+fX4gsQdIDW5KuJ3rUBThXeaA0qqIuQ2fgS8oEAhftnnQlHMhU2XDKjyPwL77KhXI3ErRKv8NMEWBClrjVLxm7encVN5AuegeBpOuJxe8NoSd2yYGSZ4nRoZgq/8KVaRervGwRq8AB3h36GYfDTOmqNK0fb08MMLSNRfk9BoD0PGe53ATrBgBmks1Iup5QlABOCaTdwBP7utGD9hLp0s7AY53gpgICF+2e9qUKcIyymYvDenACoZ9xONxUkCkade3QjQaQdWToVj/6k4hjNWaCZeGgEx0CAI4Mk80acLsgDqR0vQgkR1T6w8ATu61OtKVKHm9pZ+Csv507OSBw0a617wZcZpfGassTbebsvwr9izuqRUJ2f7gvZajvHL6GTvj26CO+HNtzEO+71S+5GYwVLq7dG5GJZ+ed2e7R1MwX/Rk+QaOuB8O2D/N2IFl6ZgvrE+Vf7XnCyICFRHo94on98kZq7b0qzsDtO+nD3gEevot2AMhDwfG8DKrANT6ub5i3pAr9i0M6QKHKMd27p9LdMesTABA7AwA8m3oo1dneDTDfR7x2Ct/0BqJ/7p5N7wwC7VbfjThgtxd5C5zpe+mWeWBO6fOBi27jC/Podgxd36XnJvimp1HX/pPQL6clPz9T8Vv/0vadu7WzJAzLLiv9VbXYdkOBA0vnRR7A0vlmyVxgbbxlFbjmApvnvbWVa0sF9TacSyl9oJpzKZWHPgHm/rki3HMugJs3YJ8DcPtDLw72Q4Rhrt3wLq3j7ljhbI9llxUAhwsF44zA0j4dmhv7o20ibrVmWX8n7AsNGfrdWisowD5/eJV9iM0LkWqYXC4AS+vC/x1vA7h7obElba045noXFmh0x9haurz6BtK/yfnfneyBM17PHZFGXXsn01VpxS6rgbHtCocDKGgv5VI7c19FuTNwBQW4eb+NRL2BF1HpGh8orPrO3L0wrsIujieqrQh3TKC9eV62m/tUUzv5iLdWFo0e4KSp1EkB4w3XNb4e/LknXM8Nmnarr0ocsCsAheB378sHNFXeyQIPb7bXs7LAe+uLgqRoz6tMlYN9V6kGxra7fr9hHnhLSgtYPz81zqz9iT86LWur5iN+cuJaLhM4hedtx9hwfFc+4g8AxvA+TEwbhu+enyaMD5Lbv6ysiEvFBsa2u7qXljzolpRfANYAJ3d0f+KPTst4FaV3FoHpY4FT+M3pU3h+Ync+4h99WGofJqaNky7tnp9GXQ/U6lVGWMWGxrYrFlMH3pIF3foke3A1VSFfTPhO4d1bP8GE3J2P+PKx8wFqq48MDuoa2pJ07cpKWRMAWMbU9+8UeJdA+JkAuEAm5V2VjscrwgY/Rh+scLMkNiS6UnH9cyJrAmAmy25Rv3ffOVBjAG9NL1L8dLi7vGH9K6GrCDyTgrfxMwGeFQATLC4QPNdKRFYYBVrfVE29ADt5DMMm/D/3wPLsw7StXxjvWW1CsIwJMIF4PGoBIJENWpf+0xucZcw6khaIJu5HaiLTWVLfPGz2af+/dU5macK4G+UaE+61HHCix70k0fGV7/OR744J+4LyPzFyI9fxb/5mwL2WA/o7AfxAlYU9MuBey4GPmFi7GPTnkd95p+Mr3+ej3xkTaxeBtjcA/PKHANAxCPdaDk2D/2ff8+8dvBtqBf2dsN5G5JFe2/A55XJvH5Vo1IpM/+3nLPeSRHJI4PZVNA19DyOTJ8yeC+pkN87e/z5O9mDtIhAbg/UP4Yip75n/X2EwL2M9U699NPh9/889pLu+/hBN6xfGOCPcazl0fOafetxLsulNnLW/G7EAjBEBbH7PD5Z8E+6tGcQG/9/Dv/3fa0zaGBH6MaFOPDkkcHMGTYP/eRBiJmzufEzY3/XCdAz6TYDVd8+iUdcuONGysNELxLrzzL8bn2heNazwUwAwBvIbp4CmzrnFyPnJOiz3fgydwwNAzJwJYktaYIPr0yhIAOC9zsJGH8BYT/FxsGBT59xCmxk4ywfQu6n8FzG9040aQvJYnvUC2eZ52e7t+xTJQ6uGhaJyi6uI9SxPJ7tDD/zQL2p6TuGduz9xJ4oHzke8X5i+5qk7ffCbqVt0iypqARxvvq6c4G17NuTMbxwHwAZrb2F9LQv3ur3/eYkPOfMyCzDWm1fZSHPvaF7Vr3H4bvrhuecn6drfyY61NHlBAcat8Vl9DB57dXxcdoWfABBbGp8xgPT961PslioP6/3oXJuaVMH4vPe+vrW/Oj6hLPCWiVuzm7MAkGw+P3nFMIHY5rnZx8CEXXevT8oeuJMTk7LTU+0rwTvkhrXYsHSujb/XCtjnpyb8dfnu8XHZidvvbZ67iszm5VuzXTDMm+NXvZ9ZrwPclELcVLDfZsx+J/jzgOAXhluT8xOGCWRujc8mcP9t99yFqAWYdWvuqhNs84q1jE9d5BZgPHuu1lMXDevm5EQcCBNfG5+6Ygggdn/8iiHC5v5S0E2SzRNz73lbNrI2Sdd+N46WGdhXAHsGC769lqU7yyOfAJxprDABMx89Xt0Pq38MncMDiFlXgticXBzcUVj0VslStsKasgDMPw5u6ZiVB2bjvrN8AKfuS8BVAkD/Ge8tpUawmcMKF5hSsJVeUVmTbs6vNWsFWOCBB36gbcSaP+irGbowhpODLc3SZhqxABM5rIffp22JTWkBmKv5XQnuzGL9HSBM/LKCDQG401iHGTb3sJtoN/0R9/yPonTxFFDtPZAD1zgcBSjPAbn/LGxRPx9bDJ+T2eEMSFWErfZj1h90YRFQzHtzzkvAzGuPwXYOjwFM5AI7Ke9+Hdz+WZfTOCd8K4AD777hS9cC4M+SmCkBlwUe+IHE4VuzQOAU3vdFc5B8xPuFSTi60NFmGrEAc1X01UmzAEAmfEPUnLRXn0HievHdloFLMoWSbqLd9Ifu+Wtf49tHo3b82+8BMDIHvePFo0sx4SvD/vOx0AuTt5P6U4D4ZZWw3v/iwUMe3v1uJF7lReRdz8wb8N6wfjy8zTBk9F2SpZoYgHZPtb3GPtlQofAslhCl6euP1iw4i3rx35yOBDqQPuL9wgiWBcSnKpqpX3SXCRV5rVwUAShW1+Oe8PFKmHhCBFseRElz193EXMvq6tDu+R9J6dJZDsq3qxeqTvyLHdZ2dnpd8IHhQtzTn9HaV1+EYd7YsniBc3iw8m0tDjelVYz0ocfGvQBDHwA7cJYPwcqkpqHVAt7XCVTb63scj+/ZFUYfqhdaty7VOZMONVv8EnXvSok8GdaJ8+HUVs3d6ya+m/4G7IjZf+naDDb0Vr6HW20Q+qkdLkjff7RH8/nY2D17qyy29a6sJl/UE8yf95aO1R02cif1/uPZr6YAOIgNrk8lPwdgwTQlvHNw94N2a04dr/bDnHp8pcueAbDNLg15Ipm+3ehbUlni/ebcxqnqzd1nWTZoY88DGHUVH9OmwlzPHTIrrYiufDp+bWvdncALfl3OfK10WG/Yv9I+UbSvzcemk7lYUQCQsS/g058DsKl6gh21q43uTor1LF1He7UfV8ruKcXVx6c9Ots/cODHfmKtz0QXKOreUKfKE+fWjVl2qnpz97qJi/V5fxRYZ3n3/wkjS8HbdcsESreKl+99TiFjxsHiAgnwLuH/6m3EZim0dgGIg3kP4bjgXcIPG4kj0QV/+3CDKFY3kn8idNnJ0Oatc+dyW4WNOIcH1lh3yY+t9rlz3pTae/Rmrj4+0iX0Qof/xWwSABMKDV8T51BMLG6bjdCL/8fvHXjb+oVROywvcOf9cxEHPyreAOliZf/j7uyWzV13E/9w8Prd8++/dMWGgbax/jM48nXj7JveDHisb9QE+jtPD4+aQHJseHQAiA22DQ+MWkfOWhhJj/QP4cTwqAmgf7j/DSA2mDzTPwB+Vpw8awFAx9dO9FvgI8Ojg4D/yYeNkf4BNtJ3ppG3M+9lnRJ6gJQT+YwsAJhFBQDprcIGqurMtpeLoT9ClWmAPz6HH7rhXjZd/GXe7W+nTjQwoQz887yZqJzcuHrTiu+Bv8pVBw+/MAW+5WxO/y2KwVDCe7jIrPpujkUDgOVnYotkBUq6ifQ7Q93u+R/Mvq7e9ev3UFBusQgARu/KPOtB+d5nsMEiiqpY9DYAJzPdedYTbMQGO1WQ7XBXUfSkAUa3Ak4c9tw9+Z8YdmR7m7d9uGHShS4ky3ymtJm8uxD5jCwpvDw2cgrgI1aVsFHn8MCCUVJ9hdjZkWHvm8WYiXb38ZEumQV4j+8sH4Atu7R08bHuBt4ne1ivDcBEuygfOguBXIcA7w088AMn3zj480T/OAE+AEQ8n7nMZKU34LNjw51Ak+cWP29YSIpcfatrbAAdPQgTd5lVMlEPmnvYTbSb/ga4538g7zBy8wcqj7vv/IdznlGXruAf/xBA1/u5F/4IsM8r41umhPHsd0355U+9B6D70s/H+pauvvxFoOvudSR63oHx2e/FvmnlLv3lbT0ojd19p4V3357iYz3v+J/gz1/4528OTP5k1Mo1rnHMdo8l7pd+95PRDfFu5DNk2kok0ofeqR426hy+dMUHwKaKx5E6rwCss1GZuqseap9gLY1Lf9oaU4nfmdHO8v2FPr1K2MCRT/5EW+o2nNyJpk5ZPnQefeb7K3xMJp6Z9j3wlzzpPKj4hXFy1giLPOVx1GjxfNQAOSuRSBel9gxgb5zuSG/KOpP2XPLDT/yO7G+LGt5v7u7fjm4IL2vaTT/qf37+QEZdzFWIDE2cK0CeoXLvc7ivd006snkGizzYiA1cD3dIewo47SrDnYEzm4D/CcwpG8/kXNnIacicTDhXvK1X/uLA0oZYkuFnpJP3fnzu/EwyWMTUYTVR5/CAO1PibC3WfO7cZe/9FfsKS61f2fVCfUf3djt/a9wU/NyAtyxZr3Md1wXuzyTSzozvLB8A1jYsMAd1uWmsZPMnqfUZYP5e17KMuH91AVvGi9i8wtPN06EH/podPT5K+IWZlOnUsgyaqTvDmqMW4McunZvY6AG8e4dzeSPtXI7uXaqBCRVfvwoEibvXeectpbSLWBU291sbYikHKBW46d9fL4hPv1Vf+Be+CQBN3wZ/ywKAp74V6abfBv7kdfg/4U+60fTXAsCXXwfwlgW8/A0g9m3wtwTQ9C00fRteiL/sjsT+6W8CeOrbwafxlr7q09+obzD6VslUjqWFvksIgAnwt0yeAoJP7whRAHHwt0zAeAuxb5eE9X4EWBoQQFx/JcCE/kHg5a/rzANAIo2KEQl/q8rcig8DL/+VZ58qnLC2/GmL+4x/BkLTtwB0vPXWX1ulltnzIEkAQCLlG8MvsF9J8VpaZvVM/MnrLO2nxnRt+Tlgac+s0d/3vNb15W88FHV6elujB4VJpxApMNIiaoGnvgnghW94rQwAuGcKUU95edpP0Us8YniwsLkLBN0kqIKdTP/Ctx6FCaMj+1TJ0JSl/NFedO9z6SYD5ZU13IgNb6d6eDeHfkTisuAzDNrQmYCbR+ht3Mum/w6+9+kGTtwZJAAXrXZJWO/HwDl8UX+l4Cr9g/I2OAcLX4XQCttiZLZdzu/6zR5nzjy75oWwuVDoKqxm+uqce6uwPGHx4eqXURrr+F777CgUqviIj5hV/37APAdtQVCYPBA9myBfYgH/vQ7L3y7t5HfZwrbr2UHHqDA83LC+VdhNgiqo2/QPxl/X9Bsj0dM/kkNCD1u33/vsiZfeiL31Dd19lLZBuTAlWh1Y03sPGwe4udc9qe72bUD9ZjdyNTQeKYCOz9kwc1iex0/eaOC6YQlmDsQDgwPMKhjNj4vRH4x03b0yOPzdUI2GmufuRaY+W+599vA2Yh8UXDUyzbLr7KXv7DmoPHZKte/Zi8Lm+W2l69puBJ2lw82xTnBymTRzmAfW1D5JDD/8PgnKgxubGcOrrea7fO1xKdAD8pK6dGUwnNbEWs6pWLc/qNpm77M3CFhRO8xFmBt8PgLSdeWNPqxfxUQNRrpzHO61PQcr1vHrtiHk/1OLsdgfU/G6N1QTe7jFLXZmsCT51celQA/KwfOSCqUrEXExxqGYeH/bCdh2x7srpqWLPVDpcqdkyWfJCPN8ys2rWl5vcS6lxbrcQwCWdlaBVHlS7PCdrhWFtMgrIFGmD4nUugSYuZrxXrRMpNYlE4izIli6Ja+AdB68+Z6VV8E+7U/tU2dq+HE1czZJ1JZt9tqKKKzi8bldPDDf9Kp0LYoB/t5nDoCFbmJE+QSsbBValY6ChUKnE3w+qGYwX/pZOjIv1Bzt3nyvsVETN2aNkf9U9n3T4P/Z9/z/p78TznmF/t+8F/1N+7wPPJIbI8K99N/PYNT9T9qROB95V3b8mw+6b8wW9aMOofavTTSWFRBPjHX2d19XOp0WAGKd4KYCuIgDQJELNghU3/usEqXi5W/EDklFHqs6rAeGlQ8+nyhih+cWzSrPUhnrKcLoXFnQO9ej1a193gceydubp+4IFFEslvjNN7pV8MSW7Zt0EcQjOuqKjQAf/n8B1t/OnRwcNla4CMDmw/cOAdX3Pq92ny19J1ZvxA7+lt0jt3P+H3bOGkk0vxt8PlGk719nh6vaffOiwM0Z3PtC+S/J5vOq7Q395gIAWLfmFw9tXvzWJQXnWo6l9Ip87O47Lb50JVty1E+IJ2rUVQQA5bqw76QPzwH2Im8BAHum9fBloGLvswsAayouvZmEW7IRO9j1O69SkVHAlEy3TKnw07/KfQLqThSxxWuO81CbM8BixYAs9Hmv31wQCo6OI+o33512FVwmAD7yxhqNuogna9T18f8GALgA51LKkQAml71loJtrBXUegHttqaCAtfGWVebqVVvn/OFVnHO9E+3t8wBuLres6jXday6wOd4iAf/Ee/tSqqDCT/scgEsu8PHFx7/uVOsWP/gO7d0K6fJ83n8m9EiuUiVDKl+kQl+JPL0yRd2EeLJGXWF3cL0DOd287hp5pQdW3heF1WALrrc93UWw29z7GeEGXf2nvxvXj1N/Blc9ASOFfGyLM5q9ISdLm+U/MFPBc9Xk72OVHZ3hz3HvFQ0/Bs+9lTNBgy7iCRt1EfvLpjwhtvb2335KVHkJQQElC/tzVr/w1xEjjsTLZZAgnsBRF7FPuOOyo3vLij1lT83sHId9eeOEHpy19a5Mlcwe2ZMwdCVIuogHjhOcvF1JsnlivoqbY1E+klq7sGF5/0vfn5iXlSGKZGeCpKuUuRzZv07turpl/cWrnYpezee97T9VNMu3xTFXAfYFMjNB0lXKiiT71ztnLK1Ms0y+Kn6q5vNehedNlO4G9l6sSpGVCZIuopGcEkiXaFfHUEnNRt9B0D+V+7zng2BWEYDJyvzm+9IVGxFkaeIRhJ4wPsps/wpOckSltzyDdcUYNporvq3wed8xtmG+C0eePHKu1G8+4DmnY/Qe0GMJO/CPjmnUtT0nrIdZOWMAwDJd1Qc+C4n0+qzvgd11g/mjCwD2YubQNABV8lPg8143XGcmkVqSwAJrKfObD0A42NGTIfHIsPXT5moceR3wD0SlUdfjSHoX/pGN1y/vT+LOPVNCO3SowtJ5kfdeILh5A5sXgoX2314E4E4uF9QF4Job/UmHCf3NeK8qYOm8A+dSKo934Jwr6jcVzFXs6MmQeGRIWLOehP3ebMUQ69RceS2ufgXAiV6413IHtLw06qofI7NfMUsLMA6vLCar31ALeSBwJ6TAfEfMCvBeL6j8SYcJd2zpdxMKqsJvPrzlfNocccDoqtwwwysHV44BoHd9gZ2iUdeTi7tvywa5/xfgXMuxeHZ2NxK69bF6Rg0n7vEWSXV7AFG/U9lCKzcX21wo3JrGnYGDupZJo6762dy3rU9FDtg5uDO7qqbWlVp+2loJyf/ygeRapQtn51zFXchVJjANLMOkCeOTO+ratzmVo6unuKuj2e7P1vLTliTWqWoPZHOscsep0kKlJ1nOfp08QNL1KOCfAM29T5ZJRf4CWArgAplg82a6SwBIgGeF3gzKUvrC1hTAvSeG+irBsqb3mfEaU6ZLACyFjAmHCz3g3w0FVctPW/LxVar5g9JAvRbltSAuwFL6q5RubOlIA2OZdDwOyIO+15jWunYm+SbcWzOAcUbg5ox2CQ/jjNBPZ2LD30HHv/mbAf0n017hRyZPmD0XVNvQ94C2we8BiA3+H2/ixn8ZMdcuBlf1337Ocs4rdHzmn3rcSxI42YO1i+DDPxy+/33/TJGHsgmHdkUcFIwR4YxL4ESPe0nixO//kA/97yNi7SL4yHfHhH1B8ZF3pd/A2KgJuP/JO4nmIEOjrh1hQ878xnEAwy3z8oTwXcL3NU/d6Ys0n4H8hvewxvcKL5KHVg0LBQ4g5Z34wU7l1fH2Q6tJM/QdnzyW590AYt151gvEepank90Aet3wHLE4uZ4htqavZeFed9CCIADW56wmLQAnmlcNS49SvJ/bDs9JNx8sQnAhSboeV2It41MXuQXDvDY1wbqRvn99chbcmpyfMMzwqqXxGcP7782Jydk0gO7xcdkJh5vBG8/GZ8ev8NfHJ5QVXtV1bTyXAZC5NT6bADKbl2/NdgH80Ll3AOUN8Xto2YnYetBl3ZyciAOG14IAgL80Pq46AfbqeOBdRDew1N3rE5gCXK/rtzskXY8taVtiU1potXOwc2ntEt5wcrBlKF3ONFa8o4oCr/Br0s1xODDB/Deer2OdrUlHivCqzZz3ac9ggQPWCrDAASwrf221Y/TwDFUDsWUfdmax/k7QggAAt5SbSwBYlu6s/k7/bK7C3hC+6o0OHNjTsEm6dsQsAJAJxB0AOaZdwiccAJGnM7YK/S9or/ALwCrgShNcv/Hs5OBiwXOx7F+1osfum4DDPJFzGYBF+AmkD92SVA3EVnhN0W9BHrOA1K2ooPu4/lkob07pMgDcXD6wD2NomX5HRBGAYjBdAEWGfKx3GhAsC4iSc6GDBSmWMoL/ArmvRI+oDx3nh1eFK1niHlqz4AxARK2WaNBFbNNAoyuh4QbTAovcRsOflYierOlcPrDFJunauWX8UotN0dMi7RLe6APwy2oByrzCr3Lvvrj9VQHHUb4ZQn2KKoHYpXRV4kLci/wpj84aj8XDY5KuPeLCHR/p6J6FPQOg2vI5P2VPGa+Gfzvcm3Nuf1XAnKKzLIg9sadtDrJ7JGGripkCSddjiPLvbHHdSpzxka5ZOPNbXJ9sPq+aotLFzDK/pNWuCijxHSsUaIMVsYsGurWgMVdFV7StxU5nHHqv4EFuWiRdO7eM3/WkSzEApgPAufofttnRV+4V3pEp88aOV4Wj+4rvblIdENs10CpNUSjvsRIAXjKyYtYPboWCtXmAT1qmJ4w7j7ATALNWIQ0AZlEPswvbWi5eEkHGlTtfpaUrst/CH3VxQZVAbE3RAGCVftcNWC6AnmA7dIBZVPBHXewAvw1E0rUjecNCUuSwznuQtBa1S3ibDwDtZnWbRr3CA7LV3sVVmlyHAPdPtWZCAugfoEogtsZmA+goa0sdwrDyANpM3l0ovTeeHhvpBbxBWdMwTRgf55axcbojvSlhyxPJhJ3TLuGdnDXCUu9Xub7CK3yh2mJodd/xwAofk4lnpv3lCgVAxKkSiG0aqDxyNv5x6Xd/f3bDnQXwk9EN8W7JLzNfTyTSLVe955LxAzx0oVHXDusICs7ljbRzGXCnNtLNU4FL+EmZTi1LoMLVYOAV3v/Jhvau7LoIFlUjvuN1St7vm1d4unlaP/rx9oMpclRKbMeEiq9fDVoQFICpDbGkACxtiCVZ0sB6bp87n0vrd9MKzuNpkaffemIbQ+wt0x/3AOBp4f3hfSbS+i9/oUAATCBYu2JpAaGDCgDsLSu8EAze1WVXBVcjkRb+f5/6JvTV3n3mre5HwDLiUWiZDykTX/7GI1lenkZJSzLeQiIFgL9l8hS8lql/Nv5aAE3fBv7SClrcI1deAHjhWzRhrAsXABx/1OR9FgrRv/T9TCFw6KZ9vAf75nmwM17pGF1VcVW4y96LXgFAygVocwSxE04eZS3Jb6RwVuG1TP0z062JezvqD3DTognj/lO7p2QzT9YjGgsTQNoBP7AeI2jU9eBI2bWGNN8n6xGNHaCx4UXe+yG4fdBLQtK1/5i1LrPzZknWI2pg6/d7nCuDfVh7D6kDP6An6dp/lj+s9Q55noxH1NJwprSCTVXe+5bOp9wF4PaHB72QJF37z0rNIRUZj6hl0LWgP6u9Z+st4B/8N/xpmZ4gCJIugiAIki6CIAiSLoIgSLoIgiBIugiCIEi6CIIg6SIIgiDpIgiCIOkiCIKkiyAIgqSLIAiCpIsgCIKkiyAIki6CIAiSLoIgiK0hV4Nb0N7+6OWJIdn38PNw/OEbwnhYmRBGH5X3wdDKSLpqtJz5KOYq8wjkIYsnORNZKu8DYqfTbWnCSBAETRgfH5ZuPXp54kNLuYc+VxucVA/dEE2vP6RMnHj+CpX3wdDxEklXbRTkIyhdKDz0I6hiyD986Xr6YWWi4/fyVN4HJNefb+iEkYnHU6ce13IRxGPL3qTrua/vZbBpVUw2hhopN31i19/uOJR4g1oCQTzG0mUtAul0Op3ezcVpsyKxdFTQ6hQylhVVtEp/u0ccg1oCQRwo9rTWxcz3YYwAwA9qnnp3/SbnS1dDnvTz7Fqu7kice6aktkAQj+uoi7dIAPl8vo61OxWIntuQA3hdKMQG9xrKGC6NRJrUFAjisR11cQcA6ntOey0QrM0LjSiAc64Int5rKJYqE9Tfp6ZAEI/tqCvh1J9gMRx1FRtSgobEIhPUFAjisZUuM5zicRNAGkC6SwBIgHcJsIwJIKE/PbXLlk3FEsIb8aTBUkBCMO8Cnk3F4wCAFDIpAJkuAYTxC5YxgUxnWY7SQBpMIC4AZLq8rHkheSYFgAsvujDj2RTABOJxAAmdCAr0VgFBPL4TRrEa/LfjD7+L2PB/Yv2dcM4rjEyeMNcunuhxL0n0337Oci9JL/oR4V6SMF6b9gP2/+Y/j74rwUfe/Wj4O+i//ZzlnFfgIyZgfxcAH/3usP1dnOyBfUEBHYNwxiU6PvNPPc75pmG0zEYz1DZ87t7Iu0+/jrOb38ORAXReBPiYWLsI8BETN2fQ8W/+ZsC9lgtCGGeEey3X9CbO2t9FxyDWvVmry6gpEMRjO+oSwWQPK4aJzCaMzpUF3g2I5KHVZKY7z3oAJI/lWa932dCz1+8MAsnoS5wOTIC7nrYlj+V5N9B+aE66WhiPu0BTz/qCMQAYg+sLvBdAU3ee9wzmVenboL2bCkBBucUijIH11aQJdDSvJi2g/fBc7jgAYyC/cSoM0dc8dacPbtEtKvDewkKr5UkXF9QWCOLxlq54PC5gKwtWHrg5MTmbBtA9PqH6lsZnEwC6ro2/1+qNcQ5PzF2JCZTogitNgNveH13XxnMZoHP5+hX3svfVa5cuomvzwuRsEmh3JiavJgWQvjWe63hm/ErJBqyYeQUA7r7jnHsbrc7EpRkAL41PKAusZ+n6JLeA2NL4TBiIW5PzE4Z5/2333AUYLeOTV7ysuSDpIojHVbr06zKjZ8++DjeXMQ7PYnMGWGQA1qQjm2ewyAFs5rDijWJaHYlNJrBSsptCJoCUXjbbzGGRgZmLsP2Bz5osMmsemOcmrDWFZWYC9gxy7DoKzIxEdOq+DP/oXFO4JYFl6eTi4C0zcHIm4ExjJXzPx3BysP2dEOlNhYf/PjNBEPs96oqyYGRsb/7oMgALgLKV52NnRU8KgbjtDbI2L0eDrnLA1GK2AhRZ6cBnAWBCAjYEMxcBR5nAJlBwcqU+fAwr+lq7uQq4ABYBxcAdBagUYKvokCrhAFD6T7MAQNGoiyAOIHtZpmdeP/d20m9uDHwIACzlT8eUd42WAm+Gyat5KitwMPNGOIEEAOHHDiiAQwGuEt53vtToCwPao4MuFlmIUwBHNrg2svVVsCwgPhWd/xIE8ZhLVwlurjsHoP2UgF3lV082RKwPwC/LfnSY+XMhSyNL5VqdiJK4yhMaT4k+VT0LYpvHgp5n2t9WfG9EMiQ+pPoniCdAukpnVSYEwE/ZU8arWwfZnPUmhaXSpcz/7pSOeOQfjyTWK0LvMCpaMLd589CZBlAZpT1T9WuCIJ6QUZdhys4cks3nVdOrW84tlTNfLaw0f1q6L5915zrXowtXTChA6KneVhO7TdXzzlbZU+58dUWLfK1EeX4Jgjgg7GWZvmTU1WrPJIG4rbaI1pOu6lM6mUgVSqWr+dq5C6osJQ7ppbiVdLmzyegU0yydlFZXtOjXSlD9E8QTIF2Qkb7eub7OLQBAvPyyTLDjtKijL3vbeZVH9uUDAGNmyauIDiyAu8qVnYBRui4GFrxsvcy7tdIBkGkg3FjqlO0x1YGCF34YANUKMBMJQaMugnispcsbpsTj8Ti4NWPLToCD91RMQntYr7d0XzAsoF2grdTHjM3LFqmcjdGx4cj7iW6uXbBTtoJMCmSd0ov5iD++smWX/krEIQ2T9QcjL0f1AnqrfCSQzQeAdhMuMxlWDQsnTPQPAMwh6SKIg8Se1rrUiwAwCuDd/+FI5P4QK8aw0Vx+Wf5EW+o2AGBTnu5gqR8gXjp9c1iLLJv6vZ5IZCLvJy5YIxupG8B899gdHVc4gAr/O/M1cQ+Aw8YKF5f56B3z3SDGmYERlrpRHsjJWSMs9T4cNVo8v7Zx+kjqfYjfAZhLTYEgHt8JI/O3LSC1BqwYwl7MHJqGfhboaveBmz9Jrc94TgUnNtLmskLRW5V3Xe9bVzqRvwCw7mvnzvtDKABYm0mklmeBzWmeWptB+LBRIcgCgLUNCwDsRd4Ce5qZS6EiLq+mU+u5iDp6gSZlOrUs4c6wZjjTLLWUg1KNcedDEMSjwdNvlX3R9K3SUUwcYGkB4f3JBAAB/MnrLC38S7z/+uv7wv9WRP8SiH0bwHPfhI4UABKp6CfzI2HAUzoXce/PeJBKIuV/KwCwdMr/bzwMpL8GwhAMwMtl54XE3jIfxfvMW90Pf5j+lngUWuZDysSXv0HlfUC88K1GThidcIzmAkARcPOA8v50VTD8ygeXeP+NbJR3wyGQ/5fy3K/qa/R6faEQ/XSDERfQuR5e6OqPfHBp0Y9eZ0J53/mB/AwhDOECEEUQBPG4ThidjX0ai7hcAJ27XG8yZ2pIYftAzMxRUyCIx1a6XGntTy4cPpzt717Y1bXG83LvCewQiDdLagoE8dhKFxZT+5MLe6a1r2t5dncyN1GLNm4fiNMqPfEQSfaQDfbM3l4EuvvOLi6as/eejZtr8cLqbod+tcxItw+0eZFaArF7WKqgAH54deeLACRadmrarZ0zZNT9HXWF2xK2YaUWbcnPrz5EK+yqXAThM2IB6Nj++M/YyJteFxsZIYM9dOkiCAJwc10ArML2QzN4nuyMSudMLHghpDodr5ONSboIovHkOMDNxe0vcpgFAIkqCyg7nHmcSJGJSboIovGsGwLGTquujncIgrVK9iLpIohHAhsW0raCf9Qxy1QbJ8kUAJgSADMBJKCPMI4j7r80ktUBvQhYpksAcbC4AFimiyy9NTEyAUHsFTfXOWsWAJzswdpFoL8Tt69WXLX6JQBGiwTw1ND3gJEbOT5iutf+rzGcxrsSQHJI4OYMgCMDuDELPmJi7SJOWjj7yx+ivxPW22RrGnURROPIGcxaBGI9y9PJbsQ65xa80VM2m81ms8K7qMAB72yqwEln8tBcznSLKHobJ4aceZkFwAfy6jjQdnh+MWnBVW5RoalzbqHNJFPTqIsgGsc6b3UlkNm8DP6F2cTmdSwDAOvzflUAAIdZOaRLtjunlq6zw875v74mAQBr11XTN4SCsXSl7Q1TOtdyLGXmrr38yjmg6+51JHreIVuTdBFEw7BZ2laAtQIsfAlxF1gFAPe6AAJPAo40c8GJox5ClWyPvgzYEArONNZgyrt34eZe8X5i1ofA7J+SqUm6CKJxuLmejwFmvg+4DCpYd5kruUq+UnLiKACkZstjAryDjvWh7Fr2wEQu4pCcIOkiiEaQsxYBhtYsOMO6MXil2kWrXwJvyRmRbxaPdZbuBksIfx1MfQZAPC586UKGlqJJugiiway7OQA4DsCFnbPerPYabIHDKH2zf23jdHw6/JMPh5sqFANO9AbnHjP0AbDJ0iRdBNF45hTgwp10OweuAkC/AIBJPetzmPU/S8XHmeg7XsgFf/abcxunwl/beldWky/6f00BIJcmJF0EsQ9oXwPOJPs8ADBvE6l/cqijzKe9vfQqGK2Nj3UG0sWtG7PMly5RRPr+BF7Q0uViRZF9SboIYh8ID1V2rx8DAHccACD9C3JfaAoOqdKnUjnvfTWcL7qzYWTmL0ueRrpbHp1MeMYjExBErdKlnwqG5PP5fD4UIBnTWyEcDsREeJ0oFbSYANcnJQt4y15lx7kTJF0E0TByHQK8F8fFFsfhFZhe6nK5xQYBDAGmA1elGAC4zEI7ALBBtLsSKgGj2+uWwpVZVDmfmaAJI0HUzQofk4lnpo0xmf642u+2v85uq9MvPC+B5NhG6gYgj519XwKO7G/rlADW2s7GNxVWjw23yt8F8sbZ3/5w2hpTid8h96k06iKIhrN5haebp7FyL+1crfa7oxbhna08zQ5PKWA6kVqbBRZYiwLgXuedt5QC7Ktx5zKwfCezeRWALeMK92cSaYeUqyYqjpB9ggajdITslpYRO+UyBQDp7Q3ItbM9fcTwnlumeDiljxypqnOQ8I5LrlYMgcjZx0iYYAi8CCa0eRIpMIDpI43B0jpYCoh4HNyuvAlzF2b0Y9qruXc6QlYXvuHpYhdHyNKoi2gwxggA1r+9dPWPeCqcHPn6AS2m8j4KeQXAzavqVwRnH6Mg4QLQp8cUpA6+ChdwAS8CfSIz3FWE125P2xAAPrKtNLAR7VH65IjV0PvYyCAAY2jrah710usYeZ0mjMSjjsMsBM/LtsbbAZVGheP22CDZcA/kDQCtzrbmZsLzNQ0LFRJ3oq6RfOv2Pwt46XZWpttWbzWTdBENxpbWjn0JWPVaXpVjx3mabLjXO0VqpxeGPIetntvDsjldPXNvl+8wipOtQPXT5blJ0kU8YsjWXfQlmwsA/LAke9V/p2BWfoerFjmwD0clO7JzhzEhBwDDUY0vOEkX0WgWuO5L3HO4zrOpaoMFE9AvJyeEXphOdAkwgbggI+7pTsFFTlsv/CgztwEAaRva0lxo3/faD36NsFwiGL1lqw2jHO8O5VUzfB/9XQJMsDqrmaSLaPgUhltc5AA+MjzaA/CRviqLw95GdG8jZ78FGMNA21j/GRz5unH2TTLinu4UhiOB5Fj/mNBG3OJOYa4C4CMATnwNODHc/wZGrSNnrdoTzxtafzrG+karxOMoC0BnAQDGLKBpSOfw5Ouxs2/UVW7akko0fgpj/tSRQPvhudjxGSQPzRlmDgDvBABH+6uSXwBQOtHpXc91oaBa1CYZcS93CvPTNsCG1nPZgXc8IwJARgDAutQzO0uCme+/HBm19Cz//FUUvYegNdc1s2YBgPc6uXRfDgAMK1LN3h3KXAXC5zG967kuFFWLukvSRTxiU5gvpGyA9Sxd55aVSy1dZ4cBwOgDgE0tXYtfQpkLUW7+QOVx953/cI5MuKc7hfV0Hog1n1c//wttRABotwDgQ6nr5BWAt8iIdBn2Zazj0pc/9V49upnrnAWAZPN51faGKQEk+wDgvn+H+hzAzRufLq/m2xt/Wmc1k3QRjZ/CvGrlAd7yLpycmQvcsTvzQOiByuFC8RYZadPMVSiQ9fZ+p/hi7F0gbSusc/GrwIgrNgDvtA8A8ouAUfLshLlAvu7Ecy8BAFK2wpqyJIC1eSD0kZjnAHf2o5pJuoj9mMK8q0/xUi+G7tjtqbLVlxx3VKRNu8zKkfFquFN02xIwC4ADcS8w4vJy9KJ1A0iVPGB0eSO86qxz878DMPOAKz9TpZq5UKWvpjeqmmmZntiHKYwtAQ7vRMLFtqoP0B1plh305cg+k4xXw50CDgDxq2y2C2IrIzrMhLVaKipDogF1rSwAzFQAVLxKuspCan0/qpmki9iHKYwDQBh9fX0WsLZxurfqRemyg74wbYyQdtWiHnkAaO/r69vaiI6yyjaG2jPJ4fq1y815O4gVgMo3I7x1ekvuRzXThJHYt9HANID10B27MQgAtn92zuqXyg/6untlcPi7ZLc9oxQALEsAMjRihwUAS75c5V4x3FIJuSU6B9+uO/HFV8slZQAANn1PGvJz3q6zhlczSRexXx3KndfLIdodu/eGT7DxocAT5e+lLF0ZNCVZrsbB13ypEdNpALjvy4b8YqLsDQf3mnqtAcl67wIJQL9s3uql6/++yvVeerfB1UzSRezXoCuYPmh37PZ4SQt2WMZ/P0T4c40lRdJV653iM4EyeEa8tQQgfJZXMPxV+nBaN99T/0q9LS3AVQJ6+xbujgORo4xsno5IpmhcNZN0EfslXeVPsJzSlS1Hdur2rX7XdyQBhfq2SD65yM+FKuZpVaHsTmLpXaIwpX5r2g3VpHZyf+wC8ijAzferVTO6dUakmUNPmENWZ7q0TE/sl3SpXqDV0u7Yq3U2oRu5TIpkNxDrBDcVwEWcrLdn8oYJdjwwYuXoSGk/RK7qwXEBdJhodRVUok7xWmEAFmMm2t1qQylX+ktdKoM2M6jmIhf1pUujrgNBvPjQkmZ6FLTnqYU7MzDCUjd8d+xVpMs/9CvPR/hyGqy/nTs5OGyscJGqfK9sboyu8kNz2ojVzP2Slq5c96g5+6/AR1cTm8Bq99mP36trxqiagXU2KlN3qzaRnD85XLTGEgsvBdXMR5/5Po26HgNY3zb3IHZWPKSUgZPdAICmvbwrq1wAWF5Np9Zzvjv2StZd3cU2c4nN9wD7TvrwHGAv8paHVAnJfTqAJzYI7OA2FsBxs5403Msb6cPzvhHLfgP0jhWlgDmVuqUUljfSzVeANRWXtScKwF10AfsKS61fqXrNqq28KeKaTKzNhNUs9/F+/KT7pveeiCW2b1CBY7x0XQ0P/C1zG/+gf6mv6QZ8h98N80bOt/fD/7K3LNL0Dd8yu1BRJnTapS7Wo1f4qyxM6EwKgHk23DnLvq92zyt6tfijSfnR7XAdgBd2cMW+ta/27UXv6W8Cb1k73TL+0tqpvFuVMFpQVq0hxgNza0Ow0DA8VUN5NSL8N5HGdtXsV1e0mrevC/JNX4eajJgA+ru3vahjZMDr0iNDdSeYyVT/vl8U2XD4Z9NQkLmtaZBX8OcshSMmjDdc19i9F3nXmzm4+S1drAc+210Fz7W7Aty8BLZw9F69z45YADoGtr2obcS7IxhjQ/vXVjJde29eWbNBiespvavd5FcZRBUDc+sKcUM7O6u1p6zCfwt5bFfNXgZRUs11dk9SqK0n8TABbuZ2uO94N48qqmMM7zHBolsyxQiHDYNuRya8764bJmC4cod8edJVxSt4x/ZiVpLtVB/+ecAEbzvGhh+9xXNHWtjRRWhcO1A3Kp9p8ZHt49+933bl7tHKgAtFfYyka19wc107SwQc7WS98t7FUntM8O6FynkoAHeiKf76TVUqqWl7h4a/tVfwnSZY0Z9v/WP3qaVZbE6fij1/5dGrItkKlG/WrlJDAvD93VWfS2416xa7zcjtt/doZcA5J6mPkXTtDzm+C4mwubmFROxdK7dIav1KfGkmKqmdgLWT35DGeAV3JsQzVwD31k8w8Qj2tAVjFyd42MoEgH3d7OrWYOYi9TCSrn1i3RCeRGznZN1VJgDeIv27dNTJujfHSiGT0q7AwVJIdAE8KyIxMi9+z3F3NoXAQbvnZzw5WOyILgPnDHBzEYE/8OreyKt5Bc90iYg3cpbRAbn2au4nHQ8i5cPqn4cAdvIYhs1Hr4YcZnkneHjG9Q1ZPjaz4M/8uQkgrU0X1pBXdADQNZHpQuC3vWqkXs2xFDImdLwp344RK/OysKEbfoE0wIVOEJnOaEbS8NzHI0W9kKSrBmxY3Fzcwck6cin4Xtw6XgeMUSA20j8mmt7knpN1PmoMDwH9w/1vAHzYGOkfYCN9ZwAYY16MJ4ZHLHjr7x1jfSOW76DdczfOhu4X3zspSiSVuxKIjfSNmjq1yl5dxSt4x3D/EE52N519HQBY/3D/GQGAe+7MvaR1trWn8xPPz15v70as5/rmPz6C5yPa0vScRHvGRf/waJU1e5mAdh+GjgHvCYcx0jdqGmcwOuYtTOmKwJHh0W7AGB3ufwPabzs7XSVSY6xvtAeIDbYNDwBAxyD4KB/rH0KJlZNjZWE7vq7P2O23+IiJjq8d8aq+Y/h0DwAkx/pGerzHDyfe0JcSJF17X+xKcVcC7YfncseB5KG5nHeLzWaz2WxntGOUenE79ez1nwu36Bb1NOK4CzR1zi20mQCGHdnedmjVsIC+5qk7fYDRU7hzCgAD2OD6/MYpFJRbLIINrV+/NwA3f4UtzasSSW21FTD07PU7gzo1AGjNZrPZrFbXKl7BWe/KdQ5X6Xw1WSsLhgWgo3nV6PaT9rKtk8bqFJ65KuHc/Yk78ShOcGTaELnAuLHOuQVv1JTNZrPZrD93NgCUvnzc3jx1R6CIYhEIrA7wgbw6Dhw/NJ9LChRVsQi0HZtb7Ci/NQy3zMsTAmCDEaucaF5NWiVWzjbPy/aS/HKAM1M7uAKMgfzGKYD3OqvHAbAhZ37juFd3JtuH08ceI2g3/baLXZ9L2GonJ+sFDkC/H6YlwpzKLTTff/s/+t63X7u0ga6715HoeQf8+Qv//M2ByZ+MWjluvZ/LfcuUrc74xqieji1fxt//hXbQrt2NYxJxd6pUUp8uAMbhS/In/1FsmFO5hWbA9wr+G73qVsUrePMPkMe1l1/x8uVOLiLROQu8NH5ntGtWJ+1l20/6449f/s1twH67idnvPII1tPAqdyR84yY2r3sejVkfAGA+mDsrpEoWB61b84uHNi9+65KKWB0wlq60vWHKwrjkljXr+W3vWbrOUt3vlQ66zPdzxre634Px7HfDefSr4z8fs3JRK9vnlfGt6CJbgQMpCDD9JkFs6cqR1wGj5dKd0xYQa7kkZ79p5dTvAoL7JyYRNOra+2JXqgDwlhk4ORMRJ+vz8/PzvljZ3AQrWQNmQqFkeXxNFpmVB2bjAOaUjWdyrkzAcHKwpYmUrdxprUqXgXWu7/Ha3TjAyl73yiXMRW99epOJMLX1+fn5+flZf6zheQWP5stF1Cv4/UUgxwAsS3eWb5F0/a/n7u9iF++y9f6I2TjiLrAKAO78/Pz8vL+r3IEJlHoVFgoR0wRFd6axBhNLEo7U5ebmItxc2ZpTq52DnUsDmEMYzbJ0cvESK08p2EqUrUAKZerpK+BMY4UJpG3pXgeQtiU2pQUZB2/hAiYt5dOoq7alFNb9/s5O1pUpDUc+FRkVlT/NWgCYyOkB2ixc+RtAfsG7pyoBMw+sR4NrtLtxBfd/K5PUWIsE4jbgSvNOkNraDl7BUemNXAHAos5ZtaRx0/vu/tuPZg3J7huhcZV/Jy4ZpMLNWTkuSqpEpXLlqwMAYCs90QbUp7TNXKlPXosQdwDkPge40WgWAcUqrFwiXS6EMuWLwUzQVt5XBWBTeWaH/AJWOZjzK6EESRdJV42LXcckwJEFBLB4rHOx6npLarbs4IDcCSnLBIIhExnjKgAMgmUB8SmID6NDo3iw9VP8Klt91GPDVoDg2aqpBZJqzZZ6Bbc3hi6X9CqWatLrYC7bKmk3KnKP4GKXKUPjrhuDVTefyaMwHNUU/ea10rqMFF19BtEdXczNArGy3aymC6DIKjdFqHIrl20Nc5R5R8x8Dim3RDN9iRJFAIrB4eCONKV5gzohSVdtHeNzChBGH4DfYm3jdHy66opY2UIK5q3Ra2V3dYY+hEc8+UsmfQB+yaJ3aSO6B6G9fStJ/QwAEesD8MtqqQHaK/iHJd/MDJyJ7s7q6BW+01IXQu0q6UeNxVdlaFw7Z71ZzenEqlG2So85q1/MVLe6YmD9nYBvOt6HiG9Xjbe6X71yIlbmfZ3AL0ul9r80FxhE9TcAxC89+XOY+T9d9ftlQ0WCpGvXBAsS2zlZL/DyvfRrl0/1y/JhyhQiviP1MGgGJVNFgA0fmts4FS6eANi69W7OAlgJU2vrBoC7s4Hulm80X5ruPRV27djg+lTyc7Ul/egMjN2Icd1Jt3PgKgDvlU93IliOFGVbeO3LQydWZfWiA22dK6uBdDvTlRVXNrPHVlZut+bU8bIBIHdcDvOX25dKiadX5RdYnRuKSbqedP3awcm6w9PlA/tl9aY1W9bMVyqboXeoanQ1hJvvSh50It/d+Bb50sGD1LyX9/9FMNao8Aru3tp4PRziZewL+PTn/GmRMvaQ9KOmX75xnUn2ecD3ShD6kkbKryF/pLR24Ux4UE2J1UURPR9PIPmpYAxdJcU4tnbzGVqZ9SxdR+nwdZUx12GpLWaCCgCEC1eZ/1MWGO2NIOmqix2drKu0/7Qq0KC1nFneuyp99SkdcVS6oo8EQ3fj1ds5K0ttaRdewSNvLJlr4a/c3VPSj5x0+cZ1rx8DAGe8bO7cVTH1snOvVLe6+UvmeSrWYdlWpje3FBbfyky8X9GYeKLgylTZm7HqRU8J1e960gX5yj996PAEObveGtocsQvp4qLsi3w+n8+vhh3DPwpV8RIXEqxkCmBWRKwf60GlgWRpMM85XOu2+SqWV56XLxmONUKv4EBPFcnTgtYTbKll/j87JP2oSVeFcfP5fD4frifJVv0qquKBqRXgRp6BMACICfjrS0L/4zCrMkFpADvsXPCtzERZYzIVZPl+LcmBpPC2NzNrFZD8sHS5WaDeR9JVj3Tt6GQ9rptYMSaMHsDoATMlXGayyI0/C/BSh3Q2HwDaTUhD8AH/Hm/iOOA5aNfuxreiYFhAu9CpVenQ5V7BmyxwSIB7/VK1gncDQJvJuwtB0i4z2U5JP2rSpY17XGyxiXOROf7dwuQDAB8Es4rwtqxHrM4G0e5KV6XQZgJQCQhHdgGt3aXxrfMeJK2qz5tLrWyiXZSprCWh0mW5XDVM3gsgb1hIihxQMBw4ypLU+2jCWE/H2MnJeh56BLbOxu5tAjiZ5C05OGq0eD64aNoaU4nfmSnRxJw1wlLvY5mP3XtGL9Zs9N85BngO2rW78a3ytSlPd7DUD3RqVajwCn66I2Er5I2zv/0hgHz3sOEAcP92dEO8HyatRovnd0j6UUMb1xiT6Y+r3n2gR2A2G914RgEdYxvmu3DkySPnIlbHWtvZ+KZC7rX+tIT22z7z9TGV/rBstilPJBN2VaOXWDl3oqlTlg8QFVZZ2YBtk43eeT4H2BunO9KbEnCYC0iSLhp11ce2TtaLgOM/irNn+LNXXTiLaXNJwZ1hzeF192cSaWcmXCJzAUzKdGpZwp7hzZe9r5zphDWptIN27W58SyY20uay0qlVodwr+OaddMsUYMu48lbIMs9edQHc2hBLuSBpd4Y175h0A6jPFXsp2rgr99LO1YqZG3QNuS7gXGXNlwFnJpFaksACawFCq8O+GncuA3MbXXdz0H7b7y4m0utlKuVObaSbp4Lnm6UyFLHy/L2uZenVtV9D0pFwUObdzbnCDk0BcC5vpJ3LABxVAKTbOOkyTj00iel78Gk+6b7pgxWPHZ2sh66/PSfr/pXak53vWCUVXBZeH1zI9FcJM+I4vNQXnvZNjzJv5Ds4f/cu1F7BBRA6E9S5Fb57cp00qiQdsYxokI3/0oIxVGvLFKWrdb4NqmWalVjcN7W2mb9j1Cs686uMe97btWFC+4a+2nVKIpJIXH9GrazdwIvIO/DCbxHx6B8JoVPUJWDRZbKn6zd67NtAR/deQ+3km35XovlWTeeV7OibniaM2+PdK7Ub7UKVVVNXj7y0C27P+7Z/ZT4YmIXBFaIuu/P+lDO4MxcK2nF4JOHqqe6cryCtwCu499tqGFwByndPrpPGtkk3EiPTqBrSNqiWabfE4r6ptc18y3lFd/0qczzv7dowVe2bj6YOV3nV7JZZWZtURUbbym8RxegfBX2FfwCri9r8F25H1+/NPqRulOmaoQkj8TitIobzrT078if2rvLhCtuOnvMbnHJY0W2NcvtGoy7i4bEZOuNn5A9037kWnm+ZeLDmvh2+PMBNGnURB3/URY4RHqi51eOUMkkX0XgSgncBSHcJeC7bAeb5cGclrti9lewSj/gNO5zwyUE7yPdMy0xkUkDCOxcyOPvABLS3+3SXCLzuI1PtVIM9UHbqQmsqOOegLGXTr1smvJTTdaZME0ZiP+i/fcTMq45B2BcUH/3OmLDePm055xXYaQu3rwIdgxAAYsPfQXJI4OZM05s4a38XsRHhXpJkwL0NP4Z/OHz/+9q0TYP/edC99n+NCOttwDgj3Gs5sFHTuQrgxO//kPV3wjmvTlo4+8sf4mQP1i7Wk/Lod4ft76K/E9bbiA3+H2/ixn8ZMdcuBinjtOVeBtDxle/jRC/cS/JkN87e/z68tkGjLuJRo+2wAu9dv250AzjhyGTbsTzvDvy8+67YAWDImZdZ7RHf97ZP7I1eNzAtY735jd7jjkyK4OyDNjN/rwcABIzOlQXerT3nx3qWp5PddaUcPXWBncqr4+2HVpNmkHKTVfDr0+hdmWc9KCq3uBq2DZIu4tGi8+Z5lWyemHsvA+DV8Qk2eGs8lwZ6lq5PbnTDaBkf/3s9YVgbn7piiPtvu+cuwDg8MXclJsh+e+zEh86945sWsfvjM7FXxyeYCW5Nzk8YJjo3xy8k9LU3JyZn07iWu3/uPWQ2L9+a7aor6dcuXUTX3euTsgcwPjt+hb8+PqGsIOUu+/z4Pe9Kd2li6r1W3H5v89xVBG2DpIt4tLBnHaTWFFY4gGVly2dnkGOBn3ffFTsAXFaw/Z2d2ts+2W+PLKti4ELfnUaeLStbmf7ZB8ych/2ed+XmDLDo7461VoCF+hSg5NSF61hna9KRIjh1wVyD/3qDcwXI+ykHbYOki3i0WId3jprLheeyfQ0oMu3nncF3xa4JNv1ob/tkvz2yiMC0sCVcLAIS/tkHTMhgdzQQuvFhptzCp8/uiZ664OTgYgFQ8SBlU25x6oLfNuqAlumJfaAIwFzLBm+/qM8AbujnvfS0iITwR13a2z6xR2SpC31PIYR/9gGLbuYHSxn+/9CaBa9PulTpqQv+SwtByjJyLUsFA6Vo2yDpIh4tGKr4zKni550Ph7sjfW/7RA3zpyou9A1tznCAi/ZTIjwh4Ti29FK9l4re8tSFkq+SQ8J3g1m1bZB0EY/MCowEXFWyJFHFz3u/GXELr73tE3unmgt97+yD6PCGn7KnjFf9v+ZUA6Rrt6cuDDXP3esubRskXcQjiIv1Kk5zKr7i1o1Z5ktX4G2f2DtVTOeZ04j4Fk82n1dNgXStyIZU9JanLkRvW7GWcyrWvV3b2Oswk6qc2B/846Ojrdx3xp8KnMBwdzb8XTEyW60CUsV0yl/6EoG547aKaI5ozD1KbJ2yGXyViKRcrW2QdBGPjHRVvuHr+3n3XbH7cwn/nyI1x5rni1Vc6AdHmptAb/h1XGtLg57lbnPqgitNRE9ECo8+aMDb39RWiH1iNWaClT4x9P28+67YAbjMQju0R3ztbZ9sV4N0VXGhr88+cHNZJK2gw3sHJHAI5DoEeG/d0rXdqQtJYbzufVXkgg16n6Ja29gztNZF7BNr90Zk4tnSNQ3t5913xQ7Akf1tndL3iK+97RN7p4oLff/sg0VrjC+nAQArxrDRDHie81f4mEw8M11vytucurDQPcLXklrNhu8dAuDw0We+X61t0KiLePgoBcCZvpduKWud2s+774rddeFe5523lO/I3/O2f/ANwPpEAy/bcdwTMa33wFB5daDPPrgrE5vveV/bi5lD09Ce8zev8HRz3crlHwzgRj3w65Q3FxP2FS9L9kzr4cteysWqbaORkG/6R4zQN/3DtMwuepteE06kBSIHRQsg6hM+cMaf2tkjfkXLfEiTyt36aue7az+7vGyn8uofw7MPApPrsw/8owkQnp7guXZM7GDuncpb9dQF7zCAdHBYgNBf+Smnd5My+aYnHgZ6r5Dn072IqGv40Cd8mTP+B+cR//Eb5EbNraIm993k5wP3+P7pCe6WXvf3QtVTF7zDAEpTVl4Fh78U6j4dlyaMBEEcQEi6CIIg6SIIgiDpIogDR0a7hfecryeyKQAJ4Tlp5wLaOX/om53c8dcGLdMTRCM5MoDOiwAfE2sXgY5BuNdy6PjMP/W4lyQ6/s3fDLjXcoj4ZjfIHT+NugjioWMMrK8mTaCjeTVpgQ2uz2+cAhDrzrNeAMZAfuMUor7Z25un7giyG426COKh0upMqJMAXhr/+ZiVw/Jl/P1fAMjcnHn5SwBiS1eOvO45cPjVvwUAWLfmFw+R3WjURRAPlc41hVsSWJZOLg73MrDOBWDPeF7gnWmsMBH1zS4UHJovknQRxMPFXPX2ZS5GPL8A2AQcBsBWcCGivtlVioxGE0aCeMiw0Kuft8mdxePBF95vLkp8s8vXOhfJbiRdBPFIYQyb1QQu4pt9zuoXM2Qoki6CeJQGYcOHIp73I4R+++3LQydWJZmKpIsgHh24+a7kldJV4pt97cIZi6SrBtuSCQiiYZR5O+ZbPDws8c1u52ilnqSLIB4uMg2UHutc7aiQUt/sisxG0kUQD1m6DJP1ByMvl5lVT0sNfbPzQTCrSHbbO7TWRRANZJmP3jHf9f9yNvrvHKtyVcQ3e8fYRng9QaMugngo2NPMXJKBdE0nrEnlzQndyDnToW92ZyaRCq8naNRFEA+HWystq3DOFYGP3waWzjfLn7i4eQNwzinc/iXgnisCS+uioOCcK+LmcssqWY2kiyAeNoUCtHd4Bd/zvu+b31X6t4jf/vq9tNOEkSAIgqSLIAiCpIsgCIKkiyAIki6CIAiSLoIgCJIugiBIugiCIEi6CIIgSLoIgiDpIgiCIOkiCIIg6SIIgqSLIAiCpIsgCIKkiyAIgqSLIAiSLoIgCJIugiAIki6CIEi6CIIgSLoIgiBIugiCIOkiCIIg6SIIgiDpIgiCpIsgCIKkiyAIgqSLIAiCpIsgiMeDGJmgOu3tj16eGJJ9Dz8Pxx++IYyHlQlh9FF5HwytjKSrRsuZj2KuMo9AHrJ4kjORpfI+IByaMBIEQRPGJ4bb1xoTz1Nfu6QalSd+5mauziiOfOlinXOXN+ssDxuruxB1GvXk0XdqDvq5tx9C0ANZXoCfuSbraakkXbVRLDYmnqdRbJx0wa03Vwp1xhCrtzys/kLUa9Q6MvBwgh7I8oLX1YnUztETBEEcOEi6CIIg6SIIgiDpIgiCIOkiCIKkiyAIgqSLIAiCpIsgCJIugiAIki6CIAiSLoIgSLoIgiBIugiCIEi6CIIgSLpqI5ECgLRoTGxePOkGRMLri4SnACBhPsQs1GmMOrNQa+L1tQeWFgBYqpYMmwBYzeX1qrqWfNfbA7w881R9jYWka880DQHgI2ZDIuMjbwLgI/UJ4YnXASQH64qjf8QEcMJ6iFmo0xj1ZaHmxOtrD7GRHgCxoRqCdgwCiA3Xqh9jwwBYLfmutwc8N9INoG2wzq5DUrRX1g0TMFzZkMgYjAYM32QrgJRTZ1ash52FOo1RXxZqTry+9sBqP28gxwGk7drLaz6kHiDQRRPGh4ENE0jbqjGxuRtW/ZEsGACz8vVFojIPPQv1GaPOLNSaeJ3twa1ZrtcNE7AKtVe49bB6gGuQdD0M3FxnXS2mfKzQgDuQwyxwkasvknxdQ56GZKE+Y9SbhRoTr7M9OHUpCDcXay5vLvOweoDNG3C/JumqocoN3WISWRMAEl319PnF4A7Es7UuXNrSRKsjAWS6BACWqSGmZ5k/gWCZrjqykNZZ6KrLGIm9GyPMgq6RzN4qptaaCNqDDra39uCGipneW37dXKeetqVrKm8huFftteXV2wNc2en/N1Nz56FjNWoZqIt/diUQGxHuJYnYiHDO1z54Xuemt2hgnBG4OVPjgOGLSNkATvZg7SLQ34nbV/d8+5ed3oiFj5jovFhrFjoGYV9QONED6+3ajZF8E+6tmRqzkHwTznmFjkGsXXwANeG3B+OMcK/l9tweFl/S/+kYhDMu96Ign/embdrkey2vzazZ2lpe3T0g94f6Pyd7YF+orfPQqKuWgbrVaitg6NnrdwaBU89e/3kdw65gwjDcMi9P1BjRAoeVB2I9y9PJbsQ65xZqGHblWr3PdnNlNdldYxZ47/p1oxu8Z3k6Xrsx2JAzv3G8xiywofXr9wbAeleu8wdRE3576GueutO39/awzr3RrjG4vsB79yb0wiwEJt97eWWqxpZXdw/Ic+/6pp71BWOAJowPCjeXsgqAcXhi7kpMMHNqbrz2QRdz9ZKDYV6bmmDdtcXi8LTIAZnNy7dmu5DYvD55de+RrHgTCNZze2J87+s+OgvJ5om59zIw7Mu3rtZujFjL+NTFPS+I6CzEmifmribBmyfmLj+ImtDtgVuT8xOGuef24CtmuzMxeTW5JwVhKXMxMPleyxvcq/be8uruAbZ+nN21eWFyNknS9eAWuxLmItDqSGwywYSCo+qIbcFTjFY7BztX4wZDW2YdCVgrwAJH3AVWaxlNdgMAF4twZ4was5BaU1jhYC6Qr90YaVtiU1q1ZSFtK6xzwVyg8EBqwmsPhpODLc09twdfMa01heVgvXF3IbtaSk2+t/KueNsjamh59fYAJ9cFAMyaB+a5SdL1wBa7Yi0SiNuAK00XZn2x2V5zjTsAcqzGSGTaBpgpAZdB1Varrtd8uSOBAqstCzBXAZcLl4t6jGEWAMhEjVkoAA4E9pyFWmvCaw8JB4ASe28PnmIycxFw1J4C55K2CkxeQ3m98V4NLa/uHuA9EmFCAjYESdeDW+yyFSB4NpsF3NyJ+rTL8R63mC6AYq3StYAFgKE1m+1kWDdq26jstaeEC8BloqYswDSy2U7A4UOiDmOIIgDFasuC+FU22wVhb+w1C7XWhG4PLJvNCrH39uDNnhgUALWnLK+jEJp87+XV96oaWl7dPcBb4ONQgKtIuh7gYtc6ABHr6+szgXk+atUVnV5yKAJArdLlQALA8b6+XsDOdbxZ02jSG7q7ALD3W6EDCTAc7+s7BdgzyWFRuzEEsIuj27eyQntfXx/gziTPmA+kJrz2AKPG9uAtl3vp7XHehcXQ5DWUd9WoseXV3QP0Ap+rampqAGhzRD1szgJYwdrlU/2ynsWulcheULe+LM0pwIU76XYO1LBObyvrf9ZtlWUJuAq3ROfg23Uao54sQGJpuvfUxQdYE/YMgPUa2kOwVWDP99ASk++9vOuRJyHuA+0Bbu6LHwR/1NZ5SLpqRjnzuuWoN/UGmRrnn3rJoY5BV9D3pHdPm2Sfr2nI88X/WuP9P+wA655V3GvqtdqN4d2Ma9VxW2fh1sbrQj24mqi5PawYZjDWVXWYfO/llXW1vHp6wMJrAMCEqnXQRRPGOqQrqO21nFlPRN7eYsUAmHW9v+wGrcC9XlPFLhj+qgev9aGpDBrifE0r9doY8TqkK1w7WWLmg6uJ2tuDrUy/8mqQrojJ91pe5BL1tLx6eoDN/DJzb5ZP0vXgKDbMdosGAGkAMIv1aaBZX0ZsZgION4F0rRoqU3XOfT1jJABmrdaYhdaHUhOFmtuDm+vUsmkIWY/J90y+rpZXTw9wpAU4sADu1naXJOmqmYJhAe3C6PE2JdTBOgOwznuQtBbriijXIcB7cVx4T+r33p5yArBVL3h3rS/XrsZMsCw6TLTW1iTXGYC8YSFZ65vUecMEO44ma+839HpqwuYDQLtZS3tYZQBkUiDryJpNXkt57bpaXl09ICcAN9cu2KkaXVDQWlfNbMrTHSz1A5xM8pZcXTHZ0gRseSKZsOuLaIWPycQz08aYTH9c26jDAtxc9xhvuVRjFtbujcjEs/N8dDWxWYcxNk53pDdrvCFsboyu8kNz7HRHYq+9op6acHLWCEu9/481tId1AJjvHruTul27yWspr5Oz6mh5dfWAFQBYsEY2Ujdqq2YaddXOxEbaXFbOYtpcUjVH4gLAoqPgTm2km6dqj0YB2LzC083TWLmXdvb+gFEBWFcKuCUTLbdkjVlwpu+lW+axvJFuvlK7MZzLG2nnco1ZcC9vpA/PY/NOumXqQdbEpEynlmUt7cHOucDmNE+t1fL6vTb5HsvrzecX3HpaXo09QHl3iSKwNpNILc82vms+/dYTq0qxt7ZdNdLLk55fce2qe0sjim1TEuG/bCdv3/yt7i1/0+86JzxX51vG9MK3ti8T8zJibmMZUV8W2DaFiBiDb2eMLY0aj1bMVln48n+suSa+/I3t20Nq6/awRdDo87UtG9KW5WW7MfkW5Y28Hb/X8u6uB/CtOhHbTZm3bqk0YawDvQLtem/pFepxO6jCf916fIzqVVYvL7XE5Ib/1piRurMQNYaTryMLXuJ7zEIdNVFPe4gMV/bekNw6TB5Zl997ZdXVA9y6ykwTRoIgDjAkXQRBkHQRBEGQdBEEQZB0EQRB0kUQBEHSRRAEQdJFEARJF0EQBEkXQRAESRdBECRdBEEQJF0EQRAkXQRBEAB2cnqTflLNYiDdIFFvQlo17j6TqLdGTJau2zL1lYfVX4g6jSp4+mAFPZDlBXhdnchkOzakrXn6r0jaCYJ4KDjfoQkjQRBP2IRx/kk1C+taadA0z7DmH6VctR5arDcP9ZYnW79p6zNq6/O5h2C+eix/EMtbb3NtPVSXdE3hCSXWNScbE9PT1pxqVK5419psnVG88PxU3ZaprzwsW3ch6jTql3+3Zht8+fceRtADWV6A19WJXnh+p+hp4EkQxMGDpIsgCJIugiAIki6CIAiSLoIgSLoIgiBIugiCIEi6CIIg6SIIgiDpIgiCIOkiCIKkiyAIgqSLIAiCpIsgCIKkiyAIki6CIAiSLoIgCJIugiBIugiCIEi6CIIgSLoIgiDpIgiCIOkiCIIg6SIIgqSLIAiCpIsgCIKkiyAIgqSLIAiSLoIgCJIugiAIki6CIEi6CIIgSLoIgiBIugiCIOkiCIIg6SIIgiDpIgiCpIsgCIKkiyAIgqSLIAiCpIsgCJIugiAIki6CIAiSLoIgSLoIgiBIugiCIEi6CIIg6SIIgiDpIgiCIOkiCIKki0xAEARJF0EQBEkXQRAESRdBECRdBEEQJF0EQRAkXQRBkHQRBEGQdBEEQdRHbNtf+55UszAcb29MTAaONzBXyXprpJX11W2ZeotRdyHqNKow+g5W0ANZ3no7USvbMfqtefqvSNoJgngoON+hCSNBEE/YhHH8STWLMXhLNiamptcnVaNyxYeWcnVG0fHS5XotU2d52HDdhajTqCc+e7XmoM9feQhBD2R5AT5UTyfqeKku6co/uYqeb5B0PY1846QLhXprpOnz+botU6d01V+IOo3a8bs1Z6Dj9x5G0ANZXoDX1YmaPr9j9ARBEAcOki6CIEi6CIIgSLoIgiBIugiCIOkiCIIg6SIIgiDpIgiCpIsgCIKkiyAIgqSLIAiSLoIgCJIugiAIki6CIAiSLoIgSLoIgiBIugiCIEi6CIIg6SIIgiDpIgiCIOkiCIKkiyAIgqSLIAiCpIsgCJIugiAIki6CIAiSLoIgCJIugiBIugiCIEi6CIIgSLoIgiDpIgiCIOkiCIIg6SIIgqSLIAiCpIsgCIKkiyAIki6CIAiSLoIgCJIugiAIki6CIEi6CIIgSLoIgiBIugiCIOkiCIIg6SIIgiDpIgiCpIsgCIKkiyAIgqSLIAiSLjIBQRAkXQRBECRdBEEQJF0EQZB0EQRBkHQRBEGQdBEEQdJFEARB0kUQBFEfsW1/7XtSzcJwvL0xMRk43sBcJeutkVbWV7dl6i1G3YWo06jC6DtYQQ9keevtRK1sx+i35um/ImknCOKh4HyHJowEQTxhE8ZzT6pZjDevycbE9NTXLqlG5YqfuZmrM4ojX7pYr2XqLA8bq7sQdRr15NF3ag76ubcfQtADWV6An6mnEx35Ul3SVXxyFb3YoLI/jWLjpAtuvblS9dZprN7ysPoLUa9R68jAwwl6IMsLXlcnUjtHTxAEceAg6SIIgqSLIAiCpIsgCIKkiyAIki6CIAiSLoIgCJIugiBIugiCIEi6CIIgSLoIgiDpIgiCIOkiCIIg6SIIgiDpIgiCpIsgCIKkiyAIgqSLIAiSLoIgCJIugiAIki6CIEi6CIIgSLoIgiBIugiCIOkiCIIg6SIIgiDpIgiCIOkiCIKkiyAIgqSLIAiCpIsgCJIugiAIki6CIAiSLoIgSLoIgiBIugiCIEi6CIIg6SIIgiDpIgiCIOkiCIIg6SIIgqSLIAiCpIsgCIKkiyAIki6CIAiSLoIgCJIugiBIugiCIEi6CIIgSLoIgiDpIhMQBEHSRRAEQdJFEARB0kUQBEkXQRAESRdBEARJF0EQJF0EQRAkXQRBEPXBtvnt6b8i+xAE8VBwvkOjLoIgaMJIEATxaPP0W09s0WNvmQ0zomjcfeat7nqjeOFbdVumzvKw+gtRp1G//B9rD/qNhxH0QJYX4HV1oh1bKo26CIKgCSNBEARJF0EQBEkXQRAkXQRBECRdBEEQJF0EQZB0EQRBkHQRBEGQdBEEQdJFEARB0kUQBEHSRRAEQdJFEARJF0EQBEkXQRAESRdBECRdBEEQJF0EQRAkXQRBkHQRBEGQdBEEQZB0EQRB0kUQBEHSRRAEQdJFEARB0kUQBEkXQRAESRdBEARJF0EQJF0EQRAkXQRBECRdBEGQdBEEQZB0EQRBkHQRBEHSRRAEQdJFEARB0kUQBEHSRRAESRdBEARJF0EQBEkXQRAkXQRBECRdBEEQJF0EQZB0EQRBkHQRBEGQdBEEQdJFJiAIgqSLIAiCpIsgCIKkiyAIki6CIAiSLoIgCJIugiBIugiCIEi6CIIgSLoIgnjiiG37a/xJNYuB+EZjYnoKcda4+wyrt0ZEvXVq1FseVn8h6jVqHRl4OEEPZHnB6+pEYueGtDVP/xVJO0EQDwXnOzRhJAjisYOkiyCIx4yn33piix57y2yYEUXDcsXf6q43ihe+Vbdl6iwPq78QdRr1y/+x9qDfeBhBD2R5AV5XJ9qxpdKoiyAImjASBEGQdBEEQZB0EQRB0kUQBEHSRRAEQdJFEARJF0EQBEkXQRAESRdBECRdBEEQJF0EQRAkXQRBECRdBEGQdBEEQZB0EQRBkHQRBEHSRRAEQdJFEARB0kUQBEkXQRAESRdBEARJF0EQJF0EQRAkXQRBECRdBEEQJF0EQZB0EQRBkHQRBEGQdBEEQdJFEARB0kUQBEHSRRAESRdBEARJF0EQBEkXQRAkXQRBECRdBEEQJF0EQRAkXQRBkHQRBEGQdBEEQZB0EQRB0kUQBEHSRRAEQdJFEARJF0EQBEkXQRAESRdBECRdZAKCIEi6CIIgSLoIgiBIugiCIOkiCIIg6SIIgiDpIgiCpIsgCIKkiyAIgqSLIIgnjti2v6afVLMYSDdI1JuQVo27zyTqrRGTpeu2TH3lYfUXok6jCp4+WEEPZHkBXlcnMtmODWlrnv4rknaCIB4KzndowkgQxBM1YbTnn1izMOGqRsVUfJRyFUfxYZcnXr9p68tEHRmow3z1WP4glrfe5hp3SZ0JgiAIgiAIgiAIgiAI4lFBkAkIgijj0d8cYZzZMu/DDU2odaS3/kj6LWO3uTpubflTxwAbMWs019BukjYb0nYaaP/2kZ7SL1hf1VtW8HX/TkUwdl+bJ6xdFbcvWs9mg6qisX2lF08OB2BflwsAiSpNhWUamk5fOlt/JGnBUruVSnPLzp8wa36RYVfbn7PVR7Id3XtKimUqgmwhODt2dT6U7iyLvHoeg6+7dhqMG527zkh6V0rOos1jx+R3WxW7IDa4a+nq3LkU5TX04AX2CZEu9qYL/gbQMbDvljBvXak/FqX2dJfMVO806uEYu6ur3iA8W71AvGv7gY3hjs88mPbeZR24Pppp5C26ooaMDA4mj7p08WS325+o/Hqk3vtWxYyMu9MrAGDUOQ3ay046N9yzV1oit5b9ePXkvG0QgNr7/sOyIG4V1e03AXcHSU/Y+cUHM4SPZGSPjWgvc8Xd3JR3vSRQbOTmzGgNGcO1NrVHgdgjnj/7ygBeGq+s97TYy8ik2tA9fassSkd/ph5c6TbPq+olUrUMvOrJOTcBXGvZc7iyIM65SvVL35dwL22viuJB9Z9oRvbYiKLlaQTlDXBL7l5ooAGiNcRSADYv4GDyyK91Lf0tm5QHx55qb8Ou4rYxPWhqeG+jPEhx7wXd81i1LooHr4+6aj8N4B5AixwM6ep4ze3XyxMJwfREnQnE4wCQ6NIT+GDAkfAWUNNdwguRMb3Lhf9bQgAshbiOwb/Wv8aLmpkAEgBLIWMiBe7FyjJ+DOACYJkUAGZGl5SVCyAheFeQlwR4lwDTy1oZLyad8RQAnkmFJfJKovxRF0sh0QXwrE4+FeZC/xRGGRil1fs64a9qeIXw0gGQ6QQQlNDLJRMsLsCFn51I7GFUXGgzgGX837gozWMa4ALRfMbB4n5KWc9eOguRawSLC4SXlCXsf6vLX7pqo3/zr/RyyIMaYaYuiJcSEkEGAnsB0A3LK3s8Ho+LSIDWTgTV55fHT8Kvac9ApmcBsEy6pCq88U1CX+JZL2yAXp5ZJhVUlVcm35DMBBJ+kRJdqbhASYH0L9Eys6wIatJEJmqAsIZ0U09FelYi68f+6GM86vl7c+65//qvfoTPPv9j/JHz0isv/vQTALH/FUe7/o6/8uM/O/rCImCMvfjir3/hXf9nR7v+SaFjONP1Twovfb75K0c/Ui/8LzxrfYDkmUzXTz956Y8/QNPXb/x7ZLIfKQD62k9/nWdf+jHw1L9jWetHT3/tA+DP7V80fW35T4/cPPvB2FFzEWCnezOp/CevnsgBp0/k2Oneo8/9DE/9xfpXnv2Zzm/7+sYffICXXnop+5Hy0sOfF/uz5uLJrxz9SAEnv5IxF4G2M5lP32/6mXHm75D8atfR537mlcgvyb94/mdHVz4BYLx548+OPvfhaFfnjwCdHh99JWMuej+9sAh0DGXMRfg5N/7g9pnMcz8DYl998ahXwKavLf/pkR/zkVeOPiWBjqHMU60f/cIvIZJnMp0/PfavefbYBye/lAMfeeXop/9bGDsAdHw1ffTXv8CREeeULQHg9KnMC4vgr/wdTn4pF80jH/tIHRkpDL3461/ofPL/8P/4veyvf+GllH7xKYmn/mJtKPOUN5LWFn3tD4ysLQEkh1886v3EvvAPyi9Dx1fTL/76F8DJr7z4PzL/4In6F/Ovdr3400/8EMafvXj0I6WrGHykq/PDYx8AQNPXl85kXlgE6z+V6fzpJ/hz+xc6A9rkAD7nHHrlxZ9+oqsCr30lm33px0GA4h/+/o/w5d6j//eLfwdAl+eLq1/q6vzpJ/gj59XsP3ziG+iPfgw+9pHC6VNHs9l/uO9XBfDU138E/u9//Qsc+eMPPOsZQQOMffVo108/wcmvHP30fzPeXPrTFz7Q5T3S/z+Gjj4l8dTXPsCfF/u7On/6CWJfPXo0+9KPgbBA+pdImb0WqSv8qT9f+dNM2GwiNeQ1mKY3fhT0LOPPXjyatT6gUVdDpuZrs2yyoP9IHsvzbj3qLRYB9DkyaQJ9zVN39KabvpaFe90wBtcXeC+AWHee9aKg3GIRbGj9+r0BLMQEutbdIooFBcC/1im6RQnALbpFFW6DZYNFACeaV5MW0GatLLZ2QyYAmBJtx+YWOwQY6wnH3O9L53sAYqby04NIHlpNZrrzrAeI9SxPJ7vBBp3VdgGAAdnmednul0iX5OP3cEnPEoYd2d52aNWw4KfXdnh+MWkFpWe9K9c5gpwDpwqyHcDQs9fvDEYK0X54Lncc4L3O6nGEG33ZKWfhV91F5RZXAQGcODwv283QtgB4r7Ow0QfA6PYy1dQ5t9CmRyGiNI8AYAzkN05B59NdRdEzKBty5jeOA4z15pXeaKAtWlRe1jHkzMvIHgSvDGxwfX7jFGD0FO6cCofjhxZ4TxCivXnqjghM3m7mf9UdFLAgkyaarJUFwwJEmAHdiADEvIalq8ItFhXCAMkWhSaroI2py4N2Xd62wyo0kKapc27RzYdVATgcYDABs6CtFzbAU89e/7lArGd9tR1ArxuUF8ZAXmU9n3oi6aU39Ox1VZBANH/eL5Eyey1SVzjjgyuyPdJsghryG0zYs443X1fOKk0YGzPPv8jgvK3/6Lo2nssAwOZFXDoP4PPjE8wEtybnJwwTAAzr5uREHO3OxOTVpAAyt8ZnE7j7jnPubcSaJ+auJmEzi1mLznlc89bI9bV333HOXQVw/233XGTh0nj23FXg1fEJZQE9dyeu5TJY5QAXEj1L1yc3uoHY5rnZ8iXY2+eUTg/oHp9QfUvjswkgs3n51mwXDDF56W9137fPT01w0ytRpCR+/Tw/fpkPTI5LC356zrWpSWUCXJe+eWLuMsKc88+ev8wtGIcn5q7ERFAI1rN0fZJbMFrGx/8+MicwzPHJy7j93ua5qwDAu5emxjd6gtgBINl8fvKKYQKxu+dyANB19/qk7KmaRwCILY3PGND5dC+p2+ckAMRaxqcucguI3R+/YnhZ0Ba9OXvfi3ltfMr/CQjKsHxh6ioHWp3xSz8PUjXHJ3PpIIR1a35c+VWMzrvjl/25pfHZ8+PMgjs5MSn1PNXLgN+IAKS9huVVBa6dO6dWwwDd1y6gyz4/fk+v83vlMcfHZSeAzpvnVWCgYCH/+hS7pfyqAOAwExxxQEhtvaABMnNqblwha18Y/1sBfujcO355kVwaD6LtHh+XneDmxNzss1cBRPLn5SRa5tjmuVm/whG7P3GZm0GzCWso0tS9nsWsW3NXnSskXY0idOS6mcNiiVvXOWVLE4aTgy29xQpnFuvvwFpTWGYmYM9gwS9i2lZY58LJdcWac2EcwbXVmYMElqWTi4ObC8C0AYcLcFdycxFuLgVgvmJN3ZlxlU4PWJOObJ7BIgesFWCBI2Hn4D9fmlKwle6tkZJEiohncq5MwE/vbk4n65WeuUAhGuI6bCXQ6khsMhEUgrfMwMmZSNvSvR65utWWWAuV13Bn4MwmgtgBIGUrrCkLcKddBYBZeWA2Xi2PXuGnscJEmM/A/hKb0gLcaazDqy5t0YDLCnY45NVlcC8D61wgZSt3Orh0TWLGCEIIBUf6VczNBazlQns4uRTuLwI53XjCDJQ2rKAqjMOzYQA758Jcg3O1pIbXpJvjAOxZJzSQr6p5ONIMqgKAo0wwNw5mynLrMaHgKJjLcG8qYFkV/fLCmcaan08vPe5I5L0GHeRP/1JS5nkov8LhTmNTWSXVoWuo3ABM5LB+YJyPPvoZtcMx0ApQLJGuWUAKJBwAXhPx/svMRa+1YBNw/BBmAXAgsGh02RFhDK6tPurLAcAioBi4KwGHmQ4T4I73p2SAnavMtArTwwKgbAUHYKYEXIa4AziR56a+dEVKEhbRlSuARJie//jRK723uB5RzRxcJRC3AVdrj5sDuKMAlYJZADYjSms6JU8zE7YCVnkYu9cT4UoB2HqBSuSAAq+WR7/wrtafqKSbBQAyAdgSjveDb9EtnjVGyuDqfKwHPy4Ajtf9XAAqFali7kpgJrQHJC/JTZiByoalBIBWW4UB5nW9rZeYeQFYBbxvAwOVPSDWVQEAUiC1ysFdWW49Fyb8luE1tcAKtgqqMEivxEAq/KWkzHakwmFLQMZLDRDWUNQAzFXbu3x/pIg9+llU2z9EZxAsC4hPAf7+IAYVlQC9e0f8KgtAYJ1334gGVxVyUe3BtIJXsy6EVKZMOWBuFoixqg/2nUh6QRkYGFqz4MzLpvqULxdB4pGSlFuAhekhHj4FYrA3hi6XD/rUpyB4tqQQHFlAAOX+NoVT+qcLwGXRAS8zbwBQLwYFZchU3vNUyfjYuzJe8rRKFAEo/xrvoai2aPSZogj/DMrA4nEA4kPvjy5gXXpWFkqHkK91LgYm/6h8O4Fn/lRTtFuW1bgbrYqelUgABTDIreYDxaiBQvND/DL8v/erKKTAnQrrubkTUupmCK3/LHj0CfWZaHpu1MqRAmnZiRbGr3Av1qPl1VHRaF3AZUIxl6TrQT6G7APwy6h0ISoNPu3t+p57OFc+Ha24dsuZqwtApmbNvH4fd3Pr63V6JRzXjSQy7u3r1JkvKUm1EbJO70Qv8Nuwyc0MnJmo7FkiVhqVMPoQDVb9zgBRrVmr0nk7Qx8Ae0d7lebT68yqukWDIg6nqpTBGDb1zMo3w4eyLMSc1S9mqptcC2ZHr9iusqJVYZjvYlcBtjCQ+iIM80bJFfKLMG8eN//JrbTevDV6LceilvHKG2Y9oi/cynXpm80O+auo8LLqqPZMTHXOpB2Srgc5p5yBP57fbri7LAFIMMFq2kUdaYefg3kDcKb1+Gr79EqXziqGGe3WnDpepSRVWpaXXlvvymoyvMVjabr31MXKqzdnvYlASej1xhh8avuCe5Tnczf0m3Mbp8rLwIYPRb91pqLl0CHsy0MnVmVVk+uGPrg+lfzc9qn7VdFqy90F2IJ8bOyeXZqLAmeH35fiaafSemuXT/XLXyFsk2XlLWkC8nRH+sbuClRW4bupDnkimb5N0vUAceZLhxBudJ4YUTjvMqNFWrJs4CzUdkOR6KUMCnnOhfRWQbZV1Cq/r8iyWynrWbqO9sqSVMP7MX1/Ai9E2qB7a+P1Si1WZVEpd756CUWV8UPZsKtsKOZiZTfSX55PFQ6LKyzqD3usG7PsVHkZDPNdyU8Fy0ZhQVgkxNqFM5bUJm+qGIG4yNgX8OntlSioip51YDcBKgykq7Urn45fK2uiLO4qZf7P1SrWW1ZvWv8lOvjzyhuZZofjsUN30vqhyk75K6lwUVkd1Uadx35irc8clF7/OBxmFhECpacgopoc6U6a2cxlor2n2rVVB3AuE15Hc1jCUeGK0FbZElXET4TZ9P7LxCK2mB1UJO99mvnyX5aqPCAtj8p/WqFSYckYAMUrg5VKlycZ5mpFKXagPJ8qXildvkX91ujOVikD1080KgzKoMIQdi7lX+FyUSrJpgtzbecFAV0V3JwBSgK4MLermHIDmbfOncuVmz/hQKaErGa9tZwZ/ZZHn+CYJS2T25fO6RF2WYEqyuxEa1m4VZpNpRg475+7oEi69pWS+o88rikaACxXdgKGKBtayVavJq38iiEC6ai4FgzeFsJYeRNzmAm0OgoOSzuAw6zIj4mKBqnTK8m2fl6kOBAzI31Gl6iwXXVE0xNbL8b5liiLytEPIiUHkiIsoeQAs4KwBUMAnaWTQZkGuBmxpjJ3W08iMr6TCYBZq1UtWr0QkTIwACoNJINvuoCEE11kUoHJHZhgPcFITj/crLIyL6rIF5AMJnsiWm/tW47IfQMp7mVQmEVVUfNpF6vclBHriRL5s0pebWEAEBPgomTiyXjJK4eRGErKHK1wcBNM66rY4e6sD017kB4InsBRV+Rcb5sPAO0mANhsAB09kEmBbPTexUUcecMEOw7j8Ky90Q1XpbwYyq91mcngcosNVi40ZMF6bMCVnQXAkV1Aa7f+caS7YsXDS6+EXIcA78W6YYXzIhPtQpcoUpJq6xxeeioBI5JYkwXuLe+4LGKUgmFBR+yFVr1Aq4VVw+S9QFDCdaMH7RaKXAhPTHpgWKU36MWYiXY3YiGZBXjPjoNhP58pbw9D3rCQFLmqFg1sbyGy0K7L4DITxwFIQ/DQbVvSYj22H4IPgllF3+SO6opYsU8kzRxUK3j3No0oUhWd6/G4KA0gk8J43f8jxcokwDdQgZt8AIB8eWykfK1KpgpwDEcF1vMboNEDZkrk2gQfFn4LxHEAYIMldgfgGH89MuzdwcoKVFrmSIWDnUJS5MqbTUWDAeDws2PDncBzwwdBAQ7oWpcjTx45F/yRs0ZY6n0AsOWRs/GPMd89did1O3p3Hytc3NwYXeWH5jK2Qu6LgDx29n0JVF6rRovnbXX6hecrlnsXB8bume8DkN03AMx8fUylP/TvVxXPBXV6pUtdfEwmnpm21ekjh2c/42X+RFOn1CWKlKQKOr3VY8Ot8nfDO+XpjoS3DclRo8XzQerydAdL/SBspjMDIyx1A5ts9M7zOcAvoa1OtKU+hMNHn/k+ADtnjSSa3y1Jd52NytTdyDhi2hpTid/ZaVHEz6fsHrmdA2BvnO5Ib8rqFvWrtb+tM7xCl8HZ6L9zDMAyH7v3TJCN2dNHzBtBiI6xDfPdwOSzr48l5o/pOP9xhDsS+e5hw9mmEUWqwhRnsfm9kgAL3SN8TY/4/PJUGshmoxvPKGDaSiTSh94pWz64AZfZEev5DfBkkrfksNB9dkM7EPLLC7vtbPx+yfhtU5pppKUCKgoULXO0wuHExhK2LG82FQ3G60iJRLooxYHY23WQRl1Rd3ULrEXvRVTApEynlr0WP6Hi61exOc1TazPe4N51AXuRt8C9vJE+PO9taDeABdaiAITXBlXOmuFOs8NTqtwP29JiwlzKAViFBHB3MZFez/nr17J8tVun500wXC+uzSs83TwN9wozJ/XX8/e6lqVfokhJok8RvE+d3vKdzOZVBKXfvJNumYrk3A8xsZE2l720XQBYXk2n1nNwrrBDU0BYwssbqbX3YMt4EVDAlEy3TKnQtgDsKyy1Hn095P5MIu3M6GtK8xhdfdb5nFcpBQU4lzfSzmV9mY46sKiO4TrvvBWpZq8MznTCmlSAPcObLwcWXr6TWpv1QzgzidSSDEy+JBPLwarZ9UTLnMJaLvPsVTfMr1La5JGGpatC6QWoIACwuZiwr+jy+eWB6+qgvoGcq6z5MoDej8+dn0mWNp48FBxVjFhPN0BnMW0uKdyfYS23FFwAfnnhXI07VwIzw3URO3zu3MRGd3n+XLekzF66usLhTCfcqbA6UN7U3cAA/NilcxMbPSg6IBoG82fmIrK2JPwvWTqYnWuvgokUoiFYOvg3+C5wd+9dGy4EpPWvwVp6XC+CAH468WhIAcD4a1G5hqBT9UMLAEikvY8w+kTKuz4hSksSZCosty5V2ttc6sedLsl5xCiiZEnEjzohdHqm/gwijQYT0aWQRBqBCUoLzsryGNeXxcN8IuE/GeBVovYyFUTtG0ObmJVYC2kRecTA0tEQuhJ1AJbyS9307aCAoiLDeoEybBKp6Pa1MEDwR1B4Fil/YCCdQf6WCRhvlRU0Hv6h8+o3QP1n2gyqyivvC99AYDCd3gvfAPDyX1XkT5SUuaTCn/6mTkhXR7SG/Aaj43jqmwBe+AYdwvVE8dQ3yQaPJE3ffuBJGm8BMP667nhe+EbFVy//FYCXv7aXWJ7efcv89Deqp0oTxseZzhWyAeHP10ygdV8mXYoDzJL7qAbMKhwQM8eopTWEw9fIBoSWLjUyzbLr+xF1sWlItrbk9injBWN4tdV8l6TrieKWIhsQWrquvNGH9av7EfVd2d7u7q2t7eF16s3FzgyWJNUgQTx8eNdDSDSR3c0RszvRalV+xzLZPe4X5Z17uHjPsRMEQRAEQRAEQRDEQeTlbgCA8a0tp9ONfeepdaS3/kj6rcg59sbQ1hcaw+gY2Be7xQarfHncCv/f0V178czqhala0BNW7YVIVn9X0S9GcqdXGXnf7tPi/fotvhrh/UN7DMH6RIOsUrWy68tL0usHVWrP73B7MW4j2KqjG0M1Fbh/2KwnM7u5SHkfLgDwKut4LNNQA/Wls/VHkhaR99/1FvvqBkx5R382HqPakmdrJK1E7ZKyxTJw9YKmzYrGtWvVbK2+yusXo3WnRWCW3b0idnRl6pKuE13tewwfPfx3LxJTWWyj7vXt8rywoXSXrr3yqvM7HMvu4ubRQLbq6Nv1r61p68rsr3Q9ZykcMcHedMHfADoG99s+3LzVgPOUoi88EjDKGl1X18PNT6Za+tbdiXpqjXUvTTyAWudZ88FUWcvkdNW7VtfOljwQdK5VcUzeSOlK9eGfB0zwZLfbn6hsLiN7qkdjeOcYuDu9ssW1e2KbDS0dr5de1vCjBHacPrRVuWCPtqxWmK2NUVZEVdxVlvYNFcnPCT0CZObCrtyvbt2Wp/fwTkNVa29Vb9G26JZ77Irtj+G4s7BQ1XTFLS3ZUE5072iu+pqMubJSj3TtvCX1ltmNpVnYVwbw0niVycitPd0YUzvHwJytr20QiVSpxDX8Xr3jCLrqHHVvtqxamK3YvFD697WW3WVpv7gd8RGU0GeasHrrwd1T+GrW3qreom3ROVdsxHRp576yhSaVVd3tX+5XJ/nUzuaqq8nUeULEzqMuZ0I8cwXA0t+ySYkDg9rDWOoJmFu6xXp6+T7kRzW+Ih6Ul6nig0lG7M507kFtvfVW187SxYfVPw8B6HjN7be0HoP7S4pxeEfGJbwZN88Et6eEt5Kc7hIAEoJlTIAJ7/JEl9CSnQ5i8K9lAnEB/1pmAkgALIWMGSbMMl1C+yzhAmCZLgDMjK50egPpIENeBr1wcbB4ECrsMjqTmc4wM0Aia3o/eSsciWChI+2vlLMUWlMA9/7m2VRYBj+uhN6knPATZIJ5hfQWKr2chJZgXq69zLMUEl3QJU9El1oSkcIgo/PjGUVfE2SXpYCEYF0CyKQ8q+l4/av8LHnlYpnoWE5b0TdYomR5JfhNlOfcLx4ApHwDeVniZpC+XwAWRzyOhOBdwVXMRCYVTS8R7Pb28pLWZU2gtVMbXV/iZyTcHq7/m9IZ9q0dzVZQb/7VXCCTCtqtLmM6moNqlR3aPGMGtVhSlEQ25Q/l0kG8KClaJC+eGRPCC8JLvgU3o8btSvmBMpEGjxR4VxBPUC2ZaN/0/y6rFj9rZR29pMkAQGtX2CvK+mXYVhNeXrxKjoPFww6aKGk+u9nUb+x4xcnPX1v5YpM03px77r/+qx8heeQD/Hmxv6vzp58AMP49MtmP7r3y4z87+sIiwEdeOfqUNzQz/uxo1z8pdAxnuv5J4aXPN3/l6EfqqX/HstaPkDyT6frpJ0cGfoTYmz/+98hkP1IA9LWf/jrPvvRjwLv26a99APy5/Yumry3/6ZEf+wmz072ZVP6TV0/kgNMncjh9KvPCIp76i/WvPPszne329Y0/+ADJr3Ydfe5nMP7g9pnMC4tg/acynT/95LU/MLLP+6HYH3xwv+NnAPDS55u/8uJPn/3TTJMEOoYznT/9BCeG00c/Uvgj56VXXvzpJzD+7MWjHykA7PSpTNdPPwGApq8tnTna9PORrLkIGGMvvvjrX7zwv/Cs9UHs5NpQ5ikJJM+kX3xKAm1nMp++3/QzAC//a5499sFnn3e+cvQjBWPklcwLi0ZoiZNfOfrp/wY+8srRT/83GG/e+LOjz3042tX5I+DP83/U9eJPP8EXV36hzagLc/IrGXMRxh/86KU//gA48scfAEDsqy8e/Ujhc/bPmt74Ef7IeekLnT/qGHrx17/AyS/ljDdveLXmldHLkl/u06eOPvczAEge+cCvVi+b0WIAySPrQ0ebJPjoKxlzETqc3w6+3Hv0/37x7wCAj30wdjRi0CN//IGf/l+YrVlbArH/FUetH/2R82r2Hz7p+Gr66K9/gaf+Ym3oxV/f9bIJIDn8ot++Tn4l88IiH/uHT3Dkjz/Anxf/8Pdv/TuWPfZBx1fTL/76F74JverwGuTYi0efkuBnf3zmX/7Ib7kbX/hvI0fNwAZevemrX/z1L3BkpDD04q9/4bVF3Xr42EcqeeQDLwcAyipb2xyA12a1LUqK4jUIY+wjBT720YYXL/vCPygA4KOvZP7lIl7+1zz77M8AfM451Hv0I4U/elYi9uaPcPJLOf7K3wXGPfLHH/CzH4wdNb2aOZp9PgcAR4Yy/3LR7yj87AdjR19Y/HLviz/9BH/01Eu9Rz9SwMmvZDp/+onfN+E3oZJq6RjKmIuVHb2kycDrX8/9zO8VZf0SANrOZDp/hJe+lIMx9nd46aWXsr/4M2Q6f6Q7KDq+mkktgJ3u9ZpPl19rdY26VqfwzFUJZ22WTWp/GCJ5aNWw9DykWFBAnyOTJtB+eC6nvbH3tSzc64YxuL7AewHEuvOsF27RLSqwofXr9wawYpjIbPoxwL/WKbpFCehrg7sRGyyGCbdZK4ut3ZAJAKZEU+fcQpsJxnrCsfz70vkekG2el+0AcKogkyaarJUFw4Kr3KLyQznfw2/f88I0ded5z2BeZQHeu37d6IbRuzLPegAkj+V5N9DePHVHAIDRubLguwZnp/LqePuh1aQJ9DVP3elDQbnFIsB78yoLsCFnfuM4wAadVf30vqjc4qpvFRw/NJ9LitASsZ711XbgxOF52W4CGHZke5tXctFxaEG7hddm9AoT61meTnYD+pB6fY7E0LPX70RWUpMvrRq9vfmNXj0d8WpNl9HLki53rHNuIbzx6Wr1shktBgA+kFfHgbbD84tJyw+nAzRZhTD5E82rhhUaFEH6RVUs+m2pCLQdVuC9zsJGH8BYb36j97gjk156Q8689DZb8J7l6XhkcpVsUW4RRcUG1+c3TgUm9KoDADDcMi9PCADH3WjL7WheTQbZ0vWGMKAxkN84pduibj1+wf0clFZ21OZssBjYIloU3SAcmAB3ZWm8aDdXVpPdKCpXH84RO5ZnvRUzyXLjJi2wwfVpFCQAGAPrq0nT7yjACUcm27w2jKQXX1PP+oIxEPTNsAlFqoX1rlznlR29pMlA96/2oFeU9UsA6F2/fi+6xG4qr7p1B+W9hYVWC23H5hY7BJKH5nJmAyaMH+cEbku4Fxmct/V33ePjshMAnPO4dl4Bnx+fYCZYz9L1SW4BgGHdnJyIo92ZmLyaFEDm1vhsAvffds9dQKx5Yu5qErayYOX9GOBfe/cd59xVQF8bDg+fPXc1SLjn7sS1XAarHOBCouvu9UnZA8Q2z82WZt4+PzXBTYB/9vw4s+BOTkzKTlzL3T/3XhAquhI5nut4ZvyKASSbJ+bey8Bdmph6rxVA17XxXAawbs2Pe+3p5sTkrJ45GJ8dv8JfH59QFrg1OT9hmHffcc69DfD741cMgVjL+NRFbsEQk5f+1quS2+9tnrsKGLfGZxNAYXxqklmhJbL2hfG/Fbx7aWp8owfgz49f5gOT49IC0D0+mfPS1Wb0CpPZvHxrtgvwThViZg4AjMMTc1ciBxt1TY7LjvvjM3qozb1a02X0sqTLndi8Phm4PvCr1ctmtBgAjKXxK4YJ59rUpDJ1OD9Al31+3G+w7NXxcdkVGjRM/5L8+NwsAPsCLl0EOm+eV8nm85NXDBOI3R+fib06PqFPa1sbn7pieLcO+/KtqG+G7msX7r/tnnsbyxemrvLAhF51AIBhXpuaYN0AXrt0MdJyXxqfUJafLV1vQBAwtjQ+Y+i2qFuPX3A/B6WVHbW58ey5q0GfiBRFNwhXmgC3y+JlPbcnxmUXbr9n69acvjb+XuXpUlHj4lWv9bVM3JrdnAWAVmfi0kzQUYBXxyfY4K3xXNpry++1Al2bFyZnk/D7ZtCEotXCmyfmLld29JImAwD8s+cvc8vvFeX9EuDmxNxEtKfdPlc4j5sX/A5qtIxPXhHoWbo+udGN1NL1yVwDpCtcMAzX1dakmysJOadsaYK3zMDx9JI7s1h/B9aawjIzAXsGC36ItK2wzoWbyxiHI1ITXFudOUg/YW4uANMGHC7AXcmsPDAbBzBfvtA7pWArAeA6nFwK9xeBnO99OQgVCt0Mcuw6CsxEak1hhcO5AuQZgM0cFhkgFLzDgzZngEXfINexztakIwUMJwdb+mVwp7EOE2lbYlNaSNg5lD6m0VZZknBkZLXDXIZ7UxnuDJzZhGddPJNzZQLAmoSWHm1GbbwVaAPbzPTPuG91JDYjbpk3c26uZdorUFhrQRkB+OWOu0Bw7JhfrTqbpcVwprEGE3dzcHMpHc4PYK7BCfRlWbqzvDQxL/0y7FkHKVthTVmAO408W1a2PgDssoLttUbmAtFzi+yc91DGvQysc+GbMKyOVjsHO5cGsCaj6+zL0snFS7MFhPXoTGPFt2C09VTmQFd2ic3nIIM+ESmK3yBkAki5ZfFysQh3JrqQs5nDCq9YtC8zrpOLgzsKi16L7lxTuCX9jgIsK1s+O+Ml48XHrHlgnod9M2hCkWphLlCo7Ojl5gKuw1bC7xUV/RLMVSXxODP66YLuoOlNhaUcNxfh5lIQCq5siHTd9PTFDsdAC5GGDQCYBaQAd5R3RimQcAAoZi4CjjKBzciZlmYBcCCwYGTsUGrCa6s/V8mFCXNXAg4zHSbAHTCR8wYbdlWpVgJwcnomFT7FCkJF2ghQcHJw4B0I6nrtxQWAFaDI/ML5WdIFcnJwsQCouC6238xsCUfp2ZtMIO4AjixL0GEVD9eYKQEXCVsBqxzALFy5og/uWAAcr2f4ZgyDMMAzoT6ENG6HRz/qQihbege3hrUWlBFhuUsOlo1UKxTKimGr8HRJ/0BaHYCZMnJy/GJgbbe01ZSxDu/0WVcKwJZwsRg9s0QHdsv68nzpbhffhGF1xB1ogSjdLLUYnFQbfSAdBLRV6XGNYTWV50BXdtTmbi5ivEhR/AaxyoODdlXE2hIoRHVhxTtlsez5XLlxFfO2U7CwHoOO4l2w5rVhHR8TEt6NwGuFYROKVku5ZFZpMn4X8M/vZSjvl94pdSXGUqUd1CwAUNyVgGTArrZF7cbVoFv10TWrfNjJkfWHaCJyJnvQlfVGDvGrLACBzY2BD6PBS67d+pEwA3MV4EJIZcqUA4aMp8HVdkOER7sq76lPkz/q8kNVKayAuZbVp0ykeOQX+VrnYrDVxygLpQDBsoD4VOmQVRQBKObZRH2q2vaWkgNoWeSM++CIbRWa3PEM6ZvRC9KaBdcHS6ZmE97WOMGz1bYfuOXVGJbRL/e6MXglcnvzqzUh/HypT6HVhOvJgPoMgHhc+OF0AIayW6dblljVZlQEmHkDgHox0vpEWJ0CABw+dFlV3VbB4vHQhGF1mC50z1VVt2SUZitaj250e1dTkN/yHPiVHbG5q6LGC4viN4gC10WNxpvwqr1005NbZWuBrNwNxIC41/G84H5H+blfS5HDmzgUAr0BxL1IEwqqxd6oUsgqtQi/ZbOUAThl/RJwZJ+K5tcp66B+drNAjGHxWOdiY6Rr9ztRjD4Av41KF6r11vZ2fVPqzpW34cqevdWOENfrpWYeDH0A7KqDyr5OQO/ZUwzo6BXY9CPZMhTAoJ82JIdE1M5zVr/wDj5rPyWqBTb6ECQYmOWX1TQjmlp/J/BhhXQB2+9Na4+ctorjwcXyc0hp6YpVyU0lkTL65bZz1psXy6vVz6bOU3s3nIXAsid6gd/qcJF2UHZbEKrMoNvttGNVNuqkgql2z3DVt36MYTM0YbQ6itvuJSrPVrV6jLaeLXNQbvNqtvAbhMPMnwtZFq8vcGXSJXbRLxxuSqtYraNUqQk9GBDVmlB43czAmWpv62xVi7pXVPbL6TdGxuXWHVR8qP8CsIm1jdPx6QcrXXCm4Y9it9twtizhDZzrPTVJfg7mDQBTpUIeMaU1pyIHUMcG16eSnwsWwrYK5efRVWyoee5ed2Sge3noxKoEwE/ZU8arlcHsGUTG8bvbY4i2zpXV9vKB2M74ZtSLEH67y3Poc9qxOevND7a/EZSU0Ss3Jt3Ogatl1aqz6RdjeTPS0Nt6V1aTL8L1wkXawbaJ7ZV+c25DHyt9S3QOvl0l/uFD3iWeCcurw92VDbaox9LWs1UOym2+tS3gKPO/O6o83jq6Hxu5s+UBxNvdJMqbUMjSdO+pi7uuRb9XVPbLu1cGh7+7Ywd1pgE4cCb6jhdyD1S6lDtfahG3eke0vcsMU3bmtr3ZbLG/OpiM5jkXEi62evWN9SxdD06CFy4y9gV8+nN+JNu8MOdifR4AYi3nVCxaR2sXzlgSQLL5vGqqIl3O/BatQ7j+mkolPR9PIPmp8v6l/HXSqoPO0Ix6QUSG9964uOHFUC03lU0gUka/3HAm2efLq1Vn0y/GetAfRRHp+xN44UUdTgeomFIwVzWVG3Qbla8oOrduzDItXe419VqV9sLNdyU/FQSNGCC+3f20op6rWS7aesIclFd2WchIn6hoEJDmT52KeFnV1Znd3MwM+1eppVy1jlLZhJQ3qRRVm1Akhlsbr1feSGNb1KLfK6r0y6Urg6bcsoP681Ztq/Xxsc4dpauhh5k5YdNQLBznVsiRzmerPZMsHxNXG29UNDiXCc/6Dks4apvRNBORObPpwlzb7Rhcr0Um7ArZzXl7se2qbamqOql4RLqqJMrMRVSZHShWTbpE0Bajy4LRwjhIt8itc1O5GhgtSrA+617nZdXqZ7OyGKbyV5u9cI6/7l22tszdKgatuq4p9EpzefjZ6Mo8r2JL/YAiasJIzZjObmywleWirWfLHJSHdNjWDQIykSqUx1tk/oOOkm6qoifWVjMugFb73Lkr2oJmSUep7PbKhQB4uGK2VX9YqvLQf6ta9HtFtX65pMzK1u93UK8tu+FC4nsJ7L90iehcO/ijaACwXNkJGEKWypDUm1Q619e5FcRQ/VqHAzFRIZEm0OooOCztBPVUsS7vWwcAt4JtmkK3L3fLh5neWvdWyqlK7uNl+E8sSwRHJgBmrUJxIGZuoccitKWrLMBAwRBAZ3lnSwMJr137ZlSs9DmiK9O64Rd3XbssqJtUtSUUEc1mSTEQE+C60kRZAFeaCMa86IG/BMd2nEnLNMBNuWVOfRu7MMG6q1zimzA0gDQAmMVtkmWReivwLZIVVRbPSyq73OZOFXnzGwRWuVgtj9fhJpCOVnvGe6au4kBXpJajxg20vRi1oAg6SunoEeCudGABPHz65VbsUxHb/sm2WBaKA1X7pf8AOlM5vgBUK8BMp+w55P5Kl6tSLDJ37/XPMbHZADp6IJMC2ciGAJeZDHnDBDsObs3YsjOMoeq1LrdYhWcNR2bBemzAlZ0FwJVZgPubS/tfL7lTwtv43SdiZg6qFd4WeA5RGqqc1ZgJlkWRi2jyfBBMr4PyqmFtPgC0mwAXobDlDQtJkcO6YfnTHRS5EBEbos2M2DLXJviwcFgPDCtfPia3WI+tozXBjnuFQa5DgOst1zm/4RcMC9jR+55fxiIXwi/3cYGEU1atfjajxQDYINpdCZWA0Q0dzm8HMimMoDbaTN5dKDeo16YT5d1iMWaivXJrj8ssv7d2mGh1lSvT6ChVE2bieGDCoDqAdd6DpLW4RcsNsqXrLRIwbItB64nkAJFAVW0e6RMVDQI2NyXK4rVVL3h3dBNUrAe9NiCTIhmR6RLj6rRjZ0e041FpmKzfDDpK6XSvh/XacHPtgp2KDJ+iTcivlibLH5hVNVe0FUd7RUW/jHWCmwrFmDDKeo3XQVcNCydMR3YBrd0Y2mZ83MC1Lnns7Pv/GD6PGGGpGwBgyyNn4x9jvnvsTup2dGFytHh+c2N0lR+aSzoSuT/0YpBA9WttdfqF5yva8OLA2D3zfQCy+waAaWtMJX5npsowMHeiqVMCsP9xhDsS+e5hwwGQN87+9oelocpYuzciE8/O23z43qHI1x1jG94RmyvGsNFcZYCSs0ZY6n04bKwQrG7aG6c70psStjp95PDsZ/S9dfSZ7weN5rX+tERoiYXusxstsHPWSKK5/EBPefoF01vJ0mb0CrPCx2TiGe+xzKq/+29Tnu5gqR/sUIN+GR0++sz3dbmNMZn+GGXVqrMZLQaw1nY2vqmwemy4Vf4uvHB+gIXuEb7mrwn8ZHRDvItygwLAavfZj98r+WadjcrUXVV5y+pv6/QaAx9dTWwCudfPxme+GL1ko//OscCEfnUAsOWJZKJk4588dvb9n5fbQNdbJGDYFv3Wg2gOIoGq2jzSJ1DeIOCwFgmUxuvmusd4y6XI5Qsn2lK3gTwf4cuhf50S4+q0VTyO1HkFYJmP3jHfDTtKiXJ68S1YIxvRjEWbkF8tf3O6w58cRjt6SZOJrpb5vaK8X7L+du7ksM7G7m2W9hevg65tnD6Seh8zXx9T6Q+RHNsos1hDRl1Ku7fyfdctsBYFF55b0uXVdGrdax4TKr5+FZvTPLU2440WXRdwZ1gz3Msb6cPzSK0BK4bwYgDCa4NKZ81wp9nhKeUn5ye8tJgwl3IAViEB3J9JpJ0Zf+AeFbr5e13LEvj/t/du33Ed953vt2qD8nHGIqo3mvZyYpO7uwEpIWOp0QDoxHZsXJlZx2dikQCYdU5sJ7ZEZq2sZB6Scfw3zEwmc87jISXZM+u8WLiQM28irllzLFnCpdvJiWMJly4640Q2ge4CMxorJnrXeaiqfelugCBFiZL8+6wlgb13XX6Xqt+u+u3du7Ve8DtXFXYqhUevaaAhM8rVarMXVEC4eDvfuYbGUtepGUTvXA2X/Jz5ic1GuXDSelnrRNU5mc9tSjTKvNM+7KUQzuzlwxlAz7JgzjlfZuqRVVb3SrcqzpYA7iyxzhWFeZnvnFeph3GAzZvBzrKpacxolNmf5fnji0aWMDLB9F4+2FRQClobJXRCV+s1p2NDZupO763b+fhRbedWK2ZSDaBxLRPOAJs3C/vX4OrZCvtlvzHretzYExsy6szpoxSwozKxz5QC0Jhlud1ZJ61y9tcLvHfF/HNzL398FthQmY2KM7/WQLjoF+dUZELrDgB6fi9/fD7pbWdtc0PR2qDMO5N+TI1FN3qsUYwErpJztrN5NEWsLZKquAEBLUOgqV2sSL9zRcYzTO18P7e7BOxX/P3rsetSxrXrkONXrsyYl4Y0FlmwIeOJkopwpr2dJT+3uRzPTTuEUm7Zv5nvnG+d6KkhkwgG0axIzUsAjZv5U6tAY4kbTZWzlp2g4SLLbVRwq+zndytY9HM7y3jgMJcsdIsb32bfzN43H79jxFwgzIvw4hp5AYDlXSsZ20KibLxsygtzNsoLJ5rJpfJNpqYAwL+Z2mD7OVvbypE3jzSa9400vaQvap0BgJ938iY1jqq4lhJVmIgtwPIi0RbPC2er6Jm/fKJVnrN9OtnN+8hZXqREEsA3iywfZxrMeatMXtiyXmwCU4IBKSUy0SFzzBklH7djm04rlRAz8eiiFSJv+nCfcmkj8W8G9lcNrEFT5sm13AT080hIG3Vn3RkLyHOJ9LVICudMmEvbIpGh9APbuEjawOmdz7lTGTcWrToZOwtc0SZnx8dFoqW0Km5AMJEwUyZqLkjOAxYbNRqLIvFRRFp8+msAfvep1GTK51KWFcDvPuUkbJ6b8RCK3BIrk57oqSGTmALO4al5GeuUUMFNs1z0N/7j/0L+kG3Hv/lAq/fNIyQyj33jvSd30xWFeCf47acSoetA7lrgfQL/oLmva/cXfgT7Ic3iX0wyB96Y/SAuUj5oChWXfuEHcJFC1y8m8skR1R1WKHS9P9n4YF901BHKtD7LSfxCsHGzH/oG2YF4L3KU31i9j59FfcdhfeS7d8HKhSP8RuTb+t1WgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgiAIgvjgwUcPPJWNf/+EjYrDKh9wun1b90Gq/YHiPdbuGQSfuMc63ZPFh+WRnsF3QFY+NiEAdFxo34Lt8oDTCfoDmjIPZNqRCd42rS90iaZ5V2+ilDisMrvrq2oSbSUZC44oZKL9fHCPs90vghUO0PEAvPH80MPyiH+PQfNIsvaUCgKA1/696e73yry7vla9vaOzQzSRKHQ9fLx3751IvPRg1jal0gPWMRtOvW9eWHskWYu3ptU7J0KhRNOGQtc7vciavGtc0lrhxIV3RRqt4vnkHWFPN3DA6krVj1rSWkDrwydysL9deVcuE2ml25n9rh47iqwsWN96B0OXSvwuWc9TNMvuTgeZ4J7Jr9ytxP5VBf7uLLz0C3HMYUf4DSj/w+2P3+g8aklrgf2rd5nI9XfpSpJWuq3Z7+qxI8jKjvRy7fvm9Z8kLJ+jSUarrodE/X3XlVYPuGPxPnLX0WR9R0OXVjRrKHQ9SPOYd33nSwIAWKEEIINMBgAvmB/gjFI+vrBLAF8gByZYRgBAIXUJtS35fU2VAYCnzvnmk6tQctOLC7ieSwKAD/iCFQKACSNZ1IY9DrBIiAyMWKwviDWyzbIc/FJU3Ja00rhShVxkgVwsXiRT0krMyOKD9yWslzfFbfVINL8vAMACFHIwUrDASmM7jaSzn9NKcyumM3ukeMJjcVMB6xMJq2TAMmA5FILIrc7u0T9Z3EgstXWBE8xSKCXOsUIp6YFCdK/Fd4PLWJgHQM4MAecj4vBMAZngsMh16fSZN99Az0Sh9A8K3uRnCo+XvT9Boe/Hik9+5swjEvjt4TP/ePq7APD5gQoe+dpLYF9989Yz3/30v+R9T76cfWx3/PSbbwAA++zfKvRMFHp/8BZ6vpw//eYbycpgn/27yTO9L8Gd6/jymdIP3nIVvK+cCX702MsA8NjY/xw/84gEv/iZQlAGvtp444lPHf/imR+rR/6I9RVfAqI27HHg7BdP/8/C3yoAvx909TXkr4cnP3P6B28B50YKj5cBAGd/q+I9/cpXzjxeBs6NnPnoD01J29InPvk9/pnvIvvl0pmP/tBY4J+e+S6cZR6brDk1wS9+pvCrZfx+8C/6fqwAfLU+VuqN+mKX33wDZ7/0MrxnvmtLnz4m0fPl/Jk338Ajv78zfvrNW18583gZx7628Uzh8TLAJz9z5mN/h0g6a/mU0nyyZPo3Zv/EJ8Mvnvlxk8cAKz8e+f3dLz76Q3Zu+MxHfwjAyPrmH2z+3mPfc261dgecjh1/jDPFlwB0nH0ZtqpzQccfFx6/9cnvmQzM2c3fK3z0h2BjI4XeH7yFs18sPF52FdAzXnik68dvAMBXt3+ndPoHbyE7cfrMIxKPfellfvnlS2eCMr7wG17fJys0+2jV9TYY6FzfG4Z3YXedDwP9J9cqWaHrqNcUuk+tVvqBY8XaTZsYVj6Q4wATEgx1petVgJ/f3huJrT28u+ANgl3YXdsbSVUG0HO8mjg38ujCj4SrgP7O7WzRXW3Ob6s+4MSptXK2aDY7HYPbbBi6rusKQNSGPQ5vqHbTClFX9TqAjie3+SBwrHd1/UQQb5pGQ5kN0NG7up4zJaOWDH3H12Q3rAUYIsvAi9XsDraq2UHUlTY7SpE9WfWKri8tAyBgADc/Ftl9autmH/hwuL43CjA2vL033B/KrADYSE1mA2Dg1JrsDiLpnOVTSmdPrlYCo5+uVyPFUx6L5AdjQ3XgxJOr5R7hKgHsQj1q3NkdkY66jnq0RbZVnQvG9faJaPnML2zJboFjxa11rwg+tLmYiSrw4bDa7zapPSfX+RAwHq5J93TNwPFqtgitrB8JCl33mwEubsw9q9AdTs9dywrUpubnWDG8ihtXFRvaWJjjRZQaV6dum9KSAYIF4KEC8Pr1/SvXAG9jaile2WaPT69eLwCbz85f40hWBvDE1JQsuXMsmF+dUq4CK25MPe8mR3ZjatYLEN6Yn1P2WGFlatnHnef1lWcBRO3b4+gKp174kSn5gvzplWUgf2OqUgBKtxbmZPxAEf/U1DQL4O8vzF2zJV1LhsbV+WkeGAsAiCyDjkhNNvT69JQs4QX50ysSADA4NSV7o75UBhA8+onu3lvTUzeRPX51btYLgI47U0sdn5+aZgHgfeLqFCuCD27MT+0NRdI5y6eUzm0szFWQMLtRPOmxWH6gY//KMoY2Fub2Bp1RJLxHr1xzjTtHxTo2nsULzzk72KrWBV7w6tS1OHTdmZ7hAfTc9JzshdeYWbkWVfA6p6b+2u0FB6fmKnlgZ2p+1rPHPj81rYq4Ublz5TrNPgpdb2c33bmExiyKOwqbLMCGRCjtIOOdSwgrAYIdhNfMoRoHAhWAJ25zI1zEFosSF7kdhS0OPQPscpGsDGBT6uXoHBMKoXIVuChjR8ZN7iDArQp0xSZrGktYT3jSte+O5xpKL6Y026+gzMCK28BynMHBqmrIABkNVJtaMswrNFQiDeMsk1CTizL0UjIPsSN1hUd9yQx4JxcI6gDAgyWEM8g1FHZUEdCL2GabqqECAAsIKzl4egnhsh9J5yyfUlooaJno055Leiwl/xoUD8qxAY32kK5xa/eUjolJY6taF3C9jI1omaQXsa+KuFMGKgxMA9tRhXxD6oXYLljygBmFhrtPsCnDSobmHYWut2+cUAGKBWUgtAsclTqVY4EEdm1I4YIJKZAJk3NIQcc3sIIqoE0k0ECyMoCyiX7mnEaQqMBDCV1JNCnT0uwDIUvLrhPHg+1kPwCwBdQZmKhEfQIAlgEpoHhrSzGJ0BVbJlaThxKoJYVZB6qI+qpysFAJCBO6tAQ0gm1ASwE0JDTKgAQQVgDJ4TcUUOWRdNbyTUo3PU6QONe89VICaFRs1zIhqK5Ebk04KuV95/xEVWUWkJF30JCAzLhzphFXIagB+yq2S8hEysJlQDGad0eDnus6hIwG3AM9SsDcRHQhvw8QYIiv9SGCvz4u/1cENRw09YOdPjAAYJkMosqsBOz+CAA0c+d0ZUDKqALTzfPv4wAy6ftQIjFJWeJumFDitTbKaYCh0HrxYtj1Lsy2bSllgbRlIjV9DUAzoZoadX2FHDyUgQxeAZxiLHgFgDqdaEdEcUdoZxjTkLV8Wunyk73lFg2FapI3kl8DYLoP6EiGLhW5NXZUWkcnRFQ1kxGAsNeqrgB63eQOzgBguWMMIR+fUVEF0fRASehkjJunJBeFrgeX8TJD6sNgY72AjQLCGwXw81QskILrKkcgD27J5mi9ieRVnI8Cr5nQBaHsubXixRsVVyHTEroYMDDcJEBip5tqvzmOJOQZBdBoPtyoFJ9+Du0l7QV+0mqZ1kjd0qXrK2TBz7T6NS5kHLrshG2z3lAMoin8t7E8sLN3LrPYRsGEx1rk56NmeZaKdqbxyFEH6Oiqpl3QPYhwPfrUMyywj8bS0MS0atuXa2uCHkCl0PUOc6J3q9qdSGI178Lkr303DDkTrxzcxKYEtGITJ1cT9+3C+URL7tzOzMiY3LMVPtZOmuGtavZ0+5CUbv8Q5gGELeuxOd17/lq7lrqLq6r/fs1n+tJKfKgqP8vC+15etLE8wunR/lrlcI81yx8utlHeNW7sfqgIYbMLNvcTIbbjwu589teBFdF74fn2fRnGgiP6iqDQdY9Ey4ihn04jay+/Sq+1pAqV79dClumUB7e0a2oF35E8Hq56zbXEtPLcuU31dHHZVqg3L0hEHfk703i8fehqbl+Jg+Rp/628cI59ql1LbGhjAd13WWCx9vse15dWwc9kjdlnI9xGULRJqgGA0HZBFp+zlm9id+pSb2voSnqsVf427USN764duog0Z5tcsBvFUwEUGs/iY78O6BvqC+37MrbixVeWGYWu+4DS9IdQ53YXB0CYdH00u5kd2In9VJULpff88OAG7d0uHsqWyuaEducA7FQCV0E3h65AIdg+2KlhOnoeHLoOOrHA27XERLlN/fS0rjPYbPcBfcmcUCH3dUIxE1uDaqsggTZ560ToCtvnscPrfmtkCNLypuTX7dpxjce3Jdvp6KpaFyjeJuIGO1GY48JVULnEtlgADIrrZZpoFLoeNCEPgJyWvYBncjNm/gkgtPcJZYDoOh6yQGqV16nlRzp05ZIX3FRlYAjIhU01bQXNA7BBt1AW4ElpWlY8Lb2rPJBNzJh4ORYcpr5opwcT8amUZZI2y7eG76gvyU9Jze29DM2KgA+ZB3hzjpAXgaCGmieA3jDRgThQ6XZmF2mloydVQlZs53HR5Ki2OiaqCnOPlg9GecYAzEZh4dZtroLkQNYJkAf8sG2STzHg3fr6PoWuD2boQgnZC5BZgb5QapXDiQDQKscQqmGgqwiZFd5T8axVkDb8oM5Fy5qm2hGA9UGzAP1AqjJwIuCDNXfOGwILpKsQohc9rjV2Ad1aQvnwBpvCAwuY/YtkSkp6gp+PdrXxPNayD+At77jrF/BDU7K5pQDdwloARgH0pZd4DTUMPlhrDV2ur5oXIlRFUylUffAuoNwRoFs3b7RHRTaoIGRD8IrxEtNZPqX0OBCELWZPeSwhv2lHloCuJgu6xq3dD9DRVbUuCNkQBuLYOIKsqEB1gQ8CPQG6tHIVql7Ah13BbJENNaBZMbkHd7NSYOwCTUDKdd136FoevOTfwdrgpZu514HKF8byEoB88vKLcun8JMu9gvXBSb6TdeO+U0IOvuYC2cWP/GVTgzu3J6X/6Fq4N3bzSSBVGfqvLu6JF+HOnc3yzkpUYXnwYm7pczY4nLicuaNQfXKiS/5y08S7WL8KRG04Nvml2x+xW57q4OWfXncnFouXlP9LzW/Z8y7J/E9Nyf+aaimsDBzrdRb4eyCyTDJEVQYv8c4XWo3p+gqZBqQNXXrp/CX+EeyyizJ3q2mTGf79JA8lGpXipH/8O3EH1vIppbOX9uyhlNlTHkvIDwBY+tollW96bMQ1bu2O9jq6qtYFDTlwIlj+TXtuv+MSb0hsD054IcAvVv39qMI+u3gz+naiPPd48ApCOXaitylob3uXf/5tQQ940arr/lmVfjiL/UWe21kCVvdKtyoA1lmnwmY1n9utYL/sN2bdFlGGQA3SbBIaMtPyWphw8Xa+cw3hol+cU0hXxsqe2Ki4c2E5H2woVwGrKrfhhnx4LRPOAps3C/vXYBLiCtAa0EvsuOnGtO+ON5b48RkXPVVGwr2e8M6Snw9t5FImJa0Utm7nw2umZKolYO12aVM6C2hElkmyIv3OFYmoDwUjhOsrVDVAuiXWpvSPz6Axy3K7sy4nrmzlBb9zVQHzMt85H0vnLJ9SetHP7ZickTG7PZfyGCL5TTe3yn5+1xhV2fJR487uaNVRx1WdCxb2ghVp1dXhot85D+xUCo9e09jcyx+fjSqEs+zkfHRBuRnsLEMv8N4VlY7aDZlRkHWaf8T9w/ICiN79xnMmLeEHAFg+50qIZB4l47IcLJ88Yv/4tsEArKmyADe92HO2T1uB51zq5PE/hTnE8jbVa4QSAOwZ10bquLuI81zyRPRWO+akjrXmuVRLwhSPLJCon1QTzLz5niVzT8m+WDLjFFk4H2WHMqbMsW9EFfJRvso8z5vPoVnp+O18LJ84l/JYLL+IPjXL6Nzq7J4olkkmz6ze+bRZzGnTmfWtFT5yZ+SIbxZZPiETE1ZzAfPuJCZo+hEfKB7/018MPY9944Ot3zeLNJZpw0gQBIUugiAICl0EQbRBkQneHvSC5/cdtTd+MfT82T9+sEPXP75FY5kgCIIgCIIgCIIgCIIgCIIgCIIgCIIgCIIgCIIgCIIgCIIgiHcGYf5k0h8JgiAeAkf++vWxrwMA2GWR/Gjwxt9RESfut2bHu/V67+zkcPOh7sG71OkPUh/Z6MO5GNyHdd8JUY8qBh8bf5/NMDY2ETQrMfrudO09DFu9a53e65sjdPJdtHkDeP4d9X3hvq2Yu6d5wfL5fD4H8Lx9U2bi3ZuHzjk2ni8192F/+IEf2EKfiLoVAFjf24wHLJ/P5wV6Bt8x6w4UbZX7FfWQS8lRxRgodT+cEO8b88L8HwDywZEsdqJUCLLpny6xP9iB/INVxYk25gRLzct3dpLeUz9jwYPo6Ig/q+H9/n/W3le/hbGVOhufdh8nzcnv/Pg9fuHzCs0/lNXeFpMAGn+BsSKwMQsAYyf//VHa9jrnDuqj53/7D3dr4cTkK8sPQMmPTgL4ybdLv7L8Tpmx9M+Vt9dAofB2RWCDG99XD2UMjQUAXlnumATCKQnAm/z5fzyKxXp3lhuF0lLbNl97kL7yJgG9sgRe+hdtR+JTL1beG9PxAAHfoVUXP/Ekm8iA5S/onoJwH6v1Our1987r/wcOWHFofaThzvT29nYVQG1b9QwCYMGhP4bHJ23bPFxfP2Ifbcgh9zY2u2zSiih0vV5XUPW3d6k7pF+l4t7a0vNUk0yJeTUBoK4PNOORh+vi1sMZW6K2vb2tAL1dNQtsfvjc6XnKWAzB1pZU+h0Z181sV9mASG+MYgLc11rHmzjAyW3kO6InkwJGQ+XekztHXHXtL47gk9PQ0xOZp15V7uMLYH/+agU49h4JXf6HD5D+6tHCSjhl/t66zidLywAPbx+2XGN5obB/VYHpo/fRZkzJX7m/dbfdJqzYfzSuAMCNzoPK3TnSpe6Qfm/oRG9t7Z9rlim2VQ7ArWcPNOPRExYPa3BtmBWSnkLPhUAC+fDQ0OXnrMWEAl7/yTsyrpuZU9mnB6/rF+rth1nufvpmuYOc3CrfET2ZEtANlXvfzx5x1aVXvo9pCezOZjaW4o/vF+55YRgueQB4KIOjtC3uqw83NoL1B/m+x3dwbtfVQ5ftPfHbhJuqCCDYvvvvU9eVkfhdC7g7MnfASOSdFf6+nY5vJ3Sxs09iIgCyF+o9Q/HHmK6Sib990VG/LwcgD4ALADmwAIXeRClf8JIVoi8AwAolE+VtWwBYwf4ElDnC7UdTMv5jrwMZsIwAUCiJOL3Sa/ZkrrdYQGaK5cATxSNqXAD5UPpokoX35azsTCCTAXJgAhlhdn22a7/UYue+XFIDawgA4Hrb/lq8FcTvMz9llREA/EhIgBUCtNgpg0wm1ZEAy6EQwBesJICkcWKbIflP0wfyJRH3m7dGYTkAPgAewBdRb1ZUa8O8Oe/s70pZnf2SgLEVCwAfvM+pmRGAPdXqZCegD9iKXb1WOjO00h5xWkQHjBUSCppaXEQmiZq2bkvYprVsMgLLHIBgW8UzwHZeKAnbsbGFL4zdeRArxgpBelHTlRjF0WiMBkopGtfW2D66elMOzCWFlMy6qy+XXjXxcJvH/aSnjjWY+eCnrOe8E7WXnmSskIsPRZ60EziWIXDOzVsVomDgBlQ03vuOvK094gX/2P+xcPK/f/pl9pU7e381+rdv2Y8A++zWG4D3G68/U/joD4GOL58+82Nzhck+kz/9iGSX33wDZ7/0MrxnvvvI7++MFx6RUanfCT/f97dvAeCXTp958w1v8jOFx8s49gcbpi0A7NzwmUck/8z3vnLm8TLAJz9z5mN/B/CLnyn8ahk4+8XC42Wwc8NnTPnfD7r6GhJnv1jo/cFbANBx1vToPfNdPPHEE30/VrGA/OJnCkEZ/PLLl84EZWOLT78EAOg+9kPwz6y9hf7/sfIbLzkTnP3imY/9HbxLp0+/+QaeeOKJvje+jDOl73rPfPfT/5L3PbrxzHeBs18sBGWg448Lj9/65PcAIPvYywBcteyXS2c++kOgZ7zwSNePzYuaOz798md33wD77N9Nnul9Cej5cv70m288/r/zvuLLyD5TKP3gLX755UtnHi//9vDpH7wFnBtJ2Mn7ExT6fqzins7+VsV7euP3Hn/5d8InPtv7Us/46TffAMD/9f/yK31vvuFMB8Dam3/me185E5TBzo0USj94y/bbM2Ft+MjXXgL/kzffwGNfevl3PrnxJyj0/Xjvs383eSYoA+zccCG3/Ra/9Ldv4bEvvfyF3/D6PlkBrExW544vnyn9IPwj1ld86ZE/eBlfrY+Ven/wFryvnD7TV3wZ6PhjnCl9N+XkR6S1eOHxMv71m2/AVPxXv7byR6zvyZeNhZo80vHlMyXjde/S6TOPSBgrAEDPlwu5dWQnTp95ROKxyZo1CaKme8YLQdkZ1tTIn37zjVTZ/l0JAPw3XwK6T3wPfGT+I9EPWXtfOVP6B2Xdf+wPNn/vsUd+w+v7ZMVYrPelx770cqTYbw+f+cfT33VtRsP93MiZj/6QX3r50pneH7wVFX5svPCrZTOurbHx1fq/+rWX3Aj+1TL45e8986svAWbIfuLky/hq4w0+WTr95hveb7yET/9LM4Kzn/reZ3+sbD/pqWMKs4ufOXNM2nlrlPge8Mgfsb7iS86mAKyT7Ug/N3Lmoz+08856MgoG1jEAjn1t4xnT30ih9IO38NXGGzYYRMPXjbsv58+8+cYDXXWFt76vp+vQ27NsY025jzEjNdkNYPzRhZsm3cbGw7W9fmgZAAEDeAjGhrdVX6LUiVMmyg10ru8No//kWiUrAGbbApAtblX7AYyGMhsAA6fWZHcAdAdb1ewg+NDmYgY48eRqucds2VS9Dhwb2l33zlvtbI8MQEegkgKeOLVWzhYBDByvZosHZaFC7i7lHUO71W5g9Pj8zVHbWh31OsBQV9rsDDqGNhezg8C43j6RvnjYan3H12Q3wIfDan90wQ7ttrTneNUbBLuwu7Y3gprS9TrY+O7C7fMABkKZPfHkNh8EjvWurp8InJ10HfVaajsiAAxrANknqt7w8PbeMADoKup15UwHILI3RsNqtgivd2udD9p++fDugjcIACEHGAIgqAFRbz3GZieKW+WuwcR+UNdVXMrqPPLowo+ErlsTQWRPVr0i0H98QYVVs32o12Mnd59arRjbGP/GP82a7VS6jrqyFmryyMijCz8yvproXJMDkRUAPlxb7ypiPFyTfQC889t7I6ndJxveWuCxYeF8kCzbOnFCFS/JRzvXbw9G7mcX6ilbmIliFTtWrN1MZqTtcO/oXV3PmdHoFaPC3vndajYw49oZW2TNL3nHvuxP3AQIaoAAsidXKwEA8EF7WyNXCxG4ftJTxxTuOLVaDty8NUoAMI6LbWoVs6qa1ox8sSftBHZTxSqZDdwYg4ALBtHwdeMuXN8bfbAbxsbzjDW+BcypjJ6PP0atfOLqDC/COzW9OtshAKCjc2r+OV6EygCCA34IdNyZmvVEXKr31asKACtuzD2rUJuan2NFwDNtAUDpzvRUBcCnpqZZAD64MT+1NwQ29Pr0lCzBa8ysXAOGNhbm9gYB4AX50yvLKO0/O7ectXKZHs1a9fUrKiFgeGN+TgUAPj81rVzoykSbLz9U4EKGzIWgvsazU38leHFubdoLgPzrV3aewwtXAeD1640rywBQ2J9ZWS7BC16dupYKXa5a4+r8NA/gdU5N/bWIB5tJoD4xNSVLwOaz89c4bn0rvPI8Oo5Pr17LGiHZhZWpSh4o3VqYk0POTuFV3Gi5QcBPXvkWUJqbkj13ppY8kxpVr1+RznQAInvjU1NTqhd4dXpuOW/7zR6fXr1eAICQBeDIAEICcL09YWw2dGv6RiV+4OFG5c6V61EpqzML5len1J3n9RWbox+cmpK9YMWV1WvhLIB9a0bjZDa0sTBn3G/9GzF449k7z+srz1sLpT1iugEAL7gxP80GnRUAr3NqblZgZ2p+1hNAx4Y1SWyu49OrM7Fh4XyQLssyGeH2P1XAD1F1k8crvjo3nXHuh/folWsJW7z6LIBIsVLj6tTt5LbHDHd/f2HuGsA+PzUlS1HhrnD6hSU7riNjD96wDTpffuGF51xr2cAsBHMbC3MVAOgOl9xlWMsAtp/01DGF83cW5pajeQvv0SvXABjHRTZ1TraqmtaMfJEnbTCIpopRcooV3RgzS4E7U7OeiIavG3dX52a94IGGrvgiIUS7wwtoKIGuUGLf/OB4viGxL4uQGfBOLhDUAb2IXQRxqcZyaIbWEhqz2JAIpYjaAsCK69DzAFZVQwbw9BLCZR9clKGXPDANbIMHZehKvNlnxTVgzWZQbY82865VQsBbFVttU4YVG7C8y5cvmyjGekOAhyqM8hnBJvSrygsraMjAtNZiheIWsM7B9TI2UmddtXmFhhLIN6ReiCpVIDkAbEq9zKFngF231ss3lPmwqRry0SVUGFhxG1jOJO0UpXj6XKJgU9WB/YqudC5imyUTYcZ0AGJ7rypd8bG/BJRd0dyOwhYHgFAFYDoDlrrVamzGg3Vg8aCUg9WZCYVQpTLJusLBRAW7qcFnnMw7lxCaBYPxb+LyWTGrC2ehlEfibroaFTQqeWcFIL+vsFHBjEIDAggXscVSg5hpoJYwbOyDZNmzly8/bfO6QgK5GqIlOQ+Xsfst535gFbJNmtwoFuwgvNY6dTIaqEaDwBXu3VFYMW3FxrZ2iH25I+0GqH9sfN+ELqGgpVm5KhtuJaQP20966pjCog4to3mbViK2aXqkO6nbBYNoqphDYSWXGmPJqRmRayjsqOIDDl13njc9/luV/Oj2kxVoJZBpwGwR7f5C+qhysFAJiDrQkAgV4lK70dLb/SqditsCwCEBDWAZkAJ+QwFVDh5KoMaguQDAtTTJySi2SpgxiqhH8+9k14i6QxlQTTeuuvrGimtmpViJMo0S0PBDAEqY1lrvFUpAM/ghoCupu8eumvkT1IB95aKJQs2zgtSMO3Ry+R9CGCF3gDoDExVTLraTc+XoqPuGThkAtgDVkNCJeepMlzLAcmQ/7c4EVVjzQgrkqtyYOcLYjGtplmXt7+gbnXXz+FwHqgDTqul+oXGyGQw5IBIgYi2ZKW/2SNxNJgRQYc4KdiiqyK4NBd30DVxuksyVyAGHlGV9/eONChBUY9WNDM79Tb5HcpTnWCCjgZ+cOopHlq3xyApBNRoMsbHXmn257trqK3XOuCBgV2HHl+LqVQ7bT9PUyQGA6krO27QSsU3TI121Cx9Wo+SYDyuw1+dojCWnZrzd3Qa0FEeLSB1HXm+pQz8C6sMQvC/KTNQBKIaQg4cykMErUSomKmWvFRk3U32RaAsAS168GIQ2mvvmr/gnPj6jwHQf0MESsVghPaftP0MgKSAQbQAiXRp/YR0+ip1l4zDFgT4g/L4pJFgfID5sWmu9c9/VB84gWk5G1YyOoo5UNAndmkYzszOJ6v1THwDxIwBQHzczFoX4gqOSD/yE5UgTmVAs+TSkM51qsneNAWC5aAEV7PS5uKJOQ9Ry4GHrcGBaAdqKhwN01pUBKVstpZlQTLce5+hzHgv5+IxqO+JYJgOwlEfibgINoM5iK0TPGvnCttzUb2NvfEalDRv7ICprv/LARqFvACx4BaEKKl4R2BI66f72j0JYxVibBRnUh7HrXZh1PUZWYInHpJyxIzs4X96ODTPPh+1AKj/ZWwYAt+gCD4EaF8r0E6amjim83TG8GM/btBKxTdMjPSF1azCIxnzCfYkxlpiacUR8xYy4Bxy6joDoGAXwEwAQPzHShiz4mVa/xoVsVyol91gv8Fr6LjRa9NTRYaGWhiamlfku637zw0PibgJiYBj4eXs9dqVeA1CUGdQ40C9w56/tAPHayu7ob50XSFbjo73NtbnOgMGFLgjlJZ866e5und2jABrtrnbzR3iyKp7MzfbuHhGuWYb++G775xC82h/8gz7gIauDnxO3Oq8VL95oXYaEqncp3yb+C28U1i8N4982DU8EiEJXazf1pgfAhFGSTxz0SKZeOv/MtEwa1ps4MN+i17FbAbjey0CehjcK7Ap9F/enFWubTK4Un37O+Uh8qLVwq7GtLxOJsy3l95kAu7N3LrMIQLg28mEGnAll+mmkpo4pvC8HxGw0b9s9icVaRnpC6gP9H4Uulhpjh6yP2EMIXdhfNhuVZCARH6rKz7Lk4rC1FACc6N2qdt9Tbyui98LzCBfRZhGk7ibgieGtavaA+L4zb59GeQxAIFeB0F28G0tIL/iRztccMHZtte7iqupv2t57l4HEpopNnFyN72ptysQayl1b0X7Rd4802ZuPNOa9zyf7NdeAGmenXpTiQ/feo9V5Z2ZkTLa6Qw5k86+3C8CLkYGNf1sjprNQyiNN3bR6YSxYPehu4cbi8MhzCcOmfdAUuswFgrPLAPbRmAfUx9hd3N+kWLtm53Tv+WtHLHwwi08ao4TTo/21CrD8O27ddOwygKK0/aSmjimspyZ7Bpfvft1LqtokdVv/46Ax9nZ5oKFLhWup0CE0tAp+Jmssud9IlEoy9NNpZD/cellvisjaHVbQN9QXmnIggFlli7sKmL8zjcdPH54GrAPIBHItkRNsL3t02ZMHjds1d5cJTeE5qANgfrQr8ILvSB5Nm8Zac0JSY0vdv4dYbMmkvRmQPX5VHfu862M3UjNkGa1U8LPqAUOZQSX2ZK06A5vq6WLLlOBPfr+42+Y7yUrHBtY31BfahAJuLKQP6ibTNHJs4rT4yjIbaT8V9creU+KfYsPylA/aktcKYMeNg9Td3B8rdmBuOZxjn4oGwYdMYdbW2G18GYcLZm8x7k5d6q0Am7xYSQyznOsnPXVM4XBqsrTs5m0TzTZ1qsZSH+x/t+zUqTF2wNr08IX827rDeHjoSiSAM9YEMidUyH3dtpRZjJqnfO1zoSlvBS2tM2323DxUANa40Ky5krAJ/sMFDLbvaprwypUrV2xS32ZrFTv0qmQT8fyAfpkwOqpcNBBY8OqVK1eucbd7NNnX9Kxr28l9kTBdbG8BZEJkkjceEpnSkPkhZE60Mahmomk2tbf1TiVoY9wXrzzbpmaYMvCazdSLdFXZ1iO2G8UABGFL6NLLB08WbLAgYdiUD9oT3Lpy5cpz9ka2Ynf3jFVMH/glaL0QDwJXOPG8fouxE75M7sNdjfC6D6Ahh+2NiBtXrlxZ5nE/qaljCofXeDxvW32ZsGmsaiT1YXPN5csyjcOvuiZFHVQfQuiqx61JH2DFKiD5Kal5UGtbyqZjA3dHJO37EAHQFX2seQLoDU15kyfRQMiKqXEYogjwA7401tT14UHAD50iALQqAl7iDlSbseruXdY4wAeTZxLVmDA3W7LCjikFoM4FMATkwuh6qgFAdrVxb9B28h3pNkvCdHHVQZuGNddW028uoVReo8qjZyNEcjYGQFeoNAKwwdRoFUibqjWBwaJ7FSkzhqnbihqADIChNqvHdh4BAOm5RYZTvMt+3adFhlTTTYa9S8IlqLoBCqDuASjq1Bdem2euVUzLADgkL2IGgbOCzJvbnyI2dnratKQLK03fYl7qsIFbAvGDaG2mjj0WzduUIZI2VSx1lz45KFKeTDqGF+29y/gty611nb5HekHVgw5dNa8ImFfBbXtFZEUFqHkhQlWUbUvZCFVC9oJWOTQ9gq4r/SKxbg/ZELziNhpqGHywhp4AXVqFsgTYR7qVD6Er3YKNHBDeE10rH95gy7o4k/hCoLnK2Hu6qJwQfEI0+HmgOxIzSI/PSo8AH0bIhjAQby3y+bhagG4BVL2ADycv8A0WACcCPliDZoFJknORwbYXgPU3xcc+gA8lQlnuiGnNHBPOdKaetXeP8IrbALetcpFBtSOI3ocnczWEnp0zqd5C2Qc21ICWeZjvM3C74lc5BquzNwQWSGiWtlXIL1+a6G01Y6iG3WsajX+hCk3jwlko5RHbDYBdPoRsMbGEr3pFDASaFVtChm36WBEcMmHYyAepWJaJb0nD3HZy385vsPPoGXLuj2aWSK+HjGIyK7ynWp3Tb58mMIMgKuwFbCyA8iGcseO9YezLZDQ2w3XcDd8dNhhfhkMuTD/pqWMKjwjkdTxvE9ZmaZtyCKeqlVr5TteEJ1NTZVRkg0o8xtBm+HKRQbkjQLc+Yuh6oLmufXmuh+X+bwBo7J3rye9LIGQakMnQlShlnLo8eMm/g8oXxvJNQpeLk/x48ibMpH/8O9CVwUu88wXwi1V/H1j62iWVNzeRqoOXf3p9vTi5l3vlrgJWn5zokr/cdE/kMhBGLxc0156aWRdhffDyXifCSnGS5V50U/fsY1dSCQB+SfofWWzIgRPB8m9aA08Cr1031WRl4FivBPbZxZv2+29mTGkV/LX+q4t74kWEe2M3nwQQsku15/b3Llb5ydV0JrZ4Sfm/FKeJ5JOXXzyKs+Xg5OsVazoAkb3DH13e6/wOtrwJ77jrd+f2pPQfXbP7rVegWSPRW/QwRPn8pdvBi0DlqcuZpc8B297ln387ksno/Pdns7yzglBdrF9NzuRK0ffzddlsRr10fpIZ91n/louX/PUnklWthVIesd0AaMiBrN9ITL6dvXOP5V4M5diJ3pbxdclffwLsXI/fUAnDuh5SnD0LvHbdxSWzrJefNSNTPnY581PnflMiskWTYuuDk3wn2+Id75LM/xTA9y/uie9EhTf5xZvBd8y4dsaOL+yxL5NrHaEAZC/ZORAu/xbglvMhAtNPeuqYwtlJlX8tnrdxyL1Yv5q06bZ3+efftqpaqauDl396vdmTyakS/v0kD2U0xlLD0g7fkF2qPbfLLsrcLfUurrq02xtP7+UD8yRJOLOXD2cAhKoGSC3tklCpqFT0wrFV6YezWN0r3aok2gJwq+x3zpmKSgHzMt85r4AV6XeuSGzu5Y/PArfKfn7XWHVHZSR2lvzc5nK8CFX2to9KC7h5s7B/rU3aWcc7GmnWeuaCcmeJda4ozMl8blPa1tZZJ6CtzBrYn+X544vAwl6wYu921c1ezVZbu13alEA4y07aJxnMMlrLHLCyJzYqCBf94pwCGmXeCT2zlz/VdGPgzpKfD5diO60z85029zo71aR3lDdSOeVMB8DZG3p6T2woNMqFk4uw/YaLt/Odtt9tKISqblt0vZkMUdkPNirAhspsVAA0ZEbFMhmdw3I+2FDQS+x45AatwZ984cr03hCSZoRSwGY1nzPutP7dkf7OkvOf1ogslPKI7QaAnt/LH59PeDJcZLmNil7gvStNL+GzTe/fzHfOR4ZF3EPLTSIVZ8NM6LIrhmmV2b3m3G8805AZM8Rdn1ax/bLfmNUtU2frdj68BmBjT2zIqHBjkQUb0oxrZ+xIhoQvk3l6AQUs+rmdZWgNrHkCgNg219rA9JOeOqbwup/fXY7mbTS8jOOcTZ1iVlUr9Y7KyMiTUU03VQBgwe9cjcdYcoi6AdUo8040ZlludxbvDizarzLAvWfdeNf+k9kET0upxBNueQGA58z3cxOnzImW1pl5Lbj7GL/8zLwJPv7saoo4LR4JyPImQx2J15zmyCC9OXed5uIivslDRP/BN837ueZciakWnfCFLcCiDJhw77H3A2MzY6S8SN7gEbF+kVX8IEqixeomykdfy8zFpkPS3qY9ljdPbJoe/diRUedRb7HN7Ev8ec7KE71OJkjpDJgXsYtIrEf+DMDjf5o2I0tYKvaVdVQ80pyFUh5JeN3Wi53qJ2zPkiZxTbcMJNNDqmw6NSOSNo9elOdH49XYgiXHv5XT2Tnhq8jd/JuBHQTOCkYkc9DWT4ypIDVAhf3EUmMkkxxm6WGVsAwAP5+ct7H1UhUiJ/vJw9EvMLT35LFvpMcYazN8o3EHgngv87E/TYQuwkXAbwYfRLWOfeMdMRaNF+KhzFKAFWtkB+J+6SATEA+BmjdR7Qq+Q4YgKHQR7yf2y70FbEgyBEEQ7y8KfTkyQhOs74OZHCiRawmCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAjiMNioePuN8ImjHAIAb+Rd0is79GBUM00RHyToLakfjNDVd9T5PVA8uJHCUQ4BAHt7bzFheSTeZ34YXb33oNpdmyIodBHvW0rFhy+DNw5g7Dw5g6DQRcT0PHXY2abf83LroMngAUamdrvMExeaChXX763VseAdFZqg0EU8ZPxDd2I3ltsefpC/IcXaCcCbwkwhXL63VvPBOyo08X6D3k3/C0b9vSEGG9okXxC06vrFhhXsMscvCSADlhEAK5TMAqgQIAdeAlghgPnx2hy6TJ6du/fDZ5DJmBbM8UJq4VQoCQDw+wLTXbz+8c2ZfEkA8AUrBAATpjFzzgfvEwATLCOSl8zOpbii7Y8FKET5f7/p1fUZsEysK+/LxULDF7wUC2OLmAaMAdwppwlBqy7iPcDAkJ4HgOzTCK+qs0Vc/sm3MdaL4vPouPCfJ37+f13895dE8flzxfCqwtg/X++48J+exqlrgPeM0DcqALxLOIfv/D28SVF8HuCTAV5dito/O4Sd54CBYegXJNjFILxmz3iTIpyS6LmAcEqi5+P/MKRfkMeexuXGX1hpMDk3EAw9q84O4vKdv4xlHt5XAHouoPGscv0dO///TOCUabvnAoxo7hJ7GWfPvlhh54p4/Rr4ZKBv/M0lnMN3JICx1x8LtpXtkJ0r4tUlZJ+GXlkyBviP3qTY+X9NqzvP0YChVRfxnsAbqt0cAcDGdxdun4dWuq5wrHd1/UQAsAt1AAOhzJ54cpsP2mXaSE12Axg9Pn9zFAB0HfWaAkZDmQ2A7lOrlf746ja0uZgdhDe8tcaGgBPB9m33iNRo5/rtQXgXdtf5MICOwW02DF3XdeWkgcierHpF1JWuVxMyFxcB8OHdBW8w6o/xC1uy26ycLuyu7SUfHtNV1OsKJ55cLfcIZE+uVgInNIATp+IOs8Wtaj/YeLi21w9rgP7O7WwRYMNbCzTgKXQR7xG6wqkXfgSg4/j06rUsblTuXLmO0q2FOTkEeI9euQZ8fmqaXViZqti8tveJqzO8CF6cW5v2AgDhVdy4qoBPTU2zAGxoY2GOF6Pt4v7MynIJemN6/noX0Ls/9azvAtCrc9MZdIfTc9eyAiisTC37uPO8vvKskwYYnJqSvXj9+v6Va4ktaKMCIHt8evV6Ie6v4870jE3nbz47fy05NvUL6vUrEkMbC3N7g8htLMxVnNAAel+9qlyHpTvTUxV0dE7NP8eLxgCsuDH1fADw49OrMzRgKHQR7w1yDaUXAeQbCrvc5HJYcRtYzgBYhQQ2VUM+uoQKs1UW0FACXlhBQwaJplZVQwbgnUsIK9Hx4hawzhHOAtsMLFhD47odPOEydr+F4o7CJguAxhLW3YCKpNmRutI8ymySPrejsMXj/vQi9lURAPQMIl2SozUoQ1dyEApaJo43lkPXISuuQ88j35DYl0VjAC7K2JEA00CNBgyFLuK9QbAN7AIIakAIG7pEBahxQFcAoAyoHaBuQ1dYgVYCfghAJePDMiAFeKgA5bLkLJCANjU1wIQEts0p0wALykCoAmAfCF10jKRZB6otMqs8AARVQHMR9deQgMxEC602o1VLQDKg6fGLhPocEtAIagCkbwzAQ2kM0RoMifctlKZ/3yNes3//qQ9woQsFc1nSygWLj6eDgfowBOsDxIebV0Tg6HPtAGDo6gNnAFiOAwwqakhoU0AlQ6BQLdKANfWhl/5AKCDY6QNDuj95xlbJZFpVZboP6GAoP9lbTh6vxx3+d0gAEHUAihkDMCtvY298RtGAodBFvCdgws3G7u5E/BkF0LhLVW8UwE9aY6E3CuDn0ed+swTKjguEAOLJ70KXDYVpEtK0ssMGr4Ohv21/AOBNBG03CaMA9rGzdy6ziPYdmjAlfpKQNKNdxDz/zLSkIUOhi3hvsSkBuJk5DyC8S4XGkt1sNREupo6vKkCDjR9fvT2YXkCxo0rT0sPy51wZrT7W1B8ANnFydW/kAMlChNOj/bXKvXRo2FgcHqGHIyh0Ee8JdLRVa6zFB7F1hJ1RuNb2sNLp41sSADo6r6iOQWhEqzy7qtHJfWIiLq4d1vX6YCA1dtda+jPK8OA7krd7s44tuTt1qbfStsMOFosm7GrL5figV/aeErRl/GBAafr3PS50JRPuGuIIFQ9YNIWp464pv6GijyzZgCnQHBHU4QLsqyFAinR/LAo3PGy7fNKRZOF1v70ZNAIAUJlE6IqrYYMFNGIodBHvjdCVB7IAZFccULQ6wgytJZ2fiDRh6kacTjw+wWxoGLarGQ9AUctewBMyvY900qQOxt/B1stZQObS/fEALKgesBtVAEJWTB1Lymk7DBEAXZA+wIq2Lc0DsMGopC9o0FDoIh4+0hP8PIBtLwDrB8AhtOwD+F3eC9rg54HuwOw6c3GkCNUw0BUFiUqPAB9GnQt2AdCVPmTtuQY7j54hyKxAX2KZpFnAYmksdS4EeuLX4WzxIqodAVhf3B8bQVZUbBPobxI3x0QoS0DXIMaBIEwL7TrUlX7BR7DtFV1bQIhe9AjgWBEckt4U9sGAcl3vezb5pdsfUcD+3sUqP7mKbe/yz7+9WLyk/F9aOrRiWClOstyLJv49efnFv3ehZ+n8JMu9EgeZS9L/yGKDT9w+CaBcvMQ3zWP5DfnY5cxPsTZ46Wbu9US76mL9qpMmXspd/MhfJpY7+3K4snN7UvqPrkX9hR2X/IYEgHBv7OaTTRF6cPL1ytLXLqn8a8he2su9YoSOIqbrsFyc5MfR2DvXk9+3J8PlwYu5pc+BnevxGwril2jQ0KqLePg0lvjxGQB6Zi9/ag1oyIzCnSU/Hy5BtzzXqVR8cE7mc5tmeq+zTgVtzmOzms/tVmzqCPuzPH98EY2lrlMzAG5Jf/+6bWBaZXavYX+R53aWzJ5Oa0AvseORNArmYENm6qgnxFnyEC7ezneuxf2Fi775IjnCRb84p1KCr6mcwq2yn9+tYNHP7Sxboa3MUYe3yn7nHMKZvXw4A6vrqsptVID9m/nO+YPetkgQxLtNXpjEEMsLwL3qz7xxUABAxmWOhE2Fu88sHz2W7gdtjwsA8E27eWFK5OMkEzfLL9MXiyrkRSRNdJDlASYSrYq4ZdPfh/4MfpCQh5kK7tFUP9f8f1fcbhud+uYPz4s4HcZzRvlYJIIgiAfDh/6MbEDQhpEgCApdBEEQFLoIog2aTEAcHY9MQLxXQtfP3iAjEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEARBEATxHoTl0p/9t9keJ5O+S44TZAPiFxlvMh15LhUpdL2LfPTp+/Xbxcv37yl2+W16GR2XHszou3yfAfjcU+/yheJi0HIo8/CGTYv6J75u/l4c/GDOk2PPtB7zw6bII9/mmKZw5ExZ6ty6qzEzTVOXBwCwfffW+wMUK/c9ETOAd2Hlnl3NJm4t23/5KaFrCkDXowc2yHu3FMBKu7JFf5X4VMjV1oGBT862baTQuR7X+3jz3P3CdPzhoBaSjH3k2sGO2662bE5ebzqSHRebM2lXG+chtEqyfK62ntALvojOFnK1bas5KwTh4tHsH+nVor5v42jm5230Ka4DYKUtBV8ANQUgD/uPh7fX09V217IDRiXPRKYWm7ZELkwZmYeq1eTo6tVrKvUPCl13jVyXBEpT9xodTkwCwL+9u++LO0vH3558+Y17H3H5X1luK/St5wH0f/zbB1YcfXQZ4KMvHtr82SGg9znkP9H+9Gjn9sFjzy8kPhzUQpLSnYMMMxmgb2XpbqYYEao7MO71njGutjuYf/4/TYmxXuDUNQCPnTf/GCgCuPOXVtOffsvGoyEg99yR7H8UvdpMydF1AGy0UcFYAIRTEnwSQDhffliTo2MSwMbsvY9K7xmB0gvG7EHNGvnkdRvJWk2efRro/QsA3oRA8T/QhvFIV0hRq/L725fpIwQRsbZdeVvy6fuqpYCeodbDXQCAgzd/ocoB6NKHRvITQ6ir7FNQADrGW0eW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";

/* ---- HPD forms SF LL1-06 / SF LL1-07, printed on top of the official forms ---- */
const HPD_PRINT_CSS = `
  @page { size: 11in 8.5in; margin: 0; }
  html, body { margin: 0; padding: 0; }
  .hpd-page { position: relative; width: 792pt; height: 612pt; overflow: hidden; page-break-after: always; background-repeat: no-repeat; background-size: 792pt 612pt; font-family: Helvetica, Arial, sans-serif; color: #000;
    -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .hpd-page:last-child { page-break-after: auto; }
  .hpd-t { position: absolute; font-size: 9.5pt; line-height: 1; white-space: nowrap; }
  .hpd-x { position: absolute; font-size: 11pt; font-weight: bold; line-height: 1; text-align: center; }
`;
function hpdDate(iso) { if (!iso) return ""; const [y, m, d] = String(iso).split("-"); return `${m}/${d}/${y}`; }
// Text whose baseline sits just above a form line at y (pts from the top).
function hpdText(x, lineY, text, w) {
  if (!text) return "";
  const t = escHtml(text);
  const len = String(text).length;
  const size = w && len * 5 > w ? Math.max(6.5, 9.5 * (w / (len * 5))) : 9.5;
  return `<div class="hpd-t" style="left:${x}pt;top:${lineY - size - 2}pt;font-size:${size}pt">${t}</div>`;
}
// Text vertically centred in a table cell.
function hpdCell(x0, x1, y0, y1, text, align) {
  if (!text) return "";
  const len = String(text).length;
  const w = x1 - x0 - 4;
  const size = len * 5 > w ? Math.max(6, 9.5 * (w / (len * 5))) : 9.5;
  return `<div class="hpd-t" style="left:${x0 + 2}pt;width:${w}pt;top:${(y0 + y1) / 2 - size / 2}pt;font-size:${size}pt;text-align:${align || "left"}">${escHtml(text)}</div>`;
}
function hpdX(x0, x1, y0, y1) { return `<div class="hpd-x" style="left:${x0}pt;width:${x1 - x0}pt;top:${(y0 + y1) / 2 - 6}pt">X</div>`; }
function hpdHeader(year, page, pages) {
  return `<div class="hpd-t" style="left:699pt;top:22pt;width:74pt;text-align:center">${escHtml(year)}</div>`
    + `<div class="hpd-t" style="left:703pt;top:37pt;width:30pt;text-align:center">${page}</div>`
    + `<div class="hpd-t" style="left:748pt;top:37pt;width:25pt;text-align:center">${pages}</div>`;
}
// info = { year, address, ownerAgent, borough, regId }; rows = [{ apt, exempt, ended, inspected, by, inspector, hazards }]
function buildLL106Pages(info, rows) {
  const PER = 12, top0 = 272, rowH = (522 - 272) / PER;
  const cols = { apt: [18, 57], no: [57, 107], hpd: [107, 186], tested: [186, 281], ended: [281, 366], insp: [366, 440], owner: [440, 489], third: [489, 533], name: [533, 694], yes: [694, 733], nope: [733, 772] };
  const pages = Math.max(1, Math.ceil(rows.length / PER));
  let html = "";
  for (let p = 0; p < pages; p++) {
    let body = hpdHeader(info.year, p + 1, pages)
      + hpdText(20, 115, info.address, 418) + hpdText(492, 115, info.ownerAgent, 278)
      + hpdText(20, 152, info.borough, 164) + hpdText(284, 152, info.regId, 155);
    rows.slice(p * PER, p * PER + PER).forEach((r, i) => {
      const y0 = top0 + i * rowH, y1 = y0 + rowH;
      body += hpdCell(...cols.apt, y0, y1, r.apt, "center");
      if (r.exempt === "hpd") body += hpdX(...cols.hpd, y0, y1);
      else if (r.exempt === "tested") body += hpdX(...cols.tested, y0, y1);
      else body += hpdX(...cols.no, y0, y1);
      body += hpdCell(...cols.ended, y0, y1, hpdDate(r.ended), "center") + hpdCell(...cols.insp, y0, y1, hpdDate(r.inspected), "center");
      if (r.by === "third") body += hpdX(...cols.third, y0, y1); else body += hpdX(...cols.owner, y0, y1);
      body += hpdCell(...cols.name, y0, y1, r.inspector);
      if (r.hazards === "yes") body += hpdX(...cols.yes, y0, y1); else if (r.hazards === "no") body += hpdX(...cols.nope, y0, y1);
    });
    html += `<div class="hpd-page" style="background-image:url(${HPD_LL106_BG})">${body}</div>`;
  }
  return html;
}
// info = { year, address, apt, inspector, borough, regId }; rooms = [{ name, peeling, teeth, subsurface, doorFrame, windowFrame, impact, floors }]
function buildLL107Pages(info, rooms) {
  const PER = 5, top0 = 372, rowH = (477 - 372) / PER;
  const room = [18, 182];
  const pairs = { peeling: [182, 231, 281], teeth: [281, 324, 367], subsurface: [367, 410, 454], doorFrame: [454, 494, 533], windowFrame: [533, 572, 611], impact: [611, 651, 690], floors: [690, 730, 771] };
  const pages = Math.max(1, Math.ceil(rooms.length / PER));
  let html = "";
  for (let p = 0; p < pages; p++) {
    let body = hpdHeader(info.year, p + 1, pages)
      + hpdText(20, 140, info.address, 258) + hpdText(327, 140, info.apt, 126) + hpdText(497, 140, info.inspector, 272)
      + hpdText(20, 177, info.borough, 160) + hpdText(233, 177, info.regId, 133);
    rooms.slice(p * PER, p * PER + PER).forEach((r, i) => {
      const y0 = top0 + i * rowH, y1 = y0 + rowH;
      body += hpdCell(...room, y0, y1, r.name);
      Object.entries(pairs).forEach(([k, [a, m, b]]) => {
        if (r[k] === "yes") body += hpdX(a, m, y0, y1); else if (r[k] === "no") body += hpdX(m, b, y0, y1);
      });
    });
    html += `<div class="hpd-page" style="background-image:url(${HPD_LL107_BG})">${body}</div>`;
  }
  return html;
}

function openHpdPrint(title, pagesHtml) {
  const w = window.open("", "_blank");
  if (!w) { window.alert("Your browser blocked the print window — allow pop-ups for this site and try again."); return; }
  w.document.write(`<!doctype html><html><head><title>${escHtml(title)}</title><style>${HPD_PRINT_CSS}</style></head><body>${pagesHtml}<script>window.onload=function(){setTimeout(function(){window.print();},400);};</script></body></html>`);
  w.document.close();
}
