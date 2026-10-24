// ==UserScript==
// @name         Peculiar Line Flush
// @namespace    peculiar-line-flush
// @version      2.9.22
// @updateURL    https://raw.githubusercontent.com/rigoabell/peculiar-line-flush/main/peculiar-line-flush.user.js
// @downloadURL  https://raw.githubusercontent.com/rigoabell/peculiar-line-flush/main/peculiar-line-flush.user.js
// @description  Scan containers in Tampermonkey. Looks up Peculiar Inventory and saves each operator’s daily dashboard to the Line Flush Records folder.
// @author       abelrigm
// @match        https://peculiar-inventory-na.aka.corp.amazon.com/*
// @match        https://aft-aging-inventory.na.aftx.amazonoperations.app/*
// @match        https://aft-aging-inventory.na.aftx.amazonoperations.app/MEM3/*
// @match        https://*.cloudfront.net/idp-selection*
// @match        https://*.federated-auth.aft.a2z.com/*
// @match        https://qifcr.na.aftx.amazonoperations.app/*
// @connect      amazon-my.sharepoint.com
// @connect      raw.githubusercontent.com
// @run-at       document-start
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addValueChangeListener
// @grant        GM_removeValueChangeListener
// @grant        GM_openInTab
// @grant        GM_xmlhttpRequest
// @grant        GM_setClipboard
// @grant        GM_registerMenuCommand
// @grant        unsafeWindow
// ==/UserScript==

/*
  Firefox + Tampermonkey — multi-user Line Flush

  Each operator scans in the panel. Each person gets their own daily file in:
    Documents / IOL Assignments / Line Flush Records
  Example: LineFlush_2026-08-28_abelrigm.xlsx
  The file matches the Dashboard screen: totals, each location, then each container with its parent, age, and status.

  1. Install from the GitHub raw link. Do not paste this file into Tampermonkey.
  2. Open AFT Aging Inventory MEM3, finish Midway, leave that tab open.
  3. Open the SOS folder once in Firefox so Microsoft 365 cookies work.
  4. Scan into the panel box and press Enter. If the dropdown is still Select location, pick a location and that scan continues.
  5. Do not open today’s file in Excel while you are still scanning (it locks).
     If it is open, the panel names that file and waits on RETRY until the save works.
*/

(function () {
  "use strict";

  const VERSION = "2.9.22";
  const UPDATE_URL =
    "https://raw.githubusercontent.com/rigoabell/peculiar-line-flush/main/peculiar-line-flush.user.js";
  const SCRIPT_NAME = "Peculiar Line Flush";
  const OLD_SCRIPT_NAME = "Peculiar Inventory Scan Panel";
  const PECULIAR_URL = "https://aft-aging-inventory.na.aftx.amazonoperations.app/MEM3/overview";
  const SP_SITE = "https://amazon-my.sharepoint.com/personal/abelrigm_amazon_com";
  const SOS_FOLDER =
    "/personal/abelrigm_amazon_com/Documents/IOL Assignments/Line Flush Records";
  const SOS_FOLDER_URL =
    "https://amazon-my.sharepoint.com/personal/abelrigm_amazon_com/Documents/IOL%20Assignments/Line%20Flush%20Records";
  const SHARE_FOLDER_URL =
    "https://amazon-my.sharepoint.com/:f:/p/abelrigm/IgC3aqxdaBJeQoeCFtcSkjRsAcsBA4G40s7Ra0q3aT1adac?e=8GQxFZ";
  const OLD_SOS_FOLDER =
    "/personal/abelrigm_amazon_com/Documents/IOL Assignments/SOS Report Upload Folder";
  const HEARTBEAT_MS = 4000;
  const HEARTBEAT_STALE_MS = 12000;
  const LOOKUP_TIMEOUT_MS = 20000;
  const KEYS = {
    heartbeat: "plf.heartbeat",
    request: "plf.request",
    response: "plf.response",
    session: "plf.sosSession",
    history: "plf.scanHistory",
    locations: "plf.locations",
    location: "plf.location",
    locationIdle: "plf.locationIdle",
    panelSize: "plf.panelSize",
    panelOpen: "plf.panelOpen",
  };
  const IDLE_LOCATION_MS = 10 * 60 * 1000;
  const TOTAL_SHEET = "TOTAL";
  const PRESET_LOCATIONS = [
    "L1 Odd",
    "L1 Even",
    "L2 Odd",
    "L2 Even",
    "L3 Odd",
    "L3 Even",
    "L4 Odd",
    "L4 Even",
    "L7 AR",
    "L8 Odd",
    "L8 Even",
    "AR EOL",
    "BIGs EOL",
    "EOL 1",
    "EOL 2",
    "EOL 3",
    "EOL 4",
    "EOL 5",
    "EOL 6",
  ];
  const OTHER_LOCATION = "__other__";
  const HEADERS = [
    "Container ID",
    "Scanned At",
    "Scanner",
    "Container Type",
    "Total Units",
    "Oldest Units Age",
    "Parent Container",
    "Parent Type",
    "Grandparent Container",
    "Line Age",
    "FcSku",
    "Quantity",
    "Owner",
    "Disposition",
    "Status",
    "Notes",
  ];

  const host = location.hostname;
  const isPeculiar = /peculiar-inventory|aft-aging-inventory/i.test(host);
  const isLoginChooser = /idp-selection|federated-auth\.aft\.a2z\.com|qifcr\.na\.aftx\.amazonoperations\.app/i.test(location.href);
  const pageWindow = typeof unsafeWindow !== "undefined" ? unsafeWindow : window;
  let isTop = true;
  try {
    isTop = pageWindow.top === pageWindow.self;
  } catch {
    isTop = false;
  }
  let onScanKey = null;
  let restoreScanFocus = null;
  let seenScanStamp = -1;
  let seenScanKey = "";

  function deliverScanKey(e) {
    if (!e || (e.timeStamp === seenScanStamp && e.key === seenScanKey)) return;
    seenScanStamp = e.timeStamp;
    seenScanKey = e.key;
    if (typeof onScanKey === "function") onScanKey(e);
  }

  document.addEventListener("keydown", deliverScanKey, true);
  try {
    pageWindow.addEventListener("keydown", deliverScanKey, true);
  } catch {
    /* The page window is not reachable from this frame. */
  }
  let updateNotice = "";
  let updateNoticeUntil = 0;

  let lastError = "";
  let lastDebug = null;
  let focusNote = "";
  const recent = [];
  const sosCache = { alias: "", day: "", name: "", book: null, records: [] };
  let selectedLocation = "";
  let customLocations = [];
  let lastScanAt = 0;
  let locationPromptOpen = false;
  let digestVal = "";
  let digestExp = 0;
  let scanQueueLen = 0;
  let fileOk = false;

  function clip(s, n) {
    const t = String(s ?? "");
    return t.length > n ? t.slice(0, n) + "…" : t;
  }

  function nowIso() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  }

  function dayStamp() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  }

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  function nonce() {
    return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function cellText(value) {
    if (value == null || value === "") return "";
    return String(value).replace(/\u00a0/g, " ").trim();
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function csvCell(value) {
    const s = String(value ?? "");
    if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  function toCsv(rows) {
    return "\uFEFF" + rows.map((row) => row.map(csvCell).join(",")).join("\r\n") + "\r\n";
  }

  function parseCsv(text) {
    const raw = String(text || "").replace(/^\uFEFF/, "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    if (!raw.trim()) return [];
    const rows = [];
    let row = [];
    let cell = "";
    let inQuotes = false;
    for (let i = 0; i < raw.length; i++) {
      const ch = raw[i];
      if (inQuotes) {
        if (ch === '"' && raw[i + 1] === '"') {
          cell += '"';
          i += 1;
        } else if (ch === '"') inQuotes = false;
        else cell += ch;
      } else if (ch === '"') inQuotes = true;
      else if (ch === ",") {
        row.push(cell);
        cell = "";
      } else if (ch === "\n") {
        row.push(cell);
        rows.push(row);
        row = [];
        cell = "";
      } else cell += ch;
    }
    if (cell || row.length) {
      row.push(cell);
      rows.push(row);
    }
    return rows.filter((r) => r.some((c) => String(c).trim()));
  }

  function parseAgeDays(raw) {
    const s = cellText(raw).toLowerCase().replace(/,/g, " ");
    if (!s) return null;
    let days = 0;
    let hit = false;
    const d = s.match(/(\d+(?:\.\d+)?)\s*(?:days?|d)\b/);
    const h = s.match(/(\d+(?:\.\d+)?)\s*(?:hours?|hrs?|h)\b/);
    const m = s.match(/(\d+(?:\.\d+)?)\s*(?:minutes?|mins?|m)\b/);
    if (d) {
      days += Number(d[1]);
      hit = true;
    }
    if (h) {
      days += Number(h[1]) / 24;
      hit = true;
    }
    if (m) {
      days += Number(m[1]) / 1440;
      hit = true;
    }
    if (hit) return days;
    const n = s.match(/^(\d+(?:\.\d+)?)/);
    return n ? Number(n[1]) : null;
  }

  function hslToRgb(h, s, l) {
    const a = s * Math.min(l, 1 - l);
    const f = (n) => {
      const k = (n + h / 30) % 12;
      return l - a * Math.max(Math.min(k - 3, 9 - k, 1), -1);
    };
    return [Math.round(255 * f(0)), Math.round(255 * f(8)), Math.round(255 * f(4))];
  }

  function colorChannels(color) {
    const raw = String(color || "").trim();
    const hex = raw.match(/^#([0-9a-f]{6})$/i);
    if (hex) {
      const n = parseInt(hex[1], 16);
      return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
    }
    const hsl = raw.match(/hsl\(\s*([\d.]+)\s*,\s*([\d.]+)%\s*,\s*([\d.]+)%\s*\)/i);
    if (hsl) return hslToRgb(Number(hsl[1]), Number(hsl[2]) / 100, Number(hsl[3]) / 100);
    return [255, 255, 255];
  }

  function relativeLuminance(color) {
    const linear = colorChannels(color).map((channel) => {
      const s = channel / 255;
      return s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
    });
    return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2];
  }

  function contrastInk(bg) {
    return relativeLuminance(bg) < 0.36 ? "#ffffff" : "#111111";
  }

  function ageTone(days) {
    if (days == null || Number.isNaN(days)) {
      return { bg: "#ffffff", fg: "#111111", border: "#d8dee8", label: "—", pct: 0 };
    }
    const age = Math.max(0, days);
    const h = Math.round(120 * (1 - Math.min(age, 5) / 5));
    const s = 78;
    const l = h > 35 && h < 90 ? 46 : 34;
    const bg = `hsl(${h}, ${s}%, ${l}%)`;
    return {
      bg,
      fg: contrastInk(bg),
      border: `hsl(${h}, ${s}%, ${Math.max(18, l - 14)}%)`,
      label: age >= 5 ? "ALERT" : age >= 2 ? "CAUTION" : "GOOD",
      pct: Math.max(2, Math.min(100, (Math.min(age, 5) / 5) * 100)),
    };
  }

  const DASH_COLORS = {
    Alert: "#b42318",
    Caution: "#f5c518",
    Good: "#0d7a3f",
    "No age": "#98a2b3",
  };

  function ageStatus(ageRaw) {
    const tone = ageTone(parseAgeDays(ageRaw));
    if (tone.label === "ALERT") return "Alert";
    if (tone.label === "CAUTION") return "Caution";
    if (tone.label === "GOOD") return "Good";
    return "No age";
  }

  function unitCount(value) {
    const n = Number(String(value ?? "").replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : 0;
  }

  function formatUnits(n) {
    if (!Number.isFinite(n)) return "0";
    return Number.isInteger(n) ? String(n) : String(Math.round(n * 100) / 100);
  }

  function scanRecord(containerId, result, location) {
    const notFound = result?.error === "NOT_FOUND" || result?.error === "NOT FOUND";
    const ok = Boolean(result?.ok) && !notFound;
    let status = "Error";
    if (ok) status = ageStatus(result?.oldestAge);
    else if (notFound) status = "Not found";
    return {
      location: normalizeLocation(location) || "Unknown",
      parent: cellText(result?.parentContainer) || "(no parent)",
      container: cellText(result?.containerId) || cellText(containerId),
      age: ok ? cellText(result?.oldestAge) : "",
      status,
      units: ok ? cellText(result?.totalUnits) : "",
    };
  }

  function historyRecord(row) {
    if (row?.record?.container) return row.record;
    const result = row?.result || {};
    const notFound = result.error === "NOT_FOUND" || row?.status === "NOT FOUND" || row?.status === "NOT_FOUND";
    const ok = !notFound && (result.ok === true || row?.status === "OK");
    return scanRecord(
      row?.id,
      {
        ok,
        error: notFound ? "NOT_FOUND" : ok ? "" : result.error || row?.status || "ERROR",
        containerId: result.containerId || row?.id,
        oldestAge: result.oldestAge || row?.age || "",
        parentContainer: result.parentContainer || "",
        totalUnits: result.totalUnits || "",
      },
      row?.location || ""
    );
  }

  function emptyTally() {
    return { containers: 0, units: 0, alert: 0, caution: 0, good: 0, noAge: 0, miss: 0, err: 0 };
  }

  function addTally(tally, rec) {
    tally.containers += 1;
    tally.units += unitCount(rec.units);
    if (rec.status === "Alert") tally.alert += 1;
    else if (rec.status === "Caution") tally.caution += 1;
    else if (rec.status === "Good") tally.good += 1;
    else if (rec.status === "No age") tally.noAge += 1;
    else if (rec.status === "Not found") tally.miss += 1;
    else tally.err += 1;
  }

  function summarizeRecords(records) {
    const all = emptyTally();
    const locations = [];
    const byName = new Map();
    for (const rec of records || []) {
      addTally(all, rec);
      const name = rec.location || "Unknown";
      let loc = byName.get(name);
      if (!loc) {
        loc = { name, ...emptyTally() };
        byName.set(name, loc);
        locations.push(loc);
      }
      addTally(loc, rec);
    }
    return { ...all, locations };
  }

  function tallyCells(name, tally) {
    return [
      name,
      formatUnits(tally.containers),
      formatUnits(tally.units),
      formatUnits(tally.alert),
      formatUnits(tally.caution),
      formatUnits(tally.good),
      formatUnits(tally.noAge),
      formatUnits(tally.miss),
      formatUnits(tally.err),
    ];
  }

  function dashboardSheetRows(records) {
    const summary = summarizeRecords(records);
    const rows = [
      ["Today", "Containers", "Units", "Alert", "Caution", "Good", "No age", "Not found", "Errors"],
      tallyCells("Today", summary),
      [],
      ["Location", "Containers", "Units", "Alert", "Caution", "Good", "No age", "Not found", "Errors"],
    ];
    for (const loc of summary.locations) rows.push(tallyCells(loc.name, loc));
    rows.push([]);
    rows.push(["Location", "Parent", "Container", "Age", "Status", "Units"]);
    for (const rec of records || []) {
      rows.push([rec.location, rec.parent, rec.container, rec.age, rec.status, rec.units]);
    }
    return rows;
  }

  function recordsFromLegacy(book) {
    const records = [];
    for (const sheet of book?.sheets || []) {
      if (String(sheet.name || "").toLowerCase() === TOTAL_SHEET.toLowerCase()) continue;
      const header = sheet.rows?.[0] || [];
      if (cellText(header[0]) !== "Container ID") continue;
      const col = (name) => header.findIndex((cell) => cellText(cell) === name);
      const idCol = col("Container ID");
      const unitCol = col("Total Units");
      const ageCol = col("Oldest Units Age");
      const parentCol = col("Parent Container");
      const statusCol = col("Status");
      for (const row of sheet.rows.slice(1)) {
        const container = cellText(row[idCol]);
        const rawStatus = cellText(row[statusCol]);
        if (!container || !rawStatus) continue;
        let status = "Error";
        if (rawStatus === "OK") status = ageStatus(row[ageCol]);
        else if (rawStatus === "NOT FOUND" || rawStatus === "NOT_FOUND") status = "Not found";
        records.push({
          location: sheet.name || "Unknown",
          parent: cellText(row[parentCol]) || "(no parent)",
          container,
          age: rawStatus === "OK" ? cellText(row[ageCol]) : "",
          status,
          units: rawStatus === "OK" ? cellText(row[unitCol]) : "",
        });
      }
    }
    return records;
  }

  function recordsFromBook(book) {
    for (const sheet of book?.sheets || []) {
      const rows = sheet.rows || [];
      const start = rows.findIndex((row) => row[0] === "Location" && row[1] === "Parent" && row[2] === "Container");
      if (start < 0) continue;
      const records = [];
      for (const row of rows.slice(start + 1)) {
        const container = cellText(row[2]);
        if (!container) continue;
        records.push({
          location: cellText(row[0]) || "Unknown",
          parent: cellText(row[1]) || "(no parent)",
          container,
          age: cellText(row[3]),
          status: cellText(row[4]) || "Error",
          units: cellText(row[5]),
        });
      }
      return records;
    }
    return recordsFromLegacy(book);
  }

  function sameDashboardRecord(a, b) {
    const locOk = a.location === b.location || a.location === "Unknown" || b.location === "Unknown";
    return (
      locOk &&
      a.parent === b.parent &&
      a.container === b.container &&
      a.age === b.age &&
      a.status === b.status &&
      String(a.units) === String(b.units)
    );
  }

  function mergeRecords(fileRecords, localRecords) {
    const pool = (fileRecords || []).map((rec) => ({ rec, used: false }));
    const extras = [];
    for (const local of localRecords || []) {
      const hit = pool.find((item) => !item.used && sameDashboardRecord(item.rec, local));
      if (hit) hit.used = true;
      else extras.push(local);
    }
    return [...(fileRecords || []), ...extras];
  }

  function statusClass(status) {
    if (status === "Alert") return "tag-alert";
    if (status === "Caution") return "tag-caution";
    if (status === "Good") return "tag-good";
    if (status === "Not found") return "tag-miss";
    if (status === "No age") return "tag-none";
    return "tag-err";
  }

  function pieChart(slices) {
    const parts = slices.filter((slice) => slice.count > 0);
    const total = parts.reduce((sum, slice) => sum + slice.count, 0);
    const label = parts.map((slice) => `${slice.name} ${slice.count}`).join(", ") || "No age chart";
    if (!total) return `<div class="empty">No found containers yet</div>`;
    if (parts.length === 1) {
      return `<svg viewBox="0 0 180 180" role="img" aria-label="${escapeHtml(label)}"><circle cx="90" cy="90" r="70" fill="${parts[0].color}"></circle></svg>`;
    }
    let angle = -Math.PI / 2;
    const arcs = parts.map((part) => {
      const sweep = (part.count / total) * Math.PI * 2;
      const start = angle;
      angle += sweep;
      const x1 = 90 + 70 * Math.cos(start);
      const y1 = 90 + 70 * Math.sin(start);
      const x2 = 90 + 70 * Math.cos(angle);
      const y2 = 90 + 70 * Math.sin(angle);
      const large = sweep > Math.PI ? 1 : 0;
      return `<path d="M 90 90 L ${x1.toFixed(2)} ${y1.toFixed(2)} A 70 70 0 ${large} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} Z" fill="${part.color}"></path>`;
    });
    return `<svg viewBox="0 0 180 180" role="img" aria-label="${escapeHtml(label)}">${arcs.join("")}</svg>`;
  }

  function barChart(rows, fillClass) {
    if (!rows.length) return `<div class="empty">None yet</div>`;
    const max = Math.max(...rows.map((row) => row.count), 1);
    return rows
      .map((row) => {
        const width = Math.max(2, (row.count / max) * 100);
        return `<div class="barrow"><div class="barlabel" title="${escapeHtml(row.label)}">${escapeHtml(row.label)}</div><div class="bartrack"><div class="barfill ${fillClass || ""}" style="width:${width}%"></div></div><div class="barcount">${row.count}</div></div>`;
      })
      .join("");
  }

  function tallyBoxes(tally) {
    const boxes = [
      ["Containers", tally.containers],
      ["Units", formatUnits(tally.units)],
      ["Alert", tally.alert],
      ["Caution", tally.caution],
      ["Good", tally.good],
      ["No age", tally.noAge],
      ["Not found", tally.miss],
      ["Errors", tally.err],
    ];
    return `<div class="dash-totals">${boxes
      .map(([label, value]) => `<div><b>${escapeHtml(value)}</b><span>${escapeHtml(label)}</span></div>`)
      .join("")}</div>`;
  }

  function dashboardHtml(records, who) {
    const list = records || [];
    const summary = summarizeRecords(list);
    const whoLine = `${who || "Operator"} · ${dayStamp()}`;
    if (!list.length) {
      return `<div class="dash-who">${escapeHtml(whoLine)}</div><div class="empty">No scans yet today.</div>`;
    }
    const slices = [
      { name: "Alert", count: summary.alert, color: DASH_COLORS.Alert },
      { name: "Caution", count: summary.caution, color: DASH_COLORS.Caution },
      { name: "Good", count: summary.good, color: DASH_COLORS.Good },
    ];
    if (summary.noAge) slices.push({ name: "No age", count: summary.noAge, color: DASH_COLORS["No age"] });
    const legend = slices
      .map(
        (slice) =>
          `<div class="dash-key"><span class="swatch" style="background:${slice.color}"></span><span>${escapeHtml(slice.name)}</span><b>${slice.count}</b></div>`
      )
      .join("");
    const byLocation = summary.locations.map((loc) => ({ label: loc.name, count: loc.containers }));
    const parents = [];
    const parentCounts = new Map();
    for (const rec of list) {
      const name = rec.parent || "(no parent)";
      parentCounts.set(name, (parentCounts.get(name) || 0) + 1);
    }
    for (const [label, count] of parentCounts) parents.push({ label, count });
    const groups = [];
    for (const rec of list) {
      let loc = groups.find((item) => item.name === rec.location);
      if (!loc) {
        loc = { name: rec.location, parents: [] };
        groups.push(loc);
      }
      let parent = loc.parents.find((item) => item.name === rec.parent);
      if (!parent) {
        parent = { name: rec.parent, rows: [] };
        loc.parents.push(parent);
      }
      parent.rows.push(rec);
    }
    const grouped = groups
      .map((loc) => {
        const locCount = loc.parents.reduce((sum, parent) => sum + parent.rows.length, 0);
        const locAlert = loc.parents.reduce(
          (sum, parent) => sum + parent.rows.filter((row) => row.status === "Alert").length,
          0
        );
        const parentsHtml = loc.parents
          .map((parent) => {
            const alert = parent.rows.filter((row) => row.status === "Alert").length;
            const lines = parent.rows
              .map((row) => {
                const units = row.units ? `${escapeHtml(row.units)} units` : "";
                return `<button type="button" class="dash-line" data-container="${escapeHtml(row.container)}"><span>${escapeHtml(row.container)}</span><span>${escapeHtml(row.age || "—")}</span><span>${units}</span><span class="tag ${statusClass(row.status)}">${escapeHtml(row.status)}</span></button>`;
              })
              .join("");
            return `<div class="dash-parent">${escapeHtml(parent.name)} · ${parent.rows.length} containers · ${alert} alert</div>${lines}`;
          })
          .join("");
        return `<div class="dash-loc">${escapeHtml(loc.name)} · ${locCount} containers · ${locAlert} alert</div>${parentsHtml}`;
      })
      .join("");
    return `<div class="dash-who">${escapeHtml(whoLine)}</div>${tallyBoxes(summary)}<div class="dash-charts"><div class="dash-pie">${pieChart(slices)}</div><div class="dash-legend">${legend}</div></div><div class="dash-section">BY LOCATION</div>${barChart(byLocation, "")}<div class="dash-section">BY PARENT</div>${barChart(parents, "parent")}<div class="dash-section">CONTAINERS</div>${grouped}`;
  }

  function splitIdType(raw) {
    const s = cellText(raw);
    const m = s.match(/^(\S+)\s*\[([^\]]+)\]\s*$/);
    if (m) return { id: m[1], type: m[2] };
    if (/^none$/i.test(s)) return { id: "", type: "" };
    return { id: s, type: "" };
  }

  function idsMatch(a, b) {
    return cellText(a).toLowerCase() === cellText(b).toLowerCase();
  }

  function textHasId(text, id) {
    const needle = cellText(id).toLowerCase();
    return Boolean(needle) && String(text || "").toLowerCase().includes(needle);
  }

  function fieldLine(text, label) {
    const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const withColon = new RegExp("^" + escaped + "\\s*:\\s*(.*)$", "i");
    const sameLine = new RegExp("^" + escaped + "\\s+(\\S.*)$", "i");
    const labelOnly = new RegExp("^" + escaped + "\\s*:?\\s*$", "i");
    const lines = String(text || "")
      .split(/\n/)
      .map((l) => l.replace(/\u00a0/g, " ").trim());
    for (let i = 0; i < lines.length; i++) {
      const colon = lines[i].match(withColon);
      if (colon && colon[1]) return colon[1].trim();
      const same = lines[i].match(sameLine);
      if (same && same[1] && !labelOnly.test(lines[i])) return same[1].trim();
      if (labelOnly.test(lines[i])) {
        for (let j = i + 1; j < lines.length; j++) {
          if (lines[j]) return lines[j];
        }
      }
    }
    return "";
  }

  function parseInventoryLine(line) {
    const raw = cellText(line);
    if (raw.includes("\t")) {
      const parts = raw.split("\t").map((p) => p.trim());
      if (parts.length >= 5 && /^\d+$/.test(parts[2])) {
        return {
          age: parts[0],
          fcSku: parts[1],
          quantity: parts[2],
          owner: parts[3],
          disposition: parts[4],
        };
      }
    }
    const parts = raw.split(/\s+/);
    if (parts.length < 5) return null;
    const disposition = parts.pop();
    const owner = parts.pop();
    const quantity = parts.pop();
    const fcSku = parts.pop();
    const age = parts.join(" ");
    if (!/^\d+$/.test(quantity)) return null;
    if (/^showing$/i.test(age) || /^filter:?$/i.test(age) || /^age$/i.test(age)) return null;
    return { age, fcSku, quantity, owner, disposition };
  }

  function parseInventory(text) {
    const lines = String(text || "")
      .split(/\n/)
      .map((l) => l.replace(/\u00a0/g, " ").trim())
      .filter(Boolean);
    const idx = lines.findIndex((l) => /FcSku/i.test(l) && /Disposition/i.test(l));
    if (idx < 0) return [];
    const rows = [];
    for (let i = idx + 1; i < lines.length; i++) {
      const line = lines[i];
      if (/^Showing\b/i.test(line) || /^Previous\b/i.test(line) || /^Next\b/i.test(line) || /^Filter:?$/i.test(line)) break;
      if (/FcSku/i.test(line) && /Disposition/i.test(line)) continue;
      const row = parseInventoryLine(line);
      if (row) rows.push(row);
    }
    return rows;
  }

  function parseDetailsText(text) {
    const t = String(text || "");
    const container = splitIdType(fieldLine(t, "Container ID") || fieldLine(t, "Container details") || "");
    const parent = splitIdType(fieldLine(t, "Parent Container"));
    const grandRaw = fieldLine(t, "Grandparent Container");
    const grand = /^none$/i.test(grandRaw) ? { id: "", type: "" } : splitIdType(grandRaw);
    const inventory = parseInventory(t);
    const headerId = splitIdType(
      (t.split(/\n/).map((l) => l.trim()).find((l) => /\[[A-Z0-9_]+\]/.test(l)) || "")
    );
    const oldestAge = fieldLine(t, "Oldest Units Age") || inventory[0]?.age || "";
    const totalUnits = fieldLine(t, "Total units") || inventory.reduce((n, r) => n + Number(r.quantity || 0), 0) || "";
    const found = Boolean(oldestAge || inventory.length);
    return {
      ok: found,
      containerId: container.id || headerId.id,
      containerType: container.type || headerId.type,
      totalUnits: String(totalUnits),
      oldestAge,
      parentContainer: parent.id,
      parentType: parent.type,
      grandparent: grand.id,
      inventory,
      raw: t,
    };
  }

  function docsToSearch(rootDoc) {
    const docs = [];
    const add = (doc) => {
      if (!doc || docs.includes(doc)) return;
      docs.push(doc);
      let frames = [];
      try {
        frames = doc.querySelectorAll("iframe, frame");
      } catch {
        return;
      }
      frames.forEach((frame) => {
        try {
          add(frame.contentDocument);
        } catch {
          /* A cross-origin frame has its own copy of this script. */
        }
      });
    };
    add(rootDoc || document);
    return docs;
  }

  function findSearchInput() {
    for (const doc of docsToSearch()) {
      const input = doc.querySelector("#searchContainerInput");
      if (input) return input;
    }
    return null;
  }

  function popoverRoots() {
    const seen = new Set();
    const out = [];
    const add = (n) => {
      if (!n || seen.has(n) || !isShown(n)) return;
      seen.add(n);
      out.push(n);
    };
    for (const doc of docsToSearch()) {
      doc.querySelectorAll(".a-popover-inner, .a-popover-content, [id^='a-popover-content']").forEach(add);
      doc.querySelectorAll(".a-popover:not(.a-hidden), [id^='a-popover']:not(.a-hidden)").forEach(add);
    }
    return out;
  }

  function isShown(el) {
    if (!(el instanceof Element)) return false;
    if (el.closest(".a-hidden")) return false;
    const r = el.getBoundingClientRect();
    return r.width > 8 && r.height > 8;
  }

  function isNoRecord(text) {
    return /no record found|no records found|not found|no inventory|does not exist|could not|internal failure/i.test(text);
  }

  function scrapePopoverFor(containerId) {
    const want = cellText(containerId);
    let best = { ok: false, error: "NO_POPOVER" };
    for (const root of popoverRoots()) {
      const t = root.innerText || "";
      if (t.length < 12) continue;
      const hasAge = /Oldest Units Age|Total units/i.test(t);
      const hasDetails = hasAge || /Container details/i.test(t);
      if (isNoRecord(t) && !hasAge) {
        if (best.error !== "NOT_FOUND" || t.length > cellText(best.raw).length) {
          best = { ok: false, error: "NOT_FOUND", raw: t, containerId: want };
        }
        continue;
      }
      if (best.error === "NOT_FOUND") continue;
      if (hasDetails && (!want || textHasId(t, want))) {
        const parsed = parseDetailsText(t);
        if (parsed.oldestAge || (parsed.inventory && parsed.inventory.length)) {
          parsed.ok = true;
          parsed.containerId = want || parsed.containerId;
          parsed.error = "";
          return parsed;
        }
        best = { ok: false, error: "PENDING", raw: t };
        continue;
      }
      if (hasDetails && want && !textHasId(t, want)) {
        if (cellText(t).toLowerCase() === "container details" && best.raw) continue;
        best = { ok: false, error: "STALE", raw: t };
      }
    }
    return best;
  }

  function setInputValue(el, value) {
    const view = el.ownerDocument?.defaultView || pageWindow;
    const desc = Object.getOwnPropertyDescriptor(view.HTMLInputElement.prototype, "value");
    if (desc?.set) desc.set.call(el, value);
    else el.value = value;
    try {
      el.dispatchEvent(new view.InputEvent("input", { bubbles: true, data: value, inputType: "insertText" }));
    } catch {
      el.dispatchEvent(new view.Event("input", { bubbles: true }));
    }
    el.dispatchEvent(new view.Event("change", { bubbles: true }));
  }

  function releasePageButton(doc) {
    const active = doc?.activeElement;
    if (!(active instanceof HTMLElement) || active === document.body) return;
    if (active.id === "searchContainerInput" || active.closest?.("#searchContainerButton")) active.blur();
  }

  function submitPeculiarSearch(containerId) {
    const pecInput = findSearchInput();
    if (!pecInput) throw new Error("Peculiar search box not found. Open MEM3 overview or Reverse Logistics.");
    const doc = pecInput.ownerDocument;
    const view = doc.defaultView || pageWindow;
    setInputValue(pecInput, containerId);
    const submit =
      doc.querySelector("#searchContainerButton input.a-button-input") ||
      doc.querySelector("#searchContainerButton .a-button-input") ||
      doc.querySelector("#searchContainerButton");
    if (cellText(pecInput.value) !== cellText(containerId)) setInputValue(pecInput, containerId);
    if (submit) submit.click();
    else pecInput.dispatchEvent(new view.KeyboardEvent("keydown", { key: "Enter", keyCode: 13, bubbles: true }));
    pecInput.blur();
    if (submit instanceof HTMLElement) submit.blur();
    releasePageButton(doc);
    if (typeof restoreScanFocus === "function") restoreScanFocus();
  }

  function popoverStillShell(result, containerId) {
    const raw = cellText(result?.raw);
    if (!raw) return true;
    if (isNoRecord(raw)) return false;
    if (/Oldest Units Age|Total units/i.test(raw)) return false;
    return !textHasId(raw, containerId) || raw.length < 80;
  }

  function waitForPopover(containerId, ms) {
    return new Promise((resolve) => {
      let finished = false;
      let notFoundSince = 0;
      let retried = false;
      const started = Date.now();
      const finish = (result) => {
        if (finished) return;
        finished = true;
        clearInterval(poll);
        clearTimeout(timer);
        resolve(result);
      };
      const check = () => {
        const result = scrapePopoverFor(containerId);
        if (result.ok) {
          finish(result);
          return;
        }
        if (result.error === "NOT_FOUND") {
          if (!notFoundSince) notFoundSince = Date.now();
          if (Date.now() - notFoundSince >= 600) finish(result);
          return;
        }
        notFoundSince = 0;
        if (!retried && Date.now() - started > 1500 && popoverStillShell(result, containerId)) {
          retried = true;
          try {
            submitPeculiarSearch(containerId);
          } catch {
            /* The first click already ran. Keep waiting for the popover. */
          }
        }
      };
      const poll = setInterval(check, 80);
      const timer = setTimeout(() => {
        const last = scrapePopoverFor(containerId);
        finish(last.ok ? last : { ok: false, error: "TIMEOUT", raw: last.raw || "" });
      }, ms);
      check();
    });
  }

  let pecLock = Promise.resolve();
  async function searchPeculiarNow(containerId) {
    const t0 = Date.now();
    submitPeculiarSearch(containerId);
    const result = await waitForPopover(containerId, 12000);
    result.lookupMs = Date.now() - t0;
    return result;
  }

  function searchPeculiar(containerId) {
    const job = pecLock.then(
      () => searchPeculiarNow(containerId),
      () => searchPeculiarNow(containerId)
    );
    pecLock = job.then(
      () => undefined,
      () => undefined
    );
    return job;
  }

  function peculiarFresh() {
    if (isPeculiar && findSearchInput()) return true;
    const t = Number(GM_getValue(KEYS.heartbeat, 0)) || 0;
    return Date.now() - t < HEARTBEAT_STALE_MS;
  }

  function requestLookup(containerId) {
    return new Promise((resolve, reject) => {
      const id = nonce();
      let done = false;
      const timer = setTimeout(() => {
        if (done) return;
        done = true;
        reject(new Error("Peculiar tab is not listening. Open Peculiar, finish Midway, leave it open."));
      }, LOOKUP_TIMEOUT_MS);
      const listenId = GM_addValueChangeListener(KEYS.response, (_n, _o, res) => {
        if (!res || res.nonce !== id || done) return;
        done = true;
        clearTimeout(timer);
        try {
          GM_removeValueChangeListener(listenId);
        } catch {
          /* ignore */
        }
        resolve(res.result || { ok: false, error: "EMPTY" });
      });
      GM_setValue(KEYS.request, { nonce: id, containerId, t: Date.now() });
    });
  }

  async function lookup(containerId) {
    if (isPeculiar) return searchPeculiar(containerId);
    return requestLookup(containerId);
  }

  function gmReq(url, opts = {}) {
    return new Promise((resolve, reject) => {
      const req = {
        method: opts.method || "GET",
        url,
        headers: opts.headers || {},
        data: opts.body,
        anonymous: opts.anonymous === true,
        onload: (res) => resolve(res),
        onerror: () => reject(new Error("Network error talking to SharePoint.")),
        ontimeout: () => reject(new Error("SharePoint timed out.")),
      };
      if (opts.responseType) req.responseType = opts.responseType;
      if (opts.binary) req.binary = true;
      if (opts.timeout) req.timeout = opts.timeout;
      GM_xmlhttpRequest(req);
    });
  }

  async function spJson(url, opts = {}) {
    const headers = Object.assign(
      { Accept: "application/json;odata=nometadata" },
      opts.headers || {}
    );
    const res = await gmReq(url, { ...opts, headers });
    let json = null;
    try {
      json = res.responseText ? JSON.parse(res.responseText) : null;
    } catch {
      json = { raw: clip(res.responseText, 240) };
    }
    if (res.status < 200 || res.status >= 300) {
      const msg =
        (typeof json?.error?.message === "string" && json.error.message) ||
        json?.error?.message?.value ||
        json?.["odata.error"]?.message?.value ||
        json?.raw ||
        res.statusText ||
        String(res.status);
      lastError = `${res.status} ${typeof msg === "string" ? msg : JSON.stringify(msg)}`;
      const err = new Error(lastError);
      if (isAccessDenied(res.status, lastError)) err.code = "ACCESS_DENIED";
      throw err;
    }
    return json;
  }

  async function digest() {
    if (digestVal && Date.now() < digestExp) return digestVal;
    const json = await spJson(`${SP_SITE}/_api/contextinfo`, { method: "POST" });
    digestVal = json.FormDigestValue || json.GetContextWebInformation?.FormDigestValue;
    digestExp = Date.now() + 18 * 60 * 1000;
    return digestVal;
  }

  function aliasFromLogin(login) {
    const s = String(login || "");
    const mail = s.match(/([a-z0-9._-]+)@/i);
    if (mail) return mail[1].toLowerCase();
    const tail = s.split("|").pop() || s;
    return tail.replace(/[^a-z0-9._-]/gi, "").slice(0, 32) || "operator";
  }

  async function currentAlias() {
    const cached = GM_getValue(KEYS.session, null);
    if (cached?.alias && cached.day === dayStamp()) return cached.alias;
    const user = await spJson(`${SP_SITE}/_api/web/currentuser?$select=Title,LoginName,Email,UserPrincipalName`);
    const alias = aliasFromLogin(user.UserPrincipalName || user.Email || user.LoginName);
    return alias;
  }

  function normalizeLocation(value) {
    return String(value || "")
      .replace(/[\u0000-\u001f]/g, "")
      .replace(/[:\\/?*\[\]]/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 31);
  }

  function loadCustomLocations() {
    try {
      const raw = GM_getValue(KEYS.locations, []);
      if (!Array.isArray(raw)) return [];
      const seen = new Set(PRESET_LOCATIONS.map((name) => name.toLowerCase()));
      const out = [];
      for (const item of raw) {
        const name = normalizeLocation(item);
        const key = name.toLowerCase();
        if (!name || seen.has(key)) continue;
        seen.add(key);
        out.push(name);
      }
      return out;
    } catch {
      return [];
    }
  }

  function saveCustomLocations() {
    GM_setValue(KEYS.locations, customLocations);
  }

  function allLocations() {
    return [...PRESET_LOCATIONS, ...customLocations];
  }

  function loadSelectedLocation() {
    const saved = normalizeLocation(GM_getValue(KEYS.location, ""));
    return allLocations().find((name) => name.toLowerCase() === saved.toLowerCase()) || "";
  }

  function saveSelectedLocation() {
    GM_setValue(KEYS.location, selectedLocation);
  }

  function fileNameFor(alias) {
    return `LineFlush_${dayStamp()}_${alias}.xlsx`;
  }

  function csvNameFor(alias) {
    return `LineFlush_${dayStamp()}_${alias}.csv`;
  }

  function filePathFor(name) {
    return `${SOS_FOLDER}/${name}`;
  }

  function quotePath(path) {
    return String(path || "").replace(/'/g, "''");
  }

  function spFolderUrl(query) {
    const q = query ? `${query}&` : "";
    return `${SP_SITE}/_api/web/GetFolderByServerRelativeUrl(@p)?${q}@p='${quotePath(SOS_FOLDER)}'`;
  }

  function spFileUrl(path, extra) {
    return `${SP_SITE}/_api/web/GetFileByServerRelativeUrl(@p)${extra || ""}?@p='${quotePath(path)}'`;
  }

  async function folderExists(path) {
    const res = await gmReq(`${SP_SITE}/_api/web/GetFolderByServerRelativeUrl(@p)?@p='${quotePath(path)}'`, {
      headers: { Accept: "application/json;odata=nometadata" },
    });
    return res.status >= 200 && res.status < 300;
  }

  async function createFolder(path) {
    const d = await digest();
    const res = await gmReq(`${SP_SITE}/_api/web/folders`, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": d,
      },
      body: JSON.stringify({ ServerRelativeUrl: path }),
    });
    if (res.status === 409) return;
    if (res.status >= 200 && res.status < 300) return;
    lastError = `${res.status} ${clip(res.responseText, 180)}`;
    throw new Error(lastError);
  }

  function childFoldersOfDocuments(fullPath) {
    const lower = fullPath.toLowerCase();
    const needle = "/documents/";
    const idx = lower.indexOf(needle);
    if (idx < 0) return [fullPath];
    const root = fullPath.slice(0, idx + "/Documents".length);
    const rest = fullPath.slice(idx + needle.length).split("/").filter(Boolean);
    const out = [];
    let acc = root;
    for (const seg of rest) {
      acc += "/" + seg;
      out.push(acc);
    }
    return out;
  }

  async function ensureFolderPath(path) {
    if (await folderExists(path)) return;
    for (const p of childFoldersOfDocuments(path)) {
      if (await folderExists(p)) continue;
      await createFolder(p);
    }
    if (!(await folderExists(path))) {
      throw new Error("Could not create the Line Flush Records folder. Create it in SharePoint, then refresh.");
    }
  }

  async function assertSosFolder() {
    try {
      await spJson(spFolderUrl());
      return;
    } catch (err) {
      if (!/404|File Not Found|DirectoryNotFound/i.test(String(err.message || err))) throw err;
    }
    await ensureFolderPath(SOS_FOLDER);
    lastError = "";
  }

  const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }

  function u16(view, offset, value) {
    view.setUint16(offset, value, true);
  }

  function u32(view, offset, value) {
    view.setUint32(offset, value, true);
  }

  function bytesToBinaryString(bytes) {
    let out = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      out += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return out;
  }

  function concatBytes(parts) {
    const total = parts.reduce((sum, part) => sum + part.length, 0);
    const out = new Uint8Array(total);
    let offset = 0;
    for (const part of parts) {
      out.set(part, offset);
      offset += part.length;
    }
    return out;
  }

  function zipStore(files) {
    const encoder = new TextEncoder();
    const locals = [];
    const centrals = [];
    let offset = 0;
    for (const file of files) {
      const name = encoder.encode(file.name);
      const data = file.data instanceof Uint8Array ? file.data : encoder.encode(String(file.data));
      const crc = crc32(data);
      const local = new Uint8Array(30 + name.length);
      const localView = new DataView(local.buffer);
      u32(localView, 0, 0x04034b50);
      u16(localView, 4, 20);
      u16(localView, 8, 0);
      u32(localView, 14, crc);
      u32(localView, 18, data.length);
      u32(localView, 22, data.length);
      u16(localView, 26, name.length);
      local.set(name, 30);
      const central = new Uint8Array(46 + name.length);
      const centralView = new DataView(central.buffer);
      u32(centralView, 0, 0x02014b50);
      u16(centralView, 4, 20);
      u16(centralView, 6, 20);
      u32(centralView, 16, crc);
      u32(centralView, 20, data.length);
      u32(centralView, 24, data.length);
      u16(centralView, 28, name.length);
      u32(centralView, 42, offset);
      central.set(name, 46);
      locals.push(local, data);
      centrals.push(central);
      offset += local.length + data.length;
    }
    const centralBytes = concatBytes(centrals);
    const end = new Uint8Array(22);
    const endView = new DataView(end.buffer);
    u32(endView, 0, 0x06054b50);
    u16(endView, 8, files.length);
    u16(endView, 10, files.length);
    u32(endView, 12, centralBytes.length);
    u32(endView, 16, offset);
    return concatBytes([...locals, centralBytes, end]);
  }

  function xmlEscape(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function xmlUnescape(value) {
    return String(value || "")
      .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
      .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&");
  }

  function xmlAttr(chunk, name) {
    const match = chunk.match(new RegExp("(?:^|\\s)" + name.replace(/[.*+?^${}()|[\]\\:]/g, "\\$&") + '="([^"]*)"'));
    return match ? match[1] : "";
  }

  function colLetters(index) {
    let n = index + 1;
    let out = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  }

  function colIndex(ref) {
    const letters = String(ref || "").match(/[A-Z]+/i);
    if (!letters) return 0;
    let n = 0;
    for (const ch of letters[0].toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  function fillStyleIndex(value) {
    if (value === "Alert") return 1;
    if (value === "Caution") return 2;
    if (value === "Good") return 3;
    return 0;
  }

  function sheetXml(rows) {
    const body = rows
      .map((row, rowIndex) => {
        const cells = row
          .map((value, col) => {
            if (value == null || value === "") return "";
            const ref = `${colLetters(col)}${rowIndex + 1}`;
            const style = fillStyleIndex(value);
            const painted = style ? ` s="${style}"` : "";
            return `<c r="${ref}"${painted} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
          })
          .join("");
        return `<row r="${rowIndex + 1}">${cells}</row>`;
      })
      .join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;
  }

  function workbookBytes(book) {
    const sheets = book.sheets.length ? book.sheets : [{ name: "Scans", rows: [HEADERS.slice()] }];
    const sheetFiles = sheets.map((sheet, index) => ({
      name: `xl/worksheets/sheet${index + 1}.xml`,
      data: sheetXml(sheet.rows),
    }));
    const sheetTags = sheets
      .map(
        (sheet, index) =>
          `<sheet name="${xmlEscape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`
      )
      .join("");
    const sheetRels = sheets
      .map(
        (_, index) =>
          `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
      )
      .join("");
    const overrides = sheets
      .map(
        (_, index) =>
          `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
      )
      .join("");
    const files = [
      {
        name: "[Content_Types].xml",
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${overrides}</Types>`,
      },
      {
        name: "_rels/.rels",
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      },
      {
        name: "xl/workbook.xml",
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheetTags}</sheets></workbook>`,
      },
      {
        name: "xl/_rels/workbook.xml.rels",
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheetRels}<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
      },
      {
        name: "xl/styles.xml",
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FF111111"/><name val="Calibri"/></font></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFB42318"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF5C518"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0D7A3F"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>`,
      },
      ...sheetFiles,
    ];
    return zipStore(files);
  }

  async function inflateRaw(bytes) {
    if (typeof DecompressionStream !== "function") {
      throw new Error("Today's workbook was resaved by Excel. Close it, then scan again.");
    }
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function unzip(bytes) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let eocd = -1;
    const min = Math.max(0, bytes.length - 22 - 65535);
    for (let i = bytes.length - 22; i >= min; i--) {
      if (view.getUint32(i, true) === 0x06054b50) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new Error("Today's workbook could not be read.");
    const count = view.getUint16(eocd + 10, true);
    let cursor = view.getUint32(eocd + 16, true);
    const decoder = new TextDecoder();
    const files = {};
    for (let i = 0; i < count; i++) {
      if (view.getUint32(cursor, true) !== 0x02014b50) throw new Error("Today's workbook could not be read.");
      const method = view.getUint16(cursor + 10, true);
      const compSize = view.getUint32(cursor + 20, true);
      const nameLen = view.getUint16(cursor + 28, true);
      const extraLen = view.getUint16(cursor + 30, true);
      const commentLen = view.getUint16(cursor + 32, true);
      const localOffset = view.getUint32(cursor + 42, true);
      const name = decoder.decode(bytes.subarray(cursor + 46, cursor + 46 + nameLen));
      const localNameLen = view.getUint16(localOffset + 26, true);
      const localExtraLen = view.getUint16(localOffset + 28, true);
      const dataStart = localOffset + 30 + localNameLen + localExtraLen;
      let data = bytes.subarray(dataStart, dataStart + compSize);
      if (method === 8) data = await inflateRaw(data);
      else if (method !== 0) throw new Error("Today's workbook uses a compression Line Flush cannot read.");
      files[name] = data;
      cursor += 46 + nameLen + extraLen + commentLen;
    }
    return files;
  }

  function parseSharedStrings(xml) {
    const out = [];
    const items = xml.match(/<si\b[^>]*>[\s\S]*?<\/si>/g) || [];
    for (const item of items) {
      const texts = [...item.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((match) => xmlUnescape(match[1]));
      out.push(texts.join(""));
    }
    return out;
  }

  function parseSheetRows(xml, shared) {
    const rows = [];
    const chunks = xml.split(/<row\b/);
    for (let i = 1; i < chunks.length; i++) {
      const end = chunks[i].search(/<\/row>/);
      const body = end >= 0 ? chunks[i].slice(0, end) : chunks[i];
      const cells = [];
      const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let match;
      while ((match = cellRe.exec(body))) {
        const attrs = match[1];
        const inner = match[2] || "";
        const col = colIndex(xmlAttr(attrs, "r"));
        const type = xmlAttr(attrs, "t");
        let value = "";
        if (type === "inlineStr") {
          value = [...inner.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((item) => xmlUnescape(item[1])).join("");
        } else if (type === "s") {
          const index = Number((inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1] || 0);
          value = shared[index] || "";
        } else {
          const raw = inner.match(/<v>([\s\S]*?)<\/v>/);
          value = raw ? xmlUnescape(raw[1]) : "";
        }
        while (cells.length < col) cells.push("");
        cells[col] = value;
      }
      if (cells.some((cell) => String(cell).trim())) rows.push(cells);
    }
    return rows;
  }

  async function workbookFromXlsx(bytes) {
    const files = await unzip(bytes);
    const decoder = new TextDecoder();
    const text = (name) => (files[name] ? decoder.decode(files[name]) : "");
    const workbook = text("xl/workbook.xml");
    const rels = text("xl/_rels/workbook.xml.rels");
    if (!workbook) throw new Error("Today's workbook could not be read.");
    const shared = files["xl/sharedStrings.xml"] ? parseSharedStrings(text("xl/sharedStrings.xml")) : [];
    const targets = {};
    for (const rel of rels.match(/<Relationship\b[^>]*>/g) || []) {
      const id = xmlAttr(rel, "Id");
      let target = xmlAttr(rel, "Target");
      if (!id || !target) continue;
      if (target.startsWith("/")) target = target.slice(1);
      if (!target.startsWith("xl/")) target = `xl/${target.replace(/^\.?\//, "")}`;
      targets[id] = target;
    }
    const sheets = [];
    for (const tag of workbook.match(/<sheet\b[^>]*>/g) || []) {
      const name = normalizeLocation(xmlUnescape(xmlAttr(tag, "name"))) || "Sheet";
      const id = xmlAttr(tag, "r:id");
      const xml = text(targets[id] || "");
      sheets.push({ name, rows: xml ? parseSheetRows(xml, shared) : [] });
    }
    return { sheets };
  }

  function sheetRowCount(book) {
    return (book?.sheets || []).reduce((sum, sheet) => {
      if (sheet.name.toLowerCase() === TOTAL_SHEET.toLowerCase()) return sum;
      const header = sheet.rows[0]?.[0] === "Container ID" ? 1 : 0;
      return sum + Math.max(0, sheet.rows.length - header);
    }, 0);
  }

  function rebuildTotal(book) {
    const rows = [["Location", ...HEADERS]];
    for (const sheet of book.sheets) {
      if (sheet.name.toLowerCase() === TOTAL_SHEET.toLowerCase()) continue;
      const start = sheet.rows[0]?.[0] === "Container ID" ? 1 : 0;
      for (const row of sheet.rows.slice(start)) {
        if (!row.some((cell) => String(cell || "").trim())) continue;
        rows.push([sheet.name, ...row]);
      }
    }
    const rest = book.sheets.filter((sheet) => sheet.name.toLowerCase() !== TOTAL_SHEET.toLowerCase());
    book.sheets = [{ name: TOTAL_SHEET, rows }, ...rest];
  }

  function touchLocationClock() {
    lastScanAt = Date.now();
    try {
      GM_setValue(KEYS.locationIdle, lastScanAt);
    } catch {
      /* ignore */
    }
  }

  function ensureSheet(book, location) {
    const name = normalizeLocation(location) || "Location";
    let sheet = book.sheets.find((item) => item.name.toLowerCase() === name.toLowerCase());
    if (!sheet) {
      sheet = { name, rows: [HEADERS.slice()] };
      book.sheets.push(sheet);
    }
    if (!sheet.rows.length || sheet.rows[0][0] !== "Container ID") sheet.rows.unshift(HEADERS.slice());
    return sheet;
  }

  async function readSosFile(path) {
    const res = await gmReq(spFileUrl(path, "/$value"), {
      method: "GET",
      headers: { Accept: "text/plain" },
    });
    if (res.status === 404) return null;
    if (res.status < 200 || res.status >= 300) {
      lastError = `${res.status} ${clip(res.responseText, 180)}`;
      throw new Error(lastError);
    }
    return res.responseText || "";
  }

  async function readSosBytes(path) {
    const res = await gmReq(spFileUrl(path, "/$value"), {
      method: "GET",
      responseType: "arraybuffer",
      headers: { Accept: "application/octet-stream" },
    });
    if (res.status === 404) return null;
    if (res.status < 200 || res.status >= 300) {
      lastError = `${res.status} ${clip(res.responseText, 180)}`;
      throw new Error(lastError);
    }
    if (res.response instanceof ArrayBuffer) return new Uint8Array(res.response);
    const text = res.responseText || "";
    const bytes = new Uint8Array(text.length);
    for (let i = 0; i < text.length; i++) bytes[i] = text.charCodeAt(i) & 0xff;
    return bytes.length ? bytes : null;
  }

  function excelFileIsOpen(status, text) {
    if (status === 423) return true;
    return status >= 400 && /locked|exclusive use|being edited|checked out|SPFileLock/i.test(String(text || ""));
  }

  function isAccessDenied(status, text) {
    const blob = `${status || ""} ${text || ""}`;
    return (
      Number(status) === 403 ||
      Number(status) === 401 ||
      /access denied|unauthorizedaccessexception|-2147024891|\b401\b/i.test(blob)
    );
  }

  function accessDeniedError(name) {
    const file = name || sosCache.name || "today's workbook";
    lastError = `Access denied. ${file} was not created.`;
    const err = new Error(lastError);
    err.code = "ACCESS_DENIED";
    err.fileName = file;
    return err;
  }

  function fileLockError(name) {
    const file = name || sosCache.name || "today's workbook";
    lastError = `Excel has ${file} open.`;
    const err = new Error(lastError);
    err.code = "FILE_LOCK";
    err.fileName = file;
    return err;
  }

  async function writeSosFile(name, bytes) {
    const d = await digest();
    const url = `${SP_SITE}/_api/web/GetFolderByServerRelativeUrl(@p)/Files/add(url='${name.replace(
      /'/g,
      ""
    )}',overwrite=true)?@p='${quotePath(SOS_FOLDER)}'`;
    const headers = {
        Accept: "application/json;odata=nometadata",
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "X-RequestDigest": d,
    };
    const body = bytesToBinaryString(bytes);
    const send = () => gmReq(url, { method: "POST", headers, body, binary: true });
    const res = await send();
    if (excelFileIsOpen(res.status, res.responseText)) throw fileLockError(name);
    if (isAccessDenied(res.status, res.responseText)) throw accessDeniedError(name);
    if (res.status < 200 || res.status >= 300) {
      if (/404|DirectoryNotFound|File Not Found/i.test(`${res.status} ${res.responseText || ""}`)) {
        await assertSosFolder();
        const retry = await send();
        if (excelFileIsOpen(retry.status, retry.responseText)) throw fileLockError(name);
        if (isAccessDenied(retry.status, retry.responseText)) throw accessDeniedError(name);
        if (retry.status >= 200 && retry.status < 300) {
          lastError = "";
          fileOk = true;
          return;
        }
        lastError = `${retry.status} ${clip(retry.responseText, 180)}`;
        throw new Error(lastError);
      }
      lastError = `${res.status} ${clip(res.responseText, 180)}`;
      throw new Error(lastError);
    }
    lastError = "";
    fileOk = true;
  }

  function rowsFromResult(containerId, result, alias) {
    const inventory = result.inventory && result.inventory.length ? result.inventory : [{}];
    const extra = Math.max(0, inventory.length - 1);
    return inventory.map((inv, i) => {
      const status = result.ok ? "OK" : result.error === "NOT_FOUND" ? "NOT FOUND" : "ERROR";
      const notes = result.ok
        ? i === 0 && extra
          ? `${extra} more SKU(s) in extra rows`
          : ""
        : result.notes || result.error || "Lookup failed";
      return [
        result.containerId || containerId,
        nowIso(),
        alias,
        i === 0 ? result.containerType || "" : "",
        i === 0 ? result.totalUnits || "" : "",
        i === 0 ? result.oldestAge || "" : "",
        i === 0 ? result.parentContainer || "" : "",
        i === 0 ? result.parentType || "" : "",
        i === 0 ? result.grandparent || "" : "",
        inv.age || (i === 0 ? result.oldestAge || "" : ""),
        inv.fcSku || "",
        inv.quantity || "",
        inv.owner || "",
        inv.disposition || "",
        i === 0 ? status : "",
        i === 0 ? notes : "",
      ];
    });
  }

  let saveChain = Promise.resolve();
  function enqueueSave(fn) {
    const run = saveChain.then(fn, fn);
    saveChain = run.then(
      () => undefined,
      () => undefined
    );
    return run;
  }

  async function ensureSosCache(alias) {
    if (sosCache.book && sosCache.alias === alias && sosCache.day === dayStamp()) return;
    await assertSosFolder();
    const name = fileNameFor(alias);
    let bytes = await readSosBytes(filePathFor(name));
    if (!bytes || !bytes.length) {
      try {
        bytes = await readSosBytes(`${OLD_SOS_FOLDER}/${name}`);
      } catch {
        bytes = null;
      }
    }
    let book = null;
    if (bytes && bytes.length) book = await workbookFromXlsx(bytes);
    if (!book) {
      book = { sheets: [] };
      const csv = await readSosFile(filePathFor(csvNameFor(alias)));
      if (csv) {
        const rows = parseCsv(csv);
        if (rows.length) {
          if (rows[0][0] !== "Container ID") rows.unshift(HEADERS.slice());
          book.sheets.push({ name: "Unassigned", rows });
        }
      }
    }
    sosCache.alias = alias;
    sosCache.day = dayStamp();
    sosCache.name = name;
    sosCache.book = book;
    sosCache.records = recordsFromBook(book);
  }

  function dashboardBook(records) {
    return { sheets: [{ name: "Dashboard", rows: dashboardSheetRows(records) }] };
  }

  async function appendSosRows(record, alias, location) {
    await ensureSosCache(alias);
    const before = sosCache.records.length;
    const previous = sosCache.book;
    sosCache.records.push(record);
    sosCache.book = dashboardBook(sosCache.records);
    try {
      await writeSosFile(sosCache.name, workbookBytes(sosCache.book));
    } catch (err) {
      sosCache.records.splice(before);
      sosCache.book = previous;
      fileOk = false;
      err.record = record;
      err.rows = [[record.container]];
      err.alias = alias;
      err.location = location || record.location;
      if (!err.fileName) err.fileName = sosCache.name;
      throw err;
    }
    const session = {
      alias,
      day: dayStamp(),
      name: sosCache.name,
      count: sosCache.records.length,
      location: record.location,
      t: Date.now(),
    };
    GM_setValue(KEYS.session, session);
    return session;
  }

  function startPeculiarWorker() {
    GM_setValue(KEYS.heartbeat, Date.now());
    setInterval(() => GM_setValue(KEYS.heartbeat, Date.now()), HEARTBEAT_MS);
    GM_addValueChangeListener(KEYS.request, async (_n, _o, req) => {
      if (!req?.nonce || !req.containerId) return;
      try {
        const result = await searchPeculiar(req.containerId);
        GM_setValue(KEYS.response, { nonce: req.nonce, t: Date.now(), result });
      } catch (err) {
        GM_setValue(KEYS.response, {
          nonce: req.nonce,
          t: Date.now(),
          result: { ok: false, error: "EXCEPTION", notes: String(err.message || err) },
        });
      }
    });
  }

  function dash(v) {
    const s = cellText(v);
    return s ? escapeHtml(s) : "<span class='empty'>—</span>";
  }

  function fcResearchUrl(sku) {
    return `https://qifcr.na.aftx.amazonoperations.app/MEM3/results?s=${encodeURIComponent(sku)}`;
  }

  function skuLink(sku) {
    const text = cellText(sku);
    if (!text) return "<span class='empty'>—</span>";
    const href = fcResearchUrl(text);
    return `<a class="sku" href="${escapeHtml(href)}" target="_blank" rel="noopener" title="Search ${escapeHtml(text)} in FcResearch">${escapeHtml(text)}</a>`;
  }

  function loadHistory(alias) {
    const pack = GM_getValue(KEYS.history, null);
    if (!pack || pack.day !== dayStamp() || pack.alias !== alias || !Array.isArray(pack.rows)) return [];
    return pack.rows.map((row) => {
      if (!row?.result?.raw) return row;
      const result = { ...row.result };
      delete result.raw;
      return { ...row, result };
    });
  }

  function saveHistory(alias, rows) {
    GM_setValue(KEYS.history, { day: dayStamp(), alias, rows: rows.slice(0, 800), t: Date.now() });
  }

  function uiPixelRatio() {
    const ratio = Number(pageWindow.devicePixelRatio || window.devicePixelRatio || 1);
    return ratio > 0 ? ratio : 1;
  }

  function panelBoost() {
    const saved = Number(GM_getValue(KEYS.panelSize, 1));
    if (saved === 1.25 || saved === 1.5) return saved;
    return 1;
  }

  function panelStartsOpen() {
    return GM_getValue(KEYS.panelOpen, true) !== false;
  }

  // Lay the panel out in device pixels, then scale it back onto the window.
  // Boost 1 matches a 100% screen. 1.25 and 1.5 make that same panel bigger.
  function paintPanelScale(el) {
    if (!el) return;
    const ratio = uiPixelRatio();
    const boost = panelBoost();
    const fit = boost / ratio;
    el.style.setProperty("--plf-dpr", String(ratio));
    el.style.setProperty("--plf-boost", String(boost));
    el.style.setProperty("--plf-vw", `${pageWindow.innerWidth * ratio}px`);
    el.style.setProperty("--plf-vh", `${pageWindow.innerHeight * ratio}px`);
    if (Math.abs(ratio - 1) < 0.02 && Math.abs(boost - 1) < 0.02) {
      el.style.setProperty("--plf-place", "0px");
      el.style.setProperty("--plf-stage-transform", "none");
      el.style.setProperty("--plf-scale", "none");
      return;
    }
    el.style.setProperty("--plf-place", "50%");
    el.style.setProperty("--plf-stage-transform", `translate(-50%, -50%) scale(${fit})`);
    el.style.setProperty("--plf-scale", `scale(${fit})`);
  }

  function syncPanelScale() {
    ["plf-scan-host", "plf-old-script", "plf-update", "plf-boot-error"].forEach((id) => {
      paintPanelScale(document.getElementById(id));
    });
  }

  function watchPanelScale() {
    if (watchPanelScale.armed) return;
    watchPanelScale.armed = true;
    const refresh = () => syncPanelScale();
    pageWindow.addEventListener("resize", refresh);
    if (pageWindow.visualViewport) pageWindow.visualViewport.addEventListener("resize", refresh);
    let media = null;
    const onMedia = () => {
      syncPanelScale();
      armMedia();
    };
    const armMedia = () => {
      if (media) media.removeEventListener("change", onMedia);
      media = pageWindow.matchMedia(`(resolution: ${uiPixelRatio()}dppx)`);
      media.addEventListener("change", onMedia);
    };
    armMedia();
  }

  async function copyDebugText() {
    const payload = {
      version: VERSION,
      at: nowIso(),
      href: location.href,
      userAgent: navigator.userAgent,
      devicePixelRatio: uiPixelRatio(),
      panelSize: panelBoost(),
      panelOpen: panelStartsOpen(),
      innerWidth: pageWindow.innerWidth,
      innerHeight: pageWindow.innerHeight,
      host,
      lastError,
      lastDebug,
      peculiarConnected: peculiarFresh(),
      fileOk,
      session: GM_getValue(KEYS.session, null),
      location: selectedLocation,
      customLocations,
      focus: focusNote,
      sosCache: {
        alias: sosCache.alias,
        day: sosCache.day,
        name: sosCache.name,
        records: (sosCache.records || []).length,
        sheets: (sosCache.book?.sheets || []).map((sheet) => ({ name: sheet.name, rows: sheet.rows.length })),
      },
      queued: scanQueueLen,
      recent: recent.slice(0, 8).map((r) => ({ t: r.t, id: r.id, age: r.age, sku: r.sku, status: r.status })),
    };
    const text = JSON.stringify(payload, null, 2);
    let copied = false;
    try {
      GM_setClipboard(text, "text");
      copied = true;
    } catch {
      try {
        await navigator.clipboard.writeText(text);
        copied = true;
      } catch {
        copied = false;
      }
    }
    return { text, copied };
  }

  function startScanner() {
    if (document.getElementById("plf-scan-host")) return;
    const hostEl = document.createElement("div");
    hostEl.id = "plf-scan-host";
    hostEl.style.setProperty("position", "fixed", "important");
    hostEl.style.setProperty("inset", "0", "important");
    hostEl.style.setProperty("z-index", "2147483646", "important");
    hostEl.style.setProperty("pointer-events", "none", "important");
    hostEl.style.setProperty("overflow", "hidden", "important");
    hostEl.style.setProperty("font-family", "Segoe UI, system-ui, sans-serif", "important");
    paintPanelScale(hostEl);
    if (!panelStartsOpen()) hostEl.classList.add("closed");
    const shadow = hostEl.attachShadow({ mode: "open" });
    shadow.innerHTML = `
      <style>
        :host {
          all: initial;
          position: fixed !important;
          inset: 0 !important;
          z-index: 2147483646 !important;
          display: block !important;
          pointer-events: none !important;
          overflow: hidden !important;
          --plf-dpr: 1;
          --plf-vw: 100vw;
          --plf-vh: 100vh;
        }
        .plf-stage {
          position: absolute; left: var(--plf-place, 0px); top: var(--plf-place, 0px);
          width: var(--plf-vw); height: var(--plf-vh);
          transform: var(--plf-stage-transform, none);
          transform-origin: center center;
          pointer-events: none;
          container: plf / inline-size;
        }
        .plf-fab-dock {
          position: absolute; right: 0; bottom: 0; width: 0; height: 0; z-index: 2;
          transform: var(--plf-scale, none);
          transform-origin: bottom right;
          pointer-events: none;
        }
        * { box-sizing: border-box; font-family: "Segoe UI", system-ui, sans-serif; }
        .wrap {
          position: absolute; inset: 0;
          display: flex; align-items: center; justify-content: center;
          pointer-events: none;
        }
        .backdrop { position: absolute; inset: 0; background: rgba(4, 8, 12, .45); pointer-events: none; }
        .modal {
          position: relative; z-index: 1;
          pointer-events: auto;
          width: min(860px, calc((var(--plf-vw) - 16px) / var(--plf-boost, 1)));
          max-height: calc((var(--plf-vh) - 16px) / var(--plf-boost, 1));
          display: flex; flex-direction: column;
          background: #ffffff; color: #111111; color-scheme: light;
          border: 1px solid #d8dee8; border-radius: 12px;
          box-shadow: 0 18px 50px rgba(0,0,0,.28);
          overflow: hidden;
          --plf-glow: transparent;
        }
        .modal::after {
          content: "";
          position: absolute; inset: 0; z-index: 0;
          border-radius: inherit; pointer-events: none;
          background: var(--plf-glow); opacity: 0;
        }
        .modal.glow::after { animation: plf-glow 1s linear 1; }
        @keyframes plf-glow {
          0% { opacity: 1; }
          100% { opacity: 0; }
        }
        .bar {
          display: flex; align-items: center; gap: 8px;
          min-height: 40px; padding: 6px 8px 6px 12px;
          background: #ffffff; border-bottom: 1px solid #e4e7ec;
        }
        .bar button, .clock, .lamps { flex: 0 0 auto; }
        .who {
          font-size: 12px; font-weight: 700; white-space: nowrap;
          overflow: hidden; text-overflow: ellipsis; max-width: 110px;
        }
        .brand {
          flex: 0 0 auto; white-space: nowrap;
          font-size: 12px; font-weight: 800; color: #111111; line-height: 1.15;
        }
        .brand span { font-weight: 700; }
        .lamps { display: flex; gap: 4px; flex: 0 0 auto; }
        .lamp {
          display: flex; align-items: center; gap: 4px;
          font-size: 11px; color: #111111; background: #ffffff;
          border: 1px solid #d8dee8; border-radius: 999px; padding: 2px 7px;
        }
        .dot { width: 7px; height: 7px; border-radius: 50%; background: #6b7785; }
        .ok .dot { background: #3dd68c; }
        .bad .dot { background: #ff5d5d; }
        .counts {
          flex: 1 1 auto; min-width: 0;
          font-size: 12px; color: #111111; white-space: nowrap;
          overflow: hidden; text-overflow: ellipsis;
        }
        .counts b { color: #111111; font-variant-numeric: tabular-nums; }
        .clock {
          flex: 0 0 auto; font-variant-numeric: tabular-nums;
          font-size: 13px; font-weight: 700; color: #111111;
        }
        button {
          border: 1px solid #d8dee8; border-radius: 7px; padding: 5px 8px; cursor: pointer;
          font-size: 12px; font-weight: 700; color: #111111; background: #ffffff;
        }
        button:hover { background: #f3f5f8; }
        #panelSize {
          flex: 0 0 auto; height: 28px; max-width: 118px;
          border: 1px solid #111111; border-radius: 7px;
          background: #ffffff; color: #111111;
          font-size: 12px; font-weight: 800; padding: 0 4px;
        }
          width: 28px; height: 28px; padding: 0; border-radius: 7px;
          background: #ffffff; color: #111111; font-size: 18px; line-height: 1;
        }
        .x:hover { background: #f3f5f8; }
        .scanrow {
          display: flex; align-items: center; gap: 8px;
          padding: 8px 10px 0; background: #ffffff;
        }
        #location {
          flex: 0 0 2.6in; width: 2.6in; max-width: calc(var(--plf-vw) * 0.42 / var(--plf-boost, 1)); height: 38px;
          border: 1px solid #111111; border-radius: 8px;
          background: #ffffff; color: #111111; font-size: 14px; font-weight: 700; padding: 0 8px;
        }
        #delLoc { flex: 0 0 auto; height: 38px; }
        #delLoc:disabled { opacity: .4; cursor: default; }
        #q {
          flex: 1 1 auto; min-width: 0; width: auto;
          border: 2px solid #111111; background: #ffffff; color: #111111;
          border-radius: 8px; padding: 7px 10px; font-size: 22px; font-weight: 700;
          letter-spacing: .04em; outline: none;
        }
        #q::placeholder, #filter::placeholder { color: #667085; }
        #q:focus { border-color: #111111; box-shadow: 0 0 0 3px rgba(17,17,17,.12); }
        .ageboard {
          display: flex; align-items: baseline; gap: 10px;
          margin: 8px 10px 0; padding: 6px 12px;
          border: 1px solid #d8dee8; border-radius: 8px; background: #ffffff; color: #111111;
          transition: background .2s linear, border-color .2s linear;
        }
        .age-kicker { font-size: 11px; letter-spacing: .12em; font-weight: 700; opacity: .8; }
        .age-val {
          flex: 1 1 auto; min-width: 0;
          font-size: 26px; font-weight: 800; line-height: 1.1;
          font-variant-numeric: tabular-nums; letter-spacing: .01em;
          overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        }
        .age-tag { font-size: 12px; font-weight: 800; letter-spacing: .12em; }
        .ready, .info { background: #ffffff; border: 1px solid #d8dee8; color: #111111; }
        .found { background: #e7f7ec; border: 1px solid #8dcea0; color: #111111; }
        .missing { background: #fdeceb; border: 1px solid #e7b0aa; color: #111111; }
        .error { background: #fff4e5; border: 1px solid #e6c48a; color: #111111; }
        .split {
          position: relative; z-index: 1;
          flex: 1 1 auto;
          min-height: 0;
          display: grid; grid-template-columns: minmax(0, 1.25fr) minmax(280px, .9fr);
        }
        .ticket, .history {
          min-height: 0; display: flex; flex-direction: column;
          padding: 8px 10px 0;
        }
        .history { border-left: 1px solid #e4e7ec; padding-right: 8px; background: #ffffff; }
        .history table { table-layout: fixed; }
        .history th:nth-child(1), .history td:nth-child(1) { width: 52px; }
        .history th:nth-child(3), .history td:nth-child(3) { width: 72px; }
        .history th:nth-child(4), .history td:nth-child(4) { width: 44px; }
        .bar, .scanrow, .status, .ageboard, .identity, .foot { position: relative; z-index: 1; flex: 0 0 auto; }
        .identity {
          margin: 8px 10px 0; padding: 8px 10px;
          font-size: 16px; font-weight: 800; line-height: 1.35;
          overflow-wrap: anywhere;
          background: #ffffff; color: #111111;
          border: 2px solid #111111; border-radius: 8px;
        }
        .lineswrap, .histwrap { min-height: 0; overflow: auto; max-height: min(240px, calc((var(--plf-vh) - 230px) / var(--plf-boost, 1))); }
        .histhead { display: flex; align-items: center; gap: 8px; margin-bottom: 6px; }
        .histhead span { font-size: 11px; letter-spacing: .12em; font-weight: 700; color: #111111; }
        #filter {
          flex: 1 1 auto; min-width: 0;
          border: 1px solid #d8dee8; background: #ffffff; color: #111111;
          border-radius: 6px; padding: 4px 8px; font-size: 12px;
        }
        .empty { color: #667085; font-weight: 500; }
        table { width: 100%; border-collapse: collapse; font-size: 12px; color: #111111; }
        th { text-align: left; color: #111111; font-weight: 650; padding: 2px 6px; border-bottom: 1px solid #e4e7ec; position: sticky; top: 0; background: #ffffff; z-index: 1; }
        td { padding: 3px 6px; border-bottom: 1px solid #eef1f5; font-variant-numeric: tabular-nums; }
        #lines tr.aged td { background: var(--row); color: var(--ink); }
        #lines a.sku { color: inherit; font-weight: 800; text-decoration: underline; }
        .histwrap tr { cursor: pointer; }
        .histwrap tr:hover { background: #f3f5f8; }
        .histwrap tr.active { background: #e7edf5; }
        .st-ok { color: #0d7a3f; font-weight: 700; }
        .st-miss { color: #b42318; font-weight: 700; }
        .st-err { color: #b54708; font-weight: 700; }
        .foot {
          display: flex; align-items: center; gap: 8px;
          min-height: 34px; padding: 6px 10px;
          border-top: 1px solid #e4e7ec; background: #ffffff;
          font-size: 12px; color: #111111;
        }
        .foot b { color: #111111; }
        #footMsg { flex: 1 1 auto; min-width: 0; white-space: normal; overflow-wrap: anywhere; }
        .status {
          margin: 8px 10px 0; padding: 10px 12px;
          border: 2px solid #111111; border-radius: 8px;
          background: #ffffff; color: #111111;
          font-size: 16px; font-weight: 800; line-height: 1.35;
          overflow-wrap: anywhere;
        }
        .status.found { background: #e8f6ec; border-color: #0b6b2f; }
        .status.missing, .status.error { background: #fff4f4; border-color: #b42318; }
        #dashboard { font-weight: 800; }
        .dash {
          display: none;
          position: fixed;
          inset: 0;
          z-index: 35;
          overflow: auto;
          pointer-events: auto;
          background: rgba(17, 17, 17, 0.45);
        }
        .dash.open { display: block; }
        .dash-card {
          width: min(980px, calc((var(--plf-vw) - 24px) / var(--plf-boost, 1)));
          margin: 12px auto;
          display: flex;
          flex-direction: column;
          gap: 12px;
          padding: 14px;
          background: #ffffff;
          color: #111111;
          border: 3px solid #111111;
          border-radius: 12px;
          box-shadow: 0 18px 50px rgba(0,0,0,.35);
        }
        .dash-bar {
          display: flex; align-items: center; gap: 8px;
          position: sticky; top: 0; z-index: 1;
          background: #ffffff; padding-bottom: 8px;
        }
        .dash-title { flex: 1 1 auto; font-size: 22px; font-weight: 900; }
        .dash-who { font-size: 13px; font-weight: 800; }
        .dash-totals { display: flex; flex-wrap: wrap; gap: 8px; }
        .dash-totals div {
          min-width: 72px;
          border: 1px solid #111111;
          border-radius: 8px;
          padding: 6px 8px;
        }
        .dash-totals b { display: block; font-size: 20px; font-variant-numeric: tabular-nums; }
        .dash-totals span { font-size: 11px; font-weight: 800; letter-spacing: .04em; }
        .dash-charts { display: flex; flex-wrap: wrap; gap: 16px; align-items: center; }
        .dash-pie { flex: 0 0 180px; }
        .dash-pie svg { width: 180px; height: 180px; display: block; }
        .dash-legend { flex: 1 1 180px; display: grid; gap: 6px; }
        .dash-key { display: flex; align-items: center; gap: 8px; font-weight: 800; }
        .swatch { width: 14px; height: 14px; border-radius: 3px; border: 1px solid #111111; }
        .dash-section { font-size: 12px; font-weight: 800; letter-spacing: .08em; }
        .barrow {
          display: grid;
          grid-template-columns: minmax(90px, 180px) minmax(40px, 1fr) 36px;
          gap: 8px;
          align-items: center;
          margin: 4px 0;
        }
        .barlabel { font-size: 12px; font-weight: 800; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .bartrack { height: 14px; background: #eef1f5; border-radius: 999px; overflow: hidden; }
        .barfill { height: 100%; background: #111111; border-radius: 999px; min-width: 2px; }
        .barfill.parent { background: #1f4b99; }
        .barcount { font-size: 12px; font-weight: 800; text-align: right; font-variant-numeric: tabular-nums; }
        .dash-loc { margin-top: 10px; font-size: 14px; font-weight: 900; }
        .dash-parent { margin: 6px 0 2px 8px; font-size: 13px; font-weight: 800; }
        button.dash-line {
          display: flex;
          gap: 8px;
          width: 100%;
          box-sizing: border-box;
          padding-left: 16px;
          text-align: left;
          border: 0;
          border-bottom: 1px solid #eef1f5;
          border-radius: 0;
          background: #ffffff;
          font-weight: 700;
        }
        .tag {
          margin-left: auto;
          border-radius: 999px;
          padding: 1px 8px;
          font-size: 11px;
          font-weight: 800;
        }
        .tag-alert { background: #b42318; color: #ffffff; }
        .tag-caution { background: #f5c518; color: #111111; }
        .tag-good { background: #0d7a3f; color: #ffffff; }
        .tag-miss { background: #fdeceb; color: #b42318; }
        .tag-err { background: #fff4e5; color: #b54708; }
        .tag-none { background: #eef1f5; color: #111111; }
        .locprompt, .dbg, .sharebox, .excelock, .accessdeny {
          display: none;
          position: fixed;
          left: 50%; top: 50%; transform: translate(-50%, -50%);
          max-width: calc((var(--plf-vw) - 24px) / var(--plf-boost, 1)); max-height: calc((var(--plf-vh) - 24px) / var(--plf-boost, 1));
          height: auto; overflow: auto; pointer-events: auto;
        }
        .locprompt.open, .dbg.open, .sharebox.open, .excelock.open, .accessdeny.open { display: block; }
        .locprompt, .sharebox, .excelock { z-index: 40; width: min(640px, calc((var(--plf-vw) - 24px) / var(--plf-boost, 1))); }
        .howto { margin: 0; padding-left: 28px; }
        .howto li { font-size: 18px; font-weight: 800; line-height: 1.35; margin: 8px 0; }
        .sharebox { z-index: 41; }
        .dbg { z-index: 50; width: min(640px, calc((var(--plf-vw) - 24px) / var(--plf-boost, 1))); }
        .excelock { z-index: 60; }
        .accessdeny { z-index: 70; width: min(720px, calc((var(--plf-vw) - 24px) / var(--plf-boost, 1))); }
        .locprompt-card {
          width: min(640px, calc((var(--plf-vw) - 24px) / var(--plf-boost, 1)));
          display: flex; flex-direction: column; gap: 8px;
          padding: 14px; background: #ffffff; color: #111111;
          border: 3px solid #111111; border-radius: 12px;
          box-shadow: 0 18px 50px rgba(0,0,0,.35);
        }
        .locprompt-title { font-size: 16px; font-weight: 800; }
        .locprompt-for { font-size: 13px; font-weight: 700; word-break: break-word; }
        .locprompt-list { display: grid; gap: 6px; overflow: auto; max-height: min(280px, calc(var(--plf-vh) * 0.46 / var(--plf-boost, 1))); }
        .locprompt-list button, .locprompt-actions button { text-align: left; }
        .locprompt-actions { display: flex; gap: 8px; }
        .locprompt-actions button { flex: 1; text-align: center; }
        .dbg-card {
          width: min(640px, calc((var(--plf-vw) - 24px) / var(--plf-boost, 1)));
          display: flex; flex-direction: column; gap: 8px;
          padding: 14px; background: #ffffff; color: #111111;
          border: 3px solid #111111; border-radius: 12px;
          box-shadow: 0 18px 50px rgba(0,0,0,.35);
        }
        .dbg-title { font-size: 16px; font-weight: 800; }
        #dbgText {
          width: 100%; min-height: 160px; max-height: min(360px, calc((var(--plf-vh) - 220px) / var(--plf-boost, 1))); resize: vertical;
          border: 1px solid #111111; border-radius: 8px; padding: 8px;
          background: #ffffff; color: #111111;
          font: 12px/1.4 Consolas, "Courier New", monospace;
        }
        .dbg-actions { display: flex; gap: 8px; justify-content: flex-end; }
        .sharebox-card {
          width: min(640px, calc((var(--plf-vw) - 24px) / var(--plf-boost, 1)));
          display: flex; flex-direction: column; gap: 10px;
          padding: 16px; background: #ffffff; color: #111111;
          border: 3px solid #111111; border-radius: 12px;
          box-shadow: 0 18px 50px rgba(0,0,0,.35);
        }
        .sharebox-title { font-size: 22px; font-weight: 900; }
        .sharebox-lead { margin: 0; font-size: 16px; font-weight: 700; line-height: 1.35; }
        .sharebox-url {
          font: 700 13px/1.4 Consolas, "Courier New", monospace;
          word-break: break-all;
          border: 2px solid #111111; border-radius: 8px; padding: 8px 10px;
          background: #f7f7f7; color: #111111;
        }
        #shareStatus { min-height: 1.3em; font-size: 15px; font-weight: 800; }
        #shareGo {
          height: 52px; font-size: 18px; font-weight: 900;
          background: #111111; color: #ffffff; border: 2px solid #111111;
        }
        #shareClose { height: 40px; font-weight: 800; }
        .excelock-card {
          width: min(640px, calc((var(--plf-vw) - 24px) / var(--plf-boost, 1)));
          display: flex; flex-direction: column; gap: 10px;
          padding: 16px; background: #ffffff; color: #111111;
          border: 3px solid #111111; border-radius: 12px;
          box-shadow: 0 18px 50px rgba(0,0,0,.35);
        }
        .excelock-title { font-size: 22px; font-weight: 900; }
        .excelock-file {
          font: 700 15px/1.4 Consolas, "Courier New", monospace;
          word-break: break-all;
          border: 1px solid #111111; border-radius: 8px; padding: 8px 10px;
        }
        .excelock-msg { font-size: 14px; font-weight: 700; }
        #excelockRetry {
          height: 42px; font-size: 16px; font-weight: 800; letter-spacing: .08em;
          background: #111111; color: #ffffff; border: 1px solid #111111;
        }
        #excelockRetry:disabled { opacity: .55; cursor: default; }
        .accessdeny-card {
          width: min(720px, calc((var(--plf-vw) - 24px) / var(--plf-boost, 1)));
          display: flex; flex-direction: column; gap: 12px;
          padding: 18px; background: #ffffff; color: #111111;
          border: 6px solid #111111; border-radius: 16px;
          box-shadow: 0 18px 50px rgba(0,0,0,.35);
        }
        .accessdeny-stop {
          font-size: 56px; font-weight: 900; letter-spacing: .08em; line-height: 1;
          background: #b42318; color: #ffffff; text-align: center; padding: 10px 12px;
        }
        .accessdeny-title { font-size: 30px; font-weight: 900; line-height: 1.15; text-align: center; }
        .accessdeny-file {
          font: 800 20px/1.35 Consolas, "Courier New", monospace;
          word-break: break-all;
          border: 4px solid #b42318; background: #fff4f4; padding: 10px 12px;
        }
        .accessdeny-steps { margin: 0; padding-left: 32px; }
        .accessdeny-steps li { font-size: 22px; font-weight: 800; line-height: 1.3; margin: 8px 0; }
        #accessMsg { font-size: 18px; font-weight: 800; line-height: 1.35; }
        #accessOpen, #accessRetry {
          height: 68px; font-size: 26px; font-weight: 900; letter-spacing: .04em;
          border: 4px solid #111111; cursor: pointer;
        }
        #accessOpen { background: #f5c518; color: #111111; }
        #accessRetry { background: #111111; color: #ffffff; }
        #accessOpen:disabled, #accessRetry:disabled { opacity: .55; cursor: default; }
        .fab {
          display: none; position: absolute; right: 18px; bottom: 18px; z-index: 2;
          pointer-events: auto;
          border-radius: 999px; padding: 12px 16px; cursor: pointer;
          background: #ffffff; color: #111111; border: 1px solid #111111; font-weight: 800; font-size: 14px;
          box-shadow: 0 10px 28px rgba(0,0,0,.4);
        }
        :host(.closed) .wrap { display: none; }
        :host(.closed) .fab { display: block; pointer-events: auto !important; }
        :host(.closed) { pointer-events: none !important; }
        @container plf (max-width: 720px) {
          .split { grid-template-columns: 1fr; }
          .history { border-left: 0; border-top: 1px solid #e4e7ec; }
          .modal { max-height: calc((var(--plf-vh) - 12px) / var(--plf-boost, 1)); }
          .lineswrap, .histwrap { max-height: min(160px, calc((var(--plf-vh) - 280px) / var(--plf-boost, 1))); }
        }
      </style>
      <div class="plf-stage" id="stage">
      <div class="wrap" id="wrap">
        <div class="backdrop" id="backdrop"></div>
        <div class="modal" id="modal">
          <div class="bar">
            <div class="brand">${SCRIPT_NAME} <span>${VERSION}</span></div>
            <div class="who" id="opTitle" title="${SCRIPT_NAME} ${VERSION}">Operator</div>
            <div class="lamps">
              <div class="lamp" id="lampPec" title="Inventory page"><span class="dot"></span>Page</div>
              <div class="lamp" id="lampSp" title="SharePoint"><span class="dot"></span>SP</div>
            </div>
            <div class="counts" id="counts"></div>
                <div class="clock" id="clock">00:00:00</div>
            <button id="dashboard" type="button" title="Charts for today's scans">Dashboard</button>
            <button id="share" type="button" title="Share the Tampermonkey install link">Share</button>
            <button id="debug" title="Copy debug JSON">Debug</button>
            <select id="panelSize" title="Panel size. Screen matches your display. +25% and +50% make this panel bigger.">
              <option value="1">Screen</option>
              <option value="1.25">+25%</option>
              <option value="1.5">+50%</option>
            </select>
              <button class="x" id="close" title="Close">×</button>
            </div>
          <div class="scanrow">
            <select id="location" title="Location"></select>
            <button id="delLoc" type="button" disabled title="Remove this custom location">Delete</button>
            <input id="q" type="text" placeholder="Scan container" autocomplete="off" spellcheck="false" />
            </div>
          <div class="status ready" id="note">Ready</div>
            <div class="ageboard" id="ageBoard">
            <div class="age-kicker" id="ageKicker">AGE</div>
            <div class="age-val" id="ageVal">Scan</div>
                <div class="age-tag" id="ageTag">—</div>
              </div>
          <div class="identity" id="identity"><span class="empty">No container yet</span></div>
          <div class="split">
          <div class="ticket">
              <div class="lineswrap">
            <table>
              <thead><tr><th>Age</th><th>FcSku</th><th>Qty</th><th>Owner</th><th>Disposition</th></tr></thead>
                  <tbody id="lines"><tr><td colspan="5" class="empty">No lines</td></tr></tbody>
            </table>
              </div>
          </div>
          <div class="history">
              <div class="histhead">
                <span>HISTORY</span>
                <input id="filter" placeholder="Filter" />
              </div>
            <div class="histwrap">
              <table>
                  <thead><tr><th>Time</th><th>Container</th><th>Age</th><th>St</th></tr></thead>
                <tbody id="hist"></tbody>
              </table>
              </div>
            </div>
          </div>
          <div class="foot">
            <span id="footMsg">File saves on the first scan. Keep it closed in Excel.</span>
          </div>
        </div>
      </div>
      <div class="dbg" id="dbg">
        <div class="dbg-card">
          <div class="dbg-title">Debug</div>
          <ol class="howto">
            <li>Click Copy.</li>
            <li>Open the chat where you were asked for this.</li>
            <li>Click the message box.</li>
            <li>Press Ctrl+V to paste.</li>
            <li>Press Enter to send it.</li>
            <li>Click Close.</li>
            <li>Click the scan box and keep scanning.</li>
          </ol>
          <textarea id="dbgText" spellcheck="false"></textarea>
          <div class="dbg-actions">
            <button id="dbgCopy" type="button">Copy</button>
            <button id="dbgClose" type="button">Close</button>
          </div>
        </div>
      </div>
      <div class="sharebox" id="sharebox">
        <div class="sharebox-card">
          <div class="sharebox-title">Share the install link</div>
          <ol class="howto">
            <li>Click the black SHARE THE INSTALL LINK button.</li>
            <li>Open Chime or email.</li>
            <li>Click the message box. Press Ctrl+V. Send it.</li>
            <li>Tell them to open that link in Firefox.</li>
            <li>On the page that opens, they click Install.</li>
            <li>They click the Peculiar tab to come back.</li>
            <li>They press F5 to refresh. Then they can scan.</li>
          </ol>
          <div class="sharebox-url" id="shareUrl"></div>
          <div id="shareStatus"></div>
          <button id="shareGo" type="button">SHARE THE INSTALL LINK</button>
          <button id="shareClose" type="button">Close</button>
        </div>
      </div>
      <div class="excelock" id="excelock">
        <div class="excelock-card">
          <div class="excelock-title">Excel file is open</div>
          <div class="excelock-file" id="excelockFile"></div>
          <div class="excelock-msg" id="excelockMsg">Close that file in Excel, then press RETRY.</div>
          <ol class="howto">
            <li>Read the file name in the box above.</li>
            <li>Click Excel on the taskbar.</li>
            <li>Find that file. Click the X to close it.</li>
            <li>Come back to this Peculiar tab. Do not press F5.</li>
            <li>Click the black RETRY button.</li>
            <li>This box stays here until the save works.</li>
          </ol>
          <button id="excelockRetry" type="button">RETRY</button>
        </div>
      </div>
      <div class="accessdeny" id="accessdeny">
        <div class="accessdeny-card">
          <div class="accessdeny-stop">STOP</div>
          <div class="accessdeny-title" id="accessTitle">YOUR SCANS WERE NOT SAVED</div>
          <div class="accessdeny-file" id="accessFile"></div>
              <ol class="accessdeny-steps">
                <li>Click the yellow OPEN THE FOLDER button.</li>
                <li>A new tab opens. Wait.</li>
                <li>If Microsoft asks you to sign in, sign in with your Amazon account.</li>
                <li>Wait until the SharePoint folder finishes loading. You should see the files.</li>
                <li>Leave that tab open.</li>
                <li>Click the Peculiar tab at the top of Firefox to come back.</li>
                <li>Click the black RETRY button on this box.</li>
                <li>If it still says STOP, you need Can edit. Ask for that, then click RETRY again.</li>
              </ol>
          <div id="accessMsg">This box stays here until the save works. RETRY is the only way to close it.</div>
          <button id="accessOpen" type="button">OPEN THE FOLDER</button>
          <button id="accessRetry" type="button">RETRY</button>
        </div>
      </div>
      <div class="dash" id="dash">
        <div class="dash-card">
          <div class="dash-bar">
            <div class="dash-title">Dashboard</div>
            <button id="dashClose" type="button">Close</button>
          </div>
          <div id="dashBody"></div>
        </div>
      </div>
      <div class="locprompt" id="locprompt">
        <div class="locprompt-card">
          <div class="locprompt-title">Select a location</div>
          <ol class="howto">
            <li>Click the location button for where you are.</li>
            <li>If it is not listed, click Other and type the name.</li>
            <li>The scan saves after you pick one.</li>
            <li>Click Cancel only if this scan should not be saved.</li>
          </ol>
          <div class="locprompt-for" id="locpromptFor"></div>
          <div class="locprompt-list" id="locpromptList"></div>
          <div class="locprompt-actions">
            <button id="locpromptOther" type="button">Other…</button>
            <button id="locpromptCancel" type="button">Cancel</button>
          </div>
        </div>
      </div>
      </div>
      <div class="plf-fab-dock"><button class="fab" id="fab">Line Flush</button></div>`;
    document.documentElement.append(hostEl);

    const input = shadow.getElementById("q");
    const note = shadow.getElementById("note");

    function isTextField(el) {
      return (
        el instanceof HTMLElement &&
        (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.tagName === "SELECT" || el.isContentEditable)
      );
    }

    function anotherFieldHasFocus() {
      const inner = shadow.activeElement;
      if (inner instanceof HTMLElement && inner !== input && isTextField(inner)) return true;
      const outer = document.activeElement;
      if (!(outer instanceof HTMLElement)) return false;
      if (outer === hostEl || outer === document.body || outer === document.documentElement) return false;
      if (outer.id === "searchContainerInput" || outer.closest("#searchContainerButton")) return false;
      return isTextField(outer);
    }

    function focusScanBox() {
      if (!open || excelLockOpen || accessDeniedOpen || anotherFieldHasFocus()) return;
      if (shadow.getElementById("sharebox")?.classList.contains("open")) return;
      if (shadow.getElementById("dbg")?.classList.contains("open")) return;
      releasePageButton(document);
      if (shadow.activeElement === input) return;
      input.focus({ preventScroll: true });
    }
    restoreScanFocus = focusScanBox;
    const scanQueue = [];
    const heldSaves = [];
    let working = false;
    let excelLockOpen = false;
    let excelLockRetrying = false;
    let accessDeniedOpen = false;
    let open = !hostEl.classList.contains("closed");
    let toastedNotice = "";
    let toastedError = "";
    let alias = "";
    let spOk = false;
    let selectedId = "";

    function setBanner(kind, html) {
      note.className = "status " + (kind || "info");
      note.innerHTML = html;
    }

    function showExcelLock(fileName, stillOpen) {
      const file = fileName || sosCache.name || "today's workbook";
      hideAccessDenied();
      excelLockOpen = true;
      setOpen(true);
      shadow.getElementById("excelockFile").textContent = file;
      shadow.getElementById("excelockMsg").textContent = stillOpen
        ? "That file is still open in Excel. Close it, then press RETRY."
        : "Close that file in Excel, then press RETRY. This stays up until the save works.";
      shadow.getElementById("excelock").classList.add("open");
      const button = shadow.getElementById("excelockRetry");
      button.disabled = false;
      button.focus();
      setBanner("error", `<strong>Close Excel.</strong> ${escapeHtml(file)}`);
    }

    function showAccessDenied(stillBlocked) {
      const who = alias || sosCache.alias || "you";
      const file = sosCache.name || (who === "you" ? "today's workbook" : fileNameFor(who));
      accessDeniedOpen = true;
      hideExcelLock();
      setOpen(true);
      const signedOut = /\b401\b/.test(String(lastError || ""));
      shadow.getElementById("accessTitle").textContent = signedOut
        ? "SHAREPOINT IS NOT SIGNED IN"
        : stillBlocked
          ? "STILL BLOCKED. THE FILE WAS NOT CREATED."
          : heldSaves.length
            ? "YOUR SCANS WERE NOT SAVED"
            : "YOU CANNOT SAVE YET";
      shadow.getElementById("accessFile").textContent = `${who}  ·  ${file}`;
      shadow.getElementById("accessMsg").textContent = signedOut
        ? "Click the yellow button. Sign in if Microsoft asks. Wait for the folder. Come back here and press RETRY."
        : stillBlocked
          ? "SharePoint blocked you again. Click OPEN THE FOLDER, wait for that page to load, then press RETRY."
          : "SharePoint blocked you. No file was created. This box stays here until the save works. RETRY is the only way to close it.";
      shadow.getElementById("accessdeny").classList.add("open");
      const button = shadow.getElementById("accessRetry");
      button.disabled = false;
      shadow.getElementById("accessOpen").disabled = false;
      button.focus();
      setBanner("error", `<strong>STOP.</strong> Open the SharePoint folder, then press RETRY.`);
    }

    function hideAccessDenied() {
      accessDeniedOpen = false;
      shadow.getElementById("accessdeny").classList.remove("open");
      shadow.getElementById("accessRetry").disabled = false;
      shadow.getElementById("accessOpen").disabled = false;
    }

    function holdAccessDenied(err) {
      rememberHeldSave(err);
      showAccessDenied(accessDeniedOpen);
    }

    function hideExcelLock() {
      excelLockOpen = false;
      excelLockRetrying = false;
      shadow.getElementById("excelock").classList.remove("open");
      shadow.getElementById("excelockRetry").disabled = false;
    }

    function rememberHeldSave(err) {
      if (err?.record) {
        heldSaves.push({
          record: err.record,
          rows: [[err.record.container]],
          alias: err.alias,
          location: err.location || err.record.location,
        });
        return;
      }
      if (Array.isArray(err?.rows) && err.rows.length) {
        heldSaves.push({ rows: err.rows, alias: err.alias, location: err.location });
      }
    }

    function holdLockedSave(err) {
      rememberHeldSave(err);
      const already = excelLockOpen && shadow.getElementById("excelock").classList.contains("open");
      if (already) {
        setOpen(true);
        return;
      }
      showExcelLock(err?.fileName || sosCache.name, false);
    }

    async function retryFolderAccess() {
      const who = alias || (await currentAlias());
      alias = who;
      shadow.getElementById("opTitle").textContent = who;
      await assertSosFolder();
      await ensureSosCache(who);
      if (sosCache.book && sosCache.name) {
        await writeSosFile(sosCache.name, workbookBytes(sosCache.book));
      }
      spOk = true;
      fileOk = true;
      lastError = "";
    }

    function retryHeldSaves() {
      if (excelLockRetrying) return;
      const jobs = heldSaves.splice(0, heldSaves.length);
      excelLockRetrying = true;
      shadow.getElementById("excelockRetry").disabled = true;
      shadow.getElementById("accessRetry").disabled = true;
      const finishFail = (err) => {
        excelLockRetrying = false;
        shadow.getElementById("excelockRetry").disabled = false;
        shadow.getElementById("accessRetry").disabled = false;
        if (err?.code === "ACCESS_DENIED" || isAccessDenied(0, err?.message)) {
          showAccessDenied(true);
          return;
        }
        if (err?.code === "FILE_LOCK") {
          showExcelLock(err.fileName || sosCache.name, true);
          return;
        }
        showExcelLock(err?.fileName || sosCache.name, true);
        shadow.getElementById("excelockMsg").textContent =
          `${err?.message || "Save failed."} Close the file, then press RETRY.`;
        shadow.getElementById("excelockRetry").focus();
      };
      const run = jobs.length
        ? enqueueSave(async () => {
            for (let i = 0; i < jobs.length; i++) {
              try {
                const record = jobs[i].record;
                if (!record) throw new Error("This scan could not be retried. Scan it again.");
                await appendSosRows(record, jobs[i].alias, jobs[i].location || record.location);
              } catch (err) {
                err.remaining = jobs.slice(i);
                throw err;
              }
            }
          })
        : retryFolderAccess();
      Promise.resolve(run)
        .then(() => {
          hideExcelLock();
          hideAccessDenied();
          excelLockRetrying = false;
          shadow.getElementById("excelockRetry").disabled = false;
          shadow.getElementById("accessRetry").disabled = false;
          fileOk = true;
          lastError = "";
          tick();
          if (!jobs.length) {
            setBanner("ready", `Saved. ${escapeHtml(sosCache.name || "The file")} is in the folder.`);
            focusScanBox();
            return;
          }
          const ids = [];
          for (const job of jobs) {
            for (const row of job.rows) {
              const id = cellText(row[0]);
              if (id && !ids.includes(id)) ids.push(id);
            }
          }
          setBanner("found", `<strong>SAVED</strong> ${ids.map((id) => escapeHtml(id)).join(", ")}`);
          focusScanBox();
        })
        .catch((err) => {
          const rest = Array.isArray(err?.remaining) ? err.remaining : jobs;
          if (rest.length) heldSaves.unshift(...rest);
          finishFail(err);
        });
    }

    function paintAge(result) {
      const board = shadow.getElementById("ageBoard");
      const kicker = shadow.getElementById("ageKicker");
      const val = shadow.getElementById("ageVal");
      const tag = shadow.getElementById("ageTag");
      const idle = { bg: "#ffffff", fg: "#111111", border: "#d8dee8" };
      const apply = (tone, kickerText, valueText, tagText) => {
        const ink = contrastInk(tone.bg);
        board.style.background = tone.bg;
        board.style.color = ink;
        board.style.borderColor = tone.border;
        kicker.style.color = ink;
        val.style.color = ink;
        tag.style.color = ink;
        kicker.textContent = kickerText;
        val.textContent = valueText;
        tag.textContent = tagText;
      };
      if (!result || (!result.oldestAge && result.ok !== true && !result.error)) {
        apply(idle, "AGE", result?.containerId || "Scan", "—");
        return;
      }
      if (result.error === "NOT_FOUND" || result.error === "NOT FOUND") {
        apply({ bg: "#e24b4b", border: "#b71c1c" }, "AGE", "No record", "ALERT");
        return;
      }
      if (result.ok === false && result.error) {
        apply({ bg: "#e6a23c", border: "#c4892c" }, "AGE", result.error, "ALERT");
        return;
      }
      const raw = cellText(result.oldestAge);
      const days = parseAgeDays(raw);
      const tone = ageTone(days);
      apply(tone, "AGE", raw || "n/a", tone.label);
    }

    function flashPanel() {
      const modal = shadow.getElementById("modal");
      const board = shadow.getElementById("ageBoard");
      const color = board && board.style.background;
      if (!modal || !color || color === "#ffffff") return;
      modal.style.setProperty("--plf-glow", color);
      modal.classList.remove("glow");
      void modal.offsetWidth;
      modal.classList.add("glow");
    }

    function clockNow() {
      const d = new Date();
      const p = (n) => String(n).padStart(2, "0");
      return `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
    }

    function counts() {
      const ok = recent.filter((r) => r.status === "OK").length;
      const miss = recent.filter((r) => r.status === "NOT FOUND" || r.status === "NOT_FOUND").length;
      const err = recent.filter((r) => r.status !== "OK" && r.status !== "NOT FOUND" && r.status !== "NOT_FOUND").length;
      return { total: recent.length, ok, miss, err };
    }

    function renderStats() {
      const c = counts();
      shadow.getElementById("counts").innerHTML =
        `<b>${c.total}</b> · <b class="st-ok">${c.ok}</b> ok · <b class="st-miss">${c.miss}</b> miss · <b class="st-err">${c.err}</b> err`;
    }

    function showTicket(id, result) {
      const container = cellText((result && result.containerId) || id);
      const type = cellText(result?.containerType);
      const units = cellText(result?.totalUnits);
      const parent = cellText(result?.parentContainer);
      const parentType = cellText(result?.parentType);
      const grand = cellText(result?.grandparent);
      const bits = [
        container ? container + (type ? ` [${type}]` : "") : "",
        units ? `${units} units` : "",
        parent ? `Parent ${parent}${parentType ? ` [${parentType}]` : ""}` : "",
        grand ? `Grandparent ${grand}` : "",
      ].filter(Boolean);
      shadow.getElementById("identity").innerHTML = bits.length
        ? bits.map((bit) => escapeHtml(bit)).join(" · ")
        : `<span class="empty">No container yet</span>`;
      paintAge(result);
      const lines = result?.inventory || [];
      shadow.getElementById("lines").innerHTML = lines.length
        ? lines
            .map((r) => {
              const sku = cellText(r.fcSku);
              let paint = "";
              if (sku) {
                const tone = /^unowned$/i.test(cellText(r.owner)) ? ageTone(99) : ageTone(parseAgeDays(r.age));
                if (tone.bg !== "#ffffff") paint = ` class="aged" style="--row:${tone.bg};--ink:${tone.fg}"`;
              }
              return `<tr${paint}><td>${dash(r.age)}</td><td>${skuLink(r.fcSku)}</td><td>${dash(r.quantity)}</td><td>${dash(r.owner)}</td><td>${dash(r.disposition)}</td></tr>`;
            })
            .join("")
        : `<tr><td colspan="5" class="empty">No inventory lines</td></tr>`;
    }

    function renderHistory() {
      const q = cellText(shadow.getElementById("filter").value).toLowerCase();
      const rows = recent
        .filter((r) => !q || `${r.id} ${r.sku} ${r.age}`.toLowerCase().includes(q))
        .slice(0, 60);
      shadow.getElementById("hist").innerHTML = rows.length
        ? rows
            .map((r) => {
              const st =
                r.status === "OK" ? "st-ok" : r.status === "NOT FOUND" || r.status === "NOT_FOUND" ? "st-miss" : "st-err";
              const active = r.id === selectedId ? "active" : "";
              const status = r.status === "OK" ? "OK" : r.status === "NOT FOUND" || r.status === "NOT_FOUND" ? "MISS" : "ERR";
              return `<tr class="${active}" data-id="${escapeHtml(r.id)}"><td>${escapeHtml((r.t || "").slice(11, 16))}</td><td>${escapeHtml(r.id)}</td><td>${escapeHtml(r.age || "")}</td><td class="${st}">${status}</td></tr>`;
            })
            .join("")
        : `<tr><td colspan="4" class="empty">No scans</td></tr>`;
    }

    function persist() {
      if (alias) saveHistory(alias, recent);
    }

    function tick() {
      shadow.getElementById("clock").textContent = clockNow();
      const pec = peculiarFresh();
      shadow.getElementById("lampPec").className = "lamp" + (pec ? " ok" : " bad");
      shadow.getElementById("lampSp").className = "lamp" + (spOk ? " ok" : " bad");
      const ses = GM_getValue(KEYS.session, null);
      const today = ses?.day === dayStamp() && ses?.name;
      if (updateNotice && updateNoticeUntil && Date.now() > updateNoticeUntil) updateNotice = "";
      if (updateNotice && updateNotice !== toastedNotice) {
        toastedNotice = updateNotice;
        setBanner(/failed/i.test(updateNotice) ? "error" : "info", escapeHtml(updateNotice));
      }
      if (!updateNotice) toastedNotice = "";
      if (lastError && lastError !== toastedError) {
        toastedError = lastError;
        setBanner("error", `<strong>Save error.</strong> ${escapeHtml(lastError)}`);
      }
      if (!lastError) toastedError = "";
      if (today && fileOk) {
        const where = selectedLocation ? ` · ${escapeHtml(selectedLocation)}` : "";
        shadow.getElementById("footMsg").innerHTML =
          `<b>${escapeHtml(ses.name)}</b> · ${ses.count} containers${where}`;
      } else {
        shadow.getElementById("footMsg").textContent = "File saves on the first scan. Keep it closed in Excel.";
      }
      const outer = document.activeElement;
      focusNote =
        outer === hostEl
          ? `panel:${shadow.activeElement?.id || shadow.activeElement?.tagName || ""}`
          : `${outer?.id || outer?.tagName || ""}`;
      focusScanBox();
      maybePromptLocation();
    }

    function maybePromptLocation() {
      if (excelLockOpen || locationPromptOpen || working || !selectedLocation || !lastScanAt) return;
      if (Date.now() - lastScanAt < IDLE_LOCATION_MS) return;
      locationPromptOpen = true;
      const place = selectedLocation;
      const keep = window.confirm(`Continue with location ${place}?`);
      if (keep) {
        touchLocationClock();
      } else {
        selectedLocation = "";
        saveSelectedLocation();
        renderLocations();
        setBanner("info", "Select a location.");
      }
      locationPromptOpen = false;
    }

    function setOpen(on) {
      if (!on && (excelLockOpen || accessDeniedOpen)) return;
      open = on;
      hostEl.classList.toggle("closed", !on);
      GM_setValue(KEYS.panelOpen, !!on);
      document.documentElement.style.overflow = "";
      if (on) setTimeout(focusScanBox, 50);
    }

    shadow.getElementById("close").addEventListener("click", () => setOpen(false));
    shadow.getElementById("fab").addEventListener("click", () => setOpen(true));
    const panelSize = shadow.getElementById("panelSize");
    panelSize.value = String(panelBoost());
    panelSize.addEventListener("change", () => {
      const next = Number(panelSize.value);
      GM_setValue(KEYS.panelSize, next === 1.25 || next === 1.5 ? next : 1);
      syncPanelScale();
      if (open) setTimeout(focusScanBox, 50);
    });
    const locationSelect = shadow.getElementById("location");
    const deleteLocationButton = shadow.getElementById("delLoc");

    function renderLocations() {
      const names = allLocations();
      locationSelect.innerHTML = [
        `<option value="">Select location</option>`,
        ...names.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`),
        `<option value="${OTHER_LOCATION}">Other…</option>`,
      ].join("");
      locationSelect.value = names.includes(selectedLocation) ? selectedLocation : "";
      if (locationSelect.value !== selectedLocation) selectedLocation = locationSelect.value;
      deleteLocationButton.disabled = !customLocations.includes(selectedLocation);
    }

    customLocations = loadCustomLocations();
    selectedLocation = loadSelectedLocation();
    lastScanAt = Number(GM_getValue(KEYS.locationIdle, 0)) || 0;
    if (selectedLocation && !lastScanAt) touchLocationClock();
    renderLocations();
    document.addEventListener("visibilitychange", () => {
      if (!document.hidden) maybePromptLocation();
    });
    window.addEventListener("focus", () => maybePromptLocation());

    const pausedScans = [];

    function hideLocationPrompt() {
      locationPromptOpen = false;
      shadow.getElementById("locprompt").classList.remove("open");
    }

    function showLocationPrompt() {
      const list = shadow.getElementById("locpromptList");
      list.innerHTML = allLocations()
        .map((name) => `<button type="button" data-location="${escapeHtml(name)}">${escapeHtml(name)}</button>`)
        .join("");
      const waiting = pausedScans.map((id) => `<strong>${escapeHtml(id)}</strong>`).join(", ");
      shadow.getElementById("locpromptFor").innerHTML = waiting
        ? `Continue scan for ${waiting}`
        : "Choose where this scan belongs.";
      locationPromptOpen = true;
      shadow.getElementById("locprompt").classList.add("open");
    }

    function askLocationForScan(id) {
      const scanId = cellText(id);
      if (!scanId) return;
      if (!pausedScans.some((item) => idsMatch(item, scanId))) pausedScans.push(scanId);
      setBanner("info", `Select a location for <strong>${escapeHtml(scanId)}</strong>`);
      showLocationPrompt();
    }

    function releasePausedScans() {
      if (!selectedLocation || !pausedScans.length) return;
      const location = selectedLocation;
      const ids = pausedScans.splice(0);
      hideLocationPrompt();
      for (const scanId of ids) {
        if (scanQueue.some((job) => idsMatch(job.id, scanId))) continue;
        scanQueue.push({ id: scanId, location });
      }
      touchLocationClock();
      syncQueueLen();
      drainQueue();
      focusScanBox();
    }

    function cancelPausedScans() {
      const ids = pausedScans.splice(0);
      hideLocationPrompt();
      if (ids[0] && !input.value) input.value = ids[ids.length - 1];
      setBanner("info", "Select a location.");
      focusScanBox();
    }

    locationSelect.addEventListener("change", () => {
      const value = locationSelect.value;
      if (value === OTHER_LOCATION) {
        const entered = normalizeLocation(window.prompt("New location name") || "");
        if (!entered || entered.toLowerCase() === TOTAL_SHEET.toLowerCase()) {
          renderLocations();
          if (entered) setBanner("error", "TOTAL is reserved for the combined sheet.");
          return;
        }
        const existing = allLocations().find((name) => name.toLowerCase() === entered.toLowerCase());
        if (!existing) {
          customLocations = [...customLocations, entered];
          saveCustomLocations();
          selectedLocation = entered;
        } else {
          selectedLocation = existing;
        }
        saveSelectedLocation();
        touchLocationClock();
        renderLocations();
        setBanner("info", `Location <strong>${escapeHtml(selectedLocation)}</strong>`);
        releasePausedScans();
        input.focus();
        return;
      }
      selectedLocation = value;
      saveSelectedLocation();
      if (selectedLocation) touchLocationClock();
      deleteLocationButton.disabled = !customLocations.includes(selectedLocation);
      if (selectedLocation) {
        setBanner("info", `Location <strong>${escapeHtml(selectedLocation)}</strong>`);
        releasePausedScans();
      }
      input.focus();
    });

    deleteLocationButton.addEventListener("click", () => {
      if (!customLocations.includes(selectedLocation)) return;
      const removing = selectedLocation;
      if (!window.confirm(`Remove "${removing}" from the location list? Sheets already saved stay in the file.`)) return;
      customLocations = customLocations.filter((name) => name !== removing);
      saveCustomLocations();
      selectedLocation = "";
      saveSelectedLocation();
      renderLocations();
      setBanner("info", `Removed ${escapeHtml(removing)}`);
    });

    shadow.getElementById("locpromptList").addEventListener("click", (event) => {
      const button = event.target instanceof Element ? event.target.closest("[data-location]") : null;
      const name = button?.getAttribute("data-location") || "";
      if (!name) return;
      selectedLocation = name;
      saveSelectedLocation();
      touchLocationClock();
      renderLocations();
      setBanner("info", `Location <strong>${escapeHtml(selectedLocation)}</strong>`);
      releasePausedScans();
      input.focus();
    });

    shadow.getElementById("locpromptOther").addEventListener("click", () => {
      const entered = normalizeLocation(window.prompt("New location name") || "");
      if (!entered || entered.toLowerCase() === TOTAL_SHEET.toLowerCase()) {
        if (entered) setBanner("error", "TOTAL is reserved for the combined sheet.");
        return;
      }
      const existing = allLocations().find((name) => name.toLowerCase() === entered.toLowerCase());
      if (!existing) {
        customLocations = [...customLocations, entered];
        saveCustomLocations();
        selectedLocation = entered;
      } else {
        selectedLocation = existing;
      }
      saveSelectedLocation();
      touchLocationClock();
      renderLocations();
      setBanner("info", `Location <strong>${escapeHtml(selectedLocation)}</strong>`);
      releasePausedScans();
      input.focus();
    });

    shadow.getElementById("locpromptCancel").addEventListener("click", () => cancelPausedScans());

    shadow.getElementById("excelockRetry").addEventListener("click", () => retryHeldSaves());
    shadow.getElementById("accessRetry").addEventListener("click", () => retryHeldSaves());
    shadow.getElementById("accessOpen").addEventListener("click", () => {
      GM_openInTab(SHARE_FOLDER_URL, { active: true, insert: true });
    });
    const debugPanel = shadow.getElementById("dbg");
    const debugText = shadow.getElementById("dbgText");
    const sharePanel = shadow.getElementById("sharebox");
    const shareStatus = shadow.getElementById("shareStatus");
    shadow.getElementById("shareUrl").textContent = UPDATE_URL;

    function closeShare() {
      sharePanel.classList.remove("open");
      focusScanBox();
    }

    async function shareInstallLink() {
      let shared = false;
      try {
        if (navigator.share) {
          await navigator.share({
            title: SCRIPT_NAME,
            text: "Install Peculiar Line Flush in Tampermonkey.",
            url: UPDATE_URL,
          });
          shared = true;
        }
      } catch (err) {
        if (err && err.name === "AbortError") return;
      }
      let copied = false;
      try {
        GM_setClipboard(UPDATE_URL, "text");
        copied = true;
      } catch {
        copied = false;
      }
      if (shared) shareStatus.textContent = "Shared. Tell them: open it in Firefox, click Install, come back to Peculiar, then press F5.";
      else if (copied) shareStatus.textContent = "Copied. Now open Chime or email, click the message box, press Ctrl+V, and send it. Tell them to open it in Firefox, click Install, come back to Peculiar, and press F5.";
      else shareStatus.textContent = "Copy failed. Click the link above, press Ctrl+A, then press Ctrl+C. Then paste it in Chime.";
    }

    function paintDashboard() {
      const panel = shadow.getElementById("dash");
      if (!panel.classList.contains("open")) return;
      const saved =
        sosCache.day === dayStamp() && sosCache.alias === (alias || "") ? sosCache.records || [] : [];
      const local = recent.slice().reverse().map(historyRecord);
      const records = mergeRecords(saved, local);
      shadow.getElementById("dashBody").innerHTML = dashboardHtml(records, alias || "");
    }

    async function openDashboard() {
      const panel = shadow.getElementById("dash");
      panel.classList.add("open");
      shadow.getElementById("dashBody").innerHTML = `<div class="empty">Loading today's scans…</div>`;
      const prevError = lastError;
      try {
        const who = alias || (await currentAlias());
        alias = who;
        shadow.getElementById("opTitle").textContent = who;
        await ensureSosCache(who);
      } catch {
        /* Today's scans on this computer still show if the file cannot be read. */
      }
      lastError = prevError;
      paintDashboard();
    }

    function closeDashboard() {
      shadow.getElementById("dash").classList.remove("open");
      focusScanBox();
    }

    shadow.getElementById("dashboard").addEventListener("click", () => {
      openDashboard().catch(() => paintDashboard());
    });
    shadow.getElementById("dashClose").addEventListener("click", () => closeDashboard());
    shadow.getElementById("dashBody").addEventListener("click", (event) => {
      const line = event.target.closest("[data-container]");
      if (!line) return;
      const id = line.getAttribute("data-container") || "";
      const found = recent.find((item) => idsMatch(item.id, id));
      if (found?.result) showTicket(found.id, found.result);
      closeDashboard();
    });

    shadow.getElementById("share").addEventListener("click", () => {
      shareStatus.textContent = "";
      sharePanel.classList.add("open");
    });
    shadow.getElementById("shareGo").addEventListener("click", () => {
      shareInstallLink().catch(() => {
        shareStatus.textContent = "Copy failed. Click the link above, press Ctrl+A, then press Ctrl+C. Then paste it in Chime.";
      });
    });
    shadow.getElementById("shareClose").addEventListener("click", () => closeShare());

    function closeDebug() {
      debugPanel.classList.remove("open");
      focusScanBox();
    }

    async function openDebug() {
      let text = "";
      let copied = false;
      try {
        const result = await copyDebugText();
        text = result.text;
        copied = result.copied;
      } catch (err) {
        text = JSON.stringify({ version: VERSION, error: String(err?.message || err), focus: focusNote }, null, 2);
      }
      debugText.value = text;
      debugPanel.classList.add("open");
      debugText.focus();
      debugText.select();
      setBanner(
        copied ? "info" : "error",
        copied
          ? "<strong>Debug copied.</strong> Open the chat. Click the message box. Press Ctrl+V. Press Enter."
          : "<strong>Clipboard blocked.</strong> The text is selected. Press Ctrl+C. Then open the chat and press Ctrl+V."
      );
    }

    shadow.getElementById("debug").addEventListener("click", () => {
      openDebug().catch((err) => setBanner("error", escapeHtml(err.message || err)));
    });
    shadow.getElementById("dbgCopy").addEventListener("click", () => {
      const text = debugText.value;
      debugText.focus();
      debugText.select();
      let copied = false;
      try {
        GM_setClipboard(text, "text");
        copied = true;
      } catch {
        copied = false;
      }
      setBanner(
        copied ? "info" : "error",
        copied
          ? "<strong>Debug copied.</strong> Open the chat. Click the message box. Press Ctrl+V. Press Enter."
          : "<strong>Clipboard blocked.</strong> The text is selected. Press Ctrl+C. Then open the chat and press Ctrl+V."
      );
    });
    shadow.getElementById("dbgClose").addEventListener("click", () => closeDebug());
    shadow.getElementById("filter").addEventListener("input", renderHistory);
    shadow.getElementById("hist").addEventListener("click", (e) => {
      const tr = e.target.closest("tr[data-id]");
      if (!tr) return;
      selectedId = tr.getAttribute("data-id") || "";
      const hit = recent.find((r) => r.id === selectedId);
      if (hit?.result) showTicket(hit.id, hit.result);
      renderHistory();
    });

    function syncQueueLen() {
      scanQueueLen = scanQueue.length + (working ? 1 : 0);
    }

    function queueBanner(lookingId) {
      const n = scanQueue.length;
      const extra = n ? ` · ${n} queued` : "";
      if (lookingId) setBanner("info", `Looking up <strong>${escapeHtml(lookingId)}</strong>${extra}`);
    }

    function enqueueScan(raw) {
      const id = cellText(raw);
      if (!id) return;
      if (!selectedLocation) {
        askLocationForScan(id);
        return;
      }
      if (id.length < 4) {
        setBanner("error", "Container ID looks too short.");
        return;
      }
      if (working && idsMatch(selectedId, id)) return;
      if (scanQueue.some((job) => idsMatch(job.id, id))) return;
      touchLocationClock();
      scanQueue.push({ id, location: selectedLocation });
      syncQueueLen();
      if (working) queueBanner(selectedId);
      drainQueue();
    }

    async function drainQueue() {
      if (working) return;
      const job = scanQueue.shift();
      if (!job) {
        syncQueueLen();
        return;
      }
      working = true;
      syncQueueLen();
      try {
        await lookupOne(job.id, job.location);
      } finally {
        working = false;
        focusScanBox();
        syncQueueLen();
        if (scanQueue.length) drainQueue();
      }
    }

    async function lookupOne(id, location) {
      selectedId = id;
      showTicket(id, { containerId: id, inventory: [] });
      queueBanner(id);
      await new Promise((r) => requestAnimationFrame(() => r()));
      try {
        if (!isPeculiar && !peculiarFresh()) {
          GM_openInTab(PECULIAR_URL, false);
          const waitStart = Date.now();
          while (!peculiarFresh() && Date.now() - waitStart < 5000) await sleep(200);
        }
        if (!peculiarFresh()) throw new Error("Open Peculiar, finish Midway, leave that tab open.");
        const result = await lookup(id);
        if (result.ok) result.containerId = result.containerId || id;
        lastDebug = {
          t: nowIso(),
          containerId: id,
          ok: Boolean(result.ok),
          error: result.error || "",
          oldestAge: result.oldestAge || "",
          inventoryCount: (result.inventory || []).length,
          matched: idsMatch(result.containerId, id),
          lookupMs: result.lookupMs || 0,
          raw: clip(result.raw || "", 600),
        };
        const who = alias || (await currentAlias());
        alias = who;
        const sku = result.inventory?.[0]?.fcSku || "";
        const status = result.ok ? "OK" : result.error === "NOT_FOUND" ? "NOT FOUND" : result.error || "ERROR";
        const stored = { ...result };
        delete stored.raw;
        selectedId = id;
        const record = scanRecord(id, result, location);
        recent.unshift({
          t: nowIso(),
          id,
          age: result.oldestAge || "",
          sku,
          status,
          location,
          record,
          result: stored,
        });
        persist();
        showTicket(id, result);
        flashPanel();
        renderStats();
        renderHistory();
        paintDashboard();
        const extra = scanQueue.length ? ` · ${scanQueue.length} queued` : "";
        setBanner(
          result.ok ? "found" : result.error === "NOT_FOUND" ? "missing" : "error",
          result.ok
            ? `<strong>SAVING</strong> ${escapeHtml(id)} · ${escapeHtml(location)}${extra}`
            : `<strong>${result.error === "NOT_FOUND" ? "NO RECORD FOUND" : escapeHtml(status)}</strong> ${escapeHtml(id)} · ${escapeHtml(location)}${extra}`
        );
        enqueueSave(() => appendSosRows(record, who, location))
          .then((session) => {
            spOk = true;
            fileOk = true;
            lastError = "";
            tick();
            paintDashboard();
            if (working || !idsMatch(selectedId, id)) return;
            setBanner(
              result.ok ? "found" : result.error === "NOT_FOUND" ? "missing" : "error",
              result.ok
                ? `<strong>SAVED</strong> ${escapeHtml(id)} · ${escapeHtml(location)}`
                : `<strong>${result.error === "NOT_FOUND" ? "NO RECORD FOUND" : escapeHtml(status)}</strong> · ${escapeHtml(session.name)}`
            );
          })
          .catch((err) => {
            fileOk = false;
            tick();
            if (err?.code === "FILE_LOCK") {
              holdLockedSave(err);
              return;
            }
            if (err?.code === "ACCESS_DENIED" || isAccessDenied(0, err?.message)) {
              if (!err.record) {
                err.record = record;
                err.rows = [[record.container]];
                err.alias = who;
                err.location = location;
              }
              holdAccessDenied(err);
              return;
            }
            if (working || !idsMatch(selectedId, id)) return;
            setBanner("error", `<strong>Lookup shown, save failed.</strong> ${escapeHtml(err.message || err)}`);
          });
      } catch (err) {
        lastDebug = { t: nowIso(), error: String(err.message || err) };
        setBanner(
          "error",
          `<strong>NOT SAVED.</strong> ${escapeHtml(err.message || err)} Open SOS folder once if this is a SharePoint login issue.`
        );
      }
    }

    function run() {
      const value = cellText(input.value);
      if (!value) return;
      if (value.length < 4) {
        setBanner("error", "Container ID looks too short.");
        return;
      }
      input.value = "";
      if (!selectedLocation) {
        askLocationForScan(value);
        return;
      }
      enqueueScan(value);
      focusScanBox();
    }

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
      e.preventDefault();
        run();
      }
    });
    function takeScanKey(e) {
      e.preventDefault();
      e.stopPropagation();
      if (typeof e.stopImmediatePropagation === "function") e.stopImmediatePropagation();
    }

    function scanKeyChar(e) {
      if (e.key && e.key.length === 1) return e.key;
      const code = e.keyCode || e.which || 0;
      if (code >= 32 && code <= 126) return String.fromCharCode(code);
      return "";
    }

    onScanKey = (e) => {
      if (!open) return;
      if (excelLockOpen || accessDeniedOpen) {
        takeScanKey(e);
        if (e.key === "Enter") retryHeldSaves();
        return;
      }
      if (e.ctrlKey || e.metaKey || e.altKey) return;
      if (e.key === "Escape") {
        takeScanKey(e);
        if (shadow.getElementById("sharebox").classList.contains("open")) {
          shadow.getElementById("sharebox").classList.remove("open");
          focusScanBox();
          return;
        }
        if (debugPanel.classList.contains("open")) {
          closeDebug();
          return;
        }
        setOpen(false);
        return;
      }
      const path = typeof e.composedPath === "function" ? e.composedPath() : [];
      const filter = shadow.getElementById("filter");
      if (path.includes(input) || path.includes(filter) || anotherFieldHasFocus()) return;
      if (e.key === "Enter" || e.key === "Tab") {
        takeScanKey(e);
        run();
        return;
      }
      if (e.key === "Backspace") {
        takeScanKey(e);
        input.focus({ preventScroll: true });
        input.value = input.value.slice(0, -1);
        return;
      }
      const ch = scanKeyChar(e);
      if (!ch) return;
      takeScanKey(e);
      input.focus({ preventScroll: true });
      input.value += ch;
    };

    renderStats();
    tick();
    setOpen(true);
    setInterval(tick, 1000);
    setTimeout(focusScanBox, 200);

    currentAlias()
      .then(async (who) => {
        alias = who;
        shadow.getElementById("opTitle").textContent = who;
        const stored = loadHistory(who);
        if (stored.length) {
          recent.splice(0, recent.length, ...stored);
          renderStats();
          renderHistory();
          if (recent[0]?.result) showTicket(recent[0].id, recent[0].result);
        }
        await digest();
        await assertSosFolder();
        spOk = true;
        await ensureSosCache(who);
        fileOk = true;
        lastError = "";
        tick();
        setBanner("ready", "Ready");
      })
      .catch((err) => {
        spOk = false;
        fileOk = false;
        tick();
        if (err?.code === "ACCESS_DENIED" || isAccessDenied(0, err?.message)) {
          showAccessDenied(false);
          return;
        }
        setBanner(
          "error",
          `SharePoint save path failed. Open the Line Flush Records folder once, allow amazon-my.sharepoint.com, then refresh. ${escapeHtml(err.message || err)}`
        );
      });

    try {
      GM_registerMenuCommand("Open Line Flush Records", () => GM_openInTab(SHARE_FOLDER_URL, true));
      GM_registerMenuCommand("Open Peculiar", () => GM_openInTab(PECULIAR_URL, true));
      GM_registerMenuCommand("Copy Debug", () => copyDebugText().catch(() => {}));
    } catch {
      /* optional */
    }
  }

  const OTHER_SCAN_PANEL_IDS = [
    "tm-container-scan-panel",
    "tm-container-scan-panel-launcher",
    "tm-container-scan-panel-backdrop",
  ];

  function collectScanPanelElements() {
    const found = [];
    const seen = new Set();
    const add = (el) => {
      if (!(el instanceof Element) || seen.has(el)) return;
      if (el.id === "plf-scan-host" || el.id === "plf-old-script") return;
      seen.add(el);
      found.push(el);
    };
    for (const id of OTHER_SCAN_PANEL_IDS) add(document.getElementById(id));
    document.querySelectorAll("[id*='scan-panel'], [id*='ScanPanel'], [id*='scanpanel'], [data-role='scan-input'], .tm-title").forEach((el) => {
      const title = el.classList?.contains("tm-title") ? el.textContent || "" : "";
      const id = el.id || "";
      if (/scan-panel|scanpanel/i.test(id) || /scan panel/i.test(title) || el.getAttribute("data-role") === "scan-input") {
        add(el.closest("[id*='scan-panel'], [id*='ScanPanel'], [id*='scanpanel']") || el);
      }
    });
    return found;
  }

  function oldScriptLines() {
    const lines = [OLD_SCRIPT_NAME];
    const seen = new Set(lines.map((line) => line.toLowerCase()));
    const versions = new Set();
    for (const el of collectScanPanelElements()) {
      const title = (el.querySelector?.(".tm-title") || (el.classList?.contains("tm-title") ? el : null))?.textContent?.trim() || "";
      if (title && /scan panel/i.test(title) && !seen.has(title.toLowerCase())) {
        seen.add(title.toLowerCase());
        lines.push(title);
      }
      const version = el.getAttribute?.("data-version") || "";
      if (version) versions.add(version);
      const blob = `${el.id || ""} ${el.className || ""} ${el.getAttribute?.("title") || ""}`;
      const inline = blob.match(/\b\d+\.\d+\.\d+\b/);
      if (inline) versions.add(inline[0]);
    }
    if (versions.size) lines.push(`Version on this page: ${[...versions].join(", ")}`);
    else lines.push("Any version of this older scan panel");
    return lines;
  }

  function otherScanPanelOn() {
    return OTHER_SCAN_PANEL_IDS.some((id) => document.getElementById(id));
  }

  let lineFlushStarted = false;

  function startLineFlush() {
    if (otherScanPanelOn()) {
      showOldScriptWarning();
      return;
    }
    document.getElementById("plf-old-script")?.remove();
    if (!lineFlushStarted) {
      lineFlushStarted = true;
    startPeculiarWorker();
    }
    if (!document.getElementById("plf-scan-host")) startScanner();
  }

  function tryLineFlushAgain() {
    if (otherScanPanelOn()) {
      document.getElementById("plf-old-script")?.remove();
      showOldScriptWarning();
      return;
    }
    startLineFlush();
  }

  function showOldScriptWarning() {
    document.getElementById("plf-scan-host")?.remove();
    if (document.getElementById("plf-old-script")) return;
    const host = document.createElement("div");
    host.id = "plf-old-script";
    Object.assign(host.style, {
      position: "fixed",
      inset: "0",
      zIndex: "2147483647",
      display: "block",
      pointerEvents: "auto",
      overflow: "hidden",
    });
    paintPanelScale(host);
    const shadow = host.attachShadow({ mode: "open" });
    shadow.innerHTML = `
      <style>
        :host { all: initial; --plf-dpr: 1; --plf-vw: 100vw; --plf-vh: 100vh; }
        * { box-sizing: border-box; font-family: "Segoe UI", system-ui, sans-serif; }
        .plf-stage {
          position: absolute; left: var(--plf-place, 0px); top: var(--plf-place, 0px);
          width: var(--plf-vw); height: var(--plf-vh);
          transform: var(--plf-stage-transform, none);
          transform-origin: center center;
          pointer-events: auto;
        }
        .card {
          position: fixed; z-index: 1;
          left: 50%; top: 50%; transform: translate(-50%, -50%);
          width: min(760px, calc((var(--plf-vw) - 24px) / var(--plf-boost, 1)));
          max-height: calc((var(--plf-vh) - 24px) / var(--plf-boost, 1));
          overflow: auto;
          background: #ffffff; color: #111111;
          border: 3px solid #111111; border-radius: 14px;
          padding: 18px 18px 16px;
          display: flex; flex-direction: column; gap: 12px;
          box-shadow: 0 18px 50px rgba(0,0,0,.35);
        }
        .attn { font-size: 64px; font-weight: 900; letter-spacing: .08em; line-height: 1; text-align: center; background: #b42318; color: #ffffff; padding: 12px; }
        .lead { font-size: 28px; font-weight: 900; line-height: 1.2; margin: 0; text-align: center; }
        .box { border: 4px solid #111111; border-radius: 10px; padding: 12px; }
        .kicker { font-size: 16px; font-weight: 900; letter-spacing: .08em; }
        .name { font-size: 32px; font-weight: 900; line-height: 1.15; }
        .old { border-color: #b42318; background: #fdeceb; color: #b42318; }
        .steps { margin: 0; padding-left: 34px; }
        .steps li { margin: 10px 0; font-size: 24px; font-weight: 900; line-height: 1.25; }
        .try {
          height: 72px; border: 4px solid #111111; border-radius: 12px;
          background: #111111; color: #ffffff;
          font-size: 22px; font-weight: 900; letter-spacing: .02em;
          cursor: pointer;
        }
      </style>
      <div class="plf-stage">
        <div class="card">
          <div class="attn">STOP</div>
          <p class="lead">Turn the old script OFF before you scan.</p>
          <div class="box old">
            <div class="kicker">TURN THIS ONE OFF</div>
            <div class="name">${escapeHtml(OLD_SCRIPT_NAME)}</div>
            ${oldScriptLines().filter((line) => line.toLowerCase() !== OLD_SCRIPT_NAME.toLowerCase()).map((line) => `<div class="name">${escapeHtml(line)}</div>`).join("")}
          </div>
          <div class="box">
            <div class="kicker">LEAVE THIS ONE ON</div>
            <div class="name">${escapeHtml(SCRIPT_NAME)}</div>
            <div class="name">${escapeHtml(VERSION)}</div>
          </div>
          <ol class="steps">
            <li>Click the Tampermonkey icon. It is the puzzle piece at the top of Firefox.</li>
            <li>Find the red name above. Click its switch so it turns OFF.</li>
            <li>Leave Peculiar Line Flush ON.</li>
            <li>Click this Peculiar tab to come back.</li>
            <li>Click the black button below.</li>
            <li>If this box is still here, the old script is still on. Start again at step 1.</li>
          </ol>
          <button class="try" id="tryAgain" type="button">I TURNED THE OLD SCRIPT OFF</button>
        </div></div>`;
    shadow.getElementById("tryAgain").addEventListener("click", () => tryLineFlushAgain());
    document.documentElement.append(host);
  }

  function versionParts(value) {
    return String(value || "")
      .trim()
      .split(".")
      .map((part) => {
        const n = parseInt(part, 10);
        return Number.isFinite(n) ? n : 0;
      });
  }

  function versionIsNewer(remote, local) {
    const a = versionParts(remote);
    const b = versionParts(local);
    const n = Math.max(a.length, b.length);
    for (let i = 0; i < n; i++) {
      const diff = (a[i] || 0) - (b[i] || 0);
      if (diff > 0) return true;
      if (diff < 0) return false;
    }
    return false;
  }

  function versionInScript(text) {
    const head = String(text || "").slice(0, 4000);
    const match = head.match(/\/\/\s*@version\s+([0-9]+(?:\.[0-9]+){1,3})/);
    return match ? match[1] : "";
  }

  let updateCheckRunning = false;

  function showUpdateReady(remoteVersion) {
    const existing = document.getElementById("plf-update");
    if (existing?.dataset.remote === remoteVersion) return;
    existing?.remove();
    const host = document.createElement("div");
    host.id = "plf-update";
    host.dataset.remote = remoteVersion;
    host.style.setProperty("position", "fixed", "important");
    host.style.setProperty("inset", "0", "important");
    host.style.setProperty("z-index", "2147483647", "important");
    host.style.setProperty("display", "block", "important");
    host.style.setProperty("pointer-events", "auto", "important");
    host.style.setProperty("overflow", "hidden", "important");
    paintPanelScale(host);
    const shadow = host.attachShadow({ mode: "open" });
    shadow.innerHTML = `
      <style>
        :host { all: initial; --plf-dpr: 1; --plf-vw: 100vw; --plf-vh: 100vh; }
        * { box-sizing: border-box; font-family: "Segoe UI", system-ui, sans-serif; }
        .plf-stage {
          position: absolute; left: var(--plf-place, 0px); top: var(--plf-place, 0px);
          width: var(--plf-vw); height: var(--plf-vh);
          transform: var(--plf-stage-transform, none);
          transform-origin: center center;
          pointer-events: auto;
        }
        .card {
          position: fixed; z-index: 1;
          left: 50%; top: 50%; transform: translate(-50%, -50%);
          width: min(640px, calc((var(--plf-vw) - 24px) / var(--plf-boost, 1)));
          max-height: calc((var(--plf-vh) - 24px) / var(--plf-boost, 1));
          overflow: auto;
          background: #ffffff; color: #111111;
          border: 3px solid #111111; border-radius: 14px;
          padding: 18px 18px 16px;
          display: flex; flex-direction: column; gap: 12px;
          box-shadow: 0 18px 50px rgba(0,0,0,.35);
        }
        .attn { font-size: 28px; font-weight: 900; letter-spacing: .04em; }
        .lead { font-size: 16px; font-weight: 700; margin: 0; }
        .box { border: 2px solid #111111; border-radius: 10px; padding: 10px 12px; }
        .new { border-color: #0b6b2f; background: #e8f6ec; }
        .kicker { font-size: 12px; font-weight: 800; letter-spacing: .08em; }
        .name { font-size: 22px; font-weight: 900; line-height: 1.2; }
        .ver { font-size: 16px; font-weight: 800; }
        .steps { margin: 0; padding-left: 32px; }
        .steps li { margin: 8px 0; font-size: 22px; font-weight: 900; line-height: 1.3; }
        .go {
          height: 56px; border: 2px solid #111111; border-radius: 10px;
          background: #111111; color: #ffffff;
          font-size: 18px; font-weight: 900; cursor: pointer;
        }
      </style>
      <div class="plf-stage">
        <div class="card">
          <div class="attn">⚠️ NEW VERSION</div>
          <p class="lead">${escapeHtml(SCRIPT_NAME)} found a newer copy on its own.</p>
          <div class="box">
            <div class="kicker">ON THIS COMPUTER</div>
            <div class="name">${escapeHtml(VERSION)}</div>
          </div>
          <div class="box new">
            <div class="kicker">NEW COPY</div>
            <div class="name">${escapeHtml(remoteVersion)}</div>
          </div>
          <ol class="steps">
            <li>Click the black button at the bottom of this box.</li>
            <li>A new tab opens. That tab is Tampermonkey.</li>
            <li>On that tab, click Install.</li>
            <li>When it says the script was installed, close that tab.</li>
            <li>Click the Peculiar tab at the top of Firefox to come back.</li>
            <li>Press F5 to refresh this page.</li>
            <li>Wait until the page loads. Then scan again.</li>
          </ol>
          <button class="go" id="installUpdate" type="button">Install the new Peculiar Line Flush</button>
        </div></div>`;
    shadow.getElementById("installUpdate").addEventListener("click", () => {
      try {
        GM_openInTab(UPDATE_URL, { active: true, insert: true });
      } catch {
        window.open(UPDATE_URL, "_blank", "noopener");
      }
    });
    document.documentElement.append(host);
  }

  function noteUpdate(text, sticky) {
    updateNotice = text || "";
    updateNoticeUntil = !text || sticky ? 0 : Date.now() + 8000;
    const foot = document.getElementById("plf-scan-host")?.shadowRoot?.getElementById("footMsg");
    if (foot && updateNotice) foot.textContent = updateNotice;
  }

  async function readPublishedScript() {
    const url = `${UPDATE_URL}?plf=${Date.now()}`;
    let gmError = "";
    try {
      const res = await gmReq(url, { timeout: 15000, anonymous: true });
      if (res.status >= 200 && res.status < 300 && res.responseText) return res.responseText;
      gmError = `GitHub HTTP ${res.status || 0}`;
    } catch (err) {
      gmError = err.message || String(err);
    }
    try {
      const res = await fetch(url, { cache: "no-store", credentials: "omit" });
      if (res.ok) return await res.text();
      throw new Error(`GitHub HTTP ${res.status}`);
    } catch (err) {
      throw new Error(gmError || err.message || String(err));
    }
  }

  async function checkForScriptUpdate() {
    if (updateCheckRunning) return;
    updateCheckRunning = true;
    try {
      const text = await readPublishedScript();
      const remote = versionInScript(text);
      if (!remote || !versionIsNewer(remote, VERSION)) {
        document.getElementById("plf-update")?.remove();
        noteUpdate(`${SCRIPT_NAME} ${VERSION} is the newest version.`);
        return;
      }
      noteUpdate(`New version ${remote} is ready.`, true);
      showUpdateReady(remote);
    } catch (err) {
      noteUpdate(`Update check failed: ${err.message || err}`, true);
    } finally {
      updateCheckRunning = false;
    }
  }

  function frameShouldRun() {
    if (!isPeculiar) return false;
    if (isTop) return true;
    if (!document.querySelector("#searchContainerInput")) return false;
    try {
      if (pageWindow.top.document.querySelector("#searchContainerInput")) return false;
    } catch {
      return true;
    }
    return true;
  }

  function showBootError(message) {
    if (!document.documentElement) return;
    let el = document.getElementById("plf-boot-error");
    if (!el) {
      el = document.createElement("div");
      el.id = "plf-boot-error";
      el.style.setProperty("position", "fixed", "important");
      el.style.setProperty("inset", "0", "important");
      el.style.setProperty("z-index", "2147483647", "important");
      el.style.setProperty("overflow", "hidden", "important");
      el.style.setProperty("pointer-events", "none", "important");
      const shadow = el.attachShadow({ mode: "open" });
      shadow.innerHTML = `
        <style>
          :host { all: initial; --plf-dpr: 1; --plf-boost: 1; --plf-vw: 100vw; --plf-vh: 100vh; }
          .plf-boot-dock {
            position: absolute; left: 50%; bottom: 0; width: 0; height: 0;
            transform: var(--plf-scale, none);
            transform-origin: bottom center;
            pointer-events: none;
          }
          .card {
            position: absolute; left: 50%; bottom: 12px; transform: translateX(-50%);
            width: max-content;
            max-width: calc((var(--plf-vw) - 24px) / var(--plf-boost, 1));
            max-height: calc((var(--plf-vh) - 24px) / var(--plf-boost, 1));
            overflow: auto; pointer-events: auto;
            background: #ffffff; color: #111111;
            border: 3px solid #111111; border-radius: 12px;
            padding: 14px 16px;
            font: 800 18px/1.35 "Segoe UI", sans-serif;
            overflow-wrap: anywhere;
            box-shadow: 0 16px 40px rgba(0,0,0,.28);
          }
        </style>
        <div class="plf-boot-dock"><div class="card" id="msg"></div></div>`;
      document.documentElement.append(el);
    }
    paintPanelScale(el);
    el.shadowRoot.getElementById("msg").textContent = message;
  }

  function bootLineFlush() {
    try {
      if (!frameShouldRun()) return;
      document.getElementById("plf-boot-error")?.remove();
      if (otherScanPanelOn()) showOldScriptWarning();
      else if (!document.getElementById("plf-old-script")) startLineFlush();
    } catch (err) {
      showBootError(`${SCRIPT_NAME} did not start: ${err.message || err}`);
    }
  }

  function chooseItServiceLogin() {
    let clicked = false;
    const visit = (el) => {
      if (clicked || !el) return;
      const label = (el.innerText || el.textContent || "").replace(/\s+/g, " ").trim();
      if (label !== "IT Service Login") return;
      const button = el.closest("alchemy-button, button, a, [role='button']") || el;
      const buttonLabel = (button.innerText || button.textContent || "").replace(/\s+/g, " ").trim();
      if (buttonLabel !== "IT Service Login" || /MAASK/i.test(buttonLabel)) return;
      if (button.dataset.plfItLogin === "1") {
        clicked = true;
        return;
      }
      button.dataset.plfItLogin = "1";
      const inner = button.shadowRoot && button.shadowRoot.querySelector("button, a, [role='button']");
      (inner || button).click();
      clicked = true;
    };
    const walk = (root) => {
      if (!root || !root.querySelectorAll) return;
      root.querySelectorAll("alchemy-button, button, a, [role='button'], span.button__label").forEach((el) => {
        visit(el);
        if (el.shadowRoot) walk(el.shadowRoot);
      });
    };
    walk(document);
    return clicked;
  }

  function watchItServiceLogin() {
    if (!isLoginChooser) return;
    const tryClick = () => {
      try {
        chooseItServiceLogin();
      } catch {
        /* The login page changed shape. */
      }
    };
    const onChooser = /idp-selection|federated-auth\.aft\.a2z\.com/i.test(location.href);
    if (!onChooser) {
      let tries = 0;
      const timer = setInterval(() => {
        tries += 1;
        tryClick();
        if (tries >= 20) clearInterval(timer);
      }, 500);
      return;
    }
    tryClick();
    const arm = () => {
      if (!document.documentElement) return;
      tryClick();
      const obs = new MutationObserver(tryClick);
      obs.observe(document.documentElement, { childList: true, subtree: true });
      setTimeout(() => obs.disconnect(), 60000);
    };
    if (document.documentElement) arm();
    else document.addEventListener("DOMContentLoaded", arm, { once: true });
  }
  watchItServiceLogin();

  if (isPeculiar) {
    let armed = false;
    const tryBoot = () => {
      if (!document.documentElement) return;
      if (!frameShouldRun()) return;
      if (!armed) {
        armed = true;
        watchPanelScale();
        checkForScriptUpdate();
        setInterval(checkForScriptUpdate, 2 * 60 * 1000);
        document.addEventListener("visibilitychange", () => {
          if (document.visibilityState === "visible") {
            bootLineFlush();
            checkForScriptUpdate();
          }
        });
      }
      bootLineFlush();
    };
    tryBoot();
    [400, 1200, 3000, 6000].forEach((ms) => setTimeout(tryBoot, ms));
    setInterval(tryBoot, 2000);
  }
})();
