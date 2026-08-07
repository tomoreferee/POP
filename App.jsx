import { useState, useEffect, useRef, useCallback, useMemo, Fragment } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

// ─── Supabase Config ─────────────────────────────────────────────────────────
// TODO: set the URL and anon key from Supabase Project Settings → API
const SUPABASE_URL = "https://qqzmpkpoanltxqzfeozf.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFxem1wa3BvYW5sdHhxemZlb3pmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI5ODg0MTYsImV4cCI6MjA5ODU2NDQxNn0.20p6PlzoAjy8GSTF0pbzeZt9r6t6AgrPLzjIH5sSeK8";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  realtime: { params: { eventsPerSecond: 10 } },
});

// ─── In-memory storage (used only for the local setup-form draft, per-browser) ─────
const memoryStorage = (() => {
  const store = {};
  return {
    getItem: (key) => (key in store ? store[key] : null),
    setItem: (key, value) => { store[key] = String(value); },
    removeItem: (key) => { delete store[key]; },
  };
})();

// ─── Constants ────────────────────────────────────────────────────────────────
// Default minutes allowed per hole, by par — varies with how many players are in
// the group (a 4-ball takes noticeably longer than a 2-ball on the same hole).
const PAR_TIMES_BY_PLAYERS = {
  2: { 3: 11, 4: 13, 5: 15 },
  3: { 3: 13, 4: 15, 5: 18 },
  4: { 3: 15, 4: 18, 5: 20 },
};
function parTimeTable(playersPerGroup) {
  return PAR_TIMES_BY_PLAYERS[playersPerGroup] ?? PAR_TIMES_BY_PLAYERS[3];
}
// Bump this whenever App.jsx is updated — shown at the bottom of the Setup page so
// you can confirm at a glance whether the browser is running the newest deploy.
const APP_BUILD = "2026-08-03-e (beta)";

// Selectable minutes per hole — dropdown beats a free number field on a phone.
const PAR_TIME_CHOICES = Array.from({ length: 16 }, (_, i) => i + 10); // 10…25
// Kept for any legacy reference — mirrors the 3-ball column (the common default).
const PAR_TIMES = PAR_TIMES_BY_PLAYERS[3];

const DEFAULT_PARS = [4,4,3,4,5,4,3,4,4, 4,4,3,4,5,4,3,4,4];

const thStyle = { padding: "8px 6px", color: "#8890b8", fontWeight: 600, textAlign: "center", borderBottom: "1px solid #2a2d4a", minWidth: 38 };
const tdStyle = { padding: "8px 6px", textAlign: "center", borderBottom: "1px solid #1a1d2e" };

function buildSchedule(startTimeStr, parTimes) {
  const [h, m] = startTimeStr.split(":").map(Number);
  let totalMin = h * 60 + m;
  const schedule = [];
  for (let i = 0; i < parTimes.length; i++) {
    schedule.push(totalMin);
    totalMin += parTimes[i];
  }
  return schedule;
}

function getHoleOrder(startHole) {
  // returns array of 0-based hole indices in play order, wrapping after 18 back to 1
  const start = startHole || 1;
  return Array.from({ length: 18 }, (_, i) => (start - 1 + i) % 18);
}

// ─── Shotgun Start helpers ──────────────────────────────────────────────────
// 4-hole shotgun: H1, H5(or H6), H10, H14(or H15)
// Rules:
//  - H5/H14 must not be par3 → if it is par3, move to the next hole (H6/H15)
//  - If H5,H6 (or H14,H15) are both consecutive par3 holes → force start at H5 (or H14) even if par3
function resolveShotgunStartHoles(pars) {
  const isPar3 = (holeNum) => (pars?.[holeNum - 1] ?? 4) === 3;

  const resolvePair = (primary, secondary) => {
    if (!isPar3(primary)) return primary;          // H5/H14 is not par3 → use it directly
    if (!isPar3(secondary)) return secondary;       // H5 is par3 but H6 is not → move to H6
    return primary;                                  // both are par3 → start at H5/H14 as default
  };

  return [1, resolvePair(5, 6), 10, resolvePair(14, 15)];
}

// Color and label for each start point (supports the 4 shotgun points + any other start points)
const START_HOLE_META = {
  1:  { color: "#6effa0", label: "🟢 H1 → 18",          shortLabel: "🟢 Start hole 1 → 18" },
  5:  { color: "#ffd966", label: "🟡 H5 → 18 → 1 → 4",  shortLabel: "🟡 Start hole 5 → 18 → 1 → 4" },
  6:  { color: "#ffd966", label: "🟡 H6 → 18 → 1 → 5",  shortLabel: "🟡 Start hole 6 → 18 → 1 → 5" },
  10: { color: "#4e9af1", label: "🔵 H10 → 18 → 1 → 9", shortLabel: "🔵 Start hole 10 → 18 → 1 → 9" },
  14: { color: "#c084fc", label: "🟣 H14 → 18 → 1 → 13",shortLabel: "🟣 Start hole 14 → 18 → 1 → 13" },
  15: { color: "#c084fc", label: "🟣 H15 → 18 → 1 → 14",shortLabel: "🟣 Start hole 15 → 18 → 1 → 14" },
};
function getStartHoleMeta(startHole) {
  return START_HOLE_META[startHole] || { color: "#aaaaaa", label: `H${startHole} → ...`, shortLabel: `Start hole ${startHole}` };
}

function buildScheduleOrdered(startTimeStr, parTimes, startHole, turnTime = 1, turnTimeBack = null) {
  const order = getHoleOrder(startHole);
  const [h, m] = startTimeStr.split(":").map(Number);
  let t = h * 60 + m;
  const sch = Array(18);
  order.forEach((hi, i) => {
    // Turn-time buffer: groups starting at Hole 1 or Hole 10 get the configured
    // walking time added once they reach the 10th hole of their round (H10 for a
    // H1 start, H1 for a H10 start) — this shifts that hole and every hole after
    // it, covering the walk between the front and back nine.
    // The walk at the turn differs by direction: a H1 start crosses H9→H10,
    // a H10 start crosses H18→H1, and those two walks are rarely the same length.
    if (i === 9) {
      if (startHole === 1) t += (turnTime ?? 0);
      else if (startHole === 10) t += (turnTimeBack ?? turnTime ?? 0);
    }
    sch[hi] = t;
    t += parTimes[hi];
  });
  return sch;
}

function minToHM(totalMin) {
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  if (h <= 0) return `${m} mins.`;
  if (m === 0) return `${h} hrs.`;
  return `${h} hrs. ${m} mins.`;
}

function minToTime(min) {
  const h = Math.floor(min / 60) % 24;
  const m = min % 60;
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}`;
}

function nowInMin() {
  const d = new Date();
  return d.getHours() * 60 + d.getMinutes();
}

function secondsToMMSS(s) {
  const sign = s < 0 ? "-" : "";
  const abs = Math.abs(Math.round(s));
  const mm = Math.floor(abs / 60);
  const ss = abs % 60;
  return `${sign}${String(mm).padStart(2,"0")}:${String(ss).padStart(2,"0")}`;
}

// ─── Excel Export ───────────────────────────────────────────────────────────
function rowsToTSV(rows) {
  if (!rows.length) return "";
  const headers = Object.keys(rows[0]);
  const esc = v => String(v ?? "").replace(/\t/g, " ").replace(/\r?\n/g, " ");
  const lines = [headers.join("\t")];
  rows.forEach(r => lines.push(headers.map(h => esc(r[h])).join("\t")));
  return lines.join("\n");
}

function buildDashboardExportData({ groups, groupData, pars, parTimes, schedules }) {
  const summaryRows = [];
  const detailRows = [];
  const logRows = [];

  groups.forEach(g => {
    const gd = groupData[g.id] || {};
    const order = getHoleOrder(g.startHole || 1);
    const records = gd.records || [];
    const holeData = gd.holeData || [];
    const completed = records.filter(Boolean).length;

    // Last finished hole (by play order) → used for group status
    let lastDiff = null, lastHoleIdxForStatus = null;
    for (let s = 17; s >= 0; s--) {
      const hi = order[s];
      const hd = holeData[hi];
      if (hd?.startTime && hd?.endTime) {
        lastDiff = computeHoleDiff(hd, hi, parTimes);
        lastHoleIdxForStatus = hi;
        break;
      }
    }
    const frontDiffForStatus = getFrontGroupDiffAtHole(groups, g, lastHoleIdxForStatus, groupData, parTimes);
    const statusLabel = { ok: "On time", warn: "Less OOP", late: "OOP", idle: "Waiting" }[
      lastDiff === null ? "idle" : getRelativeStatus(lastDiff, frontDiffForStatus)
    ];

    summaryRows.push({
      Group: g.name,
      "Start Hole": g.startHole || 1,
      "Tee Time": g.startTime,
      "Holes Completed": `${completed}/18`,
      "Last Diff (min)": lastDiff !== null ? lastDiff : "",
      Status: statusLabel,
      "MN Active": gd.mnActive ? (gd.mnName || "Yes") : "",
      "TM Active": gd.tmActive ? (gd.tmName || "Yes") : "",
    });

    order.forEach(hi => {
      const hd = holeData[hi] || {};
      const sched = schedules?.[g.id]?.[hi];
      const diff = computeHoleDiff(hd, hi, parTimes);
      const frontDiff = diff !== null ? getFrontGroupDiffAtHole(groups, g, hi, groupData, parTimes) : 0;
      detailRows.push({
        Group: g.name,
        Hole: hi + 1,
        Par: pars?.[hi] ?? "",
        "Scheduled Time": sched != null ? minToTime(sched) : "",
        "Start Time": hd.startTime || "",
        "End Time": hd.endTime || "",
        "Diff (min)": diff !== null ? diff : "",
        Status: diff !== null ? { ok: "On time", warn: "Less OOP", late: "OOP" }[getRelativeStatus(diff, frontDiff)] : "",
      });
    });

    (gd.actionLogs || []).forEach(log => {
      logRows.push({
        Group: g.name,
        Hole: (log.holeIdx ?? -1) + 1,
        Type: log.type,
        Time: log.time,
        Name: log.name || "",
        Target: log.target || "",
        "Diff at log (min)": log.diff ?? "",
      });
    });
  });

  return {
    summaryRows, detailRows, logRows,
    summaryTSV: rowsToTSV(summaryRows),
    detailTSV: rowsToTSV(detailRows),
    logsTSV: rowsToTSV(logRows),
  };
}

function buildDashboardExcelBlob({ summaryRows, detailRows, logRows }) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "Summary");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detailRows), "Hole Details");
  if (logRows.length) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(logRows), "Action Logs");
  }

  const ts = new Date();
  const stamp = `${ts.getFullYear()}${String(ts.getMonth() + 1).padStart(2, "0")}${String(ts.getDate()).padStart(2, "0")}_${String(ts.getHours()).padStart(2, "0")}${String(ts.getMinutes()).padStart(2, "0")}`;
  const filename = `pace_monitor_${stamp}.xlsx`;
  const mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: mime });
  return { blob, filename, mime };
}

// Clipboard copy that works even where navigator.clipboard is blocked (sandboxed iframes)
function copyTextToClipboard(text) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    return navigator.clipboard.writeText(text).catch(() => copyViaTextarea(text));
  }
  return copyViaTextarea(text);
}
function copyViaTextarea(text) {
  return new Promise((resolve, reject) => {
    try {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      const ok = document.execCommand("copy");
      document.body.removeChild(ta);
      ok ? resolve() : reject(new Error("copy command failed"));
    } catch (err) {
      reject(err);
    }
  });
}


// ─── Status helper (reused everywhere) ────────────────────────────────────────────
// warn = 1-2 minutes, late = 3+ minutes
function getStatus(diff) {
  if (diff <= 0) return "ok";
  if (diff <= 2) return "warn";
  return "late";
}

// Diff (actual - scheduled, in minutes) for a single hole, or null if not finished
function computeHoleDiff(hd, hi, parTimes) {
  if (!hd?.startTime || !hd?.endTime) return null;
  if (hd.manualDiff !== undefined) return hd.manualDiff;
  const [sh, sm] = hd.startTime.split(":").map(Number);
  const [eh, em] = hd.endTime.split(":").map(Number);
  const actual = (eh * 60 + em) - (sh * 60 + sm);
  return Math.round(actual - (parTimes?.[hi] ?? 14)) + 1;
}

// A group's progress: diff at the most recently completed hole (in that group's play order)
function computeGroupProgress(g, gd, parTimes) {
  const order = getHoleOrder(g.startHole || 1);
  const holeData = gd?.holeData || [];
  const records = gd?.records || [];
  const completed = records.filter(Boolean).length;
  let lastDiff = null, lastHoleIdx = null, lastSlot = -1;
  for (let s = 17; s >= 0; s--) {
    const hi = order[s];
    const d = computeHoleDiff(holeData[hi], hi, parTimes);
    if (d !== null) { lastDiff = d; lastHoleIdx = hi; lastSlot = s; break; }
  }
  return { completed, lastDiff, lastHoleIdx, lastSlot, isComplete: lastSlot === 17 };
}

// Groups bucketed by starting hole ("side"), each bucket sorted by start time.
// Handles both H1/H10 two-tee starts and 4-hole shotguns the same way.
function getGroupSides(groups) {
  const timeToMin = t => { const [h, m] = (t || "0:00").split(":").map(Number); return h * 60 + m; };
  const byHole = {};
  groups.forEach(g => {
    const hole = g.startHole || 1;
    (byHole[hole] ??= []).push(g);
  });
  return Object.keys(byHole)
    .map(Number)
    .sort((a, b) => a - b)
    .map(hole => ({
      startHole: hole,
      meta: getStartHoleMeta(hole),
      groups: byHole[hole].slice().sort((a, b) => timeToMin(a.startTime) - timeToMin(b.startTime)),
    }));
}

// ─── Front-group-relative status ───────────────────────────────────────────────
// A group's status should not be judged purely against the fixed schedule — it
// should also account for the group ahead of it (same starting hole/side, earlier
// tee time). Rules:
//  • If the group ahead is itself behind schedule (diff > 0), the following group
//    is allowed that same amount of extra time before being flagged.
//  • If the group ahead is on/ahead of schedule (diff <= 0), that gives no extra
//    cushion — the following group is still judged against the normal schedule.
//  • The very first group of each side has no group ahead, so it's treated as if
//    the group ahead played every hole exactly on schedule (diff = 0).
// Single source of truth for a group's status badge — rated relative to the group
// ahead of it. Used by the Dashboard cards/table and passed into the record screen
// so the same group can never show two different badges.
function computeGroupStatusFor(g, groups, groupData, parTimes) {
  const gd = groupData?.[g.id];
  if (!gd || !gd.holeData) return "idle";
  const order = getHoleOrder(g.startHole || 1);
  for (let s = 17; s >= 0; s--) {
    const hi = order[s];
    const hd = gd.holeData[hi];
    if (hd?.startTime && hd?.endTime) {
      const diff = computeHoleDiff(hd, hi, parTimes);
      if (diff === null) return "idle";
      const frontDiff = getFrontGroupDiffAtHole(groups, g, hi, groupData, parTimes);
      return getRelativeStatus(diff, frontDiff);
    }
  }
  return "idle";
}

// Inserts a genuine empty column at the front-nine / back-nine turn. Using a real
// cell (rather than padding) means the gap never steals width from the numbers,
// which is what made them overlap in the width-locked "fit 18" view.
// One-letter badge text for the width-locked "fit 18" view: W / M / T / B.
function shortLogLabel(l) {
  if (l.badTime) return "B";
  const letter = l.type === "WN" ? "W" : l.type === "MN" ? "M" : l.type === "TM" ? "T" : l.type?.[0] || "?";
  return l.off ? `×${letter}` : letter;
}

function withTurnGap(cells, width, keyPrefix, Tag = "td") {
  const out = cells.slice();
  out.splice(9, 0, <Tag key={`turngap-${keyPrefix}`} style={{ width, minWidth: width, padding: 0, border: "none", background: "transparent" }} />);
  return out;
}

function getGroupSideIndex(groups, group) {
  const timeToMin = t => { const [h, m] = (t || "0:00").split(":").map(Number); return h * 60 + m; };
  const side = groups
    .filter(g => (g.startHole || 1) === (group.startHole || 1) && (g.section || "morning") === (group.section || "morning"))
    .slice()
    .sort((a, b) => timeToMin(a.startTime) - timeToMin(b.startTime));
  return { side, idx: side.findIndex(g => g.id === group.id) };
}

function getFrontGroup(groups, group) {
  const { side, idx } = getGroupSideIndex(groups, group);
  return idx > 0 ? side[idx - 1] : null;
}

// Diff of the group ahead at the same hole. Returns 0 (no extra allowance) if
// there's no group ahead, or if the group ahead hasn't finished that hole yet.
function getFrontGroupDiffAtHole(groups, group, holeIdx, groupData, parTimes) {
  if (holeIdx === null || holeIdx === undefined || holeIdx < 0) return 0;
  const front = getFrontGroup(groups, group);
  if (!front) return 0;
  const d = computeHoleDiff(groupData?.[front.id]?.holeData?.[holeIdx], holeIdx, parTimes);
  return d === null ? 0 : d;
}

function getRelativeStatus(ownDiff, frontDiff) {
  if (ownDiff === null || ownDiff === undefined) return "idle";
  const credit = Math.max(frontDiff ?? 0, 0);
  return getStatus(ownDiff - credit);
}

function fmtDiff(diff) {
  if (diff === null || diff === undefined) return "–";
  if (diff === 0) return "On time";
  return diff > 0 ? `+${diff} min` : `${diff} min`;
}
function diffColor(diff) {
  if (diff === null || diff === undefined) return "#666f99";
  if (diff < -10) return "#1560a8";              // < -10   → dark blue
  if (diff <= -1) return "#4e9af1";               // -9..-1  → light blue
  if (diff === 0) return "#6effa0";               // 0       → green
  if (diff <= 2) return "#ffd966";                // +1..+2  → yellow
  if (diff <= 5) return "#ff8a80";                // +3..+5  → light red
  return "#b3261e";                                // 6+      → dark red
}

// ─── WN / MN / TM log-type styling (reused everywhere logs are shown) ─────────────────
// WN = Warning, MN = Monitor (watch), TM = Timing (time a specific player)
const LOG_TYPE_META = {
  WN: { color: "#ffd966", darkBg: "#2a1a00" },
  MN: { color: "#4e9af1", darkBg: "#001a2a" },
  TM: { color: "#ff6ec7", darkBg: "#2a0020" },
};
function logColor(type) { return (LOG_TYPE_META[type] || LOG_TYPE_META.MN).color; }
function logBg(type) { return (LOG_TYPE_META[type] || LOG_TYPE_META.MN).darkBg; }

// Builds a hole-sorted, condensed summary of a group's WN/MN/TM/Bad Time history:
// WN and Bad Time are one-off flags (one chip each); MN/TM are on/off sessions that
// get collapsed into a single "start hole → off hole" entry instead of one chip per hole.
function summarizeStatusLogs(logs, mnActive, tmActive) {
  const items = [];

  logs.filter(l => l.type === "WN").forEach(l => {
    items.push({ key: `wn-${l.idx}`, type: "WN", sortHole: l.holeIdx, label: `WN @H${l.holeIdx + 1}${l.name ? ` by ${l.name}` : ""}`, idx: l.idx });
  });
  logs.filter(l => l.type === "TM" && l.badTime).forEach(l => {
    items.push({
      key: `bt-${l.idx}`, type: "TM", sortHole: l.holeIdx,
      label: `⚡ Bad Time ${l.target || ""} @H${l.holeIdx + 1}${l.name ? ` by ${l.name}` : ""}`, idx: l.idx,
    });
  });

  const runsFor = (type, isActiveNow) => {
    const entries = logs
      .filter(l => l.type === type && !l.badTime)
      .sort((a, b) => a.holeIdx - b.holeIdx || (a.off ? 1 : -1));
    const runs = [];
    let cur = null;
    entries.forEach(l => {
      if (l.off) {
        if (cur) { cur.offHole = l.holeIdx + 1; cur.offIdx = l.idx; if (l.name) cur.name = l.name; runs.push(cur); cur = null; }
        return;
      }
      if (!cur) cur = { startHole: l.holeIdx + 1, lastHole: l.holeIdx + 1, offHole: null, offIdx: null, idx: l.idx, target: l.target || "", name: l.name || "" };
      else { cur.lastHole = l.holeIdx + 1; if (l.target) cur.target = l.target; if (l.name) cur.name = l.name; }
    });
    if (cur) runs.push(cur);
    return runs.map((r, i) => {
      const isLast = i === runs.length - 1;
      const targetSuffix = r.target ? ` (${r.target})` : "";
      const bySuffix = r.name ? ` by ${r.name}` : "";
      const label = r.offHole
        ? `${type} @H${r.startHole} → Off @H${r.offHole}${targetSuffix}${bySuffix}`
        : (isLast && isActiveNow)
          ? `${type} @H${r.startHole} → In progress${targetSuffix}${bySuffix}`
          : `${type} @H${r.startHole} → Off @H${r.lastHole + 1}${targetSuffix}${bySuffix}`;
      // Only offer a delete action when there's a specific "off" event to undo —
      // deleting it re-opens the run (fixes an accidental off-at-wrong-hole tap).
      return { key: `${type}-${r.startHole}`, type, sortHole: r.startHole - 1, label, idx: r.offIdx ?? undefined, deleteTitle: r.offHole ? "Delete this off marker (turned off on the wrong hole)" : undefined };
    });
  };

  items.push(...runsFor("MN", mnActive));
  items.push(...runsFor("TM", tmActive));

  return items.sort((a, b) => a.sortHole - b.sortHole);
}

// ─── Logout Button ────────────────────────────────────────────────────────────
function LogoutButton({ onLogout }) {
  return (
    <button
      onClick={onLogout}
      style={{
        background: "#1a0d0d", border: "1px solid #ff707044", color: "#ff7070",
        borderRadius: 7, height: 30, padding: "0 12px", cursor: "pointer",
        fontFamily: "inherit", fontSize: 12, fontWeight: 700, letterSpacing: 0.5,
        whiteSpace: "nowrap", flexShrink: 0,
      }}
      onMouseEnter={e => e.currentTarget.style.borderColor = "#ff707088"}
      onMouseLeave={e => e.currentTarget.style.borderColor = "#ff707044"}
    >⏏ Log out</button>
  );
}
function useTimer() {
  const [elapsed, setElapsed] = useState(0);
  const [running, setRunning] = useState(false);
  const intervalRef = useRef(null);
  const startRef = useRef(null);
  const baseRef = useRef(0);

  const start = useCallback(() => {
    startRef.current = Date.now();
    setRunning(true);
  }, []);

  const pause = useCallback(() => {
    baseRef.current += (Date.now() - startRef.current) / 1000;
    setRunning(false);
  }, []);

  const reset = useCallback(() => {
    baseRef.current = 0;
    startRef.current = null;
    setElapsed(0);
    setRunning(false);
  }, []);

  useEffect(() => {
    if (running) {
      intervalRef.current = setInterval(() => {
        setElapsed(baseRef.current + (Date.now() - startRef.current) / 1000);
      }, 500);
    } else {
      clearInterval(intervalRef.current);
    }
    return () => clearInterval(intervalRef.current);
  }, [running]);

  return { elapsed, running, start, pause, reset };
}

// ─── Sub-components ───────────────────────────────────────────────────────────
// Shows who else is signed in right now (live, via Supabase Presence).
function OnlineUsers({ users, currentUser }) {
  const [open, setOpen] = useState(false);
  if (!users || users.length === 0) return null;
  return (
    <div style={{ position: "relative" }}>
      <button onClick={() => setOpen(v => !v)}
        title="See who is online"
        style={{
          display: "flex", alignItems: "center", gap: 5,
          background: "#0a2a10", border: "1px solid #6effa044", color: "#6effa0",
          borderRadius: 99, padding: "4px 10px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700,
        }}>
        <span style={{ width: 7, height: 7, borderRadius: "50%", background: "#6effa0", boxShadow: "0 0 6px #6effa0", flexShrink: 0 }} />
        {users.length} online
      </button>
      {open && (
        <>
          <div onClick={() => setOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 900 }} />
          <div style={{
            position: "absolute", top: "calc(100% + 6px)", right: 0, zIndex: 901,
            background: "#141626", border: "1px solid #2a2d4a", borderRadius: 10,
            padding: 8, minWidth: 170, boxShadow: "0 12px 40px #000a",
          }}>
            <div style={{ fontSize: 10, color: "#8890b8", letterSpacing: 1, padding: "2px 8px 6px" }}>ONLINE NOW</div>
            {users.map(u => (
              <div key={u.username} style={{ display: "flex", alignItems: "center", gap: 7, padding: "5px 8px", borderRadius: 6, background: u.username === currentUser ? "#1a4a8a33" : "transparent" }}>
                <span style={{ width: 6, height: 6, borderRadius: "50%", background: "#6effa0", flexShrink: 0 }} />
                <span style={{ fontSize: 13, color: "#eee", fontWeight: u.username === currentUser ? 700 : 400 }}>
                  {u.username}{u.username === currentUser ? " (you)" : ""}
                </span>
                {u.isAdmin && <span style={{ fontSize: 9, color: "#ffd966", border: "1px solid #ffd96644", borderRadius: 4, padding: "1px 4px", marginLeft: "auto" }}>admin</span>}
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

function StatusBadge({ status, small }) {
  const cfg = {
    ok:   { bg: "#1a6b3a", text: "#6effa0", label: "✓ On time" },
    warn: { bg: "#7a5a00", text: "#ffd966", label: "⚠ Less OOP" },
    late: { bg: "#7a1a1a", text: "#ff7070", label: "✗ OOP" },
    idle: { bg: "#2a2a3a", text: "#8888aa", label: "— Waiting" },
  };
  const c = cfg[status] || cfg.idle;
  return (
    <span style={{
      background: c.bg,
      color: c.text,
      padding: small ? "2px 7px" : "2px 10px",
      borderRadius: 99,
      fontSize: small ? 10 : 12,
      fontWeight: 700,
      letterSpacing: small ? 0 : "0.05em",
      border: `1px solid ${c.text}33`,
      whiteSpace: "nowrap",
    }}>{c.label}</span>
  );
}

function HoleRow({ hole, par, scheduled, actual, diff, onClick }) {
  const hasData = actual !== null;
  const status = hasData ? getStatus(diff) : "idle";

  return (
    <tr
      onClick={onClick}
      style={{ cursor: "pointer", background: "transparent", transition: "background 0.15s" }}
      onMouseEnter={e => e.currentTarget.style.background = "#ffffff08"}
      onMouseLeave={e => e.currentTarget.style.background = "transparent"}
    >
      <td style={{ padding: "6px 8px", textAlign: "center", color: "#8888aa", fontWeight: 600 }}>{hole}</td>
      <td style={{ padding: "6px 8px", textAlign: "center", color: "#ccc" }}>Par {par}</td>
      <td style={{ padding: "6px 8px", textAlign: "center", color: "#8899cc" }}>{minToTime(scheduled)}</td>
      <td style={{ padding: "6px 8px", textAlign: "center", color: hasData ? "#eee" : "#555" }}>
        {hasData ? actual : "—"}
      </td>
      <td style={{ padding: "6px 8px", textAlign: "center" }}>
        {hasData ? (
          <span style={{ color: diff >= 3 ? "#ff7070" : diff >= 1 ? "#ffd966" : "#6effa0", fontWeight: 700 }}>
            {diff > 0 ? `+${diff}` : diff} min
          </span>
        ) : "—"}
      </td>
      <td style={{ padding: "6px 8px" }}><StatusBadge status={status} /></td>
    </tr>
  );
}

// ─── Quick Generate Panel ─────────────────────────────────────────────────────
function QuickGeneratePanel({ onGenerate, existingGroups1, existingGroups10, pars, onGenerateShotgun, existingGroupsShotgun, mode: modeProp, onModeChange }) {
  const [modeLocal, setModeLocal] = useState("h1only"); // "h1only" | "h10only" | "both" | "shotgun"
  const mode = modeProp ?? modeLocal;
  const setMode = onModeChange ?? setModeLocal;

  // H1-only mode
  const [countH1Only, setCountH1Only] = useState(12);
  const [startH1Only, setStartH1Only] = useState("06:40");
  const [gapH1Only, setGapH1Only] = useState(10);
  const [useAfternoonH1, setUseAfternoonH1] = useState(false);
  const [afternoonCountH1, setAfternoonCountH1] = useState(12);
  const [afternoonStartH1, setAfternoonStartH1] = useState("11:20");

  // H10-only mode
  const [countH10Only, setCountH10Only] = useState(12);
  const [startH10Only, setStartH10Only] = useState("06:40");
  const [gapH10Only, setGapH10Only] = useState(10);
  const [useAfternoonH10, setUseAfternoonH10] = useState(false);
  const [afternoonCountH10, setAfternoonCountH10] = useState(12);
  const [afternoonStartH10, setAfternoonStartH10] = useState("11:20");

  // H1+H10 mode
  const [countH1, setCountH1] = useState(12);
  const [countH10, setCountH10] = useState(12);
  const [startBoth, setStartBoth] = useState("06:40");
  const [gapBoth, setGapBoth] = useState(10);
  const [useAfternoonBoth, setUseAfternoonBoth] = useState(false);
  const [afternoonCountH1B, setAfternoonCountH1B] = useState(12);
  const [afternoonCountH10B, setAfternoonCountH10B] = useState(12);
  const [afternoonStartBoth, setAfternoonStartBoth] = useState("11:20");

  const [generated, setGenerated] = useState(false);
  // Add to the existing schedule, or rebuild it from scratch with these settings
  const [replaceExisting, setReplaceExisting] = useState(false);

  // ── preview helpers ──────────────────────────────────────────────────────────
  const buildPreview = () => {
    const rows = [];
    if (mode === "h1only") {
      const base = (() => { const [h,m] = startH1Only.split(":").map(Number); return h*60+m; })();
      for (let i = 0; i < countH1Only; i++)
        rows.push({ label: `Group ${i+1}`, hole: "H1", time: minToTime(base + i * gapH1Only), section: "🌅 Morning" });
      if (useAfternoonH1) {
        const abase = (() => { const [h,m] = afternoonStartH1.split(":").map(Number); return h*60+m; })();
        for (let i = 0; i < afternoonCountH1; i++)
          rows.push({ label: `Group ${countH1Only + i + 1}`, hole: "H1", time: minToTime(abase + i * gapH1Only), section: "☀️ Afternoon" });
      }
    } else if (mode === "h10only") {
      const base = (() => { const [h,m] = startH10Only.split(":").map(Number); return h*60+m; })();
      for (let i = 0; i < countH10Only; i++)
        rows.push({ label: `Group ${i+1}`, hole: "H10", time: minToTime(base + i * gapH10Only), section: "🌅 Morning" });
      if (useAfternoonH10) {
        const abase = (() => { const [h,m] = afternoonStartH10.split(":").map(Number); return h*60+m; })();
        for (let i = 0; i < afternoonCountH10; i++)
          rows.push({ label: `Group ${countH10Only + i + 1}`, hole: "H10", time: minToTime(abase + i * gapH10Only), section: "☀️ Afternoon" });
      }
    } else {
      const base = (() => { const [h,m] = startBoth.split(":").map(Number); return h*60+m; })();
      for (let i = 0; i < countH1; i++)
        rows.push({ label: `Group ${i+1}`, hole: "H1", time: minToTime(base + i * gapBoth), section: "🌅 Morning" });
      for (let i = 0; i < countH10; i++)
        rows.push({ label: `Group ${countH1 + i + 1}`, hole: "H10", time: minToTime(base + i * gapBoth), section: "🌅 Morning" });
      if (useAfternoonBoth) {
        const abase = (() => { const [h,m] = afternoonStartBoth.split(":").map(Number); return h*60+m; })();
        const mOff = countH1 + countH10;
        for (let i = 0; i < afternoonCountH1B; i++)
          rows.push({ label: `Group ${mOff + i + 1}`, hole: "H1", time: minToTime(abase + i * gapBoth), section: "☀️ Afternoon" });
        for (let i = 0; i < afternoonCountH10B; i++)
          rows.push({ label: `Group ${mOff + afternoonCountH1B + i + 1}`, hole: "H10", time: minToTime(abase + i * gapBoth), section: "☀️ Afternoon" });
      }
    }
    return rows;
  };

  const preview = buildPreview();
  const totalNew = preview.length;

  const handleGenerate = () => {
    if (totalNew === 0) return;
    const existingAll = [...existingGroups1, ...existingGroups10, ...(existingGroupsShotgun || [])];
    // Appending continues the numbering; replacing starts again from Group 1.
    const maxN = replaceExisting ? 0 : existingAll.reduce((acc, g) => {
      const m = g.name.match(/Group (\d+)/);
      return m ? Math.max(acc, Number(m[1])) : acc;
    }, 0);

    const newG1 = [];
    const newG10 = [];
    let counter = maxN + 1;

    if (mode === "h1only") {
      const base = (() => { const [h,m] = startH1Only.split(":").map(Number); return h*60+m; })();
      for (let i = 0; i < countH1Only; i++) {
        newG1.push({ id: Date.now() + counter, name: `Group ${counter}`, startTime: minToTime(base + i * gapH1Only), color: GROUP_COLORS_H1[(counter-1) % GROUP_COLORS_H1.length], startHole: 1, section: "morning" });
        counter++;
      }
      if (useAfternoonH1) {
        const abase = (() => { const [h,m] = afternoonStartH1.split(":").map(Number); return h*60+m; })();
        for (let i = 0; i < afternoonCountH1; i++) {
          newG1.push({ id: Date.now() + counter, name: `Group ${counter}`, startTime: minToTime(abase + i * gapH1Only), color: GROUP_COLORS_H1[(counter-1) % GROUP_COLORS_H1.length], startHole: 1, section: "afternoon" });
          counter++;
        }
      }
    } else if (mode === "h10only") {
      const base = (() => { const [h,m] = startH10Only.split(":").map(Number); return h*60+m; })();
      for (let i = 0; i < countH10Only; i++) {
        newG10.push({ id: Date.now() + counter, name: `Group ${counter}`, startTime: minToTime(base + i * gapH10Only), color: GROUP_COLORS_H10[(counter-1) % GROUP_COLORS_H10.length], startHole: 10, section: "morning" });
        counter++;
      }
      if (useAfternoonH10) {
        const abase = (() => { const [h,m] = afternoonStartH10.split(":").map(Number); return h*60+m; })();
        for (let i = 0; i < afternoonCountH10; i++) {
          newG10.push({ id: Date.now() + counter, name: `Group ${counter}`, startTime: minToTime(abase + i * gapH10Only), color: GROUP_COLORS_H10[(counter-1) % GROUP_COLORS_H10.length], startHole: 10, section: "afternoon" });
          counter++;
        }
      }
    } else {
      const base = (() => { const [h,m] = startBoth.split(":").map(Number); return h*60+m; })();
      for (let i = 0; i < countH1; i++) {
        newG1.push({ id: Date.now() + counter, name: `Group ${counter}`, startTime: minToTime(base + i * gapBoth), color: GROUP_COLORS_H1[(counter-1) % GROUP_COLORS_H1.length], startHole: 1, section: "morning" });
        counter++;
      }
      for (let i = 0; i < countH10; i++) {
        newG10.push({ id: Date.now() + counter, name: `Group ${counter}`, startTime: minToTime(base + i * gapBoth), color: GROUP_COLORS_H10[(counter-1) % GROUP_COLORS_H10.length], startHole: 10, section: "morning" });
        counter++;
      }
      if (useAfternoonBoth) {
        const abase = (() => { const [h,m] = afternoonStartBoth.split(":").map(Number); return h*60+m; })();
        for (let i = 0; i < afternoonCountH1B; i++) {
          newG1.push({ id: Date.now() + counter, name: `Group ${counter}`, startTime: minToTime(abase + i * gapBoth), color: GROUP_COLORS_H1[(counter-1) % GROUP_COLORS_H1.length], startHole: 1, section: "afternoon" });
          counter++;
        }
        for (let i = 0; i < afternoonCountH10B; i++) {
          newG10.push({ id: Date.now() + counter, name: `Group ${counter}`, startTime: minToTime(abase + i * gapBoth), color: GROUP_COLORS_H10[(counter-1) % GROUP_COLORS_H10.length], startHole: 10, section: "afternoon" });
          counter++;
        }
      }
    }

    // "Replace" rebuilds the whole schedule from these settings, so a mistake in
    // any of them can be corrected in one go instead of editing group by group.
    if (replaceExisting) {
      const existingCount = existingAll.length;
      if (existingCount > 0) {
        const keep = window.confirm(
          `Replace the current schedule?\n\n` +
          `${existingCount} existing group(s) will be replaced by ${totalNew} new one(s).\n\n` +
          `OK = keep recorded times for groups whose number still exists\n` +
          `Cancel = start completely fresh (recorded times are cleared)`
        );
        onGenerate(newG1, newG10, { replace: true, keepData: keep });
      } else {
        onGenerate(newG1, newG10, { replace: true, keepData: false });
      }
    } else {
      onGenerate(newG1, newG10);
    }
    setGenerated(true);
    setTimeout(() => setGenerated(false), 2000);
  };

  const numInput = (val, setVal, min=1, max=30) => (
    <div style={{ display: "flex", alignItems: "center", gap: 3 }}>
      <button onClick={() => setVal(v => Math.max(min, v-1))}
        style={{ width:36, height:36, background:"#1e2135", border:"1px solid #2a2d4a", color:"#aaa", borderRadius:6, cursor:"pointer", fontSize:15, fontFamily:"inherit" }}>−</button>
      <input type="number" min={min} max={max} value={val}
        onChange={e => setVal(Math.min(max, Math.max(min, Number(e.target.value)||min)))}
        style={{ width:46, background:"#1e2135", border:"1px solid #4e9af155", color:"#ffd966", borderRadius:7, padding:"4px 0", fontFamily:"inherit", fontSize:15, fontWeight:700, textAlign:"center" }} />
      <button onClick={() => setVal(v => Math.min(max, v+1))}
        style={{ width:36, height:36, background:"#1e2135", border:"1px solid #2a2d4a", color:"#aaa", borderRadius:6, cursor:"pointer", fontSize:15, fontFamily:"inherit" }}>+</button>
    </div>
  );

  const timeInput = (val, setVal, accentColor="#4e9af1") => (
    <input type="time" value={val} onChange={e => setVal(e.target.value)}
      style={{ background:"#1e2135", border:`1px solid ${accentColor}55`, color:accentColor, borderRadius:7, padding:"5px 10px", fontFamily:"inherit", fontSize:14, fontWeight:700 }} />
  );

  const sectionLabel = (text, color) => (
    <div style={{ fontSize:11, letterSpacing:2, fontWeight:700, color, marginBottom:10 }}>{text}</div>
  );

  const row = (label, children) => (
    <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap", marginBottom:10 }}>
      <span style={{ fontSize:12, color:"#9aa2c7", minWidth:110, whiteSpace:"nowrap" }}>{label}</span>
      {children}
    </div>
  );

  return (
    <div style={{ background:"#141626", border:"1px solid #4e9af133", borderRadius:12, padding:"18px 20px", marginBottom:16 }}>
      {/* Header */}
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:16 }}>
        <div style={{ fontSize:12, color:"#4e9af1", letterSpacing:2, fontWeight:700 }}>⚡ Auto Generate Groups</div>
        {/* Mode toggle */}
        <div style={{ display:"flex", gap:6 }}>
          {[["h1only","🟢 H1 only"],["h10only","🔵 H10 only"],["both","🟢 H1 + 🔵 H10"],["shotgun","🔫 Shotgun 4 Holes"]].map(([m,label]) => (
            <button key={m} onClick={() => setMode(m)} style={{
              padding:"5px 12px", borderRadius:7, cursor:"pointer", fontFamily:"inherit", fontSize:12, fontWeight:700,
              background: mode===m ? "#1a3a6a" : "#0d0f1a",
              border: `1px solid ${mode===m ? "#4e9af1" : "#2a2d4a"}`,
              color: mode===m ? "#4e9af1" : "#555",
              transition:"all 0.15s",
            }}>{label}</button>
          ))}
        </div>
      </div>

      {/* ── H1 only ── */}
      {mode === "h1only" && (
        <div>
          {/* Morning section */}
          <div style={{ background:"#0d0f1a", borderRadius:10, padding:"14px 16px", marginBottom:10 }}>
            {sectionLabel("🌅 Morning Section", "#6effa0")}
            {row("Number of H1 groups", numInput(countH1Only, setCountH1Only))}
            {row("First group start time", timeInput(startH1Only, setStartH1Only, "#6effa0"))}
            {row("Gap between groups",
              <>{numInput(gapH1Only, setGapH1Only, 1, 60)}<span style={{fontSize:12,color:"#8890b8"}}>min</span></>
            )}
          </div>

          {/* Afternoon toggle */}
          <button onClick={() => setUseAfternoonH1(v => !v)} style={{
            width:"100%", padding:"9px 0", borderRadius:8, cursor:"pointer", fontFamily:"inherit",
            fontSize:13, fontWeight:700, marginBottom: useAfternoonH1 ? 10 : 0,
            background: useAfternoonH1 ? "#1a1a00" : "#0d0f1a",
            border: `1px solid ${useAfternoonH1 ? "#ffd966" : "#2a2d4a"}`,
            color: useAfternoonH1 ? "#ffd966" : "#555",
            transition:"all 0.15s",
          }}>
            {useAfternoonH1 ? "☀️ Afternoon Section added ✓ — click to remove" : "+ Add Afternoon Section"}
          </button>

          {useAfternoonH1 && (
            <div style={{ background:"#0d0f1a", borderRadius:10, padding:"14px 16px", border:"1px solid #ffd96633" }}>
              {sectionLabel("☀️ Afternoon Section", "#ffd966")}
              {row("Number of H1 groups", numInput(afternoonCountH1, setAfternoonCountH1))}
              {row("First group start time", timeInput(afternoonStartH1, setAfternoonStartH1, "#ffd966"))}
              <div style={{ fontSize:11, color:"#8890b8", marginTop:2 }}>Gap uses the same value as the morning section ({gapH1Only} min)</div>
            </div>
          )}
        </div>
      )}

      {/* ── H10 only ── */}
      {mode === "h10only" && (
        <div>
          {/* Morning section */}
          <div style={{ background:"#0d0f1a", borderRadius:10, padding:"14px 16px", marginBottom:10 }}>
            {sectionLabel("🌅 Morning Section", "#4e9af1")}
            {row("Number of H10 groups", numInput(countH10Only, setCountH10Only))}
            {row("First group start time", timeInput(startH10Only, setStartH10Only, "#4e9af1"))}
            {row("Gap between groups",
              <>{numInput(gapH10Only, setGapH10Only, 1, 60)}<span style={{fontSize:12,color:"#8890b8"}}>min</span></>
            )}
          </div>

          {/* Afternoon toggle */}
          <button onClick={() => setUseAfternoonH10(v => !v)} style={{
            width:"100%", padding:"9px 0", borderRadius:8, cursor:"pointer", fontFamily:"inherit",
            fontSize:13, fontWeight:700, marginBottom: useAfternoonH10 ? 10 : 0,
            background: useAfternoonH10 ? "#1a1a00" : "#0d0f1a",
            border: `1px solid ${useAfternoonH10 ? "#ffd966" : "#2a2d4a"}`,
            color: useAfternoonH10 ? "#ffd966" : "#555",
            transition:"all 0.15s",
          }}>
            {useAfternoonH10 ? "☀️ Afternoon Section added ✓ — click to remove" : "+ Add Afternoon Section"}
          </button>

          {useAfternoonH10 && (
            <div style={{ background:"#0d0f1a", borderRadius:10, padding:"14px 16px", border:"1px solid #ffd96633" }}>
              {sectionLabel("☀️ Afternoon Section", "#ffd966")}
              {row("Number of H10 groups", numInput(afternoonCountH10, setAfternoonCountH10))}
              {row("First group start time", timeInput(afternoonStartH10, setAfternoonStartH10, "#ffd966"))}
              <div style={{ fontSize:11, color:"#8890b8", marginTop:2 }}>Gap uses the same value as the morning section ({gapH10Only} min)</div>
            </div>
          )}
        </div>
      )}

      {/* ── H1 + H10 ── */}
      {mode === "both" && (
        <div>
          {/* Morning section */}
          <div style={{ background:"#0d0f1a", borderRadius:10, padding:"14px 16px", marginBottom:10 }}>
            {sectionLabel("🌅 Morning Section", "#6effa0")}
            <div style={{ display:"flex", gap:16, flexWrap:"wrap", marginBottom:10 }}>
              <div>
                <div style={{ fontSize:11, color:"#6effa0", marginBottom:6, fontWeight:700 }}>🟢 H1 group</div>
                {numInput(countH1, setCountH1)}
              </div>
              <div>
                <div style={{ fontSize:11, color:"#4e9af1", marginBottom:6, fontWeight:700 }}>🔵 H10 group</div>
                {numInput(countH10, setCountH10)}
              </div>
            </div>
            {row("First group start time", timeInput(startBoth, setStartBoth, "#4e9af1"))}
            {row("Gap between groups",
              <>{numInput(gapBoth, setGapBoth, 1, 60)}<span style={{fontSize:12,color:"#8890b8"}}>min</span></>
            )}
            <div style={{ fontSize:11, color:"#8890b8", marginTop:2 }}>H1 and H10 start at the same time, but the gap is counted separately within each hole</div>
          </div>

          {/* Afternoon toggle */}
          <button onClick={() => setUseAfternoonBoth(v => !v)} style={{
            width:"100%", padding:"9px 0", borderRadius:8, cursor:"pointer", fontFamily:"inherit",
            fontSize:13, fontWeight:700, marginBottom: useAfternoonBoth ? 10 : 0,
            background: useAfternoonBoth ? "#1a1a00" : "#0d0f1a",
            border: `1px solid ${useAfternoonBoth ? "#ffd966" : "#2a2d4a"}`,
            color: useAfternoonBoth ? "#ffd966" : "#555",
            transition:"all 0.15s",
          }}>
            {useAfternoonBoth ? "☀️ Afternoon Section added ✓ — click to remove" : "+ Add Afternoon Section"}
          </button>

          {useAfternoonBoth && (
            <div style={{ background:"#0d0f1a", borderRadius:10, padding:"14px 16px", border:"1px solid #ffd96633" }}>
              {sectionLabel("☀️ Afternoon Section", "#ffd966")}
              <div style={{ display:"flex", gap:16, flexWrap:"wrap", marginBottom:10 }}>
                <div>
                  <div style={{ fontSize:11, color:"#6effa0", marginBottom:6, fontWeight:700 }}>🟢 H1 group</div>
                  {numInput(afternoonCountH1B, setAfternoonCountH1B, 0)}
                </div>
                <div>
                  <div style={{ fontSize:11, color:"#4e9af1", marginBottom:6, fontWeight:700 }}>🔵 H10 group</div>
                  {numInput(afternoonCountH10B, setAfternoonCountH10B, 0)}
                </div>
              </div>
              {row("First group start time", timeInput(afternoonStartBoth, setAfternoonStartBoth, "#ffd966"))}
              <div style={{ fontSize:11, color:"#8890b8", marginTop:2 }}>Gap uses the same value as the morning section ({gapBoth} min)</div>
            </div>
          )}
        </div>
      )}

      {/* ── Shotgun 4 Holes ── */}
      {mode === "shotgun" && (
        <ShotgunStartPanel
          pars={pars}
          onGenerate={onGenerateShotgun}
          existingGroups={[...existingGroups1, ...existingGroups10, ...(existingGroupsShotgun || [])]}
        />
      )}

      {/* Add to what's there, or rebuild the schedule from these settings —
          the latter is how a wrong start time / gap / count gets corrected. */}
      {mode !== "shotgun" && (
        <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
          {[
            { val: false, label: "➕ Add to schedule" },
            { val: true,  label: "↻ Replace schedule" },
          ].map(opt => (
            <button key={String(opt.val)} onClick={() => setReplaceExisting(opt.val)}
              style={{
                flex: 1, padding: "8px 0", borderRadius: 8, cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700,
                background: replaceExisting === opt.val ? (opt.val ? "#2a1a0a" : "#1a4a8a") : "#0d0f1a",
                border: `1px solid ${replaceExisting === opt.val ? (opt.val ? "#ff9966" : "#4e9af1") : "#2a2d4a"}`,
                color: replaceExisting === opt.val ? (opt.val ? "#ff9966" : "#fff") : "#8890b8",
              }}>{opt.label}</button>
          ))}
        </div>
      )}
      {mode !== "shotgun" && replaceExisting && (
        <div style={{ fontSize: 11, color: "#ff9966", marginTop: 6, lineHeight: 1.5 }}>
          Rebuilds the whole schedule and renumbers from Group 1. You'll be asked
          whether to keep times already recorded.
        </div>
      )}

      {mode !== "shotgun" && (
      <button
        onClick={handleGenerate}
        disabled={totalNew === 0}
        style={{
          width:"100%", marginTop:14, padding:"11px 0",
          background: generated ? "#0a2a10" : totalNew === 0 ? "#1a1d2e" : "linear-gradient(135deg, #1a3a6a, #4e9af1)",
          border: `1px solid ${generated ? "#6effa088" : totalNew === 0 ? "#2a2d4a" : "#4e9af188"}`,
          color: generated ? "#6effa0" : totalNew === 0 ? "#444" : "#fff",
          borderRadius:9, cursor: totalNew===0 ? "not-allowed" : "pointer",
          fontFamily:"'Bebas Neue'", fontSize:18, letterSpacing:3,
          transition:"all 0.2s",
        }}
      >
        {generated ? "✓ Groups generated!" : `⚡ ${replaceExisting ? "Replace with" : "Generate"} ${totalNew} group(s)`}
      </button>
      )}
    </div>
  );
}

// ─── Shotgun Start Panel (4 holes: H1 / H5(6) / H10 / H14(15)) ─────────────────
// Used only when necessary (rare): all groups tee off at the same time on different holes
function ShotgunStartPanel({ pars, onGenerate, existingGroups }) {
  const [startTime, setStartTime] = useState("07:00");
  const [groupsPerHole, setGroupsPerHole] = useState(6);
  const [gap, setGap] = useState(10);
  const [useAfternoon, setUseAfternoon] = useState(false);
  const [afternoonGroupsPerHole, setAfternoonGroupsPerHole] = useState(6);
  const [afternoonStartTime, setAfternoonStartTime] = useState("11:20");
  const [generated, setGenerated] = useState(false);

  const resolvedHoles = resolveShotgunStartHoles(pars); // [1, 5|6, 10, 14|15]
  const adjustedNote = (() => {
    const notes = [];
    if (resolvedHoles[1] !== 5) notes.push(`H5 is Par 3 → moved to H${resolvedHoles[1]}`);
    if (resolvedHoles[3] !== 14) notes.push(`H14 is Par 3 → moved to H${resolvedHoles[3]}`);
    const bothPar3Pair1 = (pars?.[4] === 3) && (pars?.[5] === 3);
    const bothPar3Pair2 = (pars?.[13] === 3) && (pars?.[14] === 3);
    if (bothPar3Pair1) notes.push("H5 and H6 are both Par 3 → still starting at H5");
    if (bothPar3Pair2) notes.push("H14 and H15 are both Par 3 → still starting at H14");
    return notes;
  })();

  const totalMorning = resolvedHoles.length * groupsPerHole;
  const totalAfternoon = useAfternoon ? resolvedHoles.length * afternoonGroupsPerHole : 0;
  const totalNew = totalMorning + totalAfternoon;

  const handleGenerate = () => {
    if (totalNew === 0) return;
    const maxN = existingGroups.reduce((acc, g) => {
      const m = g.name.match(/Group (\d+)/);
      return m ? Math.max(acc, Number(m[1])) : acc;
    }, 0);
    let counter = maxN + 1;
    const base = (() => { const [h, m] = startTime.split(":").map(Number); return h * 60 + m; })();
    const newGroups = [];
    const batchId = Date.now(); // links groups generated together in the same batch (allows switching holes for the whole set)
    resolvedHoles.forEach(hole => {
      const meta = getStartHoleMeta(hole);
      const setId = `${batchId}-${hole}`; // sub-set by original start point, e.g. Groups 1-6 starting at H1 share the same setId
      for (let i = 0; i < groupsPerHole; i++) {
        newGroups.push({
          id: Date.now() + counter,
          name: `Group ${counter}`,
          startTime: minToTime(base + i * gap),
          color: meta.color,
          startHole: hole,
          section: "morning",
          isShotgun: true,
          shotgunSetId: setId,
        });
        counter++;
      }
    });

    if (useAfternoon) {
      const abase = (() => { const [h, m] = afternoonStartTime.split(":").map(Number); return h * 60 + m; })();
      const abatchId = Date.now() + 1; // separate batch set for the afternoon round
      resolvedHoles.forEach(hole => {
        const meta = getStartHoleMeta(hole);
        const setId = `${abatchId}-${hole}`;
        for (let i = 0; i < afternoonGroupsPerHole; i++) {
          newGroups.push({
            id: Date.now() + counter,
            name: `Group ${counter}`,
            startTime: minToTime(abase + i * gap),
            color: meta.color,
            startHole: hole,
            section: "afternoon",
            isShotgun: true,
            shotgunSetId: setId,
          });
          counter++;
        }
      });
    }

    onGenerate(newGroups);
    setGenerated(true);
    setTimeout(() => setGenerated(false), 2000);
  };

  const numInput = (val, setVal, min=1, max=Infinity) => (
    <div style={{ display: "flex", alignItems: "center", gap: 3 }}>
      <button onClick={() => setVal(v => Math.max(min, v-1))}
        style={{ width:36, height:36, background:"#1e2135", border:"1px solid #2a2d4a", color:"#aaa", borderRadius:6, cursor:"pointer", fontSize:15, fontFamily:"inherit" }}>−</button>
      <input type="number" min={min} {...(Number.isFinite(max) ? { max } : {})} value={val}
        onChange={e => setVal(Math.min(max, Math.max(min, Number(e.target.value)||min)))}
        style={{ width:46, background:"#1e2135", border:"1px solid #f1734e55", color:"#ffd966", borderRadius:7, padding:"4px 0", fontFamily:"inherit", fontSize:15, fontWeight:700, textAlign:"center" }} />
      <button onClick={() => setVal(v => Math.min(max, v+1))}
        style={{ width:36, height:36, background:"#1e2135", border:"1px solid #2a2d4a", color:"#aaa", borderRadius:6, cursor:"pointer", fontSize:15, fontFamily:"inherit" }}>+</button>
    </div>
  );

  return (
    <div style={{ background:"#1a1410", border:"1px solid #f1734e44", borderRadius:12, padding:"18px 20px", marginBottom:16 }}>
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:6 }}>
        <div style={{ fontSize:12, color:"#f1734e", letterSpacing:2, fontWeight:700 }}>🔫 SHOTGUN START — Tee off together on 4 holes</div>
      </div>
      <div style={{ fontSize:11, color:"#888", marginBottom:14 }}>Use only when necessary (rare) — all groups tee off at the same time from different start points</div>

      {/* Show the computed start points */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4, 1fr)", gap:8, marginBottom:14 }}>
        {resolvedHoles.map(hole => {
          const meta = getStartHoleMeta(hole);
          return (
            <div key={hole} style={{
              background:"#0d0f1a", border:`1px solid ${meta.color}55`,
              borderRadius:8, padding:"8px 10px", textAlign:"center", minWidth:0,
            }}>
              <div style={{ fontSize:16, fontWeight:700, color: meta.color, fontFamily:"'Bebas Neue'", letterSpacing:2 }}>H{hole}</div>
              <div style={{ fontSize:11, color:"#9aa2c7" }}>Par {pars?.[hole-1] ?? "?"}</div>
            </div>
          );
        })}
      </div>

      {adjustedNote.length > 0 && (
        <div style={{ background:"#1a1a00", border:"1px solid #ffd96644", borderRadius:8, padding:"8px 12px", marginBottom:14 }}>
          {adjustedNote.map((n, i) => (
            <div key={i} style={{ fontSize:11, color:"#ffd966" }}>⚠ {n}</div>
          ))}
        </div>
      )}

      <div style={{ background:"#0d0f1a", borderRadius:10, padding:"14px 16px", marginBottom:10 }}>
        <div style={{ fontSize:11, color:"#6effa0", fontWeight:700, letterSpacing:1, marginBottom:10 }}>🌅 Morning Section</div>
        <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap", marginBottom:10 }}>
          <span style={{ fontSize:12, color:"#9aa2c7", minWidth:110 }}>Tee-off time</span>
          <input type="time" value={startTime} onChange={e => setStartTime(e.target.value)}
            style={{ background:"#1e2135", border:"1px solid #f1734e55", color:"#f1734e", borderRadius:7, padding:"5px 10px", fontFamily:"inherit", fontSize:14, fontWeight:700 }} />
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap", marginBottom:10 }}>
          <span style={{ fontSize:12, color:"#9aa2c7", minWidth:110 }}>Groups per hole</span>
          {numInput(groupsPerHole, setGroupsPerHole)}
        </div>
        {groupsPerHole > 1 && (
          <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
            <span style={{ fontSize:12, color:"#9aa2c7", minWidth:110 }}>Gap between groups</span>
            {numInput(gap, setGap, 1, 60)}<span style={{fontSize:12,color:"#8890b8"}}>min (within the same hole)</span>
          </div>
        )}
      </div>

      {/* Afternoon toggle */}
      <button onClick={() => setUseAfternoon(v => !v)} style={{
        width:"100%", padding:"9px 0", borderRadius:8, cursor:"pointer", fontFamily:"inherit",
        fontSize:13, fontWeight:700, marginBottom: useAfternoon ? 10 : 14,
        background: useAfternoon ? "#1a1a00" : "#0d0f1a",
        border: `1px solid ${useAfternoon ? "#ffd966" : "#2a2d4a"}`,
        color: useAfternoon ? "#ffd966" : "#555",
        transition:"all 0.15s",
      }}>
        {useAfternoon ? "☀️ Afternoon Section added ✓ — click to remove" : "+ Add Afternoon Section"}
      </button>

      {useAfternoon && (
        <div style={{ background:"#0d0f1a", borderRadius:10, padding:"14px 16px", marginBottom:14, border:"1px solid #ffd96633" }}>
          <div style={{ fontSize:11, color:"#ffd966", fontWeight:700, letterSpacing:1, marginBottom:10 }}>☀️ Afternoon Section</div>
          <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap", marginBottom:10 }}>
            <span style={{ fontSize:12, color:"#9aa2c7", minWidth:110 }}>Tee-off time</span>
            <input type="time" value={afternoonStartTime} onChange={e => setAfternoonStartTime(e.target.value)}
              style={{ background:"#1e2135", border:"1px solid #ffd96655", color:"#ffd966", borderRadius:7, padding:"5px 10px", fontFamily:"inherit", fontSize:14, fontWeight:700 }} />
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
            <span style={{ fontSize:12, color:"#9aa2c7", minWidth:110 }}>Groups per hole</span>
            {numInput(afternoonGroupsPerHole, setAfternoonGroupsPerHole)}
          </div>
          <div style={{ fontSize:11, color:"#8890b8", marginTop:8 }}>Gap uses the same value as the morning section ({gap} min)</div>
        </div>
      )}

      <button
        onClick={handleGenerate}
        disabled={totalNew === 0}
        style={{
          width:"100%", padding:"11px 0",
          background: generated ? "#0a2a10" : "linear-gradient(135deg, #6a1a1a, #f1734e)",
          border: `1px solid ${generated ? "#6effa088" : "#f1734e88"}`,
          color: generated ? "#6effa0" : "#fff",
          borderRadius:9, cursor:"pointer",
          fontFamily:"'Bebas Neue'", fontSize:18, letterSpacing:3,
          transition:"all 0.2s",
        }}
      >
        {generated
          ? "✓ Shotgun created!"
          : `🔫 Generate Shotgun ${totalNew} group(s) (4 holes)${useAfternoon ? ` — Morning ${totalMorning} / Afternoon ${totalAfternoon}` : ""}`}
      </button>
    </div>
  );
}

// ─── Setup Screen ─────────────────────────────────────────────────────────────
const GROUP_COLORS_H1  = ["#4e9af1","#f16b4e","#a06bf1","#f1c34e","#f14e9a","#4ef1a0","#f18c4e","#9af14e"];
const GROUP_COLORS_H10 = ["#6effa0","#ff7070","#ffd966","#c084fc","#fb923c","#38bdf8","#f472b6","#a3e635"];

function makeNextTime(groups) {
  if (!groups.length) return "06:40";
  const last = groups[groups.length - 1];
  const [lh, lm] = last.startTime.split(":").map(Number);
  const nm = lh * 60 + lm + 10;
  return `${String(Math.floor(nm / 60) % 24).padStart(2,"0")}:${String(nm % 60).padStart(2,"0")}`;
}

const STORAGE_KEY_SETUP = "pace_monitor_setup";

function loadSetup() {
  try {
    const raw = memoryStorage.getItem(STORAGE_KEY_SETUP);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch { return null; }
}

function saveSetup(data) {
  try { memoryStorage.setItem(STORAGE_KEY_SETUP, JSON.stringify(data)); } catch {}
}

// ─── Tournament / Round picker ──────────────────────────────────────────────
// Gate before Setup: pick or create a Tournament, then pick or create a Round
// (Q, 1, 2, 3, 4) within it. Only one round is ever "live" at a time — picking
// a different round than the one currently live archives the current round's
// data first, then starts the new one fresh.
const ROUND_LABELS = ["Q", "1", "2", "3", "4"];

function TournamentRoundScreen({ currentUser, isAdmin, onLogout, onChangePassword, onManageUsers, onRoundSelected, liveTournamentId, liveRoundId, hasLiveGroups, allUsers }) {
  const [tournaments, setTournaments] = useState([]);
  const [selectedTournamentId, setSelectedTournamentId] = useState(liveTournamentId || null);
  const [rounds, setRounds] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showCreate, setShowCreate] = useState(false);
  const [editingTournamentId, setEditingTournamentId] = useState(null); // non-null when the create form is editing an existing tournament
  const [newName, setNewName] = useState("");
  const [newVenue, setNewVenue] = useState("");
  const [newFormat, setNewFormat] = useState("stroke");
  const [newHasQualifying, setNewHasQualifying] = useState(true);
  const [newNumRounds, setNewNumRounds] = useState(4);
  const [busy, setBusy] = useState(false);
  const [viewingRound, setViewingRound] = useState(null); // full archived round record being inspected
  const [rolesMap, setRolesMap] = useState({});           // { tournamentId: { TD: "NP", R1: "CS", ... } }
  const [roundPickerFor, setRoundPickerFor] = useState(null); // tournament whose rounds are being picked

  const reloadTournaments = async () => {
    const [t, roles] = await Promise.all([fetchTournaments(), fetchAllRoles()]);
    setRolesMap(roles);
    const visible = t.filter(x => canUserSeeTournament(x, roles, currentUser, isAdmin));
    setTournaments(visible);
    return { visible, roles };
  };

  useEffect(() => {
    (async () => {
      try {
        const { visible } = await reloadTournaments();
        if (!selectedTournamentId && visible.length) setSelectedTournamentId(visible[0].id);
        if (!visible.length && isAdmin) setShowCreate(true);
      } finally {
        setLoading(false);   // show the screen even if the list couldn't load
      }
    })();
  }, []);

  const openRoundPicker = (t) => {
    setSelectedTournamentId(t.id);
    setShowCreate(false);
    setRoundPickerFor(t);
  };

  const reloadTournaments = async () => {
    const [t, roles] = await Promise.all([fetchTournaments(), fetchAllRoles()]);
    setRolesMap(roles);
    // Referees only see the tournament they hold a position in
    const visible = t.filter(x => canUserSeeTournament(x, roles, currentUser, isAdmin));
    setTournaments(visible);
    return visible;
  };

  useEffect(() => {
    (async () => {
      const visible = await reloadTournaments();
      if (!selectedTournamentId && visible.length) setSelectedTournamentId(visible[0].id);
      if (!visible.length && isAdmin) setShowCreate(true);
      setLoading(false);
    })();
  }, []);

  useEffect(() => {
    if (!selectedTournamentId) { setRounds([]); return; }
    (async () => setRounds(await fetchRounds(selectedTournamentId)))();
  }, [selectedTournamentId]);

  const selectedTournament = tournaments.find(t => t.id === selectedTournamentId);
  // Which round slots to offer. Prefer the tournament's saved configuration, but
  // tournaments created before that setting existed have it blank — for those,
  // infer the list from the rounds that actually exist instead of assuming Q+4.
  const availableRoundLabels = (() => {
    if (!selectedTournament) return ROUND_LABELS;
    const existing = rounds.map(r => r.label);
    const configured = selectedTournament.num_rounds != null || selectedTournament.has_qualifying != null;

    if (configured) {
      const list = [
        ...(selectedTournament.has_qualifying !== false ? ["Q"] : []),
        ...ROUND_LABELS.filter(l => l !== "Q").slice(0, selectedTournament.num_rounds ?? 4),
      ];
      // Never hide a round that already holds data
      existing.forEach(l => { if (!list.includes(l)) list.push(l); });
      return ROUND_LABELS.filter(l => list.includes(l));
    }

    // Unconfigured: show what exists (plus the next round so it can be started)
    if (existing.length === 0) return ROUND_LABELS;
    const numbered = existing.filter(l => l !== "Q").map(Number).filter(n => !isNaN(n));
    const nextNum = numbered.length ? Math.max(...numbered) + 1 : 1;
    const list = [...existing];
    if (nextNum <= 4 && !list.includes(String(nextNum))) list.push(String(nextNum));
    return ROUND_LABELS.filter(l => list.includes(l));
  })();

  const resetForm = () => {
    setNewName(""); setNewVenue(""); setNewFormat("stroke"); setNewHasQualifying(true); setNewNumRounds(4);
    setEditingTournamentId(null);
  };

  const handleStartEdit = (t) => {
    setNewName(t.name || "");
    setNewVenue(t.host_venue || "");
    setNewFormat(t.format || "stroke");
    setNewHasQualifying(t.has_qualifying !== false);
    setNewNumRounds(t.num_rounds ?? 4);
    setEditingTournamentId(t.id);
    setShowCreate(true);
  };

  const handleSaveTournament = async () => {
    if (!newName.trim() || busy) return;
    setBusy(true);
    if (editingTournamentId) {
      await updateTournament(editingTournamentId, { name: newName.trim(), hostVenue: newVenue.trim(), format: newFormat, hasQualifying: newHasQualifying, numRounds: newNumRounds });
      setTournaments(prev => prev.map(t => t.id === editingTournamentId
        ? { ...t, name: newName.trim(), host_venue: newVenue.trim(), format: newFormat, has_qualifying: newHasQualifying, num_rounds: newNumRounds }
        : t));
      setBusy(false);
      setShowCreate(false);
      resetForm();
    } else {
      const t = await createTournament({ name: newName.trim(), hostVenue: newVenue.trim(), format: newFormat, hasQualifying: newHasQualifying, numRounds: newNumRounds });
      setBusy(false);
      if (t) {
        setTournaments(prev => [t, ...prev]);
        setSelectedTournamentId(t.id);
        setShowCreate(false);
        resetForm();
      }
    }
  };

  const handleDeleteTournament = async (t) => {
    if (busy) return;
    const ok = window.confirm(`ลบ Tournament "${t.name}" ทิ้งถาวร?\n\nRound ทั้งหมดในทัวร์นาเมนต์นี้ (รวมข้อมูลที่เก็บสำรองไว้) จะถูกลบไปด้วย กู้คืนไม่ได้\n\nดำเนินการต่อหรือไม่?`);
    if (!ok) return;
    setBusy(true);
    await deleteTournament(t.id);
    setBusy(false);
    setTournaments(prev => prev.filter(x => x.id !== t.id));
    if (selectedTournamentId === t.id) setSelectedTournamentId(null);
  };

  // Starting the tournament opens it to the assigned referees.
  const handleToggleStarted = async (t) => {
    if (busy) return;
    const starting = t.run_state !== "started";
    if (starting) {
      const roles = rolesMap[t.id] || {};
      if (Object.keys(roles).length === 0) {
        window.alert(`ยังไม่ได้กำหนดตำแหน่งให้ใครใน "${t.name}"\n\nกดปุ่ม 👥 เพื่อกำหนด TD / CR / R1-R6 ก่อนเริ่มการแข่งขัน`);
        return;
      }
      const staff = Object.entries(roles).map(([p, u]) => `• ${p} — ${u}`).join("\n");
      if (!window.confirm(`เริ่มการแข่งขัน "${t.name}"?\n\n${staff}\n\nกรรมการเหล่านี้จะเข้าทัวร์นี้ได้ทันทีเมื่อ login`)) return;
    } else {
      if (!window.confirm(`หยุดการแข่งขัน "${t.name}" ชั่วคราว?\n\nกรรมการจะยังไม่ถูกออกจากระบบ แต่คนที่ login ใหม่จะยังเข้าไม่ได้จนกว่าจะเริ่มอีกครั้ง`)) return;
    }
    setBusy(true);
    await setTournamentRunState(t.id, starting ? "started" : "draft");
    setTournaments(prev => prev.map(x => x.id === t.id ? { ...x, run_state: starting ? "started" : "draft" } : x));
    setBusy(false);
  };

  const handleToggleClosed = async (t) => {
    if (busy) return;
    const closing = t.status !== "closed";
    const msg = closing
      ? `ปิดการแข่งขัน "${t.name}"?\n\n• กรรมการทุกคนจะถูกออกจากระบบทันที\n• ตำแหน่ง TD / CR / R1-R6 ทั้งหมดจะถูกล้าง\n• เปิดใหม่ต้องกำหนดตำแหน่งใหม่\n\n(admin ไม่ถูกออกจากระบบ และยังเข้าดู/แก้ไขได้)`
      : `เปิดการแข่งขัน "${t.name}" อีกครั้ง?\n\nอย่าลืมกำหนดตำแหน่ง TD / CR / R1-R6 ใหม่ก่อนให้กรรมการเข้าใช้งาน`;
    if (!window.confirm(msg)) return;
    setBusy(true);
    await setTournamentStatus(t.id, closing ? "closed" : "open");
    // Closing wipes the roster — every event is staffed fresh.
    if (closing) {
      await clearTournamentRoles(t.id);
      setRolesMap(prev => ({ ...prev, [t.id]: {} }));
    }
    setBusy(false);
    setTournaments(prev => prev.map(x => x.id === t.id ? { ...x, status: closing ? "closed" : "open" } : x));
  };

  const handlePickRound = async (label) => {
    if (busy) return;
    // A closed competition is read-only for everyone except admins
    if (selectedTournament?.status === "closed" && !isAdmin) {
      window.alert(`Competition "${selectedTournament.name}" is closed.\n\nPlease choose another one.`);
      return;
    }
    // TD and CR set their event up before it starts, so they may enter early.
    const canPrepare = isAdmin;
    if (selectedTournament?.run_state !== "started" && !canPrepare) {
      window.alert(`Competition "${selectedTournament.name}" has not started.\n\nPlease wait for the TD / CR to start it.`);
      return;
    }
    setRoundPickerFor(null);
    const existing = rounds.find(r => r.label === label);

    // Already the live round — just continue straight in, no data changes.
    if (existing && existing.id === liveRoundId) {
      onRoundSelected(selectedTournament, existing, "resume");
      return;
    }
    if (existing && existing.status === "finished") {
      setBusy(true);
      const full = await fetchRoundArchive(existing.id);
      setBusy(false);
      setViewingRound(full || existing);
      return;
    }
    setBusy(true);
    let round = existing;
    if (!round) {
      round = await createRound({ tournamentId: selectedTournament.id, label, isQualifying: label === "Q" });
    }
    setBusy(false);
    if (round) onRoundSelected(selectedTournament, round, "fresh");
  };

  const handleReopenRound = async () => {
    if (!viewingRound || busy) return;
    setBusy(true);
    await onRoundSelected(selectedTournament, viewingRound, "reopen");
    setBusy(false);
  };

  const handleDeleteRound = async () => {
    if (!viewingRound || busy) return;
    const ok = window.confirm(`Delete round "${viewingRound.label}" permanently?\n\nAll archived data for this round will be lost. This cannot be undone.\n\nContinue?`);
    if (!ok) return;
    setBusy(true);
    await deleteRound(viewingRound.id);
    setRounds(prev => prev.filter(r => r.id !== viewingRound.id));
    setViewingRound(null);
    setBusy(false);
  };

  if (loading) {
    return (
      <div style={{ background: "#0d0f1a", minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", color: "#8890b8", fontFamily: "'IBM Plex Mono', monospace" }}>
        Loading…
      </div>
    );
  }

  return (
    <div style={{ background: "#0d0f1a", minHeight: "100vh", fontFamily: "'IBM Plex Mono', monospace", color: "#eee" }}>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&family=Bebas+Neue&display=swap" rel="stylesheet" />
      {/* Grid so the right-hand column can never be pushed off screen */}
      <div style={{ borderBottom: "1px solid #2a2d4a" }}>
        {/* Title + who's signed in */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr auto", alignItems: "start", padding: "16px 20px", gap: 10 }}>
          <div style={{ minWidth: 0 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
              <div style={{ fontFamily: "'Bebas Neue'", fontSize: 22, letterSpacing: 3, color: "#4e9af1", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>🏆 TOURNAMENT</div>
              <span style={{ fontSize: 9, fontWeight: 700, letterSpacing: 1, color: "#ffd966", background: "#2a1a0066", border: "1px solid #ffd96655", borderRadius: 4, padding: "0 6px", height: 18, display: "inline-flex", alignItems: "center", lineHeight: 1, flexShrink: 0 }}>BETA</span>
            </div>
            <div style={{ fontSize: 11, color: "#8890b8", marginTop: 2, lineHeight: 1.4 }}>Golf Referee · Pace of Play System</div>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 8, alignItems: "flex-end", flexShrink: 0, width: 170 }}>
            {currentUser && (
              <span style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 5 }}>
                <span style={{ fontSize: 12, color: "#8890b8" }}>👤</span>
                <span style={{ fontSize: 13, color: "#8899cc", fontWeight: 700, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{currentUser}</span>
              </span>
            )}
            <div style={{ display: "flex", gap: 8, width: "100%" }}>
              <button onClick={onChangePassword}
                title="Change your password"
                style={{ background: "#0d0f1a", border: "1px solid #4e9af144", color: "#4e9af1", borderRadius: 10, height: 42, width: 46, cursor: "pointer", fontFamily: "inherit", fontSize: 15, flexShrink: 0 }}
              >🔑</button>
              <button onClick={onLogout}
                style={{ flex: 1, background: "#1a0d0d", border: "1px solid #ff707044", color: "#ff7070", borderRadius: 10, height: 42, cursor: "pointer", fontFamily: "inherit", fontSize: 14, fontWeight: 700, letterSpacing: 0.5, whiteSpace: "nowrap" }}
                onMouseEnter={e => e.currentTarget.style.borderColor = "#ff707088"}
                onMouseLeave={e => e.currentTarget.style.borderColor = "#ff707044"}
              >⏏ Log out</button>
            </div>
          </div>
        </div>

        {/* Admin actions sit in their own band below the divider */}
        {isAdmin && onManageUsers && (
          <div style={{ borderTop: "1px solid #2a2d4a", padding: "12px 20px" }}>
            <button
              onClick={onManageUsers}
              style={{ width: "100%", minHeight: 56, background: "#1a1a0a", border: "1px solid #ffd96644", color: "#ffd966", borderRadius: 10, cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700, letterSpacing: 1, padding: "10px 8px", lineHeight: 1.4 }}
            >🔑<br />Manage Users</button>
          </div>
        )}
      </div>

      <div style={{ maxWidth: 640, margin: "0 auto", padding: "24px 20px" }}>
        {/* Same selector as the dashboard, so the flow is identical everywhere */}
        <div style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 12, padding: 16, marginBottom: 20 }}>
          <div style={{ fontSize: 13, color: "#4e9af1", letterSpacing: 1, fontWeight: 700, marginBottom: 12 }}>Select competition</div>
          <RoundSelectorBar
            tournamentId={selectedTournamentId}
            roundLabel={null}
            isAdmin={isAdmin}
            onOpen={async (t, r) => {
              if (t.status === "closed" && !isAdmin) {
                window.alert(`Competition "${t.name}" is closed.\n\nPlease choose another one.`);
                return;
              }
              if (!isAdmin && t.run_state !== "started") {
                window.alert(`Competition "${t.name}" has not started.\n\nPlease wait for an admin to start it.`);
                return;
              }
              onRoundSelected(t, r, r.status === "finished" ? "reopen" : "resume");
            }}
          />
        </div>

        {/* Tournament picker */}
        <div style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 12, padding: 20, marginBottom: 20 }}>
          <div style={{ fontSize: 13, color: "#4e9af1", letterSpacing: 1, fontWeight: 700, marginBottom: 14 }}>
            {isAdmin ? "Manage tournaments" : "Your competition"}
          </div>
          {!isAdmin && tournaments.length === 1 && tournaments[0].run_state !== "started" && tournaments[0].status !== "closed" && (
            <div style={{ background: "#1a1a0a", border: "1px solid #ffd96644", borderRadius: 8, padding: "10px 12px", marginBottom: 12, fontSize: 12, color: "#ffd966", lineHeight: 1.5 }}>
              ⏳ Waiting for the TD / CR to start this competition.<br />
              You will be taken in automatically.
            </div>
          )}

          {tournaments.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: showCreate ? 16 : 0 }}>
              {tournaments.map(t => (
                <div key={t.id} style={{
                  padding: "12px 14px", borderRadius: 8, fontFamily: "inherit",
                  background: selectedTournamentId === t.id && !showCreate ? "#1a4a8a" : "#0d0f1a",
                  border: `1px solid ${selectedTournamentId === t.id && !showCreate ? "#4e9af1" : "#2a2d4a"}`,
                }}>
                  {/* Row 1 — name + venue, full width so long names don't wrap awkwardly */}
                  <button onClick={() => openRoundPicker(t)}
                    style={{ width: "100%", textAlign: "left", background: "none", border: "none", cursor: "pointer", padding: 0, fontFamily: "inherit" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, rowGap: 5, flexWrap: "wrap" }}>
                      <span style={{ fontSize: 15, fontWeight: 700, color: "#eee" }}>{t.name || "(untitled tournament)"}</span>
                      {t.status === "closed" ? (
                        <span style={{ fontSize: 9, fontWeight: 700, letterSpacing: 1, color: "#ff7070", background: "#2a0a0a", border: "1px solid #ff707055", borderRadius: 4, padding: "0 6px", height: 18, display: "inline-flex", alignItems: "center", lineHeight: 1 }}>CLOSED</span>
                      ) : t.run_state === "started" ? (
                        <span style={{ fontSize: 9, fontWeight: 700, letterSpacing: 1, color: "#6effa0", background: "#0a2a10", border: "1px solid #6effa055", borderRadius: 4, padding: "0 6px", height: 18, display: "inline-flex", alignItems: "center", lineHeight: 1 }}>● STARTED</span>
                      ) : (
                        <span style={{ fontSize: 9, fontWeight: 700, letterSpacing: 1, color: "#8890b8", background: "#0d0f1a", border: "1px solid #8890b844", borderRadius: 4, padding: "0 6px", height: 18, display: "inline-flex", alignItems: "center", lineHeight: 1 }}>DRAFT</span>
                      )}
                    </div>
                    {t.host_venue && <div style={{ fontSize: 12, color: "#8890b8", marginTop: 2 }}>{t.host_venue}</div>}
                  </button>

                  {/* Row 2 — Select on the left, admin tools on the right, all same height */}
                  <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 10 }}>
                    <button onClick={() => openRoundPicker(t)}
                      style={{ background: "#1a4a8a", border: "1px solid #4e9af1", color: "#fff", borderRadius: 7, height: 32, padding: "0 18px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700 }}>
                      Select
                    </button>
                    {isAdmin && (
                      <div style={{ display: "flex", gap: 6, marginLeft: "auto" }}>
                        <button onClick={() => handleToggleStarted(t)}
                          title={t.run_state === "started" ? "Pause competition" : "Start competition"}
                          style={{ background: "#0d0f1a", border: `1px solid ${t.run_state === "started" ? "#6effa066" : "#8890b844"}`, color: t.run_state === "started" ? "#6effa0" : "#8890b8", borderRadius: 7, height: 32, width: 36, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>
                          {t.run_state === "started" ? "⏸" : "▶"}
                        </button>
                        <button onClick={() => handleToggleClosed(t)}
                          title={t.status === "closed" ? "Reopen competition" : "Close competition"}
                          style={{ background: "#0d0f1a", border: `1px solid ${t.status === "closed" ? "#6effa044" : "#ffd96644"}`, color: t.status === "closed" ? "#6effa0" : "#ffd966", borderRadius: 7, height: 32, width: 36, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>
                          {t.status === "closed" ? "🔓" : "🔒"}
                        </button>
                        <button onClick={() => handleStartEdit(t)}
                          title="Edit competition details"
                          style={{ background: "#0d0f1a", border: "1px solid #4e9af144", color: "#4e9af1", borderRadius: 7, height: 32, width: 36, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>
                          ✏️
                        </button>
                        <button onClick={() => handleDeleteTournament(t)}
                          title="Delete competition"
                          style={{ background: "#0d0f1a", border: "1px solid #ff707044", color: "#ff7070", borderRadius: 7, height: 32, width: 36, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>
                          🗑
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>
          )}

          {isAdmin && (
            showCreate ? (
              <div style={{ display: "flex", flexDirection: "column", gap: 10, paddingTop: tournaments.length ? 16 : 0, borderTop: tournaments.length ? "1px solid #2a2d4a" : "none" }}>
                <input value={newName} onChange={e => setNewName(e.target.value)} placeholder="Tournament name"
                  style={{ background: "#0d0f1a", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 8, padding: "10px 12px", fontFamily: "inherit", fontSize: 14, outline: "none" }} />
                <input value={newVenue} onChange={e => setNewVenue(e.target.value)} placeholder="Host venue"
                  style={{ background: "#0d0f1a", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 8, padding: "10px 12px", fontFamily: "inherit", fontSize: 14, outline: "none" }} />
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={() => setNewFormat("stroke")}
                    style={{ flex: 1, padding: "9px 0", borderRadius: 8, cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700,
                      background: newFormat === "stroke" ? "#1a4a8a" : "#0d0f1a", border: `1px solid ${newFormat === "stroke" ? "#4e9af1" : "#2a2d4a"}`, color: newFormat === "stroke" ? "#fff" : "#8890b8" }}>
                    Stroke Play
                  </button>
                  <button disabled title="Match Play — coming soon"
                    style={{ flex: 1, padding: "9px 0", borderRadius: 8, cursor: "not-allowed", fontFamily: "inherit", fontSize: 13, fontWeight: 700, background: "#0d0f1a", border: "1px solid #2a2d4a", color: "#444" }}>
                    Match Play 🔒
                  </button>
                </div>

                <div>
                  <label style={{ fontSize: 11, color: "#8890b8", letterSpacing: 1 }}>Rounds in this tournament</label>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 6, flexWrap: "wrap" }}>
                    <button onClick={() => setNewHasQualifying(q => !q)}
                      style={{
                        padding: "8px 14px", borderRadius: 8, cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700,
                        background: newHasQualifying ? "#7a5a00" : "#0d0f1a",
                        border: `1px solid ${newHasQualifying ? "#ffd966" : "#2a2d4a"}`,
                        color: newHasQualifying ? "#ffd966" : "#8890b8",
                      }}>Q{newHasQualifying ? " ✓" : ""}</button>
                    <span style={{ fontSize: 12, color: "#666" }}>+</span>
                    {[1, 2, 3, 4].map(n => (
                      <button key={n} onClick={() => setNewNumRounds(n)}
                        style={{
                          width: 34, height: 34, borderRadius: 7, cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 15,
                          background: newNumRounds === n ? "#1a4a8a" : "#0d0f1a",
                          border: `1px solid ${newNumRounds === n ? "#4e9af1" : "#2a2d4a"}`,
                          color: newNumRounds === n ? "#fff" : "#8890b8",
                        }}>{n}</button>
                    ))}
                    <span style={{ fontSize: 11, color: "#8890b8" }}>rounds</span>
                  </div>
                </div>

                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={handleSaveTournament} disabled={!newName.trim() || busy}
                    style={{ flex: 1, padding: "10px 0", borderRadius: 8, cursor: newName.trim() ? "pointer" : "not-allowed", fontFamily: "inherit", fontSize: 14, fontWeight: 700, background: "#1a4a2a", border: "1px solid #6effa066", color: "#6effa0" }}>
                    {editingTournamentId ? "✓ Save changes" : "✓ Create tournament"}
                  </button>
                  {(tournaments.length > 0 || editingTournamentId) && (
                    <button onClick={() => { setShowCreate(false); resetForm(); }}
                      style={{ padding: "10px 16px", borderRadius: 8, cursor: "pointer", fontFamily: "inherit", fontSize: 14, background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7" }}>
                      Cancel
                    </button>
                  )}
                </div>
              </div>
            ) : (
              <button onClick={() => { resetForm(); setShowCreate(true); }}
                style={{ marginTop: tournaments.length ? 12 : 0, width: "100%", padding: "10px 0", borderRadius: 8, cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700, background: "#0d0f1a", border: "1px dashed #4e9af166", color: "#4e9af1" }}>
                + New tournament
              </button>
            )
          )}
        </div>

      </div>

      {/* Round picker — opened by the Select button on a tournament row */}
      {roundPickerFor && (
        <div onClick={() => setRoundPickerFor(null)} style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }}>
          <div onClick={e => e.stopPropagation()} style={{ background: "#141626", border: "1px solid #4e9af166", borderRadius: 14, padding: 22, width: "100%", maxWidth: 380, boxShadow: "0 20px 60px #000" }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 2, color: "#4e9af1", marginBottom: 2 }}>Select round</div>
            <div style={{ fontSize: 12, color: "#8890b8", marginBottom: 16 }}>
              {roundPickerFor.name}{roundPickerFor.host_venue ? ` · ${roundPickerFor.host_venue}` : ""}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10, marginBottom: 16 }}>
              {availableRoundLabels.map(label => {
                const r = rounds.find(rr => rr.label === label);
                const isLive = r && r.id === liveRoundId;
                // Every round keeps its own data now, so the old "finished" flag no
                // longer locks anything — it just means the round has been used.
                const hasData = !!r && (r.status === "finished" || r.status === "live");
                return (
                  <button key={label} onClick={() => handlePickRound(label)} disabled={busy}
                    style={{
                      padding: "16px 0", borderRadius: 10, cursor: busy ? "wait" : "pointer", fontFamily: "inherit",
                      background: isLive ? "#1a4a2a" : hasData ? "#141a2a" : "#0d0f1a",
                      border: `1px solid ${isLive ? "#6effa0" : hasData ? "#4e9af155" : "#2a2d4a"}`,
                      display: "flex", flexDirection: "column", alignItems: "center", gap: 4,
                    }}>
                    <div style={{ fontFamily: "'Bebas Neue'", fontSize: 26, letterSpacing: 2, color: isLive ? "#6effa0" : "#eee" }}>
                      {label === "Q" ? "Q" : `R${label}`}
                    </div>
                    <div style={{ fontSize: 9, letterSpacing: 1, color: isLive ? "#6effa0" : hasData ? "#4e9af1" : "#8890b8" }}>
                      {isLive ? "● LIVE" : hasData ? "HAS DATA" : r ? "SET UP" : "NOT STARTED"}
                    </div>
                  </button>
                );
              })}
            </div>
            <button onClick={() => setRoundPickerFor(null)}
              style={{ width: "100%", background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "inherit", fontSize: 13 }}>
              Cancel
            </button>
          </div>
        </div>
      )}


      {viewingRound && (() => {
        const snapshot = viewingRound.archived_app_state || {};
        return (
          <div style={{ position: "fixed", inset: 0, zIndex: 1200, background: "#0d0f1a", overflowY: "auto" }}>
            <div style={{ position: "sticky", top: 0, zIndex: 10, background: "#141626ee", backdropFilter: "blur(6px)", borderBottom: "1px solid #2a2d4a", padding: "10px 16px", display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <span style={{ fontFamily: "'Bebas Neue'", fontSize: 18, letterSpacing: 2, color: "#eee" }}>
                {viewingRound.label === "Q" ? "Round Q" : `Round ${viewingRound.label}`} <span style={{ fontSize: 12, color: "#666", fontFamily: "inherit", letterSpacing: 0 }}>· FINISHED{viewingRound.finished_at ? ` ${new Date(viewingRound.finished_at).toLocaleDateString("th-TH")}` : ""}</span>
              </span>
              {isAdmin && (
                <div style={{ display: "flex", gap: 8, marginLeft: "auto" }}>
                  <button onClick={handleReopenRound} disabled={busy}
                    style={{ padding: "6px 12px", borderRadius: 7, cursor: busy ? "wait" : "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700, background: "#1a4a8a", border: "1px solid #4e9af1", color: "#fff" }}>
                    ↩️ Reopen
                  </button>
                  <button onClick={handleDeleteRound} disabled={busy}
                    style={{ padding: "6px 12px", borderRadius: 7, cursor: busy ? "wait" : "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700, background: "#2a0a0a", border: "1px solid #ff7070", color: "#ff7070" }}>
                    🗑 Delete
                  </button>
                </div>
              )}
            </div>
            <SummaryScreen
              groups={snapshot.groups || []}
              groupData={viewingRound.archived_group_data || {}}
              pars={snapshot.pars || []}
              parTimes={snapshot.parTimes || []}
              playersPerGroup={3}
              suspensions={snapshot.suspensions || []}
              isSuspended={false}
              pendingStopTime=""
              totalOffsetMin={0}
              onBack={() => setViewingRound(null)}
              currentUser={currentUser}
              onLogout={onLogout}
            />
          </div>
        );
      })()}
    </div>
  );
}

function SetupScreen({ onStart, currentUser, isAdmin, isTrueAdmin, onChangePassword, myPosition, onManageUsers, onLogout, onClearSession, hasLiveSession, onGoToDashboard, tournamentName, hostVenue, roundLabel, savedPars, savedParTimes, savedTurnTime, savedTurnTimeBack, livePars, liveParTimes, liveTurnTime, liveGroups, onApplyLiveEdits, onSwitchTournament, onPickTournament, tournamentId, onPickRound, tournamentRunState, onToggleStarted }) {
  const [groups1, setGroups1] = useState(() => loadSetup()?.groups1 ?? []);
  const [groups10, setGroups10] = useState(() => loadSetup()?.groups10 ?? []);
  const [groupsShotgun, setGroupsShotgun] = useState(() => loadSetup()?.groupsShotgun ?? []);
  const [pars, setPars] = useState(() => savedPars ?? loadSetup()?.pars ?? [...DEFAULT_PARS]);
  const [parTimes, setParTimes] = useState(() => savedParTimes ?? loadSetup()?.parTimes ?? DEFAULT_PARS.map(p => PAR_TIMES[p]));
  const [playersPerGroup, setPlayersPerGroup] = useState(() => loadSetup()?.playersPerGroup ?? 3);
  const [turnTime, setTurnTime] = useState(() => savedTurnTime ?? loadSetup()?.turnTime ?? 1);
  const [turnTimeBack, setTurnTimeBack] = useState(() => savedTurnTimeBack ?? loadSetup()?.turnTimeBack ?? 1);

  // Quick tournament switcher popup — a simple list, so the common "jump to my
  // other event" case never leaves the Setup page. Full management (create,
  // roles, access, delete) still lives on the Tournament screen.
  const [showTournamentPicker, setShowTournamentPicker] = useState(false);
  const [tPickerList, setTPickerList] = useState([]);
  const [tPickerLoading, setTPickerLoading] = useState(false);
  const openTournamentPicker = async () => {
    setShowTournamentPicker(true);
    setTPickerLoading(true);
    const [all, roles] = await Promise.all([fetchTournaments(), fetchAllRoles()]);
    // Only offer what this person may actually open
    setTPickerList(
      all.filter(t =>
        t.status !== "closed" &&
        (isTrueAdmin || t.run_state === "started")
      )
    );
    setTPickerLoading(false);
  };

  // Round switcher popup — loads this tournament's rounds on demand
  const [showRoundPicker, setShowRoundPicker] = useState(false);
  const [pickerRounds, setPickerRounds] = useState([]);
  const [pickerLoading, setPickerLoading] = useState(false);
  const openRoundPicker = async () => {
    if (!tournamentId) return;
    setShowRoundPicker(true);
    setPickerLoading(true);
    setPickerRounds(await fetchRounds(tournamentId));
    setPickerLoading(false);
  };
  // Offer the rounds this tournament actually uses, plus the next one so a new
  // round can still be started — not a blanket Q + R1..R4.
  const pickerLabels = (() => {
    const existing = pickerRounds.map(r => r.label);
    if (existing.length === 0) return ROUND_LABELS;
    const numbered = existing.filter(l => l !== "Q").map(Number).filter(n => !isNaN(n));
    const nextNum = numbered.length ? Math.max(...numbered) + 1 : 1;
    const list = [...existing];
    if (nextNum <= 4 && !list.includes(String(nextNum))) list.push(String(nextNum));
    return ROUND_LABELS.filter(l => list.includes(l));
  })();

  // The tournament record may arrive after this screen first mounts (async fetch), and
  // it changes when switching tournaments — pull its saved course setup in whenever
  // it appears so every round of a tournament shares the same pars / par times / turn time.
  // The groups actually running in the live round live in Supabase, not in this
  // device's local draft. Pull them into the editable lists so this screen always
  // shows the real group list (and so pressing the start/update button can't wipe
  // a live session just because this device's local draft happens to be empty).
  const liveGroupsKey = (liveGroups || []).map(g => `${g.id}:${g.name}:${g.startTime}:${g.startHole}:${g.section || ""}`).join("|");
  useEffect(() => {
    if (!liveGroups || liveGroups.length === 0) return;
    setGroups1(liveGroups.filter(g => (g.startHole || 1) === 1));
    setGroups10(liveGroups.filter(g => (g.startHole || 1) === 10));
    setGroupsShotgun(liveGroups.filter(g => (g.startHole || 1) !== 1 && (g.startHole || 1) !== 10));
  }, [liveGroupsKey]);

  // Whatever the live round is actually running takes priority over the tournament
  // default — otherwise this screen could show stale values (e.g. opened on another
  // device) and "Start tracking" would silently overwrite the real schedule.
  useEffect(() => {
    const p = (livePars?.length === 18 ? livePars : null) ?? (savedPars?.length === 18 ? savedPars : null);
    const pt = (liveParTimes?.length === 18 ? liveParTimes : null) ?? (savedParTimes?.length === 18 ? savedParTimes : null);
    const tt = liveTurnTime ?? savedTurnTime;
    if (p) setPars(p);
    if (pt) setParTimes(pt);
    if (tt !== null && tt !== undefined) setTurnTime(tt);
  }, [savedPars, savedParTimes, savedTurnTime, savedTurnTimeBack, livePars, liveParTimes, liveTurnTime]);

  // Lifted up from QuickGeneratePanel so the H1/H10 group-list columns below can hide
  // themselves while the "Shotgun" tab is selected, and reappear for H1 only / H10 only / H1+H10.
  const [genMode, setGenMode] = useState("h1only"); // "h1only" | "h10only" | "both" | "shotgun"

  // Calculate nextNum from the loaded groups (once on mount)
  const nextNum = useRef((() => {
    const saved = loadSetup();
    const all = [...(saved?.groups1 ?? []), ...(saved?.groups10 ?? []), ...(saved?.groupsShotgun ?? [])];
    const maxN = all.reduce((acc, g) => {
      const m = g.name.match(/Group (\d+)/);
      return m ? Math.max(acc, Number(m[1])) : acc;
    }, 0);
    return maxN + 1;
  })());

  // ─── Save to memoryStorage every time the data changes ───────────────────────
  useEffect(() => {
    saveSetup({ groups1, groups10, groupsShotgun, pars, parTimes, playersPerGroup, turnTime, turnTimeBack });
  }, [groups1, groups10, groupsShotgun, pars, parTimes, playersPerGroup, turnTime, turnTimeBack]);

  const addGroup1 = () => {
    const n = nextNum.current++;
    setGroups1(g => [...g, {
      id: Date.now(),
      name: `Group ${n}`,
      startTime: makeNextTime(g),
      color: GROUP_COLORS_H1[(n - 1) % GROUP_COLORS_H1.length],
      startHole: 1,
    }]);
  };

  const addGroup10 = () => {
    const n = nextNum.current++;
    setGroups10(g => [...g, {
      id: Date.now() + 1,
      name: `Group ${n}`,
      startTime: makeNextTime(g),
      color: GROUP_COLORS_H10[(n - 1) % GROUP_COLORS_H10.length],
      startHole: 10,
    }]);
  };

  const removeGroup1  = (id) => setGroups1(g => g.filter(x => x.id !== id));
  const removeGroup10 = (id) => setGroups10(g => g.filter(x => x.id !== id));
  const removeGroupShotgun = (id) => setGroupsShotgun(g => g.filter(x => x.id !== id));

  const updateGroup1  = (id, field, val) => setGroups1(g => g.map(x => x.id === id ? { ...x, [field]: val } : x));
  const updateGroup10 = (id, field, val) => setGroups10(g => g.map(x => x.id === id ? { ...x, [field]: val } : x));
  const updateGroupShotgun = (id, field, val) => setGroupsShotgun(g => g.map(x => x.id === id ? { ...x, [field]: val } : x));

  // Switch the start point for a whole shotgun set (groups generated together from the same original start point) to another hole at once
  const swapShotgunSetHole = (setId, newHole) => {
    const meta = getStartHoleMeta(newHole);
    setGroupsShotgun(g => g.map(x =>
      x.shotgunSetId === setId ? { ...x, startHole: newHole, color: meta.color } : x
    ));
  };

  // ─── Clear modal state ────────────────────────────────────────────────────
  const [clearModal, setClearModal] = useState(null); // "h1" | "h10" | "shotgun" | null

  // ─── Quick Generate: append new groups ───────────────────────────────────
  const handleGenerate = (newG1, newG10, opts) => {
    if (opts?.replace) {
      // Rebuilding: reuse the old group ids where the group number matches, so
      // times already recorded stay attached to the right group. Otherwise the
      // new groups get fresh ids and start with an empty scorecard.
      const numberOf = g => { const m = g.name.match(/Group (\d+)/); return m ? Number(m[1]) : null; };
      const oldById = new Map();
      [...groups1, ...groups10, ...groupsShotgun].forEach(g => {
        const n = numberOf(g);
        if (n !== null) oldById.set(n, g.id);
      });
      const remap = g => {
        if (!opts.keepData) return g;
        const n = numberOf(g);
        const oldId = n !== null ? oldById.get(n) : undefined;
        return oldId !== undefined ? { ...g, id: oldId } : g;
      };
      setGroups1(newG1.map(remap));
      setGroups10(newG10.map(remap));
      setGroupsShotgun([]);          // a rebuilt H1/H10 schedule replaces shotgun too
      nextNum.current = newG1.length + newG10.length + 1;
      return;
    }
    setGroups1(prev => [...prev, ...newG1]);
    setGroups10(prev => [...prev, ...newG10]);
    const all = [...groups1, ...groups10, ...groupsShotgun, ...newG1, ...newG10];
    const maxN = all.reduce((acc, g) => {
      const m = g.name.match(/Group (\d+)/);
      return m ? Math.max(acc, Number(m[1])) : acc;
    }, 0);
    nextNum.current = maxN + 1;
  };

  // ─── Shotgun Generate: append new shotgun groups ──────────────────────────
  const handleGenerateShotgun = (newGroups) => {
    setGroupsShotgun(prev => [...prev, ...newGroups]);
    const all = [...groups1, ...groups10, ...groupsShotgun, ...newGroups];
    const maxN = all.reduce((acc, g) => {
      const m = g.name.match(/Group (\d+)/);
      return m ? Math.max(acc, Number(m[1])) : acc;
    }, 0);
    nextNum.current = maxN + 1;
  };

  // merge for onStart
  const allGroups = [...groups1, ...groups10, ...groupsShotgun];

  // While a round is live, edits here apply straight to the running schedule — no
  // button press needed. Debounced so typing a time doesn't spam the server, and
  // it never fires for an empty list (which would wipe the live session).
  const liveEditKey = JSON.stringify({
    g: allGroups.map(g => [g.id, g.name, g.startTime, g.startHole, g.section || ""]),
    pars, parTimes, turnTime, turnTimeBack,
  });
  const skipFirstApply = useRef(true);
  useEffect(() => {
    if (!hasLiveSession || !isAdmin || !onApplyLiveEdits) return;
    if (allGroups.length === 0) return;
    // Don't fire on the initial mount / on the sync-down from the live session
    if (skipFirstApply.current) { skipFirstApply.current = false; return; }
    const t = setTimeout(() => {
      onApplyLiveEdits(allGroups, pars, parTimes, playersPerGroup, turnTime, turnTimeBack);
    }, 700);
    return () => clearTimeout(t);
  }, [liveEditKey, hasLiveSession, isAdmin]);


  return (
    <div style={{ minHeight: "100vh", background: "#0d0f1a", color: "#eee", fontFamily: "'IBM Plex Mono', monospace" }}>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&family=Bebas+Neue&display=swap" rel="stylesheet" />

      <div style={{ background: "#141626", borderBottom: "1px solid #2a2d4a", padding: "14px 20px" }}>
        {/* One 2-column grid for the whole header, so the right-hand column —
            user, Log out and Clear Data — shares a single edge all the way down. */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, alignItems: "start" }}>
          {/* Left column: title */}
          <div style={{ minWidth: 0 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <div style={{ fontSize: 28, fontFamily: "'Bebas Neue'", letterSpacing: 4, color: "#4e9af1" }}>⛳ SETUP</div>
              <span style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1, color: "#ffd966", background: "#2a1a0066", border: "1px solid #ffd96655", borderRadius: 4, padding: "1px 6px", flexShrink: 0 }}>BETA</span>
            </div>
            <div style={{ fontSize: 11, color: "#8890b8", marginTop: 2, lineHeight: 1.4 }}>Golf Referee · Pace of Play System</div>
          </div>

          {/* Right column: user sits above Log out. Capped width so rotating to
              landscape doesn't stretch the Log out button across the screen. */}
          <div style={{ display: "flex", flexDirection: "column", gap: 8, minWidth: 0, maxWidth: 260, justifySelf: "end", width: "100%" }}>
            {currentUser && (
              <span style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 5 }}>
                <span style={{ fontSize: 12, color: "#8890b8" }}>👤</span>
                <span style={{ fontSize: 13, color: "#8899cc", fontWeight: 700, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{currentUser}</span>
                <span style={{ fontSize: 9, fontWeight: 700, color: isTrueAdmin ? "#ffd966" : "#4e9af1", background: isTrueAdmin ? "#2a1a0066" : "#001a2a66", border: `1px solid ${isTrueAdmin ? "#ffd96644" : "#4e9af144"}`, borderRadius: 4, padding: "0 6px", height: 18, display: "inline-flex", alignItems: "center", lineHeight: 1, letterSpacing: 1, flexShrink: 0 }}>
                  {isTrueAdmin ? "ADMIN" : "USER"}
                </span>
                {myPosition && (
                  <span style={{ fontSize: 9, fontWeight: 700, color: "#6effa0", background: "#0a2a1066", border: "1px solid #6effa055", borderRadius: 4, padding: "0 6px", height: 18, display: "inline-flex", alignItems: "center", lineHeight: 1, letterSpacing: 1, flexShrink: 0 }}>
                    {myPosition}
                  </span>
                )}
              </span>
            )}
            <div style={{ display: "flex", gap: 8, width: "100%" }}>
              <button onClick={onChangePassword}
                title="Change your password"
                style={{ background: "#0d0f1a", border: "1px solid #4e9af144", color: "#4e9af1", borderRadius: 10, height: 42, width: 46, cursor: "pointer", fontFamily: "inherit", fontSize: 15, flexShrink: 0 }}
              >🔑</button>
              <button onClick={onLogout}
                style={{ flex: 1, background: "#1a0d0d", border: "1px solid #ff707044", color: "#ff7070", borderRadius: 10, height: 42, cursor: "pointer", fontFamily: "inherit", fontSize: 14, fontWeight: 700, letterSpacing: 0.5, whiteSpace: "nowrap" }}
                onMouseEnter={e => e.currentTarget.style.borderColor = "#ff707088"}
                onMouseLeave={e => e.currentTarget.style.borderColor = "#ff707044"}
              >⏏ Log out</button>
            </div>
          </div>
        </div>
      </div>

      {/* Admin actions get their own band below a divider.
          Account management and wiping the session are system-admin only —
          TD/CR may edit this tournament's setup, not the whole system. */}
      {isTrueAdmin && (
        <div style={{ background: "#141626", borderBottom: "1px solid #2a2d4a", padding: "12px 20px", display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <button onClick={onManageUsers}
            style={{ minHeight: 56, background: "#1a1a0a", border: "1px solid #ffd96644", color: "#ffd966", borderRadius: 10, cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700, letterSpacing: 1, padding: "10px 8px", lineHeight: 1.4 }}
          >🔑<br />Manage Users</button>

          <button onClick={() => { if (window.confirm("Clear all group data and session?")) { onClearSession(); } }}
            style={{ minHeight: 56, background: "#1a0a0a", border: "1px solid #ff707044", color: "#ff7070", borderRadius: 10, cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700, letterSpacing: 1, padding: "10px 8px", lineHeight: 1.4 }}
          >🗑<br />Clear Data in Dashboard</button>
        </div>
      )}

      {/* Legend */}
      <div style={{ padding: "16px 24px 0" }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 12, color: "#9aa2c7" }}>
          <span style={{ color: "#4e9af1", fontWeight: 700, letterSpacing: 1 }}>Status criteria:</span>
          <span style={{ color: "#1560a8" }}>● In Position (Fast) (&lt; -10 min)</span>
          <span style={{ color: "#4e9af1" }}>● In Position (-9 to -1 min)</span>
          <span style={{ color: "#6effa0" }}>● On Time (0 min)</span>
          <span style={{ color: "#ffd966" }}>● Less Out of Position (+1 to +2 min)</span>
          <span style={{ color: "#ff8a80" }}>● Out of Position (+3 to +5 min)</span>
          <span style={{ color: "#b3261e" }}>● Out of Position (Slow) (+6 min and above)</span>
        </div>
      </div>

      <div style={{ padding: "24px 24px" }}>
        {/* ─── Tournament / Round context ──────────────────────────────────────
            Two clearly separate actions: change the whole tournament (goes back
            to the picker screen), or just change the round within it (popup). */}
        {(tournamentName || roundLabel) && (
          <div style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 12, marginBottom: 24, overflow: "hidden" }}>
            {/* Tournament block */}
            <div style={{ padding: "14px 20px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                <div style={{ fontSize: 15, fontWeight: 700, color: "#eee" }}>🏆 {tournamentName || "(untitled tournament)"}</div>
                {tournamentRunState && (
                  <span style={{ fontSize: 9, fontWeight: 700, letterSpacing: 1, color: tournamentRunState === "started" ? "#6effa0" : "#8890b8", background: tournamentRunState === "started" ? "#0a2a10" : "#0d0f1a", border: `1px solid ${tournamentRunState === "started" ? "#6effa055" : "#8890b844"}`, borderRadius: 4, padding: "0 6px", height: 18, display: "inline-flex", alignItems: "center", lineHeight: 1 }}>
                    {tournamentRunState === "started" ? "● STARTED" : "DRAFT"}
                  </span>
                )}
              </div>
              {hostVenue && <div style={{ fontSize: 12, color: "#8890b8", marginTop: 2 }}>{hostVenue}</div>}

              {/* TD/CR and admins run the event, so they start and pause it here */}
              {isAdmin && onToggleStarted && (
                <button onClick={onToggleStarted}
                  style={{
                    marginTop: 10, width: "100%", height: 34, borderRadius: 8, cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700,
                    color: tournamentRunState === "started" ? "#ffd966" : "#6effa0",
                    background: tournamentRunState === "started" ? "#1a1a0a" : "#0a2a10",
                    border: `1px solid ${tournamentRunState === "started" ? "#ffd96666" : "#6effa066"}`,
                  }}>
                  {tournamentRunState === "started" ? "⏸ Pause competition" : "▶ Start competition"}
                </button>
              )}

              {isAdmin && onSwitchTournament && (
                <button onClick={openTournamentPicker}
                  style={{ marginTop: 8, width: "100%", fontSize: 12, color: "#4e9af1", background: "#0d0f1a", border: "1px solid #4e9af166", borderRadius: 8, height: 34, padding: "0 16px", cursor: "pointer", fontFamily: "inherit", fontWeight: 700 }}>
                  ⇄ Switch Tournament
                </button>
              )}
            </div>

            {/* Divider between the two levels */}
            {roundLabel && <div style={{ borderTop: "1px solid #2a2d4a" }} />}

            {/* Round block — left-aligned with the tournament above it */}
            {roundLabel && (
              <div style={{ padding: "12px 20px 14px", display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ flex: 1, minWidth: 0, justifyContent: "center", fontSize: 13, fontWeight: 700, color: "#6effa0", background: "#1a4a2a66", border: "1px solid #6effa044", borderRadius: 8, height: 34, padding: "0 12px", display: "inline-flex", alignItems: "center" }}>
                  {roundLabel === "Q" ? "Round Q" : `Round ${roundLabel}`}
                </span>
                {isAdmin && tournamentId && (
                  <button onClick={openRoundPicker}
                    style={{ flex: 1, minWidth: 0, fontSize: 12, color: "#6effa0", background: "#0d0f1a", border: "1px solid #6effa066", borderRadius: 8, height: 34, padding: "0 12px", cursor: "pointer", fontFamily: "inherit", fontWeight: 700 }}>
                    ⇄ Switch Round
                  </button>
                )}
              </div>
            )}
          </div>
        )}

        {/* ─── Players per group ─────────────────────────────────────────────── */}
        <div style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 12, padding: 20, marginBottom: 24, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "nowrap" }}>
          <div style={{ fontSize: 13, color: "#4e9af1", letterSpacing: 1, fontWeight: 700, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>👥 PLAYERS PER GROUP</div>
          {isAdmin ? (
            <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
              {[2, 3, 4].map(n => (
                <button key={n} onClick={() => {
                  const oldTbl = parTimeTable(playersPerGroup);
                  const newTbl = parTimeTable(n);
                  // Re-apply defaults for holes still sitting on the old default,
                  // leaving any hand-tuned minutes untouched.
                  setParTimes(pt => pt.map((mins, i) => mins === oldTbl[pars[i]] ? newTbl[pars[i]] : mins));
                  setPlayersPerGroup(n);
                }}
                  style={{
                    width: 36, height: 36, borderRadius: 8, cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 16, flexShrink: 0,
                    background: playersPerGroup === n ? "#1a4a8a" : "#0d0f1a",
                    border: `1px solid ${playersPerGroup === n ? "#4e9af1" : "#2a2d4a"}`,
                    color: playersPerGroup === n ? "#fff" : "#8890b8",
                  }}>{n}</button>
              ))}
            </div>
          ) : (
            <span style={{ fontSize: 11, color: "#ffd966", background: "#2a1a0066", border: "1px solid #ffd96644", borderRadius: 5, padding: "3px 10px", letterSpacing: 1, flexShrink: 0, whiteSpace: "nowrap" }}>🔒 {playersPerGroup} players</span>
          )}
        </div>

        {/* Par Setup */}
        <div style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 12, padding: 24, marginBottom: 24 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, gap: 8, flexWrap: "wrap" }}>
            <div style={{ fontSize: 12, color: "#4e9af1", letterSpacing: 1, fontWeight: 700 }}>📋 HOLE SETUP — PAR & TIME</div>
            {isAdmin && (
              <button
                onClick={() => setParTimes(pars.map(p => parTimeTable(playersPerGroup)[p]))}
                style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#888", borderRadius: 6, padding: "4px 10px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, whiteSpace: "nowrap", flexShrink: 0 }}
              >↺ reset ({playersPerGroup}-ball)</button>
            )}
            {!isAdmin && (
              <span style={{ fontSize: 11, color: "#ffd966", background: "#2a1a0066", border: "1px solid #ffd96644", borderRadius: 5, padding: "3px 10px", letterSpacing: 1 }}>🔒 View only</span>
            )}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "44px repeat(9, 1fr)", gap: 6, marginBottom: 6 }}>
            <div />
            {Array.from({ length: 9 }, (_, i) => (
              <div key={i} style={{ fontSize: 11, color: "#8890b8", textAlign: "center" }}>H{i + 1}</div>
            ))}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "44px repeat(9, 1fr)", gap: 6, marginBottom: 6, alignItems: "center" }}>
            <div style={{ fontSize: 12, color: "#9aa2c7", textAlign: "right", paddingRight: 8 }}>Par</div>
            {pars.slice(0, 9).map((p, i) => (
              <select key={i} value={p}
                disabled={!isAdmin}
                onChange={e => {
                  const nxt = [...pars]; nxt[i] = Number(e.target.value);
                  setPars(nxt);
                  setParTimes(pt => { const n = [...pt]; const tbl = parTimeTable(playersPerGroup); if (n[i] === tbl[p]) n[i] = tbl[Number(e.target.value)]; return n; });
                }}
                style={{ width: "100%", background: isAdmin ? "#1e2135" : "#0d0f1a", border: `1px solid ${isAdmin ? "#2a2d4a" : "#1a1d2e"}`, color: isAdmin ? "#eee" : "#666", borderRadius: 6, padding: "4px 0", textAlign: "center", textAlignLast: "center", fontSize: 14, fontFamily: "inherit", cursor: isAdmin ? "pointer" : "default", appearance: "none", WebkitAppearance: "none", MozAppearance: "none" }}
              >
                <option value={3}>3</option><option value={4}>4</option><option value={5}>5</option>
              </select>
            ))}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "44px repeat(9, 1fr)", gap: 6, marginBottom: 12, alignItems: "center" }}>
            <div style={{ fontSize: 12, color: "#9aa2c7", textAlign: "right", paddingRight: 8 }}>min</div>
            {parTimes.slice(0, 9).map((t, i) => (
              <select key={i} value={t}
                disabled={!isAdmin}
                onChange={e => { if (!isAdmin) return; const n = [...parTimes]; n[i] = Number(e.target.value); setParTimes(n); }}
                style={{ width: "100%", background: isAdmin ? "#1a2a1a" : "#0d0f1a", border: `1px solid ${isAdmin ? "#2a4a2a" : "#1a2a1a"}`, color: isAdmin ? "#6effa0" : "#3a5a3a", borderRadius: 6, padding: "4px 0", textAlign: "center", textAlignLast: "center", fontSize: 14, fontFamily: "inherit", cursor: isAdmin ? "pointer" : "default", appearance: "none", WebkitAppearance: "none", MozAppearance: "none" }}
              >
                {PAR_TIME_CHOICES.map(v => <option key={v} value={v}>{v}</option>)}
              </select>
            ))}
          </div>

          <div style={{ borderTop: "1px dashed #2a2d4a", marginBottom: 12 }} />

          <div style={{ display: "grid", gridTemplateColumns: "44px repeat(9, 1fr)", gap: 6, marginBottom: 6 }}>
            <div />
            {Array.from({ length: 9 }, (_, i) => (
              <div key={i} style={{ fontSize: 11, color: "#8890b8", textAlign: "center" }}>H{i + 10}</div>
            ))}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "44px repeat(9, 1fr)", gap: 6, marginBottom: 6, alignItems: "center" }}>
            <div style={{ fontSize: 12, color: "#9aa2c7", textAlign: "right", paddingRight: 8 }}>Par</div>
            {pars.slice(9).map((p, i) => (
              <select key={i} value={p}
                disabled={!isAdmin}
                onChange={e => {
                  const nxt = [...pars]; nxt[i + 9] = Number(e.target.value);
                  setPars(nxt);
                  setParTimes(pt => { const n = [...pt]; const tbl = parTimeTable(playersPerGroup); if (n[i+9] === tbl[p]) n[i+9] = tbl[Number(e.target.value)]; return n; });
                }}
                style={{ width: "100%", background: isAdmin ? "#1e2135" : "#0d0f1a", border: `1px solid ${isAdmin ? "#2a2d4a" : "#1a1d2e"}`, color: isAdmin ? "#eee" : "#666", borderRadius: 6, padding: "4px 0", textAlign: "center", textAlignLast: "center", fontSize: 14, fontFamily: "inherit", cursor: isAdmin ? "pointer" : "default", appearance: "none", WebkitAppearance: "none", MozAppearance: "none" }}
              >
                <option value={3}>3</option><option value={4}>4</option><option value={5}>5</option>
              </select>
            ))}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "44px repeat(9, 1fr)", gap: 6, marginBottom: 16, alignItems: "center" }}>
            <div style={{ fontSize: 12, color: "#9aa2c7", textAlign: "right", paddingRight: 8 }}>min</div>
            {parTimes.slice(9).map((t, i) => (
              <select key={i} value={t}
                disabled={!isAdmin}
                onChange={e => { if (!isAdmin) return; const n = [...parTimes]; n[i + 9] = Number(e.target.value); setParTimes(n); }}
                style={{ width: "100%", background: isAdmin ? "#1a2a1a" : "#0d0f1a", border: `1px solid ${isAdmin ? "#2a4a2a" : "#1a2a1a"}`, color: isAdmin ? "#6effa0" : "#3a5a3a", borderRadius: 6, padding: "4px 0", textAlign: "center", textAlignLast: "center", fontSize: 14, fontFamily: "inherit", cursor: isAdmin ? "pointer" : "default", appearance: "none", WebkitAppearance: "none", MozAppearance: "none" }}
              >
                {PAR_TIME_CHOICES.map(v => <option key={v} value={v}>{v}</option>)}
              </select>
            ))}
          </div>

          <div style={{ borderTop: "1px dashed #2a2d4a", marginBottom: 12 }} />

          {/* Transit time — the two turn walks are different distances on most
              courses, so each direction gets its own value. */}
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 12, color: "#ffd966", fontWeight: 700, marginBottom: 8 }}>🚶 Transit Time</div>
            {[
              { label: "H9 → H10", sub: "starting hole 1", val: turnTime, set: setTurnTime },
              { label: "H18 → H1", sub: "starting hole 10", val: turnTimeBack, set: setTurnTimeBack },
            ].map(row => (
              <div key={row.label} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8, flexWrap: "wrap" }}>
                <div style={{ flex: 1, minWidth: 150 }}>
                  <div style={{ fontSize: 12, color: "#eee", fontWeight: 700 }}>{row.label}</div>
                  <div style={{ fontSize: 10, color: "#8890b8" }}>{row.sub}</div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 4, flexShrink: 0 }}>
                  <button onClick={() => isAdmin && row.set(t => Math.max(0, t - 1))} disabled={!isAdmin}
                    style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#aaa", borderRadius: 6, width: 34, height: 34, cursor: isAdmin ? "pointer" : "default", fontSize: 16, fontFamily: "inherit" }}>−</button>
                  <input type="number" min="0" value={row.val} disabled={!isAdmin}
                    onChange={e => row.set(Math.max(0, Number(e.target.value) || 0))}
                    style={{ width: 54, background: "#0d0f1a", border: `1px solid ${row.val > 0 ? "#ffd96666" : "#2a2d4a"}`, color: "#ffd966", fontFamily: "'Bebas Neue'", fontSize: 20, textAlign: "center", borderRadius: 6, padding: "3px 0", outline: "none" }} />
                  <button onClick={() => isAdmin && row.set(t => t + 1)} disabled={!isAdmin}
                    style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#aaa", borderRadius: 6, width: 34, height: 34, cursor: isAdmin ? "pointer" : "default", fontSize: 16, fontFamily: "inherit" }}>+</button>
                  <span style={{ fontSize: 12, color: "#8890b8", marginLeft: 2 }}>min</span>
                </div>
              </div>
            ))}
          </div>

          {/* Summary: front/back nine on one row, grand total on its own row below */}
          <div style={{ display: "flex", alignItems: "stretch", gap: 8, flexWrap: "nowrap", marginBottom: 8 }}>
            <div style={{ flex: 1, background: "#0d0f1a", border: "1px solid #2a2d4a", borderRadius: 8, padding: "8px 10px", textAlign: "center", minWidth: 0 }}>
              <div style={{ fontSize: 10, color: "#9aa2c7", letterSpacing: 1, marginBottom: 2, whiteSpace: "nowrap" }}>H1–H9</div>
              <div style={{ fontSize: 14, color: "#8899cc", fontWeight: 700, whiteSpace: "nowrap" }}>{minToHM(parTimes.slice(0, 9).reduce((a, b) => a + b, 0))}</div>
            </div>
            <div style={{ flex: 1, background: "#0d0f1a", border: "1px solid #2a2d4a", borderRadius: 8, padding: "8px 10px", textAlign: "center", minWidth: 0 }}>
              <div style={{ fontSize: 10, color: "#9aa2c7", letterSpacing: 1, marginBottom: 2, whiteSpace: "nowrap" }}>H10–H18</div>
              <div style={{ fontSize: 14, color: "#8899cc", fontWeight: 700, whiteSpace: "nowrap" }}>{minToHM(parTimes.slice(9).reduce((a, b) => a + b, 0))}</div>
            </div>
          </div>
          <div style={{ background: "#0a1a0a", border: "1px solid #2a4a2a", borderRadius: 8, padding: "10px 14px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
            {(() => {
              const holes = parTimes.reduce((a, b) => a + b, 0);
              const fromH1 = holes + (turnTime || 0);
              const fromH10 = holes + (turnTimeBack || 0);
              // Only split the total when the two transit walks actually differ
              if (fromH1 === fromH10) {
                return (
                  <>
                    <div style={{ fontSize: 11, color: "#6effa0", letterSpacing: 1 }}>TOTAL (Included Transit Time)</div>
                    <div style={{ fontSize: 16, color: "#6effa0", fontWeight: 700, whiteSpace: "nowrap" }}>{minToHM(fromH1)}</div>
                  </>
                );
              }
              return (
                <>
                  <div style={{ fontSize: 11, color: "#6effa0", letterSpacing: 1 }}>TOTAL (Included Transit Time)</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 3, alignItems: "flex-end" }}>
                    <div style={{ display: "flex", alignItems: "baseline", gap: 8, whiteSpace: "nowrap" }}>
                      <span style={{ fontSize: 10, color: "#8890b8" }}>from H1</span>
                      <span style={{ fontSize: 16, color: "#6effa0", fontWeight: 700 }}>{minToHM(fromH1)}</span>
                    </div>
                    <div style={{ display: "flex", alignItems: "baseline", gap: 8, whiteSpace: "nowrap" }}>
                      <span style={{ fontSize: 10, color: "#8890b8" }}>from H10</span>
                      <span style={{ fontSize: 16, color: "#6effa0", fontWeight: 700 }}>{minToHM(fromH10)}</span>
                    </div>
                  </div>
                </>
              );
            })()}
          </div>

        </div>

        {/* ─── Auto-fill Time Panel ──────────────────────────────────────────── */}
        {isAdmin && (
          <QuickGeneratePanel
            onGenerate={handleGenerate}
            existingGroups1={groups1}
            existingGroups10={groups10}
            pars={pars}
            onGenerateShotgun={handleGenerateShotgun}
            existingGroupsShotgun={groupsShotgun}
            mode={genMode}
            onModeChange={setGenMode}
          />
        )}

        {/* Groups — stacked */}
        <div style={{ display: "flex", flexDirection: "column", gap: 16, marginBottom: 24 }}>

          {!isAdmin && (
            <div style={{ background: "#141626", border: "1px solid #ffd96633", borderRadius: 12, padding: "16px 20px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                <span style={{ fontSize: 20 }}>🔒</span>
                <div style={{ fontSize: 14, color: "#ffd966", fontWeight: 700, letterSpacing: 1 }}>View-only mode (User)</div>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 4, paddingLeft: 30 }}>
                <div style={{ fontSize: 12, color: "#888" }}>✗ Cannot edit Par / time per hole</div>
                <div style={{ fontSize: 12, color: "#888" }}>✗ Cannot add or edit player groups</div>
                <div style={{ fontSize: 12, color: "#6effa0", marginTop: 2 }}>✓ Can view the Schedule table and track Pace</div>
              </div>
            </div>
          )}

          {/* H1 Column */}
          {isAdmin && genMode !== "shotgun" && (
          <div style={{ background: "#141626", border: "1px solid #6effa033", borderRadius: 12, padding: 20 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <div style={{ fontSize: 13, color: "#6effa0", letterSpacing: 2, fontWeight: 700 }}>🟢 Start hole 1 → 18</div>
              {groups1.length > 0 && (
                <button
                  onClick={() => setClearModal("h1")}
                  style={{ background: "#1a0a0a", border: "1px solid #ff707044", color: "#ff7070", borderRadius: 6, padding: "3px 10px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700 }}
                >🗑 Clear all</button>
              )}
            </div>
            {groups1.length === 0 && (
              <div style={{ color: "#666f99", fontSize: 12, textAlign: "center", padding: "20px 0" }}>No groups yet</div>
            )}
            {groups1.map(g => (
              <div key={g.id} style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8 }}>
                <div style={{ width: 5, height: 28, borderRadius: 3, background: g.color, flexShrink: 0 }} />
                <input
                  value={g.name}
                  onChange={e => updateGroup1(g.id, "name", e.target.value)}
                  style={{ flex: 1, minWidth: 0, background: "#1e2135", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 6, padding: "5px 8px", fontFamily: "inherit", fontSize: 13 }}
                />
                <input
                  type="color"
                  value={g.color}
                  onChange={e => updateGroup1(g.id, "color", e.target.value)}
                  style={{ width: 40, height: 40, border: "none", borderRadius: 6, cursor: "pointer", background: "none", flexShrink: 0 }}
                />
                <input
                  type="time"
                  value={g.startTime}
                  onChange={e => updateGroup1(g.id, "startTime", e.target.value)}
                  style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 6, padding: "4px 8px", fontFamily: "inherit", fontSize: 13, flexShrink: 0 }}
                />
                <button onClick={() => removeGroup1(g.id)} style={{ background: "#3a1a1a", border: "none", color: "#ff7070", borderRadius: 6, padding: "5px 8px", cursor: "pointer", fontSize: 13, flexShrink: 0 }}>✕</button>
              </div>
            ))}
            <button onClick={addGroup1} style={{ marginTop: 6, background: "#0d1a0d", border: "1px dashed #6effa044", color: "#6effa0", borderRadius: 6, padding: "7px 12px", cursor: "pointer", fontFamily: "inherit", fontSize: 13, width: "100%" }}>
              + Add H1 group
            </button>
          </div>
          )}

          {/* H10 Column */}
          {isAdmin && genMode !== "shotgun" && (
          <div style={{ background: "#141626", border: "1px solid #4e9af133", borderRadius: 12, padding: 20 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <div style={{ fontSize: 13, color: "#4e9af1", letterSpacing: 2, fontWeight: 700 }}>🔵 Start hole 10 → 18 → 1 → 9</div>
              {groups10.length > 0 && (
                <button
                  onClick={() => setClearModal("h10")}
                  style={{ background: "#1a0a0a", border: "1px solid #ff707044", color: "#ff7070", borderRadius: 6, padding: "3px 10px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700 }}
                >🗑 Clear all</button>
              )}
            </div>
            {groups10.length === 0 && (
              <div style={{ color: "#666f99", fontSize: 12, textAlign: "center", padding: "20px 0" }}>No groups yet</div>
            )}
            {groups10.map(g => (
              <div key={g.id} style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8 }}>
                <div style={{ width: 5, height: 28, borderRadius: 3, background: g.color, flexShrink: 0 }} />
                <input
                  value={g.name}
                  onChange={e => updateGroup10(g.id, "name", e.target.value)}
                  style={{ flex: 1, minWidth: 0, background: "#1e2135", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 6, padding: "5px 8px", fontFamily: "inherit", fontSize: 13 }}
                />
                <input
                  type="color"
                  value={g.color}
                  onChange={e => updateGroup10(g.id, "color", e.target.value)}
                  style={{ width: 40, height: 40, border: "none", borderRadius: 6, cursor: "pointer", background: "none", flexShrink: 0 }}
                />
                <input
                  type="time"
                  value={g.startTime}
                  onChange={e => updateGroup10(g.id, "startTime", e.target.value)}
                  style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 6, padding: "4px 8px", fontFamily: "inherit", fontSize: 13, flexShrink: 0 }}
                />
                <button onClick={() => removeGroup10(g.id)} style={{ background: "#3a1a1a", border: "none", color: "#ff7070", borderRadius: 6, padding: "5px 8px", cursor: "pointer", fontSize: 13, flexShrink: 0 }}>✕</button>
              </div>
            ))}
            <button onClick={addGroup10} style={{ marginTop: 6, background: "#0d0f1a", border: "1px dashed #4e9af144", color: "#4e9af1", borderRadius: 6, padding: "7px 12px", cursor: "pointer", fontFamily: "inherit", fontSize: 13, width: "100%" }}>
              + Add H10 group
            </button>
          </div>
          )}

          {/* Shotgun Column */}
          {isAdmin && groupsShotgun.length > 0 && (() => {
            const shotgunHoles = resolveShotgunStartHoles(pars); // [1, 5|6, 10, 14|15] based on current par conditions
            // Group into sub-sets by shotgunSetId (groups generated together from the same original start point)
            const sets = [];
            const seen = new Map();
            groupsShotgun.forEach(g => {
              const key = g.shotgunSetId ?? `legacy-${g.startHole}`; // fallback for old groups without a setId
              if (!seen.has(key)) { seen.set(key, sets.length); sets.push([]); }
              sets[seen.get(key)].push(g);
            });

            return (
              <div style={{ background: "#1a1410", border: "1px solid #f1734e33", borderRadius: 12, padding: 20 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
                  <div style={{ fontSize: 13, color: "#f1734e", letterSpacing: 2, fontWeight: 700 }}>🔫 SHOTGUN (4 holes)</div>
                  <button
                    onClick={() => setClearModal("shotgun")}
                    style={{ background: "#1a0a0a", border: "1px solid #ff707044", color: "#ff7070", borderRadius: 6, padding: "3px 10px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700 }}
                  >🗑 Clear all</button>
                </div>

                {sets.map((setGroups, si) => {
                  const setId = setGroups[0].shotgunSetId ?? `legacy-${setGroups[0].startHole}`;
                  const currentHole = setGroups[0].startHole;
                  const meta = getStartHoleMeta(currentHole);
                  const firstName = setGroups[0].name, lastName = setGroups[setGroups.length - 1].name;
                  return (
                    <div key={setId} style={{ marginBottom: si === sets.length - 1 ? 0 : 16 }}>
                      {/* Set header: group names in the set + toggle to switch the start point for the whole set */}
                      <div style={{
                        display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap",
                        background: "#0d0f1a", border: `1px solid ${meta.color}44`, borderRadius: 8,
                        padding: "8px 10px", marginBottom: 8,
                      }}>
                        <span style={{ fontSize: 12, color: "#888" }}>
                          Set {setGroups.length > 1 ? `${firstName} – ${lastName}` : firstName}
                        </span>
                        <span style={{ marginLeft: "auto", fontSize: 11, color: "#9aa2c7" }}>Switch start point for whole set →</span>
                        <div style={{ display: "flex", gap: 4 }}>
                          {shotgunHoles.map(h => {
                            const hMeta = getStartHoleMeta(h);
                            const active = h === currentHole;
                            return (
                              <button
                                key={h}
                                onClick={() => swapShotgunSetHole(setId, h)}
                                disabled={active}
                                style={{
                                  minWidth: 34, padding: "5px 8px", borderRadius: 6, cursor: active ? "default" : "pointer",
                                  fontFamily: "inherit", fontSize: 12, fontWeight: 700,
                                  background: active ? `${hMeta.color}33` : "#1e2135",
                                  border: `1px solid ${active ? hMeta.color : "#2a2d4a"}`,
                                  color: active ? hMeta.color : "#888",
                                }}
                              >H{h}</button>
                            );
                          })}
                        </div>
                      </div>

                      {setGroups.map(g => {
                        const gMeta = getStartHoleMeta(g.startHole);
                        return (
                          <div key={g.id} style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8 }}>
                            <div style={{
                              minWidth: 32, height: 28, borderRadius: 6, background: `${gMeta.color}22`,
                              border: `1px solid ${gMeta.color}66`, color: gMeta.color, fontSize: 12, fontWeight: 700,
                              display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                            }}>H{g.startHole}</div>
                            <input
                              value={g.name}
                              onChange={e => updateGroupShotgun(g.id, "name", e.target.value)}
                              style={{ flex: 1, minWidth: 0, background: "#1e2135", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 6, padding: "5px 8px", fontFamily: "inherit", fontSize: 13 }}
                            />
                            <input
                              type="color"
                              value={g.color}
                              onChange={e => updateGroupShotgun(g.id, "color", e.target.value)}
                              style={{ width: 40, height: 40, border: "none", borderRadius: 6, cursor: "pointer", background: "none", flexShrink: 0 }}
                            />
                            <input
                              type="time"
                              value={g.startTime}
                              onChange={e => updateGroupShotgun(g.id, "startTime", e.target.value)}
                              style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 6, padding: "4px 8px", fontFamily: "inherit", fontSize: 13, flexShrink: 0 }}
                            />
                            <button onClick={() => removeGroupShotgun(g.id)} style={{ background: "#3a1a1a", border: "none", color: "#ff7070", borderRadius: 6, padding: "5px 8px", cursor: "pointer", fontSize: 13, flexShrink: 0 }}>✕</button>
                          </div>
                        );
                      })}
                    </div>
                  );
                })}
              </div>
            );
          })()}
        </div>

        {!hasLiveSession && (
        <button
          onClick={() => {
            if (!isAdmin && allGroups.length === 0) return;
            onStart(allGroups, pars, parTimes, playersPerGroup, turnTime, turnTimeBack);
          }}
          disabled={!isAdmin && allGroups.length === 0}
          style={{
            width: "100%", background: (!isAdmin && allGroups.length === 0) ? "#1a1d2e" : "linear-gradient(135deg, #1a4a8a, #4e9af1)", border: "none",
            color: (!isAdmin && allGroups.length === 0) ? "#444" : "#fff", borderRadius: 10, padding: "16px",
            cursor: (!isAdmin && allGroups.length === 0) ? "not-allowed" : "pointer",
            fontFamily: "'Bebas Neue'", letterSpacing: 3, fontSize: 20,
          }}
        >
          {(!isAdmin && allGroups.length === 0) ? "🔒 Waiting for Admin to set up groups" : "▶ Start tracking PACE OF PLAY"}
        </button>
        )}

        {hasLiveSession && (
          <button
            onClick={onGoToDashboard}
            style={{
              width: "100%", marginTop: 12, background: "linear-gradient(135deg, #0a3a5a, #4e9af1)", border: "none",
              color: "#fff", borderRadius: 10, padding: "16px", cursor: "pointer",
              fontFamily: "'Bebas Neue'", letterSpacing: 3, fontSize: 20,
            }}
          >📊 Back to Dashboard</button>
        )}

        {/* Build marker — bump this string whenever the file changes so it's obvious
            at a glance whether the browser is running the latest deploy. */}
        <div style={{ textAlign: "center", fontSize: 10, color: "#3a3f5a", marginTop: 20, letterSpacing: 1 }}>
          build {APP_BUILD}
        </div>
      </div>

      {/* Quick tournament switcher — list only; full management is on the Tournament screen */}
      {showTournamentPicker && (
        <div onClick={() => setShowTournamentPicker(false)} style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 1000, padding: 16, overflowY: "auto" }}>
          <div onClick={e => e.stopPropagation()} style={{ background: "#141626", border: "1px solid #4e9af166", borderRadius: 14, padding: 22, width: "100%", maxWidth: 420, boxShadow: "0 20px 60px #000", fontFamily: "'IBM Plex Mono', monospace", marginTop: 24 }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 2, color: "#4e9af1", marginBottom: 14 }}>⇄ Switch Tournament</div>

            {tPickerLoading ? (
              <div style={{ padding: 24, textAlign: "center", color: "#8890b8", fontSize: 13 }}>Loading…</div>
            ) : tPickerList.length === 0 ? (
              <div style={{ padding: 20, textAlign: "center", color: "#666", fontSize: 13 }}>No other competitions available to you</div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 14, maxHeight: 340, overflowY: "auto" }}>
                {tPickerList.map(t => {
                  const isCurrent = t.id === tournamentId;
                  return (
                    <button key={t.id}
                      onClick={() => { setShowTournamentPicker(false); if (!isCurrent) onPickTournament?.(t); }}
                      style={{
                        textAlign: "left", padding: "11px 13px", borderRadius: 9, cursor: isCurrent ? "default" : "pointer", fontFamily: "inherit",
                        background: isCurrent ? "#1a4a8a" : "#0d0f1a",
                        border: `1px solid ${isCurrent ? "#4e9af1" : "#2a2d4a"}`,
                      }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                        <span style={{ fontSize: 14, fontWeight: 700, color: "#eee" }}>{t.name || "(untitled)"}</span>
                        {isCurrent && <span style={{ fontSize: 9, fontWeight: 700, letterSpacing: 1, color: "#6effa0", background: "#0a2a1066", border: "1px solid #6effa055", borderRadius: 4, padding: "0 6px", height: 18, display: "inline-flex", alignItems: "center", lineHeight: 1 }}>● CURRENT</span>}
                      </div>
                      {t.host_venue && <div style={{ fontSize: 11, color: "#8890b8", marginTop: 2 }}>{t.host_venue}</div>}
                    </button>
                  );
                })}
              </div>
            )}

            <div style={{ display: "flex", gap: 8 }}>
              {isTrueAdmin && onSwitchTournament && (
                <button onClick={() => { setShowTournamentPicker(false); onSwitchTournament(); }}
                  style={{ flex: 1, background: "#0d0f1a", border: "1px dashed #4e9af166", color: "#4e9af1", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700 }}>
                  ⚙ Manage tournaments
                </button>
              )}
              <button onClick={() => setShowTournamentPicker(false)}
                style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "10px 18px", cursor: "pointer", fontFamily: "inherit", fontSize: 13 }}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Round switcher — pick another round of the SAME tournament */}
      {showRoundPicker && (
        <div onClick={() => setShowRoundPicker(false)} style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }}>
          <div onClick={e => e.stopPropagation()} style={{ background: "#141626", border: "1px solid #6effa066", borderRadius: 14, padding: 22, width: "100%", maxWidth: 380, boxShadow: "0 20px 60px #000", fontFamily: "'IBM Plex Mono', monospace" }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 2, color: "#6effa0", marginBottom: 2 }}>⇄ Switch Round</div>
            <div style={{ fontSize: 12, color: "#8890b8", marginBottom: 16 }}>{tournamentName}</div>

            {pickerLoading ? (
              <div style={{ padding: 24, textAlign: "center", color: "#8890b8", fontSize: 13 }}>Loading…</div>
            ) : (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8, marginBottom: 16 }}>
                {pickerLabels.map(label => {
                  const r = pickerRounds.find(rr => rr.label === label);
                  const isCurrent = label === roundLabel;
                  // Rounds keep their own data now, so this only signals "already used"
                  const hasData = !!r && (r.status === "finished" || r.status === "live");
                  return (
                    <button key={label}
                      onClick={() => { setShowRoundPicker(false); if (!isCurrent) onPickRound?.(label); }}
                      style={{
                        padding: "14px 0", borderRadius: 10, cursor: isCurrent ? "default" : "pointer", fontFamily: "inherit",
                        background: isCurrent ? "#1a4a2a" : hasData ? "#141a2a" : "#0d0f1a",
                        border: `1px solid ${isCurrent ? "#6effa0" : hasData ? "#4e9af155" : "#2a2d4a"}`,
                        display: "flex", flexDirection: "column", alignItems: "center", gap: 3,
                      }}>
                      <span style={{ fontFamily: "'Bebas Neue'", fontSize: 22, letterSpacing: 1, color: isCurrent ? "#6effa0" : "#eee" }}>
                        {label === "Q" ? "Q" : `R${label}`}
                      </span>
                      <span style={{ fontSize: 9, letterSpacing: 1, color: isCurrent ? "#6effa0" : hasData ? "#4e9af1" : "#8890b8" }}>
                        {isCurrent ? "● LIVE" : hasData ? "HAS DATA" : r ? "SET UP" : "NOT STARTED"}
                      </span>
                    </button>
                  );
                })}
              </div>
            )}

            <button onClick={() => setShowRoundPicker(false)}
              style={{ width: "100%", background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "inherit", fontSize: 13 }}>
              Cancel
            </button>
          </div>
        </div>
      )}

      {clearModal && (
        <div style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }}>
          <div style={{ background: "#141626", border: "1px solid #ff707088", borderRadius: 14, padding: 28, minWidth: 300, boxShadow: "0 20px 60px #000", fontFamily: "'IBM Plex Mono', monospace" }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 22, letterSpacing: 3, color: "#ff7070", marginBottom: 12 }}>🗑 Clear all groups</div>
            <div style={{ fontSize: 14, color: "#aaa", marginBottom: 6 }}>
              Delete groups <span style={{ color: "#eee", fontWeight: 700 }}>
                {clearModal === "h1" ? `All ${groups1.length} H1 group(s)`
                  : clearModal === "h10" ? `All ${groups10.length} H10 group(s)`
                  : `All ${groupsShotgun.length} Shotgun group(s)`}
              </span>
            </div>
            <div style={{ fontSize: 13, color: "#9aa2c7", marginBottom: 22 }}>The group list will be deleted; per-hole times and Par remain unchanged</div>
            <div style={{ display: "flex", gap: 10 }}>
              <button
                onClick={() => {
                  const remaining1 = clearModal === "h1" ? [] : groups1;
                  const remaining10 = clearModal === "h10" ? [] : groups10;
                  const remainingShotgun = clearModal === "shotgun" ? [] : groupsShotgun;
                  const allRemaining = [...remaining1, ...remaining10, ...remainingShotgun];
                  const maxN = allRemaining.reduce((acc, g) => {
                    const m = g.name.match(/Group (\d+)/);
                    return m ? Math.max(acc, Number(m[1])) : acc;
                  }, 0);
                  nextNum.current = maxN + 1;
                  if (clearModal === "h1") setGroups1([]);
                  else if (clearModal === "h10") setGroups10([]);
                  else setGroupsShotgun([]);
                  setClearModal(null);
                }}
                style={{ flex: 1, background: "#2a0a0a", border: "1px solid #ff7070", color: "#ff7070", borderRadius: 8, padding: "11px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 16, letterSpacing: 2 }}
              >✓ Confirm</button>
              <button
                onClick={() => setClearModal(null)}
                style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "11px 16px", cursor: "pointer", fontFamily: "inherit", fontSize: 14 }}
              >Cancel</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Time Input Helper ────────────────────────────────────────────────────────
function TimeInput({ value, onChange, label, color = "#4e9af1" }) {
  const parts = value ? value.split(":") : ["", ""];
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
      {label && <span style={{ fontSize: 12, color: "#9aa2c7", minWidth: 60 }}>{label}</span>}
      <input type="number" min={0} max={23} placeholder="HH"
        value={parts[0]}
        onChange={e => { const hh = String(Number(e.target.value)).padStart(2,"0"); onChange(`${hh}:${parts[1] || "00"}`); }}
        style={{ width: 48, background: "#1e2135", border: `1px solid ${color}44`, color: "#eee", borderRadius: 8, padding: "5px 6px", fontFamily: "inherit", fontSize: 16, fontWeight: 700, textAlign: "center" }}
      />
      <span style={{ color, fontSize: 18, fontWeight: 700 }}>:</span>
      <input type="number" min={0} max={59} placeholder="MM"
        value={parts[1]}
        onChange={e => { const mm = String(Number(e.target.value)).padStart(2,"0"); onChange(`${parts[0] || "00"}:${mm}`); }}
        style={{ width: 48, background: "#1e2135", border: `1px solid ${color}44`, color: "#eee", borderRadius: 8, padding: "5px 6px", fontFamily: "inherit", fontSize: 16, fontWeight: 700, textAlign: "center" }}
      />
    </div>
  );
}

// ─── Group Monitor ────────────────────────────────────────────────────────────
function GroupMonitor({ group, pars, parTimes, playersPerGroup, schedule, onUpdate, onBack, currentUser,
  isSuspended, suspensions, totalOffsetMin, pendingStopTime, onLogout, allGroups, onSwitchGroup, hideLog, onRecorded, closeLabel, compact, statusOverride }) {
  const initHoleData = () =>
    group.holeData ?? Array(18).fill(null).map(() => ({ startTime: null, endTime: null }));
  const numPlayers = playersPerGroup || 3; // how many P1..Pn quick-select buttons to offer for TM / Bad Time
  const playerNums = Array.from({ length: numPlayers }, (_, i) => i + 1);

  const holeOrder = getHoleOrder(group.startHole || 1);
  const [currentSlot, setCurrentSlot] = useState(group.currentHole ?? 0);
  const currentHole = holeOrder[Math.min(currentSlot, 17)];
  const [holeData, setHoleData] = useState(initHoleData);
  const [records, setRecords] = useState(group.records ?? Array(18).fill(null));
  const [now, setNow] = useState(nowInMin());
  const [recordedEnd, setRecordedEnd] = useState(null);
  const [inputMode, setInputMode] = useState("stamp");
  const [diffManual, setDiffManual] = useState(0);

  const [editingHole, setEditingHole] = useState(null);
  const [editField, setEditField] = useState("end");
  const [editVal, setEditVal] = useState("");
  const [editingDiff, setEditingDiff] = useState(null); // holeIdx being diff-edited
  const [editDiffVal, setEditDiffVal] = useState(0);

  const startEditDiff = (i) => {
    setEditingDiff(i);
    setEditDiffVal(diffAtHole(i) ?? 0);
  };
  const confirmEditDiff = (i) => {
    const hd = holeData[i];
    const deadline = (adjustedSchedule[i] ?? 0) + (parTimes?.[i] ?? 14);
    const newEndMin = deadline + Number(editDiffVal) - 1;
    const newEndTime = minToTime(Math.max(0, newEndMin));
    const updated = { ...(hd ?? {}), endTime: newEndTime };
    const { nxtHD, nxtRec } = commitRecord(i, updated);

    // Same as Confirm-Hole / manual edit: if the hole is complete, auto-log MN / TM
    // so the status keeps following forward here too.
    if (updated.startTime && updated.endTime) {
      const nxtLogs = autoLogMonitoring(i, actionLogs);
      if (nxtLogs !== actionLogs) {
        setActionLogs(nxtLogs);
        onUpdate({ holeData: nxtHD, records: nxtRec, currentHole: currentSlot, actionLogs: nxtLogs, mnActive, mnName, tmActive, tmName, tmTarget });
      }
    }
    setEditingDiff(null);
  };
  const [delayMin, setDelayMin] = useState(group.delayMin ?? 0);

  // ── Switch between groups without leaving the time-recording page ──
  const [editingGroupNum, setEditingGroupNum] = useState(false);
  const [groupNumInput, setGroupNumInput] = useState("");

  const extractGroupNumber = (name) => {
    const m = /\d+/.exec(name || "");
    return m ? m[0] : null;
  };

  const sortedGroups = useMemo(() => {
    return (allGroups || []).slice().sort((a, b) => {
      const na = Number(extractGroupNumber(a.name));
      const nb = Number(extractGroupNumber(b.name));
      if (!Number.isNaN(na) && !Number.isNaN(nb) && na !== nb) return na - nb;
      return (a.name || "").localeCompare(b.name || "");
    });
  }, [allGroups]);
  const currentGroupIdx = sortedGroups.findIndex(g => g.id === group.id);
  const prevGroup = currentGroupIdx > 0 ? sortedGroups[currentGroupIdx - 1] : null;
  const nextGroup = (currentGroupIdx >= 0 && currentGroupIdx < sortedGroups.length - 1) ? sortedGroups[currentGroupIdx + 1] : null;

  const openGroupNumberInput = () => {
    if (!onSwitchGroup) return;
    setGroupNumInput(extractGroupNumber(group.name) || "");
    setEditingGroupNum(true);
  };

  const confirmGroupNumberInput = () => {
    const target = sortedGroups.find(g => extractGroupNumber(g.name) === groupNumInput.trim());
    if (target) {
      onSwitchGroup(target, currentSlot);
    } else if (groupNumInput.trim() !== "") {
      window.alert(`No group numbered ${groupNumInput.trim()}`);
    }
    setEditingGroupNum(false);
  };

  const updateDelay = (val) => {
    const n = Math.max(0, Number(val) || 0);
    setDelayMin(n);
    onUpdate({ holeData, records, currentHole: currentSlot, actionLogs, delayMin: n });
  };

  // ─── Suspension (global, received from App via props) ─────────────────────────────
  // The schedule received from App is already adjusted, so it can be used directly
  const adjustedSchedule = schedule;

  // WN / MN / TM logs: { holeIdx, type, name, time, target? }
  const [actionLogs, setActionLogs] = useState(group.actionLogs ?? []);
  const [actionModal, setActionModal] = useState(null); // { type: "WN"|"MN"|"TM", holeIdx }
  const [actionName, setActionName] = useState("");
  const [mnActive, setMnActive] = useState(group.mnActive ?? false);
  const [mnName, setMnName] = useState(group.mnName ?? "");
  // TM (Timing) — like MN but can specify which player(s) are being timed, or the whole group (2-4 players)
  const [tmActive, setTmActive] = useState(group.tmActive ?? false);
  const [tmName, setTmName] = useState(group.tmName ?? "");
  const [tmTarget, setTmTarget] = useState(group.tmTarget ?? "");
  const [actionTargets, setActionTargets] = useState([]); // ["ALL"] or [1,2,3,4] for TM

  const openActionModal = (type, holeIdx) => {
    setActionModal({ type, holeIdx });
    setActionName(currentUser || "");
    setActionTargets([]);
  };

  const toggleTarget = (t) => {
    if (t === "ALL") { setActionTargets(["ALL"]); return; }
    setActionTargets(prev => {
      const withoutAll = prev.filter(x => x !== "ALL");
      const next = withoutAll.includes(t) ? withoutAll.filter(x => x !== t) : [...withoutAll, t];
      return next; // allow empty — no forced fallback to "ALL"
    });
  };

  const targetLabel = (targets) => {
    if (!targets || targets.length === 0) return "";
    if (targets.includes("ALL")) return "All";
    return targets.slice().sort((a, b) => a - b).map(n => `P${n}`).join(", ");
  };

  const confirmAction = () => {
    if (!actionModal) return;
    if (actionModal.type === "TM" && actionTargets.length === 0) return; // must pick a player or "All" first
    const holeIdx = actionModal.holeIdx;
    const deadline = (adjustedSchedule[holeIdx] ?? 0) + (parTimes?.[holeIdx] ?? 14);
    const diffAtLog = nowInMin() - deadline + 1;
    const target = actionModal.type === "TM" ? targetLabel(actionTargets) : undefined;
    const log = {
      holeIdx,
      type: actionModal.type,
      name: actionName.trim() || currentUser || "—",
      time: minToTime(nowInMin()),
      diff: diffAtLog,
      ...(target ? { target } : {}),
    };
    let next = [...actionLogs, log];
    let newMnActive = mnActive;
    let newMnName = mnName;
    let newTmActive = tmActive;
    let newTmName = tmName;
    let newTmTarget = tmTarget;
    if (actionModal.type === "MN") {
      newMnActive = true;
      newMnName = log.name;
      setMnActive(true);
      setMnName(log.name);
    }
    if (actionModal.type === "TM") {
      newTmActive = true;
      newTmName = log.name;
      newTmTarget = target;
      setTmActive(true);
      setTmName(log.name);
      setTmTarget(target);
      // TM on "All" replaces MN (turns it off at the previous hole, since TM(All)
      // starting at hole X means the group was still under plain MN through hole X-1).
      // TM on a specific player runs alongside MN instead — the group can be both
      // under general monitoring and have one player specifically timed at once.
      if (mnActive && target === "All") {
        newMnActive = false;
        newMnName = "";
        setMnActive(false);
        setMnName("");
        next = [...next, { holeIdx: Math.max(0, holeIdx - 1), type: "MN", name: log.name, time: log.time, off: true }];
      }
    }
    setActionLogs(next);
    onUpdate({
      holeData, records, currentHole: currentSlot, actionLogs: next,
      mnActive: newMnActive, mnName: newMnName,
      tmActive: newTmActive, tmName: newTmName, tmTarget: newTmTarget,
    });
    setActionModal(null);
  };

  // "Bad Time" — quick action while MN/TM is active: flags a specific player (P1-P4) and
  // immediately shows the TM status right on the CURRENT hole (the log is stamped with
  // type "TM" + badTime:true at holeIdx = currentHole), then keeps following forward on
  // every hole after that too, without switching off the group's MN status (MN and TM
  // can run at the same time here). A player can be Bad-Timed more than once — every
  // press logs a fresh occurrence (see badTimeOccurrence below for the "#N" count,
  // which is derived from the current log list so it re-numbers correctly if an entry
  // is later deleted).
  const triggerBadTimeFor = (playerNum) => {
    const playerLabel = `P${playerNum}`;
    const deadline = (adjustedSchedule[currentHole] ?? 0) + (parTimes?.[currentHole] ?? 14);
    const diffAtLog = nowInMin() - deadline + 1;

    const alreadyAll = tmActive && tmTarget === "All";
    const existingTargets = tmActive && tmTarget && tmTarget !== "All"
      ? tmTarget.split(",").map(s => s.trim()).filter(Boolean)
      : [];
    const mergedTargets = alreadyAll ? [] : Array.from(new Set([...existingTargets, playerLabel]));
    // If Bad Time has now been pressed for every player in the group one by one,
    // treat it the same as TM(All) — collapse the target to "All" and let it take
    // over from MN, same as picking "All" directly in the TM picker would.
    const coversAll = alreadyAll || mergedTargets.length >= numPlayers;
    const newTarget = coversAll ? "All" : mergedTargets.join(", ");

    const log = {
      holeIdx: currentHole, // Bad Time shows on the current hole itself
      type: "TM",
      name: currentUser || "—",
      time: minToTime(nowInMin()),
      diff: diffAtLog,
      target: playerLabel, // this log records just the player flagged by this Bad Time press
      badTime: true,
    };
    let next = [...actionLogs, log];
    let nextMnActive = mnActive;
    let nextMnName = mnName;
    if (mnActive && coversAll && !alreadyAll) {
      // Every player is now covered by TM — MN is redundant from here on. Off it,
      // recorded at the previous hole (same convention as picking TM "All" directly).
      nextMnActive = false;
      nextMnName = "";
      next = [...next, { holeIdx: Math.max(0, currentHole - 1), type: "MN", name: currentUser || mnName, time: log.time, off: true }];
      setMnActive(false);
      setMnName("");
    }
    setActionLogs(next);
    setTmActive(true);
    setTmName(currentUser || tmName);
    setTmTarget(newTarget);
    onUpdate({
      holeData, records, currentHole: currentSlot, actionLogs: next,
      mnActive: nextMnActive, mnName: nextMnName,
      tmActive: true, tmName: currentUser || tmName, tmTarget: newTarget,
    });
  };

  const offMN = () => {
    const offLog = { holeIdx: currentHole, type: "MN", name: mnName, time: minToTime(nowInMin()), off: true };
    const nextLogs = [...actionLogs, offLog];
    setActionLogs(nextLogs);
    setMnActive(false);
    setMnName("");
    onUpdate({ holeData, records, currentHole: currentSlot, actionLogs: nextLogs, mnActive: false, mnName: "", tmActive, tmName, tmTarget });
  };

  const offTM = () => {
    const offLog = { holeIdx: currentHole, type: "TM", name: tmName, time: minToTime(nowInMin()), off: true };
    const nextLogs = [...actionLogs, offLog];
    setActionLogs(nextLogs);
    setTmActive(false);
    setTmName("");
    setTmTarget("");
    onUpdate({ holeData, records, currentHole: currentSlot, actionLogs: nextLogs, mnActive, mnName, tmActive: false, tmName: "", tmTarget: "" });
  };

  // Delete a WN/MN/TM log entry (in case of a mistaken tap) — confirmed via popup before removal
  const [deleteLogConfirm, setDeleteLogConfirm] = useState(null); // index into actionLogs, or null
  const deleteLogAt = (idx) => {
    const next = actionLogs.filter((_, i2) => i2 !== idx);
    setActionLogs(next);
    onUpdate({ holeData, records, currentHole: currentSlot, actionLogs: next, mnActive, mnName, tmActive, tmName, tmTarget });
    setDeleteLogConfirm(null);
  };

  useEffect(() => {
    const iv = setInterval(() => setNow(nowInMin()), 1000);
    return () => clearInterval(iv);
  }, []);

  const diffAtHole = (holeIdx) => {
    const hd = holeData[holeIdx];
    if (!hd || !hd.endTime) return null;
    if (hd.manualDiff !== undefined) return hd.manualDiff;
    const [eh, em] = hd.endTime.split(":").map(Number);
    const endMin = eh * 60 + em;
    const deadline = (adjustedSchedule[holeIdx] ?? 0) + (parTimes?.[holeIdx] ?? 14);
    return endMin - deadline + 1;
  };

  const commitRecord = (holeIdx, newHoleData) => {
    const nxtHD = [...holeData];
    nxtHD[holeIdx] = newHoleData;
    setHoleData(nxtHD);

    const nxtRec = [...records];
    nxtRec[holeIdx] = newHoleData.endTime;
    setRecords(nxtRec);

    onUpdate({ holeData: nxtHD, records: nxtRec, currentHole: holeIdx, actionLogs, mnActive, mnName, tmActive, tmName, tmTarget });
    return { nxtHD, nxtRec };
  };

  // Clears a mistakenly-recorded finish time for a hole (e.g. typo'd via manual
  // entry, or timestamped by accident) so it can be re-entered from scratch.
  const clearHoleRecord = (holeIdx) => {
    const nxtHD = [...holeData];
    nxtHD[holeIdx] = { ...nxtHD[holeIdx], endTime: null, manualDiff: undefined };
    setHoleData(nxtHD);

    const nxtRec = [...records];
    nxtRec[holeIdx] = null;
    setRecords(nxtRec);

    setRecordedEnd(null);
    setDiffManual(0);
    onUpdate({ holeData: nxtHD, records: nxtRec, currentHole: holeIdx, actionLogs, mnActive, mnName, tmActive, tmName, tmTarget });
  };

  // Shared auto-log logic: while MN/TM is active, whenever a hole becomes "complete"
  // (whether via the normal Confirm-Hole flow or a manual time edit), stamp an
  // automatic MN/TM entry for that hole so the status keeps following forward,
  // the same way for both flows.
  const autoLogMonitoring = (holeIdx, logsBase) => {
    let nxt = logsBase;
    if (mnActive) {
      const alreadyMN = nxt.some(l => l.type === "MN" && l.holeIdx === holeIdx);
      if (!alreadyMN) {
        const autoDeadline = (adjustedSchedule[holeIdx] ?? 0) + (parTimes?.[holeIdx] ?? 14);
        const autoDiff = nowInMin() - autoDeadline + 1;
        nxt = [...nxt, { holeIdx, type: "MN", name: currentUser || mnName, time: minToTime(nowInMin()), diff: autoDiff, auto: true }];
      }
    }
    if (tmActive) {
      const alreadyTM = nxt.some(l => l.type === "TM" && l.holeIdx === holeIdx);
      if (!alreadyTM) {
        const autoDeadline = (adjustedSchedule[holeIdx] ?? 0) + (parTimes?.[holeIdx] ?? 14);
        const autoDiff = nowInMin() - autoDeadline + 1;
        nxt = [...nxt, { holeIdx, type: "TM", name: currentUser || tmName, time: minToTime(nowInMin()), diff: autoDiff, target: tmTarget, auto: true }];
      }
    }
    return nxt;
  };

  const markHole = () => {
    const endAbsMin = inputMode === "stamp"
      ? (recordedEnd ?? now)
      : deadlineMin - 1;
    const st = minToTime(startAbsMin);
    const et = minToTime(endAbsMin);
    const extraFields = inputMode === "manual" ? { manualDiff: diffManual } : {};
    const { nxtRec, nxtHD } = commitRecord(currentHole, { startTime: st, endTime: et, ...extraFields });

    // auto-log MN / TM if currently monitoring/timing
    let nxtLogs = autoLogMonitoring(currentHole, actionLogs);
    if (nxtLogs !== actionLogs) setActionLogs(nxtLogs);

    // Keep mnName/tmName synced to whoever is actually recording right now — otherwise
    // the "upcoming hole" preview badge (shown before that hole is recorded) keeps
    // showing the name of whoever first started MN/TM, even after someone else takes over.
    const nextMnName = mnActive ? (currentUser || mnName) : mnName;
    const nextTmName = tmActive ? (currentUser || tmName) : tmName;
    if (nextMnName !== mnName) setMnName(nextMnName);
    if (nextTmName !== tmName) setTmName(nextTmName);

    if (currentSlot < 17) {
      const next = currentSlot + 1;
      setCurrentSlot(next);
      setRecordedEnd(null);
      setDiffManual(0);
      onUpdate({ holeData: nxtHD, records: nxtRec, currentHole: next, actionLogs: nxtLogs, mnActive, mnName: nextMnName, tmActive, tmName: nextTmName, tmTarget });
    } else {
      setCurrentSlot(18);
      onUpdate({ holeData: nxtHD, records: nxtRec, currentHole: 18, actionLogs: nxtLogs, mnActive, mnName: nextMnName, tmActive, tmName: nextTmName, tmTarget });
    }
    onRecorded?.();
  };

  const startEditHole = (i, field) => {
    const hd = holeData[i];
    setEditingHole(i);
    setEditField(field);
    setEditVal((field === "start" ? hd?.startTime : hd?.endTime) || minToTime(now));
  };

  const confirmEditHole = (i) => {
    const existing = holeData[i] || { startTime: null, endTime: null };
    const updated = { ...existing, [editField === "start" ? "startTime" : "endTime"]: editVal };
    const { nxtHD, nxtRec } = commitRecord(i, updated);

    // If this edit completes the hole (both start & end now set), auto-log MN / TM
    // just like the normal Confirm-Hole flow does, so the status keeps following forward
    // even when times are entered/fixed manually instead of via the Confirm button.
    if (updated.startTime && updated.endTime) {
      const nxtLogs = autoLogMonitoring(i, actionLogs);
      if (nxtLogs !== actionLogs) {
        setActionLogs(nxtLogs);
        onUpdate({ holeData: nxtHD, records: nxtRec, currentHole: currentSlot, actionLogs: nxtLogs, mnActive, mnName, tmActive, tmName, tmTarget });
      }
    }
    setEditingHole(null);
  };

  const clearHole = (i) => {
    const nxtHD = [...holeData];
    nxtHD[i] = { startTime: null, endTime: null };
    setHoleData(nxtHD);
    const nxtRec = [...records];
    nxtRec[i] = null;
    setRecords(nxtRec);
    onUpdate({ holeData: nxtHD, records: nxtRec, currentHole: currentSlot });
    setEditingHole(null);
  };

  const overallStatus = () => {
    const diffs = holeOrder.map(hi => holeData[hi]?.endTime ? diffAtHole(hi) : null).filter(d => d !== null);
    if (!diffs.length) return "idle";
    return getStatus(diffs[diffs.length - 1]);
  };

  // The Dashboard rates a group relative to the group ahead of it; when it opens this
  // screen it passes that same rating in so the badge here can't disagree with the
  // badge on the card/table the user just tapped.
  const status = statusOverride ?? overallStatus();
  const bgColor = { ok: "#0a1f15", warn: "#1f180a", late: "#1f0a0a", idle: "#0d0f1a" }[status];
  const done = currentSlot >= 18;

  // Slot (play-order position) where the MN/TM "coming up" preview should be shown:
  // exactly one hole past the most recently logged MN/TM entry, so the preview follows
  // forward one hole at a time (instead of showing on every future hole at once). Once
  // MN/TM is turned off, mnActive/tmActive become false and the preview disappears everywhere.
  const slotOfHole = {};
  holeOrder.forEach((hIdx, s) => { slotOfHole[hIdx] = s; });
  const lastMNSlot = actionLogs.reduce((mx, l) => (l.type === "MN" ? Math.max(mx, slotOfHole[l.holeIdx] ?? -1) : mx), -1);
  const lastTMSlot = actionLogs.reduce((mx, l) => (l.type === "TM" ? Math.max(mx, slotOfHole[l.holeIdx] ?? -1) : mx), -1);

  // Players who have been flagged via a Bad Time press (as opposed to just being part of a
  // normal TM selection) — used so the P1–P4 quick-select buttons can show a visibly different
  // state for "Bad-Timed" vs merely "under TM", since otherwise pressing Bad Time on a player
  // who's already in the TM target list looks like nothing happened (the button stayed the
  // same pink "already selected" color).
  const badTimePlayers = new Set(actionLogs.filter(l => l.type === "TM" && l.badTime).map(l => l.target));
  const badTimeCounts = actionLogs.filter(l => l.type === "TM" && l.badTime).reduce((m, l) => {
    m[l.target] = (m[l.target] || 0) + 1;
    return m;
  }, {});

  // "#N" — which Bad Time occurrence this is for a given player, e.g. their 1st, 2nd,
  // 3rd time. Recomputed fresh from the current actionLogs on every render (in chronological/
  // creation order), so if an earlier Bad Time entry is deleted the remaining ones simply
  // shift down and renumber — there's no separately stored counter to get out of sync.
  const badTimeOccurrence = new Map();
  {
    const counters = {};
    actionLogs.forEach(l => {
      if (l.type === "TM" && l.badTime) {
        counters[l.target] = (counters[l.target] || 0) + 1;
        badTimeOccurrence.set(l, counters[l.target]);
      }
    });
  }

  const parTimeNow = parTimes?.[currentHole] ?? 14;
  // Always anchor Start/Finish to the fixed master schedule for this hole — a group
  // running late (or early) on a previous hole must NOT shift this hole's target
  // time. Each hole's deadline stays fixed to the original tee sheet, so recorded
  // diffs measure cumulative schedule adherence rather than resetting hole-to-hole.
  const startAbsMin = adjustedSchedule[currentHole];
  const deadlineMin = startAbsMin + parTimeNow;
  const diffLive = now - deadlineMin + 1;
  const displayEnd = recordedEnd ?? now;
  const diffDisplay = displayEnd - deadlineMin + 1;
  const canConfirm = inputMode === "manual" || !!recordedEnd;

  const diffColor = (d) => d >= 3 ? "#ff7070" : d >= 1 ? "#ffd966" : "#6effa0";

  return (
    <div style={compact
      ? { fontFamily: "'IBM Plex Mono', monospace", color: "#eee" }
      : { background: bgColor, minHeight: "100vh", fontFamily: "'IBM Plex Mono', monospace", color: "#eee", transition: "background 1s" }}>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&family=Bebas+Neue&display=swap" rel="stylesheet" />

      {!compact && (
      <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 20px", background: "#141626cc", borderBottom: "1px solid #2a2d4a", backdropFilter: "blur(8px)" }}>
        <button onClick={onBack} style={{ background: "#1a1d2e", border: "1px solid #4e9af144", color: "#4e9af1", cursor: "pointer", fontSize: closeLabel ? 15 : 26, fontWeight: 700, borderRadius: 8, width: 44, height: 44, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{closeLabel || "←"}</button>

        {onSwitchGroup && (
          <button
            onClick={() => prevGroup && onSwitchGroup(prevGroup, currentSlot)}
            disabled={!prevGroup}
            title="Previous group"
            style={{ background: "#1a1d2e", border: "1px solid #4e9af144", color: prevGroup ? "#4e9af1" : "#333", cursor: prevGroup ? "pointer" : "not-allowed", fontSize: 20, fontWeight: 700, borderRadius: 8, width: 38, height: 38, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}
          >‹</button>
        )}

        {editingGroupNum ? (
          <input
            autoFocus
            type="number"
            value={groupNumInput}
            onChange={e => setGroupNumInput(e.target.value)}
            onKeyDown={e => { if (e.key === "Enter") confirmGroupNumberInput(); if (e.key === "Escape") setEditingGroupNum(false); }}
            onBlur={confirmGroupNumberInput}
            placeholder="Group no."
            style={{ width: 80, fontFamily: "'Bebas Neue'", fontSize: 18, letterSpacing: 2, background: "#0d0f1a", border: "1px solid #4e9af1", color: "#eee", borderRadius: 6, padding: "4px 8px", flexShrink: 0 }}
          />
        ) : (
          <div
            onClick={openGroupNumberInput}
            title={onSwitchGroup ? "Tap to type a group number" : undefined}
            style={{ fontFamily: "'Bebas Neue'", fontSize: 18, letterSpacing: 3, flexShrink: 0, cursor: onSwitchGroup ? "pointer" : "default", borderBottom: onSwitchGroup ? "1px dashed #4e9af166" : "none" }}
          >{group.name}</div>
        )}

        {onSwitchGroup && (
          <button
            onClick={() => nextGroup && onSwitchGroup(nextGroup, currentSlot)}
            disabled={!nextGroup}
            title="Next group"
            style={{ background: "#1a1d2e", border: "1px solid #4e9af144", color: nextGroup ? "#4e9af1" : "#333", cursor: nextGroup ? "pointer" : "not-allowed", fontSize: 20, fontWeight: 700, borderRadius: 8, width: 38, height: 38, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}
          >›</button>
        )}

        {/* Right column — 2 rows */}
        <div style={{ marginLeft: "auto", display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 4 }}>
          {/* Row 1: user + logout */}
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            {currentUser && <span style={{ fontSize: 12, color: "#8899cc" }}>👤 {currentUser}</span>}
            <LogoutButton onLogout={onLogout} />
          </div>
          {/* Row 2: hole info + status */}
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 11, color: "#8890b8" }}>H{group.startHole || 1} · Tee time {group.startTime}</span>
            <StatusBadge status={status} />
          </div>
        </div>
      </div>
      )}

      {/* Global Suspension Banner */}
      {isSuspended && (
        <div style={{
          background: "#1f0f00", borderBottom: "1px solid #ffd96666",
          padding: "8px 20px", display: "flex", alignItems: "center", gap: 10,
        }}>
          <div style={{ width: 7, height: 7, borderRadius: "50%", background: "#ffd966", boxShadow: "0 0 8px #ffd966", animation: "pulse 1.5s infinite" }} />
          <span style={{ fontSize: 13, color: "#ffd966", fontWeight: 700 }}>⏸ Match paused</span>
          <span style={{ fontSize: 12, color: "#aaa" }}>Since <b style={{ color: "#eee" }}>{pendingStopTime}</b></span>
          <span style={{ fontSize: 12, color: "#8890b8" }}>— Go back to the main screen to resume play</span>
        </div>
      )}

      <div style={compact ? { padding: 0 } : { maxWidth: 700, margin: "0 auto", padding: "20px 16px" }}>

        {/* Suspension History */}
        {suspensions && suspensions.length > 0 && (
          <div style={{
            background: "#141626", border: "1px solid #ffd96633",
            borderRadius: 10, padding: "10px 16px", marginBottom: 16,
          }}>
            <div style={{ fontSize: 11, color: "#ffd966", letterSpacing: 2, marginBottom: 8, fontWeight: 700 }}>⏱ Pause history</div>
            {suspensions.map((s, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: "#aaa", marginBottom: 4 }}>
                <span style={{ color: "#ffd966" }}>#{i + 1}</span>
                <span>Stopped <b style={{ color: "#eee" }}>{s.stopTime}</b></span>
                <span>→ Resumed <b style={{ color: "#6effa0" }}>{s.resumeTime}</b></span>
                <span style={{ marginLeft: "auto", color: "#ff9966" }}>+{s.offsetMin} min</span>
              </div>
            ))}
            {totalOffsetMin > 0 && (
              <div style={{ borderTop: "1px solid #2a2d4a", marginTop: 8, paddingTop: 8, fontSize: 12, color: "#ff9966", fontWeight: 700 }}>
                Total time shift +{totalOffsetMin} min
              </div>
            )}
          </div>
        )}

        {!done && (
        <div style={{
          position: "relative",
          background: "#141626",
          border: `1px solid ${group.color}44`,
          borderRadius: 14,
          padding: compact ? 16 : 24,
          marginBottom: compact ? 0 : 20,
          boxShadow: `0 0 40px ${group.color}11`,
        }}>
          {compact && (
            <button onClick={onBack} style={{ position: "absolute", top: 14, right: 14, background: "#1a1d2e", border: "1px solid #4e9af144", color: "#4e9af1", cursor: "pointer", fontSize: 15, fontWeight: 700, borderRadius: 8, width: 32, height: 32, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, zIndex: 1 }}>{closeLabel || "✕"}</button>
          )}

          <div style={{ display: "flex", alignItems: compact ? "flex-start" : "center", justifyContent: "flex-start", gap: compact ? 28 : 16, marginBottom: compact ? 10 : 14, paddingTop: compact ? 6 : 0 }}>
            {compact ? (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <div style={{ width: 10, height: 10, borderRadius: "50%", background: group.color, flexShrink: 0 }} />
                  <div style={{ fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 2, flexShrink: 0 }}>{group.name}</div>
                </div>
                <StatusBadge status={status} />
              </div>
            ) : (
              <div style={{ fontSize: 12, color: "#9aa2c7", letterSpacing: 2 }}>CURRENT HOLE</div>
            )}

            <div style={{ display: "flex", flexDirection: "column", alignItems: "stretch" }}>
            {compact && (
              <div style={{ fontSize: 12, color: "#9aa2c7", letterSpacing: 2, marginBottom: 6, width: "100%" }}>CURRENT HOLE</div>
            )}
            <div style={{ display: "flex", alignItems: "center", gap: compact ? 6 : 16 }}>
            <button
              onClick={() => setCurrentSlot(Math.max(0, currentSlot - 1))}
              disabled={currentSlot === 0}
              title="Previous hole"
              style={{ background: "#0d0f1a", border: `1px solid ${group.color}44`, color: currentSlot === 0 ? "#333" : group.color, cursor: currentSlot === 0 ? "not-allowed" : "pointer", fontSize: compact ? 15 : 22, fontWeight: 700, borderRadius: 10, width: compact ? 28 : 40, height: compact ? 28 : 40, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}
            >‹</button>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: compact ? 30 : 72, lineHeight: 1, color: group.color }}>
              {currentHole + 1}
            </div>
            <button
              onClick={() => setCurrentSlot(Math.min(17, currentSlot + 1))}
              disabled={currentSlot === 17}
              title="Next hole"
              style={{ background: "#0d0f1a", border: `1px solid ${group.color}44`, color: currentSlot === 17 ? "#333" : group.color, cursor: currentSlot === 17 ? "not-allowed" : "pointer", fontSize: compact ? 15 : 22, fontWeight: 700, borderRadius: 10, width: compact ? 28 : 40, height: compact ? 28 : 40, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}
            >›</button>
            {compact && (
              <div style={{ marginLeft: 6 }}>
                <div style={{ color: "#888", fontSize: 13 }}>Par {pars[currentHole]}</div>
                <div style={{ color: "#8890b8", fontSize: 11 }}>Slot {currentSlot + 1}/18</div>
              </div>
            )}
            </div>
            {!compact && (
              <>
                <div style={{ color: "#888", fontSize: 13, textAlign: "left", marginTop: 2 }}>Par {pars[currentHole]}</div>
                <div style={{ color: "#8890b8", fontSize: 11, textAlign: "left" }}>Slot {currentSlot + 1}/18</div>
              </>
            )}
          </div>
          </div>

          {holeData[currentHole]?.endTime && (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, background: "#2a1a0a", border: "1px solid #ffd96644", borderRadius: 10, padding: "8px 12px", marginBottom: compact ? 10 : 14 }}>
              <span style={{ fontSize: 12, color: "#ffd966" }}>Finish time already recorded for this hole ({holeData[currentHole].endTime})</span>
              <button
                onClick={() => { if (window.confirm(`Clear the recorded finish time for H${currentHole + 1}?\n\nYou can record it again afterwards.`)) clearHoleRecord(currentHole); }}
                style={{ background: "#2a0a0a", border: "1px solid #ff7070", color: "#ff7070", borderRadius: 7, padding: "5px 10px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700, flexShrink: 0 }}
              >🗑 Clear</button>
            </div>
          )}

          <div style={{ display: "flex", alignItems: "center", gap: compact ? 10 : 16, marginBottom: compact ? 12 : 20 }}>
            <div style={{ flex: 1, background: "#0d0f1a", borderRadius: 12, padding: compact ? "8px 10px" : "12px 14px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                <div style={{ textAlign: "center" }}>
                  <div style={{ fontSize: 11, color: "#8890b8", marginBottom: 2 }}>🏌️ Start</div>
                  <div style={{ fontFamily: "'Bebas Neue'", fontSize: compact ? 19 : 24, color: "#4e9af1", lineHeight: 1 }}>{minToTime(startAbsMin)}</div>
                </div>
                <div style={{ flex: 1, textAlign: "center" }}>
                  <div style={{ fontSize: 10, color: "#8890b8", marginBottom: 2 }}>Par Time</div>
                  <div style={{ color: "#666f99", fontSize: compact ? 14 : 18 }}>+{parTimeNow}m→</div>
                </div>
                <div style={{ textAlign: "center" }}>
                  <div style={{ fontSize: 11, color: "#8890b8", marginBottom: 2 }}>🏁 Finish</div>
                  <div style={{ fontFamily: "'Bebas Neue'", fontSize: compact ? 19 : 24, color: "#8899cc", lineHeight: 1 }}>{minToTime(deadlineMin)}</div>
                </div>
              </div>
            </div>
            <div style={{ textAlign: "center", background: "#0d0f1a", borderRadius: 12, padding: compact ? "6px 8px" : "10px 12px", minWidth: compact ? 64 : 80 }}>
              <div style={{ fontSize: 11, color: "#8890b8", marginBottom: 2 }}>Now</div>
              <div style={{ fontFamily: "'Bebas Neue'", fontSize: compact ? 19 : 24, lineHeight: 1, color: "#eee" }}>{minToTime(now)}</div>
              <div style={{ fontSize: 12, fontWeight: 700, marginTop: 4, color: diffColor(diffLive) }}>
                {diffLive > 0 ? `+${diffLive}` : diffLive} min
              </div>
            </div>
          </div>

          {/* Delay input */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, background: "#0d0f1a", borderRadius: 10, padding: compact ? "8px 12px" : "10px 14px", marginBottom: compact ? 10 : 14 }}>
            <span style={{ fontSize: 12, color: "#ff7070" }}>⏱ Delay Time</span>
            <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <button onClick={() => updateDelay(Math.max(0, delayMin - 1))} style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#aaa", borderRadius: 6, width: compact ? 32 : 40, height: compact ? 32 : 40, cursor: "pointer", fontSize: 16, fontFamily: "inherit" }}>−</button>
              <input
                type="number" min="0" value={delayMin}
                onChange={e => updateDelay(e.target.value)}
                style={{ width: 52, background: "#1e2135", border: `1px solid ${delayMin > 0 ? "#ffd96666" : "#2a2d4a"}`, color: "#ffd966", fontFamily: "'Bebas Neue'", fontSize: 22, textAlign: "center", borderRadius: 6, padding: "2px 0", outline: "none" }}
              />
              <button onClick={() => updateDelay(delayMin + 1)} style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#aaa", borderRadius: 6, width: compact ? 32 : 40, height: compact ? 32 : 40, cursor: "pointer", fontSize: 16, fontFamily: "inherit" }}>+</button>
            </div>
            <span style={{ fontSize: 12, color: "#8890b8" }}>min</span>
            {delayMin > 0 && (
              <span style={{ fontSize: 11, color: "#ffd966", background: "#2a1a00", border: "1px solid #ffd96644", borderRadius: 5, padding: "2px 7px" }}>
                schedule +{delayMin}m
              </span>
            )}
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: compact ? 10 : 14 }}>
            <span style={{ fontSize: 12, color: "#8890b8", fontWeight: 700 }}>Flag:</span>
            {(() => {
              const wnDisabled = (actionLogs.some(l => l.type === "WN" && l.holeIdx === currentHole)) || mnActive || tmActive;
              const mnDisabledHere = actionLogs.some(l => l.type === "MN" && !l.badTime && !l.off && l.holeIdx === currentHole);
              const mnAlreadyUsed = actionLogs.some(l => l.type === "MN" && l.off);
              const mnDisabled = mnActive || mnDisabledHere || mnAlreadyUsed;
              return (
                <>
                  <button onClick={() => !wnDisabled && openActionModal("WN", currentHole)}
                    disabled={wnDisabled}
                    title={mnActive || tmActive ? "MN/TM is active — WN cannot be used right now" : wnDisabled ? "WN already logged on this hole" : undefined}
                    style={{ flex: 1, background: wnDisabled ? "#1a1a1a" : "#2a1a00", border: `1px solid ${wnDisabled ? "#3a3a3a" : "#ffd96688"}`, color: wnDisabled ? "#666" : "#ffd966", borderRadius: 8, padding: compact ? "7px 0" : "10px 0", cursor: wnDisabled ? "not-allowed" : "pointer", fontSize: 14, fontWeight: 700, fontFamily: "inherit" }}>WN</button>
                  <button onClick={() => !mnDisabled && openActionModal("MN", currentHole)}
                    disabled={mnDisabled}
                    title={mnActive ? "MN is already running — turn MN off before starting again" : mnDisabledHere ? "MN already logged on this hole" : mnAlreadyUsed ? "This group has already had MN turned on and off — MN can only be used once per group" : undefined}
                    style={{ flex: 1, background: mnDisabled ? "#1a1a1a" : "#001a2a", border: `1px solid ${mnDisabled ? "#3a3a3a" : "#4e9af188"}`, color: mnDisabled ? "#666" : "#4e9af1", borderRadius: 8, padding: compact ? "7px 0" : "10px 0", cursor: mnDisabled ? "not-allowed" : "pointer", fontSize: 14, fontWeight: 700, fontFamily: "inherit" }}>MN</button>
                  <button onClick={() => openActionModal("TM", currentHole)}
                    style={{ flex: 1, background: "#2a0020", border: "1px solid #ff6ec788", color: "#ff6ec7", borderRadius: 8, padding: compact ? "7px 0" : "10px 0", cursor: "pointer", fontSize: 14, fontWeight: 700, fontFamily: "inherit" }}>TM</button>
                </>
              );
            })()}
          </div>

          <div style={{ display: "flex", gap: 8, marginBottom: compact ? 10 : 14 }}>
            {[["stamp", "🏁 Timestamp now"], ["manual", "✏️ Enter difference manually"]].map(([mode, label]) => (
              <button key={mode} onClick={() => { setInputMode(mode); setRecordedEnd(null); }}
                style={{
                  flex: 1, padding: compact ? "7px 0" : "8px 0", borderRadius: 8, cursor: "pointer",
                  fontFamily: "inherit", fontSize: compact ? 12 : 13, fontWeight: 700,
                  background: inputMode === mode ? "#1e2a3a" : "#0d0f1a",
                  border: `1px solid ${inputMode === mode ? "#4e9af1" : "#2a2d4a"}`,
                  color: inputMode === mode ? "#4e9af1" : "#555",
                  transition: "all 0.15s",
                }}>
                {label}
              </button>
            ))}
          </div>

          {inputMode === "stamp" && (
            !recordedEnd ? (
              <button
                onClick={() => setRecordedEnd(nowInMin())}
                style={{
                  width: "100%", padding: compact ? "16px 0" : "22px 0", marginBottom: 12,
                  background: "linear-gradient(135deg, #1a2a1a, #1f3f1f)",
                  border: "2px solid #6effa066", borderRadius: 14, cursor: "pointer",
                  fontFamily: "'Bebas Neue'", fontSize: compact ? 22 : 28, letterSpacing: 4, color: "#6effa0",
                  transition: "all 0.15s",
                }}
                onMouseEnter={e => { e.currentTarget.style.background = "linear-gradient(135deg, #1f3f1f, #2a5a2a)"; e.currentTarget.style.borderColor = "#6effa0aa"; }}
                onMouseLeave={e => { e.currentTarget.style.background = "linear-gradient(135deg, #1a2a1a, #1f3f1f)"; e.currentTarget.style.borderColor = "#6effa066"; }}
              >
                🏁 Record holed time
              </button>
            ) : (
              <div style={{ background: "#0d0f1a", borderRadius: 14, padding: "16px 20px", marginBottom: 12 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 11, color: "#8890b8", marginBottom: 4 }}>🏁 Recorded finish time</div>
                    <div style={{ fontFamily: "'Bebas Neue'", fontSize: 36, color: "#eee", lineHeight: 1 }}>{minToTime(recordedEnd)}</div>
                    <div style={{ fontSize: 14, fontWeight: 700, marginTop: 6, color: diffColor(diffDisplay) }}>
                      {diffDisplay === 0 ? "On time" : diffDisplay > 0 ? `${diffDisplay} min late` : `${Math.abs(diffDisplay)} min early`}
                    </div>
                  </div>
                  <button onClick={() => setRecordedEnd(null)}
                    style={{ background: "#2a1a1a", border: "1px solid #ff707044", color: "#ff7070", borderRadius: 8, padding: "8px 14px", cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>
                    ↺ Cancel
                  </button>
                </div>
              </div>
            )
          )}

          {inputMode === "manual" && (
            <div style={{ background: "#0d0f1a", borderRadius: 12, padding: "16px 14px", marginBottom: 14 }}>
              <div style={{ fontSize: 12, color: "#aaa", letterSpacing: 1, marginBottom: 14 }}>Difference from scheduled (min)</div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <div style={{ display: "flex", gap: 4 }}>
                  <button onClick={() => setDiffManual(d => d - 5)}
                    style={{ width: 44, height: 52, borderRadius: 10, background: "#2a1010", border: "1px solid #ff707044", color: "#ff7070", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>−5</button>
                  <button onClick={() => setDiffManual(d => d - 1)}
                    style={{ width: 44, height: 52, borderRadius: 10, background: "#1e1515", border: "1px solid #ff707033", color: "#ff9090", fontSize: 18, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>−</button>
                </div>
                <div style={{ flex: 1, textAlign: "center" }}>
                  <input type="text" inputMode="numeric" pattern="-?[0-9]*"
                    value={diffManual > 0 ? `+${diffManual}` : diffManual}
                    onChange={e => {
                      const v = e.target.value.replace(/[^0-9-]/g, "");
                      setDiffManual(v === "" || v === "-" ? 0 : Number(v));
                    }}
                    style={{
                      width: 90, background: "#141626",
                      border: `2px solid ${diffManual >= 3 ? "#ff707066" : diffManual >= 1 ? "#ffd96666" : "#6effa066"}`,
                      color: diffColor(diffManual),
                      borderRadius: 12, padding: "6px 4px",
                      fontFamily: "'Bebas Neue'", fontSize: 44, textAlign: "center", outline: "none",
                    }}
                  />
                  <div style={{ fontSize: 12, color: "#8890b8", marginTop: 4 }}>
                    {diffManual === 0 ? `On time → finish ${minToTime(deadlineMin - 1)}` : diffManual > 0 ? `${diffManual} min late → finish ${minToTime(deadlineMin - 1)}` : `${Math.abs(diffManual)} min early → finish ${minToTime(deadlineMin - 1)}`}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 4 }}>
                  <button onClick={() => setDiffManual(d => d + 1)}
                    style={{ width: 44, height: 52, borderRadius: 10, background: "#101e15", border: "1px solid #6effa033", color: "#6effa0", fontSize: 18, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>+</button>
                  <button onClick={() => setDiffManual(d => d + 5)}
                    style={{ width: 44, height: 52, borderRadius: 10, background: "#0d1a10", border: "1px solid #6effa044", color: "#6effa0", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>+5</button>
                </div>
              </div>
              <div style={{ display: "flex", gap: 6, marginTop: 12, justifyContent: "center" }}>
                {[-5, -3, -1, 0, 1, 3, 5].map(v => (
                  <button key={v} onClick={() => setDiffManual(v)}
                    style={{
                      background: diffManual === v ? (v >= 3 ? "#7a1a1a" : v >= 1 ? "#4a3a00" : v < 0 ? "#1a3a1a" : "#1a2a3a") : "#1e2135",
                      border: `1px solid ${diffManual === v ? "transparent" : "#2a2d4a"}`,
                      color: diffManual === v ? "#fff" : "#666",
                      borderRadius: 6, padding: "4px 8px", cursor: "pointer", fontSize: 13, fontFamily: "inherit", fontWeight: 700,
                    }}>
                    {v > 0 ? `+${v}` : v}
                  </button>
                ))}
              </div>
            </div>
          )}

          {canConfirm && (
            <button onClick={markHole}
              style={{
                width: "100%",
                background: `linear-gradient(135deg, ${group.color}88, ${group.color})`,
                border: "none", color: "#fff", borderRadius: 10, padding: "16px",
                cursor: "pointer",
                fontFamily: "'Bebas Neue'", letterSpacing: 3, fontSize: 22, fontWeight: 700,
              }}
            >
              ✓ Confirm H{currentHole + 1} &nbsp;→&nbsp; {currentSlot < 17 ? `H${holeOrder[currentSlot + 1] + 1}` : "Finish"}
            </button>
          )}

          <div style={{ marginTop: canConfirm ? 16 : 0 }}>
          {/* MN Active Banner */}
        {mnActive && (
          <div style={{
            background: "#001a2a", border: "1px solid #4e9af1aa",
            borderRadius: 10, padding: "10px 16px", marginBottom: 16,
            boxShadow: "0 0 20px #4e9af133",
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
              <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#4e9af1", boxShadow: "0 0 8px #4e9af1", animation: "pulse 1.5s infinite" }} />
              <span style={{ fontSize: 13, color: "#4e9af1", fontWeight: 700, letterSpacing: 1 }}>👁 MONITORING</span>
              {mnName && <span style={{ fontSize: 12, color: "#8899cc" }}>by {mnName}</span>}
              <button
                onClick={offMN}
                style={{
                  marginLeft: "auto", background: "#1a2a3a", border: "1px solid #4e9af188",
                  color: "#ff7070", borderRadius: 7, padding: "5px 12px",
                  cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700,
                }}
              >✕ Off MN</button>
            </div>
            {currentSlot < 18 && (
              <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 10, flexWrap: "wrap" }}>
                <span style={{ fontSize: 11, color: "#8899cc", fontWeight: 700, letterSpacing: 0.5 }}>⚡ Bad Time →</span>
                {playerNums.map(n => {
                  const label = `P${n}`;
                  const isFlagged = tmActive && (tmTarget === "All" || (tmTarget || "").split(",").map(s => s.trim()).includes(label));
                  const isBadTimed = badTimePlayers.has(label);
                  const count = badTimeCounts[label] || 0;
                  const usedUpInMN = isFlagged; // once P becomes a TM target, further Bad Time presses go through the TM row instead
                  return (
                    <button
                      key={n}
                      onClick={() => !usedUpInMN && triggerBadTimeFor(n)}
                      disabled={usedUpInMN}
                      title={usedUpInMN ? `${label} is already a TM target — use the TIMING row below for further Bad Times` : `Bad Time — ${label} → flags TM on this hole immediately`}
                      style={{
                        background: usedUpInMN ? "#1a1a1a" : "#2a0020",
                        border: `1px solid ${usedUpInMN ? "#3a3a3a" : "#ff6ec788"}`,
                        color: usedUpInMN ? "#666" : "#ff6ec7",
                        borderRadius: 6, padding: "4px 10px", cursor: usedUpInMN ? "not-allowed" : "pointer",
                        fontFamily: "inherit", fontSize: 12, fontWeight: 700,
                      }}
                    >{isBadTimed ? "⚡ " : ""}{label}{count >= 2 ? ` x${count}` : ""}</button>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {/* TM Active Banner */}
        {tmActive && (
          <div style={{
            background: "#2a0020", border: "1px solid #ff6ec7aa",
            borderRadius: 10, padding: "10px 16px", marginBottom: 16,
            boxShadow: "0 0 20px #ff6ec733",
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
              <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#ff6ec7", boxShadow: "0 0 8px #ff6ec7", animation: "pulse 1.5s infinite" }} />
              <span style={{ fontSize: 13, color: "#ff6ec7", fontWeight: 700, letterSpacing: 1 }}>⏱ TIMING</span>
              <span style={{ fontSize: 12, color: "#ffb3e6", fontWeight: 700 }}>{tmTarget || "All"}</span>
              {tmName && <span style={{ fontSize: 12, color: "#8899cc" }}>by {tmName}</span>}
              <button
                onClick={offTM}
                style={{
                  marginLeft: "auto", background: "#3a0030", border: "1px solid #ff6ec788",
                  color: "#ff7070", borderRadius: 7, padding: "5px 12px",
                  cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700,
                }}
              >✕ Off TM</button>
            </div>
            {currentSlot < 18 && (
              <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 10, flexWrap: "wrap" }}>
                <span style={{ fontSize: 11, color: "#8899cc", fontWeight: 700, letterSpacing: 0.5 }}>⚡ Bad Time →</span>
                {playerNums.map(n => {
                  const label = `P${n}`;
                  const isFlagged = tmTarget === "All" || (tmTarget || "").split(",").map(s => s.trim()).includes(label);
                  const isBadTimed = badTimePlayers.has(label);
                  const count = badTimeCounts[label] || 0;
                  const notYetTarget = !isFlagged; // not a TM target yet — must be added via the MN row's Bad Time button first
                  return (
                    <button
                      key={n}
                      onClick={() => !notYetTarget && triggerBadTimeFor(n)}
                      disabled={notYetTarget}
                      title={notYetTarget ? `${label} is not a TM target yet — use Bad Time in the MONITORING row above first` : isBadTimed ? `${label} already flagged Bad Time — tap to log it again` : `Bad Time — ${label} → flags TM on this hole immediately`}
                      style={{
                        background: notYetTarget ? "#1a1a1a" : isBadTimed ? "#ff3d3d" : "#ff6ec7",
                        border: `1px solid ${notYetTarget ? "#3a3a3a" : isBadTimed ? "#ff3d3d" : "#ff6ec7"}`,
                        color: notYetTarget ? "#666" : "#1a0014",
                        borderRadius: 6, padding: "4px 10px", cursor: notYetTarget ? "not-allowed" : "pointer",
                        fontFamily: "inherit", fontSize: 12, fontWeight: 700,
                      }}
                    >{isBadTimed ? "⚡ " : ""}{label}{count >= 2 ? ` x${count}` : ""}</button>
                  );
                })}
              </div>
            )}
          </div>
        )}

          </div>

        </div>
        )}

        {done && (
          <div style={{ background: "#141626", border: "1px solid #6effa044", borderRadius: 14, padding: 32, marginBottom: 20, textAlign: "center" }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 48, color: "#6effa0", letterSpacing: 4 }}>✓ All 18 holes complete</div>
            <div style={{ color: "#9aa2c7", marginTop: 8 }}>{group.name} has finished all holes</div>
          </div>
        )}


        {!hideLog && (
        /* Hole Log Table */
        <div style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 12, overflow: "hidden" }}>
          <div style={{ padding: "14px 16px", borderBottom: "1px solid #2a2d4a", fontSize: 12, color: "#4e9af1", letterSpacing: 2, fontWeight: 700 }}>
            📊 HOLE LOG — <span style={{ color: "#8890b8", fontWeight: 400 }}>press ✏ to edit</span>
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: "#0d0f1a" }}>
                  {["Hole","Par","Due","Time","Diff","Status","WN/MN/TM",""].map((h, i) => (
                    <th key={i} style={{ padding: "8px 6px", color: h === "WN/MN/TM" ? "#ffd966" : "#555", fontWeight: 600, textAlign: "center", borderBottom: "1px solid #2a2d4a" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {/* Starting row — tee time */}
                {(() => {
                  const firstHi = holeOrder[0];
                  const firstHd = holeData[firstHi];
                  const startTime = firstHd?.startTime ?? group.startTime;
                  return (
                    <tr style={{ background: "#0a1a0a", borderLeft: "3px solid #6effa044" }}>
                      <td colSpan={2} style={{ padding: "6px 8px", textAlign: "center", color: "#6effa0", fontWeight: 700, fontSize: 12, letterSpacing: 1 }}>🏌 Tee off</td>
                      <td style={{ padding: "6px 8px", textAlign: "center", color: "#8899cc" }}>{minToTime(schedule[holeOrder[0]] ?? 0)}</td>
                      <td style={{ padding: "6px 8px", textAlign: "center", color: "#6effa0", fontWeight: 700 }}>
                        {minToTime((schedule[holeOrder[0]] ?? 0) - 1 + delayMin)}
                      </td>
                      <td style={{ padding: "6px 8px", textAlign: "center" }}>
                        {delayMin > 0
                          ? <span style={{ color: "#ffd966", fontWeight: 700 }}>+{delayMin}m</span>
                          : <span style={{ color: "#666f99" }}>—</span>}
                      </td>
                      <td colSpan={3} style={{ padding: "6px 8px" }} />
                    </tr>
                  );
                })()}

                {holeOrder.map((i, slot) => {
                  const isEditing = editingHole === i;
                  const hd = holeData[i];
                  const hasEnd = !!hd?.endTime;
                  const hasStart = !!hd?.startTime;
                  const diff = diffAtHole(i);
                  const rowStatus = hasEnd ? getStatus(diff) : "idle";
                  const isActive = slot === currentSlot && !done;

                  const jumpToSlot = () => {
                    if (isEditing) return;
                    setCurrentSlot(slot);
                    setRecordedEnd(null);
                    setDiffManual(0);
                  };

                  return (
                    <tr key={i}
                      onClick={jumpToSlot}
                      style={{
                        background: isEditing ? "#1e2135" : isActive ? `${group.color}11` : "transparent",
                        transition: "background 0.15s",
                        borderLeft: isActive ? `3px solid ${group.color}` : "3px solid transparent",
                        cursor: isEditing ? "default" : "pointer",
                      }}
                      onMouseEnter={e => { if (!isEditing && !isActive) e.currentTarget.style.background = "#ffffff07"; }}
                      onMouseLeave={e => { if (!isEditing && !isActive) e.currentTarget.style.background = "transparent"; }}
                    >
                      <td style={{ padding: "6px 8px", textAlign: "center", color: isActive ? group.color : "#8888aa", fontWeight: isActive ? 700 : 600 }}>
                        {i + 1}{isActive ? " ◀" : ""}
                      </td>
                      <td style={{ padding: "6px 8px", textAlign: "center", color: "#ccc" }}>Par {pars[i]}</td>
                      <td style={{ padding: "6px 8px", textAlign: "center", color: "#8899cc" }}>{minToTime((schedule[i] ?? 0) + (parTimes?.[i] ?? 14))}</td>
                      <td style={{ padding: "6px 8px", textAlign: "center" }}>
                        {isEditing ? (
                          <input type="time" value={editVal} autoFocus
                            onChange={e => setEditVal(e.target.value)}
                            onKeyDown={e => { if (e.key === "Enter") confirmEditHole(i); if (e.key === "Escape") setEditingHole(null); }}
                            style={{ background: "#0d0f1a", border: "1px solid #f16b4e", color: "#eee", borderRadius: 6, padding: "3px 6px", fontFamily: "inherit", fontSize: 13, width: 80 }}
                          />
                        ) : (
                          <span
                              onClick={() => hasEnd && startEditHole(i, "end")}
                              style={{ color: hasEnd ? "#eee" : "#333", cursor: hasEnd ? "pointer" : "default" }}
                            >{hasEnd ? (
                              hd.manualDiff !== undefined
                                ? minToTime((schedule[i] ?? 0) + (parTimes?.[i] ?? 14) - 1 + hd.manualDiff)
                                : hd.endTime
                            ) : "—"}</span>
                        )}
                      </td>
                      <td style={{ padding: "6px 8px", textAlign: "center" }}>
                        {editingDiff === i ? (
                          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 3 }}>
                            <input
                              type="number" autoFocus value={editDiffVal}
                              onChange={e => setEditDiffVal(e.target.value)}
                              onKeyDown={e => { if (e.key === "Enter") confirmEditDiff(i); if (e.key === "Escape") setEditingDiff(null); }}
                              style={{ width: 60, background: "#0d0f1a", border: "1px solid #ffd966", color: "#ffd966", fontFamily: "inherit", fontSize: 14, borderRadius: 6, padding: "3px 4px", textAlign: "center", outline: "none" }}
                            />
                            <div style={{ display: "flex", gap: 3 }}>
                              <button onClick={() => confirmEditDiff(i)} style={{ background: "#1a3a1a", border: "none", color: "#6effa0", borderRadius: 5, padding: "2px 6px", cursor: "pointer", fontSize: 11 }}>✓</button>
                              <button onClick={() => setEditingDiff(null)} style={{ background: "#2a2a2a", border: "none", color: "#888", borderRadius: 5, padding: "2px 6px", cursor: "pointer", fontSize: 11 }}>✕</button>
                            </div>
                          </div>
                        ) : hasEnd && hasStart && diff !== null && !isEditing ? (
                          <div onClick={() => startEditDiff(i)} style={{ cursor: "pointer" }} title="Edit difference">
                            <span style={{ color: diffColor(diff), fontWeight: 700 }}>
                              {diff > 0 ? `+${diff}` : diff}m
                            </span>
                          </div>
                        ) : "—"}
                      </td>
                      <td style={{ padding: "6px 8px" }}>
                        {!isEditing && <StatusBadge status={rowStatus} />}
                      </td>
                      <td style={{ padding: "6px 4px", textAlign: "center", minWidth: 90 }} onClick={e => e.stopPropagation()}>
                        {(() => {
                          const logs = actionLogs.map((l, idx) => ({ ...l, idx })).filter(l => l.holeIdx === i);
                          const isFutureMN = mnActive && slot === lastMNSlot + 1;
                          const isFutureTM = tmActive && slot === lastTMSlot + 1;
                          const wnDisabledRow = (actionLogs.some(l => l.type === "WN" && l.holeIdx === i)) || mnActive || tmActive;
                          const mnDisabledHereRow = actionLogs.some(l => l.type === "MN" && !l.badTime && !l.off && l.holeIdx === i);
                          const mnAlreadyUsedRow = actionLogs.some(l => l.type === "MN" && l.off);
                          const mnDisabledRow = mnActive || mnDisabledHereRow || mnAlreadyUsedRow;
                          return (
                            <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                              <div style={{ display: "flex", gap: 3, justifyContent: "center" }}>
                                <button onClick={() => !wnDisabledRow && openActionModal("WN", i)}
                                  disabled={wnDisabledRow}
                                  title={mnActive || tmActive ? "MN/TM is active — WN cannot be used right now" : wnDisabledRow ? "WN already logged on this hole" : undefined}
                                  style={{ background: wnDisabledRow ? "#1a1a1a" : "#2a1a00", border: `1px solid ${wnDisabledRow ? "#3a3a3a" : "#ffd96688"}`, color: wnDisabledRow ? "#666" : "#ffd966", borderRadius: 5, padding: "3px 7px", cursor: wnDisabledRow ? "not-allowed" : "pointer", fontSize: 11, fontWeight: 700, fontFamily: "inherit" }}>WN</button>
                                <button onClick={() => !mnDisabledRow && openActionModal("MN", i)}
                                  disabled={mnDisabledRow}
                                  title={mnActive ? "MN is already running — turn MN off first" : mnDisabledHereRow ? "MN already logged on this hole" : mnAlreadyUsedRow ? "This group has already had MN turned on and off — MN can only be used once per group" : undefined}
                                  style={{ background: mnDisabledRow ? "#1a1a1a" : "#001a2a", border: `1px solid ${mnDisabledRow ? "#3a3a3a" : "#4e9af188"}`, color: mnDisabledRow ? "#666" : "#4e9af1", borderRadius: 5, padding: "3px 7px", cursor: mnDisabledRow ? "not-allowed" : "pointer", fontSize: 11, fontWeight: 700, fontFamily: "inherit" }}>MN</button>
                                <button onClick={() => openActionModal("TM", i)}
                                  style={{ background: "#2a0020", border: "1px solid #ff6ec788", color: "#ff6ec7", borderRadius: 5, padding: "3px 7px", cursor: "pointer", fontSize: 11, fontWeight: 700, fontFamily: "inherit" }}>TM</button>
                              </div>
                              {logs.map((l, li) => (
                                <div key={li} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: logColor(l.type), background: `${logBg(l.type)}44`, borderRadius: 4, padding: "2px 4px", textAlign: "left", lineHeight: 1.4 }}>
                                  <span style={{ flex: 1 }}>
                                    {l.badTime ? (
                                      <><span style={{ fontWeight: 700 }}>⚡ TM Bad Time</span> {l.target}{badTimeOccurrence.has(l) ? ` (#${badTimeOccurrence.get(l)})` : ""}{l.name ? ` - ${l.name}` : ""}</>
                                    ) : l.off ? (
                                      <><span style={{ fontWeight: 700 }}>✕ Off {l.type}</span>{l.name ? ` - ${l.name}` : ""}</>
                                    ) : (
                                      <><span style={{ fontWeight: 700 }}>{l.type}</span> {l.target ? `${l.target} ` : ""}{l.name ? `- ${l.name}` : ""}</>
                                    )}
                                  </span>
                                  <button
                                    onClick={() => setDeleteLogConfirm(l.idx)}
                                    title="Delete this log"
                                    style={{ background: "none", border: "none", color: "#ff7070", cursor: "pointer", fontSize: 11, padding: "0 2px", lineHeight: 1, flexShrink: 0 }}
                                  >🗑</button>
                                </div>
                              ))}
                              {isFutureMN && !logs.some(l => l.type === "MN") && (
                                <div style={{ fontSize: 11, color: "#4e9af1", background: "#001a2a66", border: "1px dashed #4e9af144", borderRadius: 4, padding: "2px 4px", textAlign: "center", lineHeight: 1.4 }}>
                                  👁 MN{mnName ? ` - ${mnName}` : ""}
                                </div>
                              )}
                              {isFutureTM && !logs.some(l => l.type === "TM") && (
                                <div style={{ fontSize: 11, color: "#ff6ec7", background: "#2a002066", border: "1px dashed #ff6ec744", borderRadius: 4, padding: "2px 4px", textAlign: "center", lineHeight: 1.4 }}>
                                  ⏱ TM {tmName ? `- ${tmName}` : ""}
                                </div>
                              )}
                            </div>
                          );
                        })()}
                      </td>
                      <td style={{ padding: "6px 8px", textAlign: "center" }} onClick={e => e.stopPropagation()}>
                        {isEditing ? (
                          <div style={{ display: "flex", gap: 3, justifyContent: "center" }}>
                            <button onClick={() => confirmEditHole(i)} style={{ ...btnStyle("#1a3a1a","#6effa0"), padding: "3px 7px", fontSize: 12 }}>✓</button>
                            <button onClick={() => setEditingHole(null)} style={{ ...btnStyle("#2a2a2a","#888"), padding: "3px 7px", fontSize: 12 }}>✕</button>
                            {(hasStart || hasEnd) && <button onClick={() => clearHole(i)} style={{ ...btnStyle("#3a1a1a","#ff7070"), padding: "3px 7px", fontSize: 12 }}>Delete</button>}
                          </div>
                        ) : hasEnd ? (
                          <button onClick={() => startEditHole(i, "end")} style={{ background: "none", border: "1px solid #f16b4e44", color: "#f16b4e", borderRadius: 6, padding: "3px 6px", cursor: "pointer", fontSize: 11, fontFamily: "inherit" }} title="Edit finish time">🏁 ✏</button>
                        ) : null}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
        )}
      </div>

      {/* WN/MN/TM Modal */}
      {actionModal && (
        <div style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }}>
          <div style={{ background: "#141626", border: `1px solid ${logColor(actionModal.type)}88`, borderRadius: 14, padding: 28, minWidth: 300, boxShadow: "0 20px 60px #000" }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 22, letterSpacing: 3, color: logColor(actionModal.type), marginBottom: 6 }}>
              {actionModal.type === "WN" ? "⚠ PACE OF PLAY WARNING" : actionModal.type === "MN" ? "👁 MONITORING" : "⏱ TIMING"}
            </div>
            <div style={{ fontSize: 12, color: "#8890b8", marginBottom: 18 }}>
              Hole {actionModal.holeIdx + 1} — {minToTime(nowInMin())}
            </div>

            {actionModal.type === "TM" && (
              <div style={{ marginBottom: 18 }}>
                <div style={{ fontSize: 13, color: "#aaa", marginBottom: 8 }}>Players being timed (TM)</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  <button
                    onClick={() => toggleTarget("ALL")}
                    style={{
                      background: actionTargets.includes("ALL") ? "#ff6ec7" : "#0d0f1a",
                      color: actionTargets.includes("ALL") ? "#1a0014" : "#ff6ec7",
                      border: "1px solid #ff6ec788", borderRadius: 8, padding: "6px 12px",
                      cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700,
                    }}
                  >All</button>
                  {playerNums.map(n => (
                    <button
                      key={n}
                      onClick={() => toggleTarget(n)}
                      style={{
                        background: actionTargets.includes(n) ? "#ff6ec7" : "#0d0f1a",
                        color: actionTargets.includes(n) ? "#1a0014" : "#ff6ec7",
                        border: "1px solid #ff6ec788", borderRadius: 8, padding: "6px 12px",
                        cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700,
                      }}
                    >P{n}</button>
                  ))}
                </div>
                <div style={{ fontSize: 11, color: "#8890b8", marginTop: 6 }}>Multiple players can be selected</div>
              </div>
            )}

            <div style={{ fontSize: 13, color: "#aaa", marginBottom: 8 }}>Recorded by</div>
            <input
              autoFocus
              value={actionName}
              onChange={e => setActionName(e.target.value)}
              onKeyDown={e => { if (e.key === "Enter") confirmAction(); if (e.key === "Escape") setActionModal(null); }}
              placeholder="Enter name..."
              style={{ width: "100%", background: "#0d0f1a", border: `1px solid ${logColor(actionModal.type)}44`, color: "#eee", borderRadius: 8, padding: "10px 14px", fontFamily: "inherit", fontSize: 14, boxSizing: "border-box", marginBottom: 18 }}
            />
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={confirmAction}
                disabled={actionModal.type === "TM" && actionTargets.length === 0}
                style={{
                  flex: 1,
                  background: (actionModal.type === "TM" && actionTargets.length === 0) ? "#1a1c2c" : logBg(actionModal.type),
                  border: `1px solid ${(actionModal.type === "TM" && actionTargets.length === 0) ? "#333" : logColor(actionModal.type)}`,
                  color: (actionModal.type === "TM" && actionTargets.length === 0) ? "#555" : logColor(actionModal.type),
                  borderRadius: 8, padding: "10px",
                  cursor: (actionModal.type === "TM" && actionTargets.length === 0) ? "not-allowed" : "pointer",
                  fontFamily: "'Bebas Neue'", fontSize: 16, letterSpacing: 2,
                }}>
                ✓ Save {actionModal.type}
              </button>
              <button onClick={() => setActionModal(null)}
                style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "10px 16px", cursor: "pointer", fontFamily: "inherit", fontSize: 14 }}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Delete log confirmation popup */}
      {deleteLogConfirm !== null && actionLogs[deleteLogConfirm] && (() => {
        const l = actionLogs[deleteLogConfirm];
        return (
          <div style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1100 }}>
            <div style={{ background: "#141626", border: "1px solid #ff707088", borderRadius: 14, padding: 28, minWidth: 280, maxWidth: 340, boxShadow: "0 20px 60px #000" }}>
              <div style={{ fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 3, color: "#ff7070", marginBottom: 10 }}>Delete {l.off ? `Off ${l.type}` : l.type}?</div>
              <div style={{ fontSize: 13, color: "#aaa", marginBottom: 20, lineHeight: 1.6 }}>
                Are you sure?
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => deleteLogAt(deleteLogConfirm)}
                  style={{ flex: 1, background: "#2a0a0a", border: "1px solid #ff7070", color: "#ff7070", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 15, letterSpacing: 2 }}>
                  ✓ Yes, delete
                </button>
                <button onClick={() => setDeleteLogConfirm(null)}
                  style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "10px 16px", cursor: "pointer", fontFamily: "inherit", fontSize: 14 }}>
                  Cancel
                </button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}

function btnStyle(bg, color) {
  return {
    background: bg, border: `1px solid ${color}44`, color,
    borderRadius: 8, padding: "8px 14px", cursor: "pointer",
    fontFamily: "'IBM Plex Mono'", fontSize: 14, fontWeight: 700,
  };
}

// ─── Summary Report (Pace of Play + Suspension & Resumption) ───────────────────────
function SummaryScreen({ groups, groupData, pars, parTimes, playersPerGroup, suspensions, isSuspended, pendingStopTime, totalOffsetMin, onBack, currentUser, onLogout }) {
  const sides = getGroupSides(groups);
  const totalAllowed = (parTimes || []).reduce((a, b) => a + b, 0);
  const [expandedTMGroups, setExpandedTMGroups] = useState(() => new Set());
  const toggleTMGroup = (groupId) => {
    setExpandedTMGroups(prev => {
      const next = new Set(prev);
      if (next.has(groupId)) next.delete(groupId); else next.add(groupId);
      return next;
    });
  };
  const [expandedBadTimePlayers, setExpandedBadTimePlayers] = useState(() => new Set());
  const toggleBadTimePlayer = (key) => {
    setExpandedBadTimePlayers(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  };

  const sideStats = sides.map(side => {
    const first3 = side.groups.slice(0, 3);
    const lastGroup = side.groups[side.groups.length - 1];

    const first3Details = first3.map(g => {
      const p = computeGroupProgress(g, groupData[g.id], parTimes);
      return { name: g.name, diff: p.lastDiff, isComplete: p.isComplete };
    });

    const lastProgress = lastGroup ? computeGroupProgress(lastGroup, groupData[lastGroup.id], parTimes) : null;

    return {
      ...side,
      groupCount: side.groups.length,
      first3Details,
      lastDiff: lastProgress?.lastDiff ?? null,
      lastComplete: lastProgress?.isComplete ?? false,
    };
  });

  const thS = { padding: "10px 8px", color: "#888", fontWeight: 700, textAlign: "center", borderBottom: "1px solid #2a2d4a", fontSize: 12, letterSpacing: 0.5 };
  const tdS = { padding: "10px 8px", textAlign: "center", borderBottom: "1px solid #1a1d2e", fontSize: 14 };

  // ─── TM (Timing) / Bad Time summary — every TM log across all groups, plus totals ──
  const tmRows = [];
  groups.forEach(g => {
    const logs = (groupData[g.id]?.actionLogs || []).filter(l => l.type === "TM");
    logs.forEach(l => {
      const players = l.target === "All" ? ["All"] : (l.target || "").split(",").map(s => s.trim()).filter(Boolean);
      tmRows.push({
        groupName: g.name,
        holeIdx: l.holeIdx,
        players,
        time: l.time,
        by: l.name || "—",
        badTime: !!l.badTime,
      });
    });
  });
  tmRows.sort((a, b) => (a.groupName || "").localeCompare(b.groupName || "") || (a.holeIdx - b.holeIdx));

  // Split distinct-player counting into two non-overlapping buckets, since Bad Time is
  // always a subset of TM (a Bad-Timed player is also "under TM") but should be reported
  // as its own category rather than double-counted inside the "normal TM" bucket:
  //   - badTimePlayers: players ever flagged via a Bad Time press, per group
  //   - normalTmPlayers: players targeted via a normal TM selection, MINUS anyone who was
  //     also Bad-Timed (so the two buckets never overlap)
  // "All" is treated as covering all 4 player slots (P1–P4) offered in the TM/Bad-Time controls.
  const totalBadTimePlayers = groups.reduce((sum, g) => {
    const logs = (groupData[g.id]?.actionLogs || []).filter(l => l.type === "TM" && l.badTime);
    const set = new Set();
    logs.forEach(l => set.add(l.target));
    return sum + set.size;
  }, 0);
  const totalNormalTMPlayers = groups.reduce((sum, g) => {
    const logs = (groupData[g.id]?.actionLogs || []).filter(l => l.type === "TM");
    const badTimeSet = new Set(logs.filter(l => l.badTime).map(l => l.target));
    const allSet = new Set();
    logs.forEach(l => {
      if (l.target === "All") {
        Array.from({ length: playersPerGroup || 3 }, (_, i) => i + 1).forEach(n => allSet.add(`P${n}`));
      } else {
        (l.target || "").split(",").map(s => s.trim()).filter(Boolean).forEach(p => allSet.add(p));
      }
    });
    let count = 0;
    allSet.forEach(p => { if (!badTimeSet.has(p)) count += 1; });
    return sum + count;
  }, 0);
  const totalPlayersTimed = totalNormalTMPlayers + totalBadTimePlayers;
  const totalBadTime = tmRows.filter(r => r.badTime).length;

  // Per-group, per-player breakdown: which holes a player was under TM (first → last hole
  // logged) and how many times they were specifically Bad-Timed. "All" is expanded to all
  // 4 player slots. Used to render the collapsible group-by-group summary below.
  const tmGroupSummaries = groups.map(g => {
    const logs = (groupData[g.id]?.actionLogs || []).filter(l => l.type === "TM");
    if (logs.length === 0) return null;
    const playerMap = {};
    logs.forEach(l => {
      const labels = l.target === "All" ? Array.from({ length: playersPerGroup || 3 }, (_, i) => i + 1).map(n => `P${n}`) : (l.target || "").split(",").map(s => s.trim()).filter(Boolean);
      labels.forEach(label => {
        if (!playerMap[label]) playerMap[label] = { holes: [], badTimeCount: 0, names: new Set(), badTimeHoles: [] };
        playerMap[label].holes.push(l.holeIdx);
        if (l.badTime) {
          playerMap[label].badTimeCount += 1;
          playerMap[label].badTimeHoles.push({ holeIdx: l.holeIdx, time: l.time, name: l.name || "—" });
        }
        if (l.name) playerMap[label].names.add(l.name);
      });
    });
    const players = Object.keys(playerMap).sort().map(label => {
      const holes = playerMap[label].holes;
      const badTimeHoles = playerMap[label].badTimeHoles.slice().sort((a, b) => a.holeIdx - b.holeIdx);
      return {
        label,
        firstHole: Math.min(...holes) + 1,
        lastHole: Math.max(...holes) + 1,
        badTimeCount: playerMap[label].badTimeCount,
        badTimeHoles,
        recordedBy: Array.from(playerMap[label].names).join(", ") || "—",
      };
    });
    return {
      groupId: g.id,
      groupName: g.name,
      players,
      totalBadTime: players.reduce((s, p) => s + p.badTimeCount, 0),
    };
  }).filter(Boolean);

  return (
    <div style={{ minHeight: "100vh", background: "#0a0b14", color: "#eee", fontFamily: "'IBM Plex Mono', monospace" }}>
      <div style={{ padding: "16px 24px", borderBottom: "1px solid #1a1d2e", display: "flex", alignItems: "center", gap: 12 }}>
        <button onClick={onBack} style={btnStyle("#141626", "#8899cc")}>← Back</button>
        <div style={{ fontFamily: "'Bebas Neue'", fontSize: 24, letterSpacing: 2, color: "#4e9af1" }}>📈 SUMMARY REPORT</div>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 12, color: "#9aa2c7" }}>{currentUser}</span>
          <LogoutButton onLogout={onLogout} />
        </div>
      </div>

      <div style={{ padding: "20px 24px", maxWidth: 900, margin: "0 auto" }}>

        {/* ─── Pace of Play ─────────────────────────────────────────── */}
        <div style={{ fontSize: 14, fontWeight: 700, color: "#eee", letterSpacing: 1, marginBottom: 10 }}>⛳ PACE OF PLAY</div>
        <div style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 12, padding: "4px 0", marginBottom: 24, overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={thS}>Start Point</th>
                <th style={thS}>Groups</th>
                <th style={thS}>Total<br />(Time Allowed)</th>
                <th style={thS}>1st Group<br />Finish Time (Δ)</th>
                <th style={thS}>Last Group<br />Finish Time (Δ)</th>
              </tr>
            </thead>
            <tbody>
              {sideStats.map(s => (
                <tr key={s.startHole}>
                  <td style={{ ...tdS, textAlign: "left", color: s.meta.color, fontWeight: 700 }}>{s.meta.shortLabel}</td>
                  <td style={tdS}>{s.groupCount}</td>
                  <td style={tdS}>{minToHM(totalAllowed)}</td>
                  <td style={{ ...tdS, textAlign: "left" }}>
                    {s.first3Details.length ? (
                      <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                        {s.first3Details.map((d, i) => (
                          <span key={i} style={{ whiteSpace: "nowrap" }}>
                            <span style={{ color: "#888" }}>{d.name}</span>{" "}
                            <span style={{ color: diffColor(d.diff), fontWeight: 700 }}>
                              {d.diff === null ? "–" : (d.diff > 0 ? `+${d.diff}` : `${d.diff}`)}
                            </span>
                            {d.diff !== null && !d.isComplete ? "*" : ""}
                          </span>
                        ))}
                      </div>
                    ) : <span style={{ color: "#8890b8" }}>–</span>}
                  </td>
                  <td style={{ ...tdS, color: diffColor(s.lastDiff), fontWeight: 700 }}>
                    {fmtDiff(s.lastDiff)}{s.lastDiff !== null && !s.lastComplete ? " *" : ""}
                  </td>
                </tr>
              ))}
              {sideStats.length === 0 && (
                <tr><td colSpan={5} style={{ ...tdS, color: "#8890b8", padding: 20 }}>No groups yet</td></tr>
              )}
            </tbody>
          </table>
        </div>
        <div style={{ fontSize: 11, color: "#8890b8", marginTop: -14, marginBottom: 24 }}>
          * = based on the most recently completed hole (round not finished yet). Δ is minutes ahead (–) / behind (+) of scheduled pace.
        </div>

        {/* ─── Suspension & Resumption ──────────────────────────────── */}
        <div style={{ fontSize: 14, fontWeight: 700, color: "#eee", letterSpacing: 1, marginBottom: 10 }}>⏸ SUSPENSION &amp; RESUMPTION</div>
        <div style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 12, padding: "4px 0", overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={thS}>#</th>
                <th style={thS}>Stop Time</th>
                <th style={thS}>Resume Time</th>
                <th style={thS}>Duration</th>
              </tr>
            </thead>
            <tbody>
              {suspensions.map((s, i) => (
                <tr key={i}>
                  <td style={tdS}>{i + 1}</td>
                  <td style={tdS}>{s.stopTime}</td>
                  <td style={tdS}>{s.resumeTime}</td>
                  <td style={tdS}>{minToHM(s.offsetMin)}</td>
                </tr>
              ))}
              {isSuspended && (
                <tr>
                  <td style={tdS}>{suspensions.length + 1}</td>
                  <td style={{ ...tdS, color: "#ffd966" }}>{pendingStopTime}</td>
                  <td style={{ ...tdS, color: "#ffd966" }}>Ongoing…</td>
                  <td style={{ ...tdS, color: "#ffd966" }}>—</td>
                </tr>
              )}
              {suspensions.length === 0 && !isSuspended && (
                <tr><td colSpan={4} style={{ ...tdS, color: "#8890b8", padding: 20 }}>No suspensions recorded</td></tr>
              )}
            </tbody>
            {(suspensions.length > 0 || isSuspended) && (
              <tfoot>
                <tr>
                  <td colSpan={3} style={{ ...tdS, textAlign: "right", color: "#8899cc", fontWeight: 700, borderBottom: "none" }}>Total Suspended</td>
                  <td style={{ ...tdS, color: "#6effa0", fontWeight: 700, borderBottom: "none" }}>{minToHM(totalOffsetMin)}</td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>

        {/* ─── TM (Timing) / Bad Time Summary ──────────────────────── */}
        <div style={{ fontSize: 14, fontWeight: 700, color: "#eee", letterSpacing: 1, margin: "24px 0 10px" }}>⏱ TM &amp; BAD TIME SUMMARY</div>
        <div style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 12, padding: "10px", overflowX: "auto" }}>
          {tmGroupSummaries.length === 0 && (
            <div style={{ color: "#8890b8", padding: 20, textAlign: "center", fontSize: 14 }}>No TM records yet</div>
          )}
          {tmGroupSummaries.map(gs => {
            const isOpen = expandedTMGroups.has(gs.groupId);
            return (
              <div key={gs.groupId} style={{ border: "1px solid #2a2d4a", borderRadius: 8, marginBottom: 8, overflow: "hidden" }}>
                <button
                  onClick={() => toggleTMGroup(gs.groupId)}
                  style={{
                    width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between",
                    background: isOpen ? "#1a1d33" : "#10111f", border: "none", cursor: "pointer",
                    padding: "10px 14px", fontFamily: "inherit", color: "#eee", textAlign: "left",
                  }}
                >
                  <span style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ fontSize: 12, color: "#9aa2c7", transform: isOpen ? "rotate(90deg)" : "none", transition: "transform 0.15s", display: "inline-block" }}>▶</span>
                    <span style={{ fontWeight: 700, fontSize: 14 }}>{gs.groupName}</span>
                  </span>
                  <span style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 12 }}>
                    <span style={{ color: "#ff6ec7" }}>{gs.players.length} Timing</span>
                    {gs.totalBadTime > 0 && <span style={{ color: "#ffd966" }}>⚡ {gs.totalBadTime} Bad Time</span>}
                  </span>
                </button>
                {isOpen && (
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr>
                        <th style={thS}>Player</th>
                        <th style={thS}>TM Hole Range</th>
                        <th style={thS}>Bad Time</th>
                        <th style={thS}>Recorded By</th>
                      </tr>
                    </thead>
                    <tbody>
                      {gs.players.map(p => {
                        const key = `${gs.groupId}-${p.label}`;
                        const badTimeOpen = expandedBadTimePlayers.has(key);
                        return (
                          <Fragment key={p.label}>
                            <tr>
                              <td style={{ ...tdS, color: "#ff6ec7", fontWeight: 700 }}>{p.label}</td>
                              <td style={tdS}>{p.firstHole === p.lastHole ? `H${p.firstHole}` : `H${p.firstHole} → H${p.lastHole}`}</td>
                              <td style={tdS}>
                                {p.badTimeCount > 0 ? (
                                  <button
                                    onClick={() => toggleBadTimePlayer(key)}
                                    style={{
                                      background: "none", border: "none", cursor: "pointer",
                                      color: "#ffd966", fontFamily: "inherit", fontSize: 14,
                                      display: "inline-flex", alignItems: "center", gap: 4, padding: 0,
                                    }}
                                  >
                                    ⚡ ×{p.badTimeCount}
                                    <span style={{ fontSize: 11, transform: badTimeOpen ? "rotate(90deg)" : "none", display: "inline-block", transition: "transform 0.15s" }}>▶</span>
                                  </button>
                                ) : "–"}
                              </td>
                              <td style={tdS}>{p.recordedBy}</td>
                            </tr>
                            {badTimeOpen && p.badTimeHoles.length > 0 && (
                              <tr>
                                <td colSpan={4} style={{ padding: "6px 12px 12px", borderBottom: "1px solid #1a1d2e", background: "#0d0f1a" }}>
                                  <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                                    {p.badTimeHoles.map((bt, bi) => (
                                      <span key={bi} style={{
                                        display: "inline-flex", alignItems: "center", gap: 5,
                                        background: "#2a0020", border: "1px solid #ffd96666",
                                        borderRadius: 6, padding: "3px 8px", fontSize: 12, color: "#ffd966",
                                      }}>
                                        ⚡ H{bt.holeIdx + 1} <span style={{ color: "#888" }}>· {bt.time} · {bt.name}</span>
                                      </span>
                                    ))}
                                  </div>
                                </td>
                              </tr>
                            )}
                          </Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                )}
              </div>
            );
          })}
          {tmGroupSummaries.length > 0 && (
            <div style={{ borderTop: "1px solid #2a2d4a", marginTop: 6, paddingTop: 10, fontSize: 13 }}>
              <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 4px" }}>
                <span style={{ color: "#8899cc", fontWeight: 700 }}>Timing</span>
                <span style={{ color: "#ff6ec7", fontWeight: 700 }}>{totalPlayersTimed}</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 4px", fontSize: 12 }}>
                <span style={{ color: "#9aa2c7" }}>⤷ via normal TM selection</span>
                <span style={{ color: "#ff6ec7" }}>{totalNormalTMPlayers}</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 4px", fontSize: 12 }}>
                <span style={{ color: "#9aa2c7" }}>⤷ via Bad Time</span>
                <span style={{ color: "#ffd966" }}>{totalBadTimePlayers}</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 4px" }}>
                <span style={{ color: "#8899cc", fontWeight: 700 }}>Bad Time</span>
                <span style={{ color: "#ffd966", fontWeight: 700 }}>{totalBadTime}</span>
              </div>
            </div>
          )}
        </div>

      </div>
    </div>
  );
}

// ─── Dashboard ────────────────────────────────────────────────────────────────
// Year → Tournament → Round → Search. Used on both the tournament screen and the
// dashboard, so switching rounds never means walking back through another page.
function RoundSelectorBar({ tournamentId, roundLabel, isAdmin, onOpen, compact }) {
  const [tournaments, setTournaments] = useState([]);
  const [year, setYear] = useState("");
  const [tid, setTid] = useState(tournamentId || "");
  const [rounds, setRounds] = useState([]);
  const [label, setLabel] = useState(roundLabel || "");
  const [loadingRounds, setLoadingRounds] = useState(false);
  const [busy, setBusy] = useState(false);

  const yearOf = (t) => {
    const m = (t.name || "").match(/(20\d{2})/);
    if (m) return m[1];
    return t.created_at ? String(new Date(t.created_at).getFullYear()) : String(new Date().getFullYear());
  };

  useEffect(() => {
    (async () => {
      const all = await fetchTournaments();
      const visible = all.filter(t => isAdmin || t.status !== "closed");
      setTournaments(visible);
      const current = visible.find(t => t.id === tournamentId) || visible[0];
      if (current) { setYear(yearOf(current)); setTid(current.id); }
    })();
  }, [tournamentId, isAdmin]);

  useEffect(() => {
    if (!tid) { setRounds([]); return; }
    let cancelled = false;
    (async () => {
      setLoadingRounds(true);
      const rs = await fetchRounds(tid);
      if (cancelled) return;
      setRounds(rs);
      const preferred = (tid === tournamentId && roundLabel && rs.find(r => r.label === roundLabel))
        || rs.find(r => r.status === "live") || rs[rs.length - 1];
      setLabel(preferred?.label || "");
      setLoadingRounds(false);
    })();
    return () => { cancelled = true; };
  }, [tid]);

  const years = (() => {
    const ys = new Set(tournaments.map(yearOf));
    if (year) ys.add(year);
    return [...ys].sort((a, b) => Number(b) - Number(a));
  })();
  const forYear = tournaments.filter(t => yearOf(t) === year);

  const go = async () => {
    const t = tournaments.find(x => x.id === tid);
    const r = rounds.find(x => x.label === label);
    if (!t || !r || busy) return;
    if (t.id === tournamentId && r.label === roundLabel) return; // already here
    setBusy(true);
    await onOpen(t, r);
    setBusy(false);
  };

  const sel = {
    background: "#0d0f1a", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 8,
    padding: "0 8px", height: compact ? 34 : 40, fontFamily: "inherit", fontSize: 12, outline: "none",
  };

  return (
    <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
      <select value={year}
        onChange={e => { setYear(e.target.value); const first = tournaments.find(t => yearOf(t) === e.target.value); setTid(first?.id || ""); }}
        style={{ ...sel, width: 84, flexShrink: 0 }}>
        {years.map(y => <option key={y} value={y}>{y}</option>)}
      </select>

      <select value={tid} onChange={e => setTid(e.target.value)} style={{ ...sel, flex: 1, minWidth: 120 }}>
        <option value="">— Tournament —</option>
        {forYear.map(t => <option key={t.id} value={t.id}>{t.name || "(untitled)"}</option>)}
      </select>

      <select value={label} onChange={e => setLabel(e.target.value)} disabled={!tid || loadingRounds}
        style={{ ...sel, width: 96, flexShrink: 0, opacity: tid ? 1 : 0.5 }}>
        <option value="">{loadingRounds ? "…" : "— Round —"}</option>
        {rounds.map(r => (
          <option key={r.id} value={r.label}>
            {r.label === "Q" ? "Q" : `R${r.label}`}{r.status === "live" ? " ●" : ""}
          </option>
        ))}
      </select>

      <button onClick={go} disabled={!tid || !label || busy}
        style={{
          height: compact ? 34 : 40, padding: "0 14px", borderRadius: 8, flexShrink: 0,
          cursor: (!tid || !label || busy) ? "not-allowed" : "pointer",
          fontFamily: "inherit", fontSize: 12, fontWeight: 700,
          background: (!tid || !label) ? "#1a1d2e" : "#1a4a8a",
          border: `1px solid ${(!tid || !label) ? "#2a2d4a" : "#4e9af1"}`,
          color: (!tid || !label) ? "#555" : "#fff",
        }}>
        {busy ? "…" : "🔍 Search"}
      </button>
    </div>
  );
}

function Dashboard({ groups, groupData, pars, parTimes, schedules, playersPerGroup, onChangePassword, tournamentName, hostVenue, roundLabel, tournamentId, isTrueAdmin, onOpenRound, onlineUsers, onSelectGroup, onBack, currentUser,
  suspensions, isSuspended, pendingStopTime, totalOffsetMin, onSuspendStop, onSuspendResume, onSuspendCancel, onSuspendEdit, onSuspendDelete, onLogout, onNavigateSummary, onUpdateGroupData }) {
  const [now, setNow] = useState(nowInMin());
  // Quick-record popup: clicking a hole cell opens the recording UI as a modal
  // instead of navigating to a new screen. Closes itself once a time is recorded.
  const [quickRecord, setQuickRecord] = useState(null); // { groupId, targetSlot } or null
  const [fitAllHoles, setFitAllHoles] = useState(() => {
    try { return localStorage.getItem("pop_fit_all_holes") === "1"; } catch { return false; }
  });
  const toggleFitAll = () => {
    setFitAllHoles(v => {
      const next = !v;
      try { localStorage.setItem("pop_fit_all_holes", next ? "1" : "0"); } catch {}
      return next;
    });
  };
  const [editSuspension, setEditSuspension] = useState(null); // { idx, stopTime, resumeTime } while editing a recorded stop
  const [collapsedTables, setCollapsedTables] = useState({}); // { [tableKey]: true } — folded schedule tables
  // Per-device hole focus: each marshal only supervises a few holes, so they can
  // hide the rest of the columns. Stored locally (not shared) — one device, one view.
  const [focusHoles, setFocusHoles] = useState(() => {
    try {
      const raw = localStorage.getItem("pop_focus_holes");
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr : [];
    } catch { return []; }
  });
  const [showFocusPicker, setShowFocusPicker] = useState(false);
  const saveFocusHoles = (next) => {
    setFocusHoles(next);
    try { localStorage.setItem("pop_focus_holes", JSON.stringify(next)); } catch {}
  };
  // Delete a WN/MN/TM log entry from the dashboard (in case of a mistaken tap) — confirmed via popup before removal
  const [deleteLogConfirm, setDeleteLogConfirm] = useState(null); // { groupId, idx } or null
  const deleteLogAt = (groupId, idx) => {
    const gd = groupData[groupId] || {};
    const nextLogs = (gd.actionLogs ?? []).filter((_, i2) => i2 !== idx);
    onUpdateGroupData(groupId, { actionLogs: nextLogs });
    setDeleteLogConfirm(null);
  };
  // Full reset of a status (MN or TM) for a group — removes every log entry of that
  // type (including the off-marker) and turns the active flag off, for when the
  // history has gotten tangled (e.g. accidental duplicate on/off taps) and a single
  // entry delete isn't enough.
  const [clearStatusConfirm, setClearStatusConfirm] = useState(null); // { groupId, type } or null
  const [editLogPopup, setEditLogPopup] = useState(null); // { groupId, idx } or null — editing a single WN/MN/TM log entry
  const clearStatusFor = (groupId, type) => {
    const gd = groupData[groupId] || {};
    const nextLogs = (gd.actionLogs ?? []).filter(l => l.type !== type);
    const reset = type === "MN"
      ? { actionLogs: nextLogs, mnActive: false, mnName: "" }
      : { actionLogs: nextLogs, tmActive: false, tmName: "", tmTarget: "" };
    onUpdateGroupData(groupId, reset);
    setClearStatusConfirm(null);
  };
  const [suspendModal, setSuspendModal] = useState(false); // "stop" | "resume" | false
  const [suspendStopInput, setSuspendStopInput] = useState(minToTime(nowInMin()));
  const [suspendResumeInput, setSuspendResumeInput] = useState(minToTime(nowInMin()));
  const [exportModal, setExportModal] = useState(false); // false | true
  const [exportCopied, setExportCopied] = useState(""); // which sheet was just copied, for a brief checkmark

  const exportData = exportModal ? buildDashboardExportData({ groups, groupData, pars, parTimes, schedules }) : null;

  const handleCopySheet = (label, text) => {
    copyTextToClipboard(text)
      .then(() => { setExportCopied(label); setTimeout(() => setExportCopied(""), 1800); })
      .catch(() => { setExportCopied(""); window.alert("Copy failed — please select the text manually and copy it."); });
  };

  useEffect(() => {
    const iv = setInterval(() => setNow(nowInMin()), 1000);
    return () => clearInterval(iv);
  }, []);

  const openStop = () => { setSuspendStopInput(minToTime(nowInMin())); setSuspendModal("stop"); };
  const openResume = () => { setSuspendResumeInput(minToTime(nowInMin())); setSuspendModal("resume"); };

  const confirmStop = () => {
    onSuspendStop(suspendStopInput);
    setSuspendModal(false);
  };
  const confirmResume = () => {
    onSuspendResume(suspendResumeInput);
    setSuspendModal(false);
  };

  // Find the "most recently finished" hole based on actual play order (following getHoleOrder for startHole)
  // Not simply by hole index order, to support groups that don't start at H1 and to support retroactive logging
  // (e.g. if H6 is logged after H7 was already logged, H7 is still considered the most recent hole)
  const getLastFinishedHole = (g) => {
    const gd = groupData[g.id];
    if (!gd || !gd.holeData) return { holeIdx: -1, diff: null };
    const order = getHoleOrder(g.startHole || 1);
    for (let s = 17; s >= 0; s--) {
      const hi = order[s];
      const hd = gd.holeData[hi];
      if (hd?.startTime && hd?.endTime) {
        return { holeIdx: hi, diff: computeHoleDiff(hd, hi, parTimes) };
      }
    }
    return { holeIdx: -1, diff: null };
  };

  // Group status now also accounts for the group ahead (same side, earlier tee time):
  // a group is only flagged if it's slower than the schedule allows *beyond* whatever
  // extra time the group ahead of it has already taken.
  const getGroupStatus = (g) => {
    const { holeIdx, diff } = getLastFinishedHole(g);
    if (diff === null) return "idle";
    const frontDiff = getFrontGroupDiffAtHole(groups, g, holeIdx, groupData, parTimes);
    return getRelativeStatus(diff, frontDiff);
  };

  const getHoleProgress = (g) => {
    const data = groupData[g.id];
    if (!data || !data.records) return 0;
    return data.records.filter(Boolean).length;
  };

  return (
    <div style={{ minHeight: "100vh", background: "#0d0f1a", fontFamily: "'IBM Plex Mono', monospace", color: "#eee" }}>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&family=Bebas+Neue&display=swap" rel="stylesheet" />

      {/* Grid, not flex: the left cell is allowed to shrink so the right cell can
          never be pushed off the edge of the screen. */}
      <div style={{ background: "#141626", borderBottom: "1px solid #2a2d4a", padding: "12px 16px", display: "grid", gridTemplateColumns: "1fr auto", alignItems: "center", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
          <button onClick={onBack} style={{ background: "#1a1d2e", border: "1px solid #4e9af144", color: "#4e9af1", cursor: "pointer", fontSize: 26, fontWeight: 700, borderRadius: 8, width: 44, height: 44, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>←</button>
          {/* BETA sits above the title, right edge flush with the final "D".
              The subtitle wraps rather than truncating, so it always reads fully. */}
          <div style={{ display: "flex", flexDirection: "column", minWidth: 0 }}>
            <div style={{ display: "inline-flex", flexDirection: "column", alignItems: "flex-end" }}>
              <span style={{ fontSize: 9, fontWeight: 700, letterSpacing: 1, color: "#ffd966", background: "#2a1a0066", border: "1px solid #ffd96655", borderRadius: 4, padding: "0 6px", height: 18, display: "inline-flex", alignItems: "center", lineHeight: 1, marginBottom: 2 }}>BETA</span>
              <div style={{ fontFamily: "'Bebas Neue'", fontSize: 22, letterSpacing: 4, color: "#4e9af1", whiteSpace: "nowrap" }}>⛳ DASHBOARD</div>
            </div>
            {tournamentName ? (
              <div style={{ fontSize: 11, color: "#eee", lineHeight: 1.4, marginTop: 2, fontWeight: 700 }}>
                {tournamentName}
                {roundLabel && (
                  <span style={{ color: "#6effa0" }}> — {roundLabel === "Q" ? "Round Q" : `Round ${roundLabel}`}</span>
                )}
              </div>
            ) : (
              <div style={{ fontSize: 10, color: "#8890b8", lineHeight: 1.4, marginTop: 2 }}>Golf Referee · Pace of Play System</div>
            )}
          </div>
        </div>
        {/* Right side: clock + user on top, Log out underneath — same layout as Setup.
            Capped width so rotating to landscape doesn't stretch the button. */}
        <div style={{ display: "flex", flexDirection: "column", gap: 8, flexShrink: 0, width: 170, alignItems: "flex-end" }}>
          {/* Online indicator hidden for now — re-enable by restoring this line:
              <OnlineUsers users={onlineUsers} currentUser={currentUser} /> */}
          <span style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 8 }}>
            <span style={{ fontSize: 13, color: "#8899cc" }}>⏱ {minToTime(now)}</span>
            {currentUser && (
              <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                <span style={{ fontSize: 12, color: "#8890b8" }}>👤</span>
                <span style={{ fontSize: 13, color: "#8899cc", fontWeight: 700, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{currentUser}</span>
              </span>
            )}
          </span>
          <div style={{ display: "flex", gap: 8, width: "100%" }}>
            <button onClick={onChangePassword}
              title="Change your password"
              style={{ background: "#0d0f1a", border: "1px solid #4e9af144", color: "#4e9af1", borderRadius: 10, height: 42, width: 46, cursor: "pointer", fontFamily: "inherit", fontSize: 15, flexShrink: 0 }}
            >🔑</button>
            <button onClick={onLogout}
              style={{ flex: 1, background: "#1a0d0d", border: "1px solid #ff707044", color: "#ff7070", borderRadius: 10, height: 42, cursor: "pointer", fontFamily: "inherit", fontSize: 14, fontWeight: 700, letterSpacing: 0.5, whiteSpace: "nowrap" }}
              onMouseEnter={e => e.currentTarget.style.borderColor = "#ff707088"}
              onMouseLeave={e => e.currentTarget.style.borderColor = "#ff707044"}
            >⏏ Log out</button>
          </div>
        </div>
      </div>

      {/* Where you are, and how to get somewhere else without leaving this page */}
      <div style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 10, margin: "12px 16px 0", padding: "10px 14px" }}>
        {(tournamentName || roundLabel) && (
          <div style={{ marginBottom: 10 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: "#c8ceee" }}>
              🏆 {tournamentName || "(untitled tournament)"}
              {roundLabel && <span style={{ color: "#6effa0" }}> — {roundLabel === "Q" ? "Round Q" : `Round ${roundLabel}`}</span>}
            </div>
            {hostVenue && <div style={{ fontSize: 12, color: "#8890b8", marginTop: 2 }}>{hostVenue}</div>}
          </div>
        )}
        {onOpenRound && (
          <RoundSelectorBar
            tournamentId={tournamentId}
            roundLabel={roundLabel}
            isAdmin={isTrueAdmin}
            onOpen={onOpenRound}
            compact
          />
        )}
      </div>

      {/* Action row — Summary / Export Data / Stopping Play, equal-size buttons */}
      <div style={{ background: "#141626", borderBottom: "1px solid #2a2d4a", padding: "10px 24px", display: "flex", gap: 10 }}>
        {onNavigateSummary && (
          <button
            onClick={onNavigateSummary}
            style={{
              flex: 1, background: "#0a1a2a", border: "1px solid #6effa088", color: "#6effa0",
              borderRadius: 8, padding: "10px 0", cursor: "pointer",
              fontFamily: "'IBM Plex Mono'", fontSize: 13, fontWeight: 700, letterSpacing: 1,
            }}
          >📈 Summary</button>
        )}
        <button
          onClick={() => setExportModal(true)}
          style={{
            flex: 1, background: "#0a1a2a", border: "1px solid #4e9af188", color: "#4e9af1",
            borderRadius: 8, padding: "10px 0", cursor: "pointer",
            fontFamily: "'IBM Plex Mono'", fontSize: 13, fontWeight: 700, letterSpacing: 1,
          }}
        >⬇ Export Data</button>
        {!isSuspended ? (
          <button onClick={openStop} style={{
            flex: 1, background: "#2a1500", border: "1px solid #ffd96688", color: "#ffd966",
            borderRadius: 8, padding: "10px 0", cursor: "pointer",
            fontFamily: "'IBM Plex Mono'", fontSize: 13, fontWeight: 700, letterSpacing: 1,
          }}>⏸ Stopping play</button>
        ) : (
          <button onClick={openResume} style={{
            flex: 1, background: "#0a2a10", border: "1px solid #6effa088", color: "#6effa0",
            borderRadius: 8, padding: "10px 0", cursor: "pointer",
            fontFamily: "'IBM Plex Mono'", fontSize: 13, fontWeight: 700, letterSpacing: 1,
            boxShadow: "0 0 12px #6effa033",
          }}>▶ Resume play</button>
        )}
      </div>

      {!isSuspended && suspensions.length > 0 && (
        <div style={{
          background: "#141210", borderBottom: "1px solid #ff996633",
          padding: "8px 24px", display: "flex", flexDirection: "column", alignItems: "flex-start", gap: 5,
        }}>
          <span style={{ color: "#ff9966", fontSize: 12, fontWeight: 700 }}>⏱ Time shift</span>
          {suspensions.map((s, i) => (
            <span key={i}
              onClick={() => onSuspendEdit && setEditSuspension({ idx: i, stopTime: s.stopTime, resumeTime: s.resumeTime })}
              title={onSuspendEdit ? "Tap to edit / delete" : undefined}
              style={{
                display: "grid",
                // Fixed columns keep every row's arrow, offset and pencil in line
                gridTemplateColumns: "26px 46px 12px 46px 62px 16px",
                alignItems: "center", gap: 4,
                fontSize: 12, color: "#aaa", background: "#0d0f1a", border: "1px solid #2a2d4a",
                borderRadius: 6, padding: "3px 8px", cursor: onSuspendEdit ? "pointer" : "default",
              }}>
              <span>#{i + 1}</span>
              <b style={{ color: "#ffd966" }}>{s.stopTime}</b>
              <span style={{ textAlign: "center" }}>→</span>
              <b style={{ color: "#6effa0" }}>{s.resumeTime}</b>
              <b style={{ color: "#ff9966", textAlign: "right" }}>+{s.offsetMin}min</b>
              <span style={{ color: "#4e9af1", textAlign: "right" }}>{onSuspendEdit ? "✏️" : ""}</span>
            </span>
          ))}
          <span style={{ fontSize: 13, color: "#ff9966", fontWeight: 700, marginTop: 2 }}>Total +{totalOffsetMin} min</span>
        </div>
      )}

      {/* Per-device hole focus — hide columns for holes this marshal doesn't supervise */}
      <div style={{ background: "#141626", borderBottom: "1px solid #2a2d4a", padding: "8px 24px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={() => setShowFocusPicker(v => !v)}
            style={{
              background: focusHoles.length ? "#1a4a8a" : "#0d0f1a",
              border: `1px solid ${focusHoles.length ? "#4e9af1" : "#2a2d4a"}`,
              color: focusHoles.length ? "#fff" : "#8890b8",
              borderRadius: 7, height: 34, padding: "0 12px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700, flexShrink: 0, display: "inline-flex", alignItems: "center", justifyContent: "center",
            }}>
            🎯 My ROTA{focusHoles.length ? ` (${focusHoles.length})` : ""}
          </button>
          <button onClick={toggleFitAll}
            title="Shrink columns to fit all 18 holes on one screen (for TD / CR)"
            style={{
              background: fitAllHoles ? "#1a4a8a" : "#0d0f1a",
              border: `1px solid ${fitAllHoles ? "#4e9af1" : "#2a2d4a"}`,
              color: fitAllHoles ? "#fff" : "#8890b8",
              borderRadius: 7, height: 34, padding: "0 12px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700, flexShrink: 0, display: "inline-flex", alignItems: "center", justifyContent: "center",
            }}>
            {fitAllHoles ? "🔍 Normal view" : "⛶ Fit 18 holes"}
          </button>
          {focusHoles.length > 0 && (
            <button onClick={() => saveFocusHoles([])}
              style={{ marginLeft: "auto", background: "#0d0f1a", border: "1px solid #ff707044", color: "#ff7070", borderRadius: 7, padding: "5px 10px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, flexShrink: 0 }}>
              ✕ Show all
            </button>
          )}
        </div>
        {focusHoles.length > 0 && (
          <div style={{ fontSize: 12, color: "#8890b8", marginTop: 6 }}>
            Showing {focusHoles.slice().sort((a, b) => a - b).map(h => `H${h + 1}`).join(", ")}
          </div>
        )}
      </div>

      {showFocusPicker && (
        <div style={{ background: "#0d0f1a", borderBottom: "1px solid #2a2d4a", padding: "12px 24px" }}>
          <div style={{ fontSize: 11, color: "#8890b8", marginBottom: 8 }}>Select the holes you supervise (none selected = show all) — saved on this device only</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(9, 1fr)", gap: 6, marginBottom: 10 }}>
            {Array.from({ length: 18 }, (_, i) => i).map(hi => {
              const on = focusHoles.includes(hi);
              return (
                <button key={hi}
                  onClick={() => saveFocusHoles(on ? focusHoles.filter(x => x !== hi) : [...focusHoles, hi])}
                  style={{
                    width: "100%", height: 34, borderRadius: 7, cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700, padding: 0,
                    background: on ? "#1a4a8a" : "#141626",
                    border: `1px solid ${on ? "#4e9af1" : "#2a2d4a"}`,
                    color: on ? "#fff" : "#8890b8",
                  }}>H{hi + 1}</button>
              );
            })}
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button onClick={() => saveFocusHoles(Array.from({ length: 9 }, (_, i) => i))}
              style={{ background: "#141626", border: "1px solid #2a2d4a", color: "#8890b8", borderRadius: 7, padding: "6px 12px", cursor: "pointer", fontFamily: "inherit", fontSize: 12 }}>H1–H9</button>
            <button onClick={() => saveFocusHoles(Array.from({ length: 9 }, (_, i) => i + 9))}
              style={{ background: "#141626", border: "1px solid #2a2d4a", color: "#8890b8", borderRadius: 7, padding: "6px 12px", cursor: "pointer", fontFamily: "inherit", fontSize: 12 }}>H10–H18</button>
            <button onClick={() => saveFocusHoles([])}
              style={{ background: "#141626", border: "1px solid #2a2d4a", color: "#8890b8", borderRadius: 7, padding: "6px 12px", cursor: "pointer", fontFamily: "inherit", fontSize: 12 }}>Show all</button>
            <button onClick={() => setShowFocusPicker(false)}
              style={{ marginLeft: "auto", background: "#1a4a2a", border: "1px solid #6effa066", color: "#6effa0", borderRadius: 7, padding: "6px 14px", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700 }}>Done</button>
          </div>
        </div>
      )}

      {/* Export Modal — copy each sheet's data as TSV, paste directly into Excel/Sheets.
          (File downloads are unreliable inside sandboxed artifact webviews, so
          clipboard copy is the one export path that works everywhere.) */}
      {exportModal && exportData && (
        <div style={{
          position: "fixed", inset: 0, background: "#000000aa", zIndex: 200,
          display: "flex", alignItems: "center", justifyContent: "center", padding: 20,
        }} onClick={() => setExportModal(false)}>
          <div
            style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 14, padding: 24, maxWidth: 560, width: "100%", maxHeight: "85vh", overflowY: "auto" }}
            onClick={e => e.stopPropagation()}
          >
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
              <div style={{ fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 2, color: "#4e9af1" }}>📊 EXPORT DATA</div>
              <button onClick={() => setExportModal(false)} style={{ background: "none", border: "none", color: "#9aa2c7", cursor: "pointer", fontSize: 18 }}>✕</button>
            </div>
            <div style={{ fontSize: 12, color: "#888", marginBottom: 18, lineHeight: 1.5 }}>
              Tap "Copy" for a sheet, then paste (long-press → Paste, or Ctrl/Cmd+V) into a new tab in Excel or Google Sheets — the columns will line up automatically.
            </div>

            {[
              { key: "summary", label: "Summary", tsv: exportData.summaryTSV, rows: exportData.summaryRows },
              { key: "detail", label: "Hole Details", tsv: exportData.detailTSV, rows: exportData.detailRows },
              { key: "logs", label: "Action Logs", tsv: exportData.logsTSV, rows: exportData.logRows },
            ].map(sheet => (
              <div key={sheet.key} style={{ background: "#0d0f1a", border: "1px solid #2a2d4a", borderRadius: 10, padding: 14, marginBottom: 12 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <div>
                    <div style={{ fontSize: 14, fontWeight: 700, color: "#eee" }}>{sheet.label}</div>
                    <div style={{ fontSize: 11, color: "#9aa2c7" }}>{sheet.rows.length} row{sheet.rows.length === 1 ? "" : "s"}</div>
                  </div>
                  <button
                    onClick={() => handleCopySheet(sheet.label, sheet.tsv)}
                    disabled={!sheet.rows.length}
                    style={{
                      background: exportCopied === sheet.label ? "#0a2a10" : "#1e2135",
                      border: `1px solid ${exportCopied === sheet.label ? "#6effa088" : "#2a2d4a"}`,
                      color: exportCopied === sheet.label ? "#6effa0" : (sheet.rows.length ? "#8899cc" : "#444"),
                      borderRadius: 7, padding: "6px 14px", cursor: sheet.rows.length ? "pointer" : "default",
                      fontFamily: "inherit", fontSize: 12, fontWeight: 700,
                    }}
                  >{exportCopied === sheet.label ? "✓ Copied" : "📋 Copy"}</button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Suspended Banner */}
      {isSuspended && (
        <div style={{
          background: "#1f0f00", borderBottom: "1px solid #ffd96666",
          padding: "10px 24px", display: "flex", alignItems: "center", gap: 12,
        }}>
          <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#ffd966", boxShadow: "0 0 8px #ffd966", animation: "pulse 1.5s infinite" }} />
          <span style={{ color: "#ffd966", fontWeight: 700, fontSize: 14 }}>⏸ Match paused</span>
          <span style={{ color: "#aaa", fontSize: 13 }}>Since <b style={{ color: "#eee" }}>{pendingStopTime}</b></span>
          <span style={{ color: "#8890b8", fontSize: 12 }}>— All groups paused together</span>
          <button onClick={openResume} style={{
            marginLeft: "auto", background: "#0a2a10", border: "1px solid #6effa088",
            color: "#6effa0", borderRadius: 7, padding: "5px 14px",
            cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700,
          }}>▶ Resume play</button>
          {onSuspendCancel && (
            <button onClick={() => { if (window.confirm("Cancel this stoppage?\n\nUse this if it was pressed by mistake — nothing is recorded and the schedule is not shifted.")) onSuspendCancel(); }}
              title="Pressed by mistake? Cancel without recording"
              style={{
                background: "#2a0a0a", border: "1px solid #ff707088",
                color: "#ff7070", borderRadius: 7, padding: "5px 12px",
                cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 700,
              }}>✕ Cancel</button>
          )}
        </div>
      )}

      {/* Suspension History Banner */}

      {editSuspension && (
        <div onClick={() => setEditSuspension(null)} style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1100, padding: 16 }}>
          <div onClick={e => e.stopPropagation()} style={{ background: "#141626", border: "1px solid #ff996688", borderRadius: 14, padding: 24, width: "100%", maxWidth: 320, boxShadow: "0 20px 60px #000" }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 2, color: "#ff9966", marginBottom: 16 }}>
              ⏱ Edit stop #{editSuspension.idx + 1}
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 20 }}>
              <div>
                <label style={{ fontSize: 11, color: "#8890b8", letterSpacing: 1 }}>Stop time</label>
                <input type="time" value={editSuspension.stopTime}
                  onChange={e => setEditSuspension(v => ({ ...v, stopTime: e.target.value }))}
                  style={{ display: "block", width: "100%", background: "#0d0f1a", border: "1px solid #2a2d4a", color: "#ffd966", borderRadius: 8, padding: "8px 10px", fontFamily: "inherit", fontSize: 15, outline: "none", marginTop: 4 }} />
              </div>
              <div>
                <label style={{ fontSize: 11, color: "#8890b8", letterSpacing: 1 }}>Resume time</label>
                <input type="time" value={editSuspension.resumeTime}
                  onChange={e => setEditSuspension(v => ({ ...v, resumeTime: e.target.value }))}
                  style={{ display: "block", width: "100%", background: "#0d0f1a", border: "1px solid #2a2d4a", color: "#6effa0", borderRadius: 8, padding: "8px 10px", fontFamily: "inherit", fontSize: 15, outline: "none", marginTop: 4 }} />
              </div>
              {(() => {
                const [sh, sm] = (editSuspension.stopTime || "0:00").split(":").map(Number);
                const [rh, rm] = (editSuspension.resumeTime || "0:00").split(":").map(Number);
                const off = Math.max(0, (rh * 60 + rm) - (sh * 60 + sm));
                return <div style={{ fontSize: 12, color: "#ff9966" }}>Schedule shifts by <b>+{off} min</b></div>;
              })()}
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => { onSuspendEdit(editSuspension.idx, editSuspension.stopTime, editSuspension.resumeTime); setEditSuspension(null); }}
                style={{ flex: 1, background: "#1a4a2a", border: "1px solid #6effa0", color: "#6effa0", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 15, letterSpacing: 2 }}>
                ✓ Save
              </button>
              <button onClick={() => { if (window.confirm(`Delete stoppage #${editSuspension.idx + 1}?\n\nThe schedule will be recalculated without this period.`)) { onSuspendDelete(editSuspension.idx); setEditSuspension(null); } }}
                style={{ background: "#2a0a0a", border: "1px solid #ff7070", color: "#ff7070", borderRadius: 8, padding: "10px 14px", cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700 }}>
                🗑
              </button>
              <button onClick={() => setEditSuspension(null)}
                style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "10px 14px", cursor: "pointer", fontFamily: "inherit", fontSize: 13 }}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      <div style={{ padding: "16px 20px" }}>
        {/* Alert Cards — split Morning / Afternoon */}
        {(() => {
          const alertGroups = groups.filter(g => {
            const gd = groupData[g.id];
            if (gd?.roundFinished === true) return false; // manually closed out
            const holeFromRecords = gd?.records?.filter(Boolean).length ?? 0;
            const holeFromData = (gd?.holeData ?? []).filter(h => h?.endTime).length;
            const hole = Math.max(holeFromRecords, holeFromData);
            if (hole >= 18) return false;
            const mnActive = gd?.mnActive === true;
            const tmActive = gd?.tmActive === true;
            // Only groups that have actually been flagged (WN / MN / TM / Bad Time)
            // belong here — running behind on its own isn't enough, since that's
            // already visible from the colours in the schedule table below.
            const hasLoggedAction = (gd?.actionLogs ?? []).length > 0;
            return mnActive || tmActive || hasLoggedAction;
          });

          // Sort groups with "late" status first, followed by "starting late", then on-time groups with a pending WN/MN
          const STATUS_RANK = { late: 3, warn: 2, ok: 1, idle: 0 };
          alertGroups.sort((a, b) => (STATUS_RANK[getGroupStatus(b)] ?? 0) - (STATUS_RANK[getGroupStatus(a)] ?? 0));

          // fallback: groups without a section use time < 12:00 = morning
          const getSection = (g) => {
            if (g.section) return g.section;
            const [h] = (g.startTime || "06:00").split(":").map(Number);
            return h < 12 ? "morning" : "afternoon";
          };

          const morningAlerts = alertGroups.filter(g => getSection(g) === "morning");
          const afternoonAlerts = alertGroups.filter(g => getSection(g) === "afternoon");
          const hasAfternoon = groups.some(g => getSection(g) === "afternoon");

          const renderGroupCard = (g) => {
            const status = getGroupStatus(g);
            const hole = getHoleProgress(g);
            const sch = schedules[g.id];
            const gd = groupData[g.id];
            const { holeIdx: lastHoleIdx, diff: nowDiff } = getLastFinishedHole(g);
            const lastEndTime = (() => {
              if (lastHoleIdx < 0) return null;
              const hd = gd.holeData[lastHoleIdx];
              if (!hd?.endTime) return null;
              if (hd.manualDiff !== undefined) {
                const deadline = (sch?.[lastHoleIdx] ?? 0) + (parTimes?.[lastHoleIdx] ?? 14);
                return minToTime(deadline - 1 + hd.manualDiff);
              }
              return hd.endTime;
            })();
            const diffColor = nowDiff === null ? "#555" : { ok: "#6effa0", warn: "#ffd966", late: "#ff7070" }[status];
            const mnActive = gd?.mnActive === true;
            const mnName = gd?.mnName ?? "";
            const tmActive = gd?.tmActive === true;
            const tmName = gd?.tmName ?? "";
            const tmTarget = gd?.tmTarget ?? "";
            return (
              <div key={g.id} onClick={() => setQuickRecord({ groupId: g.id, targetSlot: null })} style={{
                background: "#141626", border: `1px solid ${g.color}44`, borderRadius: 12,
                padding: "8px 10px", cursor: "pointer", transition: "box-shadow 0.15s",
                boxShadow: status === "late" ? `0 0 14px #ff707022` : "none",
                minWidth: 0, overflow: "hidden",
              }}
                onMouseEnter={e => { e.currentTarget.style.boxShadow = `0 4px 18px ${g.color}22`; }}
                onMouseLeave={e => { e.currentTarget.style.boxShadow = status === "late" ? `0 0 14px #ff707022` : "none"; }}
              >
                {/* One compact row: dot · name · progress · last hole diff · status */}
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <div style={{ width: 8, height: 8, borderRadius: "50%", background: g.color, flexShrink: 0 }} />
                  <div style={{ fontWeight: 700, fontSize: 13, whiteSpace: "nowrap", flexShrink: 0 }}>{g.name}</div>
                  <div style={{ fontSize: 11, color: "#9aa2c7", whiteSpace: "nowrap", flexShrink: 0 }}>H{hole}/18</div>
                  {nowDiff !== null ? (
                    <div style={{ display: "flex", alignItems: "baseline", gap: 4, whiteSpace: "nowrap", flexShrink: 0 }}>
                      <span style={{ fontSize: 11, color: "#8890b8" }}>H{lastHoleIdx + 1}</span>
                      <span style={{ fontFamily: "'Bebas Neue'", fontSize: 17, color: diffColor, lineHeight: 1 }}>{nowDiff > 0 ? `+${nowDiff}` : nowDiff}</span>
                      {lastEndTime && <span style={{ fontSize: 11, color: "#888" }}>{lastEndTime}</span>}
                    </div>
                  ) : (
                    <span style={{ fontSize: 11, color: "#8890b8", whiteSpace: "nowrap", flexShrink: 0 }}>No data yet</span>
                  )}
                  <div style={{ marginLeft: "auto", flexShrink: 0 }}><StatusBadge status={status} small /></div>
                </div>
                {(() => {
                  const logs = (gd?.actionLogs ?? []).map((l, idx) => ({ ...l, idx }));
                  if (logs.length === 0 && !mnActive && !tmActive) return null;
                  const items = summarizeStatusLogs(logs, mnActive, tmActive);
                  return (
                    <div style={{ marginTop: 5, display: "flex", flexWrap: "wrap", gap: 4 }}>
                      {items.map(it => (
                        <span key={it.key} onClick={(e) => { if (it.idx !== undefined) { e.stopPropagation(); setEditLogPopup({ groupId: g.id, idx: it.idx }); } }} style={{
                          display: "inline-flex", alignItems: "center", gap: 3,
                          background: `${logBg(it.type)}66`,
                          border: `1px solid ${logColor(it.type)}44`,
                          borderRadius: 4, padding: "2px 6px", fontSize: 11, fontWeight: 700,
                          color: logColor(it.type), lineHeight: 1.5, alignSelf: "flex-start",
                          cursor: it.idx !== undefined ? "pointer" : "default",
                        }}>
                          {it.label}
                          {it.idx !== undefined ? (
                            <button
                              onClick={(e) => { e.stopPropagation(); setDeleteLogConfirm({ groupId: g.id, idx: it.idx }); }}
                              title={it.deleteTitle || "Delete this log"}
                              style={{ background: "none", border: "none", color: "inherit", opacity: 0.75, cursor: "pointer", fontSize: 10, padding: 0, marginLeft: 1, lineHeight: 1 }}
                            >🗑</button>
                          ) : (it.type === "MN" || it.type === "TM") && (
                            <button
                              onClick={(e) => { e.stopPropagation(); setClearStatusConfirm({ groupId: g.id, type: it.type }); }}
                              title={`Clear all ${it.type} status`}
                              style={{ background: "none", border: "none", color: "inherit", opacity: 0.75, cursor: "pointer", fontSize: 10, padding: 0, marginLeft: 1, lineHeight: 1 }}
                            >🗑</button>
                          )}
                        </span>
                      ))}
                    </div>
                  );
                })()}
              </div>
            );
          };

          if (alertGroups.length === 0 && !hasAfternoon) return (
            <div style={{ textAlign: "center", padding: "18px 0 8px", color: "#3a3d5a", fontSize: 13, letterSpacing: 1 }}>
              ✓ No groups flagged (WN / MN / TM).
            </div>
          );

          if (alertGroups.length === 0) return (
            <div style={{ textAlign: "center", padding: "18px 0 8px", color: "#3a3d5a", fontSize: 13, letterSpacing: 1 }}>
              ✓ No groups flagged (WN / MN / TM).
            </div>
          );

          if (!hasAfternoon) return (
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {alertGroups.map(renderGroupCard)}
            </div>
          );

          return (
            <div>
              {morningAlerts.length > 0 && (
                <div style={{ marginBottom: 20 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                    <div style={{ fontSize: 12, color: "#8899cc", fontWeight: 700, letterSpacing: 2 }}>🌅 MORNING</div>
                    <div style={{ flex: 1, height: 1, background: "#2a2d4a" }} />
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {morningAlerts.map(renderGroupCard)}
                  </div>
                </div>
              )}
              {afternoonAlerts.length > 0 && (
                <div>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                    <div style={{ fontSize: 12, color: "#ffd966", fontWeight: 700, letterSpacing: 2 }}>☀️ AFTERNOON</div>
                    <div style={{ flex: 1, height: 1, background: "#2a2d4a" }} />
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {afternoonAlerts.map(renderGroupCard)}
                  </div>
                </div>
              )}
              {morningAlerts.length === 0 && afternoonAlerts.length === 0 && (
                <div style={{ textAlign: "center", padding: "18px 0 8px", color: "#3a3d5a", fontSize: 13, letterSpacing: 1 }}>
                  ✓ No groups flagged (WN / MN / TM).
                </div>
              )}
            </div>
          );
        })()}

        {/* Summary Table — split Morning/Afternoon and H1/H10 */}
        {(() => {
          const getSection = (g) => {
            if (g.section) return g.section;
            const [h] = (g.startTime || "06:00").split(":").map(Number);
            return h < 12 ? "morning" : "afternoon";
          };
          const hasAfternoon = groups.some(g => getSection(g) === "afternoon");
          const sections = hasAfternoon ? ["morning", "afternoon"] : ["morning"];

          const renderTable = (col, startHole, sectionKey) => {
            if (col.length === 0) return null;
            const order = getHoleOrder(startHole);
            const meta = getStartHoleMeta(startHole);
            const colColor = meta.color;
            const holeLabel = meta.label;
            const tableKey = `${sectionKey}-${startHole}`;
            const isCollapsed = !!collapsedTables[tableKey];
            // Fit-18 view: shrink every column so all 18 holes sit on one screen.
            const th = fitAllHoles ? { ...thStyle, padding: "4px 0", fontSize: 9, minWidth: 0 } : thStyle;
            const td = fitAllHoles ? { ...tdStyle, padding: "3px 0" } : tdStyle;
            const turnGapW = fitAllHoles ? 8 : 0;
            const nameColW = fitAllHoles ? 26 : 80;
            const startColW = fitAllHoles ? 28 : 56;
            return (
              <div key={tableKey} style={{ background: "#141626", border: `1px solid ${colColor}22`, borderRadius: 12, marginTop: 16, overflow: "hidden" }}>
                <div
                  onClick={() => setCollapsedTables(prev => ({ ...prev, [tableKey]: !prev[tableKey] }))}
                  style={{ padding: "12px 16px", borderBottom: isCollapsed ? "none" : "1px solid #2a2d4a", fontSize: 12, color: colColor, letterSpacing: 2, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "space-between", userSelect: "none" }}
                >
                  <span>📋 {holeLabel}</span>
                  <span style={{ fontSize: 14, transform: isCollapsed ? "rotate(-90deg)" : "none", transition: "transform 0.15s" }}>▾</span>
                </div>
                {!isCollapsed && (
                <div style={{ overflowX: fitAllHoles ? "hidden" : "auto" }}>
                  <table style={{ borderCollapse: "collapse", fontSize: 13, ...(fitAllHoles ? { width: "100%", tableLayout: "fixed" } : {}) }}>
                    <thead>
                      <tr style={{ background: "#0d0f1a" }}>
                        <th style={{ ...th, position: fitAllHoles ? "static" : "sticky", left: 0, zIndex: 2, background: "#0d0f1a", width: nameColW, minWidth: nameColW }}>{fitAllHoles ? "Grp" : "Group"}</th>
                        <th style={{ ...th, color: colColor, position: fitAllHoles ? "static" : "sticky", left: nameColW, zIndex: 2, background: "#0d0f1a", width: startColW, minWidth: startColW, borderRight: "1px solid #2a2d4a" }}>{fitAllHoles ? "Time" : "Start"}</th>
                        {withTurnGap(order.map((hi, i) => (
                          focusHoles.length && !focusHoles.includes(hi) ? null : (
                            <th key={hi} style={!fitAllHoles && i === 9 ? { ...th, borderLeft: `2px solid ${colColor}88` } : th}>{fitAllHoles ? hi + 1 : `H${hi + 1}`}</th>
                          )
                        )), turnGapW, `h-${tableKey}`, "th")}
                      </tr>
                    </thead>
                    <tbody>
                      {col.map(g => {
                        const data = groupData[g.id];
                        const sch = schedules[g.id];
                        const allLogs = data?.actionLogs ?? [];
                        const hasWN = allLogs.some(l => l.type === "WN");
                        const hasMN = allLogs.some(l => l.type === "MN");
                        const hasTM = allLogs.some(l => l.type === "TM");
                        const mnActiveNow = data?.mnActive === true;
                        const tmActiveNow = data?.tmActive === true;
                        // Slot where the MN/TM "coming up" preview should be shown: exactly one
                        // hole past the most recently logged MN/TM entry (see group-detail view
                        // for the same logic), so it follows forward one hole at a time.
                        const slotOfHole = {};
                        order.forEach((hIdx, s) => { slotOfHole[hIdx] = s; });
                        const lastMNSlot = allLogs.reduce((mx, l) => (l.type === "MN" ? Math.max(mx, slotOfHole[l.holeIdx] ?? -1) : mx), -1);
                        const lastTMSlot = allLogs.reduce((mx, l) => (l.type === "TM" ? Math.max(mx, slotOfHole[l.holeIdx] ?? -1) : mx), -1);
                        return (
                          <tr key={g.id}>
                            <td onClick={() => setQuickRecord({ groupId: g.id, targetSlot: null })}
                              style={{ ...td, color: g.color, fontWeight: 700, cursor: "pointer", transition: "background 0.15s", position: fitAllHoles ? "static" : "sticky", left: 0, zIndex: 1, background: "#141626", width: nameColW, minWidth: nameColW }}
                              onMouseEnter={e => e.currentTarget.style.background = `${g.color}22`}
                              onMouseLeave={e => e.currentTarget.style.background = "#141626"}
                            >
                              <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: fitAllHoles ? 0 : 5 }}>
                                <span>{fitAllHoles ? g.name.replace(/^\s*group\s*/i, "") : g.name}</span>
                                {!fitAllHoles && <span style={{ fontSize: 11, color: "#8890b8" }}>›</span>}
                              </div>
                              <div style={{ display: fitAllHoles ? "none" : "flex", gap: 3, marginTop: 4, justifyContent: "center", flexWrap: "wrap" }}>
                                {data?.roundFinished === true && (
                                  <span
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      if (window.confirm(`Undo "round finished" for ${g.name}?\n\nThis group will appear in the alert list again.`)) {
                                        onUpdateGroupData(g.id, { roundFinished: false });
                                      }
                                    }}
                                    title="Tap to undo the finished status"
                                    style={{ fontSize: fitAllHoles ? 9 : 11, fontWeight: 700, color: "#6effa0", background: "#0a2a1066", border: "1px solid #6effa044", borderRadius: 4, padding: "1px 5px", cursor: "pointer" }}
                                  >{fitAllHoles ? "🏁" : "🏁 FINISHED"}</span>
                                )}
                                {hasWN && <span style={{ fontSize: fitAllHoles ? 9 : 11, fontWeight: 700, color: "#ffd966", background: "#2a1a0066", border: "1px solid #ffd96644", borderRadius: 4, padding: fitAllHoles ? "0 2px" : "1px 5px" }}>{fitAllHoles ? "W" : "WN"}</span>}
                                {mnActiveNow ? (
                                  <span style={{ fontSize: fitAllHoles ? 9 : 11, fontWeight: 700, color: "#4e9af1", background: "#001a2a66", border: "1px solid #4e9af144", borderRadius: 4, padding: fitAllHoles ? "0 2px" : "1px 5px" }}>{fitAllHoles ? "M" : "👁 MN"}</span>
                                ) : hasMN && (
                                  <span style={{ fontSize: fitAllHoles ? 9 : 11, fontWeight: 700, color: "#888", background: "#1a1a1a66", border: "1px solid #55555544", borderRadius: 4, padding: fitAllHoles ? "0 2px" : "1px 5px" }}>{fitAllHoles ? "×M" : "✕ MN"}</span>
                                )}
                                {tmActiveNow ? (
                                  <span style={{ fontSize: fitAllHoles ? 9 : 11, fontWeight: 700, color: "#ff6ec7", background: "#2a002066", border: "1px solid #ff6ec744", borderRadius: 4, padding: fitAllHoles ? "0 2px" : "1px 5px" }}>{fitAllHoles ? "T" : "⏱ TM"}</span>
                                ) : hasTM && (
                                  <span style={{ fontSize: fitAllHoles ? 9 : 11, fontWeight: 700, color: "#888", background: "#1a1a1a66", border: "1px solid #55555544", borderRadius: 4, padding: fitAllHoles ? "0 2px" : "1px 5px" }}>{fitAllHoles ? "×T" : "✕ TM"}</span>
                                )}
                              </div>
                            </td>
                            <td
                              onClick={e => e.stopPropagation()}
                              style={{ ...td, color: colColor, fontWeight: 700, position: fitAllHoles ? "static" : "sticky", left: nameColW, zIndex: 1, background: "#141626", width: startColW, minWidth: startColW, borderRight: "1px solid #2a2d4a" }}
                            >
                              {fitAllHoles ? (
                                <div style={{ fontSize: 11, lineHeight: 1.05 }}>
                                  <div>{g.startTime.slice(0, 2)}</div>
                                  <div>{g.startTime.slice(3)}</div>
                                  {(data?.delayMin ?? 0) > 0 && (
                                    <div style={{ fontSize: 9, color: "#eee", marginTop: 1 }}>+{data.delayMin}</div>
                                  )}
                                </div>
                              ) : <div>{g.startTime}</div>}
                              {!fitAllHoles && (() => {
                                const delayMin = data?.delayMin ?? 0;
                                const setDelay = (val) => {
                                  const n = Math.max(0, parseInt(val, 10) || 0);
                                  onUpdateGroupData(g.id, { delayMin: n });
                                };
                                return (
                                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 2, marginTop: 4 }}>
                                    <button
                                      onClick={() => setDelay(delayMin - 1)}
                                      style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#aaa", borderRadius: 4, width: 16, height: 16, lineHeight: "14px", cursor: "pointer", fontSize: 12, fontFamily: "inherit", padding: 0 }}
                                    >−</button>
                                    <input
                                      type="number" min="0" value={delayMin}
                                      onChange={e => setDelay(e.target.value)}
                                      onClick={e => e.stopPropagation()}
                                      style={{ width: 28, background: "#1e2135", border: `1px solid ${delayMin > 0 ? "#ffd96666" : "#2a2d4a"}`, color: "#ffd966", fontFamily: "'IBM Plex Mono'", fontSize: 11, fontWeight: 700, textAlign: "center", borderRadius: 4, padding: "1px 0", outline: "none" }}
                                    />
                                    <button
                                      onClick={() => setDelay(delayMin + 1)}
                                      style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#aaa", borderRadius: 4, width: 16, height: 16, lineHeight: "14px", cursor: "pointer", fontSize: 12, fontFamily: "inherit", padding: 0 }}
                                    >+</button>
                                  </div>
                                );
                              })()}
                            </td>
                            {withTurnGap(order.map((hi, slot) => {
                              if (focusHoles.length && !focusHoles.includes(hi)) return null;
                              const hd = data?.holeData?.[hi];
                              const startTime = hd?.startTime;
                              const endTime = hd?.endTime;
                              const holeLogsRaw = (data?.actionLogs ?? []).filter(l => l.holeIdx === hi);
                              const holeLogs = holeLogsRaw.filter(l => !l.off || !holeLogsRaw.some(o => o !== l && o.type === l.type && !o.off));
                              const handleHoleClick = () => setQuickRecord({ groupId: g.id, targetSlot: slot });
                              const deadline = (sch?.[hi] ?? 0) + (parTimes?.[hi] ?? 14);
                              if (!endTime || !startTime) {
                                const showMnPreview = mnActiveNow && slot === lastMNSlot + 1 && !holeLogs.some(l => l.type === "MN");
                                const showTmPreview = tmActiveNow && slot === lastTMSlot + 1 && !holeLogs.some(l => l.type === "TM");
                                return (
                                  <td key={hi} onClick={handleHoleClick} style={{ ...td, color: "#666f99", cursor: "pointer", transition: "background 0.15s", ...(!fitAllHoles && slot === 9 ? { borderLeft: `2px solid ${colColor}88` } : {}) }}
                                    onMouseEnter={e => e.currentTarget.style.background = "#ffffff08"}
                                    onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                                  >
                                    {fitAllHoles ? (() => {
                                      const t = minToTime(deadline);
                                      return (
                                        <div style={{ fontSize: 11, fontWeight: 700, color: "#9aa2c7", lineHeight: 1.05 }}>
                                          <div>{t.slice(0, 2)}</div>
                                          <div>{t.slice(3)}</div>
                                        </div>
                                      );
                                    })() : (
                                      <div style={{ fontSize: 16, fontWeight: 700, color: "#9aa2c7" }}>{minToTime(deadline)}</div>
                                    )}
                                    {holeLogs.map((l, li) => (
                                      <div key={li} style={{ marginTop: fitAllHoles ? 1 : 3, fontSize: fitAllHoles ? 9 : 11, fontWeight: 700, color: logColor(l.type), background: `${logBg(l.type)}55`, borderRadius: 3, padding: fitAllHoles ? "0 2px" : "1px 4px" }}>
                                        {fitAllHoles
                                          ? shortLogLabel(l)
                                          : (l.badTime ? `⚡ BT ${l.target || ""}${l.name ? ` - ${l.name}` : ""}` : l.off ? `Off ${l.type}${l.name ? ` - ${l.name}` : ""}` : <>{l.type}{l.name ? ` - ${l.name}` : ""}</>)}
                                      </div>
                                    ))}
                                    {showMnPreview && (
                                      <div style={{ marginTop: fitAllHoles ? 1 : 3, fontSize: fitAllHoles ? 9 : 11, fontWeight: 700, color: "#4e9af1", background: "#001a2a55", border: "1px dashed #4e9af155", borderRadius: 3, padding: fitAllHoles ? "0 2px" : "1px 4px" }}>
                                        {fitAllHoles ? "M" : <>👁 MN{data?.mnName ? ` - ${data.mnName}` : ""}</>}
                                      </div>
                                    )}
                                    {showTmPreview && (
                                      <div style={{ marginTop: fitAllHoles ? 1 : 3, fontSize: fitAllHoles ? 9 : 11, fontWeight: 700, color: "#ff6ec7", background: "#2a002055", border: "1px dashed #ff6ec755", borderRadius: 3, padding: fitAllHoles ? "0 2px" : "1px 4px" }}>
                                        {fitAllHoles ? "T" : <>⏱ TM {data?.tmName ? `- ${data.tmName}` : ""}</>}
                                      </div>
                                    )}
                                  </td>
                                );
                              }
                              const diff = computeHoleDiff(hd, hi, parTimes);
                              const frontDiff = getFrontGroupDiffAtHole(groups, g, hi, groupData, parTimes);
                              const relativeDiff = diff === null ? null : diff - Math.max(frontDiff ?? 0, 0);
                              const color = diffColor(relativeDiff);
                              const cellBg = `${color}26`;
                              const cellBgHover = `${color}3d`;
                              return (
                                <td key={hi} onClick={handleHoleClick}
                                  style={{ ...td, color, fontWeight: 700, cursor: "pointer", transition: "background 0.15s", background: cellBg, ...(!fitAllHoles && slot === 9 ? { borderLeft: `2px solid ${colColor}88` } : {}) }}
                                  onMouseEnter={e => e.currentTarget.style.background = cellBgHover}
                                  onMouseLeave={e => e.currentTarget.style.background = cellBg}
                                >
                                  <div style={{ fontSize: fitAllHoles ? 11 : 20, lineHeight: 1.2 }}>{diff > 0 ? `+${diff}` : diff}</div>
                                  {holeLogs.map((l, li) => (
                                    <div key={li} style={{ marginTop: fitAllHoles ? 1 : 3, fontSize: fitAllHoles ? 9 : 11, fontWeight: 700, color: logColor(l.type), background: `${logBg(l.type)}55`, borderRadius: 3, padding: fitAllHoles ? "0 2px" : "1px 4px", whiteSpace: "nowrap" }}>
                                      {fitAllHoles
                                        ? shortLogLabel(l)
                                        : (l.badTime ? `⚡ BT ${l.target || ""}${l.name ? ` - ${l.name}` : ""}` : l.off ? `Off ${l.type}${l.name ? ` - ${l.name}` : ""}` : <>{l.type}{l.name ? ` - ${l.name}` : ""}</>)}
                                    </div>
                                  ))}
                                </td>
                              );
                            }), turnGapW, `b-${tableKey}-${g.id}`, "td")}
                          </tr>
                        );
                      })}
                    </tbody>
                    {/* Repeat the hole numbers at the bottom — with long group lists
                        the header scrolls out of view, so this keeps it obvious which
                        hole each column belongs to. */}
                    <tfoot>
                      <tr style={{ background: "#0d0f1a" }}>
                        <th style={{ ...th, position: fitAllHoles ? "static" : "sticky", left: 0, zIndex: 2, background: "#0d0f1a", width: nameColW, minWidth: nameColW, borderBottom: "none", borderTop: "1px solid #2a2d4a" }}>{fitAllHoles ? "Grp" : "Group"}</th>
                        <th style={{ ...th, color: colColor, position: fitAllHoles ? "static" : "sticky", left: nameColW, zIndex: 2, background: "#0d0f1a", width: startColW, minWidth: startColW, borderRight: "1px solid #2a2d4a", borderBottom: "none", borderTop: "1px solid #2a2d4a" }}>{fitAllHoles ? "Time" : "Start"}</th>
                        {withTurnGap(order.map((hi, i) => (
                          focusHoles.length && !focusHoles.includes(hi) ? null : (
                            <th key={hi} style={{ ...th, borderBottom: "none", borderTop: "1px solid #2a2d4a", ...(!fitAllHoles && i === 9 ? { borderLeft: `2px solid ${colColor}88` } : {}) }}>{fitAllHoles ? hi + 1 : `H${hi + 1}`}</th>
                          )
                        )), turnGapW, `f-${tableKey}`, "th")}
                      </tr>
                    </tfoot>
                  </table>
                </div>
                )}
              </div>
            );
          };

          return sections.map(sec => {
            const secGroups = groups.filter(g => getSection(g) === sec);
            if (secGroups.length === 0) return null;
            const startHolesInSec = Array.from(new Set(secGroups.map(g => g.startHole || 1))).sort((a, b) => a - b);
            const secLabel = sec === "morning" ? "🌅 MORNING SECTION" : "☀️ AFTERNOON SECTION";
            const secColor = sec === "morning" ? "#8899cc" : "#ffd966";
            const secBorder = sec === "morning" ? "#8899cc22" : "#ffd96622";
            return (
              <div key={sec} style={{ marginTop: 28 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 4 }}>
                  <div style={{ fontSize: 13, color: secColor, fontWeight: 700, letterSpacing: 3, whiteSpace: "nowrap" }}>{secLabel}</div>
                  <div style={{ flex: 1, height: 1, background: secBorder }} />
                </div>
                {startHolesInSec.map(startHole =>
                  renderTable(secGroups.filter(g => (g.startHole || 1) === startHole), startHole, sec)
                )}
              </div>
            );
          });
        })()}
      </div>

      {/* Stop Modal */}
      {suspendModal === "stop" && (
        <div style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }}>
          <div style={{ background: "#141626", border: "1px solid #ffd96688", borderRadius: 14, padding: 28, minWidth: 320, boxShadow: "0 20px 60px #000", fontFamily: "'IBM Plex Mono', monospace" }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 24, letterSpacing: 3, color: "#ffd966", marginBottom: 6 }}>⏸ Stopping play</div>
            <div style={{ fontSize: 13, color: "#888", marginBottom: 20 }}>All groups will be paused together; hole finish times will shift when play resumes</div>
            <div style={{ fontSize: 13, color: "#ffd966", marginBottom: 8 }}>Stop time</div>
            <TimeInput value={suspendStopInput} onChange={setSuspendStopInput} color="#ffd966" />
            <div style={{ display: "flex", gap: 10, marginTop: 22 }}>
              <button onClick={confirmStop} style={{ flex: 1, background: "#2a1500", border: "1px solid #ffd966", color: "#ffd966", borderRadius: 8, padding: "11px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 17, letterSpacing: 2 }}>
                ✓ Confirm pause
              </button>
              <button onClick={() => setSuspendModal(false)} style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "11px 16px", cursor: "pointer", fontFamily: "inherit", fontSize: 14 }}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Resume Modal */}
      {suspendModal === "resume" && (
        <div style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }}>
          <div style={{ background: "#141626", border: "1px solid #6effa088", borderRadius: 14, padding: 28, minWidth: 320, boxShadow: "0 20px 60px #000", fontFamily: "'IBM Plex Mono', monospace" }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 24, letterSpacing: 3, color: "#6effa0", marginBottom: 6 }}>▶ Resume play</div>
            <div style={{ fontSize: 13, color: "#888", marginBottom: 6 }}>
              Paused since <b style={{ color: "#ffd966" }}>{pendingStopTime}</b>
            </div>
            <div style={{ fontSize: 13, color: "#888", marginBottom: 20 }}>The system will add the pause duration to the schedule for every hole and group</div>
            <div style={{ fontSize: 13, color: "#6effa0", marginBottom: 8 }}>Resume time</div>
            <TimeInput value={suspendResumeInput} onChange={setSuspendResumeInput} color="#6effa0" />
            {(() => {
              const [sh, sm] = pendingStopTime.split(":").map(Number);
              const [rh, rm] = suspendResumeInput.split(":").map(Number);
              const offset = Math.max(0, (rh * 60 + rm) - (sh * 60 + sm));
              return (
                <div style={{ marginTop: 14, background: "#0d0f1a", borderRadius: 8, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 13, color: "#8890b8" }}>Time to add to every hole</span>
                  <span style={{ color: "#ff9966", fontWeight: 700, fontSize: 20, fontFamily: "'Bebas Neue'", letterSpacing: 2 }}>+{offset} min</span>
                </div>
              );
            })()}
            <div style={{ display: "flex", gap: 10, marginTop: 22 }}>
              <button onClick={confirmResume} style={{ flex: 1, background: "#0a2a10", border: "1px solid #6effa0", color: "#6effa0", borderRadius: 8, padding: "11px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 17, letterSpacing: 2 }}>
                ✓ Confirm resume
              </button>
              <button onClick={() => setSuspendModal(false)} style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "11px 16px", cursor: "pointer", fontFamily: "inherit", fontSize: 14 }}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Delete log confirmation popup */}
      {deleteLogConfirm && (() => {
        const gLogs = groupData[deleteLogConfirm.groupId]?.actionLogs ?? [];
        const l = gLogs[deleteLogConfirm.idx];
        if (!l) return null;
        const gName = groups.find(g => g.id === deleteLogConfirm.groupId)?.name ?? "";
        return (
          <div style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1100 }}>
            <div style={{ background: "#141626", border: "1px solid #ff707088", borderRadius: 14, padding: 28, minWidth: 280, maxWidth: 340, boxShadow: "0 20px 60px #000" }}>
              <div style={{ fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 3, color: "#ff7070", marginBottom: 10 }}>Delete {l.off ? `Off ${l.type}` : l.type}?</div>
              <div style={{ fontSize: 13, color: "#aaa", marginBottom: 20, lineHeight: 1.6 }}>
                {gName && <>{gName} — </>}Are you sure?
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => deleteLogAt(deleteLogConfirm.groupId, deleteLogConfirm.idx)}
                  style={{ flex: 1, background: "#2a0a0a", border: "1px solid #ff7070", color: "#ff7070", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 15, letterSpacing: 2 }}>
                  ✓ Yes, delete
                </button>
                <button onClick={() => setDeleteLogConfirm(null)}
                  style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "10px 16px", cursor: "pointer", fontFamily: "inherit", fontSize: 14 }}>
                  Cancel
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {editLogPopup && (() => {
        const gLogs = groupData[editLogPopup.groupId]?.actionLogs ?? [];
        const l = gLogs[editLogPopup.idx];
        if (!l) return null;
        const gName = groups.find(g => g.id === editLogPopup.groupId)?.name ?? "";
        const saveEdit = (updated) => {
          const nextLogs = gLogs.map((x, i2) => i2 === editLogPopup.idx ? { ...x, ...updated } : x);
          onUpdateGroupData(editLogPopup.groupId, { actionLogs: nextLogs });
          setEditLogPopup(null);
        };
        return (
          <div onClick={() => setEditLogPopup(null)} style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1100, padding: 16 }}>
            <div onClick={e => e.stopPropagation()} style={{ background: "#141626", border: `1px solid ${logColor(l.type)}88`, borderRadius: 14, padding: 24, width: "100%", maxWidth: 320, boxShadow: "0 20px 60px #000" }}>
              <div style={{ fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 2, color: logColor(l.type), marginBottom: 4 }}>
                Edit {l.badTime ? "Bad Time" : l.off ? `Off ${l.type}` : l.type}
              </div>
              <div style={{ fontSize: 12, color: "#8890b8", marginBottom: 18 }}>{gName}</div>

              <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 20 }}>
                <div>
                  <label style={{ fontSize: 11, color: "#8890b8", letterSpacing: 1 }}>Hole</label>
                  <input type="number" min="1" max="18" defaultValue={l.holeIdx + 1}
                    onChange={e => { l._newHole = Math.min(18, Math.max(1, Number(e.target.value) || 1)); }}
                    style={{ display: "block", width: "100%", background: "#0d0f1a", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 8, padding: "8px 10px", fontFamily: "inherit", fontSize: 14, outline: "none", marginTop: 4 }} />
                </div>
                {(l.target || l.badTime) && (
                  <div>
                    <label style={{ fontSize: 11, color: "#8890b8", letterSpacing: 1 }}>Player</label>
                    <input defaultValue={l.target || ""} placeholder="e.g. P2"
                      onChange={e => { l._newTarget = e.target.value; }}
                      style={{ display: "block", width: "100%", background: "#0d0f1a", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 8, padding: "8px 10px", fontFamily: "inherit", fontSize: 14, outline: "none", marginTop: 4 }} />
                  </div>
                )}
                <div>
                  <label style={{ fontSize: 11, color: "#8890b8", letterSpacing: 1 }}>Recorded by</label>
                  <input defaultValue={l.name || ""}
                    onChange={e => { l._newName = e.target.value; }}
                    style={{ display: "block", width: "100%", background: "#0d0f1a", border: "1px solid #2a2d4a", color: "#eee", borderRadius: 8, padding: "8px 10px", fontFamily: "inherit", fontSize: 14, outline: "none", marginTop: 4 }} />
                </div>
              </div>

              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => saveEdit({
                  holeIdx: (l._newHole ?? (l.holeIdx + 1)) - 1,
                  ...(l._newTarget !== undefined ? { target: l._newTarget } : {}),
                  ...(l._newName !== undefined ? { name: l._newName } : {}),
                })}
                  style={{ flex: 1, background: "#1a4a2a", border: "1px solid #6effa0", color: "#6effa0", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 15, letterSpacing: 2 }}>
                  ✓ Save
                </button>
                <button onClick={() => setEditLogPopup(null)}
                  style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "10px 16px", cursor: "pointer", fontFamily: "inherit", fontSize: 14 }}>
                  Cancel
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {clearStatusConfirm && (() => {
        const gName = groups.find(g => g.id === clearStatusConfirm.groupId)?.name ?? "";
        return (
          <div style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1100 }}>
            <div style={{ background: "#141626", border: "1px solid #ff707088", borderRadius: 14, padding: 28, minWidth: 280, maxWidth: 340, boxShadow: "0 20px 60px #000" }}>
              <div style={{ fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 3, color: "#ff7070", marginBottom: 10 }}>Clear {clearStatusConfirm.type} status?</div>
              <div style={{ fontSize: 13, color: "#aaa", marginBottom: 20, lineHeight: 1.6 }}>
                {gName && <>{gName} — </>}Are you sure?
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => clearStatusFor(clearStatusConfirm.groupId, clearStatusConfirm.type)}
                  style={{ flex: 1, background: "#2a0a0a", border: "1px solid #ff7070", color: "#ff7070", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 15, letterSpacing: 2 }}>
                  ✓ Yes, clear
                </button>
                <button onClick={() => setClearStatusConfirm(null)}
                  style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "10px 16px", cursor: "pointer", fontFamily: "inherit", fontSize: 14 }}>
                  Cancel
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {quickRecord && (() => {
        const qGroup = groups.find(x => x.id === quickRecord.groupId);
        if (!qGroup) return null;
        const gd = groupData[quickRecord.groupId] || {};
        return (
          <div onClick={() => setQuickRecord(null)} style={{ position: "fixed", inset: 0, zIndex: 1200, background: "#000000cc", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "24px 16px", overflowY: "auto", WebkitOverflowScrolling: "touch" }}>
            <div onClick={e => e.stopPropagation()} style={{ background: "#0d0f1a", border: "1px solid #2a2d4a", borderRadius: 16, padding: 16, width: "100%", maxWidth: 440, boxShadow: "0 24px 70px #000a" }}>
              <GroupMonitor
                key={quickRecord.groupId}
                group={{
                  ...qGroup,
                  records: gd.records,
                  holeData: gd.holeData,
                  currentHole: quickRecord.targetSlot !== null && quickRecord.targetSlot !== undefined ? quickRecord.targetSlot : gd.currentHole,
                  actionLogs: gd.actionLogs,
                  mnActive: gd.mnActive,
                  mnName: gd.mnName,
                  tmActive: gd.tmActive,
                  tmName: gd.tmName,
                  tmTarget: gd.tmTarget,
                  delayMin: gd.delayMin,
                }}
                pars={pars}
                parTimes={parTimes}
                playersPerGroup={playersPerGroup}
                schedule={schedules[quickRecord.groupId]}
                onUpdate={(update) => onUpdateGroupData(quickRecord.groupId, update)}
                onBack={() => setQuickRecord(null)}
                currentUser={currentUser}
                isSuspended={isSuspended}
                suspensions={suspensions}
                totalOffsetMin={totalOffsetMin}
                pendingStopTime={pendingStopTime}
                hideLog={true}
                onRecorded={() => setQuickRecord(null)}
                statusOverride={getGroupStatus(qGroup)}
                compact={true}
              />
            </div>
          </div>
        );
      })()}
    </div>
  );
}

// ─── Login Screen ─────────────────────────────────────────────────────────────
const DEFAULT_USERS = [
  { username: "admin", password: "0000", isAdmin: true },
];
const DEFAULT_RESET_PASSWORD = "1234";

// Lets anyone change their own password. Asks for the current one first so a
// borrowed, unlocked phone can't be used to take over someone's account.
function ChangePasswordModal({ currentUser, users, onSave, onClose }) {
  const [current, setCurrent] = useState("");
  const [next, setNext] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState("");
  const [show, setShow] = useState(false);

  const submit = () => {
    const me = users.find(u => u.username === currentUser);
    if (!me) { setError("Account not found"); return; }
    if (current !== me.password) { setError("Current password is incorrect"); return; }
    if (!next) { setError("Please enter a new password"); return; }
    if (next.length < 4) { setError("New password must be at least 4 characters"); return; }
    if (next === current) { setError("New password must be different from the current one"); return; }
    if (next !== confirm) { setError("The two new passwords do not match"); return; }
    onSave(next);
  };

  const field = {
    width: "100%", background: "#0d0f1a", border: "1px solid #2a2d4a", color: "#eee",
    borderRadius: 8, padding: "10px 12px", fontFamily: "inherit", fontSize: 16, outline: "none",
  };

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 1200, padding: 16, overflowY: "auto" }}>
      <div onClick={e => e.stopPropagation()} style={{ background: "#141626", border: "1px solid #4e9af166", borderRadius: 14, padding: 22, width: "100%", maxWidth: 360, boxShadow: "0 20px 60px #000", fontFamily: "'IBM Plex Mono', monospace", marginTop: 24 }}>
        <div style={{ fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 2, color: "#4e9af1", marginBottom: 2 }}>🔑 Change Password</div>
        <div style={{ fontSize: 12, color: "#8890b8", marginBottom: 16 }}>{currentUser}</div>

        <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 6 }}>
          <input type={show ? "text" : "password"} value={current} placeholder="Current password"
            onChange={e => { setCurrent(e.target.value); setError(""); }} style={field} />
          <input type={show ? "text" : "password"} value={next} placeholder="New password"
            onChange={e => { setNext(e.target.value); setError(""); }} style={field} />
          <input type={show ? "text" : "password"} value={confirm} placeholder="Confirm new password"
            onChange={e => { setConfirm(e.target.value); setError(""); }}
            onKeyDown={e => { if (e.key === "Enter") submit(); }} style={field} />
        </div>

        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "#8890b8", margin: "8px 0 14px", cursor: "pointer" }}>
          <input type="checkbox" checked={show} onChange={e => setShow(e.target.checked)} />
          Show passwords
        </label>

        {error && (
          <div style={{ background: "#2a0a0a", border: "1px solid #ff707055", color: "#ff7070", borderRadius: 8, padding: "8px 10px", fontSize: 12, marginBottom: 12 }}>{error}</div>
        )}

        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={submit}
            style={{ flex: 1, background: "#1a4a2a", border: "1px solid #6effa0", color: "#6effa0", borderRadius: 8, padding: "11px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 15, letterSpacing: 2 }}>
            ✓ Save
          </button>
          <button onClick={onClose}
            style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "11px 18px", cursor: "pointer", fontFamily: "inherit", fontSize: 13 }}>
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}

function LoginScreen({ onLogin, users, hasSession }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [showPass, setShowPass] = useState(false);

  const handleSubmit = () => {
    const match = users.find(u => u.username === username.trim() && u.password === password);
    if (match) {
      setError("");
      onLogin(username.trim(), match.isAdmin === true);
    } else {
      setError("Incorrect username or password");
    }
  };

  return (
    <div style={{
      minHeight: "100vh", background: "#0d0f1a", fontFamily: "'IBM Plex Mono', monospace",
      color: "#eee", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
    }}>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&family=Bebas+Neue&display=swap" rel="stylesheet" />

      <div style={{
        background: "#141626", border: "1px solid #2a2d4a", borderRadius: 18,
        padding: "40px 36px", width: "100%", maxWidth: 360, boxShadow: "0 20px 60px #00000088",
      }}>
        {/* Logo */}
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 38, letterSpacing: 5, color: "#4e9af1" }}>⛳ POP APP</div>
            <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: 1, color: "#ffd966", background: "#2a1a0066", border: "1px solid #ffd96655", borderRadius: 5, padding: "2px 7px" }}>BETA</span>
          </div>
          <div style={{ fontSize: 12, color: "#8890b8", marginTop: 4, letterSpacing: 2 }}>Golf Referee · Pace of Play System</div>
          {hasSession && (
            <div style={{ marginTop: 12, background: "#0a1a0a", border: "1px solid #6effa044", borderRadius: 8, padding: "6px 14px", display: "inline-flex", alignItems: "center", gap: 6 }}>
              <div style={{ width: 6, height: 6, borderRadius: "50%", background: "#6effa0", boxShadow: "0 0 6px #6effa0" }} />
              <span style={{ fontSize: 12, color: "#6effa0", fontWeight: 700 }}>Game in progress</span>
            </div>
          )}
        </div>

        {/* Username */}
        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 12, color: "#8899cc", marginBottom: 6, letterSpacing: 1 }}>USER NAME</div>
          <input
            type="text"
            value={username}
            onChange={e => { setUsername(e.target.value); setError(""); }}
            onKeyDown={e => e.key === "Enter" && handleSubmit()}
            placeholder="Enter username"
            style={{
              width: "100%", background: "#1e2135", border: `1px solid ${error ? "#ff707066" : "#2a2d4a"}`,
              color: "#eee", borderRadius: 8, padding: "10px 14px", fontFamily: "inherit",
              fontSize: 14, outline: "none", boxSizing: "border-box",
              transition: "border-color 0.2s",
            }}
            onFocus={e => e.target.style.borderColor = "#4e9af1"}
            onBlur={e => e.target.style.borderColor = error ? "#ff707066" : "#2a2d4a"}
          />
        </div>

        {/* Password */}
        <div style={{ marginBottom: 24 }}>
          <div style={{ fontSize: 12, color: "#8899cc", marginBottom: 6, letterSpacing: 1 }}>PASSWORD</div>
          <div style={{ position: "relative" }}>
            <input
              type={showPass ? "text" : "password"}
              value={password}
              onChange={e => { setPassword(e.target.value); setError(""); }}
              onKeyDown={e => e.key === "Enter" && handleSubmit()}
              placeholder="Enter password"
              style={{
                width: "100%", background: "#1e2135", border: `1px solid ${error ? "#ff707066" : "#2a2d4a"}`,
                color: "#eee", borderRadius: 8, padding: "10px 40px 10px 14px", fontFamily: "inherit",
                fontSize: 14, outline: "none", boxSizing: "border-box",
                transition: "border-color 0.2s",
              }}
              onFocus={e => e.target.style.borderColor = "#4e9af1"}
              onBlur={e => e.target.style.borderColor = error ? "#ff707066" : "#2a2d4a"}
            />
            <button
              onClick={() => setShowPass(v => !v)}
              style={{
                position: "absolute", right: 10, top: "50%", transform: "translateY(-50%)",
                background: "none", border: "none", color: "#8890b8", cursor: "pointer", fontSize: 16, padding: 0,
              }}
            >{showPass ? "🙈" : "👁"}</button>
          </div>
        </div>

        {/* Error */}
        {error && (
          <div style={{
            background: "#2a0a0a", border: "1px solid #ff707044", borderRadius: 8,
            padding: "8px 12px", fontSize: 13, color: "#ff7070", marginBottom: 16, textAlign: "center",
          }}>{error}</div>
        )}

        {/* Submit */}
        <button
          onClick={handleSubmit}
          style={{
            width: "100%", padding: "13px 0",
            background: "linear-gradient(135deg, #1a3a6a, #4e9af1)",
            border: "none", borderRadius: 10, color: "#fff",
            fontFamily: "'Bebas Neue'", fontSize: 20, letterSpacing: 3,
            cursor: "pointer", transition: "opacity 0.15s",
          }}
          onMouseEnter={e => e.currentTarget.style.opacity = "0.85"}
          onMouseLeave={e => e.currentTarget.style.opacity = "1"}
        >
          Log in
        </button>
      </div>
    </div>
  );
}

// ─── User Management Screen ───────────────────────────────────────────────────
function UserManagementScreen({ users, onUpdateUsers, onBack, currentUser, onLogout }) {
  const [newUsername, setNewUsername] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [newIsAdmin, setNewIsAdmin] = useState(false);
  const [showNewPass, setShowNewPass] = useState(true);
  const [addError, setAddError] = useState("");
  const [addSuccess, setAddSuccess] = useState("");
  const [resetConfirm, setResetConfirm] = useState(null); // username to reset
  const [deleteConfirm, setDeleteConfirm] = useState(null); // username to delete
  const [showPassFor, setShowPassFor] = useState({}); // { username: bool }
  const [showSuggestions, setShowSuggestions] = useState(false);

  const PRESET_USERS = ["AK","BH","CJ","CP","CS","EL","IS","JK","JT","KS","NP","PH","RL","SS","ST","TS","WM","WS","ZH"];
  const suggestions = newUsername.trim()
    ? PRESET_USERS.filter(u => u.toLowerCase().startsWith(newUsername.trim().toLowerCase()) && !users.find(x => x.username.toLowerCase() === u.toLowerCase()))
    : PRESET_USERS.filter(u => !users.find(x => x.username.toLowerCase() === u.toLowerCase()));

  const handleAdd = () => {
    const trimmed = newUsername.trim();
    if (!trimmed) { setAddError("Please enter a username"); return; }
    if (users.find(u => u.username.toLowerCase() === trimmed.toLowerCase())) {
      setAddError("This username already exists");
      return;
    }
    if (!newPassword) { setAddError("Please enter a password"); return; }
    onUpdateUsers([...users, { username: trimmed, password: newPassword, isAdmin: newIsAdmin }]);
    setNewUsername("");
    setNewPassword("");
    setNewIsAdmin(false);
    setAddError("");
    setAddSuccess(`Added "${trimmed}" (${newIsAdmin ? "Admin" : "User"}) successfully`);
    setTimeout(() => setAddSuccess(""), 3000);
  };

  const adminCount = users.filter(u => u.isAdmin).length;

  const isBuiltInAdmin = (name) => (name || "").trim().toLowerCase() === "admin";

  const handleToggleRole = (u) => {
    if (isBuiltInAdmin(u.username)) return;   // built-in admin is permanent
    if (u.isAdmin && adminCount <= 1) return; // never leave the system without an admin
    const demotingSelf = u.isAdmin && u.username === currentUser;
    if (demotingSelf && !window.confirm(
      `Change your own account "${u.username}" to User?\n\nYou will lose admin access immediately and will need another admin to restore it.`
    )) return;
    onUpdateUsers(users.map(x => x.username === u.username ? { ...x, isAdmin: !x.isAdmin } : x));
  };

  const handleReset = (username) => {
    if (isBuiltInAdmin(username) && username !== currentUser) { setResetConfirm(null); return; }
    onUpdateUsers(users.map(u => u.username === username ? { ...u, password: DEFAULT_RESET_PASSWORD } : u));
    setResetConfirm(null);
  };

  const handleDelete = (username) => {
    onUpdateUsers(users.filter(u => u.username !== username));
    setDeleteConfirm(null);
  };

  const inputStyle = {
    background: "#1e2135", border: "1px solid #2a2d4a", color: "#eee",
    borderRadius: 8, padding: "9px 12px", fontFamily: "inherit", fontSize: 14, outline: "none",
  };

  return (
    <div style={{ minHeight: "100vh", background: "#0d0f1a", fontFamily: "'IBM Plex Mono', monospace", color: "#eee" }}>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&family=Bebas+Neue&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{ background: "#141626", borderBottom: "1px solid #2a2d4a", padding: "16px 24px", display: "flex", alignItems: "center", gap: 16 }}>
        <button onClick={onBack} style={{ background: "none", border: "none", color: "#4e9af1", cursor: "pointer", fontSize: 18 }}>←</button>
        <div style={{ fontFamily: "'Bebas Neue'", fontSize: 24, letterSpacing: 4, color: "#4e9af1" }}>👤 Manage Users</div>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 12, color: "#ffd966", fontWeight: 700 }}>🔑 ADMIN</span>
          <span style={{ fontSize: 13, color: "#8899cc", fontWeight: 700 }}>{currentUser}</span>
          <LogoutButton onLogout={onLogout} />
        </div>
      </div>

      <div style={{ maxWidth: 600, margin: "0 auto", padding: "28px 20px" }}>

        {/* Add User Card */}
        <div style={{ background: "#141626", border: "1px solid #4e9af133", borderRadius: 14, padding: 24, marginBottom: 24 }}>
          <div style={{ fontSize: 13, color: "#4e9af1", letterSpacing: 2, fontWeight: 700, marginBottom: 18 }}>➕ Add New User</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 10 }}>
            {/* Username with autocomplete */}
            <div style={{ position: "relative" }}>
              <input
                value={newUsername}
                onChange={e => { setNewUsername(e.target.value); setAddError(""); setShowSuggestions(true); }}
                onFocus={() => setShowSuggestions(true)}
                onBlur={() => setTimeout(() => setShowSuggestions(false), 150)}
                onKeyDown={e => e.key === "Enter" && handleAdd()}
                placeholder="Username"
                style={{ ...inputStyle, width: "100%", boxSizing: "border-box" }}
                autoComplete="off"
              />
              {showSuggestions && suggestions.length > 0 && (
                <div style={{
                  position: "absolute", top: "100%", left: 0, right: 0, zIndex: 100,
                  background: "#1e2135", border: "1px solid #4e9af155", borderRadius: 8,
                  marginTop: 4, boxShadow: "0 8px 24px #000a",
                  display: "flex", flexWrap: "wrap", gap: 6, padding: 10,
                }}>
                  {suggestions.map(s => (
                    <button
                      key={s}
                      onMouseDown={() => { setNewUsername(s); setShowSuggestions(false); }}
                      style={{
                        background: "#0d1a2a", border: "1px solid #4e9af144", color: "#4e9af1",
                        borderRadius: 6, padding: "5px 12px", cursor: "pointer",
                        fontFamily: "inherit", fontSize: 14, fontWeight: 700, letterSpacing: 1,
                      }}
                      onMouseEnter={e => { e.currentTarget.style.background = "#1a3a5a"; e.currentTarget.style.borderColor = "#4e9af1"; }}
                      onMouseLeave={e => { e.currentTarget.style.background = "#0d1a2a"; e.currentTarget.style.borderColor = "#4e9af144"; }}
                    >{s}</button>
                  ))}
                </div>
              )}
            </div>
            {/* Password + submit */}
            <div style={{ display: "flex", gap: 10 }}>
              <div style={{ position: "relative", flex: 1 }}>
                <input
                  type={showNewPass ? "text" : "password"}
                  value={newPassword}
                  onChange={e => { setNewPassword(e.target.value); setAddError(""); }}
                  onKeyDown={e => e.key === "Enter" && handleAdd()}
                  placeholder="Password"
                  style={{ ...inputStyle, width: "100%", boxSizing: "border-box", paddingRight: 36 }}
                />
                <button
                  onClick={() => setShowNewPass(v => !v)}
                  style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", color: "#8890b8", cursor: "pointer", fontSize: 14, padding: 0 }}
                >{showNewPass ? "🙈" : "👁"}</button>
              </div>
              <button
                onClick={handleAdd}
                style={{ background: "linear-gradient(135deg, #1a4a8a, #4e9af1)", border: "none", color: "#fff", borderRadius: 8, padding: "9px 18px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 16, letterSpacing: 2, whiteSpace: "nowrap" }}
              >Add</button>
            </div>
          </div>
          {/* Role selector */}
          <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
            <button
              onClick={() => setNewIsAdmin(false)}
              style={{
                flex: 1, padding: "8px 0", borderRadius: 8, cursor: "pointer",
                fontFamily: "inherit", fontSize: 13, fontWeight: 700,
                background: !newIsAdmin ? "#001020" : "#0d0f1a",
                border: `1px solid ${!newIsAdmin ? "#4e9af1" : "#2a2d4a"}`,
                color: !newIsAdmin ? "#4e9af1" : "#555",
                transition: "all 0.15s",
              }}
            >👤 User</button>
            <button
              onClick={() => setNewIsAdmin(true)}
              style={{
                flex: 1, padding: "8px 0", borderRadius: 8, cursor: "pointer",
                fontFamily: "inherit", fontSize: 13, fontWeight: 700,
                background: newIsAdmin ? "#2a1a00" : "#0d0f1a",
                border: `1px solid ${newIsAdmin ? "#ffd966" : "#2a2d4a"}`,
                color: newIsAdmin ? "#ffd966" : "#555",
                transition: "all 0.15s",
              }}
            >🔑 Admin</button>
          </div>
          {addError && <div style={{ fontSize: 12, color: "#ff7070", marginTop: 4 }}>⚠ {addError}</div>}
          {addSuccess && <div style={{ fontSize: 12, color: "#6effa0", marginTop: 4 }}>✓ {addSuccess}</div>}
          <div style={{ fontSize: 11, color: "#8890b8", marginTop: 10 }}>
            💡 The password is visible while typing · reset will set it back to <span style={{ color: "#ffd966" }}>{DEFAULT_RESET_PASSWORD}</span>
            <br />
            <span style={{ color: "#4e9af1" }}>👤 User</span> — View only, cannot edit Par/time or add player groups &nbsp;|&nbsp;
            <span style={{ color: "#ffd966" }}>🔑 Admin</span> — Full access to all functions
          </div>
        </div>

        {/* User List */}
        <div style={{ background: "#141626", border: "1px solid #2a2d4a", borderRadius: 14, overflow: "hidden" }}>
          <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a2d4a", fontSize: 12, color: "#8899cc", letterSpacing: 2, fontWeight: 700 }}>
            All users ({users.length} total)
          </div>
          {users.map((u, idx) => {
            const isCurrentUser = u.username === currentUser;
            const isShowingPass = showPassFor[u.username];
            return (
              <div key={u.username} style={{
                display: "flex", alignItems: "center", gap: 12, padding: "14px 20px",
                borderBottom: idx < users.length - 1 ? "1px solid #1a1d2e" : "none",
                background: isCurrentUser ? "#1a2135" : "transparent",
              }}>
                {/* Avatar dot */}
                <div style={{ width: 8, height: 8, borderRadius: "50%", background: u.isAdmin ? "#ffd966" : "#4e9af1", flexShrink: 0 }} />

                {/* Name + badge */}
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ fontWeight: 700, fontSize: 14, color: isCurrentUser ? "#4e9af1" : "#eee" }}>{u.username}</span>
                    {u.isAdmin && (
                      <span style={{ fontSize: 11, fontWeight: 700, color: "#ffd966", background: "#2a1a0066", border: "1px solid #ffd96644", borderRadius: 4, padding: "1px 6px", letterSpacing: 1 }}>ADMIN</span>
                    )}
                    {isCurrentUser && (
                      <span style={{ fontSize: 11, color: "#4e9af1", background: "#001a2a66", border: "1px solid #4e9af144", borderRadius: 4, padding: "1px 6px" }}>You</span>
                    )}
                  </div>
                  {/* Password peek. The built-in admin is the system's master
                      account, so its password stays hidden from other admins —
                      only the person signed in as it may reveal it. */}
                  <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 4 }}>
                    {isBuiltInAdmin(u.username) && !isCurrentUser ? (
                      <span style={{ fontSize: 12, color: "#5a6180", letterSpacing: 1 }} title="Hidden — sign in as this account to view">
                        •••••• 🔒
                      </span>
                    ) : (
                      <>
                        <span style={{ fontSize: 12, color: "#8890b8", letterSpacing: 1 }}>
                          {isShowingPass ? u.password : "••••••"}
                        </span>
                        <button
                          onClick={() => setShowPassFor(s => ({ ...s, [u.username]: !s[u.username] }))}
                          style={{ background: "none", border: "none", color: "#767fa8", cursor: "pointer", fontSize: 12, padding: 0, lineHeight: 1 }}
                          title="Show/hide password"
                        >{isShowingPass ? "🙈" : "👁"}</button>
                      </>
                    )}
                  </div>
                </div>

                {/* Actions */}
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", justifyContent: "flex-end" }}>
                  {/* Toggle role. Any account can be changed — including your own
                      and the built-in admin — except the last remaining admin,
                      which would lock everyone out of administration for good. */}
                  {isBuiltInAdmin(u.username) ? (
                    <span
                      title="The built-in admin account cannot be changed"
                      style={{ background: "#1a1000", border: "1px solid #ffd96633", color: "#7a6a33", borderRadius: 7, padding: "5px 10px", fontSize: 12, fontFamily: "inherit", fontWeight: 700, display: "inline-flex", alignItems: "center", gap: 4 }}
                    >🔒 Admin</span>
                  ) : (
                    <button
                      onClick={() => handleToggleRole(u)}
                      disabled={u.isAdmin && adminCount <= 1}
                      style={{
                        background: u.isAdmin ? "#1a1000" : "#001020",
                        border: `1px solid ${u.isAdmin ? "#ffd96644" : "#4e9af144"}`,
                        color: u.isAdmin && adminCount <= 1 ? "#555" : (u.isAdmin ? "#ffd966" : "#4e9af1"),
                        borderRadius: 7, padding: "5px 10px",
                        cursor: u.isAdmin && adminCount <= 1 ? "not-allowed" : "pointer",
                        fontSize: 12, fontFamily: "inherit", fontWeight: 700,
                      }}
                      title={
                        u.isAdmin && adminCount <= 1
                          ? "This is the only admin left — promote someone else first"
                          : u.isAdmin ? "Demote to User" : "Promote to Admin"
                      }
                    >{u.isAdmin ? "→ User" : "→ Admin"}</button>
                  )}
                  {/* Reset password — not for the built-in admin, or another
                      admin could set a known password and sign in as it. */}
                  {!isCurrentUser && !isBuiltInAdmin(u.username) && (
                    <button
                      onClick={() => setResetConfirm(u.username)}
                      style={{ background: "#1a1a0a", border: "1px solid #ffd96644", color: "#ffd966", borderRadius: 7, padding: "5px 10px", cursor: "pointer", fontSize: 12, fontFamily: "inherit", fontWeight: 700 }}
                      title={`Reset password → ${DEFAULT_RESET_PASSWORD}`}
                    >↺ Reset</button>
                  )}
                  {/* Delete — never yourself, never the built-in admin */}
                  {!isCurrentUser && !isBuiltInAdmin(u.username) && (
                    <button
                      onClick={() => setDeleteConfirm(u.username)}
                      style={{ background: "#1a0a0a", border: "1px solid #ff707044", color: "#ff7070", borderRadius: 7, padding: "5px 10px", cursor: "pointer", fontSize: 12, fontFamily: "inherit", fontWeight: 700 }}
                    >✕ Delete</button>
                  )}
                  {(isCurrentUser || isBuiltInAdmin(u.username)) && (
                    <span style={{ fontSize: 11, color: "#666f99", padding: "5px 10px" }}>—</span>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Reset Confirm Modal */}
      {resetConfirm && (
        <div style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }}>
          <div style={{ background: "#141626", border: "1px solid #ffd96688", borderRadius: 14, padding: 28, minWidth: 300, boxShadow: "0 20px 60px #000", fontFamily: "'IBM Plex Mono', monospace" }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 22, letterSpacing: 3, color: "#ffd966", marginBottom: 8 }}>↺ Reset Password</div>
            <div style={{ fontSize: 14, color: "#aaa", marginBottom: 6 }}>
              User: <span style={{ color: "#eee", fontWeight: 700 }}>{resetConfirm}</span>
            </div>
            <div style={{ fontSize: 13, color: "#888", marginBottom: 20 }}>
              The password will be immediately set to <span style={{ color: "#ffd966", fontWeight: 700, fontSize: 16 }}>{DEFAULT_RESET_PASSWORD}</span>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => handleReset(resetConfirm)}
                style={{ flex: 1, background: "#2a1500", border: "1px solid #ffd966", color: "#ffd966", borderRadius: 8, padding: "11px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 16, letterSpacing: 2 }}>
                ✓ Confirm
              </button>
              <button onClick={() => setResetConfirm(null)}
                style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "11px 16px", cursor: "pointer", fontFamily: "inherit", fontSize: 14 }}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Delete Confirm Modal */}
      {deleteConfirm && (
        <div style={{ position: "fixed", inset: 0, background: "#000000bb", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }}>
          <div style={{ background: "#141626", border: "1px solid #ff707088", borderRadius: 14, padding: 28, minWidth: 300, boxShadow: "0 20px 60px #000", fontFamily: "'IBM Plex Mono', monospace" }}>
            <div style={{ fontFamily: "'Bebas Neue'", fontSize: 22, letterSpacing: 3, color: "#ff7070", marginBottom: 8 }}>✕ Delete User</div>
            <div style={{ fontSize: 14, color: "#aaa", marginBottom: 6 }}>
              User: <span style={{ color: "#eee", fontWeight: 700 }}>{deleteConfirm}</span>
            </div>
            <div style={{ fontSize: 13, color: "#888", marginBottom: 20 }}>This action cannot be undone</div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => handleDelete(deleteConfirm)}
                style={{ flex: 1, background: "#2a0a0a", border: "1px solid #ff7070", color: "#ff7070", borderRadius: 8, padding: "11px", cursor: "pointer", fontFamily: "'Bebas Neue'", fontSize: 16, letterSpacing: 2 }}>
                ✓ Confirm delete
              </button>
              <button onClick={() => setDeleteConfirm(null)}
                style={{ background: "#1e2135", border: "1px solid #2a2d4a", color: "#9aa2c7", borderRadius: 8, padding: "11px 16px", cursor: "pointer", fontFamily: "inherit", fontSize: 14 }}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}



// ─── Main App ─────────────────────────────────────────────────────────────────
// ─── Supabase-backed persistence (shared realtime state for all judges) ────────
// `app_state`         → 1 row ("main") holding groups/pars/schedules/suspensions
//                        for the CURRENT live round (tournament_id/round_id point
//                        at which tournament + round it belongs to)
// `group_data`        → 1 row per group, holding that group's live scoring data
//                        for the CURRENT live round
// `tournaments`       → tournament metadata (name/venue/format)
// `tournament_rounds` → one row per round (Q,1,2,3,4) — when a round is finished/
//                        replaced, a snapshot of app_state + group_data is archived
//                        onto its row before the live tables are cleared for the next round
// `app_users`         → login accounts, shared across every judge/device
// Live state is stored per ROUND, so several tournaments can run at the same
// time without overwriting each other. Every read/write below is scoped by
// roundId — passing a null roundId is treated as "no round selected" and is a
// no-op, which prevents a half-loaded screen from clobbering real data.
async function fetchAppState(roundId) {
  if (!roundId) return null;
  const { data, error } = await supabase.from("round_state").select("*").eq("round_id", roundId).maybeSingle();
  if (error || !data) return null;
  return {
    groups: data.groups ?? [],
    pars: data.pars ?? [],
    parTimes: data.par_times ?? [],
    baseSchedules: data.base_schedules ?? {},
    schedules: data.schedules ?? {},
    suspensions: data.suspensions ?? [],
    isSuspended: data.is_suspended ?? false,
    pendingStopTime: data.pending_stop_time ?? "",
    playersPerGroup: data.players_per_group ?? 3,
    turnTime: data.turn_time ?? 1,
    turnTimeBack: data.turn_time_back ?? data.turn_time ?? 1,
  };
}
async function saveAppState({ roundId, groups, pars, parTimes, baseSchedules, schedules, suspensions, isSuspended, pendingStopTime, playersPerGroup, turnTime, turnTimeBack }) {
  if (!roundId) return;
  try {
    await supabase.from("round_state").upsert({
      round_id: roundId,
      groups, pars,
      par_times: parTimes,
      base_schedules: baseSchedules,
      schedules,
      suspensions,
      is_suspended: isSuspended,
      pending_stop_time: pendingStopTime,
      ...(playersPerGroup !== undefined ? { players_per_group: playersPerGroup } : {}),
      ...(turnTime !== undefined ? { turn_time: turnTime } : {}),
      ...(turnTimeBack !== undefined ? { turn_time_back: turnTimeBack } : {}),
      updated_at: new Date().toISOString(),
    });
  } catch {}
}
async function clearAppState(roundId) {
  if (!roundId) return;
  try {
    await supabase.from("round_state").delete().eq("round_id", roundId);
    await supabase.from("round_group_data").delete().eq("round_id", roundId);
  } catch {}
}
async function fetchAllGroupData(roundId) {
  if (!roundId) return {};
  const { data, error } = await supabase.from("round_group_data").select("*").eq("round_id", roundId);
  if (error || !data) return {};
  const gd = {};
  data.forEach(row => { gd[row.group_id] = row.data; });
  return gd;
}
async function saveGroupData(roundId, groupId, data, updatedAt) {
  if (!roundId) return;
  try {
    await supabase.from("round_group_data").upsert({
      round_id: roundId,
      group_id: String(groupId),
      data,
      updated_at: updatedAt || new Date().toISOString(),
    });
  } catch {}
}

// ─── Tournament / Round helpers ─────────────────────────────────────────────
// Closing a competition stops non-admin users from entering it, so nobody can
// accidentally record times into a tournament that has already finished.
// A tournament stays in 'draft' while positions are being assigned; starting it
// opens the doors so assigned referees log straight in.
async function setTournamentRunState(id, runState) {
  try {
    await supabase.from("tournaments").update({
      run_state: runState,
      started_at: runState === "started" ? new Date().toISOString() : null,
    }).eq("id", id);
  } catch {}
}
async function setTournamentStatus(id, status) {
  try {
    await supabase.from("tournaments").update({
      status,
      closed_at: status === "closed" ? new Date().toISOString() : null,
    }).eq("id", id);
  } catch {}
}
// Access rules: a tournament with NO access rows is visible to everyone.
// Once any row exists, only the listed usernames (plus admins) can see it.
async function fetchTournamentAccess() {
  const { data, error } = await supabase.from("tournament_access").select("*");
  if (error || !data) return {};
  const map = {};
  data.forEach(r => {
    if (!map[r.tournament_id]) map[r.tournament_id] = [];
    map[r.tournament_id].push(r.username);
  });
  return map;
}
async function setTournamentAccess(tournamentId, usernames) {
  try {
    await supabase.from("tournament_access").delete().eq("tournament_id", tournamentId);
    if (usernames.length) {
      await supabase.from("tournament_access").insert(
        usernames.map(username => ({ tournament_id: tournamentId, username }))
      );
    }
  } catch {}
}
// ─── Tournament positions ───────────────────────────────────────────────────
// TD/CR run the event and may edit Setup; R1-R6 are referees who record times.
const TOURNAMENT_POSITIONS = ["TD", "CR", "R1", "R2", "R3", "R4", "R5", "R6"];
const SETUP_EDIT_POSITIONS = ["TD", "CR"];

// { tournamentId: { TD: "NP", R1: "CS", ... } }
async function fetchAllRoles() {
  try {
    const { data, error } = await supabase.from("tournament_roles").select("*");
    if (error || !data) return {};
    const map = {};
    data.forEach(r => {
      if (!map[r.tournament_id]) map[r.tournament_id] = {};
      map[r.tournament_id][r.position] = r.username;
    });
    return map;
  } catch { return {}; }
}
// Replaces the whole roster for one tournament in a single pass.
async function setTournamentRoles(tournamentId, roles) {
  try {
    await supabase.from("tournament_roles").delete().eq("tournament_id", tournamentId);
    const rows = Object.entries(roles)
      .filter(([, username]) => !!username)
      .map(([position, username]) => ({ tournament_id: tournamentId, position, username }));
    if (rows.length) await supabase.from("tournament_roles").insert(rows);
    return true;
  } catch { return false; }
}
async function clearTournamentRoles(tournamentId) {
  try { await supabase.from("tournament_roles").delete().eq("tournament_id", tournamentId); } catch {}
}
// Which position (if any) this person holds in this tournament
function positionOf(rolesMap, tournamentId, username) {
  const roles = rolesMap?.[tournamentId];
  if (!roles || !username) return null;
  return Object.entries(roles).find(([, u]) => u === username)?.[0] ?? null;
}
// Two levels only: admins set things up, everyone else records times.
function canEditSetup(rolesMap, tournamentId, username, isAdmin) {
  return !!isAdmin;
}
// Everyone can see every competition; admins additionally get the closed ones.
function canUserSeeTournament(t, rolesMap, username, isAdmin) {
  if (isAdmin) return true;
  return t.status !== "closed";
}

async function fetchTournaments() {
  const { data, error } = await supabase.from("tournaments").select("*").order("created_at", { ascending: false });
  if (error || !data) return [];
  return data;
}
// One-time bridge for devices upgrading from the single-tournament build: they
// have no saved tournament/round in localStorage yet, but the old `app_state`
// row still points at whichever round was live, so we can recover it from there.
// Read-only fallbacks against the pre-migration tables. Used only when the new
// round-scoped tables have nothing for the round we're opening, so a partial or
// missed migration can never look like lost data.
async function fetchLegacyAppState() {
  try {
    const { data, error } = await supabase.from("app_state").select("*").eq("id", "main").maybeSingle();
    if (error || !data) return null;
    return {
      groups: data.groups ?? [],
      pars: data.pars ?? [],
      parTimes: data.par_times ?? [],
      baseSchedules: data.base_schedules ?? {},
      schedules: data.schedules ?? {},
      suspensions: data.suspensions ?? [],
      isSuspended: data.is_suspended ?? false,
      pendingStopTime: data.pending_stop_time ?? "",
      playersPerGroup: 3,
      turnTime: 1,
      turnTimeBack: 1,
    };
  } catch { return null; }
}
async function fetchLegacyGroupData() {
  try {
    const { data, error } = await supabase.from("group_data").select("*");
    if (error || !data) return {};
    const gd = {};
    data.forEach(row => { gd[row.group_id] = row.data; });
    return gd;
  } catch { return {}; }
}

async function fetchLegacyLiveRoundIds() {
  try {
    const { data, error } = await supabase
      .from("app_state").select("tournament_id, round_id").eq("id", "main").maybeSingle();
    if (error || !data) return null;
    return { tournamentId: data.tournament_id ?? null, roundId: data.round_id ?? null };
  } catch { return null; }
}

async function fetchTournamentById(id) {
  if (!id) return null;
  const { data, error } = await supabase.from("tournaments").select("*").eq("id", id).maybeSingle();
  if (error) return null;
  return data;
}
async function fetchRoundById(id) {
  if (!id) return null;
  const { data, error } = await supabase.from("tournament_rounds").select("*").eq("id", id).maybeSingle();
  if (error) return null;
  return data;
}
async function createTournament({ name, hostVenue, format, hasQualifying, numRounds }) {
  try {
    const { data, error } = await supabase.from("tournaments")
      .insert({ name, host_venue: hostVenue, format: format || "stroke", has_qualifying: hasQualifying !== false, num_rounds: numRounds || 4 })
      .select().maybeSingle();
    if (error) return null;
    return data;
  } catch { return null; }
}
async function updateTournament(id, { name, hostVenue, format, hasQualifying, numRounds }) {
  try {
    await supabase.from("tournaments").update({ name, host_venue: hostVenue, format: format || "stroke", has_qualifying: hasQualifying !== false, num_rounds: numRounds || 4 }).eq("id", id);
  } catch {}
}
async function deleteTournament(id) {
  try { await supabase.from("tournaments").delete().eq("id", id); } catch {}
}
// Course setup (pars, par times, turn time) belongs to the TOURNAMENT, not a single
// round — set it once on Round Q / Round 1 and every later round inherits it.
async function saveTournamentCourseSetup(id, { pars, parTimes, turnTime, turnTimeBack }) {
  if (!id) return;
  try {
    await supabase.from("tournaments").update({
      pars, par_times: parTimes, turn_time: turnTime ?? 1, turn_time_back: turnTimeBack ?? turnTime ?? 1,
    }).eq("id", id);
  } catch {}
}
async function fetchRounds(tournamentId) {
  const { data, error } = await supabase.from("tournament_rounds").select("*").eq("tournament_id", tournamentId).order("created_at", { ascending: true });
  if (error || !data) return [];
  return data;
}
// Recovers the course setup (pars / par times / transit time) from the most recent
// earlier round of the same tournament. Used as a fallback so a new round still
// inherits the setup even if the tournament-level columns aren't available.
async function fetchCourseSetupFromPreviousRounds(tournamentId, excludeRoundId) {
  try {
    const rounds = await fetchRounds(tournamentId);
    const candidates = rounds
      .filter(r => r.id !== excludeRoundId && r.archived_app_state)
      .sort((a, b) => new Date(b.finished_at || b.created_at) - new Date(a.finished_at || a.created_at));
    for (const r of candidates) {
      const s = r.archived_app_state;
      if (s?.pars?.length === 18 && s?.parTimes?.length === 18) {
        return { pars: s.pars, parTimes: s.parTimes, turnTime: s.turnTime ?? 1, turnTimeBack: s.turnTimeBack ?? s.turnTime ?? 1 };
      }
    }
  } catch {}
  return null;
}
async function createRound({ tournamentId, label, isQualifying }) {
  try {
    const { data, error } = await supabase.from("tournament_rounds")
      .insert({ tournament_id: tournamentId, label, is_qualifying: !!isQualifying, status: "live" })
      .select().maybeSingle();
    if (error) return null;
    return data;
  } catch { return null; }
}
// Snapshot the current live app_state + group_data onto the given round's row (so the
// record isn't lost), mark it finished, then wipe the live tables so the next round
// starts clean. This app only ever works with ONE live round at a time by design.
async function archiveAndFinishRound(roundId, appStateSnapshot, groupDataSnapshot) {
  try {
    await supabase.from("tournament_rounds").update({
      status: "finished",
      archived_app_state: appStateSnapshot,
      archived_group_data: groupDataSnapshot,
      finished_at: new Date().toISOString(),
    }).eq("id", roundId);
  } catch {}
}
// A tournament has exactly one active round at a time. Marking one live also
// demotes the others, so every device — and every later login — agrees on which
// round is being played.
async function setActiveRound(tournamentId, roundId) {
  if (!tournamentId || !roundId) return;
  try {
    await supabase.from("tournament_rounds")
      .update({ status: "idle" })
      .eq("tournament_id", tournamentId)
      .eq("status", "live")
      .neq("id", roundId);
    await supabase.from("tournament_rounds").update({ status: "live" }).eq("id", roundId);
  } catch {}
}
async function markRoundStatus(roundId, status) {
  try { await supabase.from("tournament_rounds").update({ status }).eq("id", roundId); } catch {}
}
async function fetchRoundArchive(roundId) {
  const { data, error } = await supabase.from("tournament_rounds").select("*").eq("id", roundId).maybeSingle();
  if (error || !data) return null;
  return data;
}
async function deleteRound(roundId) {
  try { await supabase.from("tournament_rounds").delete().eq("id", roundId); } catch {}
}
async function fetchUsers() {
  const { data, error } = await supabase.from("app_users").select("*").order("username");
  if (error || !data || !data.length) return null;
  return data.map(u => ({ username: u.username, password: u.password, isAdmin: u.is_admin === true }));
}
async function pushUsers(list) {
  try {
    await supabase.from("app_users").upsert(
      list.map(u => ({ username: u.username, password: u.password, is_admin: u.isAdmin === true }))
    );
  } catch {}
}
async function removeUsers(usernames) {
  if (!usernames?.length) return;
  try { await supabase.from("app_users").delete().in("username", usernames); } catch {}
}

export default function App() {
  // Tracks the timestamp of this device's own last write per group, so a delayed
  // realtime echo of an OLDER write can't stomp a NEWER local write (race condition
  // that was silently dropping WN/MN/TM/Bad Time log entries recorded in quick succession).
  const lastLocalWriteAt = useRef({});
  const [screen, setScreen] = useState("login");
  const [currentUser, setCurrentUser] = useState(null);
  const [isAdmin, setIsAdmin] = useState(false);
  const [users, setUsers] = useState(DEFAULT_USERS);
  const [groups, setGroups] = useState([]);
  const [pars, setPars] = useState([]);
  const [parTimes, setParTimes] = useState([]);
  const [playersPerGroup, setPlayersPerGroup] = useState(3);
  const [turnTime, setTurnTime] = useState(1);
  const [turnTimeBack, setTurnTimeBack] = useState(1);
  const [rolesMap, setRolesMap] = useState({});   // { tournamentId: { TD: "NP", ... } } — App-level copy
  const [usersReturnTo, setUsersReturnTo] = useState("setup"); // where Manage Users should go Back to
  const [onlineUsers, setOnlineUsers] = useState([]); // [{ username, isAdmin, since }] — live presence
  const [currentTournament, setCurrentTournament] = useState(null); // { id, name, host_venue, format }
  const [currentRound, setCurrentRound] = useState(null);           // { id, tournament_id, label, status, is_qualifying }
  const [baseSchedules, setBaseSchedules] = useState({}); // original schedules
  const [schedules, setSchedules] = useState({});         // adjusted schedules
  const [groupData, setGroupData] = useState({});
  const [activeGroup, setActiveGroup] = useState(null);
  const [loading, setLoading] = useState(true);

  // ─── Global Suspension State ───────────────────────────────────────────────
  const [suspensions, setSuspensions] = useState([]);     // [{ stopTime, resumeTime, offsetMin }]
  const [isSuspended, setIsSuspended] = useState(false);
  const [pendingStopTime, setPendingStopTime] = useState(""); // holds the stop time while waiting to resume

  // ─── Update users list locally + push to Supabase (handles add/edit/delete) ───
  // Self-service password change, reachable from every signed-in screen.
  const [showChangePassword, setShowChangePassword] = useState(false);
  const handleChangeOwnPassword = (newPassword) => {
    const next = users.map(u => u.username === currentUser ? { ...u, password: newPassword } : u);
    setUsers(next);
    pushUsers(next);
    setShowChangePassword(false);
    window.alert("Your password has been changed.");
  };

  const handleUpdateUsers = useCallback((newUsers) => {
    // Last line of defence: usernames identify people all through the app
    // (roles, logs, "recorded by"), so duplicates would silently merge people.
    const seen = new Set();
    const deduped = newUsers.filter(u => {
      const key = (u.username || "").trim().toLowerCase();
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
    if (deduped.length !== newUsers.length) {
      window.alert("Duplicate usernames were found and ignored. Each username must be unique.");
    }
    // And never save a list with no admin in it — that locks everyone out.
    if (!deduped.some(u => u.isAdmin)) {
      window.alert("There must be at least one admin. The change was not saved.");
      return;
    }
    // The built-in "admin" account is the guaranteed way back in: it can't be
    // demoted or removed, whatever path the change came from.
    const builtIn = deduped.find(u => (u.username || "").trim().toLowerCase() === "admin");
    if (!builtIn) {
      window.alert("The built-in \"admin\" account cannot be removed. The change was not saved.");
      return;
    }
    if (!builtIn.isAdmin) builtIn.isAdmin = true;
    setUsers(prevUsers => {
      const removed = prevUsers.filter(u => !deduped.some(n => n.username === u.username)).map(u => u.username);
      if (removed.length) removeUsers(removed);
      if (deduped.length) pushUsers(deduped);
      return deduped;
    });
  }, []);

  // ─── Load users, then restore this device's own tournament/round ────────────
  // Which round this device is working on is remembered locally, so two devices
  // can be in different tournaments at the same time.
  useEffect(() => {
    (async () => {
      try {
        const u = await fetchUsers();
        if (u && u.length) {
          setUsers(u);
        } else {
          setUsers(DEFAULT_USERS);
          pushUsers(DEFAULT_USERS);
        }
      } catch {
        setUsers(DEFAULT_USERS);
      }

      let tournament = null, round = null, state = null;
      try {
        let savedTid = localStorage.getItem("pop_tournament_id");
        let savedRid = localStorage.getItem("pop_round_id");
        // First run after upgrading from the single-tournament build: nothing is
        // saved on this device yet, so fall back to the old app_state pointer.
        if (!savedTid && !savedRid) {
          const legacy = await fetchLegacyLiveRoundIds();
          if (legacy) {
            savedTid = legacy.tournamentId;
            savedRid = legacy.roundId;
          }
        }
        if (savedTid) tournament = await fetchTournamentById(savedTid);
        if (savedRid) round = await fetchRoundById(savedRid);
        // If we only recovered the round, derive its tournament from it
        if (round && !tournament && round.tournament_id) {
          tournament = await fetchTournamentById(round.tournament_id);
        }
        // A round that was deleted, or belongs to another tournament, is ignored
        if (round && tournament && round.tournament_id !== tournament.id) round = null;
        // If the TD/CR has since moved the event to a different round, follow it
        // rather than reopening on the stale one this device remembered. People
        // who run the event pick their own round, so they're left where they were.
        const savedUserForRound = (() => { try { return localStorage.getItem("pop_app_user"); } catch { return null; } })();
        const savedAdminForRound = (() => { try { return localStorage.getItem("pop_app_is_admin") === "true"; } catch { return false; } })();
        const runsEventOnRestore = savedAdminForRound;
        if (round && tournament && !runsEventOnRestore && round.status !== "live") {
          const rs = await fetchRounds(tournament.id);
          const active = rs.find(r => r.status === "live");
          if (active && active.id !== round.id) round = active;
        }
        if (round) {
          state = await fetchAppState(round.id);
          let gd = await fetchAllGroupData(round.id);
          // Safety net: if the migration didn't reach this round, read the old
          // single-tournament tables rather than showing an empty dashboard.
          if (!state || (state.groups?.length ?? 0) === 0) {
            const legacyState = await fetchLegacyAppState();
            if (legacyState && (legacyState.groups?.length ?? 0) > 0) state = legacyState;
          }
          if (!gd || Object.keys(gd).length === 0) {
            const legacyGd = await fetchLegacyGroupData();
            if (legacyGd && Object.keys(legacyGd).length > 0) gd = legacyGd;
          }
          setGroupData(gd || {});
          if (state) {
            setGroups(state.groups);
            setPars(state.pars);
            setParTimes(state.parTimes);
            setBaseSchedules(state.baseSchedules);
            setSchedules(state.schedules);
            setSuspensions(state.suspensions);
            setIsSuspended(state.isSuspended);
            setPendingStopTime(state.pendingStopTime);
            setPlayersPerGroup(state.playersPerGroup ?? 3);
            setTurnTime(state.turnTime ?? 1);
            setTurnTimeBack(state.turnTimeBack ?? state.turnTime ?? 1);
          }
        }
      } catch {}
      setCurrentTournament(tournament);
      setCurrentRound(round);
      // Positions decide who may edit Setup, so they must be loaded before the
      // restored session lands on a screen.
      setRolesMap(await fetchAllRoles());

      // ─── Restore login session (survives refresh / pull-to-refresh) ─────────
      try {
        const savedUser = localStorage.getItem("pop_app_user");
        const savedIsAdmin = localStorage.getItem("pop_app_is_admin") === "true";
        if (savedUser) {
          setCurrentUser(savedUser);
          setIsAdmin(savedIsAdmin);

          // Reopen where this device left off. Admins may work in any tournament,
          // so they don't need a position there — everyone else must hold one,
          // otherwise a leftover tournament could put them somewhere they don't
          // belong.
          const roles = await fetchAllRoles();
          setRolesMap(roles);
          const mayReopen = !!tournament;
          // A paused or closed event isn't live work, so don't drop straight back
          // into it — show the list instead. The tournament stays remembered, so
          // whoever is preparing it is one tap away from opening it again.
          const notRunning = tournament && (tournament.status === "closed" || tournament.run_state !== "started");
          const blocked = notRunning && !savedIsAdmin;

          if (tournament && round && mayReopen && !blocked && !notRunning) {
            setScreen((state?.groups?.length ?? 0) ? "dashboard" : "setup");
          } else {
            if (!mayReopen) {
              setCurrentTournament(null);
              setCurrentRound(null);
              try {
                localStorage.removeItem("pop_tournament_id");
                localStorage.removeItem("pop_round_id");
              } catch {}
            }
            setScreen("tournament");
          }
        }
      } catch {}

      setLoading(false);
    })().catch(() => setLoading(false));   // never leave the app stuck on "Loading…"
  }, []);

  // Remember this device's tournament/round so it reopens straight into it.
  useEffect(() => {
    try {
      if (currentTournament?.id) localStorage.setItem("pop_tournament_id", currentTournament.id);
      else localStorage.removeItem("pop_tournament_id");
      if (currentRound?.id) localStorage.setItem("pop_round_id", currentRound.id);
      else localStorage.removeItem("pop_round_id");
    } catch {}
  }, [currentTournament?.id, currentRound?.id]);

  // ─── Realtime: only for the round THIS device is working on ─────────────────
  // Filtering by round_id is what keeps two tournaments from bleeding into each
  // other — a change in another tournament never reaches this screen.
  useEffect(() => {
    const roundId = currentRound?.id;
    if (!roundId) return;

    const stateChannel = supabase
      .channel(`round_state_sync_${roundId}`)
      .on("postgres_changes", { event: "*", schema: "public", table: "round_state", filter: `round_id=eq.${roundId}` }, (payload) => {
        if (payload.eventType === "DELETE") {
          setGroups([]); setPars([]); setParTimes([]); setBaseSchedules({}); setSchedules({});
          setSuspensions([]); setIsSuspended(false); setPendingStopTime("");
          return;
        }
        const row = payload.new;
        setGroups(row.groups ?? []);
        setPars(row.pars ?? []);
        setParTimes(row.par_times ?? []);
        setBaseSchedules(row.base_schedules ?? {});
        setSuspensions(row.suspensions ?? []);
        setIsSuspended(row.is_suspended ?? false);
        setPendingStopTime(row.pending_stop_time ?? "");
      })
      .subscribe();

    const groupChannel = supabase
      .channel(`round_group_data_sync_${roundId}`)
      .on("postgres_changes", { event: "*", schema: "public", table: "round_group_data", filter: `round_id=eq.${roundId}` }, (payload) => {
        if (payload.eventType === "DELETE") {
          const gid = payload.old?.group_id;
          setGroupData(prev => { const next = { ...prev }; delete next[gid]; return next; });
          return;
        }
        const row = payload.new;
        const lastWrite = lastLocalWriteAt.current[row.group_id];
        if (lastWrite && row.updated_at && row.updated_at < lastWrite) {
          return; // stale echo of an older write — we already have newer local data
        }
        setGroupData(prev => ({ ...prev, [row.group_id]: row.data }));
      })
      .subscribe();

    return () => {
      supabase.removeChannel(stateChannel);
      supabase.removeChannel(groupChannel);
    };
  }, [currentRound?.id]);

  // Users are global, so this one isn't round-scoped.
  useEffect(() => {
    const usersChannel = supabase
      .channel("users_sync")
      .on("postgres_changes", { event: "*", schema: "public", table: "app_users" }, async () => {
        const u = await fetchUsers();
        if (u) setUsers(u);
      })
      .subscribe();
    return () => { supabase.removeChannel(usersChannel); };
  }, []);

  // ─── Live presence: who else is using the app right now ──────────────────────
  // Uses Supabase Realtime Presence (no extra table needed). Each signed-in device
  // announces itself on a shared channel; leaving/closing the tab drops it
  // automatically, so the list self-cleans without any timeout logic.
  useEffect(() => {
    if (!currentUser) { setOnlineUsers([]); return; }
    const channel = supabase.channel("presence_pace_monitor", {
      config: { presence: { key: `${currentUser}-${Math.random().toString(36).slice(2, 8)}` } },
    });
    const readState = () => {
      const raw = channel.presenceState();
      const seen = new Map();
      Object.values(raw).forEach(entries => {
        entries.forEach(e => {
          if (!e?.username) return;
          // One row per person even if they have several tabs/devices open
          const prev = seen.get(e.username);
          if (!prev || (e.since && e.since < prev.since)) seen.set(e.username, e);
        });
      });
      setOnlineUsers(Array.from(seen.values()).sort((a, b) => a.username.localeCompare(b.username)));
    };
    channel
      .on("presence", { event: "sync" }, readState)
      .on("presence", { event: "join" }, readState)
      .on("presence", { event: "leave" }, readState)
      .subscribe(async (status) => {
        if (status === "SUBSCRIBED") {
          await channel.track({ username: currentUser, isAdmin, since: new Date().toISOString() });
        }
      });
    return () => { supabase.removeChannel(channel); };
  }, [currentUser, isAdmin]);

  // Calculate the total offset
  const totalOffsetMin = suspensions.reduce((acc, s) => acc + (s.offsetMin ?? 0), 0);

  // When the offset changes, adjust every group's schedule from baseSchedules
  useEffect(() => {
    if (Object.keys(baseSchedules).length === 0) return;
    const adjusted = {};
    Object.keys(baseSchedules).forEach(id => {
      adjusted[id] = baseSchedules[id].map(t => t + totalOffsetMin);
    });
    setSchedules(adjusted);
  }, [totalOffsetMin, baseSchedules]);

  const handleSuspendStop = (stopTimeStr) => {
    setPendingStopTime(stopTimeStr);
    setIsSuspended(true);
    saveAppState({ groups, pars, parTimes, baseSchedules, schedules, suspensions, isSuspended: true, pendingStopTime: stopTimeStr, roundId: currentRound?.id });
  };

  const handleSuspendResume = (resumeTimeStr) => {
    const [sh, sm] = pendingStopTime.split(":").map(Number);
    const [rh, rm] = resumeTimeStr.split(":").map(Number);
    const offsetMin = Math.max(0, (rh * 60 + rm) - (sh * 60 + sm));
    const nextSuspensions = [...suspensions, { stopTime: pendingStopTime, resumeTime: resumeTimeStr, offsetMin }];
    setSuspensions(nextSuspensions);
    setIsSuspended(false);
    setPendingStopTime("");
    saveAppState({ groups, pars, parTimes, baseSchedules, schedules, suspensions: nextSuspensions, isSuspended: false, pendingStopTime: "", roundId: currentRound?.id });
  };

  // Cancel a stop that was pressed by mistake — nothing is recorded, the clock
  // simply carries on as if it never happened.
  const handleSuspendCancel = () => {
    setIsSuspended(false);
    setPendingStopTime("");
    saveAppState({ groups, pars, parTimes, baseSchedules, schedules, suspensions, isSuspended: false, pendingStopTime: "", roundId: currentRound?.id });
  };

  // Edit an already-recorded suspension (wrong stop/resume time typed in).
  const handleSuspendEdit = (idx, stopTimeStr, resumeTimeStr) => {
    const [sh, sm] = stopTimeStr.split(":").map(Number);
    const [rh, rm] = resumeTimeStr.split(":").map(Number);
    const offsetMin = Math.max(0, (rh * 60 + rm) - (sh * 60 + sm));
    const nextSuspensions = suspensions.map((s, i) => i === idx ? { stopTime: stopTimeStr, resumeTime: resumeTimeStr, offsetMin } : s);
    setSuspensions(nextSuspensions);
    saveAppState({ groups, pars, parTimes, baseSchedules, schedules, suspensions: nextSuspensions, isSuspended, pendingStopTime, roundId: currentRound?.id });
  };

  // Remove a suspension entirely — its offset stops being applied to every schedule.
  const handleSuspendDelete = (idx) => {
    const nextSuspensions = suspensions.filter((_, i) => i !== idx);
    setSuspensions(nextSuspensions);
    saveAppState({ groups, pars, parTimes, baseSchedules, schedules, suspensions: nextSuspensions, isSuspended, pendingStopTime, roundId: currentRound?.id });
  };

  const handleLogin = async (username, admin) => {
    setCurrentUser(username);
    setIsAdmin(admin === true);
    try {
      localStorage.setItem("pop_app_user", username);
      localStorage.setItem("pop_app_is_admin", admin === true ? "true" : "false");
    } catch {}

    // Anyone holding a position in a started tournament goes straight there —
    // including admins, who are often the TD or CR of the event they're running.
    const [allT, roles] = await Promise.all([fetchTournaments(), fetchAllRoles()]);
    setRolesMap(roles);
    // Everyone picks their own Year / Tournament / Round, so login always lands
    // on the chooser rather than guessing which event they meant.
    setCurrentTournament(null);
    setCurrentRound(null);
    try {
      localStorage.removeItem("pop_tournament_id");
      localStorage.removeItem("pop_round_id");
    } catch {}
    setScreen("tournament");
  };

  // Closing a tournament signs its referees out everywhere. Each device watches
  // the tournaments table and drops out the moment its own event is closed.
  useEffect(() => {
    if (!currentUser || isAdmin || !currentTournament?.id) return;
    const channel = supabase
      .channel(`tournament_closed_${currentTournament.id}`)
      .on("postgres_changes",
        { event: "UPDATE", schema: "public", table: "tournaments", filter: `id=eq.${currentTournament.id}` },
        (payload) => {
          if (payload.new?.status === "closed") {
            window.alert(`Competition "${payload.new.name || ""}" has been closed.\n\nYou will be signed out automatically.`);
            handleLogout();
          }
        })
      .subscribe();
    return () => { supabase.removeChannel(channel); };
  }, [currentUser, isAdmin, currentTournament?.id]);

  // Referees follow whichever round the TD/CR has made active, so a device that
  // was left on an old round doesn't quietly record against the wrong one.
  useEffect(() => {
    if (!currentUser || !currentTournament?.id || !currentRound?.id) return;
    const runsEvent = isAdmin;
    if (runsEvent) return; // they're the ones choosing — don't yank them around
    const channel = supabase
      .channel(`active_round_${currentTournament.id}`)
      .on("postgres_changes",
        { event: "UPDATE", schema: "public", table: "tournament_rounds", filter: `tournament_id=eq.${currentTournament.id}` },
        async (payload) => {
          const r = payload.new;
          if (!r || r.status !== "live" || r.id === currentRound.id) return;
          await loadRound(currentTournament, r, "resume");
        })
      .subscribe();
    // Belt and braces: realtime can be blocked by a flaky connection, so also
    // check periodically. Without this a referee could sit on the wrong round.
    const poll = setInterval(async () => {
      try {
        const rs = await fetchRounds(currentTournament.id);
        const active = rs.find(r => r.status === "live");
        if (active && active.id !== currentRound.id) {
          await loadRound(currentTournament, active, "resume");
        }
      } catch {}
    }, 20000);

    return () => { supabase.removeChannel(channel); clearInterval(poll); };
  }, [currentUser, isAdmin, currentTournament?.id, currentRound?.id, rolesMap]);

  const handleLogout = () => {
    setCurrentUser(null);
    setIsAdmin(false);
    setActiveGroup(null);
    setScreen("login");
    // Don't clear the session — groups/data remain shared for other judges
    try {
      localStorage.removeItem("pop_app_user");
      localStorage.removeItem("pop_app_is_admin");
    } catch {}
  };

  const handleStart = (grps, ps, pt, pxg, tt, ttb) => {
    // Was a session already running before this "Start tracking" press?
    // If so, this is just the admin/user coming back through the Pace
    // Monitor (setup) page, not a brand-new round — so we must NOT wipe
    // out any progress that's already been recorded in the Dashboard.
    const previousGroups = groups;
    const hadSessionBefore = previousGroups.length > 0;

    // Safety net: never let an empty group list replace a running session. This
    // used to silently wipe the schedule (and orphan every group's recorded times)
    // whenever the Setup screen was opened on a device with no local group draft.
    if (hadSessionBefore && (!grps || grps.length === 0)) {
      window.alert("No group list found on this page, so the schedule was not updated.\n\nFor safety, the running session's schedule and recorded times are kept.\n\nRefresh the page to load the group list from the server, then try again.");
      return;
    }

    const sch = {};
    const data = {};
    grps.forEach(g => {
      sch[g.id] = buildScheduleOrdered(g.startTime, pt, g.startHole || 1, tt ?? 1, ttb ?? tt ?? 1);
      const existingData = groupData[g.id];
      const existedBefore = previousGroups.some(pg => pg.id === g.id);
      // Keep previously recorded times/logs for any group that already
      // existed in the live session. Only brand-new groups (or a fresh
      // session started right after "Clear Data in Dashboard") get a
      // blank scorecard.
      data[g.id] = (existedBefore && existingData)
        ? existingData
        : { records: Array(18).fill(null), holeData: Array(18).fill(null).map(() => ({ startTime: null, endTime: null })), currentHole: 0 };
    });

    setGroups(grps);
    setPars(ps);
    setParTimes(pt);
    setPlayersPerGroup(pxg ?? 3);
    setTurnTime(tt ?? 1);
    setTurnTimeBack(ttb ?? tt ?? 1);
    // Remember this course setup on the tournament so the next round starts with it
    saveTournamentCourseSetup(currentTournament?.id, { pars: ps, parTimes: pt, turnTime: tt ?? 1, turnTimeBack: ttb ?? tt ?? 1 });
    setCurrentTournament(prev => prev ? { ...prev, pars: ps, par_times: pt, turn_time: tt ?? 1, turn_time_back: ttb ?? tt ?? 1 } : prev);
    setBaseSchedules(sch);
    setSchedules(sch);
    setGroupData(data);
    setScreen("dashboard");

    // Only reset the global suspension (stop/resume) state when this is
    // truly a new round — keep it intact when resuming an in-progress one.
    const nextSuspensions = hadSessionBefore ? suspensions : [];
    const nextIsSuspended = hadSessionBefore ? isSuspended : false;
    const nextPendingStopTime = hadSessionBefore ? pendingStopTime : "";
    if (!hadSessionBefore) {
      setSuspensions([]);
      setIsSuspended(false);
      setPendingStopTime("");
    }

    saveAppState({ groups: grps, pars: ps, parTimes: pt, baseSchedules: sch, schedules: sch, suspensions: nextSuspensions, isSuspended: nextIsSuspended, pendingStopTime: nextPendingStopTime, roundId: currentRound?.id });
    const seedWrittenAt = new Date().toISOString();
    grps.forEach(g => { lastLocalWriteAt.current[g.id] = seedWrittenAt; saveGroupData(currentRound?.id, g.id, data[g.id], seedWrittenAt); });
  };

  // Applies Setup edits to the running round without navigating away and without
  // ever touching recorded times — used for live editing of group names/start
  // times/pars/par times/transit time.
  const handleApplyLiveEdits = (grps, ps, pt, pxg, tt, ttb) => {
    if (!grps || grps.length === 0) return;      // never wipe a live session
    if (groups.length === 0) return;             // nothing live to update
    const sch = {};
    grps.forEach(g => { sch[g.id] = buildScheduleOrdered(g.startTime, pt, g.startHole || 1, tt ?? 1, ttb ?? tt ?? 1); });

    setGroups(grps);
    setPars(ps);
    setParTimes(pt);
    setPlayersPerGroup(pxg ?? 3);
    setTurnTime(tt ?? 1);
    setTurnTimeBack(ttb ?? tt ?? 1);
    setBaseSchedules(sch);
    setSchedules(sch);
    saveTournamentCourseSetup(currentTournament?.id, { pars: ps, parTimes: pt, turnTime: tt ?? 1, turnTimeBack: ttb ?? tt ?? 1 });
    setCurrentTournament(prev => prev ? { ...prev, pars: ps, par_times: pt, turn_time: tt ?? 1, turn_time_back: ttb ?? tt ?? 1 } : prev);
    saveAppState({
      groups: grps, pars: ps, parTimes: pt, baseSchedules: sch, schedules: sch,
      suspensions, isSuspended, pendingStopTime,
      roundId: currentRound?.id, playersPerGroup: pxg ?? 3, turnTime: tt ?? 1, turnTimeBack: ttb ?? tt ?? 1,
    });
    // Give brand-new groups a blank scorecard; existing groups keep their data untouched.
    const newOnes = grps.filter(g => !groupData[g.id]);
    if (newOnes.length) {
      const writtenAt = new Date().toISOString();
      setGroupData(prev => {
        const next = { ...prev };
        newOnes.forEach(g => {
          next[g.id] = { records: Array(18).fill(null), holeData: Array(18).fill(null).map(() => ({ startTime: null, endTime: null })), currentHole: 0 };
          lastLocalWriteAt.current[g.id] = writtenAt;
          saveGroupData(currentRound?.id, g.id, next[g.id], writtenAt);
        });
        return next;
      });
    }
  };

  // Loads a round (and its tournament) into the app. Shared by the Tournament
  // picker screen and the Switch Round popup on the Setup screen, so both paths
  // behave identically.
  const loadRound = async (tournament, round, action, rolesOverride) => {
      // Every round now owns its data, so switching is simply "load that round".
      // Nothing belonging to the round we're leaving is ever cleared.
      // TD/CR/admin decide which round the whole event is on; referees just follow.
      // rolesOverride lets callers pass a freshly fetched map — right after login
      // the rolesMap state hasn't been applied yet.
      const rolesNow = rolesOverride || rolesMap;
      const runsEvent = isAdmin;
      if (runsEvent) await setActiveRound(tournament?.id, round.id);
      const loadedRound = runsEvent ? { ...round, status: "live" } : round;

      const state = await fetchAppState(round.id);
      const gd = await fetchAllGroupData(round.id);

      if (state) {
        setGroups(state.groups);
        setPars(state.pars);
        setParTimes(state.parTimes);
        setBaseSchedules(state.baseSchedules);
        setSchedules(state.schedules);
        setSuspensions(state.suspensions);
        setIsSuspended(state.isSuspended);
        setPendingStopTime(state.pendingStopTime);
        setPlayersPerGroup(state.playersPerGroup ?? 3);
        setTurnTime(state.turnTime ?? 1);
        setGroupData(gd || {});
      } else {
        // Round has never been set up — start it blank, but inherit the course
        // setup so pars / par times / transit time don't have to be re-entered.
        let courseSetup = null;
        if (tournament?.pars?.length === 18 && tournament?.par_times?.length === 18) {
          courseSetup = { pars: tournament.pars, parTimes: tournament.par_times, turnTime: tournament.turn_time ?? 1 };
        } else {
          courseSetup = await fetchCourseSetupFromPreviousRounds(tournament.id, round.id);
        }
        setGroups([]);
        setPars(courseSetup?.pars || []);
        setParTimes(courseSetup?.parTimes || []);
        setTurnTime(courseSetup?.turnTime ?? 1);
        setBaseSchedules({});
        setSchedules({});
        setGroupData({});
        setSuspensions([]);
        setIsSuspended(false);
        setPendingStopTime("");
      }

      let tournamentForSetup = tournament;
      const hasCourseSetup = tournament?.pars?.length === 18 && tournament?.par_times?.length === 18;
      if (!hasCourseSetup) {
        const recovered = await fetchCourseSetupFromPreviousRounds(tournament.id, round.id);
        if (recovered) {
          tournamentForSetup = { ...tournament, pars: recovered.pars, par_times: recovered.parTimes, turn_time: recovered.turnTime };
        }
      }
      setCurrentTournament(tournamentForSetup);
      setCurrentRound(loadedRound);
      setScreen((state?.groups?.length ?? 0) ? "dashboard" : "setup");
  };

  // Switch Round popup on the Setup screen: resolve (or create) the round for a
  // label within the CURRENT tournament, then load it.
  // Quick switch from the Setup page: open the chosen tournament at its live
  // round (or the most recent one) without going via the Tournament screen.
  // Start / pause the current competition straight from the Setup page.
  const handleToggleStartedFromSetup = async () => {
    const t = currentTournament;
    if (!t) return;
    const starting = t.run_state !== "started";
    const msg = starting
      ? `Start competition "${t.name}"?\n\nAssigned referees will be able to enter as soon as they log in.`
      : `Pause competition "${t.name}"?\n\nReferees stay logged in, but anyone logging in again cannot enter until it is started once more.`;
    if (!window.confirm(msg)) return;
    await setTournamentRunState(t.id, starting ? "started" : "draft");
    setCurrentTournament(prev => prev ? { ...prev, run_state: starting ? "started" : "draft" } : prev);
  };

  const handlePickTournamentFromSetup = async (t) => {
    if (!t) return;
    const rs = await fetchRounds(t.id);
    const target = rs.find(r => r.status === "live") || rs[rs.length - 1];
    if (!target) {
      // Nothing set up yet — let them choose/create a round on the full screen
      setCurrentTournament(t);
      setScreen("tournament");
      return;
    }
    await loadRound(t, target, target.status === "finished" ? "reopen" : "resume");
  };

  const handlePickRoundFromSetup = async (label) => {
    if (!currentTournament) return;
    const rounds = await fetchRounds(currentTournament.id);
    let round = rounds.find(r => r.label === label);
    if (!round) {
      round = await createRound({ tournamentId: currentTournament.id, label, isQualifying: label === "Q" });
      if (!round) { window.alert("Could not create the round. Please try again."); return; }
    }
    await loadRound(currentTournament, round, round.status === "finished" ? "reopen" : "resume");
  };

  const handleSelectGroup = (g, targetSlot) => {
    setActiveGroup({ ...g, targetSlot: targetSlot ?? null });
    setScreen("group");
  };

  const handleUpdateGroup = (id, update) => {
    const writtenAt = new Date().toISOString();
    lastLocalWriteAt.current[id] = writtenAt;
    setGroupData(prev => {
      const next = { ...prev, [id]: { ...prev[id], ...update } };
      saveGroupData(currentRound?.id, id, next[id], writtenAt);
      return next;
    });
  };

  // Clear the entire session (admin only)
  const handleClearSession = () => {
    clearAppState(currentRound?.id);
    setGroups([]);
    setPars([]);
    setParTimes([]);
    setBaseSchedules({});
    setSchedules({});
    setGroupData({});
    setActiveGroup(null);
    setSuspensions([]);
    setIsSuspended(false);
    setPendingStopTime("");
  };

  if (loading) return (
    <div style={{
      minHeight: "100vh", background: "#0d0f1a", color: "#8899cc", fontFamily: "'IBM Plex Mono', monospace",
      display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, letterSpacing: 2,
    }}>⛳ Loading...</div>
  );

  if (screen === "login") return <LoginScreen onLogin={handleLogin} users={users} hasSession={groups.length > 0} />;

  // Available on every signed-in screen, so it renders above whatever follows
  const passwordModal = showChangePassword ? (
    <ChangePasswordModal
      currentUser={currentUser}
      users={users}
      onSave={handleChangeOwnPassword}
      onClose={() => setShowChangePassword(false)}
    />
  ) : null;

  if (screen === "tournament") return (
    <>
    {passwordModal}
    <TournamentRoundScreen
      currentUser={currentUser}
      isAdmin={isAdmin}
      onLogout={handleLogout}
      onChangePassword={() => setShowChangePassword(true)}
      liveTournamentId={currentTournament?.id || null}
      liveRoundId={currentRound?.id || null}
      hasLiveGroups={groups.length > 0}
      allUsers={users}
      onManageUsers={() => { setUsersReturnTo("tournament"); setScreen("users"); }}
      onRoundSelected={loadRound}
    />
  </>
  );

  if (screen === "users" && isAdmin) return (
    <UserManagementScreen
      users={users}
      onUpdateUsers={handleUpdateUsers}
      onBack={() => setScreen(usersReturnTo)}
      currentUser={currentUser}
      onLogout={handleLogout}
    />
  );

  if (screen === "setup") return (
    <>
    {passwordModal}
    <SetupScreen
      onStart={handleStart}
      currentUser={currentUser}
      // TD and CR run the event, so they get the same editing rights as an admin
      // on this screen. Assigning positions and deleting tournaments stay
      // admin-only and live on the Tournament screen.
      isAdmin={canEditSetup(rolesMap, currentTournament?.id, currentUser, isAdmin)}
      isTrueAdmin={isAdmin}
      onManageUsers={() => { setUsersReturnTo("setup"); setScreen("users"); }}
      onLogout={handleLogout}
      onChangePassword={() => setShowChangePassword(true)}
      onClearSession={handleClearSession}
      hasLiveSession={groups.length > 0}
      onGoToDashboard={() => setScreen("dashboard")}
      tournamentName={currentTournament?.name || ""}
      hostVenue={currentTournament?.host_venue || ""}
      roundLabel={currentRound?.label || ""}
      savedPars={currentTournament?.pars || null}
      savedParTimes={currentTournament?.par_times || null}
      savedTurnTime={currentTournament?.turn_time ?? null}
      savedTurnTimeBack={currentTournament?.turn_time_back ?? currentTournament?.turn_time ?? null}
      livePars={groups.length > 0 ? pars : null}
      liveParTimes={groups.length > 0 ? parTimes : null}
      liveTurnTime={groups.length > 0 ? turnTime : null}
      liveGroups={groups}
      onApplyLiveEdits={handleApplyLiveEdits}
      onSwitchTournament={() => setScreen("tournament")}
      tournamentId={currentTournament?.id || null}
      onPickRound={handlePickRoundFromSetup}
      onPickTournament={handlePickTournamentFromSetup}
      tournamentRunState={currentTournament?.run_state || "draft"}
      onToggleStarted={handleToggleStartedFromSetup}
    />
  </>
  );

  if (screen === "dashboard") return (
    <>
    {passwordModal}
    <Dashboard
      groups={groups}
      groupData={groupData}
      pars={pars}
      parTimes={parTimes}
      schedules={schedules}
      playersPerGroup={playersPerGroup}
      tournamentName={currentTournament?.name || ""}
      hostVenue={currentTournament?.host_venue || ""}
      roundLabel={currentRound?.label || ""}
      onlineUsers={onlineUsers}
      onChangePassword={() => setShowChangePassword(true)}
      tournamentId={currentTournament?.id || null}
      isTrueAdmin={isAdmin}
      onOpenRound={(t, r) => loadRound(t, r, r.status === "finished" ? "reopen" : "resume")}
      onSelectGroup={handleSelectGroup}
      onBack={() => setScreen("setup")}
      currentUser={currentUser}
      suspensions={suspensions}
      isSuspended={isSuspended}
      pendingStopTime={pendingStopTime}
      totalOffsetMin={totalOffsetMin}
      onSuspendStop={handleSuspendStop}
      onSuspendResume={handleSuspendResume}
      onSuspendCancel={handleSuspendCancel}
      onSuspendEdit={handleSuspendEdit}
      onSuspendDelete={handleSuspendDelete}
      onLogout={handleLogout}
      onNavigateSummary={() => setScreen("summary")}
      onUpdateGroupData={handleUpdateGroup}
    />
  </>
  );

  if (screen === "summary") return (
    <SummaryScreen
      groups={groups}
      groupData={groupData}
      pars={pars}
      parTimes={parTimes}
      playersPerGroup={playersPerGroup}
      suspensions={suspensions}
      isSuspended={isSuspended}
      pendingStopTime={pendingStopTime}
      totalOffsetMin={totalOffsetMin}
      onBack={() => setScreen("dashboard")}
      currentUser={currentUser}
      onLogout={handleLogout}
    />
  );

  if (screen === "group" && activeGroup) {
    const gd = groupData[activeGroup.id] || {};
    const targetSlot = activeGroup.targetSlot;
    return (
      <GroupMonitor
        key={activeGroup.id}
        group={{
          ...activeGroup,
          records: gd.records,
          holeData: gd.holeData,
          currentHole: targetSlot !== null && targetSlot !== undefined ? targetSlot : gd.currentHole,
          actionLogs: gd.actionLogs,
          mnActive: gd.mnActive,
          mnName: gd.mnName,
          tmActive: gd.tmActive,
          tmName: gd.tmName,
          tmTarget: gd.tmTarget,
          delayMin: gd.delayMin,
        }}
        pars={pars}
        parTimes={parTimes}
        playersPerGroup={playersPerGroup}
        schedule={schedules[activeGroup.id]}
        statusOverride={computeGroupStatusFor(activeGroup, groups, groupData, parTimes)}
        onUpdate={(update) => handleUpdateGroup(activeGroup.id, update)}
        onBack={() => setScreen("dashboard")}
        currentUser={currentUser}
        isSuspended={isSuspended}
        suspensions={suspensions}
        totalOffsetMin={totalOffsetMin}
        pendingStopTime={pendingStopTime}
        onLogout={handleLogout}
        allGroups={groups}
        onSwitchGroup={handleSelectGroup}
      />
    );
  }

  return null;
}
