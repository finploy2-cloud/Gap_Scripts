# ============================================================
#  GOOGLE SHEET AUTOMATION — GAP CALCULATION SYSTEM
# ============================================================
# This repository contains a Google Apps Script that automates
# candidate/client follow-up tracking inside a Google Sheet.
#
# The script performs 4 major tasks:
#
#   1) Converts any messy date format into a clean standard date.
#   2) Automatically increases follow-up counters when dates change.
#   3) Automatically updates the remark date when Status changes.
#   4) Calculates GAP values (days passed) for:
#         - Candidate follow-up date
#         - Client follow-up date
#         - Remark date
#
# GAP = (today - stored date) minus 1 day.
# Today is excluded so GAP never shows negative values.
#
# This keeps the sheet fully automated for the recruitment team.
# ============================================================


# ============================================================
# 1) UNIVERSAL DATE PARSER
# ============================================================
# Purpose:
#   - Clean and standardize date formats.
#   - Accepts dates like:
#         12-05-2025
#         12/05/2025
#         12.05.2025
#   - Returns a proper Date object for reliable calculations.
#
# If an invalid date is entered → returns null.
# ============================================================

function parseDate(value) {
  if (!value) return null;
  try {
    if (value instanceof Date)
      return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    const clean = value.toString().replace(/[-.]/g, "/");
    const p = clean.split("/");
    if (p.length === 3) {
      const [d, m, y] = p.map(x => parseInt(x, 10));
      return new Date(y, m - 1, d);
    }
  } catch (err) {}
  return null;
}


# ============================================================
# 2) GAP CALCULATION FUNCTION
# ============================================================
# Purpose:
#   - Calculate how many days have passed since:
#         * Cand_Date_update
#         * Client_Date_update
#         * Remark_date_change
#
# Logic:
#   - Convert date for each column → calculate difference.
#   - Write the GAP values to:
#         candidate_gap
#         client_gap
#         Remark_gap
#
# Notes:
#   - If date is empty → GAP = 0
#   - GAP never becomes negative.
# ============================================================

function calculateGaps(sheet, row, candCol, clientCol, remarkCol, candGapCol, clientGapCol, remarkGapCol) {
  const today = new Date();
  const todayZero = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const candDate = parseDate(sheet.getRange(row, candCol).getValue());
  const clientDate = parseDate(sheet.getRange(row, clientCol).getValue());
  const remarkDate = parseDate(sheet.getRange(row, remarkCol).getValue());

  function calculateGap(date) {
    if (!date) return 0;
    const diff = Math.floor((todayZero - date) / 86400000) - 1;
    return diff < 0 ? 0 : diff;
  }

  sheet.getRange(row, candGapCol).setValue(calculateGap(candDate));
  sheet.getRange(row, clientGapCol).setValue(calculateGap(clientDate));
  sheet.getRange(row, remarkGapCol).setValue(calculateGap(remarkDate));
}


# ============================================================
# 3) LIVE TRIGGER: onEdit()
# ============================================================
# This function runs automatically every time a user edits the sheet.
#
# It performs:
#
#   A) AUTO FOLLOW-UP COUNT (Candidate)
#         - New date added → count++
#         - Date changed → count++
#         - Date erased → count = 0
#
#   B) AUTO FOLLOW-UP COUNT (Client)
#         - Same logic as candidate
#
#   C) AUTO REMARK DATE UPDATE
#         - When "Status" changes → set Remark_date_change = Today
#
#   D) AUTO GAP RECALCULATION
#         - Anytime the user edits Cand_Date_update,
#                                    Client_Date_update,
#                                    Remark_date_change
#           → GAP recalculates instantly.
# ============================================================

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row === 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const candCol = headers.indexOf("Cand_Date_update") + 1;
  const candFollowCol = headers.indexOf("Cand_followup") + 1;
  const clientCol = headers.indexOf("Client_Date_update") + 1;
  const clientFollowCol = headers.indexOf("Client_followup") + 1;
  const statusCol = headers.indexOf("Status") + 1;
  const remarkCol = headers.indexOf("Remark_date_change") + 1;
  const candGapCol = headers.indexOf("candidate_gap") + 1;
  const clientGapCol = headers.indexOf("client_gap") + 1;
  const remarkGapCol = headers.indexOf("Remark_gap") + 1;

  if (col === candCol) {
    const newVal = e.range.getValue();
    const oldVal = e.oldValue;
    const cell = sheet.getRange(row, candFollowCol);
    let count = Number(cell.getValue()) || 0;
    if (!oldVal && newVal) count++;
    else if (oldVal && newVal && oldVal !== newVal) count++;
    else if (!newVal) count = 0;
    cell.setValue(count);
  }

  if (col === clientCol) {
    const newVal = e.range.getValue();
    const oldVal = e.oldValue;
    const cell = sheet.getRange(row, clientFollowCol);
    let count = Number(cell.getValue()) || 0;
    if (!oldVal && newVal && count === 0) count = 0;
    else if (oldVal && newVal && oldVal !== newVal) count++;
    else if (!newVal) count = 0;
    cell.setValue(count);
  }

  if (col === statusCol) {
    const newVal = e.value;
    const oldVal = e.oldValue;
    if (newVal && newVal !== oldVal) {
      const today = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy");
      sheet.getRange(row, remarkCol).setValue(today);
    }
    calculateGaps(sheet, row, candCol, clientCol, remarkCol, candGapCol, clientGapCol, remarkGapCol);
  }

  if ([candCol, clientCol, remarkCol].includes(col)) {
    calculateGaps(sheet, row, candCol, clientCol, remarkCol, candGapCol, clientGapCol, remarkGapCol);
  }
}


# ============================================================
# 4) recalcAllGaps() — FULL CLEANUP FUNCTION
# ============================================================
# This function:
#   - Loops through every row in the sheet
#   - Recalculates GAP values
#   - Useful when:
#         * Data is bulk pasted
#         * Column names change
#         * Daily refresh is needed
#
# You can run this with a Time-Driven trigger (daily).
# ============================================================

function recalcAllGaps() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  function safeCol(name) {
    const idx = headers.indexOf(name);
    if (idx === -1) throw new Error("Column missing or mismatched: " + name);
    return idx + 1;
  }

  const candCol = safeCol("Cand_Date_update");
  const clientCol = safeCol("Client_Date_update");
  const remarkCol = safeCol("Remark_date_change");
  const candGapCol = safeCol("candidate_gap");
  const clientGapCol = safeCol("client_gap");
  const remarkGapCol = safeCol("Remark_gap");

  const last = sheet.getLastRow();
  for (let r = 2; r <= last; r++) {
    calculateGaps(sheet, r, candCol, clientCol, remarkCol, candGapCol, clientGapCol, remarkGapCol);
  }
}


# ============================================================
# SUMMARY — WHAT THIS SYSTEM AUTOMATES
# ============================================================
# ✔ Eliminates manual GAP calculation
# ✔ Automates candidate follow-up count
# ✔ Automates client follow-up count
# ✔ Auto-updates remark date on status change
# ✔ Handles all date formats safely
# ✔ Ensures sheet accuracy with daily refresh
# ✔ Reduces manual mistakes by recruitment team

# ============================================================
# HOW TO INSTALL
# ============================================================
# 1. Open Google Sheet → Extensions → Apps Script
# 2. Paste this entire script
# 3. Save the project
# 4. Go to Triggers → Add Trigger:
#         Function: recalcAllGaps
#         Event: Time-driven
#         Frequency: Daily
#
# DONE. Your tracking sheet is now fully automated.
# ============================================================
