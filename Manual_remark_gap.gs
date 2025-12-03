function calculateRemarkGapSimple() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const headers = data[0];

  const remarkDateCol = headers.indexOf("Remark_date_change");
  const remarkGapCol  = headers.indexOf("Remark_gap");

  if (remarkDateCol === -1 || remarkGapCol === -1) {
    SpreadsheetApp.getUi().alert("Header missing: Remark_date_change or Remark_gap");
    return;
  }

  const today = new Date();
  const todayZero = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  for (let r = 1; r < data.length; r++) {
    const value = data[r][remarkDateCol];

    if (!value) {
      sheet.getRange(r + 1, remarkGapCol + 1).setValue(0);
      continue;
    }

    const parsedDate = parseSimpleDate(value);

    if (isSpecialDate(parsedDate)) {
      continue;
    }

    if (parsedDate) {
      let diff = Math.floor((todayZero - parsedDate) / 86400000);

      // 👇 NEW LOGIC: do not count the current day
      if (diff > 0) diff = diff - 1;

      sheet.getRange(r + 1, remarkGapCol + 1).setValue(diff);
    } else {
      sheet.getRange(r + 1, remarkGapCol + 1).setValue(0);
    }
  }
}

function parseSimpleDate(v) {
  if (v instanceof Date) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  if (typeof v === "string") {
    const cleaned = v.replace(/[-.]/g, "/");
    const p = cleaned.split("/");

    if (p.length === 3) {
      const d = parseInt(p[0], 10);
      const m = parseInt(p[1], 10);
      const y = parseInt(p[2], 10);
      return new Date(y, m - 1, d);
    }
  }
  return null;
}

function isSpecialDate(dateObj) {
  if (!dateObj) return false;

  return (
    dateObj.getFullYear() === 2025 &&
    dateObj.getMonth() === 10 &&
    dateObj.getDate() === 9
  );
}
