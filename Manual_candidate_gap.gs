function recalcCandidateGapOnly() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const candCol = headers.indexOf("Cand_Date_update") + 1;
  const candGapCol = headers.indexOf("candidate_gap") + 1;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const today = new Date();
  const todayZero = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  // Read entire Cand_Date_update column at once
  const dates = sheet.getRange(2, candCol, lastRow - 1).getValues();

  // Prepare result array
  const gapValues = [];

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

  // Process all rows in memory
  for (let i = 0; i < dates.length; i++) {
    const candDate = parseDate(dates[i][0]);
    let gap = 0;

    if (candDate) {
      gap = Math.floor((todayZero - candDate) / 86400000) - 1;
      if (gap < 0) gap = 0;
    }

    gapValues.push([gap]); // must be 2D
  }

  // Write all gaps in ONE SHOT (super fast)
  sheet.getRange(2, candGapCol, gapValues.length, 1).setValues(gapValues);

  Logger.log("Candidate gaps recalculated FAST.");
}
