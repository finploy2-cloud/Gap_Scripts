/**********************************************************
 * UNIVERSAL DATE PARSER
 **********************************************************/
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


/**********************************************************
 * GAP CALCULATION FUNCTION
 **********************************************************/
function calculateGaps(sheet, row, candCol, clientCol, remarkCol,
                       candGapCol, clientGapCol, remarkGapCol) {

  const today = new Date();
  const todayZero = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  const candDate = parseDate(sheet.getRange(row, candCol).getValue());
  const clientDate = parseDate(sheet.getRange(row, clientCol).getValue());
  const remarkDate = parseDate(sheet.getRange(row, remarkCol).getValue());

  function calculateGap(date) {
    if (!date) return 0;
    const diff = Math.floor((todayZero - date) / 86400000) - 1; // EXCLUDE TODAY
    return diff < 0 ? 0 : diff; // avoid negative values
  }

  // Candidate gap
  sheet.getRange(row, candGapCol).setValue(calculateGap(candDate));

  // Client gap
  sheet.getRange(row, clientGapCol).setValue(calculateGap(clientDate));

  // Remark gap
  sheet.getRange(row, remarkGapCol).setValue(calculateGap(remarkDate));
}


/**********************************************************
 * ON-EDIT TRIGGER FOR MANUAL UPDATES
 **********************************************************/
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


  /******** CANDIDATE FOLLOW-UP ********/
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


  /******** CLIENT FOLLOW-UP ********/
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


  /******** STATUS CHANGE → UPDATE REMARK DATE ********/
  if (col === statusCol) {
    const newVal = e.value;
    const oldVal = e.oldValue;

    if (newVal && newVal !== oldVal) {
      const today = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy");
      sheet.getRange(row, remarkCol).setValue(today);
    }

    calculateGaps(sheet, row, candCol, clientCol, remarkCol,
                  candGapCol, clientGapCol, remarkGapCol);
  }


  /******** ANY DATE CHANGE → RECALCULATE GAP ********/
  if ([candCol, clientCol, remarkCol].includes(col)) {
    calculateGaps(sheet, row, candCol, clientCol, remarkCol,
                  candGapCol, clientGapCol, remarkGapCol);
  }
}


/**********************************************************
 * RECALCULATE ALL ROWS — PYTHON + DAILY AUTO
 **********************************************************/
function recalcAllGaps() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  function safeCol(name) {
    const idx = headers.indexOf(name);
    if (idx === -1) {
      throw new Error("Column missing or mismatched: " + name);
    }
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
    calculateGaps(sheet, r, candCol, clientCol, remarkCol,
                  candGapCol, clientGapCol, remarkGapCol);
  }
}

