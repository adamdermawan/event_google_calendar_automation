/**
 * Brand Engagement % Calculator
 * For Form_Responses sheet
 * Run addBrandEngagementColumns() once to set up headers,
 * then run calculateBrandEngagement() to populate all rows.
 */

// ── Column letter map (adjust if your sheet differs) ──────────────────────────
const COL = {
  EVENT_CATEGORY:   "D",   // Event Category
  TARGET_AUDIENCE:  "M",   // Target Audience
  PARTICIPANTS:     "N",   // Estimated Number of Participants
  PRIORITY:         "O",   // Priority Level
  // Output columns (appended after the last used column)
  AUDIENCE_WEIGHT:  "T",
  CATEGORY_WEIGHT:  "U",
  PARTICIPATION_SCORE: "V",
  PRIORITY_WEIGHT:  "W",
  TOTAL_SCORE:      "X",
  ENGAGEMENT_PCT:   "Y",
};

const MAX_SCORE = 20; // 5+5+5+5

// ── Scoring tables ─────────────────────────────────────────────────────────────

function getAudienceWeight(raw) {
  if (!raw) return 2;
  const val = raw.toLowerCase();
  if (val.includes("client"))    return 5;
  if (val.includes("partner"))   return 4;
  if (val.includes("media"))     return 4;
  if (val.includes("management"))return 3;
  if (val.includes("public"))    return 3;
  if (val.includes("employee"))  return 2;
  return 2; // default
}

function getCategoryWeight(raw) {
  if (!raw) return 2;
  const val = raw.toLowerCase();
  if (val.includes("external"))  return 5;
  if (val.includes("csr"))       return 4;
  if (val.includes("community")) return 4;
  if (val.includes("training") || val.includes("workshop")) return 3;
  if (val.includes("internal"))  return 2;
  if (val.includes("others"))    return 3;
  return 2;
}

function getParticipationScore(raw) {
  const n = parseInt(raw, 10);
  if (isNaN(n)) return 2; // default when missing
  if (n > 200)  return 5;
  if (n >= 100) return 4;
  if (n >= 50)  return 3;
  if (n >= 20)  return 2;
  return 1;
}

function getPriorityWeight(raw) {
  if (!raw) return 3;
  const val = raw.toLowerCase();
  if (val.includes("high"))   return 5;
  if (val.includes("medium")) return 3;
  if (val.includes("low"))    return 1;
  return 3;
}

// ── Main: add header row labels ────────────────────────────────────────────────

function addBrandEngagementColumns() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = ["Audience Weight","Category Weight","Participation Score",
                   "Priority Weight","Total Score","Brand Engagement %"];
  const cols    = [COL.AUDIENCE_WEIGHT, COL.CATEGORY_WEIGHT,
                   COL.PARTICIPATION_SCORE, COL.PRIORITY_WEIGHT,
                   COL.TOTAL_SCORE, COL.ENGAGEMENT_PCT];

  cols.forEach((col, i) => {
    sheet.getRange(col + "1").setValue(headers[i]).setFontWeight("bold")
         .setBackground("#1a73e8").setFontColor("#ffffff");
  });

  SpreadsheetApp.getUi().alert("Headers added! Now run calculateBrandEngagement().");
}

// ── Main: calculate scores for every data row ──────────────────────────────────

function calculateBrandEngagement() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No data rows found.");
    return;
  }

  for (let row = 2; row <= lastRow; row++) {
    const category   = sheet.getRange(COL.EVENT_CATEGORY   + row).getValue();
    const audience   = sheet.getRange(COL.TARGET_AUDIENCE   + row).getValue();
    const parts      = sheet.getRange(COL.PARTICIPANTS      + row).getValue();
    const priority   = sheet.getRange(COL.PRIORITY          + row).getValue();

    const aw = getAudienceWeight(String(audience));
    const cw = getCategoryWeight(String(category));
    const ps = getParticipationScore(String(parts));
    const pw = getPriorityWeight(String(priority));
    const total = aw + cw + ps + pw;
    const pct   = Math.round((total / MAX_SCORE) * 100);

    sheet.getRange(COL.AUDIENCE_WEIGHT      + row).setValue(aw);
    sheet.getRange(COL.CATEGORY_WEIGHT      + row).setValue(cw);
    sheet.getRange(COL.PARTICIPATION_SCORE  + row).setValue(ps);
    sheet.getRange(COL.PRIORITY_WEIGHT      + row).setValue(pw);
    sheet.getRange(COL.TOTAL_SCORE          + row).setValue(total);
    sheet.getRange(COL.ENGAGEMENT_PCT       + row).setValue(pct)
         .setNumberFormat("0\"%\"");

    // Color-code engagement %
    const cell = sheet.getRange(COL.ENGAGEMENT_PCT + row);
    if (pct >= 75)      cell.setBackground("#c6efce").setFontColor("#276221"); // green
    else if (pct >= 50) cell.setBackground("#ffeb9c").setFontColor("#9c6500"); // yellow
    else                cell.setBackground("#ffc7ce").setFontColor("#9c0006"); // red
  }

  SpreadsheetApp.getUi().alert(`Done! Calculated ${lastRow - 1} rows.`);
}

// ── Optional: auto-run on form submit ─────────────────────────────────────────
// In Apps Script, go to Triggers → Add Trigger →
// Function: onFormSubmitTrigger | Event: From spreadsheet | On form submit

function onFormSubmitTrigger(e) {
  const sheet = e.range.getSheet();
  const row   = e.range.getRow();

  const category = sheet.getRange(COL.EVENT_CATEGORY   + row).getValue();
  const audience = sheet.getRange(COL.TARGET_AUDIENCE   + row).getValue();
  const parts    = sheet.getRange(COL.PARTICIPANTS      + row).getValue();
  const priority = sheet.getRange(COL.PRIORITY          + row).getValue();

  const aw = getAudienceWeight(String(audience));
  const cw = getCategoryWeight(String(category));
  const ps = getParticipationScore(String(parts));
  const pw = getPriorityWeight(String(priority));
  const total = aw + cw + ps + pw;
  const pct   = Math.round((total / MAX_SCORE) * 100);

  sheet.getRange(COL.AUDIENCE_WEIGHT      + row).setValue(aw);
  sheet.getRange(COL.CATEGORY_WEIGHT      + row).setValue(cw);
  sheet.getRange(COL.PARTICIPATION_SCORE  + row).setValue(ps);
  sheet.getRange(COL.PRIORITY_WEIGHT      + row).setValue(pw);
  sheet.getRange(COL.TOTAL_SCORE          + row).setValue(total);
  sheet.getRange(COL.ENGAGEMENT_PCT       + row).setValue(pct)
       .setNumberFormat("0\"%\"");
}