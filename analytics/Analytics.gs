// Google Apps Script - Paradox ATS Bulk Tools usage analytics
//
// Deploy as a NEW Apps Script project (keep it separate from Ask Ari):
//   1. https://script.google.com > New project > name it "Paradox Tools Analytics"
//   2. Paste this file over the default Code.gs
//   3. Run setup() once from the editor to create the sheet and authorize
//   4. Deploy > New deployment > Web app
//        Execute as: Me
//        Who has access: Anyone
//   5. Copy the /exec URL and send it to Claude to bake into telemetry.js
//
// The extension sends anonymous JSON events here. No names, emails, candidate
// data or resume content are ever included. The user id is a random UUID
// generated once per browser install.

var EVENT_HEADERS = [
  "Timestamp", "User ID", "Event", "Tab", "Opening", "Format",
  "Count", "Errors", "Version", "Browser", "Platform", "Subdomain", "Date", "Week Of"
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (!data || !data.event) return ok_();
    var ss = getOrCreateSheet_();
    var sheet = ss.getSheetByName("Events");
    var now = new Date();
    var date = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var weekOf = new Date(date.getTime() - date.getDay() * 86400000); // Sunday
    sheet.appendRow([
      now,
      String(data.uid || "unknown").slice(0, 64),
      String(data.event).slice(0, 40),
      String(data.tab || "").slice(0, 20),
      String(data.opening || "").slice(0, 20),
      String(data.format || "").slice(0, 20),
      Number(data.count) || 0,
      Number(data.errors) || 0,
      String(data.version || "").slice(0, 20),
      String(data.browser || "").slice(0, 20),
      String(data.platform || "paradox").slice(0, 20),
      String(data.subdomain || "").slice(0, 40),
      date,
      weekOf
    ]);
  } catch (err) {
    // never fail the caller over analytics
  }
  return ok_();
}

function doGet() {
  return ContentService.createTextOutput("Paradox Tools Analytics endpoint is live.");
}

function ok_() {
  return ContentService.createTextOutput("ok").setMimeType(ContentService.MimeType.TEXT);
}

// Run this once from the editor to create the sheet and grant permissions.
function setup() {
  var ss = getOrCreateSheet_();
  Logger.log("Analytics sheet: " + ss.getUrl());
  return ss.getUrl();
}

function getAnalyticsSheetUrl() {
  return getOrCreateSheet_().getUrl();
}

function getOrCreateSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("ANALYTICS_SHEET_ID");
  var ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (ss) return ss;

  ss = SpreadsheetApp.create("Paradox ATS Bulk Tools — Usage Analytics");
  props.setProperty("ANALYTICS_SHEET_ID", ss.getId());

  // Events sheet: raw log
  var ev = ss.getActiveSheet();
  ev.setName("Events");
  ev.appendRow(EVENT_HEADERS);
  ev.setFrozenRows(1);
  ev.getRange(1, 1, 1, EVENT_HEADERS.length).setFontWeight("bold");
  ev.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
  ev.getRange("M:N").setNumberFormat("yyyy-mm-dd");

  // Summary sheet: the numbers a manager wants
  var sm = ss.insertSheet("Summary", 0);
  buildSummary_(sm);
  return ss;
}

function buildSummary_(sm) {
  var rows = [
    ["Paradox ATS Bulk Tools — Usage Summary", ""],
    ["", ""],
    ["Assumptions (edit these)", ""],
    ["Minutes saved per resume downloaded", 0.5],
    ["Minutes saved per candidate dispositioned", 0.25],
    ["", ""],
    ["All time", ""],
    ["Distinct users", '=IFERROR(COUNTUNIQUE(FILTER(Events!B2:B, Events!B2:B<>"")),0)'],
    ["Resumes downloaded", '=SUMIF(Events!C2:C,"resume_download",Events!G2:G)'],
    ["Candidates dispositioned", '=SUMIF(Events!C2:C,"disposition",Events!G2:G)'],
    ["Download sessions", '=COUNTIF(Events!C2:C,"resume_download")'],
    ["Disposition sessions", '=COUNTIF(Events!C2:C,"disposition")'],
    ["Estimated hours saved", "=ROUND((B9*B4+B10*B5)/60,1)"],
    ["", ""],
    ["Last 7 days", ""],
    ["Active users", '=IFERROR(COUNTUNIQUE(FILTER(Events!B2:B, Events!M2:M>=TODAY()-7)),0)'],
    ["Resumes downloaded", '=SUMIFS(Events!G2:G,Events!C2:C,"resume_download",Events!M2:M,">="&(TODAY()-7))'],
    ["Candidates dispositioned", '=SUMIFS(Events!G2:G,Events!C2:C,"disposition",Events!M2:M,">="&(TODAY()-7))'],
    ["Estimated hours saved", "=ROUND((B17*B4+B18*B5)/60,1)"],
    ["", ""],
    ["Last 30 days", ""],
    ["Active users", '=IFERROR(COUNTUNIQUE(FILTER(Events!B2:B, Events!M2:M>=TODAY()-30)),0)'],
    ["Resumes downloaded", '=SUMIFS(Events!G2:G,Events!C2:C,"resume_download",Events!M2:M,">="&(TODAY()-30))'],
    ["Candidates dispositioned", '=SUMIFS(Events!G2:G,Events!C2:C,"disposition",Events!M2:M,">="&(TODAY()-30))'],
    ["Estimated hours saved", "=ROUND((B23*B4+B24*B5)/60,1)"],
    ["", ""],
    ["Feature split (all time)", ""],
    ["Zip downloads", '=COUNTIFS(Events!C2:C,"resume_download",Events!F2:F,"zip")'],
    ["Merged PDF downloads", '=COUNTIFS(Events!C2:C,"resume_download",Events!F2:F,"merged")'],
    ["Chrome users", '=IFERROR(COUNTUNIQUE(FILTER(Events!B2:B, Events!J2:J="chrome")),0)'],
    ["Edge users", '=IFERROR(COUNTUNIQUE(FILTER(Events!B2:B, Events!J2:J="edge")),0)'],
    ["Paradox events", '=COUNTIF(Events!K2:K,"paradox")'],
    ["Rippling events", '=COUNTIF(Events!K2:K,"rippling")']
  ];
  sm.getRange(1, 1, rows.length, 2).setValues(rows);
  sm.getRange("A1").setFontWeight("bold").setFontSize(14);
  [3, 7, 15, 21, 27].forEach(function (r) { sm.getRange(r, 1).setFontWeight("bold"); });
  sm.getRange("B4:B5").setBackground("#fff2cc");
  sm.setColumnWidth(1, 300);
  sm.setColumnWidth(2, 120);

  // Weekly table (last 12 weeks) starting at row 34
  var start = 35;
  sm.getRange(start, 1, 1, 6).setValues([[
    "Week of", "Active users", "Download sessions", "Resumes", "Dispositions", "Hours saved"
  ]]).setFontWeight("bold");
  for (var i = 0; i < 12; i++) {
    var r = start + 1 + i;
    var wk = "=TODAY()-WEEKDAY(TODAY())+1-" + (i * 7);
    sm.getRange(r, 1).setFormula(wk).setNumberFormat("yyyy-mm-dd");
    sm.getRange(r, 2).setFormula('=IFERROR(COUNTUNIQUE(FILTER(Events!B2:B, Events!N2:N=A' + r + ')),0)');
    sm.getRange(r, 3).setFormula('=COUNTIFS(Events!C2:C,"resume_download",Events!N2:N,A' + r + ')');
    sm.getRange(r, 4).setFormula('=SUMIFS(Events!G2:G,Events!C2:C,"resume_download",Events!N2:N,A' + r + ')');
    sm.getRange(r, 5).setFormula('=SUMIFS(Events!G2:G,Events!C2:C,"disposition",Events!N2:N,A' + r + ')');
    sm.getRange(r, 6).setFormula('=ROUND((D' + r + '*$B$4+E' + r + '*$B$5)/60,1)');
  }
}
