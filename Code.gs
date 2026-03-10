// ═══════════════════════════════════════════════════════════════
//  REPORTING LEADS - Final Version
// ═══════════════════════════════════════════════════════════════

var FIELD_MAP = {
  date:      ["date","created at","date de creation","date de creation","date de creations"],
  nom:       ["nom","name","last name"],
  prenom:    ["first name","prenom","prénom"],
  tel:       ["telephone","téléphone","phone","phone/whatsapp","numero de telephone (gsm)","numéro de téléphone (gsm)","whatsapp"],
  email:     ["email","e-mail","adresse e-mail","adresse email"],
  entreprise:["entreprise","company name","company_name","nom de l'entreprise"],
  ville:     ["ville","city"],
  pays:      ["country","pays"],
  adresse:   ["adresse","address","address line 1"],
  type:      ["business type","type d'activite","type","type d activite","other business type","autre_type","type d’activité"],
  message:   ["additional details","additional information","message","request"],
  quantity:  ["production_volume","estimated order quantity","volume_par_commande"],
  prodtype:  ["product_type","toys category","other toys category"]
};

// Point 1 -- Nouvel ordre : Date, Source, Type en premier
var STANDARD_HEADERS = ["Date","Source","Type","Full Name","Phone","Email","Company","City","Country","Address","Message","Quantity","Product Type"];

var MONTHS_EN = ["January","February","March","April","May","June","July","August","September","October","November","December"];
var MONTHS_SH = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

var DROP_ROW     = 13;
var DROP_COL     = 3;
var TAG_COL      = 5;
var TABLE_START_ROW = 16;

// ───────────────────────────────────────────────────────────────
// LECTURE CONFIG -- A=Nom, B=URL, C=Tag, D=GA4 Property ID
// ───────────────────────────────────────────────────────────────
function readConfig(ss) {
  var cfg = ss.getSheetByName("Config");
  if (!cfg) return null;

  var lastRow = Math.max(cfg.getLastRow() - 1, 1);
  var cfgData = cfg.getRange(2, 1, lastRow, 4).getValues();

  var allSourceNames = [], seenNames = {};
  var tagMap = {}, ga4Map = {}, allTags = [], seenTags = {};

  cfgData.forEach(function(r) {
    var name = r[0] ? r[0].toString().trim() : "";
    var tag  = r[2] ? r[2].toString().trim().toLowerCase() : "";
    var ga4  = r[3] ? r[3].toString().trim() : "";

    if (name && !seenNames[name]) {
      allSourceNames.push(name);
      seenNames[name] = true;
      if (tag) tagMap[name] = tag;
      if (ga4) ga4Map[name] = ga4;
    }
    if (tag && !seenTags[tag]) { allTags.push(tag); seenTags[tag] = true; }
  });

  var sources = cfgData.filter(function(r) {
    return r[0] && r[1] && r[1].toString().includes("docs.google.com");
  });

  return { sources: sources, allSourceNames: allSourceNames, tagMap: tagMap, ga4Map: ga4Map, allTags: allTags };
}

// ───────────────────────────────────────────────────────────────
function consolidateLeads() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = readConfig(ss);
  if (!cfg) { showAlert("Config sheet not found!"); return; }

  var result = fetchAndProcess(cfg.sources);
  var uniqueLeads = result.uniqueLeads, allLeads = result.allLeads, dupLeads = result.dupLeads;
  var invalidCount = result.invalidCount;

  writeSheet(ss, "Data",       STANDARD_HEADERS, uniqueLeads.map(formatRow), "#1E8449", true);
  writeSheet(ss, "Raw",        STANDARD_HEADERS, allLeads.map(formatRow),    "#1E3A5F", true);
  writeSheet(ss, "Duplicates", STANDARD_HEADERS, dupLeads.map(formatRow),    "#C0392B");

  storeCache(ss, uniqueLeads, cfg.allSourceNames, dupLeads.length, cfg.ga4Map);
  buildDashboard(ss, uniqueLeads, allLeads, dupLeads, cfg);

  var cfgSheet = ss.getSheetByName("Config");
  cfgSheet.getRange("F1").setValue("Last sync:").setFontWeight("bold");
  cfgSheet.getRange("G1").setValue(new Date()).setNumberFormat("dd/mm/yyyy hh:mm");

  showAlert("Done!\n\nTotal leads: " + allLeads.length + "\nUnique: " + uniqueLeads.length + "\nDuplicates: " + dupLeads.length + "\nInvalid dates: " + invalidCount);
}

// ───────────────────────────────────────────────────────────────
function refreshDashboard() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = readConfig(ss);
  if (!cfg) { showAlert("Config sheet not found!"); return; }

  ss.toast("Fetching data from all sources...", "Refreshing", -1);

  var result = fetchAndProcess(cfg.sources);
  var uniqueLeads = result.uniqueLeads, allLeads = result.allLeads, dupLeads = result.dupLeads;

  writeSheet(ss, "Data",       STANDARD_HEADERS, uniqueLeads.map(formatRow), "#1E8449", true);
  writeSheet(ss, "Raw",        STANDARD_HEADERS, allLeads.map(formatRow),    "#1E3A5F", true);
  writeSheet(ss, "Duplicates", STANDARD_HEADERS, dupLeads.map(formatRow),    "#C0392B");

  storeCache(ss, uniqueLeads, cfg.allSourceNames, dupLeads.length, cfg.ga4Map);

  var dash = ss.getSheetByName("Dashboard");
  if (dash) {
    var selected = dash.getRange(DROP_ROW, DROP_COL).getValue() || "Monthly";
    var selTag   = dash.getRange(DROP_ROW, TAG_COL).getValue()  || "All";
    updateKPIs(dash, uniqueLeads, cfg, cfg.ga4Map);
    redrawAllTables(ss, dash, uniqueLeads, cfg, selected, selTag, null);
  }

  var cfgSheet = ss.getSheetByName("Config");
  cfgSheet.getRange("G1").setValue(new Date()).setNumberFormat("dd/mm/yyyy hh:mm");
  ss.toast("Done! " + uniqueLeads.length + " unique leads", "Refresh complete", 5);
}

// ───────────────────────────────────────────────────────────────
// Point 1 -- Nouvel ordre colonnes : Date, Source, Type, FullName, Phone, Email, Company, City, Country, Address
// Index interne :                     0      9      8      1        2      3       4        5      6        7
function formatRow(r) {
  return [fmtDate(r[0]), r[9], r[8], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[10], r[11], r[12]];
}

// ───────────────────────────────────────────────────────────────
function fetchAndProcess(sources) {
  var todayEnd = new Date(); todayEnd.setHours(23,59,59,999);
  var minDate  = new Date(2020,0,1);
  var allLeads = [], invalidCount = 0;

  for (var si = 0; si < sources.length; si++) {
    var sourceName = sources[si][0], url = sources[si][1];
    try {
      var id = extractSheetId(url);
      var ss = SpreadsheetApp.openById(id);
      var sheet = null;
      var gidMatch = url.toString().match(/[#&?]gid=(\d+)/);
      if (gidMatch) {
        var gid = parseInt(gidMatch[1]);
        var allSheets = ss.getSheets();
        for (var shi = 0; shi < allSheets.length; shi++) {
          if (allSheets[shi].getSheetId() === gid) { sheet = allSheets[shi]; break; }
        }
        if (!sheet) sheet = ss.getSheets()[0];
      } else {
        sheet = ss.getSheets()[0];
      }

      var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
      if (lastRow < 2) continue;

      var values     = sheet.getRange(1,1,lastRow,lastCol).getValues();
      var rawHeaders = values[0].map(function(h){ return h.toString().trim(); });
      var colIndex   = buildColIndex(rawHeaders);
      var dateFormat = detectDateFormat(values, colIndex);
      var imported   = 0;

      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var allEmpty = true;
        for (var ci = 0; ci < row.length; ci++) {
          if (row[ci] !== "" && row[ci] !== null && row[ci] !== undefined) { allEmpty = false; break; }
        }
        if (allEmpty) continue;

        var normalized = normalizeRow(row, rawHeaders, colIndex, sourceName);
        if (!normalized) continue;

        var d = toDateObj(normalized[0], dateFormat);
        if (!d || d < minDate || d > todayEnd) { invalidCount++; continue; }
        normalized[0] = d;
        allLeads.push(normalized);
        imported++;
      }
      Logger.log(sourceName + ": " + imported + " leads");
    } catch(e) {
      Logger.log("ERROR " + sourceName + ": " + e.message);
    }
  }

  // Sort by date DESCENDING -- keep most recent submission
  allLeads.sort(function(a,b){ return b[0] - a[0]; });

  var seenEmail = {}, seenTel = {}, seenNom = {};
  var uniqueLeads = [], dupLeads = [];

  for (var li = 0; li < allLeads.length; li++) {
    var lead = allLeads[li];
    var email = norm(lead[3]), tel = normTel(lead[2]), nom = norm(lead[1]);
    var isDup = false;
    if (email && seenEmail[email]) isDup = true;
    if (tel   && seenTel[tel])     isDup = true;
    if (!email && !tel && nom && seenNom[nom]) isDup = true;
    if (!isDup) {
      if (email) seenEmail[email] = true;
      if (tel)   seenTel[tel]     = true;
      if (!email && !tel && nom) seenNom[nom] = true;
      uniqueLeads.push(lead);
    } else {
      dupLeads.push(lead);
    }
  }
  // Re-sort unique leads chronologically for display
  uniqueLeads.sort(function(a,b){ return a[0] - b[0]; });
  allLeads.sort(function(a,b){ return a[0] - b[0]; });

  return { uniqueLeads: uniqueLeads, allLeads: allLeads, dupLeads: dupLeads, invalidCount: invalidCount };
}

// ───────────────────────────────────────────────────────────────
//  CACHE -- stocke date, source, type
// ───────────────────────────────────────────────────────────────
var CACHE_SHEET_NAME = "LeadCache";

function storeCache(ss, uniqueLeads, allSourceNames, dupsCount, ga4Map) {
  ["Cache","⚙️ Cache","⚙️ Data","LeadCache"].forEach(function(n) {
    var old = ss.getSheetByName(n); if (old) ss.deleteSheet(old);
  });
  var ds = ss.insertSheet(CACHE_SHEET_NAME);
  ds.hideSheet();
  ds.getRange(1,1).setValue(JSON.stringify({
    sources: allSourceNames, dups: dupsCount,
    total: uniqueLeads.length, updated: new Date().toISOString(),
    ga4Map: ga4Map || {}
  }));
  if (uniqueLeads.length > 0) {
    var rows = uniqueLeads.map(function(r) {
      var d = (r[0] instanceof Date) ? r[0] : toDateObj(r[0]);
      var dateStr = "";
      if (d) dateStr = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
      return [dateStr, (r[9]||"Unknown").toString(), (r[8]||"").toString()];
    });
    ds.getRange(2, 1, rows.length, 3).setValues(rows);
  }
}

function loadCache(ss) {
  var ds = ss.getSheetByName(CACHE_SHEET_NAME);
  if (!ds) {
    ["Cache","⚙️ Cache","⚙️ Data"].forEach(function(n) { if (!ds) ds = ss.getSheetByName(n); });
  }
  if (!ds) return null;
  try {
    var metaStr = ds.getRange(1,1).getValue();
    if (!metaStr) return null;
    var meta = JSON.parse(metaStr);
    var last = ds.getLastRow();
    if (last < 2) return { meta: meta, leads: [] };

    var rawRows = ds.getRange(2, 1, last-1, 3).getValues();
    var leads = [];
    for (var i=0; i<rawRows.length; i++) {
      var raw = rawRows[i][0], source = rawRows[i][1]||"Unknown", type = rawRows[i][2]||"";
      if (!raw) continue;
      var d = null;
      var iso = raw.toString().trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (iso) d = new Date(parseInt(iso[1]), parseInt(iso[2])-1, parseInt(iso[3]), 12, 0, 0);
      if (!d && raw instanceof Date && !isNaN(raw)) d = new Date(raw.getFullYear(), raw.getMonth(), raw.getDate(), 12, 0, 0);
      if (!d) d = toDateObj(raw);
      if (!d || isNaN(d)) continue;
      var row = new Array(10).fill("");
      row[0] = d; row[8] = type; row[9] = source;
      leads.push(row);
    }
    var ga4Map = meta.ga4Map || {};
    return { meta: meta, leads: leads, ga4Map: ga4Map };
  } catch(e) { Logger.log("Cache ERROR: " + e.message); return null; }
}


// ───────────────────────────────────────────────────────────────
// PERFORMANCE KPIs -- Conversion Rate, Lead Growth, Top Source
// ───────────────────────────────────────────────────────────────
function computeLeadGrowthRate(leads) {
  // Compare this month vs last month
  var now = new Date();
  var thisM = 0, prevM = 0;
  for (var i = 0; i < leads.length; i++) {
    var d = toDateObj(leads[i][0]); if (!d) continue;
    if (d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth()) thisM++;
    else {
      var prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      if (d.getFullYear() === prev.getFullYear() && d.getMonth() === prev.getMonth()) prevM++;
    }
  }
  if (prevM === 0) return thisM > 0 ? "+100%" : "N/A";
  var rate = ((thisM - prevM) / prevM * 100);
  return (rate >= 0 ? "+" : "") + rate.toFixed(1) + "%";
}

function computeTopSource(leads) {
  // Top source by leads THIS MONTH only
  var now = new Date();
  var counts = {};
  for (var i = 0; i < leads.length; i++) {
    var d = toDateObj(leads[i][0]); if (!d) continue;
    if (d.getFullYear() !== now.getFullYear() || d.getMonth() !== now.getMonth()) continue;
    var src = (leads[i][9] || "Unknown").toString().trim();
    counts[src] = (counts[src] || 0) + 1;
  }
  var best = "", bestN = 0;
  for (var s in counts) { if (counts[s] > bestN) { bestN = counts[s]; best = s; } }
  return best ? best + " (" + bestN + ")" : "N/A";
}

function computeConversionRate(leads, ga4Map, cfg) {
  // Leads this month / GA4 sessions this month -- requires a quick GA4 fetch
  var now = new Date();
  var thisM = 0;
  for (var i = 0; i < leads.length; i++) {
    var d = toDateObj(leads[i][0]); if (!d) continue;
    if (d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth()) thisM++;
  }
  // Try fetching GA4 sessions for this month
  try {
    var startDate = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), 1), "UTC", "yyyy-MM-dd");
    var endDate   = Utilities.formatDate(now, "UTC", "yyyy-MM-dd");
    var totalSessions = 0;
    var ga4Sources = cfg.allSourceNames.filter(function(n) { return !!ga4Map[n]; });
    for (var si = 0; si < ga4Sources.length; si++) {
      var pid = ga4Map[ga4Sources[si]];
      var url = "https://analyticsdata.googleapis.com/v1beta/properties/" + pid + ":runReport";
      var resp = UrlFetchApp.fetch(url, {
        method: "post", contentType: "application/json",
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({
          dateRanges: [{ startDate: startDate, endDate: endDate }],
          metrics: [{ name: "sessions" }], limit: 1
        }),
        muteHttpExceptions: true
      });
      var data = JSON.parse(resp.getContentText());
      if (data.rows && data.rows[0]) totalSessions += parseInt(data.rows[0].metricValues[0].value) || 0;
    }
    if (totalSessions === 0) return "N/A";
    return (thisM / totalSessions * 100).toFixed(2) + "%";
  } catch(e) {
    Logger.log("ConvRate GA4 error: " + e.message);
    return "N/A";
  }
}

function drawPerformanceKPIs(dash, leads, cfg, ga4Map, nbCols) {
  // Row layout:
  // row 7  = KPI labels (already set)
  // row 8  = spacer between bande 1 and bande 2
  // row 9  = performance KPI values
  // row 10 = performance KPI labels

  dash.setRowHeight(8, 12);  // white spacer between bande 1 and bande 2
  dash.setRowHeight(9, 46);  // perf KPI values
  dash.setRowHeight(10, 18); // perf KPI labels
  dash.setRowHeight(11, 12); // white spacer between bande 2 and period selector
  dash.setRowHeight(12, 4);  // micro spacer

  // White spacers
  dash.getRange(8, 1, 1, nbCols+2).setBackground("#FFFFFF");
  dash.getRange(11, 1, 1, nbCols+2).setBackground("#FFFFFF");

  var growthRate = computeLeadGrowthRate(leads);
  var topSource  = computeTopSource(leads);
  var convRate   = computeConversionRate(leads, ga4Map, cfg);

  var isPositive = growthRate.charAt(0) === "+";
  var isNegative = growthRate.charAt(0) === "-";
  var growthColor = isPositive ? "#1E8449" : (isNegative ? "#C0392B" : "#7F8C8D");

  var kpis = [
    { val: convRate,   label: "TRAFFIC -> LEAD RATE",  col: 2, color: "#1A5276",  bg: "#EBF5FB",
      note: "Conversion rate: leads generated this month / total website sessions (GA4).\nExample: 6.5% means 1 lead for every ~15 visitors.\nRequires GA4 configured in Config sheet." },
    { val: growthRate, label: "LEAD GROWTH (MoM)",     col: 4, color: growthColor,
      bg: isPositive ? "#EAFAF1" : (isNegative ? "#FDEDEC" : "#F8F9FA"),
      note: "Month-over-Month growth: compares this month's leads to last month.\n+ means more leads than last month. - means a drop.\nExample: +12% = 12% more leads than previous month." },
    { val: topSource,  label: "TOP SOURCE THIS MONTH", col: 6, color: "#6C3483",  bg: "#F5EEF8",
      note: "The source that generated the most leads this month, with its count.\nUseful to quickly spot your best-performing channel this month." }
  ];

  // KPI tooltips for bande 1 (rows 6-7)
  var band1Notes = [
    { col: 2, note: "Total unique leads collected across all sources since the beginning.\nDuplicates are removed by matching email and phone number." },
    { col: 3, note: "Number of new unique leads received today." },
    { col: 4, note: "Unique leads received since Monday of the current week." },
    { col: 5, note: "Unique leads received since the 1st of the current month." },
    { col: 6, note: "Total unique leads received since January 1st of the current year." }
  ];
  for (var bi = 0; bi < band1Notes.length; bi++) {
    try { dash.getRange(6, band1Notes[bi].col).setNote(band1Notes[bi].note); } catch(e) {}
  }

  for (var i = 0; i < kpis.length; i++) {
    var k = kpis[i];
    var valFontSize = k.val.length > 10 ? 12 : (k.val.length > 7 ? 16 : 22);
    var valCell = dash.getRange(9, k.col, 1, 2).merge();
    valCell.setValue(k.val).setFontSize(valFontSize).setFontWeight("bold").setFontColor(k.color)
           .setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground(k.bg);
    try { valCell.setNote(k.note); } catch(e) {}
    dash.getRange(10, k.col, 1, 2).merge()
        .setValue(k.label).setFontSize(7).setFontColor("#999999").setFontWeight("bold")
        .setHorizontalAlignment("center").setBackground(k.bg);
  }
}

// ───────────────────────────────────────────────────────────────
//  DASHBOARD
// ───────────────────────────────────────────────────────────────
function buildDashboard(ss, unique, all, dups, cfg) {
  var dash = ss.getSheetByName("Dashboard");
  if (!dash) dash = ss.insertSheet("Dashboard");
  dash.clearContents(); dash.clearFormats();
  try { dash.getDataRange().clearDataValidations(); } catch(e) {}
  dash.setHiddenGridlines(true);

  var today = new Date(), yr = today.getFullYear();
  var nbSrc = cfg.allSourceNames.length, nbCols = nbSrc + 2;

  dash.setColumnWidth(1, 14); dash.setColumnWidth(2, 185);
  for (var c = 3; c <= 2+nbSrc; c++) dash.setColumnWidth(c, 115);
  dash.setColumnWidth(2+nbSrc+1, 85);

  // Title
  dash.setRowHeight(1,8); dash.setRowHeight(2,58); dash.setRowHeight(3,20); dash.setRowHeight(4,4);
  setR(dash,2,2,"Lead Reporting Dashboard",{merge:[1,nbCols],fontSize:22,bold:true,fontColor:"#1E3A5F",vAlign:"middle"});
  setR(dash,3,2,"Updated: "+today.toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long",year:"numeric"})+"  ",
    {merge:[1,nbCols],fontSize:9,italic:true,fontColor:"#AAAAAA"});
  dash.getRange(4,2,1,nbCols).setBackground("#1E3A5F");

  // Point 3 -- 5 KPIs sans Duplicates
  dash.setRowHeight(5,8); dash.setRowHeight(6,50); dash.setRowHeight(7,22);
  var kpiDefs = [
    {label:"TOTAL LEADS", col:2, color:"#1E3A5F", bg:"#EBF5FB"},
    {label:"TODAY",       col:3, color:"#E67E22", bg:"#FEF9E7"},
    {label:"THIS WEEK",   col:4, color:"#8E44AD", bg:"#F5EEF8"},
    {label:"THIS MONTH",  col:5, color:"#2980B9", bg:"#EBF5FB"},
    {label:"YEAR "+yr,    col:6, color:"#1E8449", bg:"#EAFAF1"}
  ];
  var kpiVals = [unique.length, countPeriod(unique,"day"), countPeriod(unique,"week"), countPeriod(unique,"month"), countPeriod(unique,"year")];
  for (var ki = 0; ki < kpiDefs.length; ki++) {
    var k = kpiDefs[ki];
    dash.getRange(6,k.col).setValue(kpiVals[ki]).setFontSize(30).setFontWeight("bold").setFontColor(k.color)
        .setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground(k.bg);
    dash.getRange(7,k.col).setValue(k.label).setFontSize(7).setFontColor("#999999").setFontWeight("bold")
        .setHorizontalAlignment("center").setBackground(k.bg);
  }

  // Sélecteurs Period + Tag
  // Performance KPIs row (rows 8-9)
  drawPerformanceKPIs(dash, unique, cfg, cfg.ga4Map, nbCols);

  dash.setRowHeight(DROP_ROW,34); dash.setRowHeight(DROP_ROW-1,4);
  setR(dash,DROP_ROW,2,"Period:",{bold:true,fontSize:10,fontColor:"#1E3A5F",vAlign:"middle",hAlign:"right"});
  dash.getRange(DROP_ROW, DROP_COL)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Daily","Weekly","Monthly","Yearly"],true).setAllowInvalid(false).build())
      .setValue("Monthly").setBackground("#1E3A5F").setFontColor("#FFFFFF").setFontWeight("bold")
      .setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Point 4 -- Sélecteur Tag
  setR(dash,DROP_ROW,4,"Tag:",{bold:true,fontSize:10,fontColor:"#1E3A5F",vAlign:"middle",hAlign:"right"});
  var tagList = ["All"].concat(cfg.allTags);
  dash.getRange(DROP_ROW, TAG_COL)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(tagList,true).setAllowInvalid(false).build())
      .setValue("All").setBackground("#2980B9").setFontColor("#FFFFFF").setFontWeight("bold")
      .setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");

  dash.setRowHeight(15,4);
  dash.getRange(15,2,1,nbCols).setBackground("#E8E8E8");

  redrawAllTables(ss, dash, unique, cfg, "Monthly", "All", null);
}

// ───────────────────────────────────────────────────────────────
function redrawAllTables(ss, dash, leads, cfg, periodLabel, selectedTag, customFilter) {
  var filteredSources = cfg.allSourceNames.filter(function(name) {
    if (!selectedTag || selectedTag === "All") return true;
    return (cfg.tagMap[name] || "").toLowerCase() === selectedTag.toLowerCase();
  });

  var nextRow = TABLE_START_ROW;
  nextRow = drawLeadsTable(dash, leads, filteredSources, periodLabel, customFilter, nextRow);
  nextRow += 2;
  nextRow = drawBusinessTypeTable(dash, leads, filteredSources, periodLabel, customFilter, nextRow);
  nextRow += 2;
  // GA4: respects tag filter like other tables
  drawGA4Table(ss, dash, filteredSources, cfg.ga4Map, periodLabel, customFilter, nextRow);
}

// ───────────────────────────────────────────────────────────────
// TABLEAU 1 -- Leads par Source
// ───────────────────────────────────────────────────────────────
function drawLeadsTable(dash, leads, sourceNames, periodLabel, customFilter, startRow) {
  var nbSrc  = sourceNames.length, nbCols = nbSrc + 2;
  var zone   = dash.getRange(startRow, 2, 300, 50);
  zone.clearContent(); zone.clearFormat(); zone.breakApart();

  var period    = labelToKey(periodLabel);
  var hasFilter = customFilter && customFilter.type && customFilter.type !== "clear";
  var grouped   = hasFilter ? filterCustom(leads, customFilter, sourceNames) : groupBySource(leads, period, sourceNames);

  var r = startRow;
  var titleText = hasFilter ? buildFilterLabel(customFilter) : ("Leads by Source -- " + periodLabel);
  dash.setRowHeight(r, 32);
  setR(dash,r,2,titleText,{merge:[1,nbCols],bg:"#1E3A5F",fontColor:"#FFFFFF",bold:true,fontSize:11,hAlign:"left",vAlign:"middle"});
  r++;
  r = drawCrossTable(dash, grouped.periods, grouped.crossData, sourceNames, r, "#D6EAF8", "#2980B9", "#1E3A5F", "#1E8449", "#EAFAF1");
  return r;
}

// ───────────────────────────────────────────────────────────────
// ───────────────────────────────────────────────────────────────
// TABLEAU 2 -- Leads par Business Type (priority order + merges)
// ───────────────────────────────────────────────────────────────
var BT_PRIORITY = ["Wholesaler","Vape Shop","Tobacco & Smoke Shop","Retail Store","Convenience Store","E-Commerce Retail","Consumer","Unknown"];

// Fusions : chaque label absorbe ses alias (sous-chaîne ou égalité, insensible à la casse)
var BT_MERGE = {
  "Wholesaler":           ["wholesaler","distributor","grossiste","distributeur"],
  "Vape Shop":            ["vape shop","vapeshop","vape store","vapeur"],
  "Tobacco & Smoke Shop": ["tobacco & smoke shop","tobacco shop","smoke shop","tabac"],
  "E-Commerce Retail":    ["e-commerce retail","online retailer","online vape store","online retail","e-commerce","ecommerce"],
  "Retail Store":         ["retail store","retail shop","retailer","magasin"],
  "Convenience Store":    ["convenience store","superette","epicerie"],
  "Consumer":             ["consumer","consommateur","particulier"]
};

function canonicalBT(raw) {
  if (!raw || !raw.toString().trim()) return "Unknown";
  var v = raw.toString().trim().toLowerCase();
  for (var label in BT_MERGE) {
    var aliases = BT_MERGE[label];
    for (var ai = 0; ai < aliases.length; ai++) {
      if (v === aliases[ai] || v.indexOf(aliases[ai]) !== -1) return label;
    }
  }
  // Unrecognised -> keep original label (displayed after Consumer, before Unknown)
  return raw.toString().trim();
}

function drawBusinessTypeTable(dash, leads, sourceNames, periodLabel, customFilter, startRow) {
  var period    = labelToKey(periodLabel);
  var filtered  = filterLeadsByPeriodAndSources(leads, period, sourceNames, customFilter);
  var hasFilter = customFilter && customFilter.type && customFilter.type !== "clear";

  // Build pbtData[periodKey][canonType] = count
  var pbtData = {}, allPeriodKeys = {}, allFoundTypes = {};

  for (var i = 0; i < filtered.length; i++) {
    var d     = toDateObj(filtered[i][0]);
    if (!d) continue;
    var canon = canonicalBT((filtered[i][8] || "").toString().trim());
    var pKey  = hasFilter
      ? (d.getDate() + " " + MONTHS_SH[d.getMonth()] + " " + d.getFullYear())
      : dateToKey(d, period);
    if (!pbtData[pKey])        pbtData[pKey] = {};
    if (!pbtData[pKey][canon]) pbtData[pKey][canon] = 0;
    pbtData[pKey][canon]++;
    allPeriodKeys[pKey] = true;
    allFoundTypes[canon] = true;
  }

  // Period list sorted correctly
  var periods = Object.keys(allPeriodKeys).sort(function(a, b) {
    return hasFilter ? keyToTs(a, "day") - keyToTs(b, "day")
                     : keyToTs(b, period) - keyToTs(a, period);
  });

  // typeList: BT_PRIORITY order first (only those with data),
  //           then unrecognised extra types alphabetically,
  //           Unknown always last
  var priorityTypes = BT_PRIORITY.filter(function(t) {
    return t !== "Unknown" && !!allFoundTypes[t];
  });
  var extraTypes = Object.keys(allFoundTypes)
    .filter(function(t) { return BT_PRIORITY.indexOf(t) === -1; })
    .sort();
  var typeList = priorityTypes.concat(extraTypes);
  if (allFoundTypes["Unknown"]) typeList.push("Unknown");

  var nbTypes = Math.max(typeList.length, 1);
  var nbCols  = nbTypes + 2;

  var zone = dash.getRange(startRow, 2, 250, 50);
  zone.clearContent(); zone.clearFormat(); zone.breakApart();

  var r = startRow;
  dash.setRowHeight(r, 32);
  var btTitle = hasFilter
    ? ("Leads by Business Type -- " + buildFilterLabel(customFilter).replace("Leads by Source -- ", ""))
    : ("Leads by Business Type -- " + periodLabel);
  setR(dash, r, 2, btTitle, {merge:[1,nbCols], bg:"#6C3483", fontColor:"#FFFFFF", bold:true, fontSize:11, hAlign:"left", vAlign:"middle"});
  r++;

  // Header row
  var headerBg = "#E8DAEF", headerFg = "#6C3483";
  dash.setRowHeight(r, 28);
  dash.getRange(r, 2).setValue("Period").setBackground(headerBg).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setFontColor(headerFg);
  for (var ti = 0; ti < typeList.length; ti++) {
    dash.getRange(r, 3 + ti).setValue(typeList[ti]).setBackground(headerBg).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setFontColor(headerFg);
  }
  dash.getRange(r, 2 + nbTypes + 1).setValue("TOTAL").setBackground(headerFg).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
  r++;

  if (periods.length === 0) {
    dash.setRowHeight(r, 34);
    setR(dash, r, 2, "No data for this selection.", {merge:[1,nbCols], fontColor:"#AAAAAA", italic:true, hAlign:"center", bg:"#FAFAFA", fontSize:10});
    r++;
  } else {
    var values = [], bgColors = [], fntColors = [];
    for (var pi = 0; pi < periods.length; pi++) {
      var pKey2 = periods[pi], rowD = pbtData[pKey2] || {}, bg = pi % 2 === 0 ? "#FFFFFF" : "#F8F9FA";
      var total = 0, rowV = [pKey2], rowB = [bg], rowF = [headerFg];
      for (var tj = 0; tj < typeList.length; tj++) {
        var cnt = rowD[typeList[tj]] || 0; total += cnt;
        rowV.push(cnt > 0 ? cnt : "-");
        rowB.push(bg);
        rowF.push(cnt > 0 ? "#8E44AD" : "#CCCCCC");
      }
      rowV.push(total);
      rowB.push(total > 0 ? "#F5EEF8" : bg);
      rowF.push(total > 0 ? "#6C3483" : "#CCCCCC");
      values.push(rowV); bgColors.push(rowB); fntColors.push(rowF);
    }
    var dr = dash.getRange(r, 2, periods.length, nbCols);
    dr.setValues(values).setBackgrounds(bgColors).setFontColors(fntColors).setFontSize(9).setHorizontalAlignment("center");
    dash.getRange(r, 2, periods.length, 1).setHorizontalAlignment("left").setFontWeight("bold");
    for (var ri = 0; ri < periods.length; ri++) dash.setRowHeight(r + ri, 24);
    r += periods.length;
  }

  // TOTAL row
  dash.setRowHeight(r, 28);
  var totV = ["TOTAL"], grand = 0;
  for (var tk = 0; tk < typeList.length; tk++) {
    var t = 0;
    for (var pk2 = 0; pk2 < periods.length; pk2++) t += ((pbtData[periods[pk2]] || {})[typeList[tk]] || 0);
    grand += t; totV.push(t);
  }
  totV.push(grand);
  dash.getRange(r, 2, 1, nbCols).setValues([totV]).setBackground(headerFg).setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(9);
  dash.getRange(r, 2).setHorizontalAlignment("left");
  dash.getRange(r, 2 + nbTypes + 1).setBackground("#6C3483");
  drawBorder(dash, startRow + 1, 2, r - startRow, nbCols, "#999999");
  return r + 1;
}

// ───────────────────────────────────────────────────────────────
// TABLEAU 3 -- Point 2 : GA4 Visits
// ───────────────────────────────────────────────────────────────
function drawGA4Table(ss, dash, sourceNames, ga4Map, periodLabel, customFilter, startRow) {
  var ga4Sources = sourceNames.filter(function(name) { return !!ga4Map[name]; });
  var nbSrc  = Math.max(ga4Sources.length, 1), nbCols = nbSrc + 2;
  var zone   = dash.getRange(startRow, 2, 250, 50);
  zone.clearContent(); zone.clearFormat(); zone.breakApart();

  var r = startRow;
  dash.setRowHeight(r, 32);
  var hasFilter = customFilter && customFilter.type && customFilter.type !== "clear";
  var ga4Title = hasFilter ? ("Website Visits GA4 -- " + buildFilterLabel(customFilter).replace("Leads by Source -- ","")) : ("Website Visits GA4 -- " + periodLabel);
  setR(dash,r,2,ga4Title,{merge:[1,nbCols],bg:"#1A5276",fontColor:"#FFFFFF",bold:true,fontSize:11,hAlign:"left",vAlign:"middle"});
  r++;

  if (ga4Sources.length === 0) {
    dash.setRowHeight(r,34);
    setR(dash,r,2,"No GA4 Property IDs configured. Add them in column D of the Config sheet.",
      {merge:[1,nbCols],fontColor:"#AAAAAA",italic:true,hAlign:"center",bg:"#FAFAFA",fontSize:10});
    return r + 2;
  }

  var ga4Data = fetchGA4Data(ga4Sources, ga4Map, periodLabel, customFilter);

  dash.setRowHeight(r, 22);
  setR(dash,r,2,"Sessions",{merge:[1,ga4Sources.length+2],bg:"#2471A3",fontColor:"#FFFFFF",bold:true,fontSize:9,hAlign:"left",vAlign:"middle"});
  r++;
  r = drawGA4SubTable(dash, ga4Data, ga4Sources, "sessions", r);
  r++;
  dash.setRowHeight(r, 22);
  setR(dash,r,2,"Unique Users",{merge:[1,ga4Sources.length+2],bg:"#2471A3",fontColor:"#FFFFFF",bold:true,fontSize:9,hAlign:"left",vAlign:"middle"});
  r++;
  r = drawGA4SubTable(dash, ga4Data, ga4Sources, "users", r);
  return r + 1;
}

function drawGA4SubTable(dash, ga4Data, sourceNames, metric, startRow) {
  var nbSrc  = sourceNames.length, nbCols = nbSrc + 2, r = startRow;
  dash.setRowHeight(r, 26);
  dash.getRange(r,2).setValue("Period").setBackground("#D6EAF8").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setFontColor("#1A5276");
  for (var si=0; si<sourceNames.length; si++) {
    dash.getRange(r,3+si).setValue(sourceNames[si]).setBackground("#D6EAF8").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setFontColor("#1A5276");
  }
  dash.getRange(r,2+nbSrc+1).setValue("TOTAL").setBackground("#1A5276").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
  r++;

  var periods = ga4Data.periods || [], crossData = ga4Data[metric] || {};

  if (periods.length === 0) {
    dash.setRowHeight(r,28);
    setR(dash,r,2,"No GA4 data available.",{merge:[1,nbCols],fontColor:"#AAAAAA",italic:true,hAlign:"center",bg:"#FAFAFA",fontSize:9});
    return r + 1;
  }

  var values=[], bgColors=[], fntColors=[];
  for (var pi=0; pi<periods.length; pi++) {
    var key=periods[pi], rowD=crossData[key]||{}, bg=pi%2===0?"#FFFFFF":"#F8F9FA";
    var total=0, rowV=[key], rowB=[bg], rowF=["#1A5276"];
    for (var sj=0; sj<sourceNames.length; sj++) {
      var cnt=rowD[sourceNames[sj]]||0; total+=cnt;
      rowV.push(cnt>0?cnt:"-"); rowB.push(bg); rowF.push(cnt>0?"#2471A3":"#CCCCCC");
    }
    rowV.push(total); rowB.push(total>0?"#EBF5FB":bg); rowF.push(total>0?"#1A5276":"#CCCCCC");
    values.push(rowV); bgColors.push(rowB); fntColors.push(rowF);
  }
  var dr = dash.getRange(r, 2, periods.length, nbCols);
  dr.setValues(values).setBackgrounds(bgColors).setFontColors(fntColors).setFontSize(9).setHorizontalAlignment("center");
  dash.getRange(r, 2, periods.length, 1).setHorizontalAlignment("left").setFontWeight("bold");
  for (var ri=0; ri<periods.length; ri++) dash.setRowHeight(r+ri, 22);
  r += periods.length;

  dash.setRowHeight(r, 26);
  var totV=["TOTAL"], grand=0;
  for (var tk=0; tk<sourceNames.length; tk++) {
    var t=0;
    for (var pk=0; pk<periods.length; pk++) t+=((crossData[periods[pk]]||{})[sourceNames[tk]]||0);
    grand+=t; totV.push(t);
  }
  totV.push(grand);
  dash.getRange(r,2,1,nbCols).setValues([totV]).setBackground("#1A5276").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(9);
  drawBorder(dash, startRow, 2, r-startRow+1, nbCols, "#2471A3");
  return r + 1;
}

// ───────────────────────────────────────────────────────────────
// FETCH GA4 via Analytics Data API v1beta
// ───────────────────────────────────────────────────────────────
function fetchGA4Data(sourceNames, ga4Map, periodLabel, customFilter) {
  var period    = labelToKey(periodLabel);
  var hasFilter = customFilter && customFilter.type && customFilter.type !== "clear";
  var today     = new Date();
  var endDate   = Utilities.formatDate(today, "UTC", "yyyy-MM-dd");
  var startDate;

  if (hasFilter && customFilter.type === "range") {
    startDate = customFilter.from; endDate = customFilter.to;
  } else if (period === "day")   { var d=new Date(); d.setDate(d.getDate()-89); startDate=Utilities.formatDate(d,"UTC","yyyy-MM-dd"); }
  else if (period === "week")  { var d2=new Date(); d2.setMonth(d2.getMonth()-6); startDate=Utilities.formatDate(d2,"UTC","yyyy-MM-dd"); }
  else if (period === "month") { var d3=new Date(); d3.setMonth(d3.getMonth()-23); d3.setDate(1); startDate=Utilities.formatDate(d3,"UTC","yyyy-MM-dd"); }
  else { startDate = "2020-01-01"; }

  // Custom date filter -> always daily granularity; otherwise follow period dropdown
  var ga4Dim = hasFilter ? "date" : (period==="day" ? "date" : period==="week" ? "isoYearIsoWeek" : period==="month" ? "yearMonth" : "year");

  var result = { periods: [], sessions: {}, users: {} };
  var allKeys = {};

  for (var si=0; si<sourceNames.length; si++) {
    var name = sourceNames[si], pid = ga4Map[name];
    if (!pid) continue;
    try {
      var url = "https://analyticsdata.googleapis.com/v1beta/properties/" + pid + ":runReport";
      var payload = {
        dateRanges: [{startDate: startDate, endDate: endDate}],
        dimensions: [{name: ga4Dim}],
        metrics: [{name:"sessions"},{name:"totalUsers"}],
        orderBys: [{dimension:{dimensionName:ga4Dim},desc:true}],
        limit: 200
      };
      var resp = UrlFetchApp.fetch(url, {
        method:"post", contentType:"application/json",
        headers:{Authorization:"Bearer "+ScriptApp.getOAuthToken()},
        payload:JSON.stringify(payload), muteHttpExceptions:true
      });
      var data = JSON.parse(resp.getContentText());
      if (data.error) {
        Logger.log("GA4 API ERROR for " + name + " (pid:" + pid + "): " + data.error.code + " - " + data.error.message);
      }
      if (data.rows) {
        data.rows.forEach(function(row) {
          var dim  = row.dimensionValues[0].value;
          var sess = parseInt(row.metricValues[0].value)||0;
          var usr  = parseInt(row.metricValues[1].value)||0;
          var key  = ga4DimToKey(dim, ga4Dim);
          if (!result.sessions[key]) result.sessions[key]={};
          if (!result.users[key])    result.users[key]={};
          result.sessions[key][name] = (result.sessions[key][name]||0) + sess;
          result.users[key][name]    = (result.users[key][name]||0)    + usr;
          allKeys[key] = true;
        });
      }
    } catch(e) { Logger.log("GA4 ERROR " + name + ": " + e.message); }
  }

  result.periods = Object.keys(allKeys).sort(function(a,b){ return keyToTs(b,period)-keyToTs(a,period); });
  return result;
}

function ga4DimToKey(val, dimension) {
  if (dimension === "date") {
    var y=parseInt(val.substring(0,4)), m=parseInt(val.substring(4,6))-1, d=parseInt(val.substring(6,8));
    return d + " " + MONTHS_SH[m] + " " + y;
  }
  if (dimension === "yearMonth") {
    var y2=parseInt(val.substring(0,4)), m2=parseInt(val.substring(4,6))-1;
    return MONTHS_EN[m2] + " " + y2;
  }
  if (dimension === "isoYearIsoWeek") {
    var y3=parseInt(val.substring(0,4)), wk=parseInt(val.substring(4,6));
    var mon=isoWeekToMonday(y3,wk), sun=new Date(mon.getFullYear(),mon.getMonth(),mon.getDate()+6);
    if (mon.getMonth()===sun.getMonth()) return MONTHS_EN[mon.getMonth()]+" "+mon.getDate()+" - "+sun.getDate()+", "+sun.getFullYear();
    return MONTHS_SH[mon.getMonth()]+" "+mon.getDate()+" - "+MONTHS_SH[sun.getMonth()]+" "+sun.getDate()+", "+sun.getFullYear();
  }
  return val;
}

function isoWeekToMonday(year, week) {
  var jan4=new Date(year,0,4), dow=jan4.getDay()||7;
  var mon1=new Date(year,0,4-dow+1);
  return new Date(mon1.getFullYear(),mon1.getMonth(),mon1.getDate()+(week-1)*7,12);
}

// ───────────────────────────────────────────────────────────────
// SHARED TABLE RENDERER
// ───────────────────────────────────────────────────────────────
function drawCrossTable(dash, periods, crossData, sourceNames, startRow, headerBg, dataColor, headerColor, totalBg, totalRowBg) {
  var nbSrc  = sourceNames.length, nbCols = nbSrc + 2, r = startRow;

  dash.setRowHeight(r, 28);
  dash.getRange(r,2).setValue("Period").setBackground(headerBg).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setFontColor(headerColor);
  for (var si=0; si<sourceNames.length; si++) {
    dash.getRange(r,3+si).setValue(sourceNames[si]).setBackground(headerBg).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setFontColor(headerColor);
  }
  dash.getRange(r,2+nbSrc+1).setValue("TOTAL").setBackground(headerColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
  r++;

  if (periods.length === 0) {
    dash.setRowHeight(r,34);
    setR(dash,r,2,"No data for this selection.",{merge:[1,nbCols],fontColor:"#AAAAAA",italic:true,hAlign:"center",bg:"#FAFAFA",fontSize:10});
    r++;
  } else {
    var values=[], bgColors=[], fntColors=[];
    for (var pi=0; pi<periods.length; pi++) {
      var key=periods[pi], rowD=crossData[key]||{}, bg=pi%2===0?"#FFFFFF":"#F8F9FA";
      var total=0, rowV=[key], rowB=[bg], rowF=[headerColor];
      for (var sj=0; sj<sourceNames.length; sj++) {
        var cnt=rowD[sourceNames[sj]]||0; total+=cnt;
        rowV.push(cnt>0?cnt:"-"); rowB.push(bg); rowF.push(cnt>0?dataColor:"#CCCCCC");
      }
      rowV.push(total); rowB.push(total>0?totalRowBg:bg); rowF.push(total>0?totalBg:"#CCCCCC");
      values.push(rowV); bgColors.push(rowB); fntColors.push(rowF);
    }
    var dr = dash.getRange(r, 2, periods.length, nbCols);
    dr.setValues(values).setBackgrounds(bgColors).setFontColors(fntColors).setFontSize(9).setHorizontalAlignment("center");
    dash.getRange(r, 2, periods.length, 1).setHorizontalAlignment("left").setFontWeight("bold");
    for (var ri=0; ri<periods.length; ri++) dash.setRowHeight(r+ri, 24);
    r += periods.length;
  }

  dash.setRowHeight(r, 28);
  var totV=["TOTAL"], grand=0;
  for (var tk=0; tk<sourceNames.length; tk++) {
    var t=0;
    for (var pk=0; pk<periods.length; pk++) t+=((crossData[periods[pk]]||{})[sourceNames[tk]]||0);
    grand+=t; totV.push(t);
  }
  totV.push(grand);
  dash.getRange(r,2,1,nbCols).setValues([totV]).setBackground(headerColor).setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(9);
  dash.getRange(r,2).setHorizontalAlignment("left");
  dash.getRange(r,2+nbSrc+1).setBackground(totalBg);

  drawBorder(dash, startRow-1, 2, r-startRow+2, nbCols, "#999999");
  return r + 1;
}

function drawBorder(dash, row, col, numRows, numCols, color) {
  if (numRows < 1 || numCols < 1) return;
  try {
    dash.getRange(row, col, numRows, numCols)
        .setBorder(true,true,true,true,null,null,color,SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
        .setBorder(null,null,null,null,true,true,color,SpreadsheetApp.BorderStyle.SOLID);
  } catch(e) { Logger.log("Border: "+e.message); }
}

// ───────────────────────────────────────────────────────────────
// HELPERS DATA
// ───────────────────────────────────────────────────────────────
function filterLeadsByPeriodAndSources(leads, period, sourceNames, customFilter) {
  var todayEnd = new Date(); todayEnd.setHours(23,59,59,999);
  var hasFilter = customFilter && customFilter.type && customFilter.type !== "clear";
  var limitDate = null;
  if (!hasFilter) {
    if (period==="day")   { limitDate=new Date(); limitDate.setDate(limitDate.getDate()-89); limitDate.setHours(0,0,0,0); }
    if (period==="week")  { limitDate=new Date(); limitDate.setMonth(limitDate.getMonth()-6); limitDate.setHours(0,0,0,0); }
    if (period==="month") { limitDate=new Date(); limitDate.setMonth(limitDate.getMonth()-23); limitDate.setDate(1); limitDate.setHours(0,0,0,0); }
  }
  var fromD, toD;
  if (hasFilter && customFilter.type==="range") {
    var fp=customFilter.from.split("-"), tp=customFilter.to.split("-");
    fromD=new Date(+fp[0],+fp[1]-1,+fp[2],0,0,0); toD=new Date(+tp[0],+tp[1]-1,+tp[2],23,59,59);
  }
  var srcSet={};
  sourceNames.forEach(function(s){ srcSet[s]=true; });
  return leads.filter(function(lead) {
    var src=(lead[9]||"").toString().trim();
    if (!srcSet[src]) return false;
    var d=toDateObj(lead[0]);
    if (!d || d>todayEnd || d.getFullYear()<2020) return false;
    if (hasFilter) return d>=fromD && d<=toD;
    return !limitDate || d>=limitDate;
  });
}

function groupBySource(leads, period, sourceNames) {
  var crossData={}, todayEnd=new Date(); todayEnd.setHours(23,59,59,999);
  var limitDate=null;
  if (period==="day")   { limitDate=new Date(); limitDate.setDate(limitDate.getDate()-89); limitDate.setHours(0,0,0,0); }
  if (period==="week")  { limitDate=new Date(); limitDate.setMonth(limitDate.getMonth()-6); limitDate.setHours(0,0,0,0); }
  if (period==="month") { limitDate=new Date(); limitDate.setMonth(limitDate.getMonth()-23); limitDate.setDate(1); limitDate.setHours(0,0,0,0); }
  var srcSet={};
  if (sourceNames) sourceNames.forEach(function(s){ srcSet[s]=true; });
  for (var i=0; i<leads.length; i++) {
    var src=(leads[i][9]||"Unknown").toString().trim();
    if (sourceNames && !srcSet[src]) continue;
    var d=toDateObj(leads[i][0]);
    if (!d || d>todayEnd || d.getFullYear()<2020) continue;
    if (limitDate && d<limitDate) continue;
    var key=dateToKey(d,period);
    if (!crossData[key]) crossData[key]={};
    crossData[key][src]=(crossData[key][src]||0)+1;
  }
  var periods=Object.keys(crossData).sort(function(a,b){ return keyToTs(b,period)-keyToTs(a,period); });
  return { periods:periods, crossData:crossData };
}

function filterCustom(leads, f, sourceNames) {
  var crossData={}, todayEnd=new Date(); todayEnd.setHours(23,59,59,999);
  var fromD, toD;
  if (f.type==="range") {
    var fp=f.from.split("-"), tp=f.to.split("-");
    fromD=new Date(+fp[0],+fp[1]-1,+fp[2],0,0,0); toD=new Date(+tp[0],+tp[1]-1,+tp[2],23,59,59);
  }
  var srcSet={};
  if (sourceNames) sourceNames.forEach(function(s){ srcSet[s]=true; });
  for (var i=0; i<leads.length; i++) {
    var src=(leads[i][9]||"Unknown").toString().trim();
    if (sourceNames && !srcSet[src]) continue;
    var d=toDateObj(leads[i][0]);
    if (!d || d>todayEnd || d.getFullYear()<2020) continue;
    if (f.type==="range" && (d<fromD || d>toD)) continue;
    var key=d.getDate()+" "+MONTHS_SH[d.getMonth()]+" "+d.getFullYear();
    if (!crossData[key]) crossData[key]={};
    crossData[key][src]=(crossData[key][src]||0)+1;
  }
  var periods=Object.keys(crossData).sort(function(a,b){ return keyToTs(a,"day")-keyToTs(b,"day"); });
  return { periods:periods, crossData:crossData };
}

// ───────────────────────────────────────────────────────────────
function updateKPIs(dash, unique, cfg, ga4Map) {
  if (!dash) return;
  var today=new Date();
  var vals=[unique.length,countPeriod(unique,"day"),countPeriod(unique,"week"),countPeriod(unique,"month"),countPeriod(unique,"year")];
  for (var i=0; i<5; i++) dash.getRange(6,i+2).setValue(vals[i]);
  setR(dash,3,2,"Updated: "+today.toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long",year:"numeric"})+"  ",
    {merge:[1,10],fontSize:9,italic:true,fontColor:"#AAAAAA"});
  if (cfg && ga4Map) {
    var growthRate = computeLeadGrowthRate(unique);
    var topSource  = computeTopSource(unique);
    var isPositive = growthRate.charAt(0) === "+";
    var isNegative = growthRate.charAt(0) === "-";
    var growthColor = isPositive ? "#1E8449" : (isNegative ? "#C0392B" : "#7F8C8D");
    dash.getRange(8,4,1,2).merge().setValue(growthRate).setFontWeight("bold")
        .setFontColor(growthColor).setHorizontalAlignment("center").setVerticalAlignment("middle")
        .setBackground(isPositive ? "#EAFAF1" : (isNegative ? "#FDEDEC" : "#F8F9FA"));
    dash.getRange(8,6,1,2).merge().setValue(topSource).setFontSize(topSource.length>8?12:22)
        .setFontWeight("bold").setFontColor("#6C3483").setHorizontalAlignment("center")
        .setVerticalAlignment("middle").setBackground("#F5EEF8");
  }
}

// ───────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
// onDashboardEdit -- INSTALLABLE trigger (runs as owner, has OAuth)
// Set up via menu: "Setup installable trigger"
// Simple onEdit is kept as fallback for leads/BT only (no GA4)
// ─────────────────────────────────────────────────────────────
function onEdit(e) {
  // Simple trigger: redraws leads + business type only (no GA4, no OAuth needed)
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "Dashboard") return;
  var row = e.range.getRow(), col = e.range.getColumn();
  if (row !== DROP_ROW || (col !== DROP_COL && col !== TAG_COL)) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cache = loadCache(ss); if (!cache) return;
  var cfg = readConfig(ss);
  if (cache.ga4Map) {
    for (var k in cache.ga4Map) { if (!cfg.ga4Map[k]) cfg.ga4Map[k] = cache.ga4Map[k]; }
  }
  var selected = sheet.getRange(DROP_ROW, DROP_COL).getValue() || "Monthly";
  var selTag   = sheet.getRange(DROP_ROW, TAG_COL).getValue()  || "All";
  // Redraw leads + business type (no GA4 -- needs OAuth which simple trigger lacks)
  var filteredSources = cfg.allSourceNames.filter(function(name) {
    if (!selTag || selTag === "All") return true;
    return (cfg.tagMap[name] || "").toLowerCase() === selTag.toLowerCase();
  });
  var nextRow = TABLE_START_ROW;
  nextRow = drawLeadsTable(sheet, cache.leads, filteredSources, selected, null, nextRow);
  nextRow += 2;
  drawBusinessTypeTable(sheet, cache.leads, filteredSources, selected, null, nextRow);
}

function onDashboardEdit(e) {
  // Installable trigger: full OAuth -> can call GA4 API
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "Dashboard") return;
  var row = e.range.getRow(), col = e.range.getColumn();
  if (row !== DROP_ROW || (col !== DROP_COL && col !== TAG_COL)) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cache = loadCache(ss); if (!cache) return;
  var cfg = readConfig(ss);
  if (cache.ga4Map) {
    for (var k in cache.ga4Map) { if (!cfg.ga4Map[k]) cfg.ga4Map[k] = cache.ga4Map[k]; }
  }
  var selected = sheet.getRange(DROP_ROW, DROP_COL).getValue() || "Monthly";
  var selTag   = sheet.getRange(DROP_ROW, TAG_COL).getValue()  || "All";
  redrawAllTables(ss, sheet, cache.leads, cfg, selected, selTag, null);
}

function applyDateFilter(params) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var cache=loadCache(ss);
  if (!cache) throw new Error("No data found. Please run a full consolidation first.");
  var cfg=readConfig(ss);
  if (cache.ga4Map) {
    for (var k in cache.ga4Map) {
      if (!cfg.ga4Map[k]) cfg.ga4Map[k] = cache.ga4Map[k];
    }
  }
  var dash=ss.getSheetByName("Dashboard");
  var selected=dash.getRange(DROP_ROW,DROP_COL).getValue()||"Monthly";
  var selTag=dash.getRange(DROP_ROW,TAG_COL).getValue()||"All";
  if (params.type==="clear") {
    redrawAllTables(ss, dash, cache.leads, cfg, selected, selTag, null);
    return "Filter cleared.";
  }
  redrawAllTables(ss, dash, cache.leads, cfg, selected, selTag, params);
  return "Filter applied successfully!";
}

// ───────────────────────────────────────────────────────────────
// DATE UTILS
// ───────────────────────────────────────────────────────────────
function buildFilterLabel(f) {
  if (f.type==="range") {
    var fp=f.from.split("-"), tp=f.to.split("-");
    var fl=+fp[2]+" "+MONTHS_SH[+fp[1]-1]+" "+fp[0], tl=+tp[2]+" "+MONTHS_SH[+tp[1]-1]+" "+tp[0];
    if (f.from===f.to) return "Leads by Source -- "+fl;
    return "Leads by Source -- "+fl+" to "+tl;
  }
  return "Leads by Source -- Custom Filter";
}

function dateToKey(d, period) {
  if (period==="day")   return d.getDate()+" "+MONTHS_SH[d.getMonth()]+" "+d.getFullYear();
  if (period==="week") {
    var mon=getMondayOf(d), sun=new Date(mon.getFullYear(),mon.getMonth(),mon.getDate()+6);
    if (mon.getMonth()===sun.getMonth()) return MONTHS_EN[mon.getMonth()]+" "+mon.getDate()+" - "+sun.getDate()+", "+sun.getFullYear();
    return MONTHS_SH[mon.getMonth()]+" "+mon.getDate()+" - "+MONTHS_SH[sun.getMonth()]+" "+sun.getDate()+", "+sun.getFullYear();
  }
  if (period==="month") return MONTHS_EN[d.getMonth()]+" "+d.getFullYear();
  if (period==="year")  return d.getFullYear().toString();
  return "";
}

function getMondayOf(d) {
  var day=d.getDay(), diff=(day===0)?-6:1-day;
  return new Date(d.getFullYear(),d.getMonth(),d.getDate()+diff,12);
}

function keyToTs(key, period) {
  try {
    if (period==="year")  return new Date(parseInt(key),0,1).getTime();
    if (period==="month") { var p=key.split(" "); return new Date(parseInt(p[1]),MONTHS_EN.indexOf(p[0]),1).getTime(); }
    if (period==="day")   { var p2=key.split(" "); return new Date(parseInt(p2[2]),MONTHS_SH.indexOf(p2[1]),parseInt(p2[0]),12).getTime(); }
    if (period==="week") {
      var ym=key.match(/(\d{4})$/); if (!ym) return 0; var yr=parseInt(ym[1]);
      var sm=key.match(/^([A-Za-z]+)\s+(\d+)\s*-\s*\d+,\s*\d{4}$/);
      if (sm) { var mi=MONTHS_EN.indexOf(sm[1]); if(mi===-1)mi=MONTHS_SH.indexOf(sm[1]); return new Date(yr,mi,parseInt(sm[2]),12).getTime(); }
      var dm=key.match(/^([A-Za-z]+)\s+(\d+)\s*-\s*([A-Za-z]+)\s+(\d+),\s*(\d{4})$/);
      if (dm) {
        var si2=MONTHS_EN.indexOf(dm[1]); if(si2===-1)si2=MONTHS_SH.indexOf(dm[1]);
        var ei=MONTHS_EN.indexOf(dm[3]);  if(ei===-1)ei=MONTHS_SH.indexOf(dm[3]);
        var sy=parseInt(dm[5]); if(si2>ei)sy=sy-1;
        return new Date(sy,si2,parseInt(dm[2]),12).getTime();
      }
    }
  } catch(e) { Logger.log("keyToTs: "+e.message); }
  return 0;
}

function countPeriod(leads, period) {
  var now=new Date(), today=new Date(now.getFullYear(),now.getMonth(),now.getDate(),12), count=0;
  for (var i=0; i<leads.length; i++) {
    var d=toDateObj(leads[i][0]); if (!d) continue;
    var ld=new Date(d.getFullYear(),d.getMonth(),d.getDate(),12);
    if      (period==="day"   && ld.getTime()===today.getTime()) count++;
    else if (period==="week") { var mon=getMondayOf(today),sun=new Date(mon.getFullYear(),mon.getMonth(),mon.getDate()+6,23); if(ld>=mon&&ld<=sun) count++; }
    else if (period==="month" && ld.getMonth()===today.getMonth()&&ld.getFullYear()===today.getFullYear()) count++;
    else if (period==="year"  && ld.getFullYear()===today.getFullYear()) count++;
  }
  return count;
}

function detectDateFormat(values, colIndex) {
  var dd=0, mm=0, lim=Math.min(values.length,21);
  for (var i=1; i<lim; i++) {
    var raw=colIndex.date>=0?values[i][colIndex.date]:null;
    if (!raw||raw instanceof Date) continue;
    var sep=raw.toString().trim().match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
    if (!sep) continue;
    var a=+sep[1],b=+sep[2];
    if(a>12&&b<=12)dd+=3; if(b>12&&a<=12)mm+=3;
    if(a<=12&&b<=12&&a>b)dd++; if(a<=12&&b<=12&&b>a)mm++;
  }
  return mm>dd*2?"MM/DD/YYYY":"DD/MM/YYYY";
}

function toDateObj(val, forceFormat) {
  if (!val) return null;
  if (val instanceof Date) { if(isNaN(val))return null; return new Date(val.getFullYear(),val.getMonth(),val.getDate(),12); }
  var str=val.toString().trim();
  if (!str||str==="Invalid Date") return null;
  if (/^\d{4}-\d{2}-\d{2}T/.test(str)) { var d=new Date(str); if(isNaN(d))return null; return new Date(d.getFullYear(),d.getMonth(),d.getDate(),12); }
  var iso=str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return new Date(+iso[1],+iso[2]-1,+iso[3],12);
  var sep=str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (sep) {
    var a=+sep[1],b=+sep[2],y=+sep[3];
    if(a>12)return new Date(y,b-1,a,12); if(b>12)return new Date(y,a-1,b,12);
    if(forceFormat==="MM/DD/YYYY")return new Date(y,a-1,b,12);
    return new Date(y,b-1,a,12);
  }
  var MOIS={"janvier":0,"février":1,"fevrier":1,"mars":2,"avril":3,"mai":4,"juin":5,"juillet":6,"août":7,"aout":7,"septembre":8,"octobre":9,"novembre":10,"décembre":11,"decembre":11};
  var frT=str.match(/^(\d{1,2})(?:er|ème|e)?\s+([a-zéûàôùêîèä]+)\s+(\d{4})$/i);
  if (frT) { var mi=MOIS[frT[2].toLowerCase()]; if(mi!==undefined)return new Date(+frT[3],mi,+frT[1],12); }
  var MONS={"january":0,"february":1,"march":2,"april":3,"may":4,"june":5,"july":6,"august":7,"september":8,"october":9,"november":10,"december":11,"jan":0,"feb":1,"mar":2,"apr":3,"jun":5,"jul":6,"aug":7,"sep":8,"oct":9,"nov":10,"dec":11};
  var enA=str.match(/^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})$/);
  if (enA) { var ma=MONS[enA[1].toLowerCase()]; if(ma!==undefined)return new Date(+enA[3],ma,+enA[2],12); }
  var enB=str.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (enB) { var mb=MONS[enB[2].toLowerCase()]; if(mb!==undefined)return new Date(+enB[3],mb,+enB[1],12); }
  if (/^\d{5}$/.test(str)) { var ds=new Date((+str-25569)*86400000); if(!isNaN(ds))return new Date(ds.getFullYear(),ds.getMonth(),ds.getDate(),12); }
  var fb=new Date(str); if (!isNaN(fb))return new Date(fb.getFullYear(),fb.getMonth(),fb.getDate(),12);
  return null;
}

function fmtDate(val) {
  var d=toDateObj(val); if(!d)return val?val.toString():"";
  return String(d.getDate()).padStart(2,"0")+" "+MONTHS_SH[d.getMonth()]+" "+d.getFullYear();
}

function setR(sheet, row, col, value, opts) {
  opts=opts||{};
  var r=opts.merge?sheet.getRange(row,col,opts.merge[0],opts.merge[1]):sheet.getRange(row,col);
  if(opts.merge&&(opts.merge[0]>1||opts.merge[1]>1))r.merge();
  if(value!==undefined)r.setValue(value);
  if(opts.bg)       r.setBackground(opts.bg);
  if(opts.fontColor)r.setFontColor(opts.fontColor);
  if(opts.fontSize) r.setFontSize(opts.fontSize);
  if(opts.bold)     r.setFontWeight("bold");
  if(opts.italic)   r.setFontStyle("italic");
  if(opts.hAlign)   r.setHorizontalAlignment(opts.hAlign);
  if(opts.vAlign)   r.setVerticalAlignment(opts.vAlign);
}

function labelToKey(l) { return {"Daily":"day","Weekly":"week","Monthly":"month","Yearly":"year"}[l]||"month"; }

function writeSheet(ss, name, headers, rows, color, sortDesc) {
  var s=ss.getSheetByName(name); if(!s)s=ss.insertSheet(name);
  s.clearContents(); s.clearFormats();
  if(s.getFilter())s.getFilter().remove();
  if(!rows.length){s.getRange(1,1).setValue("No data");return;}
  // Sort by date DESC (most recent first) when sortDesc=true
  if (sortDesc) {
    rows = rows.slice().sort(function(a, b) {
      var da = a[0] ? new Date(a[0]) : new Date(0);
      var db = b[0] ? new Date(b[0]) : new Date(0);
      return db - da;
    });
  }
  var data=[headers].concat(rows);
  s.getRange(1,1,data.length,headers.length).setValues(data);
  s.getRange(1,1,1,headers.length).setBackground(color).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center");
  for(var i=2;i<=rows.length+1;i++)s.getRange(i,1,1,headers.length).setBackground(i%2===0?"#F2F3F4":"#FFFFFF");
  s.setFrozenRows(1); s.autoResizeColumns(1,headers.length);
  s.getRange(1,1,data.length,headers.length).createFilter();
}

function buildColIndex(rawHeaders) {
  // normH: lowercase + trim + normalize smart apostrophes/quotes to straight ones
  function normH(s){ return s.toLowerCase().trim().replace(/[‘’‚‛]/g,"'").replace(/[“”„‟]/g,'"'); }
  var lower = rawHeaders.map(function(h){ return normH(h.toString()); });
  var idx = {};
  for(var field in FIELD_MAP){
    idx[field]=-1;
    for(var v=0;v<FIELD_MAP[field].length;v++){
      var i=lower.indexOf(normH(FIELD_MAP[field][v])); if(i!==-1){idx[field]=i;break;}
    }
  }
  return idx;
}

function normalizeRow(row, rawHeaders, colIndex, sourceName) {
  var get=function(f){return(colIndex[f]>=0&&colIndex[f]<row.length)?row[colIndex[f]]:"";}; 
  var nom=get("nom").toString().trim(), prn=get("prenom").toString().trim();
  if(prn&&prn!==nom)nom=(nom+" "+prn).trim();
  if(!nom&&!get("email")&&!get("tel"))return null;
  return [get("date"),nom,normTel(get("tel")),get("email").toString().toLowerCase().trim(),
          get("entreprise"),get("ville"),get("pays"),get("adresse"),get("type"),sourceName,
          get("message").toString().trim(),get("quantity").toString().trim(),get("prodtype").toString().trim()];
}

function norm(v)    { return v?v.toString().trim().toLowerCase().replace(/\s+/g,""):""; }
function normTel(v) { return v?v.toString().replace(/[\s\-\.\(\)\+]/g,"").trim():""; }
function showAlert(msg) { SpreadsheetApp.getUi().alert(msg); }
function extractSheetId(url) {
  var m=url.match(/\/d\/([a-zA-Z0-9-_]+)/); if(!m)throw new Error("Invalid URL"); return m[1];
}

// ───────────────────────────────────────────────────────────────
// DATE FILTER DIALOG
// ───────────────────────────────────────────────────────────────
function openDateFilter() {
  var htmlContent = '<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_top"><style>*{box-sizing:border-box;margin:0;padding:0;font-family:Segoe UI,Arial,sans-serif}body{background:#fff;display:flex;flex-direction:column;height:100vh}.main{display:flex;flex:1;overflow:hidden}.cal-side{flex:1;padding:20px;overflow-y:auto;border-right:1px solid #eee}.sc-side{width:145px;background:#fafafa;border-left:1px solid #eee;padding:8px 0}.sc-title{padding:8px 16px;font-size:10px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.5px}.sc{padding:10px 16px;font-size:12px;font-weight:500;color:#444;cursor:pointer;transition:background .1s;white-space:nowrap}.sc:hover{background:#f0f0f0}.sc.on{background:#e8f5e9;color:#1E8449;font-weight:700}.hdr{font-size:15px;font-weight:800;color:#1E3A5F;margin-bottom:14px}.range-row{display:flex;gap:10px;margin-bottom:16px;align-items:flex-end}.rb{flex:1}.rl{font-size:10px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}.rv{background:#f5f5f5;border:2px solid #e0e0e0;border-radius:7px;padding:9px 12px;font-size:13px;font-weight:600;color:#333;min-height:38px}.rv.on{border-color:#1E8449;background:#e8f5e9;color:#1E8449}.sep{color:#ccc;font-size:20px;padding-bottom:8px}.cals{display:flex;gap:16px}.cal{flex:1;min-width:0}.cnav{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px}.cnav button{background:none;border:none;font-size:18px;cursor:pointer;color:#666;padding:2px 8px;border-radius:4px}.cnav button:hover{background:#f0f0f0}.ctitle{font-size:13px;font-weight:700;color:#333}.grid{display:grid;grid-template-columns:repeat(7,1fr);gap:1px}.dh{text-align:center;font-size:9px;font-weight:700;color:#bbb;padding:3px 0;text-transform:uppercase}.d{text-align:center;padding:6px 2px;font-size:12px;cursor:pointer;border-radius:50%;transition:all .1s;color:#333;user-select:none}.d:hover:not(.x):not(.om):not(.fut){background:#e8f5e9}.x{cursor:default;color:transparent}.om{cursor:default;color:#ddd}.fut{color:#ddd;cursor:not-allowed}.tod{font-weight:800;color:#1E3A5F}.s{background:#1E8449!important;color:#fff!important;border-radius:50% 0 0 50%!important}.e{background:#1E8449!important;color:#fff!important;border-radius:0 50% 50% 0!important}.se{border-radius:50%!important}.ir{background:#e8f5e9;border-radius:0;color:#1E8449;font-weight:600}.footer{padding:12px 20px;border-top:1px solid #eee;display:flex;justify-content:flex-end;gap:8px;background:#fff}.bc{padding:9px 20px;border:2px solid #ddd;border-radius:7px;background:#fff;font-size:13px;font-weight:600;cursor:pointer;color:#666}.bc:hover{background:#f5f5f5}.ba{padding:9px 24px;border:none;border-radius:7px;background:#1E8449;color:#fff;font-size:13px;font-weight:700;cursor:pointer}.ba:hover{opacity:.88}.st{padding:6px 20px;font-size:11px;font-weight:600;display:none;text-align:right}</style></head><body><div class="main"><div class="cal-side"><div class="hdr">Date Range</div><div class="range-row"><div class="rb"><div class="rl">FROM</div><div class="rv" id="df">-</div></div><div class="sep">&#8594;</div><div class="rb"><div class="rl">TO</div><div class="rv" id="dt">-</div></div></div><div class="cals"><div class="cal" id="c0"></div><div class="cal" id="c1"></div></div></div><div class="sc-side"><div class="sc-title">Quick Select</div><div class="sc" id="s0" onclick="sc(0)">Today</div><div class="sc" id="s1" onclick="sc(1)">Last 7 Days</div><div class="sc" id="s2" onclick="sc(2)">Last 30 Days</div><div class="sc on" id="s3" onclick="sc(3)">Month to date</div><div class="sc" id="s4" onclick="sc(4)">Last 12 months</div><div class="sc" id="s5" onclick="sc(5)">Year to date</div><div class="sc" id="s6" onclick="sc(6)">Last 3 years</div></div></div><div class="st" id="st"></div><div class="footer"><button class="bc" onclick="clearFilter()">Clear Filter</button><button class="bc" onclick="google.script.host.close()">Cancel</button><button class="ba" onclick="apply()">Apply</button></div><script>var MN=["January","February","March","April","May","June","July","August","September","October","November","December"];var MS=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];var TODAY=new Date();TODAY.setHours(12,0,0,0);var sF=null,sT=null,picking="from";var m0=new Date(TODAY.getFullYear(),TODAY.getMonth()-1,1);var m1=new Date(TODAY.getFullYear(),TODAY.getMonth(),1);sc(3);function sc(i){document.querySelectorAll(".sc").forEach(function(e){e.classList.remove("on");});document.getElementById("s"+i).classList.add("on");var t=new Date(TODAY),f;if(i===0){f=new Date(t);}else if(i===1){f=new Date(t);f.setDate(t.getDate()-6);}else if(i===2){f=new Date(t);f.setDate(t.getDate()-29);}else if(i===3){f=new Date(t.getFullYear(),t.getMonth(),1);}else if(i===4){f=new Date(t.getFullYear()-1,t.getMonth(),t.getDate());}else if(i===5){f=new Date(t.getFullYear(),0,1);}else if(i===6){f=new Date(t.getFullYear()-3,t.getMonth(),t.getDate());}sF=f;sT=new Date(t);picking="from";m0=new Date(sF.getFullYear(),sF.getMonth(),1);m1=new Date(sT.getFullYear(),sT.getMonth(),1);if(m0.getTime()===m1.getTime()){m0=new Date(m1.getFullYear(),m1.getMonth()-1,1);}render();}function prev(){m0=new Date(m0.getFullYear(),m0.getMonth()-1,1);m1=new Date(m1.getFullYear(),m1.getMonth()-1,1);render();}function next(){m0=new Date(m0.getFullYear(),m0.getMonth()+1,1);m1=new Date(m1.getFullYear(),m1.getMonth()+1,1);render();}function render(){buildCal(document.getElementById("c0"),m0,true);buildCal(document.getElementById("c1"),m1,false);updDisp();}function sd(y,mo,d){var date=new Date(y,mo,d,12);if(date>TODAY)return;if(picking==="from"||!sF||(sF&&sT)){sF=date;sT=null;picking="to";document.querySelectorAll(".sc").forEach(function(e){e.classList.remove("on");});}else{if(date<sF){sT=sF;sF=date;}else{sT=date;}picking="from";}render();}function sameD(a,b){return a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate();}function buildCal(el,md,showPrev){var y=md.getFullYear(),mo=md.getMonth();var fd=new Date(y,mo,1).getDay();var dim=new Date(y,mo+1,0).getDate();var h="<div class=\'cnav\'>";h+=showPrev?"<button onclick=\'prev()\'>&#8249;</button>":"<span></span>";h+="<span class=\'ctitle\'>"+MN[mo]+" "+y+"</span>";h+=!showPrev?"<button onclick=\'next()\'>&#8250;</button>":"<span></span>";h+="</div><div class=\'grid\'>";["S","M","T","W","T","F","S"].forEach(function(d){h+="<div class=\'dh\'>"+d+"</div>";});for(var i=0;i<fd;i++)h+="<div class=\'d x\'></div>";for(var d=1;d<=dim;d++){var dt=new Date(y,mo,d,12);var fut=dt>TODAY;var cls="d";if(fut){cls+=" fut";}else{var isS=sF&&sameD(dt,sF);var isE=sT&&sameD(dt,sT);var inR=sF&&sT&&dt>sF&&dt<sT;if(isS&&isE)cls+=" s e se";else if(isS)cls+=" s";else if(isE)cls+=" e";else if(inR)cls+=" ir";if(sameD(dt,TODAY))cls+=" tod";}var clk=fut?"":"onclick=\'sd("+y+","+mo+","+d+")\' ";h+="<div class=\'"+cls+"\' "+clk+">"+d+"</div>";}h+="</div>";el.innerHTML=h;}function fmtD(d){return d?d.toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}):"-";}function updDisp(){var df=document.getElementById("df"),dt=document.getElementById("dt");df.textContent=fmtD(sF);df.className="rv"+(sF?" on":"");dt.textContent=fmtD(sT||sF);dt.className="rv"+((sT||sF)?" on":"");}function iso(d){return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");}function clearFilter(){show("Clearing filter...","#2980B9");google.script.run.withSuccessHandler(function(m){show(m,"#1E8449");setTimeout(function(){google.script.host.close();},1000);}).withFailureHandler(function(e){show(e.message||"Error","#C0392B");}).applyDateFilter({type:"clear"});}function apply(){if(!sF){show("Please select a start date.","#C0392B");return;}show("Applying...","#2980B9");var params={type:"range",from:iso(sF),to:iso(sT||sF)};google.script.run.withSuccessHandler(function(m){show(m,"#1E8449");setTimeout(function(){google.script.host.close();},1200);}).withFailureHandler(function(e){show(e.message||"Error applying filter","#C0392B");}).applyDateFilter(params);}function show(m,c){var el=document.getElementById("st");el.textContent=m;el.style.color=c;el.style.display="block";}<\/script></body></html>';

  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(htmlContent).setWidth(660).setHeight(460).setTitle("Date Filter"),
    "Date Filter"
  );
}

function doPost(e) {
  try {
    var params=JSON.parse(e.postData.contents);
    var result=applyDateFilter(params);
    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.TEXT);
  } catch(err) {
    return ContentService.createTextOutput("Error: "+err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Lead Reporting")
    // ── Section client : actions quotidiennes ──
    .addItem("🔄 Refresh Dashboard",            "refreshDashboard")
    .addItem("📅 Date Filter",                  "openDateFilter")
    .addSeparator()
    // ── Section admin / technique ──
    .addItem("⚙️ Full Consolidation",           "consolidateLeads")
    .addItem("⚙️ Setup triggers (8h sync + GA4 on edit)", "setNightlyTrigger")
    .addItem("🔍 Test GA4 Access",              "testGA4Access")
    .addItem("🔑 Authorize / Re-authorize script", "requestAuthorization")
    .addToUi();

  // Silently check if authorization is needed and show a prompt
  try {
    var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
      showAuthorizationPrompt(authInfo.getAuthorizationUrl());
    }
  } catch(e) {
    // Ignore — onOpen may have limited auth scope
  }
}

// Called from menu item to manually trigger re-authorization
function requestAuthorization() {
  try {
    var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    var status = authInfo.getAuthorizationStatus();
    if (status === ScriptApp.AuthorizationStatus.REQUIRED ||
        status === ScriptApp.AuthorizationStatus.NOT_REQUIRED) {
      // Force a small operation that requires full auth scope
      SpreadsheetApp.getActiveSpreadsheet().getName();
      SpreadsheetApp.getUi().alert(
        "✅ Authorization OK",
        "The script is already authorized. You can use all features including Date Filter.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert(
      "⚠️ Authorization needed",
      "Please click OK then accept all permissions in the next window.\n\nThis is required to use Date Filter and other features.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showAuthorizationPrompt(url) {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;padding:20px;">' +
    '<h3 style="color:#1E3A5F;margin-top:0">🔑 Authorization Required</h3>' +
    '<p>This script needs your permission to run features like <strong>Date Filter</strong>.</p>' +
    '<p>Click the button below and accept all permissions:</p>' +
    '<a href="' + url + '" target="_blank" style="display:inline-block;padding:10px 20px;' +
    'background:#1E8449;color:#fff;border-radius:6px;text-decoration:none;font-weight:bold;">' +
    '→ Authorize Script</a>' +
    '<p style="font-size:11px;color:#999;margin-top:16px">After authorizing, close this panel and reload the spreadsheet.</p>' +
    '</div>'
  ).setWidth(380).setHeight(200).setTitle("Authorization Required");
  SpreadsheetApp.getUi().showModelessDialog(html, "Authorization Required");
}

// ─────────────────────────────────────────────────────────
// TEST GA4 -- Run from menu to diagnose access issues
// ─────────────────────────────────────────────────────────
function testGA4Access() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName("Config");
  if (!cfg) { showAlert("Config not found"); return; }
  var lastRow = Math.max(cfg.getLastRow()-1, 1);
  var cfgData = cfg.getRange(2, 1, lastRow, 4).getValues();
  var results = [], found = 0;

  for (var i=0; i<cfgData.length; i++) {
    var name = cfgData[i][0] ? cfgData[i][0].toString().trim() : "";
    var pid  = cfgData[i][3] ? cfgData[i][3].toString().trim() : "";
    if (!name || !pid) continue;
    found++;
    try {
      var url = "https://analyticsdata.googleapis.com/v1beta/properties/"+pid+":runReport";
      var resp = UrlFetchApp.fetch(url, {
        method:"post", contentType:"application/json",
        headers:{Authorization:"Bearer "+ScriptApp.getOAuthToken()},
        payload:JSON.stringify({
          dateRanges:[{startDate:"2026-01-01",endDate:"2026-03-03"}],
          dimensions:[{name:"yearMonth"}],
          metrics:[{name:"sessions"},{name:"totalUsers"}],
          limit:3
        }),
        muteHttpExceptions:true
      });
      var data = JSON.parse(resp.getContentText());
      if (data.error) {
        results.push("❌ "+name+" (ID:"+pid+")\n   -> "+data.error.code+": "+data.error.message);
      } else if (data.rows && data.rows.length > 0) {
        var s=data.rows[0].metricValues[0].value, u=data.rows[0].metricValues[1].value;
        results.push("✅ "+name+": "+s+" sessions, "+u+" users (Jan-Mar 2026)");
      } else {
        results.push("⚠️ "+name+" (ID:"+pid+"): API OK but no data returned\n   -> Check the property ID or enable the GA4 Data API in Services");
      }
    } catch(e) {
      results.push("❌ "+name+": "+e.message);
    }
  }

  if (found === 0) {
    showAlert("No GA4 Property IDs found in column D of Config sheet.\nAdd your Property IDs in column D (numbers only, e.g. 526737266)");
    return;
  }
  showAlert("GA4 Diagnostic Results ("+found+" properties):\n\n"+results.join("\n\n")+"\n\n─────\nIf errors: Apps Script > Services > Add 'Google Analytics Data API'");
}

function setNightlyTrigger() {
  // Delete existing time-based and onEdit installable triggers
  ScriptApp.getProjectTriggers().forEach(function(t){
    var type = t.getEventType();
    if (type === ScriptApp.EventType.CLOCK || type === ScriptApp.EventType.ON_EDIT) {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Auto-sync every 8 hours
  ScriptApp.newTrigger("consolidateLeads").timeBased().everyHours(8).create();
  // Installable onEdit trigger for full OAuth (GA4 on period change)
  ScriptApp.newTrigger("onDashboardEdit").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  showAlert("✅ Auto-sync enabled every 8 hours\n✅ Installable trigger set for GA4 on period change\n\nNote: Google quotas allow ~20k UrlFetch calls/day -- 3 syncs/day is well within limits.");
}
