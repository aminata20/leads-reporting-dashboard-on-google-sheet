// ═══════════════════════════════════════════════════════════════
//  REPORTING LEADS — Final Clean Version
// ═══════════════════════════════════════════════════════════════

var FIELD_MAP = {
  date:      ["date","created at","date de creation","date de creation","date de creations"],
  nom:       ["nom","name","first name"],
  prenom:    ["last name","prenom","prénom"],
  tel:       ["telephone","téléphone","phone","phone/whatsapp","numero de telephone (gsm)","whatsapp"],
  email:     ["email","e-mail","adresse e-mail","adresse email"],
  entreprise:["entreprise","company name","company_name","nom de l'entreprise"],
  ville:     ["ville","city"],
  pays:      ["country","pays"],
  adresse:   ["adresse","address","address line 1"],
  type:      ["business type","type d'activite","type","type d activite"]
};

var STANDARD_HEADERS = ["Date","Full Name","Phone","Email","Company","City","Country","Address","Type","Source"];
var MONTHS_EN = ["January","February","March","April","May","June","July","August","September","October","November","December"];
var MONTHS_SH = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

var DROP_ROW = 11;
var DROP_COL = 3;
var TABLE_START_ROW = 14;

// ───────────────────────────────────────────────────────────────
function consolidateLeads() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName("Config");
  if (!cfg) { showAlert("Config sheet not found!"); return; }

  var cfgData = cfg.getRange(2,1,Math.max(cfg.getLastRow()-1,1),2).getValues();

  // allSourceNames : noms uniques dans l'ordre d'apparition (pour les colonnes)
  var allSourceNames = [];
  var seenNames = {};
  cfgData.forEach(function(r) {
    var name = r[0] ? r[0].toString().trim() : "";
    if (name && !seenNames[name]) {
      allSourceNames.push(name);
      seenNames[name] = true;
    }
  });

  // sources : toutes les lignes avec URL valide (y compris doublons de nom)
  var sources = cfgData.filter(function(r){
    return r[0] && r[1] && r[1].toString().includes("docs.google.com");
  });

  var result = fetchAndProcess(sources);
  var uniqueLeads = result.uniqueLeads;
  var allLeads    = result.allLeads;
  var dupLeads    = result.dupLeads;
  var invalidCount = result.invalidCount;

  writeSheet(ss, "Data",       STANDARD_HEADERS, uniqueLeads.map(function(r){ return [fmtDate(r[0])].concat(r.slice(1)); }), "#1E8449");
  writeSheet(ss, "Raw",        STANDARD_HEADERS, allLeads.map(function(r){    return [fmtDate(r[0])].concat(r.slice(1)); }), "#1E3A5F");
  writeSheet(ss, "Duplicates", STANDARD_HEADERS, dupLeads.map(function(r){   return [fmtDate(r[0])].concat(r.slice(1)); }), "#C0392B");

  storeCache(ss, uniqueLeads, allSourceNames, dupLeads.length);
  buildDashboard(ss, uniqueLeads, allLeads, dupLeads, allSourceNames);

  cfg.getRange("D1").setValue("Last sync:").setFontWeight("bold");
  cfg.getRange("E1").setValue(new Date()).setNumberFormat("dd/mm/yyyy hh:mm");

  showAlert("Done!\n\nTotal leads: " + allLeads.length + "\nUnique: " + uniqueLeads.length + "\nDuplicates: " + dupLeads.length + "\nInvalid dates: " + invalidCount);
}

// ───────────────────────────────────────────────────────────────
function refreshDashboard() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName("Config");
  if (!cfg) { showAlert("Config sheet not found!"); return; }

  var cfgData = cfg.getRange(2,1,Math.max(cfg.getLastRow()-1,1),2).getValues();

  // allSourceNames : noms uniques dans l'ordre d'apparition
  var allSourceNames = [];
  var seenNames = {};
  cfgData.forEach(function(r) {
    var name = r[0] ? r[0].toString().trim() : "";
    if (name && !seenNames[name]) {
      allSourceNames.push(name);
      seenNames[name] = true;
    }
  });

  // sources : toutes les lignes avec URL valide
  var sources = cfgData.filter(function(r){
    return r[0] && r[1] && r[1].toString().includes("docs.google.com");
  });

  ss.toast("Fetching data from all sources...", "Refreshing", -1);

  var result = fetchAndProcess(sources);
  var uniqueLeads = result.uniqueLeads;
  var allLeads    = result.allLeads;
  var dupLeads    = result.dupLeads;

  writeSheet(ss, "Data",       STANDARD_HEADERS, uniqueLeads.map(function(r){ return [fmtDate(r[0])].concat(r.slice(1)); }), "#1E8449");
  writeSheet(ss, "Raw",        STANDARD_HEADERS, allLeads.map(function(r){    return [fmtDate(r[0])].concat(r.slice(1)); }), "#1E3A5F");
  writeSheet(ss, "Duplicates", STANDARD_HEADERS, dupLeads.map(function(r){   return [fmtDate(r[0])].concat(r.slice(1)); }), "#C0392B");

  storeCache(ss, uniqueLeads, allSourceNames, dupLeads.length);

  var dash = ss.getSheetByName("Dashboard");
  if (dash) {
    var selected = dash.getRange(DROP_ROW, DROP_COL).getValue() || "Monthly";
    updateKPIs(dash, uniqueLeads, dupLeads.length);
    drawDynamicTable(dash, uniqueLeads, allSourceNames, selected, null);
  }

  cfg.getRange("E1").setValue(new Date()).setNumberFormat("dd/mm/yyyy hh:mm");
  ss.toast("Done! " + uniqueLeads.length + " unique leads | " + dupLeads.length + " duplicates", "Refresh complete", 5);
}

// ───────────────────────────────────────────────────────────────
function fetchAndProcess(sources) {
  var todayEnd = new Date(); todayEnd.setHours(23,59,59,999);
  var minDate  = new Date(2020,0,1);
  var allLeads = [];
  var invalidCount = 0;

  for (var si = 0; si < sources.length; si++) {
    var sourceName = sources[si][0];
    var url        = sources[si][1];
    try {
      var id    = extractSheetId(url);
      var ss    = SpreadsheetApp.openById(id);
      
      // ── Détecter le bon onglet depuis l'URL (gid) ──────────
      var sheet = null;
      var gidMatch = url.toString().match(/[#&?]gid=(\d+)/);
      if (gidMatch) {
        var gid = parseInt(gidMatch[1]);
        var allSheets = ss.getSheets();
        for (var shi = 0; shi < allSheets.length; shi++) {
          if (allSheets[shi].getSheetId() === gid) {
            sheet = allSheets[shi];
            break;
          }
        }
        if (!sheet) {
          Logger.log("WARNING: gid=" + gid + " not found in " + sourceName + ", using first sheet");
          sheet = ss.getSheets()[0];
        }
      } else {
        // Pas de gid dans l'URL → premier onglet
        sheet = ss.getSheets()[0];
      }

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow < 2) { Logger.log("Empty: " + sourceName + " (sheet: " + sheet.getName() + ")"); continue; }

      Logger.log(sourceName + " → sheet: '" + sheet.getName() + "' (" + (lastRow-1) + " rows)");

      var values     = sheet.getRange(1,1,lastRow,lastCol).getValues();
      var rawHeaders = values[0].map(function(h){ return h.toString().trim(); });
      var colIndex   = buildColIndex(rawHeaders);
      var dateFormat = detectDateFormat(values, colIndex);
      Logger.log(sourceName + " → date format: " + dateFormat);

      var imported = 0;

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
        if (!d || d < minDate || d > todayEnd) {
          if (!d) Logger.log("Invalid date: " + normalized[0] + " in " + sourceName);
          invalidCount++;
          continue;
        }
        normalized[0] = d;
        allLeads.push(normalized);
        imported++;
      }
      Logger.log(sourceName + ": " + imported + " leads imported");

    } catch(e) {
      Logger.log("ERROR " + sourceName + ": " + e.message);
    }
  }

  allLeads.sort(function(a,b){ return a[0] - b[0]; });

  var seenEmail = {}, seenTel = {}, seenNom = {};
  var uniqueLeads = [], dupLeads = [];

  for (var li = 0; li < allLeads.length; li++) {
    var lead  = allLeads[li];
    var email = norm(lead[3]);
    var tel   = normTel(lead[2]);
    var nom   = norm(lead[1]);
    var isDup = false;

    if (email)      { isDup = !!seenEmail[email]; seenEmail[email] = true; }
    else if (tel)   { isDup = !!seenTel[tel];     seenTel[tel]     = true; }
    else if (nom)   { isDup = !!seenNom[nom];     seenNom[nom]     = true; }

    if (isDup) dupLeads.push(lead); else uniqueLeads.push(lead);
  }

  return { uniqueLeads: uniqueLeads, allLeads: allLeads, dupLeads: dupLeads, invalidCount: invalidCount };
}

// ───────────────────────────────────────────────────────────────
//  CACHE — stocke les dates en format ISO string pur
// ───────────────────────────────────────────────────────────────
var CACHE_SHEET_NAME = "LeadCache";  // nom simple sans emojis

function storeCache(ss, uniqueLeads, allSourceNames, dupsCount) {
  // Supprimer les anciens caches peu importe leur nom
  ["Cache", "⚙️ Cache", "⚙️ Data", "LeadCache"].forEach(function(n) {
    var old = ss.getSheetByName(n);
    if (old) ss.deleteSheet(old);
  });

  var ds = ss.insertSheet(CACHE_SHEET_NAME);
  ds.hideSheet();

  ds.getRange(1,1).setValue(JSON.stringify({
    sources: allSourceNames,
    dups:    dupsCount,
    total:   uniqueLeads.length,
    updated: new Date().toISOString()
  }));

  if (uniqueLeads.length > 0) {
    var rows = uniqueLeads.map(function(r) {
      var d = (r[0] instanceof Date) ? r[0] : toDateObj(r[0]);
      var dateStr = "";
      if (d) {
        var yyyy = d.getFullYear();
        var mm   = String(d.getMonth()+1).padStart(2,"0");
        var dd   = String(d.getDate()).padStart(2,"0");
        dateStr  = yyyy + "-" + mm + "-" + dd;
      }
      return [dateStr, (r[9] || "Unknown").toString()];
    });
    ds.getRange(2, 1, rows.length, 2).setValues(rows);
    Logger.log("Cache stored: " + rows.length + " leads");
  }
}

function loadCache(ss) {
  var ds = ss.getSheetByName(CACHE_SHEET_NAME);
  if (!ds) {
    // Chercher n'importe quel ancien cache
    var fallbacks = ["Cache", "⚙️ Cache", "⚙️ Data"];
    for (var fi=0; fi<fallbacks.length; fi++) {
      ds = ss.getSheetByName(fallbacks[fi]);
      if (ds) { Logger.log("Using fallback cache: " + fallbacks[fi]); break; }
    }
  }
  if (!ds) { Logger.log("ERROR: No cache sheet found!"); return null; }

  try {
    var metaStr = ds.getRange(1,1).getValue();
    if (!metaStr) { Logger.log("ERROR: Cache meta is empty"); return null; }
    
    var meta = JSON.parse(metaStr);
    var last  = ds.getLastRow();
    
    Logger.log("Cache sheet: '" + ds.getName() + "' rows: " + last);
    
    if (last < 2) {
      Logger.log("WARNING: Cache has no data rows");
      return { meta: meta, leads: [] };
    }

    var rawRows = ds.getRange(2, 1, last-1, 2).getValues();
    var leads   = [];

    for (var i=0; i<rawRows.length; i++) {
      var raw    = rawRows[i][0];
      var source = rawRows[i][1] ? rawRows[i][1].toString() : "Unknown";
      
      if (!raw) continue;
      
      var dateStr = raw.toString().trim();
      var d = null;

      // Format YYYY-MM-DD stocké par storeCache
      var iso = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (iso) {
        d = new Date(parseInt(iso[1]), parseInt(iso[2])-1, parseInt(iso[3]), 12, 0, 0);
      }
      
      // Format Date objet (si Google Sheets a converti la string en date)
      if (!d && raw instanceof Date && !isNaN(raw)) {
        d = new Date(raw.getFullYear(), raw.getMonth(), raw.getDate(), 12, 0, 0);
      }
      
      // Fallback toDateObj
      if (!d) d = toDateObj(raw);
      
      if (!d || isNaN(d)) {
        Logger.log("Skip invalid date in cache: '" + raw + "'");
        continue;
      }

      var row = new Array(10).fill("");
      row[0]  = d;
      row[9]  = source;
      leads.push(row);
    }

    Logger.log("Cache loaded: " + leads.length + " leads from " + (last-1) + " rows");
    return { meta: meta, leads: leads };

  } catch(e) {
    Logger.log("Cache load ERROR: " + e.message);
    return null;
  }
}
// ───────────────────────────────────────────────────────────────
function buildDashboard(ss, unique, all, dups, allSourceNames) {
  var dash = ss.getSheetByName("Dashboard");
  if (!dash) dash = ss.insertSheet("Dashboard");
  dash.clearContents();
  dash.clearFormats();
  try { dash.getDataRange().clearDataValidations(); } catch(e) {}
  dash.setHiddenGridlines(true);

  var today  = new Date();
  var yr     = today.getFullYear();
  var nbSrc  = allSourceNames.length;
  var nbCols = nbSrc + 2;

  dash.setColumnWidth(1, 14);
  dash.setColumnWidth(2, 185);
  for (var c = 3; c <= 2+nbSrc; c++) dash.setColumnWidth(c, 115);
  dash.setColumnWidth(2+nbSrc+1, 85);

  // Title
  dash.setRowHeight(1,8); dash.setRowHeight(2,58); dash.setRowHeight(3,20); dash.setRowHeight(4,4);
  setR(dash,2,2,"Lead Reporting Dashboard",{merge:[1,nbCols],fontSize:22,bold:true,fontColor:"#1E3A5F",vAlign:"middle"});
  setR(dash,3,2,"Updated: "+today.toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long",year:"numeric"})+"   Deduplication: Email > Phone > Name",
    {merge:[1,nbCols],fontSize:9,italic:true,fontColor:"#AAAAAA"});
  dash.getRange(4,2,1,nbCols).setBackground("#1E3A5F");

  // KPIs
  dash.setRowHeight(5,8); dash.setRowHeight(6,50); dash.setRowHeight(7,22); dash.setRowHeight(8,8);
  var kpiDefs = [
    {label:"TOTAL LEADS", col:2, color:"#1E3A5F", bg:"#EBF5FB"},
    {label:"TODAY",       col:3, color:"#E67E22", bg:"#FEF9E7"},
    {label:"THIS WEEK",   col:4, color:"#8E44AD", bg:"#F5EEF8"},
    {label:"THIS MONTH",  col:5, color:"#2980B9", bg:"#EBF5FB"},
    {label:"YEAR "+yr,    col:6, color:"#1E8449", bg:"#EAFAF1"},
    {label:"DUPLICATES",  col:7, color:"#C0392B", bg:"#FDEDEC"}
  ];
  var kpiVals = [
    unique.length,
    countPeriod(unique,"day"),
    countPeriod(unique,"week"),
    countPeriod(unique,"month"),
    countPeriod(unique,"year"),
    dups.length
  ];
  for (var ki = 0; ki < kpiDefs.length; ki++) {
    var k = kpiDefs[ki];
    dash.getRange(6,k.col).setValue(kpiVals[ki])
        .setFontSize(30).setFontWeight("bold").setFontColor(k.color)
        .setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground(k.bg);
    dash.getRange(7,k.col).setValue(k.label)
        .setFontSize(7).setFontColor("#999999").setFontWeight("bold")
        .setHorizontalAlignment("center").setBackground(k.bg);
  }

  // Period selector
  dash.setRowHeight(9,8); dash.setRowHeight(DROP_ROW,34); dash.setRowHeight(DROP_ROW-1,8);
  setR(dash,DROP_ROW,2,"Period:",{bold:true,fontSize:10,fontColor:"#1E3A5F",vAlign:"middle",hAlign:"right"});
  dash.getRange(DROP_ROW, DROP_COL)
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(["Daily","Weekly","Monthly","Yearly"],true)
        .setAllowInvalid(false).build())
      .setValue("Monthly")
      .setBackground("#1E3A5F").setFontColor("#FFFFFF").setFontWeight("bold")
      .setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");
  setR(dash,DROP_ROW,DROP_COL+1,"Select a period  |  Use menu: Date Filter for custom range",
    {merge:[1,Math.max(nbSrc,1)],fontSize:9,italic:true,fontColor:"#BBBBBB",vAlign:"middle"});

  dash.setRowHeight(13,4);
  dash.getRange(13,2,1,nbCols).setBackground("#E8E8E8");

  drawDynamicTable(dash, unique, allSourceNames, "Monthly", null);
}

// ───────────────────────────────────────────────────────────────
function updateKPIs(dash, unique, dupsCount) {
  if (!dash) return;
  var today = new Date();
  var vals = [unique.length, countPeriod(unique,"day"), countPeriod(unique,"week"), countPeriod(unique,"month"), countPeriod(unique,"year"), dupsCount];
  var cols = [2,3,4,5,6,7];
  for (var i=0; i<cols.length; i++) dash.getRange(6,cols[i]).setValue(vals[i]);
  setR(dash,3,2,"Updated: "+today.toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long",year:"numeric"})+"   Deduplication: Email > Phone > Name",
    {merge:[1,10],fontSize:9,italic:true,fontColor:"#AAAAAA"});
}

// ───────────────────────────────────────────────────────────────
function drawDynamicTable(dash, leads, allSourceNames, periodLabel, customFilter) {
  var nbSrc  = allSourceNames.length;
  var nbCols = nbSrc + 2;

  var zone = dash.getRange(TABLE_START_ROW, 2, 400, nbCols);
  zone.clearContent(); zone.clearFormat(); zone.breakApart();

  var period    = labelToKey(periodLabel);
  var hasFilter = customFilter && customFilter.type && customFilter.type !== "clear";
  var grouped   = hasFilter ? filterCustom(leads, customFilter) : groupBySource(leads, period);
  var periods   = grouped.periods;
  var crossData = grouped.crossData;

  var r = TABLE_START_ROW;

  // Titre
  var titleText = hasFilter ? buildFilterLabel(customFilter) : ("Leads by Source — " + periodLabel);
  dash.setRowHeight(r, 32);
  setR(dash,r,2,titleText,{merge:[1,nbCols],bg:"#1E3A5F",fontColor:"#FFFFFF",bold:true,fontSize:11,hAlign:"left",vAlign:"middle"});
  r++;

  // Header
  dash.setRowHeight(r, 28);
  dash.getRange(r,2).setValue("Period")
      .setBackground("#D6EAF8").setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setFontColor("#1E3A5F");
  for (var si=0; si<allSourceNames.length; si++) {
    dash.getRange(r,3+si).setValue(allSourceNames[si])
        .setBackground("#D6EAF8").setFontWeight("bold").setFontSize(9)
        .setHorizontalAlignment("center").setFontColor("#1E3A5F");
  }
  dash.getRange(r,2+nbSrc+1).setValue("TOTAL")
      .setBackground("#1E3A5F").setFontColor("#FFFFFF")
      .setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
  r++;

  // Lignes de données
  var dataStartRow = r;

  if (periods.length === 0) {
    dash.setRowHeight(r,34);
    setR(dash,r,2,"No data for this selection.",{merge:[1,nbCols],fontColor:"#AAAAAA",italic:true,hAlign:"center",bg:"#FAFAFA",fontSize:10});
    r++;
  } else {
    var values    = [];
    var bgColors  = [];
    var fntColors = [];

    for (var pi=0; pi<periods.length; pi++) {
      var key     = periods[pi];
      var rowData = crossData[key] || {};
      var bg      = pi % 2 === 0 ? "#FFFFFF" : "#F8F9FA";
      var total   = 0;
      var rowVals = [key];
      var rowBgs  = [bg];
      var rowFCs  = ["#1E3A5F"];

      for (var sj=0; sj<allSourceNames.length; sj++) {
        var count = rowData[allSourceNames[sj]] || 0;
        total += count;
        rowVals.push(count > 0 ? count : "-");
        rowBgs.push(bg);
        rowFCs.push(count > 0 ? "#2980B9" : "#CCCCCC");
      }
      rowVals.push(total);
      rowBgs.push(total > 0 ? "#EAFAF1" : bg);
      rowFCs.push(total > 0 ? "#1E8449" : "#CCCCCC");
      values.push(rowVals);
      bgColors.push(rowBgs);
      fntColors.push(rowFCs);
    }

    var dataRange = dash.getRange(r, 2, periods.length, nbCols);
    dataRange.setValues(values);
    dataRange.setBackgrounds(bgColors);
    dataRange.setFontColors(fntColors);
    dataRange.setFontSize(9);
    dataRange.setHorizontalAlignment("center");
    dash.getRange(r, 2, periods.length, 1)
        .setHorizontalAlignment("left").setFontWeight("bold");

    for (var ri=0; ri<periods.length; ri++) dash.setRowHeight(r+ri, 24);
    r += periods.length;
  }

  // Total général
  dash.setRowHeight(r, 28);
  var totVals = ["TOTAL"];
  var grand   = 0;
  for (var tk=0; tk<allSourceNames.length; tk++) {
    var t = 0;
    for (var pk=0; pk<periods.length; pk++) t += ((crossData[periods[pk]]||{})[allSourceNames[tk]]||0);
    grand += t;
    totVals.push(t);
  }
  totVals.push(grand);
  dash.getRange(r,2,1,nbCols)
      .setValues([totVals])
      .setBackground("#1E3A5F").setFontColor("#FFFFFF")
      .setFontWeight("bold").setHorizontalAlignment("center").setFontSize(9);
  dash.getRange(r,2).setHorizontalAlignment("left");
  dash.getRange(r,2+nbSrc+1).setBackground("#1E8449");

  // ── BORDURE COMPLÈTE sur toute la plage du tableau ──────────
  // La plage totale : de la ligne header jusqu'au total général
  var headerRow  = TABLE_START_ROW + 1; // ligne header (Period, sources, TOTAL)
  var totalRows  = r - headerRow + 1;   // header + données + total général

  // Bordure extérieure épaisse
  dash.getRange(headerRow, 2, totalRows, nbCols)
      .setBorder(true, true, true, true, null, null,
                 "#999999", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Bordures internes entre chaque cellule (toutes les lignes et colonnes)
  dash.getRange(headerRow, 2, totalRows, nbCols)
      .setBorder(null, null, null, null, true, true,
                 "#999999", SpreadsheetApp.BorderStyle.SOLID);
}

// ───────────────────────────────────────────────────────────────
// ── 1. groupBySource — limite Daily 90j, Weekly 6 mois, tri correct ──
function groupBySource(leads, period) {
  var crossData = {};
  var todayEnd  = new Date(); todayEnd.setHours(23,59,59,999);

  // Calculer la date limite selon la période
  var limitDate = null;
  if (period === "day") {
    limitDate = new Date();
    limitDate.setDate(limitDate.getDate() - 89); // 90 jours
    limitDate.setHours(0,0,0,0);
  }
  if (period === "week") {
    limitDate = new Date();
    limitDate.setMonth(limitDate.getMonth() - 6); // 6 mois
    limitDate.setHours(0,0,0,0);
  }
  if (period === "month") {
    limitDate = new Date();
    limitDate.setMonth(limitDate.getMonth() - 23); // 24 mois
    limitDate.setDate(1);
    limitDate.setHours(0,0,0,0);
  }

  for (var i=0; i<leads.length; i++) {
    var d = toDateObj(leads[i][0]);
    if (!d || d > todayEnd || d.getFullYear() < 2020) continue;
    if (limitDate && d < limitDate) continue;
    var key    = dateToKey(d, period);
    var source = (leads[i][9] || "Unknown").toString().trim();
    if (!crossData[key]) crossData[key] = {};
    crossData[key][source] = (crossData[key][source] || 0) + 1;
  }

  // Tri : toujours du plus récent au plus ancien
  var periods = Object.keys(crossData).sort(function(a,b){
    return keyToTs(b, period) - keyToTs(a, period);
  });
  return { periods: periods, crossData: crossData };
}

// ───────────────────────────────────────────────────────────────
function filterCustom(leads, f) {
  var crossData = {};
  var todayEnd  = new Date(); todayEnd.setHours(23,59,59,999);
  var testFn;

  if (f.type === "range") {
    var fromParts = f.from.split("-"); var toParts = f.to.split("-");
    var fromD = new Date(+fromParts[0],+fromParts[1]-1,+fromParts[2],0,0,0);
    var toD   = new Date(+toParts[0],  +toParts[1]-1,  +toParts[2],  23,59,59);
    testFn = function(d){ return d >= fromD && d <= toD; };
  } else {
    testFn = function(){ return false; };
  }

  for (var i=0; i<leads.length; i++) {
    var d = toDateObj(leads[i][0]);
    if (!d || d > todayEnd || d.getFullYear() < 2020 || !testFn(d)) continue;
    var key    = d.getDate() + " " + MONTHS_SH[d.getMonth()] + " " + d.getFullYear();
    var source = (leads[i][9] || "Unknown").toString().trim();
    if (!crossData[key]) crossData[key] = {};
    crossData[key][source] = (crossData[key][source] || 0) + 1;
  }

  var periods = Object.keys(crossData).sort(function(a,b){
    return keyToTs(a,"day") - keyToTs(b,"day");
  });
  return { periods: periods, crossData: crossData };
}

// ── 2. buildFilterLabel — affiche une seule date si from === to ──
function buildFilterLabel(f) {
  if (f.type === "range") {
    var fp = f.from.split("-"); var tp = f.to.split("-");
    var fromLabel = +fp[2] + " " + MONTHS_SH[+fp[1]-1] + " " + fp[0];
    var toLabel   = +tp[2] + " " + MONTHS_SH[+tp[1]-1] + " " + tp[0];
    // Si même jour → afficher une seule date
    if (f.from === f.to) {
      return "Leads by Source — " + fromLabel;
    }
    return "Leads by Source — " + fromLabel + " to " + toLabel;
  }
  return "Leads by Source — Custom Filter";
}

// ───────────────────────────────────────────────────────────────
function dateToKey(d, period) {
  if (period === "day")   return d.getDate() + " " + MONTHS_SH[d.getMonth()] + " " + d.getFullYear();
  if (period === "week") {
    var mon = getMondayOf(d);
    var sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate()+6);
    if (mon.getMonth() === sun.getMonth()) {
      return MONTHS_EN[mon.getMonth()] + " " + mon.getDate() + " - " + sun.getDate() + ", " + sun.getFullYear();
    }
    return MONTHS_SH[mon.getMonth()] + " " + mon.getDate() + " - " + MONTHS_SH[sun.getMonth()] + " " + sun.getDate() + ", " + sun.getFullYear();
  }
  if (period === "month") return MONTHS_EN[d.getMonth()] + " " + d.getFullYear();
  if (period === "year")  return d.getFullYear().toString();
  return "";
}

function getMondayOf(d) {
  var day  = d.getDay();
  var diff = (day === 0) ? -6 : 1 - day;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + diff, 12);
}

// ── 3. keyToTs — correction tri Weekly ────────────────────────
function keyToTs(key, period) {
  try {
    if (period === "year") {
      return new Date(parseInt(key), 0, 1).getTime();
    }

    if (period === "month") {
      var p = key.split(" ");
      return new Date(parseInt(p[1]), MONTHS_EN.indexOf(p[0]), 1).getTime();
    }

    if (period === "day") {
      var p2 = key.split(" ");
      return new Date(parseInt(p2[2]), MONTHS_SH.indexOf(p2[1]), parseInt(p2[0]), 12).getTime();
    }

    if (period === "week") {
      // Extraire l'année en fin de string
      var yearMatch = key.match(/(\d{4})$/);
      if (!yearMatch) return 0;
      var yr = parseInt(yearMatch[1]);

      // Format 1 — même mois : "February 9 - 15, 2026"
      var sameMonth = key.match(/^([A-Za-z]+)\s+(\d+)\s*-\s*\d+,\s*\d{4}$/);
      if (sameMonth) {
        var mIdx = MONTHS_EN.indexOf(sameMonth[1]);
        if (mIdx === -1) mIdx = MONTHS_SH.indexOf(sameMonth[1]);
        return new Date(yr, mIdx, parseInt(sameMonth[2]), 12).getTime();
      }

      // Format 2 — mois différents : "Jan 27 - Feb 2, 2026" ou "Dec 29 - Jan 4, 2026"
      var diffMonth = key.match(/^([A-Za-z]+)\s+(\d+)\s*-\s*([A-Za-z]+)\s+(\d+),\s*(\d{4})$/);
      if (diffMonth) {
        var startMonthName = diffMonth[1];
        var startDay       = parseInt(diffMonth[2]);
        var endMonthName   = diffMonth[3];
        var endYear        = parseInt(diffMonth[5]);

        var startMIdx = MONTHS_EN.indexOf(startMonthName);
        if (startMIdx === -1) startMIdx = MONTHS_SH.indexOf(startMonthName);
        var endMIdx = MONTHS_EN.indexOf(endMonthName);
        if (endMIdx === -1) endMIdx = MONTHS_SH.indexOf(endMonthName);

        // Cas "Dec 29 - Jan 4, 2026" :
        // Le mois de début (Dec) > mois de fin (Jan) → la semaine commence l'année PRÉCÉDENTE
        var startYear = endYear;
        if (startMIdx > endMIdx) {
          startYear = endYear - 1;
        }

        return new Date(startYear, startMIdx, startDay, 12).getTime();
      }
    }
  } catch(e) {
    Logger.log("keyToTs error: " + e.message + " key=" + key);
  }
  return 0;
}
// ───────────────────────────────────────────────────────────────
function countPeriod(leads, period) {
  var now   = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 12);

  var count = 0;
  for (var i=0; i<leads.length; i++) {
    var d = toDateObj(leads[i][0]);
    if (!d) continue;
    var ld = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12);
    if (period === "day"   && ld.getTime() === today.getTime()) count++;
    else if (period === "week") {
      var mon = getMondayOf(today);
      var sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate()+6, 23);
      if (ld >= mon && ld <= sun) count++;
    }
    else if (period === "month" && ld.getMonth()===today.getMonth() && ld.getFullYear()===today.getFullYear()) count++;
    else if (period === "year"  && ld.getFullYear()===today.getFullYear()) count++;
  }
  return count;
}

// ───────────────────────────────────────────────────────────────
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "Dashboard") return;
  if (e.range.getRow() !== DROP_ROW || e.range.getColumn() !== DROP_COL) return;

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var cache = loadCache(ss);
  if (!cache) return;
  drawDynamicTable(sheet, cache.leads, cache.meta.sources, e.range.getValue(), null);
}

// ───────────────────────────────────────────────────────────────
function applyDateFilter(params) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var cache = loadCache(ss);
  if (!cache) throw new Error("No data found. Please run a full consolidation first.");

  var dash     = ss.getSheetByName("Dashboard");
  var selected = dash.getRange(DROP_ROW, DROP_COL).getValue() || "Monthly";

  if (params.type === "clear") {
    drawDynamicTable(dash, cache.leads, cache.meta.sources, selected, null);
    return "Filter cleared.";
  }

  drawDynamicTable(dash, cache.leads, cache.meta.sources, selected, params);
  return "Filter applied successfully!";
}

// ───────────────────────────────────────────────────────────────
function openDateFilter() {
  var htmlContent = '<!DOCTYPE html>' +
  '<html><head><meta charset="utf-8"><base target="_top">' +
  '<style>' +
  '*{box-sizing:border-box;margin:0;padding:0;font-family:Segoe UI,Arial,sans-serif}' +
  'body{background:#fff;display:flex;flex-direction:column;height:100vh}' +
  '.main{display:flex;flex:1;overflow:hidden}' +
  '.cal-side{flex:1;padding:20px;overflow-y:auto;border-right:1px solid #eee}' +
  '.sc-side{width:145px;background:#fafafa;border-left:1px solid #eee;padding:8px 0}' +
  '.sc-title{padding:8px 16px;font-size:10px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.5px}' +
  '.sc{padding:10px 16px;font-size:12px;font-weight:500;color:#444;cursor:pointer;transition:background .1s;white-space:nowrap}' +
  '.sc:hover{background:#f0f0f0}' +
  '.sc.on{background:#e8f5e9;color:#1E8449;font-weight:700}' +
  '.hdr{font-size:15px;font-weight:800;color:#1E3A5F;margin-bottom:14px}' +
  '.range-row{display:flex;gap:10px;margin-bottom:16px;align-items:flex-end}' +
  '.rb{flex:1}' +
  '.rl{font-size:10px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}' +
  '.rv{background:#f5f5f5;border:2px solid #e0e0e0;border-radius:7px;padding:9px 12px;font-size:13px;font-weight:600;color:#333;min-height:38px}' +
  '.rv.on{border-color:#1E8449;background:#e8f5e9;color:#1E8449}' +
  '.sep{color:#ccc;font-size:20px;padding-bottom:8px}' +
  '.cals{display:flex;gap:16px}' +
  '.cal{flex:1;min-width:0}' +
  '.cnav{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px}' +
  '.cnav button{background:none;border:none;font-size:18px;cursor:pointer;color:#666;padding:2px 8px;border-radius:4px}' +
  '.cnav button:hover{background:#f0f0f0}' +
  '.ctitle{font-size:13px;font-weight:700;color:#333}' +
  '.grid{display:grid;grid-template-columns:repeat(7,1fr);gap:1px}' +
  '.dh{text-align:center;font-size:9px;font-weight:700;color:#bbb;padding:3px 0;text-transform:uppercase}' +
  '.d{text-align:center;padding:6px 2px;font-size:12px;cursor:pointer;border-radius:50%;transition:all .1s;color:#333;user-select:none}' +
  '.d:hover:not(.x):not(.om):not(.fut){background:#e8f5e9}' +
  '.x{cursor:default;color:transparent}' +
  '.om{cursor:default;color:#ddd}' +
  '.fut{color:#ddd;cursor:not-allowed}' +
  '.tod{font-weight:800;color:#1E3A5F}' +
  '.s{background:#1E8449!important;color:#fff!important;border-radius:50% 0 0 50%!important}' +
  '.e{background:#1E8449!important;color:#fff!important;border-radius:0 50% 50% 0!important}' +
  '.se{border-radius:50%!important}' +
  '.ir{background:#e8f5e9;border-radius:0;color:#1E8449;font-weight:600}' +
  '.footer{padding:12px 20px;border-top:1px solid #eee;display:flex;justify-content:flex-end;gap:8px;background:#fff}' +
  '.bc{padding:9px 20px;border:2px solid #ddd;border-radius:7px;background:#fff;font-size:13px;font-weight:600;cursor:pointer;color:#666}' +
  '.bc:hover{background:#f5f5f5}' +
  '.ba{padding:9px 24px;border:none;border-radius:7px;background:#1E8449;color:#fff;font-size:13px;font-weight:700;cursor:pointer}' +
  '.ba:hover{opacity:.88}' +
  '.st{padding:6px 20px;font-size:11px;font-weight:600;display:none;text-align:right}' +
  '</style></head>' +
  '<body>' +
  '<div class="main">' +
  '<div class="cal-side">' +
  '<div class="hdr">Date Range</div>' +
  '<div class="range-row">' +
  '<div class="rb"><div class="rl">FROM</div><div class="rv" id="df">-</div></div>' +
  '<div class="sep">&#8594;</div>' +
  '<div class="rb"><div class="rl">TO</div><div class="rv" id="dt">-</div></div>' +
  '</div>' +
  '<div class="cals"><div class="cal" id="c0"></div><div class="cal" id="c1"></div></div>' +
  '</div>' +
  '<div class="sc-side">' +
  '<div class="sc-title">Quick Select</div>' +
  '<div class="sc" id="s0" onclick="sc(0)">Today</div>' +
  '<div class="sc" id="s1" onclick="sc(1)">Last 7 Days</div>' +
  '<div class="sc" id="s2" onclick="sc(2)">Last 30 Days</div>' +
  '<div class="sc on" id="s3" onclick="sc(3)">Month to date</div>' +
  '<div class="sc" id="s4" onclick="sc(4)">Last 12 months</div>' +
  '<div class="sc" id="s5" onclick="sc(5)">Year to date</div>' +
  '<div class="sc" id="s6" onclick="sc(6)">Last 3 years</div>' +
  '</div>' +
  '</div>' +
  '<div class="st" id="st"></div>' +
  '<div class="footer">' +
  '<button class="bc" onclick="google.script.host.close()">Cancel</button>' +
  '<button class="ba" onclick="apply()">Apply</button>' +
  '</div>' +
  '<script>' +
  'var MN=["January","February","March","April","May","June","July","August","September","October","November","December"];' +
  'var MS=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];' +
  'var TODAY=new Date();TODAY.setHours(12,0,0,0);' +
  'var sF=null,sT=null,picking="from";' +
  'var m0=new Date(TODAY.getFullYear(),TODAY.getMonth()-1,1);' +
  'var m1=new Date(TODAY.getFullYear(),TODAY.getMonth(),1);' +
  'sc(3);' +
  'function sc(i){' +
  'document.querySelectorAll(".sc").forEach(function(e){e.classList.remove("on");});' +
  'document.getElementById("s"+i).classList.add("on");' +
  'var t=new Date(TODAY),f;' +
  'if(i===0){f=new Date(t);}' +
  'else if(i===1){f=new Date(t);f.setDate(t.getDate()-6);}' +
  'else if(i===2){f=new Date(t);f.setDate(t.getDate()-29);}' +
  'else if(i===3){f=new Date(t.getFullYear(),t.getMonth(),1);}' +
  'else if(i===4){f=new Date(t.getFullYear()-1,t.getMonth(),t.getDate());}' +
  'else if(i===5){f=new Date(t.getFullYear(),0,1);}' +
  'else if(i===6){f=new Date(t.getFullYear()-3,t.getMonth(),t.getDate());}' +
  'sF=f;sT=new Date(t);picking="from";' +
  'm0=new Date(sF.getFullYear(),sF.getMonth(),1);' +
  'm1=new Date(sT.getFullYear(),sT.getMonth(),1);' +
  'if(m0.getTime()===m1.getTime()){m0=new Date(m1.getFullYear(),m1.getMonth()-1,1);}' +
  'render();' +
  '}' +
  'function prev(){m0=new Date(m0.getFullYear(),m0.getMonth()-1,1);m1=new Date(m1.getFullYear(),m1.getMonth()-1,1);render();}' +
  'function next(){m0=new Date(m0.getFullYear(),m0.getMonth()+1,1);m1=new Date(m1.getFullYear(),m1.getMonth()+1,1);render();}' +
  'function render(){buildCal(document.getElementById("c0"),m0,true);buildCal(document.getElementById("c1"),m1,false);updDisp();}' +
  'function sd(y,mo,d){' +
  'var date=new Date(y,mo,d,12);' +
  'if(date>TODAY)return;' +
  'if(picking==="from"||!sF||(sF&&sT)){sF=date;sT=null;picking="to";document.querySelectorAll(".sc").forEach(function(e){e.classList.remove("on");});}' +
  'else{if(date<sF){sT=sF;sF=date;}else{sT=date;}picking="from";}' +
  'render();' +
  '}' +
  'function sameD(a,b){return a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate();}' +
  'function buildCal(el,md,showPrev){' +
  'var y=md.getFullYear(),mo=md.getMonth();' +
  'var fd=new Date(y,mo,1).getDay();' +
  'var dim=new Date(y,mo+1,0).getDate();' +
  'var h="<div class=\'cnav\'>";' +
  'h+=showPrev?"<button onclick=\'prev()\'>&#8249;</button>":"<span></span>";' +
  'h+="<span class=\'ctitle\'>"+MN[mo]+" "+y+"</span>";' +
  'h+=!showPrev?"<button onclick=\'next()\'>&#8250;</button>":"<span></span>";' +
  'h+="</div><div class=\'grid\'>";' +
  '["S","M","T","W","T","F","S"].forEach(function(d){h+="<div class=\'dh\'>"+d+"</div>";});' +
  'for(var i=0;i<fd;i++)h+="<div class=\'d x\'></div>";' +
  'for(var d=1;d<=dim;d++){' +
  'var dt=new Date(y,mo,d,12);' +
  'var fut=dt>TODAY;' +
  'var cls="d";' +
  'if(fut){cls+=" fut";}' +
  'else{' +
  'var isS=sF&&sameD(dt,sF);' +
  'var isE=sT&&sameD(dt,sT);' +
  'var inR=sF&&sT&&dt>sF&&dt<sT;' +
  'if(isS&&isE)cls+=" s e se";' +
  'else if(isS)cls+=" s";' +
  'else if(isE)cls+=" e";' +
  'else if(inR)cls+=" ir";' +
  'if(sameD(dt,TODAY))cls+=" tod";' +
  '}' +
  'var clk=fut?"":"onclick=\'sd("+y+","+mo+","+d+")\' ";' +
  'h+="<div class=\'"+cls+"\' "+clk+">"+d+"</div>";' +
  '}' +
  'h+="</div>";' +
  'el.innerHTML=h;' +
  '}' +
  'function fmtD(d){return d?d.toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}):"-";}' +
  'function updDisp(){' +
  'var df=document.getElementById("df"),dt=document.getElementById("dt");' +
  'df.textContent=fmtD(sF);df.className="rv"+(sF?" on":"");' +
  'dt.textContent=fmtD(sT||sF);dt.className="rv"+((sT||sF)?" on":"");' +
  '}' +
  'function iso(d){return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");}' +
  'function apply(){' +
  'if(!sF){show("Please select a start date.","#C0392B");return;}' +
  'show("Applying...","#2980B9");' +
  'google.script.run' +
  '.withSuccessHandler(function(m){show(m,"#1E8449");setTimeout(function(){google.script.host.close();},1200);})' +
  '.withFailureHandler(function(e){show(e.message,"#C0392B");})' +
  '.applyDateFilter({type:"range",from:iso(sF),to:iso(sT||sF)});' +
  '}' +
  'function show(m,c){var el=document.getElementById("st");el.textContent=m;el.style.color=c;el.style.display="block";}' +
  '<\/script>' +
  '</body></html>';

  var html = HtmlService.createHtmlOutput(htmlContent).setWidth(660).setHeight(460).setTitle("Date Filter");
  SpreadsheetApp.getUi().showModelessDialog(html, "Date Filter");
}

// ───────────────────────────────────────────────────────────────
// ── Détecte le format dominant d'un fichier source ─────────────
function detectDateFormat(values, colIndex) {
  // Analyse jusqu'à 20 lignes pour détecter le format dominant
  var ddmmCount = 0;
  var mmddCount = 0;
  var limit = Math.min(values.length, 21);

  for (var i = 1; i < limit; i++) {
    var raw = colIndex.date >= 0 ? values[i][colIndex.date] : null;
    if (!raw) continue;
    if (raw instanceof Date) continue; // objet Date natif = pas ambiguïté

    var str = raw.toString().trim();
    var sep = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
    if (!sep) continue;

    var a = +sep[1], b = +sep[2];
    if (a > 12 && b <= 12) ddmmCount += 3; // certain DD/MM
    if (b > 12 && a <= 12) mmddCount += 3; // certain MM/DD
    // Ambiguïté : on regarde si a > b (probable DD/MM car jours > mois en fréquence)
    if (a <= 12 && b <= 12 && a > b) ddmmCount++;
    if (a <= 12 && b <= 12 && b > a) mmddCount++;
  }

  // Par défaut DD/MM (format FR/international le plus courant)
  return mmddCount > ddmmCount * 2 ? "MM/DD/YYYY" : "DD/MM/YYYY";
}

// ── Remplace toDateObj ──────────────────────────────────────────
// Accepte un paramètre optionnel forceFormat: "DD/MM/YYYY" ou "MM/DD/YYYY"
function toDateObj(val, forceFormat) {
  if (!val) return null;
  if (val instanceof Date) {
    if (isNaN(val)) return null;
    return new Date(val.getFullYear(), val.getMonth(), val.getDate(), 12);
  }

  var str = val.toString().trim();
  if (!str || str === "Invalid Date") return null;

  // ISO avec T : 2026-09-01T23:00:00.000Z
  if (/^\d{4}-\d{2}-\d{2}T/.test(str)) {
    var d = new Date(str);
    if (isNaN(d)) return null;
    return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12);
  }

  // ISO date seule : 2026-09-01
  var iso = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return new Date(+iso[1], +iso[2]-1, +iso[3], 12);

  // Format avec séparateur : 9/2/2026 ou 24/02/2026
  var sep = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (sep) {
    var a = +sep[1], b = +sep[2], y = +sep[3];

    if (a > 12) return new Date(y, b-1, a, 12); // a = jour certain
    if (b > 12) return new Date(y, a-1, b, 12); // b = jour certain

    // Ambiguïté : utiliser le format détecté sur le fichier source
    if (forceFormat === "MM/DD/YYYY") return new Date(y, a-1, b, 12);
    return new Date(y, b-1, a, 12); // défaut DD/MM/YYYY
  }

  // Texte FR : "24 décembre 2025"
  var MOIS = {"janvier":0,"février":1,"fevrier":1,"mars":2,"avril":3,"mai":4,"juin":5,
              "juillet":6,"août":7,"aout":7,"septembre":8,"octobre":9,"novembre":10,
              "décembre":11,"decembre":11};
  var frT = str.match(/^(\d{1,2})(?:er|ème|e)?\s+([a-zéûàôùêîèä]+)\s+(\d{4})$/i);
  if (frT) {
    var mi = MOIS[frT[2].toLowerCase()];
    if (mi !== undefined) return new Date(+frT[3], mi, +frT[1], 12);
  }

  // Texte EN : "February 24, 2026" ou "24 Feb 2026"
  var MONS = {"january":0,"february":1,"march":2,"april":3,"may":4,"june":5,"july":6,
              "august":7,"september":8,"october":9,"november":10,"december":11,
              "jan":0,"feb":1,"mar":2,"apr":3,"jun":5,"jul":6,"aug":7,
              "sep":8,"oct":9,"nov":10,"dec":11};
  var enA = str.match(/^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})$/);
  if (enA) { var ma = MONS[enA[1].toLowerCase()]; if (ma !== undefined) return new Date(+enA[3], ma, +enA[2], 12); }
  var enB = str.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (enB) { var mb = MONS[enB[2].toLowerCase()]; if (mb !== undefined) return new Date(+enB[3], mb, +enB[1], 12); }

  // Serial GSheets
  if (/^\d{5}$/.test(str)) {
    var ds = new Date((+str - 25569) * 86400000);
    if (!isNaN(ds)) return new Date(ds.getFullYear(), ds.getMonth(), ds.getDate(), 12);
  }

  var fb = new Date(str);
  if (!isNaN(fb)) return new Date(fb.getFullYear(), fb.getMonth(), fb.getDate(), 12);
  Logger.log("Unparseable date: " + str);
  return null;
}

// ───────────────────────────────────────────────────────────────
function fmtDate(val) {
  var d = toDateObj(val);
  if (!d) return val ? val.toString() : "";
  return String(d.getDate()).padStart(2,"0") + " " + MONTHS_SH[d.getMonth()] + " " + d.getFullYear();
}

function setR(sheet, row, col, value, opts) {
  opts = opts || {};
  var r = opts.merge ? sheet.getRange(row,col,opts.merge[0],opts.merge[1]) : sheet.getRange(row,col);
  if (opts.merge && (opts.merge[0]>1||opts.merge[1]>1)) r.merge();
  if (value !== undefined) r.setValue(value);
  if (opts.bg)        r.setBackground(opts.bg);
  if (opts.fontColor) r.setFontColor(opts.fontColor);
  if (opts.fontSize)  r.setFontSize(opts.fontSize);
  if (opts.bold)      r.setFontWeight("bold");
  if (opts.italic)    r.setFontStyle("italic");
  if (opts.hAlign)    r.setHorizontalAlignment(opts.hAlign);
  if (opts.vAlign)    r.setVerticalAlignment(opts.vAlign);
}

function labelToKey(l) {
  var map = {"Daily":"day","Weekly":"week","Monthly":"month","Yearly":"year"};
  return map[l] || "month";
}

function writeSheet(ss, name, headers, rows, color) {
  var s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  s.clearContents(); s.clearFormats();
  if (s.getFilter()) s.getFilter().remove();
  if (!rows.length) { s.getRange(1,1).setValue("No data"); return; }
  var data = [headers].concat(rows);
  s.getRange(1,1,data.length,headers.length).setValues(data);
  s.getRange(1,1,1,headers.length).setBackground(color).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center");
  for (var i=2; i<=rows.length+1; i++) s.getRange(i,1,1,headers.length).setBackground(i%2===0?"#F2F3F4":"#FFFFFF");
  s.setFrozenRows(1);
  s.autoResizeColumns(1,headers.length);
  s.getRange(1,1,data.length,headers.length).createFilter();
}

function buildColIndex(rawHeaders) {
  var lower = rawHeaders.map(function(h){ return h.toLowerCase().trim(); });
  var idx = {};
  for (var field in FIELD_MAP) {
    idx[field] = -1;
    var variants = FIELD_MAP[field];
    for (var v=0; v<variants.length; v++) {
      var i = lower.indexOf(variants[v]);
      if (i !== -1) { idx[field] = i; break; }
    }
  }
  return idx;
}

function normalizeRow(row, rawHeaders, colIndex, sourceName) {
  var get = function(f){ return (colIndex[f]>=0 && colIndex[f]<row.length) ? row[colIndex[f]] : ""; };
  var nom = get("nom").toString().trim();
  var prn = get("prenom").toString().trim();
  if (prn && prn !== nom) nom = (nom + " " + prn).trim();
  if (!nom && !get("email") && !get("tel")) return null;
  return [get("date"), nom, normTel(get("tel")), get("email").toString().toLowerCase().trim(),
          get("entreprise"), get("ville"), get("pays"), get("adresse"), get("type"), sourceName];
}

function norm(v)    { return v ? v.toString().trim().toLowerCase().replace(/\s+/g,"") : ""; }
function normTel(v) { return v ? v.toString().replace(/[\s\-\.\(\)\+]/g,"").trim() : ""; }
function showAlert(msg) { SpreadsheetApp.getUi().alert(msg); }
function extractSheetId(url) {
  var m = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!m) throw new Error("Invalid URL");
  return m[1];
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Lead Reporting")
    .addItem("Full Consolidation",       "consolidateLeads")
    .addItem("Refresh Dashboard",        "refreshDashboard")
    .addItem("Date Filter",              "openDateFilter")
    .addSeparator()
    .addItem("Enable nightly auto-sync", "setNightlyTrigger")
    .addToUi();
}

function setNightlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("consolidateLeads").timeBased().everyDays(1).atHour(0).nearMinute(0).create();
  showAlert("Auto-sync enabled every night at midnight!");
}

