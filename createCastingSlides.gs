/***** CONFIG *****/
const SPREADSHEET_ID    = "1lhTKUfqWwpi43GXqUH2rzRmXxT4ir07XRAO6N0koxPc";
const SHEET_NAME        = "SHOW SS26";
const SLIDE_ID          = "1b3C4H5awx6C9TvcBRrwMayDZayejdxTpO49Id-8xBNA";
const TEMPLATE_SLIDE_ID = "1ZWBjUTHYjb_crQhevLQWRmxHKuk7K0dXq3p_5VfYjcU";
const DIVIDER_MARKER    = "__DIVIDER__";
const TITLE_MARKER      = "__TITLE__";
const BATCH_SIZE        = 20;

/***** FIELD STYLES *****/
const FIELD_STYLES = {
  NAME:        { font: "Helvetica", size: 32, bold: true,  align: "LEFT",   color: "#000000" },
  NOTE:        { font: "Helvetica", size: 12, bold: false, align: "LEFT",   color: "#000000" },
  NOTES:       { font: "Helvetica", size: 12, bold: false, align: "RIGHT",  color: "#000000" },
  "MORE NOTES":{ font: "Helvetica", size: 12, bold: false, align: "RIGHT",  color: "#000000" },
  INSTAGRAM:   { font: "Helvetica", size: 14, bold: false, align: "LEFT",   color: "#000000" },
  AGENCY:      { font: "Helvetica", size: 22, bold: false, align: "LEFT",   color: "#000000" },
  BOOKING:     { font: "Helvetica", size: 22, bold: false, align: "LEFT",   color: "#000000" },
  FEE:         { font: "Helvetica", size: 22, bold: true,  align: "RIGHT",  color: "#000000" },
  STATUS:      { font: "Helvetica", size: 22, bold: true,  align: "RIGHT" },
  REC:         { font: "Helvetica", size: 24, bold: true,  align: "CENTER", color: "#FFD700" },
  ID:          { font: "Helvetica", size: 1,  bold: false, align: "RIGHT",  color: "#FFFFFF" },
};

/***** STATE (for batching + sidebar) *****/
const UP = PropertiesService.getUserProperties();
const ROWS_KEY = "SYNC_ROWS";
const IDX_KEY = "SYNC_IDX";
const LAST_SYNC_KEY = "LAST_SYNCED";
const MODE_KEY = "CURRENT_MODE"; // sync or cleanup

let SYNC_STATE = { state: "idle", pct: 0, msg: "READY" };
let CLEANUP_STATE = { state: "idle", pct: 0, msg: "READY", duplicates: [] };

/***** SAFE TEXT ACCESS *****/
function safeGetText(shape) {
  try {
    if (!shape || typeof shape.getText !== "function") return null;
    return shape.getText();
  } catch (e) { return null; }
}
function safeAsString(tr) {
  try { return tr.asString(); } catch (e) { return ""; }
}

/***** STATE HELPERS *****/
function getSyncState() { return SYNC_STATE; }
function setSyncState(newState) { SYNC_STATE = { ...SYNC_STATE, ...newState }; }
function getCleanupState() { return CLEANUP_STATE; }
function setCleanupState(newState) { CLEANUP_STATE = { ...CLEANUP_STATE, ...newState }; }

/***** SIDEBAR *****/
function openSidebar(mode) {
  const template = HtmlService.createTemplateFromFile("SyncSidebar");
  template.mode = mode; // pass to HTML
  const html = template.evaluate().setTitle("Casting Sync");
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSidebarStatus() {
  return JSON.stringify({
    sync: getSyncState(),
    cleanup: getCleanupState(),
    lastSynced: UP.getProperty(LAST_SYNC_KEY) || null,
    mode: UP.getProperty(MODE_KEY) || null
  });
}

/***** MENU *****/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Casting Sync");

  menu.addItem("Sync Slides", "menuStartSync");
  menu.addItem("Clean Up + Flag Duplicates", "menuCleanup");

  menu.addToUi();
}

function menuStartSync() {
  if (getCleanupState().state === "running") {
    SpreadsheetApp.getUi().alert("❌ Cleanup is currently running. Please wait until it finishes.");
    return;
  }
  UP.setProperty(MODE_KEY, "sync");
  openSidebar("sync");
  startSync();
}

function menuCleanup() {
  if (getSyncState().state === "running") {
    SpreadsheetApp.getUi().alert("❌ Sync is currently running. Please wait until it finishes.");
    return;
  }
  UP.setProperty(MODE_KEY, "cleanup");
  openSidebar("cleanup");
  cleanupAndFlagDuplicates();
}

/***** CANCEL *****/
function cancelSync() {
  setSyncState({ state: "idle", pct: 0, msg: "❌ SYNC CANCELLED" });
  setCleanupState({ state: "idle", pct: 0, msg: "❌ CLEANUP CANCELLED", duplicates: [] });
  clearBatchTriggers();
  resetProgress();
}

/***** SYNC CONTROL *****/
function startSync() {
  clearBatchTriggers();
  resetProgress();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  ensureIDColumn_(sheet);

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idCol = headers.indexOf("ID");
  if (idCol === -1) {
    setSyncState({ state: "error", msg: "❌ NO 'ID' COLUMN FOUND" });
    return;
  }

  const rowList = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = (row[idCol] || "").toString().trim();
    if (!id) continue;
    const hasData = row.some((cell, j) => j !== idCol && cell && String(cell).trim() !== "");
    if (!hasData) continue;
    rowList.push(i + 1);
  }

  UP.setProperty(ROWS_KEY, JSON.stringify(rowList));
  UP.setProperty(IDX_KEY, "0");

  setSyncState({ state: "running", pct: 0, msg: `SYNC STARTED – ${rowList.length} ROWS` });
  processNextBatch();
}

function processNextBatch() {
  const s = getSyncState();
  if (s.state !== "running") return;

  clearBatchTriggers();
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const rowList = JSON.parse(UP.getProperty(ROWS_KEY) || "[]");
    let idx = parseInt(UP.getProperty(IDX_KEY) || "0", 10);
    const total = rowList.length;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
    const dateIndex = headers.indexOf("DATE");

    const deck = SlidesApp.openById(SLIDE_ID);
    const templateDeck = SlidesApp.openById(TEMPLATE_SLIDE_ID);
    const templateSlide = templateDeck.getSlides()[0];

    const start = idx;
    const end = Math.min(idx + BATCH_SIZE, total);
    for (let p = start; p < end; p++) {
      const rNum = rowList[p];
      const rowVals = sheet.getRange(rNum, 1, 1, sheet.getLastColumn()).getValues()[0];
      const rowData = {};
      headers.forEach((h, j) => rowData[h] = rowVals[j]);

      const id = (rowData["ID"] || "").toString().trim();
      if (!id) continue;

      const hasData = headers.some(h => h !== "ID" && rowData[h] && String(rowData[h]).trim() !== "");
      if (!hasData) continue;

      let slide = findOrCreateSlide(deck, id);
      rebuildTextFromTemplate(slide, templateSlide, rowData, id, dateIndex, rowData["BOOKING"]);
    }

    idx = end;
    UP.setProperty(IDX_KEY, String(idx));
    const pct = total === 0 ? 100 : Math.min(100, Math.round((idx / total) * 100));
    setSyncState({ pct, msg: `PROCESSING… ${pct}%` });

    if (idx < total) {
      ScriptApp.newTrigger("processNextBatch").timeBased().after(600).create();
      return;
    }

    finalizeDeck_(deck, sheet, headers, dateIndex, templateSlide);

    setSyncState({ state: "idle", pct: 100, msg: "✅ SLIDES SYNCED" });
    resetProgress();
    UP.setProperty(LAST_SYNC_KEY, new Date().toLocaleString());

  } catch (err) {
    setSyncState({ state: "error", msg: `❌ ERROR: ${String(err && err.message || err)}` });
  }
}

/***** CLEANUP ENGINE *****/
function cleanupAndFlagDuplicates() {
  if (getSyncState().state !== "idle") {
    setCleanupState({ state: "idle", pct: 0, msg: "❌ CLEANUP DISABLED WHILE SYNC RUNNING", duplicates: [] });
    return;
  }
  setCleanupState({ state: "running", pct: 0, msg: "RUNNING CLEANUP…", duplicates: [] });
  processCleanupBatch();
}

function processCleanupBatch() {
  try {
    if (getCleanupState().state !== "running") return;
    Utilities.sleep(200);
    setCleanupState({ pct: getCleanupState().pct + 25 });

    if (getCleanupState().pct < 100) {
      ScriptApp.newTrigger("processCleanupBatch").timeBased().after(400).create();
    } else {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = ss.getSheetByName(SHEET_NAME);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = headers.indexOf("NAME");
      const idCol = headers.indexOf("ID");

      let rowsToDelete = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const id = row[idCol] ? String(row[idCol]).trim() : "";
        const hasContent = row.some((cell, j) => j !== idCol && cell && String(cell).trim() !== "");
        if (id && !hasContent) rowsToDelete.push(i + 1);
      }
      rowsToDelete.reverse().forEach(r => sheet.deleteRow(r));

      let lastRow = sheet.getLastRow();
      for (let r = lastRow; r > 1; r--) {
        const values = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (values.join("").trim() === "") sheet.deleteRow(r);
        else break;
      }

      let seen = {}, duplicates = [];
      for (let i = 1; i < data.length; i++) {
        const name = (data[i][nameCol] || "").toString().trim().toUpperCase();
        if (!name) continue;
        if (seen[name]) duplicates.push(name);
        else seen[name] = true;
      }

      setCleanupState({ state: "idle", pct: 100, msg: "✅ CLEANUP COMPLETE", duplicates });
    }
  } catch (err) {
    setCleanupState({ state: "error", pct: 0, msg: `❌ ERROR: ${err.message}`, duplicates: [] });
  }
}

/***** HELPERS *****/
function ensureIDColumn_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  let idCol = headers.indexOf("ID");
  if (idCol === -1) {
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1).setValue("ID");
    idCol = 0;
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[idCol];
    const hasContent = row.some((cell, j) => j !== idCol && cell && String(cell).trim() !== "");
    if (!id && hasContent) sheet.getRange(i + 1, idCol + 1).setValue(generateUUID_());
  }
  if (!sheet.isColumnHiddenByUser(idCol + 1)) sheet.hideColumns(idCol + 1);
}
function generateUUID_() {
  const chars = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx';
  return chars.replace(/[xy]/g, c => {
    const r = Math.random() * 16 | 0;
    const v = (c === 'x') ? r : ((r & 0x3) | 0x8);
    return v.toString(16);
  });
}
function clearBatchTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction && 
      (t.getHandlerFunction() === "processNextBatch" || t.getHandlerFunction() === "processCleanupBatch"))
    .forEach(t => ScriptApp.deleteTrigger(t));
}
function resetProgress() {
  UP.deleteProperty(ROWS_KEY);
  UP.deleteProperty(IDX_KEY);
}

/***** SLIDES – FIND/CREATE & FINALIZE *****/
function findOrCreateSlide(deck, uuid) {
  for (const s of deck.getSlides()) {
    const sid = extractSlideID_(s);
    if (sid === uuid) return s;
  }
  const slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const marker = slide.insertTextBox(uuid, 5, 5, 50, 10);
  const mtr = safeGetText(marker);
  if (mtr) mtr.getTextStyle().setFontSize(1).setForegroundColor("#FFFFFF");
  return slide;
}
function finalizeDeck_(deck, sheet, headers, dateIndex, templateSlide) {
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const rowData = {};
    headers.forEach((h, j) => rowData[h] = data[i][j]);
    const id = (rowData["ID"] || "").toString().trim();
    if (!id) continue;

    const hasData = headers.some(h => h !== "ID" && rowData[h] && String(rowData[h]).trim() !== "");
    if (!hasData) continue;

    const booking = rowData["BOOKING"] ? String(rowData["BOOKING"]).toUpperCase() : "";
    const agencyCell = rowData["AGENCY"] ? String(rowData["AGENCY"]).toUpperCase() : "";

    let agencyVal = "OTHER";
    if (booking === "DB") agencyVal = "DIRECT BOOKING";
    else if (booking === "AGENCY") agencyVal = agencyCell || "AGENCY";
    else if (!booking && agencyCell) agencyVal = agencyCell;

    const status = rowData["STATUS"] ? String(rowData["STATUS"]).toUpperCase() : "";
    const isNA = (status === "NA" || status === "NOT AVAILABLE");

    rows.push({
      raw: rowData,
      id: id,
      name: rowData["NAME"] ? String(rowData["NAME"]).toUpperCase() : "",
      firstName: rowData["NAME"] ? String(rowData["NAME"]).split(" ")[0].toUpperCase() : "",
      agency: agencyVal,
      booking: booking,
      isNA: isNA
    });
  }

  const buckets = {};
  deck.getSlides().forEach(s => {
    const sid = extractSlideID_(s);
    if (!sid) return;
    if (!buckets[sid]) buckets[sid] = [];
    buckets[sid].push(s);
  });
  const slideMap = {};
  Object.keys(buckets).forEach(id => {
    const list = buckets[id];
    if (list.length === 1) { slideMap[id] = list[0]; return; }
    let best = list[0];
    let bestCount = countImages_(best);
    for (let i = 1; i < list.length; i++) {
      const c = countImages_(list[i]);
      if (c > bestCount) { best = list[i]; bestCount = c; }
      else if (c === bestCount) { best = list[i]; }
    }
    slideMap[id] = best;
    list.forEach(s => { if (s !== best) s.remove(); });
  });

  deck.getSlides().forEach(s => { if (isDividerSlide_(s) || isTitleSlide_(s)) s.remove(); });

  const available = rows.filter(r => !r.isNA)
    .sort((a, b) => (a.agency.localeCompare(b.agency)) || a.name.localeCompare(b.name));
  const notAvailable = rows.filter(r => r.isNA)
    .sort((a, b) => a.name.localeCompare(b.name));

  const grouped = {};
  available.forEach(r => {
    if (!grouped[r.agency]) grouped[r.agency] = [];
    grouped[r.agency].push(r);
  });

  const order = [];
  if (grouped["DIRECT BOOKING"]) order.push("DIRECT BOOKING");
  const agencies = Object.keys(grouped).filter(k => k !== "DIRECT BOOKING" && k !== "OTHER").sort();
  order.push(...agencies);
  if (grouped["OTHER"]) order.push("OTHER");

  insertTitleSlide_(deck, SHEET_NAME, 0);
  let pos = 1;
  const usedIds = new Set();

  order.forEach(agency => {
    insertDividerSlide_(deck, agency, grouped[agency], pos); pos++;
    grouped[agency].forEach(r => {
      let slide = slideMap[r.id] || findOrCreateSlide(deck, r.id);
      slide.move(pos);
      rebuildTextFromTemplate(slide, templateSlide, r.raw, r.id, dateIndex, r.booking);
      usedIds.add(r.id);
      pos++;
    });
  });

  if (notAvailable.length > 0) {
    insertDividerSlide_(deck, "NOT AVAILABLE", notAvailable, pos); pos++;
    notAvailable.forEach(r => {
      let slide = slideMap[r.id] || findOrCreateSlide(deck, r.id);
      slide.move(pos);
      rebuildTextFromTemplate(slide, templateSlide, r.raw, r.id, dateIndex, r.booking);
      usedIds.add(r.id);
      pos++;
    });
  }

  Object.keys(slideMap).forEach(id => { if (!usedIds.has(id)) slideMap[id].remove(); });
}
function countImages_(slide) {
  try {
    const els = slide.getPageElements();
    let n = 0;
    els.forEach(pe => { if (pe.getPageElementType && pe.getPageElementType() === SlidesApp.PageElementType.IMAGE) n++; });
    return n;
  } catch (e) { return 0; }
}
function extractSlideID_(slide) {
  for (const sh of slide.getShapes()) {
    const tr = safeGetText(sh);
    if (tr) {
      const t = safeAsString(tr).trim();
      if (/^[a-f0-9-]{36}$/i.test(t)) return t;
    }
  }
  return null;
}
function isDividerSlide_(slide) {
  return slide.getShapes().some(sh => {
    const tr = safeGetText(sh);
    return tr && safeAsString(tr).includes(DIVIDER_MARKER);
  });
}
function isTitleSlide_(slide) {
  return slide.getShapes().some(sh => {
    const tr = safeGetText(sh);
    return tr && safeAsString(tr).includes(TITLE_MARKER);
  });
}
function insertTitleSlide_(deck, title, position) {
  const slide = deck.insertSlide(position, SlidesApp.PredefinedLayout.BLANK);
  const pageWidth  = deck.getPageWidth();
  const pageHeight = deck.getPageHeight();
  const boxWidth = pageWidth * 0.8, boxHeight = 100;
  const left = (pageWidth - boxWidth) / 2, top = (pageHeight - boxHeight) / 2;
  const titleShape = slide.insertTextBox((title || "").toString().toUpperCase(), left, top, boxWidth, boxHeight);
  const tr = safeGetText(titleShape);
  if (tr) {
    tr.getTextStyle().setFontFamily("Helvetica").setFontSize(48).setBold(true).setForegroundColor("#000000");
    tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }
  const marker = slide.insertTextBox(TITLE_MARKER, 5, 5, 50, 10);
  const mtr = safeGetText(marker);
  if (mtr) mtr.getTextStyle().setFontSize(1).setForegroundColor("#FFFFFF");
}
function insertDividerSlide_(deck, agency, models, position) {
  const slide = deck.insertSlide(position, SlidesApp.PredefinedLayout.BLANK);
  const pageWidth = deck.getPageWidth();
  const titleShape = slide.insertTextBox(agency || " ", 50, 100, pageWidth - 100, 100);
  const ttr = safeGetText(titleShape);
  if (ttr) {
    ttr.getTextStyle().setFontFamily("Helvetica").setFontSize(48).setBold(true).setForegroundColor("#000000");
    ttr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }
  const names = (models || []).map(m => m.firstName).sort().join(", ");
  const subtitleShape = slide.insertTextBox(names || " ", 50, 220, pageWidth - 100, 200);
  const str = safeGetText(subtitleShape);
  if (str) {
    str.getTextStyle().setFontFamily("Helvetica").setFontSize(12).setForegroundColor("#000000");
    str.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }
  const marker = slide.insertTextBox(DIVIDER_MARKER, 5, 5, 50, 10);
  const mtr = safeGetText(marker);
  if (mtr) mtr.getTextStyle().setFontSize(1).setForegroundColor("#FFFFFF");
}

/***** TEXT REPLACEMENT *****/
function rebuildTextFromTemplate(slide, templateSlide, rowData, uuid, dateIndex, booking) {
  slide.getShapes().forEach(sh => {
    const tr = safeGetText(sh);
    if (!tr) return;
    const t = safeAsString(tr).trim();
    if (!/^[a-f0-9-]{36}$/i.test(t)) {
      try { sh.remove(); } catch (e) {}
    }
  });

  templateSlide.getShapes().forEach(tShape => {
    const srcTR = safeGetText(tShape);
    if (!srcTR) return;
    const newShape = slide.insertShape(
      tShape.getShapeType(), tShape.getLeft(), tShape.getTop(), tShape.getWidth(), tShape.getHeight()
    );
    const dstTR = safeGetText(newShape);
    if (dstTR) dstTR.setText(srcTR.asString());
  });

  let hasUUID = false;
  slide.getShapes().forEach(sh => {
    const tr = safeGetText(sh);
    if (tr && /^[a-f0-9-]{36}$/i.test(safeAsString(tr).trim())) hasUUID = true;
  });
  if (!hasUUID) {
    const marker = slide.insertTextBox(uuid, 5, 5, 50, 10);
    const mtr = safeGetText(marker);
    if (mtr) mtr.getTextStyle().setFontSize(1).setForegroundColor("#FFFFFF");
  }

  const values = {
    ID: uuid,
    NAME: rowData["NAME"] || "",
    NOTE: rowData["NOTE"] || "",
    NOTES: rowData["MORE NOTES"] || "",
    "MORE NOTES": rowData["MORE NOTES"] || "",
    INSTAGRAM: rowData["INSTAGRAM"] || "",
    FEE: rowData["FEE"] ? (String(rowData["FEE"]).startsWith("£") ? rowData["FEE"] : "£" + rowData["FEE"]) : "",
    AGENCY: (String(booking || "").toUpperCase() === "DB") ? "DIRECT BOOKING" : (rowData["AGENCY"] || ""),
    BOOKING: rowData["BOOKING"] || "",
    REC: rowData["REC"] || ""
  };

  let statusVal = rowData["STATUS"] ? String(rowData["STATUS"]).toUpperCase() : "";
  if (dateIndex !== -1 && rowData["DATE"] && statusVal.includes("COMING TO CASTING")) {
    statusVal = `${statusVal} – ${String(rowData["DATE"]).toUpperCase()}`;
  }
  values["STATUS"] = statusVal;

  Object.keys(values).forEach(key => {
    if (key === "INSTAGRAM") {
      handleInstagram_(slide, values[key]);
    } else if (key === "STATUS") {
      const raw = (rowData["STATUS"] || "").toString().toUpperCase();
      replaceField_(slide, key, values[key], (style) => {
        if (raw.includes("REQUESTED") || raw.includes("CONTACTED")) style.setForegroundColor("#FF0000");
        else if (raw === "TBC") style.setForegroundColor("#00AA00");
        else if (raw.includes("COMING TO CASTING") || raw.includes("OPTION") || raw.includes("CONFIRMED"))
          style.setForegroundColor("#0000FF");
        else style.setForegroundColor("#000000");
      });
    } else {
      replaceField_(slide, key, values[key]);
    }
  });
}
function replaceField_(slide, fieldKey, val, extraStyleFn) {
  const placeholder = "{{" + fieldKey + "}}";
  slide.getShapes().forEach(shape => {
    let tr;
    try {
      if (!shape || typeof shape.getText !== "function") return;
      tr = shape.getText();
    } catch (e) { return; }
    if (!tr || typeof tr.asString !== "function") return;

    let txt; try { txt = tr.asString(); } catch (e) { return; }
    if (!txt || !txt.includes(placeholder)) return;

    const outVal = (fieldKey === "INSTAGRAM") ? (val || "") : (val || "").toUpperCase();

    try { tr.clear(); tr.insertText(0, outVal); } catch (e) { return; }

    const fs = FIELD_STYLES[fieldKey];
    if (fs) {
      try {
        const style = tr.getTextStyle();
        if (fs.font) style.setFontFamily(fs.font);
        if (fs.size) style.setFontSize(fs.size);
        if (typeof fs.bold === "boolean") style.setBold(fs.bold);
        if (fs.color) style.setForegroundColor(fs.color);

        const ps = tr.getParagraphStyle();
        if (fs.align) {
          if (fs.align === "LEFT")   ps.setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
          if (fs.align === "RIGHT")  ps.setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
          if (fs.align === "CENTER") ps.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        }
      } catch (e) {}
    }
    if (extraStyleFn) {
      try { extraStyleFn(tr.getTextStyle(), shape); } catch (e) {}
    }
  });
}
function handleInstagram_(slide, rawVal) {
  let handle = "";
  if (rawVal) {
    handle = String(rawVal).trim();
    if (/^https?:\/\//i.test(handle)) {
      handle = handle.replace(/https?:\/\/(www\.)?instagram\.com\//i, "").replace(/\/$/, "");
    }
    handle = handle.replace(/^@/, "").toLowerCase();
    handle = "@" + handle;
  }
  if (handle && handle.length > 1) {
    const url = "https://www.instagram.com/" + handle.slice(1) + "/";
    replaceField_(slide, "INSTAGRAM", handle, (style) => {
      style.setLinkUrl(url);
      style.setForegroundColor("#000000").setUnderline(false);
    });
  } else {
    replaceField_(slide, "INSTAGRAM", "", (style) => {
      try { style.setLinkUrl(null); } catch (e) {}
      style.setForegroundColor("#000000").setUnderline(false);
    });
  }
}
