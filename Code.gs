/**
 * NICU Inventory Manager — Google Apps Script backend.
 *
 * Deploy this bound to the "nicu_inventory_sheet_template" Google Sheet
 * (Equipment / NonConsumables / Consumables / History / Repairs / Movements /
 * Categories / Locations tabs).
 *
 * Equipment tab holds Assets only (Type="equipment"). NonConsumables tab holds
 * Non-Consumables only (Type="nonconsumable") — same column layout as Equipment,
 * so migrating rows between them is a straight copy-paste. Feature-specific
 * columns:
 *   AssignedPatient      (Equipment tab only, e.g. "Baby of Tan — Bed 8"; blank = unassigned)
 *   Unit                 (Equipment / NonConsumables / Consumables — "RHU" | "NICU" |
 *                         "PICU" | "Paediatric": which manager is responsible for this
 *                         item day to day — RHU for ventilators/CPAP/HHHFNC/aEEG-EEG/
 *                         cooling therapy and their supplies, or a ward manager for
 *                         monitors/pumps/incubators/radiant warmers and theirs.)
 *   BoxCount, UnitsPerBox (NonConsumables / Consumables — how the stock is packaged,
 *                         e.g. 3 boxes of 20.)
 *   QtyClean, QtyDirty,
 *   QtyReprocessing      (NonConsumables tab only — pooled-quantity reprocessing
 *                         state; replaces the older single-item SterilizationStatus.
 *                         A "unit" moves Clean -> Dirty on use, Dirty -> Reprocessing
 *                         when sent for cleaning, Reprocessing -> Clean once returned.)
 *   TotalIssued          (Consumables tab only — cumulative units issued over time;
 *                         CurrentStock is "number left". Bumped automatically by the
 *                         issueStock action.)
 *   LastModified         (Equipment / NonConsumables / Consumables — ISO timestamp,
 *                         server-stamped on every write. Used for conflict detection:
 *                         if editItem's payload includes an older lastModified than
 *                         what's in the sheet, the edit is rejected rather than
 *                         silently overwriting someone else's more recent change.)
 * Everything here reads/writes columns by header name, so exact column order
 * doesn't matter and unused columns in either tab are simply left blank. Add the
 * new columns above to the Equipment/NonConsumables/Consumables sheet headers;
 * anything not yet added is silently skipped (try/catch per optional column), so
 * this deploys safely even before you've added every column.
 *
 * Setup:
 *   1. Open the Sheet -> Extensions -> Apps Script.
 *   2. Delete any starter code, paste this whole file in as Code.gs.
 *   3. Deploy -> New deployment -> type: "Web app".
 *        Execute as: Me
 *        Who has access: Anyone
 *   4. Copy the Web app URL (ends in /exec) into the prototype's Cloud Sync
 *      settings panel.
 * *   Note: "Categories" and "Locations" tabs (Type|Category, and Location respectively) are
 *   created automatically the first time they're needed, so no manual setup is required.
 *
 * API:
 *   GET  ?action=getChangesSince&since=<serverTime>&historyRows=<n>&movementsRows=<n>
 *        -> only what changed since the last sync: item rows whose LastModified is
 *           newer than `since`, IDs deleted since then, History/Movements rows
 *           appended after row n, plus repairs/categories/locations (small, sent
 *           in full). Returns { resetRequired:true } when a full getAll is needed.
 *        Requires a "LastModified" column on Equipment/NonConsumables/Consumables.
 *        A "Deletions" tab is created automatically.
 *   GET  ?action=getAll   -> { equipment:[] (Assets + Non-Consumables merged, same as before),
 *                              consumables:[], history:[], repairs:[], movements:[],
 *                              categories:{ equipment:[], nonconsumable:[], consumable:[] },
 *                              locations:[] }
 *   POST body (text/plain, JSON-encoded): { action: "...", payload: {...} }
 *     actions: issueStock, receiveStock, moveAsset, service, repair,
 *              completeRepair, auditMark, scanLog,
 *              addEquipment, batchAddEquipment, addConsumable, deleteItem, editItem,
 *              logUse, logReprocess, logReplace, markDirty, markClean, setPhoto,
 *              addCategory, renameCategory, deleteCategory,
 *              addLocation, renameLocation, deleteLocation,
 *              aiQuery, aiPhoto
 *
 *   Photo      (Equipment / NonConsumables — set via the setPhoto action, or
 *              included at registration time via addEquipment's photo field.
 *              Stores a Drive view URL, not the raw image, to keep every
 *              getAll sync small — see savePhotoToDrive_. Requires adding a
 *              "Photo" column header to the Equipment/NonConsumables tabs the
 *              same way HospitalFileRef was added, and requires this script
 *              to have Drive access (Apps Script will prompt for that
 *              authorization the first time a photo is saved).)
 *
 * AI features (aiQuery / aiPhoto) — optional:
 *   These relay to Anthropic's Claude API server-side, so your API key never
 *   touches the browser. To enable them:
 *     1. Get an API key from https://console.anthropic.com
 *     2. In the Apps Script editor: Project Settings (gear icon) -> Script Properties
 *        -> Add property: name "ANTHROPIC_API_KEY", value <your key> -> Save.
 *     3. Redeploy (Deploy -> Manage deployments -> Edit -> New version).
 *   Without a key configured, aiQuery/aiPhoto return {ok:false, error:"..."} and
 *   the app automatically falls back to its built-in local answer engine.
 */

const SHEETS = {
  EQUIPMENT: "Equipment",
  NONCONSUMABLES: "NonConsumables",
  CONSUMABLES: "Consumables",
  HISTORY: "History",
  REPAIRS: "Repairs",
  MOVEMENTS: "Movements",
  CATEGORIES: "Categories",
  LOCATIONS: "Locations",
  DELETIONS: "Deletions",
};

function getSs_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** Stamps LastModified with the server's own clock. Optional column — safe to call
 *  even if a sheet doesn't have it. Called by every action that changes a row, so
 *  editItem's conflict check always has an accurate timestamp to compare against. */
function touchLastModified_(sh, row) {
  const stamp = new Date().toISOString();
  try { setCellByHeader_(sh, row, "LastModified", stamp); } catch (e) { /* column optional */ }
  return stamp;
}

function getSheet_(name) {
  const sh = getSs_().getSheetByName(name);
  if (!sh) throw new Error("Missing sheet tab: " + name);
  return sh;
}

// Header lookups (setCellByHeader_/getCellByHeader_) used to call
// sh.getDataRange().getValues() -- which reads EVERY row in the sheet -- just
// to find which column a header name lives in. With a 500+ row asset
// register, a single action like "service" or "moveAsset" (which each touch
// 2-3 columns) was doing several full-sheet reads back to back, and
// "addEquipment" was doing up to ~10. That's the main cause of the sync lag:
// each of those reads is a separate round-trip to the Sheets backend.
// Caching headers per sheet (and reading only row 1, not the whole sheet)
// cuts that down to one cheap read per sheet per request.
const _headerCache_ = {};
function getHeaders_(sh) {
  const key = sh.getSheetId();
  if (_headerCache_[key]) return _headerCache_[key];
  const lastCol = sh.getLastColumn();
  const headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  _headerCache_[key] = headers;
  return headers;
}

/** Appends one Equipment/NonConsumable row and its optional columns (Supplier,
 *  BiomedicalContact, AssignedPatient, Unit, BoxCount, UnitsPerBox, QtyClean,
 *  QtyDirty, QtyReprocessing, LastModified). Shared by addEquipment (one item)
 *  and batchAddEquipment (many). */
function appendEquipmentRow_(p) {
  const sh = sheetForType_(p.type);
  const headers = getHeaders_(sh);
  const addStamp = new Date().toISOString();

  const rowMap = {
    ID: p.id, Type: p.type, Category: p.category, Name: p.name,
    AssetNumber: p.assetNumber, HospitalAssetId: p.hospitalAssetId, SerialNumber: p.serialNumber,
    Manufacturer: p.manufacturer, Model: p.model, Warranty: p.warranty, PurchaseDate: p.purchaseDate,
    Location: p.location, Status: p.status || "Working", LastPM: p.lastPM, NextPM: p.nextPM,
    LastModified: addStamp,
  };
  if (p.supplier) rowMap.Supplier = p.supplier;
  if (p.biomedicalContact) rowMap.BiomedicalContact = p.biomedicalContact;
  if (p.type === "equipment" && p.assignedPatient) rowMap.AssignedPatient = p.assignedPatient;
  if (p.type === "nonconsumable") {
    if (p.boxCount !== undefined && p.boxCount !== "") rowMap.BoxCount = p.boxCount;
    if (p.unitsPerBox !== undefined && p.unitsPerBox !== "") rowMap.UnitsPerBox = p.unitsPerBox;
    if (p.qtyClean !== undefined && p.qtyClean !== "") rowMap.QtyClean = p.qtyClean;
    if (p.qtyDirty !== undefined && p.qtyDirty !== "") rowMap.QtyDirty = p.qtyDirty;
    if (p.qtyReprocessing !== undefined && p.qtyReprocessing !== "") rowMap.QtyReprocessing = p.qtyReprocessing;
  }
  if (p.unit) rowMap.Unit = p.unit;
  if (p.hospitalFileRef) rowMap.HospitalFileRef = p.hospitalFileRef;
  // A photo taken at registration time gets uploaded to Drive right away so
  // it's there from the first sync -- see savePhotoToDrive_ for why we store
  // a Drive URL here rather than the raw base64 (a photo can be 1-3MB; a
  // sheet cell tops out around 50,000 characters and every getAll call would
  // have to ship that blob to every device on every sync).
  if (p.photo) {
    try { rowMap.Photo = savePhotoToDrive_(p.photo, p.id); }
    catch (e) { /* Drive upload failed or Photo column not added yet -- skip, not fatal */ }
  }

  // Headers not present in rowMap are simply left blank -- same "add the
  // column whenever you're ready" contract as the rest of this file.
  const row = headers.map(h => (Object.prototype.hasOwnProperty.call(rowMap, h) ? rowMap[h] : ""));
  sh.appendRow(row);
  return addStamp;
}

/** Uploads a base64 data-URL photo to a Drive folder (created once, reused after)
 *  and returns a shareable view URL to store in the sheet's Photo column. Storing
 *  a URL instead of the raw image keeps every getAll sync fast and small even as
 *  more items get photos -- inlining base64 image data in the sheet would bloat
 *  every single device's sync payload by however many photos exist, not just the
 *  one being added. Requires the Photo column to be added to Equipment /
 *  NonConsumables the same way HospitalFileRef was. */
function savePhotoToDrive_(dataUrl, id) {
  const match = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
  if (!match) throw new Error("Not a base64 image data URL");
  const mimeType = match[1];
  const bytes = Utilities.base64Decode(match[2]);
  const ext = mimeType.split("/")[1] || "jpg";
  const blob = Utilities.newBlob(bytes, mimeType, id + "." + ext);

  const folderName = "NICU Inventory Photos";
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // IMPORTANT: drive.google.com/uc?export=view frequently fails to actually
  // serve the image when used as an <img src="..."> -- Google Drive often
  // returns an interstitial "can't scan this file for viruses" or sign-in
  // page instead of the image bytes for that URL pattern, especially in a
  // browser with no active Google session. The /thumbnail endpoint is the
  // reliable way to hotlink a publicly-shared Drive file as an image.
  return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";
}

/** Returns the correct tab for a given item Type ("equipment" | "nonconsumable" | "consumable"). */
function sheetForType_(type) {
  if (type === "consumable") return getSheet_(SHEETS.CONSUMABLES);
  if (type === "nonconsumable") return getSheet_(SHEETS.NONCONSUMABLES);
  return getSheet_(SHEETS.EQUIPMENT);
}

/** Locates an Asset or Non-Consumable by ID without the caller needing to know which
 *  of the two tabs it lives in. Checks Equipment first, then NonConsumables.
 *  Returns { sh, row } or null if not found in either tab. Consumables are handled
 *  separately by their own actions (issueStock/receiveStock), so they're not searched here. */
function findItemLocation_(id) {
  let sh = getSheet_(SHEETS.EQUIPMENT);
  let row = findRowIndexById_(sh, "ID", id);
  if (row !== -1) return { sh: sh, row: row };
  sh = getSheet_(SHEETS.NONCONSUMABLES);
  row = findRowIndexById_(sh, "ID", id);
  if (row !== -1) return { sh: sh, row: row };
  return null;
}

/** Converts one sheet row into a plain object keyed by header. Dates become
 *  yyyy-MM-dd, except timestamp columns (LastModified / Timestamp), which keep
 *  full ISO time -- trimming those to a date would break change detection. */
function rowToObj_(headers, r) {
  const obj = {};
  headers.forEach((h, i) => {
    let v = r[i];
    if (v instanceof Date) {
      v = (h === "LastModified" || h === "Timestamp")
        ? v.toISOString()
        : Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    obj[h] = v;
  });
  return obj;
}
function rowNotEmpty_(r) { return r.some(c => c !== "" && c !== null); }

/** Read a sheet into an array of plain objects keyed by header row. */
function sheetToObjects_(name) {
  const sh = getSheet_(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 1) return [];
  const headers = values[0];
  return values.slice(1).filter(rowNotEmpty_).map(r => rowToObj_(headers, r));
}

/* ---------------- DELTA SYNC HELPERS ---------------- */
// Phones ask "what changed since <serverTime>?" instead of downloading the
// whole register every refresh. Rows are matched by their LastModified stamp
// (set by every write action, and by onEdit below for hand edits in the
// Sheet). A 2-minute overlap is subtracted from `since` so a write that was
// stamped a moment before the previous sync but committed just after it is
// never missed -- re-sending a row is harmless, the app just upserts it.
const DELTA_OVERLAP_MS = 2 * 60 * 1000;
function cellTimeMs_(v) {
  if (v instanceof Date) return v.getTime();
  if (v === "" || v === null || v === undefined) return NaN;
  return new Date(v).getTime();
}
/** Item rows (Equipment / NonConsumables / Consumables) changed after sinceMs. */
function changedRowsSince_(name, sinceMs) {
  const sh = getSheet_(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const lmCol = headers.indexOf("LastModified");
  if (lmCol === -1) {
    // No LastModified column yet -- can't tell what changed, so send everything.
    return values.slice(1).filter(rowNotEmpty_).map(r => rowToObj_(headers, r));
  }
  return values.slice(1).filter(r => {
    if (!rowNotEmpty_(r)) return false;
    const t = cellTimeMs_(r[lmCol]);
    return !isNaN(t) && t > sinceMs;
  }).map(r => rowToObj_(headers, r));
}
/** Number of data rows (excluding header) in an append-only log tab. */
function dataRowCount_(name) {
  const sh = getSs_().getSheetByName(name);
  return sh ? Math.max(0, sh.getLastRow() - 1) : 0;
}
/** Rows appended to an append-only tab (History / Movements) after the first
 *  `already` data rows. If the tab now has FEWER rows than that (someone
 *  deleted rows by hand), the phone's watermark is meaningless -> reset. */
function rowsAppendedAfter_(name, already) {
  const sh = getSheet_(name);
  const total = Math.max(0, sh.getLastRow() - 1);
  if (already > total) return { reset: true, total: total, rows: [] };
  if (already === total) return { total: total, rows: [] };
  const headers = getHeaders_(sh);
  const vals = sh.getRange(already + 2, 1, total - already, headers.length).getValues();
  return { total: total, rows: vals.filter(rowNotEmpty_).map(r => rowToObj_(headers, r)) };
}
/** Deletions tab auto-creates itself, like Categories/Locations. */
function getDeletionsSheet_() {
  let sh = getSs_().getSheetByName(SHEETS.DELETIONS);
  if (!sh) {
    sh = getSs_().insertSheet(SHEETS.DELETIONS);
    sh.appendRow(["ItemId", "Timestamp"]);
  }
  return sh;
}
function deletionsSince_(sinceMs) {
  const sh = getDeletionsSheet_();
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const t = cellTimeMs_(values[i][1]);
    if (values[i][0] && !isNaN(t) && t > sinceMs) out.push({ ItemId: String(values[i][0]) });
  }
  return out;
}

/** Simple trigger: when someone edits an item row directly in the Google
 *  Sheet (not through the app), stamp its LastModified so phones pick the
 *  edit up on their next quick refresh. Runs automatically -- no setup. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const name = sh.getName();
    if ([SHEETS.EQUIPMENT, SHEETS.NONCONSUMABLES, SHEETS.CONSUMABLES].indexOf(name) === -1) return;
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const lmCol = headers.indexOf("LastModified");
    if (lmCol === -1) return;
    const r0 = e.range.getRow(), n = e.range.getNumRows();
    if (e.range.getNumColumns() === 1 && e.range.getColumn() === lmCol + 1) return; // editing the stamp itself
    const start = Math.max(2, r0);
    const count = r0 + n - start;
    if (count <= 0) return;
    const stamp = new Date().toISOString();
    const col = [];
    for (let i = 0; i < count; i++) col.push([stamp]);
    sh.getRange(start, lmCol + 1, count, 1).setValues(col);
  } catch (err) { /* never block a manual edit */ }
}

function findRowIndexById_(sh, idColName, id) {
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf(idColName);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) return i + 1; // 1-based sheet row
  }
  return -1;
}

function setCellByHeader_(sh, rowIndex, headerName, value) {
  const headers = getHeaders_(sh);
  const col = headers.indexOf(headerName);
  if (col === -1) throw new Error("Unknown column: " + headerName);
  sh.getRange(rowIndex, col + 1).setValue(value);
}

function getCellByHeader_(sh, rowIndex, headerName) {
  const headers = getHeaders_(sh);
  const col = headers.indexOf(headerName);
  return sh.getRange(rowIndex, col + 1).getValue();
}

/** Reads an entire row once and returns it as a {headerName: value} object, so a
 *  handler that needs several fields off the same row (e.g. CurrentStock and
 *  TotalIssued) does one getRange().getValues() instead of one per field. */
function getRowByHeader_(sh, rowIndex) {
  const headers = getHeaders_(sh);
  const lastCol = headers.length;
  if (lastCol === 0) return {};
  const values = sh.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  const obj = {};
  headers.forEach(function (h, i) { obj[h] = values[i]; });
  return obj;
}

/** Batched version of setCellByHeader_: reads the row once, applies every
 *  {headerName: value} pair from valuesByHeader in memory, then writes the
 *  whole row back in a single setValues() call -- one read + one write no
 *  matter how many fields are being changed, instead of a separate
 *  getRange().setValue() round-trip per field. This is what every doPost
 *  action below uses now instead of several setCellByHeader_/
 *  touchLastModified_ calls back to back on the same row.
 *  A header name that doesn't exist as a column yet is silently skipped --
 *  same "add the column whenever you're ready" tolerance setCellByHeader_
 *  gave individual optional fields, just applied per-field within the batch
 *  instead of aborting the whole write. */
function setCellsByHeader_(sh, rowIndex, valuesByHeader) {
  const headers = getHeaders_(sh);
  const lastCol = headers.length;
  if (lastCol === 0) return;
  const range = sh.getRange(rowIndex, 1, 1, lastCol);
  const row = range.getValues()[0];
  let touched = false;
  Object.keys(valuesByHeader).forEach(function (headerName) {
    const col = headers.indexOf(headerName);
    if (col === -1) return; // column not added yet -- skip this field, don't fail the batch
    row[col] = valuesByHeader[headerName];
    touched = true;
  });
  if (touched) range.setValues([row]);
}

/** Categories sheet auto-creates itself (with a header row) the first time it's touched,
 *  so there's no manual setup step required for it like the other tabs. */
function getCategoriesSheet_() {
  let sh = getSs_().getSheetByName(SHEETS.CATEGORIES);
  if (!sh) {
    sh = getSs_().insertSheet(SHEETS.CATEGORIES);
    sh.appendRow(["Type", "Category"]);
  }
  return sh;
}

/** Reads the Categories sheet into { equipment:[...], nonconsumable:[...], consumable:[...] }. */
function getCategoriesForSync() {
  const sh = getCategoriesSheet_();
  const values = sh.getDataRange().getValues();
  const out = { equipment: [], nonconsumable: [], consumable: [] };
  for (let i = 1; i < values.length; i++) {
    const type = String(values[i][0] || "").trim();
    const category = String(values[i][1] || "").trim();
    if (!type || !category) continue;
    if (!out[type]) out[type] = [];
    out[type].push(category);
  }
  return out;
}

function findCategoryRow_(sh, type, name) {
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === type && String(values[i][1]).toLowerCase() === String(name).toLowerCase()) {
      return i + 1; // 1-based sheet row
    }
  }
  return -1;
}

/** Re-tags every Equipment/NonConsumables/Consumables row using oldName to newName,
 *  mirroring what the app already does to its own local copy of the data so the
 *  sheet never disagrees with it. */
function retagItemsCategory_(type, oldName, newName) {
  const sh = sheetForType_(type);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const catCol = headers.indexOf("Category");
  if (catCol === -1) return 0;
  const lmCol = headers.indexOf("LastModified");
  const stamp = new Date().toISOString();
  let updated = 0;
  const colValues = [], lmValues = [];
  for (let i = 1; i < values.length; i++) {
    let v = values[i][catCol];
    let lm = lmCol === -1 ? "" : values[i][lmCol];
    if (v === oldName) { v = newName; lm = stamp; updated++; }
    colValues.push([v]); lmValues.push([lm]);
  }
  if (updated > 0) {
    sh.getRange(2, catCol + 1, colValues.length, 1).setValues(colValues);
    if (lmCol !== -1) sh.getRange(2, lmCol + 1, lmValues.length, 1).setValues(lmValues); // so phones see the re-tag
  }
  return updated;
}

/** Counts items still using a category — checked server-side too (not just client-side)
 *  in case two people edit categories around the same time. */
function countItemsUsingCategory_(type, name) {
  const sh = sheetForType_(type);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const catCol = headers.indexOf("Category");
  if (catCol === -1) return 0;
  let count = 0;
  for (let i = 1; i < values.length; i++) {
    if (values[i][catCol] === name) count++;
  }
  return count;
}

/** Locations sheet auto-creates itself (with a header row) the first time it's touched.
 *  Locations only apply to the "Location" column on the Equipment and NonConsumables
 *  tabs (beds/rooms for large equipment and non-consumables) — Consumables use a
 *  separate "StoreLocation" field that isn't part of this registry. */
function getLocationsSheet_() {
  let sh = getSs_().getSheetByName(SHEETS.LOCATIONS);
  if (!sh) {
    sh = getSs_().insertSheet(SHEETS.LOCATIONS);
    sh.appendRow(["Location"]);
  }
  return sh;
}

/** Reads the Locations sheet into a flat array of location names. */
function getLocationsForSync() {
  const sh = getLocationsSheet_();
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][0] || "").trim();
    if (name) out.push(name);
  }
  return out;
}

function findLocationRow_(sh, name) {
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase() === String(name).toLowerCase()) return i + 1;
  }
  return -1;
}

/** Re-tags every Equipment + NonConsumables row using oldName to newName, mirroring
 *  what the app already does to its own local copy of the data so the sheet never
 *  disagrees with it. */
function retagItemsLocation_(oldName, newName) {
  let updated = 0;
  [SHEETS.EQUIPMENT, SHEETS.NONCONSUMABLES].forEach(function (sheetName) {
    const sh = getSheet_(sheetName);
    const values = sh.getDataRange().getValues();
    const headers = values[0];
    const locCol = headers.indexOf("Location");
    if (locCol === -1) return;
    const lmCol = headers.indexOf("LastModified");
    const stamp = new Date().toISOString();
    let changed = 0;
    const colValues = [], lmValues = [];
    for (let i = 1; i < values.length; i++) {
      let v = values[i][locCol];
      let lm = lmCol === -1 ? "" : values[i][lmCol];
      if (v === oldName) { v = newName; lm = stamp; changed++; }
      colValues.push([v]); lmValues.push([lm]);
    }
    if (changed > 0) {
      sh.getRange(2, locCol + 1, colValues.length, 1).setValues(colValues);
      if (lmCol !== -1) sh.getRange(2, lmCol + 1, lmValues.length, 1).setValues(lmValues);
      updated += changed;
    }
  });
  return updated;
}

/** Counts items still at a location — checked server-side too, same reasoning as categories. */
function countItemsUsingLocation_(name) {
  let count = 0;
  [SHEETS.EQUIPMENT, SHEETS.NONCONSUMABLES].forEach(function (sheetName) {
    const sh = getSheet_(sheetName);
    const values = sh.getDataRange().getValues();
    const headers = values[0];
    const locCol = headers.indexOf("Location");
    if (locCol === -1) return;
    for (let i = 1; i < values.length; i++) {
      if (values[i][locCol] === name) count++;
    }
  });
  return count;
}

function todayIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}
function nowTimeStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm");
}
function addOneYear_(isoDate) {
  const d = isoDate ? new Date(isoDate) : new Date();
  d.setFullYear(d.getFullYear() + 1);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function appendHistory_(itemId, action, who, location, reason) {
  const sh = getSheet_(SHEETS.HISTORY);
  sh.appendRow([
    new Date().toISOString(),
    itemId,
    action,
    who || "Unknown",
    todayIso_(),
    nowTimeStr_(),
    location || "",
    reason || "",
  ]);
}

/* ---------------- ONE-TIME MIGRATION ----------------
 * Run migrateNonConsumablesToOwnTab manually, once, after this update:
 *   Apps Script editor -> select "migrateNonConsumablesToOwnTab" in the function
 *   dropdown at the top -> click Run. Check the execution log (View -> Logs) to
 *   confirm how many rows moved. Safe to run again — it's a no-op if no
 *   Type="nonconsumable" rows remain in the Equipment tab.
 *   (Named without a trailing underscore, unlike the rest of this file's internal
 *   helpers, specifically so it shows up in the Run dropdown — Apps Script hides
 *   underscore-suffixed functions from it.)
 */
function migrateNonConsumablesToOwnTab() {
  const eqSh = getSheet_(SHEETS.EQUIPMENT);
  let ncSh = getSs_().getSheetByName(SHEETS.NONCONSUMABLES);
  const values = eqSh.getDataRange().getValues();
  if (values.length < 1) return;
  const headers = values[0];
  if (!ncSh) {
    ncSh = getSs_().insertSheet(SHEETS.NONCONSUMABLES);
    ncSh.appendRow(headers); // same column layout as Equipment, so this is a straight row copy
  }
  const typeCol = headers.indexOf("Type");
  if (typeCol === -1) {
    throw new Error('Equipment tab has no "Type" column — cannot identify non-consumable rows.');
  }
  const rowsToMove = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][typeCol]) === "nonconsumable") rowsToMove.push(i + 1); // 1-based sheet row
  }
  if (rowsToMove.length === 0) {
    Logger.log("No nonconsumable rows found in Equipment — nothing to migrate.");
    return;
  }
  // Copy to NonConsumables first...
  rowsToMove.forEach(function (rowIdx) {
    const rowValues = eqSh.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
    ncSh.appendRow(rowValues);
  });
  // ...then delete from Equipment bottom-up, so earlier row indices don't shift
  // out from under the ones still queued for deletion.
  rowsToMove.reverse().forEach(function (rowIdx) {
    eqSh.deleteRow(rowIdx);
  });
  Logger.log("Migrated " + rowsToMove.length + " non-consumable row(s) to the NonConsumables tab.");
}

/* ---------------- GET ---------------- */
// History and Movements only ever grow -- every action anyone takes appends a
// row -- and the regular sync (getAll) used to ship the entire tab to every
// device on every single sync. That's the main thing that keeps getting
// slower over time as the sheet accumulates activity. Most of what the app
// actually displays (recent activity feeds, per-item timelines, "events this
// week") only looks back a few days/weeks anyway, so getAll now caps both to
// the last ~90 days. A couple of report types (and the full backup export)
// legitimately want the complete, uncapped record -- those fetch it
// on-demand via the separate getFullHistory action below instead.
const SYNC_HISTORY_DAYS = 90;
function recentCutoffIso_(days) {
  const d = new Date();
  d.setDate(d.getDate() - days);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "getAll";
  let payload;
  // Captured BEFORE reading anything, so the next delta's `since` can never
  // skip a change that landed while this request was reading the tabs.
  const serverTime = new Date().toISOString();
  if (action === "getAll") {
    const cutoff = recentCutoffIso_(SYNC_HISTORY_DAYS);
    const allHistory = sheetToObjects_(SHEETS.HISTORY).sort((a, b) =>
      (b.Timestamp || "").localeCompare(a.Timestamp || "")
    );
    const allMovements = sheetToObjects_(SHEETS.MOVEMENTS);
    payload = {
      // Frontend expects one combined "equipment" array (Assets + Non-Consumables,
      // distinguished by each row's Type field) — same shape as before the tab split.
      equipment: sheetToObjects_(SHEETS.EQUIPMENT).concat(sheetToObjects_(SHEETS.NONCONSUMABLES)),
      consumables: sheetToObjects_(SHEETS.CONSUMABLES),
      history: allHistory.filter(h => (h.Date || (h.Timestamp || "").slice(0, 10)) >= cutoff),
      repairs: sheetToObjects_(SHEETS.REPAIRS),
      movements: allMovements.filter(m => (m.Date || "") >= cutoff),
      categories: getCategoriesForSync(),
      locations: getLocationsForSync(),
      historyCappedToDays: SYNC_HISTORY_DAYS,
      // Watermarks the app sends back on its next quick refresh (getChangesSince).
      serverTime: serverTime,
      historyRowCount: dataRowCount_(SHEETS.HISTORY),
      movementsRowCount: dataRowCount_(SHEETS.MOVEMENTS),
    };
  } else if (action === "getChangesSince") {
    const p = e.parameter || {};
    const sinceMs = new Date(p.since).getTime();
    if (isNaN(sinceMs)) {
      payload = { resetRequired: true, serverTime: serverTime };
    } else {
      const cutoffMs = sinceMs - DELTA_OVERLAP_MS;
      const hist = rowsAppendedAfter_(SHEETS.HISTORY, Number(p.historyRows) || 0);
      const mov = rowsAppendedAfter_(SHEETS.MOVEMENTS, Number(p.movementsRows) || 0);
      if (hist.reset || mov.reset) {
        payload = { resetRequired: true, serverTime: serverTime };
      } else {
        const equipment = changedRowsSince_(SHEETS.EQUIPMENT, cutoffMs)
          .concat(changedRowsSince_(SHEETS.NONCONSUMABLES, cutoffMs));
        const consumables = changedRowsSince_(SHEETS.CONSUMABLES, cutoffMs);
        // Don't report an ID as deleted if it has since been re-added.
        const present = {};
        equipment.concat(consumables).forEach(r => { present[String(r.ID)] = true; });
        payload = {
          delta: true,
          equipment: equipment,
          consumables: consumables,
          deletions: deletionsSince_(cutoffMs).filter(d => !present[d.ItemId]),
          history: hist.rows.sort((a, b) => String(b.Timestamp || "").localeCompare(String(a.Timestamp || ""))),
          movements: mov.rows,
          repairs: sheetToObjects_(SHEETS.REPAIRS),
          categories: getCategoriesForSync(),
          locations: getLocationsForSync(),
          serverTime: serverTime,
          historyRowCount: hist.total,
          movementsRowCount: mov.total,
        };
      }
    }
  } else if (action === "getFullHistory") {
    // Uncapped -- used for reports/exports that need the complete record
    // (e.g. "Consumables Issued", "Asset Movement", full history CSV/backup),
    // not for the regular per-device sync.
    payload = {
      history: sheetToObjects_(SHEETS.HISTORY).sort((a, b) =>
        (b.Timestamp || "").localeCompare(a.Timestamp || "")
      ),
      movements: sheetToObjects_(SHEETS.MOVEMENTS),
    };
  } else {
    payload = { error: "Unknown action: " + action };
  }
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- POST ---------------- */
function doPost(e) {
  let result = { ok: true };
  let lock = null;
  let lockAcquired = false;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const p = body.payload || {};

    // Google Sheets is a spreadsheet, not a database -- with several nurses
    // scanning/editing at once, two requests can genuinely overlap (e.g. two
    // stock decrements both reading the same "before" CurrentStock and only
    // one surviving). LockService forces every write-action request to run
    // one at a time, like a queue, instead of racing. aiQuery/aiPhoto don't
    // touch the Sheet at all (they call Anthropic's API), so they skip the
    // lock entirely -- holding it for however long an AI call takes would
    // block every other person's scan for no reason, since there's nothing
    // to protect from a race there.
    const needsLock = action !== "aiQuery" && action !== "aiPhoto";
    if (needsLock) {
      lock = LockService.getScriptLock();
      try {
        // Generous wait: with the batched reads/writes below, even a full
        // editItem save is one read + one write, so the queue should clear
        // quickly under normal load -- this just gives it headroom.
        lock.waitLock(10000);
        lockAcquired = true;
      } catch (lockErr) {
        return ContentService.createTextOutput(JSON.stringify({
          ok: false,
          error: "System busy — please try again in a moment.",
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    if (action === "issueStock" || action === "receiveStock") {
      const sh = getSheet_(SHEETS.CONSUMABLES);
      const row = findRowIndexById_(sh, "ID", p.id);
      if (row === -1) throw new Error("Consumable not found: " + p.id);
      const rowData = getRowByHeader_(sh, row);
      const current = Number(rowData.CurrentStock) || 0;
      const delta = action === "issueStock" ? -Math.abs(p.qty || 1) : Math.abs(p.qty || 1);
      const next = Math.max(0, current + delta);
      const updates = { CurrentStock: next, LastModified: new Date().toISOString() };
      if (action === "issueStock") {
        const issuedSoFar = Number(rowData.TotalIssued) || 0;
        updates.TotalIssued = issuedSoFar + Math.abs(p.qty || 1);
      }
      setCellsByHeader_(sh, row, updates);
      appendHistory_(p.id, action === "issueStock" ? "Issued" : "Stock received", p.who, p.location, p.reason);
      result.currentStock = next;

    } else if (action === "moveAsset") {
      const loc = findItemLocation_(p.id);
      if (!loc) throw new Error("Asset not found: " + p.id);
      const sh = loc.sh, row = loc.row;
      const from = getRowByHeader_(sh, row).Location;
      setCellsByHeader_(sh, row, { Location: p.to, LastModified: new Date().toISOString() });
      const mv = getSheet_(SHEETS.MOVEMENTS);
      mv.appendRow([p.id, from, p.via || "", p.to, todayIso_()]);
      appendHistory_(p.id, "Moved", p.who, p.to, from + " -> " + p.to);

    } else if (action === "service") {
      const loc = findItemLocation_(p.id);
      if (!loc) throw new Error("Asset not found: " + p.id);
      const today = todayIso_();
      setCellsByHeader_(loc.sh, loc.row, {
        LastPM: today, NextPM: addOneYear_(today), Status: "Working",
        LastModified: new Date().toISOString(),
      });
      appendHistory_(p.id, "Serviced / Calibrated", p.who, null, "Routine PM completed");

    } else if (action === "repair") {
      const loc = findItemLocation_(p.id);
      if (!loc) throw new Error("Asset not found: " + p.id);
      setCellsByHeader_(loc.sh, loc.row, { Status: "Repair", LastModified: new Date().toISOString() });
      const rp = getSheet_(SHEETS.REPAIRS);
      rp.appendRow([p.id, p.fault || "Unspecified fault", todayIso_(), p.engineer || "Unassigned", ""]);
      appendHistory_(p.id, "Fault reported", p.who, null, p.fault || "");

    } else if (action === "completeRepair") {
      const rp = getSheet_(SHEETS.REPAIRS);
      const values = rp.getDataRange().getValues();
      const headers = values[0];
      const idCol = headers.indexOf("ItemId");
      const completedCol = headers.indexOf("Completed");
      let targetRow = -1;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][idCol]) === String(p.id) && !values[i][completedCol]) targetRow = i + 1;
      }
      if (targetRow !== -1) rp.getRange(targetRow, completedCol + 1).setValue(todayIso_());
      const loc = findItemLocation_(p.id);
      if (loc) setCellsByHeader_(loc.sh, loc.row, { Status: "Working", LastModified: new Date().toISOString() });
      appendHistory_(p.id, "Repair completed", p.who, null, "");

    } else if (action === "auditMark") {
      appendHistory_(p.id, "Audited", p.who, p.location, p.result === "found" ? "Confirmed present" : "Marked missing");

    } else if (action === "scanLog") {
      appendHistory_(p.id, "Scanned", p.who, p.location, "QR scan");

    } else if (action === "logUse") {
      appendHistory_(p.id, "Used", p.who, null, "Used in patient care");

    } else if (action === "logReprocess") {
      appendHistory_(p.id, "Reprocessed", p.who, null, "Cleaned / sterilised");

    } else if (action === "markDirty") {
      const loc = findItemLocation_(p.id);
      if (!loc) throw new Error("Item not found: " + p.id);
      setCellsByHeader_(loc.sh, loc.row, { SterilizationStatus: "Dirty", LastModified: new Date().toISOString() });
      appendHistory_(p.id, "Marked Dirty", p.who, null, "Flagged for cleaning/reprocessing");

    } else if (action === "markClean") {
      const loc = findItemLocation_(p.id);
      if (!loc) throw new Error("Item not found: " + p.id);
      setCellsByHeader_(loc.sh, loc.row, { SterilizationStatus: "Clean", LastModified: new Date().toISOString() });
      appendHistory_(p.id, "Marked Clean", p.who, null, "Cleaned / sterilised and ready for use");

    } else if (action === "logReplace") {
      const loc = findItemLocation_(p.id);
      if (loc) setCellsByHeader_(loc.sh, loc.row, { Status: "Disposed", LastModified: new Date().toISOString() });
      appendHistory_(p.id, "Replaced", p.who, null, p.reason || "End of service life");

    } else if (action === "setPhoto") {
      const loc = findItemLocation_(p.id);
      if (!loc) throw new Error("Item not found: " + p.id);
      const url = savePhotoToDrive_(p.imageBase64, p.id);
      setCellsByHeader_(loc.sh, loc.row, { Photo: url, LastModified: new Date().toISOString() });
      result.photoUrl = url;

    } else if (action === "addEquipment") {
      result.lastModified = appendEquipmentRow_(p);
      appendHistory_(p.id, "Added", p.who, p.location, "Registered new item");

    } else if (action === "batchAddEquipment") {
      // Bulk import — e.g. loading in an existing asset register. One call instead of
      // one HTTP round-trip per item, since that could be dozens of items.
      const list = Array.isArray(p.items) ? p.items : [];
      let added = 0;
      list.forEach(function (item) {
        try {
          appendEquipmentRow_(item);
          appendHistory_(item.id, "Added", p.who, item.location, "Registered new item (batch import)");
          added++;
        } catch (e) {
          // Skip a single bad row rather than failing the whole batch.
        }
      });
      result.added = added;
      result.requested = list.length;

    } else if (action === "addConsumable") {
      const sh = getSheet_(SHEETS.CONSUMABLES);
      sh.appendRow([
        p.id, p.category, p.name, p.currentStock, p.minStock, p.maxStock,
        p.batchNumber, p.expiryDate, p.supplier, p.storeLocation, "",
      ]);
      const addStamp = new Date().toISOString();
      const newRow = findRowIndexById_(sh, "ID", p.id);
      if (newRow !== -1) {
        const updates = { LastModified: addStamp };
        if (p.boxCount !== undefined && p.boxCount !== "") updates.BoxCount = p.boxCount;
        if (p.unitsPerBox !== undefined && p.unitsPerBox !== "") updates.UnitsPerBox = p.unitsPerBox;
        if (p.totalIssued !== undefined && p.totalIssued !== "") updates.TotalIssued = p.totalIssued;
        if (p.unit) updates.Unit = p.unit;
        setCellsByHeader_(sh, newRow, updates);
      }
      result.lastModified = addStamp;
      appendHistory_(p.id, "Added", p.who, p.storeLocation, "Registered new item");

    } else if (action === "deleteItem") {
      if (p.sheet === "Consumables") {
        const sh = getSheet_(SHEETS.CONSUMABLES);
        const row = findRowIndexById_(sh, "ID", p.id);
        if (row !== -1) sh.deleteRow(row);
      } else {
        const loc = findItemLocation_(p.id);
        if (loc) loc.sh.deleteRow(loc.row);
      }
      getDeletionsSheet_().appendRow([p.id, new Date().toISOString()]); // so other phones drop it on their next quick refresh
      appendHistory_(p.id, "Removed", p.who, null, "Removed from asset registry");

    } else if (action === "editItem") {
      let sh, row;
      if (p.sheet === "Consumables") {
        sh = getSheet_(SHEETS.CONSUMABLES);
        row = findRowIndexById_(sh, "ID", p.id);
      } else {
        const loc = findItemLocation_(p.id);
        if (loc) { sh = loc.sh; row = loc.row; }
      }
      if (!sh || row === -1 || row === undefined) throw new Error("Item not found: " + p.id);

      // Conflict check: if the client's edit was based on a stale copy (e.g. it queued
      // while offline, or two people had the same row open), and someone else's edit
      // has since landed, reject rather than silently overwrite their change. One read
      // of the row also gives us every field the payload might touch, in one call.
      const rowData = getRowByHeader_(sh, row);
      let conflict = false;
      if (p.lastModified) {
        const existing = rowData.LastModified;
        if (existing) {
          const existingTime = new Date(existing).getTime();
          const incomingTime = new Date(p.lastModified).getTime();
          if (!isNaN(existingTime) && !isNaN(incomingTime) && existingTime > incomingTime) {
            conflict = true;
          }
        }
      }

      if (conflict) {
        result = {
          ok: false,
          conflict: true,
          error: "This item was updated more recently by someone else. Refresh to see the latest version before editing again.",
        };
      } else {
        // Update any field present in the payload (except id/sheet/who/lastModified) by
        // matching header name, all assembled in memory first and written back in one
        // call -- lastModified is stamped with the server's own clock so comparisons
        // stay consistent regardless of client clock skew.
        const updates = {};
        Object.keys(p).forEach(function (key) {
          if (key === "id" || key === "sheet" || key === "who" || key === "lastModified") return;
          const headerName = key.charAt(0).toUpperCase() + key.slice(1); // camelCase -> Header
          updates[headerName] = p[key];
        });
        const stamp = new Date().toISOString();
        updates.LastModified = stamp;
        setCellsByHeader_(sh, row, updates);
        result.lastModified = stamp;
        appendHistory_(p.id, "Edited", p.who, p.location || null, "Details updated");
      }

    } else if (action === "addCategory") {
      const sh = getCategoriesSheet_();
      if (findCategoryRow_(sh, p.type, p.name) !== -1) {
        result = { ok: false, error: '"' + p.name + '" already exists for ' + p.type };
      } else {
        sh.appendRow([p.type, p.name]);
      }

    } else if (action === "renameCategory") {
      const sh = getCategoriesSheet_();
      const row = findCategoryRow_(sh, p.type, p.oldName);
      if (row === -1) {
        result = { ok: false, error: '"' + p.oldName + '" not found for ' + p.type };
      } else {
        sh.getRange(row, 2).setValue(p.newName);
        result.itemsUpdated = retagItemsCategory_(p.type, p.oldName, p.newName);
      }

    } else if (action === "deleteCategory") {
      const inUse = countItemsUsingCategory_(p.type, p.name);
      if (inUse > 0) {
        result = { ok: false, error: inUse + " item(s) still use \"" + p.name + "\" — re-categorize them first" };
      } else {
        const sh = getCategoriesSheet_();
        const row = findCategoryRow_(sh, p.type, p.name);
        if (row !== -1) sh.deleteRow(row);
      }

    } else if (action === "addLocation") {
      const sh = getLocationsSheet_();
      if (findLocationRow_(sh, p.name) !== -1) {
        result = { ok: false, error: '"' + p.name + '" already exists' };
      } else {
        sh.appendRow([p.name]);
      }

    } else if (action === "renameLocation") {
      const sh = getLocationsSheet_();
      const row = findLocationRow_(sh, p.oldName);
      if (row === -1) {
        result = { ok: false, error: '"' + p.oldName + '" not found' };
      } else {
        sh.getRange(row, 1).setValue(p.newName);
        result.itemsUpdated = retagItemsLocation_(p.oldName, p.newName);
      }

    } else if (action === "deleteLocation") {
      const inUse = countItemsUsingLocation_(p.name);
      if (inUse > 0) {
        result = { ok: false, error: inUse + " item(s) still at \"" + p.name + "\" — move them first" };
      } else {
        const sh = getLocationsSheet_();
        const row = findLocationRow_(sh, p.name);
        if (row !== -1) sh.deleteRow(row);
      }

    } else if (action === "aiQuery") {
      result = aiQuery_(p.question || "");

    } else if (action === "aiPhoto") {
      result = aiPhoto_(p.imageBase64 || "");

    } else {
      result = { ok: false, error: "Unknown action: " + action };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  } finally {
    if (lockAcquired) { try { lock.releaseLock(); } catch (e) { /* already released or expired -- nothing more to do */ } }
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- AI relay (Anthropic Claude API, server-side only) ---------------- */

/** Builds a compact plain-text snapshot of current inventory state for the AI's context. */
function buildInventorySnapshot_() {
  const eq = sheetToObjects_(SHEETS.EQUIPMENT);
  const nc = sheetToObjects_(SHEETS.NONCONSUMABLES);
  const cons = sheetToObjects_(SHEETS.CONSUMABLES);
  const rep = sheetToObjects_(SHEETS.REPAIRS).filter(function (r) { return !r.Completed; });
  const lines = [];
  lines.push("EQUIPMENT / ASSETS (" + eq.length + " items):");
  eq.forEach(function (e) {
    lines.push("- " + e.ID + " | " + e.Name + " | " + e.Manufacturer + " " + e.Model +
      " | status:" + e.Status + " | location:" + e.Location + " | nextPM:" + e.NextPM +
      (e.AssignedPatient ? " | patient:" + e.AssignedPatient : ""));
  });
  lines.push("");
  lines.push("NON-CONSUMABLES (" + nc.length + " items):");
  nc.forEach(function (n) {
    lines.push("- " + n.ID + " | " + n.Name + " | " + n.Manufacturer + " " + n.Model +
      " | status:" + n.Status + " | location:" + n.Location +
      " | sterilization:" + (n.SterilizationStatus || "Clean"));
  });
  lines.push("");
  lines.push("CONSUMABLES (" + cons.length + " items):");
  cons.forEach(function (c) {
    lines.push("- " + c.ID + " | " + c.Name + " | stock:" + c.CurrentStock + "/" + c.MaxStock +
      " (min " + c.MinStock + ") | expiry:" + c.ExpiryDate + " | location:" + c.StoreLocation);
  });
  lines.push("");
  lines.push("OPEN REPAIRS (" + rep.length + "):");
  rep.forEach(function (r) {
    lines.push("- " + r.ItemId + " | fault: " + r.Fault + " | reported:" + r.Reported + " | engineer:" + r.Engineer);
  });
  return lines.join("\n");
}

function getAnthropicKey_() {
  return PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
}

/** Answers a free-text inventory question using Claude, grounded in the live sheet data. */
function aiQuery_(question) {
  const key = getAnthropicKey_();
  if (!key) return { ok: false, error: "No ANTHROPIC_API_KEY configured in Script Properties. See setup guide." };
  if (!question) return { ok: false, error: "No question provided." };
  const snapshot = buildInventorySnapshot_();
  const prompt = "You are the inventory assistant for a hospital NICU inventory system. " +
    "Answer the user's question using ONLY the inventory data below. Be concise (1-3 sentences), " +
    "and mention specific item IDs, names, or locations where relevant. If the data doesn't answer " +
    "the question, say so plainly.\n\n" + snapshot + "\n\nQUESTION: " + question;
  try {
    const res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      contentType: "application/json",
      headers: { "x-api-key": key, "anthropic-version": "2023-06-01" },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: 300,
        messages: [{ role: "user", content: prompt }],
      }),
    });
    const json = JSON.parse(res.getContentText());
    if (json.error) return { ok: false, error: json.error.message || "Anthropic API error" };
    const text = (json.content && json.content[0] && json.content[0].text) || "";
    return { ok: true, answer: text.trim() };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/** Identifies equipment (manufacturer/model) from a photo using Claude's vision capability. */
function aiPhoto_(imageBase64) {
  const key = getAnthropicKey_();
  if (!key) return { ok: false, error: "No ANTHROPIC_API_KEY configured in Script Properties. See setup guide." };
  if (!imageBase64) return { ok: false, error: "No image provided." };
  const m = imageBase64.match(/^data:(image\/[a-z]+);base64,(.*)$/);
  const mediaType = m ? m[1] : "image/jpeg";
  const data = m ? m[2] : imageBase64;
  const prompt = "This is a photo of a piece of NICU medical equipment (e.g. ventilator, incubator, " +
    "monitor, infusion pump). Identify the manufacturer and model as precisely as you can from what's " +
    "visible (branding, screen, shape). Respond ONLY with strict JSON: " +
    "{\"manufacturer\":\"...\",\"model\":\"...\",\"confidence\":\"high|medium|low\"}. " +
    "If you cannot identify it, respond {\"manufacturer\":\"\",\"model\":\"\",\"confidence\":\"low\"}.";
  try {
    const res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      contentType: "application/json",
      headers: { "x-api-key": key, "anthropic-version": "2023-06-01" },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: 200,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: mediaType, data: data } },
            { type: "text", text: prompt },
          ],
        }],
      }),
    });
    const json = JSON.parse(res.getContentText());
    if (json.error) return { ok: false, error: json.error.message || "Anthropic API error" };
    const text = (json.content && json.content[0] && json.content[0].text) || "{}";
    let parsed;
    try { parsed = JSON.parse(text); } catch (e) {
      const jm = text.match(/\{[\s\S]*\}/);
      parsed = jm ? JSON.parse(jm[0]) : {};
    }
    if (!parsed.manufacturer) return { ok: false, error: "Couldn't confidently identify this equipment from the photo." };
    return { ok: true, manufacturer: parsed.manufacturer, model: parsed.model || "", confidence: parsed.confidence || "" };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/**
 * BULK IMPORT — HARTA MODAL NICU SCN 2026 register (v2)
 * ----------------------------------------------------------------
 * Reads each row from the "HartaModalImport" sheet into a plain object keyed
 * by its own header (id, type, category, name, ...) and hands it to
 * appendEquipmentRow_() above — the same function addEquipment and
 * batchAddEquipment already use. That function routes to the Equipment vs
 * NonConsumables tab based on `type`, defaults Status to "Working", and
 * stamps LastModified, so this stays consistent with every other way items
 * get added to the sheet.
 *
 * Setup:
 *   1. Make sure the "HartaModalImport" sheet exists with the CSV pasted in
 *      (headers: id, type, category, name, assetNumber, hospitalAssetId,
 *      serialNumber, manufacturer, model, warranty, purchaseDate, location,
 *      status, lastPM, nextPM, supplier, biomedicalContact, photo,
 *      assignedPatient, ward, custody).
 *   2. Save this file. Deploy -> Manage deployments -> Edit -> New version -> Deploy.
 *   3. Run runBulkImportHartaModalV2() from the function dropdown.
 *   4. Check View -> Execution log for "Imported N of 596 rows".
 *
 * If an earlier attempt already left blank rows in Equipment, clean those out
 * first (sort by ID or PurchaseDate, the blanks will sort to one end — select
 * and delete them) before running this, so you don't end up with blanks and
 * real rows both.
 */
function runBulkImportHartaModalV2() {
  const IMPORT_SHEET_NAME = 'HartaModalImport';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName(IMPORT_SHEET_NAME);
  if (!importSheet) {
    throw new Error('Sheet "' + IMPORT_SHEET_NAME + '" not found. Paste the CSV in first.');
  }

  const values = importSheet.getDataRange().getValues();
  if (values.length < 2) {
    Logger.log('No data rows found in "' + IMPORT_SHEET_NAME + '".');
    return;
  }

  const headers = values[0].map(function (h) { return String(h).trim(); });
  const rows = values.slice(1).filter(function (r) {
    return r.some(function (c) { return c !== '' && c !== null; });
  });

  let added = 0;
  let failed = 0;
  const errors = [];

  rows.forEach(function (row, i) {
    const item = {};
    headers.forEach(function (h, colIdx) {
      let v = row[colIdx];
      if (v instanceof Date) {
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      item[h] = v;
    });

    // appendEquipmentRow_ expects a `type` of "equipment" or "nonconsumable" and
    // routes to the right tab automatically (sheetForType_ above).
    try {
      appendEquipmentRow_(item);
      appendHistory_(item.id, 'Added', 'Bulk import', item.location, 'Registered new item (HARTA MODAL batch import)');
      added++;
    } catch (e) {
      failed++;
      errors.push('Row ' + (i + 2) + ' (id=' + item.id + '): ' + String(e));
    }
  });

  Logger.log('Imported ' + added + ' of ' + rows.length + ' rows. Failed: ' + failed);
  if (errors.length > 0) {
    Logger.log('First few errors:\n' + errors.slice(0, 10).join('\n'));
  }
}

/** DIAGNOSTIC -- run this manually from the Apps Script editor
 *  (select testPhotoSetup in the function dropdown, then click Run) to find
 *  out exactly why photos aren't reaching the sheet/Drive. It checks each
 *  step separately and logs a clear pass/fail for each, then writes the
 *  results to a "PhotoSetupTest" sheet tab too, in case Logger.log's output
 *  is hard to find. Safe to run repeatedly; it cleans up the test file/row
 *  it creates. Delete this function (and the PhotoSetupTest tab) once photo
 *  sync is confirmed working, if you like -- it's not needed for normal use.
 */
function testPhotoSetup() {
  const log = [];
  function say(msg) { log.push(msg); Logger.log(msg); }

  say('--- Photo setup diagnostic ---');

  // 1. Does the Equipment tab have a "Photo" column header?
  let eqSh, eqHasPhoto = false;
  try {
    eqSh = getSheet_(SHEETS.EQUIPMENT);
    const headers = getHeaders_(eqSh);
    eqHasPhoto = headers.indexOf('Photo') !== -1;
    say((eqHasPhoto ? 'PASS' : 'FAIL') + ': Equipment tab ' + (eqHasPhoto ? 'has' : 'is MISSING') + ' a "Photo" column header. Headers found: [' + headers.join(', ') + ']');
  } catch (e) {
    say('FAIL: Could not read Equipment tab at all -- ' + String(e));
  }

  // 2. Does DriveApp actually work (this is the step most likely to be silently
  //    blocked by missing authorization -- running this function manually is
  //    what triggers Google's permission prompt if it hasn't been granted yet).
  let testFile = null;
  try {
    const folderName = 'NICU Inventory Photos';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const blob = Utilities.newBlob('test', 'text/plain', 'photo-setup-test.txt');
    testFile = folder.createFile(blob);
    testFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    say('PASS: Drive access works -- created and shared a test file (id ' + testFile.getId() + ') in the "' + folderName + '" folder.');
  } catch (e) {
    say('FAIL: Drive access did not work -- ' + String(e) + '. This usually means the script needs re-authorization: run this function (testPhotoSetup) directly from this editor via the Run button, NOT via the web app, so Google can show you the permission prompt.');
  }

  // 3. A real 1x1-pixel image through the exact same path setPhoto/addEquipment use.
  if (eqHasPhoto) {
    try {
      const onePixelPngBase64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
      const url = savePhotoToDrive_(onePixelPngBase64, 'PHOTOTEST');
      say('PASS: savePhotoToDrive_ produced a URL: ' + url + ' -- open this link in a browser to confirm it actually shows an image.');
    } catch (e) {
      say('FAIL: savePhotoToDrive_ threw an error -- ' + String(e));
    }
  } else {
    say('SKIPPED: end-to-end upload test, because the Photo column is missing (see step 1). Add the column first, then re-run this.');
  }

  // Clean up the test file from step 2 so it doesn't clutter the Drive folder.
  if (testFile) { try { testFile.setTrashed(true); } catch (e) {} }

  say('--- End of diagnostic -- check the "PhotoSetupTest" sheet tab or Executions log for full results ---');

  try {
    let sh = getSs_().getSheetByName('PhotoSetupTest');
    if (!sh) sh = getSs_().insertSheet('PhotoSetupTest');
    sh.clear();
    sh.appendRow(['Photo setup diagnostic — ' + new Date().toISOString()]);
    log.forEach(line => sh.appendRow([line]));
  } catch (e) { /* Logger.log output above is the fallback if this fails */ }
}
