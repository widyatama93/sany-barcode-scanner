/**
 * SANY Shipping Proof module.
 * Uses a dedicated spreadsheet configured in Script Properties:
 * SANY_SHIPPING_SPREADSHEET_ID
 *
 * Sheet: Bukti Pengiriman
 * Folder: Foto Bukti Pengiriman
 *
 * Add a route in the existing doPost(e):
 * if (["submitShippingBatch","getShippingStats","getShippingInitialData","checkShippingDO"].includes(request.action)) {
 *   result = handleShippingRequest_(request);
 * }
 */

const SHIPPING_CONFIG = {
  SPREADSHEET_PROPERTY: "SANY_SHIPPING_SPREADSHEET_ID",
  SHEET_NAME: "Bukti Pengiriman",
  FOLDER_NAME: "Foto Bukti Pengiriman",
  HEADERS: ["Timestamp", "Nomor Surat Jalan", "Status", "Foto Bukti Pengiriman", "Keterangan"],
  STATUS: "Terkirim",
  MAX_DOS: 200,
  MAX_PHOTOS: 15,
  MAX_PHOTO_BYTES: 350 * 1024
};

/**
 * Run once after setting SANY_SHIPPING_SPREADSHEET_ID in Script Properties.
 * Creates the Shipping tab and Drive folder.
 */
function setupShippingAwal() {
  const ss = getShippingSpreadsheet_();
  let sheet = ss.getSheetByName(SHIPPING_CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHIPPING_CONFIG.SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SHIPPING_CONFIG.HEADERS.length).setValues([SHIPPING_CONFIG.HEADERS]);
  } else {
    const current = sheet.getRange(1, 1, 1, SHIPPING_CONFIG.HEADERS.length).getValues()[0];
    if (current.every(value => value === "")) {
      sheet.getRange(1, 1, 1, SHIPPING_CONFIG.HEADERS.length).setValues([SHIPPING_CONFIG.HEADERS]);
    }
  }

  sheet.getRange(1, 1, 1, SHIPPING_CONFIG.HEADERS.length)
    .setFontWeight("bold")
    .setBackground("#d71920")
    .setFontColor("#ffffff");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, SHIPPING_CONFIG.HEADERS.length);
  getOrCreateShippingFolder_();
  return "Setup Shipping berhasil. Tab Bukti Pengiriman dan folder foto siap.";
}

/**
 * Called by the existing Code.gs doPost router.
 */
function handleShippingRequest_(request) {
  switch (String(request && request.action || "")) {
    case "submitShippingBatch":
      return saveShippingBatch_(request);
    case "getShippingStats":
      return getShippingStats_();
    case "getShippingInitialData":
      return getShippingInitialData_();
    case "checkShippingDO":
      return checkShippingDO_(request);
    default:
      throw new Error("Aksi Shipping tidak dikenal.");
  }
}

function getShippingSpreadsheet_() {
  const id = PropertiesService.getScriptProperties()
    .getProperty(SHIPPING_CONFIG.SPREADSHEET_PROPERTY);
  if (!id) {
    throw new Error("Script Property SANY_SHIPPING_SPREADSHEET_ID belum diisi.");
  }
  return SpreadsheetApp.openById(id);
}

function getShippingSheet_() {
  const ss = getShippingSpreadsheet_();
  const sheet = ss.getSheetByName(SHIPPING_CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error('Tab "' + SHIPPING_CONFIG.SHEET_NAME + '" belum ada. Jalankan setupShippingAwal().');
  }
  return sheet;
}

function getOrCreateShippingFolder_() {
  const folders = DriveApp.getFoldersByName(SHIPPING_CONFIG.FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(SHIPPING_CONFIG.FOLDER_NAME);
}

function normalizeShippingDO_(value) {
  return String(value || "").trim().replace(/\s+/g, "").toUpperCase();
}

function checkShippingDO_(request) {
  const code = normalizeShippingDO_(request.doNumber || request.barcode || "");
  if (!code) throw new Error("Nomor surat jalan kosong atau tidak valid.");
  const sheet = getShippingSheet_();
  const lastRow = sheet.getLastRow();
  const values = lastRow < 2 ? [] : sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
  const rowIndex = values.findIndex(row => normalizeShippingDO_(row[0]) === code);
  return {
    success: true,
    action: "checkShippingDO",
    doNumber: code,
    duplicate: rowIndex >= 0,
    row: rowIndex >= 0 ? rowIndex + 2 : null,
    message: rowIndex >= 0 ? "Nomor sudah tercatat pada Bukti Pengiriman." : "Nomor belum pernah disimpan."
  };
}

function saveShippingBatch_(request) {
  let incoming = Array.isArray(request.dos) ? request.dos : [];
  if (!incoming.length && request.doNumber) incoming = [request.doNumber];

  const normalized = [];
  const seen = {};
  const duplicateInRequest = [];
  incoming.forEach(value => {
    const code = normalizeShippingDO_(value);
    if (!code) return;
    if (seen[code]) {
      duplicateInRequest.push(code);
      return;
    }
    seen[code] = true;
    normalized.push(code);
  });
  if (!normalized.length) throw new Error("Belum ada nomor surat jalan yang valid.");
  if (normalized.length > SHIPPING_CONFIG.MAX_DOS) {
    throw new Error("Maksimal " + SHIPPING_CONFIG.MAX_DOS + " nomor dalam satu batch.");
  }

  const photos = Array.isArray(request.photos) ? request.photos : [];
  if (photos.length > SHIPPING_CONFIG.MAX_PHOTOS) {
    throw new Error("Maksimal " + SHIPPING_CONFIG.MAX_PHOTOS + " foto dalam satu batch.");
  }
  const photoBlobs = photos.map((dataUrl, index) => {
    if (typeof dataUrl !== "string") throw new Error("Data foto ke-" + (index + 1) + " tidak valid.");
    const match = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,([A-Za-z0-9+/=]+)$/);
    if (!match) throw new Error("Format foto ke-" + (index + 1) + " tidak dapat dibaca.");
    const bytes = Utilities.base64Decode(match[2]);
    if (bytes.length > SHIPPING_CONFIG.MAX_PHOTO_BYTES) {
      throw new Error("Foto ke-" + (index + 1) + " melebihi batas 350 KB.");
    }
    return {
      bytes: bytes,
      mimeType: match[1].toLowerCase(),
      extension: match[1].toLowerCase().indexOf("png") >= 0 ? "png" : "jpg"
    };
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const createdFiles = [];
  try {
    const sheet = getShippingSheet_();
    const lastRow = sheet.getLastRow();
    const oldCodes = lastRow < 2 ? [] :
      sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues()
        .map(row => normalizeShippingDO_(row[0]));
    const oldSet = {};
    oldCodes.forEach(code => { if (code) oldSet[code] = true; });

    const accepted = [];
    const duplicatesExisting = [];
    normalized.forEach(code => {
      if (oldSet[code]) duplicatesExisting.push(code);
      else accepted.push(code);
    });

    if (!accepted.length) {
      return {
        success: true,
        action: "submitShippingBatch",
        totalRequested: normalized.length + duplicateInRequest.length,
        totalSaved: 0,
        saved: [],
        duplicates: duplicatesExisting.concat(duplicateInRequest),
        totalPhotos: 0,
        message: "Tidak ada nomor baru yang disimpan."
      };
    }

    const photoUrls = [];
    if (photoBlobs.length) {
      const folder = getOrCreateShippingFolder_();
      const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
      photoBlobs.forEach((photo, index) => {
        const name = "SHIPPING_" + stamp + "_" + (index + 1) + "." + photo.extension;
        const file = folder.createFile(Utilities.newBlob(photo.bytes, photo.mimeType, name));
        createdFiles.push(file);
        photoUrls.push(file.getUrl());
      });
    }

    const timestamp = new Date();
    const photoCell = photoUrls.join("\n");
    const note = String(request.keterangan || "").trim();
    const rows = accepted.map(code => [
      timestamp,
      code,
      SHIPPING_CONFIG.STATUS,
      photoCell,
      note
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, SHIPPING_CONFIG.HEADERS.length)
      .setValues(rows);

    const duplicates = duplicatesExisting.concat(duplicateInRequest);
    return {
      success: true,
      action: "submitShippingBatch",
      totalRequested: normalized.length + duplicateInRequest.length,
      totalSaved: accepted.length,
      saved: accepted,
      duplicates: duplicates,
      totalPhotos: photoUrls.length,
      message: accepted.length + " surat jalan berhasil disimpan" +
        (duplicates.length ? ". " + duplicates.length + " nomor duplikat dilewati." : ".")
    };
  } catch (error) {
    createdFiles.forEach(file => {
      try { file.setTrashed(true); } catch (ignore) {}
    });
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function getShippingInitialData_() {
  const stats = getShippingStats_().stats;
  const sheet = getShippingSheet_();
  const lastRow = sheet.getLastRow();
  const knownDOs = lastRow < 2 ? [] :
    [...new Set(sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues()
      .map(row => normalizeShippingDO_(row[0])).filter(Boolean))];
  return { success: true, action: "getShippingInitialData", knownDOs: knownDOs, stats: stats };
}

function getShippingStats_() {
  const sheet = getShippingSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { success: true, action: "getShippingStats", stats: { total: 0, totalPhotos: 0, today: 0, todayPhotos: 0 } };
  }
  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const allPhotos = {}, todayPhotos = {};
  let total = 0, today = 0;
  rows.forEach(row => {
    if (!normalizeShippingDO_(row[1])) return;
    total++;
    const isToday = row[0] instanceof Date &&
      Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyyMMdd") === todayKey;
    if (isToday) today++;
    String(row[3] || "").split(/\r?\n/).filter(Boolean).forEach(url => {
      allPhotos[url] = true;
      if (isToday) todayPhotos[url] = true;
    });
  });
  return {
    success: true,
    action: "getShippingStats",
    stats: {
      total: total,
      totalPhotos: Object.keys(allPhotos).length,
      today: today,
      todayPhotos: Object.keys(todayPhotos).length
    }
  };
}
