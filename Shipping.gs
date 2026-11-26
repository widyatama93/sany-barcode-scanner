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
  HEADERS: ["Timestamp", "Nomor Surat Jalan", "Status", "Foto Bukti Pengiriman", "Keterangan", "Metode Pengiriman", "Nomor Resi / Air Waybill", "Waktu Ambil Foto"],
  STATUS: "Terkirim",
  MAX_DOS: 200,
  MAX_PHOTOS: 15,
  MAX_PHOTO_BYTES: 350 * 1024
};

/**
 * Ensures required headers exist without moving or deleting existing shipping data.
 * New shipping metadata columns are appended for older sheets without moving existing rows.
 */
function ensureShippingHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SHIPPING_CONFIG.HEADERS.length)
      .setValues([SHIPPING_CONFIG.HEADERS]);
  } else {
    let lastColumn = Math.max(1, sheet.getLastColumn());
    let current = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
      .map(value => String(value || "").trim());
    if (current.every(value => !value)) {
      sheet.getRange(1, 1, 1, SHIPPING_CONFIG.HEADERS.length)
        .setValues([SHIPPING_CONFIG.HEADERS]);
      current = SHIPPING_CONFIG.HEADERS.slice();
    } else {
      SHIPPING_CONFIG.HEADERS.forEach(header => {
        if (current.indexOf(header) < 0) {
          const column = Math.max(sheet.getLastColumn(), current.length) + 1;
          sheet.getRange(1, column).setValue(header);
          current[column - 1] = header;
        }
      });
    }
  }

  const headerWidth = Math.max(sheet.getLastColumn(), SHIPPING_CONFIG.HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, headerWidth).getDisplayValues()[0]
    .map(value => String(value || "").trim());
  const columns = {
    method: headers.indexOf("Metode Pengiriman") + 1,
    airWaybill: headers.indexOf("Nomor Resi / Air Waybill") + 1,
    photoTime: headers.indexOf("Waktu Ambil Foto") + 1
  };
  Object.keys(columns).forEach(key => {
    if (columns[key] < 1) throw new Error("Kolom Shipping gagal disiapkan: " + key);
  });
  return columns;
}

/**
 * Run after setting SANY_SHIPPING_SPREADSHEET_ID in Script Properties.
 * Creates the Shipping tab/folder and adds the method header to existing sheets.
 */
function setupShippingAwal() {
  const ss = getShippingSpreadsheet_();
  let sheet = ss.getSheetByName(SHIPPING_CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHIPPING_CONFIG.SHEET_NAME);

  ensureShippingHeaders_(sheet);
  const headerWidth = Math.max(sheet.getLastColumn(), SHIPPING_CONFIG.HEADERS.length);
  sheet.getRange(1, 1, 1, headerWidth)
    .setFontWeight("bold")
    .setBackground("#d71920")
    .setFontColor("#ffffff");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headerWidth);
  getOrCreateShippingFolder_();
  return "Setup Shipping berhasil. Kolom Metode Pengiriman, Nomor Resi / Air Waybill, dan Waktu Ambil Foto siap digunakan.";
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
  const shippingMethod = String(request.shippingMethod || "").trim().toUpperCase();
  const allowedMethods = ["PCP", "CKB", "TRUCKING", "SELFPICK / GOJEK"];
  if (allowedMethods.indexOf(shippingMethod) < 0) {
    throw new Error("Pilih metode pengiriman PCP, CKB, TRUCKING, atau SELFPICK / GOJEK.");
  }

  const supportsAirWaybill = shippingMethod === "PCP" || shippingMethod === "CKB";
  const rawAirWaybills = Array.isArray(request.airWaybills) ? request.airWaybills : [];
  if (!supportsAirWaybill && rawAirWaybills.some(value => String(value || "").trim())) {
    throw new Error("Nomor Resi / Air Waybill hanya digunakan untuk PCP dan CKB.");
  }
  const airWaybills = [];
  const seenAirWaybills = {};
  rawAirWaybills.forEach(value => {
    const code = normalizeShippingDO_(value);
    if (code && !seenAirWaybills[code]) {
      seenAirWaybills[code] = true;
      airWaybills.push(code);
    }
  });
  if (airWaybills.length > SHIPPING_CONFIG.MAX_DOS) {
    throw new Error("Maksimal " + SHIPPING_CONFIG.MAX_DOS + " nomor Resi / Air Waybill.");
  }
  if (supportsAirWaybill && !airWaybills.length) {
    throw new Error("Nomor Resi / Air Waybill wajib diisi untuk PCP dan CKB. Satu resi dapat digunakan untuk beberapa DO dalam satu batch.");
  }

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
    const columns = ensureShippingHeaders_(sheet);
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
    const photoTimestampValues = shippingMethod === "TRUCKING" || shippingMethod === "SELFPICK / GOJEK"
      ? (Array.isArray(request.photoTimestamps) ? request.photoTimestamps : []).map(value => {
          if (!value) return "";
          const capturedAt = new Date(value);
          if (isNaN(capturedAt.getTime())) return "";
          return Utilities.formatDate(capturedAt, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        }).filter(Boolean)
      : [];
    const photoTimeCell = photoTimestampValues.join("\n");
    const airWaybillCell = supportsAirWaybill ? airWaybills.join("\n") : "";
    const rows = accepted.map(code => [
      timestamp,
      code,
      SHIPPING_CONFIG.STATUS,
      photoCell,
      note
    ]);
    const firstNewRow = sheet.getLastRow() + 1;
    sheet.getRange(firstNewRow, 1, rows.length, 5).setValues(rows);
    sheet.getRange(firstNewRow, columns.method, accepted.length, 1)
      .setValues(accepted.map(() => [shippingMethod]));
    sheet.getRange(firstNewRow, columns.airWaybill, accepted.length, 1)
      .setValues(accepted.map(() => [airWaybillCell]));
    sheet.getRange(firstNewRow, columns.photoTime, accepted.length, 1)
      .setValues(accepted.map(() => [photoTimeCell]));

    const duplicates = duplicatesExisting.concat(duplicateInRequest);
    return {
      success: true,
      action: "submitShippingBatch",
      totalRequested: normalized.length + duplicateInRequest.length,
      totalSaved: accepted.length,
      saved: accepted,
      duplicates: duplicates,
      totalPhotos: photoUrls.length,
      totalAirWaybills: airWaybills.length,
      totalPhotoTimestamps: photoTimestampValues.length,
      shippingMethod: shippingMethod,
      message: accepted.length + " surat jalan berhasil disimpan melalui " + shippingMethod +
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
