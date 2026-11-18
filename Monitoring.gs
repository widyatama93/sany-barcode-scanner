/**
 * SANY Warehouse Monitoring backend module for Google Apps Script.
 *
 * Install:
 * 1. Add this file to the Apps Script project behind the existing scanner URL.
 * 2. In the existing doPost(e), after parsing payload, dispatch:
 *      if (["getMonitoring", "saveMonitoring"].includes(payload.action)) {
 *        return jsonOutput_(handleMonitoringRequest_(payload));
 *      }
 *    Replace jsonOutput_ with the existing JSON response helper used by getStats/submitBatch.
 * 3. Configure Script Property SANY_MONITORING_EDITORS with comma-separated editor emails.
 *    Reads are available to users of the existing web app. Writes require an email in that list.
 * 4. Create the tab "Monitoring" or call setupMonitoringSheet_() once in the editor.
 * 5. Deploy a new Apps Script version to the same Web App deployment.
 *
 * The stable Monitoring ID is generated once and remains unchanged on updates.
 */
const MONITORING_SHEET_NAME = "Monitoring";
const MONITORING_HEADERS = [
  "Monitoring ID", "Sales Order", "Delivery Order", "Company", "Reason Delay",
  "PIC", "Status", "Created At", "Last Updated"
];

function setupMonitoringSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MONITORING_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(MONITORING_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, MONITORING_HEADERS.length).setValues([MONITORING_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function handleMonitoringRequest_(payload) {
  const action = String(payload && payload.action || "");
  if (action === "getMonitoring") return getMonitoringRecords_();
  if (action === "saveMonitoring") return saveMonitoringRecord_(payload.record || {});
  return { success: false, action: action || "error", message: "Unsupported monitoring action." };
}

function getMonitoringRecords_() {
  const sheet = setupMonitoringSheet_();
  const rows = sheet.getDataRange().getValues();
  const records = [];
  const now = new Date();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue;
    const created = row[7] instanceof Date ? row[7] : new Date(row[7]);
    const age = isNaN(created.getTime()) ? 0 : Math.max(0, Math.floor((now.getTime() - created.getTime()) / 86400000));
    records.push({
      monitoringId: String(row[0]), so: String(row[1] || ""), do: String(row[2] || ""),
      company: String(row[3] || ""), reasonDelay: String(row[4] || ""), pic: String(row[5] || ""),
      status: String(row[6] || "Pending"), agingDays: age,
      isDelayed: String(row[6] || "Pending") !== "Shipped" && age > 7,
      lastUpdated: row[8] instanceof Date ? row[8].toISOString() : String(row[8] || "")
    });
  }
  return { success: true, action: "getMonitoring", records: records, updatedAt: now.toISOString() };
}

function saveMonitoringRecord_(record) {
  const action = "saveMonitoring";
  const email = String(Session.getActiveUser().getEmail() || "").toLowerCase();
  const editors = String(PropertiesService.getScriptProperties().getProperty("SANY_MONITORING_EDITORS") || "")
    .split(",").map(function (v) { return v.trim().toLowerCase(); }).filter(Boolean);
  if (!email || editors.indexOf(email) < 0) {
    return { success: false, action: action, message: "Akun ini tidak memiliki akses edit monitoring." };
  }

  const required = ["so", "do", "company", "reasonDelay", "pic"];
  for (let i = 0; i < required.length; i++) {
    if (!String(record[required[i]] || "").trim()) {
      return { success: false, action: action, message: "Lengkapi semua kolom monitoring." };
    }
  }
  const allowed = ["Pending", "Ready to Ship", "Shipped"];
  const status = String(record.status || "Pending");
  if (allowed.indexOf(status) < 0) {
    return { success: false, action: action, message: "Status monitoring tidak valid." };
  }

  const sheet = setupMonitoringSheet_();
  const values = sheet.getDataRange().getValues();
  const wantedId = String(record.monitoringId || "").trim();
  let rowNumber = 0;
  let createdAt = new Date();
  if (wantedId) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === wantedId) {
        rowNumber = i + 1;
        createdAt = values[i][7] instanceof Date ? values[i][7] : new Date(values[i][7]);
        break;
      }
    }
    if (!rowNumber) return { success: false, action: action, message: "Monitoring ID tidak ditemukan." };
  }
  const id = wantedId || Utilities.getUuid();
  const now = new Date();
  const row = [id, String(record.so).trim(), String(record.do).trim(), String(record.company).trim(),
    String(record.reasonDelay).trim(), String(record.pic).trim(), status, createdAt, now];
  if (rowNumber) sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);

  return {
    success: true, action: action, monitoringId: id, message: "Monitoring tersimpan.",
    record: { monitoringId: id, so: row[1], do: row[2], company: row[3],
      reasonDelay: row[4], pic: row[5], status: status,
      agingDays: Math.max(0, Math.floor((now.getTime() - createdAt.getTime()) / 86400000)),
      lastUpdated: now.toISOString() }
  };
}
