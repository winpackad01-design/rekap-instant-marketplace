const SHEET_NAME = "Rekap";

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService
      .createTextOutput("doPost harus dipanggil dari web app. Untuk test manual, jalankan fungsi testManual().")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  const sheet = getSheet();
  const payload = JSON.parse(e.postData.contents || "{}");
  const record = payload.record || {};
  const created = record.createdAt ? new Date(record.createdAt) : new Date();
  const items = Array.isArray(record.items) ? record.items : [];
  const nonEmptyItems = items.filter(item => Number(item.qty || 0) > 0);
  const rows = nonEmptyItems.length ? nonEmptyItems : [{ service: "-", qty: 0 }];

  rows.forEach(item => {
    sheet.appendRow([
      new Date(),
      payload.eventType || "rekap",
      record.id || "",
      formatDate(created),
      formatTime(created),
      record.marketplace || "",
      item.service || "",
      Number(item.qty || 0),
      Number(record.total || 0),
      record.note || "",
      payload.source || "",
      payload.sentAt || ""
    ]);
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService
    .createTextOutput("Rekap pengiriman siap menerima data.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function testManual() {
  const sampleEvent = {
    postData: {
      contents: JSON.stringify({
        eventType: "test",
        source: "apps-script-manual-test",
        sentAt: new Date().toISOString(),
        record: {
          id: "manual-test-" + Date.now(),
          createdAt: new Date().toISOString(),
          marketplace: "TEST",
          note: "Test manual dari Apps Script",
          items: [
            { service: "Test Layanan", qty: 1 }
          ],
          total: 1
        }
      })
    }
  };

  return doPost(sampleEvent);
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Timestamp Masuk",
      "Tipe",
      "ID Rekap",
      "Tanggal Rekap",
      "Jam Rekap",
      "Marketplace",
      "Layanan",
      "Jumlah",
      "Total Rekap",
      "Catatan",
      "Sumber",
      "Sent At"
    ]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function formatDate(date) {
  return Utilities.formatDate(date, "Asia/Bangkok", "yyyy-MM-dd");
}

function formatTime(date) {
  return Utilities.formatDate(date, "Asia/Bangkok", "HH:mm");
}
