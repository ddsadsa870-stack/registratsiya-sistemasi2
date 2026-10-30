
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Science Conference 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// GOOGLE FORM ORQALI RO'YXATDAN O'TGANLARNI OLISH
function getParticipants() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName("Forma javoblari 1");

  if (!sheet) {
    throw new Error("Forma javoblari 1 varag'i topilmadi!");
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const data = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getDisplayValues();

  const participants = [];

  data.forEach(function(row) {

    if (!row[1]) return;

    participants.push({
      name: row[1] || "",
      group: row[2] || "",
      topic: row[3] || "",
      direction: row[4] || ""
    });

  });

  return participants;
}


// YO'NALISHLARNI OLISH
function getDirections() {

  return [
    "Dasturiy injiniring",
    "Axborot xavfsizligi",
    "Sun'iy intellekt"
  ];

}


// SAYT UCHUN MA'LUMOTLARNI TEKSHIRISH
function testConnection() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName("Forma javoblari 1");

  if (!sheet) {
    return "XATOLIK: Forma javoblari 1 topilmadi!";
  }

  return "ALOQA MUVAFFAQIYATLI! Javoblar soni: " +
    Math.max(0, sheet.getLastRow() - 1);
}
