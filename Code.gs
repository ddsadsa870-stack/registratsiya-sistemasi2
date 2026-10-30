
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Science Conference 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// VARAQLARNI TAYYORLASH
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Qatnashchilar varag'i
  let participantsSheet = ss.getSheetByName("Qatnashchilar");

  if (!participantsSheet) {
    participantsSheet = ss.insertSheet("Qatnashchilar");
    participantsSheet.appendRow([
      "F.I.Sh.",
      "O'quv kursi/Guruh",
      "Ma'ruza mavzusi",
      "Yo'nalishi",
      "E-pochta",
      "Telefon raqami",
      "Qo'shilgan vaqt"
    ]);
  }

  // 2. Yo'nalishlar varag'i
  let directionsSheet = ss.getSheetByName("Yo'nalishlar");

  if (!directionsSheet) {
    directionsSheet = ss.insertSheet("Yo'nalishlar");
    directionsSheet.appendRow(["Yo'nalish nomi"]);
    directionsSheet.appendRow(["Dasturiy injiniring"]);
    directionsSheet.appendRow(["Axborot xavfsizligi"]);
    directionsSheet.appendRow(["Sun'iy intellekt"]);
  }

  // 3. Hakamlar varag'i
  let judgesSheet = ss.getSheetByName("Hakamlar");

  if (!judgesSheet) {
    judgesSheet = ss.insertSheet("Hakamlar");
    judgesSheet.appendRow([
      "F.I.Sh.",
      "Lavozimi",
      "Yo'nalishi",
      "E-pochta",
      "Telefon raqami"
    ]);
  }

  return "Tayyor! Qatnashchilar, Yo'nalishlar va Hakamlar varaqlari tekshirildi.";
}

// SARLAVHALARNI BIR XIL KO'RINISHGA KELTIRISH
function normalizeHeader(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

// USTUNNI SARLAVHA ORQALI TOPISH
function findColumn(headers, keywords) {
  return headers.findIndex(function(header) {
    return keywords.some(function(keyword) {
      return header.includes(keyword);
    });
  });
}

// BITTA VARAQDAN QATNASHCHILARNI O'QISH
function readParticipantsFromSheet(sheet, isFormSheet) {
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(normalizeHeader);

  let nameCol = findColumn(headers, ["fish", "fio", "ismfamiliya"]);
  let groupCol = findColumn(headers, ["oquvkursiguruh", "guruh", "kurs"]);
  let topicCol = findColumn(headers, ["maruzamavzusi", "mavzu"]);
  let directionCol = findColumn(headers, ["yonalishi", "yonalish"]);
  let emailCol = findColumn(headers, ["epochta", "email", "elektronpochta"]);
  let phoneCol = findColumn(headers, ["telefonraqami", "telefon", "phone"]);

  // Eski Google Form jadvalingdagi ustunlar tartibi:
  // A - vaqt, B - F.I.Sh., C - Guruh, D - Mavzu, E - Yo'nalish
  if (isFormSheet) {
    if (nameCol === -1) nameCol = 1;
    if (groupCol === -1) groupCol = 2;
    if (topicCol === -1) topicCol = 3;
    if (directionCol === -1) directionCol = 4;
  } else {
    if (nameCol === -1) nameCol = 0;
    if (groupCol === -1) groupCol = 1;
    if (topicCol === -1) topicCol = 2;
    if (directionCol === -1) directionCol = 3;
    if (emailCol === -1) emailCol = 4;
    if (phoneCol === -1) phoneCol = 5;
  }

  return values.slice(1).map(function(row) {
    return {
      name: row[nameCol] || "",
      group: row[groupCol] || "",
      topic: row[topicCol] || "",
      direction: row[directionCol] || "",
      email: emailCol >= 0 ? row[emailCol] || "" : "",
      phone: phoneCol >= 0 ? row[phoneCol] || "" : ""
    };
  }).filter(function(person) {
    return person.name.trim() !== "";
  });
}

// BARCHA QATNASHCHILARNI OLISH
function getParticipants() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const formSheet = ss.getSheetByName("Forma javoblari 1");
  const participantsSheet = ss.getSheetByName("Qatnashchilar");

  let participants = [];

  // Google Form javoblari
  participants = participants.concat(
    readParticipantsFromSheet(formSheet, true)
  );

  // HTML orqali qo'shilgan qatnashchilar
  participants = participants.concat(
    readParticipantsFromSheet(participantsSheet, false)
  );

  return participants;
}

// YO'NALISHLARNI GOOGLE SHEETS'DAN OLISH
function getDirections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Yo'nalishlar");

  if (!sheet || sheet.getLastRow() < 2) {
    return [
      "Dasturiy injiniring",
      "Axborot xavfsizligi",
      "Sun'iy intellekt"
    ];
  }

  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .flat()
    .filter(String);
}

// HTML FORMADAN YANGI QATNASHCHI QO'SHISH
function addParticipant(data) {
  if (!data) {
    throw new Error("Ma'lumotlar kelmadi!");
  }

  const name = String(data.name || "").trim();
  const group = String(data.group || "").trim();
  const topic = String(data.topic || "").trim();
  const direction = String(data.direction || "").trim();
  const email = String(data.email || "").trim();
  const phone = String(data.phone || "").trim();

  if (!name || !group || !topic || !direction || !email || !phone) {
    throw new Error("Barcha maydonlarni to'ldiring!");
  }

  const directions = getDirections();

  if (!directions.includes(direction)) {
    throw new Error("Yo'nalishlar ro'yxatidan to'g'ri yo'nalishni tanlang!");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Qatnashchilar");

  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName("Qatnashchilar");
  }

  sheet.appendRow([
    name,
    group,
    topic,
    direction,
    email,
    phone,
    new Date()
  ]);

  return "Qatnashchi muvaffaqiyatli qo'shildi!";
}

// ALOQANI TEKSHIRISH
function testConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const formSheet = ss.getSheetByName("Forma javoblari 1");
  const participantsSheet = ss.getSheetByName("Qatnashchilar");

  return "Aloqa muvaffaqiyatli! " +
    "Google Form javoblari: " +
    (formSheet ? Math.max(0, formSheet.getLastRow() - 1) : 0) +
    ". HTML orqali qo'shilganlar: " +
    (participantsSheet ? Math.max(0, participantsSheet.getLastRow() - 1) : 0);
}
