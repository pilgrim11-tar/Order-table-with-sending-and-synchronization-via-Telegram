///////////////////////////////////////////
// НАЛАШТУВАННЯ
///////////////////////////////////////////

var ORDERS_SHEET_NAME = "Замовлення";
var MEMO_SHEET_NAME = "Документ для підпису";

// Script Properties keys
var PROP_BOT_TOKEN = "TELEGRAM_BOT_TOKEN";
var PROP_MANAGER_CHAT_ID = "TELEGRAM_MANAGER_CHAT_ID";
var PROP_ENGINEER_CHAT_ID = "TELEGRAM_ENGINEER_CHAT_ID";

// Колонки листа "Замовлення"
var COL_ID = 1;       // A
var COL_DATE = 2;     // B
var COL_NAME = 3;     // C
var COL_BRAND = 4;    // D
var COL_CATALOG = 5;  // E
var COL_PHOTO = 6;    // F
var COL_QTY = 7;      // G
var COL_URGENT = 8;   // H
var COL_NOTE = 9;     // I
var COL_SEND = 10;    // J
var COL_STATUS = 11;  // K
var COL_PACKAGE = 12; // L

// Статуси
var STATUS_RECEIVED = "Отримано";
var STATUS_IN_WORK = "В роботі";
var STATUS_ISSUED = "Видано";

///////////////////////////////////////////
// МЕНЮ
///////////////////////////////////////////

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Електроцех")
    .addItem("Надіслати вибрані замовлення", "sendSelectedPackages")
    .addItem("Оновити документ для підпису", "rebuildMemoNow")
    .addToUi();
}

///////////////////////////////////////////
// ГОЛОВНИЙ ТРИГЕР
///////////////////////////////////////////

function onEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();

  if (sheet.getName() !== ORDERS_SHEET_NAME) return;
  if (row === 1) return;

  if (col === COL_STATUS) {
    handleStatusEdit_(sheet, row, e.oldValue, e.value);
    rebuildMemoSheet_();
  }
}

///////////////////////////////////////////
// ПАКЕТНА ВІДПРАВКА
///////////////////////////////////////////

function sendSelectedPackages() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    SpreadsheetApp.getUi().alert("Спробуйте ще раз через кілька секунд.");
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ORDERS_SHEET_NAME);
    if (!sheet) throw new Error('Не знайдено лист "' + ORDERS_SHEET_NAME + '"');

    var selectedRows = getSelectedRows_(sheet);
    if (selectedRows.length === 0) {
      SpreadsheetApp.getUi().alert("Немає вибраних рядків для відправки.");
      return;
    }

    var grouped = groupRowsByDate_(sheet, selectedRows);
    var dateKeys = Object.keys(grouped).sort();
    var packageCount = 0;

    var sentPackageIds = [];

    for (var i = 0; i < dateKeys.length; i++) {
      var dateKey = dateKeys[i];
      var rows = grouped[dateKey];

      prepareRowsForSend_(sheet, rows);

      var packageId = generatePackageId_(dateKey, i + 1);
      var managerMessage = buildManagerPackageMessage_(sheet, rows, packageId, dateKey);
      var engineerMessage = buildEngineerPackageMessage_(sheet, rows, packageId, dateKey);

      if (sendTelegramToManager_(managerMessage)) {
        sendTelegramToEngineer_(engineerMessage);
        markPackageSent_(sheet, rows, packageId);
        sentPackageIds.push(packageId);
        packageCount++;
      }
    }

    rebuildMemoSheet_();

    if (sentPackageIds.length > 0) {
      sendMemoPdfToManager_(sentPackageIds);
    }

    SpreadsheetApp.getUi().alert("Відправлено пакетів: " + packageCount);
  } finally {
    lock.releaseLock();
  }
}

function getSelectedRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, COL_PACKAGE).getValues();
  var rows = [];

  for (var i = 0; i < values.length; i++) {
    var rowNumber = i + 2;
    var row = values[i];
    var sendFlag = row[COL_SEND - 1];
    var packageValue = row[COL_PACKAGE - 1];
    var status = row[COL_STATUS - 1];

    if (sendFlag === true && !String(packageValue || "").trim() && String(status || "").trim() !== STATUS_ISSUED) {
      rows.push(rowNumber);
    }
  }

  return rows;
}

function groupRowsByDate_(sheet, rowNumbers) {
  var grouped = {};

  for (var i = 0; i < rowNumbers.length; i++) {
    var row = rowNumbers[i];
    var dateValue = sheet.getRange(row, COL_DATE).getValue();

    if (!dateValue) {
      dateValue = new Date();
      sheet.getRange(row, COL_DATE).setValue(dateValue);
    }

    var dateKey = formatDateOnly_(dateValue);
    if (!grouped[dateKey]) grouped[dateKey] = [];
    grouped[dateKey].push(row);
  }

  return grouped;
}

function prepareRowsForSend_(sheet, rows) {
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];

    var id = sheet.getRange(row, COL_ID).getValue();
    var dateValue = sheet.getRange(row, COL_DATE).getValue();
    var name = sheet.getRange(row, COL_NAME).getValue();
    var status = sheet.getRange(row, COL_STATUS).getValue();

    if (!name) {
      sheet.getRange(row, COL_SEND).setValue(false);
      throw new Error("Не заповнено поле 'Назва' у рядку " + row);
    }

    if (!id) {
      sheet.getRange(row, COL_ID).setValue(generateOrderId_());
    }

    if (!dateValue) {
      sheet.getRange(row, COL_DATE).setValue(new Date());
    }

    if (!status) {
      sheet.getRange(row, COL_STATUS).setValue(STATUS_RECEIVED);
    }
  }
}

function buildManagerPackageMessage_(sheet, rows, packageId, dateKey) {
  var lines = [];
  lines.push("🆕 Новий пакет замовлень");
  lines.push("Пакет: " + packageId);
  lines.push("Дата: " + dateKey);
  lines.push("Кількість позицій: " + rows.length);
  lines.push("");

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var id = sheet.getRange(row, COL_ID).getValue();
    var name = sheet.getRange(row, COL_NAME).getValue();
    var brand = sheet.getRange(row, COL_BRAND).getValue();
    var catalog = sheet.getRange(row, COL_CATALOG).getValue();
    var qty = sheet.getRange(row, COL_QTY).getValue();
    var urgent = sheet.getRange(row, COL_URGENT).getValue();

    lines.push((i + 1) + ". ID: " + valueOrNA_(id));
    lines.push("Назва: " + valueOrNA_(name));
    lines.push("Фірма: " + valueOrNA_(brand));
    lines.push("Каталожний номер: " + valueOrNA_(catalog));
    lines.push("Кількість: " + valueOrNA_(qty));
    lines.push("Терміновість: " + valueOrNA_(urgent));
    lines.push("");
  }

  return lines.join("\n");
}

function buildEngineerPackageMessage_(sheet, rows, packageId, dateKey) {
  var lines = [];
  lines.push("📤 Замовлення успішно відправлено");
  lines.push("Пакет: " + packageId);
  lines.push("Дата: " + dateKey);
  lines.push("Кількість позицій: " + rows.length);
  lines.push("");

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var id = sheet.getRange(row, COL_ID).getValue();
    var name = sheet.getRange(row, COL_NAME).getValue();
    var brand = sheet.getRange(row, COL_BRAND).getValue();
    var catalog = sheet.getRange(row, COL_CATALOG).getValue();
    var qty = sheet.getRange(row, COL_QTY).getValue();

    lines.push((i + 1) + ". ID: " + valueOrNA_(id));
    lines.push("Назва: " + valueOrNA_(name));
    lines.push("Фірма: " + valueOrNA_(brand));
    lines.push("Каталожний номер: " + valueOrNA_(catalog));
    lines.push("Кількість: " + valueOrNA_(qty));
    lines.push("");
  }

  return lines.join("\n");
}

function markPackageSent_(sheet, rows, packageId) {
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    sheet.getRange(row, COL_PACKAGE).setValue(packageId);
    sheet.getRange(row, COL_SEND).setValue(false);
  }
}

///////////////////////////////////////////
// ЗМІНА СТАТУСУ
///////////////////////////////////////////

function handleStatusEdit_(sheet, row, oldValue, newValue) {
  var oldStatus = oldValue ? String(oldValue).trim() : "";
  var newStatus = newValue ? String(newValue).trim() : "";

  if (!newStatus) return;
  if (oldStatus === newStatus) return;

  var packageId = String(sheet.getRange(row, COL_PACKAGE).getValue() || "").trim();
  if (!packageId) return;

  var id = sheet.getRange(row, COL_ID).getValue();
  var dateValue = sheet.getRange(row, COL_DATE).getValue();
  var name = sheet.getRange(row, COL_NAME).getValue();
  var brand = sheet.getRange(row, COL_BRAND).getValue();
  var catalog = sheet.getRange(row, COL_CATALOG).getValue();
  var qty = sheet.getRange(row, COL_QTY).getValue();

  var statusLabel = "ℹ️ Змінено стан замовлення";
  if (newStatus === STATUS_RECEIVED) statusLabel = "📥 Замовлення отримано";
  if (newStatus === STATUS_IN_WORK) statusLabel = "🛠 Замовлення в роботі";
  if (newStatus === STATUS_ISSUED) statusLabel = "✅ Замовлення видано";

  var message =
    statusLabel + "\n" +
    "Пакет: " + packageId + "\n" +
    "ID: " + valueOrNA_(id) + "\n" +
    "Дата: " + formatDateTime_(dateValue) + "\n" +
    "Назва: " + valueOrNA_(name) + "\n" +
    "Фірма виробник: " + valueOrNA_(brand) + "\n" +
    "Каталожний номер: " + valueOrNA_(catalog) + "\n" +
    "Кількість: " + valueOrNA_(qty) + "\n" +
    "Було: " + (oldStatus || "порожньо") + "\n" +
    "Стало: " + newStatus;

  sendTelegramToEngineer_(message);
}

///////////////////////////////////////////
// ДОКУМЕНТ ДЛЯ ПІДПИСУ
///////////////////////////////////////////

function rebuildMemoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ordersSheet = ss.getSheetByName(ORDERS_SHEET_NAME);
  var memoSheet = ss.getSheetByName(MEMO_SHEET_NAME);

  if (!ordersSheet) throw new Error('Не знайдено лист "' + ORDERS_SHEET_NAME + '"');
  if (!memoSheet) throw new Error('Не знайдено лист "' + MEMO_SHEET_NAME + '"');

  memoSheet.clear();
  memoSheet.clearFormats();
  memoSheet.setColumnWidths(1, 5, 160);

  memoSheet.getRange("D1:E5").merge();
  memoSheet.getRange("D1").setValue(
    "Начальнику технічної служби\n" +
    "ПрАТ «РФНМ»\n" +
    "Олінчуку С.В.\n" +
    "Головному енергетику\n" +
    "Олійнику О.А."
  );
  memoSheet.getRange("D1").setHorizontalAlignment("right").setVerticalAlignment("middle").setWrap(true);

  memoSheet.getRange("A7:E7").merge();
  memoSheet.getRange("A7").setValue("СЛУЖБОВА ЗАПИСКА");
  memoSheet.getRange("A7").setHorizontalAlignment("center").setFontWeight("bold").setFontSize(14);

  memoSheet.getRange("A9:E9").merge();
  memoSheet.getRange("A9").setValue("Для потреб виробництва необхідно придбати:");
  memoSheet.getRange("A9").setWrap(true);

  var tableStartRow = 11;
  memoSheet.getRange(tableStartRow, 1, 1, 5).setValues([
    ["ID", "Дата", "Назва", "Фірма виробник", "Кількість"]
  ]);
  memoSheet.getRange(tableStartRow, 1, 1, 5).setFontWeight("bold").setHorizontalAlignment("center");

  var data = getMemoRows_(ordersSheet);
  if (data.length > 0) {
    memoSheet.getRange(tableStartRow + 1, 1, data.length, 5).setValues(data);
  }

  var lastTableRow = Math.max(tableStartRow + 1, tableStartRow + data.length);
  memoSheet.getRange(tableStartRow, 1, lastTableRow - tableStartRow + 1, 5)
    .setBorder(true, true, true, true, true, true);

  var signRow = lastTableRow + 3;
  memoSheet.getRange("A" + signRow + ":E" + signRow).merge();
  memoSheet.getRange("A" + signRow).setValue(
    "Дата " + formatDateOnly_(new Date()) + "                              Олійник О.А. ___________________"
  );

  memoSheet.getRange("A" + (signRow + 2) + ":E" + (signRow + 2)).merge();
  memoSheet.getRange("A" + (signRow + 2)).setValue(
    "Погоджено ___                              Олінчук С.В. ____________________"
  );

  memoSheet.getDataRange().setVerticalAlignment("middle");
}

function getMemoRows_(ordersSheet) {
  var lastRow = ordersSheet.getLastRow();
  if (lastRow < 2) return [];

  var values = ordersSheet.getRange(2, 1, lastRow - 1, COL_PACKAGE).getValues();
  var result = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var packageId = String(row[COL_PACKAGE - 1] || "").trim();
    var status = String(row[COL_STATUS - 1] || "").trim();

    if (packageId && status !== STATUS_ISSUED) {
      result.push([
        valueOrNA_(row[COL_ID - 1]),
        formatDateOnly_(row[COL_DATE - 1]),
        valueOrNA_(row[COL_NAME - 1]),
        valueOrNA_(row[COL_BRAND - 1]),
        valueOrNA_(row[COL_QTY - 1])
      ]);
    }
  }

  return result;
}

///////////////////////////////////////////
// TELEGRAM
///////////////////////////////////////////

function sendTelegramToManager_(message) {
  return sendTelegram_(PROP_MANAGER_CHAT_ID, message);
}

function sendTelegramToEngineer_(message) {
  return sendTelegram_(PROP_ENGINEER_CHAT_ID, message);
}

function sendTelegram_(chatIdPropKey, message) {
  var props = PropertiesService.getScriptProperties();
  var botToken = props.getProperty(PROP_BOT_TOKEN);
  var chatId = props.getProperty(chatIdPropKey);

  if (!botToken) throw new Error("Не задано TELEGRAM_BOT_TOKEN у Script Properties");
  if (!chatId) throw new Error("Не задано " + chatIdPropKey + " у Script Properties");

  var url = "https://api.telegram.org/bot" + botToken + "/sendMessage";
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      chat_id: String(chatId),
      text: message,
      disable_web_page_preview: false
    }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());
  Logger.log(response.getContentText());

  if (!result.ok) {
    throw new Error("Telegram error: " + result.description);
  }

  return true;
}

function sendMemoPdfToManager_(packageIds) {
  var props = PropertiesService.getScriptProperties();
  var botToken = props.getProperty(PROP_BOT_TOKEN);
  var chatId = props.getProperty(PROP_MANAGER_CHAT_ID);

  if (!botToken) throw new Error("Не задано TELEGRAM_BOT_TOKEN у Script Properties");
  if (!chatId) throw new Error("Не задано " + PROP_MANAGER_CHAT_ID + " у Script Properties");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memoSheet = ss.getSheetByName(MEMO_SHEET_NAME);
  if (!memoSheet) throw new Error('Не знайдено лист "' + MEMO_SHEET_NAME + '"');

  var blob = exportSheetToPdf_(ss, memoSheet, buildMemoPdfFileName_(packageIds));
  var url = "https://api.telegram.org/bot" + botToken + "/sendDocument";
  var caption = "📎 Документ для підпису\nПакети: " + packageIds.join(", ");

  var options = {
    method: "post",
    payload: {
      chat_id: String(chatId),
      caption: caption,
      document: blob
    },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());
  Logger.log(response.getContentText());

  if (!result.ok) {
    throw new Error("Telegram PDF error: " + result.description);
  }

  return true;
}

function exportSheetToPdf_(spreadsheet, sheet, fileName) {
  var exportUrl =
    "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export" +
    "?format=pdf" +
    "&gid=" + sheet.getSheetId() +
    "&size=A4" +
    "&portrait=true" +
    "&fitw=true" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=false" +
    "&top_margin=0.50" +
    "&bottom_margin=0.50" +
    "&left_margin=0.50" +
    "&right_margin=0.50";

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: "Bearer " + token
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("Не вдалося сформувати PDF. Код: " + response.getResponseCode());
  }

  return response.getBlob().setName(fileName);
}

function buildMemoPdfFileName_(packageIds) {
  var suffix = packageIds.length === 1 ? packageIds[0] : packageIds[0] + "_and_more";
  return "document_dlya_pidpysu_" + suffix + ".pdf";
}

///////////////////////////////////////////
// ДОПОМІЖНІ
///////////////////////////////////////////

function generateOrderId_() {
  var now = new Date();
  return "ORD-" + Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
}

function generatePackageId_(dateKey, index) {
  return "PKG-" + dateKey.replace(/\./g, "") + "-" + ("0" + index).slice(-2);
}

function valueOrNA_(value) {
  return value === "" || value === null || value === undefined ? "N/A" : String(value);
}

function formatDateOnly_(value) {
  if (!value) return "N/A";
  return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), "dd.MM.yyyy");
}

function formatDateTime_(value) {
  if (!value) return "N/A";
  return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
}

///////////////////////////////////////////
// СЕРВІСНІ ФУНКЦІЇ
///////////////////////////////////////////

function setupTelegramConfig() {
  PropertiesService.getScriptProperties().setProperties({
    TELEGRAM_BOT_TOKEN: "PASTE_NEW_BOT_TOKEN_HERE",
    TELEGRAM_MANAGER_CHAT_ID: "PASTE_MANAGER_CHAT_ID_HERE",
    TELEGRAM_ENGINEER_CHAT_ID: "PASTE_ENGINEER_CHAT_ID_HERE"
  });
}

function testManagerTelegram() {
  sendTelegramToManager_("Тестове повідомлення менеджеру");
}

function testEngineerTelegram() {
  sendTelegramToEngineer_("Тестове повідомлення інженеру");
}

function rebuildMemoNow() {
  rebuildMemoSheet_();
}
