
// Скрипт будет автоматически вставлять текущую дату в столбец
function onEdit(e) {
  var sheetName = "Лист1"; // Название листа
  var timestampColumn = "A"; // Столбец для записи даты/времени
  var watchedRanges = ["B:C", "E2:E", "H3:L"]; // Диапазоны для отслеживания
  var timestampFormat = "dateTime"; // Формат записи ('date', 'time', 'dateTime')

  var sheet = e.source.getSheetByName(sheetName);
  var editedCell = e.range;

  // Проверяем, что изменение произошло на нужном листе
  if (sheet.getName() !== sheetName) {
    return;
  }

  // Проверяем, находится ли измененная ячейка в одном из отслеживаемых диапазонов
  var isInWatchedRange = watchedRanges.some(function(range) {
    var rangeObj = sheet.getRange(range);
    return editedCell.getRow() >= rangeObj.getRow() 
        && editedCell.getRow() <= rangeObj.getLastRow() 
        && editedCell.getColumn() >= rangeObj.getColumn() 
        && editedCell.getColumn() <= rangeObj.getLastColumn();
  });

  if (!isInWatchedRange) {
    return;
  }

  // Определяем, какой формат даты/времени использовать
  var timestamp;
  switch (timestampFormat) {
    case "date":
      timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      break;
    case "time":
      timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");
      break;
    case "dateTime":
      timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      break;
  }

  // Вставляем метку времени в указанный столбец
  sheet.getRange(editedCell.getRow(), getColumnNumber(timestampColumn)).setValue(timestamp);
}

// Функция для преобразования буквенного обозначения столбца в числовое
function getColumnNumber(columnLetter) {
  var column = 0;
  for (var i = 0; i < columnLetter.length; i++) {
    column += (columnLetter.charCodeAt(i) - 64) * Math.pow(26, columnLetter.length - i - 1);
  }
  return column;
}
