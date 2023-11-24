// Инициирует цепочку выполнения
function triggerFunction(){
removePrefixesFromTextInCells("Main","B2:B","Main","D2:D");
}

/**
 * Функция для Удаление символов или слов в начале абзаца * 
 * @param {string} sheetName (название листа с данными)
 * @param {string} dataRangeNotation (обозначение диапазона данных)
 * 
 * @param {string} settingsSheetName (название листа с настройками)
 * @param {string} wordsRangeNotation (обозначение диапазона минус слов в листе настроек) 
 */
function removePrefixesFromTextInCells(sheetName, dataRangeNotation, settingsSheetName, wordsRangeNotation) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = ss.getSheetByName(sheetName);
  var settingsSheet = ss.getSheetByName(settingsSheetName);

  // Получаем массив слов для удаления
  var wordsToExclude = settingsSheet.getRange(wordsRangeNotation).getValues().flat().filter(String);

  // Получаем данные из указанного диапазона
  var dataRange = sheet.getRange(dataRangeNotation);
  var values = dataRange.getValues();
  
  // Функция для удаления слов из начала абзаца
  function removeWordsFromParagraph(paragraph, words) {
    var modifiedParagraph = paragraph;
    words.forEach(function(word) {
      if (modifiedParagraph.startsWith(word)) {
        modifiedParagraph = modifiedParagraph.substring(word.length);
      }
    });
    return modifiedParagraph;
  }

  // Проходим по каждой ячейке диапазона и обрабатываем текст
  var cleanedValues = values.map(function(row) {
    return row.map(function(cell) {
      if (typeof cell === 'string') { // Проверяем, что в ячейке текст
        // Разбиваем текст на абзацы и обрабатываем каждый абзац
        return cell.split('\n').map(function(paragraph) {
          return removeWordsFromParagraph(paragraph, wordsToExclude);
        }).join('\n');
      } else {
        return cell; // Если в ячейке не текст, возвращаем как есть
      }
    });
  });

  // Записываем обработанные данные обратно в диапазон
  dataRange.setValues(cleanedValues);
}
