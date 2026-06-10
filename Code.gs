// ╔══════════════════════════════════════════════════════════════════════════╗
// ║        🚀 TRENDHERO PROCESSOR — СИСТЕМА ОБРАБОТКИ ДАННЫХ               ║
// ║                                                                          ║
// ║  СТРУКТУРА ТАБЛИЦЫ:                                                      ║
// ║  • "⚙️ Настройки"      — все параметры здесь                            ║
// ║  • "📥 Входные данные"  — вставляй сырые данные с TrendHero             ║
// ║  • "📊 Результат"       — готовая таблица после обработки               ║
// ║                                                                          ║
// ║  КАК ЗАПУСТИТЬ ПЕРВЫЙ РАЗ:                                              ║
// ║  1. Расширения → Apps Script → вставь этот файл                         ║
// ║  2. Запусти функцию: setupSystem()                                       ║
// ║  3. Появится меню "🚀 TrendHero" и листы с настройками                  ║
// ╚══════════════════════════════════════════════════════════════════════════╝


// ══════════════════════════════════════════════════════════════════════════
// 📌 КОНСТАНТЫ — НАЗВАНИЯ ЛИСТОВ (не меняй если нет причины)
// ══════════════════════════════════════════════════════════════════════════

var SHEET_INPUT    = '📥 Входные данные';
var SHEET_RESULT   = '📊 Результат';
var SHEET_SETTINGS = '⚙️ Настройки';


// ══════════════════════════════════════════════════════════════════════════
// 🔧 ПЕРВОНАЧАЛЬНАЯ УСТАНОВКА — запусти один раз!
// ══════════════════════════════════════════════════════════════════════════

function setupSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  _createSettingsSheet(ss);
  _createInputSheet(ss);
  _createResultSheet(ss);

  ss.setActiveSheet(ss.getSheetByName(SHEET_SETTINGS));

  SpreadsheetApp.getUi().alert(
    '✅ Система установлена!\n\n' +
    '1. Открой лист "⚙️ Настройки" и укажи нужные ГЕО\n' +
    '2. Вставь данные с TrendHero в лист "📥 Входные данные"\n' +
    '3. Меню 🚀 TrendHero → ▶ Запустить обработку'
  );
}


// ══════════════════════════════════════════════════════════════════════════
// 📋 МЕНЮ В ТАБЛИЦЕ
// ══════════════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 TrendHero')
    .addItem('▶ Запустить обработку', 'runProcessing')
    .addItem('↕️ Contacts наверх (пересортировать)', 'resortByContacts')
    .addSeparator()
    .addItem('🧹 Очистить результаты', 'clearResults')
    .addItem('📋 Очистить входные данные', 'clearInput')
    .addToUi();
}


// ══════════════════════════════════════════════════════════════════════════
// 🚀 ГЛАВНАЯ ФУНКЦИЯ ОБРАБОТКИ
// ══════════════════════════════════════════════════════════════════════════

function runProcessing() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var cfg = _loadSettings(ss);
  if (!cfg) return;

  var inputSheet = ss.getSheetByName(SHEET_INPUT);
  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('❌ Лист "' + SHEET_INPUT + '" не найден!\nЗапусти setupSystem().');
    return;
  }

  var lastRow = inputSheet.getLastRow();
  var lastCol = inputSheet.getLastColumn();

  if (lastRow < 4) {
    SpreadsheetApp.getUi().alert('⚠️ Входные данные пустые!\nВставь данные с TrendHero в лист "' + SHEET_INPUT + '" начиная со строки 4');
    return;
  }

  // Строка 3 — заголовки, данные с 4й строки
  var headers = inputSheet.getRange(3, 1, 1, lastCol).getValues()[0];
  var data    = inputSheet.getRange(4, 1, lastRow - 3, lastCol).getValues();

  console.log('📊 Загружено строк: ' + data.length);

  // ── ШАГ 1: Фильтр по ГЕО ──
  var geoColIdx = _findColumn(headers, cfg.geoColumn);
  if (geoColIdx === -1) {
    SpreadsheetApp.getUi().alert(
      '❌ Колонка с ГЕО не найдена!\n\n' +
      'В настройках указано: "' + cfg.geoColumn + '"\n' +
      'Доступные колонки: ' + headers.filter(Boolean).join(', ')
    );
    return;
  }

  var afterGeo = _filterByGeo(data, geoColIdx, cfg.allowedGeo);
  console.log('🌍 После фильтра ГЕО: ' + afterGeo.length + ' строк (было ' + data.length + ')');

  if (afterGeo.length === 0) {
    SpreadsheetApp.getUi().alert(
      '⚠️ После фильтра по ГЕО не осталось ни одной строки.\n\n' +
      '🔎 Ты указал ГЕО: ' + (cfg.allowedGeo.length > 0 ? cfg.allowedGeo.join(', ') : '(пусто — фильтр отключён)') + '\n' +
      '📊 Строк во входных данных было: ' + data.length + '\n\n' +
      'Проверь:\n' +
      '• Правильно ли написан код ГЕО (TH, UA, US — заглавными буквами)\n' +
      '• Название колонки ГЕО в настройках: "' + cfg.geoColumn + '" — совпадает с данными?\n' +
      '• Есть ли вообще такие страны во входных данных'
    );
    return;
  }

  // ── ШАГ 2: Формируем данные: нужный порядок колонок + извлечение контактов ──
  var bioColIdx  = _findColumn(headers, cfg.bioColumn);
  var linkColIdx = _findColumn(headers, cfg.linkColumn);

  if (bioColIdx === -1) {
    SpreadsheetApp.getUi().alert('❌ Колонка Bio не найдена!\n\nУказано: "' + cfg.bioColumn + '"\nДоступные: ' + headers.filter(Boolean).join(', '));
    return;
  }

  var workData = _buildResultData(afterGeo, headers, bioColIdx, linkColIdx, cfg.extractContacts);
  console.log('📱 Данные сформированы: ' + workData.length + ' строк, колонки: ' + RESULT_COLUMNS.join(', '));

  // ── ШАГ 3: Сортировка — строки с contacts наверх ──
  workData = _sortByContacts(workData, 1);
  console.log('↕️ Сортировка по contacts выполнена');

  // ── Записываем результат ──
  _writeResults(ss, RESULT_COLUMNS, workData);

  var removed = data.length - workData.length;
  SpreadsheetApp.getUi().alert(
    '✅ ОБРАБОТКА ЗАВЕРШЕНА!\n\n' +
    '📊 Было строк: ' + data.length + '\n' +
    '✅ Осталось строк: ' + workData.length + '\n' +
    '🗑️ Удалено (не то ГЕО): ' + removed + '\n\n' +
    'Результат → лист "' + SHEET_RESULT + '"'
  );
}


// ══════════════════════════════════════════════════════════════════════════
// ⚙️ ЗАГРУЗКА НАСТРОЕК
// ══════════════════════════════════════════════════════════════════════════

function _loadSettings(ss) {
  var s = ss.getSheetByName(SHEET_SETTINGS);
  if (!s) {
    SpreadsheetApp.getUi().alert('❌ Лист настроек не найден! Запусти setupSystem()');
    return null;
  }

  var vals = s.getRange('B4:B8').getValues();

  var allowedRaw = String(vals[0][0] || '').toUpperCase();
  var allowedGeo = allowedRaw.split(/[,;\s]+/).map(function(x){ return x.trim(); }).filter(Boolean);

  var geoColumn  = String(vals[1][0] || 'country').trim();
  var bioColumn  = String(vals[2][0] || 'biography').trim();
  var linkColumn = String(vals[3][0] || 'ig_link').trim();

  var rawCheck        = vals[4][0];
  var extractContacts = (rawCheck === true || String(rawCheck).toUpperCase() === 'TRUE' || String(rawCheck).toUpperCase() === 'ИСТИНА');

  return {
    allowedGeo:      allowedGeo,
    geoColumn:       geoColumn,
    bioColumn:       bioColumn,
    linkColumn:      linkColumn,
    extractContacts: extractContacts,
  };
}


// ══════════════════════════════════════════════════════════════════════════
// 🌍 ФИЛЬТР ПО ГЕО — только белый список
// ══════════════════════════════════════════════════════════════════════════

function _filterByGeo(data, geoColIdx, allowedGeo) {
  if (allowedGeo.length === 0) return data;
  return data.filter(function(row) {
    var geo = String(row[geoColIdx] || '').toUpperCase().trim();
    return allowedGeo.indexOf(geo) !== -1;
  });
}


// ══════════════════════════════════════════════════════════════════════════
// 📱 ИЗВЛЕЧЕНИЕ КОНТАКТОВ ИЗ BIO + ФОРМИРОВАНИЕ НУЖНОГО ПОРЯДКА КОЛОНОК
// ══════════════════════════════════════════════════════════════════════════

var RESULT_COLUMNS = ['username', 'contacts', 'biography'];

function _buildResultData(data, headers, bioColIdx, linkColIdx, extractContacts) {
  var usernameIdx = _findColumn(headers, 'username');

  var newData = data.map(function(row) {
    var bio    = String(row[bioColIdx] || '');
    var igLink = linkColIdx !== -1 ? String(row[linkColIdx] || '').trim() : '';

    var urlMatch = igLink.match(/https?:\/\/[^\s\)]+/);
    if (urlMatch) igLink = urlMatch[0];
    igLink = igLink.replace(/^[_\s]+|[_\s]+$/g, '').replace(/#.*$/, '').replace(/\/$/, '');

    var contacts = '';
    if (extractContacts) {
      var found = [];
      var seen  = {};

      // Email
      var emails = bio.match(/[a-zA-Z0-9._+\-]+@[a-zA-Z0-9._\-]+\.[a-zA-Z]{2,}/gi);
      if (emails) emails.forEach(function(e) {
        var em = e.toLowerCase();
        if (!seen[em]) { seen[em] = true; found.push(em); }
      });

      // Телефоны с кодом страны
      var intlPhones = bio.match(/\+[\d\s\-\.\(\)]{6,20}/g);
      if (intlPhones) intlPhones.forEach(function(p) {
        var digits = p.replace(/\D/g, '');
        if (digits.length >= 7 && digits.length <= 15) {
          var key = '+' + digits;
          if (!seen[key]) { seen[key] = true; found.push(key); }
        }
      });

      // Локальные номера (длинные цифровые строки)
      var bioClean  = bio.replace(/\+[\d\s\-\.\(\)]{6,20}/g, ' ');
      var localNums = bioClean.match(/\d{7,15}/g);
      if (localNums) localNums.forEach(function(p) {
        var digits = p.replace(/\D/g, '');
        if (!seen[digits] && !seen['+' + digits]) {
          seen[digits] = true; found.push(digits);
        }
      });

      // Номера через дефис: 0xx-xxx-xxxx
      var dashedNums = bio.match(/0[\d][\d\-]{6,12}\d/g);
      if (dashedNums) dashedNums.forEach(function(p) {
        var digits = p.replace(/\D/g, '');
        if (digits.length >= 8 && digits.length <= 12 && !seen[digits] && !seen['+' + digits]) {
          seen[digits] = true;
          found.push(p.trim());
        }
      });

      // Номера через пробел: 0xx xxx xxxx
      var spacedNums = bio.match(/0\d[\d ]{7,13}\d/g);
      if (spacedNums) spacedNums.forEach(function(p) {
        var digits = p.replace(/\D/g, '');
        if (digits.length >= 8 && digits.length <= 12 && !seen[digits] && !seen['+' + digits]) {
          seen[digits] = true;
          found.push(p.trim());
        }
      });

      contacts = found.join('\n');
    }

    return [
      usernameIdx !== -1 ? row[usernameIdx] : '',
      contacts,
      bio,
      igLink,  // 4й элемент — служебный, в результат не идёт, используется для гиперссылки
    ];
  });

  return newData;
}


// ══════════════════════════════════════════════════════════════════════════
// ↕️ СОРТИРОВКА — строки с contacts наверх, пустые вниз
// ══════════════════════════════════════════════════════════════════════════

function _sortByContacts(data, contactsIdx) {
  var withContacts    = [];
  var withoutContacts = [];
  data.forEach(function(row) {
    if (String(row[contactsIdx] || '').trim()) withContacts.push(row);
    else                                        withoutContacts.push(row);
  });
  console.log('↕️ С контактами: ' + withContacts.length + ' | Без: ' + withoutContacts.length);
  return withContacts.concat(withoutContacts);
}


// ══════════════════════════════════════════════════════════════════════════
// 📝 ЗАПИСЬ РЕЗУЛЬТАТА НА ЛИСТ
// ══════════════════════════════════════════════════════════════════════════

function _writeResults(ss, headers, data) {
  var resultSheet = ss.getSheetByName(SHEET_RESULT);
  if (!resultSheet) {
    resultSheet = ss.insertSheet(SHEET_RESULT);
  } else {
    resultSheet.clearContents();
    resultSheet.clearFormats();
  }

  if (data.length === 0) {
    resultSheet.getRange(1, 1).setValue('⚠️ Нет данных после фильтрации — проверь ГЕО в настройках');
    ss.setActiveSheet(resultSheet);
    return;
  }

  var displayHeaders = RESULT_COLUMNS;
  var totalCols      = displayHeaders.length;
  var totalRows      = data.length;
  var igLinkCol      = 3;  // индекс в массиве data (0-based)

  // Шапка
  resultSheet.getRange(1, 1, 1, totalCols)
             .setValues([displayHeaders])
             .setFontWeight('bold')
             .setBackground('#1a73e8')
             .setFontColor('#ffffff')
             .setFontFamily('Arial')
             .setFontSize(10)
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle');
  resultSheet.setRowHeight(1, 28);

  // Данные
  var displayData = data.map(function(row) { return row.slice(0, totalCols); });
  resultSheet.getRange(2, 1, totalRows, totalCols)
             .setValues(displayData)
             .setFontFamily('Arial')
             .setFontSize(10)
             .setVerticalAlignment('top');

  // Username → кликабельная ссылка
  var richTextValues = [];
  for (var r = 0; r < totalRows; r++) {
    var username = String(data[r][0] || '').trim();
    var igLink   = String(data[r][igLinkCol] || '').trim();
    var rtv;
    if (username && igLink) {
      rtv = SpreadsheetApp.newRichTextValue()
            .setText(username)
            .setLinkUrl(igLink)
            .build();
    } else {
      rtv = SpreadsheetApp.newRichTextValue()
            .setText(username)
            .build();
    }
    richTextValues.push([rtv]);
  }
  resultSheet.getRange(2, 1, totalRows, 1).setRichTextValues(richTextValues);

  // Подсветка строк с контактами
  var contactsColIdx = 1;
  var bgColors = [];
  for (var r = 0; r < totalRows; r++) {
    var hasContact = String(data[r][contactsColIdx] || '').trim() !== '';
    var bg = hasContact
      ? (r % 2 === 0 ? '#f1f8e9' : '#e8f5e9')
      : (r % 2 === 0 ? '#ffffff' : '#f8f9fa');
    var rowColors = [];
    for (var c = 0; c < totalCols; c++) rowColors.push(bg);
    bgColors.push(rowColors);
  }
  resultSheet.getRange(2, 1, totalRows, totalCols).setBackgrounds(bgColors);

  // Форматирование contacts
  resultSheet.getRange(2, contactsColIdx + 1, totalRows, 1)
             .setWrap(true)
             .setVerticalAlignment('top');
  resultSheet.setColumnWidth(contactsColIdx + 1, 220);

  // Авторазмер остальных колонок
  for (var c = 1; c <= totalCols; c++) {
    if (c === contactsColIdx + 1) continue;
    resultSheet.autoResizeColumn(c);
  }

  resultSheet.setFrozenRows(1);

  // Итого
  var summaryRow = totalRows + 2;
  resultSheet.getRange(summaryRow, 1, 1, totalCols).merge();
  resultSheet.getRange(summaryRow, 1)
             .setValue('Итого строк: ' + totalRows)
             .setFontStyle('italic')
             .setFontColor('#666666')
             .setFontFamily('Arial')
             .setFontSize(9);

  ss.setActiveSheet(resultSheet);
  console.log('✅ Записано: ' + totalRows + ' строк');
}


// ══════════════════════════════════════════════════════════════════════════
// 🔍 УТИЛИТЫ
// ══════════════════════════════════════════════════════════════════════════

function _findColumn(headers, name) {
  var n = name.toLowerCase().trim();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().trim() === n) return i;
  }
  return -1;
}

function clearResults() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName(SHEET_RESULT);
  if (s) {
    var lastRow = s.getLastRow();
    if (lastRow > 1) {
      s.getRange(2, 1, lastRow - 1, s.getLastColumn()).clearContent().clearFormat();
    }
    SpreadsheetApp.getUi().alert('🧹 Результаты очищены (заголовок сохранён)');
  } else {
    SpreadsheetApp.getUi().alert('Лист результатов не найден');
  }
}

function clearInput() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName(SHEET_INPUT);
  if (!s) {
    SpreadsheetApp.getUi().alert('Лист входных данных не найден');
    return;
  }
  var lastRow = s.getLastRow();
  if (lastRow > 3) {
    var rowsToDelete = lastRow - 3;
    // Google Sheets запрещает удалять ВСЕ незакреплённые строки —
    // если удаляем всё до последней, очищаем содержимое вместо deleteRows
    if (rowsToDelete >= s.getMaxRows() - 3) {
      s.getRange(4, 1, rowsToDelete, s.getLastColumn()).clearContent();
    } else {
      s.deleteRows(4, rowsToDelete);
    }
  }
  SpreadsheetApp.getUi().alert('🧹 Входные данные очищены (заголовки сохранены)');
}


// ══════════════════════════════════════════════════════════════════════════
// ↕️ ПЕРЕСОРТИРОВКА — contacts наверх без полной переобработки
// ══════════════════════════════════════════════════════════════════════════

function resortByContacts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName(SHEET_RESULT);

  if (!s) {
    SpreadsheetApp.getUi().alert('❌ Лист "' + SHEET_RESULT + '" не найден');
    return;
  }

  var lastRow = s.getLastRow();
  var lastCol = s.getLastColumn();

  if (lastRow < 2 || lastCol < 1) {
    SpreadsheetApp.getUi().alert('⚠️ Нет данных для сортировки');
    return;
  }

  var headers     = s.getRange(1, 1, 1, lastCol).getValues()[0];
  var contactsCol = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().trim() === 'contacts') {
      contactsCol = i;
      break;
    }
  }

  if (contactsCol === -1) {
    SpreadsheetApp.getUi().alert('❌ Колонка "contacts" не найдена в результатах');
    return;
  }

  var dataRows  = lastRow - 1;
  var data      = s.getRange(2, 1, dataRows, lastCol).getValues();
  var richTexts = s.getRange(2, 1, dataRows, lastCol).getRichTextValues();

  var withContacts    = [];
  var withoutContacts = [];

  for (var r = 0; r < data.length; r++) {
    var hasContact = String(data[r][contactsCol] || '').trim() !== '';
    if (hasContact) withContacts.push(r);
    else            withoutContacts.push(r);
  }

  var newOrder     = withContacts.concat(withoutContacts);
  var newData      = newOrder.map(function(i) { return data[i]; });
  var newRichTexts = newOrder.map(function(i) { return richTexts[i]; });

  var range = s.getRange(2, 1, dataRows, lastCol);
  range.setValues(newData);
  range.setRichTextValues(newRichTexts);

  var bgColors = [];
  for (var r = 0; r < newData.length; r++) {
    var hasContact = String(newData[r][contactsCol] || '').trim() !== '';
    var bg = hasContact
      ? (r % 2 === 0 ? '#f1f8e9' : '#e8f5e9')
      : (r % 2 === 0 ? '#ffffff' : '#f8f9fa');
    var rowColors = [];
    for (var c = 0; c < lastCol; c++) rowColors.push(bg);
    bgColors.push(rowColors);
  }
  range.setBackgrounds(bgColors);

  SpreadsheetApp.getUi().alert(
    '✅ Пересортировано!\n\n' +
    '📞 С контактами (наверху): ' + withContacts.length + '\n' +
    '○  Без контактов (внизу): ' + withoutContacts.length
  );
}


// ══════════════════════════════════════════════════════════════════════════
// 🏗️ СОЗДАНИЕ ЛИСТОВ (вызывается из setupSystem)
// ══════════════════════════════════════════════════════════════════════════

function _createSettingsSheet(ss) {
  var s = ss.getSheetByName(SHEET_SETTINGS);
  if (!s) s = ss.insertSheet(SHEET_SETTINGS);
  else s.clear();

  s.getRange('A1:C1').merge();
  s.getRange('A1')
   .setValue('⚙️  НАСТРОЙКИ ОБРАБОТКИ TRENDHERO')
   .setFontWeight('bold').setFontSize(14).setFontFamily('Arial')
   .setBackground('#1a73e8').setFontColor('#ffffff')
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(1, 44);

  s.getRange('A2:C2').merge();
  s.getRange('A2')
   .setValue('Здесь настраиваются все параметры фильтрации. Заполни колонку «Значение» и запусти обработку.')
   .setFontStyle('italic').setFontSize(10).setFontColor('#455a64').setFontFamily('Arial')
   .setBackground('#e3f2fd').setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(2, 24);

  s.getRange('A3:C3')
   .setValues([['ПАРАМЕТР', 'ЗНАЧЕНИЕ', 'ПОДСКАЗКА / ПРИМЕР']])
   .setFontWeight('bold').setFontSize(10).setFontFamily('Arial')
   .setBackground('#4a90d9').setFontColor('#ffffff')
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(3, 26);

  var settings = [
    ['🌍  Разрешённые ГЕО',         'TH',         'Коды через запятую: TH, UA, KZ — оставить ТОЛЬКО эти. Если пусто — берутся все.'],
    ['📌  Колонка ГЕО',              'country',    'Как называется колонка с кодом страны в данных. По умолчанию: country'],
    ['📝  Колонка Bio',               'biography',  'Как называется колонка с описанием профиля. По умолчанию: biography'],
    ['🔗  Колонка ссылки',            'ig_link',    'Как называется колонка со ссылкой на профиль. По умолчанию: ig_link'],
    ['📱  Извлекать контакты из Bio', true,         'ИСТИНА / ЛОЖЬ — вытаскивать email и телефоны из Bio → колонка «contacts» (перед biography)'],
  ];

  for (var i = 0; i < settings.length; i++) {
    var row = i + 4;
    s.setRowHeight(row, 38);
    var bgA = (i % 2 === 0) ? '#f5f5f5' : '#ffffff';
    var bgB = (i % 2 === 0) ? '#e8f5e9' : '#f0faf0';
    var bdr = (i % 2 === 0) ? '#c8e6c9' : '#dcedc8';

    s.getRange(row, 1)
     .setValue(settings[i][0])
     .setFontWeight('bold').setFontSize(10).setFontFamily('Arial')
     .setBackground(bgA).setVerticalAlignment('middle')
     .setBorder(false, false, true, false, false, false, bdr, SpreadsheetApp.BorderStyle.SOLID);

    var valCell = s.getRange(row, 2);
    if (i === settings.length - 1) {
      valCell.insertCheckboxes();
      valCell.setValue(true);
    } else {
      valCell.setValue(settings[i][1])
             .setFontSize(10).setFontFamily('Arial').setFontColor('#2e7d32');
    }
    valCell.setBackground(bgB).setVerticalAlignment('middle')
           .setBorder(false, false, true, false, false, false, bdr, SpreadsheetApp.BorderStyle.SOLID);

    s.getRange(row, 3)
     .setValue(settings[i][2])
     .setFontStyle('italic').setFontSize(9).setFontColor('#888888').setFontFamily('Arial')
     .setBackground(bgA).setVerticalAlignment('middle').setWrap(true)
     .setBorder(false, false, true, false, false, false, bdr, SpreadsheetApp.BorderStyle.SOLID);
  }

  s.setRowHeight(9, 14);

  s.getRange('A10:C10').merge();
  s.getRange('A10')
   .setValue('📋  КАК ПОЛЬЗОВАТЬСЯ')
   .setFontWeight('bold').setFontSize(12).setFontFamily('Arial')
   .setFontColor('#ffffff').setBackground('#2e7d32')
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(10, 30);

  var instructions = [
    ['1.  Установка',      'Расширения → Apps Script → вставь код → сохрани → запусти функцию setupSystem()'],
    ['2.  Настройка',      'Заполни колонку «Значение» выше. Минимум: укажи нужные ГЕО в строке «Разрешённые ГЕО»'],
    ['3.  Входные данные', 'Перейди на лист «📥 Входные данные» и вставь туда экспорт из TrendHero (с заголовками!)'],
    ['4.  Запуск',         'Меню «🚀 TrendHero» → «▶ Запустить обработку» или запусти функцию runProcessing()'],
    ['5.  Результат',      'Готовая таблица появится на листе «📊 Результат» — отформатированная и отсортированная'],
  ];

  for (var j = 0; j < instructions.length; j++) {
    var iRow = j + 11;
    s.setRowHeight(iRow, 30);
    var iBg  = (j % 2 === 0) ? '#e8f5e9' : '#ffffff';
    var iBdr = '#a5d6a7';

    s.getRange(iRow, 1)
     .setValue(instructions[j][0])
     .setFontWeight('bold').setFontSize(10).setFontColor('#2e7d32').setFontFamily('Arial')
     .setBackground(iBg).setVerticalAlignment('middle')
     .setBorder(true, true, true, true, false, false, iBdr, SpreadsheetApp.BorderStyle.SOLID);

    s.getRange(iRow, 2, 1, 2).merge();
    s.getRange(iRow, 2)
     .setValue(instructions[j][1])
     .setFontSize(10).setFontColor('#455a64').setFontFamily('Arial')
     .setBackground(iBg).setVerticalAlignment('middle').setWrap(true)
     .setBorder(true, true, true, true, false, false, iBdr, SpreadsheetApp.BorderStyle.SOLID);
  }

  s.setRowHeight(16, 12);
  var tips = [
    '💡  ГЕО-коды двухбуквенные ISO: TH = Таиланд, UA = Украина, US = США, ID = Индонезия, PH = Филиппины',
    '💡  Названия колонок должны совпадать с заголовками в данных TrendHero (регистр не важен)',
    '💡  В колонке «contacts» — email и телефоны через перенос строки. Строки с контактами идут НАВЕРХ результата.',
  ];
  for (var k = 0; k < tips.length; k++) {
    var tRow = k + 17;
    s.setRowHeight(tRow, 26);
    s.getRange(tRow, 1, 1, 3).merge();
    s.getRange(tRow, 1)
     .setValue(tips[k])
     .setFontStyle('italic').setFontSize(9).setFontColor('#5d4037').setFontFamily('Arial')
     .setBackground('#fff9c4').setVerticalAlignment('middle').setWrap(true)
     .setBorder(true, true, true, true, false, false, '#f9a825', SpreadsheetApp.BorderStyle.SOLID);
  }

  s.setColumnWidth(1, 230);
  s.setColumnWidth(2, 180);
  s.setColumnWidth(3, 450);
  s.setFrozenRows(3);
}

function _createInputSheet(ss) {
  var s = ss.getSheetByName(SHEET_INPUT);
  if (!s) s = ss.insertSheet(SHEET_INPUT);
  else s.clearContents();

  s.getRange('A1:K1').merge();
  s.getRange('A1')
   .setValue('📥  ВХОДНЫЕ ДАННЫЕ — вставляй сюда экспорт из TrendHero')
   .setFontWeight('bold').setFontSize(13).setFontFamily('Arial')
   .setBackground('#1a73e8').setFontColor('#ffffff').setHorizontalAlignment('center');
  s.setRowHeight(1, 38);

  s.getRange('A2:K2').merge();
  s.getRange('A2')
   .setValue('Строка 3 — заголовки (не трогай!). Вставляй данные начиная со строки 4. Колонки соответствуют стандартному экспорту TrendHero.')
   .setFontStyle('italic').setFontSize(10).setFontColor('#455a64').setFontFamily('Arial')
   .setBackground('#e3f2fd').setHorizontalAlignment('center');
  s.setRowHeight(2, 22);

  var headers = ['username', 'full_name', 'ig_link', 'country', 'city', 'languages', 'biography', 'follower_count', 'following_count', 'media_count', 'general_er'];
  s.getRange(3, 1, 1, headers.length)
   .setValues([headers])
   .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff')
   .setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center');
  s.setRowHeight(3, 26);

  s.setColumnWidth(1, 140);
  s.setColumnWidth(2, 180);
  s.setColumnWidth(3, 300);
  s.setColumnWidth(4, 70);
  s.setColumnWidth(5, 100);
  s.setColumnWidth(6, 110);
  s.setColumnWidth(7, 380);
  s.setColumnWidth(8, 110);
  s.setColumnWidth(9, 120);
  s.setColumnWidth(10, 100);
  s.setColumnWidth(11, 90);
  s.setFrozenRows(3);
}

function _createResultSheet(ss) {
  var s = ss.getSheetByName(SHEET_RESULT);
  if (!s) s = ss.insertSheet(SHEET_RESULT);
  else s.clear();

  s.getRange('A1:D1').merge();
  s.getRange('A1')
   .setValue('👈  Сначала настрой параметры на листе "⚙️ Настройки", потом запусти обработку через меню 🚀 TrendHero')
   .setFontStyle('italic').setFontColor('#888888').setFontFamily('Arial').setFontSize(11)
   .setBackground('#f5f5f5').setHorizontalAlignment('center');
  s.setRowHeight(1, 36);
}
