function onOpen() {
  const locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  let LANGUAGE = 'en';
  if (locale.startsWith('uk') || locale.startsWith('ua')) {
    LANGUAGE = 'ua';
  }

  const lang = getLang(LANGUAGE);

  SpreadsheetApp.getUi()
    .createMenu(lang.menu)
    .addItem(lang.menuItem, 'calculateStatistics')
    .addToUi();
}

function getLang(langCode) {
  const langs = {
    ua: {
      menu: 'Статистика',
      menuItem: 'Обчислити статистику',
      alert: 'Будь ласка, створіть листи з іменами "Data" та "Statistic"',
      sheetTitle: 'Статистика по листу',
      metric: 'Метрика',
      value: 'Значення',
      metrics: [
        ['Кількість рядків із даними'],
        ['Кількість порожніх рядків'],
        ['Кількість стовпців'],
        ['Кількість унікальних рядків'],
        ['Кількість дублікатів (рядків)']
      ],
      columnTitle: 'Статистика по колонках',
      columnHeaders: ['Колонка', 'Порожніх клітинок', 'Унікальних значень', 'Середнє (числове)', 'Медіана (числова)'],
      noData: 'Немає даних'
    },
    en: {
      menu: 'Statistics',
      menuItem: 'Calculate statistics',
      alert: 'Please create sheets named "Data" and "Statistic"',
      sheetTitle: 'Sheet Statistics',
      metric: 'Metric',
      value: 'Value',
      metrics: [
        ['Number of rows with data'],
        ['Number of empty rows'],
        ['Number of columns'],
        ['Number of unique rows'],
        ['Number of duplicates (rows)']
      ],
      columnTitle: 'Column Statistics',
      columnHeaders: ['Column', 'Empty cells', 'Unique values', 'Average (numeric)', 'Median (numeric)'],
      noData: 'No data'
    }
  };
  return langs[langCode] || langs.en;
}

function calculateStatistics() {
  const locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  let LANGUAGE = 'en';
  if (locale.startsWith('uk') || locale.startsWith('ua')) {
    LANGUAGE = 'ua';
  }
  const lang = getLang(LANGUAGE);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  const statsSheet = ss.getSheetByName('Statistic');

  if (!dataSheet || !statsSheet) {
    SpreadsheetApp.getUi().alert(lang.alert);
    return;
  }

  const data = dataSheet.getDataRange().getValues();
  if (data.length < 2) {
    statsSheet.clear();
    statsSheet.getRange(2, 1).setValue(lang.noData);
    return;
  }

  const headers = data[0];
  const rows = data.slice(1);

  const emptyRowsCount = rows.filter(row => row.every(cell => cell === '' || cell === null)).length;
  const nonEmptyRows = rows.filter(row => row.some(cell => cell !== '' && cell !== null));
  const totalRows = nonEmptyRows.length;

  const rowStrings = nonEmptyRows.map(row => JSON.stringify(row));
  const uniqueRowSet = new Set(rowStrings);
  const uniqueRowsCount = uniqueRowSet.size;

  const duplicatesCount = totalRows - uniqueRowsCount;
  const colCount = headers.length;

  const emptyCellsPerCol = Array(colCount).fill(0);
  nonEmptyRows.forEach(row => {
    row.forEach((cell, i) => {
      if (cell === '' || cell === null) emptyCellsPerCol[i]++;
    });
  });

  const uniqueValuesPerCol = headers.map((_, i) => {
    const colVals = nonEmptyRows.map(r => r[i]).filter(v => v !== '' && v !== null);
    return new Set(colVals).size;
  });

  function median(values) {
    if(values.length === 0) return '-';
    values.sort((a,b) => a-b);
    const half = Math.floor(values.length / 2);
    if (values.length % 2) return values[half];
    return (values[half - 1] + values[half]) / 2.0;
  }

  const numericStats = headers.map((_, i) => {
    const colVals = nonEmptyRows
      .map(r => r[i])
      .filter(v => typeof v === 'number' && !isNaN(v));
    if (colVals.length === 0) return ['-', '-'];
    const avg = (colVals.reduce((a,b) => a + b, 0) / colVals.length).toFixed(2);
    const med = median(colVals).toFixed(2);
    return [avg, med];
  });

  statsSheet.clear();

  const startRow = 2;
  const startCol = 1;

  // --- Sheet Stats ---
  statsSheet.getRange(startRow, startCol, 1, 2).merge();
  statsSheet.getRange(startRow, startCol)
    .setValue(lang.sheetTitle)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(12);

  statsSheet.getRange(startRow + 1, startCol, 1, 2)
    .setValues([[lang.metric, lang.value]])
    .setFontWeight('bold')
    .setFontSize(12);

  const sheetStatsData = [
    [lang.metrics[0][0], totalRows],
    [lang.metrics[1][0], emptyRowsCount],
    [lang.metrics[2][0], colCount],
    [lang.metrics[3][0], uniqueRowsCount],
    [lang.metrics[4][0], duplicatesCount]
  ];

  statsSheet.getRange(startRow + 2, startCol, sheetStatsData.length, 2).setValues(sheetStatsData).setFontSize(12);

  const sheetStatsLastRow = startRow + 1 + 1 + sheetStatsData.length;

  // --- Column Stats ---
  const colStatsStartRow = sheetStatsLastRow + 2;

  statsSheet.getRange(colStatsStartRow, startCol, 1, 5).merge();
  statsSheet.getRange(colStatsStartRow, startCol)
    .setValue(lang.columnTitle)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(12);

  statsSheet.getRange(colStatsStartRow + 1, startCol, 1, 5)
    .setValues([lang.columnHeaders])
    .setFontWeight('bold')
    .setFontSize(12);

  const colStatsData = headers.map((header, i) => [
    header,
    emptyCellsPerCol[i],
    uniqueValuesPerCol[i],
    numericStats[i][0],
    numericStats[i][1]
  ]);

  statsSheet.getRange(colStatsStartRow + 2, startCol, colStatsData.length, 5).setValues(colStatsData).setFontSize(12);

  const lastRow = colStatsStartRow + 1 + 1 + colStatsData.length;

  // Borders for sheet stats
  statsSheet.getRange(startRow, startCol, sheetStatsLastRow - startRow + 1, 2)
    .setBorder(true, true, true, true, true, true);

  for(let r = startRow + 2; r < sheetStatsLastRow; r++) {
    statsSheet.getRange(r, startCol, 1, 2).setBorder(null, null, true, null, null, null);
  }
  statsSheet.getRange(startRow, startCol, sheetStatsLastRow - startRow + 1, 1)
    .setBorder(null, true, null, null, null, null);

  // Borders for column stats
  statsSheet.getRange(colStatsStartRow, startCol, lastRow - colStatsStartRow + 1, 5)
    .setBorder(true, true, true, true, true, true);

  for(let r = colStatsStartRow + 2; r < lastRow; r++) {
    statsSheet.getRange(r, startCol, 1, 5).setBorder(null, null, true, null, null, null);
  }
  for(let c = startCol; c < startCol + 4; c++) {
    statsSheet.getRange(colStatsStartRow, c, lastRow - colStatsStartRow + 1, 1)
      .setBorder(null, true, null, null, null, null);
  }

  for(let c = startCol; c < startCol + 5; c++) {
    statsSheet.autoResizeColumn(c);
  }
}
