const PROCEDURES_SHEET = 'ProceduresConfig';


function getDefaultProcedures_() {
  const definitions = [
    ['S_START', 'Start', 'S_STR', false, []],
    ['Oklejanie', 'Zalando', 'S_OKL', true, ['Ilosc_szt']],
    ['Oklejanieplus', 'Testery', 'S_OKLK', true, ['Ilosc_szt']],
    ['Briefing', 'Briefing', 'M_BRF', false, []],
    ['Auto_Labeling', 'Auto_Labeling', 'A_LAB', true, ['Ilosc_szt', 'Ilosc_ludzi']],
    ['Kartony', 'Kartony', 'M_KAR', false, ['Ilosc_szt']],
    ['Drukowanie', 'Drukowanie', 'A_DRK', false, []],
    ['Przerwa', 'Przerwa', 'S_PRW', false, []],
    ['Anomalia', 'Anomalia', 'S_ANO', false, []],
    ['Copacking', 'Copacking', 'S_COP', false, ['Ilosc_szt']],
    ['Delegacja', 'Delegacja', 'S_DEL', false, ['Klient']],
    ['Simplex', 'Simplex', 'A_SIM', true, ['Ilosc_szt']],
    ['Process', 'Process', 'S_PRC', false, []],
    ['S_END', 'End', 'S_END', false, []]
  ];


  return definitions.map((row, index) => ({
    id: row[0],
    name: row[1],
    workmode: row[2],
    active: true,
    ssccRequired: row[3],
    startParameters: row[3] ? ['SSCC'] : [],
    stopParameters: row[4],
    order: index + 1
  }));
}


function getProceduresConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PROCEDURES_SHEET);


  if (!sheet || sheet.getLastRow() < 2) {
    return getDefaultProcedures_();
  }


  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8)
    .getValues()
    .filter(row => row[0])
    .map(row => ({
      id: String(row[0]),
      name: String(row[1]),
      workmode: String(row[2]),
      active: row[3] === true || String(row[3]).toLowerCase() === 'true',
      ssccRequired: row[4] === true || String(row[4]).toLowerCase() === 'true',
      startParameters: parseProcedureParameters_(row[5]),
      stopParameters: parseProcedureParameters_(row[6]),
      order: Number(row[7]) || 999
    }))
    .sort((a, b) => a.order - b.order);
}


function parseProcedureParameters_(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value;


  try {
    const parsed = JSON.parse(String(value));
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return String(value).split(',')
      .map(item => item.trim())
      .filter(Boolean);
  }
}

function testProceduresConfig() {
  const procedures = getProceduresConfig();


  console.log('Liczba procedur: ' + procedures.length);
  console.log('Oklejanie: ' + getWorkMode('Oklejanie'));
  console.log('Kartony: ' + getWorkMode('Kartony'));
  console.log('Nieznana: ' + getWorkMode('UNKNOWN'));
}
function initializeProceduresConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();


  if (ss.getSheetByName(PROCEDURES_SHEET)) {
    return 'Arkusz ProceduresConfig już istnieje. Nic nie zmieniono.';
  }


  const sheet = ss.insertSheet(PROCEDURES_SHEET);


  sheet.appendRow([
    'ID', 'Name', 'Workmode', 'Active',
    'SSCCRequired', 'StartParameters',
    'StopParameters', 'Order'
  ]);


  const rows = getDefaultProcedures_().map(p => [
    p.id,
    p.name,
    p.workmode,
    p.active,
    p.ssccRequired,
    JSON.stringify(p.startParameters),
    JSON.stringify(p.stopParameters),
    p.order
  ]);


  sheet.getRange(2, 1, rows.length, 8).setValues(rows);
  sheet.setFrozenRows(1);


  return 'Utworzono konfigurację ' + rows.length + ' procedur.';
}
