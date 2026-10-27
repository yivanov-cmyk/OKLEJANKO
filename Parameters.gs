const PARAMETERS_SHEET = 'ParametersConfig';


function initializeParametersConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();


  if (ss.getSheetByName(PARAMETERS_SHEET)) {
    return 'Arkusz ParametersConfig już istnieje.';
  }


  const sheet = ss.insertSheet(PARAMETERS_SHEET);


  sheet.appendRow([
    'ParameterID',
    'Name',
    'Type',
    'Required',
    'Options',
    'TargetColumn',
    'Active'
  ]);


  const rows = [
    ['SSCC', 'SSCC / ID Kontenera', 'text', true, '', 'A', true],
    ['Ilosc_szt', 'Ilość sztuk', 'number', true, '', 'F', true],
    ['Ilosc_ludzi', 'Ilość ludzi', 'number', false, '', 'I', true],
    ['Klient', 'Klient', 'select', true,
      JSON.stringify([
        'Lindt Copacking',
        'Mattel',
        'Foodcare',
        'Lindt',
        'Yves Rocher',
        'Kontrola st.',
        'HUB'
      ]),
      '', true
    ],
    ['Komentarz', 'Komentarz', 'select', true,
      JSON.stringify([
        "ZWOZENIE", 
        "PREPARACJA", 
        "SPRZATANIE"]),
        '', true
    ],
    ['Referencja', 'Referencja', 'text', false, '', 'K', true]
  ];


  sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  sheet.setFrozenRows(1);


  return 'Utworzono ParametersConfig: ' + rows.length + ' parametrów.';
}


function getParametersConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(PARAMETERS_SHEET);


  if (!sheet || sheet.getLastRow() < 2) return [];


  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 7)
    .getValues()
    .filter(row => row[0])
    .map(row => {
      let options = [];


      try {
        options = JSON.parse(String(row[4] || '[]'));
      } catch (e) {
        options = [];
      }


      return {
        id: String(row[0]),
        name: String(row[1]),
        type: String(row[2]),
        required: row[3] === true,
        options: Array.isArray(options) ? options : [],
        targetColumn: String(row[5] || ''),
        active: row[6] === true
      };
    });
}
