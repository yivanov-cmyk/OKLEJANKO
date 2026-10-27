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
function saveParameters(parameters, adminToken) {
      requireAdminSession_(adminToken);
    {
    if (!Array.isArray(parameters) || !parameters.length) {
      throw new Error('Brak parametrów do zapisania.');
    }


    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PARAMETERS_SHEET);


    if (!sheet) {
      throw new Error('Brak arkusza ParametersConfig.');
    }


    const existingIds = new Set(
      getParametersConfig().map(p => p.id)
    );
    const receivedIds = new Set();


    const rows = parameters.map(p => {
      const id = String(p.id || '').trim();
      const name = String(p.name || '').trim();
      const type = String(p.type || '').trim();
      const column = String(p.targetColumn || '').trim().toUpperCase();


      if (!/^[A-Za-z0-9_]{1,60}$/.test(id)) {
        throw new Error('Nieprawidłowe ID: ' + id);
      }


      if (!name || name.length > 100) {
        throw new Error('Nieprawidłowa nazwa: ' + id);
      }


      if (!['text', 'number', 'select'].includes(type)) {
        throw new Error('Nieprawidłowy typ: ' + id);
      }


      if (column && !['A', 'F', 'I', 'K', 'L'].includes(column)) {
        throw new Error('Niedozwolona kolumna: ' + column);
      }


      if (receivedIds.has(id)) {
        throw new Error('Duplikat ID: ' + id);
      }
      receivedIds.add(id);


      const options = Array.isArray(p.options)
        ? [...new Set(p.options.map(v => String(v).trim()).filter(Boolean))]
        : [];


      if (type === 'select' && !options.length) {
        throw new Error('Brak opcji dla: ' + id);
      }


      return [
        id,
        name,
        type,
        p.required === true,
        JSON.stringify(type === 'select' ? options : []),
        column,
        p.active === true
      ];
    });


    const missing = [...existingIds].filter(id => !receivedIds.has(id));
    if (missing.length) {
      throw new Error('Brak istniejących parametrów: ' + missing.join(', '));
    }


    const backup = sheet.copyTo(ss);
    backup.setName(
      'ParametersBackup_' +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyyMMdd_HHmmss'
      )
    );
    backup.hideSheet();


    const oldRows = Math.max(sheet.getLastRow() - 1, 0);
    if (oldRows) {
      sheet.getRange(2, 1, oldRows, 7).clearContent();
    }


    sheet.getRange(2, 1, rows.length, 7).setValues(rows);


    return {
      success: true,
      message: 'Zapisano parametry: ' + rows.length
    };
  }
}
