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
function createAdminSession(password) {
  const expected = PropertiesService
    .getScriptProperties()
    .getProperty('ADMIN_PASSWORD');


  if (!expected || String(password) !== expected) {
    throw new Error('Nieprawidłowe hasło administratora.');
  }


  const token = Utilities.getUuid() + Utilities.getUuid();


  CacheService.getScriptCache().put(
    'admin_' + token,
    'authorized',
    1800
  );


  return token;
}


function requireAdminSession_(token) {
  if (typeof token !== 'string' || !token) {
    throw new Error('Brak autoryzacji administratora.');
  }


  const authorized = CacheService
    .getScriptCache()
    .get('admin_' + token);


  if (authorized !== 'authorized') {
    throw new Error('Sesja administratora wygasła.');
  }
}
function saveProcedures(procedures, adminToken) {
  requireAdminSession_(adminToken);


  if (!Array.isArray(procedures) || !procedures.length) {
    throw new Error('Lista procedur jest pusta.');
  }


  const lock = LockService.getScriptLock();
  lock.waitLock(20000);


  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PROCEDURES_SHEET);


    if (!sheet) {
      throw new Error('Brak arkusza ProceduresConfig.');
    }


    const existing = getProceduresConfig();
    const existingIds = new Set(existing.map(p => p.id));
    const receivedIds = new Set();


    const clean = procedures.map((p, index) => {
      const id = String(p.id || '').trim();
      const name = String(p.name || '').trim();
      const workmode = String(p.workmode || '').trim();


      if (!/^[A-Za-z0-9_]{1,60}$/.test(id)) {
        throw new Error('Nieprawidłowe ID: ' + id);
      }


      if (!name || name.length > 100) {
        throw new Error('Nieprawidłowa nazwa: ' + id);
      }


      if (!/^[A-Za-z0-9_]{1,60}$/.test(workmode)) {
        throw new Error('Nieprawidłowy Workmode: ' + id);
      }


      if (receivedIds.has(id)) {
        throw new Error('Duplikat ID: ' + id);
      }
      receivedIds.add(id);


      const parseFields = value => {
        if (!Array.isArray(value)) {
          throw new Error('Nieprawidłowe parametry: ' + id);
        }


        const fields = value.map(v => String(v).trim());


        if (fields.some(v => !/^[A-Za-z0-9_]{1,60}$/.test(v))) {
          throw new Error('Nieprawidłowa nazwa parametru: ' + id);
        }


        return [...new Set(fields)];
      };


      const ssccRequired = p.ssccRequired === true;
      const startParameters = parseFields(p.startParameters || []);
      const stopParameters = parseFields(p.stopParameters || []);


      if (ssccRequired && !startParameters.includes('SSCC')) {
        startParameters.unshift('SSCC');
      }


      return {
        id,
        name,
        workmode,
        active: p.active === true,
        ssccRequired,
        startParameters,
        stopParameters,
        order: index + 1
      };
    });


    // Nie pozwalamy przypadkowo usunąć istniejących ID.
    const missing = [...existingIds].filter(id => !receivedIds.has(id));


    if (missing.length) {
      throw new Error(
        'Brakuje istniejących procedur: ' + missing.join(', ')
      );
    }


    // Kopia konfiguracji przed zapisem.
    const backupName = 'ProceduresBackup_' +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyyMMdd_HHmmss'
      );


    const backup = sheet.copyTo(ss);
    backup.setName(backupName);
    backup.hideSheet();


    const rows = clean.map(p => [
      p.id,
      p.name,
      p.workmode,
      p.active,
      p.ssccRequired,
      JSON.stringify(p.startParameters),
      JSON.stringify(p.stopParameters),
      p.order
    ]);


    // Nadpisujemy tylko konfigurację, nie arkusze procesów.
    const oldRows = Math.max(sheet.getLastRow() - 1, 0);


    if (oldRows > 0) {
      sheet.getRange(2, 1, oldRows, 8).clearContent();
    }


    sheet.getRange(2, 1, rows.length, 8).setValues(rows);
    SpreadsheetApp.flush();


    return {
      success: true,
      message: 'Zapisano ' + rows.length + ' procedur.',
      count: rows.length
    };


  } finally {
    lock.releaseLock();
  }
}
