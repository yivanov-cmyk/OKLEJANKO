// --- Zmienne globalne ---
const LOGIN_SHEET_NAME = "Login";
const ACTIVE_SESSIONS_SHEET_NAME = "Aktywni";
const LOGS_SHEET_NAME = "Calosc";
const SERVICES_CONFIG_SHEET_NAME = "ServicesConfig";

// --- Funkcje główne ---

function doGet(e) {
  if (e.parameter.panel === 'adminPanel') {
    return HtmlService.createHtmlOutputFromFile('adminPanel')
      .setTitle('Panel Administratora')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Rejestracja Procesu')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getServicesConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SERVICES_CONFIG_SHEET_NAME);
  if (!sheet) {
    Logger.log("Błąd: Nie znaleziono arkusza 'ServicesConfig'.");
    return {};
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
     Logger.log("Arkusz ServicesConfig jest pusty.");
     return {};
  }
  const headers = data.shift();

  const config = {};
  data.forEach(row => {
    const procedure = row[0];
    const workmode = row[1] || "";
    const startParams = row[2] ? row[2].split(",").map(p => p.trim()) : [];
    const stopParams = row[3] ? row[3].split(",").map(p => p.trim()) : [];

    config[procedure] = {
      workmode: workmode,
      startParameters: startParams,
      stopParameters: stopParams
    };
  });
  return config;
}

function getActiveSessions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVE_SESSIONS_SHEET_NAME);
  if (!sheet) {
    throw new Error(`Błąd: Brak arkusza o nazwie "${ACTIVE_SESSIONS_SHEET_NAME}".`);
  }
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  if (data.length === 0) return [];
  
  const sessions = data.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header.trim()] = row[index];
    });
    return obj;
  });
  return sessions;
}

function startProcess(sessionData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVE_SESSIONS_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Błąd: Brak arkusza o nazwie "${ACTIVE_SESSIONS_SHEET_NAME}".`);
    }

    const sessionId = Utilities.getUuid();
    
    // Upewnij się, że ten nowy wiersz pasuje do kolejności kolumn w Twoim arkuszu
    const newRow = [
      sessionId,
      sessionData.login,
      sessionData.loginMask || resolveLogin(sessionData.login) || sessionData.login,
      sessionData.procedure,
      sessionData.start,
      sessionData.sscc || "",
      sessionData.startTimestamp
    ];

    sheet.appendRow(newRow);

    return { success: true, message: "Sesja została pomyślnie dodana.", sessionId: sessionId };
  } catch (e) {
    Logger.log(e.message);
    return { success: false, message: 'Nie udało się dodać sesji: ' + e.message };
  }
}

function saveProcedures(proceduresToSave) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SERVICES_CONFIG_SHEET_NAME);
  if (!sheet) {
    return { success: false, message: "Błąd: Nie znaleziono arkusza 'ServicesConfig'." };
  }

  // Wyczyść wszystkie dane, oprócz nagłówków
  sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();

  const newRows = proceduresToSave.map(proc => [
    proc.name,
    proc.workmode,
    proc.startParameters.join(', '),
    proc.stopParameters.join(', ')
  ]);

  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  return { success: true, message: "Konfiguracja procedur została zapisana pomyślnie." };
}
function finishProcess(data) {
  try {
    saveProcessData(data);
    removeFromActive(data.login, data.procedure);
    return { success: true, message: "Sesja została pomyślnie zakończona i zapisana." };
  } catch (e) {
    Logger.log(e.message);
    return { success: false, message: 'Nie udało się zakończyć sesji: ' + e.message };
  }
}

function saveProcessData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calosc");
  if (!sheet) {
    throw new Error("Brak arkusza o nazwie 'Calosc'.");
  }

  // Odczytaj konfigurację serwisów, aby poznać nazwy parametrów
  const config = getServicesConfig();
  const procedureConfig = config[data.procedure] || {};
  const defaultValues = {
        qty: "1",
        warning_qty: "",
        people_count: "",
        formula1: "",
        formula2: "",
    };

    // Pobierz dynamiczne parametry
    const dynamicParams = {
        sscc: data.sscc || "",
        qty: data.Ilosc || defaultValues.qty,
        warning_qty: data.Ostrzezenie || data.Referencja || defaultValues.warning_qty,
        people_count: data.Ilosc_ludzi || defaultValues.people_count,
        workmode: data.workmode || "", // Pobierz workmode dynamicznie z danych
        formula1: data.formula || defaultValues.formula1,
        formula2: data.formula2 || defaultValues.formula2,
        client: data.klient || data.procedure,
        loginMask: data.loginMask || "",
        login: data.login || "",
    };

    // Sformatuj daty i czasy
    const startDate = new Date(data.start);
    const stopDate = new Date(data.stop);
    const formattedDate = Utilities.formatDate(stopDate, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd");
    const formattedStartTime = Utilities.formatDate(startDate, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "HH:mm:ss");
    const formattedStopTime = Utilities.formatDate(stopDate, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "HH:mm:ss");
    const durationInMinutes = (stopDate.getTime() - startDate.getTime()) / (1000 * 60);
    // Stwórz tablicę, która odzwierciedla Twoje kolumny
    const row = [
        dynamicParams.client,           // Kolumna A: Nazwa klienta / Nazwa procesu
        dynamicParams.loginMask,        // Kolumna B: Login mask
        formattedDate,                  // Kolumna C: Data
        formattedStartTime,             // Kolumna D: Czas start (tylko godzina)
        formattedStopTime,              // Kolumna E: Czas finisz (tylko godzina)
        dynamicParams.qty,              // Kolumna F: Ilosc
        durationInMinutes,            // Kolumna G: Czas poświęcony (float)
        dynamicParams.warning_qty,      // Kolumna H: Ostrzezenie/Ilosc referencji
        dynamicParams.people_count,     // Kolumna I: Ilosc ludzi
        dynamicParams.workmode,         // Kolumna J: Workmode
        dynamicParams.formula1,         // Kolumna K: Formula
        dynamicParams.formula2,         // Kolumna L: Formula2
        dynamicParams.login             // Kolumna M: Login
    ];

    // Dodaj wiersz do arkusza
    sheet.appendRow(row);
}

function removeFromActive(login, procedure) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVE_SESSIONS_SHEET_NAME);
  if (!sheet) {
    throw new Error("Brak arkusza o nazwie 'Aktywni'.");
  }

  const data = sheet.getDataRange().getValues();
  const headerRow = data.shift();

  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() == String(login).toLowerCase() && 
        String(data[i][3]).toLowerCase() == String(procedure).toLowerCase()) {
      rowIndex = i + 2;
      break;
    }
  }

  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex);
  } else {
    throw new Error("Nie znaleziono aktywnej sesji do usunięcia.");
  }
}

function clearAllActiveSessions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVE_SESSIONS_SHEET_NAME);
  if (!sheet) {
    return { success: false, message: `Błąd: Brak arkusza o nazwie "${ACTIVE_SESSIONS_SHEET_NAME}".` };
  }
  
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    return { success: true, message: "Wszystkie aktywne sesje zostały wyczyszczone." };
  }
  return { success: true, message: "Brak aktywnych sesji do wyczyszczenia." };
}

function resolveLogin(login) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOGIN_SHEET_NAME);
  if (!sheet) {
    throw new Error(`Błąd: Brak arkusza o nazwie "${LOGIN_SHEET_NAME}".`);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const loginColIndex = headers.findIndex(h => h.trim().toLowerCase() === "login");
  const maskColIndex = headers.findIndex(h => h.trim().toLowerCase() === "loginmask");
  
  if (loginColIndex === -1 || maskColIndex === -1) {
    throw new Error("Błąd: Brak kolumn 'login' lub 'loginMask' w arkuszu Login.");
  }

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][loginColIndex]).toLowerCase().trim() === String(login).toLowerCase().trim()) {
      return data[i][maskColIndex] || String(login).trim();
    }
  }
  return String(login).trim(); // Zwróć oryginalny login jako maskę, jeśli nie znaleziono
}
function getPredefinedParameters() {
  return ["Ilosc", "Klient", "Ostrzerzenie/ilosc ref", "Ilosc ludzi(kon)", "Komentarz"];
}
function getTopOklejanieToday() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calosc");
  if (!sheet) {
    console.log("Błąd: Brak arkusza 'Calosc'. Nie można pobrać danych Top Oklejanie.");
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  console.log(data,"DATA")
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const oklejanieData = data.filter(row => {
    const workmode = row[9]; // Kolumna z procedurą
    console.log(workmode)
    const timestamp = new Date(row[2]); // Kolumna z czasem startu
    
    
    return workmode === "A_FLO" && timestamp >= today;
  });
  
  const userDurations = {};
  oklejanieData.forEach(row => {
    const user = row[1]; // Kolumna z loginem
    const duration = row[5]; // Kolumna z czasem trwania (w sekundach)
    
    if (userDurations[user]) {
      userDurations[user] += duration;
    } else {
      userDurations[user] = duration;
    }
  });
  
  const sortedUsers = Object.keys(userDurations)
    .map(user => ({ user: user, duration: userDurations[user] }))
    .sort((a, b) => b.duration - a.duration)
    .slice(0, 3); // Zwróć tylko TOP 3
  console.log(sortedUsers)
  return sortedUsers;
}
function getWorkmodeStatsForMonth(workmodeFilter) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calosc");
  if (!sheet) {
    Logger.log("Błąd: Brak arkusza 'Calosc'.");
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  const today = new Date();
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();
  
  const workmodeData = data.filter(row => {
    const workmode = row[9]; // Indeks kolumny z workmode (dostosuj, jeśli jest inna)
    const timestamp = new Date(row[2]); // Kolumna z czasem startu
    
    // Filtruj po workmode i bieżącym miesiącu
    return workmode === workmodeFilter && 
           timestamp.getMonth() === currentMonth &&
           timestamp.getFullYear() === currentYear;
  });
  
  const userItems = {};
  workmodeData.forEach(row => {
    const user = row[1]; // Kolumna z loginem
    const items = parseInt(row[5]) || 0; // Kolumna z ilością sztuk (jeśli nazwa parametru to 'Ilosc_szt')
    
    if (userItems[user]) {
      userItems[user] += items;
    } else {
      userItems[user] = items;
    }
  });
  
  const sortedUsers = Object.keys(userItems)
    .map(user => ({ user: user, qty: userItems[user] }))
    .sort((a, b) => b.qty - a.qty)
    .slice(0, 3); // Zwróć tylko TOP 3
  
  return sortedUsers;
}
