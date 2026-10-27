const CACHE_KEY_LOGINS = 'LOGIN_MAP_CACHE';
const PROP_KEY_SESSIONS = 'activeSessions';
const PROP_KEY_TOP3 = 'TOP_3_CACHE';
function doGet(e) {
  const page = e && e.parameter && e.parameter.page;


  if (page === 'admin') {
    return HtmlService
      .createHtmlOutputFromFile('adminPanel')
      .setTitle('FM Logistic – Administracja')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }


  const newPassword = generateAndSendNewPasswordToAdmin();


  if (newPassword) {
    logDebug('Nowe hasło wygenerowane i wysłane pomyślnie.');
  } else {
    logDebug('Błąd generowania hasła.');
  }


  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('FM Logistic – Rejestracja procesu')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function logDebug(message) {
  const debug = PropertiesService.getScriptProperties().getProperty('DEBUG_ENABLED') === 'true';
  if (debug) {
    console.log(message);
  }
}
function NAPRAWA_I_UPRAWNIENIA() {
  console.log("--- START NAPRAWY ---");
  
  // Wymuszenie okna zgody na arkusze
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log("Połączono z arkuszem: " + ss.getName());
  
  // Wymuszenie okna zgody na sieć (UrlFetch)
  UrlFetchApp.fetch("https://google.com", {muteHttpExceptions: true});
  console.log("Dostęp do internetu: OK");
  
  // Wyłączenie trybu debugowania, który powodował błąd pętli
  PropertiesService.getScriptProperties().setProperty('DEBUG_ENABLED', 'false');
  console.log("Tryb debugowania: WYŁĄCZONY (dla bezpieczeństwa)");
  
  // Wymuszenie BigQuery (Upewnij się, że dodałeś BigQuery w Services!)
  try {
    BigQuery.Jobs.list('fml-prod-pl-bo');
    console.log("Dostęp do BigQuery: OK");
  } catch(e) {
    console.log("BigQuery API wymaga autoryzacji lub dodania do Services.");
  }
  
  console.log("--- KONIEC. Teraz spróbuj odświeżyć aplikację ---");
}

function sendAnomalyNotification(data) {
  
  const WEBHOOK_URL = 'https://chat.googleapis.com/v1/spaces/AAQA2cf1rDE/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=8_TkBCTsrix5onMJBXW08rm8lWrNokEh59GGO9PB-Gc'; 

  const maskedUser = data.loginMask || data.login;
  const timeFormatted = Utilities.formatDate(new Date(), "GMT+2", "HH:mm:ss");

  
  const payload = {
    "text": "⚠️ *ALERT: ZGŁOSZONO ANOMALIĘ!*\n\n" +
            "👤 *Pracownik:* " + maskedUser + "\n" +
            "⏰ *Czas rozpoczęcia:* " + timeFormatted + "\n" +
            "📌 *Proces:* Anomalia"
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    UrlFetchApp.fetch(WEBHOOK_URL, options);
  } catch (error) {
    Logger.log("Błąd wysyłania wiadomości na czat: " + error.toString());
  }
}
function generateAndSendNewPasswordToAdmin() {
  try {
    const newPassword = generateNewPassword();
    PropertiesService.getScriptProperties().setProperty('accessPassword', newPassword);
    
    // Używamy nowszego adresu webhooka z Twojego kodu
    const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAQA2cf1rDE/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=8_TkBCTsrix5onMJBXW08rm8lWrNokEh59GGO9PB-Gc';
    
    const message = { 'text': `Nowe hasło do aplikacji: ${newPassword}` };
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(message),
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(webhookUrl, options);
    if (response.getResponseCode() === 200) {
      return newPassword;
    }
    return null;
  } catch (e) {
    Logger.log('Błąd w generateAndSendNewPasswordToAdmin: ' + e.toString());
    return null;
  }
}

function generateNewPassword() {
  const length = 8;
  const chars = '0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

function verifyPassword(enteredPassword) {
  const storedPassword = PropertiesService.getScriptProperties().getProperty('accessPassword');
  if (storedPassword === enteredPassword) {
    logDebug('Weryfikacja hasła Pomyślna!');
    return 'OK';
  } else {
    logDebug('Weryfikacja hasła NIEUDANA!');
    return 'Błędne hasło.';
  }
}
function autoCleanupSessions() {
  const lock = LockService.getScriptLock();
  try {
    // Blokujemy dostęp na 5 sekund, żeby bezpiecznie zmodyfikować sesje
    lock.waitLock(5000); 
    
    const props = PropertiesService.getScriptProperties();
    const sessionsData = props.getProperty("activeSessions");
    
    if (!sessionsData) return;
    
    let sessions = JSON.parse(sessionsData);
    const now = Date.now();
    const TWELVE_HOURS = 14 * 60 * 60 * 1000; // 12 godzin w milisekundach

    // Filtrujemy: zostawiamy tylko te, które trwają KRÓCEJ niż 12h
    const initialCount = sessions.length;
    sessions = sessions.filter(s => {
      const startTime = s.startTimestamp || (s.start ? new Date(s.start).getTime() : now);
      return (now - startTime) < TWELVE_HOURS;
    });

    // Jeśli kogoś usunęliśmy, zapisujemy nową listę
    if (sessions.length < initialCount) {
      props.setProperty("activeSessions", JSON.stringify(sessions));
      logDebug(`Auto-cleanup: Usunięto ${initialCount - sessions.length} przedawnionych sesji.`);
    }
    
  } catch (e) {
    logDebug("Błąd podczas auto-cleanupu: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}
function forcePermissions() {
  console.log("--- Start wymuszania uprawnień ---");
  
  // 1. Dostęp do SpreadsheetApp
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log("Dostęp do arkusza: " + ss.getName());
  
  // 2. Dostęp do PropertiesService
  PropertiesService.getScriptProperties().setProperty('AUTH_TEST', 'OK');
  console.log("Dostęp do Properties: OK");
  
  // 3. Dostęp do UrlFetch (Webhook Chat)
  try {
    // Próba pobrania czegokolwiek (nawet Google), by wymusić zgodę na sieć
    UrlFetchApp.fetch("https://www.google.com", {muteHttpExceptions: true});
    console.log("Dostęp do UrlFetch: OK");
  } catch(e) {
    console.log("UrlFetch wymaga zgody...");
  }
} 
function getLoginMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Login");
  if (!sheet) {
    Logger.log("Arkusz 'Login' nie został znaleziony.");
    return {};
  }
  const data = sheet.getRange("A1:B" + sheet.getLastRow()).getValues();
  const map = {};
  data.forEach(row => {
    if (row[0] && row[1]) {
      map[row[0].toString().trim().toLowerCase()] = row[1].toString().trim();
    }
  });
  return map;
}
function getTopOklejanieFromCache() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard_Data");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  // Mapujemy prosto z gotowej tabelki
  return data.map(row => ({ user: row[0], qty: row[1] }));
}
function resolveLogin(login) {
  const map = getLoginMapCached();
  const cleanLogin = login.toString().trim().toLowerCase();
  return map[cleanLogin] || login;
}
function getTopOklejanieToday() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calosc");
  if (!sheet) {
    logDebug("Arkusz 'Calosc' nie został znaleziony.");
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) { 
    logDebug("Brak danych w arkuszu 'Calosc'.");
    return [];
  }

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayFormatted = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const userTotalsToday = {};
  const userTotalsMonth = {};

  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Sprawdzamy czy wiersz ma kolumnę M (indeks 12)
    if (row.length < 13 || !row[2]) continue;

    const recordDate = new Date(row[2]);
    const ilosc = parseFloat(row[5]) || 0;
    if (ilosc === 0) continue;

    const procedureWorkMode = row[9];
    
    // ZMIANA: Zamiast row[1] (login), pobieramy row[12] (Kolumna M - Identyfikator)
    const userIdentifier = row[12] || "Brak ID";

    const dateRecord = Utilities.formatDate(recordDate, tz, "yyyy-MM-dd");
    
    // Logika sumowania - używamy identyfikatora jako klucza
    if (dateRecord === todayFormatted && procedureWorkMode === "S_OKL") {
      userTotalsToday[userIdentifier] = (userTotalsToday[userIdentifier] || 0) + ilosc;
    }

    if (recordDate.getFullYear() === currentYear && recordDate.getMonth() === currentMonth && procedureWorkMode === "S_OKL") {
      userTotalsMonth[userIdentifier] = (userTotalsMonth[userIdentifier] || 0) + ilosc;
    }
  }

  // Budowanie rankingu Top 3
  let top3 = Object.entries(userTotalsToday)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([user, qty]) => ({ user, qty: Math.round(qty * 100) / 100 }));

  if (top3.length === 0) {
    logDebug("Brak danych dziennych. Używam danych miesięcznych.");
    top3 = Object.entries(userTotalsMonth)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([user, qty]) => ({ user, qty: Math.round(qty * 100) / 100 }));
  }

  logDebug("Top 3 Oklejanie (Anonimizowane): " + JSON.stringify(top3));
  return top3;
}
/**
 * Pobiera mapę loginów inteligentnie. 
 * Sprawdza czy zmieniła się liczba wierszy w arkuszu "Login".
 */
function getLoginMapCached() {
  const cache = CacheService.getScriptCache();
  const cachedMap = cache.get("LOGIN_MAP_DATA");

  // 1. Jeśli mamy dane w RAM, oddaj je natychmiast (trwa 0.01s)
  // To sprawi, że resolveLogin i odświeżanie listy będą błyskawiczne
  if (cachedMap) {
    return JSON.parse(cachedMap);
  }

  // 2. Tylko jeśli Cache jest pusty, dotykamy arkusza (trwa 2-5s)
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Login");
  if (!sheet) return {};

  const currentLastRow = sheet.getLastRow().toString();
  
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  const map = {};
  data.forEach(row => {
    if (row[0] && row[1]) {
      map[row[0].toString().trim().toLowerCase()] = row[1].toString().trim();
    }
  });

  // Zapisujemy do Cache na 6h
  try {
    cache.put("LOGIN_MAP_DATA", JSON.stringify(map), 21600);
    props.setProperty("LAST_LOGIN_ROW_COUNT", currentLastRow);
  } catch (e) {
    console.log("Mapa zbyt duża dla Cache - optymalizacja niemożliwa.");
  }

  return map;
}
function onEdit(e) {
  const sheetName = e.source.getActiveSheet().getName();
  if (sheetName === "Login") {
    // Jeśli ktoś edytował arkusz Login, czyścimy cache, żeby przy następnym ruchu pobrało nowe dane
    CacheService.getScriptCache().remove("LOGIN_MAP_DATA");
    console.log("Cache loginów wyczyszczony po edycji arkusza.");
  }
}
// Wklej do Code.gs
function completeProcessOnServer(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Blokada na 10s
    
    // 1. Zapisz do arkusza (wykorzystuje Twoją szybką funkcję saveProcessData)
    saveProcessData(data);
    
    // 2. Usuń z aktywnych sesji
    removeFromActive(data.login, data.procedure);
    
    return true;
  } catch (e) {
    throw new Error("Błąd podczas kończenia procesu: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}
/**
 * Funkcja pomocnicza do sortowania i ograniczania wyników
 */
function sortAndSlice(totalsMap) {
  return Object.entries(totalsMap)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([user, qty]) => ({ 
      user, 
      qty: Math.round(qty * 100) / 100 
    }));
}
function updateTopStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Calosc");
  const targetSheet = ss.getSheetByName("Dashboard_Data") || ss.insertSheet("Dashboard_Data");

  // 1. Wykonaj ciężkie obliczenia (Twoja obecna logika z getTopOklejanieToday)
  const top3 = getTopOklejanieToday(); 

  // 2. Wyczyść stare wyniki i zapisz nowe
  targetSheet.clear();
  if (top3.length > 0) {
    const rows = top3.map(item => [item.user, item.qty]);
    targetSheet.getRange(1, 1, rows.length, 2).setValues(rows);
  }
  
  logDebug("Statystyki Top 3 zostały zaktualizowane: " + new Date());
}
function getUsersAndLoginMap() {
  return {
    users: getActiveSessions(), // To pobieramy zawsze (z PropertiesService - spójne na urządzeniach)
    loginMap: getLoginMapCached() // To pobieramy inteligentnie (tylko przy zmianach)
  };
}

function getActiveSessions() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const data = scriptProperties.getProperty("activeSessions");
  try {
    return data ? JSON.parse(data) : [];
  } catch (e) {
    Logger.log("Błąd podczas parsowania aktywnych sesji: " + e.toString() + ". Zwracam pustą tablicę.");
    return []; // W razie błędu parsowania, zwróć pustą tablicę, aby uniknąć dalszych problemów
  }
}

function setActiveSessions(sessions) {
  try {
    PropertiesService.getScriptProperties().setProperty("activeSessions", JSON.stringify(sessions));
  } catch (e) {
    Logger.log("Błąd podczas zapisywania aktywnych sesji: " + e.toString());
    // Rozważ rzucenie błędu, jeśli zapis jest krytyczny
    // throw new Error("Nie udało się zapisać aktywnych sesji.");
  }
}


function checkIfUserIsActive(identifier) {
  Logger.log("checkIfUserIsActive wywołana z identyfikatorem: " + identifier);
  const sessions = getActiveSessions();
  const isActive = sessions.some(s => 
    (s.loginMask && s.loginMask.toString().trim().toLowerCase() === identifier.toString().trim().toLowerCase()) ||
    (s.login && s.login.toString().trim().toLowerCase() === identifier.toString().trim().toLowerCase())
  );
  
  if (isActive) {
    Logger.log("Użytkownik " + identifier + " JEST aktywny.");
    // Możesz zwrócić obiekt, aby przekazać również wiadomość
    return { isActive: true, message: `Użytkownik ${identifier} ma już aktywną sesję.` };
  } else {
    Logger.log("Użytkownik " + identifier + " NIE JEST aktywny.");
    return { isActive: false }; // Lub po prostu return false;
  }
}


/**
 * Dodaje nową sesję do aktywnych.
 * Ta funkcja zakłada, że sprawdzenie, czy użytkownik jest już aktywny,
 * zostało wykonane po stronie klienta PRZED jej wywołaniem.
 * Jeśli nadal chcesz mieć zabezpieczenie po stronie serwera (zalecane), możesz je tu zostawić.
 */
function addToActive(session) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); // Czekaj max 5s na zwolnienie blokady
    
    const sessions = getActiveSessions();
    const alreadyActive = sessions.some(s => s.login === session.login);

    if (alreadyActive) {
      throw new Error(`Użytkownik ${session.login} jest już aktywny.`);
    }
    if (session.procedure === "Anomalia") {
      sendAnomalyNotification(session);
    }
  
    
    sessions.push(session);
    setActiveSessions(sessions);
  } finally {
    lock.releaseLock();
  }
}

function removeFromActive(login, procedure) {
  const lock = LockService.getScriptLock();
  try {
    // Czekaj do 5 sekund na dostęp, jeśli inny proces aktualnie zapisuje dane
    lock.waitLock(5000); 

    let sessions = getActiveSessions(); // Pobiera aktualną listę
    const initialCount = sessions.length;
    
    // Filtrowanie sesji w pamięci RAM (bardzo szybkie)
    sessions = sessions.filter(s => 
      !(s.login && s.login.toString().trim().toLowerCase() === login.toString().trim().toLowerCase() && 
        s.procedure && s.procedure.toString().trim() === procedure.toString().trim())
    );
    
    if (sessions.length < initialCount) {
      setActiveSessions(sessions); // Zapisuje zaktualizowaną listę
      logDebug(`Usunięto sesję dla: ${login}`);
    }
  } catch (e) {
    logDebug(`Błąd w removeFromActive: ${e.toString()}`);
    throw new Error("Serwer jest zajęty, spróbuj zakończyć proces ponownie za chwilę.");
  } finally {
    lock.releaseLock(); // Zawsze zwalniaj blokadę dla innych użytkowników
  }
}

function clearAllActiveSessions(adminToken) {
  requireAdminSession_(adminToken);


  PropertiesService.getScriptProperties()
    .deleteProperty('activeSessions');


  Logger.log('Wszystkie aktywne sesje zostały wyczyszczone.');


  return {
    success: true,
    message: 'Wszystkie aktywne sesje zostały wyczyszczone.'
  };
}
function fetchBigQueryData(sscc) {
  const PROJECT_ID = 'fml-prod-pl-bo';
  
  
  let ssccToSearch = sscc ? sscc.toString().trim() : "";
  
  if (ssccToSearch.startsWith("~") || ssccToSearch.startsWith("00")) {
      ssccToSearch = ssccToSearch.substring(2);
  }
  
  if (!ssccToSearch) return { ref_bo: "BRAK SSCC", qty: 0 };

  
  const QUERY = `
    SELECT 
        CAST(ItemCode AS STRING) as ItemCode, 
        LotQuantity
    FROM 
        \`fml-prod-pl-bo.WAREHOUSING.Stock\`
    WHERE 
        HDNumber = '${ssccToSearch}' 
        OR HDNumber = '00${ssccToSearch}'
        OR LTRIM(HDNumber, '0') = LTRIM('${ssccToSearch}', '0')
    LIMIT 1`; 

  try {
    const request = {
      query: QUERY,
      useLegacySql: false
    };

    let queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    
    while (!queryResults.jobComplete) {
      Utilities.sleep(500);
      queryResults = BigQuery.Jobs.getQueryResults(PROJECT_ID, queryResults.jobReference.jobId);
    }

    if (queryResults.rows && queryResults.rows.length > 0) {
      const row = queryResults.rows[0];
      
      let itemCodeRaw = row.f[0].v || "";
      
      let itemCodeFormatted = itemCodeRaw.toString();
      if (itemCodeFormatted !== "" && !itemCodeFormatted.startsWith("'")) {
        itemCodeFormatted = "'" + itemCodeFormatted;
      }

      return {
        ref_bo: itemCodeFormatted, 
        qty: row.f[1].v ? Math.round(parseFloat(row.f[1].v)) : 0 
      };
    
    } else {
      Logger.log("Nie znaleziono HDNumber: " + ssccToSearch);
      return { ref_bo: "NIE ZNALEZIONO", qty: 0 };
    }

  } catch (e) {
    Logger.log("Błąd BQ: " + e.toString());
    // Jeśli tu wejdzie, w Excelu zobaczysz ten komunikat
    return { 
        ref_bo: "BŁĄD POŁĄCZENIA", 
        qty: "ERROR" 
    };
  }
}
function saveProcessData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Calosc") || ss.insertSheet("Calosc");
  
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  
  const iloscFinalna = data.ilosc || data.Ilosc || data["Ilosc_szt"] || (["Przerwa", "Delegacja", "S_START", "Anomalia"].includes(data.procedure) ? "" : "1");

  const start = data.start ? new Date(data.start) : now;
  const stop = data.stop ? new Date(data.stop) : now;

  let duration = "00:00:00";
  const diffMS = stop.getTime() - start.getTime();
  if (diffMS >= 0) {
    const hours = Math.floor(diffMS / 3600000);
    const mins = Math.floor((diffMS % 3600000) / 60000);
    const secs = Math.floor((diffMS % 60000) / 1000);
    duration = `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
  }

  let refBo = "";
  let planQty = "";
  if (["Oklejanie","Oklejanieplus", "Auto_Labeling", "Simplex"].includes(data.procedure) && data.sscc) {
      const bqData = fetchBigQueryData(data.sscc);
      if (bqData) {
          refBo = bqData.ref_bo;
          planQty = bqData.qty;
      }
  } else if (data.procedure === "Copacking") {
      refBo = data["Referencja"] || "";
      planQty = data["Ilosc_szt"] || "";
  }
  let columnAValue = data.sscc || data.procedure;


  if (data.procedure === 'Anomalia') {
    columnAValue = data.Komentarz || data.procedure;
  } else if (data.procedure === 'Delegacja') {
    columnAValue = data.Klient || data.procedure;
  }
  const rowData = [
    columnAValue,          // A: SSCC
    data.loginMask || data.login || "",   // B: PRACOWNIK
    Utilities.formatDate(start, tz, "yyyy-MM-dd"), // C: DATA
    Utilities.formatDate(start, tz, "HH:mm:ss"),   // D: START
    Utilities.formatDate(stop, tz, "HH:mm:ss"),    // E: STOP
    iloscFinalna,                         // F: ILOSC
    duration,                             // G: CZAS_TRWANIA
    iloscFinalna,                         // H: SUMA_ILOSCI
    data["Ilosc_ludzi"] || "1",           // I: LUDZIE
    getWorkMode(data.procedure),          // J: TRYB
    refBo,                                // K: REF/BO
    planQty,                              // L: QTY
    data.login || ""                      // M: LOGIN_RAW
  ];

  const nextRow = sheet.getLastRow() + 1;
  const f1 = `=IF(J${nextRow}="S_PRW";IF(G${nextRow}>TIME(0;25;0);"TAK");"")`;
  const f2 = `=ROUND(VALUE(G${nextRow})*24;2)`;
  const f3 = `=IF(AND(A${nextRow}<>"";L${nextRow}<>"");IF(SUMIF(A:A;A${nextRow};H:H)=L${nextRow};"OK";"RÓŻNICA");"")`;
  
  const finalRow = [...rowData, f1, f2, f3];
  sheet.getRange(nextRow, 1, 1, finalRow.length).setValues([finalRow]);
  
  return "OK";
}


function getNameByLogin(login) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Login");
  if (!sheet) return login;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { // Zakładając, że pierwszy wiersz to nagłówek
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === login.toString().trim().toLowerCase()) {
      return data[i][1] ? data[i][1].toString().trim() : login;
    }
  }
  return login;
}

//function getWorkMode(proc) {
  //const map = {
  //  "Oklejanie": "S_OKL",
  //  "Oklejanieplus": "S_OKLK",
  //  "Auto_Labeling": "A_LAB",
  //  "Briefing": "M_BRF",
  //  "Przerwa": "S_PRW",
  //  "Anomalia": "S_ANO",
  //  "Copacking": "S_COP",
  //  "Delegacja": "S_DEL",
  //  "Drukowanie": "A_DRK",
  //  "Kartony": "M_KAR",
  //  "Simplex": "A_SIM",
  // "S_START": "S_STR",
  //  "S_END": "S_END",
  //  "Process": "S_PRC",
  //};
//  return map[proc] || proc || "";
//}
function getWorkMode(proc) {
  const procedureId = String(proc || '').trim();
  if (!procedureId) return '';


  const procedure = getProceduresConfig()
    .find(p => p.id === procedureId);


  return procedure ? procedure.workmode : procedureId;
}
