/**
 * ============================================================
 * E-HADIR SYSTEM BACKEND (VERSI 39.0 - PERFORMANCE PATCH)
 * Berdasarkan v38.0 (Integrasi Modul Cuti)
 * Perubahan v39.0:
 *   - Skip lock untuk operasi baca (GET_LIST/GET_CUTI/GET_CONFIG)
 *   - Cache jam config (CacheService, 6 jam)
 *   - Baca hanya 500 baris terakhir untuk operasi harian
 *   - Cache senarai hari ini 20 saat
 *   - Break awal dalam loop bawah-atas
 *   - Fungsi lain, format, UI TIDAK DIUBAH
 * ============================================================
 */

var ID_FOLDER_FOTO = "1Ia26pk1bhT1GZCRuTqzE7rLu0cSX8IRJ";
var PASSCODE_RAHSIA = "101010";
var idSheet = SpreadsheetApp.getActiveSpreadsheet().getId();
var SHEETS_LIST = ["GURU", "AKP", "KEBERSIHAN", "PELAWAT", "TIDAK HADIR"];

// Berapa baris terakhir untuk diimbas bagi carian hari ini.
// 500 sudah cukup lebih untuk sekolah biasa (kurang dari 200 rekod/hari).
var TAIL_SCAN_ROWS = 500;

function getMalaysiaTime() {
  return new Date(new Date().toLocaleString("en-US", {timeZone: "Asia/Kuala_Lumpur"}));
}

// Helpers Kritikal untuk Standardisasi Format Tarikh (DD.MM.YYYY)
function getSafeDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, "GMT+8", "dd.MM.yyyy");
  var s = String(val).trim();
  var p = s.split(" ")[0].split(".");
  if (p.length === 3) return ("0"+p[0]).slice(-2) + "." + ("0"+p[1]).slice(-2) + "." + p[2];
  var p2 = s.split(" ")[0].split("/");
  if (p2.length === 3) return ("0"+p2[0]).slice(-2) + "." + ("0"+p2[1]).slice(-2) + "." + p2[2];
  var p3 = s.split(" ")[0].split("-");
  if (p3.length === 3 && p3[0].length === 4) return ("0"+p3[2]).slice(-2) + "." + ("0"+p3[1]).slice(-2) + "." + p3[0];
  var d = new Date(val);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, "GMT+8", "dd.MM.yyyy");
  return s.split(" ")[0];
}

function parseSafeDateObj(val) {
  if (!val) return new Date(0);
  if (val instanceof Date) {
     return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  var s = String(val).trim().split(" ")[0];
  var sep = s.indexOf(".") > -1 ? "." : (s.indexOf("/") > -1 ? "/" : null);
  if (sep) {
     var p = s.split(sep);
     if (p.length === 3) return new Date(p[2], parseInt(p[1], 10)-1, p[0]);
  }
  if (s.indexOf("-") > -1) {
     var p2 = s.split("-");
     if (p2.length === 3 && p2[0].length === 4) return new Date(p2[0], parseInt(p2[1], 10)-1, p2[2]);
  }
  var d = new Date(val);
  if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return new Date(0);
}

// Helper baru: dapat objek Date daripada nilai cell dengan selamat (tanpa NaN/1970)
function safeCellToDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  var d = new Date(val);
  if (!isNaN(d.getTime()) && d.getFullYear() > 1971) return d;
  // Fallback: parse guna getSafeDateStr + jam kosong
  var s = getSafeDateStr(val);
  var p = s.split(".");
  if (p.length === 3) return new Date(p[2], parseInt(p[1],10)-1, parseInt(p[0],10));
  return null;
}

// Helper baru: baca hanya "tail" baris (paling akhir) untuk imbasan pantas hari ini
function getTailValues(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { values: [], startRow: 2 };
  var startRow = Math.max(2, lastRow - TAIL_SCAN_ROWS + 1);
  var numRows = lastRow - startRow + 1;
  var values = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  return { values: values, startRow: startRow };
}

// ============= doPost dengan lock hanya untuk operasi TULIS =============
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (!data.passcode || data.passcode !== PASSCODE_RAHSIA) {
       return sendResponse("error", "❌ AKSES DITOLAK!");
    }

    var action = data.action;
    var isWrite = (action === "MASUK" || action === "KELUAR" ||
                   action === "CHECKOUT_VISITOR" || action === "MOHON_CUTI" ||
                   action === "BATAL_CUTI" || action === "SAH_CUTI");

    if (isWrite) {
      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(30000);
      } catch (err) {
        return sendResponse("error", "Server sibuk, sila cuba sebentar lagi.");
      }
      try {
        return dispatchAction(action, data);
      } finally {
        lock.releaseLock();
      }
    } else {
      // Operasi baca — tak perlu lock, jangan berebut dengan penulis.
      return dispatchAction(action, data);
    }
  } catch (err) {
    return sendResponse("error", "RALAT SERVER: " + err.toString());
  }
}

function dispatchAction(action, data) {
  if (action == "MASUK") return daftarMasuk(data);
  if (action == "KELUAR") return daftarKeluar(data);
  if (action == "CHECKOUT_VISITOR") return checkoutVisitor(data.nama);
  if (action == "GET_LIST") return getSenaraiHadir(data);
  if (action == "GET_CONFIG") return getJamConfig();

  // --- ENDPOINTS MODUL CUTI ---
  if (action == "MOHON_CUTI") return mohonCuti(data);
  if (action == "GET_CUTI") return getCutiList(data);
  if (action == "BATAL_CUTI") return batalCuti(data);
  if (action == "SAH_CUTI") return sahCuti(data);

  return sendResponse("error", "Action tidak dikenali.");
}

// Cache helpers
function invalidateTodayListCache() {
  try {
    var cache = CacheService.getScriptCache();
    var todayStr = Utilities.formatDate(getMalaysiaTime(), "GMT+8", "dd.MM.yyyy");
    cache.remove("list_" + todayStr);
  } catch(e) {}
}

// --- FUNGSI OVERRIDE CUTI (AUTO-BATAL JIKA HADIR) ---
// Sheet TIDAK HADIR biasanya kecil. Kekal imbas penuh, tapi tapis awal.
function checkAndOverrideCuti(namaUser, dateObj) {
  var sheet = SpreadsheetApp.openById(idSheet).getSheetByName("TIDAK HADIR");
  if(!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var lastCol = Math.max(sheet.getLastColumn(), 12);
  var d = sheet.getRange(2, 1, lastRow-1, lastCol).getValues();
  var todayDateObj = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());

  for(var i=0; i<d.length; i++) {
    if (d[i][8] !== "AKTIF") continue;
    if (String(d[i][1]).trim() !== namaUser) continue;
    try {
      var startObj = parseSafeDateObj(d[i][3]);
      var endObj = parseSafeDateObj(d[i][4]);
      if(todayDateObj >= startObj && todayDateObj <= endObj) {
        sheet.getRange(i+2, 9).setValue("BATAL (HADIR BERTUGAS)");
      }
    } catch(e) {}
  }
}

// --- FUNGSI 1: DAFTAR MASUK ---
function daftarMasuk(data) {
  var targetSheetName = data.kategori ? data.kategori.toString().toUpperCase().trim() : "PELAWAT";
  if (SHEETS_LIST.indexOf(targetSheetName) === -1) targetSheetName = "PELAWAT";

  var sheet = SpreadsheetApp.openById(idSheet).getSheetByName(targetSheetName);
  if (!sheet) return sendResponse("error", "Ralat Kritikal: Sheet " + targetSheetName + " tidak dijumpai.");

  var now;
  if (data.clientTime) {
    now = new Date(data.clientTime);
    if (isNaN(now.getTime()) || now.getFullYear() === 1970) now = getMalaysiaTime();
  } else {
    now = getMalaysiaTime();
  }

  var inputNama = data.nama.toUpperCase().trim();

  // RUN AUTO-OVERRIDE CUTI
  if (targetSheetName !== "PELAWAT") {
    checkAndOverrideCuti(inputNama, now);
  }

  var masaString = Utilities.formatDate(now, "GMT+8", "hh:mm:ss a");
  var inputDevice = data.deviceId || "UNKNOWN_DEVICE";

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var lastDateVal = sheet.getRange(lastRow, 1).getValue();
    var lastNameVal = sheet.getRange(lastRow, 2).getValue();
    var lastDateObj = safeCellToDate(lastDateVal);
    if (lastDateObj) {
      var diffSeconds = (now.getTime() - lastDateObj.getTime()) / 1000;
      if (String(lastNameVal).trim() == inputNama && diffSeconds < 60) {
         return sendResponse("success", "Data telah direkod (Duplicate Filter).");
      }
    }
  }

  var fotoUrl = "TIADA";
  if (data.fotoBase64 && data.fotoBase64.length > 50) {
    try {
      var folder = DriveApp.getFolderById(ID_FOLDER_FOTO);
      var blob = Utilities.newBlob(Utilities.base64Decode(data.fotoBase64.split(",")[1]), "image/jpeg", inputNama + "_" + Date.now() + ".jpg");
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fotoUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";
    } catch(e) { fotoUrl = "RALAT FOLDER"; }
  }

  var status = "HADIR";
  var alasan = data.alasan || "-";
  var warnaJam = "black";
  var lokasiLink = "SEKOLAH";
  if (data.lat && data.lng) lokasiLink = "http://maps.google.com/maps?q=" + data.lat + "," + data.lng;

  var config = getJamConfigObject();
  var rule = config[targetSheetName];

  if (rule && rule.masuk !== "BEBAS") {
       var lM = parseTime(rule.masuk, now);
       var lK = parseTime(rule.keluar, now);
       if (lM && lK) {
           if (now > lK) { status = "HADIR"; warnaJam = "black"; }
           else if (now > lM) { status = "LEWAT"; warnaJam = "red"; }
           else { warnaJam = "blue"; }
       }
  } else if (targetSheetName === "PELAWAT") {
      status = "PELAWAT"; alasan = data.tujuan || "-";
  }

  sheet.appendRow([now, inputNama, targetSheetName, data.plat || "-", masaString, "", status, alasan, lokasiLink, fotoUrl, "", inputDevice, ""]);

  var lr = sheet.getLastRow();
  if (warnaJam !== "black") {
      sheet.getRange(lr, 5).setFontColor(warnaJam);
      if (warnaJam === "red") sheet.getRange(lr, 7).setFontColor("red");
  }

  invalidateTodayListCache();
  return sendResponse("success", "Berjaya Daftar: " + inputNama);
}

// --- FUNGSI 2: DAFTAR KELUAR ---
function daftarKeluar(data) {
  var doc = SpreadsheetApp.openById(idSheet);
  var now = data.clientTime ? new Date(data.clientTime) : getMalaysiaTime();
  if (isNaN(now.getTime()) || now.getFullYear() === 1970) now = getMalaysiaTime();
  var todayStr = Utilities.formatDate(now, "GMT+8", "dd.MM.yyyy");
  var masaString = Utilities.formatDate(now, "GMT+8", "hh:mm:ss a");
  var nama = data.nama;
  var config = getJamConfigObject();

  for (var s = 0; s < SHEETS_LIST.length; s++) {
    if (SHEETS_LIST[s] === "TIDAK HADIR") continue;
    var sheet = doc.getSheetByName(SHEETS_LIST[s]); if(!sheet) continue;

    var tail = getTailValues(sheet);
    var rawData = tail.values;
    var startRow = tail.startRow; // baris sebenar dalam sheet untuk rawData[0]

    for (var i = rawData.length-1; i>=0; i--) {
       if (!rawData[i][0]) continue;
       var rowDate = getSafeDateStr(rawData[i][0]);
       // Break awal: sekali jumpa baris hari LAIN (tarikh lebih awal), stop.
       if (rowDate !== todayStr) {
         // Jika sudah lepas hari sasaran (tarikh lama), berhenti terus.
         var rowObj = parseSafeDateObj(rawData[i][0]);
         var todayObj = parseSafeDateObj(todayStr);
         if (rowObj < todayObj) break;
         continue;
       }
       var rowNama = String(rawData[i][1]).trim();

       if(rowNama == nama && rawData[i][5] == "") {
         var actualRow = startRow + i;
         var rule = config[SHEETS_LIST[s]];
         var statusKeluar = "SELESAI";
         var warnaJam = "black";
         var msg = "Jumpa lagi " + nama;
         var statusIn = rawData[i][6];
         var isLateIn = (String(statusIn).indexOf("LEWAT") !== -1);

         if (data.alasan && data.alasan.indexOf("[LUAR KAWASAN]") !== -1) {
             warnaJam = "#f39c12";
             statusKeluar = "LUAR KAWASAN";
         }
         else if (rule && rule.keluar !== "BEBAS") {
             var limitKeluar = parseTime(rule.keluar, now);
             if (limitKeluar) {
                 if (now < limitKeluar) {
                     statusKeluar = "BALIK AWAL";
                     warnaJam = "red";
                 } else {
                     warnaJam = "green";
                     if (!isLateIn) {
                         var quotes = ["Tahniah! Anda pekerja cemerlang.", "Terima kasih atas dedikasi anda!", "Hebat! Teruskan usaha murni ini."];
                         msg += ". " + quotes[Math.floor(Math.random() * quotes.length)];
                     }
                 }
             }
         }

         sheet.getRange(actualRow, 6).setValue(masaString);
         sheet.getRange(actualRow, 7).setValue(statusKeluar);
         if (data.alasan && data.alasan !== "-") {
             var oldReason = rawData[i][7];
             var newReason = (oldReason == "-" || oldReason == "") ? data.alasan : oldReason + " | Keluar: " + data.alasan;
             sheet.getRange(actualRow, 8).setValue(newReason);
         }

         if (typeof data.lat !== 'undefined' && typeof data.lng !== 'undefined') {
             sheet.getRange(actualRow, 11).setValue("http://maps.google.com/maps?q=" + data.lat + "," + data.lng);
         }

         if (data.locStatus) {
             sheet.getRange(actualRow, 13).setValue(data.locStatus);
         } else {
             sheet.getRange(actualRow, 13).setValue("TIADA");
         }

         if(warnaJam !== "black") {
             sheet.getRange(actualRow, 6).setFontColor(warnaJam);
             if (warnaJam === "red") sheet.getRange(actualRow, 7).setFontColor("red");
         }
         invalidateTodayListCache();
         return sendResponse("success", msg);
       }
    }
  }
  return sendResponse("error", "Tiada rekod masuk aktif hari ini.");
}

// --- FUNGSI 3: MODUL CUTI (BACKEND) ---
function checkAndCreateCutiSheet() {
  var ss = SpreadsheetApp.openById(idSheet);
  var sheet = ss.getSheetByName("TIDAK HADIR");
  if(!sheet) {
    sheet = ss.insertSheet("TIDAK HADIR");
    sheet.appendRow(["Timestamp", "Nama", "Pekerjaan", "Tarikh Mula", "Tarikh Tamat", "Jenis Cuti", "Alasan", "Status Pengesahan", "Status Rekod", "UUID", "Bukti URL", "Admin Pengesah"]);
    sheet.getRange("A1:L1").setFontWeight("bold");
  }
  return sheet;
}

function mohonCuti(data) {
  var sheet = checkAndCreateCutiSheet();
  var now = getMalaysiaTime();
  var uuid = "CUTI_" + Utilities.getUuid();

  var fotoUrl = "";
  if (data.fotoBase64 && data.fotoBase64.length > 50) {
    try {
      var folder = DriveApp.getFolderById(ID_FOLDER_FOTO);
      var dataParts = data.fotoBase64.split(",");
      var mimeString = dataParts[0].split(":")[1].split(";")[0];
      var rawBase64 = dataParts[1];

      var originalName = data.fileName || "lampiran";
      var finalName = data.nama + "_BUKTI_CUTI_" + Date.now() + "_" + originalName;

      var blob = Utilities.newBlob(Utilities.base64Decode(rawBase64), mimeString, finalName);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fotoUrl = file.getUrl();
    } catch(e) { fotoUrl = "RALAT FOLDER: " + e.toString(); }
  }

  if (sheet.getLastColumn() < 12) {
      sheet.getRange(1, 1, 1, 12).setValues([["Timestamp", "Nama", "Pekerjaan", "Tarikh Mula", "Tarikh Tamat", "Jenis Cuti", "Alasan", "Status Pengesahan", "Status Rekod", "UUID", "Bukti URL", "Admin Pengesah"]]);
      sheet.getRange("A1:L1").setFontWeight("bold");
  }

  sheet.appendRow([
    now,
    data.nama,
    data.kategori,
    data.tarikhMula,
    data.tarikhTamat,
    data.jenisCuti,
    data.alasan,
    "BELUM DISAHKAN",
    "AKTIF",
    uuid,
    fotoUrl,
    ""
  ]);
  invalidateTodayListCache();
  return sendResponse("success", "Cuti direkodkan.");
}

function getCutiList(data) {
  var sheet = checkAndCreateCutiSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
  var lastCol = Math.max(sheet.getLastColumn(), 12);
  var d = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var list = [];

  var todayDateObj = new Date();
  todayDateObj.setHours(0,0,0,0);

  for(var i=1; i<d.length; i++) {
    var rec = {
      nama: d[i][1],
      pekerjaan: d[i][2],
      tarikhMula: getSafeDateStr(d[i][3]),
      tarikhTamat: getSafeDateStr(d[i][4]),
      jenisCuti: d[i][5],
      alasan: d[i][6],
      statusPengesahan: d[i][7],
      statusRekod: d[i][8],
      uuid: d[i][9],
      buktiUrl: d[i][10] || "",
      adminPengesah: d[i][11] || ""
    };

    var startObj = parseSafeDateObj(d[i][3]);
    var endObj = parseSafeDateObj(d[i][4]);

    if (data.mode === "public") {
      if (rec.statusRekod === "AKTIF" && todayDateObj >= startObj && todayDateObj <= endObj) {
        list.push({nama: rec.nama, pekerjaan: rec.pekerjaan, jenisCuti: rec.jenisCuti, statusPengesahan: rec.statusPengesahan, adminPengesah: rec.adminPengesah});
      }
    } else if (data.mode === "history") {
      if (String(rec.nama).trim() === String(data.nama).trim()) {
        list.push(rec);
      }
    } else if (data.mode === "admin") {
      if (rec.statusRekod === "AKTIF") {
        list.push(rec);
      }
    }
  }
  if(data.mode === "history") list.reverse();
  return ContentService.createTextOutput(JSON.stringify(list)).setMimeType(ContentService.MimeType.JSON);
}

function batalCuti(data) {
  var sheet = checkAndCreateCutiSheet();
  var d = sheet.getDataRange().getValues();
  for(var i=1; i<d.length; i++) {
    if(d[i][9] === data.uuid) {
      sheet.getRange(i+1, 9).setValue("SUDAH DIBATALKAN");
      invalidateTodayListCache();
      return sendResponse("success", "Batal OK");
    }
  }
  return sendResponse("error", "Rekod tidak dijumpai");
}

function sahCuti(data) {
  var sheet = checkAndCreateCutiSheet();
  var d = sheet.getDataRange().getValues();
  for(var i=1; i<d.length; i++) {
    if(d[i][9] === data.uuid) {
      sheet.getRange(i+1, 8).setValue("DISAHKAN");
      sheet.getRange(i+1, 12).setValue(data.adminName || "ADMIN");
      invalidateTodayListCache();
      return sendResponse("success", "Sah OK");
    }
  }
  return sendResponse("error", "Rekod tidak dijumpai");
}

// --- FUNGSI CONFIG (dengan CacheService) ---
function getJamConfigObject() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("jam_config");
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }
  var s = SpreadsheetApp.openById(idSheet).getSheetByName("JAM");
  if(!s) return {};
  var d = s.getDataRange().getValues();
  function formatMasa(rawVal) {
    if (rawVal instanceof Date) return Utilities.formatDate(rawVal, "GMT+8", "hh:mm:ss a");
    return rawVal ? String(rawVal) : "BEBAS";
  }
  var config = {
    "GURU":       { masuk: formatMasa(d[1][1]), keluar: formatMasa(d[2][1]) },
    "AKP":        { masuk: formatMasa(d[1][2]), keluar: formatMasa(d[2][2]) },
    "KEBERSIHAN": { masuk: formatMasa(d[1][3]), keluar: formatMasa(d[2][3]) },
    "PELAWAT":    { masuk: formatMasa(d[1][4]), keluar: formatMasa(d[2][4]) }
  };
  try { cache.put("jam_config", JSON.stringify(config), 21600); } catch(e) {}
  return config;
}

function getJamConfig() {
  var config = getJamConfigObject();
  return ContentService.createTextOutput(JSON.stringify(config)).setMimeType(ContentService.MimeType.JSON);
}

function checkoutVisitor(nama) {
  var s = SpreadsheetApp.openById(idSheet).getSheetByName("PELAWAT");
  if (!s) return sendResponse("error","Gagal");
  var d = getMalaysiaTime();
  var todayStr = Utilities.formatDate(d, "GMT+8", "dd.MM.yyyy");

  var tail = getTailValues(s);
  var v = tail.values;
  var startRow = tail.startRow;
  var todayObj = parseSafeDateObj(todayStr);

  for(var i=v.length-1; i>=0; i--){
    if (!v[i][0]) continue;
    var rowDate = getSafeDateStr(v[i][0]);
    if (rowDate !== todayStr) {
      var rowObj = parseSafeDateObj(v[i][0]);
      if (rowObj < todayObj) break;
      continue;
    }
    if (String(v[i][1]).trim() === nama && v[i][5] === ""){
      var actualRow = startRow + i;
      s.getRange(actualRow, 6).setValue(Utilities.formatDate(d, "GMT+8", "hh:mm:ss a"));
      s.getRange(actualRow, 7).setValue("SELESAI");
      invalidateTodayListCache();
      return sendResponse("success", "Pelawat Checkout OK");
    }
  }
  return sendResponse("error", "Gagal");
}

function getSenaraiHadir(data) {
  var d = SpreadsheetApp.openById(idSheet);
  var todayStr = Utilities.formatDate(getMalaysiaTime(), "GMT+8", "dd.MM.yyyy");
  var targetDate = (data && data.targetDate) ? data.targetDate : todayStr;
  var isToday = (targetDate === todayStr);

  // Cache 20 saat untuk senarai hari ini
  var cache = CacheService.getScriptCache();
  var cacheKey = "list_" + targetDate;
  if (isToday) {
    var cached = cache.get(cacheKey);
    if (cached) {
      return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
    }
  }

  var targetObj = parseSafeDateObj(targetDate);
  var l = [];

  for(var s=0; s<SHEETS_LIST.length; s++){
    var shName = SHEETS_LIST[s];
    var sh = d.getSheetByName(shName);
    if(!sh) continue;

    if (shName === "TIDAK HADIR") {
      // Sheet kecil — imbas penuh tapi tapis awal
      var lastRow = sh.getLastRow();
      if (lastRow < 2) continue;
      var lastCol = Math.max(sh.getLastColumn(), 12);
      var v = sh.getRange(2, 1, lastRow-1, lastCol).getValues();
      for(var i=0; i<v.length; i++){
        if (!v[i][0]) continue;
        if (v[i][8] !== "AKTIF") continue;
        try {
          var startObj = parseSafeDateObj(v[i][3]);
          var endObj = parseSafeDateObj(v[i][4]);
          if (targetObj >= startObj && targetObj <= endObj) {
            l.push({ nama: v[i][1], kategori: "CUTI", jam: v[i][5], keluar: "TIDAK HADIR", foto: "", statusSah: v[i][7], admin: v[i][11] });
          }
        } catch(e) {}
      }
    } else {
      // Sheet kehadiran besar — imbas hanya tail; break bila jumpa hari lebih awal
      var tail = isToday ? getTailValues(sh) : (function(){
        var lr = sh.getLastRow();
        var lc = sh.getLastColumn();
        if (lr < 2) return { values: [], startRow: 2 };
        return { values: sh.getRange(2, 1, lr-1, lc).getValues(), startRow: 2 };
      })();
      var v2 = tail.values;
      for(var i2=v2.length-1; i2>=0; i2--){
        if(!v2[i2][0]) continue;
        var rowDateStr = getSafeDateStr(v2[i2][0]);
        if (rowDateStr === targetDate) {
          var locStat = (v2[i2].length > 12) ? v2[i2][12] : "TIADA";
          l.push({ nama: v2[i2][1], kategori: shName, jam: v2[i2][4], keluar: v2[i2][5], foto: v2[i2][9]||"", locStatus: locStat });
        } else if (isToday) {
          // Untuk hari ini, sebaik jumpa tarikh lebih awal, berhenti.
          var rowObj = parseSafeDateObj(v2[i2][0]);
          if (rowObj < targetObj) break;
        }
      }
    }
  }

  var payload = JSON.stringify(l);
  if (isToday) {
    try { cache.put(cacheKey, payload, 20); } catch(e) {}
  }
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

function parseTime(t,r){
  if(!t || t=="BEBAS")return null;
  var d=new Date(r);
  var p=t.match(/(\d+)[.:](\d+).*?(AM|PM)/i);
  if(p){
    var h=parseInt(p[1]),m=parseInt(p[2]);
    if(p[3]=="PM"&&h<12)h+=12;
    if(p[3]=="AM"&&h==12)h=0;
    d.setHours(h,m,0,0);
    return d;
  }
  return null;
}

function sendResponse(s,m) { return ContentService.createTextOutput(JSON.stringify({"result":s,"msg":m})).setMimeType(ContentService.MimeType.JSON); }
