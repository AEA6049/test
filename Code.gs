/**
 * ============================================================
 * E-HADIR SYSTEM BACKEND (VERSI 38.0 - INTEGRASI MODUL CUTI)
 * DIKEMASKINI OLEH: AI FORENSIC REPAIR (CUSTOMIZED)
 * TARIKH & FORMAT: DD.MM.YYYY (FIXED UNTUK DAFTAR MASUK/KELUAR)
 * ============================================================
 */

var ID_FOLDER_FOTO = "1Ia26pk1bhT1GZCRuTqzE7rLu0cSX8IRJ"; 
var PASSCODE_RAHSIA = "101010"; 
var idSheet = SpreadsheetApp.getActiveSpreadsheet().getId();
var SHEETS_LIST = ["GURU", "AKP", "KEBERSIHAN", "PELAWAT", "TIDAK HADIR"];

function getMalaysiaTime() {
  return new Date(new Date().toLocaleString("en-US", {timeZone: "Asia/Kuala_Lumpur"}));
}

// Helpers Kritikal untuk Standardisasi Format Tarikh (DD.MM.YYYY) Mengatasi Masalah Locale Google Sheet
function getSafeDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, "GMT+8", "dd.MM.yyyy");
  var s = String(val).trim();
  var p = s.split(" ")[0].split(".");
  if (p.length === 3) return ("0"+p[0]).slice(-2) + "." + ("0"+p[1]).slice(-2) + "." + p[2];
  var p2 = s.split(" ")[0].split("/");
  if (p2.length === 3) return ("0"+p2[0]).slice(-2) + "." + ("0"+p2[1]).slice(-2) + "." + p2[2];
  var p3 = s.split(" ")[0].split("-"); // yyyy-mm-dd
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

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
  } catch (err) {
    return sendResponse("error", "Server sibuk, sila cuba sebentar lagi.");
  }

  try {
    var data = JSON.parse(e.postData.contents);

    if (!data.passcode || data.passcode !== PASSCODE_RAHSIA) {
       return sendResponse("error", "❌ AKSES DITOLAK!");
    }

    if (data.action == "MASUK") return daftarMasuk(data);
    if (data.action == "KELUAR") return daftarKeluar(data); 
    if (data.action == "CHECKOUT_VISITOR") return checkoutVisitor(data.nama);
    if (data.action == "GET_LIST") return getSenaraiHadir(data);
    if (data.action == "GET_CONFIG") return getJamConfig();
    
    // --- ENDPOINTS MODUL CUTI ---
    if (data.action == "MOHON_CUTI") return mohonCuti(data);
    if (data.action == "GET_CUTI") return getCutiList(data);
    if (data.action == "BATAL_CUTI") return batalCuti(data);
    if (data.action == "SAH_CUTI") return sahCuti(data);

    // --- ENDPOINTS MODUL ARKIB AUTO-BULANAN ---
    if (data.action == "GET_ARKIB") return getArkibData(data);
    if (data.action == "GET_BULAN_ARKIB") return getSenaraiBulanArkib();

  } catch (err) {
    return sendResponse("error", "RALAT SERVER: " + err.toString());
  } finally {
    lock.releaseLock();
  }
}

// --- FUNGSI OVERRIDE CUTI (AUTO-BATAL JIKA HADIR) ---
function checkAndOverrideCuti(namaUser, dateObj) {
  var sheet = SpreadsheetApp.openById(idSheet).getSheetByName("TIDAK HADIR");
  if(!sheet) return;
  var d = sheet.getDataRange().getValues();
  var todayDateObj = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  
  for(var i=1; i<d.length; i++) {
    if(String(d[i][1]).trim() === namaUser && d[i][8] === "AKTIF") {
      try {
        var startObj = parseSafeDateObj(d[i][3]);
        var endObj = parseSafeDateObj(d[i][4]);
        if(todayDateObj >= startObj && todayDateObj <= endObj) {
          sheet.getRange(i+1, 9).setValue("BATAL (HADIR BERTUGAS)");
        }
      } catch(e) {}
    }
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
    var diffSeconds = (now.getTime() - new Date(lastDateVal).getTime()) / 1000;
    if (String(lastNameVal).trim() == inputNama && diffSeconds < 60) {
       return sendResponse("success", "Data telah direkod (Duplicate Filter).");
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

  var config = JSON.parse(getJamConfig().getContent()); 
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

  // Penstrukturan lajur: 0=Timestamp, 1=Nama, 2=Kategori, 3=Plat, 4=MasaMasuk, 5=MasaKeluar(""), 6=Status, 7=Alasan, 8=GPS Masuk, 9=Foto, 10=GPS Keluar(""), 11=InputDevice, 12=Status Lokasi
  sheet.appendRow([now, inputNama, targetSheetName, data.plat || "-", masaString, "", status, alasan, lokasiLink, fotoUrl, "", inputDevice, ""]);
  
  var lr = sheet.getLastRow(); 
  if (warnaJam !== "black") {
      sheet.getRange(lr, 5).setFontColor(warnaJam); 
      if (warnaJam === "red") sheet.getRange(lr, 7).setFontColor("red"); 
  }

  return sendResponse("success", "Berjaya Daftar: " + inputNama);
}

// --- FUNGSI 2: DAFTAR KELUAR ---
function daftarKeluar(data) {
  var doc = SpreadsheetApp.openById(idSheet);
  var now = data.clientTime ? new Date(data.clientTime) : getMalaysiaTime();
  var todayStr = Utilities.formatDate(now, "GMT+8", "dd.MM.yyyy");
  var masaString = Utilities.formatDate(now, "GMT+8", "hh:mm:ss a");
  var nama = data.nama;

  for (var s = 0; s < SHEETS_LIST.length; s++) {
    if (SHEETS_LIST[s] === "TIDAK HADIR") continue;
    var sheet = doc.getSheetByName(SHEETS_LIST[s]); if(!sheet) continue;
    var rawData = sheet.getDataRange().getValues();
    
    for (var i = rawData.length-1; i>=1; i--) {
       var rowDate = getSafeDateStr(rawData[i][0]);
       var rowNama = String(rawData[i][1]).trim(); 
       
       if(rowDate === todayStr && rowNama == nama && rawData[i][5] == "") {
         var config = JSON.parse(getJamConfig().getContent());
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

         sheet.getRange(i+1, 6).setValue(masaString);   
         sheet.getRange(i+1, 7).setValue(statusKeluar); 
         if (data.alasan && data.alasan !== "-") {
             var oldReason = rawData[i][7]; 
             var newReason = (oldReason == "-" || oldReason == "") ? data.alasan : oldReason + " | Keluar: " + data.alasan;
             sheet.getRange(i+1, 8).setValue(newReason);
         }
         
         // REKOD GPS DAFTAR KELUAR DI LAJUR 11 (K) TERMASUK JIKA GPS OFF
         if (typeof data.lat !== 'undefined' && typeof data.lng !== 'undefined') {
             sheet.getRange(i+1, 11).setValue("http://maps.google.com/maps?q=" + data.lat + "," + data.lng);
         }
         
         // REKOD STATUS LOKASI (DALAM/LUAR/TIADA) DI LAJUR 13 (M)
         if (data.locStatus) {
             sheet.getRange(i+1, 13).setValue(data.locStatus);
         } else {
             sheet.getRange(i+1, 13).setValue("TIADA");
         }
         
         if(warnaJam !== "black") {
             sheet.getRange(i+1, 6).setFontColor(warnaJam); 
             if (warnaJam === "red") sheet.getRange(i+1, 7).setFontColor("red"); 
         }
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
    data.tarikhMula, // Disimpan sebagai teks DD.MM.YYYY dari Front-End
    data.tarikhTamat,
    data.jenisCuti,
    data.alasan,
    "BELUM DISAHKAN",
    "AKTIF",
    uuid,
    fotoUrl,
    ""
  ]);
  return sendResponse("success", "Cuti direkodkan.");
}

function getCutiList(data) {
  var sheet = checkAndCreateCutiSheet();
  var d = sheet.getDataRange().getValues();
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
      return sendResponse("success", "Sah OK");
    }
  }
  return sendResponse("error", "Rekod tidak dijumpai");
}

// --- FUNGSI CONFIG & LAIN ---
function getJamConfig() { 
  var s = SpreadsheetApp.openById(idSheet).getSheetByName("JAM"); 
  if(!s) return ContentService.createTextOutput("{}"); 
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
  return ContentService.createTextOutput(JSON.stringify(config)).setMimeType(ContentService.MimeType.JSON); 
}

function checkoutVisitor(nama) { 
  var s=SpreadsheetApp.openById(idSheet).getSheetByName("PELAWAT"); 
  var d=getMalaysiaTime(); 
  var todayStr=Utilities.formatDate(d,"GMT+8","dd.MM.yyyy"); 
  var v=s.getDataRange().getValues(); 
  for(var i=v.length-1;i>=1;i--){
    var rowDate = getSafeDateStr(v[i][0]);
    if(rowDate === todayStr && String(v[i][1]).trim() === nama && v[i][5] === ""){ 
      s.getRange(i+1,6).setValue(Utilities.formatDate(d,"GMT+8","hh:mm:ss a")); 
      s.getRange(i+1,7).setValue("SELESAI"); 
      return sendResponse("success","Pelawat Checkout OK"); 
    } 
  } 
  return sendResponse("error","Gagal"); 
}

function getSenaraiHadir(data) { 
  var d=SpreadsheetApp.openById(idSheet); 
  var todayStr = Utilities.formatDate(getMalaysiaTime(),"GMT+8","dd.MM.yyyy"); 
  var targetDate = (data && data.targetDate) ? data.targetDate : todayStr;
  var l=[]; 
  for(var s=0;s<SHEETS_LIST.length;s++){
    var shName = SHEETS_LIST[s];
    var sh=d.getSheetByName(shName); 
    if(!sh)continue; 
    var v=sh.getDataRange().getValues(); 

    if (shName === "TIDAK HADIR") {
        for(var i=1; i<v.length; i++){
            if (!v[i][0]) continue;
            try {
                var targetObj = parseSafeDateObj(targetDate);
                var startObj = parseSafeDateObj(v[i][3]);
                var endObj = parseSafeDateObj(v[i][4]);
                var statusRekod = v[i][8];
                if (statusRekod === "AKTIF" && targetObj >= startObj && targetObj <= endObj) {
                    l.push({ nama: v[i][1], kategori: "CUTI", jam: v[i][5], keluar: "TIDAK HADIR", foto: "", statusSah: v[i][7], admin: v[i][11] });
                }
            } catch(e) {}
        }
    } else {
        for(var i=1;i<v.length;i++){
          if(!v[i][0]) continue;
          if(getSafeDateStr(v[i][0]) === targetDate){
            var locStat = (v[i].length > 12) ? v[i][12] : "TIADA";
            l.push({ nama: v[i][1], kategori: shName, jam: v[i][4], keluar: v[i][5], foto: v[i][9]||"", locStatus: locStat }); 
          } 
        } 
    }
  } 
  return ContentService.createTextOutput(JSON.stringify(l)).setMimeType(ContentService.MimeType.JSON); 
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


/**
 * ============================================================
 * MODUL ARKIB AUTO-BULANAN (v1.0)
 * Auto arkib data kehadiran bulanan ke sheet ARKIB_DATA
 * dan bersihkan sheet aktif supaya database sentiasa ringan.
 * ============================================================
 */

var ARKIB_SHEET_NAME = "ARKIB_DATA";
var ARKIB_HEADERS = ["KATEGORI","BULAN_ARKIB","TARIKH","NAMA","JAWATAN/KELAS","MASA MASUK","MASA KELUAR","STATUS","CATATAN","LOKASI","FOTO"];

function ensureArkibSheet() {
  var ss = SpreadsheetApp.openById(idSheet);
  var sh = ss.getSheetByName(ARKIB_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ARKIB_SHEET_NAME);
    sh.getRange(1,1,1,ARKIB_HEADERS.length).setValues([ARKIB_HEADERS]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1,110); sh.setColumnWidth(2,100); sh.setColumnWidth(4,180);
  }
  // Paksa format kolum supaya tarikh & masa dipapar dengan betul
  // (elakkan bug "12/30/1899" pada kolum masa)
  sh.getRange("C2:C").setNumberFormat("dd/MM/yyyy");   // TARIKH
  sh.getRange("F2:F").setNumberFormat("HH:mm:ss");     // MASA MASUK
  sh.getRange("G2:G").setNumberFormat("HH:mm:ss");     // MASA KELUAR
  return sh;
}

/**
 * Fungsi utama — arkib data bulan SEBELUM ini, kemudian padam dari sheet aktif.
 * Dipicu automatik oleh trigger setiap 1hb pukul 1:00 AM.
 */
function arkibBulanan() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(60000); } catch(e) { return; }
  try {
    var ss = SpreadsheetApp.openById(idSheet);
    var arkibSh = ensureArkibSheet();

    var now = getMalaysiaTime();
    var target = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    var tM = target.getMonth(), tY = target.getFullYear();
    var bulanArkib = Utilities.formatDate(target, "GMT+8", "MM.yyyy");

    var totalArkib = 0, totalPadam = 0;

    SHEETS_LIST.forEach(function(sheetName) {
      var sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) return;

      var lastRow = sh.getLastRow();
      var lastCol = sh.getLastColumn();
      var data = sh.getRange(2,1,lastRow-1,lastCol).getValues();

      var toArchive = [];
      var rowsToDelete = [];

      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var d = parseSafeDateObj(row[0]);
        if (d.getMonth() === tM && d.getFullYear() === tY) {
          var archRow = [sheetName, bulanArkib];
          for (var j = 0; j < 9; j++) archRow.push(row[j] !== undefined ? row[j] : "");
          toArchive.push(archRow);
          rowsToDelete.push(i + 2);
        }
      }

      if (toArchive.length > 0) {
        arkibSh.getRange(arkibSh.getLastRow()+1, 1, toArchive.length, ARKIB_HEADERS.length)
               .setValues(toArchive);
        totalArkib += toArchive.length;
        rowsToDelete.sort(function(a,b){return b-a;});
        rowsToDelete.forEach(function(r){ sh.deleteRow(r); });
        totalPadam += rowsToDelete.length;
      }
    });

    Logger.log("Arkib "+bulanArkib+" selesai. Arkib: "+totalArkib+" | Padam: "+totalPadam);
  } catch(err) {
    Logger.log("ERROR arkibBulanan: " + err.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * Endpoint frontend: dapatkan rekod arkib.
 * Body: { action:"GET_ARKIB", bulan:"11.2025", kategori:"GURU", nama:"..." }
 * Semua filter optional. Jika kosong = pulangkan semua.
 */
function getArkibData(data) {
  try {
    var ss = SpreadsheetApp.openById(idSheet);
    var sh = ss.getSheetByName(ARKIB_SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) return sendResponse("success", []);

    var rows = sh.getRange(2,1,sh.getLastRow()-1,ARKIB_HEADERS.length).getValues();
    var fBulan = (data.bulan||"").trim();
    var fKat = (data.kategori||"").trim().toUpperCase();
    var fNama = (data.nama||"").trim().toLowerCase();
    var result = [];

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (fBulan && String(r[1]) !== fBulan) continue;
      if (fKat && String(r[0]).toUpperCase() !== fKat) continue;
      if (fNama && String(r[3]).toLowerCase().indexOf(fNama) === -1) continue;
      result.push({
        kategori:r[0], bulan:r[1], tarikh:getSafeDateStr(r[2]),
        nama:r[3], jawatan:r[4], masuk:r[5], keluar:r[6],
        status:r[7], catatan:r[8], lokasi:r[9], foto:r[10]
      });
    }
    return sendResponse("success", result);
  } catch(err) {
    return sendResponse("error", err.toString());
  }
}

/**
 * Endpoint frontend: senarai bulan tersedia dalam arkib (untuk dropdown).
 */
function getSenaraiBulanArkib() {
  try {
    var ss = SpreadsheetApp.openById(idSheet);
    var sh = ss.getSheetByName(ARKIB_SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) return sendResponse("success", []);
    var col = sh.getRange(2,2,sh.getLastRow()-1,1).getValues();
    var set = {};
    col.forEach(function(r){ if (r[0]) set[String(r[0])] = true; });
    var list = Object.keys(set).sort(function(a,b){
      var pa=a.split("."), pb=b.split(".");
      return (parseInt(pb[1])-parseInt(pa[1])) || (parseInt(pb[0])-parseInt(pa[0]));
    });
    return sendResponse("success", list);
  } catch(err) {
    return sendResponse("error", err.toString());
  }
}

/**
 * Jalankan SEKALI SAHAJA secara manual dari editor Apps Script
 * untuk memasang jadual auto-arkib.
 */
function pasangTriggerArkib() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === "arkibBulanan") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("arkibBulanan").timeBased().onMonthDay(1).atHour(1).create();
  Logger.log("Trigger arkib dipasang: 1hb setiap bulan, 1:00 AM waktu Malaysia.");
}

/**
 * Untuk test manual — arkib data bulan lepas serta-merta tanpa tunggu 1hb.
 */
function ujiArkibSekarang() { arkibBulanan(); }

/**
 * SKRIP TEST — pindah SEMUA data kehadiran (semua bulan, semua pekerja)
 * dari sheet aktif ke ARKIB_DATA, kemudian bersihkan sheet aktif.
 * Sheet ARKIB_DATA akan auto-create jika belum wujud.
 *
 * Jalankan sekali dari editor Apps Script: pilih fungsi `ujiArkibSemua`
 * kemudian tekan ▶ Run. Semak log (View > Logs) untuk laporan.
 */
function ujiArkibSemua() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(60000); } catch(e) { Logger.log("Lock gagal"); return; }
  try {
    var ss = SpreadsheetApp.openById(idSheet);
    var arkibSh = ensureArkibSheet();
    var totalArkib = 0, totalPadam = 0;
    var laporan = [];

    SHEETS_LIST.forEach(function(sheetName) {
      var sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) {
        laporan.push(sheetName + ": kosong");
        return;
      }
      var lastRow = sh.getLastRow();
      var lastCol = sh.getLastColumn();
      var data = sh.getRange(2,1,lastRow-1,lastCol).getValues();

      var toArchive = [];
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var d = parseSafeDateObj(row[0]);
        var bulanArkib = Utilities.formatDate(d, "GMT+8", "MM.yyyy");
        var archRow = [sheetName, bulanArkib];
        for (var j = 0; j < 9; j++) archRow.push(row[j] !== undefined ? row[j] : "");
        toArchive.push(archRow);
      }

      if (toArchive.length > 0) {
        arkibSh.getRange(arkibSh.getLastRow()+1, 1, toArchive.length, ARKIB_HEADERS.length)
               .setValues(toArchive);
        totalArkib += toArchive.length;
        // Padam semua row data (kekalkan header)
        // Kosongkan kandungan data (kekalkan header di baris 1).
        // Guna clearContent kerana deleteRows tidak boleh padam semua baris non-frozen.
        sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
        totalPadam += toArchive.length;
        laporan.push(sheetName + ": " + toArchive.length + " rekod diarkib");
      } else {
        laporan.push(sheetName + ": tiada data");
      }
    });

    Logger.log("=== UJI ARKIB SEMUA SELESAI ===");
    laporan.forEach(function(l){ Logger.log(l); });
    Logger.log("JUMLAH: " + totalArkib + " rekod diarkib, " + totalPadam + " rekod dipadam dari sheet aktif.");
    Logger.log("Semak sheet 'ARKIB_DATA' di Google Sheets anda.");
  } catch(err) {
    Logger.log("ERROR ujiArkibSemua: " + err.toString());
  } finally {
    lock.releaseLock();
  }
}
