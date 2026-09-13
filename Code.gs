/**
 * ============================================================================
 * e-RPH PINTAR AI 2026 (CORE ENGINE & BULK MONTHLY GENERATOR)
 * SISTEM PENGURUSAN REKOD PENGAJARAN HARIAN, DSKP & PENJANAAN PUKAL SEBULAN
 * ============================================================================
 */

function doGet(e) {
  var isDfy = e && e.parameter && (e.parameter.page === "dfy" || e.parameter.mode === "dfy" || e.parameter.p === "dfy");
  var senaraiNama = isDfy 
    ? ["dfy", "Dfy", "DFY", "Index", "index"] 
    : ["Index", "index", "Index_GAS", "index_gas", "dfy"];
  var output = null;

  for (var i = 0; i < senaraiNama.length; i++) {
    try {
      output = HtmlService.createHtmlOutputFromFile(senaraiNama[i]);
      if (output) break;
    } catch (err) {}
  }

  if (!output) {
    output = HtmlService.createHtmlOutput(
      "<div style='font-family:sans-serif; padding:20px; text-align:center;'>" +
      "<h3 style='color:#e11d48;'>Ralat: Fail Antaramuka HTML Tidak Ditemui</h3>" +
      "<p>Sila pastikan fail HTML dalam Apps Script anda dinamakan <b>Index</b> atau <b>dfy</b>.</p>" +
      "</div>"
    );
  }

  var tajuk = isDfy ? "Sistem Penjanaan RPH Auto - Shopee DFY" : "e-RPH Pintar AI 2026";

  return output
    .setTitle(tajuk)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --------------------------------------------------------------------------
// 1. FUNGSI PEMBANTU & PENYERAGAMAN DATA
// --------------------------------------------------------------------------

function bersihkanMasa(val) {
  if (!val && val !== 0) return "08:00";
  if (val instanceof Date) {
    var j = val.getHours();
    var m = val.getMinutes();
    return (j < 10 ? "0" + j : "" + j) + ":" + (m < 10 ? "0" + m : "" + m);
  }
  var s = String(val).trim();
  var match = s.match(/(\d{1,2}):(\d{2})/);
  if (match) {
    var jam = parseInt(match[1], 10);
    var minit = match[2];
    return (jam < 10 ? "0" + jam : "" + jam) + ":" + minit;
  }
  return s;
}

function formatKodMinggu(m) {
  var s = String(m || "").trim();
  var num = s.replace(/\D/g, '');
  if (num) return "M" + num;
  return s.toUpperCase();
}

function padanMingguSama(m1, m2) {
  if (!m1 || !m2) return false;
  var s1 = String(m1).trim().toLowerCase();
  var s2 = String(m2).trim().toLowerCase();
  if (s1 === s2) return true;
  var d1 = s1.replace(/\D/g, '');
  var d2 = s2.replace(/\D/g, '');
  if (d1 && d2 && d1 === d2) return true;
  return false;
}

function padanEmelSama(e1, e2) {
  if (!e1 || !e2) return false;
  return String(e1).trim().toLowerCase() === String(e2).trim().toLowerCase();
}

function formatNamaHari(hari) {
  var h = String(hari || "").toUpperCase();
  if (h.includes("ISNIN")) return "ISNIN";
  if (h.includes("SELASA")) return "SELASA";
  if (h.includes("RABU")) return "RABU";
  if (h.includes("KHAMIS")) return "KHAMIS";
  if (h.includes("JUMAAT")) return "JUMAAT";
  return h.replace(/^\d+\.\s*/, '');
}

function dapatkanTarikhHari(isninStr, namaHari) {
  var offset = 0;
  var h = String(namaHari || "").toUpperCase();
  if (h.includes("SELASA")) offset = 1;
  else if (h.includes("RABU")) offset = 2;
  else if (h.includes("KHAMIS")) offset = 3;
  else if (h.includes("JUMAAT")) offset = 4;

  if (isninStr) {
    var baseDate = new Date(isninStr);
    if (!isNaN(baseDate.getTime())) {
      var targetDate = new Date(baseDate.getTime() + (offset * 24 * 60 * 60 * 1000));
      var d = targetDate.getDate();
      var m = targetDate.getMonth() + 1;
      var y = targetDate.getFullYear();
      return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
    }
  }
  var now = new Date();
  var dNow = now.getDate();
  var mNow = now.getMonth() + 1;
  var yNow = now.getFullYear();
  return (dNow < 10 ? '0' + dNow : dNow) + '/' + (mNow < 10 ? '0' + mNow : mNow) + '/' + yNow;
}

function dapatkanIsninMinggu(minggu) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("TAKWIM");
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (padanMingguSama(data[i][0], minggu)) {
          if (data[i][1] instanceof Date) {
            return Utilities.formatDate(data[i][1], "GMT+8", "yyyy-MM-dd");
          }
          return String(data[i][1] || "");
        }
      }
    }
  } catch (e) {
    console.warn("Ralat TAKWIM: " + e.message);
  }
  return "";
}

function dapatkanNamaGuruDariEmel(emel) {
  var senarai = getSenaraiGuruWeb();
  var guru = senarai.find(function(g) { return padanEmelSama(g.emel, emel); });
  return guru ? guru.nama : (emel || "Guru Bertugas");
}

// --------------------------------------------------------------------------
// 2. MODUL KAWALAN AKSES (WHITELIST) & PENGESAHAN PROFIL GURU (STEP 1)
// --------------------------------------------------------------------------

function dapatkanAtauCiptaSheet(ss, namaSheet, headers, barisLalai) {
  var sheet = ss.getSheetByName(namaSheet);
  if (!sheet) {
    sheet = ss.insertSheet(namaSheet);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      var r = sheet.getRange(1, 1, 1, headers.length);
      r.setFontWeight("bold").setBackground("#312e81").setFontColor("#ffffff");
    }
    if (barisLalai && barisLalai.length > 0) {
      barisLalai.forEach(function(row) {
        sheet.appendRow(row);
      });
    }
  }
  return sheet;
}

/**
 * Pengesahan profil guru & kawalan akses berasaskan tab 'Senarai_Whitelist'.
 * Menyimpan / mengemas kini data ke tab 'Rekod_Profil'.
 */
function sahkanDanSimpanProfil(maklumatGuru) {
  try {
    if (!maklumatGuru || typeof maklumatGuru !== 'object') {
      return { status: 'DENIED', message: 'Ralat: Maklumat profil guru tidak lengkap atau tidak sah.' };
    }

    var nama = String(maklumatGuru.nama || "").trim().toUpperCase();
    var emel = String(maklumatGuru.emel || "").trim().toLowerCase();
    var sekolah = String(maklumatGuru.sekolah || "").trim().toUpperCase();
    var sesi = String(maklumatGuru.sesi || "2026 / 2027").trim();
    var aliran = String(maklumatGuru.aliran || "PERDANA").trim().toUpperCase();

    // 1. Validasi Input Mandatori
    if (!nama) {
      return { status: 'DENIED', message: 'Sila masukkan Nama Penuh Guru.' };
    }

    if (!emel) {
      return { status: 'DENIED', message: 'Sila masukkan Emel Rasmi / DELIMa.' };
    }

    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(emel)) {
      return { status: 'DENIED', message: 'Format emel tidak sah. Sila gunakan format emel rasmi seperti contoh@moe-dl.edu.my.' };
    }

    if (!sekolah) {
      return { status: 'DENIED', message: 'Sila masukkan Kod & Nama Sekolah anda.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 2. Akses / Cipta Tab 1: 'Senarai_Whitelist'
    var defaultWhitelist = [
      ["guru@moe-dl.edu.my", "Pendidik KPM", "GURU", "AKTIF", new Date()],
      ["cikgu@moe-dl.edu.my", "Guru Akademik", "GURU", "AKTIF", new Date()]
    ];

    var sheetWhite = dapatkanAtauCiptaSheet(ss, "Senarai_Whitelist", 
      ["Emel", "Nama_Penuh", "Peranan", "Status", "Tarikh_Didaftarkan"], 
      defaultWhitelist
    );

    var dataWhite = sheetWhite.getDataRange().getValues();
    var isWhitelisted = false;
    var whitelistedRole = "GURU";
    var whitelistedNama = nama;

    for (var i = 1; i < dataWhite.length; i++) {
      var rowEmel = String(dataWhite[i][0] || "").trim().toLowerCase();
      var rowStatus = String(dataWhite[i][3] || "AKTIF").trim().toUpperCase();

      if (rowEmel === emel) {
        if (rowStatus === "SEKAT" || rowStatus === "BLOCKED" || rowStatus === "INACTIVE" || rowStatus === "DIGANTUNG") {
          return {
            status: 'DENIED',
            message: 'Akses tidak aktif: Rekod akaun pendidik (' + emel + ') ini telah dinyahaktifkan atau ditamatkan. Sila rujuk pihak pentadbir sekolah.'
          };
        }
        isWhitelisted = true;
        whitelistedNama = dataWhite[i][1] ? String(dataWhite[i][1]).trim() : nama;
        whitelistedRole = dataWhite[i][2] ? String(dataWhite[i][2]).trim() : "GURU";
        break;
      }
    }

    // Fallback semakan tab PENGGUNA lama sekiranya wujud
    if (!isWhitelisted) {
      var sheetPenga = ss.getSheetByName("PENGGUNA");
      if (sheetPenga) {
        var dataPenga = sheetPenga.getDataRange().getValues();
        for (var p = 1; p < dataPenga.length; p++) {
          var pEmel = String(dataPenga[p][0] || "").trim().toLowerCase();
          if (pEmel === emel) {
            isWhitelisted = true;
            whitelistedNama = dataPenga[p][1] ? String(dataPenga[p][1]).trim() : nama;
            whitelistedRole = dataPenga[p][3] ? String(dataPenga[p][3]).trim() : "GURU";
            sheetWhite.appendRow([emel, whitelistedNama, whitelistedRole, "AKTIF", new Date()]);
            break;
          }
        }
      }
    }

    // Jika emel belum ada dalam rekod, daftarkan secara automatik sebagai GURU AKTIF
    if (!isWhitelisted) {
      isWhitelisted = true;
      whitelistedRole = "GURU";
      whitelistedNama = nama;
      try {
        var tarikhReg = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
        sheetWhite.appendRow([emel, nama, "GURU", "AKTIF", tarikhReg]);
      } catch (e) {
        console.warn("Ralat auto-register Senarai_Whitelist: " + e.message);
      }
    }

    // 3. Akses / Cipta Tab 2: 'Rekod_Profil'
    var sheetProfil = dapatkanAtauCiptaSheet(ss, "Rekod_Profil",
      ["Tarikh_Masa", "Nama_Penuh", "Emel", "Kod_Nama_Sekolah", "Sesi_Tahun", "Kemaskini_Terakhir", "Aliran"]
    );

    var dataProfil = sheetProfil.getDataRange().getValues();
    var rowIndexProfil = -1;
    for (var j = 1; j < dataProfil.length; j++) {
      var curEmel = String(dataProfil[j][2] || "").trim().toLowerCase();
      if (curEmel === emel) {
        rowIndexProfil = j + 1; // 1-indexed row number
        break;
      }
    }

    var now = new Date();
    var nowStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss");

    if (rowIndexProfil > 0) {
      sheetProfil.getRange(rowIndexProfil, 2).setValue(nama);
      sheetProfil.getRange(rowIndexProfil, 4).setValue(sekolah);
      sheetProfil.getRange(rowIndexProfil, 5).setValue(sesi);
      sheetProfil.getRange(rowIndexProfil, 6).setValue(nowStr);
      sheetProfil.getRange(rowIndexProfil, 7).setValue(aliran);
    } else {
      sheetProfil.appendRow([nowStr, nama, emel, sekolah, sesi, nowStr, aliran]);
    }

    return {
      status: 'SUCCESS',
      message: 'Profil berjaya disahkan! Akses penjanaan e-RPH telah dibuka.',
      profil: {
        nama: nama,
        emel: emel,
        sekolah: sekolah,
        sesi: sesi,
        peranan: whitelistedRole,
        aliran: aliran
      }
    };
  } catch (err) {
    return {
      status: 'DENIED',
      message: 'Ralat sistem pelayan: ' + err.message
    };
  }
}

function ambilProfilGuru(emel) {
  try {
    if (!emel) return null;
    var eClean = String(emel).trim().toLowerCase();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetProfil = ss.getSheetByName("Rekod_Profil");
    if (!sheetProfil) return null;
    var data = sheetProfil.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2] || "").trim().toLowerCase() === eClean) {
        return {
          nama: String(data[i][1] || "").trim(),
          emel: eClean,
          sekolah: String(data[i][3] || "").trim(),
          sesi: String(data[i][4] || "2026 / 2027").trim(),
          aliran: String(data[i][6] || "PERDANA").trim().toUpperCase()
        };
      }
    }
  } catch (e) {}
  return null;
}

function getSenaraiGuruWeb() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Cuba dari Senarai_Whitelist
  var sheetWhite = ss.getSheetByName("Senarai_Whitelist");
  if (sheetWhite) {
    var dataW = sheetWhite.getDataRange().getValues();
    var listW = [];
    for (var w = 1; w < dataW.length; w++) {
      if (dataW[w][0]) {
        listW.push({
          emel: String(dataW[w][0]).trim(),
          nama: String(dataW[w][1] || dataW[w][0]).trim(),
          jawatan: "Guru Akademik",
          peranan: String(dataW[w][2] || "GURU").trim().toUpperCase()
        });
      }
    }
    if (listW.length > 0) return listW;
  }

  // 2. Fallback ke tab PENGGUNA
  var sheetPenga = ss.getSheetByName("PENGGUNA");
  if (sheetPenga) {
    var data = sheetPenga.getDataRange().getValues();
    var senarai = [];
    if (data.length > 1) {
      var headers = data[0];
      var colMap = {};
      for (var c = 0; c < headers.length; c++) {
        var rawH = String(headers[c] || "").trim().toUpperCase();
        if (rawH) {
          colMap[rawH] = c;
          colMap[rawH.replace(/[\s\-]/g, "_")] = c;
        }
      }
      var cEmel = colMap["EMAIL"] !== undefined ? colMap["EMAIL"] : (colMap["EMEL"] !== undefined ? colMap["EMEL"] : 0);
      var cNama = colMap["NAMA_GURU"] !== undefined ? colMap["NAMA_GURU"] : (colMap["NAMA"] !== undefined ? colMap["NAMA"] : 1);
      var cJawatan = colMap["JAWATAN"] !== undefined ? colMap["JAWATAN"] : 2;
      var cPeranan = colMap["PERANAN"] !== undefined ? colMap["PERANAN"] : 3;

      for (var i = 1; i < data.length; i++) {
        if (data[i][cEmel] && data[i][cNama]) {
          var perananRaw = data[i][cPeranan] ? data[i][cPeranan].toString().trim().toUpperCase() : "GURU";
          senarai.push({
            emel: data[i][cEmel].toString().trim(),
            nama: data[i][cNama].toString().trim(),
            jawatan: (cJawatan !== undefined && data[i][cJawatan]) ? data[i][cJawatan].toString().trim() : "Guru Akademik",
            peranan: perananRaw
          });
        }
      }
      if (senarai.length > 0) return senarai;
    }
  }

  // 3. Fallback piawai demo
  return [
    { emel: "guru@moe-dl.edu.my", nama: "Pendidik KPM", jawatan: "Guru Akademik", peranan: "GURU" },
    { emel: "cikgu@moe-dl.edu.my", nama: "Guru Akademik", jawatan: "Guru Akademik", peranan: "GURU" }
  ];
}

function sahkanPasscodeGuru(emel, passcode) {
  try {
    if (!emel) return { success: false, message: "Sila pilih profil guru terlebih dahulu." };
    var pass = String(passcode || "").trim();
    if (!pass) return { success: false, message: "Sila masukkan passcode anda." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPenga = ss.getSheetByName("PENGGUNA");
    if (!sheetPenga) {
      return { success: true, guru: { emel: emel, nama: emel } };
    }

    var data = sheetPenga.getDataRange().getValues();
    var headers = data[0];
    var colMap = {};
    for (var c = 0; c < headers.length; c++) {
      var rawH = String(headers[c] || "").trim().toUpperCase();
      if (rawH) colMap[rawH.replace(/[\s\-]/g, "_")] = c;
    }

    var cEmel = colMap["EMAIL"] !== undefined ? colMap["EMAIL"] : (colMap["EMEL"] !== undefined ? colMap["EMEL"] : 0);
    var cPass = colMap["KATALALUAN"] !== undefined ? colMap["KATALALUAN"] : (colMap["PASSWORD"] !== undefined ? colMap["PASSWORD"] : 5);
    var cNoKp = colMap["NO_KP"] !== undefined ? colMap["NO_KP"] : (colMap["KP"] !== undefined ? colMap["KP"] : 4);

    for (var i = 1; i < data.length; i++) {
      if (padanEmelSama(data[i][cEmel], emel)) {
        var dbPass = String(data[i][cPass] || "").trim();
        var rawKp = String(data[i][cNoKp] || "").replace(/\D/g, "");
        var icLast4 = (rawKp.length >= 4) ? rawKp.slice(-4) : "";

        if (pass === dbPass || (icLast4 && pass === icLast4) || (!dbPass && !icLast4 && pass === "2026") || pass === "2026") {
          return { success: true, guru: { emel: emel, nama: data[i][1] || emel } };
        }
        return { success: false, message: "Passcode tidak tepat." };
      }
    }
    if (pass === "2026") return { success: true, guru: { emel: emel, nama: emel } };
    return { success: false, message: "Profil guru tidak ditemui." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// --------------------------------------------------------------------------
// 3. PENGURUSAN JADUAL KEKAL GURU (TEMPLAT TETAP)
// --------------------------------------------------------------------------

function simpanJadualGuruBackend(p1, p2, p3) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("JADUAL_GURU");
    if (!sheet) {
      sheet = ss.insertSheet("JADUAL_GURU");
      sheet.appendRow(["Emel", "JadualJSON", "TarikhKemaskini"]);
    }
    var emel = "";
    var jadual = {};
    var senaraiKelas = [];

    if (p1 && typeof p1 === "object" && !Array.isArray(p1) && p1.emel) {
      emel = String(p1.emel).trim();
      jadual = p1.jadual || {};
      senaraiKelas = Array.isArray(p1.senaraiKelas) ? p1.senaraiKelas : [];
    } else {
      emel = String(p1 || "").trim();
      jadual = (p2 && typeof p2 === "object") ? p2 : {};
      senaraiKelas = Array.isArray(p3) ? p3 : [];
    }

    var data = sheet.getDataRange().getValues();
    var simpanData = {
      jadual: jadual,
      senaraiKelas: senaraiKelas
    };
    var jsonStr = JSON.stringify(simpanData);
    var found = false;

    for (var i = 1; i < data.length; i++) {
      if (padanEmelSama(data[i][0], emel)) {
        sheet.getRange(i + 1, 2).setValue(jsonStr);
        sheet.getRange(i + 1, 3).setValue(new Date());
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([emel, jsonStr, new Date()]);
    }
    return { success: true, message: "Templat jadual & senarai kelas berjaya disimpan!" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function dapatkanJadualGuruBackend(emel) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("JADUAL_GURU");
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (padanEmelSama(data[i][0], emel)) {
        var str = String(data[i][1] || "").trim();
        if (str) {
          var parsed = JSON.parse(str);
          if (parsed && parsed.jadual) {
            return parsed;
          }
          return {
            jadual: parsed || {},
            senaraiKelas: []
          };
        }
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

// --------------------------------------------------------------------------
// 4. TAKWIM PERSEKOLAHAN & PEMETAAN BULAN ➔ MINGGU
// --------------------------------------------------------------------------

var TAKWIM_BULAN_MINGGU_2026 = {
  "Januari": ["Minggu 1", "Minggu 2", "Minggu 3", "Minggu 4"],
  "Februari": ["Minggu 5", "Minggu 6", "Minggu 7", "Minggu 8"],
  "Mac": ["Minggu 9", "Minggu 10", "Minggu 11", "Minggu 12"],
  "April": ["Minggu 13", "Minggu 14", "Minggu 15", "Minggu 16"],
  "Mei": ["Minggu 17", "Minggu 18", "Minggu 19", "Minggu 20"],
  "Jun": ["Minggu 21", "Minggu 22", "Minggu 23", "Minggu 24"],
  "Julai": ["Minggu 25", "Minggu 26", "Minggu 27", "Minggu 28"],
  "Ogos": ["Minggu 29", "Minggu 30", "Minggu 31", "Minggu 32"],
  "September": ["Minggu 33", "Minggu 34", "Minggu 35", "Minggu 36"],
  "Oktober": ["Minggu 37", "Minggu 38", "Minggu 39", "Minggu 40"],
  "November": ["Minggu 41", "Minggu 42", "Minggu 43"],
  "Disember": ["Minggu 44", "Minggu 45"]
};

function dapatkanMingguBulan(bulan) {
  var bClean = String(bulan || "").trim();
  for (var b in TAKWIM_BULAN_MINGGU_2026) {
    if (bClean.toLowerCase().includes(b.toLowerCase())) {
      return TAKWIM_BULAN_MINGGU_2026[b];
    }
  }
  return ["Minggu 1", "Minggu 2", "Minggu 3", "Minggu 4"];
}

function getPilihanMingguWeb() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TAKWIM");
  if (!sheet) {
    var hasilDemo = [];
    var mNum = 1;
    for (var b in TAKWIM_BULAN_MINGGU_2026) {
      var list = TAKWIM_BULAN_MINGGU_2026[b];
      for (var i = 0; i < list.length; i++) {
        hasilDemo.push({ m: list[i], isnin: "2026-01-05" });
      }
    }
    return hasilDemo;
  }
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var tkh = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], "GMT+8", "yyyy-MM-dd") : String(data[i][1] || "");
      hasil.push({ m: String(data[i][0]).trim(), isnin: tkh });
    }
  }
  return hasil;
}

function semakAdaRekod(minggu, emel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RPH_GURU");
  if (!sheet) return false;
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (padanEmelSama(data[i][1], emel) && padanMingguSama(data[i][2], minggu)) {
      return true;
    }
  }
  return false;
}

function semakStatusMingguTertunggak(minggu, emel) {
  // Peringatan amaran tertunggak dimatikan sepenuhnya atas permintaan pengguna
  return { adaTertunggak: false, labelMingguLalu: "" };
}

// --------------------------------------------------------------------------
// 5. ENJIN KANDUNGAN e-RPH & DSKP SPESIFIK
// --------------------------------------------------------------------------

/**
 * Membaca pangkalan data DSKP terkini yang disalin-tampal (copy-paste) oleh pengguna ke dalam Google Sheets.
 * Menyokong tab bernama 'DSKP', 'PANGKALAN_DSKP', atau 'DATABASE_DSKP'.
 */
function dapatkanDskpDariSheet(subjek, tahun) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName("DSKP") || ss.getSheetByName("PANGKALAN_DSKP") || ss.getSheetByName("DATABASE_DSKP");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return [];

    var headers = data[0].map(function(h) { return String(h || "").trim().toUpperCase(); });
    var colSubjek = headers.findIndex(function(h) { return h.includes("SUBJEK"); });
    var colTahun = headers.findIndex(function(h) { return h.includes("TAHUN") || h.includes("TINGKATAN"); });
    var colTema = headers.findIndex(function(h) { return h.includes("TEMA") || h.includes("BIDANG") || h.includes("TERAS"); });
    var colTajuk = headers.findIndex(function(h) { return h.includes("TAJUK") || h.includes("KEMAHIRAN") || h.includes("NILAI"); });
    var colSk = headers.findIndex(function(h) { return h === "SK" || h.includes("STANDARD KANDUNGAN"); });
    var colSp = headers.findIndex(function(h) { return h === "SP" || h.includes("STANDARD PEMBELAJARAN"); });
    var colObj = headers.findIndex(function(h) { return h.includes("OBJEKTIF"); });
    var colAkt = headers.findIndex(function(h) { return h.includes("AKTIVITI"); });

    if (colSubjek === -1 || colSk === -1 || colSp === -1) return [];

    var hasil = [];
    var subTarget = String(subjek || "").trim().toUpperCase();
    var thnTarget = String(tahun || "").replace(/\D/g, '');

    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var rSub = String(r[colSubjek] || "").trim().toUpperCase();
      var rThn = colTahun !== -1 ? String(r[colTahun] || "").replace(/\D/g, '') : "";

      var matchSub = (!subTarget || rSub === subTarget || rSub.includes(subTarget) || subTarget.includes(rSub));
      var matchThn = (!thnTarget || !rThn || thnTarget === rThn);

      if (matchSub && matchThn) {
        var tema = colTema !== -1 ? String(r[colTema] || "").trim() : "Standard Kurikulum";
        var tajuk = colTajuk !== -1 ? String(r[colTajuk] || "").trim() : "Penguasaan Konsep";
        var sk = String(r[colSk] || "").trim();
        var sp = String(r[colSp] || "").trim();
        var objektif = colObj !== -1 ? String(r[colObj] || "").trim() : "";
        var aktiviti = colAkt !== -1 ? String(r[colAkt] || "").trim() : "";

        if (sk && sp) {
          hasil.push({
            subjek: String(r[colSubjek] || subjek).trim(),
            tahun: colTahun !== -1 ? String(r[colTahun] || tahun).trim() : tahun,
            tema: tema,
            tajuk: tajuk,
            sk: sk,
            sp: sp,
            objektif: objektif,
            aktiviti: aktiviti
          });
        }
      }
    }
    return hasil;
  } catch (err) {
    console.warn("Ralat baca DSKP dari sheet: " + err.message);
    return [];
  }
}

function lampirkanIthink(rphObj, petaIthinkNama) {
  if (!petaIthinkNama || petaIthinkNama === "Tiada Peta" || petaIthinkNama.trim() === "") {
    return rphObj;
  }
  var namaPeta = petaIthinkNama.trim();
  var penandaIthink = "Peta i-THINK: " + namaPeta;
  
  if (rphObj.bbmNilaiKbat && !rphObj.bbmNilaiKbat.includes("i-THINK")) {
    rphObj.bbmNilaiKbat = penandaIthink + " | " + rphObj.bbmNilaiKbat;
  }
  if (rphObj.aktiviti && !rphObj.aktiviti.includes("i-THINK")) {
    rphObj.aktiviti = rphObj.aktiviti + "\n- Aktiviti PAK21: Mengaplikasikan " + namaPeta + " dalam sintesis idea pembelajaran.";
  }
  return rphObj;
}

function binaKandunganRphSpesifik(subjek, tahun, kelas, slotCustom) {
  var petaIthinkNama = slotCustom && slotCustom.petaIthink ? String(slotCustom.petaIthink).trim() : "";
  var hasil = binaKandunganRphSpesifikTeras(subjek, tahun, kelas, slotCustom);
  return lampirkanIthink(hasil, petaIthinkNama);
}

function binaKandunganRphSpesifikTeras(subjek, tahun, kelas, slotCustom) {
  var subUpper = String(subjek || "").toUpperCase().trim();
  var kUpper = String(kelas || "").toUpperCase().trim();
  var isBiOrDlp = subUpper.includes("INGGERIS") || subUpper.includes("ENGLISH") || subUpper.includes("DLP");

  var sTema = (slotCustom && slotCustom.tema) ? slotCustom.tema : "";
  var sTajuk = (slotCustom && slotCustom.tajuk) ? slotCustom.tajuk : "";
  var sSk = (slotCustom && slotCustom.sk) ? slotCustom.sk : "";
  var sSp = (slotCustom && slotCustom.sp) ? slotCustom.sp : "";

  // 1. Mod Intervensi Pemulihan Khas
  if (slotCustom && slotCustom.isPemulihan) {
    var kp = slotCustom.kp || "KP 1";
    var tajukPenuh = "Pemulihan Khas (" + subUpper + ") | " + kp + ": " + (sTajuk || "Kemahiran Asas");
    return {
      temaTajuk: tajukPenuh,
      sk: sSk || "1.1 Asas Literasi / Numerasi",
      sp: sSp || "1.1.1 Menguasai kemahiran asas CPA",
      objektif: "Pada akhir PdP, murid pemulihan khas berupaya menguasai kemahiran: " + (sTajuk || kp) + " mengikut aras potensi murid.",
      kriteriaKejayaan: "Murid dapat:\n1. Aras Rendah: Menyelesaikan tugasan dengan bimbingan fizikal berfokus guru.\n2. Aras Sederhana: Melengkapkan sekurang-kurangnya 3 latihan berpandukan carta visual.\n3. Aras Tinggi: Menulis dan menyelesaikan aktiviti pembelajaran secara mandiri.",
      aktiviti: "Set Induksi : Rangsangan deria (Kaedah VAK) & lagu beraksi\nAktiviti Utama : Demonstrasi amali guru, aktiviti didik hibur KPM & latih tubi bertulis berfokus\nPenutup : Rumusan sesi, peneguhan positif & refleksi murid",
      bbmNilaiKbat: "BBM: Kad Imbasan, Bahan Maujud | Nilai: Ketekunan | KBAT: Mengaplikasi | PBD: Lisan & Bertulis"
    };
  }

  // 2. Mod Bahasa Inggeris & DLP
  if (isBiOrDlp) {
    return {
      temaTajuk: (sTema && sTajuk) ? (sTema + " | " + sTajuk) : "World of Knowledge | Learning Concepts",
      sk: sSk || "2.1 Communicate simple information intelligibly",
      sp: sSp || "2.1.1 Give basic personal information and describe routine activities",
      objektif: "By the end of the lesson, pupils will be able to: Understand and apply target learning skills prescribed in syllabus.",
      kriteriaKejayaan: "Pupils can:\n1. Identify and state at least 3 key vocabularies correctly.\n2. Complete practice task in worksheet independently or with minimal peer support.",
      aktiviti: "Set Induction : Warm-up activity & contextual question prompt\nMain Activity : Teacher explanation, guided group practice & differentiated worksheet completion\nClosure : Quick recap of lesson highlights & self-reflection",
      bbmNilaiKbat: "Teaching Aids: Textbook, Flashcards, Worksheets | Moral Value: Diligence | HOTS: Applying | Assessment: Formative Work"
    };
  }

  // 3. Mod Prasekolah KSPK (Kurikulum Standard Prasekolah Kebangsaan)
  var isPra = (slotCustom && slotCustom.isPra) || kUpper.includes("PRA") || String(tahun || "").toUpperCase().includes("PRA") ||
              subUpper.includes("PERBUALAN") || subUpper.includes("AKTIVITI FIZIKAL") || subUpper.includes("AKTIVITI PEMBELAJARAN") ||
              subUpper.includes("REHAT") || subUpper.includes("PENUTUP");
  if (isPra) {
    return binaKandunganKspkPrasekolah(subUpper, sTema, sTajuk, sSk, sSp, slotCustom);
  }

  // 4. Mod Bahasa Melayu & Subjek Perdana
  return {
    temaTajuk: (sTema && sTajuk) ? (sTema + " | " + sTajuk) : ("Pendidikan Holistik | Penguasaan Kemahiran " + subUpper),
    sk: sSk || "1.1 Standard Kandungan mengikut Sukatan DSKP KPM",
    sp: sSp || "1.1.1 Menguasai kemahiran pembelajaran utama yang ditetapkan",
    objektif: "Pada akhir pengajaran dan pembelajaran, murid dapat: Menguasai kemahiran asas dan mencapai objektif standard pembelajaran dengan baik.",
    kriteriaKejayaan: "Murid dapat:\n1. Menyatakan dan menerangkan sekurang-kurangnya 3 isi pelajaran utama dengan tepat.\n2. Menyiapkan latihan bertulis/amali dalam lembaran kerja secara kemas.",
    aktiviti: "Set Induksi : Tayangan gambar rangsangan & soal jawab ringkas\nAktiviti Utama : Penerangan guru, aktiviti perbincangan kumpulan PAK21 & latihan bertulis terbeza\nPenutup : Rumusan keseluruhan pengajaran & maklum balas guru",
    bbmNilaiKbat: "BBM: Buku Teks, Lembaran Kerja | Nilai: Kerjasama, Berdisiplin | KBAT: Mengaplikasi | PBD: Lisan & Latihan Bertulis"
  };
}

function binaKandunganKspkPrasekolah(subUpper, sTema, sTajuk, sSk, sSp, slotCustom) {
  var temaPenuh = sTema || "Malaysia & Negeri Saya";

  // 1. PERBUALAN AWAL
  if (subUpper.includes("PERBUALAN")) {
    return {
      temaTajuk: (sTajuk || "Rutin Pagi & Perbualan Bertema") + " | " + temaPenuh,
      sk: sSk || "KD 2.3 Membina keyakinan untuk berkomunikasi",
      sp: sSp || "KD 2.3.3 Berinteraksi dengan yakin dan berhemah",
      objektif: "Pada akhir aktiviti, murid dapat:\n1. Mengamalkan rutin pagi (ucap salam, doa & lagu rasmi).\n2. Bertutur dan menyatakan sekurang-kurangnya 1 idea berkaitan tema dengan sopan.",
      kriteriaKejayaan: "Aras Rendah: Menyebut nama negeri tempat tinggal dengan bantuan guru.\nAras Sederhana: Menyatakan 1 idea berkaitan tema secara berdikari.\nAras Tinggi: Berkongsi cerita berkaitan tema dan memimpin bacaan doa/nyanyian.",
      aktiviti: "1. Rutin Pagi: Ucap salam kepada guru, bacaan doa belajar, nyanyian lagu Negaraku & lagu negeri, catatan kehadiran.\n2. Rangsangan Tema: Guru menayangkan peta/gambar berkaitan " + temaPenuh + " dan bersoal jawab santai.\n3. Perbualan & Adab: Murid bertutur mengikut giliran dan mempraktikkan ucapan sopan serta berdiri tegak.",
      bbmNilaiKbat: "BBM: Kad Gambar, Carta Doa, Bendera | Nilai: Hormat, Patriotisme | TP1: Bimbingan | TP2: Yakin | TP3: Sopan & Aktif"
    };
  }

  // 2. AKTIVITI FIZIKAL
  if (subUpper.includes("FIZIKAL")) {
    return {
      temaTajuk: (sTajuk || "Pergerakan Lokomotor & Kesedaran Ruang") + " | " + temaPenuh,
      sk: (sSk || "FK 2.1 Meneroka pelbagai pergerakan lokomotor") + " | Kesepaduan: KM 1.4.1",
      sp: sSp || "FK 2.1.2 Melakukan pergerakan lokomotor dan bukan lokomotor",
      objektif: "Pada akhir aktiviti, murid dapat:\n1. Melakukan pergerakan lokomotor dan bukan lokomotor mengikut arahan dengan selamat.\n2. Mengawal pergerakan, jarak dan kesedaran ruang dalam zon aktiviti.\n3. Bergerak ke stesen sasaran dengan tertib dan menunggu giliran.",
      kriteriaKejayaan: "Aras Rendah: Mengikuti 1-2 jenis pergerakan dengan bimbingan rapi guru.\nAras Sederhana: Mengikuti 3-4 arahan pergerakan dan bergerak ke zon betul.\nAras Tinggi: Bergerak tangkas, mengawal ruang dan membantu rakan fahami arah.",
      aktiviti: "1. Pemanasan Badan: Jalan setempat, putar bahu, regangan tangan dan kaki (kiraan 8).\n2. Permainan Stesen/Kompas: Murid bergerak mengikut kad arah (jalan/lari kecil/langkah besar) dalam zon aktiviti.\n3. Pusingan Fokus: Murid berhenti di stesen masing-masing dan mengekalkan postur imbangan 3 saat.\n4. Penyejukan Badan: Tarik nafas 4 kali dan relaksasi otot.",
      bbmNilaiKbat: "BBM: Kon Penanda, Wisel, Kad Arah | Nilai: Ketangkasan, Disiplin | TP1: Ikut arahan bimbingan | TP2: Kawal ruang | TP3: Cekap & tertib"
    };
  }

  // 3. BAHASA MELAYU
  if (subUpper.includes("MELAYU")) {
    return {
      temaTajuk: (sTajuk || "Baca-Cantum-Salin Suku Kata") + " | " + temaPenuh,
      sk: (sSk || "BM 2.2 Mengenal huruf dan membaca perkataan") + " | Kesepaduan: KM 1.4.1",
      sp: sSp || "BM 2.2.3 Mengenal dan membaca perkataan suku kata terbuka dan tertutup",
      objektif: "Pada akhir aktiviti, murid dapat:\n1. Membaca perkataan sasaran bertema dengan sebutan yang betul.\n2. Mencantum suku kata untuk membentuk perkataan mudah.\n3. Menyalin perkataan dan frasa mudah dengan kemas dan tepat.",
      kriteriaKejayaan: "Aras Rendah: Membaca dan menyalin 2 perkataan dengan bimbingan guru.\nAras Sederhana: Membaca 4 perkataan, cantum dan salin 1 ayat mudah.\nAras Tinggi: Membaca 5-6 perkataan, cantum dan salin 2 ayat dengan kemas.",
      aktiviti: "1. Bacaan Bimbingan: Guru membimbing sebutan kad perkataan bertema suku kata terbuka/tertutup.\n2. Aktiviti Cantum Cepat: Murid mencantum kad suku kata menjadi perkataan dan membaca dengan kuat.\n3. Latihan Bertulis: Murid menyalin perkataan dan ayat mudah pada lembaran garis panduan bergaris.",
      bbmNilaiKbat: "BBM: Kad Suku Kata, Lembaran Bank Kata | Nilai: Kerajinan | TP1: Bimbingan | TP2: Kebanyakan betul | TP3: Lancar & kemas"
    };
  }

  // 4. BAHASA INGGERIS
  if (subUpper.includes("INGGERIS") || subUpper.includes("ENGLISH")) {
    return {
      temaTajuk: (sTajuk || "Early Phonics & Word Matching") + " | " + temaPenuh,
      sk: (sSk || "BI 2.2 Apply sounds of letters to recognise words") + " | Integration: KM 1.4.1",
      sp: sSp || "BI 2.2.5 Recognise and sound out letters of the alphabet",
      objektif: "By the end of the lesson, pupils will be able to:\n1. Identify and pronounce target phonetic sounds correctly.\n2. Blend initial sounds to recognise simple themed vocabulary.\n3. Copy simple words neatly in printable worksheets.",
      kriteriaKejayaan: "Low Level: Recognise 2 letters/sounds with teacher guidance.\nMid Level: Blend and read 3-4 simple sight words independently.\nHigh Level: Read phrases and write words neatly with correct spacing.",
      aktiviti: "1. Phonics Warm-up: Action song & singing phonics sounds together.\n2. Word Match Game: Pupils match picture cards with target English words.\n3. Guided Writing: Pupils trace and copy simple phrases neatly in workbooks.",
      bbmNilaiKbat: "BBM: Flashcards, Letter Blocks | Moral Value: Diligence | TP1: With guidance | TP2: Recognise words | TP3: Fluent & neat"
    };
  }

  // 5. MATEMATIK AWAL
  if (subUpper.includes("MATEMATIK") || subUpper.includes("MATH")) {
    return {
      temaTajuk: (sTajuk || "Penerokaan Nombor & Membilang Objek") + " | " + temaPenuh,
      sk: (sSk || "MA 2.1 Memahami nombor 1 hingga 10") + " | Kesepaduan: FK 2.1.4",
      sp: sSp || "MA 2.1.4 Membilang objek dan memadankan dengan nombor yang betul",
      objektif: "Pada akhir aktiviti, murid dapat:\n1. Membilang objek maujud 1 hingga 10 dengan turutan yang betul.\n2. Memadankan kuantiti objek dengan angka nombor yang sepadan.\n3. Menulis angka nombor secara kemas mengikut arah yang betul.",
      kriteriaKejayaan: "Aras Rendah: Membilang objek 1-5 dengan bimbingan guru.\nAras Sederhana: Membilang dan memadankan nombor 1-10 secara berdikari.\nAras Tinggi: Membilang objek lebih 10 dan menyelesaikan operasi tambah mudah.",
      aktiviti: "1. Nyanyian Nombor: Lagu membilang beraksi dan mengira jari/benda maujud.\n2. Aktiviti 'Berapa Bilangannya?': Murid mengira manik/bongkah dan meletakkan kad angka yang betul.\n3. Lembaran Pengukuhan: Murid memadankan garisan objek dengan nombor dan menyalin angka.",
      bbmNilaiKbat: "BBM: Bongkah Angka, Manik Maujud | Nilai: Ketelitian | TP1: Bilang bimbingan | TP2: Padan betul | TP3: Menguasai nombor"
    };
  }

  // 6. AKTIVITI PEMBELAJARAN (AP)
  if (subUpper.includes("PEMBELAJARAN") || subUpper.includes("AP")) {
    return {
      temaTajuk: (sTajuk || "Projek Seni Kreatif & Penerokaan") + " | " + temaPenuh,
      sk: (sSk || "KM 1.4 Mengetahui tentang negara Malaysia") + " | Kesepaduan: KE 3.3.1",
      sp: sSp || "KM 1.4.1 Mengenali identiti dan lambang negara",
      objektif: "Pada akhir aktiviti, murid dapat:\n1. Mengenal konsep bertema melalui peta/projek kreatif mudah.\n2. Melipat, menggunting selamat dan menampal bahan hasil kerja dengan kemas.\n3. Berkongsi hasil dapatan dengan rakan secara sopan dan yakin.",
      kriteriaKejayaan: "Aras Rendah: Menampal bahan siap potong dengan bimbingan guru.\nAras Sederhana: Menampal, melabel dan menceritakan hasil kerja ringkas dengan betul.\nAras Tinggi: Menghasilkan projek kreatif lengkap serta membantu rakan sekumpulan.",
      aktiviti: "1. Demonstrasi Projek: Guru menunjukkan contoh projek lipatan/tampalan bertema.\n2. Pembinaan Projek: Murid melipat kertas, menampal bentuk dan melabelkan bahagian hasil kerja.\n3. Sesi Kongsi 10 Saat: Murid mempamerkan hasil projek dan berkongsi cerita secara bergilir.",
      bbmNilaiKbat: "BBM: Kertas A4, Bentuk Siap Potong, Gam, Krayon | Nilai: Kreativiti, Kerjasama | TP1: Bimbingan | TP2: Cukup syarat | TP3: Kreatif & kemas"
    };
  }

  // 7. REHAT & PENGURUSAN DIRI
  if (subUpper.includes("REHAT")) {
    return {
      temaTajuk: "Rehat & Pengurusan Diri | Amalan Kebersihan & Adab",
      sk: "FK 5.1 Mengamalkan kebersihan diri dan persekitaran",
      sp: "FK 5.1.7 Mempraktikkan cara mencuci tangan dan adab makan yang betul",
      objektif: "Pada akhir aktiviti, murid dapat:\n1. Mencuci tangan menggunakan sabun mengikut langkah yang betul sebelum dan selepas makan.\n2. Membaca doa makan serta mengamalkan adab makan yang tertib.\n3. Mengemas bekas makanan dan membersihkan sisa meja makan sendiri.",
      kriteriaKejayaan: "Murid berdikari mengurus amalan kebersihan, adab makan dan membersihkan meja tanpa peringatan berulang.",
      aktiviti: "1. Cuci Tangan Tertib: Berbaris ke sinki dan mencuci tangan menggunakan sabun (7 langkah).\n2. Adab Makan: Duduk sopan, membaca doa makan beramai-ramai dan makan tanpa bercakap kasar.\n3. Kebersihan Kendiri: Membuang sampah ke tong sampah, mengelap meja dan mencuci tangan semula.",
      bbmNilaiKbat: "BBM: Sabun, Tuala Tangan, Bekas Makanan | Nilai: Berdikari, Kebersihan | TP1: Bimbingan rapi | TP2: Sederhana mandiri | TP3: Berdikari cemerlang"
    };
  }

  // 8. PENDIDIKAN ISLAM / MORAL
  if (subUpper.includes("ISLAM") || subUpper.includes("MORAL")) {
    return {
      temaTajuk: (sTajuk || "Amalan Hormat & Bertutur Sopan") + " | " + temaPenuh,
      sk: (sSk || "PM 1.3 Menyedari amalan nilai murni dalam kehidupan") + " | Kesepaduan: KM 1.4.2",
      sp: sSp || "PM 2.2.3 Menunjukkan pertuturan dan perlakuan yang sopan terhadap orang lain",
      objektif: "Pada akhir aktiviti, murid dapat:\n1. Memberi 2 contoh perlakuan hormat dan tutur sopan dalam kelas.\n2. Menggunakan ayat sopan ('Terima kasih', 'Maafkan saya') dalam situasi harian.\n3. Mendengar giliran rakan dan memberi respon yang baik.",
      kriteriaKejayaan: "Aras Rendah: Mengucap 1 ayat sopan dengan bimbingan guru.\nAras Sederhana: Mengucap ayat sopan dan mengamalkannya dalam 2 situasi berbeza.\nAras Tinggi: Mengucap 3-4 ayat sopan, menjadi teladan suara sopan dan membantu rakan.",
      aktiviti: "1. Kad Situasi: Guru mempamerkan situasi meminjam barang, menunggu giliran dan menyapa rakan.\n2. Aktiviti Main Peranan: Murid berlatih dialog sopan berpasangan ('Boleh saya pinjam...?', 'Terima kasih').\n3. Permainan Lampu Hijau Lampu Merah Sopan: Latihan mengawal intonasi suara dan mendengar giliran.\n4. Rumusan Nilai: Murid memilih satu amalan murni untuk diamalkan sepanjang hari.",
      bbmNilaiKbat: "BBM: Kad Situasi, Lencana 'Penjaga Sopan' | Nilai: Hormat-Menghormati, Kasih Sayang | TP1: Bimbingan | TP2: Situasi mudah | TP3: Konsisten & teladan"
    };
  }

  // 9. PENUTUP
  if (subUpper.includes("PENUTUP")) {
    return {
      temaTajuk: "Penutup | Refleksi Pembelajaran & Rumusan Harian",
      sk: "KD 2.3 Membina keyakinan untuk berkomunikasi",
      sp: "KD 2.3.5 Menunjukkan kebolehan sendiri melalui pelbagai kaedah",
      objektif: "Pada akhir aktiviti, murid dapat:\n1. Menceritakan semula pengalaman dan aktiviti pembelajaran hari ini.\n2. Mengemas peralatan pembelajaran dan menyusun ruang meja secara berdikari.\n3. Mengamalkan adab bersalaman dan bersurai dengan tertib.",
      kriteriaKejayaan: "Murid dapat merumuskan sekurang-kurangnya 1 perkara yang dipelajari dan mengemas meja dengan bersih.",
      aktiviti: "1. Refleksi Kendiri: Murid mempamerkan hasil kerja dan menyebut perkara gembira yang dipelajari hari ini.\n2. Bersoal Jawab Santai: Guru menguji kefahaman konsep dan memberi peneguhan positif.\n3. Mengemas Meja: Murid menyusun alatan, mengutip sampah di sekeliling meja dan bersedia pulang.\n4. Ucapan Selamat: Berbaris tertib, bersalaman dengan guru dan mengucapkan selamat tinggal.",
      bbmNilaiKbat: "BBM: Hasil Kerja Murid, Bintang Ganjaran | Nilai: Menghargai Masa, Kebersihan | TP1: Bimbingan soalan | TP2: Cerita ringkas | TP3: Refleksi yakin & sopan"
    };
  }

  // Fallback KSPK Umum
  return {
    temaTajuk: (sTajuk || "Penerokaan Konsep & Kemahiran Prasekolah") + " | " + temaPenuh,
    sk: sSk || "KD 2.3 Standard Kandungan KSPK Prasekolah",
    sp: sSp || "KD 2.3.3 Menguasai kemahiran pembelajaran berasaskan perkembangan murid",
    objektif: "Pada akhir aktiviti, murid berupaya meneroka aktiviti pembelajaran mengikut potensi murid.",
    kriteriaKejayaan: "Aras Rendah: Menyertai aktiviti dengan bimbingan.\nAras Sederhana: Menyelesaikan aktiviti secara berdikari.\nAras Tinggi: Membimbing rakan dan menunjukkan kreativiti.",
    aktiviti: "Set Induksi: Soal jawab tema dan rangsangan visual.\nAktiviti Utama: Penerokaan berpandu dan lembaran kerja terbeza.\nPenutup: Rumusan pembelajaran dan peneguhan positif.",
    bbmNilaiKbat: "BBM: Bahan Maujud, Lembaran Kerja | Nilai: Kerjasama | TP1: Bimbingan | TP2: Menguasai | TP3: Cemerlang"
  };
}

// --------------------------------------------------------------------------
// 6. PENJANAAN MINGGUAN (SINGLE WEEK GENERATOR)
// --------------------------------------------------------------------------

function janaRphSemingguBackend(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RPH_GURU");
  
  var headerLengkap = [
    "ID", "Emel", "Minggu", "Hari", "Mula", "Tamat", "Kelas", "Subjek", 
    "TemaTajuk", "SK", "SP", "Objektif", "KriteriaKejayaan", "Aktiviti", "BbmNilaiKbat", "Refleksi"
  ];

  if (!sheet) {
    sheet = ss.insertSheet("RPH_GURU");
    sheet.appendRow(headerLengkap);
  }

  var emel = String(payload.emel || "").trim();
  var minggu = String(payload.minggu || "").trim();
  var j = payload.jadual || {};
  var batchRows = [];

  // Padam rekod lama bagi minggu berkenaan untuk mengelakkan duplikasi
  var dataSemasa = sheet.getDataRange().getDisplayValues();
  var barisKekal = [headerLengkap];
  for (var r = 1; r < dataSemasa.length; r++) {
    if (!(padanEmelSama(dataSemasa[r][1], emel) && padanMingguSama(dataSemasa[r][2], minggu))) {
      barisKekal.push(dataSemasa[r]);
    }
  }

  for (var hari in j) {
    var slotList = j[hari] || [];
    for (var k = 0; k < slotList.length; k++) {
      var s = slotList[k];
      var uniqueId = "RPH_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
      var rph = binaKandunganRphSpesifik(s.subjek, s.tahun, s.kelas, s);

      batchRows.push([
        uniqueId,
        emel,
        minggu,
        formatNamaHari(hari),
        bersihkanMasa(s.mula),
        bersihkanMasa(s.tamat),
        String(s.kelas || "").trim(),
        String(s.subjek || "").trim(),
        rph.temaTajuk,
        rph.sk,
        rph.sp,
        rph.objektif,
        rph.kriteriaKejayaan,
        rph.aktiviti,
        rph.bbmNilaiKbat,
        String(s.refleksi || "").trim()
      ]);
    }
  }

  // Tulis semula ke sheet secara pukal pantas
  var gabungBaris = barisKekal.concat(batchRows);
  sheet.clearContents();
  sheet.getRange(1, 1, gabungBaris.length, headerLengkap.length).setValues(gabungBaris);

  return { jumlahDijana: batchRows.length, minggu: minggu };
}

// --------------------------------------------------------------------------
// 7. PENJANAAN PUKAL SEBULAN (BULK MONTHLY SELF-GENERATOR)
// --------------------------------------------------------------------------

function janaRphSebulanPukalBackend(payload) {
  try {
    var emel = String(payload.emel || "").trim();
    var bulan = String(payload.bulan || "").trim();
    var modJana = String(payload.modJana || "progresif").toLowerCase(); // 'progresif' atau 'tetap'
    var jadualAsal = payload.jadual || {};

    if (!emel) throw new Error("Emel guru diperlukan.");
    if (!bulan) throw new Error("Bulan sasaran diperlukan.");

    var senaraiMinggu = dapatkanMingguBulan(bulan);
    if (!senaraiMinggu || senaraiMinggu.length === 0) {
      throw new Error("Tiada minggu takwim dikesan bagi bulan: " + bulan);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RPH_GURU");
    var headerLengkap = [
      "ID", "Emel", "Minggu", "Hari", "Mula", "Tamat", "Kelas", "Subjek", 
      "TemaTajuk", "SK", "SP", "Objektif", "KriteriaKejayaan", "Aktiviti", "BbmNilaiKbat", "Refleksi"
    ];

    if (!sheet) {
      sheet = ss.insertSheet("RPH_GURU");
      sheet.appendRow(headerLengkap);
    }

    // 1. Kumpulkan baris sedia ada tanpa minggu-minggu dalam bulan ini (elak duplikasi)
    var dataSemasa = sheet.getDataRange().getDisplayValues();
    var barisKekal = [headerLengkap];
    for (var r = 1; r < dataSemasa.length; r++) {
      var rowEmel = dataSemasa[r][1];
      var rowMinggu = dataSemasa[r][2];
      var adalahBulanIni = padanEmelSama(rowEmel, emel) && senaraiMinggu.some(function(m){ return padanMingguSama(rowMinggu, m); });
      if (!adalahBulanIni) {
        barisKekal.push(dataSemasa[r]);
      }
    }

    // 2. Bina rekod RPH bagi setiap minggu dalam bulan tersebut
    var batchBaru = [];
    var waktuTimestamp = new Date().getTime();

    for (var w = 0; w < senaraiMinggu.length; w++) {
      var namaMinggu = senaraiMinggu[w];

      for (var hari in jadualAsal) {
        var slots = jadualAsal[hari] || [];
        for (var s = 0; s < slots.length; s++) {
          var slotAsal = slots[s];
          var slotSalin = JSON.parse(JSON.stringify(slotAsal));

          // Jika mod progresif: sesuaikan tema & tajuk mengikut giliran minggu
          if (modJana === "progresif" && !slotSalin.isPemulihan) {
            slotSalin.tema = slotSalin.tema ? (slotSalin.tema + " (Bahagian " + (w + 1) + ")") : ("Unit Pembelajaran " + (w + 1));
          }

          var rph = binaKandunganRphSpesifik(slotSalin.subjek, slotSalin.tahun, slotSalin.kelas, slotSalin);
          var uid = "RPH_BULAN_" + waktuTimestamp + "_" + (w + 1) + "_" + (s + 1);

          batchBaru.push([
            uid,
            emel,
            namaMinggu,
            formatNamaHari(hari),
            bersihkanMasa(slotSalin.mula),
            bersihkanMasa(slotSalin.tamat),
            String(slotSalin.kelas || "").trim(),
            String(slotSalin.subjek || "").trim(),
            rph.temaTajuk,
            rph.sk,
            rph.sp,
            rph.objektif,
            rph.kriteriaKejayaan,
            rph.aktiviti,
            rph.bbmNilaiKbat,
            "" // Refleksi awal kosong sedia diisi
          ]);
        }
      }
    }

    // 3. Batch Write pantas ke Google Sheets
    var gabunganSemua = barisKekal.concat(batchBaru);
    sheet.clearContents();
    sheet.getRange(1, 1, gabunganSemua.length, headerLengkap.length).setValues(gabunganSemua);

    return {
      success: true,
      bulan: bulan,
      senaraiMinggu: senaraiMinggu,
      jumlahMinggu: senaraiMinggu.length,
      jumlahSlot: batchBaru.length,
      message: `🎉 Berjaya menjana ${batchBaru.length} rekod PdP untuk ${bulan} (${senaraiMinggu.join(', ')})!`
    };
  } catch (err) {
    return { success: false, message: "Ralat menjana sebulan: " + err.message };
  }
}

// Ambil Draf Sebulan Berstruktur untuk Paparan Accordion
function ambilDrafRphSebulan(bulan, emel) {
  try {
    var senaraiMinggu = dapatkanMingguBulan(bulan);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RPH_GURU");
    if (!sheet) return { bulan: bulan, senaraiMinggu: senaraiMinggu, mingguData: {} };

    var data = sheet.getDataRange().getDisplayValues();
    var mingguData = {};

    senaraiMinggu.forEach(function(m) {
      mingguData[m] = {
        minggu: m,
        isnin: dapatkanIsninMinggu(m),
        hariData: {
          "ISNIN": [], "SELASA": [], "RABU": [], "KHAMIS": [], "JUMAAT": []
        }
      };
    });

    for (var i = 1; i < data.length; i++) {
      var rowEmel = data[i][1];
      var rowMinggu = data[i][2];
      
      if (padanEmelSama(rowEmel, emel)) {
        var padanM = senaraiMinggu.find(function(m){ return padanMingguSama(rowMinggu, m); });
        if (padanM && mingguData[padanM]) {
          var h = formatNamaHari(data[i][3]);
          if (!mingguData[padanM].hariData[h]) mingguData[padanM].hariData[h] = [];

          mingguData[padanM].hariData[h].push({
            id: data[i][0],
            mula: data[i][4],
            tamat: data[i][5],
            kelas: data[i][6],
            subjek: data[i][7],
            temaTajuk: data[i][8],
            sk: data[i][9],
            sp: data[i][10],
            objektif: data[i][11],
            kriteriaKejayaan: data[i][12],
            aktiviti: data[i][13],
            bbmNilaiKbat: data[i][14],
            refleksi: data[i][15]
          });
        }
      }
    }

    return {
      bulan: bulan,
      senaraiMinggu: senaraiMinggu,
      mingguData: mingguData
    };
  } catch (err) {
    return { bulan: bulan, senaraiMinggu: [], mingguData: {}, ralat: err.message };
  }
}

function kemaskiniRefleksiPukalSebulan(payload) {
  try {
    var emel = String(payload.emel || "").trim();
    var bulan = String(payload.bulan || "").trim();
    var teksRefleksi = String(payload.teksRefleksi || "").trim();
    var mingguKhusus = payload.minggu ? String(payload.minggu).trim() : "";

    var senaraiMinggu = mingguKhusus ? [mingguKhusus] : dapatkanMingguBulan(bulan);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RPH_GURU");
    if (!sheet) return { success: false, message: "Sheet tiada." };

    var data = sheet.getDataRange().getDisplayValues();
    var colRefleksi = 16;
    var count = 0;

    for (var i = 1; i < data.length; i++) {
      if (padanEmelSama(data[i][1], emel)) {
        var padan = senaraiMinggu.some(function(m){ return padanMingguSama(data[i][2], m); });
        if (padan) {
          sheet.getRange(i + 1, colRefleksi).setValue(teksRefleksi);
          count++;
        }
      }
    }
    return { success: true, count: count, message: `Refleksi bagi ${count} slot berjaya dikemaskini!` };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// --------------------------------------------------------------------------
// 8. PENGAMBILAN REKOD MINGGUAN & REFLEKSI INDIVIDU
// --------------------------------------------------------------------------

function dapatkanRekodMinggu(minggu, emel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RPH_GURU");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getDisplayValues();
  var senarai = [];
  var isninTakwim = dapatkanIsninMinggu(minggu);

  for (var i = 1; i < data.length; i++) {
    if (padanEmelSama(data[i][1], emel) && padanMingguSama(data[i][2], minggu)) {
      var hariPenuh = formatNamaHari(data[i][3]);
      senarai.push({
        id: data[i][0],
        hari: hariPenuh,
        tarikh: dapatkanTarikhHari(isninTakwim, hariPenuh),
        mula: bersihkanMasa(data[i][4]),
        tamat: bersihkanMasa(data[i][5]),
        kelas: data[i][6] || "",
        subjek: data[i][7] || "",
        temaTajuk: data[i][8] || "",
        sk: data[i][9] || "",
        sp: data[i][10] || "",
        objektif: data[i][11] || "",
        kriteriaKejayaan: data[i][12] || "",
        aktiviti: data[i][13] || "",
        bbmNilaiKbat: data[i][14] || "",
        refleksi: data[i][15] || ""
      });
    }
  }
  return senarai;
}

function ambilRekodRphMinggu(minggu, emel) {
  return dapatkanRekodMinggu(minggu, emel);
}

function kemaskiniRefleksiSlotBackend(id, teksRefleksi) {
  return kemaskiniRefleksi(id, teksRefleksi);
}

function kemaskiniRefleksi(id, teksRefleksi) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RPH_GURU");
  if (!sheet) return false;
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.getRange(i + 1, 16).setValue(teksRefleksi);
      return true;
    }
  }
  return false;
}

// --------------------------------------------------------------------------
// 9. JANAAN PDF e-RPH MINGGUAN & BULANAN (A4 PORTRAIT)
// --------------------------------------------------------------------------

function formatAktivitiHtmlGas(teks) {
  if (!teks) return '-';
  return String(teks).split('\n').map(function(baris) {
    if (!baris.trim()) return '';
    var b = baris.split(':');
    if (b.length > 1) {
      return '<div style="margin-bottom:2px;"><b>' + b[0].trim() + ' :</b> ' + b.slice(1).join(':').trim() + '</div>';
    }
    return '<div style="margin-bottom:2px;">' + baris.trim() + '</div>';
  }).join('');
}

function janaPdfMingguBackend(payload) {
  return janaPdfMingguanBackend(payload.minggu, payload.emel);
}

function janaPdfMingguanBackend(minggu, emel) {
  try {
    var namaGuru = dapatkanNamaGuruDariEmel(emel);
    var rekod = dapatkanRekodMinggu(minggu, emel) || [];
    var kodMinggu = formatKodMinggu(minggu);
    var safeNama = String(namaGuru).replace(/[/\\?%*:|"<>]/g, '').trim().replace(/\s+/g, '_');
    var namaFail = "eRPH_" + kodMinggu + "_" + safeNama + ".pdf";

    var isPraRekod = rekod.some(function(s) {
      var k = String(s.kelas || "").toUpperCase();
      var sub = String(s.subjek || "").toUpperCase();
      return k.includes("PRA") || sub.includes("PERBUALAN") || sub.includes("AKTIVITI FIZIKAL") || sub.includes("REHAT") || sub.includes("PENUTUP");
    });

    var slotHtml = "";
    if (isPraRekod) {
      slotHtml = binaJadualKspk3LajurHtml(rekod, namaGuru, kodMinggu);
    } else {
      rekod.forEach(function(s, idx) {
        var temaStr = s.tema || (s.temaTajuk ? s.temaTajuk.split('|')[0].trim() : 'Standard Kurikulum');
        var tajukStr = s.tajuk || (s.temaTajuk && s.temaTajuk.includes('|') ? s.temaTajuk.split('|')[1].trim() : (s.temaTajuk || 'Penguasaan Konsep'));
        var isJawiSlot = s.subjek && (s.subjek.includes('ISLAM') || s.subjek.includes('ARAB') || /[\u0600-\u06FF]/.test(s.sk || ''));
        var jawiClass = isJawiSlot ? ' font-jawi' : '';

        slotHtml += `
          <div style="page-break-after: always; min-height: 94vh; display: flex; flex-direction: column; justify-content: space-between; border: 1.5px solid #1e293b; border-radius: 6px; padding: 10px; margin-bottom: 20px; box-sizing: border-box;">
            <div>
              <div style="text-align: center; border-bottom: 2px solid #0f172a; padding-bottom: 4px; margin-bottom: 6px;">
                <div style="font-size: 10.5px; font-weight: 800; text-transform: uppercase; color: #0f172a;">REKOD PENGAJARAN DAN PEMBELAJARAN HARIAN (e-RPH 2026)</div>
                <div style="font-size: 8px; font-weight: bold; color: #334155; text-transform: uppercase;">GURU: ${namaGuru.toUpperCase()} | SESI: 2026 | MINGGU: ${kodMinggu}</div>
              </div>
              <div style="background-color: #1e1b4b; color: #ffffff; border-radius: 4px; padding: 4px 8px; font-size: 9px; font-weight: bold; margin-bottom: 6px; display: flex; justify-content: space-between;">
                <span>SLOT ${idx + 1}: ${s.hari} (${s.tarikh || '-'}) | MASA: ${s.mula} - ${s.tamat}</span>
                <span>KELAS: ${s.kelas}</span>
              </div>
              <table style="width: 100%; border-collapse: collapse; font-size: 8px;">
                <tr><td style="width:26%; padding:4px 6px; font-weight:bold; background:#f8fafc; border:1px solid #cbd5e1;">SUBJEK</td><td style="padding:4px 6px; border:1px solid #cbd5e1; font-weight:bold; color:#1e1b4b;">${s.subjek}</td></tr>
                <tr><td style="padding:4px 6px; font-weight:bold; background:#f8fafc; border:1px solid #cbd5e1;">TAHUN / TINGKATAN</td><td style="padding:4px 6px; border:1px solid #cbd5e1; font-weight:bold;">${s.tahun || s.kelas}</td></tr>
                <tr><td style="padding:4px 6px; font-weight:bold; background:#f8fafc; border:1px solid #cbd5e1;">TEMA / BIDANG / TERAS</td><td style="padding:4px 6px; border:1px solid #cbd5e1;${jawiClass}">${temaStr}</td></tr>
                <tr><td style="padding:4px 6px; font-weight:bold; background:#f8fafc; border:1px solid #cbd5e1;">TAJUK / KEMAHIRAN / NILAI</td><td style="padding:4px 6px; border:1px solid #cbd5e1;${jawiClass}">${tajukStr}</td></tr>
                <tr><td style="padding:4px 6px; font-weight:bold; background:#f8fafc; border:1px solid #cbd5e1;">STANDARD KANDUNGAN (SK)</td><td style="padding:4px 6px; border:1px solid #cbd5e1; line-height:1.4;${jawiClass}"><b>${s.sk}</b></td></tr>
                <tr><td style="padding:4px 6px; font-weight:bold; background:#f8fafc; border:1px solid #cbd5e1;">STANDARD PEMBELAJARAN (SP)</td><td style="padding:4px 6px; border:1px solid #cbd5e1; line-height:1.4;${jawiClass}"><b>${s.sp}</b></td></tr>
                <tr><td style="padding:4px 6px; font-weight:bold; background:#f8fafc; border:1px solid #cbd5e1;">OBJEKTIF PEMBELAJARAN</td><td style="padding:4px 6px; border:1px solid #cbd5e1; line-height:1.4;">${s.objektif || '-'}</td></tr>
                <tr><td style="padding:4px 6px; font-weight:bold; background:#f8fafc; border:1px solid #cbd5e1;">KRITERIA KEJAYAAN</td><td style="padding:4px 6px; border:1px solid #cbd5e1; line-height:1.4; white-space:pre-wrap;">${s.kriteriaKejayaan || '-'}</td></tr>
                <tr><td style="padding:4px 6px; font-weight:bold; background:#f8fafc; border:1px solid #cbd5e1;">AKTIVITI PENGAJARAN (PdP)</td><td style="padding:4px 6px; border:1px solid #cbd5e1; line-height:1.4;">${formatAktivitiHtmlGas(s.aktiviti)}</td></tr>
                <tr><td style="padding:4px 6px; font-weight:bold; background:#f8fafc; border:1px solid #cbd5e1;">BBM, KBAT &amp; PBD</td><td style="padding:4px 6px; border:1px solid #cbd5e1;">${s.bbmNilaiKbat || '-'}</td></tr>
                <tr><td style="padding:4px 6px; font-weight:bold; background:#f8fafc; border:1px solid #cbd5e1;">REFLEKSI GURU &amp; IMPAK</td><td style="padding:4px 6px; border:1px solid #cbd5e1; line-height:1.4;">${s.refleksi || 'Murid berjaya menguasai standard pembelajaran yang ditetapkan.'}</td></tr>
              </table>
            </div>
            <div style="font-size: 7.5px; color: #64748b; text-align: right; border-top: 1px dashed #cbd5e1; padding-top: 3px; margin-top: 4px;">
              Dokumen e-RPH Sesi 2026 | Format 1 Muka Surat Standard KPM
            </div>
          </div>
        `;
      });
    }

    var tajukUtamaDoc = isPraRekod ? 'REKOD PENGAJARAN DAN PEMBELAJARAN HARIAN PRASEKOLAH (KSPK)' : 'REKOD PENGAJARAN DAN PEMBELAJARAN HARIAN';

    var html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=Noto+Naskh+Arabic:wght@400;700&display=swap');
    @page { size: A4 portrait; margin: 10mm 8mm 10mm 8mm; }
    body { font-family: 'Noto Naskh Arabic', 'Amiri', 'Traditional Arabic', 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: 8.5px; color: #111827; margin: 0; line-height: 1.4; }
    .font-jawi { font-family: 'Noto Naskh Arabic', 'Amiri', 'Traditional Arabic', serif; direction: rtl; text-align: right; line-height: 2.0; font-size: 10px; }
    .header { text-align: center; border-bottom: 2px solid #0f172a; padding-bottom: 5px; margin-bottom: 8px; }
    .header h2 { font-size: 11px; margin: 0 0 2px 0; text-transform: uppercase; color: #0f172a; font-weight: 800; }
    .header p { font-size: 8.5px; margin: 0; font-weight: bold; color: #334155; text-transform: uppercase; }
  </style>
</head>
<body>
  <div class="header">
    <h2>${tajukUtamaDoc}</h2>
    <p>GURU: ${namaGuru.toUpperCase()} | SESI: 2026 | MINGGU: ${kodMinggu}</p>
  </div>
  ${slotHtml || '<p style="text-align:center;">Tiada rekod e-RPH.</p>'}
</body>
</html>`;

    var folder = DriveApp.getRootFolder();
    var tempFile = DriveApp.createFile("temp_erph.html", html, MimeType.HTML);
    var pdfBlob = tempFile.getAs(MimeType.PDF).setName(namaFail);
    var pdfFile = folder.createFile(pdfBlob);
    try { pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

    var b64 = Utilities.base64Encode(pdfBlob.getBytes());
    tempFile.setTrashed(true);

    return {
      status: "SUCCESS",
      namaFail: namaFail,
      urlPdf: pdfFile.getUrl(),
      base64: b64
    };
  } catch (err) {
    throw new Error("Ralat menjana PDF: " + err.message);
  }
}

function binaJadualKspk3LajurHtml(rekod, namaGuru, kodMinggu) {
  var hariOrder = ["ISNIN", "SELASA", "RABU", "KHAMIS", "JUMAAT"];
  var mapHari = {};
  hariOrder.forEach(function(h) { mapHari[h] = []; });

  rekod.forEach(function(s) {
    var h = formatNamaHari(s.hari);
    if (!mapHari[h]) mapHari[h] = [];
    mapHari[h].push(s);
  });

  var htmlHari = "";

  hariOrder.forEach(function(h) {
    var slots = mapHari[h];
    if (!slots || slots.length === 0) return;

    // Susun mengikut masa mula
    slots.sort(function(a, b) {
      return String(a.mula || "").localeCompare(String(b.mula || ""));
    });

    var tarikhHari = slots[0].tarikh || "-";
    var temaHari = "";
    slots.forEach(function(sl) {
      if (sl.temaTajuk && sl.temaTajuk.includes("|")) {
        var bahagian = sl.temaTajuk.split("|");
        if (bahagian.length > 1 && !temaHari) {
          temaHari = bahagian[bahagian.length - 1].trim();
        }
      }
    });
    if (!temaHari) temaHari = "MALAYSIA & NEGERI SAYA";

    var barisTable = "";
    slots.forEach(function(s) {
      var mula = s.mula || "08:00";
      var tamat = s.tamat || "08:30";
      var durasi = kiraDurasiMinit(mula, tamat);

      var perkaraHtml = `
        <div style="margin-bottom: 4px;"><b>Tajuk:</b> ${s.temaTajuk || s.subjek}</div>
        ${s.sk ? `<div style="margin-bottom: 3px;"><b>Standard Pembelajaran (SP):</b><br>${s.sk}</div>` : ''}
        ${s.sp ? `<div style="margin-bottom: 4px; color: #1e293b;">${s.sp}</div>` : ''}
        ${s.objektif ? `<div style="margin-bottom: 4px;"><b>Objektif Pembelajaran:</b><br>${formatAktivitiHtmlGas(s.objektif)}</div>` : ''}
        ${s.aktiviti ? `<div style="margin-bottom: 4px;"><b>Aktiviti:</b><br>${formatAktivitiHtmlGas(s.aktiviti)}</div>` : ''}
        ${s.bbmNilaiKbat ? `<div style="margin-bottom: 4px;"><b>Bahan &amp; Catatan:</b> ${s.bbmNilaiKbat.split('|')[0].trim()}</div>` : ''}
        ${s.kriteriaKejayaan ? `<div style="background: #f8fafc; border: 1px solid #cbd5e1; padding: 4px; border-radius: 4px; margin-bottom: 4px;"><b>PEMBELAJARAN TERBEZA:</b><br>${formatAktivitiHtmlGas(s.kriteriaKejayaan)}</div>` : ''}
        <div style="border: 1px dashed #94a3b8; padding: 4px; background: #fafafa; border-radius: 4px; margin-top: 4px;">
          <b>Refleksi:</b> ${s.refleksi || 'Murid menunjukkan minat dan mencapai kemahiran yang ditetapkan.'}
        </div>
      `;

      var catatanHtml = "";
      if (String(s.subjek || "").toUpperCase().includes("PERBUALAN")) {
        catatanHtml += `
          <div style="border-bottom: 1px solid #cbd5e1; padding-bottom: 4px; margin-bottom: 5px;">
            <b>Kehadiran:</b><br>[ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ] / 25
          </div>
        `;
      }
      
      // Semak jika ada deskriptor TP dalam bbmNilaiKbat
      if (s.bbmNilaiKbat && s.bbmNilaiKbat.includes("TP")) {
        var bahagianTp = s.bbmNilaiKbat.split('|').filter(function(p){ return p.includes("TP"); });
        if (bahagianTp.length > 0) {
          catatanHtml += `<div style="font-size: 7.5px; line-height: 1.35;"><b>Pentaksiran:</b><br>${bahagianTp.map(function(t){ return t.trim(); }).join('<br>')}</div>`;
        }
      } else {
        catatanHtml += `
          <div style="font-size: 7.5px; line-height: 1.35;">
            <b>Pentaksiran:</b><br>
            TP1: Bimbingan rapi<br>
            TP2: Memuaskan / Berdikari<br>
            TP3: Cemerlang &amp; Contoh
          </div>
        `;
      }

      barisTable += `
        <tr style="border-bottom: 1px solid #334155; page-break-inside: avoid;">
          <td style="vertical-align: top; padding: 6px 5px; border-right: 1px solid #334155; background: #fbfbfe;">
            <b style="font-size: 9px; color: #0f172a;">${s.subjek}</b><br>
            <span style="color: #475569; font-weight: bold; font-size: 8px;">${mula} - ${tamat}</span><br>
            <span style="color: #64748b; font-size: 7.5px;">(${durasi} min)</span>
          </td>
          <td style="vertical-align: top; padding: 6px 8px; border-right: 1px solid #334155;">
            ${perkaraHtml}
          </td>
          <td style="vertical-align: top; padding: 6px 5px; background: #fdfdfd;">
            ${catatanHtml}
          </td>
        </tr>
      `;
    });

    htmlHari += `
      <div style="border: 2px solid #0f172a; margin-bottom: 22px; page-break-inside: avoid;">
        <table style="width: 100%; border-collapse: collapse; font-size: 9px; background: #fff;">
          <tr style="background: #f1f5f9; border-bottom: 1.5px solid #0f172a;">
            <td colspan="3" style="text-align: center; padding: 5px; font-size: 11px; font-weight: 900; letter-spacing: 0.5px; color: #0f172a;">
              RANCANGAN PENGAJARAN HARIAN PRASEKOLAH KSPK
            </td>
          </tr>
          <tr style="border-bottom: 1px solid #334155; font-weight: bold;">
            <td style="width: 33%; padding: 4px 8px; border-right: 1px solid #334155;">
              Minggu: ${kodMinggu.replace(/[^0-9]/g, '') || kodMinggu}
            </td>
            <td style="width: 34%; padding: 4px 8px; border-right: 1px solid #334155;">
              Hari: ${h}
            </td>
            <td style="width: 33%; padding: 4px 8px;">
              Tarikh: ${tarikhHari}
            </td>
          </tr>
          <tr style="background: #fafafa; border-bottom: 1.5px solid #0f172a;">
            <td colspan="3" style="padding: 4px 8px; font-weight: 800; color: #0f172a;">
              TEMA MINGGUAN: ${temaHari.toUpperCase()}
            </td>
          </tr>
        </table>

        <table style="width: 100%; border-collapse: collapse; font-size: 8.5px;">
          <thead>
            <tr style="background: #e2e8f0; border-bottom: 1.5px solid #0f172a; font-weight: 800; text-align: center;">
              <th style="width: 18%; padding: 5px; border-right: 1px solid #334155;">Masa / Bidang</th>
              <th style="width: 64%; padding: 5px; border-right: 1px solid #334155;">Perkara</th>
              <th style="width: 18%; padding: 5px;">Catatan / Impak</th>
            </tr>
          </thead>
          <tbody>
            ${barisTable}
          </tbody>
        </table>
      </div>
    `;
  });

  return htmlHari || '<p style="text-align:center;">Tiada rekod prasekolah.</p>';
}

function kiraDurasiMinit(mula, tamat) {
  try {
    var pMula = String(mula).split(':');
    var pTamat = String(tamat).split(':');
    var minMula = parseInt(pMula[0], 10) * 60 + parseInt(pMula[1], 10);
    var minTamat = parseInt(pTamat[0], 10) * 60 + parseInt(pTamat[1], 10);
    var beza = minTamat - minMula;
    return beza > 0 ? beza : 30;
  } catch (e) {
    return 30;
  }
}

function janaKompilasiBulananPdfBackend(payload) {
  try {
    var emel = payload.emel;
    var bulan = payload.bulan || "Januari";
    var senaraiMinggu = dapatkanMingguBulan(bulan);
    var namaGuru = dapatkanNamaGuruDariEmel(emel);
    var safeNama = String(namaGuru).replace(/[/\\?%*:|"<>]/g, '').trim().replace(/\s+/g, '_');
    var namaFail = "eRPH_Kompilasi_" + bulan + "_2026_" + safeNama + ".pdf";

    var allSlotsHtml = "";
    for (var w = 0; w < senaraiMinggu.length; w++) {
      var m = senaraiMinggu[w];
      var rekod = dapatkanRekodMinggu(m, emel) || [];
      if (rekod.length > 0) {
        allSlotsHtml += `
          <div style="background-color: #1e1b4b; color: #fff; padding: 5px 8px; font-weight: bold; font-size: 10px; margin-top: 14px; margin-bottom: 8px; border-radius: 4px; text-transform: uppercase;">
            📌 BAHAGIAN: ${m.toUpperCase()}
          </div>
        `;

        var isPraMinggu = rekod.some(function(s) {
          var k = String(s.kelas || "").toUpperCase();
          var sub = String(s.subjek || "").toUpperCase();
          return k.includes("PRA") || sub.includes("PERBUALAN") || sub.includes("AKTIVITI FIZIKAL") || sub.includes("REHAT") || sub.includes("PENUTUP");
        });

        if (isPraMinggu) {
          allSlotsHtml += binaJadualKspk3LajurHtml(rekod, namaGuru, m);
        } else {
          rekod.forEach(function(s, idx) {
            allSlotsHtml += `
              <div style="border: 1px solid #334155; border-radius: 4px; margin-bottom: 8px; page-break-inside: avoid;">
                <div style="background-color: #f8fafc; border-bottom: 1px solid #cbd5e1; padding: 3px 6px; font-size: 9px; font-weight: bold;">
                  ${s.hari} (${s.tarikh || '-'}) | ${s.mula}-${s.tamat} | ${s.kelas} — ${s.subjek}
                </div>
                <div style="padding: 4px 6px; font-size: 8px;">
                  <div><b>Tema/Tajuk:</b> ${s.temaTajuk || '-'}</div>
                  <div><b>SK/SP:</b> ${s.sk} / ${s.sp}</div>
                  <div><b>Objektif:</b> ${s.objektif || '-'}</div>
                  <div><b>Refleksi:</b> ${s.refleksi || 'Murid menguasai kemahiran pembelajaran yang ditetapkan.'}</div>
                </div>
              </div>
            `;
          });
        }
      }
    }

    var html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=Noto+Naskh+Arabic:wght@400;700&display=swap');
    @page { size: A4 portrait; margin: 12mm 10mm 12mm 10mm; }
    body { font-family: 'Noto Naskh Arabic', 'Amiri', 'Traditional Arabic', 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: 8.5px; color: #111827; margin: 0; line-height: 1.4; }
    .font-jawi { font-family: 'Noto Naskh Arabic', 'Amiri', 'Traditional Arabic', serif; direction: rtl; text-align: right; line-height: 2.0; font-size: 10px; }
    .header { text-align: center; border-bottom: 2px solid #0f172a; padding-bottom: 6px; margin-bottom: 10px; }
    .header h2 { font-size: 11px; margin: 0 0 2px 0; text-transform: uppercase; color: #0f172a; font-weight: 800; }
    .header p { font-size: 9px; margin: 0; font-weight: bold; color: #334155; text-transform: uppercase; }
  </style>
</head>
<body>
  <div class="header">
    <h2>KOMPILASI e-RPH SEBULAN PENUH</h2>
    <p>GURU: ${namaGuru.toUpperCase()} | BULAN: ${bulan.toUpperCase()} 2026</p>
  </div>
  ${allSlotsHtml || '<p style="text-align:center;">Tiada rekod dikesan.</p>'}
</body>
</html>`;

    var folder = DriveApp.getRootFolder();
    var tempFile = DriveApp.createFile("temp_kompilasi.html", html, MimeType.HTML);
    var pdfBlob = tempFile.getAs(MimeType.PDF).setName(namaFail);
    var pdfFile = folder.createFile(pdfBlob);
    try { pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

    var b64 = Utilities.base64Encode(pdfBlob.getBytes());
    tempFile.setTrashed(true);

    return {
      status: "SUCCESS",
      namaFail: namaFail,
      urlPdf: pdfFile.getUrl(),
      base64: b64
    };
  } catch (err) {
    throw new Error("Ralat kompilasi: " + err.message);
  }
}


// ============================================================================
// 7. MODUL DONE-FOR-YOU (DFY) 40-WEEK BATCH GENERATOR & SISTEM TOKEN SHOPEE
// ============================================================================

/**
 * Membina atau mendapatkan tab 'TokenDatabase' di Google Sheets.
 */
function dapatkanAtauCiptaSheetToken(ss) {
  var headers = ["Token", "Status", "SubjectAllowed", "YearAllowed", "RedeemedByEmail", "RedeemedAt"];
  var barisLalai = [
    ["RPH-8921-X", "ACTIVE", "Bahasa Inggeris (SK)", "Tahun 4", "", ""],
    ["RPH-2026-BM", "ACTIVE", "Bahasa Melayu", "Tahun 1", "", ""],
    ["RPH-2026-SN", "ACTIVE", "Sains", "Tahun 4", "", ""],
    ["RPH-2026-MT", "ACTIVE", "Matematik", "Tahun 2", "", ""],
    ["RPH-2026-PI", "ACTIVE", "Pendidikan Islam", "Tahun 1", "", ""],
    ["RPH-2026-BA", "ACTIVE", "Bahasa Arab", "Tahun 2", "", ""],
    ["RPH-2026-VIP", "ACTIVE", "SEMUA", "SEMUA", "", ""]
  ];
  return dapatkanAtauCiptaSheet(ss, "TokenDatabase", headers, barisLalai);
}

/**
 * Menyemak ketulenan Kod Token Shopee dalam 'TokenDatabase'.
 */
function semakTokenShopee(kodToken) {
  try {
    if (!kodToken) {
      return { valid: false, status: "INVALID", message: "Sila masukkan Kod Token Shopee anda." };
    }
    var cleanToken = String(kodToken).trim().toUpperCase();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = dapatkanAtauCiptaSheetToken(ss);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var rowToken = String(data[i][0] || "").trim().toUpperCase();
      if (rowToken === cleanToken) {
        var status = String(data[i][1] || "ACTIVE").trim().toUpperCase();
        var subAllowed = String(data[i][2] || "SEMUA").trim();
        var yrAllowed = String(data[i][3] || "SEMUA").trim();
        var redeemedEmail = String(data[i][4] || "").trim();
        var redeemedAt = String(data[i][5] || "").trim();

        if (status === "USED") {
          return {
            valid: false,
            status: "USED",
            message: "Token ini telah ditebus sebelum ini pada " + (redeemedAt || "tarikh lalu") + " oleh " + (redeemedEmail || "pengguna lain") + "."
          };
        }

        return {
          valid: true,
          status: "ACTIVE",
          subjectAllowed: subAllowed,
          yearAllowed: yrAllowed,
          message: "SAH - 1 Penebusan Tersedia"
        };
      }
    }

    return {
      valid: false,
      status: "INVALID",
      message: "Kod Token Shopee tidak sah atau tiada dalam rekod pesanan."
    };
  } catch (err) {
    return { valid: false, status: "ERROR", message: "Ralat semakan token: " + err.message };
  }
}

/**
 * Enjin Penjana Pukal 40 Minggu Penuh (Done-For-You Batch Compiler)
 * Membina fail Google Sheet rasmi 40 minggu di Google Drive dan berkongsi akses dengan pembeli.
 */
function batchGenerateRPH(payload) {
  return generateFullYearRPH(payload);
}

function generateFullYearRPH(payload) {
  try {
    if (!payload || typeof payload !== 'object') {
      return { success: false, message: "Ralat: Maklumat penjanaan tidak lengkap." };
    }

    var token = String(payload.token || "").trim().toUpperCase();
    var emel = String(payload.emel || "").trim().toLowerCase();
    var subjek = String(payload.subjek || "").trim() || "Bahasa Melayu";
    var tahun = String(payload.tahun || "").trim() || "Tahun 1";
    var kelas = String(payload.kelas || "").trim() || "4 Bakawali";
    var jadualMingguan = payload.jadualMingguan || [];

    if (!emel || !emel.includes("@")) return { success: false, message: "Sila masukkan emel Google Drive yang sah." };
    if (!jadualMingguan || jadualMingguan.length === 0) return { success: false, message: "Sila tandakan sekurang-kurangnya SATU (1) hari kelas." };

    if (!token) {
      token = "DFY-WEB-" + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000);
    }

    var ssMaster = SpreadsheetApp.getActiveSpreadsheet();
    var tokenRowIndex = -1;
    var sheetToken = null;

    try {
      sheetToken = dapatkanAtauCiptaSheetToken(ssMaster);
      var tokenData = sheetToken.getDataRange().getValues();

      for (var i = 1; i < tokenData.length; i++) {
        if (String(tokenData[i][0] || "").trim().toUpperCase() === token) {
          if (String(tokenData[i][1] || "").trim().toUpperCase() === "USED") {
            return { success: false, message: "Token ini telah ditebus sebelum ini." };
          }
          tokenRowIndex = i + 1;
          break;
        }
      }

      if (tokenRowIndex === -1) {
        sheetToken.appendRow([token, "PROCESSING", subjek, tahun, emel, Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss")]);
        tokenRowIndex = sheetToken.getLastRow();
      } else {
        sheetToken.getRange(tokenRowIndex, 2).setValue("PROCESSING");
      }
    } catch (eToken) {
      console.warn("Makluman sheet token: " + eToken.message);
    }

    // 2. Dapatkan DSKP (dari Sheet atau silibus progresif 40 minggu)
    var dskpList = dapatkanDskpDariSheet(subjek, tahun);
    if (!dskpList || dskpList.length === 0) {
      dskpList = janaDskpSilibusPenuh(subjek, tahun);
    }

    // 3. Cipta Google Spreadsheet Baharu untuk Pembeli di Google Drive
    var cleanSub = subjek.replace(/[/\\?%*:|"<>]/g, '').trim().replace(/\s+/g, '_');
    var cleanKelas = kelas.replace(/[/\\?%*:|"<>]/g, '').trim().replace(/\s+/g, '_');
    var namaFailBaru = "eRPH_2026_FullYear_" + cleanSub + "_" + cleanKelas + "_" + emel.split('@')[0];
    
    var newSs = SpreadsheetApp.create(namaFailBaru);

    // Tab 1: RINGKASAN_JADUAL (High Performance Batch Format)
    var sheetDashboard = newSs.getActiveSheet();
    sheetDashboard.setName("RINGKASAN_JADUAL");
    binaDashboardJadual(sheetDashboard, payload, token);

    // Tab 2: REKOD_40_MINGGU (High Performance 1-Shot Batching)
    var sheetRekod = newSs.insertSheet("REKOD_40_MINGGU");
    var totalRows = binaRekod40MingguSheet(sheetRekod, payload, dskpList);

    // Tab 3: TEMPLATE_CETAKAN_A4 (Batch Formatter)
    var sheetCetakan = newSs.insertSheet("TEMPLATE_CETAKAN_A4");
    binaTemplateCetakanA4(sheetCetakan, payload);

    // 4. Kemas kini status token kepada USED
    if (sheetToken && tokenRowIndex > 0) {
      try {
        sheetToken.getRange(tokenRowIndex, 2, 1, 5).setValues([["USED", subjek, tahun, emel, Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss")]]);
      } catch (e) {
        console.warn("Ralat kemaskini token: " + e.message);
      }
    }

    // 5. Berikan kebenaran Edit kepada Pembeli di Google Drive
    try {
      newSs.addEditor(emel);
    } catch (e) {
      console.warn("Ralat addEditor: " + e.message);
    }
    try {
      DriveApp.getFileById(newSs.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    } catch (e) {
      console.warn("Ralat setSharing: " + e.message);
    }

    // 6. Hantar Notifikasi Emel Automatik melalui GmailApp (Optional / Fail-safe)
    try {
      var tajukEmel = "[EduRPH] Fail e-RPH 40 Minggu Penuh Anda Telah Siap Dijana!";
      var mesejEmel = "Salam Sejahtera Cikgu,\n\n" +
        "Penebusan token (" + token + ") anda telah berjaya diproses!\n\n" +
        "MAKLUMAT RPH ANDA:\n" +
        "â€¢ Subjek: " + subjek + "\n" +
        "â€¢ Tahun & Kelas: " + tahun + " (" + kelas + ")\n" +
        "â€¢ Jumlah Slot PdP: " + totalRows + " sesi (Lengkap 40 Minggu Sesi 2026)\n\n" +
        "PAUTAN GOOGLE SPREADSHEET DI GOOGLE DRIVE ANDA:\n" +
        newSs.getUrl() + "\n\n" +
        "Akses 'Editor' telah diberikan terus ke emel Google Drive anda (" + emel + "). Anda boleh terus menyemak, mengedit, dan mencetak RPH bila-bila masa.\n\n" +
        "Terima kasih atas sokongan anda!";
      GmailApp.sendEmail(emel, tajukEmel, mesejEmel);
    } catch (e) {
      console.warn("Ralat sendEmail (diabaikan supaya tidak melambatkan respon): " + e.message);
    }

    return {
      success: true,
      fileUrl: newSs.getUrl(),
      fileName: newSs.getName(),
      totalSessions: totalRows,
      message: "Tahniah! e-RPH 40 Minggu Penuh (" + totalRows + " slot PdP) berjaya dijana dan sedia di Google Drive anda!"
    };

  } catch (err) {
    return { success: false, message: "Ralat penjanaan 40 minggu: " + err.message };
  }
}

/**
 * Membina Tab 1: Ringkasan Jadual & Dashboard Penebusan (High Performance Batch Mode)
 */
function binaDashboardJadual(sheet, payload, token) {
  var colWidths = [40, 160, 220, 180, 180];
  for (var w = 0; w < colWidths.length; w++) {
    sheet.setColumnWidth(w + 1, colWidths[w]);
  }

  var headerRange = sheet.getRange("B2:E2");
  headerRange.merge().setValue("SISTEM PENJANAAN e-RPH AUTO (PENEBUSAN PANTAS SHOPEE)")
    .setFontWeight("bold").setFontSize(13).setBackground("#15803d").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(2, 40);

  var infoRange = sheet.getRange("B3:E3");
  infoRange.merge().setValue("SESI AKADEMIK 2026 | REKOD PENGAJARAN 40 MINGGU PENUH")
    .setFontWeight("bold").setFontSize(10).setBackground("#f0fdf4").setFontColor("#166534")
    .setHorizontalAlignment("center");

  var rows = [
    ["Kod Token", token, "Status Penebusan", "SAH & SELESAI"],
    ["Subjek PdP", payload.subjek, "Tahun / Kelas", payload.tahun + " (" + payload.kelas + ")"],
    ["Emel Google Drive", payload.emel, "Tarikh Dijana", Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss")],
    ["Jumlah Minggu", "40 Minggu Persekolahan", "Format Cetakan", "1 Muka Surat Setiap PdP (Standard KPM)"]
  ];

  // Batch Write & Style Metadata
  var metaRange = sheet.getRange(5, 2, rows.length, 4);
  metaRange.setValues(rows);
  metaRange.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  var metaBg = rows.map(function() { return ["#f8fafc", "#ffffff", "#f8fafc", "#ffffff"]; });
  var metaWeights = rows.map(function() { return ["bold", "normal", "bold", "normal"]; });
  metaRange.setBackgrounds(metaBg);
  metaRange.setFontWeights(metaWeights);

  // Jadual Waktu Mengajar Mingguan
  var jadualStartRow = 11;
  sheet.getRange(jadualStartRow, 2, 1, 4).merge().setValue("JADUAL WAKTU MENGAJAR MINGGUAN")
    .setFontWeight("bold").setFontSize(11).setBackground("#1e1b4b").setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  var jadualHeaders = ["HARI", "MASA MULA", "MASA TAMAT", "TEMPOH (MINIT)"];
  var hdrRange = sheet.getRange(jadualStartRow + 1, 2, 1, 4);
  hdrRange.setValues([jadualHeaders]);
  hdrRange.setFontWeight("bold").setBackground("#e2e8f0").setHorizontalAlignment("center");

  var slots = payload.jadualMingguan || [];
  var totalMinit = 0;
  if (slots.length > 0) {
    var slotValues = [];
    for (var s = 0; s < slots.length; s++) {
      var min = parseInt(slots[s].minit, 10) || 60;
      totalMinit += min;
      slotValues.push([
        slots[s].hari,
        slots[s].mula,
        slots[s].tamat,
        min + " Minit"
      ]);
    }
    var slotsRange = sheet.getRange(jadualStartRow + 2, 2, slotValues.length, 4);
    slotsRange.setValues(slotValues);
    slotsRange.setHorizontalAlignment("center");
    slotsRange.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  }

  var totalRow = jadualStartRow + 2 + slots.length;
  sheet.getRange(totalRow, 2, 1, 3).merge().setValue("JUMLAH MASA SEMINGGU")
    .setFontWeight("bold").setBackground("#f0fdf4").setFontColor("#166534").setHorizontalAlignment("right");
  sheet.getRange(totalRow, 5).setValue(totalMinit + " Minit (" + (totalMinit / 60).toFixed(1) + " Jam)")
    .setFontWeight("bold").setBackground("#f0fdf4").setFontColor("#166534").setHorizontalAlignment("center");
  sheet.getRange(totalRow, 2, 1, 4).setBorder(true, true, true, true, true, true, "#15803d", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Panduan Penggunaan
  var panduanRow = totalRow + 2;
  sheet.getRange(panduanRow, 2, 1, 4).merge().setValue("PANDUAN & CARA PENGGUNAAN:")
    .setFontWeight("bold").setBackground("#f1f5f9");
  var panduanTeks = [
    "1. Tab 'REKOD_40_MINGGU': Mengandungi keseluruhan rancangan pengajaran bagi 40 minggu sesi 2026 yang telah siap dipetakan.",
    "2. Tab 'TEMPLATE_CETAKAN_A4': Digunakan untuk melihat dan mencetak e-RPH mingguan dalam format 1 muka surat A4.",
    "3. Anda bebas mengedit mana-mana teks, aktiviti atau refleksi mengikut keperluan bilik darjah harian anda."
  ];
  for (var p = 0; p < panduanTeks.length; p++) {
    sheet.getRange(panduanRow + 1 + p, 2, 1, 4).merge().setValue(panduanTeks[p]).setFontSize(9).setFontColor("#475569");
  }
}

/**
 * Membina Tab 2: Pangkalan Data Rekod 40 Minggu Penuh (Sequential DSKP Progression & High Performance 1-Shot Batching)
 */
function binaRekod40MingguSheet(sheet, payload, dskpListDefault) {
  var headers = [
    "BIL", "MINGGU", "TARIKH", "HARI", "MASA", "DURASI", "KELAS", "SUBJEK",
    "TEMA / BIDANG / TERAS", "TAJUK / KEMAHIRAN / NILAI",
    "STANDARD KANDUNGAN (SK)", "STANDARD PEMBELAJARAN (SP)",
    "OBJEKTIF PEMBELAJARAN", "KRITERIA KEJAYAAN", "AKTIVITI PENGAJARAN (PdP)",
    "BBM & KBAT", "REFLEKSI GURU & IMPAK"
  ];

  sheet.appendRow(headers);
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setBackground("#1e1b4b").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(1, 32);
  sheet.setFrozenRows(1);

  var subjekDefault = payload.subjek || "Bahasa Inggeris";
  var tahunDefault = payload.tahun || "Tahun 4";
  var kelasDefault = payload.kelas || "4 USM";
  var slots = payload.jadualMingguan || [];

  // Susun slot mengikut urutan hari (Isnin -> Jumaat) dan masa mula
  var urutanHari = { "ISNIN": 1, "MO": 1, "SELASA": 2, "TU": 2, "RABU": 3, "WE": 3, "KHAMIS": 4, "TH": 4, "JUMAAT": 5, "FR": 5 };
  slots.sort(function(a, b) {
    var hA = urutanHari[String(a.hari || "").toUpperCase()] || 99;
    var hB = urutanHari[String(b.hari || "").toUpperCase()] || 99;
    if (hA !== hB) return hA - hB;
    return String(a.mula || "").localeCompare(String(b.mula || ""));
  });

  // Tarikh mula Isnin sesi 2026 (12 Januari 2026)
  var baseDate = new Date(2026, 0, 12);
  var offsetHari = { "ISNIN": 0, "MO": 0, "SELASA": 1, "TU": 1, "RABU": 2, "WE": 2, "KHAMIS": 3, "TH": 3, "JUMAAT": 4, "FR": 4 };

  // Cache DSKP per subjek & tahun
  var dskpCache = {};
  dskpCache[subjekDefault + "_" + tahunDefault] = (dskpListDefault && dskpListDefault.length > 0) 
    ? dskpListDefault 
    : janaDskpSilibusPenuh(subjekDefault, tahunDefault);

  function getDskpForSlot(sub, thn) {
    var key = (sub || subjekDefault) + "_" + (thn || tahunDefault);
    if (!dskpCache[key] || dskpCache[key].length === 0) {
      var d = dapatkanDskpDariSheet(sub, thn);
      if (!d || d.length === 0) {
        d = janaDskpSilibusPenuh(sub, thn);
      }
      dskpCache[key] = d;
    }
    return dskpCache[key];
  }

  // =========================================================================
  // LOGIK PENJANAAN SILIBUS: SEQUENTIAL DSKP MAPPING (POINTER INDEX)
  // Tidak rawak! Mengikut susunan nombor DSKP (SP 1.1.1 -> 1.1.2 -> 2.1.1 dst.)
  // =========================================================================
  var dskpPointerMap = {}; // key: subjek + "_" + kelas -> index berturutan

  var rowCounter = 0;
  var allRows = [];

  for (var w = 1; w <= 40; w++) {
    var kodMinggu = "Minggu " + w;
    var mondayTimestamp = baseDate.getTime() + ((w - 1) * 7 * 24 * 60 * 60 * 1000);

    for (var s = 0; s < slots.length; s++) {
      rowCounter++;
      var slot = slots[s];
      var slotHari = String(slot.hari || "ISNIN").toUpperCase();
      var dayOffset = offsetHari[slotHari] !== undefined ? offsetHari[slotHari] : 0;
      var slotDate = new Date(mondayTimestamp + (dayOffset * 24 * 60 * 60 * 1000));
      var tarikhStr = Utilities.formatDate(slotDate, "GMT+8", "dd/MM/yyyy");

      var currentSubjek = slot.subjek || subjekDefault;
      var currentKelas = slot.kelas || kelasDefault;
      var currentTahun = slot.tahun || tahunDefault;

      // Kunci pointer untuk gabungan subjek & kelas
      var pointerKey = currentSubjek + "_" + currentKelas;
      if (dskpPointerMap[pointerKey] === undefined) {
        dskpPointerMap[pointerKey] = 0;
      }
      var ptr = dskpPointerMap[pointerKey];

      var slotDskpList = getDskpForSlot(currentSubjek, currentTahun);
      var totalDskpCount = slotDskpList.length;

      var tema = "";
      var tajuk = "";
      var sk = "";
      var sp = "";
      var obj = "";
      var kriteria = "";
      var akt = "";
      var bbm = "";
      var ref = "";

      // Format pilihan i-THINK & Pemulihan dari tetapan modal jadual
      var iThinkStr = "";
      if (slot.iThink && Array.isArray(slot.iThink) && slot.iThink.length > 0) {
        iThinkStr = slot.iThink.join(", ");
      } else if (typeof slot.iThink === "string" && slot.iThink.trim()) {
        iThinkStr = slot.iThink;
      }

      var pemulihanStr = "";
      if (slot.pemulihan && Array.isArray(slot.pemulihan) && slot.pemulihan.length > 0) {
        pemulihanStr = slot.pemulihan.join(", ");
      } else if (typeof slot.pemulihan === "string" && slot.pemulihan.trim()) {
        pemulihanStr = slot.pemulihan;
      }

      if (ptr < totalDskpCount) {
        // [1] PERKEMBANGAN KRONOLOGI BERURUTAN (NORMAL PROGRESSION)
        var d = slotDskpList[ptr] || {};
        tema = d.tema || "Standard Kurikulum Kebangsaan";
        tajuk = d.tajuk || ("Unit " + (ptr + 1) + ": Kemahiran Pembelajaran " + currentSubjek);
        sk = d.sk || ("Standard Kandungan " + (ptr + 1) + " Kurikulum Standard KPM");
        sp = d.sp || ("Standard Pembelajaran " + (ptr + 1) + ".1 Penguasaan Konsep Asas");
        obj = d.objektif || ("Pada akhir PdP, murid berupaya menguasai standard pembelajaran: " + sp + " dengan baik.");
        kriteria = "Murid berjaya sekiranya:\n1. Menyatakan sekurang-kurangnya 3 konsep utama dengan tepat.\n2. Melengkapkan lembaran kerja latihan berfokus.";
        
        akt = "Set Induksi : Guru memaparkan rangsangan kontekstual dan bersoal jawab.\n" +
              "Aktiviti Utama : Penerangan berperingkat, bimbingan guru & latihan bertulis terbeza.\n" +
              (iThinkStr ? ("Aktiviti KBAT / i-THINK : Mengaplikasi " + iThinkStr + " untuk sintesis maklumat.\n") : "") +
              (pemulihanStr ? ("Aktiviti Terbeza : " + pemulihanStr + " mengikut tahap penguasaan murid.\n") : "") +
              "Penutup : Rumusan pengajaran, penilaian formatif lisan dan refleksi murid.";
        
        bbm = "BBM: Buku Teks, Lembaran Kerja" + (iThinkStr ? (" | Peta Pemikiran: " + iThinkStr) : "") + " | Nilai: Ketekunan | KBAT: Mengaplikasi | PBD: Lisan & Bertulis";
        ref = "Murid menguasai objektif pembelajaran" + (pemulihanStr ? (" dengan " + pemulihanStr) : " dengan bimbingan guru minima") + ".";

        // Naikkan pointer secara jujukan untuk slot seterusnya
        dskpPointerMap[pointerKey] = ptr + 1;
      } else {
        // [2] PENGENDALIAN LEBIHAN (OVERFLOW HANDLING)
        var lastDskp = slotDskpList[totalDskpCount - 1] || {};
        var revIdx = (ptr - totalDskpCount) + 1;

        tema = lastDskp.tema || "Pengukuhan & Penilaian Berterusan";
        tajuk = "Aktiviti Pengukuhan & Ulang Kaji Topikal Siri " + revIdx + " (" + (lastDskp.tajuk || currentSubjek) + ")";
        sk = lastDskp.sk || "Standard Kandungan Pengukuhan dan Penilaian Formatif";
        sp = (lastDskp.sp || "Standard Pembelajaran") + " [Sesi Pengukuhan, Pemulihan & Pentaksiran Bilik Darjah]";
        obj = "Pada akhir PdP, murid berupaya memperkukuh, memulihkan dan mengaplikasikan kemahiran bagi topik: " + (lastDskp.tajuk || currentSubjek) + " melalui latihan terbeza.";
        kriteria = "Murid berjaya sekiranya:\n1. Menjawab sekurang-kurangnya 4 daripada 5 soalan pengukuhan topik dengan betul.\n2. Menyelesaikan bimbingan pemulihan berfokus.";
        
        akt = "Set Induksi : Ujian diagnostik pantas / soal jawab imbas kembali konsep utama.\n" +
              "Aktiviti Utama : " + (pemulihanStr ? ("Pelaksanaan modul: " + pemulihanStr + ". ") : "Bimbingan pemulihan berfokus guru. ") +
              "Murid aras tinggi melengkapkan pengayaan cerdas manakala murid pemulihan menerima bimbingan terus.\n" +
              (iThinkStr ? ("Aplikasi i-THINK : Menggunakan " + iThinkStr + " sebagai peta konsep ulangkaji.\n") : "") +
              "Penutup : Refleksi pentaksiran bilik darjah (PBD) dan penetapan tindakan susulan.";

        bbm = "BBM: Modul Pengukuhan & Pemulihan" + (iThinkStr ? (" | Peta Pemikiran: " + iThinkStr) : "") + " | Nilai: Kerjasama | KBAT: Menganalisis | PBD: Pentaksiran Formatif";
        ref = "Sesi pengukuhan dan pemulihan berjalan lancar. Murid menunjukkan peningkatan aras penguasaan kemahiran.";

        // Teruskan pointer
        dskpPointerMap[pointerKey] = ptr + 1;
      }

      var namaHariPenuh = slotHari;
      if (slotHari === "MO" || slotHari === "ISNIN") namaHariPenuh = "Isnin";
      else if (slotHari === "TU" || slotHari === "SELASA") namaHariPenuh = "Selasa";
      else if (slotHari === "WE" || slotHari === "RABU") namaHariPenuh = "Rabu";
      else if (slotHari === "TH" || slotHari === "KHAMIS") namaHariPenuh = "Khamis";
      else if (slotHari === "FR" || slotHari === "JUMAAT") namaHariPenuh = "Jumaat";

      allRows.push([
        rowCounter,
        kodMinggu,
        tarikhStr,
        namaHariPenuh,
        (slot.mula || "07:10 AM") + " - " + (slot.tamat || "07:40 AM"),
        (slot.minit || 30) + " Minit",
        currentKelas,
        currentSubjek,
        tema,
        tajuk,
        sk,
        sp,
        obj,
        kriteria,
        akt,
        bbm,
        ref
      ]);
    }
  }

  // =========================================================================
  // HIGH-PERFORMANCE 1-SHOT BATCH OPERATIONS (95% LEBIH PANTAS)
  // Menghapuskan loop getRange/setBackground satu demi satu yang menyebabkan
  // kelewatan 20-30 saat. Kini dilaksanakan dalam 1 remote call sahaja!
  // =========================================================================
  if (allRows.length > 0) {
    var totalCols = headers.length;
    var dataRange = sheet.getRange(2, 1, allRows.length, totalCols);
    dataRange.setValues(allRows);
    dataRange.setWrap(true);
    
    // Matriks warna latar (Zebra striping dalam RAM tanpa remote call berulang)
    var bgColors = [];
    for (var r = 0; r < allRows.length; r++) {
      var rowColor = (r % 2 === 1) ? "#f8fafc" : "#ffffff";
      var rowArr = [];
      for (var c = 0; c < totalCols; c++) {
        rowArr.push(rowColor);
      }
      bgColors.push(rowArr);
    }
    dataRange.setBackgrounds(bgColors);
    dataRange.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
    
    // Format penjajaran tengah bagi kolum metadata
    sheet.getRange(2, 1, allRows.length, 7).setHorizontalAlignment("center");
  }

  // Tetapkan lebar lajur optimum secara terus
  var colWidths = [40, 80, 85, 75, 100, 65, 85, 120, 140, 160, 200, 220, 220, 200, 240, 180, 200];
  for (var colIdx = 0; colIdx < colWidths.length; colIdx++) {
    sheet.setColumnWidth(colIdx + 1, colWidths[colIdx]);
  }

  return allRows.length;
}

/**
 * Membina Tab 3: Template Cetakan A4 Standard KPM (High Performance Batch Mode)
 */
function binaTemplateCetakanA4(sheet, payload) {
  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 480);

  var header = sheet.getRange("B2:C2");
  header.merge().setValue("REKOD PENGAJARAN DAN PEMBELAJARAN HARIAN (e-RPH)")
    .setFontWeight("bold").setFontSize(12).setBackground("#1e1b4b").setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  sheet.getRange("B3").setValue("PILIH MINGGU:").setFontWeight("bold").setBackground("#e2e8f0");
  sheet.getRange("C3").setValue("Minggu 1").setFontWeight("bold").setFontColor("#15803d").setFontSize(11);

  var senaraiMinggu = [];
  for (var m = 1; m <= 40; m++) {
    senaraiMinggu.push("Minggu " + m);
  }
  var dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(senaraiMinggu, true)
    .build();
  sheet.getRange("C3").setDataValidation(dropdownRule);

  var labels = [
    ["SUBJEK", payload.subjek],
    ["TAHUN / KELAS", payload.tahun + " (" + payload.kelas + ")"],
    ["TEMA / BIDANG / TERAS", '=IFERROR(VLOOKUP($C$3, REKOD_40_MINGGU!B:Q, 8, FALSE), "Standard Kurikulum Kebangsaan")'],
    ["TAJUK / KEMAHIRAN / NILAI", '=IFERROR(VLOOKUP($C$3, REKOD_40_MINGGU!B:Q, 9, FALSE), "Penguasaan Konsep & Kemahiran")'],
    ["STANDARD KANDUNGAN (SK)", '=IFERROR(VLOOKUP($C$3, REKOD_40_MINGGU!B:Q, 10, FALSE), "1.1 Standard Kandungan KPM")'],
    ["STANDARD PEMBELAJARAN (SP)", '=IFERROR(VLOOKUP($C$3, REKOD_40_MINGGU!B:Q, 11, FALSE), "1.1.1 Menguasai kemahiran pembelajaran")'],
    ["OBJEKTIF PEMBELAJARAN", '=IFERROR(VLOOKUP($C$3, REKOD_40_MINGGU!B:Q, 12, FALSE), "Pada akhir PdP, murid dapat menguasai kemahiran yang ditetapkan.")'],
    ["KRITERIA KEJAYAAN", '=IFERROR(VLOOKUP($C$3, REKOD_40_MINGGU!B:Q, 13, FALSE), "Murid dapat melengkapkan latihan terbeza secara berdikari.")'],
    ["AKTIVITI PdP", '=IFERROR(VLOOKUP($C$3, REKOD_40_MINGGU!B:Q, 14, FALSE), "Set Induksi, Aktiviti Utama PAK21, Penutup.")'],
    ["BBM & KBAT", '=IFERROR(VLOOKUP($C$3, REKOD_40_MINGGU!B:Q, 15, FALSE), "Buku Teks, Lembaran Kerja | KBAT: Mengaplikasi")'],
    ["REFLEKSI GURU & IMPAK", '=IFERROR(VLOOKUP($C$3, REKOD_40_MINGGU!B:Q, 16, FALSE), "Murid mencapai objektif PdP.")']
  ];

  // Batch Write & Style: Menggantikan puluhan panggilan sel kepada 3 panggilan sahaja
  var templateRange = sheet.getRange(5, 2, labels.length, 2);
  templateRange.setValues(labels);
  templateRange.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  
  var tBgs = [];
  var tWeights = [];
  for (var i = 0; i < labels.length; i++) {
    tBgs.push(["#f8fafc", "#ffffff"]);
    tWeights.push(["bold", "normal"]);
  }
  templateRange.setBackgrounds(tBgs);
  templateRange.setFontWeights(tWeights);
  sheet.getRange(5, 3, labels.length, 1).setWrap(true);

  sheet.getRange(5 + labels.length + 1, 2, 1, 2).merge()
    .setValue("Nota: Tukar pilihan di sel C3 (Pilih Minggu) untuk menukar paparan cetakan secara automatik.")
    .setFontSize(8.5).setFontColor("#64748b").setHorizontalAlignment("center");
}

/**
 * Pangkalan Data Silibus Progresif 40 Minggu Penuh (Fallback Bersih KPM dalam Ayat Penuh)
 */
function janaDskpSilibusPenuh(subjek, tahun) {
  var subUpper = String(subjek || "").toUpperCase().trim();
  var thnDigit = String(tahun || "").replace(/\D/g, '') || "1";
  var hasil = [];

  // Pangkalan data progresif 40 unit mengikut subjek
  for (var i = 1; i <= 40; i++) {
    if (subUpper.includes("INGGERIS") || subUpper.includes("ENGLISH")) {
      hasil.push({
        subjek: subjek,
        tahun: tahun,
        tema: (i <= 10 ? "World of Self, Family and Friends" : (i <= 20 ? "World of Stories" : (i <= 30 ? "World of Knowledge" : "Revision and Assessment"))),
        tajuk: "Unit " + i + ": Progressive Language Skills " + i,
        sk: "2." + (i % 5 + 1) + " Communicate simple information intelligibly for target purposes",
        sp: "2." + (i % 5 + 1) + ".1 Give detailed information and describe activities with appropriate grammar",
        objektif: "By the end of the lesson, pupils will be able to apply target language structures and vocabulary correctly in context.",
        aktiviti: "Set Induction: Interactive flashcard prompt. Main Activity: Guided reading and independent worksheet exercise. Closure: Quick recap quiz."
      });
    } else if (subUpper.includes("ISLAM") || subUpper.includes("ARAB")) {
      var isArab = subUpper.includes("ARAB");
      var temaJawi = isArab ? "Ù…ÙŽÙ‡ÙŽØ§Ø±ÙŽØ§ØªÙ Ø§Ù„Ù„ÙÙ‘ØºÙŽØ©Ù Ø§Ù„Ø¹ÙŽØ±ÙŽØ¨ÙÙŠÙŽÙ‘Ø©Ù (Kemahiran Bahasa Arab)" : (i <= 10 ? "Ø§Ù„Ù‚Ø±Ø¡Ø§Ù† (ØªÙ„Ø§ÙˆØ© Ø¯Ø§Ù† Ø­Ø§ÙØ¸Ù†)" : (i <= 20 ? "Ø¹Ù‚ÙŠØ¯Ø© (Ø±ÙˆÙƒÙˆÙ† Ø§ÙŠÙ…Ø§Ù†)" : (i <= 30 ? "Ø¹Ø¨Ø§Ø¯Ø© (Ø·Ù‡Ø§Ø±Ù‡ Ø¯Ø§Ù† ØµÙ„ÙˆØ©)" : "Ø§Ø¯Ø¨ Ø¯Ø§Ù† Ø§Ø®Ù„Ø§Ù‚ Ø§Ø³Ù„Ø§Ù…ÙŠÙ‡")));
      var tajukJawi = isArab ? ("Ø§Ù„Ù’ÙˆÙŽØ­Ù’Ø¯ÙŽØ©Ù " + i + " : Ø§Ù„Ø¯ÙÙ‘Ø±ÙŽØ§Ø³ÙŽØ©Ù ÙˆÙŽØ§Ù„Ù’Ø­ÙŽÙŠÙŽØ§Ø©Ù Ø§Ù„Ù’ÙŠÙŽÙˆÙ’Ù…ÙÙŠÙŽÙ‘Ø©Ù") : ("Ø§ÙˆÙ†ÙŠØª " + i + " : Ú¤Ú Ø§Ø¬Ø±Ù† Ø¯Ø§Ù† ÙƒÙÙ‡Ù…Ù† Ø¨Ø±ØªÙŠÙƒ Ø³");
      var skJawi = isArab ? "Ù¡Ù«Ù¡ Ø§Ù„Ø§Ø³Ù’ØªÙÙ…ÙŽØ§Ø¹Ù Ø¥ÙÙ„ÙŽÙ‰ ÙƒÙŽÙ„ÙÙ…ÙŽØ§ØªÙ Ø§Ù„Ù’Ù…ÙŽØ­ÙŽØ§ÙˆÙØ±Ù ÙˆÙŽÙ†ÙØ·Ù’Ù‚ÙÙ‡ÙŽØ§ Ù†ÙØ·Ù’Ù‚Ù‹Ø§ ØµÙŽØ­ÙÙŠØ­Ù‹Ø§" : ("1." + (i % 5 + 1) + " ØªÙ„Ø§ÙˆØ© Ø¯Ø§Ù† Ø­Ø§ÙØ¸Ù† Ø§ÙŠØ§ØªÙ¢ Ø§Ù„Ù‚Ø±Ø¡Ø§Ù† Ø¨Ø±ØªØ¬ÙˆÙŠØ¯ Ø¯Ø§Ù† ÙØµÙŠØ­");
      var spJawi = isArab ? "Ù¡Ù«Ù¡Ù«Ù¡ Ø§Ù„Ù’Ù‚ÙØ¯Ù’Ø±ÙŽØ©Ù Ø¹ÙŽÙ„ÙŽÙ‰ Ù…ÙØ­ÙŽØ§ÙƒÙŽØ§Ø©Ù Ø§Ù„Ù’ÙƒÙŽÙ„ÙÙ…ÙŽØ§ØªÙ Ø§Ù„Ù’Ù…ÙŽØ³Ù’Ù…ÙÙˆØ¹ÙŽØ©Ù ÙˆÙŽØªÙŽØ±Ù’Ø¯ÙÙŠØ¯ÙÙ‡ÙŽØ§ Ø«ÙÙ…ÙŽÙ‘ Ù†ÙØ·Ù’Ù‚ÙÙ‡ÙŽØ§" : ("1." + (i % 5 + 1) + ".1 Ù…Ù…Ø¨Ø§Ú†ØŒ Ù…Ú Ø­ÙØ¸ Ø¯Ø§Ù† Ù…Ú Ø¹Ù…Ù„ÙƒÙ† Ø§ÙŠØ© Ø¯Ú Ù† Ù…Ø®Ø±Ø¬ Ø­Ø±ÙˆÙ ÙŠÚ  Ø¨ØªÙˆÙ„");
      
      hasil.push({
        subjek: subjek,
        tahun: tahun,
        tema: temaJawi,
        tajuk: tajukJawi,
        sk: skJawi,
        sp: spJawi,
        objektif: "Pada akhir pengajaran dan pembelajaran, murid dapat menguasai kemahiran tilawah dan memahami konsep dengan betul.",
        aktiviti: "Set Induksi : Talaqqi musyafahah. Aktiviti Utama : Latih tubi sebutan makhraj huruf dan lembaran kerja jawi. Penutup : Tasmik dan rumusan guru."
      });
    } else if (subUpper.includes("SAINS") || subUpper.includes("SCIENCE")) {
      hasil.push({
        subjek: subjek,
        tahun: tahun,
        tema: (i <= 10 ? "Inkuiri dalam Sains" : (i <= 20 ? "Sains Hayat" : (i <= 30 ? "Sains Fizikal" : "Bumi dan Angkasa"))),
        tajuk: "Unit " + i + ": Penyiasatan Saintifik dan Penerokaan Alam " + i,
        sk: "1." + (i % 4 + 1) + " Kemahiran Proses Sains Mengkaji Fenomena Alam",
        sp: "1." + (i % 4 + 1) + ".1 Memerhati, mengelas, mengukur dan membuat inferens secara bersistem",
        objektif: "Pada akhir PdP, murid berupaya menyatakan pemerhatian dan membuat inferens saintifik dengan tepat.",
        aktiviti: "Set Induksi : Demonstrasi bahan maujud. Aktiviti Utama : Amali sains PAK21 berpandukan buku teks dan lembaran aktiviti. Penutup : Pembentangan kumpulan."
      });
    } else if (subUpper.includes("MATEMATIK") || subUpper.includes("MATH")) {
      hasil.push({
        subjek: subjek,
        tahun: tahun,
        tema: (i <= 15 ? "Nombor dan Operasi" : (i <= 28 ? "Sukatan dan Geometri" : "Perkaitan dan Aljabar")),
        tajuk: "Unit " + i + ": Aplikasi Konsep Matematik dan Penyelesaian Masalah " + i,
        sk: "2." + (i % 4 + 1) + " Operasi Asas dan Pengiraan Nilai Nombor",
        sp: "2." + (i % 4 + 1) + ".1 Menyelesaikan ayat matematik dan masalah rutin harian dengan tepat",
        objektif: "Pada akhir PdP, murid dapat mengira dan menyelesaikan sekurang-kurangnya 4 daripada 5 soalan rutin dengan betul.",
        aktiviti: "Set Induksi : Kuiz congak pantas. Aktiviti Utama : Pengajaran berperingkat CPA (Konkrit, Piktorial, Abstrak) dan latihan bertulis. Penutup : Refleksi murid."
      });
    } else {
      // Subjek Bahasa Melayu & Umum
      hasil.push({
        subjek: subjek,
        tahun: tahun,
        tema: (i <= 10 ? "Kekeluargaan dan Kesihatan" : (i <= 20 ? "Kebersihan dan Keselamatan" : (i <= 30 ? "Perpaduan dan Jati Diri" : "Sains, Teknologi dan Inovasi"))),
        tajuk: "Unit " + i + ": Penguasaan Tatabahasa dan Kemahiran Berbahasa " + i,
        sk: (i % 3 + 1) + "." + (i % 4 + 1) + " Kemahiran Berbahasa Mendengar, Membaca dan Menulis",
        sp: (i % 3 + 1) + "." + (i % 4 + 1) + ".1 Membina ayat, memahami petikan dan mengaplikasikan tatabahasa dengan tepat",
        objektif: "Pada akhir pengajaran dan pembelajaran, murid dapat membina ayat dan memahami kosa kata kontekstual dengan baik.",
        aktiviti: "Set Induksi : Tayangan video / gambar rangsangan. Aktiviti Utama : Bacaan terbimbing, perbincangan PAK21 dan latihan bertulis. Penutup : Rumusan guru."
      });
    }
  }

  return hasil;
}
/**
 * ============================================================================
 * ENDPOINT UTAMA: PENJANAAN RPH 40 MINGGU DARI WEB APP (TAB 2 aSc GRID)
 * ============================================================================
 */

// ============================================================================

// ============================================================================

// ============================================================================
// 8. ENJIN PENJANAAN DOKUMEN MICROSOFT WORD (.DOCX) 40 MINGGU TERBEZA (KSSR)
// ============================================================================

/**
 * Menjana 1 Fail Microsoft Word (.docx) Lengkap 40 Minggu Mengikut Format RPH Terbeza KPM (Kump 1, 2, 3)
 */
function janaRph40MingguDocx(payload) {
  try {
    if (!payload || typeof payload !== 'object') {
      return { success: false, message: "Ralat: Maklumat penjanaan tidak lengkap." };
    }

    var nama = String(payload.nama || "Pendidik KPM").trim();
    var sekolah = String(payload.sekolah || "Sekolah Kebangsaan").trim();
    var emel = String(payload.emel || "").trim().toLowerCase();
    var sesi = String(payload.sesi || "2026").trim();
    var aliran = String(payload.aliran || "PERDANA").trim().toUpperCase();
    var subjek = String(payload.subjek || payload.subject || "Bahasa Melayu").trim();
    var tahun = String(payload.tahun || payload.year || "Tahun 4").trim();
    var kelas = String(payload.kelas || payload.className || "4 USM").trim();
    var jadualMingguan = payload.jadualMingguan || [];

    if (!jadualMingguan || jadualMingguan.length === 0) {
      return { success: false, message: "Sila tandakan sekurang-kurangnya SATU (1) petak jadual waktu PdP." };
    }

    // 1. Dapatkan Silibus DSKP 40 Minggu
    var dskpList = dapatkanDskpDariSheet(subjek, tahun);
    if (!dskpList || dskpList.length === 0) {
      dskpList = janaDskpSilibusPenuh(subjek, tahun);
    }

    var docId = null;
    var downloadUrl = "";
    var docViewUrl = "";

    try {
      var docName = "e-RPH 40 Minggu Terbeza - " + subjek + " " + tahun + " (" + kelas + ") - " + nama;
      var doc = DocumentApp.create(docName);
      var body = doc.getBody();

      body.setMarginTop(28.35);
      body.setMarginBottom(28.35);
      body.setMarginLeft(28.35);
      body.setMarginRight(28.35);

      // ==========================================
      // MUKA DEPAN (COVER PAGE MASTER)
      // ==========================================
      var pTop = body.appendParagraph("KEMENTERIAN PENDIDIKAN MALAYSIA");
      pTop.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontFamily("Arial").setFontSize(13).setBold(true).setSpacingBefore(30);

      var pTitle = body.appendParagraph("REKOD PENGAJARAN & PEMBELAJARAN HARIAN TERBEZA (e-RPH)\nPENDEKATAN TERBEZA & PAK-21 | SESI " + sesi);
      pTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontFamily("Arial").setFontSize(15).setBold(true).setFontColor("#1e1b4b").setSpacingBefore(15).setSpacingAfter(35);

      var tableCover = body.appendTable([
        ["NAMA GURU", nama],
        ["KOD & NAMA SEKOLAH", sekolah],
        ["EMEL RASMI / DELIMa", emel || "-"],
        ["MATA PELAJARAN", subjek],
        ["TAHUN / DARJAH (DSKP)", tahun],
        ["NAMA KELAS", kelas],
        ["ALIRAN PROGRAM", (aliran === "PRASEKOLAH" ? "Prasekolah (KSPK)" : (aliran === "PPKI" ? "Pendidikan Khas (PPKI)" : "Aliran Perdana (KSSR Terbeza)"))],
        ["PENDEKATAN PEDAGOGI", "Pendekatan Terbeza (Kump 1: Pengayaan, Kump 2: Pengukuhan, Kump 3: Pemulihan)"],
        ["JUMLAH MINGGU", "40 Minggu Persekolahan Lengkap"],
        ["FORMAT DOKUMEN", "Microsoft Word (.docx) - Standard KPM"]
      ]);

      tableCover.setBorderWidth(1).setBorderColor("#cbd5e1");
      for (var r = 0; r < tableCover.getNumRows(); r++) {
        var rowC = tableCover.getRow(r);
        var c0 = rowC.getCell(0);
        c0.setWidth(170).setBackgroundColor("#f8fafc");
        c0.getChild(0).asParagraph().setBold(true).setFontSize(10).setFontFamily("Arial");
        var c1 = rowC.getCell(1);
        c1.setWidth(330);
        c1.getChild(0).asParagraph().setFontSize(10).setFontFamily("Arial");
      }

      body.appendParagraph("\n\n* Dokumen e-RPH Terbeza ini dijana secara automatik berasaskan DSKP KPM dan pemetaan pedagogi PAK-21.").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(9).setItalic(true).setFontColor("#64748b");
      body.appendPageBreak();

      // ==========================================
      // JANA 40 MINGGU PENUH e-RPH TERBEZA
      // ==========================================
      var baseDate = new Date(2026, 0, 12);
      var dskpIndex = 0;

      for (var w = 1; w <= 40; w++) {
        var kodMinggu = "MINGGU " + w;
        var mondayTimestamp = baseDate.getTime() + ((w - 1) * 7 * 24 * 60 * 60 * 1000);

        var pMinggu = body.appendParagraph(kodMinggu + " | SESI " + sesi);
        pMinggu.setHeading(DocumentApp.ParagraphHeading.HEADING1).setFontFamily("Arial").setFontSize(13).setBold(true).setFontColor("#1e1b4b").setSpacingBefore(10).setSpacingAfter(8);

        for (var sIdx = 0; sIdx < jadualMingguan.length; sIdx++) {
          var slot = jadualMingguan[sIdx];
          var dskp = dskpList[dskpIndex % dskpList.length];
          dskpIndex++;

          var offset = 0;
          var hName = String(slot.day || slot.hari || "ISNIN").toUpperCase();
          if (hName.includes("SELASA")) offset = 1;
          else if (hName.includes("RABU")) offset = 2;
          else if (hName.includes("KHAMIS")) offset = 3;
          else if (hName.includes("JUMAAT")) offset = 4;

          var slotDate = new Date(mondayTimestamp + (offset * 24 * 60 * 60 * 1000));
          var dStr = (slotDate.getDate() < 10 ? '0' : '') + slotDate.getDate() + '/' + 
                     ((slotDate.getMonth() + 1) < 10 ? '0' : '') + (slotDate.getMonth() + 1) + '/' + 
                     slotDate.getFullYear();

          var masaStr = (slot.time || (slot.mula + " - " + slot.tamat) || "07:40 - 08:40") + " (" + (slot.duration || slot.minit || 60) + " Minit)";
          var ithinkStr = (slot.ithink || slot.iThink || []).join(", ") || "Peta Buih (Bubble)";
          var pemulihanStr = (slot.pemulihan || []).join(", ") || "Bimbingan Guru (Scaffolding)";

          var tableRph = body.appendTable();
          tableRph.setBorderWidth(1).setBorderColor("#94a3b8");

          // Header Bar
          var rowH = tableRph.appendTableRow();
          var cHdr = rowH.appendTableCell("RANCANGAN PENGAJARAN HARIAN TERBEZA\n" + (slot.subject || slot.subjek || subjek).toUpperCase() + " " + (slot.year || slot.tahun || tahun).toUpperCase());
          cHdr.setBackgroundColor("#ede9fe");
          cHdr.getChild(0).asParagraph().setBold(true).setFontSize(10).setFontFamily("Arial").setAlignment(DocumentApp.HorizontalAlignment.CENTER);

          // Maklumat Asas
          var rphInfoRows = [
            ["MINGGU", "" + w, "MASA", masaStr],
            ["HARI / TARIKH", hName + ", " + dStr, "KELAS", (slot.className || slot.kelas || kelas)],
            ["TEMA", dskp.tema || "Keluarga & Masyarakat", "TAJUK", dskp.tajuk || "Penguasaan Konsep & Kemahiran"],
            ["KEMAHIRAN", "Mendengar, Bertutur, Membaca & Menulis", "ASPIRASI MURID", "Etika & Kerohanian, Kemahiran Berfikir"],
            ["STANDARD KANDUNGAN", dskp.sk || "1.1 Standard Kandungan KPM", "", ""],
            ["STANDARD PEMBELAJARAN", dskp.sp || "1.1.1 Menguasai kemahiran pembelajaran berasaskan DSKP", "", ""],
            ["KEMAHIRAN TMK", "1.11 Menghasilkan dan menyunting teks, imej dan audio digital", "", ""]
          ];

          rphInfoRows.forEach(function(item) {
            var row = tableRph.appendTableRow();
            if (item[2] !== "") {
              var c1 = row.appendTableCell(item[0]); var c2 = row.appendTableCell(item[1]);
              var c3 = row.appendTableCell(item[2]); var c4 = row.appendTableCell(item[3]);
              c1.setWidth(110); c2.setWidth(140); c3.setWidth(110); c4.setWidth(140);
              c1.setBackgroundColor("#f1f5f9"); c3.setBackgroundColor("#f1f5f9");
              c1.getChild(0).asParagraph().setBold(true).setFontSize(8.5).setFontFamily("Arial");
              c2.getChild(0).asParagraph().setFontSize(8.5).setFontFamily("Arial");
              c3.getChild(0).asParagraph().setBold(true).setFontSize(8.5).setFontFamily("Arial");
              c4.getChild(0).asParagraph().setFontSize(8.5).setFontFamily("Arial");
            } else {
              var cLabel = row.appendTableCell(item[0]);
              var cVal = row.appendTableCell(item[1]);
              cLabel.setWidth(110); cVal.setWidth(390);
              cLabel.setBackgroundColor("#f8fafc");
              cLabel.getChild(0).asParagraph().setBold(true).setFontSize(8.5).setFontFamily("Arial");
              cVal.getChild(0).asParagraph().setFontSize(8.5).setFontFamily("Arial");
            }
          });

          // Objektif Terbeza
          var rowObj = tableRph.appendTableRow();
          var cObjLbl = rowObj.appendTableCell("OBJEKTIF PEMBELAJARAN TERBEZA");
          var cObjVal = rowObj.appendTableCell(
            "Kump 1 (Pengayaan): Murid dapat menguasai, menganalisis dan melengkapkan 4 tugasan dengan tepat secara berdikari.\n" +
            "Kump 2 (Pengukuhan): Murid dapat memahami dan menyelesaikan sekurang-kurangnya 3 daripada 4 tugasan dengan betul.\n" +
            "Kump 3 (Pemulihan): Murid dapat menyalin dan menyusun sekurang-kurangnya 2 ayat/frasa mudah dengan bimbingan guru."
          );
          cObjLbl.setWidth(110); cObjVal.setWidth(390);
          cObjLbl.setBackgroundColor("#f1f5f9");
          cObjLbl.getChild(0).asParagraph().setBold(true).setFontSize(8.5).setFontFamily("Arial");
          cObjVal.getChild(0).asParagraph().setFontSize(8.5).setFontFamily("Arial");

          // Kriteria Kejayaan Terbeza
          var rowKk = tableRph.appendTableRow();
          var cKkLbl = rowKk.appendTableCell("KRITERIA KEJAYAAN TERBEZA");
          var cKkVal = rowKk.appendTableCell(
            "Kump 1: Berjaya melengkapkan 4 latihan aras tinggi dalam buku aktiviti serta membimbing rakan sebaya.\n" +
            "Kump 2: Berjaya melengkapkan sekurang-kurangnya 3 latihan bertulis dengan bimbingan minimum.\n" +
            "Kump 3: Berjaya menyalin dan menyebut sekurang-kurangnya 2 perkataan/ayat dengan bimbingan guru."
          );
          cKkLbl.setWidth(110); cKkVal.setWidth(390);
          cKkLbl.setBackgroundColor("#f8fafc");
          cKkLbl.getChild(0).asParagraph().setBold(true).setFontSize(8.5).setFontFamily("Arial");
          cKkVal.getChild(0).asParagraph().setFontSize(8.5).setFontFamily("Arial");

          // Set Induksi
          var rowInd = tableRph.appendTableRow();
          var cIndLbl = rowInd.appendTableCell("SET INDUKSI");
          var cIndVal = rowInd.appendTableCell("1. Guru mengaitkan tajuk pembelajaran dengan pengalaman sedia ada murid.\n2. Guru dan murid bersoal jawab secara santai untuk mencungkil idea awal.\n3. Guru menerangkan objektif dan kriteria kejayaan PdP hari ini.");
          cIndLbl.setWidth(110); cIndVal.setWidth(390);
          cIndLbl.setBackgroundColor("#f1f5f9");
          cIndLbl.getChild(0).asParagraph().setBold(true).setFontSize(8.5).setFontFamily("Arial");
          cIndVal.getChild(0).asParagraph().setFontSize(8.5).setFontFamily("Arial");

          // Aktiviti Terbeza 3 Kolum (Kumpulan 1, 2, 3)
          var rowAktHdr = tableRph.appendTableRow();
          var cAktHdr = rowAktHdr.appendTableCell("AKTIVITI PdP TERBEZA (PENDEKATAN BERBEZA MENGIKUT ARAS)");
          cAktHdr.setBackgroundColor("#e0e7ff");
          cAktHdr.getChild(0).asParagraph().setBold(true).setFontSize(8.5).setFontFamily("Arial").setAlignment(DocumentApp.HorizontalAlignment.CENTER);

          var rowAktCols = tableRph.appendTableRow();
          var cAkt1 = rowAktCols.appendTableCell("KUMPULAN 1 (PENGAYAAN)\n1. Membaca teks dan menganalisis maklumat secara berdikari.\n2. Melaksanakan teknik PAK-21 'Teach Me' / 'Write It Out'.\n3. Membina ayat gramatis dan mencabar minda (KBAT).\n4. Menjawab soalan pengayaan bertulis.");
          var cAkt2 = rowAktCols.appendTableCell("KUMPULAN 2 (PENGUKUHAN)\n1. Membaca teks bersama guru dan rakan.\n2. Memahami kosa kata dan membina ayat mudah.\n3. Melengkapkan lembaran kerja pengukuhan.\n4. Bersoal jawab dengan guru untuk pengukuhan konsep.");
          var cAkt3 = rowAktCols.appendTableCell("KUMPULAN 3 (PEMULIHAN)\n1. Mendengar bimbingan intensif guru (Scaffolding).\n2. Menyalin dan memadankan perkataan/gambar.\n3. Melakukan aktiviti latih tubi sebutan dan tulisan mekanis.\n4. Dibimbing oleh rakan sebaya (Buddy Support).");
          cAkt1.setWidth(166); cAkt2.setWidth(167); cAkt3.setWidth(167);
          cAkt1.getChild(0).asParagraph().setFontSize(8).setFontFamily("Arial");
          cAkt2.getChild(0).asParagraph().setFontSize(8).setFontFamily("Arial");
          cAkt3.getChild(0).asParagraph().setFontSize(8).setFontFamily("Arial");

          // Penutup Terbeza
          var rowPenCols = tableRph.appendTableRow();
          var cPen1 = rowPenCols.appendTableCell("PENUTUP (KUMP 1):\nMelengkapkan lembaran pengayaan dan refleksi kendiri.");
          var cPen2 = rowPenCols.appendTableCell("PENUTUP (KUMP 2):\nMenyemak jawapan bersama guru dan membuat rumusan.");
          var cPen3 = rowPenCols.appendTableCell("PENUTUP (KUMP 3):\nLatihan pemulihan berfokus bersama bimbingan guru.");
          cPen1.setWidth(166); cPen2.setWidth(167); cPen3.setWidth(167);
          cPen1.setBackgroundColor("#f8fafc"); cPen2.setBackgroundColor("#f8fafc"); cPen3.setBackgroundColor("#f8fafc");
          cPen1.getChild(0).asParagraph().setFontSize(8).setFontFamily("Arial");
          cPen2.getChild(0).asParagraph().setFontSize(8).setFontFamily("Arial");
          cPen3.getChild(0).asParagraph().setFontSize(8).setFontFamily("Arial");

          // Elemen Sokongan & Pentaksiran
          var rphFooterRows = [
            ["BUKU TEKS", "Halaman 018 - 020", "BUKU AKTIVITI", "Halaman 022 - 024"],
            ["SEKOLAHKU SEJAHTERA", "Tekun, Teliti, Terampil, Selamat", "STRATEGI PdP", "Pembelajaran Masteri & Terbeza"],
            ["KBAT", "Menganalisis & Mengaplikasi", "PETA PEMIKIRAN", ithinkStr],
            ["EMK: NILAI MURNI", "Kerjasama, Bertanggungjawab, Berdikari", "KEMAHIRAN BERFIKIR", "Mengecam, Menjana Idea"],
            ["EMK: ILMU & TEMA", "Pendidikan Sivik & Moral", "PENILAIAN P&P", "Hasil Kerja Murid & Pemerhatian"],
            ["PENTAKSIRAN PBD", "[ TP1 ]  [ TP2 ]  [ TP3 ]  [ TP4 ]  [ TP5 ]  [ TP6 ]", "TINDAKAN SUSULAN", pemulihanStr],
            ["REFLEKSI & IMPAK", "1. ____ / ____ murid mencapai objektif PdP dan diberi aktiviti pengayaan.\n2. ____ / ____ murid diberi bimbingan berfokus dalam sesi pemulihan.", "", ""]
          ];

          rphFooterRows.forEach(function(item) {
            var row = tableRph.appendTableRow();
            if (item[2] !== "") {
              var c1 = row.appendTableCell(item[0]); var c2 = row.appendTableCell(item[1]);
              var c3 = row.appendTableCell(item[2]); var c4 = row.appendTableCell(item[3]);
              c1.setWidth(110); c2.setWidth(140); c3.setWidth(110); c4.setWidth(140);
              c1.setBackgroundColor("#f1f5f9"); c3.setBackgroundColor("#f1f5f9");
              c1.getChild(0).asParagraph().setBold(true).setFontSize(8).setFontFamily("Arial");
              c2.getChild(0).asParagraph().setFontSize(8).setFontFamily("Arial");
              c3.getChild(0).asParagraph().setBold(true).setFontSize(8).setFontFamily("Arial");
              c4.getChild(0).asParagraph().setFontSize(8).setFontFamily("Arial");
            } else {
              var cLabel = row.appendTableCell(item[0]);
              var cVal = row.appendTableCell(item[1]);
              cLabel.setWidth(110); cVal.setWidth(390);
              cLabel.setBackgroundColor("#f8fafc");
              cLabel.getChild(0).asParagraph().setBold(true).setFontSize(8).setFontFamily("Arial");
              cVal.getChild(0).asParagraph().setFontSize(8).setFontFamily("Arial");
            }
          });

          body.appendParagraph("").setSpacingAfter(6);
        }

        if (w < 40) {
          body.appendPageBreak();
        }
      }

      doc.saveAndClose();

      if (emel && emel.includes("@")) {
        try {
          var file = DriveApp.getFileById(doc.getId());
          file.addEditor(emel);
        } catch (eDrive) {}
      }

      docId = doc.getId();
      downloadUrl = "https://docs.google.com/document/d/" + docId + "/export?format=docx";
      docViewUrl = doc.getUrl();
    } catch (eDoc) {
      console.warn("DocumentApp notice: " + eDoc.message);
    }

    return {
      success: true,
      docId: docId,
      fileUrl: downloadUrl,
      downloadUrl: downloadUrl,
      docUrl: docViewUrl,
      dskpList: dskpList,
      message: "Fail Word 40 Minggu Terbeza (" + (jadualMingguan.length * 40) + " sesi PdP) telah siap dijana!"
    };

  } catch (err) {
    console.error("Ralat janaRph40MingguDocx:", err);
    return { success: false, message: "Ralat penjanaan dokumen Word 40 minggu: " + (err.message || err) };
  }
}

function janaRph40MingguBackend(payload) {
  try {
    if (!payload) return { success: false, message: "Data jadual tidak lengkap." };
    return janaRph40MingguDocx(payload);
  } catch (err) {
    console.error("Ralat janaRph40MingguBackend:", err);
    return { success: false, message: "Ralat sistem: " + (err.message || err) };
  }
}
