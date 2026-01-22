/**
 * Monitoring BBM - PDSI
 * Core Backend System
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Monitoring BBM - PDSI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- SISTEM AUTENTIKASI ---
function checkLogin(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users") || createDefaultUserSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === password) {
      return { success: true, role: data[i][2], user: username };
    }
  }
  return { success: false, message: "Username atau Password salah!" };
}

function createNewUser(newUsername, newPassword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users") || createDefaultUserSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === newUsername) return "❌ Error: Username sudah terdaftar!";
  }
  sheet.appendRow([newUsername, newPassword, "Admin"]);
  return "✅ User " + newUsername + " berhasil dibuat!";
}

function changePassword(username, oldPass, newPass) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      if (data[i][1] === oldPass) {
        sheet.getRange(i + 1, 2).setValue(newPass);
        return "✅ Password berhasil diperbarui!";
      } else { return "❌ Error: Password lama salah!"; }
    }
  }
  return "❌ Error: User tidak ditemukan!";
}

function createDefaultUserSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.insertSheet("Users");
  sheet.appendRow(["Username", "Password", "Role"]);
  sheet.appendRow(["admin", "pdsi2026", "Master"]);
  return sheet;
}

// --- LOGIKA DATA BBM ---
function getVoucherList() {
  var vouchers = [];
  for (var i = 1; i <= 1000; i++) {
    vouchers.push("PDSI-ROII/BBM/2026/" + ("00" + i).slice(-3));
  }
  return vouchers;
}

function simpanData(obj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DataBBM") || ss.insertSheet("DataBBM");
    sheet.appendRow([
      obj.kodeVoucher, obj.hari, obj.tanggal, obj.jenisKendaraan,
      obj.plate, obj.odometer, obj.liter, obj.nettPrice,
      obj.totalPrice, obj.jenisBBM, obj.namaDriver, obj.dariFungsi,
      obj.fungsiDilayani, obj.keperluan
    ]);
    return "✅ Data Berhasil Disimpan!";
  } catch (e) { return "❌ Gagal: " + e.toString(); }
}

function updateData(obj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DataBBM");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === obj.kodeVoucher) {
        sheet.getRange(i + 1, 1, 1, 14).setValues([[
          obj.kodeVoucher, obj.hari, obj.tanggal, obj.jenisKendaraan,
          obj.plate, obj.odometer, obj.liter, obj.nettPrice,
          obj.totalPrice, obj.jenisBBM, obj.namaDriver, obj.dariFungsi,
          obj.fungsiDilayani, obj.keperluan
        ]]);
        return "✅ Data " + obj.kodeVoucher + " Berhasil Diperbarui!";
      }
    }
    return "⚠️ Voucher tidak ditemukan.";
  } catch (e) { return "❌ Gagal: " + e.toString(); }
}

function hapusData(voucher) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DataBBM");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === voucher) {
        sheet.deleteRow(i + 1);
        return "🗑️ Data Berhasil Dihapus!";
      }
    }
    return "Data tidak ditemukan.";
  } catch (e) { return "Gagal: " + e.toString(); }
}

function tampilkanData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DataBBM");
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  return values.map(row => row.map(cell => cell instanceof Date ? Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd") : cell));
}

function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DataBBM");
  if (!sheet) return null;
  var values = sheet.getDataRange().getValues();
  var stats = { totalLiter: 0, totalCost: 0, bbmCounts: {}, vehicleCounts: {} };
  if (values.length <= 1) return stats;
  values.slice(1).forEach(r => {
    stats.totalLiter += (parseFloat(r[6]) || 0);
    stats.totalCost += (parseFloat(r[8]) || 0);
    stats.bbmCounts[r[9]] = (stats.bbmCounts[r[9]] || 0) + 1;
    stats.vehicleCounts[r[3]] = (stats.vehicleCounts[r[3]] || 0) + 1;
  });
  stats.bbmLabels = Object.keys(stats.bbmCounts);
  stats.bbmValues = Object.values(stats.bbmCounts);
  stats.vehicleLabels = Object.keys(stats.vehicleCounts);
  stats.vehicleValues = Object.values(stats.vehicleCounts);
  stats.totalUnit = stats.vehicleLabels.length;
  return stats;
}