// =============================================
// Gold Portfolio Pro — Google Apps Script
// Sheet "Users" kolom: A=username, B=password
// Sheet "Transaksi" kolom: A=nama,B=gram,C=harga,D=modal,E=tgl
// =============================================

function doGet(e) {
  var params = e.parameter;
  
  // === LOGIN VERIFICATION ===
  if (params.action === 'login') {
    var user = (params.user || '').toLowerCase().trim();
    var pass = params.pass || '';
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName('Users');
    
    if (!userSheet) {
      return jsonResponse({ success: false, message: 'Sheet Users tidak ditemukan.' });
    }
    
    var data = userSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var sheetUser = String(data[i][0]).toLowerCase().trim();
      var sheetPass = String(data[i][1]).trim();
      if (sheetUser === user && sheetPass === pass) {
        return jsonResponse({ success: true, message: 'Login berhasil' });
      }
    }
    return jsonResponse({ success: false, message: 'Username atau password salah.' });
  }
  
  // === GET TRANSACTIONS ===
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Transaksi') || ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({
      nama: data[i][0],
      gram: Number(data[i][1]),
      harga: Number(data[i][2]),
      modal: Number(data[i][3]),
      tgl: data[i][4]
    });
  }
  return jsonResponse(rows);
}

function doPost(e) {
  var params = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Transaksi') || ss.getActiveSheet();
  
  sheet.clearContents();
  sheet.appendRow(['Nama', 'Gram', 'Harga Beli', 'Modal', 'Tanggal']);
  
  params.forEach(function(r) {
    sheet.appendRow([r.nama, r.gram, r.harga, r.modal, r.tgl]);
  });
  
  return jsonResponse({ status: 'success' });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
