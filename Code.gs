// ═══════════════════════════════════════════════════════════
// KION Field App — Google Apps Script (Single Sheet)
// يرسل البيانات إلى شيت واحد: "بيانات الأدارة المركزة"
// 18 عمود بالضبط: A=ID, B=التاريخ, C=اسم المندوب ... R=التقييم
// ═══════════════════════════════════════════════════════════

const SHEET_NAME = 'بيانات الأدارة المركزة';

function doPost(e) {
  try {
    // Parse incoming JSON
    const data = JSON.parse(e.postData.contents);

    // Open the active spreadsheet and target sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      // Add headers (18 columns)
      const headers = [
        'ID','التاريخ','اسم المندوب','اسم العميل','النشاط','التخصص',
        'الهاتف','المنطقه','العنوان','GPS','تاريخ الزياره القاتدمه',
        'الاجراء','المنافس','سعر المنافس','نقطة الضعف','الملاحظات',
        'التصنيف','التقييم'
      ];
      sheet.appendRow(headers);
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#1a1a1a').setFontColor('#c9a96e').setFontWeight('bold');
    }

    // Build row array in exact column order (18 columns)
    const row = [
      data.id || '',
      data.date || '',
      data.rep || '',
      data.name || '',
      data.type || '',
      data.specialty || '',
      data.phone || '',
      data.area || '',
      data.address || '',
      data.gps || '',
      data.nextVisitDate || '',
      data.nextAction || '',
      data.compName || '',
      data.compPrice || '',
      data.compWeak || '',
      data.notes || '',
      data.qual || '',
      data.rating || 0
    ];

    // Append to sheet
    sheet.appendRow(row);

    // Return success
    return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'تم الحفظ' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// CORS preflight support
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// Test function (Run this in Apps Script editor to verify sheet connection)
function testConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log('Sheet not found, creating...');
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      'ID','التاريخ','اسم المندوب','اسم العميل','النشاط','التخصص',
      'الهاتف','المنطقه','العنوان','GPS','تاريخ الزياره القاتدمه',
      'الاجراء','المنافس','سعر المنافس','نقطة الضعف','الملاحظات',
      'التصنيف','التقييم'
    ];
    sheet.appendRow(headers);
  }
  Logger.log('Sheet ready: ' + sheet.getName());
  Logger.log('Last row: ' + sheet.getLastRow());
}