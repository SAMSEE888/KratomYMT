/**
 * Google Apps Script สำหรับระบบบันทึกยอดขายน้ำกระท่อม
 * โดย SSKratomYMT
 * 
 * ฟังก์ชันหลัก:
 * - doGet: สำหรับทดสอบการทำงาน
 * - doPost: รับข้อมูลจากเว็บแอปและบันทึกลง Google Sheet
 * - getWeeklyReport: สร้างรายงานประจำสัปดาห์
 * - getMonthlyReport: สร้างรายงานประจำเดือน
 */

// ตั้งค่าข้อมูล Spreadsheet
const SPREADSHEET_ID = '11vhg37MbHRm53SSEHLsCI3EBXx5_meXVvlRuqhFteaY';
const SHEET_NAME = 'SalesData';
const REPORT_SHEET = 'Reports';
const PRICE_PER_BOTTLE = 40;

/**
 * ฟังก์ชันสำหรับทดสอบการทำงาน (GET)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'SSKratomYMT API is working',
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * ฟังก์ชันหลักสำหรับรับข้อมูลจากเว็บแอป (POST)
 */
function doPost(e) {
  try {
    // ตรวจสอบข้อมูลที่ได้รับ
    if (!e || !e.postData || !e.postData.contents) {
      return createResponse(400, 'Invalid request');
    }

    // แปลงข้อมูล JSON
    const data = JSON.parse(e.postData.contents);
    
    // ตรวจสอบข้อมูลที่จำเป็น
    if (!data.date || isNaN(data.sold) || isNaN(data.revenue) || isNaN(data.expense) || isNaN(data.balance)) {
      return createResponse(400, 'Missing required fields');
    }

    // บันทึกข้อมูลลง Google Sheet
    const result = saveToSheet(data);
    
    // ส่งคำตอบกลับ
    return createResponse(200, 'Data saved successfully', {
      spreadsheetId: SPREADSHEET_ID,
      sheetName: SHEET_NAME,
      row: result.row,
      data: result.data
    });
    
  } catch (error) {
    // กรณีเกิดข้อผิดพลาด
    return createResponse(500, 'Error: ' + error.message);
  }
}

/**
 * บันทึกข้อมูลลงใน Google Sheet
 */
function saveToSheet(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  
  // เพิ่มหัวข้อถ้ายังไม่มี
  if (sheet.getLastRow() === 0) {
    const headers = [
      'วันที่', 'ขายได้ (ขวด)', 'ค้างน้ำดิบ (ขวด)', 'เคลียร์ค้างน้ำดิบ (ขวด)',
      'รายรับ (บาท)', 'ค่าท่อม', 'ค่าแชร์', 'ค่าใช้จ่ายอื่น', 
      'เก็บออมเงิน', 'รายจ่าย (บาท)', 'ยอดคงเหลือ (บาท)', 'เวลาบันทึก'
    ];
    sheet.appendRow(headers);
  }
  
  // เตรียมข้อมูลที่จะบันทึก
  const rowData = [
    data.date,
    parseFloat(data.sold) || 0,
    parseFloat(data.pending) || 0,
    parseFloat(data.cleared) || 0,
    parseFloat(data.revenue) || 0,
    parseFloat(data.pipeFee) || 0,
    parseFloat(data.shareFee) || 0,
    parseFloat(data.otherFee) || 0,
    parseFloat(data.saveFee) || 0,
    parseFloat(data.expense) || 0,
    parseFloat(data.balance) || 0,
    new Date().toISOString()
  ];
  
  // บันทึกข้อมูล
  sheet.appendRow(rowData);
  
  // จัดรูปแบบเซลล์
  const lastRow = sheet.getLastRow();
  formatSheet(sheet, lastRow);
  
  // สร้างรายงานอัตโนมัติ
  generateAutoReports();
  
  return {
    row: lastRow,
    data: rowData
  };
}

/**
 * จัดรูปแบบ Google Sheet
 */
function formatSheet(sheet, row) {
  // รูปแบบวันที่
  sheet.getRange(row, 1).setNumberFormat('yyyy-mm-dd');
  
  // รูปแบบตัวเลข
  const numberColumns = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  numberColumns.forEach(col => {
    sheet.getRange(row, col).setNumberFormat('#,##0');
  });
  
  // รูปแบบเวลาบันทึก
  sheet.getRange(row, 12).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  
  // สีพื้นหลังสลับแถว
  if (row % 2 === 0) {
    sheet.getRange(row, 1, 1, 12).setBackground('#f5f5f5');
  }
  
  // กรอบตาราง
  sheet.getRange(row, 1, 1, 12).setBorder(true, true, true, true, true, true);
}

/**
 * สร้างรายงานอัตโนมัติ
 */
function generateAutoReports() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dataSheet = ss.getSheetByName(SHEET_NAME);
  const reportSheet = ss.getSheetByName(REPORT_SHEET) || ss.insertSheet(REPORT_SHEET);
  
  // ลบข้อมูลเก่าใน sheet รายงาน
  reportSheet.clear();
  
  // สร้างรายงานประจำสัปดาห์
  const weeklyReport = getWeeklyReport(dataSheet);
  reportSheet.getRange(1, 1).setValue('รายงานประจำสัปดาห์').setFontWeight('bold');
  reportSheet.getRange(2, 1, weeklyReport.length, weeklyReport[0].length).setValues(weeklyReport);
  
  // สร้างรายงานประจำเดือน
  const monthlyReport = getMonthlyReport(dataSheet);
  reportSheet.getRange(weeklyReport.length + 3, 1).setValue('รายงานประจำเดือน').setFontWeight('bold');
  reportSheet.getRange(weeklyReport.length + 4, 1, monthlyReport.length, monthlyReport[0].length).setValues(monthlyReport);
  
  // จัดรูปแบบรายงาน
  formatReportSheet(reportSheet);
}

/**
 * สร้างรายงานประจำสัปดาห์
 */
function getWeeklyReport(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // กรองข้อมูล 7 วันที่ผ่านมา
  const oneWeekAgo = new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
  
  const weeklyData = data.filter((row, index) => {
    if (index === 0) return false; // ข้ามหัวข้อ
    const rowDate = new Date(row[0]);
    return rowDate >= oneWeekAgo;
  });
  
  // สรุปข้อมูลรายสัปดาห์
  const report = [
    ['วันที่', 'ขายได้ (ขวด)', 'รายรับ (บาท)', 'รายจ่าย (บาท)', 'ยอดคงเหลือ (บาท)']
  ];
  
  weeklyData.forEach(row => {
    report.push([
      row[0], // วันที่
      row[1], // ขายได้
      row[4], // รายรับ
      row[9], // รายจ่าย
      row[10] // ยอดคงเหลือ
    ]);
  });
  
  // เพิ่มแถวสรุป
  const summaryRow = ['รวม'];
  for (let i = 1; i <= 4; i++) {
    const sum = weeklyData.reduce((acc, row) => acc + (row[i] || 0), 0);
    summaryRow.push(sum);
  }
  report.push(summaryRow);
  
  return report;
}

/**
 * สร้างรายงานประจำเดือน
 */
function getMonthlyReport(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // กรองข้อมูล 30 วันที่ผ่านมา
  const oneMonthAgo = new Date();
  oneMonthAgo.setDate(oneMonthAgo.getDate() - 30);
  
  const monthlyData = data.filter((row, index) => {
    if (index === 0) return false; // ข้ามหัวข้อ
    const rowDate = new Date(row[0]);
    return rowDate >= oneMonthAgo;
  });
  
  // จัดกลุ่มข้อมูลตามเดือน
  const monthlyGroups = {};
  monthlyData.forEach(row => {
    const date = new Date(row[0]);
    const monthYear = `${date.getMonth() + 1}/${date.getFullYear()}`;
    
    if (!monthlyGroups[monthYear]) {
      monthlyGroups[monthYear] = [];
    }
    monthlyGroups[monthYear].push(row);
  });
  
  // สร้างรายงาน
  const report = [
    ['เดือน/ปี', 'ขายได้ (ขวด)', 'รายรับ (บาท)', 'รายจ่าย (บาท)', 'ยอดคงเหลือ (บาท)', 'วันที่มีการขาย']
  ];
  
  // เพิ่มข้อมูลแต่ละเดือน
  for (const [monthYear, rows] of Object.entries(monthlyGroups)) {
    const sold = rows.reduce((acc, row) => acc + (row[1] || 0), 0);
    const revenue = rows.reduce((acc, row) => acc + (row[4] || 0), 0);
    const expense = rows.reduce((acc, row) => acc + (row[9] || 0), 0);
    const balance = rows.reduce((acc, row) => acc + (row[10] || 0), 0);
    const days = rows.length;
    
    report.push([monthYear, sold, revenue, expense, balance, days]);
  }
  
  return report;
}

/**
 * จัดรูปแบบ sheet รายงาน
 */
function formatReportSheet(sheet) {
  // หาจำนวนแถวและคอลัมน์
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  // จัดรูปแบบหัวข้อ
  sheet.getRange(1, 1, 1, lastCol).merge().setFontSize(16).setFontWeight('bold');
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  // หาตำแหน่งรายงานเดือน
  const weeklyLastRow = sheet.createTextFinder('รวม').findNext().getRow();
  sheet.getRange(weeklyLastRow + 2, 1).setFontSize(16).setFontWeight('bold');
  
  // จัดรูปแบบตัวเลข
  sheet.getDataRange().setNumberFormat('#,##0');
  
  // จัดรูปแบบวันที่
  const dateCol = sheet.createTextFinder('วันที่').findNext().getColumn();
  sheet.getRange(2, dateCol, weeklyLastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
  
  // สีพื้นหลังหัวตาราง
  const headers = sheet.createTextFinder('วันที่').findAll();
  headers.forEach(header => {
    const headerRow = header.getRow();
    sheet.getRange(headerRow, 1, 1, lastCol)
      .setBackground('#4CAF50')
      .setFontColor('white')
      .setFontWeight('bold');
  });
  
  // สีพื้นหลังแถวสรุป
  const summaries = sheet.createTextFinder('รวม').findAll();
  summaries.forEach(summary => {
    const summaryRow = summary.getRow();
    sheet.getRange(summaryRow, 1, 1, lastCol)
      .setBackground('#81C784')
      .setFontWeight('bold');
  });
  
  // ปรับความกว้างคอลัมน์
  for (let i = 1; i <= lastCol; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // กรอบตาราง
  sheet.getDataRange().setBorder(true, true, true, true, true, true);
}

/**
 * สร้างคำตอบ JSON
 */
function createResponse(status, message, data = {}) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  output.setContent(JSON.stringify({
    status: status === 200 ? 'success' : 'error',
    message: message,
    timestamp: new Date().toISOString(),
    data: data
  }));
  return output;
}

/**
 * ฟังก์ชันสำหรับทดสอบใน Editor
 */
function testSaveData() {
  const testData = {
    date: '2023-06-15',
    sold: '50',
    pending: '2',
    cleared: '1',
    revenue: '1960',
    pipeFee: '200',
    shareFee: '100',
    otherFee: '50',
    saveFee: '500',
    expense: '850',
    balance: '1110'
  };
  
  const result = saveToSheet(testData);
  Logger.log(result);
}

/**
 * ฟังก์ชันสำหรับทดสอบการสร้างรายงาน
 */
function testGenerateReports() {
  generateAutoReports();
  Logger.log('Reports generated successfully');
}
