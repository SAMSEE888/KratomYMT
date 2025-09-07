https://script.google.com/macros/s/AKfycbwY_qZiHLfucVQmiGHb_fa46d_X68Blq_BULa8q6FQBJSjwJYtEM1XBK1xVple_Fcw3/exec

1lu6Y9Xq7fEme3I1Z10IRIb4wqEk_QOZIqWV7oy5jNGE


// ID ของ Google Sheet ที่ต้องการบันทึกข้อมูล
const SHEET_ID = '1lu6Y9Xq7fEme3I1Z10IRIb4wqEk_QOZIqWV7oy5jNGE';
// ชื่อของชีต (แท็บ) ที่ต้องการบันทึกข้อมูล
const SHEET_NAME = 'SalesData'; // <-- **สำคัญ:** กรุณาสร้างชีตชื่อนี้ในไฟล์ของคุณ

/**
 * ฟังก์ชันหลักที่ทำงานเมื่อมีการส่งข้อมูลแบบ POST เข้ามา
 */
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

    // ตรวจสอบและสร้างหัวตารางหากยังไม่มี
    if (sheet.getLastRow() < 1) {
      const headers = [
        'Timestamp', 'Date', 'Sold (bottles)', 'Pending (bottles)', 'Cleared (bottles)',
        'Pipe Fee', 'Share Fee', 'Other Fee', 'Save Fee',
        'Revenue', 'Expense', 'Balance'
      ];
      sheet.appendRow(headers);
    }
    
    // แปลงข้อมูลที่ส่งมา (JSON)
    const requestData = JSON.parse(e.postData.contents);
    
    // ตรวจสอบว่าข้อมูลที่ส่งมาเป็น array (สำหรับ offline sync) หรือ object เดียว
    const records = Array.isArray(requestData) ? requestData : [requestData];

    // วนลูปเพื่อบันทึกทุกรายการข้อมูล
    records.forEach(record => {
      const newRow = [
        new Date(),
        record.date,
        record.sold,
        record.pending,
        record.cleared,
        record.pipeFee,
        record.shareFee,
        record.otherFee,
        record.saveFee,
        record.revenue,
        record.expense,
        record.balance
      ];
      sheet.appendRow(newRow);
    });
    
    // ส่งคำตอบกลับไปว่าสำเร็จ
    return ContentService
      .createTextOutput(JSON.stringify({ 'status': 'success', 'message': 'Data saved successfully', 'records_count': records.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // กรณีเกิดข้อผิดพลาด ให้ส่ง error กลับไป
    return ContentService
      .createTextOutput(JSON.stringify({ 'status': 'error', 'message': error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
