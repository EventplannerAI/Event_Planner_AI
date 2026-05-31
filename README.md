To exchange the scripts to google sheet pleás open thís programme - thís the frist demo

const SHEET_NAME = 'EventPlanner';
const NOTIFY_EMAIL = 'insa.k16.nguyendinhhoang@gmail.com'; // ← nhập email của bạn để nhận thông báo

// Hàm ghi dữ liệu vào sheet
function writeToSheet(data) {
  const ss = SpreadsheetApp.openById('1DqS5Xv45eaya2r6YnicIR0r6oNJUqy2VQ-8fzKBpvSY');
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Tạo sheet + header nếu chưa có
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Thời gian ghi', 'Loại sự kiện', 'Người tổ chức', 'Email',
      'Ngày & Giờ', 'Thời lượng', 'Địa điểm', 'Địa chỉ chi tiết',
      'Ghi chú địa điểm', 'Số khách', 'Ngân sách', 'Tiền tệ',
      'Phong cách', 'Yêu cầu đặc biệt', 'Số thiệp mời', 'Danh sách email mời'
    ]);
    sheet.getRange(1, 1, 1, 16).setFontWeight('bold').setBackground('#1A1614').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  // Thêm dòng dữ liệu mới
  sheet.appendRow([
    data.timestamp, data.eventType, data.organizer, data.email,
    data.dateTime, data.duration, data.location, data.locationAddr,
    data.locationNotes, data.guests, data.budget, data.currency,
    data.styles, data.specialReqs, data.inviteCount, data.inviteEmails
  ]);

  // Gửi email thông báo nếu đã cấu hình
  if (NOTIFY_EMAIL) {
    const subject = `[Event Planner] Đăng ký mới: ${data.eventType} — ${data.organizer}`;
    const body = Object.entries(data).map(([k,v]) => `${k}: ${v}`).join('\n');
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  }
}

// Nhận dữ liệu qua GET (từ form dùng mode: no-cors)
function doGet(e) {
  try {
    if (!e.parameter || !e.parameter.data) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'No data' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const data = JSON.parse(decodeURIComponent(e.parameter.data));
    writeToSheet(data);
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Nhận dữ liệu qua POST (fallback)
function doPost(e) {
  try {
    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.data) {
      data = JSON.parse(decodeURIComponent(e.parameter.data));
    } else {
      throw new Error('Không có dữ liệu');
    }
    writeToSheet(data);
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

