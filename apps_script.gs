// =====================================================
// 구글 Apps Script — 자산진단 신청 데이터 수신기
// =====================================================
// 사용법:
// 1. 구글 스프레드시트 열기
// 2. 상단 메뉴 → 확장 프로그램 → Apps Script
// 3. 이 코드 전체 붙여넣기 (기존 내용 삭제 후)
// 4. 저장 → 배포 → 새 배포 → 웹 앱
//    · 다음 사용자로 실행: 나
//    · 액세스 권한: 모든 사용자
// 5. 배포 후 나오는 URL 복사 → 랜딩페이지에 붙여넣기
// =====================================================

const SHEET_NAME = '신청자목록'; // 시트 이름 (변경 가능)

function doPost(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName(SHEET_NAME);

    // 시트 없으면 자동 생성 + 헤더 세팅
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['접수일시', '이름', '연락처', '주요고민', '추가메모', '처리상태']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#d4ff00').setFontColor('#000000');
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 80);
      sheet.setColumnWidth(3, 130);
      sheet.setColumnWidth(4, 220);
      sheet.setColumnWidth(5, 220);
      sheet.setColumnWidth(6, 100);
    }

    const data = JSON.parse(e.postData.contents);

    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      now,
      data.name    || '',
      data.phone   || '',
      data.concern || '',
      data.memo    || '',
      '미확인',
    ]);

    // 새 신청 들어오면 이메일 알림 (원하면 주석 해제 + 이메일 입력)
    // sendAlert(data, now);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 이메일 알림 함수 (선택사항)
function sendAlert(data, time) {
  const EMAIL = 'your@email.com'; // ← 본인 이메일로 변경
  const subject = `[자산진단] 새 신청 — ${data.name}`;
  const body = `
새로운 자산진단 신청이 들어왔습니다.

접수일시: ${time}
이  름: ${data.name}
연락처: ${data.phone}
주요고민: ${data.concern}
추가메모: ${data.memo || '없음'}

스프레드시트에서 확인하세요.
  `.trim();
  MailApp.sendEmail(EMAIL, subject, body);
}

// GET 요청 테스트용 (배포 확인용)
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'running' }))
    .setMimeType(ContentService.MimeType.JSON);
}
