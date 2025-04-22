function myFunction() {

}



// 스프레드시트가 열릴 때 메뉴를 생성합니다.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('커스텀 메뉴')
    .addItem('주소분리', 'splitAddress')
    .addItem('운임결정', 'copyValuesToR')
    .addItem('행추가', 'addBlankRows')
    .addItem('아이디추가', 'fillUniqueIds')
    .addToUi();
}

// 행추가
function addBlankRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getName() === 'CS이슈 (택배, 개인, 대리점)') {

    var range = sheet.getDataRange(); // 시트의 전체 데이터 범위 가져오기
    var numColumns = range.getNumColumns(); // 열 개수 가져오기

    var columns = Array.from({ length: numColumns }, (_, i) => i + 1); // [1, 2, 3, ...] 모든 열 번호 배열 생성
    range.removeDuplicates(columns); // 모든 열을 기준으로 중복 제거

    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    range.sort([
      { column: 1, ascending: false },  
      { column: 2, ascending: true }  
    ]);
  }

  // sheet.insertRows(2, 10);
}

// 전화받을 때 마다 전화번호를 손으로 쓰는게 싫은데, 조금이라도 덜 쓰려고
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  var colNum = e.range.getColumn();

  // 특정 시트의 특정 열만 트리거
  if (sheetName === "CS 전화상담기록" && colNum === 2) {
    var value = e.range.getValue().toString();
    if (/^\d{8}$/.test(value)) {
      var formattedNumber = "010-" + value.slice(0, 4) + "-" + value.slice(4);
      e.range.setValue(formattedNumber);
    }
    var row = range.getRow();
    var dateCell = sheet.getRange(row, 1);  // 같은 행의 1열(A열) 선택

    // 1열이 비어있는지 확인
    if (dateCell.getValue() === "") {
      // 현재 날짜와 시간을 가져옵니다
      var now = new Date();

      // 날짜와 시간을 원하는 형식으로 포맷팅합니다
      var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

      // 날짜와 시간을 1열에 입력합니다
      dateCell.setValue(formattedDate);

      // // 현재 행 바로 아래에 새 행을 삽입합니다
      // // 실행이 너무 느려서 그 와중에 다른 편집을 해버리면 이상하게 작동한다
      // sheet.insertRowAfter(row);
    }
  }

  // 특정 시트의 특정 열만 트리거
  // cs 이슈 전화번호 포멧
  if (sheetName === "CS이슈 (택배, 개인, 대리점)" && colNum === 8) {
    var value = e.range.getValue().toString();
    // Logger.log(value);
    // 8자리: 010-0000-0000
    if (/^\d{8}$/.test(value)) {
      var formattedNumber = "010-" + value.slice(0, 4) + "-" + value.slice(4);
      e.range.setValue(formattedNumber);
    }

    // 10~12자리: [가변 앞자리]-[중간 4자리]-[마지막 4자리]
    else if (/^\d{10,12}$/.test(value)) {
      var headLen = value.length - 8; // 앞자리는 나머지
      var formattedNumber = value.slice(0, headLen) + "-" + value.slice(headLen, headLen + 4) + "-" + value.slice(-4);
      e.range.setValue(formattedNumber);
    }
  }

  // // 특정 시트의 특정 열만 트리거
  // // 판매기록[경동]
  // if (sheetName === "판매기록[경동]" && colNum === 16) {
  //   var value = e.range.getValue();
  //   Logger.log(value, typeof value);
  //   // 숫자이면 정수로 라운드업해서 다시 입력
  //   if (typeof value === "number") {
  //     var rounded = Math.ceil(value);
  //     e.range.setValue(rounded);
  //   }
  // }
}

// 도로명 주소 분리 함수
function splitAddress() {
  var sheet = SpreadsheetApp.getActiveSheet();
  // 몇번째 행까지 적용할지?
  var lastRow = 300;
  var addresses = sheet.getRange(2, 11, lastRow, 1).getValues();

  for (var i = 0; i < addresses.length; i++) {
    var address = addresses[i][0];
    var parts = address.split(' ');
    var roadAddressParts = [];
    var detailAddress = '';

    for (var j = 0; j < parts.length; j++) {
      if (parts[j].includes('로') || parts[j].includes('길')) {
        roadAddressParts.push(parts[j]);
        if (j + 1 < parts.length && /^\d+/.test(parts[j + 1])) {
          roadAddressParts.push(parts[j + 1]);
          detailAddress = parts.slice(j + 2).join(' ');
          break;
        }
      } else {
        roadAddressParts.push(parts[j]);
      }
    }

    var roadAddress = roadAddressParts.join(' ');

    // 결과를 입력할 셀
    sheet.getRange(i + 2, 25).setValue(roadAddress);
    sheet.getRange(i + 2, 26).setValue(detailAddress);
  }
}

// function copyValuesToR() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   const today = new Date();
//   const data = sheet.getDataRange().getValues();
//   const lastRow = 100;

//   for (let i = 1; i < lastRow; i++) { // Skip header (assumes row 1 is header)
//     const dateInA = new Date(data[i][0]); // A열
//     const columnM = data[i][12]; // M열
//     const columnN = data[i][13]; // N열
//     const columnR = data[i][17]; // R열
//     const columnAE = data[i][30]; // AE열

//     // 조건: A열이 오늘 날짜이고, M과 N이 비어있지 않고, AE가 비어있지 않고, R이 비어 있는 경우
//     if (dateInA.toDateString() === today.toDateString() && columnM && columnN && !columnR && columnAE) {

//       // R열에 AE열 값과 N열 값을 곱한 결과 복사
//       const multipliedValue = columnAE * columnN; // N열의 값과 AE열의 값 곱하기
//       sheet.getRange(i + 1, 18).setValue(multipliedValue); // R열: 18번째 컬럼

//       // 추가 값 입력
//       sheet.getRange(i + 1, 15).setValue('박스'); // O열: 15번째 컬럼
//       sheet.getRange(i + 1, 16).setValue(156);   // P열: 16번째 컬럼
//       sheet.getRange(i + 1, 17).setValue('현택'); // Q열: 17번째 컬럼
//       sheet.getRange(i + 1, 20).setValue(100);   // T열: 20번째 컬럼
//     }
//   }
// }

function fillUniqueIds() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("v2:v" + sheet.getLastRow()).getValues(); // v열의 데이터 가져오기

  var updates = [];

  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) { // 비어있는 경우만 처리
      // updates.push([Utilities.getUuid()]);
      updates.push([Math.random().toString(36).substring(2, 10)]);
    } else {
      updates.push([data[i][0]]);
    }
  }

  // 업데이트 수행
  if (updates.length > 0) {
    sheet.getRange("v2:v" + sheet.getLastRow()).setValues(updates);
  }
}

