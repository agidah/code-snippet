function onEdit(e) {
  // e가 정의되지 않았다면 함수를 종료하기!!!!
  if (!e) return;

  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  
  // 변경을 추적할 시트 이름들을 배열에 넣기!!!
  var trackedSheets = ["이용개시정보RAW"];
  
  // 현재 편집된 시트가 추적 대상인지 확인
  if (trackedSheets.indexOf(sheetName) !== -1) {
    // 변경된 사용자의 이메일 가져오기
    var user = Session.getActiveUser().getEmail();
    
    // 현재 시간 가져오기
    var time = new Date().toLocaleString();
    
    // 수정된 셀의 정보 가져오기
    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();
    var oldValue = e.oldValue;
    var newValue = range.getValue();
    
    // 수정 유형 확인
    var editType = "";
    if (oldValue === undefined && newValue !== null && newValue !== "") {
      editType = "입력";
    } else if (oldValue !== undefined && (newValue === null || newValue === "")) {
      editType = "삭제";
    } else {
      editType = "수정";
    }
    
    // 수정 정보 문자열 생성
    var editInfo = "셀 " + range.getA1Notation() + "에서 " + editType + ": ";
    if (editType === "입력") {
      editInfo += "새 값 '" + newValue + "' 입력됨";
    } else if (editType === "삭제") {
      editInfo += "'" + oldValue + "' 삭제됨";
    } else {
      editInfo += "'" + oldValue + "'에서 '" + newValue + "'로 변경됨";
    }
    
    // T2 셀에 사용자 이메일 기록
    sheet.getRange("T2").setValue(user);
    
    // U2 셀에 변경 시간 기록
    sheet.getRange("U2").setValue(time);
    
    // V2 셀에 수정 정보 기록
    sheet.getRange("V2").setValue(editInfo);
  }
}
