// ════════════════════════════════════════════════════════════════
// KY 스마트공장 재고·물류관리 시스템 — Google Apps Script (Code.gs) v6
// 변경사항: saveTransactionGet에 price 컬럼 추가
//           getTransactions 응답에 price 포함
//           입출고이력 시트 헤더: 일시/구분/품목코드/품목명/수량/가격/창고/현장명/담당자/비고
// ════════════════════════════════════════════════════════════════

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  var p = e ? e.parameter : {};

  var output;

  switch (action) {

    // ── 연결 테스트 ──────────────────────────────────────────────
    case 'test':
      output = { success: true, message: '연결 정상 (v6)', timestamp: new Date().toISOString() };
      break;

    // ── 대시보드 ─────────────────────────────────────────────────
    case 'dashboard':
      try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var txSheet = ss.getSheetByName('입출고이력');
        var totalIn = 0, totalOut = 0;
        if (txSheet) {
          var txData = txSheet.getDataRange().getValues();
          for (var i = 1; i < txData.length; i++) {
            if (txData[i][1] === '입고') totalIn++;
            else if (txData[i][1] === '출고') totalOut++;
          }
        }
        output = { success: true, totalIn: totalIn, totalOut: totalOut };
      } catch(err) { output = { success: false, message: err.message }; }
      break;

    // ── 입출고 저장 (v6: price 컬럼 추가) ───────────────────────
    // 시트 헤더: 일시(A) / 구분(B) / 품목코드(C) / 품목명(D) / 수량(E)
    //           가격(F) / 창고(G) / 현장명(H) / 담당자(I) / 비고(J)
    case 'saveTransactionGet':
      try {
        var ss2 = SpreadsheetApp.getActiveSpreadsheet();
        var sh2 = ss2.getSheetByName('입출고이력') || ss2.insertSheet('입출고이력');
        if (sh2.getLastRow() === 0) {
          sh2.appendRow(['일시','구분','품목코드','품목명','수량','가격','창고','현장명','담당자','비고']);
          // 헤더 서식
          var hdrRange = sh2.getRange(1, 1, 1, 10);
          hdrRange.setBackground('#1e3a5f');
          hdrRange.setFontColor('#ffffff');
          hdrRange.setFontWeight('bold');
          sh2.setFrozenRows(1);
        }
        sh2.appendRow([
          p.date     || new Date().toISOString(),
          p.type     || '',
          p.code     || '',
          p.name     || '',
          Number(p.qty   || 0),
          Number(p.price || 0),   // ← v6 신규: 입고가격 또는 출고가격
          p.warehouse || '',
          p.site      || '',
          p.manager   || '',
          p.note      || ''
        ]);
        output = { success: true, message: '저장 완료' };
      } catch(err) { output = { success: false, message: err.message }; }
      break;

    // ── 품목 저장 ────────────────────────────────────────────────
    case 'saveMasterItemGet':
      try {
        var ss3 = SpreadsheetApp.getActiveSpreadsheet();
        var sh3 = ss3.getSheetByName('품목마스터') || ss3.insertSheet('품목마스터');
        if (sh3.getLastRow() === 0) {
          sh3.appendRow(['품목코드','품목명','규격','단위','안전재고']);
          var hR = sh3.getRange(1, 1, 1, 5);
          hR.setBackground('#1e3a5f'); hR.setFontColor('#ffffff'); hR.setFontWeight('bold');
          sh3.setFrozenRows(1);
        }
        var code3 = p.code || '';
        var data3 = sh3.getDataRange().getValues();
        var found3 = -1;
        for (var i3 = 1; i3 < data3.length; i3++) {
          if (String(data3[i3][0]) === code3) { found3 = i3 + 1; break; }
        }
        var row3 = [code3, p.name||'', p.spec||'', p.unit||'EA', Number(p.safety||0)];
        if (found3 > 0) sh3.getRange(found3, 1, 1, 5).setValues([row3]);
        else sh3.appendRow(row3);
        output = { success: true, message: '품목 저장 완료' };
      } catch(err) { output = { success: false, message: err.message }; }
      break;

    // ── 품목 삭제 ────────────────────────────────────────────────
    case 'deleteMasterItem':
      try {
        var ss4 = SpreadsheetApp.getActiveSpreadsheet();
        var sh4 = ss4.getSheetByName('품목마스터');
        if (!sh4) { output = { success: false, message: '품목마스터 시트 없음' }; break; }
        var code4 = p.code || '';
        var data4 = sh4.getDataRange().getValues();
        var found4 = -1;
        for (var i4 = 1; i4 < data4.length; i4++) {
          if (String(data4[i4][0]) === code4) { found4 = i4 + 1; break; }
        }
        if (found4 > 0) sh4.deleteRow(found4);
        output = { success: true, message: found4 > 0 ? '삭제 완료' : '해당 코드 없음' };
      } catch(err) { output = { success: false, message: err.message }; }
      break;

    // ── 품목 목록 조회 ───────────────────────────────────────────
    case 'getMasterItems':
      try {
        var ss5 = SpreadsheetApp.getActiveSpreadsheet();
        var sh5 = ss5.getSheetByName('품목마스터');
        if (!sh5 || sh5.getLastRow() < 2) { output = { success: true, items: [] }; break; }
        var data5 = sh5.getDataRange().getValues();
        var items5 = [];
        for (var i5 = 1; i5 < data5.length; i5++) {
          var r5 = data5[i5];
          if (!r5[0]) continue;
          items5.push({
            code:   String(r5[0]),
            name:   String(r5[1] || ''),
            spec:   String(r5[2] || ''),
            unit:   String(r5[3] || 'EA'),
            safety: Number(r5[4] || 0)
          });
        }
        output = { success: true, items: items5 };
      } catch(err) { output = { success: false, message: err.message }; }
      break;

    // ── 이력 조회 (v6: price 컬럼 인덱스 변경) ──────────────────
    // 컬럼 순서: [0]일시 [1]구분 [2]품목코드 [3]품목명 [4]수량
    //            [5]가격 [6]창고 [7]현장명 [8]담당자 [9]비고
    case 'getTransactions':
      try {
        var ss6 = SpreadsheetApp.getActiveSpreadsheet();
        var sh6 = ss6.getSheetByName('입출고이력');
        if (!sh6 || sh6.getLastRow() < 2) { output = { success: true, transactions: [] }; break; }
        var data6 = sh6.getDataRange().getValues();
        var txs6 = [];
        for (var i6 = 1; i6 < data6.length; i6++) {
          var r6 = data6[i6];
          if (!r6[0] && !r6[2]) continue;
          txs6.push({
            date:      String(r6[0] || ''),
            type:      String(r6[1] || ''),
            code:      String(r6[2] || ''),
            name:      String(r6[3] || ''),
            qty:       Number(r6[4] || 0),
            price:     Number(r6[5] || 0),   // ← v6 신규
            warehouse: String(r6[6] || ''),
            site:      String(r6[7] || ''),
            manager:   String(r6[8] || ''),
            note:      String(r6[9] || '')
          });
        }
        output = { success: true, transactions: txs6 };
      } catch(err) { output = { success: false, message: err.message }; }
      break;

    // ── 기존 시트 마이그레이션 (v5→v6 헤더 변환용) ─────────────
    // action=migrateSheet 로 호출하면 기존 9컬럼 시트에 '가격' 컬럼 삽입
    case 'migrateSheet':
      try {
        var ss7 = SpreadsheetApp.getActiveSpreadsheet();
        var sh7 = ss7.getSheetByName('입출고이력');
        if (!sh7) { output = { success: false, message: '시트 없음' }; break; }
        var hdr7 = sh7.getRange(1, 1, 1, sh7.getLastColumn()).getValues()[0];
        // 6번째 컬럼(F)이 이미 '가격'이면 스킵
        if (String(hdr7[5]) === '가격') {
          output = { success: true, message: '이미 v6 형식입니다' }; break;
        }
        // F열에 '가격' 컬럼 삽입 (기존 데이터 우측 이동)
        sh7.insertColumnAfter(5);
        sh7.getRange(1, 6).setValue('가격');
        sh7.getRange(1, 6).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
        output = { success: true, message: '마이그레이션 완료 — F열에 가격 컬럼 삽입됨' };
      } catch(err) { output = { success: false, message: err.message }; }
      break;

    default:
      output = { success: false, message: '알 수 없는 action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}
