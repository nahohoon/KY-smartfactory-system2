// ════════════════════════════════════════════════════════════════
// KY 스마트공장 재고·물류관리 시스템 — Google Apps Script (Code.gs)
// 버전: 2026-04-19 수정본
// 수정 내용: getMasterItems / getTransactions 누락 대응 + CORS 헤더 보강
// ════════════════════════════════════════════════════════════════

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  var p = (e && e.parameter) ? e.parameter : {};

  var output;

  switch (action) {

    // ── 연결 테스트 ──────────────────────────────────────────────
    case 'test':
      output = { success: true, message: '연결 정상', timestamp: new Date().toISOString() };
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

    // ── 입출고 저장 ──────────────────────────────────────────────
    case 'saveTransactionGet':
      try {
        var ss2 = SpreadsheetApp.getActiveSpreadsheet();
        var sh2 = ss2.getSheetByName('입출고이력') || ss2.insertSheet('입출고이력');
        if (sh2.getLastRow() === 0) {
          sh2.appendRow(['일시','구분','품목코드','품목명','수량','창고','현장명','담당자','비고']);
        }
        sh2.appendRow([
          p.date  || new Date().toISOString(),
          p.type  || '',
          p.code  || '',
          p.name  || '',
          Number(p.qty || 0),
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
        }
        var code3 = p.code || '';
        var data3 = sh3.getDataRange().getValues();
        var found3 = -1;
        for (var i3 = 1; i3 < data3.length; i3++) {
          if (String(data3[i3][0]) === code3) { found3 = i3 + 1; break; }
        }
        var row3 = [code3, p.name || '', p.spec || '', p.unit || 'EA', Number(p.safety || 0)];
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
    // HTML이 요청하는 action: getMasterItems
    // 반환 형식: { success: true, items: [ {code, name, spec, unit, safety}, ... ] }
    case 'getMasterItems':
      try {
        var ss5 = SpreadsheetApp.getActiveSpreadsheet();
        var sh5 = ss5.getSheetByName('품목마스터');
        if (!sh5 || sh5.getLastRow() < 2) {
          output = { success: true, items: [] };
          break;
        }
        var data5 = sh5.getDataRange().getValues();
        var items5 = [];
        for (var i5 = 1; i5 < data5.length; i5++) {
          var r5 = data5[i5];
          if (!r5[0]) continue;
          items5.push({
            code:   String(r5[0] || ''),
            name:   String(r5[1] || ''),
            spec:   String(r5[2] || ''),
            unit:   String(r5[3] || 'EA'),
            safety: Number(r5[4] || 0)
          });
        }
        output = { success: true, items: items5 };
      } catch(err) { output = { success: false, message: err.message }; }
      break;

    // ── 입출고 이력 조회 ─────────────────────────────────────────
    // HTML이 요청하는 action: getTransactions
    // 반환 형식: { success: true, transactions: [ {date, type, code, name, qty, ...}, ... ] }
    case 'getTransactions':
      try {
        var ss6 = SpreadsheetApp.getActiveSpreadsheet();
        var sh6 = ss6.getSheetByName('입출고이력');
        if (!sh6 || sh6.getLastRow() < 2) {
          output = { success: true, transactions: [] };
          break;
        }
        var data6 = sh6.getDataRange().getValues();
        var txs6 = [];
        for (var i6 = 1; i6 < data6.length; i6++) {
          var r6 = data6[i6];
          if (!r6[0] && !r6[2]) continue;
          txs6.push({
            date: r6[1] ? Utilities.formatDate(new Date(r6[1]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : '',
            type: String(r6[4] || ''),
            code: String(r6[2] || ''),
            name: String(r6[3] || ''),
            qty: Math.abs(Number(r6[5] || 0)),
            warehouse: String(r6[6] || ''),
            site: String(r6[7] || ''),
            manager: String(r6[8] || ''),
            note: ''
          });
        }
        output = { success: true, transactions: txs6 };
      } catch(err) { output = { success: false, message: err.message }; }
      break;

    // ── 알 수 없는 action ────────────────────────────────────────
    default:
      output = { success: false, message: '알 수 없는 action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}
