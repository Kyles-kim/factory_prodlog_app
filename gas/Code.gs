// =====================================================
// (주)삼인 생산실적관리 앱 - Google Apps Script 최종
// 스프레드시트 ID: 1vl2AkjwgN3CxZb8_Hr_-n4qU1N2C1AgzjsM09ppPWD4
// 버전: v1.3.0 (2026-10-29) — 변경 내용은 저장소의 CHANGELOG.md 참고
// ★ 생산실적 시트 1행 헤더에 맞춰 동적 저장
// ★ 불량1/수량1/불량2/수량2 반환 포함
// ★ 저장일시/작업일자 Date 객체 → 한국시간 문자열 변환
// ★ 불량율 소수점 1자리 정규화
// ★ v1.3.0: 불량 사진·조치상태를 생산실적 시트 끝 열에 저장 (별도 불량 시트 사용 안 함)
// ★ v1.3.0: 사원 시트가 없으면 자동으로 만듦 (비밀번호는 시트에서만 관리)
// =====================================================

const GAS_VERSION    = 'v1.3.0';
const SPREADSHEET_ID = '1vl2AkjwgN3CxZb8_Hr_-n4qU1N2C1AgzjsM09ppPWD4';

// 생산실적 시트 끝에 자동으로 추가되는 열 (불량 사진·조치 관리용)
const RECORD_EXTRA_COLS = ['사진','조치상태','조치내용','조치자','조치일시'];
const PHOTO_FOLDER = '삼인_불량사진';

// ── JSONP 지원 응답 생성 ──────────────────────────────
function makeResponse(data, callback) {
  let output;
  if (callback) {
    output = ContentService.createTextOutput(callback + '(' + JSON.stringify(data) + ')');
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    output = ContentService.createTextOutput(JSON.stringify(data));
    output.setMimeType(ContentService.MimeType.JSON);
  }
  return output;
}

// ── GET 요청 라우터 ───────────────────────────────────
function doGet(e) {
  const action   = e.parameter.action;
  const callback = e.parameter.callback || null;
  try {
    if (action === 'login')      return makeResponse(login(e.parameter), callback);
    if (action === 'getMaster')  return makeResponse(getMasterData(), callback);
    if (action === 'getWorkers') return makeResponse(getWorkersData(), callback);
    if (action === 'getRecords') return makeResponse(getRecords(e.parameter), callback);
    if (action === 'getHeaders') return makeResponse(getSheetHeaders(e.parameter.sheet), callback);
    if (action === 'getVersion') return makeResponse({ success: true, version: GAS_VERSION }, callback);
    return makeResponse({ success: false, message: '알 수 없는 요청입니다.' }, callback);
  } catch (err) {
    return makeResponse({ success: false, message: err.message }, callback);
  }
}

// ── POST 요청 라우터 ──────────────────────────────────
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === 'saveRecord')   return makeResponse(saveRecord(data));
    if (action === 'deleteRecord') return makeResponse(deleteRecord(data));
    if (action === 'updateAction') return makeResponse(updateAction(data));
    return makeResponse({ success: false, message: '알 수 없는 요청입니다.' });
  } catch (err) {
    return makeResponse({ success: false, message: err.message });
  }
}

// =====================================================
// 1. 시트 1행 헤더 조회
// =====================================================
function getSheetHeaders(sheetName) {
  if (!sheetName) return { success: false, message: 'sheet 파라미터가 없습니다.' };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(sheetName);

  const defaultHeaders = [
    '저장일시','작업일자','작업조','작업구분',
    '고객사','모델명','부품명','원료명','색상',
    '설비번호','작업자',
    '양품수량','불량수량','불량1','수량1','불량2','수량2',
    '별도수량','총생산수량','양품율(%)','불량율(%)',
    '별도구분 ISSUE'
  ];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(defaultHeaders);
  } else if (sheet.getLastColumn() === 0) {
    sheet.appendRow(defaultHeaders);
  }

  // 생산실적 시트는 사진·조치 열이 없으면 끝에 추가
  if (sheetName === '생산실적') ensureRecordExtraCols(sheet);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(h => String(h).trim())
                       .filter(h => h !== '');
  return { success: true, headers: headers };
}

// =====================================================
// 2. 로그인 인증 (사원 시트: A=번호, B=사원명, C=비밀번호)
// =====================================================
function login(params) {
  const name = params.name;
  const pw   = params.pw;
  if (!name || !pw) return { success: false, message: '사원명과 비밀번호를 입력하세요.' };

  const staff = getStaffSheet();
  if (staff.created) return { success: false, message: '사원 시트를 새로 만들었습니다. 관리자가 비밀번호를 입력한 뒤 로그인하세요.' };

  const data = staff.sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowName = String(data[i][1]).trim();
    const rowPw   = String(data[i][2]).trim();
    if (rowPw && rowName === name.trim() && rowPw === pw.trim()) {
      return {
        success: true,
        worker: {
          name: rowName,
          dept: data[i][3] ? String(data[i][3]) : '',
          role: data[i][4] ? String(data[i][4]) : '',
        }
      };
    }
  }
  return { success: false, message: '사원명 또는 비밀번호가 일치하지 않습니다.' };
}

// =====================================================
// 3. MASTER 시트 → 제품 드롭다운 데이터
// =====================================================
function getMasterData() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('MASTER');
  if (!sheet) return { success: false, message: 'MASTER 시트를 찾을 수 없습니다.' };

  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    rows.push({
      client:   String(row[0]).trim(),
      model:    String(row[1]).trim(),
      part:     String(row[2]).trim(),
      material: String(row[3]).trim(),
      color:    String(row[4]).trim(),
    });
  }
  return { success: true, data: rows };
}

// =====================================================
// 4. 사원 시트 → 작업자 목록
// =====================================================
function getWorkersData() {
  const staff = getStaffSheet();
  const data = staff.sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][1]).trim();
    if (name) list.push(name);
  }
  return { success: true, data: list, created: staff.created };
}

// =====================================================
// 5. 생산실적 저장 (헤더 기반 동적 저장 + 불량 사진)
// =====================================================
function saveRecord(data) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName('생산실적');

  if (!sheet) {
    sheet = ss.insertSheet('생산실적');
    sheet.appendRow([
      '저장일시','작업일자','작업조','작업구분',
      '고객사','모델명','부품명','원료명','색상',
      '설비번호','작업자',
      '양품수량','불량수량','불량1','수량1','불량2','수량2',
      '별도수량','총생산수량','양품율(%)','불량율(%)',
      '별도구분 ISSUE'
    ]);
    ensureRecordExtraCols(sheet);
  }

  // rowData 배열 있으면 헤더 기반 동적 저장
  if (data.rowData && Array.isArray(data.rowData)) {
    const row = data.rowData.slice();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const pIdx = headers.indexOf('사진');
    if (pIdx >= 0) {
      // 수정 저장 시 기존 사진 유지 + 새 사진 추가
      const keep  = Array.isArray(data.keepPhotos) ? data.keepPhotos.map(photoLink) : [];
      const added = savePhotos(data.photos, data.photoName || '');
      const links = keep.concat(added);
      while (row.length <= pIdx) row.push('');
      if (links.length) row[pIdx] = links.join('\n');
    }
    sheet.appendRow(row);
    return { success: true, message: '저장되었습니다.' };
  }

  // fallback
  const now      = new Date();
  const total    = (data.qtyGood || 0) + (data.qtyBad || 0) + (data.qtyExtra || 0);
  const goodRate = total > 0 ? parseFloat((data.qtyGood / total * 100).toFixed(1)) : 0;
  const badRate  = total > 0 ? parseFloat((data.qtyBad  / total * 100).toFixed(1)) : 0;

  sheet.appendRow([
    Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    data.workDate  || '', data.shift    || '', data.workType || '',
    data.client    || '', data.model    || '', data.part     || '',
    data.material  || '', data.color    || '',
    data.equipment || '', data.worker   || '',
    data.qtyGood   || 0, data.qtyBad   || 0,
    data.defect1   || '', data.defectQty1 || 0,
    data.defect2   || '', data.defectQty2 || 0,
    data.qtyExtra  || 0, total, goodRate, badRate,
    data.issueNote || '',
  ]);
  return { success: true, message: '저장되었습니다.(fallback)' };
}

// =====================================================
// 6. 생산실적 조회
// =====================================================
function getRecords(params) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('생산실적');
  if (!sheet) return { success: true, data: [] };

  const raw = sheet.getDataRange().getValues();
  if (raw.length < 2) return { success: true, data: [] };

  const headers = raw[0].map(h => String(h).trim());
  const idx = (name) => headers.indexOf(name);

  const rows = [];
  for (let i = 1; i < raw.length; i++) {
    const r = raw[i];
    rows.push({
      rowIndex:   i + 1,
      savedAt:    r[idx('저장일시')] instanceof Date
                    ? Utilities.formatDate(r[idx('저장일시')], 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
                    : String(r[idx('저장일시')] || ''),
      workDate:   r[idx('작업일자')] instanceof Date
                    ? Utilities.formatDate(r[idx('작업일자')], 'Asia/Seoul', 'yyyy-MM-dd')
                    : String(r[idx('작업일자')] || ''),
      shift:      r[idx('작업조')]    || '',
      workType:   r[idx('작업구분')]  || '',
      client:     r[idx('고객사')]    || '',
      model:      r[idx('모델명')]    || '',
      part:       r[idx('부품명')]    || '',
      material:   r[idx('원료명')]    || '',
      color:      r[idx('색상')]      || '',
      equipment:  r[idx('설비번호')]  || '',
      worker:     r[idx('작업자')]    || '',
      qtyGood:    r[idx('양품수량')]  || 0,
      qtyBad:     r[idx('불량수량')]  || 0,
      defect1:    r[idx('불량1')]     || '',
      defectQty1: r[idx('수량1')]     || 0,
      defect2:    r[idx('불량2')]     || '',
      defectQty2: r[idx('수량2')]     || 0,
      qtyExtra:   r[idx('별도수량')]  || 0,
      qtyTotal:   r[idx('총생산수량')]|| 0,
      goodRate:   parseFloat((parseFloat(r[idx('양품율(%)')] || r[idx('양품율')] || 0)).toFixed(1)),
      badRate:    parseFloat((parseFloat(r[idx('불량율(%)')] || r[idx('불량율')] || 0)).toFixed(1)),
      issueNote:  r[idx('별도구분 ISSUE')] || r[idx('별도구분ISSUE')] ||
                  r[idx('주요 품질 ISSUE')] || r[idx('품질이슈')] || '',
      photos:     photoIds(idx('사진') >= 0 ? r[idx('사진')] : ''),
      status:     idx('조치상태') >= 0 ? String(r[idx('조치상태')] || '') : '',
      actionNote: idx('조치내용') >= 0 ? String(r[idx('조치내용')] || '') : '',
      actor:      idx('조치자')   >= 0 ? String(r[idx('조치자')]   || '') : '',
      actedAt:    idx('조치일시') >= 0 ? fmtDate(r[idx('조치일시')], 'yyyy-MM-dd HH:mm') : '',
    });
  }
  rows.reverse();
  return { success: true, data: rows };
}

// =====================================================
// 7. 생산실적 삭제
// =====================================================
function deleteRecord(data) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('생산실적');
  if (!sheet) return { success: false, message: '시트를 찾을 수 없습니다.' };

  const indices = [...data.rowIndices].sort((a, b) => b - a);
  for (const idx of indices) {
    sheet.deleteRow(idx);
  }
  return { success: true, message: `${indices.length}건 삭제 완료` };
}

// =====================================================
// 8. 불량 조치상태 저장 (생산실적 시트의 조치상태·조치내용·조치자·조치일시 열)
//    행 번호가 바뀌었을 수 있어 저장일시로 한 번 더 확인
// =====================================================
function updateAction(data) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('생산실적');
  if (!sheet) return { success: false, message: '생산실적 시트를 찾을 수 없습니다.' };
  ensureRecordExtraCols(sheet);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const col = name => headers.indexOf(name) + 1;
  const savedCol = col('저장일시');
  const savedAtOf = r => fmtDate(sheet.getRange(r, savedCol).getValue(), 'yyyy-MM-dd HH:mm:ss');

  let row = parseInt(data.rowIndex, 10);
  if (!(row >= 2 && row <= sheet.getLastRow() && savedAtOf(row) === String(data.savedAt))) {
    row = -1;
    const vals = sheet.getRange(2, savedCol, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (fmtDate(vals[i][0], 'yyyy-MM-dd HH:mm:ss') === String(data.savedAt)) { row = i + 2; break; }
    }
  }
  if (row < 0) return { success: false, message: '기록을 찾을 수 없습니다.' };

  sheet.getRange(row, col('조치상태')).setValue(data.status || '미조치');
  sheet.getRange(row, col('조치내용')).setValue(data.actionNote || '');
  sheet.getRange(row, col('조치자')).setValue(data.actor || '');
  sheet.getRange(row, col('조치일시')).setValue(Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'));
  return { success: true, message: '조치 내용이 저장되었습니다.' };
}

// =====================================================
// 9. 공통 도우미 (열 추가, 사진 저장, 날짜 변환)
// =====================================================
function ensureRecordExtraCols(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  let next = headers.filter(h => h !== '').length ? lastCol + 1 : 1;
  RECORD_EXTRA_COLS.forEach(name => {
    if (headers.indexOf(name) < 0) sheet.getRange(1, next++).setValue(name);
  });
}

function getPhotoFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 폴더가 지워졌으면 새로 만듦 */ }
  }
  const folder = DriveApp.createFolder(PHOTO_FOLDER);
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

// base64 사진(최대 3장)을 드라이브에 저장하고 링크 목록을 돌려줌
function savePhotos(photos, baseName) {
  if (!Array.isArray(photos) || !photos.length) return [];
  const folder = getPhotoFolder();
  const stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss');
  return photos.slice(0, 3).map((p, i) => {
    const bytes = Utilities.base64Decode(String(p.base64 || ''));
    const file  = folder.createFile(Utilities.newBlob(bytes, p.mimeType || 'image/jpeg', `${stamp}_${baseName}_${i + 1}.jpg`));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* 회사 정책으로 막힌 경우 무시 */ }
    return photoLink(file.getId());
  });
}

function photoLink(id) { return 'https://drive.google.com/file/d/' + id + '/view'; }

// 사진 칸의 드라이브 링크에서 파일 ID만 추출
function photoIds(cell) {
  return (String(cell || '').match(/[-\w]{25,}/g) || []);
}

function fmtDate(v, pattern) {
  return v instanceof Date ? Utilities.formatDate(v, 'Asia/Seoul', pattern) : String(v || '');
}

// =====================================================
// 10. 사원 시트 — 없으면 자동으로 만들고 사원명을 채움
//     비밀번호(C열)는 시트에서 직접 입력하세요.
//     ※ 비밀번호는 이 코드에 적지 마세요 (코드는 공개 GitHub에 올라갑니다).
// =====================================================
function getStaffSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('사원');
  if (sheet) return { sheet: sheet, created: false };

  const names = ['경희','지민','챠리','객','잔','디','맨','꼭','지원','은미','관곤'];
  sheet = ss.insertSheet('사원');
  sheet.appendRow(['번호','사원명','비밀번호','부서','역할']);
  names.forEach((n, i) => sheet.appendRow([i + 1, n, '', '', '']));
  sheet.getRange('C:C').setNumberFormat('@');   // 0으로 시작하는 비밀번호 유지
  sheet.setFrozenRows(1);
  return { sheet: sheet, created: true };
}

// 편집기에서 직접 실행해도 됨 (이미 있으면 그대로 둠)
function setupStaffSheet() {
  const res = getStaffSheet();
  Logger.log(res.created ? '사원 시트를 만들었습니다. C열에 비밀번호를 입력하세요.' : '사원 시트가 이미 있습니다.');
}
