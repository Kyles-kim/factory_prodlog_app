// =====================================================
// (주)삼인 생산실적관리 앱 - Google Apps Script 최종
// 스프레드시트 ID: 1vl2AkjwgN3CxZb8_Hr_-n4qU1N2C1AgzjsM09ppPWD4
// 버전: v1.1.0 (2026-10-29) — 변경 내용은 저장소의 CHANGELOG.md 참고
// ★ 생산실적 시트 1행 헤더에 맞춰 동적 저장
// ★ 불량1/수량1/불량2/수량2 반환 포함
// ★ 저장일시/작업일자 Date 객체 → 한국시간 문자열 변환
// ★ 불량율 소수점 1자리 정규화
// ★ v1.1.0: 불량기록 시트(사진·조치상태) 추가 — 8~11번 함수
// =====================================================

const GAS_VERSION    = 'v1.1.0';
const SPREADSHEET_ID = '1vl2AkjwgN3CxZb8_Hr_-n4qU1N2C1AgzjsM09ppPWD4';

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
    if (action === 'getDefects') return makeResponse(getDefects(), callback);
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
    if (action === 'saveDefect')   return makeResponse(saveDefect(data));
    if (action === 'updateDefect') return makeResponse(updateDefect(data));
    if (action === 'deleteDefect') return makeResponse(deleteDefect(data));
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
    return { success: true, headers: defaultHeaders, created: true };
  }

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.appendRow(defaultHeaders);
    return { success: true, headers: defaultHeaders, created: true };
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
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

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('사원');
  if (!sheet) return { success: false, message: '사원 시트를 찾을 수 없습니다.' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowName = String(data[i][1]).trim();
    const rowPw   = String(data[i][2]).trim();
    if (rowName === name.trim() && rowPw === pw.trim()) {
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
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('사원');
  if (!sheet) return { success: false, message: '사원 시트를 찾을 수 없습니다.' };

  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][1]).trim();
    if (name) list.push(name);
  }
  return { success: true, data: list };
}

// =====================================================
// 5. 생산실적 저장 (헤더 기반 동적 저장)
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
  }

  // rowData 배열 있으면 헤더 기반 동적 저장
  if (data.rowData && Array.isArray(data.rowData)) {
    sheet.appendRow(data.rowData);
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
// 8. 불량기록 시트 준비
//    (시트가 없으면 만들고, 사진은 드라이브 '삼인_불량사진' 폴더에 저장)
// =====================================================
const DEFECT_SHEET   = '불량기록';
const DEFECT_HEADERS = [
  'ID','저장일시','발생일자','작업조','설비번호',
  '고객사','모델명','부품명',
  '불량유형','불량수량','작업자','불량내용','사진',
  '조치상태','조치내용','조치자','조치일시'
];
const PHOTO_FOLDER = '삼인_불량사진';

function getDefectSheet() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(DEFECT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DEFECT_SHEET);
    sheet.appendRow(DEFECT_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
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

// 헤더명 → 열 번호(1부터) 맵
function defectColMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i + 1; });
  return map;
}

// ID로 행 번호 찾기 (없으면 -1)
function findDefectRow(sheet, id) {
  const col  = defectColMap(sheet)['ID'];
  const last = sheet.getLastRow();
  if (!col || last < 2) return -1;
  const ids = sheet.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function fmtDate(v, pattern) {
  return v instanceof Date ? Utilities.formatDate(v, 'Asia/Seoul', pattern) : String(v || '');
}

// 사진 칸의 드라이브 링크에서 파일 ID만 추출
function photoIds(cell) {
  return (String(cell || '').match(/[-\w]{25,}/g) || []);
}

// =====================================================
// 9. 불량기록 조회 (최신순)
// =====================================================
function getDefects() {
  const sheet = getDefectSheet();
  const raw = sheet.getDataRange().getValues();
  if (raw.length < 2) return { success: true, version: GAS_VERSION, data: [] };

  const headers = raw[0].map(h => String(h).trim());
  const idx = (name) => headers.indexOf(name);
  const val = (r, name) => idx(name) >= 0 ? r[idx(name)] : '';

  const rows = [];
  for (let i = 1; i < raw.length; i++) {
    const r = raw[i];
    if (!val(r, 'ID')) continue;
    rows.push({
      id:         String(val(r, 'ID')),
      savedAt:    fmtDate(val(r, '저장일시'), 'yyyy-MM-dd HH:mm:ss'),
      date:       fmtDate(val(r, '발생일자'), 'yyyy-MM-dd'),
      shift:      String(val(r, '작업조') || ''),
      equipment:  String(val(r, '설비번호') || ''),
      client:     String(val(r, '고객사') || ''),
      model:      String(val(r, '모델명') || ''),
      part:       String(val(r, '부품명') || ''),
      type:       String(val(r, '불량유형') || ''),
      qty:        parseInt(val(r, '불량수량')) || 0,
      worker:     String(val(r, '작업자') || ''),
      note:       String(val(r, '불량내용') || ''),
      photos:     photoIds(val(r, '사진')),
      status:     String(val(r, '조치상태') || '미조치'),
      actionNote: String(val(r, '조치내용') || ''),
      actor:      String(val(r, '조치자') || ''),
      actedAt:    fmtDate(val(r, '조치일시'), 'yyyy-MM-dd HH:mm'),
    });
  }
  rows.reverse();
  return { success: true, version: GAS_VERSION, data: rows };
}

// =====================================================
// 10. 불량기록 저장 (사진은 base64로 받아 드라이브에 저장)
//     앱이 만든 ID로 저장 → 같은 ID가 이미 있으면 다시 저장하지 않음
// =====================================================
function saveDefect(data) {
  const d = data.defect || {};
  if (!data.id) return { success: false, message: 'ID가 없습니다.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getDefectSheet();
    if (findDefectRow(sheet, data.id) > 0) return { success: true, message: '이미 저장된 기록입니다.' };

    // 사진 저장
    const links = [];
    const photos = Array.isArray(data.photos) ? data.photos.slice(0, 3) : [];
    if (photos.length) {
      const folder = getPhotoFolder();
      photos.forEach((p, i) => {
        const bytes = Utilities.base64Decode(String(p.base64 || ''));
        const name  = `${d.date || ''}_${d.equipment || ''}_${d.type || ''}_${data.id}_${i + 1}.jpg`;
        const file  = folder.createFile(Utilities.newBlob(bytes, p.mimeType || 'image/jpeg', name));
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* 회사 정책으로 막힌 경우 무시 */ }
        links.push('https://drive.google.com/file/d/' + file.getId() + '/view');
      });
    }

    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    const valueMap = {
      'ID': data.id, '저장일시': now, '발생일자': d.date || '', '작업조': d.shift || '',
      '설비번호': d.equipment || '', '고객사': d.client || '', '모델명': d.model || '', '부품명': d.part || '',
      '불량유형': d.type || '', '불량수량': parseInt(d.qty) || 0, '작업자': d.worker || '',
      '불량내용': d.note || '', '사진': links.join('\n'),
      '조치상태': '미조치', '조치내용': '', '조치자': '', '조치일시': '',
    };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    sheet.appendRow(headers.map(h => valueMap[h] !== undefined ? valueMap[h] : ''));
    return { success: true, message: '저장되었습니다.' };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================
// 11. 불량기록 조치상태 변경 / 삭제
// =====================================================
function updateDefect(data) {
  const sheet = getDefectSheet();
  const row = findDefectRow(sheet, data.id);
  if (row < 0) return { success: false, message: '기록을 찾을 수 없습니다.' };
  const col = defectColMap(sheet);
  const set = (name, v) => { if (col[name]) sheet.getRange(row, col[name]).setValue(v); };
  set('조치상태', data.status || '미조치');
  set('조치내용', data.actionNote || '');
  set('조치자',   data.actor || '');
  set('조치일시', Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'));
  return { success: true, message: '조치 내용이 저장되었습니다.' };
}

function deleteDefect(data) {
  const sheet = getDefectSheet();
  const row = findDefectRow(sheet, data.id);
  if (row < 0) return { success: false, message: '기록을 찾을 수 없습니다.' };
  // 사진은 드라이브 휴지통으로 이동 (30일 안에 복구 가능)
  const col = defectColMap(sheet);
  if (col['사진']) {
    photoIds(sheet.getRange(row, col['사진']).getValue()).forEach(id => {
      try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* 이미 없는 파일 */ }
    });
  }
  sheet.deleteRow(row);
  return { success: true, message: '삭제되었습니다.' };
}
