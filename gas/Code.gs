// =====================================================
// (주)삼인 생산실적관리 앱 - Google Apps Script 최종
// 스프레드시트 ID: 1vl2AkjwgN3CxZb8_Hr_-n4qU1N2C1AgzjsM09ppPWD4
// 버전: v1.4.0 (2026-10-30) — 변경 내용은 저장소의 CHANGELOG.md 참고
// ★ 생산실적 시트 1행 헤더에 맞춰 동적 저장
// ★ 불량1/수량1/불량2/수량2 반환 포함
// ★ 저장일시/작업일자 Date 객체 → 한국시간 문자열 변환
// ★ 불량율 소수점 1자리 정규화
// ★ v1.4.0: 불량률 5% 이상 생산실적은 '불량기록' 시트에 자동 등록
// ★ v1.4.0: 모바일 불량 입력·조치·사진 → '불량기록' 시트
// ★ v1.4.0: 사원 시트가 없으면 자동으로 만듦 (비밀번호는 시트에서만 관리)
// =====================================================

const GAS_VERSION    = 'v1.4.0';
const SPREADSHEET_ID = '1vl2AkjwgN3CxZb8_Hr_-n4qU1N2C1AgzjsM09ppPWD4';

// 불량률이 이 값(%) 이상이면 생산실적 저장 시 불량기록에 자동 등록
const DEFECT_RATE_LIMIT = 5;

const DEFECT_SHEET = '불량기록';
const PHOTO_FOLDER = '삼인_불량사진';

// 불량기록 시트 열 — 1행 헤더 이름으로 찾음 (열 순서를 바꿔도 동작)
const DEFECT_COLS = {
  equipment: ['설비번호'],
  worker:    ['작업자'],
  date:      ['작업일자', '발생일자'],
  shift:     ['작업조(주.야간)', '작업조(주/야간)', '작업조'],
  client:    ['고객사'],
  model:     ['모델명'],
  part:      ['부품명'],
  type:      ['불량유형'],
  qty:       ['불량수량'],
  status:    ['조치상태'],
  action:    ['조치내용'],
  actedAt:   ['조치일시'],
  photo1:    ['사진1'],
  photo2:    ['사진2'],
  note:      ['불량내용'],
  source:    ['등록구분'],
  id:        ['ID'],
};
// 시트를 새로 만들 때의 헤더
const DEFECT_HEADERS = ['설비번호','작업자','작업일자','작업조(주.야간)','고객사','모델명','부품명',
  '불량유형','불량수량','조치상태','조치내용','조치일시','사진1','사진2','불량내용','등록구분','ID'];
// 기존 시트에 없으면 맨 끝에 추가하는 열
const DEFECT_EXTRA = ['불량내용','등록구분','ID'];

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
    if (action === 'getVersion') return makeResponse({ success: true, version: GAS_VERSION, rateLimit: DEFECT_RATE_LIMIT }, callback);
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

  const staff = getStaffSheet();
  if (staff.created) return { success: false, message: '사원 시트를 새로 만들었습니다. 관리자가 비밀번호를 입력한 뒤 로그인하세요.' };

  const data = staff.sheet.getDataRange().getValues();
  const c = staffCols(data[0]);
  for (let i = 1; i < data.length; i++) {
    const rowName = String(data[i][c.name]).trim();
    const rowPw   = String(data[i][c.pw]).trim();
    if (rowPw && rowName === name.trim() && rowPw === pw.trim()) {
      return {
        success: true,
        worker: {
          name: rowName,
          dept: c.dept >= 0 && data[i][c.dept] ? String(data[i][c.dept]) : '',
          role: c.role >= 0 && data[i][c.role] ? String(data[i][c.role]) : '',
        }
      };
    }
  }
  return { success: false, message: '사원명 또는 비밀번호가 일치하지 않습니다.' };
}

// 사원 시트 열 위치: 1행 제목으로 찾고, 없으면 B열=사원명, C열=비밀번호
function staffCols(header) {
  const h = (header || []).map(v => String(v).replace(/\s/g, ''));
  const find = (names, def) => { const i = h.findIndex(v => names.indexOf(v) >= 0); return i >= 0 ? i : def; };
  return {
    name: find(['사원명', '이름', '성명', '작업자'], 1),
    pw:   find(['비밀번호', '비번', 'PW', 'pw', '암호'], 2),
    dept: find(['부서'], -1),
    role: find(['역할', '직책'], -1),
  };
}

// =====================================================
// 3. MASTER 시트 → 제품 드롭다운 데이터
// =====================================================
function getMasterData() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('MASTER');
  if (!sheet) return { success: false, message: 'MASTER 시트를 찾을 수 없습니다.' };

  const data = sheet.getDataRange().getValues();
  // 1행 제목으로 열을 찾음 (앞에 '번호' 같은 열이 끼어도 동작). 제목이 없으면 A~E열 순서
  const h = data[0].map(v => String(v).replace(/\s/g, ''));
  const find = (names, def) => { const i = h.findIndex(v => names.indexOf(v) >= 0); return i >= 0 ? i : def; };
  const c = {
    client:   find(['고객사', '고객'], 0),
    model:    find(['모델명', '모델'], 1),
    part:     find(['부품명', '부품', '품명'], 2),
    material: find(['원료명', '원료', '재질'], 3),
    color:    find(['색상', '컬러'], 4),
  };
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[c.client]) continue;
    rows.push({
      client:   String(row[c.client]).trim(),
      model:    String(row[c.model]).trim(),
      part:     String(row[c.part]).trim(),
      material: String(row[c.material]).trim(),
      color:    String(row[c.color]).trim(),
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
  const c = staffCols(data[0]);
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][c.name]).trim();
    if (name) list.push(name);
  }
  return { success: true, data: list, created: staff.created };
}

// =====================================================
// 5. 생산실적 저장 (헤더 기반 동적 저장)
//    불량률 5% 이상이면 불량기록 시트에도 자동 등록
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
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const rec = {};
    headers.forEach((h, i) => { rec[h] = data.rowData[i]; });
    const auto = syncAutoDefects(rec, data.prevSavedAt || '');
    return { success: true, message: '저장되었습니다.', autoDefects: auto };
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
// 8. 불량기록 시트 — 준비·열 찾기·행 쓰기
// =====================================================
function getDefectSheet() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(DEFECT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DEFECT_SHEET);
    sheet.appendRow(DEFECT_HEADERS);
    sheet.setFrozenRows(1);
  }
  // 앱에 필요한 열(불량내용·등록구분·ID)이 없으면 맨 끝에 추가
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim());
  let next = sheet.getLastColumn() + 1;
  DEFECT_EXTRA.forEach(name => { if (headers.indexOf(name) < 0) sheet.getRange(1, next++).setValue(name); });
  return sheet;
}

// 필드 → 열 번호(1부터). 없는 열은 0
function defectColMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const map = {};
  Object.keys(DEFECT_COLS).forEach(key => {
    const i = DEFECT_COLS[key].map(n => headers.indexOf(n)).find(i => i >= 0);
    map[key] = i === undefined ? 0 : i + 1;
  });
  map.width = headers.length;
  return map;
}

// 날짜 문자열(yyyy-MM-dd) → Date (시트에 yy-MM-dd로 표시)
function toSheetDate(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : (s || '');
}

// 한 행 쓰기 (row 없으면 새 행 추가). fields: { equipment, worker, ... }
function writeDefectRow(sheet, col, fields, row) {
  if (!row) {
    sheet.appendRow(new Array(col.width).fill(''));
    row = sheet.getLastRow();
  }
  Object.keys(fields).forEach(key => {
    const c = col[key];
    if (!c || fields[key] === undefined) return;
    const cell = sheet.getRange(row, c);
    const v = fields[key];
    if (key === 'date' || key === 'actedAt') {
      cell.setValue(toSheetDate(v));
      cell.setNumberFormat('yy-MM-dd');
    } else if ((key === 'photo1' || key === 'photo2') && v) {
      cell.setFormula('=HYPERLINK("' + photoLink(v) + '","사진 보기")');
    } else {
      cell.setValue(v);
    }
  });
  return row;
}

// ID로 행 찾기 (없으면 -1)
function findDefectRow(sheet, col, id) {
  const last = sheet.getLastRow();
  if (!col.id || last < 2) return -1;
  const ids = sheet.getRange(2, col.id, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function newDefectId() {
  return 'D' + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase();
}

// =====================================================
// 9. 불량률 5% 이상 → 불량기록 자동 등록
//    불량1·불량2를 유형별 한 줄씩. ID = 'P' + 생산실적 저장일시(숫자) + '-' + 순번
//    수정 저장이면(prevSavedAt) 기존 자동 등록 줄을 찾아 내용만 고침 (조치·사진은 유지)
// =====================================================
function syncAutoDefects(rec, prevSavedAt) {
  const bad   = parseInt(rec['불량수량']) || 0;
  const total = parseInt(rec['총생산수량']) ||
                ((parseInt(rec['양품수량']) || 0) + bad + (parseInt(rec['별도수량']) || 0));
  const rate  = total > 0 ? bad / total * 100 : 0;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getDefectSheet();
    const col   = defectColMap(sheet);
    const key   = s => 'P' + String(s || '').replace(/\D/g, '');
    const newKey = key(rec['저장일시']);

    // 수정 저장: 이전 자동 등록 줄 찾기
    const prevRows = [];
    if (prevSavedAt && col.id && sheet.getLastRow() > 1) {
      const ids = sheet.getRange(2, col.id, sheet.getLastRow() - 1, 1).getValues();
      ids.forEach((v, i) => { if (String(v[0]).indexOf(key(prevSavedAt) + '-') === 0) prevRows.push(i + 2); });
    }
    if (rate < DEFECT_RATE_LIMIT) {
      // 수정해서 5% 미만이 되면: 조치·사진이 없는 자동 등록 줄은 삭제, 손댄 줄은 남기고 ID만 새 실적에 연결
      for (let i = prevRows.length - 1; i >= 0; i--) {
        const row = prevRows[i];
        const val = key => col[key] ? sheet.getRange(row, col[key]).getValue() : '';
        const form = key => col[key] ? sheet.getRange(row, col[key]).getFormula() : '';
        const touched = (val('status') && val('status') !== '미조치') || val('action') ||
                        photoIds(form('photo1') || val('photo1')).length || photoIds(form('photo2') || val('photo2')).length;
        if (touched) sheet.getRange(row, col.id).setValue(newKey + '-' + (i + 1));
        else sheet.deleteRow(row);
      }
      return { count: 0, rate: Math.round(rate * 10) / 10 };
    }

    const types = [];
    if (rec['불량1'] || parseInt(rec['수량1'])) types.push({ type: rec['불량1'] || '(유형 미입력)', qty: parseInt(rec['수량1']) || 0 });
    if (rec['불량2'] || parseInt(rec['수량2'])) types.push({ type: rec['불량2'] || '(유형 미입력)', qty: parseInt(rec['수량2']) || 0 });
    if (!types.length) types.push({ type: '(유형 미입력)', qty: bad });

    const rateText = Math.round(rate * 10) / 10;
    types.forEach((t, i) => {
      const fields = {
        equipment: rec['설비번호'] || '', worker: rec['작업자'] || '', date: rec['작업일자'] || '',
        shift: rec['작업조'] || '', client: rec['고객사'] || '', model: rec['모델명'] || '', part: rec['부품명'] || '',
        type: t.type, qty: t.qty, source: `생산실적 자동 (불량률 ${rateText}%)`, id: newKey + '-' + (i + 1),
      };
      if (prevRows[i]) {
        writeDefectRow(sheet, col, fields, prevRows[i]);
      } else {
        fields.status = '미조치';
        writeDefectRow(sheet, col, fields, 0);
      }
    });
    return { count: types.length, rate: rateText };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================
// 10. 불량기록 조회 (최신순). ID가 없는 줄(시트에 직접 적은 줄)은 ID를 채움
// =====================================================
function getDefects() {
  const sheet = getDefectSheet();
  const col = defectColMap(sheet);
  const last = sheet.getLastRow();
  if (last < 2) return { success: true, version: GAS_VERSION, data: [] };

  const range = sheet.getRange(2, 1, last - 1, col.width);
  const vals = range.getValues();
  const forms = range.getFormulas();
  const get = (r, key) => col[key] ? vals[r][col[key] - 1] : '';
  const photo = (r, key) => col[key] ? (photoIds(forms[r][col[key] - 1] || vals[r][col[key] - 1])[0] || '') : '';

  const rows = [];
  for (let r = 0; r < vals.length; r++) {
    if (!get(r, 'equipment') && !get(r, 'type') && !get(r, 'qty')) continue;   // 빈 줄
    let id = String(get(r, 'id') || '');
    if (!id && col.id) {
      id = 'S' + (r + 2) + '-' + newDefectId();
      sheet.getRange(r + 2, col.id).setValue(id);
    }
    rows.push({
      id,
      equipment: String(get(r, 'equipment') || ''),
      worker:    String(get(r, 'worker') || ''),
      date:      fmtDate(get(r, 'date'), 'yyyy-MM-dd'),
      shift:     String(get(r, 'shift') || ''),
      client:    String(get(r, 'client') || ''),
      model:     String(get(r, 'model') || ''),
      part:      String(get(r, 'part') || ''),
      type:      String(get(r, 'type') || ''),
      qty:       parseInt(get(r, 'qty')) || 0,
      status:    String(get(r, 'status') || '미조치'),
      action:    String(get(r, 'action') || ''),
      actedAt:   fmtDate(get(r, 'actedAt'), 'yyyy-MM-dd'),
      photos:    [photo(r, 'photo1'), photo(r, 'photo2')],
      note:      String(get(r, 'note') || ''),
      source:    String(get(r, 'source') || ''),
    });
  }
  rows.reverse();
  return { success: true, version: GAS_VERSION, data: rows };
}

// =====================================================
// 11. 모바일 불량 입력 → 불량기록 (사진 최대 2장)
//     앱이 만든 ID로 저장 → 같은 ID가 이미 있으면 다시 저장하지 않음
// =====================================================
function saveDefect(data) {
  const d = data.defect || {};
  if (!data.id) return { success: false, message: 'ID가 없습니다.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getDefectSheet();
    const col = defectColMap(sheet);
    if (findDefectRow(sheet, col, data.id) > 0) return { success: true, message: '이미 저장된 기록입니다.' };
    const photos = savePhotos(data.photos, `${d.equipment || ''}_${d.type || ''}`);
    writeDefectRow(sheet, col, {
      equipment: d.equipment || '', worker: d.worker || '', date: d.date || '', shift: d.shift || '',
      client: d.client || '', model: d.model || '', part: d.part || '',
      type: d.type || '', qty: parseInt(d.qty) || 0, note: d.note || '',
      status: d.status || '미조치', action: d.action || '', actedAt: d.action ? today_() : '',
      photo1: photos[0] || '', photo2: photos[1] || '',
      source: '모바일 입력', id: data.id,
    }, 0);
    return { success: true, message: '저장되었습니다.' };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================
// 12. 조치상태·조치내용 저장, 빈 사진 칸에 사진 추가
// =====================================================
function updateDefect(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getDefectSheet();
    const col = defectColMap(sheet);
    const row = findDefectRow(sheet, col, data.id);
    if (row < 0) return { success: false, message: '기록을 찾을 수 없습니다.' };

    const fields = { status: data.status || '미조치', action: data.action || '', actedAt: today_() };
    const newPhotos = savePhotos(data.photos, String(data.id));
    if (newPhotos.length) {
      const has = key => col[key] && photoIds(sheet.getRange(row, col[key]).getFormula() || sheet.getRange(row, col[key]).getValue()).length;
      ['photo1', 'photo2'].forEach(key => { if (!has(key) && newPhotos.length) fields[key] = newPhotos.shift(); });
    }
    writeDefectRow(sheet, col, fields, row);
    return { success: true, message: '조치 내용이 저장되었습니다.' };
  } finally {
    lock.releaseLock();
  }
}

function deleteDefect(data) {
  const sheet = getDefectSheet();
  const col = defectColMap(sheet);
  const row = findDefectRow(sheet, col, data.id);
  if (row < 0) return { success: false, message: '기록을 찾을 수 없습니다.' };
  // 사진은 드라이브 휴지통으로 이동 (30일 안에 복구 가능)
  ['photo1', 'photo2'].forEach(key => {
    if (!col[key]) return;
    const cell = sheet.getRange(row, col[key]);
    photoIds(cell.getFormula() || cell.getValue()).forEach(id => {
      try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* 이미 없는 파일 */ }
    });
  });
  sheet.deleteRow(row);
  return { success: true, message: '삭제되었습니다.' };
}

// =====================================================
// 13. 공통 도우미 (사진 저장, 날짜 변환)
// =====================================================
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

// base64 사진(최대 2장)을 드라이브에 저장하고 파일 ID 목록을 돌려줌
function savePhotos(photos, baseName) {
  if (!Array.isArray(photos) || !photos.length) return [];
  const folder = getPhotoFolder();
  const stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss');
  return photos.slice(0, 2).map((p, i) => {
    const bytes = Utilities.base64Decode(String(p.base64 || ''));
    const file  = folder.createFile(Utilities.newBlob(bytes, p.mimeType || 'image/jpeg', `${stamp}_${baseName}_${i + 1}.jpg`));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* 회사 정책으로 막힌 경우 무시 */ }
    return file.getId();
  });
}

function photoLink(id) { return 'https://drive.google.com/file/d/' + id + '/view'; }

// 사진 칸(링크 또는 HYPERLINK 수식)에서 드라이브 파일 ID만 추출
function photoIds(cell) {
  return (String(cell || '').match(/[-\w]{25,}/g) || []);
}

function fmtDate(v, pattern) {
  return v instanceof Date ? Utilities.formatDate(v, 'Asia/Seoul', pattern) : String(v || '');
}

function today_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'); }

// =====================================================
// 14. 사원 시트 — 없으면 자동으로 만들고 사원명을 채움
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
