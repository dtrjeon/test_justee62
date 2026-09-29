// ============================================================
//  code.gs  —  62기 수지주일반 전도현황 GAS 서버
//  대응 클라이언트: index_62th_JEE.html (보고서 입력), justee62_admin.html (관리자)
// ============================================================

// ── 대상 스프레드시트 ID ────────────────────────────────────
//  62기 수지주일반 전도현황 시트
var SHEET_ID = '1ERh4TN23-NTKyLwuw3dtHhSkVzCzqvg1plh6AMbGK1o';

// ── 시트 이름 상수 ────────────────────────────────────────────
var SHEET_REPORT = '전도보고서';   // 보고서 데이터 시트
var SHEET_JO     = '조정보';       // 조 구성원 시트
var SHEET_GROUP  = '그룹설정';      // 그룹 편성 설정 시트 (관리자+그룹장 이메일 통합 관리)
// SHEET_EMAIL('이메일') 은 SHEET_GROUP으로 통합됨 — 기존 시트는 읽기 폴백으로만 사용

// ── 관리자 비밀번호 ─────────────────────────────────────────
// admin(7463) = 전체 권한(문서 수정/삭제, 문서번호 관리, 불일치 수정 등)
// monitor(9191) = 열람 전용 — 기록 조회(전체 목록 보기)는 허용하되
//                 삭제/수정/그룹장(관리그룹) 신분 사용 등 데이터 변경은 차단
var ADMIN_PW   = '7463';
// 7463과 동일한 admin 권한을 갖는 추가 비밀번호들 (예: 다른 관리자 인원용)
var ADMIN_PW_EXTRA = ['7963', '7834'];
var MONITOR_PW = '9191';

// ── 세션 토큰 캐시 키 접두사 / 유효시간 ─────────────────────
// CacheService는 항목당 최대 21600초(6시간)까지만 보관 가능 — GAS 자체 제약.
// 6시간이 지나면 토큰이 자동 만료되어 재인증이 필요해진다(보안상 정상 동작).
var AUTH_TOKEN_PREFIX = 'authtok_';
var AUTH_TOKEN_TTL    = 21600; // 6시간

// ── 대상 스프레드시트 열기 (ID로 명시적 지정) ─────────────────
function getSS() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// ── 응답 헬퍼 ────────────────────────────────────────────────
function ok(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(Object.assign({ status: 'ok' }, obj || {})))
    .setMimeType(ContentService.MimeType.JSON);
}

function err(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 현재 시각 (한국 시간) ─────────────────────────────────────
function nowKST() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

// ============================================================
//  서버 측 권한 토큰 검증
//  ── 예전엔 클라이언트(localStorage의 index_admin_verified 플래그)만 보고
//     삭제/수정 버튼을 숨겼을 뿐, 서버는 요청이 실제로 admin에게서 왔는지
//     전혀 확인하지 않았다 — GAS 웹앱 URL만 알면 누구든 앱 화면을 거치지
//     않고 직접 delete_doc 등을 호출해 데이터를 지울 수 있었다는 뜻.
//  ── 이제 verify_admin이 성공하면 임시 세션 토큰을 발급하고(CacheService,
//     최대 6시간), admin 전용 쓰기 액션은 반드시 이 토큰을 함께 보내야
//     서버가 실제로 admin 레벨인지 재검증한 뒤 실행한다.
// ============================================================
function issueAuthToken(level) {
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put(AUTH_TOKEN_PREFIX + token, level, AUTH_TOKEN_TTL);
  return token;
}

function getTokenLevel(token) {
  if (!token) return null;
  return CacheService.getScriptCache().get(AUTH_TOKEN_PREFIX + token); // 'admin' | 'monitor' | null
}

// admin 전용 쓰기 액션 진입 전 공통 체크 — 통과 시 null, 실패 시 에러 응답 반환
function requireAdminToken(body) {
  var level = getTokenLevel(body && body.adminToken);
  if (level !== 'admin') {
    return err('🔒 admin 권한이 필요합니다. 다시 인증해주세요. (세션이 만료되었을 수 있습니다)');
  }
  return null;
}

// ── 시트 셀 값을 doGet 응답용으로 포맷 ─────────────────────
// "저장일시" 컬럼은 구글시트가 문자열을 자동으로 Date로 인식해버리는 경우가 있어
// (예: "2026-09-05 14:32:10" 저장 → 셀 값이 실제 Date 객체가 됨) 날짜만 남기면
// 등록 "시각" 정보가 통째로 사라진다. 그래서 이 컬럼만 시:분:초까지 포맷하고,
// 그 외 날짜 컬럼(전도일 등)은 기존대로 날짜만 남긴다.
function formatCellValue(v, headerName) {
  if (!(v instanceof Date)) return v;
  if (headerName === '저장일시') {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  }
  return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
}

// ── 헤더 인덱스 맵 빌드 ─────────────────────────────────────
function buildHeaderMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  headers.forEach(function(h, i) { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

// ── 이력 컬럼 확인 / 자동 생성 ───────────────────────────────
function ensureHistCol(sheet, colName) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headers.indexOf(colName);
  if (idx > -1) return idx + 1;
  var newCol = lastCol + 1;
  sheet.getRange(1, newCol).setValue(colName);
  return newCol;
}

// ── 이력 누적 기록 ────────────────────────────────────────────
function appendHist(sheet, row, histCol, entry) {
  var prev = sheet.getRange(row, histCol).getValue() || '';
  var next = prev ? entry + '\n' + prev : entry;
  sheet.getRange(row, histCol).setValue(next);
}

// ============================================================
//  doGet — 데이터 조회
// ============================================================
function doGet(e) {
  e = e || {};
  var action = (e.parameter && e.parameter.action) || 'get';

  try {
    // ── 전도보고서 전체 조회 ──────────────────────────────────
    if (action === 'get') {
      var sheet = getSS().getSheetByName(SHEET_REPORT);
      if (!sheet) return err('시트 없음: ' + SHEET_REPORT);
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) return ok({ records: [] });

      var headers = data[0].map(function(h) { return String(h).trim(); });
      var records = [];
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (!row[0] && !row[headers.indexOf('문서번호')]) continue;
        var obj = { rowIndex: i + 1 };
        headers.forEach(function(h, j) {
          var v = formatCellValue(row[j], h);
          obj[h] = (v === null || v === undefined) ? '' : v;
        });
        records.push(obj);
      }
      return ok({ records: records, lastRow: sheet.getLastRow(), serverTime: nowKST() });
    }

    // ── 특정 문서번호 단건 조회 ──────────────────────────────
    if (action === 'get_by_doc') {
      var docNumber = e.parameter.docNumber || '';
      if (!docNumber) return err('docNumber 파라미터 필요');
      var sheet = getSS().getSheetByName(SHEET_REPORT);
      if (!sheet) return err('시트 없음: ' + SHEET_REPORT);
      var data = sheet.getDataRange().getValues();
      var headers = data[0].map(function(h) { return String(h).trim(); });
      var docCol = headers.indexOf('문서번호');
      if (docCol < 0) return err('문서번호 컬럼 없음');

      var latest = null;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][docCol]).trim() === docNumber) {
          var obj = { rowIndex: i + 1 };
          headers.forEach(function(h, j) {
            obj[h] = formatCellValue(data[i][j], h);
          });
          if (!latest || String(obj['저장일시']) > String(latest['저장일시'])) {
            latest = obj;
          }
        }
      }
      if (!latest) return err('레코드를 찾을 수 없습니다: ' + docNumber);
      return ok({ record: latest });
    }

    // ── 조정보 조회 ──────────────────────────────────────────
    if (action === 'get_jo') {
      var sheet = getSS().getSheetByName(SHEET_JO);
      if (!sheet) return ok({ joData: [] });
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) return ok({ joData: [] });
      var headers = data[0].map(function(h) { return String(h).trim(); });
      var joData = [];
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (!row[0]) continue;
        var obj = {};
        headers.forEach(function(h, j) { obj[h] = row[j] || ''; });
        joData.push({
          jo:          obj['조'] || '',
          leader:      obj['조장'] || obj['훈련자(조장)'] || obj['리더'] || '',
          trainer:     obj['훈련자'] || '',
          trainee:     obj['훈련생'] || obj['훈련생1'] || '',
          trainee2:    obj['훈련생2'] || '',
          leaderStep:  obj['조장단계'] || obj['훈련자(조장)단계'] || obj['리더단계'] || '',
          trainerStep: obj['훈련자단계'] || '',
          traineeStep: obj['훈련생단계'] || '',
          group:       obj['그룹'] || '',
          ban:         obj['반'] || '',
          kisu:        obj['기수'] || ''
        });
      }
      return ok({ joData: joData });
    }

    // ── 그룹 편성 설정 조회 (+ 하위호환: get_email_sheet → get_group_settings) ──
    if (action === 'get_group_settings' || action === 'get_email_sheet') {
      return getGroupSettings();
    }

    // ── 버전 조회 ─────────────────────────────────────────────
    if (action === 'get_version') {
      return ok({ version: 'gas v2026.09.14', gasVersion: 'gas v2026.09.14', updated: nowKST() });
    }

    return err('알 수 없는 action: ' + action);

  } catch(e) {
    return err(e.message);
  }
}

// ============================================================
//  doPost — 데이터 저장 / 수정
// ============================================================
function doPost(e) {
  e = e || {};
  var body;
  try {
    body = JSON.parse((e.postData && e.postData.contents) || '{}');
  } catch(ex) {
    return err('JSON 파싱 실패: ' + ex.message);
  }

  var action = body.action || '';

  try {
    // ── 관리자 비밀번호 검증 ──────────────────────────────────
    // 성공 시 세션 토큰을 함께 발급한다 — 클라이언트는 이후 admin 전용
    // 쓰기 요청(삭제/수정/조정보 저장 등)에 이 토큰을 동봉해야 한다.
    if (action === 'verify_admin') {
      var pw = String(body.password || '');
      if (pw === ADMIN_PW || ADMIN_PW_EXTRA.indexOf(pw) > -1) {
        return ok({ verified: true, level: 'admin', token: issueAuthToken('admin') });
      } else if (pw === MONITOR_PW) {
        return ok({ verified: true, level: 'monitor', token: issueAuthToken('monitor') });
      } else {
        return err('비밀번호가 올바르지 않습니다');
      }
    }

    // ── 보고서 제출 ──────────────────────────────────────────
    // (일반 조원 누구나 제출 가능한 기능이라 토큰 검증 대상이 아님)
    if (!action || action === 'submit' || action === 'save_report') {
      return saveReport(body);
    }

    // ── 인적정보 사진 인식(Claude 비전 OCR) ──────────────────
    // (일반 조원 누구나 사용 가능한 기능이라 토큰 검증 대상이 아님 —
    //  보고서 폼을 채워주는 보조 기능일 뿐, 데이터를 직접 변경하지 않음)
    if (action === 'ocrFieldInfo') {
      return ocrFieldInfo(body.imageDataUrl);
    }

    // ── 조정보 저장 (admin 전용) ─────────────────────────────
    if (action === 'save_jo') {
      var authErr = requireAdminToken(body);
      if (authErr) return authErr;
      return saveJo(body.joData || []);
    }

    // ── 그룹 편성 설정 저장 (admin 전용) (+ 하위호환: save_email_sheet) ──
    if (action === 'save_group_settings' || action === 'save_email_sheet') {
      var authErr = requireAdminToken(body);
      if (authErr) return authErr;
      return saveGroupSettings(body);
    }

    // ── 전도유형 수정 + 이력 기록 (admin 전용) ──────────────
    if (action === 'fix_evang_type') {
      var authErr = requireAdminToken(body);
      if (authErr) return authErr;
      return fixEvangType(body);
    }

    // ── 문서번호 재넘버링 + 이력 기록 (admin 전용) ──────────
    if (action === 'renumber_doc') {
      var authErr = requireAdminToken(body);
      if (authErr) return authErr;
      return renumberDoc(body);
    }

    // ── 문서 삭제 (문서번호 전체) (admin 전용) ───────────────
    if (action === 'delete_doc') {
      var authErr = requireAdminToken(body);
      if (authErr) return authErr;
      return deleteDoc(body);
    }

    // ── 특정 행만 삭제 (admin 전용) ──────────────────────────
    if (action === 'delete_doc_row') {
      var authErr = requireAdminToken(body);
      if (authErr) return authErr;
      return deleteDocRow(body);
    }

    return err('알 수 없는 action: ' + action);

  } catch(e) {
    return err(e.message);
  }
}

// ============================================================
//  saveReport — 보고서 행 추가
// ============================================================
function saveReport(data) {
  data = data || {};
  var ss    = getSS();
  var sheet = ss.getSheetByName(SHEET_REPORT);
  if (!sheet) return err('시트 없음: ' + SHEET_REPORT);

  var BASE_COLS = [
    '저장일시','주차','전도일','문서번호','전도유형','횟수','조',
    '훈련자(조장)','훈련자','훈련생',
    '접촉횟수','복음제시','대상자','결신','보류/거절','확신','이미신자',
    '거주지역','성씨','나이','결혼여부','성별','만난장소','복음전달',
    '대상자구분','관계조원','종교배경','직업',
    '천국확신','이유대답','영접여부','진단결과',
    '간증6','간증7',
    '훈련자(조장)역할','훈련자(조장)행동',
    '훈련자역할','훈련자행동',
    '훈련생역할','훈련생행동',
    '결과','즉석양육','후속양육','하지못한사유',
    '느낀점(10번)','문제점(11번)','기도제목(12번)','기타(13번)','풀복음 완료일',
    '특이사항(공유)'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(BASE_COLS);
  } else {
    // 기존 시트에 없는 컬럼(예: 특이사항(공유))이 있으면 맨 끝에 자동 추가
    var lastCol0 = sheet.getLastColumn();
    var headers0 = lastCol0 > 0
      ? sheet.getRange(1, 1, 1, lastCol0).getValues()[0].map(function(h) { return String(h).trim(); })
      : [];
    BASE_COLS.forEach(function(col) {
      if (headers0.indexOf(col) === -1) {
        var newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(col);
        headers0.push(col);
      }
    });
  }

  var now = nowKST();

  // ── 헤더 "이름"을 기준으로 값을 매칭한다 ──────────────────────
  // (예전엔 appendRow(고정순서배열)를 썼는데, 문서번호_수정이력 같은
  //  컬럼이 중간에 자동으로 끼어들면 뒤 컬럼들이 전부 한 칸씩 밀려서
  //  엉뚱한 칸에 저장되는 문제가 있었음 — 반드시 헤더명으로 찾아서 넣는다)
  var FIELD_MAP = {
    '저장일시': now,
    '주차': data.week || '',
    '전도일': data.date || '',
    '문서번호': data.doc_number || '',
    '전도유형': data.evang_type || '',
    '횟수': data.evang_count || '',
    '조': data.jo || '',
    '훈련자(조장)': data.leader || '',
    '훈련자': data.trainer || '',
    '훈련생': data.trainee1 || '',
    '접촉횟수': data.contact || '',
    '복음제시': data.gospel || '',
    '대상자': data.target || '',
    '결신': data.decision || '',
    '보류/거절': data.reject || '',
    '확신': data.confirm || '',
    '이미신자': data.believer || '',
    '거주지역': data.location || '',
    '성씨': data.name || '',
    '나이': data.age || '',
    '결혼여부': data.marriage || '',
    '성별': data.gender || '',
    '만난장소': data.meet_place || '',
    '복음전달': data.gospel_delivery || '',
    '대상자구분': data.target_type || '',
    '관계조원': data.relation_member || '',
    '종교배경': data.religion || '',
    '직업': data.job || '',
    '천국확신': data.heaven || '',
    '이유대답': data.reason || '',
    '영접여부': data.yeopjeop || '',
    '진단결과': data.diag_result || '',
    '간증6': data.testimony6 || '',
    '간증7': data.testimony7 || '',
    '훈련자(조장)역할': data.leader_part || '',
    '훈련자(조장)행동': data.leader_action || '',
    '훈련자역할': data.trainer_part || '',
    '훈련자행동': data.trainer_action || '',
    '훈련생역할': data.trainee_part || '',
    '훈련생행동': data.trainee_action || '',
    '결과': data.result || '',
    '즉석양육': data.immed_nurture || '',
    '후속양육': data.followup || '',
    '하지못한사유': data.followup_reason || '',
    '느낀점(10번)': data.q10 || '',
    '문제점(11번)': data.q11 || '',
    '기도제목(12번)': data.q12 || '',
    '기타(13번)': data.q13 || '',
    '풀복음 완료일': data.gospel_date || '',
    '특이사항(공유)': data.q14 || ''
  };

  var hMap = buildHeaderMap(sheet);
  var totalCols = sheet.getLastColumn();
  var rowArr = new Array(totalCols).fill('');
  Object.keys(FIELD_MAP).forEach(function(colName) {
    var idx = hMap[colName];
    if (idx) rowArr[idx - 1] = FIELD_MAP[colName];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, 1, totalCols).setValues([rowArr]);

  try {
    sendNotificationEmail(data, now);
  } catch(e) {
    Logger.log('이메일 발송 오류: ' + e.message);
  }

  // ── 특이사항 공유(14번) — 입력된 경우에만, 관리자에게만 별도 발송 ──
  if (data.q14 && String(data.q14).trim()) {
    try {
      sendSpecialNoteEmail(data, now);
    } catch(e) {
      Logger.log('특이사항 이메일 발송 오류: ' + e.message);
    }
  }

  return ok({ message: '저장 완료', savedAt: now });
}

// ============================================================
//  saveJo — 조정보 시트 덮어쓰기
// ============================================================
function saveJo(joData) {
  joData = joData || [];
  if (!joData.length) return err('joData 없음');
  var ss    = getSS();
  var sheet = ss.getSheetByName(SHEET_JO);
  if (!sheet) sheet = ss.insertSheet(SHEET_JO);

  sheet.clearContents();
  var headerRow = ['기수','그룹','조','반','조장','조장단계','훈련자','훈련자단계','훈련생','훈련생단계'];
  sheet.appendRow(headerRow);

  joData.forEach(function(j) {
    sheet.appendRow([
      j.kisu        || '',
      j.group       || '',
      j.jo          || '',
      j.ban         || '',
      j.leader      || '',
      j.leaderStep  || '',
      j.trainer     || '',
      j.trainerStep || '',
      j.trainee     || '',
      j.traineeStep || ''
    ]);
  });

  return ok({ saved: joData.length });
}

// ============================================================
//  그룹 편성 설정 시트 확인 / 자동 생성
//  ── 훈련기간(훈련시작일/훈련종료일)은 '관리자' 행에 저장 ──
//  ── 모든 값은 컬럼 위치가 아니라 "헤더 이름"으로 찾아 읽고/쓴다 ──
//     (그래야 시트에서 열 순서를 수동으로 바꾸거나, 열을 나중에
//      추가해도 안전하게 동작한다)
// ============================================================
var GROUP_SHEET_COLS = ['구분', '그룹명', '그룹장', '이메일', '부그룹장', '부그룹장이메일', '훈련시작일', '훈련종료일'];

function ensureGroupSheet() {
  var ss    = getSS();
  var sheet = ss.getSheetByName(SHEET_GROUP);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_GROUP);
    // 헤더
    sheet.appendRow(GROUP_SHEET_COLS);
    // 관리자 행
    sheet.appendRow(['관리자', '—', '—', '', '', '', '', '']);
    // 기본 그룹 행
    ['A','B','C','D'].forEach(function(id) {
      sheet.appendRow([id, id + '그룹', '', '', '', '', '', '']);
    });
    // 헤더 스타일
    var hdr = sheet.getRange(1, 1, 1, GROUP_SHEET_COLS.length);
    hdr.setBackground('#1a3a9e');
    hdr.setFontColor('#ffffff');
    hdr.setFontWeight('bold');
    // 관리자 행 스타일 (노란 배경)
    sheet.getRange(2, 1, 1, GROUP_SHEET_COLS.length).setBackground('#fff8e1');
    sheet.setColumnWidth(1, 60);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 200);
    sheet.setColumnWidth(7, 100);
    sheet.setColumnWidth(8, 100);
  } else {
    // 기존 시트에 없는 컬럼(부그룹장/부그룹장이메일 등)만 맨 끝에 자동 추가
    // — 이미 있는 열(순서 포함)은 절대 건드리지 않는다 (기존 데이터 보존)
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
    GROUP_SHEET_COLS.forEach(function(col) {
      if (headers.indexOf(col) === -1) {
        var newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(col);
        sheet.getRange(1, newCol).setBackground('#1a3a9e').setFontColor('#ffffff').setFontWeight('bold');
        headers.push(col);
      }
    });
  }
  return sheet;
}

// ============================================================
//  saveGroupSettings — 그룹 편성 설정 저장
//  body: { adminEmail, trainStart, trainEnd,
//          groups: [{id, name, leader, email, deputy, deputyEmail},...] }
//  ── 저장 시 GROUP_SHEET_COLS 순서로 시트를 재작성한다.
//     (구분/그룹명/그룹장/이메일/부그룹장/부그룹장이메일/훈련시작일/훈련종료일)
// ============================================================
function saveGroupSettings(body) {
  body = body || {};
  var sheet = ensureGroupSheet();
  var cols  = GROUP_SHEET_COLS;
  var idx   = {};
  cols.forEach(function(c, i) { idx[c] = i; });

  sheet.clearContents();
  // 헤더
  sheet.appendRow(cols);

  // 관리자 행 (훈련기간은 여기에 저장 — 그룹별이 아닌 전체 공통 값)
  var adminRow = new Array(cols.length).fill('');
  adminRow[idx['구분']]     = '관리자';
  adminRow[idx['그룹명']]   = '—';
  adminRow[idx['그룹장']]   = '—';
  adminRow[idx['이메일']]   = body.adminEmail || '';
  adminRow[idx['훈련시작일']] = body.trainStart || '';
  adminRow[idx['훈련종료일']] = body.trainEnd   || '';
  sheet.appendRow(adminRow);

  // 그룹 행
  var groups = body.groups || [];
  groups.forEach(function(g) {
    var row = new Array(cols.length).fill('');
    row[idx['구분']]        = g.id     || '';
    row[idx['그룹명']]      = g.name   || '';
    row[idx['그룹장']]      = g.leader || '';
    row[idx['이메일']]      = g.email  || '';
    row[idx['부그룹장']]     = g.deputy || '';
    row[idx['부그룹장이메일']] = g.deputyEmail || '';
    sheet.appendRow(row);
  });

  // 헤더 스타일 재적용
  var hdr = sheet.getRange(1, 1, 1, cols.length);
  hdr.setBackground('#1a3a9e');
  hdr.setFontColor('#ffffff');
  hdr.setFontWeight('bold');
  // 관리자 행 노란 배경
  sheet.getRange(2, 1, 1, cols.length).setBackground('#fff8e1');

  // ── 이메일 시트도 함께 동기화 (하위호환) ─────────────────
  syncEmailSheetFromGroupSettings(body.adminEmail, groups);

  return ok({ message: '그룹설정 저장 완료', groups: groups.length });
}

// ============================================================
//  getGroupSettings — 그룹 편성 설정 조회
//  반환: { adminEmail, trainStart, trainEnd, groups: [...] }
//  ── 시트의 실제 헤더 순서와 무관하게 헤더 "이름"으로 값을 찾는다 ──
// ============================================================
function getGroupSettings() {
  var sheet   = ensureGroupSheet();
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colIdx  = {};
  headers.forEach(function(h, i) { if (h) colIdx[h] = i; });

  var cell = function(row, colName) {
    var i = colIdx[colName];
    return (i === undefined) ? '' : row[i];
  };
  var fmtDate = function(v) {
    if (!v) return '';
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
    return String(v).trim();
  };

  var adminEmail = '';
  var trainStart = '';
  var trainEnd   = '';
  var groups     = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var id  = String(cell(row, '구분') || '').trim();
    if (!id) continue;

    if (id === '관리자') {
      adminEmail = String(cell(row, '이메일') || '').trim();
      trainStart = fmtDate(cell(row, '훈련시작일'));
      trainEnd   = fmtDate(cell(row, '훈련종료일'));
    } else {
      groups.push({
        id: id,
        name: String(cell(row, '그룹명') || '').trim(),
        leader: String(cell(row, '그룹장') || '').trim(),
        email: String(cell(row, '이메일') || '').trim(),
        deputy: String(cell(row, '부그룹장') || '').trim(),
        deputyEmail: String(cell(row, '부그룹장이메일') || '').trim()
      });
    }
  }

  return ok({ adminEmail: adminEmail, trainStart: trainStart, trainEnd: trainEnd, groups: groups });
}

// ============================================================
//  migrateEmailSheetToGroupSettings
//  — 기존 이메일 시트(이메일) 데이터를 그룹설정으로 일회성 마이그레이션
//  — Apps Script 편집기에서 수동으로 한 번 실행하세요
// ============================================================
function migrateEmailSheetToGroupSettings() {
  var ss        = getSS();
  var emailSheet = ss.getSheetByName('이메일');
  if (!emailSheet) {
    Logger.log('이메일 시트 없음 — 마이그레이션 불필요');
    return;
  }

  var data = emailSheet.getDataRange().getValues();
  var adminEmail = '';
  var groupEmails = {}; // { 'A': 'email', 'B': 'email', ... }

  for (var i = 1; i < data.length; i++) {
    var label = String(data[i][0] || '').trim();
    var email = String(data[i][1] || '').trim();
    if (!label) continue;
    if (label === '관리자') {
      adminEmail = email;
    } else {
      // 'A그룹장' → 'A'
      var match = label.match(/^([A-Z])그룹장$/);
      if (match) groupEmails[match[1]] = email;
    }
  }

  // 기존 그룹설정 시트 불러오기
  var gResult = getGroupSettings();
  var parsed  = JSON.parse(gResult.getContent());
  var groups  = parsed.groups || [];

  // 이메일 머지
  groups.forEach(function(g) {
    if (!g.email && groupEmails[g.id]) g.email = groupEmails[g.id];
  });
  if (!parsed.adminEmail && adminEmail) parsed.adminEmail = adminEmail;

  // 그룹설정에 저장 (훈련기간은 기존 값 유지)
  saveGroupSettings({
    adminEmail: parsed.adminEmail,
    trainStart: parsed.trainStart,
    trainEnd:   parsed.trainEnd,
    groups:     groups
  });

  Logger.log('마이그레이션 완료 — 관리자: ' + parsed.adminEmail
    + ', 그룹: ' + groups.map(function(g){ return g.id+'('+g.email+')'; }).join(', '));
}

// 내부 호환용 (saveGroupSettings에서 호출 제거)
function syncEmailSheetFromGroupSettings() {
  // 통합 완료 — 더 이상 이메일 시트 동기화 불필요
}

// ============================================================
//  fixEvangType — 전도유형 수정 + 이력 기록
// ============================================================
function fixEvangType(body) {
  body = body || {};
  var docNumber = String(body.docNumber || '');
  var newType   = String(body.newType   || '');
  var oldType   = String(body.oldType   || '');
  if (!docNumber || !newType) return err('docNumber, newType 필요');

  var ss    = getSS();
  var sheet = ss.getSheetByName(SHEET_REPORT);
  if (!sheet) return err('시트 없음: ' + SHEET_REPORT);

  var hMap    = buildHeaderMap(sheet);
  var docCol  = hMap['문서번호'];
  var typeCol = hMap['전도유형'];
  if (!docCol || !typeCol) return err('문서번호 또는 전도유형 컬럼 없음');

  var histCol = ensureHistCol(sheet, '전도유형_수정이력');
  var data    = sheet.getDataRange().getValues();
  var now     = nowKST();
  var updated = 0;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][docCol - 1]).trim() !== docNumber) continue;
    var prevVal  = String(data[i][typeCol - 1]);
    var entryOld = oldType || prevVal;

    sheet.getRange(i + 1, typeCol).setValue(newType);
    appendHist(sheet, i + 1, histCol,
      '[' + now + '] ' + entryOld + ' → ' + newType);
    updated++;
  }

  return ok({ updated: updated });
}

// ============================================================
//  renumberDoc — 문서번호 재넘버링 + 이력 기록
// ============================================================
function renumberDoc(body) {
  body = body || {};
  var oldDoc = String(body.oldDoc || '');
  var newDoc = String(body.newDoc || '');
  if (!oldDoc || !newDoc) return err('oldDoc, newDoc 필요');
  if (oldDoc === newDoc) return ok({ updated: 0, message: '변경 없음' });

  var ss    = getSS();
  var sheet = ss.getSheetByName(SHEET_REPORT);
  if (!sheet) return err('시트 없음: ' + SHEET_REPORT);

  var hMap   = buildHeaderMap(sheet);
  var docCol = hMap['문서번호'];
  if (!docCol) return err('문서번호 컬럼 없음');

  var histCol = ensureHistCol(sheet, '문서번호_수정이력');
  var data    = sheet.getDataRange().getValues();
  var now     = nowKST();
  var updated = 0;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][docCol - 1]).trim() !== oldDoc) continue;
    sheet.getRange(i + 1, docCol).setValue(newDoc);
    appendHist(sheet, i + 1, histCol,
      '[' + now + '] ' + oldDoc + ' → ' + newDoc);
    updated++;
  }

  return ok({ updated: updated });
}

// ── 이메일 시트 함수 제거 완료 (그룹설정으로 통합) ────────────

function getEmailsByGroup(group) {
  // 그룹설정 시트에서만 조회 (이메일 시트 통합 완료)
  var ss     = getSS();
  var gSheet = ss.getSheetByName(SHEET_GROUP);
  if (!gSheet) return { admin: [], group: [] };

  var gData       = gSheet.getDataRange().getValues();
  var headers     = gData[0].map(function(h) { return String(h).trim(); });
  var idCol       = headers.indexOf('구분');
  var emailCol    = headers.indexOf('이메일');
  var deputyCol   = headers.indexOf('부그룹장이메일');
  var adminEmails = [];
  var groupEmails = [];
  var targetId    = (group || '').toUpperCase();

  var splitEmails = function(v) {
    return String(v || '').split(/[,，\s]+/)
      .map(function(e) { return e.trim(); })
      .filter(function(e) { return e.indexOf('@') > -1; });
  };

  for (var i = 1; i < gData.length; i++) {
    var id     = idCol    >= 0 ? String(gData[i][idCol] || '').trim()    : '';
    var email  = emailCol >= 0 ? String(gData[i][emailCol] || '').trim() : '';
    var deputyEmail = deputyCol >= 0 ? String(gData[i][deputyCol] || '').trim() : '';
    if (!email && !deputyEmail) continue;
    var list = splitEmails(email).concat(splitEmails(deputyEmail));
    if (id === '관리자') {
      adminEmails = adminEmails.concat(list);
    } else if (id === targetId) {
      groupEmails = groupEmails.concat(list);
    }
  }
  return { admin: adminEmails, group: groupEmails };
}

function sendNotificationEmail(data, savedAt) {
  var group = '';
  var joSheet = getSS().getSheetByName(SHEET_JO);
  if (joSheet) {
    var joData    = joSheet.getDataRange().getValues();
    var joHeaders = joData[0].map(function(h) { return String(h).trim(); });
    var joCol  = joHeaders.indexOf('조');
    var grpCol = joHeaders.indexOf('그룹');
    if (joCol >= 0 && grpCol >= 0) {
      for (var i = 1; i < joData.length; i++) {
        if (String(joData[i][joCol]) === String(data.jo || '')) {
          group = String(joData[i][grpCol] || '').toUpperCase();
          break;
        }
      }
    }
  }

  var emails     = getEmailsByGroup(group);
  var recipients = emails.admin.concat(emails.group);
  var seen = {};
  recipients = recipients.filter(function(e) {
    if (seen[e]) return false;
    seen[e] = true;
    return true;
  });
  if (!recipients.length) return;

  var subject = '[전도현황] ' + (data.jo || '') + '조 '
    + (data.evang_type || '') + ' 보고서 제출 — '
    + (data.date || savedAt.slice(0, 10));

  var body = [
    '■ 전도 보고서가 제출되었습니다.',
    '',
    '저장일시 : ' + savedAt,
    '문서번호 : ' + (data.doc_number  || ''),
    '조       : ' + (data.jo          || '') + '조  (' + (group || '-') + '그룹)',
    '전도일   : ' + (data.date        || ''),
    '전도유형 : ' + (data.evang_type  || ''),
    '훈련자(조장) : ' + (data.leader  || ''),
    '훈련자   : ' + (data.trainer     || ''),
    '훈련생   : ' + (data.trainee1    || ''),
    '',
    '■ 통계',
    '복음제시 : ' + (data.gospel   || 0),
    '대상자   : ' + (data.target   || 0),
    '결신     : ' + (data.decision || 0),
    '확신     : ' + (data.confirm  || 0),
    '보류/거절: ' + (data.reject   || 0),
    '이미신자 : ' + (data.believer || 0),
    '',
    '■ 결과   : ' + (data.result   || ''),
    '■ 후속양육: ' + (data.followup || ''),
    '',
    '─────────────────────────────',
    '이 메일은 자동 발송된 알림입니다.'
  ].join('\n');

  MailApp.sendEmail({ to: recipients.join(','), subject: subject, body: body });
}

// ============================================================
//  sendSpecialNoteEmail — 14번 "특이사항 공유" 전용 알림
//  ── 일반 제출 알림과 분리하여 관리자에게만 발송 (그룹장 제외) ──
//  ── 14번 항목이 입력된 경우에만 saveReport()에서 호출됨 ──
// ============================================================
function sendSpecialNoteEmail(data, savedAt) {
  // group을 빈 값으로 넘기면 getEmailsByGroup은 관리자 이메일만 반환한다
  var emails     = getEmailsByGroup('');
  var recipients = emails.admin;
  var seen = {};
  recipients = recipients.filter(function(e) {
    if (seen[e]) return false;
    seen[e] = true;
    return true;
  });
  if (!recipients.length) return;

  var subject = '[특이사항] ' + (data.jo || '') + '조 '
    + (data.leader || '') + ' 조장 — '
    + (data.date || savedAt.slice(0, 10));

  var body = [
    '■ 관리자에게만 전달되는 특이사항입니다. (그룹장 미공유)',
    '',
    '저장일시 : ' + savedAt,
    '문서번호 : ' + (data.doc_number || ''),
    '조       : ' + (data.jo         || '') + '조',
    '전도일   : ' + (data.date       || ''),
    '훈련자(조장) : ' + (data.leader || ''),
    '훈련자   : ' + (data.trainer    || ''),
    '훈련생   : ' + (data.trainee1   || ''),
    '',
    '■ 특이사항 공유 내용',
    String(data.q14 || ''),
    '',
    '─────────────────────────────',
    '이 메일은 자동 발송된 알림입니다. 관리자 대시보드(justee62_admin.html)에서도 확인할 수 있습니다.'
  ].join('\n');

  MailApp.sendEmail({ to: recipients.join(','), subject: subject, body: body });
}


// ============================================================
//  deleteDoc — 문서 삭제 + 삭제이력 시트 기록
// ============================================================
function deleteDoc(body) {
  body = body || {};
  var docNumber = String(body.docNumber || '');
  if (!docNumber) return err('docNumber 필요');

  var ss    = getSS();
  var sheet = ss.getSheetByName(SHEET_REPORT);
  if (!sheet) return err('시트 없음: ' + SHEET_REPORT);

  var hMap   = buildHeaderMap(sheet);
  var docCol = hMap['문서번호'];
  if (!docCol) return err('문서번호 컬럼 없음');

  var histSheetName = '삭제이력';
  var histSheet = ss.getSheetByName(histSheetName);
  if (!histSheet) {
    histSheet = ss.insertSheet(histSheetName);
    histSheet.appendRow(['삭제일시','문서번호','전도일','조','전도유형','저장일시','삭제자정보']);
  }

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var now     = nowKST();
  var deleted = 0;

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][docCol - 1]).trim() !== docNumber) continue;
    var dateCol  = headers.indexOf('전도일');
    var joCol    = headers.indexOf('조');
    var typeCol  = headers.indexOf('전도유형');
    var savedCol = headers.indexOf('저장일시');
    histSheet.appendRow([
      now, docNumber,
      dateCol  >= 0 ? data[i][dateCol]  : '',
      joCol    >= 0 ? data[i][joCol]    : '',
      typeCol  >= 0 ? data[i][typeCol]  : '',
      savedCol >= 0 ? data[i][savedCol] : '',
      body.deletedBy || ''
    ]);
    sheet.deleteRow(i + 1);
    deleted++;
  }

  return ok({ deleted: deleted });
}

// ============================================================
//  deleteDocRow — 시트의 특정 "행 하나"만 삭제 + 삭제이력 기록
//  (문서번호 전체가 아니라, 같은 문서번호를 가진 여러 수정이력 행 중
//   일부만 정리하고 싶을 때 사용 — 예: 최신 행은 남기고 이전 행만 삭제)
//  body: { rowIndex, docNumber, deletedBy }
//  rowIndex는 doGet의 get/get_by_doc 응답에 포함된 rowIndex를 그대로 사용
//  docNumber는 안전장치용 — 실제 행의 문서번호와 일치하는지 대조 후에만 삭제
// ============================================================
function deleteDocRow(body) {
  body = body || {};
  var rowIndex  = parseInt(body.rowIndex, 10);
  var docNumber = String(body.docNumber || '');
  if (!rowIndex || rowIndex < 2) return err('rowIndex 필요 (2 이상)');

  var ss    = getSS();
  var sheet = ss.getSheetByName(SHEET_REPORT);
  if (!sheet) return err('시트 없음: ' + SHEET_REPORT);

  var lastRow = sheet.getLastRow();
  if (rowIndex > lastRow) return err('행 번호가 시트 범위를 벗어났습니다 (다른 사람이 먼저 삭제/변경했을 수 있음) — 새로고침 후 다시 시도하세요.');

  var hMap    = buildHeaderMap(sheet);
  var docCol  = hMap['문서번호'];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  var rowVals = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 안전장치: 넘어온 문서번호와 실제 그 행의 문서번호가 일치하는지 확인
  // (행이 그 사이에 밀렸다면 엉뚱한 행이 지워지는 걸 막기 위함)
  if (docCol && docNumber) {
    var actualDoc = String(rowVals[docCol - 1]).trim();
    if (actualDoc !== docNumber) {
      return err('문서번호 불일치 — 행이 이동되었을 수 있습니다. 새로고침 후 다시 시도하세요.');
    }
  }

  // 안전장치 2: 같은 문서번호 행이 여러 개일 때 엉뚱한 행이 지워지지 않도록
  // 저장일시(초 단위)까지 일치하는지 확인 (클라이언트가 savedAt을 보낸 경우에만 검사)
  var savedColChk = headers.indexOf('저장일시');
  var actualSaved = savedColChk >= 0 ? String(formatCellValue(rowVals[savedColChk], '저장일시')).trim() : '';
  if (body.savedAt && savedColChk >= 0 && String(body.savedAt).trim() !== actualSaved) {
    return err('저장일시 불일치 — 행이 이동되었을 수 있습니다. (요청: ' + body.savedAt + ' / 실제: ' + actualSaved + ') 새로고침 후 다시 시도하세요.');
  }

  var histSheetName = '삭제이력';
  var histSheet = ss.getSheetByName(histSheetName);
  if (!histSheet) {
    histSheet = ss.insertSheet(histSheetName);
    histSheet.appendRow(['삭제일시','문서번호','전도일','조','전도유형','저장일시','삭제자정보']);
  }

  var dateCol  = headers.indexOf('전도일');
  var joCol    = headers.indexOf('조');
  var typeCol  = headers.indexOf('전도유형');
  var savedCol = headers.indexOf('저장일시');
  var now = nowKST();
  histSheet.appendRow([
    now, docNumber || (docCol ? rowVals[docCol - 1] : ''),
    dateCol  >= 0 ? rowVals[dateCol]  : '',
    joCol    >= 0 ? rowVals[joCol]    : '',
    typeCol  >= 0 ? rowVals[typeCol]  : '',
    savedCol >= 0 ? rowVals[savedCol] : '',
    body.deletedBy || ''
  ]);

  var lastRowBefore = sheet.getLastRow();
  sheet.deleteRow(rowIndex);
  SpreadsheetApp.flush();
  return ok({ deleted: 1, rowIndex: rowIndex, docNumber: docNumber, savedAt: actualSaved,
              sheetName: sheet.getName(), lastRowBefore: lastRowBefore, lastRowAfter: sheet.getLastRow() });
}

// ── 유틸: 시트 수동 생성 트리거 ─────────────────────────────
function createGroupSheet() { ensureGroupSheet(); }
// function runMigration() { migrateEmailSheetToGroupSettings(); }

// ============================================================
//  renameColumn — 전도보고서 시트 특정 컬럼명 변경
//  Apps Script 편집기에서 수동으로 한 번 실행하세요.
//  예: '복음제시장소' → '복음전달' (X열 — 더 이상 쓰지 않는 질문을
//      대면/비대면 값으로 재사용)
// ============================================================
function renameColumn() {
  var OLD_NAME = '복음제시장소';
  var NEW_NAME = '복음전달';

  var ss    = getSS();
  var sheet = ss.getSheetByName(SHEET_REPORT);
  if (!sheet) {
    Logger.log('❌ 시트 없음: ' + SHEET_REPORT);
    return;
  }

  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    Logger.log('❌ 시트가 비어 있습니다.');
    return;
  }

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var found = false;

  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    if (h === OLD_NAME) {
      sheet.getRange(1, i + 1).setValue(NEW_NAME);
      Logger.log('✅ 변경 완료: "' + OLD_NAME + '" → "' + NEW_NAME + '" (열 ' + (i + 1) + ')');
      found = true;
    } else if (h === NEW_NAME) {
      Logger.log('⏭ 이미 "' + NEW_NAME + '"으로 되어 있습니다. (열 ' + (i + 1) + ')');
      found = true;
    }
  }

  if (!found) {
    Logger.log('⚠ "' + OLD_NAME + '" 컬럼을 찾을 수 없습니다. 현재 헤더: ' + headers.filter(Boolean).join(', '));
  }

  SpreadsheetApp.flush();
}

// ============================================================
//  addMissingColumns — 전도보고서 시트에 누락 컬럼 자동 추가
//  Apps Script 편집기에서 수동으로 한 번 실행하세요.
//  기존 데이터는 그대로 유지되며, 없는 컬럼만 맨 끝에 추가됩니다.
// ============================================================
function addMissingColumns() {
  var ss    = getSS();
  var sheet = ss.getSheetByName(SHEET_REPORT);
  if (!sheet) {
    Logger.log('❌ 시트 없음: ' + SHEET_REPORT);
    return;
  }

  // 추가해야 할 컬럼 목록 (순서대로)
  var REQUIRED_COLS = [
    '저장일시','주차','전도일','문서번호','전도유형','횟수','조',
    '훈련자(조장)','훈련자','훈련생',
    '접촉횟수','복음제시','대상자','결신','보류/거절','확신','이미신자',
    '거주지역','성씨','나이','결혼여부','성별','만난장소','복음전달',
    '대상자구분','관계조원','종교배경','직업',
    '천국확신','이유대답','영접여부','진단결과',
    '간증6','간증7',
    '훈련자(조장)역할','훈련자(조장)행동',
    '훈련자역할','훈련자행동',
    '훈련생역할','훈련생행동',
    '결과','즉석양육','후속양육','하지못한사유',
    '느낀점(10번)','문제점(11번)','기도제목(12번)','기타(13번)',
    '풀복음 완료일',
    '특이사항(공유)'
  ];

  // 현재 헤더 읽기
  var lastCol   = sheet.getLastColumn();
  var headers   = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); })
    : [];

  var added = [];
  REQUIRED_COLS.forEach(function(col) {
    if (headers.indexOf(col) === -1) {
      // 없으면 맨 끝에 추가
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(col);
      // 헤더 스타일 적용 (기존 헤더와 동일하게)
      var hdrRange = sheet.getRange(1, newCol);
      hdrRange.setBackground('#1a3a9e');
      hdrRange.setFontColor('#ffffff');
      hdrRange.setFontWeight('bold');
      headers.push(col); // 중복 방지
      added.push(col);
      Logger.log('✅ 컬럼 추가: ' + col + ' (열 ' + newCol + ')');
    } else {
      Logger.log('⏭ 이미 존재: ' + col);
    }
  });

  if (added.length === 0) {
    Logger.log('✔ 모든 컬럼이 이미 존재합니다. 추가 없음.');
  } else {
    Logger.log('📋 추가된 컬럼 (' + added.length + '개): ' + added.join(', '));
  }
  SpreadsheetApp.flush();
}

// ============================================================
//  ocrFieldInfo — "만난 분의 인적 정보" 사진을 Claude 비전으로 인식
//  ── 사전 준비: Apps Script 편집기 > 프로젝트 설정(⚙️) > 스크립트 속성
//     에 ANTHROPIC_API_KEY 등록 (https://console.anthropic.com 에서 발급) ──
//  ── 일반 조원 누구나 호출 가능한 보조 기능 — 폼을 채워줄 뿐 데이터를
//     직접 저장/변경하지 않으므로 admin 토큰 검증 대상이 아님 ──
// ============================================================
function ocrFieldInfo(imageDataUrl) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    return ok({ status: 'error', message: 'ANTHROPIC_API_KEY 스크립트 속성이 설정되지 않았습니다. 관리자에게 문의해주세요.' });
  }

  var m = String(imageDataUrl || '').match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!m) {
    return ok({ status: 'error', message: '이미지 데이터 형식이 올바르지 않습니다.' });
  }
  var mediaType  = m[1];
  var base64Data = m[2];

  var instruction =
    '이 사진은 한국 전도폭발훈련(EE) 현장전도 보고서의 "만난 분의 인적 정보" 작성란입니다. ' +
    '손글씨 또는 인쇄된 내용을 읽어서 아래 JSON 스키마로만 응답하세요. 설명, 코드블록 표시(```) 없이 JSON 객체 하나만 출력하세요.\n' +
    '스키마:\n' +
    '{\n' +
    '  "location": "거주지역(동/구/시 단위, 짧게)",\n' +
    '  "name": "이름",\n' +
    '  "age": "나이(숫자만, 문자열)",\n' +
    '  "gender": "여자 또는 남자",\n' +
    '  "marriage": "기혼 또는 미혼",\n' +
    '  "job": "직업",\n' +
    '  "meet_place": "만난 장소",\n' +
    '  "religion": "기독교|천주교|불교|무교|이슬람|천부교|대순진리교|유교 중 하나 또는 그대로",\n' +
    '  "heaven": "있었고 또는 없었고 (천국에 대한 확신 여부)",\n' +
    '  "reason": "믿음|행위|불확실 중 하나 (천국 가는 이유에 대한 대답 성격)",\n' +
    '  "date": "YYYY-MM-DD (전도일이 적혀 있으면)",\n' +
    '  "q12": "요청받은 기도제목 (있으면 원문 그대로)"\n' +
    '}\n' +
    '읽을 수 없거나 적혀 있지 않은 항목은 빈 문자열("")로 두세요. 절대로 값을 추측해서 채우지 마세요.';

  var payload = {
    model: 'claude-sonnet-4-6',
    max_tokens: 1024,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mediaType, data: base64Data } },
        { type: 'text', text: instruction }
      ]
    }]
  };

  var res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    return ok({ status: 'error', message: '네트워크 오류: ' + e.message });
  }

  var code = res.getResponseCode();
  var raw  = res.getContentText();
  var body;
  try {
    body = JSON.parse(raw);
  } catch (e) {
    return ok({ status: 'error', message: 'API 응답을 해석할 수 없습니다.' });
  }

  if (code !== 200) {
    var apiMsg = (body.error && body.error.message) ? body.error.message : ('HTTP ' + code);
    return ok({ status: 'error', message: 'Claude API 오류: ' + apiMsg });
  }

  var textBlock = null;
  (body.content || []).forEach(function(block) {
    if (!textBlock && block.type === 'text') textBlock = block;
  });
  if (!textBlock || !textBlock.text) {
    return ok({ status: 'error', message: '인식 결과가 비어 있습니다.' });
  }

  var cleaned = textBlock.text.replace(/```json|```/g, '').trim();
  var parsedFields;
  try {
    parsedFields = JSON.parse(cleaned);
  } catch (e) {
    return ok({ status: 'error', message: '응답 형식을 해석하지 못했습니다.' });
  }

  // 빈 문자열 필드는 제거 (프론트에서 falsy 체크로 건너뜀)
  Object.keys(parsedFields).forEach(function(k) {
    if (parsedFields[k] === '' || parsedFields[k] === null) delete parsedFields[k];
  });

  return ok({ status: 'ok', parsed: parsedFields });
} 
function testFetch() {
  var res = UrlFetchApp.fetch('https://www.google.com');
  Logger.log(res.getResponseCode());
}
