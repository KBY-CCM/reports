/**
 * 출연료 월정산 자동화 — 금병영  v15 (2026-10-30)
 * ---------------------------------------------------------------
 * 단위: (출연일 × 인물) 한 줄. 영상수는 따로 센다. (섭외 IN/OUT 시트 방식)
 * 흐름: 캘린더 + 유튜브 수집 → 팀 탭 선채움 → 팀이 ✔/단가기준/단가 입력
 *       → 취합 탭(지출&수입RAW A~I) → HJ가 통합문서에 붙임
 * 업데이트 방법: Apps Script에 이 파일을 통째로 덮어쓰고 저장 → 메뉴 > 최초 설정(setup) 한 번 더
 * ---------------------------------------------------------------
 */

const VERSION = 'v15';
const TZ = 'Asia/Seoul';
const TEAM_TABS = ['입력_침투부', '입력_깔깔튜브', '입력_챡'];
const MAX_ROWS = 300;

const TEAM_HEADERS = ['월', '채널/출처', '게스트명\n(인물마스터 표기)', '출연일\n(방송/녹화)', '내용\n(대표 영상 제목)', '편집본',
  '업로드일', '✔확인', '단가기준', '출연단가(세전)', '출연료(세전)\n=자동', '비고\n(RAW 비고로 들어감)',
  '신규여부(자동)', '구분(자동)', '담당자 메모', '영상 링크'];
const C = { YM: 1, CH: 2, NAME: 3, DATE: 4, CONTENT: 5, NVID: 6, UP: 7, CHK: 8, BASIS: 9, UNIT: 10, AMT: 11, NOTE: 12, NEW: 13, TYPE: 14, MEMO: 15, LINKS: 16 };
// 헤더 첫 줄이 이 글자로 시작하면 그 열로 인식 → 열을 옮겨도 동작
const HEADER_KEYS = { YM: '월', CH: '채널', NAME: '게스트명', DATE: '출연일', CONTENT: '내용', NVID: '편집본', LINKS: '영상 링크', UP: '업로드일',
  CHK: '✔확인', BASIS: '단가기준', UNIT: '출연단가', AMT: '출연료', NOTE: '비고', NEW: '신규여부', TYPE: '구분', MEMO: '담당자 메모' };
const BASIS_OPTIONS = ['출연당', '회당', '특강', '광고', '광고주지불', '인바운드', '거절', '별도'];
const ZERO_BASIS = ['거절', '광고주지불', '인바운드'];

/** 팀 탭의 열 위치를 헤더 이름으로 찾음 (없으면 기본 위치) */
function cols(sh) {
  const n = Math.max(sh.getLastColumn(), TEAM_HEADERS.length);
  const h = sh.getRange(1, 1, 1, n).getValues()[0].map(x => String(x).split('\n')[0].trim());
  const k = {};
  Object.keys(HEADER_KEYS).forEach(key => {
    let i = h.findIndex(x => x === HEADER_KEYS[key]);
    if (i < 0) i = h.findIndex(x => x.startsWith(HEADER_KEYS[key]));
    k[key] = i >= 0 ? i + 1 : C[key];
  });
  return k;
}
function colLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

const COLLECT_HEADERS = ['월', '출처', '일자', '원문 제목', '추출 게스트', '매칭 이름', '영상ID/이벤트ID', '배정 탭', '상태', '수집시각'];
const PEOPLE_HEADERS = ['이름(RAW 표기)', '분류', '배수(자동)', '게스트여부', '유튜브 핸들\n(@제외)', '별칭/캘린더 표기\n(쉼표 구분)', '동기화 시각'];
const NEWP_HEADERS = ['월', '이름', '분류', '필요번호', '본명/예금주', '은행', '계좌번호', '게스트여부', '촬영일', '출연료(계약)', '정산명의', '본명', '계약서파일', '확인필요', '인물마스터 등록(자동)'];
const MISSING = '모두싸인에서 확인';
const OUT_HEADERS = ['대분류', '분류', '이름', '구분', '시기', '비용', '추가작업(필요시)', '최종금액(지급액)', '비고', '출처 탭', '신규여부'];
const OUT_NEWP_HEADERS = ['▶ 통합문서 인물정보에 붙일 행', '이름', '분류', '필요번호', '본명(필요시)', '은행', '계좌번호', '게스트여부-3개월이후삭제'];

const DEFAULT_SETTINGS = [
  ['XLSX_FILE_ID', '14ypyaW04Go3V-ev0ehlhr-bBXL_HBA2f', '★통합 문서1.xlsx 드라이브 파일 ID (인물정보 동기화용)'],
  ['XLSX_PEOPLE_SHEET', '인물정보', '통합문서 안의 인물정보 시트 이름'],
  ['TEAM_SS_ID', '', '팀 입력용 별도 스프레드시트 ID (비우면 이 파일 안에 팀 탭 생성). 팀원에겐 그 파일만 공유'],
  ['NEW_PEOPLE_FOLDER_ID', '15IiDgt0rGHzpIAScHF0j31wO9sy3kl-f', '계약서 추출 CSV(신규인물정보_YYYY-MM.csv)가 놓이는 드라이브 폴더 ID'],
  ['CAL_침착맨 생방 일정', '입력_침투부', '캘린더 이름 → 배정 탭 (이름은 메뉴 > 내 캘린더 목록 보기 에서 복사)'],
  ['CAL_NEW채널제작일정', '입력_깔깔튜브', '캘린더 이름 → 배정 탭'],
  ['CAL_통닭천사 BDC', '입력_챡|통천,통닭천사,이세화', '캘린더 이름 → 배정 탭 | 제외할 호스트 이름(쉼표)'],
  ['CAL_KEYWORDS', '녹화,촬영,초대석,특강,합방,출연,방송,BDC,광고', '캘린더 일정 중 이 단어가 들어간 것만 출연 후보로 봄 (쉼표 구분, 비우면 전부)'],
  ['HOST_EXCLUDE', '침착맨,이병건,침투부,김현지', '모든 채널·캘린더에서 항상 제외할 이름 (쉼표)'],
  ['CAL_EXCLUDE', '기획안,회의,연차,피드백,전달,최종본,미팅', '이 단어가 들어간 일정은 제외'],
  ['YT_ChimChakMan_Official', '입력_침투부|desc|침착맨', '채널 핸들 → 탭 | desc=설명글 출연: 줄 | 제외할 이름(쉼표)'],
  ['YT_chicken_angel', '입력_깔깔튜브|title|통닭천사,통천,침착맨', '채널 핸들 → 탭 | title=제목 괄호 | 제외'],
  ['YT_ChimChakMan_Plus', '입력_침투부|desc', '플러스 채널 (침투부 탭으로 통합)'],
  ['YT_chaak챡', '입력_챡|desc|통천,통닭천사,이세화', '챡 채널 (핸들은 실제 값으로) | 제외할 호스트'],
  ['SLACK_WEBHOOK', '', '슬랙 Incoming Webhook URL (비우면 슬랙 알림 생략)'],
  ['SLACK_MSG_COLLECT', '📋 *{년}년 {월}월 출연료 정리* 시작합니다!\n<@멤버ID1> <@멤버ID2> 각 팀 탭에서 ✔확인·단가기준·단가·비고 입력 부탁드려요.\n마감: *{마감월}월 1일* (5일 지급)\n{현황}\n{링크}', '25일 안내 문구. 자유롭게 수정. 치환어: {년} {월} {마감월} {현황} {링크}. 멘션은 <@멤버ID> 형식(프로필→⋯→멤버 ID 복사)'],
  ['SLACK_MSG_REMIND', '⏰ *{월}월 출연료 입력 마감일*입니다! <@멤버ID1> <@멤버ID2> 미완료 건 오늘 중 입력 부탁드려요.\n{현황}\n{링크}', '1일 리마인드 문구. 치환어 동일'],
  ['COLLECT_DAY', '25', '자동 수집 + 팀 안내 발송일'],
  ['REMIND_DAY', '1', '마감 리마인드 발송일 (전월분)'],
  ['NEW_GUEST_FORM_URL', '', '신규 게스트 계좌/세금정보 등록 구글폼 URL'],
  ['OWNER_EMAIL', Session.getActiveUser().getEmail(), '운영자 메일 (취합 결과/오류 알림)'],
  ['담당_입력_침투부', '김현지', '탭 담당자 (슬랙 안내에 표시) — 테스트 기간엔 전부 김현지'],
  ['담당_입력_깔깔튜브', '김현지', '탭 담당자'],
  ['담당_입력_챡', '김현지', '탭 담당자'],
];

// ─────────────────────────── 메뉴 ───────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('출연료정산 ' + VERSION)
    .addItem('① 이번 달 수집 (팀 탭 선채움)', 'collectThisMonth')
    .addItem('① 지난 달 수집', 'collectLastMonth')
    .addItem('② 팀에 슬랙 안내 보내기', 'notifyCollect')
    .addItem('③ 취합 만들기 (지난 달)', 'buildOutputLastMonth')
    .addItem('③ 취합 만들기 (이번 달)', 'buildOutputThisMonth')
    .addItem('④ 월 마감 — 팀 탭 → 아카이브 (지난 달)', 'archiveLastMonth')
    .addItem('④ 월 마감 — 팀 탭 → 아카이브 (이번 달)', 'archiveThisMonth')
    .addSeparator()
    .addItem('인물정보 동기화 (통합문서→시트)', 'syncPeople')
    .addItem('계약서 추출본 가져오기 (신규인물정보 탭)', 'importNewPeopleMenu')
    .addItem('내 캘린더 목록 보기', 'listCalendars')
    .addItem('월간 자동 트리거 설치', 'installTriggers')
    .addItem('최초 설정 (setup)', 'setup')
    .addItem('이 파일의 입력 탭 → 팀 파일로 이사 (1회)', 'migrateTeamTabs')
    .addItem('팀 탭 전부 비우기 (테스트용)', 'clearTeamTabs')
    .addToUi();
}

function getSettings() {
  const sh = ss().getSheetByName('설정');
  const out = {};
  if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) out[String(r[0]).trim()] = String(r[1] || '').trim(); });
  return out;
}
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
/** 팀 입력 탭이 있는 스프레드시트 (설정 TEAM_SS_ID 가 있으면 그 파일, 없으면 이 파일) */
let _teamSS = null;
function teamSS() {
  if (_teamSS) return _teamSS;
  const id = (getSettings().TEAM_SS_ID || '').trim();
  _teamSS = id ? SpreadsheetApp.openById(id) : ss();
  return _teamSS;
}
function getOrCreateIn(book, name) { return book.getSheetByName(name) || book.insertSheet(name); }

// ─────────────────────────── 최초 설정 ───────────────────────────
function setup() {
  const s = ss();
  const ts = teamSS();
  const guide = getOrCreateIn(ts, '안내');
  guide.clear(); guide.setColumnWidth(1, 900);
  const lines = [
    ['출연료 월정산 입력 시트 — 사용 안내'], [''],
    ['■ 일정'],
    ['· 매월 25일: 자동 수집 결과가 각 팀 탭에 채워지고 슬랙으로 안내가 갑니다.'],
    ['· 익월 1일 팀 입력 마감 → 2일 운영 검수 → 3일 확정 → 5일 지급'], [''],
    ['■ 한 줄 = 한 명의 한 번 출연 (같은 방송에서 영상이 여러 개 나와도 한 줄, 영상수만 늘어남)'], [''],
    ['■ 팀 탭에서 할 일 (회색 = 자동, 노랑 = 직접 입력)'],
    ['1) 자동으로 채워진 행이 실제 지급 대상이면 ✔확인 체크. 편집본 0 = 녹화 일정만 있고 영상이 아직 없음(확인 필요)'],
    ['2) 단가기준 선택: 출연당(한 번에 단가) / 회당(단가×영상수) / 특강 / 광고 / 별도 — 광고주지불·인바운드·거절은 0원 처리'],
    ['3) 출연단가(세전) 입력 → 출연료는 자동 계산'],
    ['4) 비고: 정산 RAW 비고로 그대로 들어갑니다. 비우면 "내용 (출연일)"로 자동'],
    ['5) 자동 수집에 없는 출연은 맨 아래 빈 행에 추가 (게스트명은 드롭다운)'],
    ['6) 드롭다운에 없는 게스트 → 이름을 직접 적고, "신규" 표시가 뜨면 [설정] 탭의 신규 게스트 등록 폼으로 계좌·세금정보 제출'], [''],
    ['■ 하지 말 것'],
    ['· 회색 열은 고치지 말고, 틀렸으면 ✔을 빼고 새 행에 다시 적어주세요'],
    ['· ✔ 없고 단가 없는 자동 행은 다음 수집 때 지워지고 다시 채워집니다 (손댄 행은 유지)'],
  ];
  guide.getRange(1, 1, lines.length, 1).setValues(lines);
  [1, 3, 7, 9, 17].forEach(r => guide.getRange(r, 1).setFontWeight('bold'));

  // 설정: 없는 키만 추가
  const cfg = getOrCreate('설정');
  if (cfg.getLastRow() < 1 || !cfg.getRange(1, 1).getValue()) cfg.getRange(1, 1, 1, 3).setValues([['키', '값', '설명']]);
  const have = new Set(cfg.getDataRange().getValues().slice(1).map(r => String(r[0]).trim()));
  const haveYT = [...have].some(k => k.startsWith('YT_'));
  DEFAULT_SETTINGS.forEach(([k, v, d]) => {
    if (have.has(k)) return;
    if (k.startsWith('YT_') && haveYT) return; // 이미 유튜브 키가 있으면 기본 YT 키는 안 넣음
    const r = cfg.getLastRow() + 1;
    cfg.getRange(r, 1, 1, 3).setValues([[k, v, d]]); cfg.getRange(r, 2).setBackground('#fef9c3');
  });
  cfg.setColumnWidth(1, 220); cfg.setColumnWidth(2, 360); cfg.setColumnWidth(3, 520);
  styleHeader(cfg, 3);

  const pplM = getOrCreate('인물마스터');
  pplM.getRange(1, 1, 1, PEOPLE_HEADERS.length).setValues([PEOPLE_HEADERS]); styleHeader(pplM, PEOPLE_HEADERS.length);
  pplM.getRange(2, 3, 500, 1).setFormula('=IF(B2="","",IF(B2="세금계산서",1.1,IF(B2="삼쩜삼",0.967,1)))');
  pplM.setColumnWidth(1, 170); pplM.setColumnWidth(5, 160); pplM.setColumnWidth(6, 260);
  s.setNamedRange('PEOPLE_NAMES', pplM.getRange('A2:A500'));
  // 팀 파일에도 인물마스터(이름·분류만) 미러 — 드롭다운/수식용
  const ppl = ts.getId() === s.getId() ? pplM : getOrCreateIn(ts, '인물마스터');
  if (ppl !== pplM) {
    ppl.getRange(1, 1, 1, 2).setValues([['이름(RAW 표기)', '분류']]); styleHeader(ppl, 2); ppl.setColumnWidth(1, 170);
    ts.setNamedRange('PEOPLE_NAMES', ppl.getRange('A2:A500'));
    mirrorPeopleToTeam();
  }

  const col = getOrCreate('자동수집');
  col.getRange(1, 1, 1, COLLECT_HEADERS.length).setValues([COLLECT_HEADERS]); styleHeader(col, COLLECT_HEADERS.length); col.setColumnWidth(4, 360);
  const np = getOrCreate('신규인물정보');
  np.getRange(1, 1, 1, NEWP_HEADERS.length).setValues([NEWP_HEADERS]); styleHeader(np, NEWP_HEADERS.length);
  [70, 120, 110, 150, 170, 110, 170, 80, 95, 110, 80, 90, 220, 220, 120].forEach((w, i) => np.setColumnWidth(i + 1, w));
  np.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(MISSING).setFontColor('#dc2626').setBold(true).setRanges([np.getRange(2, 1, MAX_ROWS, NEWP_HEADERS.length)]).build()]);
  const out = getOrCreate('취합');
  out.getRange(1, 1, 1, OUT_HEADERS.length).setValues([OUT_HEADERS]); styleHeader(out, OUT_HEADERS.length); out.setColumnWidth(9, 260);
  const arc = getOrCreate('입력_아카이브');
  if (!arc.getRange(1, 1).getValue()) { arc.getRange(1, 1, 1, TEAM_HEADERS.length + 2).setValues([['출처 탭', '마감일시'].concat(TEAM_HEADERS)]); styleHeader(arc, TEAM_HEADERS.length + 2); }

  TEAM_TABS.forEach(name => {
    const sh = getOrCreateIn(ts, name);
    const hasHeader = String(sh.getRange(1, 1).getValue()).trim() !== '';
    if (!hasHeader) {
      sh.getRange(1, 1, 1, TEAM_HEADERS.length).setValues([TEAM_HEADERS]);
      [70, 120, 160, 95, 280, 70, 90, 60, 90, 120, 120, 240, 110, 95, 180, 240].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    }
    const k = cols(sh);
    const nCol = Math.max(sh.getLastColumn(), TEAM_HEADERS.length);
    styleHeader(sh, nCol); sh.setRowHeight(1, 44);
    const body = sh.getRange(2, 1, MAX_ROWS, nCol);
    body.setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
    body.setNumberFormat('General');
    sh.getRange(2, k.YM, MAX_ROWS, 1).setNumberFormat('@');
    const ymVals = sh.getRange(2, k.YM, MAX_ROWS, 1).getValues();
    sh.getRange(2, k.YM, MAX_ROWS, 1).setValues(ymVals.map(x => [x[0] === '' ? '' : ymCell(x[0])]));
    ['YM', 'CH', 'DATE', 'CONTENT', 'NVID', 'LINKS', 'UP', 'AMT', 'NEW', 'TYPE'].forEach(c => sh.getRange(2, k[c], MAX_ROWS, 1).setBackground('#e5e7eb'));
    ['NAME', 'CHK', 'BASIS', 'UNIT', 'NOTE', 'MEMO'].forEach(c => sh.getRange(2, k[c], MAX_ROWS, 1).setBackground('#fef9c3'));
    applyRowFormulas(sh, 2, MAX_ROWS);
    sh.getRange(2, k.UNIT, MAX_ROWS, 1).setNumberFormat('#,##0'); sh.getRange(2, k.AMT, MAX_ROWS, 1).setNumberFormat('#,##0');
    sh.getRange(2, k.NVID, MAX_ROWS, 1).setNumberFormat('0').setHorizontalAlignment('center');
    // 편집본 0 → 빨간 글씨
    const nv = sh.getRange(2, k.NVID, MAX_ROWS, 1);
    sh.setConditionalFormatRules(sh.getConditionalFormatRules().filter(r => !r.getRanges().some(g => g.getColumn() === k.NVID)).concat([
      SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setFontColor('#dc2626').setBold(true).setRanges([nv]).build()]));
    sh.getRange(2, k.DATE, MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd'); sh.getRange(2, k.UP, MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, k.NAME, MAX_ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ppl.getRange('A2:A500'), true).setAllowInvalid(true).build());
    sh.getRange(2, k.BASIS, MAX_ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(BASIS_OPTIONS, true).setAllowInvalid(false).build());
    sh.setFrozenRows(1);
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    const p = sh.protect().setDescription('자동 채움 열 보호'); p.setWarningOnly(true);
    p.setUnprotectedRanges(['NAME', 'CHK', 'BASIS', 'UNIT', 'NOTE', 'MEMO'].map(c => sh.getRange(2, k[c], MAX_ROWS, 1)));
  });
  ['인물마스터', '자동수집', '취합', '설정', '신규인물정보', '입력_아카이브'].forEach(n => { const sh = s.getSheetByName(n); sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove()); sh.protect().setDescription('운영 전용').setWarningOnly(true); });
  if (ppl !== pplM) { ppl.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove()); ppl.protect().setDescription('운영 전용').setWarningOnly(true); }
  [s, ts].forEach(b => { const def = b.getSheetByName('시트1') || b.getSheetByName('Sheet1'); if (def && b.getSheets().length > 1) b.deleteSheet(def); });
  if (ts.getId() === s.getId()) s.setActiveSheet(guide);
  SpreadsheetApp.getUi().alert('설정 완료!' + (ts.getId() !== s.getId() ? ' 팀 탭은 별도 파일(' + ts.getName() + ')에 만들어졌어요.' : '') + '\n다음: 메뉴 > 인물정보 동기화 → 내 캘린더 목록 보기(이름 확인) → 이번 달 수집 → 트리거 설치');
}

function applyRowFormulas(sh, startRow, n) {
  const k = cols(sh), r = startRow, L = colLetter;
  sh.getRange(r, k.CHK, n, 1).insertCheckboxes();
  const zero = ZERO_BASIS.map(z => `${L(k.BASIS)}${r}="${z}"`).join(',');
  sh.getRange(r, k.AMT, n, 1).setFormula(`=IF(OR(${zero}),0,IF(${L(k.UNIT)}${r}="","",IF(${L(k.BASIS)}${r}="회당",${L(k.UNIT)}${r}*MAX(${L(k.NVID)}${r},1),${L(k.UNIT)}${r})))`);
  sh.getRange(r, k.NEW, n, 1).setFormula(`=IF(${L(k.NAME)}${r}="","",IF(ISNA(MATCH(${L(k.NAME)}${r},PEOPLE_NAMES,0)),"신규→폼제출",""))`);
  sh.getRange(r, k.TYPE, n, 1).setFormula(`=IFERROR(VLOOKUP(${L(k.NAME)}${r},인물마스터!$A:$B,2,FALSE),"")`);
}
function clearTeamTabs() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('팀 탭의 데이터와 헤더를 전부 지우고 다시 만듭니다. 계속할까요?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  TEAM_TABS.forEach(n => { const sh = teamSS().getSheetByName(n); if (sh) { sh.clear(); sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove()); } });
  setup();
}
function getOrCreate(name) { return ss().getSheetByName(name) || ss().insertSheet(name); }
/** 인물마스터 이름·분류를 팀 파일로 복사 (민감정보 없음) */
function mirrorPeopleToTeam() {
  const ts = teamSS(); if (ts.getId() === ss().getId()) return;
  const src = ss().getSheetByName('인물마스터'); const dst = getOrCreateIn(ts, '인물마스터');
  const v = src.getLastRow() > 1 ? src.getRange(2, 1, src.getLastRow() - 1, 2).getValues().filter(r => r[0]) : [];
  if (dst.getLastRow() > 1) dst.getRange(2, 1, dst.getLastRow() - 1, 2).clearContent();
  if (v.length) dst.getRange(2, 1, v.length, 2).setValues(v);
}
function styleHeader(sh, n) {
  sh.getRange(1, 1, 1, n).setBackground('#1f2937').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setFrozenRows(1);
}

// ─────────────────────────── 캘린더 목록 ───────────────────────────
function listCalendars() {
  const names = CalendarApp.getAllCalendars().map(c => c.getName());
  const msg = '이 계정에서 보이는 캘린더 (설정 탭 CAL_ 뒤에 이 이름을 그대로 쓰세요):\n\n' + names.join('\n');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  return names;
}

// ─────────────────────────── 인물정보 동기화 ───────────────────────────
function syncPeople() {
  const cfg = getSettings();
  const xlsxId = cfg.XLSX_FILE_ID;
  if (!xlsxId) throw new Error('설정 탭의 XLSX_FILE_ID가 비어있습니다.');
  const copy = Drive.Files.copy({ name: '_tmp_people_sync_' + Date.now(), mimeType: 'application/vnd.google-apps.spreadsheet' }, xlsxId, { supportsAllDrives: true });
  let rows = [];
  try {
    const tmp = SpreadsheetApp.openById(copy.id);
    const src = tmp.getSheetByName(cfg.XLSX_PEOPLE_SHEET || '인물정보');
    if (!src) throw new Error('통합문서에 "' + cfg.XLSX_PEOPLE_SHEET + '" 시트가 없습니다.');
    const v = src.getDataRange().getValues();
    const h = v[0].map(x => String(x).trim());
    const iName = h.indexOf('이름'), iType = h.indexOf('분류');
    const iGuest = h.findIndex(x => x.indexOf('게스트여부') === 0);
    v.slice(1).forEach(r => {
      const name = String(r[iName] || '').trim(); if (!name) return;
      rows.push([name, String(r[iType] || '').trim(), iGuest >= 0 ? String(r[iGuest] || '').trim() : '']);
    });
  } finally { Drive.Files.remove(copy.id, { supportsAllDrives: true }); }
  const ppl = ss().getSheetByName('인물마스터');
  const old = ppl.getLastRow() > 1 ? ppl.getRange(2, 1, ppl.getLastRow() - 1, 7).getValues() : [];
  const keep = {}; old.forEach(r => { if (r[0]) keep[String(r[0]).trim()] = { handle: r[4], alias: r[5] }; });
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  const outRows = rows.map(r => [r[0], r[1], '', r[2], (keep[r[0]] || {}).handle || '', (keep[r[0]] || {}).alias || '', now]);
  const inXlsx = new Set(rows.map(r => r[0]));
  old.forEach(r => { if (r[0] && !inXlsx.has(String(r[0]).trim())) outRows.push([r[0], r[1], '', r[3], r[4], r[5], r[6]]); });
  ppl.getRange(2, 1, Math.max(ppl.getLastRow() - 1, 1), 7).clearContent();
  if (outRows.length) {
    ppl.getRange(2, 1, outRows.length, 7).setValues(outRows);
    ppl.getRange(2, 3, outRows.length, 1).setFormula('=IF(B2="","",IF(B2="세금계산서",1.1,IF(B2="삼쩜삼",0.967,1)))');
  }
  mirrorPeopleToTeam();
  return rows.length;
}

// ─────────────────────────── 계약서 추출본(CSV) 가져오기 ───────────────────────────
function importNewPeopleMenu() { const n = importNewPeople(); prefillFromContracts(ymOf(new Date())); try { SpreadsheetApp.getUi().alert('신규인물정보 ' + n + '행 반영 + 팀 탭 단가 선채움 완료'); } catch (e) {} }

/** 드라이브 폴더의 신규인물정보_YYYY-MM.csv 전부를 신규인물정보 탭으로 (최근 6개월분) */
function importNewPeople() {
  const cfg = getSettings();
  const folderId = cfg.NEW_PEOPLE_FOLDER_ID; if (!folderId) return 0;
  const folder = DriveApp.getFolderById(folderId);
  const it = folder.getFiles(); const rows = [];
  const cutoff = ymOf(addMonths(new Date(), -6));
  while (it.hasNext()) {
    const f = it.next(); const m = /^신규인물정보_(\d{4}-\d{2})\.csv$/.exec(f.getName());
    if (!m || m[1] < cutoff) continue;
    const data = Utilities.parseCsv(f.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/, ''));
    const h = data[0].map(x => String(x).trim());
    const idx = n => h.indexOf(n);
    data.slice(1).forEach(r => {
      if (!r[idx('이름')]) return;
      rows.push([m[1], r[idx('이름')], r[idx('분류')], r[idx('필요번호')], r[idx('본명_예금주')], r[idx('은행')], r[idx('계좌번호')], r[idx('게스트여부')],
        r[idx('촬영일')], r[idx('출연료')], r[idx('정산명의')], r[idx('본명')], r[idx('계약서파일')], r[idx('확인필요')]]);
    });
  }
  const sh = ss().getSheetByName('신규인물정보');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, NEWP_HEADERS.length).clearContent();
  if (rows.length) {
    rows.sort((a, b) => (b[0] + b[8]).localeCompare(a[0] + a[8]));
    sh.getRange(2, 1, rows.length, 14).setValues(rows);
    sh.getRange(2, 15, rows.length, 1).setFormula('=IF(B2="","",IF(ISNA(MATCH(B2,PEOPLE_NAMES,0)),"미등록","등록됨"))');
    sh.getRange(2, 10, rows.length, 1).setNumberFormat('#,##0');
  }
  return rows.length;
}

/** 팀 탭에서 단가가 비어 있는 행에 계약서의 출연료/촬영일을 채움 (이름 일치 + 날짜 ±7일) */
function prefillFromContracts(ym) {
  const np = ss().getSheetByName('신규인물정보');
  const c = np.getLastRow() > 1 ? np.getRange(2, 1, np.getLastRow() - 1, 14).getValues() : [];
  const contracts = c.filter(r => r[1]).map(r => ({
    names: String(r[1]).split(/[,&\/]/).map(x => x.trim()).filter(Boolean).concat(r[11] ? [String(r[11]).trim()] : []),
    date: r[8] ? new Date(r[8]) : null, fee: Number(String(r[9]).replace(/[^\d]/g, '')) || 0, file: r[12] }));
  let n = 0;
  TEAM_TABS.forEach(tab => {
    const sh = teamSS().getSheetByName(tab); const k = cols(sh);
    const last = sh.getLastRow(); if (last < 2) return;
    const v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    v.forEach((r, i) => {
      if (ymCell(r[k.YM - 1]) !== ym || r[k.UNIT - 1] !== '' || r[k.BASIS - 1] !== '') return;
      const name = String(r[k.NAME - 1] || '').replace(/\s*(공식|개인|출연)$/, '').trim(); if (!name) return;
      const d = r[k.DATE - 1] ? new Date(r[k.DATE - 1]) : null;
      const hits = contracts.filter(ct => ct.names.some(nm => nm && (name.indexOf(nm) >= 0 || nm.indexOf(name) >= 0)))
        .filter(ct => !d || !ct.date || Math.abs((ct.date - d) / 86400000) <= 7);
      if (hits.length !== 1) return;
      const ct = hits[0]; const row = i + 2;
      sh.getRange(row, k.UNIT).setValue(ct.fee);
      sh.getRange(row, k.BASIS).setValue(ct.fee === 0 ? '인바운드' : '출연당');
      sh.getRange(row, k.MEMO).setValue('계약서 자동: ' + ct.file);
      n++;
    });
  });
  return n;
}

// ─────────────────────────── 수집 ───────────────────────────
function collectThisMonth() { collectMonth(ymOf(new Date())); }
function collectLastMonth() { collectMonth(ymOf(addMonths(new Date(), -1))); }

function collectMonth(ym) {
  const cfg = getSettings();
  const people = loadPeople();
  const [y, m] = ym.split('-').map(Number);
  const start = new Date(y, m - 1, 1), end = new Date(y, m, 1);
  const found = [], skipped = { noGuest: 0, calNoMatch: 0 };
  const calKw = (cfg.CAL_KEYWORDS || '').split(',').map(x => x.trim()).filter(Boolean);
  const calEx = (cfg.CAL_EXCLUDE || '').split(',').map(x => x.trim()).filter(Boolean);
  const hostEx = (cfg.HOST_EXCLUDE || '침착맨,이병건').split(',').map(x => x.trim()).filter(Boolean);

  // 1) 캘린더: 인물마스터 이름이 들어가고 (키워드 있으면) 키워드도 들어간 일정만
  Object.keys(cfg).filter(k => k.startsWith('CAL_') && k !== 'CAL_KEYWORDS' && k !== 'CAL_EXCLUDE').forEach(k => {
    const calName = k.substring(4);
    const cp = cfg[k].split('|').map(x => x.trim());
    const tab = cp[0];
    const calExNames = (cp[1] || '').split(',').map(x => x.trim()).filter(Boolean).concat(hostEx);
    const cals = CalendarApp.getCalendarsByName(calName);
    if (!cals.length) { found.push(row(ym, '캘린더:' + calName, '', '(캘린더를 찾지 못함 — 메뉴 "내 캘린더 목록 보기"로 이름 확인)', '', '', '', tab, '오류', 'cal')); return; }
    cals[0].getEvents(start, end).forEach(ev => {
      const text = ev.getTitle() + ' ' + (ev.getDescription() || '');
      if (calKw.length && !calKw.some(w => text.indexOf(w) >= 0)) { skipped.calNoMatch++; return; }
      if (calEx.some(w => text.indexOf(w) >= 0)) { skipped.calNoMatch++; return; }
      const names = extractNames(text, people).filter(n => !calExNames.some(x => n.name.indexOf(x) >= 0 || n.raw.indexOf(x) >= 0));
      if (!names.length) { skipped.calNoMatch++; return; }
      names.forEach(n => found.push(row(ym, '캘린더:' + calName, ev.getStartTime(), ev.getTitle(), n.raw, n.name, ev.getId(), tab, '', 'cal')));
    });
  });

  // 2) 유튜브
  Object.keys(cfg).filter(k => k.startsWith('YT_')).forEach(k => {
    const handle = k.substring(3).replace(/^@/, '');
    const parts = cfg[k].split('|').map(x => x.trim());
    const tab = parts[0], mode = parts[1] || 'desc';
    const exclude = (parts[2] || '').split(',').map(x => x.trim()).filter(Boolean).concat(hostEx);
    if (!tab || /핸들입력|핸들$/.test(handle)) return;
    let vids;
    try { vids = fetchUploads(handle, start, end); }
    catch (e) { found.push(row(ym, '유튜브:@' + handle, '', '(조회 실패) ' + e.message, '', '', '', tab, '오류', 'yt')); return; }
    vids.forEach(v => {
      let names = mode === 'title' ? extractFromTitle(v.title, people) : extractFromDesc(v.desc, people, handle);
      names = names.filter(n => !exclude.some(x => (n.name || n.raw).indexOf(x) >= 0));
      if (!names.length) { skipped.noGuest++; return; }
      const bDate = mode === 'desc' ? parseBroadcastDate(v.desc) : null;
      names.forEach(n => found.push(row(ym, '유튜브:@' + handle, v.published, v.title, n.raw, n.name, v.id, tab, '', 'yt', 'https://youtu.be/' + v.id, bDate)));
    });
  });

  const col = ss().getSheetByName('자동수집');
  clearMonthRows(col, ym, 1);
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  if (found.length) col.getRange(col.getLastRow() + 1, 1, found.length, 10)
    .setValues(found.map(f => [f.ym, f.source, f.date ? fmt(f.date) : '', f.title, f.extracted, f.matched, f.id, f.tab, f.status, stamp]));

  const prefill = mergeForTeams(found);
  let added = 0;
  TEAM_TABS.forEach(tab => { removeUntouchedRows(tab, ym); added += appendRows(tab, prefill.filter(p => p.tab === tab), ym); });
  try { importNewPeople(); prefillFromContracts(ym); } catch (e) { Logger.log('계약서 추출본 처리 실패: ' + e.message); }
  const msg = ym + ' 수집 완료: 출연 후보 ' + prefill.length + '건 → 팀 탭 ' + added + '행 (게스트 없는 영상 ' + skipped.noGuest + '개, 관계없는 일정 ' + skipped.calNoMatch + '개 제외)';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  return { found: found.length, added };
}

function row(ym, source, date, title, extracted, matched, id, tab, status, kind, link, bDate) {
  return { ym, source, date, title, extracted, matched, id, tab, status, kind, link: link || '', bDate: bDate || null };
}

/** 같은 탭·같은 게스트·같은 출연일(방송일) → 한 줄, 영상수 합산 */
function mergeForTeams(found) {
  const groups = {};
  found.filter(f => f.status !== '오류').forEach(f => {
    const guest = f.matched || f.extracted; if (!guest) return;
    const when = f.kind === 'yt' ? (f.bDate || f.date) : f.date;
    const key = [f.tab, guest, fmt(when)].join('|');
    const g = groups[key] = groups[key] || { tab: f.tab, ym: f.ym, name: guest, date: when, ch: f.source.replace(/^(유튜브|캘린더):/, ''), titles: [], links: [], ups: [], cal: false, yt: false, ids: [] };
    if (f.kind === 'yt') { g.yt = true; g.titles.push(f.title); g.links.push(f.link); g.ups.push(f.date); g.ids.push(f.id); }
    else { g.cal = true; if (!g.titles.length) g.titles.push(f.title); g.ids.push(f.id); }
  });
  // 캘린더만 있는 그룹을 같은 게스트의 유튜브 그룹(업로드 전 45일 내)에 흡수
  const list = Object.values(groups);
  const ytOnes = list.filter(g => g.yt), calOnes = list.filter(g => !g.yt);
  const absorbed = new Set();
  calOnes.forEach(c => {
    const cand = ytOnes.filter(y => y.tab === c.tab && y.name === c.name && !y.cal)
      .map(y => ({ y, d: (y.ups[0] - c.date) / 86400000 })).filter(x => x.d >= -1 && x.d <= 45).sort((a, b) => a.d - b.d)[0];
    if (cand) { cand.y.cal = true; cand.y.date = c.date; absorbed.add(c); }
  });
  return list.filter(g => !absorbed.has(g)).map(g => ({
    tab: g.tab, ym: g.ym, ch: g.ch, name: g.name, date: g.date,
    content: g.titles[0] || '', nvid: g.links.length, links: g.links.join('\n'),
    up: g.ups.length ? g.ups.sort((a, b) => a - b)[0] : '',
    status: '',
    key: g.ids.sort().join(',')
  }));
}

/** ✔ 없고 단가도 없는 자동 행(해당 월)을 지움 → 재수집 시 깨끗하게 */
function removeUntouchedRows(tab, ym) {
  const sh = teamSS().getSheetByName(tab); const k = cols(sh);
  const last = sh.getLastRow(); if (last < 2) return;
  const v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    const r = v[i];
    if (ymCell(r[k.YM - 1]) !== ym) continue;
    const touched = r[k.CHK - 1] === true || r[k.UNIT - 1] !== '' || r[k.NOTE - 1] !== '' || r[k.MEMO - 1] !== '' || r[k.BASIS - 1] !== '';
    if (!touched && r[k.CH - 1] !== '') sh.deleteRow(i + 2);
  }
}

function appendRows(tab, items, ym) {
  if (!items.length) return 0;
  const sh = teamSS().getSheetByName(tab); const k = cols(sh);
  const nCol = Math.max(sh.getLastColumn(), TEAM_HEADERS.length);
  const last = Math.max(sh.getLastRow(), 1);
  const existing = last > 1 ? sh.getRange(2, 1, last - 1, nCol).getValues() : [];
  const keys = new Set(existing.filter(r => r[k.YM - 1]).map(r => [ymCell(r[k.YM - 1]), r[k.NAME - 1], fmt(r[k.DATE - 1])].join('|')));
  const rows = [];
  items.sort((a, b) => (a.date || 0) - (b.date || 0)).forEach(it => {
    const key = [ym, it.name, fmt(it.date)].join('|');
    if (keys.has(key)) return; keys.add(key);
    const row = new Array(nCol).fill('');
    row[k.YM - 1] = ym; row[k.CH - 1] = it.ch; row[k.NAME - 1] = it.name; row[k.DATE - 1] = it.date ? fmt(it.date) : '';
    row[k.CONTENT - 1] = it.content; row[k.NVID - 1] = it.nvid; row[k.LINKS - 1] = it.links; row[k.UP - 1] = it.up ? fmt(it.up) : '';
    row[k.CHK - 1] = false;
    rows.push(row);
  });
  if (!rows.length) return 0;
  const vals = sh.getRange(2, k.YM, MAX_ROWS, 1).getValues();
  let r = 2; for (let i = vals.length - 1; i >= 0; i--) { if (vals[i][0] !== '') { r = i + 3; break; } }
  sh.getRange(r, k.YM, rows.length, 1).setNumberFormat('@');
  sh.getRange(r, 1, rows.length, nCol).setValues(rows);
  applyRowFormulas(sh, r, rows.length);
  sh.getRange(r, k.LINKS, rows.length, 1).setWrap(true);
  return rows.length;
}

// ─────────────────────────── 이름 추출 ───────────────────────────
function loadPeople() {
  const sh = ss().getSheetByName('인물마스터');
  const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];
  const list = [];
  v.forEach(r => {
    const name = String(r[0] || '').trim(); if (!name) return;
    const aliases = String(r[5] || '').split(',').map(x => x.trim()).filter(Boolean);
    const base = name.replace(/\s*(공식|개인|출연)$/, '');
    list.push({ name, handle: String(r[4] || '').replace(/^@/, '').trim().toLowerCase(), terms: uniq([name, base].concat(aliases)) });
  });
  return list;
}
function uniq(a) { return a.filter((x, i) => x && a.indexOf(x) === i); }

/** 텍스트에서 인물마스터 이름/별칭 (2글자 이상, 긴 이름 우선) */
function extractNames(text, people) {
  const hits = []; const t = text || '';
  people.forEach(p => p.terms.forEach(term => {
    if (term.length < 2) return;
    if (t.indexOf(term) >= 0 && !hits.some(h => h.name === p.name)) hits.push({ name: p.name, raw: term, len: term.length });
  }));
  return hits.filter(h => !hits.some(o => o !== h && o.raw.indexOf(h.raw) >= 0 && o.len > h.len));
}
/** 인물마스터에 있으면 그 표기, 없으면 원문 이름 그대로 */
function resolveName(raw, people) {
  const hit = extractNames(raw, people)[0];
  return { name: hit ? hit.name : '', raw: hit ? hit.name : raw };
}
/**
 * 설명글: "▶출연: 빠니보틀(@PaniBottle), 캡틴따거(@captainbrother), 침착맨" 줄에서 이름을 뽑음.
 * 그 줄이 없으면 @핸들만 (이메일 주소, 채널 자신, 회사 도메인은 제외)
 */
function extractFromDesc(desc, people, ownHandle) {
  const out = [];
  const push = (name, raw) => { if (raw && !out.some(o => o.raw === raw)) out.push({ name, raw }); };
  const lines = (desc || '').split(/\r?\n/);
  const castLines = lines.filter(l => /(출연|게스트)\s*[:：]/.test(l));
  if (castLines.length) {
    castLines.forEach(l => {
      let body = l.replace(/^.*?(출연|게스트)\s*[:：]\s*/, '');
      body = body.replace(/https?:\/\/[^\s)\]]+/g, '');                 // URL
      for (let k = 0; k < 4; k++) body = body.replace(/\([^()]*\)/g, ''); // (…) 중첩 제거
      body = body.replace(/\[[^\]]*\]/g, '').replace(/[（【][^）】]*[）】]/g, '');
      body.split(/[,，\/·&]|\s+and\s+/).forEach(seg => {
        let nm = seg.replace(/@\S+/g, '').replace(/[▶▷•·\s]+$/, '').trim();
        nm = nm.replace(/님$/, '').trim();
        if (nm.length >= 2 && nm.length <= 20) { const r = resolveName(nm, people); push(r.name, r.raw); }
      });
    });
    if (out.length) return out;
  }
  const re = /(^|[^A-Za-z0-9._\-])@([A-Za-z0-9._\-]{3,})/g; let m;
  while ((m = re.exec(desc || '')) !== null) {
    const h = m[2];
    if (/\.(com|kr|net|org|co|io)$/i.test(h)) continue;                 // 이메일/도메인
    if (ownHandle && h.toLowerCase() === ownHandle.toLowerCase()) continue; // 채널 자신
    const p = people.find(x => x.handle && x.handle === h.toLowerCase());
    push(p ? p.name : '', p ? p.name : '@' + h);
  }
  return out;
}
/** 제목 괄호/with/feat */
function extractFromTitle(title, people) {
  const out = [];
  const segs = [];
  const re = /[\(\[（【]([^\)\]）】]+)[\)\]）】]/g; let m;
  while ((m = re.exec(title || '')) !== null) segs.push(m[1]);
  const w = /(?:with|feat\.?|ft\.?)\s*([^\)\]）】]+)/i.exec(title || '');
  if (w) segs.push(w[1]);
  segs.forEach(seg => {
    seg = seg.replace(/^\s*(with|feat\.?|ft\.?)\s*/i, '');
    seg.split(/[,&\/·+]|와\s|과\s/).map(x => x.trim()).filter(x => x.length >= 2 && x.length <= 20).forEach(cand => {
      const r = resolveName(cand, people);
      if (!out.some(o => o.raw === r.raw)) out.push(r);
    });
  });
  return out;
}
/** "2026년 9월 17일 방송분" → Date */
function parseBroadcastDate(desc) {
  const m = /(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*방송/.exec(desc || '');
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

// ─────────────────────────── 유튜브 ───────────────────────────
function fetchUploads(handle, start, end) {
  const ch = YouTube.Channels.list('contentDetails', { forHandle: handle });
  if (!ch.items || !ch.items.length) throw new Error('채널 핸들을 찾을 수 없음: @' + handle);
  const uploads = ch.items[0].contentDetails.relatedPlaylists.uploads;
  const vids = []; let token = null, stop = false;
  while (!stop) {
    const res = YouTube.PlaylistItems.list('snippet', { playlistId: uploads, maxResults: 50, pageToken: token || undefined });
    (res.items || []).forEach(it => {
      const d = new Date(it.snippet.publishedAt);
      if (d < start) { stop = true; return; }
      if (d < end) vids.push({ id: it.snippet.resourceId.videoId, title: it.snippet.title, desc: it.snippet.description || '', published: d });
    });
    token = res.nextPageToken; if (!token) stop = true;
  }
  // 쇼츠 제외: 길이 3분 이하 영상은 버림
  const dur = {};
  for (let i = 0; i < vids.length; i += 50) {
    const r = YouTube.Videos.list('contentDetails', { id: vids.slice(i, i + 50).map(v => v.id).join(',') });
    (r.items || []).forEach(it => { dur[it.id] = isoSeconds(it.contentDetails.duration); });
  }
  return vids.filter(v => (dur[v.id] || 0) > 180);
}
function isoSeconds(iso) {
  const m = /PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/.exec(iso || '');
  return m ? (Number(m[1] || 0) * 3600 + Number(m[2] || 0) * 60 + Number(m[3] || 0)) : 0;
}

// ─────────────────────────── 팀 탭 이사 (이 파일 → 팀 파일, 1회) ───────────────────────────
function migrateTeamTabs() {
  const ui = SpreadsheetApp.getUi();
  const ts = teamSS();
  if (ts.getId() === ss().getId()) { ui.alert('설정 탭 TEAM_SS_ID 가 비어 있어요. 새 팀 파일 ID를 먼저 넣고 setup 하세요.'); return; }
  let moved = 0;
  TEAM_TABS.forEach(tab => {
    const src = ss().getSheetByName(tab); if (!src || src.getLastRow() < 2) return;
    const dst = ts.getSheetByName(tab); if (!dst) return;
    const ks = cols(src), kd = cols(dst);
    const v = src.getRange(2, 1, src.getLastRow() - 1, src.getLastColumn()).getValues().filter(r => r[ks.NAME - 1]);
    if (!v.length) return;
    const nCol = Math.max(dst.getLastColumn(), TEAM_HEADERS.length);
    const rows = v.map(r => { const o = new Array(nCol).fill(''); Object.keys(HEADER_KEYS).forEach(key => { if (ks[key] && kd[key]) o[kd[key] - 1] = r[ks[key] - 1]; }); o[kd.YM - 1] = ymCell(r[ks.YM - 1]); return o; });
    const vals = dst.getRange(2, kd.YM, MAX_ROWS, 1).getValues();
    let r0 = 2; for (let i = vals.length - 1; i >= 0; i--) { if (vals[i][0] !== '') { r0 = i + 3; break; } }
    dst.getRange(r0, kd.YM, rows.length, 1).setNumberFormat('@');
    dst.getRange(r0, 1, rows.length, nCol).setValues(rows);
    applyRowFormulas(dst, r0, rows.length);
    moved += rows.length;
    src.setName(tab + '_백업'); src.hideSheet();
  });
  ui.alert(moved + '행을 팀 파일로 옮겼어요. 이 파일의 원래 입력 탭은 "_백업"으로 이름 바꿔 숨겨뒀어요 (확인 후 지워도 됨).');
}

// ─────────────────────────── 월 마감 (팀 탭 → 아카이브) ───────────────────────────
function archiveThisMonth() { archiveMonth(ymOf(new Date())); }
function archiveLastMonth() { archiveMonth(ymOf(addMonths(new Date(), -1))); }
/** 해당 월 행을 팀 탭에서 잘라내 HJ 파일의 입력_아카이브 탭에 붙임 (값으로 저장) */
function archiveMonth(ym) {
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch (e) { return null; } })();
  if (ui && ui.alert(ym + ' 행을 팀 탭에서 지우고 아카이브로 옮깁니다. 취합은 끝났나요?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return 0;
  const arc = getOrCreate('입력_아카이브');
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  let moved = 0;
  TEAM_TABS.forEach(tab => {
    const sh = teamSS().getSheetByName(tab); if (!sh) return;
    const k = cols(sh); const last = sh.getLastRow(); if (last < 2) return;
    const nCol = sh.getLastColumn();
    const v = sh.getRange(2, 1, last - 1, nCol).getValues();
    const take = [];
    for (let i = v.length - 1; i >= 0; i--) {
      if (ymCell(v[i][k.YM - 1]) !== ym || !v[i][k.NAME - 1]) continue;
      take.unshift([tab, stamp].concat(v[i].map((x, c) => c === k.YM - 1 ? ym : x)));
      sh.deleteRow(i + 2);
    }
    if (take.length) { arc.getRange(arc.getLastRow() + 1, 1, take.length, take[0].length).setValues(take); moved += take.length; }
  });
  const msg = ym + ' 마감: ' + moved + '행을 아카이브로 이동. 팀 탭은 비워졌어요.';
  if (ui) ui.alert(msg); else Logger.log(msg);
  return moved;
}

// ─────────────────────────── 취합 ───────────────────────────
function buildOutputThisMonth() { buildOutput(ymOf(new Date())); }
function buildOutputLastMonth() { buildOutput(ymOf(addMonths(new Date(), -1))); }

function buildOutput(ym) {
  const out = ss().getSheetByName('취합');
  const ppl = loadPeopleTypes();
  const newp = loadNewPeople();
  const [y, m] = ym.split('-').map(Number);
  const period = Utilities.formatDate(new Date(y, m - 1, 1), TZ, 'MMM-yy');
  const rows = [], warn = [], toAdd = {}, diag = [];
  TEAM_TABS.forEach(tab => {
    const sh = teamSS().getSheetByName(tab);
    if (!sh) { diag.push(tab + ': 탭 없음'); return; }
    const k = cols(sh);
    const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() : [];
    const r0 = v.find(r => r[k.NAME - 1]);
    diag.push(tab + ': 행 ' + v.length + ', 열(월=' + k.YM + ',게스트=' + k.NAME + ',✔=' + k.CHK + ',단가=' + k.UNIT + ',출연료=' + k.AMT + ')' + (r0 ? ', 첫행 월=' + JSON.stringify(r0[k.YM - 1]) + '→' + ymCell(r0[k.YM - 1]) + ', ✔=' + JSON.stringify(r0[k.CHK - 1]) + ', 출연료=' + JSON.stringify(r0[k.AMT - 1]) : ', 데이터 없음'));
    v.forEach((r, i) => {
      if (ymCell(r[k.YM - 1]) !== ym || r[k.CHK - 1] !== true) return;
      const name = String(r[k.NAME - 1] || '').trim();
      const amt = Number(r[k.AMT - 1]) || 0;
      const note = String(r[k.NOTE - 1] || '').trim() || (String(r[k.CONTENT - 1] || '').trim() + (r[k.DATE - 1] ? ' (' + fmt(r[k.DATE - 1]) + ')' : ''));
      const p = ppl[name];
      const np = p ? null : findNewPerson(name, newp);
      const type = p ? p.type : (np ? np.type : '');
      const mult = type === '세금계산서' ? 1.1 : type === '삼쩜삼' ? 0.967 : 1;
      const flag = p ? '' : np ? '신규(계약서)' : '신규(정보없음)';
      if (!name) warn.push(tab + ' ' + (i + 2) + '행: 게스트명 없음');
      if (!r[k.BASIS - 1]) warn.push(tab + ' ' + (i + 2) + '행 (' + name + '): 단가기준 없음');
      if (!amt && ZERO_BASIS.indexOf(r[k.BASIS - 1]) < 0) warn.push(tab + ' ' + (i + 2) + '행 (' + name + '): 금액 없음');
      if (!p && !np) warn.push(tab + ' ' + (i + 2) + '행 (' + name + '): 인물마스터·신규인물정보 둘 다 없음 → 계약서/계좌 확인 필요');
      if (!p && np) { toAdd[name] = np; if (np.missing) warn.push(name + ': 계약서 추출본에 빈 항목(' + np.missing + ') → 모두싸인에서 확인'); }
      if (amt === 0) return;
      rows.push(['제작인건비', '출연료', name, type, period, -amt, mult, -Math.round(amt * mult), note, tab, flag]);
    });
  });
  out.clearContents();
  out.getRange(1, 1, 1, OUT_HEADERS.length).setValues([OUT_HEADERS]); styleHeader(out, OUT_HEADERS.length);
  if (rows.length) {
    out.getRange(2, 1, rows.length, OUT_HEADERS.length).setValues(rows);
    out.getRange(2, 6, rows.length, 1).setNumberFormat('#,##0'); out.getRange(2, 8, rows.length, 1).setNumberFormat('#,##0');
  }
  // 오른쪽: 통합문서 인물정보에 붙일 신규 행 (M열부터)
  const c0 = OUT_HEADERS.length + 2;
  out.getRange(1, c0, 1, OUT_NEWP_HEADERS.length).setValues([OUT_NEWP_HEADERS]).setBackground('#7c2d12').setFontColor('#fff').setFontWeight('bold');
  const addRows = Object.keys(toAdd).map(n => { const q = toAdd[n]; return ['', n, q.type, q.num, q.payee, q.bank, q.acct, '게스트']; });
  if (addRows.length) out.getRange(2, c0, addRows.length, OUT_NEWP_HEADERS.length).setValues(addRows);
  out.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(MISSING).setFontColor('#dc2626').setBold(true).setRanges([out.getRange(2, c0, MAX_ROWS, OUT_NEWP_HEADERS.length)]).build()]);
  const total = rows.reduce((s, r) => s + r[7], 0);
  const summary = ym + ' 취합: ' + rows.length + '행, 지급 합계 ' + Math.abs(total).toLocaleString() + '원'
    + (addRows.length ? '\n인물정보 추가 필요: ' + addRows.length + '명 (취합 탭 오른쪽 블록 복사)' : '')
    + (warn.length ? '\n\n⚠ 확인 필요:\n- ' + warn.join('\n- ') : '')
    + (rows.length === 0 ? '\n\n🔍 진단 (대상 월 ' + ym + '):\n' + diag.join('\n') : '');
  const cfg = getSettings();
  if (cfg.OWNER_EMAIL) MailApp.sendEmail(cfg.OWNER_EMAIL, '[출연료] ' + ym + ' 취합 완료', summary + '\n\n' + ss().getUrl() + (teamSS().getId() !== ss().getId() ? '\n팀 파일: ' + teamSS().getUrl() : ''));
  try { SpreadsheetApp.getUi().alert(summary); } catch (e) { Logger.log(summary); }
  return rows.length;
}
/** 신규인물정보 탭 → {이름: {type,num,payee,bank,acct,missing,names[]}} (최신 행 우선) */
function loadNewPeople() {
  const sh = ss().getSheetByName('신규인물정보'); if (!sh) return [];
  const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues() : [];
  return v.filter(r => r[1]).map(r => ({
    names: String(r[1]).split(/[,&\/]/).map(x => x.trim()).filter(Boolean).concat(r[11] ? [String(r[11]).trim()] : []),
    type: String(r[2] || ''), num: String(r[3] || ''), payee: String(r[4] || ''), bank: String(r[5] || ''), acct: String(r[6] || ''), missing: String(r[13] || '') }));
}
function findNewPerson(name, newp) {
  const base = name.replace(/\s*(공식|개인|출연)$/, '').trim();
  return newp.find(q => q.names.some(n => n === base || n === name || base.indexOf(n) >= 0 || n.indexOf(base) >= 0)) || null;
}
function loadPeopleTypes() {
  const sh = ss().getSheetByName('인물마스터');
  const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  const o = {}; v.forEach(r => { if (r[0]) o[String(r[0]).trim()] = { type: String(r[1] || '').trim() }; });
  return o;
}

// ─────────────────────────── 알림 ───────────────────────────
function notifyCollect() { notify('collect', ymOf(new Date())); }
function notifyRemind() { const ym = ymOf(addMonths(new Date(), -1)); try { importNewPeople(); prefillFromContracts(ym); } catch (e) { Logger.log(e.message); } notify('remind', ym); }

function notify(kind, ym) {
  const cfg = getSettings();
  const url = teamSS().getUrl();
  const counts = TEAM_TABS.map(tab => {
    const sh = teamSS().getSheetByName(tab); const k = cols(sh);
    const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() : [];
    const mine = v.filter(r => ymCell(r[k.YM - 1]) === ym);
    const todo = mine.filter(r => r[k.CHK - 1] !== true || r[k.UNIT - 1] === '').length;
    const owner = cfg['담당_' + tab] ? ' @' + cfg['담당_' + tab] : '';
    return tab.replace('입력_', '') + owner + ' ' + mine.length + '건 (미완료 ' + todo + ')';
  }).join('\n· ');
  const [y, m] = ym.split('-').map(Number);
  const dueMonth = m === 12 ? 1 : m + 1;
  const tpl = (kind === 'collect' ? cfg.SLACK_MSG_COLLECT : cfg.SLACK_MSG_REMIND) || (kind === 'collect'
    ? '📋 *{년}년 {월}월 출연료 정리* 시작합니다! 각 팀 탭에서 ✔확인·단가기준·단가·비고 입력 부탁드려요.\n마감: *{마감월}월 1일* (5일 지급)\n{현황}\n{링크}'
    : '⏰ *{월}월 출연료 입력 마감일*입니다! 미완료 건 오늘 중 입력 부탁드려요.\n{현황}\n{링크}');
  const text = tpl.replace(/\\n/g, '\n').replace(/\{년\}/g, y).replace(/\{월\}/g, m).replace(/\{마감월\}/g, dueMonth)
    .replace(/\{현황\}/g, '· ' + counts).replace(/\{링크\}/g, url);
  let result;
  if (!cfg.SLACK_WEBHOOK) {
    result = '⚠ 설정 탭 SLACK_WEBHOOK이 비어 있어 슬랙으로 못 보냈어요.\n\n보낼 예정이던 메시지:\n' + text;
  } else {
    try {
      const res = UrlFetchApp.fetch(cfg.SLACK_WEBHOOK, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ text }), muteHttpExceptions: true });
      result = res.getResponseCode() === 200 ? '슬랙 발송 완료 ✔\n\n' + text : '⚠ 슬랙 응답 ' + res.getResponseCode() + ': ' + res.getContentText() + '\n(웹훅 URL이 맞는지, 채널이 살아있는지 확인)';
    } catch (e) { result = '⚠ 슬랙 전송 실패: ' + e.message; }
  }
  try { SpreadsheetApp.getUi().alert(result); } catch (e) { Logger.log(result); }
  return text;
}

// ─────────────────────────── 트리거 ───────────────────────────
function installTriggers() {
  const cfg = getSettings();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('monthlyCollect').timeBased().onMonthDay(Number(cfg.COLLECT_DAY) || 25).atHour(9).inTimezone(TZ).create();
  ScriptApp.newTrigger('notifyRemind').timeBased().onMonthDay(Number(cfg.REMIND_DAY) || 1).atHour(10).inTimezone(TZ).create();
  ScriptApp.newTrigger('syncPeople').timeBased().everyDays(1).atHour(6).inTimezone(TZ).create();
  try { SpreadsheetApp.getUi().alert('트리거 설치 완료: 매월 ' + (cfg.COLLECT_DAY || 25) + '일 수집+안내, 매월 ' + (cfg.REMIND_DAY || 1) + '일 리마인드, 매일 06시 인물정보 동기화'); } catch (e) {}
}
function monthlyCollect() {
  try { syncPeople(); } catch (e) { Logger.log('syncPeople 실패: ' + e.message); }
  const ym = ymOf(new Date());
  collectMonth(ym);
  notify('collect', ym);
}

// ─────────────────────────── 유틸 ───────────────────────────
function ymOf(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM'); }
/** 셀의 월 값: 텍스트 '2026-10' / 날짜 / 숫자(시리얼) 모두 'yyyy-MM'로 */
function ymCell(v) {
  if (v instanceof Date) return ymOf(v);
  if (typeof v === 'number' && v > 30000) return ymOf(new Date(Date.UTC(1899, 11, 30) + v * 86400000));
  if (/^\d{5}$/.test(String(v).trim())) return ymOf(new Date(Date.UTC(1899, 11, 30) + Number(v) * 86400000));
  const m = /^(\d{4})[-./](\d{1,2})/.exec(String(v || '').trim());
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) : String(v || '').trim();
}
function addMonths(d, n) { const x = new Date(d); x.setMonth(x.getMonth() + n); return x; }
function fmt(d) { return d instanceof Date ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : (d || ''); }
function clearMonthRows(sh, ym, col) {
  const last = sh.getLastRow(); if (last < 2) return;
  const v = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === ym) sh.deleteRow(i + 2);
}
