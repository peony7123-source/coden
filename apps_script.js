// =====================================================
//  코드앤 코딩 스튜디오 — 학생 진도 저장 스크립트
//  Google 스프레드시트 > 확장 프로그램 > Apps Script
//  에 이 코드 전체를 붙여넣고 배포하세요.
// =====================================================

function doGet(e) {
  const action = e.parameter.action || 'read';
  const name   = decodeURIComponent(e.parameter.name || '');

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName('학생XP');

  // 학생XP 시트가 없으면 자동 생성
  if (!sheet) {
    sheet = ss.insertSheet('학생XP');
    sheet.appendRow(['학생이름', '엔트리XP', '엔트리미션', '스크래치XP', '스크래치미션', '최종업데이트']);
    sheet.setFrozenRows(1);
  }

  // 관리자설정 시트가 없으면 자동 생성
  if (!ss.getSheetByName('관리자설정')) {
    const adminSheet = ss.insertSheet('관리자설정');
    adminSheet.appendRow(['type', 'images', 'memos', 'links', 'missions', '업데이트']);
    adminSheet.setFrozenRows(1);
  }

  // ── 읽기 ──────────────────────────────────────────
  if (action === 'read') {
    if (!name) return respond({ error: 'no name' });

    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === name) {
        return respond({
          found:          true,
          entryXP:        rows[i][1] || 0,
          entryMissions:  rows[i][2] || '{}',
          scratchXP:      rows[i][3] || 0,
          scratchMissions:rows[i][4] || '{}'
        });
      }
    }
    // 처음 접속하는 학생
    return respond({ found: false, entryXP: 0, entryMissions: '{}', scratchXP: 0, scratchMissions: '{}' });
  }

  // ── 저장 ──────────────────────────────────────────
  if (action === 'save') {
    const type     = e.parameter.type || 'entry';
    const xp       = parseInt(e.parameter.xp)  || 0;
    const missions = decodeURIComponent(e.parameter.missions || '{}');
    const now      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    const rows  = sheet.getDataRange().getValues();
    let   found = false;

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === name) {
        if (type === 'entry') {
          sheet.getRange(i + 1, 2).setValue(xp);
          sheet.getRange(i + 1, 3).setValue(missions);
        } else {
          sheet.getRange(i + 1, 4).setValue(xp);
          sheet.getRange(i + 1, 5).setValue(missions);
        }
        sheet.getRange(i + 1, 6).setValue(now);
        found = true;
        break;
      }
    }

    if (!found) {
      // 새 학생 행 추가
      if (type === 'entry') {
        sheet.appendRow([name, xp, missions, 0, '{}', now]);
      } else {
        sheet.appendRow([name, 0, '{}', xp, missions, now]);
      }
    }

    return respond({ ok: true });
  }

  // ── 관리자 설정 읽기 ──────────────────────────────
  if (action === 'readAdmin') {
    const type = e.parameter.type || 'entry';
    const adminSheet = ss.getSheetByName('관리자설정');
    if (!adminSheet) return respond({ found: false });

    const rows = adminSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === type) {
        return respond({
          found:    true,
          images:   rows[i][1] || '{}',
          memos:    rows[i][2] || '{}',
          links:    rows[i][3] || '{}',
          missions: rows[i][4] || '{}'
        });
      }
    }
    return respond({ found: false });
  }

  // ── 관리자 설정 저장 (GET) ────────────────────────
  if (action === 'saveAdmin') {
    const type     = e.parameter.type     || 'entry';
    const images   = e.parameter.images   || '{}';
    const memos    = e.parameter.memos    || '{}';
    const links    = e.parameter.links    || '{}';
    const missions = e.parameter.missions || '{}';
    const now      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    let adminSheet = ss.getSheetByName('관리자설정');
    if (!adminSheet) {
      adminSheet = ss.insertSheet('관리자설정');
      adminSheet.appendRow(['type', 'images', 'memos', 'links', 'missions', '업데이트']);
      adminSheet.setFrozenRows(1);
    }

    const rows = adminSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === type) {
        adminSheet.getRange(i + 1, 2, 1, 5).setValues([[images, memos, links, missions, now]]);
        found = true;
        break;
      }
    }
    if (!found) {
      adminSheet.appendRow([type, images, memos, links, missions, now]);
    }
    return respond({ ok: true });
  }

  // ── 갤러리 → 미션 시트 동기화 ────────────────────
  if (action === 'writeMissionsToSheet') {
    const type = e.parameter.type || 'entry';
    const missionsJson = decodeURIComponent(e.parameter.missions || '{}');
    const sheetName = type === 'entry' ? '엔트리미션' : '스크래치미션';

    let missions;
    try { missions = JSON.parse(missionsJson); } catch(err) { return respond({ error: 'invalid json' }); }

    let missionSheet = ss.getSheetByName(sheetName);
    if (!missionSheet) {
      missionSheet = ss.insertSheet(sheetName);
      missionSheet.appendRow(['제목/설명 (A열)', '단계별 내용 (B열)']);
      missionSheet.setFrozenRows(1);
    }

    const ids = Object.keys(missions).map(Number).sort((a, b) => a - b);
    let updated = 0;

    for (const id of ids) {
      const m = missions[id];
      if (!m) continue;
      const row = id + 1; // 1행=헤더, 2행=미션1번

      // 행이 부족하면 빈 행 추가
      while (missionSheet.getLastRow() < row) {
        missionSheet.appendRow(['', '']);
      }

      // A열: 제목\n설명
      let colA = m.title || '';
      if (m.desc) colA += '\n' + m.desc;

      // B열: 1단계 — 단계명\n설명\n2단계 — ...
      let colB = '';
      if (m.steps && m.steps.length > 0) {
        colB = m.steps.map((s, i) => {
          let line = (i + 1) + '단계 — ' + (s.title || '');
          if (s.desc) line += '\n' + s.desc;
          return line;
        }).join('\n');
      }

      missionSheet.getRange(row, 1).setValue(colA);
      missionSheet.getRange(row, 2).setValue(colB);
      updated++;
    }

    return respond({ ok: true, updated });
  }

  // ── 커스텀 과목 읽기 ──────────────────────────────
  if (action === 'readSubjects') {
    let adminSheet = ss.getSheetByName('관리자설정');
    if (!adminSheet) return respond({ found: false, subjects: '[]' });

    const rows = adminSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === 'subjects') {
        return respond({ found: true, subjects: rows[i][1] || '[]' });
      }
    }
    return respond({ found: false, subjects: '[]' });
  }

  // ── 커스텀 과목 저장 ──────────────────────────────
  if (action === 'saveSubjects') {
    const subjects = decodeURIComponent(e.parameter.subjects || '[]');
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    let adminSheet = ss.getSheetByName('관리자설정');
    if (!adminSheet) {
      adminSheet = ss.insertSheet('관리자설정');
      adminSheet.appendRow(['type', 'data', 'memos', 'links', 'missions', '업데이트']);
      adminSheet.setFrozenRows(1);
    }

    const rows = adminSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === 'subjects') {
        adminSheet.getRange(i + 1, 2).setValue(subjects);
        adminSheet.getRange(i + 1, 6).setValue(now);
        found = true;
        break;
      }
    }
    if (!found) {
      adminSheet.appendRow(['subjects', subjects, '', '', '', now]);
    }
    return respond({ ok: true });
  }

  // ── 학생 메모 저장 ───────────────────────────────
  if (action === 'saveMemo') {
    const type        = e.parameter.type        || 'entry';
    const missionId   = parseInt(e.parameter.missionId) || 0;
    const missionTitle= decodeURIComponent(e.parameter.missionTitle || '');
    const memo        = decodeURIComponent(e.parameter.memo || '');
    const now         = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    let memoSheet = ss.getSheetByName('학생메모');
    if (!memoSheet) {
      memoSheet = ss.insertSheet('학생메모');
      memoSheet.appendRow(['학생이름', '과목', '미션번호', '미션제목', '메모내용', '저장시간']);
      memoSheet.setFrozenRows(1);
    }

    // 기존 행 찾아서 업데이트, 없으면 새 행 추가
    const rows = memoSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === name && rows[i][1] === type && rows[i][2] === missionId) {
        memoSheet.getRange(i + 1, 5).setValue(memo);
        memoSheet.getRange(i + 1, 6).setValue(now);
        found = true;
        break;
      }
    }
    if (!found) {
      memoSheet.appendRow([name, type, missionId, missionTitle, memo, now]);
    }
    return respond({ ok: true });
  }

  return respond({ error: 'unknown action' });
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
//  시트3 → 갤러리 미션 단계 동기화
//  스프레드시트 메뉴 [코딩학원] > [미션 단계 동기화] 로 실행
// =====================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎓 코딩학원')
    .addItem('🔄 엔트리 미션 단계 동기화 (엔트리미션)', 'syncEntryMissionSteps')
    .addItem('🔗 엔트리 YouTube 링크 동기화 (엔트리영상)', 'syncEntryLinks')
    .addSeparator()
    .addItem('🔄 스크래치 미션 단계 동기화 (스크래치미션)', 'syncScratchMissionSteps')
    .addItem('🔗 스크래치 YouTube 링크 동기화 (스크래치영상)', 'syncScratchLinks')
    .addToUi();
}

// 시트3 → 관리자설정 entry 행의 missions 컬럼 업데이트
function syncEntryMissionSteps() {
  _syncMissionSteps('엔트리미션', 'entry', '엔트리');
}

// 엔트리영상 시트 (A=id, B=url) → 관리자설정 entry 행의 links 컬럼 업데이트
function syncEntryLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('엔트리영상');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('엔트리 시트를 찾을 수 없어요!');
    return;
  }
  const lastRow = sheet.getLastRow();
  const entryLinks = {};
  for (let row = 2; row <= lastRow; row++) {
    const id = sheet.getRange(row, 1).getValue();
    const url = (sheet.getRange(row, 2).getValue() || '').toString().trim();
    if (id && url) entryLinks[String(id)] = url;
  }
  _saveLinksToAdmin('entry', entryLinks);
  SpreadsheetApp.getUi().alert('엔트리 YouTube 링크 동기화 완료!\n총 ' + Object.keys(entryLinks).length + '개 링크가 업데이트 됐어요.\n\n갤러리에서 [🔄 시트 동기화] 버튼을 눌러 반영하세요.');
}

// 스크래치미션 시트 → 관리자설정 scratch 행의 missions 컬럼 업데이트
function syncScratchMissionSteps() {
  _syncMissionSteps('스크래치미션', 'scratch', '스크래치');
}

// 스크래치영상 시트 (A=id, B=url) → 관리자설정 scratch 행의 links 컬럼 업데이트
function syncScratchLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('스크래치영상');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('스크래치영상 시트를 찾을 수 없어요!\n시트 이름을 확인해주세요.');
    return;
  }
  const lastRow = sheet.getLastRow();
  const links = {};
  for (let row = 2; row <= lastRow; row++) {
    const id = sheet.getRange(row, 1).getValue();
    const url = (sheet.getRange(row, 2).getValue() || '').toString().trim();
    if (id && url) links[String(id)] = url;
  }
  _saveLinksToAdmin('scratch', links);
  SpreadsheetApp.getUi().alert('스크래치 YouTube 링크 동기화 완료!\n총 ' + Object.keys(links).length + '개 링크가 업데이트 됐어요.\n\n갤러리에서 [🔄 시트 동기화] 버튼을 눌러 반영하세요.');
}

function _syncMissionSteps(sheetName, adminType, label) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(sheetName + ' 시트를 찾을 수 없어요!\n시트 이름을 확인해주세요.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('데이터가 없어요. 1행은 헤더, 2행부터 미션 데이터를 입력해주세요.');
    return;
  }

  // 기존 관리자설정에서 현재 customMissions 불러오기 (덮어쓰지 않고 병합)
  const existingMissions = _loadExistingMissions(adminType);
  const customMissions = Object.assign({}, existingMissions);

  let count = 0;
  for (let row = 2; row <= lastRow; row++) {
    const missionId = row - 1; // 2행 → 미션 1번
    const colA = (sheet.getRange(row, 1).getValue() || '').toString().trim();
    const colB = (sheet.getRange(row, 2).getValue() || '').toString().trim();

    if (!colA && !colB) continue; // 빈 행 건너뜀

    if (!customMissions[missionId]) customMissions[missionId] = {};

    // A열: 첫 줄 = 제목, 나머지 = 설명
    if (colA) {
      const aLines = colA.split('\n').map(l => l.trim()).filter(Boolean);
      if (aLines[0]) customMissions[missionId].title = aLines[0];
      if (aLines.length > 1) customMissions[missionId].desc = aLines.slice(1).join(' ');
    }

    // B열: "N단계 — emoji 단계명\n설명" 형식 파싱
    if (colB) {
      const steps = _parseSteps(colB);
      if (steps.length > 0) {
        customMissions[missionId].steps = steps;
        count++;
      }
    }
  }

  // 관리자설정에 저장
  _saveMissionsToAdmin(adminType, customMissions);

  // 첫 번째 미션 파싱 결과 미리보기 (디버그용)
  const firstId = Object.keys(customMissions)[0];
  const firstSteps = firstId && customMissions[firstId].steps;
  const rawB2 = lastRow >= 2 ? (sheet.getRange(2, 2).getValue() || '').toString().substring(0, 80) : '(없음)';
  const preview = firstSteps && firstSteps.length > 0
    ? '\n\n✅ [미리보기] 미션' + firstId + ' 1단계: ' + firstSteps[0].title
    : '\n\n⚠️ 단계 파싱 실패\nB2 셀 내용: [' + rawB2 + ']\n형식 예시: 1단계 — 🎨 배경 고르기';

  SpreadsheetApp.getUi().alert(
    label + ' 미션 단계 동기화 완료!\n' +
    '총 ' + count + '개 미션의 단계가 업데이트 됐어요.' +
    preview + '\n\n갤러리에서 [🔄 시트 동기화] 버튼을 클릭하면 반영돼요.'
  );
}

// "N단계 — emoji 단계명\n설명" 형식 텍스트 파싱 → [{title, desc}] 배열
function _parseSteps(text) {
  const lines = text.split('\n');
  const steps = [];
  let currentTitle = null;
  let descLines = [];

  for (const raw of lines) {
    const line = raw.trim();
    // "1단계 — 제목" 패턴: em dash(—), en dash(–), hyphen(-), 화살표(→), 공백 등 모두 허용
    const match = line.match(/^\d+단계\s*[—–\-→]\s*(.+)$/) ||
                  line.match(/^\d+단계\s+(.+)$/);
    if (match) {
      if (currentTitle !== null) {
        steps.push({ title: currentTitle, desc: descLines.join(' ').trim() });
      }
      currentTitle = match[1].trim();
      descLines = [];
    } else if (line && currentTitle !== null) {
      descLines.push(line);
    }
  }
  if (currentTitle !== null) {
    steps.push({ title: currentTitle, desc: descLines.join(' ').trim() });
  }
  return steps;
}

function _loadExistingMissions(type) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const adminSheet = ss.getSheetByName('관리자설정');
    if (!adminSheet) return {};
    const rows = adminSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === type) {
        const missionsJson = rows[i][4] || '{}';
        return JSON.parse(missionsJson);
      }
    }
  } catch (e) {}
  return {};
}

function _saveMissionsToAdmin(type, missionsObj) {
  _saveColToAdmin(type, 5, JSON.stringify(missionsObj)); // col 5 = missions
}

function _saveLinksToAdmin(type, linksObj) {
  _saveColToAdmin(type, 4, JSON.stringify(linksObj)); // col 4 = links
}

function _saveColToAdmin(type, col, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let adminSheet = ss.getSheetByName('관리자설정');
  if (!adminSheet) {
    adminSheet = ss.insertSheet('관리자설정');
    adminSheet.appendRow(['type', 'images', 'memos', 'links', 'missions', '업데이트']);
    adminSheet.setFrozenRows(1);
  }
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const rows = adminSheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === type) {
      adminSheet.getRange(i + 1, col).setValue(value);
      adminSheet.getRange(i + 1, 6).setValue(now);
      found = true;
      break;
    }
  }
  if (!found) {
    const newRow = [type, '{}', '{}', '{}', '{}', now];
    newRow[col - 1] = value;
    adminSheet.appendRow(newRow);
  }
}
