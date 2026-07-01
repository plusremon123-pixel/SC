// ─── 상수 ───────────────────────────────────────────────
const LS_KEY_DATA   = 'curriculum_data';
const LS_KEY_FILES  = 'curriculum_filenames';
const LS_KEY_APIKEY = 'openrouter_api_key';
const LS_UPLOAD_MARKER = 'manual-upload-v3';
const UPLOAD_AUTH_KEY = 'schedule-card-upload-auth-v1';
const UPLOAD_PASSWORD = '610503';
const SUPABASE_URL = 'https://vmebzlinboxmgcrrorwv.supabase.co';
const SUPABASE_KEY = 'sb_publishable_zpANEcZ0GfP44NpyHgZECQ_LPxB1LhR';
const SUPABASE_BUCKET = 'schedule-data';
const SUPABASE_SCHEDULE_CARD_PATH = 'schedule-card-data.json';

const TERM_CONFIG = {
  regular: {
    label: '정규학기',
    dataKey: LS_KEY_DATA,
    filesKey: LS_KEY_FILES,
    markerKey: 'curriculum_data_upload_marker',
  },
  vacation: {
    label: '방학학기',
    dataKey: 'curriculum_data_vacation',
    filesKey: 'curriculum_filenames_vacation',
    markerKey: 'curriculum_data_vacation_upload_marker',
  },
};

// ─── API Key 로드 (config.js → localStorage 자동 저장) ──
function loadApiKey() {
  const configKey = (window.CONFIG && CONFIG.API_KEY && !CONFIG.API_KEY.includes('여기에'))
    ? CONFIG.API_KEY : null;
  if (configKey) {
    // config.js에 유효한 키가 있으면 localStorage에 덮어씀
    localStorage.setItem(LS_KEY_APIKEY, configKey);
    return configKey;
  }
  // config.js 키가 없으면 localStorage에서 복원
  return localStorage.getItem(LS_KEY_APIKEY) || '';
}

// 과목 약어 → 전체명 (약어, 전체명 모두 허용)
const SUBJECT_MAP = {
  '수': '수학',   '국': '국어',   '바': '바슬즐',
  '영': '영어',   '과': '과학',   '사': '사회',
  '수학': '수학', '국어': '국어', '바슬즐': '바슬즐',
  '영어': '영어', '과학': '과학', '사회': '사회'
};
const DAYS = ['월', '화', '수', '목', '금'];

// ─── DOM 참조 (단일 검색) ────────────────────────────────
const fileStatusEl   = document.getElementById('file-status');
const fileStatusText = document.getElementById('file-status-text');
const csvInput       = document.getElementById('csv-input');
const clearUploadBtn = document.getElementById('clear-upload-btn');
const termButtons    = Array.from(document.querySelectorAll('[data-term]'));
const termSwitchNote = document.getElementById('term-switch-note');
const gradeSelect    = document.getElementById('grade-select');
const subjectSelect  = document.getElementById('subject-select');
const lessonInput        = document.getElementById('lesson-input');
const lessonSelectHelper = document.getElementById('lesson-select-helper');
const uniqueidInput    = document.getElementById('uniqueid-input');
const expdateInput     = document.getElementById('expdate-input');
const quickSearchInput = document.getElementById('quick-search-input');
const extraGradeSelect = document.getElementById('extra-grade-select');
const unitnameInput  = document.getElementById('unitname-input');
const analyzeBtn     = document.getElementById('analyze-btn');
const resultSection  = document.getElementById('result-section');
const resultContent  = document.getElementById('result-content');

// ─── DOM 참조 (복수 검색) ────────────────────────────────
const bulkGradeSelect   = document.getElementById('bulk-grade');
const bulkImageInput    = document.getElementById('bulk-image');
const bulkImageName     = document.getElementById('bulk-image-name');
const bulkPreviewWrap   = document.getElementById('bulk-preview-wrap');
const bulkPreviewImg    = document.getElementById('bulk-preview');
const bulkAnalyzeBtn    = document.getElementById('bulk-analyze-btn');
const bulkEditSection   = document.getElementById('bulk-edit-section');
const bulkSearchBtn     = document.getElementById('bulk-search-btn');
const bulkResultSection = document.getElementById('bulk-result-section');
const weekGrid          = document.getElementById('week-grid');
const pasteZone         = document.getElementById('paste-zone');
const EDIT_INPUTS = {
  '월': document.getElementById('edit-mon'),
  '화': document.getElementById('edit-tue'),
  '수': document.getElementById('edit-wed'),
  '목': document.getElementById('edit-thu'),
  '금': document.getElementById('edit-fri'),
};

// ─── 데이터 파일 섹션 토글 ───────────────────────────────
(function initFileToggle() {
  const toggle = document.getElementById('file-section-toggle');
  const body   = document.getElementById('file-section-body');
  if (!toggle || !body) return;
  toggle.addEventListener('click', function () {
    const expanded = this.getAttribute('aria-expanded') === 'true';
    this.setAttribute('aria-expanded', String(!expanded));
    body.hidden = expanded;
  });
})();

// ─── 데이터 상태 ─────────────────────────────────────────
let rows = [];
let activeTerm = 'regular';
const termRows = { regular: [], vacation: [] };
const termFiles = { regular: [], vacation: [] };

// ─── 초기화 ──────────────────────────────────────────────
(async function init() {
  if (!TERM_CONFIG[activeTerm]) activeTerm = 'regular';
  loadDataJs();
  loadSavedUploadedData();
  await loadSharedScheduleCardData();
  updateTermSwitchUI();
  switchTerm(activeTerm, true);
})().catch(error => {
  console.error(error);
  loadDataJs();
  loadSavedUploadedData();
  updateTermSwitchUI();
  switchTerm(activeTerm, true);
});

termButtons.forEach(btn => {
  btn.addEventListener('click', function () {
    switchTerm(this.dataset.term);
  });
});

if (csvInput) {
  csvInput.addEventListener('change', async function () {
    const files = Array.from(this.files || []);
    if (!files.length) return;
    if (!confirmUploadPassword()) {
      alert('비밀번호가 맞지 않아 업로드를 취소했습니다.');
      this.value = '';
      return;
    }
    const label = TERM_CONFIG[activeTerm].label;
    if (rows.length > 0 && !confirm(`${label} 데이터를 업로드한 CSV(${files.length}개)로 임시 교체하시겠습니까?`)) {
      this.value = '';
      return;
    }

    let pending = files.length;
    const allRows = [];
    const fileNames = [];
    files.forEach(file => {
      fileNames.push(file.name);
      const reader = new FileReader();
      reader.onload = e => {
        const result = Papa.parse(e.target.result, { header: true, skipEmptyLines: true });
        allRows.push(...processRows(result.data));
        if (--pending === 0) {
          if (!allRows.length) {
            alert('유효한 데이터가 없습니다.');
            return;
          }
          termRows[activeTerm] = allRows;
          termFiles[activeTerm] = fileNames;
          rows = allRows;
          localStorage.setItem(TERM_CONFIG[activeTerm].dataKey, JSON.stringify(allRows));
          localStorage.setItem(TERM_CONFIG[activeTerm].filesKey, JSON.stringify(fileNames));
          localStorage.setItem(TERM_CONFIG[activeTerm].markerKey, LS_UPLOAD_MARKER);
          refreshDataControls(false);
          saveSharedScheduleCardData(activeTerm, allRows, fileNames)
            .then(() => alert('스케쥴 카드 공유 데이터가 저장되었습니다.'))
            .catch(error => {
              console.error(error);
              alert(error.message ? `화면에는 반영됐지만 공유 저장 실패: ${error.message}` : '화면에는 반영됐지만 공유 저장에 실패했습니다.');
            });
        }
      };
      reader.readAsText(file, 'EUC-KR');
    });
    this.value = '';
  });
}

if (clearUploadBtn) {
  clearUploadBtn.addEventListener('click', function () {
    const cfg = TERM_CONFIG[activeTerm];
    localStorage.removeItem(cfg.dataKey);
    localStorage.removeItem(cfg.filesKey);
    localStorage.removeItem(cfg.markerKey);
    loadDataJs();
    switchTerm(activeTerm);
  });
}

// ─── 탭 전환 ─────────────────────────────────────────────
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', function () {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    this.classList.add('active');
    document.querySelectorAll('.tab-content').forEach(c => { c.hidden = true; });
    document.getElementById('tab-' + this.dataset.tab).hidden = false;
  });
});

// ─── 상호 초기화 헬퍼 ────────────────────────────────────
function clearRow1() { quickSearchInput.value = ''; }

function clearRow2() {
  gradeSelect.value = '';
  subjectSelect.innerHTML = '<option value="">-- 과목 선택 --</option>';
  subjectSelect.disabled = true;
  lessonSelectHelper.innerHTML = '<option value="">-- 목록에서 선택 --</option>';
  lessonSelectHelper.disabled = true;
  lessonInput.value = ''; lessonInput.disabled = true;
}

function clearRow3() {
  extraGradeSelect.value = '';
  uniqueidInput.value = ''; uniqueidInput.disabled = true;
  expdateInput.value = ''; expdateInput.disabled = true;
  unitnameInput.value = ''; unitnameInput.disabled = true;
}

// ─── 단일 검색 - 이벤트 ──────────────────────────────────
// Row 1: 빠른 입력
function handleQuickSearchChange() {
  if (this.value.trim()) { clearRow2(); clearRow3(); }
  updateAnalyzeBtn();
}

['input', 'change', 'keyup', 'compositionend'].forEach(eventName => {
  quickSearchInput.addEventListener(eventName, handleQuickSearchChange);
});
quickSearchInput.addEventListener('paste', () => setTimeout(updateAnalyzeBtn, 0));
quickSearchInput.addEventListener('keydown', e => { if (e.key === 'Enter') analyzeBtn.click(); });

// Row 2: 선택 입력
gradeSelect.addEventListener('change', function () {
  if (this.value) { clearRow1(); clearRow3(); }
  resetSubjectAndBelow();
  if (this.value) {
    populateSubjects(this.value);
    subjectSelect.disabled = false;
  }
  updateAnalyzeBtn();
});

subjectSelect.addEventListener('change', function () {
  lessonSelectHelper.innerHTML = '<option value="">-- 목록에서 선택 --</option>';
  lessonSelectHelper.disabled = true;
  lessonInput.value = ''; lessonInput.disabled = true;
  if (this.value) {
    populateLessons(gradeSelect.value, this.value);
    lessonSelectHelper.disabled = false;
    lessonInput.disabled = false;
  }
  updateAnalyzeBtn();
});

// Row 3: 추가 검색
extraGradeSelect.addEventListener('change', function () {
  if (this.value) {
    clearRow1(); clearRow2();
    uniqueidInput.disabled = false;
    expdateInput.disabled = false;
    unitnameInput.disabled = false;
  } else {
    uniqueidInput.disabled = true;
    expdateInput.disabled = true;
    unitnameInput.disabled = true;
  }
  updateAnalyzeBtn();
});

// ─── 단일 검색 - 서브 조건 이벤트 ───────────────────────
// 목록에서 선택 → text input에 자동 반영
lessonSelectHelper.addEventListener('change', function () {
  if (this.value) {
    lessonInput.value = this.value;
    updateAnalyzeBtn();
  }
});

// lesson 직접입력 → row 1·3 초기화
lessonInput.addEventListener('input', function () {
  if (this.value.trim()) { clearRow1(); clearRow3(); }
  updateAnalyzeBtn();
});
lessonInput.addEventListener('keydown', e => { if (e.key === 'Enter') analyzeBtn.click(); });

// row 3 text inputs → row 1·2 초기화 + 나머지 row3 입력값 초기화
[uniqueidInput, expdateInput, unitnameInput].forEach(el => {
  el.addEventListener('input', function () {
    if (this.value.trim()) {
      clearRow1(); clearRow2();
      [uniqueidInput, expdateInput, unitnameInput]
        .filter(other => other !== this)
        .forEach(other => { other.value = ''; });
    }
    updateAnalyzeBtn();
  });
  el.addEventListener('keydown', e => { if (e.key === 'Enter') analyzeBtn.click(); });
});

function updateAnalyzeBtn() {
  const quick      = quickSearchInput.value.trim();
  const grade      = gradeSelect.value;
  const subject    = subjectSelect.value;
  const lesson     = lessonInput.value.trim();
  const extraGrade = extraGradeSelect.value;
  const uniqueId   = uniqueidInput.value.trim();
  const expDate    = expdateInput.value.trim();
  const unitName   = unitnameInput.value.trim();

  const mode1ok = !!quick;
  const mode2ok = !!(grade && subject && lesson);
  const mode3ok = !!(extraGrade && (uniqueId || expDate || unitName));
  analyzeBtn.disabled = !(mode1ok || mode2ok || mode3ok);
}

// ─── 단일 검색 실행 ──────────────────────────────────────
analyzeBtn.addEventListener('click', function () {
  let grade, subject, lessonCode;
  const quick = quickSearchInput.value.trim();

  if (quick) {
    const parsed = parseQuickSearch(quick);
    grade      = normalizeGrade(parsed.grade || gradeSelect.value.trim() || extraGradeSelect.value.trim() || '');
    subject    = parsed.grade ? resolveSubject(parsed.grade, parsed.subjectHint) : '';
    if (!subject && grade) subject = resolveSubject(grade, parsed.subjectHint);
    lessonCode = (parsed.lesson || '').replace(/'/g, '');
  } else {
    grade      = normalizeGrade(gradeSelect.value.trim());
    subject    = subjectSelect.value.trim();
    lessonCode = lessonInput.value.trim().replace(/'/g, '');
  }

  // Row 3이 활성이면 extraGrade를 grade로 사용
  if (!grade) grade = normalizeGrade(extraGradeSelect.value.trim());

  const uniqueId = uniqueidInput.value.trim();
  const expDate  = expdateInput.value.trim().replace(/\D/g, '');
  const unitName = unitnameInput.value.trim();

  if (!grade) return;

  let filtered = rows.filter(r => r.학년 === grade);
  if (subject)    filtered = filtered.filter(r => r.과목 === subject);
  if (lessonCode) filtered = filtered.filter(r => lessonMatches(r, lessonCode));
  if (uniqueId)   filtered = filtered.filter(r => r.차시고유번호.includes(uniqueId));
  if (expDate)    filtered = filtered.filter(r => r.노출일.includes(expDate));
  if (unitName)   filtered = filtered.filter(r => r.단원명.includes(unitName));

  if (lessonCode && !uniqueId && !expDate && !unitName) {
    renderDetailResult(filtered, grade, subject, lessonCode);
  } else {
    renderTableResult(filtered, grade, subject);
  }
});

// ─── 데이터 전처리 ───────────────────────────────────────
function processRows(data) {
  return data.map(row => {
    const r = {};
    Object.keys(row).forEach(k => { r[k.trim()] = (row[k] || '').trim(); });
    return {
      학년: r['학년'] || '',
      학기: r['학기'] || '',
      과목: r['과목'] || '',
      과목차시_clean: normalizeLessonCode(r['과목차시'] || r['과목차시_clean'] || ''),
      노출일: r['노출일'] || '',
      단원명: r['단원명'] || '',
      차시명: r['차시명'] || '',
      교과서: r['교과서'] || '',
      차시고유번호: r['차시고유번호'] || '',
    };
  }).filter(r => r.과목 && r.학년);
}

function normalizeLessonCode(value) {
  const clean = String(value || '').trim().replace(/'/g, '');
  if (!clean) return '';

  const monthDayMatch = clean.match(/^(\d{1,2})\s*월\s*(\d{1,2})\s*일?$/);
  if (monthDayMatch) return `${Number(monthDayMatch[1])}-${Number(monthDayMatch[2])}`;

  const fullDateMatch = clean.match(/^(?:19|20)?\d{2}[.\-/년]\s*(\d{1,2})[.\-/월]\s*(\d{1,2})\s*일?$/);
  if (fullDateMatch) return `${Number(fullDateMatch[1])}-${Number(fullDateMatch[2])}`;

  const shortDateMatch = clean.match(/^0?(\d{1,2})[.\-/]0?(\d{1,2})$/);
  if (shortDateMatch) return `${Number(shortDateMatch[1])}-${Number(shortDateMatch[2])}`;

  return clean;
}

function normalizeGrade(value) {
  const match = String(value || '').match(/\d+/);
  return match ? match[0] : '';
}

// ─── data.js 로드 ────────────────────────────────────────
function loadSavedUploadedData() {
  Object.keys(TERM_CONFIG).forEach(term => {
    const cfg = TERM_CONFIG[term];
    if (localStorage.getItem(cfg.markerKey) !== LS_UPLOAD_MARKER) return;
    const savedRows = localStorage.getItem(cfg.dataKey);
    const savedFiles = localStorage.getItem(cfg.filesKey);
    if (!savedRows) return;

    try {
      const parsedRows = JSON.parse(savedRows);
      if (!Array.isArray(parsedRows) || !parsedRows.length) return;
      termRows[term] = processRows(parsedRows);
      termFiles[term] = savedFiles ? JSON.parse(savedFiles) : ['업로드 데이터'];
    } catch {
      localStorage.removeItem(cfg.dataKey);
      localStorage.removeItem(cfg.filesKey);
      localStorage.removeItem(cfg.markerKey);
    }
  });
}

function loadDataJs() {
  const byTerm = window.CURRICULUM_DATA_BY_TERM;
  if (byTerm && typeof byTerm === 'object') {
    termRows.regular = processRows(byTerm.regular || []);
    termRows.vacation = processRows(byTerm.vacation || []);
  } else if (Array.isArray(window.CURRICULUM_DATA)) {
    termRows.regular = processRows(window.CURRICULUM_DATA);
    termRows.vacation = [];
  }

  termFiles.regular = termRows.regular.length ? ['data.js 정규학기'] : [];
  termFiles.vacation = termRows.vacation.length ? ['data.js 방학학기'] : [];
}

async function loadSharedScheduleCardData() {
  try {
    const data = await getSharedScheduleCardData();
    if (!data || typeof data !== 'object') return false;
    let loaded = false;
    Object.keys(TERM_CONFIG).forEach(term => {
      const entry = data[term];
      const sharedRows = Array.isArray(entry?.rows) ? entry.rows : [];
      if (!sharedRows.length) return;
      termRows[term] = processRows(sharedRows);
      termFiles[term] = Array.isArray(entry.files) && entry.files.length
        ? entry.files.map(name => `공유: ${name}`)
        : ['Supabase 공유 데이터'];
      loaded = true;
    });
    return loaded;
  } catch (error) {
    console.warn('Supabase schedule card data load failed', error);
    return false;
  }
}

async function getSharedScheduleCardData() {
  const response = await fetch(`${SUPABASE_URL}/storage/v1/object/public/${SUPABASE_BUCKET}/${SUPABASE_SCHEDULE_CARD_PATH}?t=${Date.now()}`, {
    cache: 'no-store',
  });
  if (response.status === 404) return null;
  if (!response.ok) throw new Error(`Supabase 데이터 로드 실패: ${response.status}`);
  return response.json();
}

async function saveSharedScheduleCardData(term, nextRows, fileNames) {
  const current = await getSharedScheduleCardData().catch(() => null) || {};
  current[term] = {
    rows: nextRows,
    files: fileNames,
    updatedAt: new Date().toISOString(),
  };
  await putJsonToSupabase(SUPABASE_SCHEDULE_CARD_PATH, current);
}

async function putJsonToSupabase(path, data) {
  const response = await fetch(`${SUPABASE_URL}/storage/v1/object/${SUPABASE_BUCKET}/${path}`, {
    method: 'PUT',
    headers: {
      apikey: SUPABASE_KEY,
      Authorization: `Bearer ${SUPABASE_KEY}`,
      'Content-Type': 'application/json; charset=utf-8',
      'x-upsert': 'true',
    },
    body: JSON.stringify(data),
  });
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`Supabase 저장 실패: ${response.status} ${text}`);
  }
}

function confirmUploadPassword() {
  if (sessionStorage.getItem(UPLOAD_AUTH_KEY) === 'true') return true;
  const input = prompt('업로드 비밀번호를 입력해 주세요.');
  if (input === UPLOAD_PASSWORD) {
    sessionStorage.setItem(UPLOAD_AUTH_KEY, 'true');
    return true;
  }
  return false;
}

// ─── 상세 분석 결과 ──────────────────────────────────────
function renderDetailResult(filtered, grade, subject, lessonCode) {
  resultSection.hidden = false;
  if (!filtered.length) { renderNoData(grade, subject, lessonCode); return; }

  const repDate   = mode(filtered.map(r => r.노출일));
  const repUnit   = mode(filtered.map(r => r.단원명));
  const repLesson = mode(filtered.map(r => r.차시명));
  const semester  = filtered[0].학기;
  const sorted    = [...filtered].sort((a, b) => a.교과서.localeCompare(b.교과서, 'ko'));

  lastSortedRows = sorted;
  lastRepUnit = repUnit;

  const idRows = sorted.map(r => `
    <tr><td>${escapeHtml(r.교과서)}</td><td>${escapeHtml(r.차시고유번호)}</td></tr>`).join('');

  const { displayRows, correctedRows } = resolveDateRows(sorted, repDate, grade, subject);
  lastCorrectedRows = correctedRows;

  const unitRows = displayRows.map(row => {
    const isDiff = row.단원명 !== repUnit;
    let badge = '';
    if (row._noAlt)          badge = '<span class="diff-badge diff-badge--warn">노출일 없음</span>';
    else if (row._dateFixed) badge = '<span class="diff-badge diff-badge--fix">노출일 보정</span>';
    else if (isDiff)         badge = '<span class="diff-badge">※ 상이</span>';
    return `<tr class="${(isDiff || row._dateFixed || row._noAlt) ? 'diff' : ''}">
      <td>${escapeHtml(row.교과서)}</td>
      <td>${escapeHtml(row.과목차시_clean)}</td>
      <td>${escapeHtml(row.단원명)}${badge}</td>
      <td>${escapeHtml(row.노출일 || '')}</td>
      <td>${escapeHtml(row.차시고유번호 || '')}</td>
    </tr>`;
  }).join('');

  resultContent.innerHTML = `
    <div class="result-header">
      <div class="result-title">${escapeHtml(subject)} (${escapeHtml(lessonCode)})</div>
      <div class="result-meta">${escapeHtml(grade)}학년 ${escapeHtml(semester)}학기</div>
    </div>
    <div class="info-grid">
      <div class="info-item">
        <div class="info-label">노출일</div>
        <div class="info-value">${escapeHtml(formatDate(repDate))}</div>
      </div>
      <div class="info-item">
        <div class="info-label info-label--with-action">
          <span>대표 단원명</span>
          <button class="btn-copy" onclick="copyRepUnit(this)">복사하기</button>
        </div>
        <div class="info-value">${escapeHtml(repUnit)}</div>
      </div>
      <div class="info-item">
        <div class="info-label">대표 차시명</div>
        <div class="info-value">${escapeHtml(repLesson)}</div>
      </div>
    </div>
    <div class="result-block">
      <div class="result-block-header">
        <div class="result-block-title">차시고유번호 리스트</div>
        <div class="copy-btns">
          <button class="btn-copy" id="copy-btn" onclick="copyIds()">복사하기</button>
          <button class="btn-copy btn-copy-excel" id="copy-excel-btn" onclick="copyIdsExcel()">엑셀 복사</button>
        </div>
      </div>
      <table class="result-table">
        <thead><tr><th>교과서</th><th>차시고유번호</th></tr></thead>
        <tbody>${idRows}</tbody>
      </table>
    </div>
    <div class="result-block">
      <div class="result-block-header">
        <div class="result-block-title">교과서별 단원명 검증</div>
        ${correctedRows.length ? `<button class="btn-copy" onclick="copyDiffRows()">상이 복사</button>` : ''}
      </div>
      <table class="result-table">
        <thead><tr><th>교과서</th><th>차시</th><th>단원명</th><th>노출일</th><th>차시고유번호</th></tr></thead>
        <tbody>${unitRows}</tbody>
      </table>
    </div>`;

  resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ─── 테이블 결과 ─────────────────────────────────────────
function renderTableResult(filtered, grade, subject) {
  resultSection.hidden = false;
  if (!filtered.length) { renderNoData(grade, subject, ''); return; }

  lastSortedRows = filtered;

  const tableRows = filtered.map(r => `
    <tr>
      <td>${escapeHtml(r.과목차시_clean)}</td>
      <td>${escapeHtml(r.교과서)}</td>
      <td>${escapeHtml(r.차시고유번호)}</td>
      <td>${escapeHtml(formatDate(r.노출일))}</td>
      <td>${escapeHtml(r.단원명)}</td>
      <td>${escapeHtml(r.차시명)}</td>
    </tr>`).join('');

  resultContent.innerHTML = `
    <div class="result-header">
      <div class="result-title">${escapeHtml(subject)}</div>
      <div class="result-meta">${escapeHtml(grade)}학년  ·  ${filtered.length}건</div>
    </div>
    <div class="result-block">
      <div class="result-block-header">
        <div class="result-block-title">검색 결과</div>
      </div>
      <table class="result-table">
        <thead>
          <tr>
            <th>과목차시</th><th>교과서</th><th>차시고유번호</th>
            <th>노출일</th><th>단원명</th><th>차시명</th>
          </tr>
        </thead>
        <tbody>${tableRows}</tbody>
      </table>
    </div>`;

  resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderNoData(grade, subject, lessonCode) {
  resultContent.innerHTML = `
    <div class="error-msg">
      <strong>${escapeHtml(grade)}학년 ${escapeHtml(subject)}${lessonCode ? ' ' + escapeHtml(lessonCode) : ''}</strong>에 해당하는 데이터가 없습니다.<br>
      조건을 다시 확인해 주세요.
    </div>`;
  resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ─── 단일 검색 복사 ──────────────────────────────────────
let lastSortedRows = [];
let lastCorrectedRows = [];     // 노출일 보정된 행 (단일 검색)
let lastRepUnit = '';
const bulkCorrectedMap = {};    // 노출일 보정된 행 (복수 검색, 요일별)

function formatPublisherCopy(row) {
  const publisher = String(row.교과서 || '').trim();
  return publisher ? `${publisher} : ${row.차시고유번호}` : row.차시고유번호;
}

function copyIds() {
  const text = lastSortedRows
    .map(formatPublisherCopy)
    .join('\n');
  clipboardWrite(text);
  const btn = document.getElementById('copy-btn');
  btn.textContent = '복사됨!';
  btn.classList.add('copied');
  setTimeout(() => { btn.textContent = '복사하기'; btn.classList.remove('copied'); }, 2000);
}

function copyIdsExcel() {
  const text = lastSortedRows
    .map(r => `${r.교과서}\t${r.차시고유번호}`)
    .join('\n');
  clipboardWrite(text);
  const btn = document.getElementById('copy-excel-btn');
  btn.textContent = '복사됨!';
  btn.classList.add('copied');
  setTimeout(() => { btn.textContent = '엑셀 복사'; btn.classList.remove('copied'); }, 2000);
}

// ─── 노출일 기반 행 교체 유틸 ────────────────────────────
// sorted 배열에서 repDate와 다른 교과서 행을 전역 rows에서 같은 날짜 행으로 교체
function resolveDateRows(sorted, repDate, grade, subject) {
  const correctedRows = [];
  const knownPublishers = new Set(sorted.map(r => r.교과서));

  const displayRows = sorted.map(row => {
    if (row.노출일 === repDate) return { ...row, _dateFixed: false };
    const alt = rows.find(r =>
      r.교과서 === row.교과서 && r.학년 === grade &&
      r.과목 === subject && r.노출일 === repDate
    );
    if (alt) {
      correctedRows.push(alt);
      return { ...alt, _dateFixed: true };
    }
    correctedRows.push({ ...row, _noAlt: true });
    return { ...row, _dateFixed: false, _noAlt: true };
  });

  // 원래 검색에 없었지만 같은 노출일(repDate)에 존재하는 교과서 추가
  const seen = new Set();
  for (const r of rows) {
    if (r.학년 === grade && r.과목 === subject && r.노출일 === repDate &&
        !knownPublishers.has(r.교과서) && !seen.has(r.교과서)) {
      seen.add(r.교과서);
      correctedRows.push(r);
      displayRows.push({ ...r, _dateFixed: true });
    }
  }

  displayRows.sort((a, b) => a.교과서.localeCompare(b.교과서, 'ko'));
  return { displayRows, correctedRows };
}

// ─── 유틸 ────────────────────────────────────────────────
function mode(arr) {
  const freq = {};
  let maxCount = 0, modeVal = arr[0];
  for (const v of arr) {
    freq[v] = (freq[v] || 0) + 1;
    if (freq[v] > maxCount) { maxCount = freq[v]; modeVal = v; }
  }
  return modeVal;
}

function formatDate(raw) {
  const s = String(raw).replace(/\D/g, '');
  if (s.length !== 8) return raw;
  return `${s.slice(0,4)}년 ${parseInt(s.slice(4,6))}월 ${parseInt(s.slice(6,8))}일`;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function clipboardWrite(text) {
  if (navigator.clipboard) {
    navigator.clipboard.writeText(text).catch(() => fallbackCopy(text));
  } else {
    fallbackCopy(text);
  }
}

function fallbackCopy(text) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.cssText = 'position:fixed;opacity:0';
  document.body.appendChild(ta);
  ta.select();
  document.execCommand('copy');
  document.body.removeChild(ta);
}

function copyRepUnit(btn) {
  if (!lastRepUnit) return;
  clipboardWrite(lastRepUnit);
  btn.textContent = '복사됨!';
  btn.classList.add('copied');
  setTimeout(() => { btn.textContent = '복사하기'; btn.classList.remove('copied'); }, 2000);
}

// ─── 데이터 종류 전환 ─────────────────────────────────────
function switchTerm(term, isInitial = false) {
  activeTerm = TERM_CONFIG[term] ? term : 'regular';
  updateTermSwitchUI();
  rows = termRows[activeTerm] || [];
  refreshDataControls(isInitial && activeTerm === 'regular');
}

function updateTermSwitchUI() {
  termButtons.forEach(btn => {
    btn.classList.toggle('active', btn.dataset.term === activeTerm);
    btn.hidden = false;
  });
  if (quickSearchInput) {
    quickSearchInput.placeholder = activeTerm === 'vacation'
      ? '예: 5학년 수학 1'
      : '예: 2학년 국어 1-4';
  }
  if (!termSwitchNote) return;
  termSwitchNote.textContent = activeTerm === 'vacation'
    ? '방학학기는 예1, 예2 형식도 1, 2로 검색합니다.'
    : '정규학기는 2학기 데이터 기준으로 검색합니다.';
}

function refreshDataControls(isBundle = false) {
  setFileStatus(termFiles[activeTerm] || [], rows.length, isBundle);
  populateGrades();
  if (rows.length) enableBaseControls();
  else disableBaseControls();
  clearRow1();
  clearRow2();
  clearRow3();
  resultSection.hidden = true;
  bulkEditSection.hidden = true;
  bulkResultSection.hidden = true;
  updateAnalyzeBtn();
}

function normalizeLessonSearchValue(lessonCode) {
  const clean = normalizeLessonCode(lessonCode).replace(/\s+/g, '');
  if (!clean) return '';

  const knownKeys = Object.keys(SUBJECT_MAP).sort((a, b) => b.length - a.length);
  for (const key of knownKeys) {
    if (!clean.startsWith(key)) continue;
    const rest = normalizeLessonCode(clean.slice(key.length)).replace(/\s+/g, '');
    if (rest && (/^\d/.test(rest) || /^예/i.test(rest))) return rest;
  }

  return clean;
}

function getLessonSearchKey(lessonCode) {
  const clean = normalizeLessonSearchValue(lessonCode);
  if (activeTerm !== 'vacation') return clean;
  return clean.replace(/^예/i, '');
}

function lessonMatches(row, lessonCode) {
  return getLessonSearchKey(row.과목차시_clean) === getLessonSearchKey(lessonCode);
}

// ─── 드롭다운 생성 ───────────────────────────────────────
function populateGrades() {
  const grades = [...new Set(rows.map(r => r.학년))].filter(Boolean)
    .sort((a, b) => Number(a) - Number(b));
  const opts = grades.map(g => `<option value="${escapeHtml(g)}">${escapeHtml(g)}학년</option>`).join('');
  gradeSelect.innerHTML      = '<option value="">-- 학년 선택 --</option>' + opts;
  extraGradeSelect.innerHTML = '<option value="">-- 학년 --</option>' + opts;
  bulkGradeSelect.innerHTML  = '<option value="">-- 학년 선택 --</option>' + opts;
}

function populateSubjects(grade) {
  const subjects = [...new Set(rows.filter(r => r.학년 === grade).map(r => r.과목))]
    .sort((a, b) => a.localeCompare(b, 'ko'));
  subjectSelect.innerHTML = '<option value="">-- 과목 선택 --</option>' +
    subjects.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
}

function populateLessons(grade, subject) {
  const lessons = [...new Set(
    rows.filter(r => r.학년 === grade && r.과목 === subject).map(r => {
      return activeTerm === 'vacation' ? getLessonSearchKey(r.과목차시_clean) : r.과목차시_clean;
    })
  )].sort((a, b) => {
    const parse = s => s.split('-').map(p => isNaN(p) ? p : Number(p));
    const [a1, a2] = parse(a), [b1, b2] = parse(b);
    if (a1 !== b1) return typeof a1 === typeof b1 ? (a1 > b1 ? 1 : -1) : (typeof a1 === 'number' ? -1 : 1);
    if (a2 === undefined) return -1;
    if (b2 === undefined) return 1;
    return typeof a2 === typeof b2 ? (a2 > b2 ? 1 : -1) : (typeof a2 === 'number' ? -1 : 1);
  });
  lessonSelectHelper.innerHTML = '<option value="">-- 목록에서 선택 --</option>' +
    lessons.map(l => `<option value="${escapeHtml(l)}">${escapeHtml(l)}</option>`).join('');
}

// ─── UI 상태 헬퍼 ─────────────────────────────────────────
function resetSubjectAndBelow() {
  subjectSelect.innerHTML = '<option value="">-- 과목 선택 --</option>';
  subjectSelect.disabled = true;
  lessonSelectHelper.innerHTML = '<option value="">-- 목록에서 선택 --</option>';
  lessonSelectHelper.disabled = true;
  lessonInput.value = ''; lessonInput.disabled = true;
  analyzeBtn.disabled = true;
  resultSection.hidden = true;
}

function enableBaseControls() {
  gradeSelect.disabled = false;
  extraGradeSelect.disabled = false;
  bulkGradeSelect.disabled = false;
  quickSearchInput.disabled = false;
}

function disableBaseControls() {
  gradeSelect.disabled = true;
  extraGradeSelect.disabled = true;
  bulkGradeSelect.disabled = true;
  quickSearchInput.disabled = true;
}

// ─── 빠른 입력 파싱 헬퍼 ──────────────────────────────────
function parseQuickSearch(text) {
  const gradeMatch = text.match(/(\d+)\s*학년/);
  const grade = gradeMatch ? gradeMatch[1] : null;
  const restNorm = text.replace(/\d+\s*학년/, '').trim().replace(/\s+/g, '');

  // 알려진 과목 약어/전체명을 긴 것부터 매칭
  const knownKeys = Object.keys(SUBJECT_MAP).sort((a, b) => b.length - a.length);
  let subjectHint = null;
  let lessonPart  = restNorm;

  for (const key of knownKeys) {
    if (restNorm.startsWith(key)) {
      subjectHint = SUBJECT_MAP[key]; // 전체명으로 변환
      lessonPart  = restNorm.slice(key.length);
      break;
    }
  }

  // 과목 미발견 시 첫 숫자 이전 한글을 힌트로
  if (!subjectHint && restNorm) {
    const m = restNorm.match(/^([가-힣]*?)(\d.*)$/);
    if (m && m[1]) subjectHint = m[1];
    lessonPart = m ? m[2] : restNorm;
  }

  return { grade, subjectHint, lesson: lessonPart.replace(/'/g, '') || null };
}

function resolveSubject(grade, hint) {
  if (!hint) return '';
  const subjects = [...new Set(rows.filter(r => r.학년 === grade).map(r => r.과목))];
  return subjects.find(s => s === hint || s.includes(hint) || hint.includes(s)) || hint;
}

function setFileStatus(names, count, isBundle = false) {
  fileStatusEl.classList.toggle('has-file', count > 0);
  const termLabel = TERM_CONFIG[activeTerm]?.label || '';
  let msg;
  if (!count) {
    msg = `${termLabel} 데이터 없음`;
  } else if (isBundle) {
    msg = `${termLabel} CSV 폴더 로드됨  (총 ${count.toLocaleString()}행)`;
  } else {
    const label = names.length <= 3 ? names.join(', ') : `${names.slice(0,2).join(', ')} 외 ${names.length-2}개`;
    msg = `${termLabel}: ${label}  (총 ${count.toLocaleString()}행)`;
  }
  fileStatusText.textContent = msg;
  // 토글 버튼 요약 업데이트
  const summary = document.getElementById('file-status-summary');
  const icon    = document.getElementById('file-status-icon');
  if (summary) summary.textContent = '— ' + msg;
  if (icon)    icon.textContent = count ? '✅' : '📂';
}

// ════════════════════════════════════════════════════════
// ─── 복수 검색 ───────────────────────────────────────────
// ════════════════════════════════════════════════════════

// 현재 이미지 File 객체 (파일선택 또는 붙여넣기)
let currentImageFile = null;

function setBulkImage(file) {
  if (!file || !file.type.startsWith('image/')) return;
  currentImageFile = file;
  bulkImageName.textContent = file.name || '붙여넣은 이미지';
  const url = URL.createObjectURL(file);
  bulkPreviewImg.src = url;
  bulkPreviewWrap.hidden = false;
  pasteZone.classList.add('paste-zone--has-image');
  pasteZone.textContent = '✓ 이미지 준비됨  (다시 붙여넣으면 교체됩니다)';
  updateBulkBtn();
}

function updateBulkBtn() {
  bulkAnalyzeBtn.disabled = !(bulkGradeSelect.value && currentImageFile);
}

bulkGradeSelect.addEventListener('change', updateBulkBtn);

// 파일 선택
bulkImageInput.addEventListener('change', function () {
  if (!this.files.length) return;
  setBulkImage(this.files[0]);
  this.value = '';
});

// 붙여넣기 (paste zone 클릭 포커스 후 Ctrl+V 또는 전역 paste)
pasteZone.addEventListener('click', () => pasteZone.focus());

function handlePaste(e) {
  if (!document.getElementById('tab-bulk') || document.getElementById('tab-bulk').hidden) return;
  const items = Array.from(e.clipboardData?.items || []);
  const imgItem = items.find(item => item.type.startsWith('image/'));
  if (!imgItem) return;
  e.preventDefault();
  setBulkImage(imgItem.getAsFile());
}

pasteZone.addEventListener('paste', handlePaste);
document.addEventListener('paste', handlePaste);

// 검색하기 (편집 후)
bulkSearchBtn.addEventListener('click', function () {
  const grade = normalizeGrade(bulkGradeSelect.value);
  if (!grade) { alert('학년을 선택해 주세요.'); return; }
  const schedule = {};
  DAYS.forEach(day => {
    const val = EDIT_INPUTS[day].value.trim();
    schedule[day] = val ? [val] : [];
  });
  renderWeekGrid(schedule, grade);
  bulkResultSection.hidden = false;
  bulkResultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
});

bulkAnalyzeBtn.addEventListener('click', async function () {
  const grade  = normalizeGrade(bulkGradeSelect.value);
  const apiKey = loadApiKey();
  if (!grade || !currentImageFile) return;
  if (!apiKey) {
    alert('config.js 파일에 OpenRouter API Key를 입력해 주세요.\n발급: https://openrouter.ai');
    return;
  }

  this.textContent = '분석 중...';
  this.disabled = true;
  bulkResultSection.hidden = true;

  try {
    const { base64, mediaType } = await fileToBase64(currentImageFile);
    const grade = normalizeGrade(bulkGradeSelect.value);
    const schedule = await analyzeImageWithClaude(base64, mediaType, apiKey);

    // 추출 결과를 편집 필드에 채우고 크로스체크
    DAYS.forEach(day => {
      const val = ((schedule[day] || []).filter(t => t.trim()))[0] || '';
      EDIT_INPUTS[day].value = val;
      validateEditInput(day, val, grade);
    });

    bulkEditSection.hidden = false;
    bulkResultSection.hidden = true;
    bulkEditSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (err) {
    alert('분석 중 오류가 발생했습니다:\n' + err.message);
  } finally {
    this.textContent = '분석하기';
    updateBulkBtn();
  }
});

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      const dataUrl = e.target.result;
      const [header, data] = dataUrl.split(',');
      const mediaType = header.match(/:(.*?);/)[1];
      resolve({ base64: data, mediaType });
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// 단일 API 호출 (내부용) — 열=요일 고정 프롬프트
async function callAnalyzeApi(base64, mediaType, apiKey) {
  const prompt = `이 이미지는 초등학교 주간 시간표입니다.

규칙:
1. 수업 내용(한글+숫자)이 있는 세로 열 5개를 찾으세요. 왼쪽의 주차/날짜 레이블 열과 오른쪽의 날짜 숫자 열은 제외합니다. 나머지 5개 열을 왼쪽부터 월, 화, 수, 목, 금으로 매핑합니다.
2. 각 열에서 배경이 선명한 빨간색인 셀을 찾으세요. 노랑·주황·살구·분홍·보라·하늘·초록·회색·흰색은 제외합니다. 처음 발견한 빨간 셀과 동일한 색조의 셀을 전 열에서 빠짐없이 찾으세요.
3. 열당 빨간 셀이 여러 개면 가장 위의 것 1개만 선택합니다.
4. 선택된 셀의 텍스트를 있는 그대로 읽습니다 (예: "국4-3", "영5-1", "바2-단풍", "사1-단평").

다른 텍스트 없이 반드시 아래 JSON 형식으로만 응답하세요:
{"월":["텍스트"],"화":["텍스트"],"수":["텍스트"],"목":["텍스트"],"금":["텍스트"]}
빨간 셀이 없는 요일은 빈 배열 []로 표시하세요.`;

  const res = await fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: 'openrouter/auto',
      messages: [{
        role: 'user',
        content: [
          { type: 'image_url', image_url: { url: `data:${mediaType};base64,${base64}` } },
          { type: 'text', text: prompt }
        ]
      }],
      temperature: 0,
      max_tokens: 2048
    })
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `API 오류 (${res.status})`);
  }

  const data = await res.json();
  const text = data.choices?.[0]?.message?.content?.trim() || '';
  // 그리디 매칭으로 가장 큰 JSON 블록 추출 (사고 과정 이후 최종 JSON)
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error('응답을 파싱할 수 없습니다:\n' + text);
  return JSON.parse(jsonMatch[0]);
}

// ─── 이미지 분석 (단일 호출) ─────────────────────────────
let bulkMismatches = {};

async function analyzeImageWithClaude(base64, mediaType, apiKey) {
  bulkMismatches = {};
  return await callAnalyzeApi(base64, mediaType, apiKey);
}

// 셀 텍스트 파싱
// - 1~2학년: 과목명 치환 + 과목차시 = 전체 텍스트
// - 3~6학년: 과목명 치환 + 과목차시 = 나머지 부분 (예: "수3-1"→"3-1", "국매체-3"→"매체-3")
// 지원 형식: "수3-5" / "국매체-3" / "바2-단풍" / "바슬즐" / "수학"
function parseCell(text, grade) {
  const gradeNum = parseInt(grade);
  const clean    = text.trim().replace(/\s+/g, '');

  // 알려진 과목 약어/전체명을 긴 것부터 순서대로 매칭 (예: "국어" 전에 "국" 매칭 방지)
  const knownKeys = Object.keys(SUBJECT_MAP).sort((a, b) => b.length - a.length);
  let abbr = '';
  let rest = clean;
  for (const key of knownKeys) {
    if (clean.startsWith(key)) {
      abbr = key;
      rest = clean.slice(key.length);
      break;
    }
  }

  if (!abbr) {
    // 폴백: 첫 번째 숫자 이전까지를 약어로 처리
    const m = clean.match(/^([가-힣]+?)(\d[\d\-가-힣]*)$/);
    abbr = m ? m[1] : clean;
    rest = m ? m[2] : '';
  }

  const subject = SUBJECT_MAP[abbr] || abbr;
  const lessonCode = normalizeLessonCode(rest || clean);

  return { subject, lessonCode };
}

// 데이터 조회 (단일 검색 로직 재사용)
function lookupLesson(grade, subject, lessonCode) {
  const gradeKey = normalizeGrade(grade);
  const filtered = rows.filter(r =>
    normalizeGrade(r.학년) === gradeKey && r.과목 === subject && lessonMatches(r, lessonCode)
  );
  if (!filtered.length) return null;
  const repDate   = mode(filtered.map(r => r.노출일));
  const repUnit   = mode(filtered.map(r => r.단원명));
  const repLesson = mode(filtered.map(r => r.차시명));
  const semester  = filtered[0].학기;
  const sorted    = [...filtered].sort((a, b) => a.교과서.localeCompare(b.교과서, 'ko'));
  return { repDate, repUnit, repLesson, semester, sorted };
}

// 주간 그리드 렌더링 (단일 검색과 동일한 서식)
let bulkAllEntries = [];

function renderWeekGrid(schedule, grade) {
  bulkAllEntries = [];
  weekGrid.innerHTML = '';

  DAYS.forEach(day => {
    const cellText = ((schedule[day] || []).filter(t => t.trim()))[0] || null;

    const col = document.createElement('div');
    col.className = 'day-col';
    col.innerHTML = `<div class="day-header">${day}</div>`;

    if (!cellText) {
      col.innerHTML += `<div class="day-empty">—</div>`;
    } else {
      const { subject, lessonCode } = parseCell(cellText, grade);
      const data = lookupLesson(grade, subject, lessonCode);

      if (!data) {
        col.innerHTML += `
          <div class="day-card day-card--empty">
            <div class="day-card-title">${escapeHtml(subject)}</div>
            <div class="day-card-lesson">${escapeHtml(lessonCode)}</div>
            <div class="day-card-error">데이터 없음<br>
              <span style="font-size:0.72rem;color:#718096;">
                검색: ${escapeHtml(grade)}학년 · ${escapeHtml(subject)} · ${escapeHtml(lessonCode)}
              </span>
            </div>
          </div>`;
      } else {
        bulkAllEntries.push({ day, subject, lessonCode, data });

        const idRows = data.sorted.map(r =>
          `<tr><td>${escapeHtml(r.교과서)}</td><td>${escapeHtml(r.차시고유번호)}</td></tr>`
        ).join('');

        const { displayRows: dRows, correctedRows: cRows } = resolveDateRows(
          data.sorted, data.repDate, grade, subject
        );
        bulkCorrectedMap[day] = cRows;

        const unitRows = dRows.map(row => {
          const isDiff = row.단원명 !== data.repUnit;
          let badge = '';
          if (row._noAlt)          badge = '<span class="diff-badge diff-badge--warn">노출일 없음</span>';
          else if (row._dateFixed) badge = '<span class="diff-badge diff-badge--fix">노출일 보정</span>';
          else if (isDiff)         badge = '<span class="diff-badge">※ 상이</span>';
          return `<tr class="${(isDiff || row._dateFixed || row._noAlt) ? 'diff' : ''}">
            <td>${escapeHtml(row.교과서)}</td>
            <td>${escapeHtml(row.과목차시_clean)}</td>
            <td>${escapeHtml(row.단원명)}${badge}</td>
            <td>${escapeHtml(row.노출일 || '')}</td>
            <td>${escapeHtml(row.차시고유번호 || '')}</td>
          </tr>`;
        }).join('');

        col.innerHTML += `
          <div class="day-detail">
            <div class="result-header">
              <div>
                <div class="result-title">${escapeHtml(subject)} (${escapeHtml(lessonCode)})</div>
                <div class="result-meta">${escapeHtml(grade)}학년 ${escapeHtml(data.semester)}학기</div>
              </div>
            </div>
            <div class="info-grid day-info-grid">
              <div class="info-item">
                <div class="info-label">노출일</div>
                <div class="info-value">${escapeHtml(data.repDate)}</div>
              </div>
            </div>
            <div class="result-block">
              <div class="result-block-header">
                <div class="result-block-title">대표 단원명</div>
                <button class="btn-copy" onclick="copyBulkRepUnit(this, '${day}')">복사하기</button>
              </div>
              <div class="result-block-value">${escapeHtml(data.repUnit)}</div>
            </div>
            <div class="result-block">
              <div class="result-block-title">대표 차시명</div>
              <div class="result-block-value">${escapeHtml(data.repLesson)}</div>
            </div>
            <div class="result-block">
              <div class="result-block-header result-block-header--col">
                <div class="result-block-title">차시고유번호 리스트</div>
                <div class="copy-btns">
                  <button class="btn-copy" onclick="copyBulkDayIds('${day}', this)">복사하기</button>
                  <button class="btn-copy btn-copy-excel" onclick="copyBulkDayIdsExcel('${day}', this)">엑셀 복사</button>
                </div>
              </div>
              <table class="result-table">
                <thead><tr><th>교과서</th><th>차시고유번호</th></tr></thead>
                <tbody>${idRows}</tbody>
              </table>
            </div>
            <div class="result-block">
              <div class="result-block-header">
                <div class="result-block-title">교과서별 단원명 검증</div>
                ${cRows.length ? `<button class="btn-copy" onclick="copyBulkDiffRows('${day}', this)">상이 복사</button>` : ''}
              </div>
              <table class="result-table">
                <thead><tr><th>교과서</th><th>차시</th><th>단원명</th><th>노출일</th><th>차시고유번호</th></tr></thead>
                <tbody>${unitRows}</tbody>
              </table>
            </div>
          </div>`;
      }
    }

    weekGrid.appendChild(col);
  });
}

// ─── 추출 결과 크로스체크 ─────────────────────────────────
const DAY_ID_MAP = { '월': 'mon', '화': 'tue', '수': 'wed', '목': 'thu', '금': 'fri' };

function validateEditInput(day, val, grade) {
  const inputEl = EDIT_INPUTS[day];
  const hintEl  = document.getElementById('hint-' + DAY_ID_MAP[day]);

  // 이전 상태 초기화
  inputEl.classList.remove('extract-input--ok', 'extract-input--warn', 'extract-input--mismatch');
  if (hintEl) { hintEl.textContent = ''; hintEl.className = 'extract-hint'; }

  if (!val.trim()) return;
  if (!grade) return;

  const mismatch = bulkMismatches[day];
  const { subject, lessonCode } = parseCell(val, grade);
  const found = lookupLesson(grade, subject, lessonCode);

  if (mismatch) {
    // AI 3회가 모두 다른 요일 → 데이터로 유효한 값 자동 선택
    const { votes } = mismatch;
    const validVote = votes.find(v => {
      const { subject: sv, lessonCode: lv } = parseCell(v, grade);
      return !!lookupLesson(grade, sv, lv);
    });

    if (validVote) {
      // 데이터에 있는 값으로 자동 교체
      EDIT_INPUTS[day].value = validVote;
      delete bulkMismatches[day];
      validateEditInput(day, validVote, grade);
      return;
    } else {
      // 전부 데이터 없음 → 경고 표시
      inputEl.classList.add('extract-input--warn');
      if (hintEl) {
        hintEl.innerHTML = `⚠ AI 3회 불일치: <b>${votes.map(escapeHtml).join('</b> / <b>')}</b> — 직접 확인`;
        hintEl.classList.add('extract-hint--warn');
      }
      return;
    }
  }

  if (found) {
    inputEl.classList.add('extract-input--ok');
    if (hintEl) {
      hintEl.textContent = '✓ ' + subject + ' ' + lessonCode;
      hintEl.classList.add('extract-hint--ok');
    }
  } else {
    inputEl.classList.add('extract-input--warn');
    if (hintEl) {
      hintEl.textContent = '⚠ 데이터 없음 — 수정 필요';
      hintEl.classList.add('extract-hint--warn');
    }
  }
}

// 편집 필드 수동 입력 시 즉시 재검증
DAYS.forEach(day => {
  EDIT_INPUTS[day].addEventListener('input', function () {
    const grade = normalizeGrade(bulkGradeSelect.value);
    validateEditInput(day, this.value.trim(), grade);
  });
});

// 요일별 복사
function copyBulkDayIds(day, btn) {
  const entry = bulkAllEntries.find(e => e.day === day);
  if (!entry) return;
  const text = entry.data.sorted.map(formatPublisherCopy).join('\n');
  clipboardWrite(text);
  btn.textContent = '복사됨!';
  btn.classList.add('copied');
  setTimeout(() => { btn.textContent = '복사하기'; btn.classList.remove('copied'); }, 2000);
}

// 단일 검색 — 노출일 보정 행 복사
function copyDiffRows() {
  if (!lastCorrectedRows.length) return;
  const text = lastCorrectedRows.map(formatPublisherCopy).join('\n');
  clipboardWrite(text);
}

// 복수 검색 — 요일별 노출일 보정 행 복사
function copyBulkDiffRows(day, btn) {
  const list = bulkCorrectedMap[day];
  if (!list || !list.length) return;
  const text = list.map(formatPublisherCopy).join('\n');
  clipboardWrite(text);
  btn.textContent = '복사됨!'; btn.classList.add('copied');
  setTimeout(() => { btn.textContent = '상이 복사'; btn.classList.remove('copied'); }, 1500);
}

function copyBulkRepUnit(btn, day) {
  const entry = bulkAllEntries.find(e => e.day === day);
  if (!entry) return;
  clipboardWrite(entry.data.repUnit);
  btn.textContent = '복사됨!';
  btn.classList.add('copied');
  setTimeout(() => { btn.textContent = '복사하기'; btn.classList.remove('copied'); }, 2000);
}

function copyBulkDayIdsExcel(day, btn) {
  const entry = bulkAllEntries.find(e => e.day === day);
  if (!entry) return;
  const text = entry.data.sorted.map(r => `${r.교과서}\t${r.차시고유번호}`).join('\n');
  clipboardWrite(text);
  btn.textContent = '복사됨!';
  btn.classList.add('copied');
  setTimeout(() => { btn.textContent = '엑셀 복사'; btn.classList.remove('copied'); }, 2000);
}

