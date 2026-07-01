const STORAGE_KEY = "schedule-card-workbook-v1";
const SOURCE_KEY = "schedule-card-source-v1";
const TAB_ORDER_KEY = "schedule-card-tab-order-v1";
const TIMELINE_KEY = "schedule-card-release-timeline-v1";
const UPLOAD_AUTH_KEY = "schedule-card-upload-auth-v1";
const UPLOAD_PASSWORD = "610503";
const ALL_DATES = "__all_dates__";
const SUPABASE_URL = "https://vmebzlinboxmgcrrorwv.supabase.co";
const SUPABASE_KEY = "sb_publishable_zpANEcZ0GfP44NpyHgZECQ_LPxB1LhR";
const SUPABASE_BUCKET = "schedule-data";
const SUPABASE_DATA_PATH = "current.xlsx";
const SUPABASE_TIMELINE_PATH = "timeline.json";
const SUPABASE_OVERALL_PATH = "overall-schedule.json";
const OVERALL_SCHEDULE_PATH = "./overall-schedule.json";
const OPEN_DATE = "운영 오픈 날짜";
const OPEN_DATE_CANDIDATES = [OPEN_DATE, "개정 오픈일"];
const UNIT_ORDER = "단원순서";
const LESSON_ORDER = "차시순서";
const SUBJECT_CANDIDATES = ["진입 과목명", "대표단원(한글/국어)", "과목"];
const TIMELINE_DAYS = 30;
const DEFAULT_QA_DAYS = 5;
const DEFAULT_STAGE_DAYS = 5;
const COPY_SUBJECT_ORDER = ["국어", "수학", "과학", "사회", "영어"];
const KOREA_HOLIDAYS_2026 = new Set([
  "2026-01-01",
  "2026-02-16",
  "2026-02-17",
  "2026-02-18",
  "2026-03-01",
  "2026-03-02",
  "2026-05-01",
  "2026-05-05",
  "2026-05-24",
  "2026-05-25",
  "2026-06-03",
  "2026-06-06",
  "2026-07-17",
  "2026-08-15",
  "2026-08-17",
  "2026-09-24",
  "2026-09-25",
  "2026-09-26",
  "2026-10-03",
  "2026-10-05",
  "2026-10-09",
  "2026-12-25",
]);

const state = {
  workbookBase64: null,
  sourceName: "",
  currentFileName: "",
  sheets: [],
  rows: [],
  selectedDate: "",
  selectedSubject: "",
  selectedCategory: "",
  tabOrder: "category-first",
  currentView: "lesson",
  timelineMode: "qa",
  timelineMarks: {},
  overallSchedule: null,
  search: "",
};

const els = {
  lessonView: document.querySelector("#lessonView"),
  overallView: document.querySelector("#overallView"),
  lessonViewButton: document.querySelector("#lessonViewButton"),
  overallViewButton: document.querySelector("#overallViewButton"),
  fileInput: document.querySelector("#fileInput"),
  overallFileInput: document.querySelector("#overallFileInput"),
  downloadButton: document.querySelector("#downloadButton"),
  searchInput: document.querySelector("#searchInput"),
  dateTabs: document.querySelector("#dateTabs"),
  releaseTimelineSection: document.querySelector("#releaseTimelineSection"),
  releaseTimeline: document.querySelector("#releaseTimeline"),
  timelineMonths: document.querySelector("#timelineMonths"),
  timelineRange: document.querySelector("#timelineRange"),
  timelineSummary: document.querySelector("#timelineSummary"),
  timelineModeButtons: document.querySelectorAll("[data-timeline-mode]"),
  qaCopyButton: document.querySelector("#qaCopyButton"),
  primaryTabs: document.querySelector("#primaryTabs"),
  secondaryTabs: document.querySelector("#secondaryTabs"),
  categoryFirstButton: document.querySelector("#categoryFirstButton"),
  subjectFirstButton: document.querySelector("#subjectFirstButton"),
  tableHeader: document.querySelector("#tableHeader"),
  tableBody: document.querySelector("#tableBody"),
  tableTitle: document.querySelector("#tableTitle"),
  rowCount: document.querySelector("#rowCount"),
  emptyState: document.querySelector("#emptyState"),
  overallSummary: document.querySelector("#overallSummary"),
  overallKpis: document.querySelector("#overallKpis"),
  overallTableHeader: document.querySelector("#overallTableHeader"),
  overallTableBody: document.querySelector("#overallTableBody"),
  overallEmptyState: document.querySelector("#overallEmptyState"),
  toast: document.querySelector("#toast"),
};

init().catch((error) => {
  console.error(error);
  showToast("엑셀 데이터를 불러오지 못했습니다.");
});

async function init() {
  state.workbookBase64 = localStorage.getItem(STORAGE_KEY);
  state.sourceName = localStorage.getItem(SOURCE_KEY) || "";
  state.tabOrder = localStorage.getItem(TAB_ORDER_KEY) || "category-first";
  state.timelineMarks = readTimelineMarks();
  await loadSharedTimeline();
  await loadOverallSchedule();
  const hasLocalUpload = state.workbookBase64 && state.sourceName.startsWith("현재 데이터:");
  if (hasLocalUpload) {
    await loadWorkbook(state.workbookBase64);
  } else {
    const sharedLoaded = await loadSharedWorkbook();
    if (!sharedLoaded && state.workbookBase64) {
      await loadWorkbook(state.workbookBase64);
    }
  }
  bindEvents();
  render();
}

function bindEvents() {
  els.lessonViewButton.addEventListener("click", () => {
    state.currentView = "lesson";
    render();
  });

  els.overallViewButton.addEventListener("click", () => {
    state.currentView = "overall";
    render();
  });

  els.fileInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!confirmUploadPassword()) {
      showToast("비밀번호가 맞지 않아 업로드를 취소했습니다.");
      event.target.value = "";
      return;
    }
    let base64;
    try {
      base64 = await fileToBase64(file);
      await loadWorkbook(base64);
    } catch (error) {
      console.error(error);
      showToast("업로드한 엑셀을 읽지 못했습니다.");
      event.target.value = "";
      return;
    }

    state.workbookBase64 = base64;
    localStorage.setItem(STORAGE_KEY, base64);
    localStorage.setItem(SOURCE_KEY, `현재 데이터: ${file.name}`);
    state.currentView = "lesson";
    state.sourceName = `현재 데이터: ${file.name}`;
    state.currentFileName = file.name;
    state.timelineMarks = createDefaultTimelineMarks();
    localStorage.setItem(TIMELINE_KEY, JSON.stringify(state.timelineMarks));
    showToast("현재 데이터가 교체 저장되었습니다. Supabase 저장을 진행합니다.");
    render();

    try {
      await saveCurrentWorkbookToSupabase();
      await saveTimelineToSupabase("기본 체크 일정도 저장했습니다.");
    } catch (error) {
      console.error(error);
      showToast(error.message ? `화면에는 반영됐지만 공유 저장 실패: ${error.message}` : "화면에는 반영됐지만 공유 저장에 실패했습니다.");
    } finally {
      event.target.value = "";
    }
  });

  els.overallFileInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!confirmUploadPassword()) {
      showToast("비밀번호가 맞지 않아 전체 일정 업로드를 취소했습니다.");
      event.target.value = "";
      return;
    }

    try {
      const base64 = await fileToBase64(file);
      state.overallSchedule = await parseOverallScheduleWorkbook(base64, file.name);
      state.currentView = "overall";
      render();
      await saveOverallScheduleToSupabase();
    } catch (error) {
      console.error(error);
      showToast(error.message || "전체 일정 엑셀을 저장하지 못했습니다.");
    } finally {
      event.target.value = "";
    }
  });

  els.downloadButton.addEventListener("click", () => {
    downloadFilteredWorkbook().catch((error) => {
      console.error(error);
      showToast("엑셀 다운로드 중 오류가 발생했습니다.");
    });
  });

  els.searchInput.addEventListener("input", (event) => {
    state.search = event.target.value.trim();
    state.selectedCategory = "";
    state.selectedSubject = "";
    render();
  });

  [els.categoryFirstButton, els.subjectFirstButton].forEach((button) => {
    button.addEventListener("click", () => {
      state.tabOrder = button.dataset.tabOrder;
      localStorage.setItem(TAB_ORDER_KEY, state.tabOrder);
      state.selectedCategory = "";
      state.selectedSubject = "";
      render();
    });
  });

  els.timelineModeButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.timelineMode = button.dataset.timelineMode;
      renderTimelineMode();
    });
  });

  els.qaCopyButton.addEventListener("click", () => {
    copyQaSummary().catch((error) => {
      console.error(error);
      showToast("QA 내용을 복사하지 못했습니다.");
    });
  });
}

function confirmUploadPassword() {
  if (sessionStorage.getItem(UPLOAD_AUTH_KEY) === "true") return true;

  const password = window.prompt("수정 비밀번호를 입력해 주세요.");
  if (password !== UPLOAD_PASSWORD) return false;

  sessionStorage.setItem(UPLOAD_AUTH_KEY, "true");
  return true;
}

async function loadWorkbook(base64) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(base64ToArrayBuffer(base64));

  const sheets = [];
  const rows = [];
  workbook.eachSheet((worksheet) => {
    const headers = getHeaders(worksheet);
    const openDateHeader = OPEN_DATE_CANDIDATES.find((header) => headers.includes(header));
    if (!headers.length || !openDateHeader) return;

    const subjectHeader = SUBJECT_CANDIDATES.find((name) => headers.includes(name));
    const sheetName = normalizeSheetName(worksheet.name);
    const sheet = {
      name: sheetName,
      headers,
      subjectHeader,
    };
    sheets.push(sheet);

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const values = headers.map((_, index) => normalizeCell(row.getCell(index + 1).value));
      if (values.every((value) => value === "")) return;
      const record = {};
      headers.forEach((header, index) => {
        record[header] = values[index];
      });
      record[OPEN_DATE] = record[openDateHeader];
      record.__sourceSheet = worksheet.name;
      record.__sheet = sheetName;
      record.__headers = headers;
      record.__subject = normalizeSubject(record, subjectHeader, sheetName);
      record.__openDate = normalizeDate(record[openDateHeader]);
      record.__unitOrder = toNumber(record[UNIT_ORDER]);
      record.__lessonOrder = toNumber(record[LESSON_ORDER]);
      rows.push(record);
    });
  });

  state.sheets = sheets;
  state.rows = rows.filter((row) => row.__openDate);
  ensureSelections();
}

async function parseOverallScheduleWorkbook(base64, fileName) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(base64ToArrayBuffer(base64));
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("전체 일정 엑셀에 시트가 없습니다.");

  const rows = [];
  let category = "";
  for (let rowNumber = 4; rowNumber <= worksheet.actualRowCount; rowNumber += 1) {
    const values = Array.from({ length: 8 }, (_, index) => normalizeCell(worksheet.getRow(rowNumber).getCell(index + 2).value));
    if (values.every((value) => value === "")) continue;
    if (values[0]) category = values[0];
    if (!values[1]) continue;

    rows.push({
      row: rowNumber,
      category,
      subject: values[1],
      term: values[2],
      round: values[3],
      openDate: normalizeOverallDate(values[4]),
      openDateText: values[4],
      scope: values[5],
      lessonShare: values[6],
      note: values[7],
    });
  }

  if (!rows.length) throw new Error("전체 일정 엑셀에서 일정 데이터를 찾지 못했습니다.");
  const merges = mergedRanges(worksheet);
  return {
    source: fileName,
    sheet: worksheet.name,
    title: normalizeCell(worksheet.getRow(2).getCell(2).value),
    version: normalizeCell(worksheet.getRow(2).getCell(9).value),
    mergeSpans: {
      category: mergeSpanMap(rows, merges, 2),
      note: mergeSpanMap(rows, merges, 9),
    },
    rows,
  };
}

async function loadSharedWorkbook() {
  try {
    const response = await fetch(supabasePublicUrl(), { cache: "no-store" });
    if (!response.ok) return false;

    const buffer = await response.arrayBuffer();
    const base64 = arrayBufferToBase64(buffer);
    await loadWorkbook(base64);
    state.workbookBase64 = base64;
    state.sourceName = "공유 데이터: current.xlsx";
    state.currentFileName = "current.xlsx";
    localStorage.setItem(STORAGE_KEY, base64);
    localStorage.setItem(SOURCE_KEY, state.sourceName);
    return true;
  } catch (error) {
    console.warn("공유 데이터를 불러오지 못했습니다.", error);
    return false;
  }
}

async function loadOverallSchedule() {
  const sharedLoaded = await loadSharedOverallSchedule();
  if (sharedLoaded) return true;

  try {
    const response = await fetch(`${OVERALL_SCHEDULE_PATH}?v=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) return false;
    state.overallSchedule = await response.json();
    return true;
  } catch (error) {
    console.warn("전체 일정 데이터를 불러오지 못했습니다.", error);
    state.overallSchedule = null;
    return false;
  }
}

async function loadSharedOverallSchedule() {
  try {
    const response = await fetch(supabasePublicUrl(SUPABASE_OVERALL_PATH), { cache: "no-store" });
    if (!response.ok) return false;
    state.overallSchedule = await response.json();
    return true;
  } catch (error) {
    console.warn("공유 전체 일정을 불러오지 못했습니다.", error);
    return false;
  }
}

function getHeaders(worksheet) {
  const headerRow = worksheet.getRow(1);
  const headers = [];
  for (let col = 1; col <= worksheet.actualColumnCount; col += 1) {
    const value = normalizeCell(headerRow.getCell(col).value);
    if (value) headers.push(value);
  }
  return headers;
}

function normalizeCell(value) {
  if (value == null) return "";
  if (value instanceof Date) return dateToYmd(value);
  if (typeof value === "object") {
    if (value.text) return String(value.text).trim();
    if (value.result != null) return normalizeCell(value.result);
    if (value.richText) return value.richText.map((part) => part.text || "").join("").trim();
  }
  return String(value).trim();
}

function normalizeSheetName(sheetName) {
  if (sheetName === "전체과목_국수사과영바") return "학교공부";
  if (sheetName === "검정교과서 학습관") return "검정교과서";
  return sheetName;
}

function normalizeSubject(record, subjectHeader, sheetName) {
  if (sheetName === "수학마스터") return "수학";
  if (subjectHeader && record[subjectHeader]) return String(record[subjectHeader]).trim();
  for (const header of SUBJECT_CANDIDATES) {
    if (record[header]) return String(record[header]).trim();
  }
  return "미분류";
}

function normalizeDate(value) {
  const text = String(value || "").replace(/[^\d]/g, "");
  if (text.length === 8) return `${text.slice(0, 4)}-${text.slice(4, 6)}-${text.slice(6, 8)}`;
  if (text.length === 6) return `20${text.slice(0, 2)}-${text.slice(2, 4)}-${text.slice(4, 6)}`;
  if (value instanceof Date) return dateToYmd(value);
  return "";
}

function normalizeOverallDate(value) {
  if (!value) return "";
  if (value instanceof Date) return dateToYmd(value);
  const normalized = normalizeDate(value);
  if (normalized) return normalized;
  const match = String(value).match(/(\d{1,2})월\s*(\d{1,2})일/);
  if (!match) return "";
  return `2026-${String(Number(match[1])).padStart(2, "0")}-${String(Number(match[2])).padStart(2, "0")}`;
}

function dateToYmd(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function toNumber(value) {
  const number = Number(String(value ?? "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(number) ? number : 999999;
}

function ensureSelections() {
  const searchedRows = rowsMatchingSearch();
  const dates = unique(searchedRows.map((row) => row.__openDate)).sort();
  if (!searchedRows.length) {
    state.selectedDate = "";
    state.selectedCategory = "";
    state.selectedSubject = "";
    return;
  }
  if (state.selectedDate !== ALL_DATES && !dates.includes(state.selectedDate)) state.selectedDate = ALL_DATES;

  if (isCategoryFirst()) {
    const categories = categoriesForDate(state.selectedDate);
    if (!categories.includes(state.selectedCategory)) state.selectedCategory = categories[0] || "";

    const subjects = subjectsForDateCategory(state.selectedDate, state.selectedCategory);
    if (!subjects.includes(state.selectedSubject)) state.selectedSubject = subjects[0] || "";
    return;
  }

  const subjects = subjectsForDate(state.selectedDate);
  if (!subjects.includes(state.selectedSubject)) state.selectedSubject = subjects[0] || "";

  const categories = categoriesForDateSubject(state.selectedDate, state.selectedSubject);
  if (!categories.includes(state.selectedCategory)) state.selectedCategory = categories[0] || "";
}

function render() {
  ensureSelections();
  renderView();
  renderOrderControl();
  renderDateTabs();
  renderReleaseTimeline();
  renderPrimaryTabs();
  renderSecondaryTabs();
  renderTable();
  renderOverallSchedule();
}

function renderView() {
  const isOverall = state.currentView === "overall";
  els.lessonView.hidden = isOverall;
  els.overallView.hidden = !isOverall;
  els.lessonViewButton.classList.toggle("active", !isOverall);
  els.overallViewButton.classList.toggle("active", isOverall);
}

function renderTimelineMode() {
  els.timelineModeButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.timelineMode === state.timelineMode);
  });
}

function isCategoryFirst() {
  return state.tabOrder !== "subject-first";
}

function renderOrderControl() {
  els.categoryFirstButton.classList.toggle("active", isCategoryFirst());
  els.subjectFirstButton.classList.toggle("active", !isCategoryFirst());
}

function renderDateTabs() {
  const searchedRows = rowsMatchingSearch();
  const dates = unique(searchedRows.map((row) => row.__openDate)).sort();
  els.dateTabs.innerHTML = "";
  if (searchedRows.length) {
    els.dateTabs.appendChild(tabButton("전체", searchedRows.length, state.selectedDate === ALL_DATES, () => {
      state.selectedDate = ALL_DATES;
      state.selectedCategory = "";
      state.selectedSubject = "";
      render();
    }));
  }
  dates.forEach((date) => {
    const count = searchedRows.filter((row) => row.__openDate === date).length;
    els.dateTabs.appendChild(tabButton(formatDateTabLabel(date), count, date === state.selectedDate, () => {
      state.selectedDate = date;
      state.selectedCategory = "";
      state.selectedSubject = "";
      render();
    }, null, date));
  });
}

function renderReleaseTimeline() {
  renderTimelineMode();
  els.releaseTimeline.innerHTML = "";
  if (!state.selectedDate || state.selectedDate === ALL_DATES) {
    els.releaseTimelineSection.hidden = true;
    els.timelineMonths.innerHTML = "";
    els.timelineSummary.innerHTML = "";
    return;
  }

  els.releaseTimelineSection.hidden = false;
  const days = releaseTimelineDays(state.selectedDate);
  const marks = timelineMarksForSelectedDate();
  els.timelineRange.textContent = `${formatDateWithDay(days[0])} ~ ${formatDateWithDay(days[days.length - 1])}`;
  renderTimelineMonths(days);

  const fragment = document.createDocumentFragment();
  days.forEach((date, index) => {
    const isOpenDate = index === days.length - 1;
    const blockedReason = isOpenDate ? "" : timelineBlockedReason(date);
    const mark = isOpenDate ? "open" : (blockedReason ? "" : (marks[date] || ""));
    const button = document.createElement("button");
    button.type = "button";
    button.className = `timeline-cell${mark ? ` ${mark}` : ""}${blockedReason ? " disabled" : ""}`;
    button.disabled = Boolean(blockedReason);
    button.dataset.date = date;
    button.title = `${formatDateWithDay(date)} ${blockedReason || markLabel(mark)}`;
    button.innerHTML = `<span class="day">${Number(date.slice(8, 10))}일</span>`;
    button.addEventListener("click", async () => {
      if (isOpenDate) {
        showToast("오픈일은 마지막 칸에 자동 표시됩니다.");
        return;
      }
      if (!confirmUploadPassword()) {
        showToast("비밀번호가 맞지 않아 체크 수정을 취소했습니다.");
        return;
      }
      if (state.timelineMode === "clear") {
        delete marks[date];
      } else if (marks[date] === state.timelineMode) {
        delete marks[date];
      } else {
        marks[date] = state.timelineMode;
      }
      state.timelineMarks[state.selectedDate] = normalizeTimelineMarksForDate(state.selectedDate, marks);
      localStorage.setItem(TIMELINE_KEY, JSON.stringify(state.timelineMarks));
      renderReleaseTimeline();
      try {
        await saveTimelineToSupabase();
      } catch (error) {
        console.error(error);
        showToast(error.message || "체크 설정 저장 중 오류가 발생했습니다.");
      }
    });
    fragment.appendChild(button);
  });
  els.releaseTimeline.appendChild(fragment);
  renderTimelineSummary(days);
}

function renderTimelineSummary(days) {
  const items = [
    ["qa", "검증계"],
    ["stage", "스테이징"],
  ].map(([mark, label]) => timelineSummaryItem(days, mark, label)).filter(Boolean);
  items.push(`<span class="timeline-summary-item"><i class="summary-dot open"></i><strong>오픈</strong>: ${formatDateDotWithDay(state.selectedDate)}</span>`);

  els.timelineSummary.innerHTML = items.length
    ? items.join("")
    : `<span class="timeline-summary-item">체크된 일정이 없습니다.</span>`;
}

function timelineSummaryItem(days, mark, label) {
  const marks = timelineMarksForSelectedDate();
  const marked = days.slice(0, -1).filter((date) => marks[date] === mark);
  if (!marked.length) return "";

  const text = `${formatDateDotWithDay(marked[0])}~${formatDateDotWithDay(marked[marked.length - 1])} (${marked.length}일간)`;

  return `<span class="timeline-summary-item"><i class="summary-dot ${mark}"></i><strong>${label}</strong>: ${text}</span>`;
}

async function copyQaSummary() {
  if (!state.selectedDate || state.selectedDate === ALL_DATES) {
    showToast("오픈 날짜를 선택한 뒤 복사해 주세요.");
    return;
  }

  const text = buildQaSummaryText();
  await copyTextToClipboard(text);
  showToast("QA 내용을 복사했습니다.");
}

async function copyTextToClipboard(text) {
  if (navigator.clipboard?.writeText && window.isSecureContext) {
    try {
      await navigator.clipboard.writeText(text);
      return;
    } catch {
      // Fall back to the selection-based copy path below.
    }
  }

  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.setAttribute("readonly", "");
  textarea.style.position = "fixed";
  textarea.style.left = "-9999px";
  textarea.style.top = "0";
  document.body.appendChild(textarea);
  textarea.select();
  const copied = document.execCommand("copy");
  textarea.remove();
  if (!copied) {
    window.prompt("자동 복사가 막혔습니다. 아래 내용을 복사해 주세요.", text);
  }
}

function buildQaSummaryText() {
  const rows = rowsMatchingSearch().filter((row) => row.__openDate === state.selectedDate);
  const stageRange = timelineCopyRange("stage") || "설정 없음";
  const lines = [
    "[일정]",
    `오픈일 : ${formatDateDotWithDay(state.selectedDate)}`,
    `스테이징 예정 기간 : ${stageRange}`,
    "QA 요청 기간 : 스테이징 기간 동일",
    "",
    "[과목별 오픈 차시]",
    ...subjectCopyLines(rows),
  ];

  return lines.join("\n");
}

function timelineCopyRange(mark) {
  const days = releaseTimelineDays(state.selectedDate);
  const marks = timelineMarksForSelectedDate();
  const marked = days.slice(0, -1).filter((date) => marks[date] === mark);
  if (!marked.length) return "";
  return `${formatDateDotWithDay(marked[0])}~${formatDateDotWithDay(marked[marked.length - 1])} (${marked.length}일간)`;
}

function subjectCopyLines(rows) {
  const subjects = unique(rows.map((row) => row.__subject)).sort(compareCopySubject);
  return subjects.map((subject) => {
    const subjectRows = rows.filter((row) => row.__subject === subject);
    const categories = unique(subjectRows.map((row) => row.__sheet))
      .sort((a, b) => state.sheets.findIndex((sheet) => sheet.name === a) - state.sheets.findIndex((sheet) => sheet.name === b))
      .map((category) => {
        const count = subjectRows.filter((row) => row.__sheet === category).length;
        return `${category} (${count.toLocaleString("ko-KR")}건)`;
      });

    return `${subject} (${subjectRows.length.toLocaleString("ko-KR")}건) : ${categories.join(", ")}`;
  });
}

function compareCopySubject(a, b) {
  const aIndex = COPY_SUBJECT_ORDER.indexOf(a);
  const bIndex = COPY_SUBJECT_ORDER.indexOf(b);
  const aOrder = aIndex === -1 ? COPY_SUBJECT_ORDER.length : aIndex;
  const bOrder = bIndex === -1 ? COPY_SUBJECT_ORDER.length : bIndex;
  return aOrder - bOrder || localeSort(a, b);
}

function renderTimelineMonths(days) {
  const groups = [];
  days.forEach((date) => {
    const month = `${Number(date.slice(5, 7))}월`;
    const latest = groups[groups.length - 1];
    if (latest?.month === month) {
      latest.count += 1;
    } else {
      groups.push({ month, count: 1 });
    }
  });

  els.timelineMonths.innerHTML = groups
    .map((group) => `<span style="--month-span: ${group.count}">${group.month}</span>`)
    .join("");
}

function renderPrimaryTabs() {
  if (isCategoryFirst()) {
    renderCategoryTabs("primary");
  } else {
    renderSubjectTabs("primary");
  }
}

function renderSecondaryTabs() {
  if (isCategoryFirst()) {
    renderSubjectTabs("secondary");
  } else {
    renderCategoryTabs("secondary");
  }
}

function renderSubjectTabs(level) {
  const target = level === "primary" ? els.primaryTabs : els.secondaryTabs;
  const subjects = isCategoryFirst()
    ? subjectsForDateCategory(state.selectedDate, state.selectedCategory)
    : subjectsForDate(state.selectedDate);
  target.innerHTML = "";
  subjects.forEach((subject) => {
    const count = level === "primary"
      ? categoriesForDateSubject(state.selectedDate, subject).length
      : rowsMatchingSearch().filter((row) => dateMatches(row, state.selectedDate) && row.__sheet === state.selectedCategory && row.__subject === subject).length;
    const lessonCount = level === "primary"
      ? rowsMatchingSearch().filter((row) => dateMatches(row, state.selectedDate) && row.__subject === subject).length
      : null;
    target.appendChild(tabButton(subject, count, subject === state.selectedSubject, () => {
      state.selectedSubject = subject;
      if (level === "primary") {
        state.selectedCategory = "";
        render();
      } else {
        renderTable();
        renderSecondaryTabs();
      }
    }, lessonCount));
  });
}

function renderCategoryTabs(level) {
  const target = level === "primary" ? els.primaryTabs : els.secondaryTabs;
  const categories = isCategoryFirst()
    ? categoriesForDate(state.selectedDate)
    : categoriesForDateSubject(state.selectedDate, state.selectedSubject);
  target.innerHTML = "";
  categories.forEach((category) => {
    const count = level === "primary"
      ? subjectsForDateCategory(state.selectedDate, category).length
      : rowsMatchingSearch().filter((row) => dateMatches(row, state.selectedDate) && row.__sheet === category && row.__subject === state.selectedSubject).length;
    const lessonCount = level === "primary"
      ? rowsMatchingSearch().filter((row) => dateMatches(row, state.selectedDate) && row.__sheet === category).length
      : null;
    target.appendChild(tabButton(category, count, category === state.selectedCategory, () => {
      state.selectedCategory = category;
      if (level === "primary") {
        state.selectedSubject = "";
        render();
      } else {
        renderTable();
        renderSecondaryTabs();
      }
    }, lessonCount));
  });
}

function tabButton(label, count, active, onClick, detailCount = null, dataLabel = label) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `tab${active ? " active" : ""}`;
  button.dataset.label = dataLabel;
  const countText = detailCount == null
    ? count.toLocaleString("ko-KR")
    : `${count.toLocaleString("ko-KR")} (${detailCount.toLocaleString("ko-KR")})`;
  button.innerHTML = `${escapeHtml(label)}<span class="count">${countText}</span>`;
  button.addEventListener("click", onClick);
  return button;
}

function renderTable() {
  const rows = getVisibleRows();
  const sheet = state.sheets.find((item) => item.name === state.selectedCategory);
  const headers = displayHeaders(sheet?.headers || []);

  const tablePath = isCategoryFirst()
    ? [dateLabel(state.selectedDate), state.selectedCategory, state.selectedSubject]
    : [dateLabel(state.selectedDate), state.selectedSubject, state.selectedCategory];
  els.tableTitle.textContent = tablePath.filter(Boolean).join(" > ") || "조회 결과";
  els.rowCount.textContent = `${rows.length.toLocaleString("ko-KR")}건`;
  els.downloadButton.disabled = state.rows.length === 0;

  els.tableHeader.innerHTML = "";
  els.tableBody.innerHTML = "";
  const headerRow = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = displayHeaderLabel(header, state.selectedCategory);
    headerRow.appendChild(th);
  });
  els.tableHeader.appendChild(headerRow);

  const fragment = document.createDocumentFragment();
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    headers.forEach((header) => {
      const td = document.createElement("td");
      td.textContent = displayValue(row[header], header);
      if (!["단원명", "차시명"].includes(header)) td.classList.add("center");
      if ([UNIT_ORDER, LESSON_ORDER].includes(header)) td.classList.add("number");
      if (isDateHeader(header)) td.classList.add("date");
      tr.appendChild(td);
    });
    fragment.appendChild(tr);
  });
  els.tableBody.appendChild(fragment);
  els.emptyState.hidden = rows.length > 0;
  els.emptyState.textContent = state.rows.length ? "조건에 맞는 데이터가 없습니다." : "엑셀 파일을 업로드하면 데이터가 표시됩니다.";
}

function renderOverallSchedule() {
  const scheduleRows = state.overallSchedule?.rows || [];
  const auditRows = scheduleRows.map((row) => ({
    ...row,
    audit: auditScheduleRow(row),
  }));
  const counts = auditRows.reduce((acc, row) => {
    acc[row.audit.status] = (acc[row.audit.status] || 0) + 1;
    return acc;
  }, {});

  els.overallSummary.textContent = state.overallSchedule
    ? `${state.overallSchedule.source || state.overallSchedule.sheet} · ${scheduleRows.length.toLocaleString("ko-KR")}건`
    : "일정 데이터 없음";
  els.overallKpis.innerHTML = [
    kpiHtml("등록됨", counts.matched || 0, "ok"),
    kpiHtml("미등록", counts.missing || 0, "warn"),
    kpiHtml("검수제외", counts.skipped || 0, "muted"),
  ].join("");

  const headers = ["카테고리", "과목", "학기", "차수", "오픈일", "콘텐츠 범위", "차시 수/비중", "비고", "검수 상태", "등록 차시"];
  els.overallTableHeader.innerHTML = `
    <tr class="overall-title-row">
      <th colspan="7">${escapeHtml(state.overallSchedule?.title || "스마트올 5, 6학년 2학기 콘텐츠 오픈 예정일")}</th>
      <th colspan="3">${escapeHtml(state.overallSchedule?.version || "")}</th>
    </tr>
    <tr>${headers.map((header) => `<th>${header}</th>`).join("")}</tr>
  `;
  els.overallTableBody.innerHTML = "";

  const fragment = document.createDocumentFragment();
  const categorySpans = state.overallSchedule?.mergeSpans?.category || categoryRowSpans(auditRows);
  const noteSpans = state.overallSchedule?.mergeSpans?.note || {};
  auditRows.forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.className = `audit-${row.audit.status} subject-${subjectTone(row.subject)}`;
    if (categorySpans[index]) {
      const categoryCell = document.createElement("td");
      categoryCell.className = "overall-category";
      categoryCell.rowSpan = categorySpans[index];
      categoryCell.textContent = compactScheduleCategory(row.category);
      tr.appendChild(categoryCell);
    }

    [
      row.subject,
      row.term,
      row.round,
      formatDateWithDay(row.openDate),
      row.scope,
      row.lessonShare,
    ].forEach((value, cellIndex) => {
      const td = document.createElement("td");
      td.textContent = value || "";
      if (![4, 5].includes(cellIndex)) td.classList.add("center");
      tr.appendChild(td);
    });

    if (noteSpans[index] !== 0) {
      const noteCell = document.createElement("td");
      noteCell.textContent = row.note || "";
      if (noteSpans[index]) noteCell.rowSpan = noteSpans[index];
      tr.appendChild(noteCell);
    }

    [row.audit.label, row.audit.detail].forEach((value, cellIndex) => {
      const td = document.createElement("td");
      td.textContent = value || "";
      if (cellIndex === 0) td.classList.add("center");
      tr.appendChild(td);
    });
    fragment.appendChild(tr);
  });

  els.overallTableBody.appendChild(fragment);
  els.overallEmptyState.hidden = scheduleRows.length > 0;
}

function mergedRanges(worksheet) {
  return (worksheet.model?.merges || []).map(parseCellRange).filter(Boolean);
}

function parseCellRange(range) {
  const [start, end = start] = String(range).split(":");
  const startCell = parseCellRef(start);
  const endCell = parseCellRef(end);
  if (!startCell || !endCell) return null;
  return {
    top: Math.min(startCell.row, endCell.row),
    bottom: Math.max(startCell.row, endCell.row),
    left: Math.min(startCell.col, endCell.col),
    right: Math.max(startCell.col, endCell.col),
  };
}

function parseCellRef(ref) {
  const match = String(ref).match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;
  return {
    col: columnNameToNumber(match[1]),
    row: Number(match[2]),
  };
}

function columnNameToNumber(name) {
  return String(name).toUpperCase().split("").reduce((number, char) => (
    number * 26 + char.charCodeAt(0) - 64
  ), 0);
}

function mergeSpanMap(rows, merges, columnNumber) {
  const spans = {};
  merges
    .filter((merge) => merge.left <= columnNumber && merge.right >= columnNumber)
    .forEach((merge) => {
      const indexes = rows
        .map((row, index) => (row.row >= merge.top && row.row <= merge.bottom ? index : -1))
        .filter((index) => index !== -1);
      if (indexes.length <= 1) return;
      spans[indexes[0]] = indexes.length;
      indexes.slice(1).forEach((index) => {
        spans[index] = 0;
      });
    });
  return spans;
}

function categoryRowSpans(rows) {
  const spans = {};
  let index = 0;
  while (index < rows.length) {
    const category = rows[index].category;
    let count = 1;
    while (index + count < rows.length && rows[index + count].category === category) count += 1;
    spans[index] = count;
    index += count;
  }
  return spans;
}

function subjectTone(subject) {
  if (subject === "수학") return "math";
  if (subject === "과학") return "science";
  if (subject === "영어") return "english";
  return "plain";
}

function kpiHtml(label, count, tone) {
  return `<span class="overall-kpi ${tone}"><strong>${count.toLocaleString("ko-KR")}</strong>${label}</span>`;
}

function auditScheduleRow(scheduleRow) {
  const categories = scheduleTargetCategories(scheduleRow.category);
  if (!categories.length || !scheduleRow.openDate) {
    return { status: "skipped", label: "검수제외", detail: "매칭 대상 없음" };
  }

  const matches = state.rows.filter((row) => (
    row.__openDate === scheduleRow.openDate
    && categories.includes(row.__sheet)
    && scheduleSubjectMatches(row, scheduleRow, categories)
  ));

  if (!matches.length) {
    return { status: "missing", label: "미등록", detail: `${categories.join(", ")} 기준 등록 차시 없음` };
  }

  const detail = categories
    .map((category) => {
      const count = matches.filter((row) => row.__sheet === category).length;
      return count ? `${category} ${count.toLocaleString("ko-KR")}건` : "";
    })
    .filter(Boolean)
    .join(", ");
  return { status: "matched", label: "등록됨", detail };
}

function scheduleTargetCategories(category) {
  const value = String(category || "");
  if (value.includes("학교공부")) return ["학교공부", "학교시험"];
  if (value.includes("성취도평가")) return ["성취도평가"];
  if (value.includes("학교시험")) return ["학교시험"];
  if (value.includes("검정")) return ["검정교과서"];
  if (value.includes("수학 마스터") || value.includes("AI수학")) return ["수학마스터"];
  return [];
}

function scheduleSubjectMatches(row, scheduleRow, categories) {
  if (categories.includes("수학마스터")) return row.__subject === "수학";
  return row.__subject === scheduleRow.subject;
}

function compactScheduleCategory(category) {
  return String(category || "").replace(/\n+/g, " / ");
}

function displayHeaders(headers) {
  const filtered = headers.filter((header) => header !== "비고");
  let next = moveBefore(filtered, "단원명", "차시명");
  if (state.selectedCategory === "성취도평가") {
    next = moveAfter(next, "진입 과목명", "학년");
  }
  if (usesRepresentativeUnitAsSubject(state.selectedCategory)) {
    next = moveAfter(next, "대표단원(한글/국어)", "학년");
  }
  if (state.selectedCategory === "학교공부") {
    next = moveAfter(next, "과목차시", "출판사");
  }
  return moveToFront(next, "학년");
}

function displayHeaderLabel(header, category) {
  if (usesRepresentativeUnitAsSubject(category)) {
    if (header === "대표단원(한글/국어)") return "과목";
    if (header === "과목") return "구분";
  }
  if (category === "성취도평가") {
    if (header === "진입 과목명") return "과목";
    if (header === "과목") return "구분";
  }
  return header;
}

function usesRepresentativeUnitAsSubject(category) {
  return category === "검정교과서" || category === "학교시험";
}

function moveBefore(headers, source, target) {
  const sourceIndex = headers.indexOf(source);
  const targetIndex = headers.indexOf(target);
  if (sourceIndex === -1 || targetIndex === -1 || sourceIndex === targetIndex - 1) return headers;

  const next = [...headers];
  const [item] = next.splice(sourceIndex, 1);
  next.splice(next.indexOf(target), 0, item);
  return next;
}

function moveToFront(headers, source) {
  const sourceIndex = headers.indexOf(source);
  if (sourceIndex <= 0) return headers;

  const next = [...headers];
  const [item] = next.splice(sourceIndex, 1);
  next.unshift(item);
  return next;
}

function moveAfter(headers, source, target) {
  const sourceIndex = headers.indexOf(source);
  const targetIndex = headers.indexOf(target);
  if (sourceIndex === -1 || targetIndex === -1 || sourceIndex === targetIndex + 1) return headers;

  const next = [...headers];
  const [item] = next.splice(sourceIndex, 1);
  next.splice(next.indexOf(target) + 1, 0, item);
  return next;
}

function getVisibleRows() {
  return rowsMatchingSearch()
    .filter((row) => dateMatches(row, state.selectedDate))
    .filter((row) => row.__subject === state.selectedSubject)
    .filter((row) => row.__sheet === state.selectedCategory)
    .sort(compareRows);
}

function categoriesForDate(date) {
  const order = state.sheets.map((sheet) => sheet.name);
  return unique(
    rowsMatchingSearch()
      .filter((row) => dateMatches(row, date))
      .map((row) => row.__sheet),
  ).sort((a, b) => order.indexOf(a) - order.indexOf(b));
}

function subjectsForDate(date) {
  return unique(
    rowsMatchingSearch()
      .filter((row) => dateMatches(row, date))
      .map((row) => row.__subject),
  ).sort(localeSort);
}

function subjectsForDateCategory(date, category) {
  return unique(
    rowsMatchingSearch()
      .filter((row) => dateMatches(row, date) && row.__sheet === category)
      .map((row) => row.__subject),
  ).sort(localeSort);
}

function categoriesForDateSubject(date, subject) {
  const order = state.sheets.map((sheet) => sheet.name);
  return unique(
    rowsMatchingSearch()
      .filter((row) => dateMatches(row, date) && row.__subject === subject)
      .map((row) => row.__sheet),
  ).sort((a, b) => order.indexOf(a) - order.indexOf(b));
}

function dateMatches(row, date) {
  return date === ALL_DATES || row.__openDate === date;
}

function dateLabel(date) {
  return date === ALL_DATES ? "전체" : formatDateWithDay(date);
}

function formatDateWithDay(date) {
  if (!date || date === ALL_DATES) return "전체";
  const [year, month, day] = date.split("-").map(Number);
  const parsed = new Date(year, month - 1, day);
  const weekdays = ["일", "월", "화", "수", "목", "금", "토"];
  return `${date} (${weekdays[parsed.getDay()]})`;
}

function formatDateTabLabel(date) {
  if (!date || date === ALL_DATES) return "전체";
  const [year, month, day] = date.split("-").map(Number);
  const parsed = new Date(year, month - 1, day);
  const weekdays = ["일", "월", "화", "수", "목", "금", "토"];
  return `${month}월 ${day}일 (${weekdays[parsed.getDay()]})`;
}

function formatDateDotWithDay(date) {
  if (!date) return "";
  const [year, month, day] = date.split("-").map(Number);
  const parsed = new Date(year, month - 1, day);
  const weekdays = ["일", "월", "화", "수", "목", "금", "토"];
  return `${year}.${String(month).padStart(2, "0")}.${String(day).padStart(2, "0")}(${weekdays[parsed.getDay()]})`;
}

function releaseTimelineDays(openDate) {
  const [year, month, day] = openDate.split("-").map(Number);
  const end = new Date(year, month - 1, day);
  return Array.from({ length: TIMELINE_DAYS }, (_, index) => {
    const date = new Date(end);
    date.setDate(end.getDate() - (TIMELINE_DAYS - 1 - index));
    return dateToYmd(date);
  });
}

function readTimelineMarks() {
  try {
    const value = JSON.parse(localStorage.getItem(TIMELINE_KEY) || "{}");
    return migrateTimelineMarks(value);
  } catch {
    return {};
  }
}

async function loadSharedTimeline() {
  try {
    const response = await fetch(supabasePublicUrl(SUPABASE_TIMELINE_PATH), { cache: "no-store" });
    if (!response.ok) return false;
    const value = await response.json();
    state.timelineMarks = migrateTimelineMarks(value);
    localStorage.setItem(TIMELINE_KEY, JSON.stringify(state.timelineMarks));
    return true;
  } catch (error) {
    console.warn("공유 체크 설정을 불러오지 못했습니다.", error);
    return false;
  }
}

async function saveTimelineToSupabase(message = "체크 설정을 저장했습니다.") {
  const result = await putJsonToSupabase(SUPABASE_TIMELINE_PATH, state.timelineMarks);
  if (!result.ok) throw new Error(result.message);
  showToast(message);
}

async function saveOverallScheduleToSupabase() {
  const result = await putJsonToSupabase(SUPABASE_OVERALL_PATH, state.overallSchedule);
  if (!result.ok) throw new Error(result.message);
  showToast("전체 일정을 저장했습니다.");
}

function migrateTimelineMarks(value) {
  const entries = Object.entries(value || {});
  if (!entries.length) return {};
  const hasLegacyFlatMarks = entries.some(([, mark]) => typeof mark === "string");
  if (hasLegacyFlatMarks) return {};
  return normalizeTimelineMarks(value);
}

function normalizeTimelineMarks(value) {
  return Object.fromEntries(
    Object.entries(value || {}).map(([openDate, marks]) => [
      openDate,
      normalizeTimelineMarksForDate(openDate, marks),
    ])
  );
}

function normalizeTimelineMarksForDate(openDate, marks) {
  return Object.fromEntries(
    Object.entries(marks || {}).filter(([date, mark]) => (
      date !== openDate
      && !timelineBlockedReason(date)
      && (mark === "qa" || mark === "stage")
    ))
  );
}

function createDefaultTimelineMarks() {
  const marksByOpenDate = {};
  unique(state.rows.map((row) => row.__openDate)).forEach((openDate) => {
    const workdays = workdaysBefore(openDate, DEFAULT_QA_DAYS + DEFAULT_STAGE_DAYS);
    const marks = {};
    workdays.slice(0, DEFAULT_QA_DAYS).forEach((date) => {
      marks[date] = "qa";
    });
    workdays.slice(DEFAULT_QA_DAYS, DEFAULT_QA_DAYS + DEFAULT_STAGE_DAYS).forEach((date) => {
      marks[date] = "stage";
    });
    marksByOpenDate[openDate] = normalizeTimelineMarksForDate(openDate, marks);
  });
  return marksByOpenDate;
}

function workdaysBefore(openDate, count) {
  const [year, month, day] = openDate.split("-").map(Number);
  const cursor = new Date(year, month - 1, day);
  const dates = [];
  while (dates.length < count) {
    cursor.setDate(cursor.getDate() - 1);
    const date = dateToYmd(cursor);
    if (!timelineBlockedReason(date)) dates.unshift(date);
  }
  return dates;
}

function timelineBlockedReason(date) {
  if (isWeekend(date)) return "주말";
  if (KOREA_HOLIDAYS_2026.has(date)) return "공휴일";
  return "";
}

function isWeekend(date) {
  const [year, month, day] = date.split("-").map(Number);
  const dayOfWeek = new Date(year, month - 1, day).getDay();
  return dayOfWeek === 0 || dayOfWeek === 6;
}

function timelineMarksForSelectedDate() {
  if (!state.selectedDate || state.selectedDate === ALL_DATES) return {};
  if (!state.timelineMarks[state.selectedDate]) state.timelineMarks[state.selectedDate] = {};
  return state.timelineMarks[state.selectedDate];
}

function markLabel(mark) {
  if (mark === "qa") return "검증계";
  if (mark === "stage") return "스테이징";
  if (mark === "open") return "오픈";
  return "";
}

async function saveCurrentWorkbookToSupabase() {
  if (!state.workbookBase64) {
    showToast("먼저 엑셀을 업로드해 주세요.");
    return;
  }

  const result = await putWorkbookToSupabase();
  if (!result.ok) throw new Error(result.message);

  state.sourceName = `공유 데이터: ${state.currentFileName || "current.xlsx"}`;
  localStorage.setItem(SOURCE_KEY, state.sourceName);
  showToast("Supabase에 공유 엑셀을 저장했습니다. 다른 사용자에게도 반영됩니다.");
  render();
}

async function putWorkbookToSupabase() {
  const bytes = base64ToArrayBuffer(state.workbookBase64);
  const saveResponse = await fetch(supabaseObjectUrl(), {
    method: "POST",
    headers: {
      apikey: SUPABASE_KEY,
      Authorization: `Bearer ${SUPABASE_KEY}`,
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "x-upsert": "true",
    },
    body: bytes,
  });

  if (!saveResponse.ok) {
    const text = await saveResponse.text();
    return {
      ok: false,
      status: saveResponse.status,
      message: `Supabase 저장 실패: ${saveResponse.status} ${text}`,
    };
  }

  return { ok: true, status: saveResponse.status };
}

async function putJsonToSupabase(path, value) {
  const saveResponse = await fetch(supabaseObjectUrl(path), {
    method: "POST",
    headers: {
      apikey: SUPABASE_KEY,
      Authorization: `Bearer ${SUPABASE_KEY}`,
      "Content-Type": "application/json",
      "x-upsert": "true",
    },
    body: JSON.stringify(value),
  });

  if (!saveResponse.ok) {
    const text = await saveResponse.text();
    return {
      ok: false,
      status: saveResponse.status,
      message: `Supabase 체크 설정 저장 실패: ${saveResponse.status} ${text}`,
    };
  }

  return { ok: true, status: saveResponse.status };
}

function supabaseObjectUrl(path = SUPABASE_DATA_PATH) {
  return `${SUPABASE_URL}/storage/v1/object/${SUPABASE_BUCKET}/${path}`;
}

function supabasePublicUrl(path = SUPABASE_DATA_PATH) {
  return `${SUPABASE_URL}/storage/v1/object/public/${SUPABASE_BUCKET}/${path}?v=${Date.now()}`;
}

function rowsMatchingSearch() {
  return state.rows.filter((row) => rowMatchesSearch(row));
}

function rowMatchesSearch(row) {
  const needle = state.search.toLowerCase();
  if (!needle) return true;
  return row.__headers.some((header) => String(row[header] ?? "").toLowerCase().includes(needle));
}

function compareRows(a, b) {
  return gradeOrder(a) - gradeOrder(b)
    || sortSubject(a).localeCompare(sortSubject(b), "ko")
    || sortGroup(a).localeCompare(sortGroup(b), "ko")
    || String(a["출판사"] || "").localeCompare(String(b["출판사"] || ""), "ko")
    || a.__unitOrder - b.__unitOrder
    || a.__lessonOrder - b.__lessonOrder
    || String(a["단원명"] || "").localeCompare(String(b["단원명"] || ""), "ko");
}

function sortSubject(row) {
  if (usesRepresentativeUnitAsSubject(row.__sheet)) return String(row["대표단원(한글/국어)"] || row.__subject || "");
  if (row.__sheet === "성취도평가") return String(row["진입 과목명"] || row.__subject || "");
  return String(row["과목"] || row.__subject || "");
}

function sortGroup(row) {
  if (usesRepresentativeUnitAsSubject(row.__sheet) || row.__sheet === "성취도평가") return String(row["과목"] || "");
  return String(row.__sheet || "");
}

function gradeOrder(row) {
  const grade = String(row["학년"] || "");
  const number = Number((grade.match(/\d+/) || ["999"])[0]);
  return Number.isFinite(number) ? number : 999;
}

async function downloadFilteredWorkbook() {
  if (!state.workbookBase64) return;
  const exportRows = getExportRows();
  if (!exportRows.length) return;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(base64ToArrayBuffer(state.workbookBase64));

  workbook.eachSheet((worksheet) => {
    const headers = getHeaders(worksheet);
    if (!OPEN_DATE_CANDIDATES.some((header) => headers.includes(header))) return;
    const sheetRows = exportRows.filter((row) => (row.__sourceSheet || row.__sheet) === worksheet.name);
    replaceWorksheetData(worksheet, headers, sheetRows);
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const safeDate = state.selectedDate === ALL_DATES ? "전체" : state.selectedDate.replaceAll("-", "");
  saveBlob(
    new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `개정_차시별운영계오픈일정_${safeDate}.xlsx`,
  );
  showToast("선택한 오픈날짜 기준으로 엑셀을 저장했습니다.");
}

function getExportRows() {
  return state.rows
    .filter((row) => dateMatches(row, state.selectedDate))
    .sort((a, b) => {
      const sheetOrder = state.sheets.findIndex((sheet) => sheet.name === a.__sheet)
        - state.sheets.findIndex((sheet) => sheet.name === b.__sheet);
      return sheetOrder || compareRows(a, b);
    });
}

function replaceWorksheetData(worksheet, headers, rows) {
  const lastRow = Math.max(worksheet.rowCount, worksheet.actualRowCount);
  if (lastRow > 1) worksheet.spliceRows(2, lastRow - 1);

  rows.forEach((record, index) => {
    const row = worksheet.getRow(index + 2);
    headers.forEach((header, colIndex) => {
      row.getCell(colIndex + 1).value = excelValue(record[header], header);
    });
    row.commit();
  });
}

function excelValue(value, header) {
  if (value == null || value === "") return null;
  if (isDateHeader(header)) return Number(String(normalizeDate(value) || value).replace(/[^\d]/g, ""));
  const numberHeaders = [UNIT_ORDER, LESSON_ORDER, "차시고유번호"];
  if (numberHeaders.includes(header) && Number.isFinite(Number(value))) return Number(value);
  return value;
}

function displayValue(value, header) {
  if (isDateHeader(header)) return normalizeDate(value) || value || "";
  return value ?? "";
}

function isDateHeader(header) {
  return header.includes("날짜") || header.includes("오픈일") || OPEN_DATE_CANDIDATES.includes(header);
}

function unique(values) {
  return [...new Set(values.filter((value) => value != null && value !== ""))];
}

function localeSort(a, b) {
  return String(a).localeCompare(String(b), "ko");
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.hidden = false;
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => {
    els.toast.hidden = true;
  }, 2600);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
