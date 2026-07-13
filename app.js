const STORAGE_KEY = "schedule-card-workbook-v1";
const SOURCE_KEY = "schedule-card-source-v1";
const TAB_ORDER_KEY = "schedule-card-tab-order-v1";
const TIMELINE_KEY = "schedule-card-release-timeline-v1";
const LEGACY_MATH_PUBLISHER_CONFIG_KEY = "schedule-card-math-publisher-config-v3";
const PUBLISHER_CONFIG_KEY = "schedule-card-open-publisher-config-v1";
const UPLOAD_AUTH_KEY = "schedule-card-upload-auth-v1";
const PUBLISHER_AUTH_KEY = "schedule-card-publisher-auth-v1";
const ALL_DATES = "__all_dates__";
const ALL_SUBJECTS = "__all_subjects__";
const ALL_CATEGORIES = "__all_categories__";
const SUPABASE_URL = "https://vmebzlinboxmgcrrorwv.supabase.co";
const SUPABASE_KEY = "sb_publishable_zpANEcZ0GfP44NpyHgZECQ_LPxB1LhR";
const SUPABASE_BUCKET = "schedule-data";
const SUPABASE_DATA_PATH = "current.xlsx";
const SUPABASE_TIMELINE_PATH = "timeline.json";
const SUPABASE_OVERALL_PATH = "overall-schedule.json";
const LEGACY_SUPABASE_MATH_PUBLISHER_PATH = "math-publisher-config-v3.json";
const SUPABASE_PUBLISHER_CONFIG_PATH = "open-publisher-config-v1.json";
const OVERALL_SCHEDULE_PATH = "./overall-schedule.json";
const OPEN_DATE = "운영 오픈 날짜";
const OPEN_DATE_LABEL = "운영 오픈 일정";
const OPEN_DATE_CANDIDATES = [OPEN_DATE, "개정 오픈일", OPEN_DATE_LABEL];
const UNIT_ORDER = "단원순서";
const LESSON_ORDER = "차시순서";
const SUBJECT_CANDIDATES = ["진입 과목명", "대표단원(한글/국어)", "과목"];
const MIXED_TABLE_HEADERS = ["학년", "과목", "구분", "학기", "출판사", UNIT_ORDER, LESSON_ORDER, "단원명", "차시명", "차시고유번호", OPEN_DATE_LABEL];
const TIMELINE_DAYS = 30;
const DEFAULT_QA_DAYS = 5;
const DEFAULT_STAGE_DAYS = 5;
const COPY_SUBJECT_ORDER = ["국어", "수학", "과학", "사회", "영어"];
const PUBLISHER_APPLY_SUBJECTS = ["수학", "과학", "사회"];
const MONTHLY_SERVICE_LABEL = "스마트올 초등";
const APP_PARAMS = new URLSearchParams(window.location.search);
const IS_LESSON_ONLY = APP_PARAMS.get("view") === "lesson-only";
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
  selectedTablePublisher: "",
  selectedTableGroup: "",
  selectedTableGrade: "",
  selectedMathAnalysisDate: "",
  selectedPublisherSubject: "수학",
  selectedWeeklyReportMonth: "",
  selectedWeeklyReportWeek: "",
  selectedMonthlyOpenMonth: "",
  mathPublisherConfig: {},
  tabOrder: "subject-first",
  currentView: "lesson",
  timelineMode: "qa",
  timelineMarks: {},
  overallSchedule: null,
  search: "",
};

const els = {
  lessonView: document.querySelector("#lessonView"),
  mathAnalysisView: document.querySelector("#mathAnalysisView"),
  weeklyReportView: document.querySelector("#weeklyReportView"),
  monthlyOpenView: document.querySelector("#monthlyOpenView"),
  scheduleCardView: document.querySelector("#scheduleCardView"),
  overallView: document.querySelector("#overallView"),
  lessonViewButton: document.querySelector("#lessonViewButton"),
  mathAnalysisViewButton: document.querySelector("#mathAnalysisViewButton"),
  weeklyReportViewButton: document.querySelector("#weeklyReportViewButton"),
  monthlyOpenViewButton: document.querySelector("#monthlyOpenViewButton"),
  scheduleCardViewButton: document.querySelector("#scheduleCardViewButton"),
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
  tableGradeTabs: document.querySelector("#tableGradeTabs"),
  tableGroupTabs: document.querySelector("#tableGroupTabs"),
  tablePublisherTabs: document.querySelector("#tablePublisherTabs"),
  rowCount: document.querySelector("#rowCount"),
  emptyState: document.querySelector("#emptyState"),
  mathAnalysisDateTabs: document.querySelector("#mathAnalysisDateTabs"),
  publisherSubjectTabs: document.querySelector("#publisherSubjectTabs"),
  mathPublisherFilters: document.querySelector("#mathPublisherFilters"),
  mathAnalysisKpis: document.querySelector("#mathAnalysisKpis"),
  mathAnalysisSummary: document.querySelector("#mathAnalysisSummary"),
  mathAnalysisTableBody: document.querySelector("#mathAnalysisTableBody"),
  mathAnalysisEmptyState: document.querySelector("#mathAnalysisEmptyState"),
  weeklyReportMonthTabs: document.querySelector("#weeklyReportMonthTabs"),
  weeklyReportWeekTabs: document.querySelector("#weeklyReportWeekTabs"),
  weeklyReportCopyButton: document.querySelector("#weeklyReportCopyButton"),
  weeklyReportContent: document.querySelector("#weeklyReportContent"),
  weeklyReportEmptyState: document.querySelector("#weeklyReportEmptyState"),
  monthlyOpenMonthTabs: document.querySelector("#monthlyOpenMonthTabs"),
  monthlyOpenCopyButton: document.querySelector("#monthlyOpenCopyButton"),
  monthlyOpenList: document.querySelector("#monthlyOpenList"),
  monthlyOpenEmptyState: document.querySelector("#monthlyOpenEmptyState"),
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
  document.body.classList.toggle("lesson-only", IS_LESSON_ONLY);
  if (IS_LESSON_ONLY) state.currentView = "lesson";
  state.workbookBase64 = localStorage.getItem(STORAGE_KEY);
  state.sourceName = localStorage.getItem(SOURCE_KEY) || "";
  state.tabOrder = "subject-first";
  localStorage.setItem(TAB_ORDER_KEY, state.tabOrder);
  state.timelineMarks = readTimelineMarks();
  state.mathPublisherConfig = readMathPublisherConfig();
  await loadSharedTimeline();
  await loadSharedMathPublisherConfig();
  await loadOverallSchedule();
  const sharedLoaded = await loadSharedWorkbook();
  if (!sharedLoaded && state.workbookBase64) {
    await loadWorkbook(state.workbookBase64);
  }
  bindEvents();
  render();
}

function bindEvents() {
  window.addEventListener("message", (event) => {
    if (event.data?.type !== "schedule-card-height") return;
    const height = Number(event.data.height);
    const frame = document.querySelector("#scheduleCardFrame");
    if (frame && Number.isFinite(height)) {
      frame.style.height = `${Math.max(height, 520)}px`;
    }
  });

  els.lessonViewButton.addEventListener("click", () => {
    state.currentView = "lesson";
    render();
  });

  els.mathAnalysisViewButton.addEventListener("click", () => {
    state.currentView = "math-analysis";
    render();
  });

  els.weeklyReportViewButton.addEventListener("click", () => {
    state.currentView = "weekly-report";
    render();
  });

  els.overallViewButton.addEventListener("click", () => {
    state.currentView = "overall";
    render();
  });

  els.monthlyOpenViewButton.addEventListener("click", () => {
    state.currentView = "monthly-open";
    render();
  });

  els.scheduleCardViewButton.addEventListener("click", () => {
    state.currentView = "schedule-card";
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
    const previousOpenDates = currentOpenDates();
    const previousTimelineMarks = clonePlainObject(state.timelineMarks || {});
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
    const timelineMerge = mergeTimelineMarksAfterUpload(previousTimelineMarks, previousOpenDates);
    state.timelineMarks = timelineMerge.marks;
    localStorage.setItem(TIMELINE_KEY, JSON.stringify(state.timelineMarks));
    showToast("현재 데이터가 교체 저장되었습니다. Supabase 저장을 진행합니다.");
    notifyOpenDateChanges(timelineMerge.addedDates, timelineMerge.removedDates, previousOpenDates.length > 0);
    render();

    try {
      await saveCurrentWorkbookToSupabase();
      await saveTimelineToSupabase("체크 일정을 저장했습니다.");
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
    state.selectedTablePublisher = "";
    state.selectedTableGroup = "";
    state.selectedTableGrade = "";
    render();
  });

  [els.categoryFirstButton, els.subjectFirstButton].forEach((button) => {
    button.addEventListener("click", () => {
      state.tabOrder = button.dataset.tabOrder;
      localStorage.setItem(TAB_ORDER_KEY, state.tabOrder);
      state.selectedCategory = "";
      state.selectedSubject = "";
      state.selectedTablePublisher = "";
      state.selectedTableGroup = "";
      state.selectedTableGrade = "";
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

  els.monthlyOpenCopyButton.addEventListener("click", () => {
    copyMonthlyOpenSummary().catch((error) => {
      console.error(error);
      showToast("월별 오픈 정보를 복사하지 못했습니다.");
    });
  });

  els.weeklyReportCopyButton.addEventListener("click", () => {
    copyWeeklyReportSummary().catch((error) => {
      console.error(error);
      showToast("주간보고 정보를 복사하지 못했습니다.");
    });
  });
}

function confirmUploadPassword() {
  const todayPassword = getTodayPassword();
  if (sessionStorage.getItem(UPLOAD_AUTH_KEY) === todayPassword) return true;

  const password = window.prompt("수정 비밀번호를 입력해 주세요.");
  if (password !== todayPassword) return false;

  sessionStorage.setItem(UPLOAD_AUTH_KEY, todayPassword);
  return true;
}

function confirmPublisherPassword() {
  const todayPassword = getTodayPassword();
  if (sessionStorage.getItem(PUBLISHER_AUTH_KEY) === todayPassword) return true;

  const password = window.prompt("출판사 설정 수정 비밀번호를 입력해 주세요.");
  if (password !== todayPassword) return false;

  sessionStorage.setItem(PUBLISHER_AUTH_KEY, todayPassword);
  return true;
}

function getTodayPassword() {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "Asia/Seoul",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());
  const month = parts.find((part) => part.type === "month")?.value || "01";
  const day = parts.find((part) => part.type === "day")?.value || "01";
  return `${month}${day}`;
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
  if (subjectHeader && record[subjectHeader]) return normalizeSubjectName(record[subjectHeader]);
  for (const header of SUBJECT_CANDIDATES) {
    if (record[header]) return normalizeSubjectName(record[header]);
  }
  return "미분류";
}

function normalizeSubjectName(value) {
  const subject = String(value || "").trim();
  const compact = subject.replace(/\s+/g, "").toLowerCase();
  if (compact === "ai수학" || compact === "all수학") return "수학";
  return subject;
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

function addDays(value, days) {
  const date = value instanceof Date
    ? new Date(value)
    : new Date(...String(value).split("-").map((part, index) => index === 1 ? Number(part) - 1 : Number(part)));
  date.setDate(date.getDate() + days);
  return date;
}

function weekStartDate(value) {
  const date = value instanceof Date ? new Date(value) : addDays(value, 0);
  const day = date.getDay();
  const offset = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + offset);
  return date;
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
    if (![ALL_CATEGORIES, ...categories].includes(state.selectedCategory)) state.selectedCategory = categories[0] || "";

    const subjects = subjectsForDateCategory(state.selectedDate, state.selectedCategory);
    if (![ALL_SUBJECTS, ...subjects].includes(state.selectedSubject)) state.selectedSubject = subjects[0] || "";
    return;
  }

  const subjects = subjectsForDate(state.selectedDate);
  if (![ALL_SUBJECTS, ...subjects].includes(state.selectedSubject)) state.selectedSubject = subjects[0] || "";

  const categories = categoriesForDateSubject(state.selectedDate, state.selectedSubject);
  if (![ALL_CATEGORIES, ...categories].includes(state.selectedCategory)) state.selectedCategory = categories[0] || "";
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
  renderWeeklyReportInfo();
  renderMonthlyOpenInfo();
  renderMathAnalysis();
}

function renderView() {
  if (IS_LESSON_ONLY) state.currentView = "lesson";
  const isOverall = state.currentView === "overall";
  const isWeeklyReport = state.currentView === "weekly-report";
  const isMonthlyOpen = state.currentView === "monthly-open";
  const isMathAnalysis = state.currentView === "math-analysis";
  const isScheduleCard = state.currentView === "schedule-card";
  els.lessonView.hidden = isOverall || isWeeklyReport || isMonthlyOpen || isMathAnalysis || isScheduleCard;
  els.weeklyReportView.hidden = !isWeeklyReport;
  els.monthlyOpenView.hidden = !isMonthlyOpen;
  els.mathAnalysisView.hidden = !isMathAnalysis;
  els.scheduleCardView.hidden = !isScheduleCard;
  els.overallView.hidden = !isOverall;
  els.lessonViewButton.classList.toggle("active", !isOverall && !isWeeklyReport && !isMonthlyOpen && !isMathAnalysis && !isScheduleCard);
  els.weeklyReportViewButton.classList.toggle("active", isWeeklyReport);
  els.monthlyOpenViewButton.classList.toggle("active", isMonthlyOpen);
  els.mathAnalysisViewButton.classList.toggle("active", isMathAnalysis);
  els.scheduleCardViewButton.classList.toggle("active", isScheduleCard);
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
  const today = dateToYmd(new Date());
  const nextOpenDate = dates.find((date) => date > today) || "";
  els.dateTabs.innerHTML = "";
  if (searchedRows.length) {
    els.dateTabs.appendChild(tabButton("전체", searchedRows.length, state.selectedDate === ALL_DATES, () => {
      state.selectedDate = ALL_DATES;
      state.selectedCategory = "";
      state.selectedSubject = "";
      state.selectedTablePublisher = "";
      state.selectedTableGroup = "";
      state.selectedTableGrade = "";
      render();
    }));
  }
  dates.forEach((date) => {
    const count = searchedRows.filter((row) => row.__openDate === date).length;
    const button = tabButton(formatDateTabLabel(date), count, date === state.selectedDate, () => {
      state.selectedDate = date;
      state.selectedCategory = "";
      state.selectedSubject = "";
      state.selectedTablePublisher = "";
      state.selectedTableGroup = "";
      state.selectedTableGrade = "";
      render();
    }, null, date);
    button.classList.toggle("past-open-date", date <= today);
    button.classList.toggle("next-open-date", date === nextOpenDate);
    els.dateTabs.appendChild(button);
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
    ` - 오픈일 : ${formatDateDotWithDay(state.selectedDate)}`,
    ` - 스테이징 예정 기간 : ${stageRange}`,
    " - QA 요청 기간 : 스테이징 기간 동일",
    "",
    "[총 오픈 차시]",
    `- ${rows.length.toLocaleString("ko-KR")}차시`,
    "",
    "[과목별 오픈 차시]",
    ...subjectCopyLines(rows),
  ];

  return lines.join("\n");
}

function renderWeeklyReportInfo() {
  const bundles = weeklyReportBundles();
  const months = weeklyReportMonths();
  if (!bundles.length) {
    els.weeklyReportMonthTabs.innerHTML = "";
    els.weeklyReportWeekTabs.innerHTML = "";
    els.weeklyReportCopyButton.disabled = true;
    els.weeklyReportContent.innerHTML = "";
    els.weeklyReportEmptyState.hidden = false;
    return;
  }

  if (!months.includes(state.selectedWeeklyReportMonth)) {
    state.selectedWeeklyReportMonth = months[0];
  }
  const weeks = weeklyReportWeeksForMonth(state.selectedWeeklyReportMonth);
  if (!weeks.some((week) => week.key === state.selectedWeeklyReportWeek)) {
    state.selectedWeeklyReportWeek = weeks.find((week) => week.count > 0)?.key || weeks[0]?.key || "";
  }

  els.weeklyReportMonthTabs.innerHTML = "";
  months.forEach((month) => {
    els.weeklyReportMonthTabs.appendChild(simpleTabButton(formatMonthlyOpenMonth(month), month === state.selectedWeeklyReportMonth, () => {
      state.selectedWeeklyReportMonth = month;
      state.selectedWeeklyReportWeek = "";
      renderWeeklyReportInfo();
    }, null, month));
  });

  els.weeklyReportWeekTabs.innerHTML = "";
  weeks.forEach((week) => {
    const button = weeklyReportWeekButton(week, week.key === state.selectedWeeklyReportWeek, () => {
      state.selectedWeeklyReportWeek = week.key;
      renderWeeklyReportInfo();
    });
    button.title = `${formatDateWithDay(week.start)} ~ ${formatDateWithDay(week.end)}`;
    els.weeklyReportWeekTabs.appendChild(button);
  });

  const report = weeklyReportData();
  els.weeklyReportCopyButton.disabled = report.sections.length === 0;
  els.weeklyReportContent.innerHTML = weeklyReportPreviewHtml(report);
  els.weeklyReportEmptyState.hidden = report.sections.length > 0;
}

function simpleTabButton(label, active, onClick) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `tab${active ? " active" : ""}`;
  button.textContent = label;
  button.addEventListener("click", onClick);
  return button;
}

function weeklyReportWeekButton(week, active, onClick) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `mini-tab${active ? " active" : ""}`;
  button.textContent = week.label;
  button.addEventListener("click", onClick);
  return button;
}

function weeklyReportRows() {
  return state.rows.filter((row) => row.__openDate);
}

function weeklyReportBundles() {
  const rows = weeklyReportRows();
  const rowsByDate = groupBy(rows, (row) => row.__openDate);
  return Object.keys(rowsByDate).sort().flatMap((openDate) => {
    const openWeek = weekStartDate(openDate);
    return [0, 1, 2, 3].map((offset) => {
      const weekStart = addDays(openWeek, -7 * offset);
      const weekEnd = addDays(weekStart, 4);
      return {
        openDate,
        progress: 100 - (offset * 20),
        weekKey: dateToYmd(weekStart),
        weekStart: dateToYmd(weekStart),
        weekEnd: dateToYmd(weekEnd),
        month: weeklyReportMonthKey(weekStart),
        rows: rowsByDate[openDate],
      };
    });
  }).filter((bundle) => bundle.progress >= 40);
}

function weeklyReportMonths() {
  return ["2026-07", "2026-08", "2026-09", "2026-10", "2026-11"];
}

function weeklyReportWeeksForMonth(month) {
  const bundles = weeklyReportBundles();
  return monthWeekStarts(month).map((weekStart, index) => {
    const key = dateToYmd(weekStart);
    const end = dateToYmd(addDays(weekStart, 4));
    const displayStart = maxDate(key, `${month}-01`);
    const displayEnd = minDate(end, lastDateOfMonth(month));
    return {
      key,
      start: key,
      end,
      label: `${index + 1}주 ${formatShortSlashDate(displayStart)}~${formatShortSlashDate(displayEnd)}`,
      count: bundles.filter((bundle) => bundle.weekKey === key).length,
    };
  });
}

function maxDate(a, b) {
  return a > b ? a : b;
}

function minDate(a, b) {
  return a < b ? a : b;
}

function lastDateOfMonth(month) {
  const [year, monthNumber] = month.split("-").map(Number);
  return dateToYmd(new Date(year, monthNumber, 0));
}

function monthWeekStarts(month) {
  const [year, monthNumber] = month.split("-").map(Number);
  const firstDay = new Date(year, monthNumber - 1, 1);
  let cursor = weekStartDate(firstDay);
  const weeks = [];
  while (weeklyReportMonthKey(cursor) <= month) {
    if (weeklyReportMonthKey(cursor) === month) weeks.push(new Date(cursor));
    cursor = addDays(cursor, 7);
  }
  return weeks;
}

function weeklyReportMonthKey(weekStart) {
  return dateToYmd(addDays(weekStart, 4)).slice(0, 7);
}

function weeklyReportData() {
  const bundles = weeklyReportBundles()
    .filter((bundle) => bundle.weekKey === state.selectedWeeklyReportWeek)
    .sort((a, b) => a.openDate.localeCompare(b.openDate) || b.progress - a.progress);
  const progress = bundles.length
    ? Math.round(bundles.reduce((sum, bundle) => sum + bundle.progress, 0) / bundles.length)
    : 0;
  const openDates = unique(bundles.map((bundle) => bundle.openDate)).sort();
  return {
    headline: `5,6학년 개정 2학기 오픈 / 공정률 ${progress}% (${openDates.map(formatShortSlashDate).join(", ")})`,
    progress,
    openDates,
    sections: weeklyReportSections(bundles),
  };
}

function weeklyReportSections(bundles) {
  const rows = bundles.flatMap((bundle) => bundle.rows);
  const subjects = unique(rows.map(weeklyReportSubject)).sort(compareSubjectOrder);
  return subjects.map((subject) => {
    const blocks = bundles.map((bundle) => weeklyReportSubjectBlock(subject, bundle)).filter(Boolean);
    return {
      subject,
      blocks,
    };
  }).filter((section) => section.blocks.length);
}

function weeklyReportSubjectBlock(subject, bundle) {
  const subjectRows = bundle.rows.filter((row) => weeklyReportSubject(row) === subject);
  if (!subjectRows.length) return null;
  const categories = weeklyReportCategories(subjectRows);
  const categoryNames = categories.map(({ sheet }) => weeklyReportCategoryName(sheet));
  return {
    subject,
    openDate: bundle.openDate,
    progress: bundle.progress,
    grades: weeklyGradeLabel(subjectRows),
    categoryNames,
    action: weeklyReportAction(bundle.progress),
    lines: categories.map(({ sheet, rows: categoryRows }) => ({
      label: weeklyReportCategoryName(sheet),
      detail: weeklyReportCategoryDetail(categoryRows),
    })),
  };
}

function weeklyReportSubject(row) {
  if (row.__sheet === "수학마스터") return "수학";
  return sortSubject(row) || row.__subject || "미분류";
}

function weeklyReportAction(progress) {
  if (progress === 100) return "오픈 및 모니터링";
  if (progress === 80) return "운영계 검수";
  if (progress === 60) return "검증계 및 운영계 검수";
  if (progress === 40) return "검증계 검수";
  return "검증계 검수";
}

function weeklyReportCategories(rows) {
  const order = ["학교공부", "학교시험", "수학마스터", "검정교과서", "성취도평가"];
  return unique(rows.map((row) => row.__sheet || "기타"))
    .sort((a, b) => {
      const aIndex = order.indexOf(a);
      const bIndex = order.indexOf(b);
      const aOrder = aIndex === -1 ? order.length : aIndex;
      const bOrder = bIndex === -1 ? order.length : bIndex;
      return aOrder - bOrder || localeSort(a, b);
    })
    .map((sheet) => ({ sheet, rows: rows.filter((row) => row.__sheet === sheet) }));
}

function weeklyReportCategoryName(sheet) {
  if (["학교공부", "학교시험", "수학마스터"].includes(sheet)) return `AI ${sheet}`;
  return sheet || "기타";
}

function weeklyReportCategoryDetail(rows) {
  const details = weeklyReportDetailGroups(rows).map(({ label, rows: groupRows }) => {
    const scope = weeklyReportScopeText(groupRows);
    return { label, scope };
  });
  return mergeWeeklyReportDetails(details).join(", ");
}

function mergeWeeklyReportDetails(details) {
  const unlabeled = details
    .filter((detail) => !detail.label)
    .map((detail) => detail.scope)
    .filter(Boolean);
  const labeledGroups = groupBy(details.filter((detail) => detail.label), (detail) => detail.scope || "");
  const labeled = Object.keys(labeledGroups)
    .sort((a, b) => compareWeeklyScopeOrder(a, b))
    .map((scope) => {
      const labels = labeledGroups[scope].map((detail) => detail.label).sort(compareWeeklyDetailLabel);
      return [labels.join(", "), scope].filter(Boolean).join(" ");
    });
  return [...unlabeled, ...labeled].filter(Boolean);
}

function compareWeeklyDetailLabel(a, b) {
  const order = ["단원요점정리", "단원 요점정리", "단원평가", "단원핵심특강", "서술형 트레이닝", "AI 서술형 평가", "실력완성문제"];
  const aIndex = order.indexOf(a);
  const bIndex = order.indexOf(b);
  const aOrder = aIndex === -1 ? order.length : aIndex;
  const bOrder = bIndex === -1 ? order.length : bIndex;
  return aOrder - bOrder || localeSort(a, b);
}

function compareWeeklyScopeOrder(a, b) {
  const aNumber = toNumber(a);
  const bNumber = toNumber(b);
  return aNumber - bNumber || localeSort(a, b);
}

function weeklyReportDetailGroups(rows) {
  const sheet = rows[0]?.__sheet || "";
  if (sheet === "학교시험") {
    const grouped = groupBy(rows, (row) => sortGroup(row) || "미분류");
    return Object.keys(grouped)
      .sort(localeSort)
      .map((label) => ({ label, rows: grouped[label] }));
  }
  if (usesRepresentativeUnitAsSubject(sheet) || sheet === "성취도평가") {
    const grouped = groupBy(rows, (row) => sortGroup(row) || "미분류");
    return Object.keys(grouped)
      .sort(compareSubjectOrder)
      .map((label) => ({ label, rows: grouped[label] }));
  }
  return [{ label: "", rows }];
}

function weeklyReportScopeText(rows) {
  const sortedRows = [...rows].sort(compareRows);
  const unitGroups = groupBy(sortedRows, (row) => monthlyUnitValue(row) || "미분류");
  const unitParts = Object.keys(unitGroups)
    .sort((a, b) => toNumber(a) - toNumber(b) || localeSort(a, b))
    .map((unit) => weeklyUnitScopeText(unit, unitGroups[unit]));
  return unitParts.join(", ");
}

function weeklyGradeLabel(rows) {
  return compressGradeLabels(unique(rows.map((row) => row["학년"])).sort((a, b) => toNumber(a) - toNumber(b)))
    .replace(/,\s+/g, ",");
}

function weeklyUnitScopeText(unit, rows) {
  const unitLabel = unit === "미분류" ? "단원 미분류" : `${unit}단원`;
  if (monthlyScopeUsesUnitOnly(rows)) return unitLabel;
  const lessons = unique(rows.map((row) => row.__lessonOrder).filter((value) => Number.isFinite(value) && value !== 999999))
    .sort((a, b) => a - b);
  const lessonText = weeklyLessonText(lessons);
  return lessonText ? `${unitLabel} ${lessonText}` : unitLabel;
}

function weeklyLessonText(lessons) {
  if (!lessons.length) return "";
  if (lessons.length === 1) return `${lessons[0]}차시`;
  const contiguous = lessons.every((lesson, index) => index === 0 || lesson === lessons[index - 1] + 1);
  if (lessons.length === 2) return `${lessons.join(",")}차시`;
  if (contiguous) return `${lessons[0]}~${lessons[lessons.length - 1]}차시`;
  return `${lessons.join(",")}차시`;
}

function weeklyReportPreviewHtml(report) {
  if (!report.sections.length) {
    return `<div class="empty">선택한 주차에 포함되는 주간보고 데이터가 없습니다.</div>`;
  }
  return `
    <div class="weekly-report-title">${escapeHtml(report.headline)}</div>
    ${report.sections.map((section) => `
      <section class="weekly-report-card">
        <h3>[${escapeHtml(section.subject)}]</h3>
        ${section.blocks.map(weeklyReportBlockHtml).join("")}
      </section>
    `).join("")}
  `;
}

function weeklyReportBlockHtml(block) {
  const dateLabel = formatShortSlashDate(block.openDate);
  const categoryText = block.categoryNames.join("/");
  return `
    <div class="weekly-report-block">
      <strong>${escapeHtml(block.grades)} ${escapeHtml(categoryText)} ${escapeHtml(block.action)} / 공정률 ${block.progress}% (${escapeHtml(dateLabel)})</strong>
      <ul>
        ${block.lines.map((line) => `
          <li><b>${escapeHtml(line.label)}</b> : ${escapeHtml(line.detail)}</li>
        `).join("")}
      </ul>
    </div>
  `;
}

async function copyWeeklyReportSummary() {
  const report = weeklyReportData();
  if (!report.sections.length) {
    showToast("복사할 주간보고 정보가 없습니다.");
    return;
  }
  await copyTableToClipboard(weeklyReportClipboardHtml(report), weeklyReportText(report));
  showToast("주간보고 정보를 복사했습니다.");
}

function weeklyReportClipboardHtml(report) {
  return `
    <ul>
      <li><strong>${escapeHtml(report.headline)}</strong></li>
      ${report.sections.map((section) => `
        <li><strong>[${escapeHtml(section.subject)}]</strong>
          <ul>
            ${section.blocks.map((block) => `
              <li><strong>${escapeHtml(block.grades)} ${escapeHtml(block.categoryNames.join("/"))} ${escapeHtml(block.action)} / 공정률 ${block.progress}% (${escapeHtml(formatShortSlashDate(block.openDate))})</strong>
                <ul>
                  ${block.lines.map((line) => `<li><strong>${escapeHtml(line.label)}</strong> : ${escapeHtml(line.detail)}</li>`).join("")}
                </ul>
              </li>
            `).join("")}
          </ul>
        </li>
      `).join("")}
    </ul>
  `;
}

function weeklyReportText(report) {
  return [
    report.headline,
    "",
    ...report.sections.map((section) => [
      `[${section.subject}]`,
      ...section.blocks.flatMap((block) => [
        `${block.grades} ${block.categoryNames.join("/")} ${block.action} / 공정률 ${block.progress}% (${formatShortSlashDate(block.openDate)})`,
        ...block.lines.map((line) => `- ${line.label} : ${line.detail}`),
      ]),
    ].join("\n")),
  ].join("\n\n");
}

function renderMonthlyOpenInfo() {
  const entries = monthlyOpenEntries();
  const months = unique(entries.map((entry) => entry.month)).sort();
  if (!months.length) {
    els.monthlyOpenMonthTabs.innerHTML = "";
    els.monthlyOpenCopyButton.disabled = true;
    els.monthlyOpenList.innerHTML = "";
    els.monthlyOpenEmptyState.hidden = false;
    return;
  }

  if (!months.includes(state.selectedMonthlyOpenMonth)) {
    state.selectedMonthlyOpenMonth = months[0];
  }

  els.monthlyOpenMonthTabs.innerHTML = "";
  months.forEach((month) => {
    const count = entries.filter((entry) => entry.month === month).length;
    els.monthlyOpenMonthTabs.appendChild(tabButton(formatMonthlyOpenMonth(month), count, month === state.selectedMonthlyOpenMonth, () => {
      state.selectedMonthlyOpenMonth = month;
      renderMonthlyOpenInfo();
    }, null, month));
  });
  els.monthlyOpenCopyButton.disabled = false;

  const monthEntries = entries.filter((entry) => entry.month === state.selectedMonthlyOpenMonth);
  els.monthlyOpenList.innerHTML = `
    <div class="monthly-open-table-wrap">
      <table class="monthly-open-table">
        <thead>
          <tr>
            <th>오픈일</th>
            <th>서비스</th>
            <th>오픈 정보</th>
            <th>대상</th>
          </tr>
        </thead>
        <tbody>
          ${monthEntries.map((entry) => `
            <tr>
              <td class="date">${escapeHtml(formatKoreanDate(entry.openDate))}</td>
              <td class="center">${escapeHtml(MONTHLY_SERVICE_LABEL)}</td>
              <td>
                <strong>[${escapeHtml(entry.category)} &gt; ${escapeHtml(entry.title)}]</strong>
                <div class="monthly-open-detail">${escapeHtml(entry.detail)}</div>
              </td>
              <td class="center">${escapeHtml(entry.grades)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
  els.monthlyOpenEmptyState.hidden = monthEntries.length > 0;
}

async function copyMonthlyOpenSummary() {
  const entries = monthlyOpenEntries().filter((entry) => entry.month === state.selectedMonthlyOpenMonth);
  if (!entries.length) {
    showToast("복사할 월별 오픈 정보가 없습니다.");
    return;
  }
  await copyTableToClipboard(monthlyOpenHtmlTable(entries), monthlyOpenTsv(entries));
  showToast("월별 오픈 정보를 복사했습니다.");
}

async function copyTableToClipboard(html, text) {
  if (navigator.clipboard?.write && window.ClipboardItem && window.isSecureContext) {
    try {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([text], { type: "text/plain" }),
        }),
      ]);
      return;
    } catch {
      // Fall back to plain text below.
    }
  }
  await copyTextToClipboard(text);
}

function monthlyOpenHtmlTable(entries) {
  const rows = entries.map((entry) => `
    <tr>
      <td>${escapeHtml(formatKoreanDate(entry.openDate))}</td>
      <td>${escapeHtml(MONTHLY_SERVICE_LABEL)}</td>
      <td><strong>[${escapeHtml(entry.category)} &gt; ${escapeHtml(entry.title)}]</strong><br><br>${escapeHtml(entry.detail).replace(/\n/g, "<br>")}</td>
      <td>${escapeHtml(entry.grades)}</td>
    </tr>
  `).join("");
  return `
    <table>
      <thead>
        <tr>
          <th>오픈일</th>
          <th>서비스</th>
          <th>오픈 정보</th>
          <th>대상</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function monthlyOpenTsv(entries) {
  return [
    ["오픈일", "서비스", "오픈 정보", "대상"].join("\t"),
    ...entries.map((entry) => [
      formatKoreanDate(entry.openDate),
      MONTHLY_SERVICE_LABEL,
      `[${entry.category} > ${entry.title}]\n\n${entry.detail}`,
      entry.grades,
    ].join("\t")),
  ].join("\n");
}

function monthlyOpenEntries() {
  const rows = [...state.rows].filter((row) => row.__openDate);
  const grouped = groupBy(rows, monthlyOpenGroupKey);
  return Object.values(grouped)
    .map(monthlyOpenEntryFromRows)
    .sort(compareMonthlyOpenEntries);
}

function monthlyOpenGroupKey(row) {
  return [
    row.__openDate,
    monthlyOpenCategory(row),
    monthlyOpenTitle(row),
  ].join("||");
}

function monthlyOpenEntryFromRows(rows) {
  const sortedRows = [...rows].sort(compareRows);
  const first = sortedRows[0];
  return {
    openDate: first.__openDate,
    month: first.__openDate.slice(0, 7),
    category: monthlyOpenCategory(first),
    title: monthlyOpenTitle(first),
    grades: monthlyOpenGrades(sortedRows),
    detail: monthlyOpenDetail(sortedRows),
    count: sortedRows.length,
  };
}

function monthlyOpenCategory(row) {
  const sheet = row.__sheet || "기타";
  if (["학교공부", "학교시험", "수학마스터"].includes(sheet)) return `AI ${sheet}`;
  return sheet;
}

function monthlyOpenTitle(row) {
  if (row.__sheet === "학교시험") return sortGroup(row) || "미분류";
  if (row.__sheet === "수학마스터") {
    return firstMeaningfulValue([
      row["과목"],
      row["구분"],
      row["대표단원(한글/국어)"],
      row["진입 과목명"],
    ], ["수학", "AI수학", "All수학"]) || "수학";
  }
  return sortSubject(row) || "미분류";
}

function monthlyOpenGrades(rows) {
  const grades = unique(rows.map((row) => row["학년"]))
    .sort((a, b) => toNumber(a) - toNumber(b));
  return compressGradeLabels(grades);
}

function monthlyOpenDetail(rows) {
  const detailGroups = monthlyDetailGroups(rows);
  return detailGroups.map(({ label, rows: groupRows }) => {
    const scope = monthlyScopeText(groupRows);
    return label ? `${label} - ${scope}` : scope;
  }).join("\n");
}

function monthlyDetailGroups(rows) {
  const sheet = rows[0]?.__sheet || "";
  if (sheet === "학교시험") {
    const grouped = groupBy(rows, (row) => sortSubject(row) || "미분류");
    return Object.keys(grouped)
      .sort(compareSubjectOrder)
      .map((label) => ({ label, rows: grouped[label] }));
  }
  if (usesRepresentativeUnitAsSubject(sheet) || sheet === "성취도평가") {
    const grouped = groupBy(rows, (row) => sortGroup(row) || "미분류");
    return Object.keys(grouped)
      .sort(compareSubjectOrder)
      .map((label) => ({ label, rows: grouped[label] }));
  }
  return [{ label: "", rows }];
}

function monthlyScopeText(rows) {
  const sortedRows = [...rows].sort(compareRows);
  const grades = compressGradeLabels(unique(sortedRows.map((row) => row["학년"])).sort((a, b) => toNumber(a) - toNumber(b)));
  const terms = unique(sortedRows.map((row) => row["학기"])).sort(localeSort).join(", ");
  const units = unique(sortedRows.map(monthlyUnitValue).filter((value) => value !== "")).sort((a, b) => toNumber(a) - toNumber(b));
  const lessons = sortedRows.map((row) => row.__lessonOrder).filter((value) => Number.isFinite(value) && value !== 999999);
  const unitText = monthlyRangeLabel(units, "단원");
  const lessonText = monthlyScopeUsesUnitOnly(sortedRows) ? "" : monthlyLessonRangeText(sortedRows, lessons);
  return [grades, terms, unitText, lessonText].filter(Boolean).join(" ");
}

function monthlyScopeUsesUnitOnly(rows) {
  const sheet = rows[0]?.__sheet || "";
  const group = sortGroup(rows[0] || {});
  if (sheet === "검정교과서") return true;
  return sheet === "학교시험" && [
    "서술형 트레이닝",
    "AI 서술형 평가",
    "단원평가",
    "단원 요점정리",
    "단원요점정리",
    "단원핵심특강",
  ].includes(group);
}

function monthlyUnitValue(row) {
  if (row.__sheet === "학교공부") {
    return leadingNumberText(row["단원명"]) || row[UNIT_ORDER];
  }
  if (row.__sheet === "학교시험" && sortSubject(row) === "영어") {
    return leadingNumberText(row["차시명"]) || row[UNIT_ORDER];
  }
  return row[UNIT_ORDER];
}

function leadingNumberText(value) {
  const match = String(value || "").trim().match(/^['"]?(\d+)/);
  return match ? match[1] : "";
}

function monthlyRangeLabel(values, suffix) {
  if (!values.length) return "";
  const numbers = values.map(toNumber).filter((value) => value !== 999999);
  if (numbers.length === values.length) {
    const min = Math.min(...numbers);
    const max = Math.max(...numbers);
    if (min === max) return `${min}${suffix}`;
    return `${min}~${max}${suffix}`;
  }
  if (values.length === 1) return `${values[0]}${suffix}`;
  return `${values[0]}~${values[values.length - 1]}${suffix}`;
}

function monthlyLessonRangeText(rows, lessons) {
  if (!lessons.length) return "";
  const min = Math.min(...lessons);
  const max = Math.max(...lessons);
  if (min === max) return `${min}차시`;
  return `${min}~${max}차시`;
}

function compareMonthlyOpenEntries(a, b) {
  return a.openDate.localeCompare(b.openDate)
    || monthlyCategoryOrder(a.category) - monthlyCategoryOrder(b.category)
    || compareSubjectOrder(a.title, b.title);
}

function monthlyCategoryOrder(category) {
  const order = ["AI 학교공부", "AI 학교시험", "AI 수학마스터", "검정교과서", "성취도평가"];
  const index = order.indexOf(category);
  return index === -1 ? order.length : index;
}

function firstMeaningfulValue(values, excluded = []) {
  const excludedSet = new Set(excluded.map((value) => String(value).replace(/\s+/g, "").toLowerCase()));
  return values.map((value) => String(value || "").trim()).find((value) => {
    if (!value) return false;
    return !excludedSet.has(value.replace(/\s+/g, "").toLowerCase());
  }) || "";
}

function compressGradeLabels(grades) {
  const numbers = unique(grades.map((grade) => Number((String(grade).match(/\d+/) || [""])[0])).filter(Number.isFinite)).sort((a, b) => a - b);
  if (!numbers.length) return "";
  if (numbers.length === 1) return `${numbers[0]}학년`;
  return `${numbers.join(", ")}학년`;
}

function formatMonthlyOpenMonth(month) {
  const [year, monthNumber] = month.split("-");
  return `${year}년 ${Number(monthNumber)}월`;
}

function formatKoreanDate(date) {
  const [year, month, day] = date.split("-").map(Number);
  return `${year}년 ${month}월 ${day}일`;
}

function formatShortSlashDate(date) {
  if (!date) return "";
  const [, month, day] = date.split("-").map(Number);
  return `${month}/${day}`;
}

function timelineCopyRange(mark) {
  const days = releaseTimelineDays(state.selectedDate);
  const marks = timelineMarksForSelectedDate();
  const marked = days.slice(0, -1).filter((date) => marks[date] === mark);
  if (!marked.length) return "";
  return `${formatDateDotWithDay(marked[0])}~${formatDateDotWithDay(marked[marked.length - 1])} (${marked.length}일간)`;
}

function subjectCopyLines(rows) {
  const subjects = unique(rows.map((row) => row.__subject)).sort(compareSubjectOrder);
  return subjects.map((subject) => {
    const subjectRows = rows.filter((row) => row.__subject === subject);
    const categories = unique(subjectRows.map((row) => row.__sheet))
      .sort((a, b) => state.sheets.findIndex((sheet) => sheet.name === a) - state.sheets.findIndex((sheet) => sheet.name === b))
      .map((category) => {
        const count = subjectRows.filter((row) => row.__sheet === category).length;
        return `${category} (${count.toLocaleString("ko-KR")}건)`;
      });

    return `- ${subject} (${subjectRows.length.toLocaleString("ko-KR")}건) : ${categories.join(", ")}`;
  });
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
  const allRows = rowsMatchingSearch().filter((row) => (
    dateMatches(row, state.selectedDate)
    && (!isCategoryFirst() || state.selectedCategory === ALL_CATEGORIES || row.__sheet === state.selectedCategory)
  ));
  target.appendChild(tabButton("전체", subjects.length, state.selectedSubject === ALL_SUBJECTS, () => {
    state.selectedSubject = ALL_SUBJECTS;
    state.selectedTablePublisher = "";
    state.selectedTableGroup = "";
    state.selectedTableGrade = "";
    if (level === "primary") {
      state.selectedCategory = ALL_CATEGORIES;
      render();
    } else {
      renderTable();
      renderSecondaryTabs();
    }
  }, allRows.length, "전체"));
  subjects.forEach((subject) => {
    const count = level === "primary"
      ? categoriesForDateSubject(state.selectedDate, subject).length
      : rowsMatchingSearch().filter((row) => (
        dateMatches(row, state.selectedDate)
        && (state.selectedCategory === ALL_CATEGORIES || row.__sheet === state.selectedCategory)
        && row.__subject === subject
      )).length;
    const lessonCount = level === "primary"
      ? rowsMatchingSearch().filter((row) => dateMatches(row, state.selectedDate) && row.__subject === subject).length
      : null;
    target.appendChild(tabButton(subject, count, subject === state.selectedSubject, () => {
      state.selectedSubject = subject;
      state.selectedTablePublisher = "";
      state.selectedTableGroup = "";
      state.selectedTableGrade = "";
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
  const allRows = rowsMatchingSearch().filter((row) => (
    dateMatches(row, state.selectedDate)
    && (!state.selectedSubject || state.selectedSubject === ALL_SUBJECTS || row.__subject === state.selectedSubject)
  ));
  target.appendChild(tabButton("전체", categories.length, state.selectedCategory === ALL_CATEGORIES, () => {
    state.selectedCategory = ALL_CATEGORIES;
    state.selectedTablePublisher = "";
    state.selectedTableGroup = "";
    state.selectedTableGrade = "";
    if (level === "primary") {
      state.selectedSubject = ALL_SUBJECTS;
      render();
    } else {
      renderTable();
      renderSecondaryTabs();
    }
  }, allRows.length, "전체"));
  categories.forEach((category) => {
    const count = level === "primary"
      ? subjectsForDateCategory(state.selectedDate, category).length
      : rowsMatchingSearch().filter((row) => (
        dateMatches(row, state.selectedDate)
        && row.__sheet === category
        && (state.selectedSubject === ALL_SUBJECTS || row.__subject === state.selectedSubject)
      )).length;
    const lessonCount = level === "primary"
      ? rowsMatchingSearch().filter((row) => dateMatches(row, state.selectedDate) && row.__sheet === category).length
      : null;
    target.appendChild(tabButton(category, count, category === state.selectedCategory, () => {
      state.selectedCategory = category;
      state.selectedTablePublisher = "";
      state.selectedTableGroup = "";
      state.selectedTableGrade = "";
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
  renderTableGradeTabs();
  renderTablePublisherTabs();
  const rows = getVisibleRows();
  const isMixedTable = state.selectedCategory === ALL_CATEGORIES;
  const sheet = state.sheets.find((item) => item.name === state.selectedCategory);
  const headers = isMixedTable ? MIXED_TABLE_HEADERS : displayHeaders(sheet?.headers || []);

  const tablePath = isCategoryFirst()
    ? [dateLabel(state.selectedDate), categoryLabel(state.selectedCategory), subjectLabel(state.selectedSubject)]
    : [dateLabel(state.selectedDate), subjectLabel(state.selectedSubject), categoryLabel(state.selectedCategory)];
  if (state.selectedTableGroup) tablePath.push(state.selectedTableGroup);
  if (state.selectedTablePublisher) tablePath.push(state.selectedTablePublisher);
  els.tableTitle.textContent = tablePath.filter(Boolean).join(" > ") || "조회 결과";
  els.rowCount.textContent = `${rows.length.toLocaleString("ko-KR")}건`;
  els.downloadButton.disabled = state.rows.length === 0 || !state.selectedDate || state.selectedDate === ALL_DATES;

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
      td.textContent = tableDisplayValue(row, header);
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

function renderTableGradeTabs() {
  const grades = ["5", "6"];
  const baseRows = getVisibleRows({ includeGrade: false });
  const availableGrades = new Set(baseRows.map(gradeValue).filter(Boolean));
  if (state.selectedTableGrade && !availableGrades.has(state.selectedTableGrade)) {
    state.selectedTableGrade = "";
  }

  els.tableGradeTabs.innerHTML = "";
  els.tableGradeTabs.appendChild(smallTabButton("전체", baseRows.length, state.selectedTableGrade === "", () => {
    state.selectedTableGrade = "";
    state.selectedTablePublisher = "";
    renderTable();
  }));

  grades.forEach((grade) => {
    const count = baseRows.filter((row) => gradeValue(row) === grade).length;
    els.tableGradeTabs.appendChild(smallTabButton(`${grade}학년`, count, state.selectedTableGrade === grade, () => {
      state.selectedTableGrade = grade;
      renderTable();
    }));
  });
}

function renderTablePublisherTabs() {
  const groupBaseRows = getVisibleRows({ includeGroup: false });
  const showGroupTabs = groupBaseRows.length > 0 && groupBaseRows.every(usesTableGroupTabs);
  els.tableGroupTabs.hidden = !showGroupTabs;
  els.tableGroupTabs.innerHTML = "";
  if (showGroupTabs) {
    renderTableGroupTabs(groupBaseRows);
  } else {
    state.selectedTableGroup = "";
  }

  const baseRows = getVisibleRows({ includePublisher: false });
  const publishers = tablePublisherCandidates(baseRows);
  const showPublisherTabs = Boolean(state.selectedTableGrade) && publishers.length > 0;
  els.tablePublisherTabs.hidden = !showPublisherTabs;
  els.tablePublisherTabs.innerHTML = "";
  if (!showPublisherTabs) {
    state.selectedTablePublisher = "";
    return;
  }
  if (state.selectedTablePublisher && !publishers.includes(state.selectedTablePublisher)) {
    state.selectedTablePublisher = "";
  }

  els.tablePublisherTabs.appendChild(smallTabButton("전체", baseRows.length, state.selectedTablePublisher === "", () => {
    state.selectedTablePublisher = "";
    renderTable();
  }));

  publishers.forEach((publisher) => {
    const count = baseRows.filter((row) => (row["출판사"] || "미분류") === publisher).length;
    const button = smallTabButton(publisher, count, state.selectedTablePublisher === publisher, () => {
      state.selectedTablePublisher = publisher;
      renderTable();
    });
    const status = tablePublisherStatus(publisher, baseRows);
    button.classList.toggle("table-publisher-main", status === "main");
    button.classList.toggle("table-publisher-disabled", status === "disabled");
    button.classList.toggle("table-publisher-mixed-disabled", status === "mixed-disabled");
    els.tablePublisherTabs.appendChild(button);
  });
}

function renderTableGroupTabs(baseRows) {
  const groups = unique(baseRows.map((row) => sortGroup(row) || "미분류")).sort(localeSort);
  if (state.selectedTableGroup && !groups.includes(state.selectedTableGroup)) {
    state.selectedTableGroup = "";
  }

  els.tableGroupTabs.appendChild(smallTabButton("전체", baseRows.length, state.selectedTableGroup === "", () => {
    state.selectedTableGroup = "";
    renderTable();
  }));

  groups.forEach((group) => {
    const count = baseRows.filter((row) => (sortGroup(row) || "미분류") === group).length;
    els.tableGroupTabs.appendChild(smallTabButton(group, count, state.selectedTableGroup === group, () => {
      state.selectedTableGroup = group;
      renderTable();
    }));
  });
}

function usesTableGroupTabs(row) {
  return row.__sheet === "학교시험" || row.__sheet === "수학마스터";
}

function smallTabButton(label, count, active, onClick) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `mini-tab${active ? " active" : ""}`;
  button.innerHTML = `${escapeHtml(label)}<span>${count.toLocaleString("ko-KR")}</span>`;
  button.addEventListener("click", onClick);
  return button;
}

function tablePublisherCandidates(rows) {
  const publishers = unique(rows.map((row) => row["출판사"] || "미분류"));
  if (!publishers.some(isMeaningfulPublisher)) return [];
  return publishers.sort((a, b) => compareTablePublisherOrder(a, b, rows));
}

function isMeaningfulPublisher(publisher) {
  const compact = String(publisher || "").replace(/\s+/g, "");
  return !["", "공통", "국정", "미분류"].includes(compact);
}

function compareTablePublisherOrder(a, b, rows = []) {
  const aStatus = tablePublisherSortStatus(a, rows);
  const bStatus = tablePublisherSortStatus(b, rows);
  if (aStatus.order !== bStatus.order) return aStatus.order - bStatus.order;
  if (aStatus.mainOrder !== bStatus.mainOrder) return aStatus.mainOrder - bStatus.mainOrder;
  if (a === "아이스크림" && b !== "아이스크림") return -1;
  if (b === "아이스크림" && a !== "아이스크림") return 1;
  return localeSort(a, b);
}

function tablePublisherSortStatus(publisher, rows) {
  const status = tablePublisherStatus(publisher, rows);
  return {
    order: status === "main" ? 0 : status === "disabled" ? 2 : 1,
    mainOrder: tablePublisherMainOrder(publisher, rows),
  };
}

function tablePublisherMainOrder(publisher, rows) {
  const orders = rows
    .filter((row) => publisherConfigAppliesToRow(row) && (row["출판사"] || "미분류") === publisher)
    .map((row) => {
      const config = publisherConfigForRow(row);
      const index = config?.main?.indexOf(publisher) ?? -1;
      return index === -1 ? 9999 : index;
    });
  return Math.min(...orders, 9999);
}

function tablePublisherStatus(publisher, rows) {
  const statuses = rows
    .filter((row) => (row["출판사"] || "미분류") === publisher)
    .map((row) => mathPublisherStatusForRow(row, publisher))
    .filter(Boolean);
  if (statuses.includes("main")) return "main";
  if (statuses.length && statuses.every((status) => status === "disabled")) return "disabled";
  if (statuses.includes("disabled")) return "mixed-disabled";
  return "sub";
}

function mathPublisherStatusForRow(row, publisher) {
  if (!publisherConfigAppliesToRow(row)) return "";
  const config = publisherConfigForRow(row);
  if (!config) return "";
  if ((config.main || []).includes(publisher)) return "main";
  if ((config.sub || []).includes(publisher) && (config.disabledSub || []).includes(publisher)) return "disabled";
  if ((config.sub || []).includes(publisher)) return "sub";
  return "";
}

function publisherConfigAppliesToRow(row) {
  return row.__sheet === "학교공부" && PUBLISHER_APPLY_SUBJECTS.includes(sortSubject(row));
}

function renderMathAnalysis() {
  const allRows = publisherAnalysisRows();
  const subjects = PUBLISHER_APPLY_SUBJECTS.filter((subject) => (
    allRows.some((row) => sortSubject(row) === subject)
  ));
  if (subjects.length && !subjects.includes(state.selectedPublisherSubject)) {
    state.selectedPublisherSubject = subjects[0];
    state.selectedMathAnalysisDate = "";
  }
  renderPublisherSubjectTabs(subjects, allRows);

  const rows = mathSchoolworkRows();
  const dates = unique(rows.map((row) => row.__openDate)).sort();
  if (!dates.includes(state.selectedMathAnalysisDate)) state.selectedMathAnalysisDate = dates[0] || "";
  const rowsForSelectedDate = rows.filter((row) => row.__openDate === state.selectedMathAnalysisDate);
  const publishersByGrade = mathPublishersByGrade(rowsForSelectedDate);
  ensureMathPublisherConfig(publishersByGrade);
  renderMathPublisherFilters(publishersByGrade, rows);
  const publisherRows = rows.filter((row) => mathPublisherSelected(row));

  els.mathAnalysisDateTabs.innerHTML = "";
  dates.forEach((date) => {
    const count = publisherRows.filter((row) => row.__openDate === date).length;
    els.mathAnalysisDateTabs.appendChild(tabButton(formatDateTabLabel(date), count, date === state.selectedMathAnalysisDate, () => {
      state.selectedMathAnalysisDate = date;
      renderMathAnalysis();
    }, null, date));
  });

  const selectedRows = rows.filter((row) => (
    row.__openDate === state.selectedMathAnalysisDate
    && mathPublisherSelected(row)
  ));
  const visibleRows = rows.filter((row) => (
    row.__openDate === state.selectedMathAnalysisDate
    && (mathPublisherSelected(row) || mathPublisherDisabled(row))
  ));
  const groups = mathLessonGroups(visibleRows);
  const duplicateGroups = groups.filter((group) => group.publishers.length > 1);
  const unitCount = unique(selectedRows.map((row) => mathUnitKey(row))).length;

  els.mathAnalysisSummary.textContent = state.selectedMathAnalysisDate
    ? `${formatDateWithDay(state.selectedMathAnalysisDate)} · 학교공부 ${currentPublisherSubject()} ${selectedRows.length.toLocaleString("ko-KR")}건`
    : `학교공부 ${currentPublisherSubject()} 데이터 없음`;
  els.mathAnalysisKpis.innerHTML = [
    analysisKpi("전체 차시", selectedRows.length),
    analysisKpi("단원", unitCount),
    analysisKpi("차시 묶음", groups.length),
    analysisKpi("중복 차시", duplicateGroups.length),
  ].join("");

  els.mathAnalysisTableBody.innerHTML = "";
  const fragment = document.createDocumentFragment();
  groups.forEach((group) => {
    const tr = document.createElement("tr");
    if (group.publishers.length > 1) tr.classList.add("duplicate");
    [group.grade, group.unitLabel, group.lessonLabel].forEach((value, index) => {
      const td = document.createElement("td");
      td.textContent = value;
      if (index === 0) td.classList.add("center");
      tr.appendChild(td);
    });
    const publisherTd = document.createElement("td");
    publisherTd.innerHTML = mathPublisherListHtml(group.grade, group.publishers);
    tr.appendChild(publisherTd);
    const countTd = document.createElement("td");
    countTd.textContent = `${group.publishers.length.toLocaleString("ko-KR")}개`;
    countTd.classList.add("center");
    tr.appendChild(countTd);
    fragment.appendChild(tr);
  });
  els.mathAnalysisTableBody.appendChild(fragment);
  els.mathAnalysisEmptyState.hidden = groups.length > 0;
}

function renderPublisherSubjectTabs(subjects, rows) {
  els.publisherSubjectTabs.innerHTML = "";
  if (!subjects.length) {
    els.publisherSubjectTabs.hidden = true;
    return;
  }
  els.publisherSubjectTabs.hidden = false;
  subjects.forEach((subject) => {
    const count = rows.filter((row) => sortSubject(row) === subject && mathPublisherSelected(row)).length;
    els.publisherSubjectTabs.appendChild(smallTabButton(subject, count, subject === state.selectedPublisherSubject, () => {
      state.selectedPublisherSubject = subject;
      renderMathAnalysis();
    }));
  });
}

function renderMathPublisherFilters(publishersByGrade, rows) {
  els.mathPublisherFilters.innerHTML = "";
  const grades = Object.keys(publishersByGrade).sort((a, b) => gradeNumber(a) - gradeNumber(b));
  if (!grades.length) return;

  grades.forEach((grade) => {
    const panel = document.createElement("section");
    panel.className = "publisher-grade-panel";

    const title = document.createElement("div");
    title.className = "publisher-grade-title";
    const selectedCount = mathSelectedCountForGrade(rows, grade);
    const disabledCount = mathDisabledCountForGrade(rows, grade);
    title.innerHTML = `
      <strong>${escapeHtml(grade)}</strong>
      <span>${selectedCount.toLocaleString("ko-KR")}차시</span>
      <span class="disabled-count">해제 ${disabledCount.toLocaleString("ko-KR")}차시</span>
    `;
    panel.appendChild(title);

    const lanes = document.createElement("div");
    lanes.className = "publisher-lanes";
    lanes.appendChild(publisherLane(grade, "main", "메인 출판사"));
    lanes.appendChild(publisherLane(grade, "sub", "서브 출판사"));
    lanes.appendChild(publisherLane(grade, "disabled", "해제 출판사"));
    panel.appendChild(lanes);

    els.mathPublisherFilters.appendChild(panel);
  });
}

function mathSelectedCountForGrade(rows, grade) {
  return rows.filter((row) => (
    row["학년"] === grade
    && row.__openDate === state.selectedMathAnalysisDate
    && mathPublisherSelected(row)
  )).length;
}

function mathDisabledCountForGrade(rows, grade) {
  return rows.filter((row) => (
    row["학년"] === grade
    && row.__openDate === state.selectedMathAnalysisDate
    && mathPublisherDisabled(row)
  )).length;
}

function publisherLane(grade, lane, title) {
  const config = mathPublisherConfigForSelectedDate()[grade] || { main: [], sub: [], disabledSub: [] };
  const subject = currentPublisherSubject();
  const date = state.selectedMathAnalysisDate;
  const wrapper = document.createElement("div");
  wrapper.className = `publisher-lane publisher-lane-${lane}`;
  wrapper.dataset.subject = subject;
  wrapper.dataset.date = date;
  wrapper.dataset.grade = grade;
  wrapper.dataset.lane = lane;

  const head = document.createElement("div");
  head.className = "publisher-lane-head";
  head.textContent = title;
  wrapper.appendChild(head);

  const list = document.createElement("div");
  list.className = "publisher-chip-list";
  list.dataset.grade = grade;
  list.dataset.lane = lane;
  list.addEventListener("dragover", (event) => {
    event.preventDefault();
    list.classList.add("drag-over");
  });
  list.addEventListener("dragleave", () => list.classList.remove("drag-over"));
  list.addEventListener("drop", (event) => {
    event.preventDefault();
    list.classList.remove("drag-over");
    const payload = parseDragPayload(event.dataTransfer.getData("text/plain"));
    if (!payload || payload.grade !== grade || payload.subject !== subject || payload.date !== date) return;
    moveMathPublisher(grade, payload.publisher, lane, subject, date);
  });

  const publishers = publishersForLane(config, lane);
  if (!publishers.length) {
    const empty = document.createElement("span");
    empty.className = "publisher-empty";
    empty.textContent = lane === "disabled" ? "해제 없음" : "여기로 드래그";
    list.appendChild(empty);
  }

  publishers.forEach((publisher) => {
    list.appendChild(publisherChip(grade, publisher, lane));
  });

  wrapper.appendChild(list);
  return wrapper;
}

function publisherChip(grade, publisher, lane) {
  const config = mathPublisherConfigForSelectedDate()[grade] || { disabledSub: [] };
  const subject = currentPublisherSubject();
  const date = state.selectedMathAnalysisDate;
  const chip = document.createElement("div");
  chip.className = `publisher-chip publisher-chip-${lane}`;
  chip.draggable = true;
  chip.setAttribute("role", "button");
  chip.setAttribute("tabindex", "0");
  chip.addEventListener("dragstart", (event) => {
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", JSON.stringify({ subject, date, grade, publisher, lane }));
  });

  if (lane === "sub" || lane === "disabled") {
    const isDisabled = config.disabledSub.includes(publisher);
    chip.classList.toggle("is-disabled", isDisabled);
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = !isDisabled;
    checkbox.addEventListener("click", (event) => event.stopPropagation());
    checkbox.addEventListener("change", () => {
      toggleSubPublisher(grade, publisher, checkbox.checked, subject, date);
    });
    chip.appendChild(checkbox);
  }
  const label = document.createElement("span");
  label.textContent = publisher;
  chip.appendChild(label);

  return chip;
}

function publishersForLane(config, lane) {
  if (lane === "main") return config.main || [];
  if (lane === "disabled") {
    return (config.sub || []).filter((publisher) => (config.disabledSub || []).includes(publisher));
  }
  return (config.sub || []).filter((publisher) => !(config.disabledSub || []).includes(publisher));
}

function parseDragPayload(value) {
  try {
    const payload = JSON.parse(value);
    if (!payload?.subject || !payload?.date || !payload?.grade || !payload?.publisher) return null;
    return payload;
  } catch {
    return null;
  }
}

function moveMathPublisher(grade, publisher, targetLane, subject = currentPublisherSubject(), date = state.selectedMathAnalysisDate) {
  if (!confirmPublisherPassword()) {
    showToast("비밀번호가 맞지 않아 출판사 설정 변경을 취소했습니다.");
    return;
  }

  const config = ensureGradePublisherConfig(grade, subject, date);
  config.main = config.main.filter((item) => item !== publisher);
  config.sub = config.sub.filter((item) => item !== publisher);
  config.disabledSub = config.disabledSub.filter((item) => item !== publisher);

  if (targetLane === "main") {
    config.main = unique([...config.main, publisher]).sort(localeSort);
  } else {
    config.sub = unique([...config.sub, publisher]).sort(localeSort);
    if (targetLane === "disabled") {
      config.disabledSub = unique([...config.disabledSub, publisher]).sort(localeSort);
    }
  }
  saveMathPublisherConfig({ shared: true });
  renderMathAnalysis();
}

function toggleSubPublisher(grade, publisher, enabled, subject = currentPublisherSubject(), date = state.selectedMathAnalysisDate) {
  if (!confirmPublisherPassword()) {
    showToast("비밀번호가 맞지 않아 출판사 설정 변경을 취소했습니다.");
    renderMathAnalysis();
    return;
  }

  const config = ensureGradePublisherConfig(grade, subject, date);
  if (enabled) {
    config.disabledSub = config.disabledSub.filter((item) => item !== publisher);
  } else {
    config.disabledSub = unique([...config.disabledSub, publisher]).sort(localeSort);
  }
  saveMathPublisherConfig({ shared: true });
  renderMathAnalysis();
}

function mathPublishersByGrade(rows) {
  return rows.reduce((acc, row) => {
    const grade = row["학년"] || "미분류";
    if (!acc[grade]) acc[grade] = [];
    acc[grade].push(row["출판사"] || "미분류");
    return acc;
  }, {});
}

function mathPublisherSelected(row) {
  const publisher = row["출판사"] || "미분류";
  const config = publisherConfigForRow(row);
  if (!config) return true;
  if ((config.main || []).includes(publisher)) return true;
  if ((config.sub || []).includes(publisher)) return !(config.disabledSub || []).includes(publisher);
  return false;
}

function mathPublisherDisabled(row) {
  const publisher = row["출판사"] || "미분류";
  const config = publisherConfigForRow(row);
  return Boolean(config && (config.sub || []).includes(publisher) && (config.disabledSub || []).includes(publisher));
}

function ensureMathPublisherConfig(publishersByGrade, subject = currentPublisherSubject(), date = state.selectedMathAnalysisDate) {
  const dateConfig = publisherConfigForSubjectDate(subject, date);
  Object.entries(publishersByGrade).forEach(([grade, publishers]) => {
    const normalized = unique(publishers).sort(localeSort);
    const config = ensureGradePublisherConfig(grade, subject, date);
    if (!config.main.length && !config.sub.length) {
      config.main = [];
      config.sub = [...normalized];
      config.disabledSub = [];
      return;
    }
    config.main = config.main.filter((publisher) => normalized.includes(publisher));
    config.sub = config.sub.filter((publisher) => normalized.includes(publisher));
    const assigned = new Set([...config.main, ...config.sub]);
    normalized.forEach((publisher) => {
      if (!assigned.has(publisher)) config.sub.push(publisher);
    });
    config.main.sort(localeSort);
    config.sub.sort(localeSort);
    config.disabledSub = config.disabledSub.filter((publisher) => config.sub.includes(publisher)).sort(localeSort);
  });
  Object.keys(dateConfig).forEach((grade) => {
    if (!publishersByGrade[grade]) delete dateConfig[grade];
  });
  saveMathPublisherConfig();
}

function ensureGradePublisherConfig(grade, subject = currentPublisherSubject(), date = state.selectedMathAnalysisDate) {
  const dateConfig = publisherConfigForSubjectDate(subject, date);
  if (!dateConfig[grade]) {
    dateConfig[grade] = { main: [], sub: [], disabledSub: [] };
  }
  dateConfig[grade].main ||= [];
  dateConfig[grade].sub ||= [];
  dateConfig[grade].disabledSub ||= [];
  return dateConfig[grade];
}

function readMathPublisherConfig() {
  try {
    const raw = localStorage.getItem(PUBLISHER_CONFIG_KEY) || localStorage.getItem(LEGACY_MATH_PUBLISHER_CONFIG_KEY) || "{}";
    const parsed = JSON.parse(raw);
    return normalizeMathPublisherConfig(parsed);
  } catch {
    return {};
  }
}

async function loadSharedMathPublisherConfig() {
  try {
    let response = await fetch(supabasePublicUrl(SUPABASE_PUBLISHER_CONFIG_PATH), { cache: "no-store" });
    let isLegacy = false;
    if (!response.ok) {
      response = await fetch(supabasePublicUrl(LEGACY_SUPABASE_MATH_PUBLISHER_PATH), { cache: "no-store" });
      isLegacy = true;
    }
    if (!response.ok) return false;
    const value = await response.json();
    if (!value || typeof value !== "object") return false;
    state.mathPublisherConfig = normalizeMathPublisherConfig(value, isLegacy);
    localStorage.setItem(PUBLISHER_CONFIG_KEY, JSON.stringify(state.mathPublisherConfig));
    return true;
  } catch (error) {
    console.warn("공유 출판사 설정을 불러오지 못했습니다.", error);
    return false;
  }
}

function normalizeMathPublisherConfig(config, forceLegacy = false) {
  if (!config || typeof config !== "object") return {};
  if (forceLegacy || isLegacyMathPublisherConfig(config)) {
    return { 수학: normalizePublisherDateConfig(config) };
  }
  const normalized = {};
  PUBLISHER_APPLY_SUBJECTS.forEach((subject) => {
    if (config[subject] && typeof config[subject] === "object") {
      normalized[subject] = normalizePublisherDateConfig(config[subject]);
    }
  });
  return normalized;
}

function isLegacyMathPublisherConfig(config) {
  return Object.keys(config).some((key) => /^\d{4}-\d{2}-\d{2}$/.test(key));
}

function normalizePublisherDateConfig(config) {
  const normalized = {};
  Object.entries(config).forEach(([date, value]) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
    if (!value || typeof value !== "object") return;
    normalized[date] = {};
    Object.entries(value).forEach(([grade, gradeConfig]) => {
      if (!gradeConfig || typeof gradeConfig !== "object") return;
      normalized[date][grade] = {
        main: Array.isArray(gradeConfig.main) ? gradeConfig.main : [],
        sub: Array.isArray(gradeConfig.sub) ? gradeConfig.sub : [],
        disabledSub: Array.isArray(gradeConfig.disabledSub) ? gradeConfig.disabledSub : [],
      };
    });
  });
  return normalized;
}

function mathPublisherConfigForSelectedDate() {
  return publisherConfigForSubjectDate(currentPublisherSubject(), state.selectedMathAnalysisDate);
}

function mathPublisherConfigForDate(date, subject = currentPublisherSubject()) {
  return state.mathPublisherConfig[subject]?.[date || state.selectedMathAnalysisDate] || {};
}

function publisherConfigForSubjectDate(subject, date) {
  const safeSubject = PUBLISHER_APPLY_SUBJECTS.includes(subject) ? subject : "수학";
  const safeDate = date || "__noDate";
  if (!state.mathPublisherConfig[safeSubject]) {
    state.mathPublisherConfig[safeSubject] = {};
  }
  if (!state.mathPublisherConfig[safeSubject][safeDate]) {
    state.mathPublisherConfig[safeSubject][safeDate] = {};
  }
  return state.mathPublisherConfig[safeSubject][safeDate];
}

function publisherConfigForRow(row) {
  if (!publisherConfigAppliesToRow(row)) return null;
  const subject = sortSubject(row);
  const grade = row["학년"] || "미분류";
  return mathPublisherConfigForDate(row.__openDate, subject)[grade] || null;
}

function currentPublisherSubject() {
  return PUBLISHER_APPLY_SUBJECTS.includes(state.selectedPublisherSubject)
    ? state.selectedPublisherSubject
    : "수학";
}

function saveMathPublisherConfig(options = {}) {
  localStorage.setItem(PUBLISHER_CONFIG_KEY, JSON.stringify(state.mathPublisherConfig));
  if (options.shared) {
    saveMathPublisherConfigToSupabase().catch((error) => {
      console.error(error);
      showToast(error.message || "출판사 설정 저장 중 오류가 발생했습니다.");
    });
  }
}

async function saveMathPublisherConfigToSupabase() {
  const result = await putJsonToSupabase(SUPABASE_PUBLISHER_CONFIG_PATH, mathPublisherConfigPayload());
  if (!result.ok) throw new Error(result.message);
  showToast("출판사 설정을 공유 저장했습니다.");
}

function mathSchoolworkRows() {
  const subject = currentPublisherSubject();
  return publisherAnalysisRows().filter((row) => sortSubject(row) === subject);
}

function publisherAnalysisRows() {
  return state.rows.filter((row) => (
    row.__sheet === "학교공부"
    && PUBLISHER_APPLY_SUBJECTS.includes(sortSubject(row))
    && isMeaningfulPublisher(row["출판사"] || "미분류")
  ));
}

function mathLessonGroups(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    const key = [
      row["학년"] || "",
      row[UNIT_ORDER] || "",
      normalizeText(row["단원명"]),
      row[LESSON_ORDER] || "",
    ].join("||");
    if (!groups.has(key)) {
      groups.set(key, {
        grade: row["학년"] || "",
        unitOrder: toNumber(row[UNIT_ORDER]),
        lessonOrder: toNumber(row[LESSON_ORDER]),
        unitLabel: formatMathUnit(row[UNIT_ORDER], row["단원명"]),
        lessonLabel: formatMathLesson(row[LESSON_ORDER], row["차시명"]),
        publishers: [],
      });
    }
    const group = groups.get(key);
    const publisher = row["출판사"] || "미분류";
    if (!group.publishers.includes(publisher)) group.publishers.push(publisher);
  });

  return [...groups.values()]
    .map((group) => ({
      ...group,
      publishers: sortMathPublishersForGrade(group.grade, group.publishers),
    }))
    .sort((a, b) => gradeNumber(a.grade) - gradeNumber(b.grade)
      || a.unitOrder - b.unitOrder
      || a.unitLabel.localeCompare(b.unitLabel, "ko")
      || a.lessonOrder - b.lessonOrder
      || a.lessonLabel.localeCompare(b.lessonLabel, "ko"));
}

function sortMathPublishersForGrade(grade, publishers) {
  const mainPublishers = mathPublisherConfigForSelectedDate()[grade]?.main || [];
  return [...publishers].sort((a, b) => {
    const mainOrderA = mainPublishers.includes(a) ? mainPublishers.indexOf(a) : 9999;
    const mainOrderB = mainPublishers.includes(b) ? mainPublishers.indexOf(b) : 9999;
    return mainOrderA - mainOrderB || localeSort(a, b);
  });
}

function mathPublisherListHtml(grade, publishers) {
  const config = mathPublisherConfigForSelectedDate()[grade] || { main: [], disabledSub: [] };
  const mainPublishers = config.main || [];
  const disabledPublishers = config.disabledSub || [];
  return publishers.map((publisher) => {
    const escaped = escapeHtml(publisher);
    if (mainPublishers.includes(publisher)) return `<strong class="main-publisher">${escaped}</strong>`;
    if (disabledPublishers.includes(publisher)) return `<span class="disabled-publisher">${escaped}</span>`;
    return escaped;
  }).join(", ");
}

function formatMathUnit(unitOrder, unitName) {
  const order = toNumber(unitOrder);
  const name = String(unitName || "").replace(/^\s*\d+\.\s*/, "").trim();
  return `${Number.isFinite(order) && order !== 999999 ? `${order}단원 ` : ""}${name}`.trim();
}

function formatMathLesson(lessonOrder, lessonName) {
  const order = toNumber(lessonOrder);
  return `${Number.isFinite(order) && order !== 999999 ? `${order}차시 ` : ""}${String(lessonName || "").trim()}`.trim();
}

function mathUnitKey(row) {
  return [row["학년"] || "", row[UNIT_ORDER] || "", normalizeText(row["단원명"])].join("||");
}

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function gradeNumber(grade) {
  return Number((String(grade || "").match(/\d+/) || ["999"])[0]);
}

function analysisKpi(label, count) {
  return `<span class="analysis-kpi"><strong>${count.toLocaleString("ko-KR")}</strong>${label}</span>`;
}

function renderOverallSchedule() {
  const scheduleRows = state.overallSchedule?.rows || [];
  const headers = ["카테고리", "과목", "학기", "차수", "오픈일", "콘텐츠 범위", "차시 수/비중", "비고"];
  els.overallTableHeader.innerHTML = `
    <tr class="overall-title-row">
      <th colspan="7">${escapeHtml(state.overallSchedule?.title || "스마트올 5, 6학년 2학기 콘텐츠 오픈 예정일")}</th>
      <th>${escapeHtml(state.overallSchedule?.version || "")}</th>
    </tr>
    <tr>${headers.map((header) => `<th>${header}</th>`).join("")}</tr>
  `;
  els.overallTableBody.innerHTML = "";

  const fragment = document.createDocumentFragment();
  const categorySpans = state.overallSchedule?.mergeSpans?.category || categoryRowSpans(scheduleRows);
  const savedNoteSpans = state.overallSchedule?.mergeSpans?.note || {};
  const noteSpans = Object.keys(savedNoteSpans).length ? savedNoteSpans : noteRowSpans(scheduleRows);
  scheduleRows.forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.className = `subject-${subjectTone(row.subject)}`;
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

function noteRowSpans(rows) {
  const spans = {};
  let index = 0;
  while (index < rows.length) {
    if (!rows[index].note) {
      index += 1;
      continue;
    }

    let count = 1;
    while (
      index + count < rows.length
      && !rows[index + count].note
      && rows[index + count].category === rows[index].category
      && rows[index + count].subject === rows[index].subject
    ) {
      count += 1;
    }

    spans[index] = count;
    for (let offset = 1; offset < count; offset += 1) {
      spans[index + offset] = 0;
    }
    index += count;
  }
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

function compactScheduleCategory(category) {
  return String(category || "").replace(/\n+/g, " / ");
}

function displayHeaders(headers) {
  const filtered = headers.filter((header) => {
    if (header === "비고") return false;
    if (state.selectedCategory === "수학마스터" && header === "과목차시") return false;
    return true;
  });
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
  if (OPEN_DATE_CANDIDATES.includes(header)) return OPEN_DATE_LABEL;
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

function getVisibleRows(options = {}) {
  const includePublisher = options.includePublisher !== false;
  const includeGroup = options.includeGroup !== false;
  const includeGrade = options.includeGrade !== false;
  return rowsMatchingSearch()
    .filter((row) => dateMatches(row, state.selectedDate))
    .filter((row) => state.selectedSubject === ALL_SUBJECTS || row.__subject === state.selectedSubject)
    .filter((row) => state.selectedCategory === ALL_CATEGORIES || row.__sheet === state.selectedCategory)
    .filter((row) => !includePublisher || !state.selectedTablePublisher || (row["출판사"] || "미분류") === state.selectedTablePublisher)
    .filter((row) => !includeGroup || !state.selectedTableGroup || (sortGroup(row) || "미분류") === state.selectedTableGroup)
    .filter((row) => !includeGrade || !state.selectedTableGrade || gradeValue(row) === state.selectedTableGrade)
    .sort(compareVisibleRows);
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
  ).sort(compareSubjectOrder);
}

function subjectsForDateCategory(date, category) {
  if (category === ALL_CATEGORIES) return subjectsForDate(date);
  return unique(
    rowsMatchingSearch()
      .filter((row) => dateMatches(row, date) && row.__sheet === category)
      .map((row) => row.__subject),
  ).sort(compareSubjectOrder);
}

function categoriesForDateSubject(date, subject) {
  if (subject === ALL_SUBJECTS) return categoriesForDate(date);
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

function subjectLabel(subject) {
  return subject === ALL_SUBJECTS ? "전체" : subject;
}

function categoryLabel(category) {
  return category === ALL_CATEGORIES ? "전체" : category;
}

function tableDisplayValue(row, header) {
  if (header === "과목" && state.selectedCategory === ALL_CATEGORIES) return sortSubject(row);
  if (header === "구분") return row.__sheet || "";
  if (header === OPEN_DATE_LABEL) return row.__openDate || "";
  return displayValue(row[header], header);
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
  const result = await putJsonToSupabase(SUPABASE_TIMELINE_PATH, sharedTimelinePayload());
  if (!result.ok) throw new Error(result.message);
  showToast(message);
}

async function saveOverallScheduleToSupabase() {
  const result = await putJsonToSupabase(SUPABASE_OVERALL_PATH, state.overallSchedule);
  if (!result.ok) throw new Error(result.message);
  showToast("전체 일정을 저장했습니다.");
}

function migrateTimelineMarks(value) {
  const entries = Object.entries(value || {}).filter(([key]) => !isSharedMetaKey(key));
  if (!entries.length) return {};
  const hasLegacyFlatMarks = entries.some(([, mark]) => typeof mark === "string");
  if (hasLegacyFlatMarks) return {};
  return normalizeTimelineMarks(value);
}

function normalizeTimelineMarks(value) {
  return Object.fromEntries(
    Object.entries(value || {})
      .filter(([key]) => !isSharedMetaKey(key))
      .map(([openDate, marks]) => [
        openDate,
        normalizeTimelineMarksForDate(openDate, marks),
      ])
  );
}

function isSharedMetaKey(key) {
  return String(key || "").startsWith("__");
}

function sharedTimelinePayload() {
  return normalizeTimelineMarks(state.timelineMarks);
}

function mathPublisherConfigPayload() {
  const config = clonePlainObject(state.mathPublisherConfig);
  delete config.__noDate;
  Object.values(config).forEach((subjectConfig) => {
    if (subjectConfig && typeof subjectConfig === "object") {
      delete subjectConfig.__noDate;
    }
  });
  return config;
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

function currentOpenDates() {
  return unique(state.rows.map((row) => row.__openDate)).sort();
}

function clonePlainObject(value) {
  try {
    return JSON.parse(JSON.stringify(value || {}));
  } catch {
    return {};
  }
}

function mergeTimelineMarksAfterUpload(previousMarks, previousOpenDates) {
  const nextOpenDates = currentOpenDates();
  const defaults = createDefaultTimelineMarks();
  const previousDateSet = new Set(previousOpenDates);
  const nextDateSet = new Set(nextOpenDates);
  const marks = {};

  nextOpenDates.forEach((openDate) => {
    marks[openDate] = previousMarks?.[openDate]
      ? normalizeTimelineMarksForDate(openDate, previousMarks[openDate])
      : defaults[openDate] || {};
  });

  return {
    marks,
    addedDates: nextOpenDates.filter((date) => !previousDateSet.has(date)),
    removedDates: previousOpenDates.filter((date) => !nextDateSet.has(date)),
  };
}

function notifyOpenDateChanges(addedDates, removedDates, shouldNotify) {
  if (!shouldNotify || (!addedDates.length && !removedDates.length)) return;
  const lines = ["오픈 날짜 변경이 감지되었습니다."];
  if (addedDates.length) {
    lines.push("", "[신규 날짜]", ...addedDates.map((date) => `- ${formatDateWithDay(date)}`));
  }
  if (removedDates.length) {
    lines.push("", "[삭제 날짜]", ...removedDates.map((date) => `- ${formatDateWithDay(date)}`));
  }
  window.alert(lines.join("\n"));
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

function compareVisibleRows(a, b) {
  if (state.selectedDate === ALL_DATES) {
    const dateOrder = String(a.__openDate || "").localeCompare(String(b.__openDate || ""));
    if (dateOrder) return dateOrder;
  }
  return compareRows(a, b);
}

function sortSubject(row) {
  if (usesRepresentativeUnitAsSubject(row.__sheet)) return normalizeSubjectName(row["대표단원(한글/국어)"] || row.__subject || "");
  if (row.__sheet === "성취도평가") return normalizeSubjectName(row["진입 과목명"] || row.__subject || "");
  return normalizeSubjectName(row["과목"] || row.__subject || "");
}

function sortGroup(row) {
  if (usesRepresentativeUnitAsSubject(row.__sheet) || row.__sheet === "성취도평가") return String(row["과목"] || "");
  if (row.__sheet === "수학마스터") {
    return firstMeaningfulValue([
      row["과목"],
      row["구분"],
      row["대표단원(한글/국어)"],
      row["진입 과목명"],
    ], ["수학", "AI수학", "All수학"]) || "수학마스터";
  }
  return String(row.__sheet || "");
}

function gradeValue(row) {
  const match = String(row["학년"] || "").match(/\d+/);
  return match ? match[0] : "";
}

function gradeOrder(row) {
  const number = Number(gradeValue(row) || "999");
  return Number.isFinite(number) ? number : 999;
}

async function downloadFilteredWorkbook() {
  if (!state.workbookBase64) return;
  if (!state.selectedDate || state.selectedDate === ALL_DATES) {
    showToast("오픈 날짜를 선택한 뒤 다운로드해주세요.");
    return;
  }
  const exportRows = getExportRows();
  if (!exportRows.length) return;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "5,6개정 실무 관리 툴";
  workbook.created = new Date();
  workbook.modified = new Date();

  exportRowGroups(exportRows).forEach((group) => {
    const worksheet = workbook.addWorksheet(uniqueWorksheetName(workbook, group.sheetName));
    writeExportWorksheet(worksheet, group.rows);
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const safeDate = state.selectedDate.replaceAll("-", "");
  saveBlob(
    new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `개정_차시별운영계오픈일정_${safeDate}.xlsx`,
  );
  showToast("선택한 오픈날짜 기준으로 엑셀을 저장했습니다.");
}

function exportRowGroups(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    const category = row.__sheet || normalizeSheetName(row.__sourceSheet || "") || "데이터";
    if (!groups.has(category)) {
      groups.set(category, {
        sheetName: category,
        rows: [],
      });
    }
    groups.get(category).rows.push(row);
  });
  return [...groups.values()];
}

function uniqueWorksheetName(workbook, name) {
  const base = String(name || "데이터").slice(0, 31);
  let candidate = base;
  let index = 2;
  while (workbook.getWorksheet(candidate)) {
    const suffix = `(${index})`;
    candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    index += 1;
  }
  return candidate;
}

function writeExportWorksheet(worksheet, rows) {
  const headers = exportHeaders(rows);
  worksheet.addRow(headers.map((header) => displayHeaderLabel(header, rows[0]?.__sheet || "")));
  rows.forEach((record) => {
    worksheet.addRow(headers.map((header) => exportCellValue(record, header)));
  });
  styleExportWorksheet(worksheet, headers, rows.length);
}

function exportHeaders(rows) {
  const headers = [...(rows[0]?.__headers || MIXED_TABLE_HEADERS)];
  const openDateIndex = headers.findIndex((header) => OPEN_DATE_CANDIDATES.includes(header));
  if (openDateIndex !== -1) headers[openDateIndex] = OPEN_DATE_LABEL;
  return headers;
}

function exportCellValue(record, header) {
  const value = isDateHeader(header)
    ? record[OPEN_DATE] || record[OPEN_DATE_LABEL] || record["개정 오픈일"] || record[header]
    : record[header];
  return excelValue(value, header);
}

function styleExportWorksheet(worksheet, headers, rowCount) {
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: Math.max(rowCount + 1, 1), column: headers.length },
  };

  const headerRow = worksheet.getRow(1);
  headerRow.height = 22;
  headerRow.eachCell((cell) => {
    cell.font = { name: "맑은 고딕", size: 10, bold: true };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
    cell.border = exportCellBorder();
  });

  for (let rowNumber = 2; rowNumber <= rowCount + 1; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    row.eachCell((cell, colNumber) => {
      const header = headers[colNumber - 1];
      cell.font = { name: "맑은 고딕", size: 10 };
      cell.alignment = {
        horizontal: ["단원명", "차시명"].includes(header) ? "left" : "center",
        vertical: "middle",
        wrapText: ["단원명", "차시명"].includes(header),
      };
      if (isDateHeader(header)) cell.numFmt = "yyyy-mm-dd";
      cell.border = exportCellBorder();
    });
  }

  headers.forEach((header, index) => {
    worksheet.getColumn(index + 1).width = exportColumnWidth(header);
  });
}

function exportCellBorder() {
  return {
    top: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
  };
}

function exportColumnWidth(header) {
  if (header === "차시명") return 56;
  if (header === "단원명") return 34;
  if (header === "차시고유번호") return 14;
  if (isDateHeader(header)) return 16;
  if (header === "과목차시") return 12;
  if (header === "출판사") return 12;
  if ([UNIT_ORDER, LESSON_ORDER].includes(header)) return 10;
  if (["과목", "학년", "학기"].includes(header)) return 9;
  return Math.max(10, Math.min(22, String(header || "").length + 4));
}

function getExportRows() {
  return state.rows
    .filter((row) => row.__openDate === state.selectedDate)
    .sort((a, b) => {
      const sheetOrder = state.sheets.findIndex((sheet) => sheet.name === a.__sheet)
        - state.sheets.findIndex((sheet) => sheet.name === b.__sheet);
      return sheetOrder || compareRows(a, b);
    });
}

function excelValue(value, header) {
  if (value == null || value === "") return null;
  if (isDateHeader(header)) return excelDateValue(value);
  const numberHeaders = [UNIT_ORDER, LESSON_ORDER, "차시고유번호"];
  if (numberHeaders.includes(header) && Number.isFinite(Number(value))) return Number(value);
  return value;
}

function excelDateValue(value) {
  const normalized = normalizeDate(value);
  if (!normalized) return value;
  const [year, month, day] = normalized.split("-").map(Number);
  return Math.round((Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86400000);
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

function groupBy(values, keyGetter) {
  return values.reduce((groups, value) => {
    const key = keyGetter(value);
    if (!groups[key]) groups[key] = [];
    groups[key].push(value);
    return groups;
  }, {});
}

function localeSort(a, b) {
  return String(a).localeCompare(String(b), "ko");
}

function compareSubjectOrder(a, b) {
  const aIndex = COPY_SUBJECT_ORDER.indexOf(a);
  const bIndex = COPY_SUBJECT_ORDER.indexOf(b);
  const aOrder = aIndex === -1 ? COPY_SUBJECT_ORDER.length : aIndex;
  const bOrder = bIndex === -1 ? COPY_SUBJECT_ORDER.length : bIndex;
  return aOrder - bOrder || localeSort(a, b);
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
