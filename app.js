const STORAGE_KEY = "schedule-card-workbook-v1";
const SOURCE_KEY = "schedule-card-source-v1";
const TAB_ORDER_KEY = "schedule-card-tab-order-v1";
const GITHUB_TOKEN_KEY = "schedule-card-github-token-v1";
const UPLOAD_AUTH_KEY = "schedule-card-upload-auth-v1";
const UPLOAD_PASSWORD = "610503";
const ALL_DATES = "__all_dates__";
const GITHUB_OWNER = "plusremon123-pixel";
const GITHUB_REPO = "SC";
const GITHUB_BRANCH = "main";
const GITHUB_DATA_PATH = "data/current.xlsx";
const OPEN_DATE = "운영 오픈 날짜";
const UNIT_ORDER = "단원순서";
const LESSON_ORDER = "차시순서";
const SUBJECT_CANDIDATES = ["진입 과목명", "대표단원(한글/국어)", "과목"];

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
  search: "",
};

const els = {
  fileInput: document.querySelector("#fileInput"),
  downloadButton: document.querySelector("#downloadButton"),
  searchInput: document.querySelector("#searchInput"),
  sourceMeta: document.querySelector("#sourceMeta"),
  summary: document.querySelector("#summary"),
  dateTabs: document.querySelector("#dateTabs"),
  primaryTabs: document.querySelector("#primaryTabs"),
  secondaryTabs: document.querySelector("#secondaryTabs"),
  primaryTabTitle: document.querySelector("#primaryTabTitle"),
  secondaryTabTitle: document.querySelector("#secondaryTabTitle"),
  categoryFirstButton: document.querySelector("#categoryFirstButton"),
  subjectFirstButton: document.querySelector("#subjectFirstButton"),
  tableHeader: document.querySelector("#tableHeader"),
  tableBody: document.querySelector("#tableBody"),
  tableTitle: document.querySelector("#tableTitle"),
  rowCount: document.querySelector("#rowCount"),
  emptyState: document.querySelector("#emptyState"),
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
  const sharedLoaded = await loadSharedWorkbook();
  if (!sharedLoaded && state.workbookBase64) {
    await loadWorkbook(state.workbookBase64);
  }
  bindEvents();
  render();
}

function bindEvents() {
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
    state.sourceName = `현재 데이터: ${file.name}`;
    state.currentFileName = file.name;
    showToast("현재 데이터가 교체 저장되었습니다. GitHub 저장을 진행합니다.");
    render();

    try {
      await saveCurrentWorkbookToGitHub();
    } catch (error) {
      console.error(error);
      showToast(error.message || "GitHub 저장 중 오류가 발생했습니다.");
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
}

function confirmUploadPassword() {
  if (sessionStorage.getItem(UPLOAD_AUTH_KEY) === "true") return true;

  const password = window.prompt("엑셀 업로드/교체 비밀번호를 입력해 주세요.");
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
    if (!headers.length || !headers.includes(OPEN_DATE)) return;

    const subjectHeader = SUBJECT_CANDIDATES.find((name) => headers.includes(name));
    const sheet = {
      name: worksheet.name,
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
      record.__sheet = worksheet.name;
      record.__headers = headers;
      record.__subject = normalizeSubject(record, subjectHeader, worksheet.name);
      record.__openDate = normalizeDate(record[OPEN_DATE]);
      record.__unitOrder = toNumber(record[UNIT_ORDER]);
      record.__lessonOrder = toNumber(record[LESSON_ORDER]);
      rows.push(record);
    });
  });

  state.sheets = sheets;
  state.rows = rows.filter((row) => row.__openDate);
  ensureSelections();
}

async function loadSharedWorkbook() {
  try {
    const response = await fetch(`./${GITHUB_DATA_PATH}?v=${Date.now()}`, { cache: "no-store" });
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
  if (value instanceof Date) return dateToYmd(value);
  return "";
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
  els.sourceMeta.textContent = state.sourceName
    ? `${state.sourceName} · ${state.rows.length.toLocaleString("ko-KR")}건`
    : "엑셀 업로드/교체로 현재 데이터를 불러와 주세요.";
  els.summary.textContent = state.rows.length
    ? `${unique(state.rows.map((row) => row.__openDate)).length}개 날짜 · ${unique(state.rows.map((row) => row.__subject)).length}개 과목 · ${state.sheets.length}개 구분`
    : "저장된 데이터 없음";
  renderOrderControl();
  renderDateTabs();
  renderPrimaryTabs();
  renderSecondaryTabs();
  renderTable();
}

function isCategoryFirst() {
  return state.tabOrder !== "subject-first";
}

function renderOrderControl() {
  els.categoryFirstButton.classList.toggle("active", isCategoryFirst());
  els.subjectFirstButton.classList.toggle("active", !isCategoryFirst());
  els.primaryTabTitle.textContent = isCategoryFirst() ? "구분" : "과목";
  els.secondaryTabTitle.textContent = isCategoryFirst() ? "과목" : "구분";
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
    els.dateTabs.appendChild(tabButton(date, count, date === state.selectedDate, () => {
      state.selectedDate = date;
      state.selectedCategory = "";
      state.selectedSubject = "";
      render();
    }));
  });
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

function tabButton(label, count, active, onClick, detailCount = null) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `tab${active ? " active" : ""}`;
  button.dataset.label = label;
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
      if ([UNIT_ORDER, LESSON_ORDER].includes(header)) td.className = "number";
      if (header.includes("날짜")) td.className = "date";
      tr.appendChild(td);
    });
    fragment.appendChild(tr);
  });
  els.tableBody.appendChild(fragment);
  els.emptyState.hidden = rows.length > 0;
  els.emptyState.textContent = state.rows.length ? "조건에 맞는 데이터가 없습니다." : "엑셀 파일을 업로드하면 데이터가 표시됩니다.";
}

function displayHeaders(headers) {
  const filtered = headers.filter((header) => header !== "비고");
  let next = moveBefore(filtered, "단원명", "차시명");
  if (state.selectedCategory === "성취도평가") {
    next = moveAfter(next, "진입 과목명", "학년");
  }
  return moveToFront(next, "학년");
}

function displayHeaderLabel(header, category) {
  if (category === "검정교과서") {
    if (header === "대표단원(한글/국어)") return "과목";
    if (header === "과목") return "구분";
  }
  if (category === "성취도평가") {
    if (header === "진입 과목명") return "과목";
    if (header === "과목") return "구분";
  }
  return header;
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
  return date === ALL_DATES ? "전체" : date;
}

async function saveCurrentWorkbookToGitHub() {
  if (!state.workbookBase64) {
    showToast("먼저 엑셀을 업로드해 주세요.");
    return;
  }

  let token = localStorage.getItem(GITHUB_TOKEN_KEY);
  let isSavedToken = Boolean(token);
  if (!token) token = window.prompt("GitHub 저장을 위해 토큰을 입력해 주세요. 입력한 토큰은 이 브라우저에 저장됩니다.");
  if (!token) {
    showToast("브라우저에는 저장되었습니다. 공유하려면 GitHub 저장을 완료해 주세요.");
    return;
  }
  localStorage.setItem(GITHUB_TOKEN_KEY, token);

  let result = await putWorkbookToGitHub(token);
  if ((result.status === 401 || result.status === 403) && isSavedToken) {
    localStorage.removeItem(GITHUB_TOKEN_KEY);
    token = window.prompt("저장된 GitHub 토큰이 만료되었거나 권한이 없습니다. 새 토큰을 입력해 주세요.");
    if (!token) {
      throw new Error("GitHub 저장 실패: 새 토큰이 입력되지 않았습니다.");
    }
    localStorage.setItem(GITHUB_TOKEN_KEY, token);
    result = await putWorkbookToGitHub(token);
  }

  if (!result.ok) {
    if (result.status === 401 || result.status === 403) localStorage.removeItem(GITHUB_TOKEN_KEY);
    throw new Error(result.message);
  }

  state.sourceName = `공유 데이터: ${state.currentFileName || "current.xlsx"}`;
  localStorage.setItem(SOURCE_KEY, state.sourceName);
  showToast("GitHub에 공유 엑셀을 저장했습니다. 잠시 후 다른 사용자에게도 반영됩니다.");
  render();
}

async function putWorkbookToGitHub(token) {
  const apiBase = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${GITHUB_DATA_PATH}`;
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
  };

  let sha;
  const existing = await fetch(`${apiBase}?ref=${GITHUB_BRANCH}`, { headers });
  if (existing.ok) {
    const payload = await existing.json();
    sha = payload.sha;
  } else if (existing.status !== 404) {
    return {
      ok: false,
      status: existing.status,
      message: `GitHub 파일 확인 실패: ${existing.status}`,
    };
  }

  const messageName = state.currentFileName || "current.xlsx";
  const saveResponse = await fetch(apiBase, {
    method: "PUT",
    headers: {
      ...headers,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      branch: GITHUB_BRANCH,
      message: `Update shared workbook: ${messageName}`,
      content: state.workbookBase64,
      sha,
    }),
  });

  if (!saveResponse.ok) {
    const text = await saveResponse.text();
    return {
      ok: false,
      status: saveResponse.status,
      message: `GitHub 저장 실패: ${saveResponse.status} ${text}`,
    };
  }

  return { ok: true, status: saveResponse.status };
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
  if (row.__sheet === "검정교과서") return String(row["대표단원(한글/국어)"] || row.__subject || "");
  if (row.__sheet === "성취도평가") return String(row["진입 과목명"] || row.__subject || "");
  return String(row["과목"] || row.__subject || "");
}

function sortGroup(row) {
  if (row.__sheet === "검정교과서" || row.__sheet === "성취도평가") return String(row["과목"] || "");
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
    if (!headers.includes(OPEN_DATE)) return;
    const sheetRows = exportRows.filter((row) => row.__sheet === worksheet.name);
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
  if (header.includes("날짜")) return Number(String(value).replace(/[^\d]/g, ""));
  const numberHeaders = [UNIT_ORDER, LESSON_ORDER, "차시고유번호"];
  if (numberHeaders.includes(header) && Number.isFinite(Number(value))) return Number(value);
  return value;
}

function displayValue(value, header) {
  if (header.includes("날짜")) return normalizeDate(value) || value || "";
  return value ?? "";
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
