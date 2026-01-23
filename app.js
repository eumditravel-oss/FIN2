/* app.js (FINAL FIX v13.2b+) - FIN 산출자료 (Web)
   - ✅ (v13.0) 내보내기/가져오기: JSON → Excel(.xlsx) 기반으로 변경
   - ✅ (v13.0) 내보내기 클릭 시 탭 선택 팝업(모달) 제공 (코드/철골/철골_부자재/구조이기-동바리)
   - ✅ (v13.0) 가져오기(Excel): Codes 시트 기반으로 codeMaster 갱신 (임시 양식)
   - ✅ (v12.4) 산출표(계산표)에서 "비고" 컬럼만 숨김(렌더링 제거)
   - ✅ (v12.3) 변수표 영역에서도 Ctrl+F3/Shift+Ctrl+F3 행추가 지원 (변수표 셀 선택 시)
   - ✅ (v12.3) 집계 탭: 구분 개소(count) 반영하여 코드별 수량 합산
   - ✅ (v12.3) 집계 탭: 환산단위/환산계수 있으면 환산후수량 기준으로 단위/할증전/후 집계
   - ✅ (v12.3) 산출표 헤더 "물량(Value)" -> "물량"
   - ✅ (v12.3) 산출표 컬럼폭: 단위/물량(및 코드) 가로폭 증가 (CALC_COL_WEIGHTS 조정)
   - ✅ (v13.1) 도움말 버튼 추가: 화면 안내문구 제거 + help.html로 이동
   - ✅ (v13.2) 구분명 리스트: 클릭 후에도 ↑/↓ 키로 이동 가능(렌더 후 포커스 복원)
   - ✅ (v13.2a) 내보내기 모달 '전체선택' 버튼이 실제 체크박스에 반영되도록 수정(모달 재오픈 제거)
   - ✅ (v13.2b) top-split(구분/변수) ↔ panel 사이 리사이저(split-resizer) 적용 + 높이 상태 저장(ui.topSplitH)
   - ✅ (v13.2b) section-editor(구분 편집) CSS(3컬럼)와 맞게 버튼들을 한 칸으로 묶음

   - 🛠 (Patch) LS_KEY 버전 분리 + 구버전(V11) 데이터 자동 마이그레이션 + 초기화 시 구키도 함께 삭제
   - 🛠 (Patch) 프로젝트 모달 show/hide: hidden + aria-hidden 동시 지원(접근성/표준)
   - 🛠 (Patch) Init/Render 중복 호출 제거, bindTopButtons 1회만 바인딩
*/

(() => {
  "use strict";

  /***************
   * Storage (✅ Project-ready)
   ***************/
  const PROJECT_INDEX_KEY = "FIN_PROJECT_INDEX_V1";
  const PROJECT_ACTIVE_KEY = "FIN_PROJECT_ACTIVE_V1";
  const PROJECT_STATE_PREFIX = "FIN_PROJECT_STATE_V1::";

  // (기존 단일 저장키 마이그레이션용)
  const LS_KEY_OLD_SINGLE_V13 = "FIN_WEB_STATE_V13_2A";
  const LS_KEY_OLD_SINGLE_V11 = "FIN_WEB_STATE_V11";

  const deepClone = (obj) => JSON.parse(JSON.stringify(obj));
  const clamp = (n, a, b) => Math.max(a, Math.min(b, n));

  /***************
   * ✅ focus jump 방지 헬퍼
   ***************/
  function safeFocus(target) {
    if (!target) return;
    try {
      target.focus({ preventScroll: true });
    } catch {
      try { target.focus(); } catch {}
    }
  }

  function raf2(fn) {
    requestAnimationFrame(() => requestAnimationFrame(fn));
  }

  /***************
   * Sticky height auto-measure
   ***************/
  function updateStickyVars() {
    const root = document.documentElement;

    const topbar = document.querySelector(".topbar");
    const tabs = document.querySelector(".tabs");
    const topSplit = document.querySelector(".top-split"); // 산출탭에서만 존재

    const topbarH = topbar ? topbar.getBoundingClientRect().height : 0;
    const tabsH = tabs ? tabs.getBoundingClientRect().height : 0;
    const topSplitH = topSplit ? topSplit.getBoundingClientRect().height : 0;

    root.style.setProperty("--topbarH", `${Math.ceil(topbarH)}px`);
    root.style.setProperty("--tabsH", `${Math.ceil(tabsH)}px`);
    root.style.setProperty("--topSplitActualH", `${Math.ceil(topSplitH)}px`);

    const base = Math.ceil(topbarH + tabsH);
    root.style.setProperty("--stickyBaseTop", `${base}px`);

    const withTopSplit = Math.ceil(topbarH + tabsH + topSplitH + 10);
    root.style.setProperty("--stickyWithTopSplitTop", `${withTopSplit}px`);
  }

  window.addEventListener("resize", () => {
    requestAnimationFrame(() => {
      updateStickyVars();
      applyPanelStickyTop();
      updateScrollHeights();
      updateViewFillHeight();
    });
  });

  /***************
   * ✅ 내부 스크롤 높이 자동 보정 (PATCH: 하단 공백 제거)
   ***************/
  function updateScrollHeights() {
  const scrolls = document.querySelectorAll(".calc-scroll");
  if (!scrolls.length) return;

  scrolls.forEach((sc) => {
    if (!(sc instanceof HTMLElement)) return;

    // ✅ 최근에 스크롤 중이면(사용자 조작 중) 강제 height 갱신을 건너뛰어
    // 클릭 hit-test / 포커스 튐을 줄임
    const now = Date.now();
    const last = Number(sc.__lastScrollAt || 0);
    if (now - last < 120) return;

    sc.style.overflow = "auto";
    sc.style.webkitOverflowScrolling = "touch";
    sc.tabIndex = -1;

    // ✅ 중요: flex 컨테이너 안에서 스크롤 영역이 제대로 줄어들도록
    sc.style.minHeight = "0";

    const scRect = sc.getBoundingClientRect();
    const viewportH = window.innerHeight || document.documentElement.clientHeight || 800;

    const bottomPad = 12;

    const panel = sc.closest(".panel");
    let h = 0;

    if (panel instanceof HTMLElement) {
      const panelRect = panel.getBoundingClientRect();

      // ✅ 패널이 화면 밖으로 내려가도, 계산 기준은 "뷰포트 바닥"까지만
      const bottom = Math.min(panelRect.bottom, viewportH);

      h = Math.floor(bottom - scRect.top - bottomPad);
    } else {
      h = Math.floor(viewportH - scRect.top - bottomPad);
    }

    h = clamp(h, 160, 20000);

    sc.style.maxHeight = "";
    sc.style.height = `${h}px`;
  });
}


     /***************
 * ✅ REMARK(비고) 고정코드/행 규칙
 ***************/
const REMARK_CODE = "ZZZZZZZZZZZZZZZZZ";
const REMARK_NAME = "[비          고]";

function normalizeRemarkName(s) {
  return String(s || "").trim().replace(/\s+/g, " ");
}

function isRemarkCode(code) {
  return String(code || "").trim().toUpperCase() === REMARK_CODE.toUpperCase();
}

function isRemarkRowObj(r) {
  return (
    isRemarkCode(r?.code) ||
    normalizeRemarkName(r?.name) === normalizeRemarkName(REMARK_NAME)
  );
}

// ✅ 비고 코드마스터 행 생성
function getRemarkCodeMasterRow() {
  return {
    code: REMARK_CODE,
    name: REMARK_NAME,
    spec: "",
    unit: "",
    surcharge: null,
    convUnit: "",
    convFactor: null,
    note: ""
  };
}

// ✅ codeMaster 최상단에 비고 고정코드 강제 + 중복 제거
function ensureRemarkCodeMasterTop() {
  if (!state || !Array.isArray(state.codeMaster)) state.codeMaster = [];
  state.codeMaster = state.codeMaster.filter(r => !isRemarkCode(r?.code));
  state.codeMaster.unshift(getRemarkCodeMasterRow());
}


   /***************
 * ✅ Z 5개 이상 코드 판단 (행 회색 처리용)
 ***************/
function hasAtLeastFiveZ(code) {
  // "Z"가 연속 5개 이상 포함되면 true (대/소문자 무시)
  return /Z{5,}/i.test(String(code || "").trim());
}





  /***************
   * Code Master
   ***************/
  const DEFAULT_CODE_MASTER = [
    {"code":"A0SM355150","name":"RH형강 / SM355","spec":"150*150*7*10","unit":"M","surcharge":7,"convUnit":"TON","convFactor":0.0315,"note":""},
    {"code":"A0SM355200","name":"RH형강 / SM355","spec":"200*100*5.5*8","unit":"M","surcharge":7,"convUnit":"TON","convFactor":0.0213,"note":""},
    {"code":"A0SM355201","name":"RH형강 / SM355","spec":"200*200*8*12","unit":"M","surcharge":null,"convUnit":"","convFactor":null,"note":""},
    {"code":"A0SM355294","name":"RH형강 / SM355","spec":"294*200*8*12","unit":"M","surcharge":null,"convUnit":"","convFactor":null,"note":""},
    {"code":"A0SM355300","name":"RH형강 / SM355","spec":"300*300*10*15, CAMBER 35mm","unit":"M","surcharge":null,"convUnit":"","convFactor":null,"note":""},

    {"code":"B0SM355800","name":"BH형강 / SM355","spec":"800*300*25*40","unit":"M","surcharge":10,"convUnit":"TON","convFactor":0.3297,"note":""},
    {"code":"B0SM355900","name":"BH형강 / SM355","spec":"900*350*30*60","unit":"M","surcharge":10,"convUnit":"TON","convFactor":0.35796,"note":""},

    {"code":"C0SS275009","name":"강판 / SS275","spec":"9mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275010","name":"강판 / SS275","spec":"10mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275011","name":"강판 / SS275","spec":"11mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275012","name":"강판 / SS275","spec":"12mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275013","name":"강판 / SS275","spec":"13mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275014","name":"강판 / SS275","spec":"14mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275025","name":"강판 / SS275","spec":"25mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355009","name":"강판 / SM355","spec":"9mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355010","name":"강판 / SM355","spec":"10mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355011","name":"강판 / SM355","spec":"11mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355012","name":"강판 / SM355","spec":"12mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355013","name":"강판 / SM355","spec":"13mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355014","name":"강판 / SM355","spec":"14mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355025","name":"강판 / SM355","spec":"25mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
  ];

  /***************
   * Tabs
   ***************/
  const TABS = [
    { id: "code", title: "코드(Ctrl+.)" },
    { id: "steel", title: "철골" },
    { id: "steel_sum", title: "철골_집계" },
    { id: "steel_sub", title: "철골_부자재" },
    { id: "support", title: "구조이기/동바리" },
    { id: "support_sum", title: "구조이기/동바리_집계" },
  ];

  /***************
   * Default State
   ***************/
  const defaultCalcRow = () => ({
    code: "",
    name: "",
    spec: "",
    unit: "",
    formula: "",
    value: 0,
    surchargePct: null,
    surchargeMul: 1,
    convUnit: "",
    convFactor: null,
    converted: 0,
    note: "",
  });

  const defaultVarRow = () => ({
    key: "",
    expr: "",
    value: 0,
    note: "",
  });

  const defaultSection = (name = "구분 1", count = 1) => ({
    name,
    count,
    vars: Array.from({ length: 12 }, () => defaultVarRow()),
    rows: Array.from({ length: 12 }, () => defaultCalcRow()),
  });

  const DEFAULT_STATE = {
    activeTab: "code",
    codeMaster: deepClone(DEFAULT_CODE_MASTER),
    steel: { activeSection: 0, sections: [defaultSection("구분 1", 1)] },
    steel_sub: { activeSection: 0, sections: [defaultSection("구분 1", 1)] },
    support: { activeSection: 0, sections: [defaultSection("구분 1", 1)] },

    ui: {
      topSplitH: 190,
    }
  };

  /***************
   * ✅ Project Store Adapter
   ***************/
  const ProjectStore = (() => {
    const local = {
      loadIndex() {
        try {
          const raw = localStorage.getItem(PROJECT_INDEX_KEY);
          const parsed = raw ? JSON.parse(raw) : null;
          if (!parsed || !Array.isArray(parsed.projects)) return { projects: [] };
          return parsed;
        } catch { return { projects: [] }; }
      },
      saveIndex(index) {
        localStorage.setItem(PROJECT_INDEX_KEY, JSON.stringify(index));
      },
      loadActiveId() {
        return localStorage.getItem(PROJECT_ACTIVE_KEY) || "";
      },
      saveActiveId(id) {
        if (!id) localStorage.removeItem(PROJECT_ACTIVE_KEY);
        else localStorage.setItem(PROJECT_ACTIVE_KEY, id);
      },
      loadProjectState(id) {
        try {
          const k = PROJECT_STATE_PREFIX + id;
          const raw = localStorage.getItem(k);
          return raw ? JSON.parse(raw) : null;
        } catch {
          return null;
        }
      },
      saveProjectState(id, projectState) {
        const k = PROJECT_STATE_PREFIX + id;
        localStorage.setItem(k, JSON.stringify(projectState));
      },
      deleteProject(id) {
        localStorage.removeItem(PROJECT_STATE_PREFIX + id);
      }
    };
    return local;
  })();

  function genId() {
    return "p_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
  }

  function normalizeProjectMeta(p) {
    return {
      id: String(p?.id || genId()),
      name: String(p?.name || "새 프로젝트"),
      code: String(p?.code || ""),
      updatedAt: Number(p?.updatedAt || Date.now()),
      createdAt: Number(p?.createdAt || Date.now()),
    };
  }

  function loadProjectIndex() {
    const idx = ProjectStore.loadIndex();
    return { projects: Array.isArray(idx.projects) ? idx.projects.map(normalizeProjectMeta) : [] };
  }

  function saveProjectIndex(index) {
    ProjectStore.saveIndex(index);
  }

  function loadProjectState(projectId) {
    try {
      const parsed = ProjectStore.loadProjectState(projectId);
      if (!parsed) return deepClone(DEFAULT_STATE);

      const s = { ...deepClone(DEFAULT_STATE), ...parsed };
      s.codeMaster = Array.isArray(parsed?.codeMaster) ? parsed.codeMaster : deepClone(DEFAULT_CODE_MASTER);

      for (const k of ["steel", "steel_sub", "support"]) {
        if (!s[k] || !Array.isArray(s[k].sections) || s[k].sections.length === 0) {
          s[k] = deepClone(DEFAULT_STATE[k]);
        }
        s[k].activeSection = clamp(Number(s[k].activeSection || 0), 0, s[k].sections.length - 1);
      }

      if (!s.ui || typeof s.ui !== "object") s.ui = deepClone(DEFAULT_STATE.ui);
      s.ui.topSplitH = clamp(Number(s.ui.topSplitH ?? 190), 120, 520);

      if (!TABS.some(t => t.id === s.activeTab)) s.activeTab = "code";
      return s;
    } catch (e) {
      console.warn("loadProjectState failed:", e);
      return deepClone(DEFAULT_STATE);
    }
  }

  function saveProjectState(projectId) {
    if (!projectId) return;
    ProjectStore.saveProjectState(projectId, deepClone(state));
  }

  // ✅ activeProjectId가 준비되기 전 호출 방지 포함
  function saveState() {
    if (!activeProjectId) return;
    saveProjectState(activeProjectId);
  }

  let projectIndex = loadProjectIndex();
  let activeProjectId = ProjectStore.loadActiveId();

  /***************
   * ✅ Legacy migration(단일키 -> 프로젝트 1회 이관)
   ***************/
  (function migrateLegacySingleToProjectOnce() {
    const legacy = localStorage.getItem(LS_KEY_OLD_SINGLE_V13) || localStorage.getItem(LS_KEY_OLD_SINGLE_V11);
    if (!legacy) return;
    if (projectIndex.projects.length > 0) return;

    try {
      const parsed = JSON.parse(legacy);
      const pid = genId();
      const meta = normalizeProjectMeta({ id: pid, name: "마이그레이션 프로젝트", code: "LEGACY" });
      projectIndex.projects.push(meta);
      saveProjectIndex(projectIndex);
      ProjectStore.saveActiveId(pid);
      activeProjectId = pid;

      ProjectStore.saveProjectState(pid, { ...deepClone(DEFAULT_STATE), ...parsed });
    } catch {}
  })();

  (function cleanupLegacyKeys() {
    if (projectIndex.projects.length <= 0) return;
    try { localStorage.removeItem(LS_KEY_OLD_SINGLE_V13); } catch {}
    try { localStorage.removeItem(LS_KEY_OLD_SINGLE_V11); } catch {}
  })();

  (function ensureAtLeastOneProject() {
    if (projectIndex.projects.length > 0) {
      if (!activeProjectId || !projectIndex.projects.some(p => p.id === activeProjectId)) {
        activeProjectId = projectIndex.projects[0].id;
        ProjectStore.saveActiveId(activeProjectId);
      }
      return;
    }

    const pid = genId();
    const meta = normalizeProjectMeta({ id: pid, name: "프로젝트 1", code: "" });
    projectIndex.projects.push(meta);
    saveProjectIndex(projectIndex);

    activeProjectId = pid;
    ProjectStore.saveActiveId(activeProjectId);
    ProjectStore.saveProjectState(pid, deepClone(DEFAULT_STATE));
  })();

  let state = activeProjectId ? loadProjectState(activeProjectId) : deepClone(DEFAULT_STATE);
// ✅ 항상 비고 고정코드 최상단 강제
ensureRemarkCodeMasterTop();
saveState();

   /***************
 * ✅ 비고행(ZZZZ...) 스타일 동기화
 * - Ctrl+F10으로 비고코드 넣은 직후 행 클래스 반영
 ***************/
function syncRemarkRowFromCodeInput(codeInput) {
  if (!(codeInput instanceof HTMLInputElement)) return;

  const tabId = codeInput.dataset.tab;
  const row = Number(codeInput.dataset.row || 0);
  if (!tabId) return;

  const table = codeInput.closest("table.calc-table");
  if (!table) return;

  const tr = table.querySelectorAll("tbody tr")[row];
  if (!tr) return;

  const code = String(codeInput.value || "").trim();
  const isRemark = isRemarkCode(code);

  // 비고행이면 z5-row도 같이 켜고(회색 스타일 재사용)
  tr.classList.toggle("z5-row", isRemark || hasAtLeastFiveZ(code));

  // 원하면 전용 클래스도 추가로 사용 가능
  tr.classList.toggle("remark-calc-row", isRemark);
}




  // ✅ (v13.2) 구분명 리스트 클릭/↑↓ 후 렌더링되면 포커스 복원
  let __pendingSectionFocus = null;

  /***************
   * ✅ Calc(산출표) 멀티선택 상태 (비저장/런타임)
   ***************/
  const __calcMulti = {
    active: false,
    tabId: null,
    sectionIndex: null,
    anchorRow: null,
    rows: new Set(),
  };

  function __calcMultiClear() {
    __calcMulti.active = false;
    __calcMulti.tabId = null;
    __calcMulti.sectionIndex = null;
    __calcMulti.anchorRow = null;
    __calcMulti.rows.clear();
  }

  function __calcMultiIsSameContext(tabId) {
    const bucket = state?.[tabId];
    const secIdx = bucket?.activeSection ?? 0;
    return __calcMulti.active && __calcMulti.tabId === tabId && __calcMulti.sectionIndex === secIdx;
  }

  function __calcMultiBegin(tabId, anchorRow) {
    const bucket = state?.[tabId];
    const secIdx = bucket?.activeSection ?? 0;

    __calcMulti.active = true;
    __calcMulti.tabId = tabId;
    __calcMulti.sectionIndex = secIdx;
    __calcMulti.anchorRow = clamp(
      Number(anchorRow || 0),
      0,
      (bucket?.sections?.[secIdx]?.rows?.length ?? 1) - 1
    );

    __calcMulti.rows.clear();
    __calcMulti.rows.add(__calcMulti.anchorRow);
  }

  function __calcMultiSetRange(tabId, fromRow, toRow) {
    if (!__calcMultiIsSameContext(tabId)) {
      __calcMultiBegin(tabId, fromRow);
    }
    const a = __calcMulti.anchorRow ?? fromRow;
    const lo = Math.min(a, toRow);
    const hi = Math.max(a, toRow);

    __calcMulti.rows.clear();
    for (let r = lo; r <= hi; r++) __calcMulti.rows.add(r);
  }

  function __applyCalcRowSelectionStyles(tabId) {
    const table = document
      .querySelector(`table.calc-table input[data-grid="calc"][data-tab="${tabId}"]`)
      ?.closest("table.calc-table");
    if (!table) return;

    const should = __calcMultiIsSameContext(tabId);
    const trs = table.querySelectorAll("tbody tr");
    trs.forEach((tr, i) => {
      if (should && __calcMulti.rows.has(i)) tr.classList.add("row-selected");
      else tr.classList.remove("row-selected");
    });
  }

  function __getSelectedCalcRows(tabId) {
    if (!__calcMultiIsSameContext(tabId)) return [];
    return [...__calcMulti.rows].sort((a, b) => a - b);
  }


   /***************
 * ✅ Z 5개 이상 행 스타일 적용
 ***************/
function __applyZ5RowStyles(tabId) {
  // 해당 탭의 calc-table 찾기
  const table = document
    .querySelector(`table.calc-table input[data-grid="calc"][data-tab="${tabId}"]`)
    ?.closest("table.calc-table");
  if (!table) return;

  const trs = table.querySelectorAll("tbody tr");
  trs.forEach((tr, i) => {
    const codeInput = tr.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${i}"][data-field="code"]`);
    const code = codeInput ? codeInput.value : "";
    tr.classList.toggle("z5-row", hasAtLeastFiveZ(code));
  });
}



   // ✅ [추가] Ctrl+B에서 "현재 행 선택 토글"을 만들기 위한 함수
function __calcMultiToggleRow(tabId, row) {
  const bucket = state?.[tabId];
  const secIdx = bucket?.activeSection ?? 0;
  const maxRow = (bucket?.sections?.[secIdx]?.rows?.length ?? 1) - 1;
  const r = clamp(Number(row || 0), 0, Math.max(0, maxRow));

  // 컨텍스트 다르면 시작(=첫 선택은 anchor로)
  if (!__calcMultiIsSameContext(tabId)) {
    __calcMultiBegin(tabId, r);
    return;
  }

  // 같은 컨텍스트면 현재 row 토글
  if (__calcMulti.rows.has(r)) __calcMulti.rows.delete(r);
  else __calcMulti.rows.add(r);

  // 전부 해제됐으면 모드도 종료
  if (__calcMulti.rows.size === 0) __calcMultiClear();
}

   

  /***************
   * DOM
   ***************/
  const $tabs = document.getElementById("tabs");
  const $view = document.getElementById("view");

  function el(tag, attrs = {}, children = []) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === "class") node.className = v;
      else if (k === "dataset") Object.assign(node.dataset, v);
      else if (k.startsWith("on") && typeof v === "function") node.addEventListener(k.slice(2), v);
      else if (v === false || v == null) continue;
      else node.setAttribute(k, String(v));
    }
    for (const ch of children) {
      if (ch == null) continue;
      node.appendChild(typeof ch === "string" ? document.createTextNode(ch) : ch);
    }
    return node;
  }

  function clear(node) {
    while (node.firstChild) node.removeChild(node.firstChild);
  }

  /***************
   * ✅ (v13.2b) topSplit height 적용
   ***************/
  function applyTopSplitH() {
    const root = document.documentElement;
    const h = clamp(Number(state?.ui?.topSplitH ?? 190), 120, 520);
    root.style.setProperty("--topSplitH", `${Math.round(h)}px`);
  }

  /***************
   * ✅ zoom(--uiScale) 대응: view 높이 보정
   ***************/
  function getUiScale() {
    const v = getComputedStyle(document.documentElement).getPropertyValue("--uiScale").trim();
    const n = Number(v);
    return (Number.isFinite(n) && n > 0.2 && n < 2.5) ? n : 1;
  }

  function updateViewFillHeight() {
  const view = document.getElementById("view");
  if (!view) return;

  const scale = getUiScale();
  const vh = window.innerHeight || document.documentElement.clientHeight || 800;

  const topbar = document.querySelector(".topbar");
  const tabs = document.querySelector(".tabs");

  const topbarH = topbar ? topbar.getBoundingClientRect().height : 0;
  const tabsH = tabs ? tabs.getBoundingClientRect().height : 0;

  // ✅ view는 "상단(topbar+tabs) 제외한 나머지"만 차지해야 함
  const available = Math.max(200, vh - Math.ceil(topbarH + tabsH));

  // ✅ uiScale이 transform 기반이면 실제 px로 맞추기 위해 scale로 보정
  const target = Math.ceil(available / scale);

  view.style.height = `${target}px`;
  view.style.minHeight = `${target}px`;
}


  /***************
   * Helpers: Code master lookup
   ***************/
  function codeLookup(code) {
    const c = String(code || "").trim();
    if (!c) return null;
    return state.codeMaster.find(x => String(x.code).trim().toUpperCase() === c.toUpperCase()) || null;
  }

  /***************
   * Expression evaluator
   ***************/
  function stripAngleComments(expr) {
    if (!expr) return "";
    return String(expr).replace(/<[^>]*>/g, "");
  }

  function safeEvalWithVars(expr, varMap) {
    const raw = String(expr || "").trim();
    if (!raw) return 0;

    const replaced = raw.replace(/\b([A-Za-z][A-Za-z0-9]{0,2})\b/g, (m, p1) => {
      const k = p1.toUpperCase();
      if (Object.prototype.hasOwnProperty.call(varMap, k)) return String(varMap[k] ?? 0);
      return "0";
    });

    const cleaned = replaced.replace(/\s+/g, "");
    if (!/^[0-9+\-*/().]*$/.test(cleaned)) return NaN;

    try {
      // eslint-disable-next-line no-new-func
      const fn = new Function(`return (${replaced});`);
      const v = fn();
      const n = Number(v);
      return Number.isFinite(n) ? n : NaN;
    } catch {
      return NaN;
    }
  }

  function buildVarMap(section) {
    const map = Object.create(null);

    for (const v of section.vars) {
      const key = (v.key || "").trim();
      if (!key) continue;
      map[key.toUpperCase()] = 0;
    }

    for (let pass = 0; pass < 6; pass++) {
      for (const v of section.vars) {
        const key = (v.key || "").trim();
        if (!key) continue;

        const exprRaw = stripAngleComments(v.expr || "");
        const val = safeEvalWithVars(exprRaw, map);
        if (Number.isFinite(val)) map[key.toUpperCase()] = val;
      }
    }

    for (const v of section.vars) {
      const key = (v.key || "").trim();
      if (!key) v.value = 0;
      else v.value = Number(map[key.toUpperCase()] ?? 0) || 0;
    }
    return map;
  }

  function recomputeSection(tabId) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];
    const varMap = buildVarMap(sec);

    for (const r of sec.rows) {
      const info = codeLookup(r.code);
      if (info) {
        r.name = info.name || "";
        r.spec = info.spec || "";
        r.unit = info.unit || "";
        r.note = info.note || "";
        r.convUnit = info.convUnit || "";
        r.convFactor = info.convFactor ?? null;

        const sPct = (r.surchargePct == null || r.surchargePct === "") ? (info.surcharge ?? null) : Number(r.surchargePct);
        r.surchargePct = (sPct == null || sPct === "") ? null : Number(sPct);
      } else {
        r.name = r.name || "";
        r.spec = r.spec || "";
        r.unit = r.unit || "";
        r.note = r.note || "";
        r.convUnit = r.convUnit || "";
      }

      const expr = stripAngleComments(r.formula || "");
      const base = safeEvalWithVars(expr, varMap);
      r.value = Number.isFinite(base) ? base : 0;

      const pct = (r.surchargePct == null || r.surchargePct === "") ? null : Number(r.surchargePct);
      const mul = pct == null || !Number.isFinite(pct) ? 1 : (1 + pct / 100);
      r.surchargeMul = mul;

      const after = r.value * mul;
      const cf = r.convFactor;
      if (cf != null && Number.isFinite(Number(cf)) && Number(cf) !== 0) r.converted = after * Number(cf);
      else r.converted = after;
    }
  }

  /***************
   * Column width helpers
   ***************/
  function buildColGroupFromWeights(weights) {
    const sum = weights.reduce((a, b) => a + b, 0);
    const cg = el("colgroup", {}, []);
    weights.forEach((w) => {
      const pct = (w / sum) * 100;
      cg.appendChild(el("col", { style: `width:${pct.toFixed(3)}%` }, []));
    });
    return cg;
  }

  const CALC_COL_WEIGHTS = [
    0.35,  // No
    0.75,  // 코드
    2.5,   // 품명(자동)
    2.5,   // 규격(자동)
    0.50,  // 단위(자동)
    2.5,   // 산출식
    0.50,  // 물량
    0.25,  // 할증(%)
    0.25,  // 환산단위
    0.25,  // 환산계수
    0.25,  // 환산후수량
  ];

  const CODE_COL_WEIGHTS = [0.6, 2.2, 2.2, 0.6, 0.6, 0.7, 0.7, 1.2, 0.6];

// codeMaster(코드표) 열 번호(0~8)
const CODE_COL_INDEX = {
  code: 0,
  name: 1,
  spec: 2,
  unit: 3,
  surcharge: 4,
  convUnit: 5,
  convFactor: 6,
  note: 7,
  action: 8,
};

// calc(산출표) 실제 열 번호(No 포함 0~10)
const CALC_COL_INDEX = {
  code: 1,
  name: 2,
  spec: 3,
  unit: 4,
  formula: 5,
  value: 6,
  surchargePct: 7,
  convUnit: 8,
  convFactor: 9,
  converted: 10,
};


  /***************
   * ✅ Help
   ***************/
  function buildHelpPayload() {
    return {
      title: "FIN 산출자료 도움말",
      sections: [
        { title: "코드 선택(팝업)", items: [
          "Ctrl+. : 코드 선택 창 열기",
          "코드 선택 창에서 Ctrl+B : 다중선택",
          "코드 선택 창에서 Ctrl+Enter : 삽입",
        ]},

        { title: "표 이동/편집(공통)", items: [
          "방향키: 셀 이동",
          "F2: 편집 모드(읽기전용 셀 제외)",
          "편집 모드에서 Enter: 편집 종료",
          "PageUp / PageDown: 한 페이지 단위로 위/아래 이동(현재 열 유지)",
          "Ctrl+Home / Ctrl+End: 최상단/최하단으로 이동(현재 열 유지)"
        ]},
        { title: "행 추가/삭제", items: [
          "Ctrl+F3: 현재 행 아래 행 추가",
          "Shift+Ctrl+F3: +10행 추가",
          "Ctrl+Del: 삭제(확인창) - 산출표/코드표는 현재 '행' 삭제, 변수표는 현재 '셀' 비움",
          "ESC: (산출표 다중선택 중) 다중선택 취소"
        ]},
        { title: "산출 탭", items: [
          "구분 리스트: ↑/↓ 로 이동 및 선택",
          "구분/변수 영역 높이 조절: 중간 점선 바(리사이저)를 드래그"
        ]},
        { title: "산출표 다중선택", items: [
          "Shift+B: 다중선택 모드 토글",
          "Shift+↑ / Shift+↓: 다중선택 범위 확장",
          "Ctrl+Del: (다중선택 중) 선택된 행들을 한 번에 삭제",
          "Ctrl+G: (다중선택 중) 선택된 행들을 현재 행 아래로 복사/삽입"
        ]},
        { title: "엑셀 내보내기/가져오기", items: [
          "내보내기(EXCEL): 선택 모달에서 탭 선택 후 .xlsx 다운로드",
          "가져오기(EXCEL): 'Codes(또는 코드)' 시트 기반으로 codeMaster 갱신"
        ]},
      ]
    };
  }

  function openHelpWindow() {
    const w = window.open("help.html", "FIN_HELP", "width=980,height=820");
    if (!w) {
      alert("팝업이 차단되었습니다. 브라우저에서 팝업 허용 후 다시 시도해 주세요.");
      return;
    }

    const payload = buildHelpPayload();
    let tries = 0;
    const timer = setInterval(() => {
      tries++;
      try { w.postMessage({ type: "HELP_INIT", payload }, window.location.origin); } catch {}
      if (tries >= 20) clearInterval(timer);
    }, 120);
  }

  /***************
   * UI: Tabs
   ***************/
  function renderTabs() {
    clear($tabs);
    for (const t of TABS) {
      const btn = el("button", {
        class: "tab" + (state.activeTab === t.id ? " active" : ""),
        onclick: () => {
          state.activeTab = t.id;
          saveState();
          render();
        }
      }, [t.title]);
      $tabs.appendChild(btn);
    }
  }

  /***************
   * Code tab
   ***************/
  function renderCodeTab() {
    const panelHeader = el("div", { class: "panel-header sticky-head", dataset: { sticky: "panel" } }, [
      el("div", {}, [ el("div", { class: "panel-title" }, ["코드"]) ]),
      el("div", { class: "row-actions" }, [
        el("button", { class: "smallbtn", onclick: () => addCodeRows(1) }, ["행 추가 (Ctrl+F3)"]),
        el("button", { class: "smallbtn", onclick: () => addCodeRows(10) }, ["+10행"]),
      ])
    ]);

    const scroll = el("div", { class: "table-wrap calc-scroll", dataset: { scroll: "code" } }, [buildCodeMasterTable()]);
    forceScrollStyle(scroll);
    attachGridNav(scroll);
    attachWheelLock(scroll);

    return el("div", { class: "panel" }, [panelHeader, scroll]);
  }

    function buildCodeMasterTable() {
    const table = el("table", { class: "code-table" }, []);
    table.style.tableLayout = "fixed";
    table.style.width = "100%";
    table.style.minWidth = "100%";

    table.appendChild(buildColGroupFromWeights(CODE_COL_WEIGHTS));

    const thead = el("thead", {}, [
      el("tr", {}, [
        el("th", {}, ["코드"]),
        el("th", {}, ["품명"]),
        el("th", {}, ["규격"]),
        el("th", {}, ["단위"]),
        el("th", {}, ["할증"]),
        el("th", {}, ["환산단위"]),
        el("th", {}, ["환산계수"]),
        el("th", {}, ["비고"]),
        el("th", {}, [""])
      ])
    ]);

    const tbody = el("tbody", {}, []);

      state.codeMaster.forEach((row, idx) => {
  const isFixed = (idx === 0 && isRemarkCode(row.code));

  const tr = el("tr", { class: isFixed ? "remark-row" : "" }, [
    tdInput("codeMaster", idx, "code", row.code, { readonly: isFixed }),
    tdInput("codeMaster", idx, "name", row.name, { readonly: isFixed }),
    tdInput("codeMaster", idx, "spec", row.spec, { readonly: isFixed }),
    tdInput("codeMaster", idx, "unit", row.unit, { readonly: isFixed }),
    tdInput("codeMaster", idx, "surcharge", row.surcharge ?? "", { readonly: isFixed }),
    tdInput("codeMaster", idx, "convUnit", row.convUnit, { readonly: isFixed }),
    tdInput("codeMaster", idx, "convFactor", row.convFactor ?? "", { readonly: isFixed }),
    tdInput("codeMaster", idx, "note", row.note, { readonly: isFixed }),
    el("td", {}, [
      el("button", {
        class: "smallbtn",
        disabled: isFixed ? "disabled" : null,
        onclick: () => {
          if (isFixed) return;

          state.codeMaster.splice(idx, 1);

          // ✅ 삭제 후에도 최상단 고정 보장
          ensureRemarkCodeMasterTop();

          saveState();
          render();
        }
      }, [isFixed ? "고정" : "삭제"])
    ])
  ]);

  tbody.appendChild(tr);
});



    table.appendChild(thead);
    table.appendChild(tbody);

    return table;
  }

  // ✅ 여기서 CALC_COL_INDEX를 "재선언"하지 않는다.
  // ✅ 위쪽(이미 존재하는) CALC_COL_INDEX를 그대로 사용한다.

  function tdInput(scope, rowIndex, field, value, opts = {}) {
    const ds =
      scope === "codeMaster"
        ? { grid: "code", row: String(rowIndex), col: String(CODE_COL_INDEX[field] ?? 0), field }
        : (opts.dataset || null);

    const input = el("input", {
      class: "cell" + (opts.readonly ? " readonly" : ""),
      value: value ?? "",
      readonly: opts.readonly ? "readonly" : null,
      dataset: ds,
      oninput: (e) => {
        const v = e.target.value;
        if (scope === "codeMaster") {
          const r = state.codeMaster[rowIndex];
          if (!r) return;
          if (field === "surcharge" || field === "convFactor") r[field] = v === "" ? null : Number(v);
          else r[field] = v;
          saveState();
        }
      }
    });

    input.addEventListener("blur", () => { delete input.dataset.editing; });

    return el("td", {}, [input]);
  }


  function addCodeRows(n, insertAfterRow = null) {
    const idx = insertAfterRow == null ? (state.codeMaster.length - 1) : insertAfterRow;
    const insertPos = clamp(idx + 1, 0, state.codeMaster.length);

    const empty = { code:"", name:"", spec:"", unit:"", surcharge:null, convUnit:"", convFactor:null, note:"" };
    const newRows = Array.from({ length: n }, () => deepClone(empty));

    state.codeMaster.splice(insertPos, 0, ...newRows);
    saveState();
    render();

    raf2(() => {
      updateViewFillHeight();
      updateScrollHeights();
      const first = document.querySelector(`input[data-grid="code"][data-row="${insertPos}"][data-col="0"]`);
      if (first) safeFocus(first);
      ensureScrollIntoView();
    });
  }

  /***************
   * ✅ Split resizer
   ***************/
  function attachSplitResizer(resizerEl, topPaneEl) {
    if (!resizerEl || !topPaneEl) return;

    const root = document.documentElement;

    const begin = (clientY) => {
      const startH = topPaneEl.getBoundingClientRect().height;
      const startY = clientY;

      document.body.classList.add("is-resizing");

      const move = (y) => {
        const dy = y - startY;
        const next = clamp(startH + dy, 120, 520);
        state.ui.topSplitH = next;
        root.style.setProperty("--topSplitH", `${Math.round(next)}px`);
        saveState();
        updateStickyVars();
        applyPanelStickyTop();
        updateViewFillHeight();
        updateScrollHeights();
      };

      const onMove = (e) => {
        if (e.touches && e.touches[0]) move(e.touches[0].clientY);
        else move(e.clientY);
      };

      const end = () => {
        document.body.classList.remove("is-resizing");
        window.removeEventListener("mousemove", onMove, true);
        window.removeEventListener("mouseup", end, true);
        window.removeEventListener("touchmove", onMove, { capture: true });
        window.removeEventListener("touchend", end, true);
        window.removeEventListener("touchcancel", end, true);

        raf2(() => {
          updateStickyVars();
          applyPanelStickyTop();
          updateViewFillHeight();
          updateScrollHeights();
        });
      };

      window.addEventListener("mousemove", onMove, true);
      window.addEventListener("mouseup", end, true);
      window.addEventListener("touchmove", onMove, { capture: true, passive: false });
      window.addEventListener("touchend", end, true);
      window.addEventListener("touchcancel", end, true);
    };

    resizerEl.addEventListener("mousedown", (e) => {
      e.preventDefault();
      begin(e.clientY);
    });

    resizerEl.addEventListener("touchstart", (e) => {
      if (!e.touches || !e.touches[0]) return;
      e.preventDefault();
      begin(e.touches[0].clientY);
    }, { passive: false });
  }

  /***************
   * Calc tab
   ***************/
  function renderCalcTab(tabId, title) {
    recomputeSection(tabId);

    const top = el("div", { class: "top-split" }, [
      el("div", { class: "calc-layout top-grid" }, [
        el("div", { class: "rail-box section-box", dataset: { region: "section" } }, [
          el("div", { class: "rail-title" }, ["구분명 리스트 (↑/↓ 이동)"]),
          buildSectionList(tabId),
          buildSectionEditor(tabId),
        ]),
        el("div", { class: "rail-box var-box", dataset: { region: "var" } }, [
          el("div", { class: "rail-title" }, ["변수표 (A, AB, A1, AB1... 최대 3자)"]),
          buildVarTable(tabId),
        ]),
      ])
    ]);

    const panelHeader = el("div", { class: "panel-header sticky-head", dataset: { sticky: "panel" } }, [
      el("div", {}, [ el("div", { class: "panel-title" }, [title]) ]),
      el("div", { class: "row-actions" }, [
        el("button", { class: "smallbtn", onclick: () => addRows(tabId, 1) }, ["행 추가 (Ctrl+F3)"]),
        el("button", { class: "smallbtn", onclick: () => addRows(tabId, 10) }, ["+10행"]),
      ])
    ]);

    const scroll = el("div", { class: "table-wrap calc-scroll", dataset: { scroll: "calc" } }, [buildCalcTable(tabId)]);
    forceScrollStyle(scroll);
    attachGridNav(scroll);
    attachWheelLock(scroll);

    const panel = el("div", { class: "panel" }, [panelHeader, scroll]);

    const topPane = el("div", { class: "pane top-pane" }, [top]);
    const resizer = el("div", { class: "split-resizer", dataset: { ui: "splitResizer" } }, []);
    const bottomPane = el("div", { class: "pane bottom-pane" }, [panel]);

    const workArea = el("div", { class: "work-area" }, [topPane, resizer, bottomPane]);

    raf2(() => {
      attachSplitResizer(resizer, topPane);
      updateViewFillHeight();
      updateScrollHeights();
    });

    return workArea;
  }

  function buildSectionList(tabId) {
    const bucket = state[tabId];

    const list = el("div", {
      class: "section-list",
      tabindex: "0",
      dataset: { nav: "sectionList", tab: tabId }
    }, []);

    bucket.sections.forEach((s, idx) => {
      const item = el("div", {
        class: "section-item" + (bucket.activeSection === idx ? " active" : ""),
        tabindex: "0",
        onclick: () => {
          bucket.activeSection = idx;
          saveState();
          __pendingSectionFocus = { tabId, index: idx };
          render();
        },
      }, [
        el("div", { class: "name" }, [s.name || `구분 ${idx + 1}`]),
        el("div", { class: "meta-inline" }, [`개소: ${s.count ?? ""}`]),
        el("div", { class: "meta" }, ["선택"])
      ]);
      list.appendChild(item);
    });

    list.addEventListener("mousedown", () => safeFocus(list));

    list.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return;
      const a = document.activeElement;
      if (a instanceof HTMLInputElement || a instanceof HTMLTextAreaElement) return;

      e.preventDefault();
      e.stopPropagation();

      const dir = e.key === "ArrowDown" ? 1 : -1;
      const nextIdx = clamp(bucket.activeSection + dir, 0, bucket.sections.length - 1);
      if (nextIdx === bucket.activeSection) return;

      bucket.activeSection = nextIdx;
      saveState();
      __pendingSectionFocus = { tabId, index: nextIdx };
      render();
    }, true);

    return list;
  }

  function buildSectionEditor(tabId) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const nameInput = el("input", {
      class: "cell",
      value: sec.name || "",
      placeholder: "구분명 (예: 2층 바닥 철골보)",
      oninput: (e) => {
        sec.name = e.target.value;
        saveState();
        const item = document.querySelectorAll(".section-item .name")[bucket.activeSection];
        if (item) item.textContent = sec.name || `구분 ${bucket.activeSection + 1}`;
      }
    });

    const countInput = el("input", {
      class: "cell",
      value: sec.count ?? "",
      placeholder: "개소(예: 0,1,2...)",
      oninput: (e) => {
        const v = e.target.value.trim();
        sec.count = v === "" ? "" : Number(v);
        saveState();
        const meta = document.querySelectorAll(".section-item .meta-inline")[bucket.activeSection];
        if (meta) meta.textContent = `개소: ${sec.count ?? ""}`;
      }
    });

    const btnWrap = el("div", { class: "row-actions", style: "justify-content:flex-end; gap:6px;" }, [
      el("button", { class: "smallbtn", onclick: () => { saveState(); render(); } }, ["저장"]),
      el("button", {
        class: "smallbtn",
        onclick: () => {
          bucket.sections.push(defaultSection(`구분 ${bucket.sections.length + 1}`, 1));
          bucket.activeSection = bucket.sections.length - 1;
          saveState(); render();
        }
      }, ["구분 추가"]),
      el("button", {
        class: "smallbtn",
        onclick: () => {
          if (bucket.sections.length <= 1) return alert("구분은 최소 1개가 필요합니다.");
          bucket.sections.splice(bucket.activeSection, 1);
          bucket.activeSection = clamp(bucket.activeSection, 0, bucket.sections.length - 1);
          saveState(); render();
        }
      }, ["구분 삭제"]),
    ]);

    return el("div", { class: "section-editor" }, [nameInput, countInput, btnWrap]);
  }

  function buildVarTable(tabId) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const wrap = el("div", { class: "var-tablewrap calc-scroll", dataset: { scroll: "var" } }, []);
    forceScrollStyle(wrap);
    attachWheelLock(wrap);

    const table = el("table", { class: "var-table" }, []);
    const thead = el("thead", {}, [
      el("tr", {}, [
        el("th", {}, ["변수"]),
        el("th", {}, ["산식"]),
        el("th", {}, ["값"]),
        el("th", {}, ["비고"])
      ])
    ]);
    const tbody = el("tbody", {}, []);

    sec.vars.forEach((v, r) => {
      const tr = el("tr", {}, [
        tdNavInputVar(tabId, r, 0, "key", v.key, { placeholder: "예: A / AB / A1" }),
        tdNavInputVar(tabId, r, 1, "expr", v.expr, { placeholder: "예: (A+0.5)*2  (<...> 주석)" }),
        tdNavInputVar(tabId, r, 2, "value", String(v.value ?? 0), { readonly: true }),
        tdNavInputVar(tabId, r, 3, "note", v.note, { placeholder: "비고" }),
      ]);
      tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    wrap.appendChild(table);

    wrap.addEventListener("input", () => {
      recomputeSection(tabId);
      saveState();
      const valueInputs = wrap.querySelectorAll('input[data-grid="var"][data-col="2"]');
      sec.vars.forEach((vv, i) => {
        if (valueInputs[i]) valueInputs[i].value = String(vv.value ?? 0);
      });
      refreshCalcComputed(tabId);

      updateViewFillHeight();
      updateScrollHeights();
    });

    attachGridNav(wrap);

    raf2(() => {
      updateViewFillHeight();
      updateScrollHeights();
    });

    return wrap;
  }

  function tdNavInputVar(tabId, row, col, field, value, opts = {}) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const input = el("input", {
      class: "cell" + (opts.readonly ? " readonly" : ""),
      value: value ?? "",
      placeholder: opts.placeholder || "",
      readonly: opts.readonly ? "readonly" : null,
      dataset: { grid: "var", tab: tabId, row: String(row), col: String(col), field },
      oninput: (e) => {
        if (opts.readonly) return;
        const rr = sec.vars[row];
        if (!rr) return;

        if (field === "key") {
          let val = e.target.value.toUpperCase();
          val = val.replace(/[^A-Z0-9]/g, "");
          if (val.length > 3) val = val.slice(0, 3);
          if (val && !/^[A-Z]/.test(val)) val = val.replace(/^[^A-Z]+/, "");
          e.target.value = val;
          rr.key = val;
        } else {
          rr[field] = e.target.value;
        }
      }
    });

    input.addEventListener("blur", () => { delete input.dataset.editing; });
    return el("td", {}, [input]);
  }

  function buildCalcTable(tabId) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const table = el("table", { class: "calc-table" }, []);
    table.style.tableLayout = "fixed";
    table.style.width = "100%";
    table.style.minWidth = "100%";

    table.appendChild(buildColGroupFromWeights(CALC_COL_WEIGHTS));

    const thead = el("thead", {}, [
      el("tr", {}, [
        el("th", {}, ["No"]),
        el("th", {}, ["코드"]),
        el("th", {}, ["품명(자동)"]),
        el("th", {}, ["규격(자동)"]),
        el("th", {}, ["단위(자동)"]),
        el("th", {}, ["산출식"]),
        el("th", {}, ["물량"]),
        el("th", {}, ["할증(%)"]),
        el("th", {}, ["환산단위"]),
        el("th", {}, ["환산계수"]),
        el("th", {}, ["환산후수량"]),
      ])
    ]);

    const tbody = el("tbody", {}, []);

    sec.rows.forEach((r, i) => {
      const tr = el("tr", { class: hasAtLeastFiveZ(r.code) ? "z5-row" : "" }, [
        el("td", {}, [String(i + 1)]),
        tdNavInputCalc(tabId, i, 0, "code", r.code, { placeholder: "코드 입력" }),
        tdNavInputCalc(tabId, i, 1, "name", r.name, { readonly: true }),
        tdNavInputCalc(tabId, i, 2, "spec", r.spec, { readonly: true }),
        tdNavInputCalc(tabId, i, 3, "unit", r.unit, { readonly: true }),
        tdNavInputCalc(tabId, i, 4, "formula", r.formula, { placeholder: "예: (A+0.5)*2  (<...> 주석)" }),
        tdNavInputCalc(tabId, i, 5, "value", String(r.value ?? 0), { readonly: true }),
        tdNavInputCalc(tabId, i, 6, "surchargePct", r.surchargePct ?? "", { readonly: true }),
        tdNavInputCalc(tabId, i, 7, "convUnit", r.convUnit || "", { readonly: true }),
        tdNavInputCalc(tabId, i, 8, "convFactor", r.convFactor ?? "", { readonly: true }),
        tdNavInputCalc(tabId, i, 9, "converted", String(r.converted ?? 0), { readonly: true }),
      ]);
      tbody.appendChild(tr);
    });


    table.appendChild(thead);
    table.appendChild(tbody);

        raf2(() => __applyCalcRowSelectionStyles(tabId));

    // ✅ buildCalcTable 내부 keydown (중간생략 없음)
    table.addEventListener("keydown", (e) => {
      const t = e.target;
      if (!(t instanceof HTMLInputElement)) return;
      if (t.dataset.grid !== "calc") return;

      // 편집 중이면 Enter로 편집 종료만 허용
      if (t.dataset.editing === "1") {
        if (e.key === "Enter") {
          e.preventDefault();
          delete t.dataset.editing;
          t.blur();
          raf2(() => safeFocus(t));
        }
        return;
      }

      const curRow = Number(t.dataset.row || 0);


      // ESC 단독은 아무 동작 안 함 (블록 유지)
if (e.key === "Escape") {
  return;
}


      // Shift+B: 다중선택 토글
      if ((e.key === "B" || e.key === "b") && e.shiftKey) {
        e.preventDefault();
        if (!__calcMultiIsSameContext(tabId)) __calcMultiBegin(tabId, curRow);
        else __calcMultiClear();
        __applyCalcRowSelectionStyles(tabId);
        return;
      }

      // Shift+↑/↓ : 다중선택 범위 확장 + 포커스 이동
      if ((e.key === "ArrowUp" || e.key === "ArrowDown") && e.shiftKey) {
        e.preventDefault();

        const bucket = state[tabId];
        const sec = bucket.sections[bucket.activeSection];

        const next = clamp(
          curRow + (e.key === "ArrowDown" ? 1 : -1),
          0,
          sec.rows.length - 1
        );

        if (!__calcMultiIsSameContext(tabId)) __calcMultiBegin(tabId, curRow);
        __calcMultiSetRange(tabId, __calcMulti.anchorRow ?? curRow, next);
        __applyCalcRowSelectionStyles(tabId);

        raf2(() => {
          const col = t.dataset.col || String(CALC_COL_INDEX.code);
          const target = document.querySelector(
            `input[data-grid="calc"][data-tab="${tabId}"][data-row="${next}"][data-col="${col}"]`
          );
          safeFocus(target);
          ensureScrollIntoView(target);
        });
        return;
      }

      // Ctrl+Del: 선택행 삭제(없으면 현재행)
      if ((e.key === "Delete" || e.key === "Del") && e.ctrlKey) {
        e.preventDefault();
        const selected = __getSelectedCalcRows(tabId);
        const targets = selected.length ? selected : [curRow];
        if (!confirm(`선택된 ${targets.length}행을 삭제할까요?`)) return;
        deleteCalcRows(tabId, targets);
        __calcMultiClear();
        return;
      }

      // Del(단독): 현재행 삭제
      if ((e.key === "Delete" || e.key === "Del") && !e.ctrlKey) {
        e.preventDefault();
        if (!confirm("현재 행을 삭제할까요?")) return;
        deleteCalcRows(tabId, [curRow]);
        return;
      }

      // Ctrl+G: 선택행 복사/삽입
      if ((e.key === "g" || e.key === "G") && e.ctrlKey) {
        e.preventDefault();
        const selected = __getSelectedCalcRows(tabId);
        if (!selected.length) return;
        duplicateCalcRows(tabId, selected, curRow);
        return;
      }
    }, true);



    // ✅ input 변화가 있을 때 재계산 + 저장 + 렌더 반영
    table.addEventListener("input", (e) => {
      const t = e.target;
      if (!(t instanceof HTMLInputElement)) return;
      if (t.dataset.grid !== "calc") return;
      if (t.dataset.tab !== tabId) return;

      const row = Number(t.dataset.row || 0);
      const field = t.dataset.field;

      const bucket2 = state[tabId];
      const sec2 = bucket2.sections[bucket2.activeSection];
      const rr = sec2.rows[row];
      if (!rr) return;

      if (field === "code") {
        rr.code = (t.value || "").trim();
      } else if (field === "formula") {
        rr.formula = t.value || "";
      } else {
        // readonly는 원칙적으로 여기에 안 옴
        rr[field] = t.value;
      }

            recomputeSection(tabId);
      saveState();
      refreshCalcComputed(tabId); // 값/환산/자동필드 갱신
      __applyZ5RowStyles(tabId);  // ✅ Z 5개 이상 행 회색 즉시 반영
    });


    return table;
  }

  function tdNavInputCalc(tabId, row, _colNo, field, value, opts = {}) {
  const bucket = state[tabId];
  const sec = bucket.sections[bucket.activeSection];

  // ✅ data-col은 무조건 “실제 테이블 열 번호”로 고정
  const dataCol = String(CALC_COL_INDEX[field] ?? 0);

  const input = el("input", {
    class: "cell" + (opts.readonly ? " readonly" : ""),
    value: value ?? "",
    placeholder: opts.placeholder || "",
    readonly: opts.readonly ? "readonly" : null,
    dataset: { grid: "calc", tab: tabId, row: String(row), col: dataCol, field },

    onfocus: () => {
      if (__calcMulti.active && __calcMultiIsSameContext(tabId)) {
        __applyCalcRowSelectionStyles(tabId);
      }
    },

    onkeydown: (e) => {
      const t = e.target;
      if (!(t instanceof HTMLInputElement)) return;

      if (e.key === "F2") {
        if (t.readOnly) return;
        e.preventDefault();
        t.dataset.editing = "1";
        t.setSelectionRange?.(t.value.length, t.value.length);
        return;
      }

      if (e.key === "Enter") {
        if (t.dataset.editing === "1") {
          e.preventDefault();
          delete t.dataset.editing;
          t.blur();
          raf2(() => safeFocus(t));
          return;
        }
      }
    },

    oninput: (e) => {
      if (opts.readonly) return;
      const rr = sec.rows[row];
      if (!rr) return;
      rr[field] = e.target.value;
    }
  });

  input.addEventListener("blur", () => { delete input.dataset.editing; });
  return el("td", {}, [input]);
}



  function refreshCalcComputed(tabId) {
    // 현재 tab의 calc-table에서 readonly 셀들 업데이트
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    sec.rows.forEach((r, i) => {
      const setVal = (field, v) => {
        const col = CALC_COL_INDEX[field];
        const inp = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${i}"][data-col="${col}"]`);
        if (inp) inp.value = (v ?? "");
      };

      setVal("name", r.name || "");
      setVal("spec", r.spec || "");
      setVal("unit", r.unit || "");
      setVal("value", String(r.value ?? 0));
      setVal("surchargePct", (r.surchargePct ?? "") === null ? "" : String(r.surchargePct ?? ""));
      setVal("convUnit", r.convUnit || "");
      setVal("convFactor", (r.convFactor ?? "") === null ? "" : String(r.convFactor ?? ""));
      setVal("converted", String(r.converted ?? 0));
    });

        // 다중선택 표시 갱신
    raf2(() => __applyCalcRowSelectionStyles(tabId));

    // ✅ Z 5개 이상 행 회색 표시 갱신
    raf2(() => __applyZ5RowStyles(tabId));
  }


  function addRows(tabId, n, insertAfterRow = null) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const idx = insertAfterRow == null ? (sec.rows.length - 1) : insertAfterRow;
    const insertPos = clamp(idx + 1, 0, sec.rows.length);

    const newRows = Array.from({ length: n }, () => defaultCalcRow());
    sec.rows.splice(insertPos, 0, ...newRows);

    saveState();
    render();

    raf2(() => {
      updateViewFillHeight();
      updateScrollHeights();
      const first = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${insertPos}"][data-col="${CALC_COL_INDEX.code}"]`);
      safeFocus(first);
      ensureScrollIntoView(first);
    });
  }

  function deleteCalcRows(tabId, rowIndices) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];
    const uniq = [...new Set(rowIndices)].sort((a, b) => b - a); // 뒤에서부터 삭제

    uniq.forEach((r) => {
      if (r >= 0 && r < sec.rows.length) sec.rows.splice(r, 1);
    });

    if (sec.rows.length === 0) sec.rows.push(defaultCalcRow());

    saveState();
    render();

    raf2(() => {
      updateViewFillHeight();
      updateScrollHeights();
      const targetRow = clamp(Math.min(...rowIndices), 0, sec.rows.length - 1);
      const target = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${targetRow}"][data-col="${CALC_COL_INDEX.code}"]`);
      safeFocus(target);
      ensureScrollIntoView(target);
    });
  }

  function duplicateCalcRows(tabId, rowIndices, insertAfterRow) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const selected = [...new Set(rowIndices)].sort((a, b) => a - b);
    const clones = selected
      .map((r) => sec.rows[r])
      .filter(Boolean)
      .map((r) => deepClone(r));

    if (!clones.length) return;

    const insertPos = clamp((insertAfterRow ?? selected[selected.length - 1]) + 1, 0, sec.rows.length);
    sec.rows.splice(insertPos, 0, ...clones);

    saveState();
    render();

    raf2(() => {
      updateViewFillHeight();
      updateScrollHeights();
      const target = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${insertPos}"][data-col="${CALC_COL_INDEX.code}"]`);
      safeFocus(target);
      ensureScrollIntoView(target);
    });
  }

  /* ============================
     ✅ Summation Tabs (steel_sum / support_sum)
     - (v12.3) 개소(count) 반영
     - (v12.3) 환산단위/환산계수 있으면 converted 기준 집계
  ============================ */
  function buildSummaryRows(tabId) {
    const bucket = state[tabId];
    const map = new Map();

    bucket.sections.forEach((sec) => {
      const count = Number(sec.count ?? 1);
      const mult = Number.isFinite(count) && count > 0 ? count : 1;

      // section별로 vars/rows 값이 계산되어 있어야 함
      // recomputeSection는 activeSection만 계산하므로, 여기선 간단히 현재 저장값(value/converted)을 사용
      sec.rows.forEach((r) => {
        const code = (r.code || "").trim();
        if (!code) return;

        const info = codeLookup(code);
        const unit = info?.unit || r.unit || "";
        const surcharge = (r.surchargePct == null ? (info?.surcharge ?? null) : r.surchargePct);

        // 환산계수 있으면 converted 기준
        const hasConv = r.convFactor != null && Number.isFinite(Number(r.convFactor)) && Number(r.convFactor) !== 0;
        const qty = hasConv ? Number(r.converted || 0) : Number((r.value || 0) * (r.surchargeMul || 1));

        const key = code.toUpperCase();
        const prev = map.get(key) || {
          code,
          name: info?.name || r.name || "",
          spec: info?.spec || r.spec || "",
          unit,
          convUnit: info?.convUnit || r.convUnit || "",
          convFactor: info?.convFactor ?? r.convFactor ?? null,
          surchargePct: surcharge,
          qty: 0,
        };
        prev.qty += qty * mult;
        map.set(key, prev);
      });
    });

    return [...map.values()].sort((a, b) => String(a.code).localeCompare(String(b.code)));
  }

  function renderSummaryTab(srcTabId, title) {
    const rows = buildSummaryRows(srcTabId);

    const header = el("div", { class: "panel-header sticky-head", dataset: { sticky: "panel" } }, [
      el("div", {}, [ el("div", { class: "panel-title" }, [title]) ]),
      el("div", { class: "row-actions" }, [
        el("button", { class: "smallbtn", onclick: () => { /* noop */ } }, ["집계(자동)"]),
      ]),
    ]);

    const table = el("table", { class: "code-table" }, []);
    table.style.tableLayout = "fixed";
    table.style.width = "100%";
    table.style.minWidth = "100%";
    table.appendChild(buildColGroupFromWeights([0.9, 2.4, 2.4, 0.8, 0.8, 0.9, 0.9, 1.4, 1.2]));

    const thead = el("thead", {}, [
      el("tr", {}, [
        el("th", {}, ["코드"]),
        el("th", {}, ["품명"]),
        el("th", {}, ["규격"]),
        el("th", {}, ["단위"]),
        el("th", {}, ["할증"]),
        el("th", {}, ["환산단위"]),
        el("th", {}, ["환산계수"]),
        el("th", {}, ["수량(환산/할증 반영)"]),
        el("th", {}, ["비고"]),
      ])
    ]);

    const tbody = el("tbody", {}, []);
    rows.forEach((r) => {
      tbody.appendChild(el("tr", {}, [
        el("td", {}, [r.code]),
        el("td", {}, [r.name || ""]),
        el("td", {}, [r.spec || ""]),
        el("td", {}, [r.unit || ""]),
        el("td", {}, [r.surchargePct == null ? "" : String(r.surchargePct)]),
        el("td", {}, [r.convUnit || ""]),
        el("td", {}, [r.convFactor == null ? "" : String(r.convFactor)]),
        el("td", {}, [String(Math.round((Number(r.qty) || 0) * 1000) / 1000)]),
        el("td", {}, [""]),
      ]));
    });

    table.appendChild(thead);
    table.appendChild(tbody);

    const scroll = el("div", { class: "table-wrap calc-scroll", dataset: { scroll: "sum" } }, [table]);
    forceScrollStyle(scroll);
    attachWheelLock(scroll);

    return el("div", { class: "panel" }, [header, scroll]);
  }

  /* ============================
     ✅ Grid Navigation (Arrow/PageUp/PageDown/Home/End)
     - (간단 구현) data-grid="code|var|calc"
  ============================ */
  function parseCellDataset(input) {
    const ds = input?.dataset || {};
    return {
      grid: ds.grid || "",
      tab: ds.tab || "",
      row: Number(ds.row || 0),
      col: Number(ds.col || 0),
    };
  }

  function queryCell(grid, tab, row, col) {
    const selector =
      grid === "code"
        ? `input[data-grid="code"][data-row="${row}"][data-col="${col}"]`
        : `input[data-grid="${grid}"][data-tab="${tab}"][data-row="${row}"][data-col="${col}"]`;
    return document.querySelector(selector);
  }

  function moveCell(fromInput, dRow, dCol, pageJump = false) {
  const { grid, tab, row, col } = parseCellDataset(fromInput);
  if (!grid) return;

  // row/col 범위 추정
  let maxRow = 0;
  let maxCol = 0;

  const all = grid === "code"
    ? document.querySelectorAll(`input[data-grid="code"]`)
    : document.querySelectorAll(`input[data-grid="${grid}"][data-tab="${tab}"]`);

  all.forEach((x) => {
    const r = Number(x.dataset.row || 0);
    const c = Number(x.dataset.col || 0);
    if (r > maxRow) maxRow = r;
    if (c > maxCol) maxCol = c;
  });

  let nextRow = clamp(row + dRow, 0, maxRow);
  let nextCol = clamp(col + dCol, 0, maxCol);

  if (pageJump) {
    // pageJump일 때는 scroller 높이 기준으로 row를 대략 이동
    const sc = fromInput.closest(".calc-scroll");
    if (sc) {
      const rect = sc.getBoundingClientRect();
      const rowH = 34; // 대략
      const jump = Math.max(1, Math.floor(rect.height / rowH) - 1);
      nextRow = clamp(row + (dRow > 0 ? jump : -jump), 0, maxRow);
    }
  }

  const target = queryCell(grid, tab, nextRow, nextCol);
  if (target) {
    // ✅ 포커스는 1번만 (2번 주면 sticky/transform 환경에서 튐이 생길 수 있음)
    safeFocus(target);

    // ✅ 스크롤만 다음 프레임에서 보정
    raf2(() => {
      ensureScrollIntoView(target);
    });
  }
}

function attachGridNav(container) {
  if (!container) return;
  container.addEventListener("keydown", (e) => {
    const t = e.target;
    if (!(t instanceof HTMLInputElement)) return;
    if (!t.dataset.grid) return;

    // 편집중이면 방향키 이동 막음
    if (t.dataset.editing === "1") return;

    const isInput = (document.activeElement instanceof HTMLInputElement);
    if (!isInput) return;

    const key = e.key;

    if (key === "ArrowUp") { e.preventDefault(); moveCell(t, -1, 0); }
    else if (key === "ArrowDown") { e.preventDefault(); moveCell(t, 1, 0); }
    else if (key === "ArrowLeft") { e.preventDefault(); moveCell(t, 0, -1); }
    else if (key === "ArrowRight") { e.preventDefault(); moveCell(t, 0, 1); }
    else if (key === "PageUp") { e.preventDefault(); moveCell(t, -1, 0, true); }
    else if (key === "PageDown") { e.preventDefault(); moveCell(t, 1, 0, true); }
    else if (key === "Home" && e.ctrlKey) { e.preventDefault(); moveCell(t, -99999, 0); }
    else if (key === "End" && e.ctrlKey) { e.preventDefault(); moveCell(t, 99999, 0); }
    else if ((key === "Delete" || key === "Del") && e.ctrlKey) {
  const grid = t.dataset.grid;
  if (grid === "var") {
    if (t.readOnly) return;
    e.preventDefault();
    t.value = "";
    t.dispatchEvent(new Event("input", { bubbles: true }));
  } else if (grid === "code") {
    e.preventDefault();
    const row = Number(t.dataset.row || 0);
    if (confirm("현재 행을 삭제할까요?")) {
      // codeMaster 0행(비고 고정) 보호까지 고려
      if (row === 0) return;
      state.codeMaster.splice(row, 1);
      ensureRemarkCodeMasterTop();
      saveState();
      render();
    }
  }
}

  }, true);
}

  /* ============================
     ✅ wheel lock (trackpad/space bounce 방지)
  ============================ */
  function attachWheelLock(scroller) {
    if (!scroller) return;
    scroller.addEventListener("wheel", (e) => {
      // 기본 스크롤 허용(단, 바깥으로 튀는 스크롤만 차단)
      const el = scroller;
      const delta = e.deltaY;
      if (delta < 0 && el.scrollTop <= 0) e.preventDefault();
      else if (delta > 0 && el.scrollTop + el.clientHeight >= el.scrollHeight) e.preventDefault();
    }, { passive: false });
  }

  function forceScrollStyle(sc) {
  if (!sc || !(sc instanceof HTMLElement)) return;
  sc.style.overflow = "auto";
  sc.style.webkitOverflowScrolling = "touch";
  sc.style.minHeight = "0";
  sc.tabIndex = -1;

  // ✅ 사용자가 스크롤 중인지 판단하기 위한 타임스탬프 기록
  if (!sc.__finScrollBound) {
    sc.__finScrollBound = true;
    sc.addEventListener("scroll", () => {
      sc.__lastScrollAt = Date.now();
    }, { passive: true });
  }
}


  function ensureScrollIntoView(target) {
  if (!target || !(target instanceof HTMLElement)) return;
  const sc = target.closest(".calc-scroll");
  if (!sc) return;

  const tRect = target.getBoundingClientRect();
  const sRect = sc.getBoundingClientRect();

  // ✅ 스티키 헤더(패널헤더/테이블헤더)에 가려지는 상단 여유를 반영
  // (실측 기반으로 복잡하게 가지 않고, 안정적인 고정 topPad로 보정)
  const topPad = 60;     // 대략 thead + 여유
  const bottomPad = 10;

  if (tRect.top < sRect.top + topPad) {
    sc.scrollTop -= (sRect.top + topPad - tRect.top);
  } else if (tRect.bottom > sRect.bottom - bottomPad) {
    sc.scrollTop += (tRect.bottom - (sRect.bottom - bottomPad));
  }
}

  /* ============================
     ✅ Sticky Panel Top 적용
  ============================ */
  function applyPanelStickyTop() {
    const root = document.documentElement;
    const top = state.activeTab === "code"
      ? getComputedStyle(root).getPropertyValue("--stickyBaseTop").trim()
      : getComputedStyle(root).getPropertyValue("--stickyWithTopSplitTop").trim();

    document.querySelectorAll('[data-sticky="panel"]').forEach((h) => {
      if (!(h instanceof HTMLElement)) return;
      h.style.top = top || "0px";
    });
  }


   function openExportModal() {
  // 프로젝트 미선택 방지
  if (!activeProjectId) {
    alert("프로젝트를 먼저 선택(열기)해 주세요.");
    return;
  }

  // 이미 모달이 있으면 제거 후 재생성(중복 방지)
  const old = document.getElementById("exportModal");
  if (old) old.remove();

  const modal = document.createElement("div");
  modal.id = "exportModal";
  modal.className = "modal";
  modal.setAttribute("aria-hidden", "false");
  modal.hidden = false;

  modal.innerHTML = `
    <div class="modal-backdrop" data-close="1"></div>
    <div class="modal-card" role="dialog" aria-modal="true" aria-labelledby="exportModalTitle">
      <div class="modal-head">
        <div class="modal-title" id="exportModalTitle">내보내기(EXCEL)</div>
        <div class="modal-head-actions">
          <button id="btnExportAll" class="btn">전체선택</button>
          <button id="btnExportDo" class="btn btn-primary">다운로드</button>
          <button id="btnExportClose" class="btn">닫기</button>
        </div>
      </div>
      <div class="modal-body">
        <div class="project-hint" style="margin-bottom:10px;">
          내보낼 탭을 선택하세요.
        </div>

        <div style="display:flex; flex-direction:column; gap:10px;">
          <label style="display:flex; gap:10px; align-items:center;">
            <input type="checkbox" data-tab="code" checked />
            <b>코드</b>
          </label>

          <label style="display:flex; gap:10px; align-items:center;">
            <input type="checkbox" data-tab="steel" checked />
            <b>철골</b>
          </label>

          <label style="display:flex; gap:10px; align-items:center;">
            <input type="checkbox" data-tab="steel_sub" checked />
            <b>철골_부자재</b>
          </label>

          <label style="display:flex; gap:10px; align-items:center;">
            <input type="checkbox" data-tab="support" checked />
            <b>구조이기/동바리</b>
          </label>
        </div>
      </div>
      <div class="modal-foot">
        <div class="muted">* 선택한 탭들이 하나의 .xlsx 파일에 시트로 포함됩니다.</div>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  const close = () => {
    modal.setAttribute("aria-hidden", "true");
    modal.hidden = true;
    modal.remove();
  };

  modal.addEventListener("click", (e) => {
    const t = e.target;
    if (t && t.getAttribute && t.getAttribute("data-close") === "1") close();
  });

  const btnClose = document.getElementById("btnExportClose");
  const btnAll = document.getElementById("btnExportAll");
  const btnDo = document.getElementById("btnExportDo");

  if (btnClose) btnClose.onclick = close;

  if (btnAll) {
    btnAll.onclick = () => {
      const checks = modal.querySelectorAll('input[type="checkbox"][data-tab]');
      const allChecked = Array.from(checks).every(c => c.checked);
      checks.forEach(c => { c.checked = !allChecked; });
    };
  }

  if (btnDo) {
    btnDo.onclick = () => {
      const checks = modal.querySelectorAll('input[type="checkbox"][data-tab]');
      const selected = Array.from(checks).filter(c => c.checked).map(c => c.getAttribute("data-tab"));
      if (!selected.length) {
        alert("내보낼 탭을 1개 이상 선택해 주세요.");
        return;
      }
      exportToExcelSelectedTabs(selected);
      close();
    };
  }
}



   
  /* ============================
     ✅ Export / Import (placeholder-safe)
     - XLSX가 페이지에 로드되어 있으면 실제로 동작
     - 없으면 alert로 안내 (런타임 에러 방지)
  ============================ */
  function exportToExcelSelectedTabs(tabIds) {
  if (!window.XLSX) {
    alert("XLSX 라이브러리가 로드되지 않았습니다.\nCDN(xlsx.full.min.js) 로드 상태를 확인해 주세요.");
    return;
  }

  const wb = window.XLSX.utils.book_new();

  const sanitizeSheetName = (name) => {
    // Excel sheet name 제한 대응: \ / ? * [ ] : 최대 31자
    return String(name || "")
      .replace(/[\\/?*\[\]:]/g, "_")
      .slice(0, 31) || "Sheet";
  };

  const addSheet = (sheetName, aoaOrJson, mode = "aoa") => {
    const ws = (mode === "json")
      ? window.XLSX.utils.json_to_sheet(aoaOrJson)
      : window.XLSX.utils.aoa_to_sheet(aoaOrJson);
    window.XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(sheetName));
  };

  const want = new Set((tabIds || []).map(x => String(x)));

  // 1) Codes (코드마스터)
  if (want.has("code")) {
    const rows = Array.isArray(state.codeMaster) ? state.codeMaster : [];
    const aoa = [
      ["code", "name", "spec", "unit", "surcharge", "convUnit", "convFactor", "note"],
      ...rows.map(r => ([
        r.code ?? "",
        r.name ?? "",
        r.spec ?? "",
        r.unit ?? "",
        r.surcharge ?? "",
        r.convUnit ?? "",
        r.convFactor ?? "",
        r.note ?? ""
      ]))
    ];
    addSheet("Codes", aoa, "aoa");
  }

  // 2) 산출탭(구분 포함 flatten)
  const exportCalcTab = (tabId, sheetTitle) => {
    const bucket = state?.[tabId];
    if (!bucket || !Array.isArray(bucket.sections)) return;

    // export 전 계산값 최신화(현재 activeSection만 recompute가 아니라, export는 저장값 기준이라도 괜찮지만
    // 최대한 정확하게 하려면 섹션별로 varMap 재계산이 필요 -> 간단히 activeSection만 최신화 + 저장값 사용)
    try { recomputeSection(tabId); } catch {}

    const aoa = [
      ["sectionName", "count", "no", "code", "name", "spec", "unit", "formula", "value", "surchargePct", "convUnit", "convFactor", "converted", "note"],
    ];

    bucket.sections.forEach((sec, sidx) => {
      const sectionName = sec?.name ?? `구분 ${sidx + 1}`;
      const count = sec?.count ?? "";

      const rows = Array.isArray(sec?.rows) ? sec.rows : [];
      rows.forEach((r, i) => {
        aoa.push([
          sectionName,
          count,
          i + 1,
          r.code ?? "",
          r.name ?? "",
          r.spec ?? "",
          r.unit ?? "",
          r.formula ?? "",
          r.value ?? 0,
          r.surchargePct ?? "",
          r.convUnit ?? "",
          r.convFactor ?? "",
          r.converted ?? 0,
          r.note ?? ""
        ]);
      });
    });

    addSheet(sheetTitle, aoa, "aoa");
  };

  if (want.has("steel")) exportCalcTab("steel", "철골");
  if (want.has("steel_sub")) exportCalcTab("steel_sub", "철골_부자재");
  if (want.has("support")) exportCalcTab("support", "구조이기_동바리");

  // 파일명: 프로젝트명 반영
  const meta = projectIndex.projects.find(p => p.id === activeProjectId);
  const baseName = meta ? `${(meta.code || "FIN")}_${(meta.name || "Project")}` : "FIN_Project";
  const filename = `${baseName}_export.xlsx`;

  window.XLSX.writeFile(wb, filename);
}

function importFromExcelFile(file) {
  if (!window.XLSX) {
    alert("XLSX 라이브러리가 로드되지 않았습니다.\nCDN(xlsx.full.min.js) 로드 상태를 확인해 주세요.");
    return;
  }

  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = evt.target.result;
      const wb = window.XLSX.read(data, { type: "array" });

      // Codes 또는 코드 시트 찾기
      const sheetName =
        wb.SheetNames.find(n => n.toLowerCase() === "codes") ||
        wb.SheetNames.find(n => n.includes("코드")) ||
        wb.SheetNames[0];

      const ws = wb.Sheets[sheetName];
      if (!ws) {
        alert("가져올 시트를 찾지 못했습니다. (Codes 또는 코드 시트 필요)");
        return;
      }

      const aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
      if (!Array.isArray(aoa) || aoa.length < 2) {
        alert("Codes/코드 시트에 데이터가 없습니다.");
        return;
      }

      // 헤더 기반 매핑 (code, name, spec, unit, surcharge, convUnit, convFactor, note)
      const header = aoa[0].map(x => String(x || "").trim());
      const idx = (key) => header.findIndex(h => h.toLowerCase() === key.toLowerCase());

      const iCode = idx("code");
      const iName = idx("name");
      const iSpec = idx("spec");
      const iUnit = idx("unit");
      const iSurcharge = idx("surcharge");
      const iConvUnit = idx("convUnit");
      const iConvFactor = idx("convFactor");
      const iNote = idx("note");

      if (iCode < 0) {
        alert("Codes/코드 시트 헤더에 'code' 컬럼이 필요합니다.");
        return;
      }

      const next = [];
      for (let r = 1; r < aoa.length; r++) {
        const row = aoa[r];
        if (!row) continue;

        const code = String(row[iCode] ?? "").trim();
        if (!code) continue;

        const obj = {
          code,
          name: iName >= 0 ? String(row[iName] ?? "") : "",
          spec: iSpec >= 0 ? String(row[iSpec] ?? "") : "",
          unit: iUnit >= 0 ? String(row[iUnit] ?? "") : "",
          surcharge: iSurcharge >= 0 ? (row[iSurcharge] === "" || row[iSurcharge] == null ? null : Number(row[iSurcharge])) : null,
          convUnit: iConvUnit >= 0 ? String(row[iConvUnit] ?? "") : "",
          convFactor: iConvFactor >= 0 ? (row[iConvFactor] === "" || row[iConvFactor] == null ? null : Number(row[iConvFactor])) : null,
          note: iNote >= 0 ? String(row[iNote] ?? "") : "",
        };

        // 숫자 NaN 정리
        if (!Number.isFinite(obj.surcharge)) obj.surcharge = (obj.surcharge == null ? null : null);
        if (!Number.isFinite(obj.convFactor)) obj.convFactor = (obj.convFactor == null ? null : null);

        next.push(obj);
      }

      if (!next.length) {
        alert("가져올 코드 데이터가 없습니다.");
        return;
      }

            // ✅ 비고 고정코드는 가져오기에서 덮어쓰지 않도록 강제 유지
const filtered = next.filter(r => !isRemarkCode(r?.code));
state.codeMaster = filtered;
ensureRemarkCodeMasterTop();

saveState();
render();



      alert(`코드 ${next.length}개를 가져왔습니다. (시트: ${sheetName})`);
    } catch (err) {
      console.error(err);
      alert("가져오기 실패: 파일 형식/시트 구성을 확인해 주세요.");
    }
  };

  reader.readAsArrayBuffer(file);
}


  /* ============================
   ✅ Top Buttons (bind once)
   - 요소가 아직 없으면 raf2로 재시도
============================ */
let __topButtonsBound = false;

function bindTopButtonsOnce(forceRetry = false) {
  if (__topButtonsBound && !forceRetry) return;

  const btnHelp = document.getElementById("btnHelp");

  // ✅ 상단 프로젝트 버튼
  const btnProject = document.getElementById("btnProject");

  // ✅ 모달 버튼들
  const btnProjectAdd = document.getElementById("btnProjectAdd");
  const btnProjectDelete = document.getElementById("btnProjectDelete");
  const btnProjectSave = document.getElementById("btnProjectSave");
  const btnProjectClose = document.getElementById("btnProjectClose");
  const btnProjectOpen = document.getElementById("btnProjectOpen");

  // ✅ 상단 주요 기능 버튼들
  const btnOpenPicker = document.getElementById("btnOpenPicker");
  const btnExport = document.getElementById("btnExport");
  const btnReset = document.getElementById("btnReset");
  const fileImport = document.getElementById("fileImport");

  // ✅ DOM 아직 안 잡히면 다음 프레임 재시도
    const needRetry =
    !btnProject || !btnProjectAdd || !btnProjectDelete || !btnProjectSave || !btnProjectClose || !btnProjectOpen ||
    !btnOpenPicker || !btnExport || !btnReset || !fileImport;


  if (needRetry) {
    raf2(() => bindTopButtonsOnce(true));
    return;
  }

  __topButtonsBound = true;

  // 도움말
  if (btnHelp) btnHelp.onclick = openHelpWindow;

  // ✅ 상단 “프로젝트 열기” → 모달 열기
  btnProject.onclick = openProjectModal;

  // ✅ 모달 버튼 바인딩
  btnProjectAdd.onclick = createProject;
  btnProjectDelete.onclick = deleteSelectedProjectInModal;
  btnProjectSave.onclick = saveProjectsFromModal;
  btnProjectClose.onclick = closeProjectModal;
  btnProjectOpen.onclick = openSelectedProjectFromModal;

    // ✅ 코드선택 버튼 연결 (Ctrl+. / 버튼 모두 이걸 타게 됨)
  if (btnOpenPicker) {
    btnOpenPicker.onclick = openPickerWindow;
  }


  // 내보내기/가져오기/초기화
btnExport.onclick = () => openExportModal();


  fileImport.onchange = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    importFromExcelFile(f);
    e.target.value = "";
  };

  btnReset.onclick = () => {
    if (!activeProjectId) return;
    if (!confirm("현재 프로젝트를 초기화할까요?")) return;
    state = deepClone(DEFAULT_STATE);
    saveState();
    render();
  };

  // backdrop 클릭 닫기
  const modal = document.getElementById("projectModal");
  if (modal && !modal.__backdropBound) {
    modal.__backdropBound = true;
    modal.addEventListener("click", (e) => {
      const t = e.target;
      if (t && t.getAttribute && t.getAttribute("data-close") === "1") closeProjectModal();
    });
  }

  // ESC로 모달 닫기(1회만)
  if (!window.__finEscBound) {
    window.__finEscBound = true;
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      const m = document.getElementById("projectModal");
      if (!m) return;
      if (m.getAttribute("aria-hidden") === "false") closeProjectModal();
    });
  }
}


   /* ============================
   ✅ Code Picker Window (picker.html + picker.js 연동)
   - Ctrl+. 또는 버튼으로 열기
   - INIT 메시지로 codes 전달
   - picker → INSERT_SELECTED 수신 후 calc/code 셀에 반영
============================ */

let __pickerWin = null;

function getActiveCalcFocusRow(tabId) {
  const ae = document.activeElement;
  if (ae instanceof HTMLInputElement && ae.dataset.grid === "calc" && ae.dataset.tab === tabId) {
    return clamp(Number(ae.dataset.row || 0), 0, 999999);
  }
  return 0;
}

function openPickerWindow() {
  // code 탭에서는 의미가 애매하니(필요하면 허용 가능) 우선 산출탭에서만 사용 권장
  const tabId = state.activeTab;

  const isCalc = (tabId === "steel" || tabId === "steel_sub" || tabId === "support");
  if (!isCalc) {
    alert("코드 선택은 산출 탭(철골/부자재/동바리)에서 사용해 주세요.");
    return;
  }

  const focusRow = getActiveCalcFocusRow(tabId);

  // picker.html 열기 (같은 폴더에 있어야 함)
  const w = window.open("picker.html", "FIN_PICKER", "width=1100,height=820");
  if (!w) {
    alert("팝업이 차단되었습니다. 브라우저에서 팝업 허용 후 다시 시도해 주세요.");
    return;
  }
  __pickerWin = w;

  // ✅ picker.js가 INIT을 받을 때까지 약간 딜레이/재시도
  const payload = {
    type: "INIT",
    originTab: tabId,
    focusRow,
    codes: Array.isArray(state.codeMaster) ? state.codeMaster.map(x => ({
      code: x.code || "",
      name: x.name || "",
      spec: x.spec || "",
      unit: x.unit || "",
      surcharge: (x.surcharge ?? ""),
      conv_unit: (x.convUnit || ""),
      conv_factor: (x.convFactor ?? ""),
    })) : []
  };

  let tries = 0;
  const timer = setInterval(() => {
    tries++;
    try {
      w.postMessage(payload, window.location.origin);
    } catch {}
    if (tries >= 20) clearInterval(timer);
  }, 120);
}

// ✅ picker → 메인으로 삽입 요청 수신
if (!window.__finPickerMsgBound) {
  window.__finPickerMsgBound = true;

  window.addEventListener("message", (event) => {
    if (event.origin !== window.location.origin) return;
    const msg = event.data;
    if (!msg || typeof msg !== "object") return;

    // picker가 닫힐 때 알림(선택사항)
    if (msg.type === "CLOSE_PICKER") {
      try { __pickerWin = null; } catch {}
      return;
    }

        if (msg.type === "INSERT_SELECTED") {
      const tabId = msg.originTab;
      const focusRow = Number(msg.focusRow || 0);
      const codes = Array.isArray(msg.selectedCodes) ? msg.selectedCodes : [];
      if (!codes.length) return;

      const isCalc = (tabId === "steel" || tabId === "steel_sub" || tabId === "support");
      if (!isCalc) return;

      const bucket = state[tabId];
      const sec = bucket.sections[bucket.activeSection];

      // ✅ "현재 행 아래"에 삽입
      const insertPos = clamp(focusRow + 1, 0, sec.rows.length);

      // ✅ 선택 개수만큼 새 행 끼워넣기
      const newRows = Array.from({ length: codes.length }, () => defaultCalcRow());
      sec.rows.splice(insertPos, 0, ...newRows);

      // ✅ 새로 생긴 행에 코드 채우기
      codes.forEach((c, i) => {
        const r = sec.rows[insertPos + i];
        if (!r) return;
        r.code = String(c || "").trim();
      });

      recomputeSection(tabId);
      saveState();
      render();

      raf2(() => {
        const target = document.querySelector(
          `input[data-grid="calc"][data-tab="${tabId}"][data-row="${insertPos}"][data-col="${CALC_COL_INDEX.code}"]`
        );
        safeFocus(target);
        ensureScrollIntoView(target);
      });

      return; // ✅ 다른 메시지 타입으로 흐르지 않게(안전)
    }
  }); // ✅ window.addEventListener("message", ...) 닫기
} // ✅ if (!window.__finPickerMsgBound) 닫기




  /* ============================
   ✅ Project UI (index.html v22 1:1 매칭)
   - 상단: btnProject / activeProjectBadge
   - 모달: projectModal / projectTbody
   - 모달버튼: btnProjectAdd / btnProjectDelete / btnProjectSave / btnProjectClose / btnProjectOpen
============================ */

/** 상단 주요 버튼 잠금/해제 (코드선택/내보내기/가져오기/초기화) */
function setTopButtonsEnabled(enabled) {
  const btnOpen = document.getElementById("btnOpenPicker");
  const btnExport = document.getElementById("btnExport");
  const btnReset = document.getElementById("btnReset");
  const fileImport = document.getElementById("fileImport");
  const btnImportWrap = document.getElementById("btnImportWrap"); // label wrapper

  if (btnOpen) btnOpen.disabled = !enabled;
  if (btnExport) btnExport.disabled = !enabled;
  if (btnReset) btnReset.disabled = !enabled;
  if (fileImport) fileImport.disabled = !enabled;

  // label wrapper는 disabled가 안 먹어서 스타일로 잠금
  if (btnImportWrap) {
    btnImportWrap.style.opacity = enabled ? "1" : "0.55";
    btnImportWrap.style.pointerEvents = enabled ? "auto" : "none";
    btnImportWrap.setAttribute("aria-disabled", enabled ? "false" : "true");
  }

  // 도움말은 항상 사용 가능
  const help = document.getElementById("btnHelp");
  if (help) help.disabled = false;
}

/** 상단 배지 + 버튼 잠금 상태 업데이트 */
function updateProjectHeaderUI() {
  const meta = projectIndex.projects.find(p => p.id === activeProjectId);
  const badge = document.getElementById("activeProjectBadge");

  if (badge) {
    badge.textContent = meta ? `${meta.code || "-"} · ${meta.name || ""}` : "(미선택)";
  }

  setTopButtonsEnabled(!!meta);
}

/** 모달 열기 */
function openProjectModal() {
  const modal = document.getElementById("projectModal");
  if (!modal) return;

  modal.hidden = false;
  modal.setAttribute("aria-hidden", "false");

  // ✅ 테이블 렌더
  renderProjectTable();

  // ✅ 선택 상태 기본값(없으면 active 선택)
  if (!__selectedProjectIdInModal && activeProjectId) {
    __selectedProjectIdInModal = activeProjectId;
    markSelectedRow(__selectedProjectIdInModal);
  }
}

/** 모달 닫기 */
function closeProjectModal() {
  const modal = document.getElementById("projectModal");
  if (!modal) return;

  modal.hidden = true;
  modal.setAttribute("aria-hidden", "true");
}

/** 모달 내부 선택 프로젝트 id */
let __selectedProjectIdInModal = "";

/** 선택 표시 */
function markSelectedRow(pid) {
  document.querySelectorAll("#projectTbody tr").forEach(tr => {
    tr.classList.toggle("selected", tr.dataset.pid === pid);
  });
}

/** 프로젝트 테이블 렌더 */
function renderProjectTable() {
  const tbody = document.getElementById("projectTbody");
  if (!tbody) return;

  tbody.innerHTML = "";

  const items = projectIndex.projects
    .slice()
    .sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));

  if (!items.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 3;
    td.style.padding = "14px";
    td.style.color = "rgba(90,90,97,1)";
    td.style.fontWeight = "700";
    td.textContent = "프로젝트가 없습니다. [추가]로 생성하세요.";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  items.forEach(p => {
    const tr = document.createElement("tr");
    tr.dataset.pid = p.id;

    // activeProjectId와 별개로 "모달 선택" 기준으로 하이라이트
    const isSelected = (p.id === (__selectedProjectIdInModal || activeProjectId));
    if (isSelected) tr.classList.add("selected");

    const tdCode = document.createElement("td");
    const tdName = document.createElement("td");
    const tdDate = document.createElement("td");

    const inpCode = document.createElement("input");
    inpCode.className = "cell";
    inpCode.value = p.code || "";
    inpCode.placeholder = "공사코드";

    const inpName = document.createElement("input");
    inpName.className = "cell";
    inpName.value = p.name || "";
    inpName.placeholder = "프로젝트명";

    // 입력 값은 meta에 즉시 반영(저장은 '저장' 버튼)
    inpCode.addEventListener("input", () => { p.code = inpCode.value; });
    inpName.addEventListener("input", () => { p.name = inpName.value; });

    // 클릭하면 선택 표시
    const pick = () => {
      __selectedProjectIdInModal = p.id;
      markSelectedRow(p.id);
    };
    tr.addEventListener("click", pick);
    inpCode.addEventListener("click", (e) => { e.stopPropagation(); pick(); });
    inpName.addEventListener("click", (e) => { e.stopPropagation(); pick(); });

    tdCode.appendChild(inpCode);
    tdName.appendChild(inpName);

    const d = new Date(p.createdAt || p.updatedAt || Date.now());
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    tdDate.textContent = `${yyyy}-${mm}-${dd}`;

    tr.appendChild(tdCode);
    tr.appendChild(tdName);
    tr.appendChild(tdDate);

    tbody.appendChild(tr);
  });
}

/** 프로젝트 추가 */
function createProject() {
  const pid = genId();
  const meta = normalizeProjectMeta({ id: pid, name: "새 프로젝트", code: "" });

  projectIndex.projects.push(meta);
  saveProjectIndex(projectIndex);

  // 신규 기본 상태 저장
  ProjectStore.saveProjectState(pid, deepClone(DEFAULT_STATE));

  // 모달에서 바로 선택
  __selectedProjectIdInModal = pid;

  renderProjectTable();
  markSelectedRow(pid);
}

/** 모달에서 프로젝트 저장 */
function saveProjectsFromModal() {
  projectIndex.projects.forEach(p => {
    p.name = (p.name || "").trim() || "새 프로젝트";
    p.code = (p.code || "").trim();
    p.updatedAt = Date.now();
  });

  saveProjectIndex(projectIndex);
  renderProjectTable();
  updateProjectHeaderUI();
}

/** 모달에서 선택 프로젝트 삭제 */
function deleteSelectedProjectInModal() {
  const pid = __selectedProjectIdInModal || activeProjectId;
  if (!pid) return alert("삭제할 프로젝트를 선택해 주세요.");

  const meta = projectIndex.projects.find(p => p.id === pid);
  if (!confirm(`프로젝트를 삭제할까요?\n${meta?.name || ""} (${meta?.code || "-"})`)) return;

  // active 삭제면 active 해제
  if (pid === activeProjectId) {
    try { saveProjectState(activeProjectId); } catch {}
    activeProjectId = "";
    ProjectStore.saveActiveId("");
  }

  projectIndex.projects = projectIndex.projects.filter(p => p.id !== pid);
  saveProjectIndex(projectIndex);
  ProjectStore.deleteProject(pid);

  __selectedProjectIdInModal = "";

  // 프로젝트가 0개면 하나 생성
  if (projectIndex.projects.length === 0) {
    const nid = genId();
    const m = normalizeProjectMeta({ id: nid, name: "프로젝트 1", code: "" });
    projectIndex.projects.push(m);
    saveProjectIndex(projectIndex);
    ProjectStore.saveProjectState(nid, deepClone(DEFAULT_STATE));
  }

  // active가 없으면 첫 프로젝트로
  if (!activeProjectId) {
    activeProjectId = projectIndex.projects[0].id;
    ProjectStore.saveActiveId(activeProjectId);
    state = loadProjectState(activeProjectId);
  }

  renderProjectTable();
  updateProjectHeaderUI();
  render();
}

/** 모달의 "선택 프로젝트 열기" */
function openSelectedProjectFromModal() {
  const pid = __selectedProjectIdInModal || activeProjectId;
  if (!pid) return alert("열 프로젝트를 선택해 주세요.");

  selectProject(pid);
  closeProjectModal();
}

/** 프로젝트 선택(열기) */
function selectProject(projectId) {
  const meta = projectIndex.projects.find(p => p.id === projectId);
  if (!meta) return alert("프로젝트를 찾을 수 없습니다.");

  // 현재 프로젝트 저장
  if (activeProjectId) saveProjectState(activeProjectId);

  // 새 프로젝트 로드
  state = loadProjectState(projectId);

  activeProjectId = projectId;
  ProjectStore.saveActiveId(activeProjectId);

  updateProjectHeaderUI();
  render();
}

/* ============================
   ✅ Render Main (그대로 유지)
============================ */
function render() {
  if (!$view) return;

  applyTopSplitH();
  renderTabs();

  clear($view);

  let node = null;
  if (state.activeTab === "code") node = renderCodeTab();
  else if (state.activeTab === "steel") node = renderCalcTab("steel", "철골");
  else if (state.activeTab === "steel_sum") node = renderSummaryTab("steel", "철골_집계");
  else if (state.activeTab === "steel_sub") node = renderCalcTab("steel_sub", "철골_부자재");
  else if (state.activeTab === "support") node = renderCalcTab("support", "구조이기/동바리");
  else if (state.activeTab === "support_sum") node = renderSummaryTab("support", "구조이기/동바리_집계");
  else node = renderCodeTab();

  $view.appendChild(node);

  raf2(() => {
    updateStickyVars();
    applyPanelStickyTop();
    updateViewFillHeight();
    updateScrollHeights();

    if (__pendingSectionFocus && __pendingSectionFocus.tabId === state.activeTab) {
      const list = document.querySelector(`.section-list[data-tab="${__pendingSectionFocus.tabId}"]`);
      const idx = __pendingSectionFocus.index;
      const item = list?.querySelectorAll(".section-item")?.[idx];
      raf2(() => safeFocus(item));
      __pendingSectionFocus = null;
    }
  });
}

/* ============================
   ✅ Init (DOM 준비 후 1회) — index.html v22 맞춤
============================ */
let __appInited = false;
function initAppOnce() {
  console.log("[FIN] initAppOnce fired");
  if (__appInited) return;
  __appInited = true;

  // ✅ 1) 먼저 상단 버튼들(프로젝트 포함)을 bindTopButtonsOnce에서 바인딩
  //    (여기 안에서 btnProject를 이미 getElementById로 잡고 onclick을 걸어줌)
  try { bindTopButtonsOnce(); } catch (e) { console.warn(e); }

  /* ============================
     ✅ Shift+좌클릭 셀 블록지정 이벤트(1회 바인딩)
     - input.cell만 대상
     - ShiftKey면 anchor~target 사각형 블록 지정
     ============================ */
    if (!window.__finCellBlockBound) {
    window.__finCellBlockBound = true;

    // ✅ 산출표(calc)에서만 Shift+클릭/일반클릭을 "행 선택" 용도로 처리
    // (코드/변수표 등 셀 블록 선택은 아래쪽 __finCellBlockBound2 로직이 담당)
    document.addEventListener("mousedown", (e) => {
      const input = e.target?.closest?.("input.cell");
      if (!(input instanceof HTMLInputElement)) return;

      const grid = input.dataset.grid || "";
      const tabId = input.dataset.tab || "";
      const row = Number(input.dataset.row || 0);

      const isCalcTab =
        grid === "calc" &&
        (tabId === "steel" || tabId === "steel_sub" || tabId === "support");

      // calc 탭이 아니면 여기서는 아무것도 하지 않음(다른 핸들러가 처리)
      if (!isCalcTab) return;

      // ================================
      // ✅ Shift + 좌클릭 : 행 범위 선택
      // ================================
      if (e.shiftKey) {
        e.preventDefault(); // 텍스트 드래그 방지

        // mousedown 시점에는 포커스가 아직 이동 전 → 기존 포커스 행을 anchor로 사용
        let anchor = row;
        const ae = document.activeElement;
        if (
          ae instanceof HTMLInputElement &&
          ae.dataset.grid === "calc" &&
          ae.dataset.tab === tabId
        ) {
          anchor = Number(ae.dataset.row || row);
        } else if (__calcMulti.anchorRow != null) {
          anchor = Number(__calcMulti.anchorRow);
        }

        // 컨텍스트가 다르면 anchor 기준으로 시작
        if (!__calcMultiIsSameContext(tabId)) {
          __calcMultiBegin(tabId, anchor);
        } else {
          __calcMulti.anchorRow = anchor;
        }

        __calcMultiSetRange(tabId, anchor, row);
        __applyCalcRowSelectionStyles(tabId);
        return;
      }

      // ================================
      // ✅ 일반 클릭 : 기존 선택은 유지, anchor만 갱신
      // ================================
      __calcMulti.anchorRow = row;
      // (의도적으로 return;  다른 셀 블록 선택 로직과 충돌 방지)
      return;

    }, true);
  }


      /* ============================
   ✅ Global Hotkeys (Ctrl+., Ctrl+B, Ctrl+F3, Ctrl+F10)
   - 프로젝트 선택된 상태에서만 작동
   - input/textarea 편집중에는 일부 단축키 무시
============================ */
let __globalHotkeysBound = false;

function bindGlobalHotkeysOnce() {
  if (__globalHotkeysBound) return;
  __globalHotkeysBound = true;

  document.addEventListener("keydown", (e) => {
    // 프로젝트 미선택이면 단축키 동작 X
    if (!activeProjectId) return;

    const ae = document.activeElement;

    // -------------------------
    // Ctrl + . : 코드 선택창
    // -------------------------
    if (e.ctrlKey && !e.shiftKey && !e.altKey && e.key === ".") {
      e.preventDefault();
      e.stopPropagation();

      const btn = document.getElementById("btnOpenPicker");
      if (btn && !btn.disabled) btn.click();
      else alert("프로젝트를 먼저 선택(열기)해 주세요.");
      return;
    }

    // -------------------------
    // Ctrl + B : 현재 행 선택 토글
    // -------------------------
    if (e.ctrlKey && !e.shiftKey && !e.altKey && (e.key === "b" || e.key === "B")) {
      e.preventDefault();
      e.stopPropagation();

      const tabId = state.activeTab;
      const isCalc = (tabId === "steel" || tabId === "steel_sub" || tabId === "support");
      if (!isCalc) return;

      let curRow = 0;
      if (ae instanceof HTMLInputElement && ae.dataset.grid === "calc" && ae.dataset.tab === tabId) {
        curRow = Number(ae.dataset.row || 0);
      }

      __calcMultiToggleRow(tabId, curRow);
      __applyCalcRowSelectionStyles(tabId);
      return;
    }

    // -------------------------
    // Ctrl + F3 : 행 추가 / Shift+Ctrl+F3 : +10행
    // -------------------------
    if (e.ctrlKey && (e.key === "F3")) {
      e.preventDefault();
      e.stopPropagation();

      const n = e.shiftKey ? 10 : 1;
      const tabId = state.activeTab;

      // 코드탭
      if (tabId === "code") {
        let insertAfter = null;
        if (ae instanceof HTMLInputElement && ae.dataset.grid === "code") {
          insertAfter = Number(ae.dataset.row || 0);
        }
        addCodeRows(n, insertAfter);
        return;
      }

      // 산출탭
      if (tabId === "steel" || tabId === "steel_sub" || tabId === "support") {
        let insertAfter = null;
        if (ae instanceof HTMLInputElement && ae.dataset.grid === "calc" && ae.dataset.tab === tabId) {
          insertAfter = Number(ae.dataset.row || 0);
        }
        addRows(tabId, n, insertAfter);
        return;
      }

      return;
    }

        // =========================
    // Ctrl + F10 : 아래로 1행 추가 + 코드 자동입력(REMARK_CODE)
    // =========================
    if (e.ctrlKey && !e.shiftKey && !e.altKey && e.key === "F10") {
      e.preventDefault();
      e.stopPropagation();

      const tabId = state?.activeTab;
      const isCalcTab = (tabId === "steel" || tabId === "steel_sub" || tabId === "support");
      if (!isCalcTab) return;

      // 1) 현재 포커스가 산출표 셀인지 확인
      const ae = document.activeElement;
      if (!(ae instanceof HTMLInputElement)) return;

      const row = Number(ae.dataset.row);
      const grid = ae.dataset.grid;
      const tab = ae.dataset.tab;

      // 산출표 셀인지 최소 검증
      if (grid !== "calc" || tab !== tabId || !Number.isFinite(row)) {
        alert("산출표 셀을 선택한 상태에서 Ctrl+F10을 눌러주세요.");
        return;
      }

      // 2) 현재 행 아래에 1행 추가
      addRows(tabId, 1, row);
      const newRow = row + 1;

      // 3) 새 행의 '코드' 셀 찾아 값 입력 + input 이벤트 발생
      raf2(() => {
        const codeInput = document.querySelector(
          `input[data-grid="calc"][data-tab="${tabId}"][data-row="${newRow}"][data-col="${CALC_COL_INDEX.code}"]`
        );

        if (!codeInput) {
          alert("새로 추가된 행의 코드 입력칸을 찾지 못했습니다.");
          return;
        }

        codeInput.value = REMARK_CODE;
codeInput.dispatchEvent(new Event("input", { bubbles: true }));
codeInput.dispatchEvent(new Event("change", { bubbles: true }));

syncRemarkRowFromCodeInput(codeInput);   // ✅ 추가: 비고행 회색 처리 클래스 부여

         

safeFocus(codeInput);
ensureScrollIntoView(codeInput);

      });

      return;
    }



  }, true); // ✅ keydown 리스너 닫기
} // ✅ bindGlobalHotkeysOnce 닫기




       



  // ✅ 2) initAppOnce에서 다시 btnProject를 “중복 바인딩”하지 않는다
  //    (중복 바인딩은 필요 없고, TDZ 에러 원인이 됐음)

  // ✅ 3) 모달 닫기(backdrop/ESC)는 1회만 걸리도록 가드 추가
  const modal = document.getElementById("projectModal");
  if (modal && !modal.__finModalBound) {
    modal.__finModalBound = true;

    modal.addEventListener("click", (e) => {
      const t = e.target;
      if (t && t.getAttribute && t.getAttribute("data-close") === "1") closeProjectModal();
    });

    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      const m = document.getElementById("projectModal");
      if (!m) return;
      if (m.getAttribute("aria-hidden") === "false") closeProjectModal();
    });
  }

    // ✅ 4) 최초 UI 반영
  updateProjectHeaderUI();
  render();

  // ✅ 5) 전역 단축키 바인딩 (Ctrl+., Ctrl+B, Ctrl+F3 등)
  bindGlobalHotkeysOnce();

  raf2(() => {
    updateStickyVars();
    applyPanelStickyTop();
    updateViewFillHeight();
    updateScrollHeights();
  });
}


   function getRemarkItemFromCodeMaster() {
  const REMARK_CODE_LOCAL = "ZZZZZZZZZZZZZZZZZ";
  const cm = Array.isArray(state?.codeMaster) ? state.codeMaster : [];

  // 1) 코드로 먼저 찾기
  let it = cm.find(x => String(x?.code || "").trim().toUpperCase() === REMARK_CODE_LOCAL.toUpperCase());
  if (it) return normalizeCodeItem(it);

  // 2) 품명/상품명으로 찾기 ("비고")
  it = cm.find(x => {
    const pn = (x?.name || x?.productName || x?.["품명"] || x?.Product || "").toString().trim();
    return pn === "비고" || pn === "[비          고]";
  });
  if (it) return normalizeCodeItem(it);

  return null;
}

// codeMaster 항목 구조가 섞여 있어도 picker 삽입 함수가 먹는 형태로 정규화
function normalizeCodeItem(it) {
  return {
    code: it.code ?? it.Code ?? "",
    productName: it.productName ?? it.name ?? it["품명"] ?? it.Product ?? "",
    specs: it.specs ?? it.spec ?? it["규격"] ?? it.Specifications ?? "",
    unit: it.unit ?? it["단위"] ?? it.Unit ?? "",
    surcharge: it.surcharge ?? it["할증"] ?? "",
    convUnit: it.convUnit ?? it.convUnit ?? it["환산단위"] ?? "",
    convFactor: it.convFactor ?? it["환산계수"] ?? "",
    note: it.note ?? it["비고"] ?? ""
  };
}




   // =========================================================
// ✅ Ctrl+Z : 블록 선택 / 행 선택 해제 (1회 바인딩)
// - 선택이 있을 때만 가로채고
// - 선택이 없으면 기본 Undo 동작 유지
// =========================================================
if (!window.__finClearSelectionHotkeyBound) {
  window.__finClearSelectionHotkeyBound = true;

  document.addEventListener("keydown", (e) => {
    // Ctrl + Z
    if (!(e.ctrlKey && !e.shiftKey && !e.altKey && (e.key === "z" || e.key === "Z"))) {
      return;
    }

    const hasCellBlock =
      !!document.querySelector("input.cell.block-selected");

    const hasCalcMulti =
      !!__calcMulti && __calcMulti.active;

    // ✅ 선택이 없으면 → 원래 Ctrl+Z(Undo) 그대로
    if (!hasCellBlock && !hasCalcMulti) {
      return;
    }

    // ✅ 선택이 있으면 → 해제 전용 단축키로 사용
    e.preventDefault();
    e.stopPropagation();

    // 🔹 셀 블록 해제
    if (hasCellBlock) {
      __clearCellBlockSelection();
      __finBlockSel.anchor = null;
    }

    // 🔹 산출표 행 블록 해제
    if (hasCalcMulti) {
      __calcMultiClear();
      const tabId = state.activeTab;
      if (tabId === "steel" || tabId === "steel_sub" || tabId === "support") {
        __applyCalcRowSelectionStyles(tabId);
      }
    }
  }, true);
}


/* =========================================================
   ✅ Shift + Click 셀 블록지정 (input.cell)
   - data-grid / data-tab / data-row / data-col 기준 사각형 선택
   - 선택된 input.cell에 .block-selected 클래스 부여
   - code(grid="code")는 data-tab이 없으므로 grid만으로 그룹핑
   ========================================================= */
const __finBlockSel = {
  anchor: null, // { grid, tab, row, col }
};

function __getCellKey(input) {
  const ds = input?.dataset || {};
  const grid = ds.grid || "";
  const tab = ds.tab || "";     // code는 없음
  const row = Number(ds.row || 0);
  const col = Number(ds.col || 0);
  return { grid, tab, row, col };
}

function __sameContext(a, b) {
  if (!a || !b) return false;
  if (a.grid === "code" && b.grid === "code") return true;
  return a.grid === b.grid && a.tab === b.tab;
}

function __queryAllCellsInContext(key) {
  if (!key?.grid) return [];
  if (key.grid === "code") {
    return Array.from(document.querySelectorAll(`input.cell[data-grid="code"]`));
  }
  return Array.from(document.querySelectorAll(
    `input.cell[data-grid="${key.grid}"][data-tab="${key.tab}"]`
  ));
}

function __clearCellBlockSelection() {
  document.querySelectorAll("input.cell.block-selected").forEach((x) => {
    x.classList.remove("block-selected");
  });
}

function __setAnchor(input) {
  const k = __getCellKey(input);
  if (!k.grid) return;
  __finBlockSel.anchor = k;
}

function __applyCellBlockSelection(anchorKey, targetKey) {
  if (!anchorKey?.grid || !targetKey?.grid) return;
  if (!__sameContext(anchorKey, targetKey)) return;

  const cells = __queryAllCellsInContext(anchorKey);
  if (!cells.length) return;

  const r1 = Math.min(anchorKey.row, targetKey.row);
  const r2 = Math.max(anchorKey.row, targetKey.row);
  const c1 = Math.min(anchorKey.col, targetKey.col);
  const c2 = Math.max(anchorKey.col, targetKey.col);

  for (const inp of cells) {
    const k = __getCellKey(inp);
    if (k.row >= r1 && k.row <= r2 && k.col >= c1 && k.col <= c2) {
      inp.classList.add("block-selected");
    } else {
      inp.classList.remove("block-selected");
    }
  }
}

  // =========================================================
  // ✅ 일반 클릭 시 앵커만 갱신 + (필요 시) 기존 블록 선택 해제
  // - calc-grid는 위에서 "행 선택" 전용 mousedown 핸들러가 처리 중이므로
  //   여기서는 code/var 등 "셀 블록 선택" 전용으로 처리
  // =========================================================
        function __handleNormalClickCell(input) {
      if (!(input instanceof HTMLInputElement)) return;

      // ✅ calc는 행선택 로직이 전담하므로 여기서 제외
      const grid = input.dataset.grid || "";
      if (grid === "calc") return;

      // 기존 블록 선택은 일반 클릭이면 해제
      __clearCellBlockSelection();
      __setAnchor(input);
    }

    

    function __handleShiftClickCell(input) {
      if (!(input instanceof HTMLInputElement)) return;

      const grid = input.dataset.grid || "";
      // ✅ calc는 위에서 행선택 전용으로 처리하므로 셀블록은 제외
      if (grid === "calc") return;

      const targetKey = __getCellKey(input);
      if (!targetKey.grid) return;

      // anchor가 없거나 컨텍스트 다르면 anchor를 현재로 잡고 1셀만 선택
      if (!__finBlockSel.anchor || !__sameContext(__finBlockSel.anchor, targetKey)) {
        __clearCellBlockSelection();
        __setAnchor(input);
        input.classList.add("block-selected");
        return;
      }

      // 기존 선택 해제 후, 사각형 블록 선택 적용
      __clearCellBlockSelection();
      __applyCellBlockSelection(__finBlockSel.anchor, targetKey);
    }

    // ✅ 선택 셀들을 TSV로 복사 (Ctrl+C)
    function __copySelectedBlockToClipboard() {
      const selected = Array.from(document.querySelectorAll("input.cell.block-selected"));
      if (!selected.length) return false;

      // 컨텍스트 기준(같은 grid/tab만)
      const firstKey = __getCellKey(selected[0]);
      const ctxCells = selected
        .map(inp => ({ inp, k: __getCellKey(inp) }))
        .filter(x => __sameContext(firstKey, x.k));

      if (!ctxCells.length) return false;

      // 범위 계산
      const rows = ctxCells.map(x => x.k.row);
      const cols = ctxCells.map(x => x.k.col);
      const r1 = Math.min(...rows), r2 = Math.max(...rows);
      const c1 = Math.min(...cols), c2 = Math.max(...cols);

      // (row,col)->value 맵
      const map = new Map();
      ctxCells.forEach(({ inp, k }) => {
        map.set(`${k.row},${k.col}`, inp.value ?? "");
      });

      // TSV 만들기
      const lines = [];
      for (let r = r1; r <= r2; r++) {
        const line = [];
        for (let c = c1; c <= c2; c++) {
          line.push(String(map.get(`${r},${c}`) ?? ""));
        }
        lines.push(line.join("\t"));
      }
      const tsv = lines.join("\n");

      try {
        navigator.clipboard?.writeText(tsv);
      } catch {
        // fallback
        const ta = document.createElement("textarea");
        ta.value = tsv;
        ta.style.position = "fixed";
        ta.style.left = "-9999px";
        ta.style.top = "0";
        document.body.appendChild(ta);
        ta.focus();
        ta.select();
        try { document.execCommand("copy"); } catch {}
        document.body.removeChild(ta);
      }
      return true;
    }

    // ✅ 블록선택/앵커 관련 전역 이벤트 바인딩(1회)
    if (!window.__finCellBlockBound2) {
      window.__finCellBlockBound2 = true;

      // (1) 클릭 처리: Shift면 블록선택 / 아니면 앵커 갱신 + 필요 시 기존 해제
      document.addEventListener("mousedown", (e) => {
        const input = e.target?.closest?.("input.cell");
        if (!(input instanceof HTMLInputElement)) return;

        // calc는 위에서 행선택 전용 mousedown이 처리하므로 여기서는 건너뜀
        if ((input.dataset.grid || "") === "calc") return;

        if (e.shiftKey) {
          e.preventDefault(); // 드래그 방지
          __handleShiftClickCell(input);
        } else {
          // 일반 클릭: 기존 선택 해제(선택이 있었고, 같은 컨텍스트가 아니면) + anchor 갱신
          __handleNormalClickCell(input);
        }
      }, true);

      // (2) 바깥 클릭하면 블록선택 해제(원하면 유지도 가능하지만, 보통 해제하는 게 UX 좋음)
      document.addEventListener("mousedown", (e) => {
        const input = e.target?.closest?.("input.cell");
        if (input) return; // 셀 클릭이면 유지
        // 모달/버튼 클릭 등에서도 유지하고 싶으면 여기 조건 추가 가능
        __clearCellBlockSelection();
      }, true);

      // (3) Ctrl+C : 선택된 블록이 있을 때만 가로채서 복사
      document.addEventListener("keydown", (e) => {
        if (!(e.ctrlKey && !e.shiftKey && !e.altKey && (e.key === "c" || e.key === "C"))) return;

        // input 편집 중이면(커서 선택 복사) 기본 동작 유지
        const ae = document.activeElement;
        if (ae instanceof HTMLInputElement && ae.dataset.editing === "1") return;

        const hasBlock = !!document.querySelector("input.cell.block-selected");
        if (!hasBlock) return;

        e.preventDefault();
        e.stopPropagation();
        __copySelectedBlockToClipboard();
      }, true);
    }

  } // ✅ initAppOnce() 블록 끝

  // ✅ DOMContentLoaded 시 init
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initAppOnce);
  } else {
    initAppOnce();
  }

})(); // ✅ IIFE 끝








