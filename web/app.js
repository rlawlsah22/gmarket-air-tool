// ─────────────────────────────────────────────
// 전역 메타데이터
// ─────────────────────────────────────────────
let AIRPORTS = {};
let AIRLINE_MODES = [];
let SPECIFIC_AIRLINES = [];
let AIRPREMIA_DESTINATIONS = {};
let TIME_SLOT_LABELS = [];
let TIME_SLOT_KEYS = [];

const $ = (id) => document.getElementById(id);

function waitForApi() {
  return new Promise((resolve) => {
    if (window.pywebview && window.pywebview.api) return resolve();
    window.addEventListener("pywebviewready", () => resolve());
  });
}

function fmtDate(d) {
  const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, "0"), day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

// ─────────────────────────────────────────────
// 페이지별 컨텍스트 (G마켓 / 에어프레미아 각각 독립 상태)
// ─────────────────────────────────────────────
function makePageContext(prefix, kind) {
  // prefix: "gm" | "ap"   kind: "gmarket" | "airpremia"
  return {
    prefix, kind,
    setMode: false,
    setBlocks: [],
    presets: {},
    running: false,
  };
}

const ctxGmarket = makePageContext("gm", "gmarket");
const ctxAirpremia = makePageContext("ap", "airpremia");

// ─────────────────────────────────────────────
// 초기화
// ─────────────────────────────────────────────
async function init() {
  await waitForApi();
  const meta = await window.pywebview.api.get_meta();
  AIRPORTS = meta.airports;
  AIRLINE_MODES = meta.airline_modes;
  SPECIFIC_AIRLINES = meta.specific_airlines;
  AIRPREMIA_DESTINATIONS = meta.airpremia_destinations;
  TIME_SLOT_LABELS = meta.time_slot_labels;
  TIME_SLOT_KEYS = meta.time_slot_keys;
  ctxGmarket.presets = meta.presets_gmarket || {};
  ctxAirpremia.presets = meta.presets_airpremia || {};

  bindSidebar();
  setupGmarketPage(meta.default_out_dir);
  setupAirpremiaPage(meta.default_out_dir);
}

// ─────────────────────────────────────────────
// 사이드바 네비게이션
// ─────────────────────────────────────────────
function bindSidebar() {
  document.querySelectorAll(".nav-item").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".nav-item").forEach((b) => b.classList.toggle("active", b === btn));
      const page = btn.dataset.page;
      document.querySelectorAll(".page").forEach((p) => (p.style.display = "none"));
      $(`page-${page}`).style.display = "block";
    });
  });
}

// ─────────────────────────────────────────────
// G마켓 페이지 설정
// ─────────────────────────────────────────────
function setupGmarketPage(defaultOutDir) {
  const ctx = ctxGmarket;

  populateCountrySelect($("gm-origin-country"), "국내");
  populateCitySelect($("gm-origin-city"), "국내", "인천");
  populateCountrySelect($("gm-dest-country"), "일본");
  populateCitySelect($("gm-dest-city"), "일본", "삿포로");
  $("gm-origin-country").addEventListener("change", () => populateCitySelect($("gm-origin-city"), $("gm-origin-country").value));
  $("gm-dest-country").addEventListener("change", () => populateCitySelect($("gm-dest-city"), $("gm-dest-country").value));

  $("gm-out-dir").value = defaultOutDir;
  const today = new Date();
  $("gm-date-from").value = fmtDate(new Date(today.getTime() + 14 * 86400000));
  $("gm-date-to").value = fmtDate(new Date(today.getTime() + 21 * 86400000));

  buildSingleCondition(ctx);
  refreshPresetList(ctx);
  bindPageEvents(ctx);

  addLog(ctx, "시스템 준비 완료. 조건을 설정하고 '추출 시작'을 눌러주세요.", "info");
  addLog(ctx, "오류 등 문의 사항은 진모에게 문의 주세요.");
}

// ─────────────────────────────────────────────
// 에어프레미아 페이지 설정
// ─────────────────────────────────────────────
function setupAirpremiaPage(defaultOutDir) {
  const ctx = ctxAirpremia;

  $("ap-out-dir").value = defaultOutDir;
  const today = new Date();
  $("ap-date-from").value = fmtDate(new Date(today.getTime() + 14 * 86400000));
  $("ap-date-to").value = fmtDate(new Date(today.getTime() + 21 * 86400000));

  buildSingleCondition(ctx);
  refreshPresetList(ctx);
  bindPageEvents(ctx);

  addLog(ctx, "시스템 준비 완료. 조건을 설정하고 '추출 시작'을 눌러주세요.", "info");
  addLog(ctx, "오류 등 문의 사항은 진모에게 문의 주세요.");
}

// ─────────────────────────────────────────────
// 공통 페이지 이벤트 바인딩 (프리셋/세트모드/실행/중지/폴더선택)
// ─────────────────────────────────────────────
function bindPageEvents(ctx) {
  const pageEl = $(`page-${ctx.kind}`);

  pageEl.querySelector('[data-act="toggle-set-mode"]').addEventListener("click", () => {
    ctx.setMode = !ctx.setMode;
    pageEl.querySelector('[data-act="toggle-set-mode"]').textContent = ctx.setMode ? "📑 단일 모드로 전환" : "📑 검색 조건 세트 활성화";
    $(`${ctx.prefix}-condition-single-wrap`).style.display = ctx.setMode ? "none" : "block";
    $(`${ctx.prefix}-condition-set-wrap`).style.display = ctx.setMode ? "block" : "none";
    $(`${ctx.prefix}-add-set-row`).style.display = ctx.setMode ? "flex" : "none";
    $(`${ctx.prefix}-set-mode-hint`).style.display = ctx.setMode ? "none" : "inline";
    if (ctx.setMode && ctx.setBlocks.length === 0) addConditionSet(ctx);
  });

  pageEl.querySelector('[data-act="add-set"]').addEventListener("click", () => addConditionSet(ctx));

  pageEl.querySelector('[data-act="choose-dir"]').addEventListener("click", async () => {
    const dir = await window.pywebview.api.choose_dir();
    if (dir) $(`${ctx.prefix}-out-dir`).value = dir;
  });

  pageEl.querySelector('[data-act="preset-save"]').addEventListener("click", () => onSavePreset(ctx));
  pageEl.querySelector('[data-act="preset-load"]').addEventListener("click", () => onLoadPreset(ctx));
  pageEl.querySelector('[data-act="preset-delete"]').addEventListener("click", () => onDeletePreset(ctx));
  pageEl.querySelector('[data-act="preset-new"]').addEventListener("click", () => { $(`${ctx.prefix}-preset-input`).value = ""; });

  pageEl.querySelector('[data-act="run"]').addEventListener("click", () => onRun(ctx));
  pageEl.querySelector('[data-act="stop"]').addEventListener("click", () => onStop(ctx));
}

function populateCountrySelect(sel, selected) {
  sel.innerHTML = "";
  Object.keys(AIRPORTS).forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c; opt.textContent = c;
    if (c === selected) opt.selected = true;
    sel.appendChild(opt);
  });
}

function populateCitySelect(sel, country, selected) {
  sel.innerHTML = "";
  const cities = Object.keys(AIRPORTS[country] || {});
  cities.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c; opt.textContent = c;
    if (c === selected || (!selected && c === cities[0])) opt.selected = true;
    sel.appendChild(opt);
  });
}

// ─────────────────────────────────────────────
// 조건 블록 생성/관리
// ─────────────────────────────────────────────
function buildSingleCondition(ctx) {
  const wrap = $(`${ctx.prefix}-condition-single-wrap`);
  wrap.innerHTML = "";
  const block = ctx.kind === "gmarket" ? makeGmarketConditionBlock(null, "3") : makeAirpremiaConditionBlock(null, "3");
  wrap.appendChild(block.el);
  wrap._getData = block.getData;
  wrap._setData = block.setData;
}

function addConditionSet(ctx) {
  const idx = ctx.setBlocks.length + 1;
  const nightsDefault = String(2 + idx);
  const block = ctx.kind === "gmarket"
    ? makeGmarketConditionBlock(idx, nightsDefault, true, ctx)
    : makeAirpremiaConditionBlock(idx, nightsDefault, true, ctx);
  ctx.setBlocks.push(block);
  $(`${ctx.prefix}-condition-set-wrap`).appendChild(block.el);
  relabelSets(ctx);
}

function relabelSets(ctx) {
  ctx.setBlocks.forEach((b, i) => {
    b.el.querySelector(".set-badge").textContent = i + 1;
    b.el.querySelector(".set-title").textContent = `세트 ${i + 1}`;
    b.el.classList.toggle("set-tier1", i % 2 === 0);
    b.el.classList.toggle("set-tier2", i % 2 === 1);
  });
}

function makeSetHeader(idx, onDelete) {
  const hdr = document.createElement("div");
  hdr.className = "set-header";
  hdr.innerHTML = `<div class="set-badge">${idx}</div><div class="set-title">세트 ${idx}</div>`;
  const delBtn = document.createElement("button");
  delBtn.className = "btn btn-danger-soft btn-delete-set";
  delBtn.textContent = "✕ 세트 삭제";
  delBtn.addEventListener("click", onDelete);
  hdr.appendChild(delBtn);
  return hdr;
}

function makeGmarketConditionBlock(setIdx, defaultNights, showHeader = false, ctx = null) {
  const el = document.createElement("div");
  el.className = "condition-block";

  if (showHeader) {
    el.appendChild(makeSetHeader(setIdx, () => {
      ctx.setBlocks = ctx.setBlocks.filter((b) => b.el !== el);
      el.remove();
      relabelSets(ctx);
    }));
  }

  const top = document.createElement("div");
  top.className = "row-inline";
  top.innerHTML = `
    <span class="field-label">여행 일수</span>
    <input type="number" class="input-xs nights" min="1" max="20" value="${defaultNights}">
    <span class="unit">일</span>
    <span class="field-label" style="margin-left:16px;">항공사 모드</span>
    <select class="sel-md airline-mode" style="width:180px;"></select>
  `;
  el.appendChild(top);
  const modeSel = top.querySelector(".airline-mode");
  AIRLINE_MODES.forEach((m) => {
    const opt = document.createElement("option");
    opt.value = m; opt.textContent = m;
    modeSel.appendChild(opt);
  });

  const airlineChecksWrap = document.createElement("div");
  airlineChecksWrap.className = "airline-checks";
  airlineChecksWrap.style.display = "none";
  SPECIFIC_AIRLINES.forEach((a) => {
    const label = document.createElement("label");
    label.innerHTML = `<input type="checkbox" value="${a}"> ${a}`;
    airlineChecksWrap.appendChild(label);
  });
  el.appendChild(airlineChecksWrap);
  modeSel.addEventListener("change", () => {
    airlineChecksWrap.style.display = modeSel.value === "특정 항공사 지정" ? "grid" : "none";
  });

  const bandTitle = document.createElement("div");
  bandTitle.textContent = "시간대 조건";
  bandTitle.style.fontSize = "12px";
  bandTitle.style.fontWeight = "700";
  bandTitle.style.marginTop = "14px";
  el.appendChild(bandTitle);

  const bandGrid = document.createElement("div");
  bandGrid.className = "band-grid";
  const bandCols = [
    ["가는편 출발", "dep_band"], ["가는편 도착", "arr_band"],
    ["오는편 출발", "ret_dep_band"], ["오는편 도착", "ret_arr_band"],
  ];
  const bandRefs = {};
  bandCols.forEach(([label, key]) => {
    const col = document.createElement("div");
    const title = document.createElement("div");
    title.className = "band-col-title";
    title.textContent = label;
    col.appendChild(title);

    const custom = document.createElement("div");
    custom.className = "band-custom";
    custom.innerHTML = `<span style="font-size:10px;color:#9CA3AF;">직접:</span><input type="text" class="cf" placeholder="00:00"><span>~</span><input type="text" class="ct" placeholder="24:00">`;
    col.appendChild(custom);

    const allLabel = document.createElement("label");
    allLabel.className = "band-slot";
    allLabel.innerHTML = `<input type="checkbox" class="band-all" checked> 전체`;
    col.appendChild(allLabel);

    const slotChecks = [];
    TIME_SLOT_LABELS.forEach((slotLbl) => {
      const slotLabel = document.createElement("label");
      slotLabel.className = "band-slot";
      slotLabel.innerHTML = `<input type="checkbox" class="band-slot-chk" checked> ${slotLbl}`;
      col.appendChild(slotLabel);
      slotChecks.push(slotLabel.querySelector("input"));
    });

    const allChk = allLabel.querySelector("input");
    allChk.addEventListener("change", () => slotChecks.forEach((c) => (c.checked = allChk.checked)));
    slotChecks.forEach((c) => c.addEventListener("change", () => {
      allChk.checked = slotChecks.every((sc) => sc.checked);
    }));

    bandRefs[key] = { fromInput: custom.querySelector(".cf"), toInput: custom.querySelector(".ct"), slotChecks };
    bandGrid.appendChild(col);
  });
  el.appendChild(bandGrid);

  function getData() {
    const specific = Array.from(airlineChecksWrap.querySelectorAll("input:checked")).map((c) => c.value);
    const bands = {};
    Object.entries(bandRefs).forEach(([key, refs]) => {
      const f = refs.fromInput.value.trim(), t = refs.toInput.value.trim();
      if (f || t) {
        bands[key] = { type: "range", from: f || "00:00", to: t || "24:00" };
      } else {
        const slots = TIME_SLOT_KEYS.filter((_, i) => refs.slotChecks[i].checked);
        bands[key] = { type: "slots", slots: slots.length ? slots : TIME_SLOT_KEYS.slice() };
      }
    });
    return {
      nights: parseInt(top.querySelector(".nights").value, 10) || 1,
      airline_mode: modeSel.value,
      specific_airlines: specific,
      bands,
    };
  }

  function setData(data) {
    if (data.nights) top.querySelector(".nights").value = data.nights;
    if (data.airline_mode) {
      modeSel.value = data.airline_mode;
      airlineChecksWrap.style.display = data.airline_mode === "특정 항공사 지정" ? "grid" : "none";
    }
    if (data.specific_airlines) {
      airlineChecksWrap.querySelectorAll("input").forEach((c) => (c.checked = data.specific_airlines.includes(c.value)));
    }
    if (data.bands) {
      Object.entries(data.bands).forEach(([key, cond]) => {
        const refs = bandRefs[key];
        if (!refs) return;
        if (cond.type === "range") {
          refs.fromInput.value = cond.from === "00:00" ? "" : cond.from;
          refs.toInput.value = cond.to === "24:00" ? "" : cond.to;
        } else if (cond.type === "slots") {
          refs.slotChecks.forEach((chk, i) => (chk.checked = cond.slots.includes(TIME_SLOT_KEYS[i])));
        }
      });
    }
  }

  return { el, getData, setData, kind: "gmarket" };
}

function makeAirpremiaConditionBlock(setIdx, defaultNights, showHeader = false, ctx = null) {
  const el = document.createElement("div");
  el.className = "condition-block";

  if (showHeader) {
    el.appendChild(makeSetHeader(setIdx, () => {
      ctx.setBlocks = ctx.setBlocks.filter((b) => b.el !== el);
      el.remove();
      relabelSets(ctx);
    }));
  }

  const row = document.createElement("div");
  row.className = "row-inline";
  row.innerHTML = `
    <span class="field-label">목적지</span>
    <span class="note" style="margin-right:14px;">방콕 (BKK) 고정</span>
    <span class="field-label">박수</span>
    <input type="number" class="input-xs nights" min="1" max="20" value="${defaultNights}">
    <span class="unit">박</span>
    <span class="field-label" style="margin-left:14px;">인원수</span>
    <input type="number" class="input-xs adt" min="1" max="9" value="4">
    <span class="unit">명</span>
  `;
  el.appendChild(row);
  const FIXED_DEST = Object.keys(AIRPREMIA_DESTINATIONS)[0];

  function getData() {
    return {
      dest: FIXED_DEST,
      nights: parseInt(row.querySelector(".nights").value, 10) || 1,
      adt: parseInt(row.querySelector(".adt").value, 10) || 4,
    };
  }
  function setData(data) {
    if (data.nights) row.querySelector(".nights").value = data.nights;
    if (data.adt) row.querySelector(".adt").value = data.adt;
  }

  return { el, getData, setData, kind: "airpremia" };
}

// ─────────────────────────────────────────────
// 프리셋
// ─────────────────────────────────────────────
function refreshPresetList(ctx) {
  const dl = $(`${ctx.prefix}-preset-list`);
  dl.innerHTML = "";
  Object.keys(ctx.presets).forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    dl.appendChild(opt);
  });
}

function currentRouteAndCondition(ctx) {
  const data = {
    adults: ctx.kind === "gmarket" ? (parseInt($("gm-adults").value, 10) || 4) : undefined,
    set_mode: ctx.setMode,
  };
  if (ctx.kind === "gmarket") {
    data.origin_country = $("gm-origin-country").value;
    data.origin_city = $("gm-origin-city").value;
    data.dest_country = $("gm-dest-country").value;
    data.dest_city = $("gm-dest-city").value;
  }
  if (ctx.setMode) {
    data.sets = ctx.setBlocks.map((b) => b.getData());
  } else {
    data.sets = [$(`${ctx.prefix}-condition-single-wrap`)._getData()];
  }
  return data;
}

async function onSavePreset(ctx) {
  const inputEl = $(`${ctx.prefix}-preset-input`);
  let name = inputEl.value.trim();
  if (!name) {
    addLog(ctx, "⚠ 프리셋 이름을 입력하세요.", "err");
    return;
  }
  if (ctx.presets[name]) {
    const ok = confirm(`'${name}' 프리셋이 이미 있습니다.\n현재 설정으로 덮어쓸까요?`);
    if (!ok) return;
  }
  ctx.presets[name] = currentRouteAndCondition(ctx);
  await window.pywebview.api.save_presets(ctx.kind, ctx.presets);
  refreshPresetList(ctx);
  addLog(ctx, `'${name}' 프리셋이 저장되었습니다.`, "ok");
}

function onLoadPreset(ctx) {
  const name = $(`${ctx.prefix}-preset-input`).value.trim();
  const data = ctx.presets[name];
  if (!data) {
    addLog(ctx, "⚠ 불러올 프리셋을 선택하세요.", "err");
    return;
  }
  if (ctx.kind === "gmarket" && typeof data.adults === "number") $("gm-adults").value = data.adults;
  if (ctx.kind === "gmarket" && data.origin_country) {
    $("gm-origin-country").value = data.origin_country;
    populateCitySelect($("gm-origin-city"), data.origin_country, data.origin_city);
    $("gm-dest-country").value = data.dest_country;
    populateCitySelect($("gm-dest-city"), data.dest_country, data.dest_city);
  }

  const wantSetMode = !!data.set_mode;
  const pageEl = $(`page-${ctx.kind}`);
  if (wantSetMode !== ctx.setMode) pageEl.querySelector('[data-act="toggle-set-mode"]').click();

  if (wantSetMode) {
    ctx.setBlocks.forEach((b) => b.el.remove());
    ctx.setBlocks = [];
    (data.sets || []).forEach(() => addConditionSet(ctx));
    ctx.setBlocks.forEach((b, i) => b.setData((data.sets || [])[i] || {}));
    if (!data.sets || !data.sets.length) addConditionSet(ctx);
  } else {
    buildSingleCondition(ctx);
    $(`${ctx.prefix}-condition-single-wrap`)._setData((data.sets && data.sets[0]) || {});
  }
  addLog(ctx, `'${name}' 프리셋을 불러왔습니다. (날짜는 직접 입력해주세요)`, "ok");
}

async function onDeletePreset(ctx) {
  const inputEl = $(`${ctx.prefix}-preset-input`);
  const name = inputEl.value.trim();
  if (!ctx.presets[name]) {
    addLog(ctx, "⚠ 삭제할 프리셋을 선택하세요.", "err");
    return;
  }
  if (!confirm(`'${name}' 프리셋을 삭제할까요?`)) return;
  delete ctx.presets[name];
  await window.pywebview.api.save_presets(ctx.kind, ctx.presets);
  refreshPresetList(ctx);
  inputEl.value = "";
}

// ─────────────────────────────────────────────
// 실행 / 로그 / 진행률 (페이지별)
// ─────────────────────────────────────────────
function addLog(ctx, text, kind) {
  const box = $(`${ctx.prefix}-log-box`);
  const line = document.createElement("div");
  if (kind === "ok") line.className = "log-ok";
  else if (kind === "err") line.className = "log-err";
  else if (kind === "info") line.className = "log-info";
  const ts = new Date();
  const hh = String(ts.getHours()).padStart(2, "0"), mm = String(ts.getMinutes()).padStart(2, "0"), ss = String(ts.getSeconds()).padStart(2, "0");
  line.textContent = `[${hh}:${mm}:${ss}] ${text}`;
  box.appendChild(line);
  box.scrollTop = box.scrollHeight;
}

// 파이썬 → JS 콜백 (kind로 어느 페이지인지 구분)
window.onBackendLog = function (kind, text, logKind) {
  const ctx = kind === "gmarket" ? ctxGmarket : ctxAirpremia;
  addLog(ctx, text, logKind);
};
window.onBackendProgress = function (kind, cur, total) {
  const ctx = kind === "gmarket" ? ctxGmarket : ctxAirpremia;
  const pct = total > 0 ? Math.round((cur / total) * 100) : 0;
  $(`${ctx.prefix}-progress-fill`).style.width = pct + "%";
  $(`${ctx.prefix}-progress-pct`).textContent = pct + "%";
  $(`${ctx.prefix}-progress-label`).textContent = `${cur} / ${total} 일`;
};
window.onBackendDone = function (kind) {
  const ctx = kind === "gmarket" ? ctxGmarket : ctxAirpremia;
  ctx.running = false;
  const pageEl = $(`page-${ctx.kind}`);
  pageEl.querySelector('[data-act="run"]').disabled = false;
  pageEl.querySelector('[data-act="stop"]').disabled = true;
  $(`${ctx.prefix}-progress-label`).textContent = "완료";
};

async function onRun(ctx) {
  if (ctx.running) return;
  const config = currentRouteAndCondition(ctx);
  config.date_from = $(`${ctx.prefix}-date-from`).value.trim();
  config.date_to = $(`${ctx.prefix}-date-to`).value.trim();
  config.out_dir = $(`${ctx.prefix}-out-dir`).value.trim();
  config.custom_filename = $(`${ctx.prefix}-custom-filename`).value.trim();
  config.show_browser = $(`${ctx.prefix}-show-browser`).checked;

  ctx.running = true;
  const pageEl = $(`page-${ctx.kind}`);
  pageEl.querySelector('[data-act="run"]').disabled = true;
  pageEl.querySelector('[data-act="stop"]').disabled = false;
  $(`${ctx.prefix}-progress-fill`).style.width = "0%";
  $(`${ctx.prefix}-progress-pct`).textContent = "";
  $(`${ctx.prefix}-progress-label`).textContent = "실행 중...";

  const result = ctx.kind === "gmarket"
    ? await window.pywebview.api.start_scraping_gmarket(config)
    : await window.pywebview.api.start_scraping_airpremia(config);

  if (result && result.error) {
    addLog(ctx, `❌ ${result.error}`, "err");
    ctx.running = false;
    pageEl.querySelector('[data-act="run"]').disabled = false;
    pageEl.querySelector('[data-act="stop"]').disabled = true;
  }
}

async function onStop(ctx) {
  if (ctx.kind === "gmarket") await window.pywebview.api.stop_scraping_gmarket();
  else await window.pywebview.api.stop_scraping_airpremia();
  addLog(ctx, "⏹ 중지 요청됨.", "err");
}

document.addEventListener("DOMContentLoaded", init);
