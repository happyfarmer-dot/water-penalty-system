const state = {
  allClauses: [],
  clausesById: {},
  items: [],
  suggestions: [],
  sourceFilename: "",
  modalTargetIndex: null, // null -> browse/add mode, number -> replace item at index
};

const el = (id) => document.getElementById(id);

function showToast(msg) {
  const t = el("toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(showToast._timer);
  showToast._timer = setTimeout(() => t.classList.remove("show"), 3200);
}

async function fetchClauses() {
  const resp = await fetch("/api/clauses");
  const data = await resp.json();
  state.allClauses = data.clauses || [];
  state.clausesById = {};
  for (const c of state.allClauses) {
    state.clausesById[c.id] = c;
  }
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const res = reader.result;
      const idx = res.indexOf(",");
      resolve(idx >= 0 ? res.slice(idx + 1) : res);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function handleFileUpload(file) {
  if (!file) return;
  showToast(`正在解析「${file.name}」並比對契約條文...`);
  try {
    const content_base64 = await fileToBase64(file);
    const resp = await fetch("/api/analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ filename: file.name, content_base64 }),
    });
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err.detail || "解析失敗");
    }
    const data = await resp.json();
    loadAnalysisResult(data);
    showToast("已完成條文辨識！請先核對並修改欄位後再生成罰單。");
  } catch (e) {
    showToast(`錯誤：${e.message}`);
  }
}

async function loadSample() {
  showToast("正在載入範例缺失單並比對契約條文...");
  try {
    const resp = await fetch("/api/analyze-sample", { method: "POST" });
    if (!resp.ok) throw new Error("載入範例失敗");
    const data = await resp.json();
    loadAnalysisResult(data);
    showToast("已載入範例缺失單，請確認修改欄位後再生成罰單。");
  } catch (e) {
    showToast(`錯誤：${e.message}`);
  }
}

function loadAnalysisResult(data) {
  state.sourceFilename = data.source_filename || "缺失單.xlsx";
  state.items = (data.items || []).map((it) => ({ ...it, custom_line: "" }));
  state.suggestions = data.suggestions || [];

  const meta = data.meta || {};
  const insp = meta.inspection_date || {};

  el("fProjectName").value = meta.project_name || "";
  el("fProjectCode").value = meta.project_code || "";
  el("fContractor").value = meta.contractor || "";
  el("fSafetyOfficer").value = meta.safety_officer || "";
  el("fOccurTime").value = insp.roc_datetime_default || "";
  el("fOccurLocation").value = "工地";
  el("fViolationSummary").value = "";
  el("fEvidenceNote").value = `依 ${state.sourceFilename} 督導紀錄表開立`;

  const dateTag = insp.roc_year ? `${insp.roc_year}${insp.month}${insp.day}` : "1150101";
  el("fOutputFilename").value = `(_罰單_)_違規稽查表${dateTag}.docx`;
  el("fOverrideTotal").value = "";
  el("generatedResult").classList.add("hidden");

  el("statusBar").classList.remove("hidden");
  el("workspace").classList.remove("hidden");
  el("pillFile").textContent = `來源：${state.sourceFilename}`;

  renderSuggestions();
  renderItems();
  updatePreview();
}

function renderSuggestions() {
  const box = el("suggestionsContainer");
  box.innerHTML = "";
  for (const s of state.suggestions) {
    const already = state.items.some((it) => it.clause_id === s.clause_id && it.selected);
    const div = document.createElement("div");
    div.className = "suggestion-box";
    div.innerHTML = `
      <div>
        <strong>提醒：${escapeHtml(s.title)}</strong>
        <p>${escapeHtml(s.message)}</p>
      </div>
      <button type="button" class="btn btn-outline btn-sm">${already ? "已於下方清單" : `＋ 加入 ${escapeHtml(s.clause_id)}`}</button>
    `;
    const btn = div.querySelector("button");
    if (already) btn.disabled = true;
    btn.addEventListener("click", () => {
      addManualClause(s.clause);
      renderSuggestions();
    });
    box.appendChild(div);
  }
}

function getMultiplier(repeatLevel) {
  return { 0: 1, 1: 2, 2: 3, 3: 4 }[Number(repeatLevel) || 0] || 1;
}

function computeItemSubtotal(it) {
  if (!it.selected) return 0;
  const fine = Number(it.unit_fine) || 0;
  const cnt = Math.max(1, Number(it.count) || 1);
  return fine * cnt * getMultiplier(it.repeat_level);
}

function escapeHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderItems() {
  const container = el("itemsContainer");
  container.innerHTML = "";

  el("pillMatched").textContent = `缺失／項目：${state.items.length} 筆`;
  const selectedCount = state.items.filter((it) => it.selected).length;
  el("pillSelected").textContent = `已勾選開罰：${selectedCount} 項`;

  state.items.forEach((it, idx) => {
    const card = document.createElement("div");
    card.className = `def-card${it.selected ? "" : " unchecked"}`;

    const confPct = Math.round((it.confidence || 0) * 100);
    const lowConf = confPct < 75 && !it.is_manual;
    const subtotal = computeItemSubtotal(it);

    const candHtml = (it.candidates || [])
      .slice(0, 5)
      .map((c) => {
        const active = c.clause_id === it.clause_id ? "active" : "";
        const fLabel = c.fine ? `${c.fine.toLocaleString()}元` : "依契約";
        return `<button type="button" class="chip ${active}" data-cid="${escapeHtml(c.clause_id)}" title="${escapeHtml(c.text)}">${escapeHtml(c.clause_id)} (${fLabel})</button>`;
      })
      .join("");

    const overlapHtml = it.overlap_warning
      ? `<div class="overlap-alert">⚠ ${escapeHtml(it.overlap_warning)}</div>`
      : "";

    const sourceLabel = it.is_manual
      ? `<span class="seq-tag">手動新增</span><span class="cat-tag">${escapeHtml(it.raw_category || "自行新增條文")}</span>`
      : `<span class="seq-tag">${escapeHtml(it.section)} #${escapeHtml(it.seq)}</span>
         <span class="cat-tag">${escapeHtml(it.raw_category)}</span>
         <span class="conf-tag ${lowConf ? "low" : ""}">辨識度 ${confPct}%</span>`;

    card.innerHTML = `
      <div class="def-card-top">
        <div class="def-source">
          <input type="checkbox" class="def-check" ${it.selected ? "checked" : ""} title="勾選是否列入本張罰單" />
          <div style="flex:1;">
            <div class="def-meta-line">${sourceLabel}</div>
            <p class="def-desc">${escapeHtml(it.raw_description || "（手動新增罰則項目）")}</p>
          </div>
        </div>
        <div style="display:flex; align-items:center; gap:8px;">
          <div class="def-subtotal">${subtotal.toLocaleString()} 元</div>
          <button type="button" class="btn btn-danger-text btn-sm btn-remove" title="刪除此列">✕</button>
        </div>
      </div>

      ${overlapHtml}

      <div style="display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:6px; margin-bottom:6px;">
        <span style="font-size:12px; color:var(--text-muted); font-weight:600;">建議契約條款（點擊切換）：</span>
        <button type="button" class="btn btn-outline btn-sm btn-pick-clause">從完整條文庫重選條號...</button>
      </div>
      <div class="candidate-chips">${candHtml}</div>

      <div class="edit-row-grid">
        <div class="field">
          <label>契約條號</label>
          <input type="text" class="inp-cid" value="${escapeHtml(it.clause_id)}" />
        </div>
        <div class="field">
          <label>單價（元）</label>
          <input type="number" class="inp-fine" value="${Number(it.unit_fine) || 0}" min="0" step="1000" />
        </div>
        <div class="field">
          <label>計次/處/人</label>
          <input type="number" class="inp-count" value="${Math.max(1, Number(it.count) || 1)}" min="1" step="1" />
        </div>
        <div class="field">
          <label>重複違規倍數</label>
          <select class="inp-repeat">
            <option value="0" ${it.repeat_level == 0 ? "selected" : ""}>首次違規 (1倍)</option>
            <option value="1" ${it.repeat_level == 1 ? "selected" : ""}>第1次重複 (2倍)</option>
            <option value="2" ${it.repeat_level == 2 ? "selected" : ""}>第2次重複 (3倍)</option>
            <option value="3" ${it.repeat_level == 3 ? "selected" : ""}>第3次以上 (4倍)</option>
          </select>
        </div>
      </div>

      <div class="field" style="margin-bottom:8px;">
        <label>罰單條文內容（可直接編輯文字或增刪子項）</label>
        <textarea class="inp-text" rows="2">${escapeHtml(it.clause_text)}</textarea>
      </div>

      <div class="field">
        <label>整行自訂覆寫（選填；留空則自動依「條號＋內容＋金額算式」組字）</label>
        <input type="text" class="inp-custom" value="${escapeHtml(it.custom_line || "")}" placeholder="例：2.18承攬商於上工前，未依規定辦理1.工作場所簽到4.危害因素告知5.工作場所簽退。(3,000元×3=9,000元)" />
      </div>
    `;

    // Bind events
    card.querySelector(".def-check").addEventListener("change", (e) => {
      state.items[idx].selected = e.target.checked;
      renderItems();
      updatePreview();
    });

    card.querySelector(".btn-remove").addEventListener("click", () => {
      state.items.splice(idx, 1);
      renderItems();
      renderSuggestions();
      updatePreview();
    });

    card.querySelector(".btn-pick-clause").addEventListener("click", () => {
      openClauseModal(idx);
    });

    card.querySelectorAll(".chip").forEach((chip) => {
      chip.addEventListener("click", () => {
        applyClauseToItem(idx, chip.dataset.cid);
      });
    });

    card.querySelector(".inp-cid").addEventListener("input", (e) => {
      state.items[idx].clause_id = e.target.value;
      updatePreview();
    });

    card.querySelector(".inp-fine").addEventListener("input", (e) => {
      state.items[idx].unit_fine = Math.max(0, Number(e.target.value) || 0);
      card.querySelector(".def-subtotal").textContent = `${computeItemSubtotal(state.items[idx]).toLocaleString()} 元`;
      updatePreview();
    });

    card.querySelector(".inp-count").addEventListener("input", (e) => {
      state.items[idx].count = Math.max(1, Number(e.target.value) || 1);
      card.querySelector(".def-subtotal").textContent = `${computeItemSubtotal(state.items[idx]).toLocaleString()} 元`;
      updatePreview();
    });

    card.querySelector(".inp-repeat").addEventListener("change", (e) => {
      state.items[idx].repeat_level = Number(e.target.value) || 0;
      card.querySelector(".def-subtotal").textContent = `${computeItemSubtotal(state.items[idx]).toLocaleString()} 元`;
      updatePreview();
    });

    card.querySelector(".inp-text").addEventListener("input", (e) => {
      state.items[idx].clause_text = e.target.value;
      updatePreview();
    });

    card.querySelector(".inp-custom").addEventListener("input", (e) => {
      state.items[idx].custom_line = e.target.value;
      updatePreview();
    });

    container.appendChild(card);
  });
}

function applyClauseToItem(idx, cid) {
  const cl = state.clausesById[cid];
  if (!cl) return;
  const it = state.items[idx];
  it.clause_id = cl.id;
  it.source = cl.source;
  it.clause_text = cl.text;
  it.unit_fine = cl.fine != null ? cl.fine : 0;
  it.selected = true;
  it.custom_line = "";

  // Check overlap warning from candidate or clause
  const matchedCand = (it.candidates || []).find((c) => c.clause_id === cid);
  if (matchedCand) {
    it.overlap_warning = matchedCand.overlap_warning;
  } else if (cl.overlaps && cl.overlaps.length > 0) {
    it.overlap_warning = `涉職安保命條款競合（${cl.overlaps.join("、")}），請確認是否取大值開罰`;
  } else {
    it.overlap_warning = null;
  }

  renderItems();
  updatePreview();
}

function addManualClause(clauseObj) {
  const c = clauseObj || state.allClauses[0];
  if (!c) return;
  state.items.push({
    deficiency_index: state.items.length + 1,
    section: "手動新增",
    seq: String(state.items.length + 1),
    raw_category: c.category,
    raw_description: "",
    selected: true,
    is_manual: true,
    clause_id: c.id,
    source: c.source,
    clause_text: c.text,
    unit_fine: c.fine != null ? c.fine : 0,
    count: 1,
    repeat_level: 0,
    confidence: 1.0,
    candidates: [
      {
        clause_id: c.id,
        source: c.source,
        category: c.category,
        text: c.text,
        fine: c.fine,
      },
    ],
    overlap_warning: null,
    custom_line: "",
  });
  renderItems();
  updatePreview();
}

function mergeSameClauses() {
  const merged = [];
  const byKey = new Map();

  for (const it of state.items) {
    if (!it.selected || !it.clause_id) {
      merged.push(it);
      continue;
    }
    const key = `${it.clause_id}|${it.unit_fine}|${it.repeat_level}`;
    if (byKey.has(key)) {
      const existing = byKey.get(key);
      existing.count = (Number(existing.count) || 1) + (Number(it.count) || 1);
      if (it.raw_description && !existing.raw_description.includes(it.raw_description)) {
        existing.raw_description = [existing.raw_description, it.raw_description].filter(Boolean).join("；");
      }
      existing.seq = `${existing.seq},${it.seq}`;
    } else {
      const clone = { ...it };
      byKey.set(key, clone);
      merged.push(clone);
    }
  }
  state.items = merged;
  renderItems();
  updatePreview();
  showToast("已合併相同條號與單價之項目，並加總計次。");
}

function buildTicketPayload() {
  const overrideRaw = el("fOverrideTotal").value.trim();
  return {
    project_name: el("fProjectName").value.trim(),
    project_code: el("fProjectCode").value.trim(),
    contractor: el("fContractor").value.trim(),
    safety_officer: el("fSafetyOfficer").value.trim(),
    occur_time: el("fOccurTime").value.trim(),
    occur_location: el("fOccurLocation").value.trim(),
    violation_summary: el("fViolationSummary").value.trim(),
    evidence_note: el("fEvidenceNote").value.trim(),
    output_filename: el("fOutputFilename").value.trim() || null,
    override_total: overrideRaw !== "" ? Math.max(0, Number(overrideRaw) || 0) : null,
    items: state.items.map((it) => ({
      clause_id: it.clause_id,
      clause_text: it.clause_text,
      unit_fine: Number(it.unit_fine) || 0,
      count: Math.max(1, Number(it.count) || 1),
      repeat_level: Number(it.repeat_level) || 0,
      selected: Boolean(it.selected),
      custom_line: it.custom_line ? it.custom_line.trim() : null,
    })),
  };
}

let previewTimer = null;
function updatePreview() {
  clearTimeout(previewTimer);
  previewTimer = setTimeout(async () => {
    const payload = buildTicketPayload();
    el("pvProjectName").textContent = payload.project_name;
    el("pvProjectCode").textContent = payload.project_code;
    el("pvContractor").textContent = payload.contractor;
    el("pvSafetyOfficer").textContent = payload.safety_officer;
    el("pvEvidence").textContent = payload.evidence_note ? `[佐證資料]\n${payload.evidence_note}` : "[佐證資料]";

    try {
      const resp = await fetch("/api/preview", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const pv = await resp.json();
      el("confirmNumericTotal").textContent = `${pv.total.toLocaleString()} 元`;
      el("confirmChineseTotal").textContent = `新台幣(大寫)：${pv.chinese_total}`;
      el("pvTotalWords").textContent = `新台幣(大寫)：${pv.chinese_total}`;
      el("pillSelected").textContent = `已勾選開罰：${pv.selected_count} 項`;

      const vLines = [
        `[發生時間]：${payload.occur_time}`,
        `[發生地點]：${payload.occur_location}`,
        payload.violation_summary ? `[違規情形]：${payload.violation_summary}` : "[違規情形]：",
        `[罰款金額]：${pv.total.toLocaleString()}元`,
        ...pv.lines,
      ];
      el("pvViolations").innerHTML = vLines
        .map((ln) => `<div class="ticket-v-item">${escapeHtml(ln)}</div>`)
        .join("");
    } catch (e) {
      console.error(e);
    }
  }, 60);
}

async function confirmAndGenerate() {
  const payload = buildTicketPayload();
  const selectedCount = payload.items.filter((x) => x.selected).length;
  if (selectedCount === 0) {
    showToast("請至少勾選一項違規條文再生成罰單。");
    return;
  }

  const btn = el("btnConfirmGenerate");
  btn.disabled = true;
  btn.textContent = "正在生成 Word 罰單...";

  try {
    const resp = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!resp.ok) throw new Error("生成失敗");
    const res = await resp.json();

    // 觸發瀏覽器下載
    const bin = atob(res.content_base64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = res.filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    const outBox = el("generatedResult");
    outBox.classList.remove("hidden");
    outBox.innerHTML = `已確認並產出罰單：<strong>${escapeHtml(res.filename)}</strong>（總額 ${res.total.toLocaleString()} 元／${escapeHtml(res.chinese_total)}）<br/>伺服器存檔位置：<code>${escapeHtml(res.saved_path)}</code>`;
    showToast(`罰單「${res.filename}」已生成並下載！`);
  } catch (e) {
    showToast(`生成錯誤：${e.message}`);
  } finally {
    btn.disabled = false;
    btn.textContent = "確認無誤，生成並下載罰單 (.docx)";
  }
}

// Clause Modal
function openClauseModal(targetIdx = null) {
  state.modalTargetIndex = targetIdx;
  el("modalTitle").textContent =
    targetIdx === null
      ? "契約罰則條文查詢（點擊可加入目前罰單）"
      : `選擇替代契約條文（第 ${targetIdx + 1} 列）`;
  el("clauseSearchInput").value = "";
  el("clauseSourceFilter").value = "";
  renderClauseModalList();
  el("clauseModal").classList.add("open");
  el("clauseSearchInput").focus();
}

function closeClauseModal() {
  el("clauseModal").classList.remove("open");
}

function renderClauseModalList() {
  const q = el("clauseSearchInput").value.trim().toLowerCase();
  const src = el("clauseSourceFilter").value;
  const list = el("clauseModalList");
  list.innerHTML = "";

  const filtered = state.allClauses.filter((c) => {
    if (src && c.source !== src) return false;
    if (!q) return true;
    return (
      c.id.toLowerCase().includes(q) ||
      c.category.toLowerCase().includes(q) ||
      c.text.toLowerCase().includes(q)
    );
  });

  if (filtered.length === 0) {
    list.innerHTML = `<div style="padding:20px; text-align:center; color:var(--text-muted);">查無符合關鍵字之契約條文</div>`;
    return;
  }

  for (const c of filtered.slice(0, 80)) {
    const div = document.createElement("div");
    div.className = "clause-option";
    const fineLabel = c.fine != null ? `${c.fine.toLocaleString()} 元` : c.fine_text;
    div.innerHTML = `
      <div style="flex:1;">
        <div style="display:flex; gap:6px; align-items:center; margin-bottom:4px;">
          <span class="seq-tag">${escapeHtml(c.id)}</span>
          <span class="cat-tag">${escapeHtml(c.category)}</span>
        </div>
        <div style="font-size:13.5px;">${escapeHtml(c.text)}</div>
      </div>
      <div style="text-align:right; font-weight:700; color:var(--primary); min-width:92px;">
        ${escapeHtml(fineLabel)}
      </div>
    `;
    div.addEventListener("click", () => {
      if (state.modalTargetIndex !== null) {
        applyClauseToItem(state.modalTargetIndex, c.id);
      } else {
        if (el("workspace").classList.contains("hidden")) {
          showToast("請先匯入缺失單或載入範例後再加入條文。");
          return;
        }
        addManualClause(c);
        showToast(`已新增條款 ${c.id}`);
      }
      closeClauseModal();
    });
    list.appendChild(div);
  }
}

// Init
document.addEventListener("DOMContentLoaded", async () => {
  await fetchClauses();

  const dropZone = el("dropZone");
  const fileInput = el("fileInput");

  el("btnPickFile").addEventListener("click", (e) => {
    e.stopPropagation();
    fileInput.click();
  });
  dropZone.addEventListener("click", (e) => {
    if (e.target.closest("button")) return;
    fileInput.click();
  });
  fileInput.addEventListener("change", (e) => {
    if (e.target.files[0]) handleFileUpload(e.target.files[0]);
  });

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });
  dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    if (e.dataTransfer.files[0]) handleFileUpload(e.dataTransfer.files[0]);
  });

  el("btnLoadSample").addEventListener("click", (e) => {
    e.stopPropagation();
    loadSample();
  });

  ["fProjectName", "fProjectCode", "fContractor", "fSafetyOfficer", "fOccurTime", "fOccurLocation", "fViolationSummary", "fEvidenceNote", "fOverrideTotal"].forEach((id) => {
    el(id).addEventListener("input", updatePreview);
  });

  el("btnMergeSame").addEventListener("click", mergeSameClauses);
  el("btnAddManualItem").addEventListener("click", () => openClauseModal(null));
  el("btnBrowseClauses").addEventListener("click", () => openClauseModal(null));
  el("btnCloseModal").addEventListener("click", closeClauseModal);
  el("clauseModal").addEventListener("click", (e) => {
    if (e.target === el("clauseModal")) closeClauseModal();
  });
  el("clauseSearchInput").addEventListener("input", renderClauseModalList);
  el("clauseSourceFilter").addEventListener("change", renderClauseModalList);
  el("btnConfirmGenerate").addEventListener("click", confirmAndGenerate);
});
