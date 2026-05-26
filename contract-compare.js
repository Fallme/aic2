const API_BASE = window.location.origin.startsWith("http") ? window.location.origin : "http://localhost:6189";

const els = {
  serverState: document.getElementById("server-state"),
  leftFile: document.getElementById("left-file"),
  rightFile: document.getElementById("right-file"),
  leftEditor: document.getElementById("left-editor"),
  rightEditor: document.getElementById("right-editor"),
  leftText: document.getElementById("left-text"),
  rightText: document.getElementById("right-text"),
  leftFileName: document.getElementById("left-file-name"),
  rightFileName: document.getElementById("right-file-name"),
  leftHint: document.getElementById("left-hint"),
  rightHint: document.getElementById("right-hint"),
  compareBtn: document.getElementById("compare-btn"),
  drawerToggleBtn: document.getElementById("compare-drawer-toggle-btn"),
  drawer: document.getElementById("compare-drawer"),
  drawerOverlay: document.getElementById("compare-drawer-overlay"),
  drawerCloseBtn: document.getElementById("compare-drawer-close-btn"),
  drawerCountBadge: document.getElementById("drawer-count-badge"),
  resultPanel: document.getElementById("compare-drawer"),
  resetBtn: document.getElementById("reset-compare-btn"),
  downloadReportBtn: document.getElementById("download-report-btn"),
  summary: document.getElementById("compare-summary"),
  partyCount: document.getElementById("party-count"),
  missingCount: document.getElementById("missing-count"),
  addedCount: document.getElementById("added-count"),
  inconsistentCount: document.getElementById("inconsistent-count"),
  bubbleLayer: document.getElementById("left-bubble-layer"),
  bubbleList: document.getElementById("bubble-annotation-list"),
  bubbleEmpty: document.getElementById("bubble-empty"),
  popover: document.getElementById("compare-popover"),
};

const state = {
  leftName: "",
  rightName: "",
  leftHtml: "",
  rightHtml: "",
  diff: [],
  stats: { party: 0, missing: 0, added: 0, inconsistent: 0 },
};

const editors = {
  left: new OfficeEditor(els.leftEditor, {
    placeholder: "粘贴或导入原合同，系统会尽量保留 Word 结构。",
    defaultZoom: 90,
    readOnly: true,
  }),
  right: new OfficeEditor(els.rightEditor, {
    placeholder: "粘贴或导入修订后合同。",
    defaultZoom: 90,
    readOnly: true,
  }),
};

function escapeHtml(value = "") {
  return String(value).replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[char]));
}

function setEditorContent(side, text = "", html = "") {
  const editor = editors[side];
  const textarea = side === "left" ? els.leftText : els.rightText;
  const finalText = String(text || "");
  const finalHtml = html || editor.textToContractHtml(finalText);
  textarea.value = finalText;
  state[side + "Html"] = finalHtml;
  editor.setHtml(finalHtml);
  clearAnnotations();
}

function currentText(side) {
  const clone = editors[side].editor.cloneNode(true);
  // Unwrap diff-highlight spans (preserve text content)
  clone.querySelectorAll(".diff-highlight").forEach((el) => {
    const parent = el.parentNode;
    if (parent) { while (el.firstChild) parent.insertBefore(el.firstChild, el); parent.removeChild(el); }
  });
  // Remove label badges and other overlays (text not needed)
  clone.querySelectorAll(".diff-label-badge-inline,.compare-comment-badge,.compare-annotated,.diff-annotation,.bubble-marker").forEach((node) => node.remove());
  return clone.innerText.trim() || (side === "left" ? els.leftText.value : els.rightText.value);
}

function currentHtml(side) {
  return state[side + "Html"] || editors[side].textToContractHtml(currentText(side));
}

function normalizeLine(value = "") {
  return String(value).replace(/\s+/g, " ").trim();
}

async function api(path, options = {}) {
  const res = await fetch(API_BASE + path, options);
  const type = res.headers.get("content-type") || "";
  const data = type.includes("application/json") ? await res.json() : await res.text();
  if (!res.ok) throw new Error(data?.error || data || "请求失败");
  return data;
}

async function checkHealth() {
  try {
    const data = await api("/api/health");
    const providerLabel = data.modelProvider && data.modelProvider !== "None" ? data.modelProvider + " / " : "";
    els.serverState.classList.add("ok");
    els.serverState.classList.remove("warn");
    els.serverState.textContent = data.hasApiKey ? "AI 已配置：" + providerLabel + (data.model || "") + "，实时校验" : "AI 未配置，本地规则可用";
  } catch (error) {
    els.serverState.classList.remove("ok");
    els.serverState.classList.add("warn");
    els.serverState.textContent = "服务未连接";
  }
}

async function parseUploadedFile(file, side) {
  const textEl = side === "left" ? els.leftText : els.rightText;
  const nameEl = side === "left" ? els.leftFileName : els.rightFileName;
  const hintEl = side === "left" ? els.leftHint : els.rightHint;
  if (!file) return;
  hintEl.textContent = "正在解析文件...";
  const formData = new FormData();
  formData.append("contractFile", file);
  try {
    const data = await api("/api/contracts/parse", { method: "POST", body: formData });
    setEditorContent(side, data.contractText || "", data.contractHtml || "");
    nameEl.textContent = data.fileName || file.name;
    state[side + "Name"] = data.fileName || file.name;
    hintEl.textContent = data.templatePreserved ? "DOCX 格式已保留。" : "已解析 " + (data.charCount || textEl.value.length) + " 字。";
  } catch (error) {
    hintEl.textContent = error.message;
  }
}

/* ===== 语义分析 ===== */

function normalizeLines(text = "") {
  return String(text)
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
}

function semanticNormalize(text = "") {
  return String(text)
    .replace(/\s+/g, "")
    .replace(/[，。；：、""''（）()【】《》<>「」『』,.!?;:"'[\]{}\-—_]/g, "")
    .replace(/甲方采购方|甲方（采购方）/g, "甲方")
    .replace(/乙方供应商|乙方（供应商）/g, "乙方")
    .trim();
}

function splitIntoSegments(text = "") {
  const lines = normalizeLines(text);
  const segments = [];
  lines.forEach((line, lineIndex) => {
    const rawParts = line
      .split(/(?<=[。；;！!?？])|(?=第[一二三四五六七八九十百\d]+[章节条、])|(?=[（(]?\d{1,3}[）).、])|(?=甲方[:：])|(?=乙方[:：])|(?=丙方[:：])|(?=采购方[:：])|(?=供应商[:：])|(?=名称[:：])|(?=法定代表人?[:：])|(?=统一社会信用代码[:：])|(?=注册地址[:：])|(?=地址[:：])|(?=联系人[:：])|(?=联系电话[:：])|(?=开户行[:：])|(?=银行账号[:：])|(?=账号[:：])|(?=合同金额[:：])|(?=付款方式[:：])|(?=交付[:：])|(?=验收[:：])/)
      .map((part) => part.trim())
      .filter(Boolean);
    const parts = rawParts.length ? rawParts : [line];
    parts.forEach((part, partIndex) => {
      const norm = semanticNormalize(part);
      if (norm.length < 3) return;
      segments.push({ text: part, norm, lineIndex, partIndex, index: segments.length });
    });
  });
  return segments;
}

function bigrams(text = "") {
  const normalized = semanticNormalize(text);
  if (normalized.length <= 2) return normalized ? [normalized] : [];
  const result = [];
  for (let index = 0; index < normalized.length - 1; index += 1) {
    result.push(normalized.slice(index, index + 2));
  }
  return result;
}

function segmentSimilarity(left = "", right = "") {
  const a = semanticNormalize(left);
  const b = semanticNormalize(right);
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (a.includes(b) || b.includes(a)) return Math.min(a.length, b.length) / Math.max(a.length, b.length);
  const leftBigrams = bigrams(a);
  const rightBigrams = bigrams(b);
  if (!leftBigrams.length || !rightBigrams.length) return 0;
  const counts = new Map();
  leftBigrams.forEach((token) => counts.set(token, (counts.get(token) || 0) + 1));
  let overlap = 0;
  rightBigrams.forEach((token) => {
    const count = counts.get(token) || 0;
    if (!count) return;
    overlap += 1;
    counts.set(token, count - 1);
  });
  return (2 * overlap) / (leftBigrams.length + rightBigrams.length);
}

function findBestSegment(segment, candidates, usedNew) {
  let best = null;
  candidates.forEach((candidate) => {
    if (usedNew.has(candidate.index)) return;
    const score = segmentSimilarity(segment.norm, candidate.norm);
    const distancePenalty = Math.min(Math.abs(segment.index - candidate.index) * 0.002, 0.08);
    const adjusted = score - distancePenalty;
    if (!best || adjusted > best.adjusted) best = { candidate, score, adjusted };
  });
  return best;
}

function classifyDiff(item) {
  const text = (item.text || "") + " " + (item.before || "") + " " + (item.after || "");
  if (/甲方|乙方|丙方|采购方|供应商|委托方|受托方|主体|公司|统一社会信用代码|法定代表人|联系人|地址|账户|开户行/.test(text)) {
    return "party";
  }
  if (item.type === "removed") return "missing";
  if (item.type === "added") return "added";
  return "inconsistent";
}

function categoryLabel(category) {
  return {
    party: "主体不同",
    missing: "原文有、新合同缺失",
    added: "新合同增加",
    inconsistent: "描述不一致",
  }[category] || "描述不一致";
}

function categoryIcon(category) {
  return {
    party: "🏢",
    missing: "➖",
    added: "➕",
    inconsistent: "✏️",
  }[category] || "📌";
}

function categoryClass(category) {
  return {
    party: "bubble-party",
    missing: "bubble-missing",
    added: "bubble-added",
    inconsistent: "bubble-inconsistent",
  }[category] || "bubble-inconsistent";
}

function diffReason(item) {
  const category = classifyDiff(item);
  if (category === "party") return "合同主体或主体信息不一致，需核对签约主体、统一社会信用代码、地址、账户等信息。";
  if (category === "missing") return "新合同未保留原合同中的该项内容，需确认是否误删或是否应继续保留。";
  if (category === "added") return "新合同新增了该项内容，需确认是否经过双方确认并符合业务事实。";
  return "两份合同同一位置或相近条款表述不一致，需确认以哪一版为准。";
}

function semanticDiff(oldText, newText) {
  const oldSegments = splitIntoSegments(oldText);
  const newSegments = splitIntoSegments(newText);
  const usedNew = new Set();
  const diff = [];

  oldSegments.forEach((oldSegment) => {
    const best = findBestSegment(oldSegment, newSegments, usedNew);
    if (!best || best.score < 0.34) {
      diff.push({ type: "removed", text: oldSegment.text, oldIndex: oldSegment.index });
      return;
    }
    usedNew.add(best.candidate.index);
    if (best.score >= 0.94) {
      diff.push({ type: "equal", text: oldSegment.text, oldIndex: oldSegment.index, newIndex: best.candidate.index });
      return;
    }
    diff.push({
      type: "changed",
      before: oldSegment.text,
      after: best.candidate.text,
      score: best.score,
      oldIndex: oldSegment.index,
      newIndex: best.candidate.index,
    });
  });

  newSegments.forEach((newSegment) => {
    if (!usedNew.has(newSegment.index)) diff.push({ type: "added", text: newSegment.text, newIndex: newSegment.index });
  });
  return diff.sort((a, b) => (a.oldIndex ?? a.newIndex ?? 0) - (b.oldIndex ?? b.newIndex ?? 0));
}

/* ===== 对比主流程 ===== */

function syncEditorsFromTextareas() {
  // If editor DOM is empty but textarea has content, sync it so bubble positioning works
  const leftEditorText = (els.leftEditor.innerText || "").trim();
  const rightEditorText = (els.rightEditor.innerText || "").trim();
  const leftRaw = (els.leftText.value || "").trim();
  const rightRaw = (els.rightText.value || "").trim();

  if (!leftEditorText && leftRaw) {
    setEditorContent("left", leftRaw);
  }
  if (!rightEditorText && rightRaw) {
    setEditorContent("right", rightRaw);
  }
}

function compareContracts() {
  // Sync textarea content into editor so bubble positioning can find text nodes
  syncEditorsFromTextareas();

  const oldText = currentText("left");
  const newText = currentText("right");
  if (!normalizeLines(oldText).length || !normalizeLines(newText).length) {
    els.summary.textContent = "请先导入或粘贴两份合同文本。";
    return;
  }
  state.diff = semanticDiff(oldText, newText);
  state.stats = state.diff.filter((item) => item.type !== "equal").reduce((stats, item) => {
    const category = classifyDiff(item);
    stats[category] += 1;
    return stats;
  }, { party: 0, missing: 0, added: 0, inconsistent: 0 });

  renderResultSummary();
  renderBubbleAnnotations();
  showCompareResult();
  // Delay annotation to let editors finish rendering HTML into contentEditable
  // Retry up to 5 times until annotations are actually applied
  _retryAnnotate(0);
}

function _retryAnnotate(attempt) {
  if (attempt > 5) return;
  const delay = 300 + attempt * 400;
  setTimeout(() => {
    annotateEditors();
    // Check if any highlights were actually created
    const leftHighlights = els.leftEditor.querySelectorAll(".diff-highlight");
    const rightHighlights = els.rightEditor.querySelectorAll(".diff-highlight");
    if (leftHighlights.length === 0 && rightHighlights.length === 0 && attempt < 5) {
      _retryAnnotate(attempt + 1);
    }
  }, delay);
}

function showCompareResult() {
  openDrawer();
}

function renderResultSummary() {
  const visible = state.diff.filter((item) => item.type !== "equal");
  els.partyCount.textContent = state.stats.party;
  els.missingCount.textContent = state.stats.missing;
  els.addedCount.textContent = state.stats.added;
  els.inconsistentCount.textContent = state.stats.inconsistent;
  els.downloadReportBtn.disabled = !visible.length;
  els.drawerToggleBtn.hidden = !visible.length;
  if (els.drawerCountBadge) els.drawerCountBadge.textContent = visible.length;
  els.summary.textContent = visible.length
    ? "发现 " + visible.length + " 处差异：主体不同 " + state.stats.party + " 处，缺失 " + state.stats.missing + " 处，新增 " + state.stats.added + " 处，描述不一致 " + state.stats.inconsistent + " 处。"
    : "两份合同正文未发现明显差异。";
}

/* ===== 气泡批注 — 不修改编辑器内容，在叠加层上渲染 ===== */

function findTextInEditor(searchText) {
  const editor = els.leftEditor;
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      if (!node.nodeValue || !normalizeLine(node.nodeValue)) return NodeFilter.FILTER_REJECT;
      if (node.parentElement?.closest(".compare-comment-badge,.compare-annotated,.bubble-marker")) return NodeFilter.FILTER_REJECT;
      const parent = node.parentElement;
      if (parent?.closest(".oe-toolbar,.oe-statusbar")) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    },
  });

  const norm = normalizeLine(searchText).slice(0, 60);
  if (!norm) return null;

  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);

  for (const node of nodes) {
    const value = normalizeLine(node.nodeValue || "");
    const shortNorm = norm.slice(0, 30);
    if (value.includes(shortNorm) || value.includes(norm.slice(0, 40))) {
      // Find the position
      const range = document.createRange();
      const idx = (node.nodeValue || "").indexOf(shortNorm);
      if (idx < 0) continue;
      range.setStart(node, idx);
      range.setEnd(node, idx + Math.min(shortNorm.length, (node.nodeValue || "").length - idx));
      const rect = range.getBoundingClientRect();
      const editorRect = editor.getBoundingClientRect();
      return {
        top: rect.top - editorRect.top - 2,
        bottom: rect.bottom - editorRect.top + 2,
        left: editorRect.width,
        width: rect.width,
        node,
        range,
      };
    }
  }
  return null;
}

function findAllTextPositions(searchText) {
  const editor = els.leftEditor;
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      if (!node.nodeValue || !normalizeLine(node.nodeValue)) return NodeFilter.FILTER_REJECT;
      if (node.parentElement?.closest(".oe-toolbar,.oe-statusbar")) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    },
  });

  const norm = normalizeLine(searchText).slice(0, 60);
  if (!norm) return [];

  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);

  const positions = [];
  for (const node of nodes) {
    const value = node.nodeValue || "";
    const shortNorm = norm.slice(0, 30);
    let idx = 0;
    while (idx < value.length) {
      const pos = value.indexOf(shortNorm, idx);
      if (pos < 0) break;

      const range = document.createRange();
      const end = pos + Math.min(shortNorm.length, value.length - pos);
      try {
        range.setStart(node, pos);
        range.setEnd(node, end);
      } catch (e) { idx = pos + 1; continue; }

      const rect = range.getBoundingClientRect();
      const editorRect = editor.getBoundingClientRect();
      positions.push({
        top: rect.top - editorRect.top,
        left: rect.right - editorRect.left,
        height: rect.height,
      });
      idx = pos + shortNorm.length;
    }
  }
  return positions;
}

function renderBubbleAnnotations() {
  if (els.bubbleLayer) els.bubbleLayer.innerHTML = "";
  if (els.bubbleList) els.bubbleList.innerHTML = "";
  if (els.downloadReportBtn) els.downloadReportBtn.disabled = true;

  const visible = state.diff.filter((item) => item.type !== "equal");
  if (!visible.length) {
    if (els.bubbleEmpty) els.bubbleEmpty.style.display = "";
    return;
  }
  if (els.bubbleEmpty) els.bubbleEmpty.style.display = "none";

  // Render annotation list in drawer
  const listHtml = visible.map((item, index) => {
    const category = classifyDiff(item);
    const icon = categoryIcon(category);
    const label = categoryLabel(category);
    const cls = categoryClass(category);
    const reason = diffReason(item);

    let changeDetail = "";
    if (item.type === "changed") {
      changeDetail = `<div class="bubble-change-detail">
        <div class="bcd-row bcd-before"><span>原文</span>${escapeHtml(item.before)}</div>
        <div class="bcd-row bcd-after"><span>新文</span>${escapeHtml(item.after)}</div>
      </div>`;
    } else if (item.type === "removed") {
      changeDetail = `<div class="bubble-change-detail">
        <div class="bcd-row bcd-before"><span>原文（已删除）</span>${escapeHtml(item.text)}</div>
      </div>`;
    } else if (item.type === "added") {
      changeDetail = `<div class="bubble-change-detail">
        <div class="bcd-row bcd-after"><span>新文（新增）</span>${escapeHtml(item.text)}</div>
      </div>`;
    }

    return `<div class="bubble-annotation-item ${cls}" data-diff-index="${index}" id="bubble-item-${index}">
      <div class="bai-header">
        <span class="bai-number">${icon} 差异 ${index + 1}</span>
        <span class="bai-category ${cls}">${label}</span>
        <button class="bai-locate-btn secondary" type="button" data-locate="${index}" title="定位到原文">📍 定位</button>
      </div>
      <p class="bai-reason">${escapeHtml(reason)}</p>
      ${changeDetail}
    </div>`;
  }).join("");

  els.bubbleList.innerHTML = listHtml;

  // Position bubble dots on the left editor overlay
  if (!els.bubbleLayer) return;

  setTimeout(() => {
    const placement = placeBubbleDots(visible);
    els.bubbleLayer.innerHTML = placement;
  }, 200);
}

/* ===== 抽屉面板 ===== */

function openDrawer() {
  if (!els.drawer) return;
  els.drawer.classList.add("open");
  if (els.drawerOverlay) els.drawerOverlay.classList.add("visible");
}

function closeDrawer() {
  if (!els.drawer) return;
  els.drawer.classList.remove("open");
  if (els.drawerOverlay) els.drawerOverlay.classList.remove("visible");
  hidePopover();
}

function toggleDrawer() {
  if (!els.drawer) return;
  els.drawer.classList.contains("open") ? closeDrawer() : openDrawer();
}

/* ===== 批注弹出卡片 ===== */

function showPopover(event, index) {
  const item = state.diff.filter((d) => d.type !== "equal")[index];
  if (!item || !els.popover) return;

  const category = classifyDiff(item);
  const icon = categoryIcon(category);
  const label = categoryLabel(category);
  const cls = categoryClass(category);
  const reason = diffReason(item);

  let changeDetail = "";
  if (item.type === "changed") {
    changeDetail = `<div class="bubble-change-detail">
      <div class="bcd-row bcd-before"><span>原文</span>${escapeHtml(item.before)}</div>
      <div class="bcd-row bcd-after"><span>新文</span>${escapeHtml(item.after)}</div>
    </div>`;
  } else if (item.type === "removed") {
    changeDetail = `<div class="bubble-change-detail">
      <div class="bcd-row bcd-before"><span>原文（已删除）</span>${escapeHtml(item.text)}</div>
    </div>`;
  } else if (item.type === "added") {
    changeDetail = `<div class="bubble-change-detail">
      <div class="bcd-row bcd-after"><span>新文（新增）</span>${escapeHtml(item.text)}</div>
    </div>`;
  }

  els.popover.innerHTML = `<div class="popover-inner ${cls}">
    <div class="popover-head">
      <span>${icon} 差异 ${index + 1}</span>
      <span class="bai-category ${cls}">${label}</span>
      <button class="popover-close-btn" onclick="document.getElementById('compare-popover').hidden=true">✕</button>
    </div>
    <p class="popover-reason">${escapeHtml(reason)}</p>
    ${changeDetail}
    <div class="popover-actions">
      <button class="secondary" type="button" data-locate="${index}">📍 定位到原文</button>
      <button class="secondary" type="button" data-view-in-drawer="${index}">📋 在批注列表中查看</button>
    </div>
  </div>`;

  // Position popover near the click (viewport-relative since popover is position:fixed)
  const clientX = event.clientX || (event.touches && event.touches[0].clientX) || 200;
  const clientY = event.clientY || (event.touches && event.touches[0].clientY) || 200;

  els.popover.style.left = Math.min(clientX, window.innerWidth - 360) + "px";
  els.popover.style.top = Math.max(10, clientY - 40) + "px";
  els.popover.hidden = false;
}

function hidePopover() {
  if (els.popover) els.popover.hidden = true;
}

function placeBubbleDots(diffItems) {
  if (!els.leftEditor) return "";
  const editorRect = els.leftEditor.getBoundingClientRect();
  const editorHeight = editorRect.height;

  const dotsHtml = [];

  diffItems.forEach((item, index) => {
    const category = classifyDiff(item);
    const searchText = item.type === "changed" ? item.before : item.type === "removed" ? item.text : item.text;
    const positions = findAllTextPositions(searchText);

    // Use first matching position
    const pos = positions[0];
    if (!pos) return;

    // Calculate dot position relative to editor
    const dotTop = pos.top + pos.height / 2;
    const dotLeft = Math.min(pos.left + 12, editorRect.width - 20);

    const cls = categoryClass(category);
    dotsHtml.push(
      `<div class="bubble-dot ${cls}" style="top:${dotTop}px; left:${dotLeft}px;" data-diff-index="${index}" title="差异 ${index + 1}：${categoryLabel(category)}">
        <span class="bubble-dot-inner">${index + 1}</span>
      </div>`
    );
  });

  return dotsHtml.join("");
}

/* ===== 定位到原文 ===== */

function locateDiffInEditor(index) {
  const item = state.diff.filter((d) => d.type !== "equal")[index];
  if (!item) return;

  const searchText = item.type === "changed" ? item.before : item.type === "removed" ? item.text : item.text;
  const position = findTextInEditor(searchText);
  if (!position) return;

  const editor = els.leftEditor;

  // Scroll to position
  editor.scrollTo({
    top: Math.max(0, editor.scrollTop + position.top - 80),
    behavior: "smooth",
  });

  // Highlight animation
  highlightPosition(position);
}

function highlightPosition(position) {
  if (!position.node) return;

  // Remove existing highlights
  els.leftEditor.querySelectorAll(".bubble-highlight-temp").forEach((el) => el.classList.remove("bubble-highlight-temp"));

  // Add temp highlight
  const parent = position.node.parentElement;
  if (parent) {
    const span = document.createElement("span");
    span.className = "bubble-highlight-temp";
    span.textContent = position.node.nodeValue?.slice(
      position.range.startOffset,
      Math.min(position.range.endOffset, position.node.nodeValue.length)
    ) || "";
    try {
      position.range.deleteContents();
      position.range.insertNode(span);
    } catch (e) {
      // Fallback: highlight the parent
      parent.classList.add("bubble-highlight-temp-block");
      setTimeout(() => parent.classList.remove("bubble-highlight-temp-block"), 3000);
    }
    setTimeout(() => {
      if (span.parentNode) {
        span.replaceWith(document.createTextNode(span.textContent || ""));
      }
    }, 3000);
  }
}

/* ===== 内联对比批注 ===== */

function clearAnnotations() {
  // Remove highlight spans from both editors
  document.querySelectorAll(".diff-highlight,.diff-label-badge-inline").forEach((el) => {
    const parent = el.parentNode;
    if (parent) {
      while (el.firstChild) parent.insertBefore(el.firstChild, el);
      parent.removeChild(el);
      parent.normalize();
    }
  });
  // Remove old-style annotation spans (legacy cleanup)
  document.querySelectorAll(".diff-annotation").forEach((el) => {
    const parent = el.parentNode;
    if (parent) {
      while (el.firstChild) parent.insertBefore(el.firstChild, el);
      parent.removeChild(el);
      parent.normalize();
    }
  });
}

/* Find and highlight matching text in editor, optionally adding a badge.
   Uses TreeWalker to find the best matching text node, then wraps the matched text
   in a highlight span with an optional numbered badge. */
function _highlightInEditor(editor, raw, seq, label, addBadge, category) {
  if (!editor) return;
  const rawTrimmed = String(raw || "").trim();
  if (!rawTrim || rawTrim.length < 4) return;

  // Split into lines for multi-line matching
  const lines = rawTrim.split(/\n/).filter((l) => l.trim().length >= 4);

  // First pass: try to find a single text node containing the full first line
  const firstLine = (lines[0] || "").trim();
  if (firstLine.length >= 6) {
    const found = _tryHighlightSingleLine(editor, firstLine, seq, label, addBadge, category);
    if (found) return;
  }

  // Second pass: try any line
  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed.length < 6) continue;
    const found = _tryHighlightSingleLine(editor, trimmed, seq, label, addBadge, category);
    if (found) return;
  }
}

function _tryHighlightSingleLine(editor, searchText, seq, label, addBadge, category) {
  const shortSearch = searchText.length > 40 ? searchText.slice(0, 40) : searchText;

  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      if (!node.nodeValue || !node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
      if (node.parentElement?.closest(
        ".diff-highlight,.diff-label-badge-inline,.compare-comment-badge,.oe-toolbar,.oe-statusbar"
      )) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    },
  });

  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);

  // Try to find the search text in a single text node
  for (const node of nodes) {
    const value = node.nodeValue || "";
    const idx = value.indexOf(shortSearch);
    if (idx < 0) continue;

    try {
      const range = document.createRange();
      const endIdx = Math.min(idx + shortSearch.length, value.length);
      range.setStart(node, idx);
      range.setEnd(node, endIdx);
      if (range.collapsed) continue;

      const matchedText = range.toString();
      if (!matchedText || matchedText.trim().length < 4) continue;

      // Wrap matched text in highlight span (non-destructive for contentEditable)
      const span = document.createElement("span");
      span.className = "diff-highlight" + (category ? " diff-" + category : "");
      span.textContent = matchedText;

      // Remove old content in range and insert highlight
      range.deleteContents();
      range.insertNode(span);

      // Add badge before the highlight
      if (addBadge) {
        const badge = document.createElement("span");
        badge.className = "diff-label-badge-inline";
        badge.textContent = seq + " " + label;
        if (span.parentNode) {
          span.parentNode.insertBefore(badge, span);
        }
      }
      return true;
    } catch (e) { /* skip node, try next */ }
  }

  return false;
}

function annotateEditors() {
  clearAnnotations();
  if (!state.diff || !state.diff.length) return;
  const visible = state.diff.filter((d) => d.type !== "equal");
  if (!visible.length) return;
  let seq = 0;
  visible.forEach((item) => {
    seq++;
    const category = classifyDiff(item);
    const categoryLabelMap = { party: "主体", missing: "缺失", added: "新增", inconsistent: "修改" };
    const catLabel = categoryLabelMap[category] || "修改";

    // Left editor: removed + changed + party — show badge
    if (item.type === "removed" || item.type === "changed" || category === "party") {
      const searchText = item.type === "changed" ? item.before : item.text;
      _highlightInEditor(els.leftEditor, searchText, seq, catLabel, true, category);
    }
    // Right editor: added + changed + party — also show badge for cross-reference
    if (item.type === "added" || item.type === "changed" || category === "party") {
      const searchText = item.type === "changed" ? item.after : item.text;
      _highlightInEditor(els.rightEditor, searchText, seq, catLabel, true, category);
    }
  });
}

/* ===== 事件绑定 ===== */

// Drawer toggle
if (els.drawerToggleBtn) {
  els.drawerToggleBtn.addEventListener("click", toggleDrawer);
}
if (els.drawerCloseBtn) {
  els.drawerCloseBtn.addEventListener("click", closeDrawer);
}
if (els.drawerOverlay) {
  els.drawerOverlay.addEventListener("click", closeDrawer);
}

// Locate buttons in drawer
if (els.bubbleList) {
  els.bubbleList.addEventListener("click", (e) => {
    const locateBtn = e.target.closest("[data-locate]");
    if (!locateBtn) return;
    const index = Number(locateBtn.dataset.locate);
    if (!Number.isNaN(index)) {
      locateDiffInEditor(index);
      // Flash the item in drawer
      const listItem = document.getElementById("bubble-item-" + index);
      if (listItem) {
        listItem.classList.add("bai-flash");
        setTimeout(() => listItem.classList.remove("bai-flash"), 2000);
      }
    }
  });
}

// Popover action buttons
if (els.popover) {
  els.popover.addEventListener("click", (e) => {
    const locateBtn = e.target.closest("[data-locate]");
    const viewBtn = e.target.closest("[data-view-in-drawer]");
    const closeBtn = e.target.closest(".popover-close-btn");

    if (closeBtn) { hidePopover(); return; }

    const index = Number((locateBtn || viewBtn)?.dataset?.locate || (locateBtn || viewBtn)?.dataset?.viewInDrawer);
    if (Number.isNaN(index)) return;

    if (locateBtn) {
      locateDiffInEditor(index);
    }
    if (viewBtn) {
      openDrawer();
      const listItem = document.getElementById("bubble-item-" + index);
      if (listItem) {
        listItem.scrollIntoView({ behavior: "smooth", block: "center" });
        listItem.classList.add("bai-flash");
        setTimeout(() => listItem.classList.remove("bai-flash"), 2000);
      }
    }
  });
}

// Click outside popover to close (delay to avoid same-click closing)
document.addEventListener("click", (e) => {
  if (els.popover && !els.popover.hidden) {
    // Skip if clicking the bubble dot that just opened this popover
    if (e.target.closest(".bubble-dot")) return;
    if (!els.popover.contains(e.target)) {
      hidePopover();
    }
  }
});

// Click bubble dot → show popover + locate
if (els.bubbleLayer) {
  els.bubbleLayer.addEventListener("click", (e) => {
    const dot = e.target.closest(".bubble-dot");
    if (!dot) return;
    e.stopPropagation();
    const index = Number(dot.dataset.diffIndex);
    if (!Number.isNaN(index)) {
      locateDiffInEditor(index);
      showPopover(e, index);
    }
  });
}

/* ===== 下载报告 ===== */

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function downloadReport() {
  const rows = state.diff
    .filter((item) => item.type !== "equal")
    .map((item, index) => {
      const type = categoryLabel(classifyDiff(item));
      const before = item.type === "changed" ? item.before : item.type === "removed" ? item.text : "";
      const after = item.type === "changed" ? item.after : item.type === "added" ? item.text : "";
      return "<tr><td>" + (index + 1) + "</td><td>" + type + "</td><td>" + escapeHtml(before) + "</td><td>" + escapeHtml(after) + "</td></tr>";
    })
    .join("");
  const html = "<!DOCTYPE html><html><head><meta charset=\"utf-8\"><title>合同对比报告</title>" +
    "<style>body{font-family:\"Microsoft YaHei\",sans-serif;line-height:1.7;padding:32px;max-width:900px;margin:0 auto;color:#1d1d1f}" +
    "h1{font-size:24px;border-bottom:2px solid #007AFF;padding-bottom:12px}" +
    "table{border-collapse:collapse;width:100%;margin-top:16px}" +
    "td,th{border:1px solid #d2d2d7;padding:10px 12px;vertical-align:top;font-size:14px}" +
    "th{background:rgba(0,122,255,0.06);font-weight:700;text-align:left}" +
    "tr:hover{background:rgba(0,122,255,0.03)}" +
    ".report-meta{color:#86868b;font-size:14px;margin-bottom:12px}" +
    "</style></head><body>" +
    "<h1>合同对比报告</h1>" +
    "<p class=\"report-meta\">原合同：" + escapeHtml(state.leftName || "文本输入") + " · 新合同：" + escapeHtml(state.rightName || "文本输入") + "</p>" +
    "<p class=\"report-meta\">生成时间：" + new Date().toLocaleString("zh-CN") + "</p>" +
    "<p class=\"report-meta\">主体不同 " + state.stats.party + " 处 · 缺失 " + state.stats.missing + " 处 · 新增 " + state.stats.added + " 处 · 描述不一致 " + state.stats.inconsistent + " 处</p>" +
    "<table><thead><tr><th>序号</th><th>类型</th><th>原文</th><th>新文</th></tr></thead><tbody>" + rows + "</tbody></table>" +
    "</body></html>";
  downloadBlob(new Blob([html], { type: "application/msword;charset=utf-8" }), "合同对比报告_" + new Date().toISOString().slice(0, 10) + ".doc");
}

/* ===== 清空 ===== */

function resetCompare() {
  els.leftText.value = "";
  els.rightText.value = "";
  els.leftFile.value = "";
  els.rightFile.value = "";
  els.leftFileName.textContent = "未导入";
  els.rightFileName.textContent = "未导入";
  els.leftHint.textContent = "支持 DOC、DOCX、PDF、Excel、PPT、TXT。";
  els.rightHint.textContent = "上传后仍可在文本框内微调。";
  state.leftName = "";
  state.rightName = "";
  state.leftHtml = "";
  state.rightHtml = "";
  state.diff = [];
  state.stats = { party: 0, missing: 0, added: 0, inconsistent: 0 };
  editors.left.setHtml("");
  editors.right.setHtml("");
  clearAnnotations();
  if (els.bubbleLayer) els.bubbleLayer.innerHTML = "";
  if (els.bubbleList) els.bubbleList.innerHTML = "";
  if (els.bubbleEmpty) els.bubbleEmpty.style.display = "";
  els.partyCount.textContent = "0";
  els.missingCount.textContent = "0";
  els.addedCount.textContent = "0";
  els.inconsistentCount.textContent = "0";
  els.downloadReportBtn.disabled = true;
  els.drawerToggleBtn.hidden = true;
  if (els.drawerCountBadge) els.drawerCountBadge.textContent = "0";
  els.summary.textContent = "导入两份合同后开始对比。";
  closeDrawer();
  hidePopover();
}

/* ===== 事件监听 ===== */

els.leftFile.addEventListener("change", () => parseUploadedFile(els.leftFile.files[0], "left"));
els.rightFile.addEventListener("change", () => parseUploadedFile(els.rightFile.files[0], "right"));
els.compareBtn.addEventListener("click", compareContracts);
els.resetBtn.addEventListener("click", resetCompare);
els.downloadReportBtn.addEventListener("click", downloadReport);

// Re-bubble after window resize (debounced)
let resizeTimer;
window.addEventListener("resize", () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => {
    if (state.diff.length && els.bubbleLayer) {
      const visible = state.diff.filter((item) => item.type !== "equal");
      els.bubbleLayer.innerHTML = placeBubbleDots(visible);
    }
  }, 300);
});

checkHealth();
// deploy 1779762013
