const state = {
  documents: [],
  rules: [],
  dimensions: {},
};

const API_BASE = window.location.protocol === "file:" ? "http://localhost:5173" : "";
let selectedUploadFiles = [];

const els = {
  serverState: document.getElementById("server-state"),
  uploadForm: document.getElementById("upload-form"),
  uploadBtn: document.getElementById("upload-btn"),
  fileInput: document.getElementById("file-input"),
  fileHint: document.getElementById("file-hint"),
  fileList: document.getElementById("file-list"),
  uploadMessage: document.getElementById("upload-message"),
  documentBody: document.getElementById("document-body"),
  selectAllDocs: document.getElementById("select-all-docs"),
  docSearch: document.getElementById("doc-search"),
  docParseFilter: document.getElementById("doc-parse-filter"),
  docTypeFilter: document.getElementById("doc-type-filter"),
  extractSelectedBtn: document.getElementById("extract-selected-btn"),
  extractAllBtn: document.getElementById("extract-all-btn"),
  extractStatus: document.getElementById("extract-status"),
  extractStatusTitle: document.getElementById("extract-status-title"),
  extractStatusText: document.getElementById("extract-status-text"),
  acceptAllBtn: document.getElementById("accept-all-btn"),
  exportRulesBtn: document.getElementById("export-rules-btn"),
  importRulesBtn: document.getElementById("import-rules-btn"),
  importRulesInput: document.getElementById("import-rules-input"),
  customRuleForm: document.getElementById("custom-rule-form"),
  ruleList: document.getElementById("rule-list"),
  dimensionList: document.getElementById("dimension-list"),
  ruleCategoryFilter: document.getElementById("rule-category-filter"),
  dimensionFilter: document.getElementById("dimension-filter"),
  statusFilter: document.getElementById("status-filter"),
  ruleDetailModal: document.getElementById("rule-detail-modal"),
  ruleDetailForm: document.getElementById("rule-detail-form"),
  ruleDetailClose: document.getElementById("rule-detail-close"),
  docCount: document.getElementById("doc-count"),
  ruleCount: document.getElementById("rule-count"),
  pendingCount: document.getElementById("pending-count"),
  dimensionCount: document.getElementById("dimension-count"),
};

const RULE_TYPE_OPTIONS = ["通用规则", "必备条款规则", "禁止条款规则", "审批规则", "风险提示规则", "信息追问规则", "条款推荐规则"];
const SCENARIO_OPTIONS = ["通用规则", "审批规则", "交付规则", "付款规则", "验收规则", "违约责任规则", "保密规则", "知识产权规则", "期限终止规则"];
const RULE_BASIS_OPTIONS = ["通用法规", "行业惯例", "企业自定"];
const INDUSTRY_RULE_DOMAIN_ALIASES = ["软件外包", "系统集成", "制造业", "工业制造业"];
const MANAGEMENT_CATEGORIES = ["通用规则", "行业预设", "企业自定"];

async function api(path, options = {}) {
  const res = await fetch(`${API_BASE}${path}`, options);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `请求失败：${res.status}`);
  return data;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function formatDate(value) {
  if (!value) return "";
  return new Date(value).toLocaleString("zh-CN", { hour12: false });
}

function statusLabel(status) {
  const map = {
    pending_review: "待核对",
    active: "已启用",
    inactive: "已关闭",
    rejected: "已驳回",
  };
  return map[status] || status || "待核对";
}

function statusClass(status) {
  if (status === "active") return "ok";
  if (status === "inactive") return "neutral";
  if (status === "rejected") return "high";
  return "mid";
}

function parseStatusLabel(status) {
  const map = {
    parsed: "已解析",
    partial: "待解析",
    failed: "解析失败",
  };
  return map[status] || "待解析";
}

function parseStatusClass(status) {
  if (status === "parsed") return "ok";
  if (status === "failed") return "high";
  return "mid";
}

function docStatusRank(status) {
  if (status === "failed") return 0;
  if (status === "partial") return 1;
  if (status !== "parsed") return 2;
  return 3;
}

function docTypeOptions(selected = "") {
  const types = ["法律法规", "规章制度", "合同模板", "历史合同", "交易习惯", "其他资料"];
  return types
    .map((type) => `<option value="${escapeHtml(type)}" ${type === selected ? "selected" : ""}>${escapeHtml(type)}</option>`)
    .join("");
}

function riskClass(level) {
  if (level === "高") return "high";
  if (level === "低") return "ok";
  return "mid";
}

function displayRuleType(ruleType) {
  return ruleType === "生成约束规则" ? "通用规则" : ruleType || "通用规则";
}

function optionHtml(options, selected = "") {
  return options
    .map((item) => `<option value="${escapeHtml(item)}" ${item === selected ? "selected" : ""}>${escapeHtml(item)}</option>`)
    .join("");
}

function displayRuleBasis(rule) {
  if (rule.ruleBasis) return rule.ruleBasis;
  const text = [rule.sourceDocName, rule.sourceQuote, rule.ruleName, rule.action].join(" ");
  if (/法律|法规|条例|办法|司法解释|国家标准|监管|规章/.test(text)) return "通用法规";
  if (/行业|惯例|习惯|标准|协会|实践|做法/.test(text)) return "行业惯例";
  return "企业自定";
}

function ruleManagementCategory(rule = {}) {
  if (rule.ruleBasis === "企业自定" || rule.ruleSource === "公司规则" || rule.sourceDocName === "公司自己的规则") return "企业自定";
  if (
    /^PRESET_INDUSTRY_(SOFTWARE|INTEGRATION|MANUFACTURING)/.test(String(rule.id || "")) ||
    INDUSTRY_RULE_DOMAIN_ALIASES.includes(rule.businessDomain)
  ) {
    return "行业预设";
  }
  if (MANAGEMENT_CATEGORIES.includes(rule.ruleCategory)) return rule.ruleCategory;
  return "通用规则";
}

function updateStats() {
  els.docCount.textContent = state.documents.length;
  els.ruleCount.textContent = state.rules.length;
  els.pendingCount.textContent = state.rules.filter((rule) => rule.reviewStatus === "pending_review").length;
  els.dimensionCount.textContent = Object.keys(state.dimensions).length;
}

function getFilteredRules() {
  const ruleCategory = els.ruleCategoryFilter?.value || "";
  const dimension = els.dimensionFilter.value;
  const status = els.statusFilter.value;
  return state.rules.filter((rule) => {
    if (ruleCategory && ruleManagementCategory(rule) !== ruleCategory) return false;
    if (dimension && rule.dimension !== dimension) return false;
    if (status && rule.reviewStatus !== status) return false;
    return true;
  });
}

function renderDocuments() {
  renderDocumentTypeFilter();
  const documents = getFilteredDocuments();
  if (!documents.length) {
    els.documentBody.innerHTML = `<tr><td colspan="7">暂无文档，请先上传资料。</td></tr>`;
    return;
  }
  els.documentBody.innerHTML = documents
    .map((doc) => `
      <tr>
        <td><input type="checkbox" class="doc-check" value="${escapeHtml(doc.id)}" /></td>
        <td><strong>${escapeHtml(doc.name)}</strong><br><span class="tag">${formatDate(doc.createdAt)}</span></td>
        <td><select class="doc-edit" data-doc-id="${escapeHtml(doc.id)}" data-doc-field="docType">${docTypeOptions(doc.docType)}</select></td>
        <td><input class="doc-edit" data-doc-id="${escapeHtml(doc.id)}" data-doc-field="domain" value="${escapeHtml(doc.domain || "")}" placeholder="通用" /></td>
        <td><input class="doc-edit" data-doc-id="${escapeHtml(doc.id)}" data-doc-field="contractType" value="${escapeHtml(doc.contractType || "")}" placeholder="通用合同" /></td>
        <td><span class="tag ${parseStatusClass(doc.parseStatus)}">${parseStatusLabel(doc.parseStatus)}</span></td>
        <td class="summary-cell"><div class="summary-text">${escapeHtml(doc.summary || doc.parseMessage || "无摘要")}</div></td>
      </tr>
    `)
    .join("");
}

function getFilteredDocuments() {
  const keyword = (els.docSearch?.value || "").trim().toLowerCase();
  const parseFilter = els.docParseFilter?.value || "";
  const typeFilter = els.docTypeFilter?.value || "";
  return state.documents
    .filter((doc) => {
      if (parseFilter === "unparsed" && doc.parseStatus === "parsed") return false;
      if (parseFilter && parseFilter !== "unparsed" && doc.parseStatus !== parseFilter) return false;
      if (typeFilter && doc.docType !== typeFilter) return false;
      if (!keyword) return true;
      const text = [doc.name, doc.docType, doc.domain, doc.contractType, doc.summary, doc.parseMessage].join(" ").toLowerCase();
      return text.includes(keyword);
    })
    .sort((a, b) => {
      const rank = docStatusRank(a.parseStatus) - docStatusRank(b.parseStatus);
      if (rank) return rank;
      return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
    });
}

function renderDocumentTypeFilter() {
  if (!els.docTypeFilter) return;
  const current = els.docTypeFilter.value;
  const types = [...new Set(state.documents.map((doc) => doc.docType).filter(Boolean))].sort();
  els.docTypeFilter.innerHTML = `<option value="">全部资料类型</option>${types
    .map((type) => `<option value="${escapeHtml(type)}" ${type === current ? "selected" : ""}>${escapeHtml(type)}</option>`)
    .join("")}`;
}

function renderDimensions() {
  const names = Object.keys(state.dimensions);
  els.dimensionFilter.innerHTML = `<option value="">全部维度</option>${names
    .map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
    .join("")}`;
  els.dimensionList.innerHTML = names
    .map((name) => {
      const item = state.dimensions[name];
      const categories = item.categories || [];
      return `
        <article class="dimension-card">
          <h3>${escapeHtml(name)}</h3>
          <p>${escapeHtml(item.description || "")}</p>
          <ul>
            ${categories
              .slice(0, 12)
              .map((cat) => `<li><strong>${escapeHtml(cat["条款名称"])}</strong>：${escapeHtml(cat["条款说明"] || "")}</li>`)
              .join("")}
          </ul>
        </article>
      `;
    })
    .join("");
}

function groupRules(rules) {
  return rules.reduce((acc, rule) => {
    const key = rule.scenario || inferScenario(rule.ruleName, rule.action, rule.triggerCondition);
    if (!acc[key]) acc[key] = [];
    acc[key].push(rule);
    return acc;
  }, {});
}

function groupRulesByManagementCategory(rules) {
  return rules.reduce((acc, rule) => {
    const key = ruleManagementCategory(rule);
    if (!acc[key]) acc[key] = [];
    acc[key].push(rule);
    return acc;
  }, {});
}

function inferScenario(...parts) {
  const text = parts.join(" ");
  if (/审批|批准|授权|权限|盖章|签批/.test(text)) return "审批规则";
  if (/付款|支付|价款|发票|结算|预付款|尾款|账期/.test(text)) return "付款规则";
  if (/交付|交货|交接|交付物|里程碑|期限|延期/.test(text)) return "交付规则";
  if (/验收|测试|试运行|确认|验收标准/.test(text)) return "验收规则";
  if (/违约|赔偿|责任|损失|处罚|补偿/.test(text)) return "违约责任规则";
  if (/保密|秘密|数据|隐私|个人信息/.test(text)) return "保密规则";
  if (/知识产权|著作权|专利|商标|源码|许可/.test(text)) return "知识产权规则";
  if (/解除|终止|续签|期限|生效/.test(text)) return "期限终止规则";
  return "通用规则";
}

function renderRules() {
  const rules = getFilteredRules();
  if (!rules.length) {
    els.ruleList.innerHTML = `<div class="notice">暂无符合筛选条件的规则。</div>`;
    return;
  }

  const pending = rules.filter((rule) => rule.reviewStatus === "pending_review");
  const handled = rules.filter((rule) => rule.reviewStatus !== "pending_review");
  const pendingHtml = pending.length
    ? `<section class="review-panel">
        <div class="review-panel-head">
          <h3>需要人工核对</h3>
          <span class="tag mid">${pending.length} 条</span>
        </div>
        <div class="pending-grid">${pending.map(renderPendingRuleCard).join("")}</div>
      </section>`
    : "";

  const categoryOrder = ["通用规则", "行业预设", "企业自定"];
  const handledByCategory = groupRulesByManagementCategory(handled);
  const handledHtml = categoryOrder
    .filter((category) => handledByCategory[category]?.length)
    .map((category) => renderRuleCategoryFolder(category, handledByCategory[category]))
    .join("");

  els.ruleList.innerHTML = `${pendingHtml}${handledHtml || (pendingHtml ? "" : `<div class="notice">暂无已处理规则。</div>`)}`;
}

function renderRuleCategoryFolder(categoryName, items) {
  const activeCount = items.filter((rule) => rule.reviewStatus === "active").length;
  const industryHint =
    categoryName === "行业预设"
      ? `<span class="tag neutral">软件外包 / 系统集成 / 制造业</span>`
      : "";
  const scenarioHtml = Object.entries(groupRules(items))
    .map(([folderName, folderItems]) => renderRuleFolder(folderName, folderItems, false))
    .join("");
  return `
    <details class="rule-folder rule-category-folder" open>
      <summary>
        <span class="folder-icon">▾</span>
        <strong>${escapeHtml(categoryName)}</strong>
        <span class="tag">${items.length} 条</span>
        <span class="tag ok">已启用 ${activeCount}</span>
        ${industryHint}
      </summary>
      <div class="rule-category-body">
        ${scenarioHtml}
      </div>
    </details>
  `;
}

function renderPendingRuleCard(rule) {
  return `
    <article class="rule-card pending-card">
      <h3>${escapeHtml(rule.ruleName)}</h3>
      <div class="rule-meta">
        <span class="tag detail">${escapeHtml(ruleManagementCategory(rule))}</span>
        <span class="tag">${escapeHtml(rule.dimension)}</span>
        <span class="tag">${escapeHtml(displayRuleBasis(rule))}</span>
        <span class="tag">${escapeHtml(displayRuleType(rule.ruleType))}</span>
        <span class="tag ${riskClass(rule.riskLevel)}">${escapeHtml(rule.riskLevel)}风险</span>
        <span class="tag mid">待核对</span>
      </div>
      <dl>
        <dt>合同类型</dt><dd>${escapeHtml((rule.contractType || []).join("、"))}</dd>
        <dt>业务领域</dt><dd>${escapeHtml(rule.businessDomain)}</dd>
        <dt>触发条件</dt><dd>${escapeHtml(rule.triggerCondition)}</dd>
        <dt>执行动作</dt><dd>${escapeHtml(rule.action)}</dd>
        <dt>优先级</dt><dd>${escapeHtml(rule.priority)}</dd>
        <dt>规则归属</dt><dd>${escapeHtml(displayRuleBasis(rule))}</dd>
        <dt>来源</dt><dd>${escapeHtml(rule.sourceDocName || rule.sourceDocId)}</dd>
      </dl>
      <div class="quote">${escapeHtml(rule.sourceQuote || "未提供原文依据")}</div>
      <div class="card-actions">
        <button class="secondary" data-detail="${escapeHtml(rule.id)}">详情</button>
        <button class="secondary" data-review="${escapeHtml(rule.id)}" data-status="active">启用</button>
        <button class="danger" data-review="${escapeHtml(rule.id)}" data-status="rejected">驳回</button>
      </div>
    </article>
  `;
}

function renderRuleFolder(folderName, items, open = true) {
      const activeCount = items.filter((rule) => rule.reviewStatus === "active").length;
      const inactiveCount = items.filter((rule) => rule.reviewStatus === "inactive").length;
      const rejectedCount = items.filter((rule) => rule.reviewStatus === "rejected").length;
  return `
    <details class="rule-folder" ${open ? "open" : ""}>
      <summary>
        <span class="folder-icon">▾</span>
        <strong>${escapeHtml(folderName)}</strong>
        <span class="tag">${items.length} 条</span>
        <span class="tag ok">已启用 ${activeCount}</span>
        <span class="tag neutral">已关闭 ${inactiveCount}</span>
        <span class="tag high">已驳回 ${rejectedCount}</span>
      </summary>
      <div class="rule-row-list">
        ${items.map(renderHandledRuleRow).join("")}
      </div>
    </details>
  `;
}

function renderHandledRuleRow(rule) {
  return `
    <div class="rule-row">
      <div>
        <strong>${escapeHtml(rule.ruleName)}</strong>
        <small>${escapeHtml(rule.dimension)} / ${escapeHtml(rule.businessDomain || "通用")}</small>
      </div>
      <span class="tag">${escapeHtml(displayRuleType(rule.ruleType))}</span>
      <span class="tag detail">${escapeHtml(ruleManagementCategory(rule))}</span>
      <span class="tag">${escapeHtml(displayRuleBasis(rule))}</span>
      <span class="tag ${riskClass(rule.riskLevel)}">${escapeHtml(rule.riskLevel)}风险</span>
      <span class="tag ${statusClass(rule.reviewStatus)}">${statusLabel(rule.reviewStatus)}</span>
      <span class="row-source">${escapeHtml(rule.sourceDocName || rule.sourceDocId || "导入规则")}</span>
      <div class="row-actions">
        ${
          rule.reviewStatus === "active"
            ? `<button class="mini" data-review="${escapeHtml(rule.id)}" data-status="inactive">关闭</button>`
            : `<button class="mini" data-review="${escapeHtml(rule.id)}" data-status="active">启用</button>`
        }
        <button class="mini" data-detail="${escapeHtml(rule.id)}">详情</button>
        <button class="mini danger" data-delete="${escapeHtml(rule.id)}">删除</button>
      </div>
    </div>
  `;
}

function findRule(ruleId) {
  return state.rules.find((rule) => rule.id === ruleId);
}

function openRuleDetail(ruleId) {
  const rule = findRule(ruleId);
  if (!rule) return;
  const form = els.ruleDetailForm;
  form.elements.ruleId.value = rule.id;
  form.elements.ruleName.value = rule.ruleName || "";
  form.elements.dimension.innerHTML = optionHtml(Object.keys(state.dimensions), rule.dimension);
  form.elements.ruleType.innerHTML = optionHtml(RULE_TYPE_OPTIONS, displayRuleType(rule.ruleType));
  form.elements.scenario.innerHTML = optionHtml(SCENARIO_OPTIONS, rule.scenario || inferScenario(rule.ruleName, rule.action));
  form.elements.riskLevel.value = rule.riskLevel || "中";
  form.elements.ruleBasis.value = displayRuleBasis(rule);
  form.elements.contractType.value = (rule.contractType || []).join("、");
  form.elements.businessDomain.value = rule.businessDomain || "通用";
  form.elements.priority.value = rule.priority || 50;
  form.elements.reviewStatus.value = rule.reviewStatus || "pending_review";
  form.elements.triggerCondition.value = rule.triggerCondition || "";
  form.elements.action.value = rule.action || "";
  form.elements.sourceQuote.value = rule.sourceQuote || "";
  els.ruleDetailModal.hidden = false;
}

function closeRuleDetail() {
  els.ruleDetailModal.hidden = true;
}

async function loadAll() {
  const [health, dimensions, documents, rules] = await Promise.all([
    api("/api/health"),
    api("/api/dimensions"),
    api("/api/documents"),
    api("/api/rules"),
  ]);
  els.serverState.textContent = health.hasApiKey ? `AI 已配置：${health.model}，调用时实时校验` : "服务可用，未配置 API Key";
  els.serverState.className = `server-state ${health.hasApiKey ? "ok" : "warn"}`;
  state.dimensions = dimensions;
  state.documents = documents.documents || [];
  state.rules = rules.rules || [];
  renderDocuments();
  renderDimensions();
  renderRules();
  updateStats();
}

function selectedDocumentIds() {
  return [...document.querySelectorAll(".doc-check:checked")].map((item) => item.value);
}

function startExtractFeedback(button, documentCount) {
  const originalText = button.textContent;
  const messages = [
    `正在读取 ${documentCount} 个选中文档...`,
    "正在整理条款、制度和合同片段...",
    "正在调用 AI 抽取候选规则...",
    "正在匹配规则维度和合同场景...",
    "正在执行相似规则去重...",
    "文档较多时可能需要 30-90 秒，请保持页面打开。",
  ];
  let index = 0;

  button.disabled = true;
  button.classList.add("loading");
  button.innerHTML = `<span class="btn-spinner"></span><span>AI 提取中</span>`;

  if (els.extractStatus) {
    els.extractStatus.hidden = false;
    els.extractStatus.classList.remove("done", "error");
    els.extractStatusTitle.textContent = "AI 提取中";
    els.extractStatusText.textContent = messages[index];
  }

  const timer = window.setInterval(() => {
    index = (index + 1) % messages.length;
    if (els.extractStatusText) els.extractStatusText.textContent = messages[index];
  }, 2800);

  return {
    done(title, text, className = "done") {
      window.clearInterval(timer);
      button.disabled = false;
      button.classList.remove("loading");
      button.textContent = originalText;
      if (els.extractStatus) {
        els.extractStatus.hidden = false;
        els.extractStatus.classList.remove("done", "error");
        els.extractStatus.classList.add(className);
        els.extractStatusTitle.textContent = title;
        els.extractStatusText.textContent = text;
      }
    },
    reset() {
      window.clearInterval(timer);
      button.disabled = false;
      button.classList.remove("loading");
      button.textContent = originalText;
    },
  };
}

async function extractRules(documentIds, button) {
  if (!documentIds.length) {
    alert("没有可提取的文档。");
    return;
  }
  const feedback = startExtractFeedback(button, documentIds.length);
  try {
    const data = await api("/api/extract-rules", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ documentIds }),
    });
    await loadAll();
    const skippedText = data.skipped ? `，跳过重复 ${data.skipped} 条` : "";
    const prefix = data.usedFallback ? `${data.fallbackReason || "AI 调用失败，已使用本地兜底抽取。"} ` : "AI 提取完成。";
    feedback.done("提取完成", `新增 ${data.rules.length} 条规则，跳过重复 ${data.skipped || 0} 条。`);
    alert(`${prefix}新增 ${data.rules.length} 条规则${skippedText}。`);
  } catch (error) {
    feedback.done("提取失败", error.message, "error");
    alert(error.message);
  }
}

function downloadJson(filename, data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function fileThumbLabel(file) {
  const ext = (file.name.split(".").pop() || "FILE").slice(0, 4).toUpperCase();
  if (file.type.startsWith("image/")) return "IMG";
  if (/pdf/i.test(file.type) || ext === "PDF") return "PDF";
  if (/word|document/i.test(file.type) || ["DOC", "DOCX"].includes(ext)) return "DOC";
  if (["XLS", "XLSX", "CSV"].includes(ext)) return "XLS";
  return ext || "FILE";
}

function syncUploadInputFiles() {
  const transfer = new DataTransfer();
  selectedUploadFiles.forEach((file) => transfer.items.add(file));
  els.fileInput.files = transfer.files;
}

function renderSelectedFiles() {
  els.fileHint.textContent = selectedUploadFiles.length ? `已选择 ${selectedUploadFiles.length} 个文件` : "可一次选择多个文件";
  els.fileList.innerHTML = selectedUploadFiles
    .map(
      (file, index) => `
        <div class="file-chip" title="${escapeHtml(file.name)}">
          <span class="file-thumb">${escapeHtml(fileThumbLabel(file))}</span>
          <span class="file-name">${escapeHtml(file.name)}</span>
          <button type="button" aria-label="删除 ${escapeHtml(file.name)}" data-remove-file="${index}">×</button>
        </div>
      `
    )
    .join("");
}

els.fileInput.addEventListener("change", () => {
  selectedUploadFiles = [...els.fileInput.files];
  if (selectedUploadFiles.length > 1) els.uploadForm.elements.docType.value = "";
  renderSelectedFiles();
});

els.fileList.addEventListener("click", (event) => {
  const button = event.target.closest("[data-remove-file]");
  if (!button) return;
  event.preventDefault();
  event.stopPropagation();
  selectedUploadFiles.splice(Number(button.dataset.removeFile), 1);
  syncUploadInputFiles();
  renderSelectedFiles();
});

els.uploadBtn.addEventListener("click", async () => {
  if (!selectedUploadFiles.length) {
    els.uploadMessage.textContent = "请先选择要上传的文档。";
    return;
  }
  syncUploadInputFiles();
  const formData = new FormData(els.uploadForm);
  els.uploadBtn.disabled = true;
  els.uploadMessage.textContent = "正在上传并解析文档...";
  try {
    const data = await api("/api/documents", { method: "POST", body: formData });
    els.uploadMessage.textContent = `已归档 ${data.documents.length} 个文档。`;
    els.uploadForm.reset();
    selectedUploadFiles = [];
    renderSelectedFiles();
    await loadAll();
  } catch (error) {
    els.uploadMessage.textContent = error.message;
  } finally {
    els.uploadBtn.disabled = false;
  }
});

els.selectAllDocs.addEventListener("change", () => {
  document.querySelectorAll(".doc-check").forEach((item) => {
    item.checked = els.selectAllDocs.checked;
  });
});

els.documentBody.addEventListener("change", async (event) => {
  const input = event.target.closest(".doc-edit");
  if (!input) return;
  input.disabled = true;
  try {
    await api("/api/documents/update", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        documentId: input.dataset.docId,
        field: input.dataset.docField,
        value: input.value.trim(),
      }),
    });
    await loadAll();
  } catch (error) {
    alert(error.message);
    await loadAll();
  } finally {
    input.disabled = false;
  }
});

els.extractSelectedBtn.addEventListener("click", () => {
  const ids = selectedDocumentIds();
  if (!ids.length) {
    alert("请至少选择一个文档。");
    return;
  }
  extractRules(ids, els.extractSelectedBtn);
});

els.extractAllBtn.addEventListener("click", () => {
  const ids = state.documents.filter((doc) => doc.parseStatus === "parsed").map((doc) => doc.id);
  extractRules(ids, els.extractAllBtn);
});

els.acceptAllBtn.addEventListener("click", async () => {
  const pendingCount = state.rules.filter((rule) => rule.reviewStatus === "pending_review").length;
  if (!pendingCount) {
    alert("当前没有待核对规则。");
    return;
  }
  els.acceptAllBtn.disabled = true;
  try {
    await api("/api/rules/accept-all", { method: "POST" });
    await loadAll();
  } catch (error) {
    alert(error.message);
  } finally {
    els.acceptAllBtn.disabled = false;
  }
});

els.exportRulesBtn.addEventListener("click", async () => {
  try {
    const data = await api("/api/rules/export");
    downloadJson(`contract-rules-${new Date().toISOString().slice(0, 10)}.json`, data);
  } catch (error) {
    alert(error.message);
  }
});

els.importRulesBtn.addEventListener("click", () => els.importRulesInput.click());

els.importRulesInput.addEventListener("change", async () => {
  const file = els.importRulesInput.files[0];
  if (!file) return;
  const formData = new FormData();
  formData.append("rules", file);
  try {
    const data = await api("/api/rules/import", { method: "POST", body: formData });
    await loadAll();
    alert(`已导入 ${data.rules.length} 条规则${data.skipped ? `，跳过重复 ${data.skipped} 条` : ""}。`);
  } catch (error) {
    alert(error.message);
  } finally {
    els.importRulesInput.value = "";
  }
});

els.ruleList.addEventListener("click", async (event) => {
  const detailButton = event.target.closest("[data-detail]");
  if (detailButton) {
    openRuleDetail(detailButton.dataset.detail);
    return;
  }
  const deleteButton = event.target.closest("[data-delete]");
  if (deleteButton) {
    if (!window.confirm("确认删除这条规则？删除后不会出现在规则库列表中。")) return;
    await api("/api/rules/delete", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ ruleId: deleteButton.dataset.delete }),
    });
    await loadAll();
    return;
  }
  const button = event.target.closest("[data-review]");
  if (!button) return;
  await api("/api/rules/review", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ ruleId: button.dataset.review, status: button.dataset.status }),
  });
  await loadAll();
});

els.ruleDetailClose.addEventListener("click", closeRuleDetail);

els.ruleDetailModal.addEventListener("click", (event) => {
  if (event.target === els.ruleDetailModal) closeRuleDetail();
});

els.ruleDetailForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(els.ruleDetailForm);
  const contractType = String(formData.get("contractType") || "")
    .split(/[,，、]/)
    .map((item) => item.trim())
    .filter(Boolean);
  await api("/api/rules/update", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      ruleId: formData.get("ruleId"),
      updates: {
        ruleName: formData.get("ruleName"),
        dimension: formData.get("dimension"),
        ruleType: formData.get("ruleType"),
        scenario: formData.get("scenario"),
        riskLevel: formData.get("riskLevel"),
        ruleBasis: formData.get("ruleBasis"),
        contractType: contractType.length ? contractType : ["通用合同"],
        businessDomain: formData.get("businessDomain"),
        priority: formData.get("priority"),
        reviewStatus: formData.get("reviewStatus"),
        triggerCondition: formData.get("triggerCondition"),
        action: formData.get("action"),
        sourceQuote: formData.get("sourceQuote"),
      },
    }),
  });
  closeRuleDetail();
  await loadAll();
});

els.customRuleForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(els.customRuleForm);
  const scopeText = String(formData.get("scope") || "").trim();
  const ruleText = String(formData.get("ruleText") || "").trim();
  const scenario = formData.get("scenario") || inferScenario(ruleText, scopeText);
  const payload = {
    ruleText,
    ruleName: ruleText.slice(0, 40),
    dimension: "通用必备条款",
    scenario,
    ruleType: scenario === "审批规则" ? "审批规则" : "通用规则",
    contractType: scopeText ? scopeText.split(/[,，、]/).map((item) => item.trim()).filter(Boolean) : ["通用合同"],
    businessDomain: "通用",
    scope: scopeText,
    classifyMode: formData.get("classifyMode") || "",
    triggerCondition: scopeText || "合同生成或审查时",
    action: ruleText,
    riskLevel: formData.get("riskLevel"),
    priority: 70,
    ruleBasis: "企业自定",
    sourceQuote: ruleText,
  };
  try {
    const data = await api("/api/rules/custom", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
    });
    els.customRuleForm.reset();
    await loadAll();
    alert(data.rules.length ? "已添加公司自定义规则。" : "规则重复，未重复添加。");
  } catch (error) {
    alert(error.message);
  }
});

els.ruleCategoryFilter?.addEventListener("change", renderRules);
els.dimensionFilter.addEventListener("change", renderRules);
els.statusFilter.addEventListener("change", renderRules);
els.docSearch.addEventListener("input", renderDocuments);
els.docParseFilter.addEventListener("change", renderDocuments);
els.docTypeFilter.addEventListener("change", renderDocuments);

loadAll().catch((error) => {
  els.serverState.textContent = error.message;
  els.serverState.className = "server-state warn";
});
