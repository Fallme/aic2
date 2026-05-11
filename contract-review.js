const API_BASE = window.location.protocol === "file:" ? "http://localhost:5173" : "";

// ── AI Status + Icons ──
window.addEventListener("DOMContentLoaded", () => {
  // Set sidebar icons
  const iconMap = { navTF: "templateFill", navSD: "smartDraft", navCR: "contractReview", navKB: "knowledge" };
  for (const [id, name] of Object.entries(iconMap)) {
    const el = document.getElementById(id);
    if (el && window.AppIcons) el.innerHTML = AppIcons[name] || "";
  }
  // AI status
  fetch(`${API_BASE}/api/health`).then((r) => r.json()).then((d) => {
    const m = document.getElementById("aiModel");
    const p = document.getElementById("aiProvider");
    const dot = document.getElementById("aiDot");
    if (m) m.textContent = d.model || "unknown";
    if (p) p.textContent = d.provider || "";
    if (dot) dot.style.background = d.apiKeyConfigured ? "var(--green)" : "var(--red)";
  }).catch(() => {
    const m = document.getElementById("aiModel");
    const dot = document.getElementById("aiDot");
    if (m) m.textContent = "离线";
    if (dot) dot.style.background = "var(--red)";
  });
});

const els = {
  serverState: document.getElementById("server-state"),
  reviewFile: document.getElementById("review-file"),
  reviewFileHint: document.getElementById("review-file-hint"),
  reviewText: document.getElementById("review-text"),
  reviewImportText: document.getElementById("review-import-text"),
  reviewBtn: document.getElementById("review-btn"),
  reviewSource: document.getElementById("review-editor-container"),
  reviewIssues: document.getElementById("review-issues"),
  reviewWorkflow: document.getElementById("review-workflow"),
  reviewSummary: document.getElementById("review-summary"),
  sourceMeta: document.getElementById("source-meta"),
  reviewScoreboard: document.getElementById("review-scoreboard"),
  reviewFilterBar: document.getElementById("review-filter-bar"),
  issueNav: document.getElementById("issue-nav"),
  copyCorrectedBtn: document.getElementById("copy-corrected-btn"),
  downloadCorrectedBtn: document.getElementById("download-corrected-btn"),
  prevIssueBtn: document.getElementById("prev-issue-btn"),
  nextIssueBtn: document.getElementById("next-issue-btn"),
  issueCounter: document.getElementById("issue-counter"),
  acceptAllBtn: document.getElementById("accept-all-btn"),
  exportReportBtn: document.getElementById("export-report-btn"),
  openImportBtn: document.getElementById("open-import-btn"),
  importModal: document.getElementById("review-import-modal"),
  closeImportBtn: document.getElementById("close-import-btn"),
  confirmImportBtn: document.getElementById("confirm-import-btn"),
  clearImportBtn: document.getElementById("clear-import-btn"),
};

const state = {
  issues: [],
  reviewData: null,
  filter: "summary",
  activeIssueIndex: 0,
  originalText: "",
  workingText: "",
  fileName: "",
  templateDocxId: "",
  templatePreserved: false,
  reviewComplete: false,
  model: "",
  modelProvider: "",
  hasApiKey: false,
};

const reviewSteps = [
  "正在读取合同文本，识别可审查内容。",
  "正在执行主体审查，核对甲乙方、授权和签署信息。",
  "正在调用大模型进行深度审查，复核逻辑、金额和风险表述。",
  "正在执行内容审查，检查付款、交付、验收和违约闭环。",
  "正在执行格式审查，检查编号、附件、签署栏和用语问题。",
  "正在匹配规则库、企业自定规则和常见问题。",
  "正在定位合同原文中的风险片段并整理处理意见。",
  "文档较长时可能需要 30-90 秒，请稍候。",
];

let reviewProgressTimer = null;
let officeEditor = null;
const STORAGE_KEY = "contract-review-workbench-v1";

function saveState() {
  const editor = document.getElementById("source-editor");
  const payload = {
    issues: state.issues,
    reviewData: state.reviewData,
    filter: state.filter,
    activeIssueIndex: state.activeIssueIndex,
    originalText: state.originalText,
    workingText: sourceEditorText(),
    fileName: state.fileName,
    templateDocxId: state.templateDocxId,
    templatePreserved: state.templatePreserved,
    reviewComplete: state.reviewComplete,
    model: state.model,
    modelProvider: state.modelProvider,
    hasApiKey: state.hasApiKey,
    sourceHtml: officeEditor ? officeEditor.getHtml() : (editor ? editor.outerHTML : ""),
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function removeLegacyRiskMarks(root) {
  root.querySelectorAll("mark.risk-mark").forEach((mark) => {
    const text = mark.textContent || "";
    if (mark.classList.contains("anchor") || /^需补充/.test(text)) mark.remove();
    else mark.replaceWith(document.createTextNode(text));
  });
}

function restoreState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return false;
  try {
    const saved = JSON.parse(raw);
    state.issues = Array.isArray(saved.issues) ? saved.issues.map(normalizeIssueForReview) : [];
    state.reviewData = saved.reviewData || null;
    if (state.reviewData) state.reviewData.issues = state.issues;
    state.filter = saved.filter === "all" ? "summary" : saved.filter || "summary";
    state.activeIssueIndex = Number(saved.activeIssueIndex || 0);
    state.originalText = saved.originalText || "";
    state.workingText = saved.workingText || "";
    state.fileName = saved.fileName || "";
    state.templateDocxId = saved.templateDocxId || "";
    state.templatePreserved = Boolean(saved.templatePreserved && state.templateDocxId);
    state.reviewComplete = Boolean(saved.reviewComplete);
    state.model = state.model || saved.model || "";
    state.modelProvider = state.modelProvider || saved.modelProvider || "";
    state.hasApiKey = state.hasApiKey || Boolean(saved.hasApiKey);
    els.reviewText.value = state.workingText;
    if (officeEditor && saved.sourceHtml) {
      officeEditor.setHtml(saved.sourceHtml);
    } else if (saved.sourceHtml) {
      const restoredEditor = document.getElementById("source-editor");
      if (restoredEditor) {
        restoredEditor.dataset.officePlugin = "online-office";
        restoredEditor.dataset.officeMode = "word";
      }
    }
    else renderEditableSource({ annotated: state.reviewComplete });
    updateScoreboard(state.reviewData || {}, state.issues);
    renderReviewSummary(state.reviewData, state.issues);
    renderReviewWorkflow(state.reviewData, { phase: state.reviewComplete ? "done" : "idle" });
    updateFilterButtons();
    renderIssueList();
    els.reviewBtn.textContent = state.reviewComplete ? "重新审查" : "开始审查";
    if (els.exportReportBtn) els.exportReportBtn.disabled = !state.reviewComplete;
    els.openImportBtn.classList.toggle("file-loaded", Boolean(state.workingText));
    els.openImportBtn.textContent = state.fileName ? `文件：${state.fileName}` : state.workingText ? "已导入文本合同" : "上传/粘贴合同";
    setSourceMeta(state.workingText);
    return true;
  } catch {
    localStorage.removeItem(STORAGE_KEY);
    return false;
  }
}

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

function aiFailureLabel(reason = "") {
  const text = String(reason || "");
  if (/Xiaomi|MiMo|XIAOMI|MIMO/i.test(text)) return "小米 MiMo 调用未完成，已启用本地规则兜底。";
  if (/OpenAI|insufficient_quota|billing|quota|rate limit|429|credit/i.test(text)) return "OpenAI 额度或计费不可用，已启用本地规则兜底。";
  if (/overdue-payment|good standing|Access denied/i.test(text)) return "大模型账户不可用，已启用本地规则兜底。";
  if (/Missing.*API|API Key|XIAOMI_API_KEY|OPENAI_API_KEY|DASHSCOPE|unauthorized|invalid api/i.test(text)) return "未配置可用大模型 Key，已启用本地规则兜底。";
  if (/timeout|network|fetch|ECONN|ENOTFOUND|socket/i.test(text)) return "大模型网络调用失败，已启用本地规则兜底。";
  return "大模型调用未完成，已启用本地规则兜底。";
}

function riskClass(level) {
  if (level === "高") return "high";
  if (level === "低") return "low";
  return "mid";
}

function riskRank(level) {
  if (level === "高") return 3;
  if (level === "中") return 2;
  return 1;
}

function issueLevel(issue) {
  return issue.risk_level || issue.riskLevel || "中";
}

function issueCategory(issue) {
  return [issue.category, issue.detail_category, issue.detailCategory, issue.source_rule, issue.sourceRule].filter(Boolean).join(" ");
}

function issueDetailCategory(issue) {
  return issue.detail_category || issue.detailCategory || issue.category || "风险问题";
}

function issueMajorCategory(issue) {
  const category = issueCategory(issue);
  if (/规则|法规|企业|行业/.test(category)) return "rule";
  if (/常见问题|清单|格式|编号|排版|标题|附件|签署|盖章|日期/.test(category)) return "common";
  return "risk";
}

function issueMajorLabel(issue) {
  const labels = { risk: "风险问题", rule: "规则库", common: "常见问题" };
  return labels[issueMajorCategory(issue)] || "风险问题";
}

function cleanIssueBasisText(value = "") {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function issueBasisItems(issue = {}) {
  const legalBasis = cleanIssueBasisText(issue.legal_basis || issue.legalBasis);
  const sourceRule = cleanIssueBasisText(issue.source_rule || issue.sourceRule);
  const sourceBasis = cleanIssueBasisText(issue.source_rule_basis || issue.sourceRuleBasis || issue.rule_basis || issue.ruleBasis);
  const sourceQuote = cleanIssueBasisText(issue.source_quote || issue.sourceQuote);
  const items = [];
  if (legalBasis) items.push(`法律/法规依据：${legalBasis}`);
  if (sourceRule) items.push(`规则依据：${sourceRule}`);
  if (sourceBasis) items.push(`依据类型：${sourceBasis}`);
  if (sourceQuote) items.push(`依据片段：${sourceQuote}`);
  return items.filter(Boolean).slice(0, 4);
}

function normalizeIssueForReview(issue = {}) {
  const detailCategory = issueDetailCategory(issue);
  const majorCategory = issueMajorLabel({ ...issue, detail_category: detailCategory });
  return {
    ...issue,
    category: majorCategory,
    detail_category: detailCategory,
    appliedText: issue.appliedText || issue.applied_text || "",
    applied: Boolean(issue.applied),
  };
}

function countsByRisk(issues) {
  return issues.reduce(
    (acc, issue) => {
      const level = issueLevel(issue);
      if (level === "高") acc.high += 1;
      else if (level === "低") acc.low += 1;
      else acc.mid += 1;
      acc.total += 1;
      return acc;
    },
    { high: 0, mid: 0, low: 0, total: 0 }
  );
}

function reviewScore(issues) {
  if (!issues.length) return 100;
  const counts = countsByRisk(issues);
  return Math.max(35, 100 - counts.high * 18 - counts.mid * 8 - counts.low * 2);
}

function issueReviewDimension(issue = {}) {
  const text = issueFullText(issue);
  if (/主体|甲方|乙方|签约|名称|证照|授权|权限|住所|统一社会信用代码|法定代表人|联系人|联系方式/.test(text)) return "subject";
  if (/格式|编号|标题|排版|附件|签署|签章|盖章|日期|落款|页码|空白|错别字|错字|语病|称谓|用语/.test(text)) return "format";
  if (issueMajorCategory(issue) === "rule") return "rule";
  return "content";
}

function dimensionCounts(issues = []) {
  return issues.reduce(
    (acc, issue) => {
      const key = issueReviewDimension(issue);
      acc[key] += 1;
      return acc;
    },
    { content: 0, subject: 0, format: 0, rule: 0 }
  );
}

function updateScoreboard(data = {}, issues = []) {
  const counts = countsByRisk(issues);
  const dims = data.reviewReport?.counts || dimensionCounts(issues);
  const risk = data.overallRisk || data.overall_risk || (counts.high ? "高" : counts.mid ? "中" : counts.total ? "低" : "待审查");
  const score = ["待审查", "审查中"].includes(risk) ? "--" : reviewScore(issues);
  const riskCount = counts.high + counts.mid;
  const aiStatus = data.usedFallback ? "已兜底" : data.usedAI ? "已调用" : risk === "审查中" ? "调用中" : "待调用";
  const modelLabel = currentModelLabel(data);
  const aiDetail = data.usedFallback
    ? "本地规则兜底"
    : data.usedAI
      ? modelLabel
      : risk === "审查中"
        ? modelLabel
        : state.hasApiKey
          ? modelLabel
          : "未配置 API Key";
  els.reviewScoreboard.innerHTML = `
    <article class="${score !== "--" && score >= 85 ? "ok" : score !== "--" && score < 70 ? "high" : "mid"}">
      <span>合同评分</span>
      <strong>${score}</strong>
    </article>
    <article class="${riskClass(risk)}">
      <span>综合风险</span>
      <strong>${escapeHtml(risk)}</strong>
    </article>
    <article class="${counts.high ? "high" : counts.mid ? "mid" : ""}">
      <span>风险问题</span>
      <strong>${riskCount}</strong>
      <small>高 ${counts.high} / 中 ${counts.mid}</small>
    </article>
    <article>
      <span>审查点</span>
      <strong>${counts.total}</strong>
      <small>主体 ${dims.subject || 0} / 格式 ${dims.format || 0}</small>
    </article>
    <article class="${data.usedAI ? "ok" : data.usedFallback ? "mid" : ""}">
      <span>AI审查</span>
      <strong>${escapeHtml(aiStatus)}</strong>
      <small>${escapeHtml(aiDetail)}</small>
    </article>
  `;
}

function currentModelLabel(data = {}) {
  const provider = data.modelProvider || data.model_provider || state.modelProvider || "";
  const model = data.model || state.model || "";
  if (provider && provider !== "None" && model) return `${provider} / ${model}`;
  if (model) return model;
  return data.reviewEngine || (state.hasApiKey ? "大模型已配置" : "大模型待配置");
}

function renderReviewSummary(data = state.reviewData, issues = state.issues) {
  if (!els.reviewSummary) return;
  if (!data) {
    els.reviewSummary.innerHTML = `
      <div>
        <strong>审查结果</strong>
        <p>审查完成后显示整体意见、综合风险和合同评分。</p>
      </div>
      <span class="tag neutral">待审查</span>
    `;
    return;
  }
  const counts = countsByRisk(issues || []);
  const dims = data.reviewReport?.counts || dimensionCounts(issues || []);
  const risk = data.overallRisk || data.overall_risk || (counts.high ? "高" : counts.mid ? "中" : counts.total ? "低" : "待审查");
  const score = ["待审查", "审查中"].includes(risk) ? "--" : reviewScore(issues || []);
  const modelLabel = currentModelLabel(data);
  const aiStatus = data.usedFallback
    ? "大模型未完成，已启用本地兜底"
    : data.usedAI
      ? `大模型已调用：${modelLabel}`
      : risk === "审查中"
        ? `大模型调用中：${modelLabel}`
        : "大模型待调用";
  if (risk === "审查中") {
    els.reviewSummary.innerHTML = `
      <div>
        <strong>审查结果</strong>
        <p>${escapeHtml(data.summary || "正在调用大模型深度审查，并同步匹配规则库、常见问题和格式风险。")}</p>
      </div>
      <div class="summary-tags">
        <span class="tag mid">整体风险：审查中</span>
        <span class="tag neutral">评分：--</span>
        <span class="tag neutral">${escapeHtml(aiStatus)}</span>
      </div>
    `;
    return;
  }
  const fallbackNote = data.usedFallback
    ? `<small class="review-fallback">${escapeHtml(aiFailureLabel(data.fallbackReason))}</small>`
    : "";
  const repairNote = data.aiJsonRepaired ? `<small class="review-fallback">模型结果已完成 JSON 格式修复后解析。</small>` : "";
  els.reviewSummary.innerHTML = `
    <div>
      <strong>审查结果</strong>
      <p>${escapeHtml(data.summary || "审查完成。")}</p>
      <small>内容 ${dims.content || 0} / 主体 ${dims.subject || 0} / 格式 ${dims.format || 0} / 规则 ${dims.rule || 0}；高 ${counts.high} / 中 ${counts.mid}。</small>
      ${fallbackNote || repairNote}
    </div>
    <div class="summary-tags">
      <span class="tag ${riskClass(risk)}">整体风险：${escapeHtml(risk)}</span>
      <span class="tag neutral">评分：${score}</span>
      <span class="tag ${data.usedAI ? "ok" : data.usedFallback ? "mid" : "neutral"}">${escapeHtml(aiStatus)}</span>
    </div>
  `;
}

function cleanWorkflowText(value = "", max = 220) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  return text.length > max ? `${text.slice(0, max)}...` : text;
}

function captureWorkflowLabel(text = "", label = "") {
  const match = new RegExp(`${label}\\s*[：:]\\s*([^\\n；;。]{2,80})`).exec(text);
  return match ? match[1].replace(/[，,。；;].*$/, "").trim() : "";
}

function inferWorkflowProfile(text = "", data = {}) {
  const profile = data.ruleProfile || {};
  let contractType = profile.contractType || "";
  if (!contractType) {
    if (/劳动合同|用人单位|劳动者|工资|试用期|社保/.test(text)) contractType = "劳动合同";
    else if (/软件开发|定制开发|源代码|源码|上线|小程序|APP/.test(text)) contractType = "软件开发合同";
    else if (/系统集成|联调|软硬件|设备|等保/.test(text)) contractType = "系统集成合同";
    else if (/制造|加工|OEM|ODM|设备制造|零部件/.test(text)) contractType = "制造合同";
    else if (/服务|培训|咨询|委托/.test(text)) contractType = "服务合同";
    else contractType = "通用合同";
  }
  return { ...profile, contractType, industry: profile.industry || "" };
}

function localRightsObligations(text = "", profile = {}) {
  const partyA = captureWorkflowLabel(text, "甲方") || "甲方";
  const partyB = captureWorkflowLabel(text, "乙方") || "乙方";
  const type = profile.contractType || "合同";
  const hasPayment = /付款|支付|价款|费用|报酬|合同金额|结算/.test(text);
  const hasDelivery = /交付|提交|完成|提供|服务|培训|开发|上线|验收/.test(text);
  const hasDataIp = /源代码|知识产权|数据|保密|个人信息|资料/.test(text);
  const hasExamTraining = /报名|考试|培训|证书|报考|学习/.test(text);
  if (hasExamTraining) {
    return `${type}中，${partyA}的核心权利是要求${partyB}提供真实、准确、完整的报名考试信息并按约缴纳相关费用；核心义务是按约为${partyB}完成报名考试安排、学习或考试事项协助、证书取得协助，并对${partyB}提交的个人资料和考试资料承担保密与合规使用义务。${partyB}的核心权利是获得${partyA}协助办理的合法有效报名、考试安排或证书取得服务；核心义务是按约缴纳费用、提供真实资料、配合考试安排并遵守考试规则。审查需重点核对报名信息责任、费用支付与退费、考试结果责任、资料保密和证书有效性。`;
  }
  const payerDuty = hasPayment ? "按约验收、接收发票并支付合同价款" : "按约配合履行、确认成果并承担约定费用";
  const deliverDuty = hasDelivery ? "按约完成服务或交付成果，并配合验收、整改和资料提交" : "按约履行合同义务并留存履约记录";
  const extraDuty = hasDataIp ? "，同时遵守保密、数据安全和成果权属约定" : "";
  return `${type}中，${partyA}的核心权利是要求${partyB}按约完成合同事项并提交合格成果；核心义务是${payerDuty}。${partyB}的核心权利是按约取得价款或报酬；核心义务是${deliverDuty}${extraDuty}。审查需重点核对双方权利义务是否对等、付款与交付验收是否闭环、违约责任是否可执行。`;
}

function defaultKnowledgeBases(profile = {}, data = {}) {
  if (data.reviewWorkflow?.knowledgeBases?.length) return data.reviewWorkflow.knowledgeBases;
  const type = profile.contractType || "通用合同";
  const bases = [`${type}法律法规库`, `${type}裁判案例库`, `${type}审查知识库`];
  if (profile.industry) bases.push(`${profile.industry}行业预设规则库`);
  if (data.ruleSelection && /企业自定 [1-9]/.test(data.ruleSelection)) bases.push("企业自定规则库");
  return bases;
}

function workflowStatusClass(status = "pending") {
  if (status === "done") return "done";
  if (status === "current" || status === "running") return "current";
  return "pending";
}

function buildWorkflowTasks(data = {}, issues = [], activeStep = 0, phase = "idle") {
  if (data.reviewWorkflow?.reviewTasks?.length) return data.reviewWorkflow.reviewTasks;
  const done = phase === "done";
  const statusFor = (index) => {
    if (done) return "done";
    const taskStep = Math.max(0, activeStep - 2);
    if (index < taskStep) return "done";
    if (index === taskStep && activeStep >= 2) return "current";
    return "pending";
  };
  const dims = data.reviewReport?.counts || dimensionCounts(issues);
  const suggestions = issues.filter((issue) => canApplyIssue(issue)).length;
  return [
    { label: "实质风险识别", note: `高/中风险 ${countsByRisk(issues).high + countsByRisk(issues).mid} 项`, status: statusFor(0) },
    { label: "文字符号检查", note: `格式和用语 ${dims.format || 0} 项`, status: statusFor(1) },
    { label: "生成修改建议", note: `可应用正文 ${suggestions} 条`, status: statusFor(2) },
    { label: "签约主体审查", note: `主体问题 ${dims.subject || 0} 项`, status: statusFor(3) },
  ];
}

function workflowStageStatus(stageIndex, activeStep = 0, phase = "idle") {
  if (phase === "done") return "done";
  if (phase === "error") return stageIndex <= activeStep ? "done" : "pending";
  if (activeStep > stageIndex) return "done";
  if (activeStep === stageIndex) return "current";
  return "pending";
}

function renderReviewWorkflow(data = state.reviewData, options = {}) {
  if (!els.reviewWorkflow) return;
  const phase = options.phase || (state.reviewComplete ? "done" : data ? "done" : "idle");
  if (phase === "idle" || phase === "done") {
    els.reviewWorkflow.hidden = true;
    els.reviewWorkflow.innerHTML = "";
    return;
  }
  els.reviewWorkflow.hidden = false;
  const activeStep = Number(options.activeStep || 0);
  const currentStep = reviewSteps[activeStep % reviewSteps.length] || "正在审查合同，请稍候。";
  const stepLabels = ["读取原文", "主体核验", "大模型审查", "风险识别", "格式检查", "规则库匹配", "定位批注", "生成结果"];
  const percent = Math.min(96, Math.max(12, Math.round(((activeStep + 1) / stepLabels.length) * 100)));
  els.reviewWorkflow.className = `review-workflow stream ${phase}`;
  els.reviewWorkflow.innerHTML = `
    <div class="workflow-stream">
      ${stepLabels.map((label, index) => {
        const status = index < activeStep ? "done" : index === activeStep ? "current" : "pending";
        return `
          <div class="workflow-stream-node ${status}">
            <span>${status === "done" ? "✓" : ""}</span>
            <div>
              <strong>${escapeHtml(label)}</strong>
              ${status === "current" ? `<p>${escapeHtml(currentStep)}</p>` : ""}
            </div>
          </div>
        `;
      }).join("")}
      <small>当前进度 ${percent}% · 审查完成后自动隐藏流程。</small>
    </div>
  `;
  els.reviewWorkflow.scrollTop = els.reviewWorkflow.scrollHeight;
  return;
  const issues = data?.issues || state.issues || [];
  const profile = inferWorkflowProfile(state.workingText || els.reviewText.value || "", data || {});
  const workflow = data?.reviewWorkflow || {};
  const rightsText = cleanWorkflowText(workflow.rightsObligations || localRightsObligations(state.workingText || els.reviewText.value || "", profile), 360);
  const knowledgeBases = defaultKnowledgeBases(profile, data || {});
  const tasks = buildWorkflowTasks(data || {}, issues, activeStep, phase);
  const statusText = phase === "done" ? "审查完成，可查看风险意见" : phase === "error" ? "审查中断，请重新发起审查" : "审查中，审查完成后将短信通知";
  const ruleSelection = data?.ruleSelection || workflow.ruleSelection || "正在按合同类型匹配通用规则、行业预设和企业自定规则。";
  const resultStatus = phase === "done" ? "done" : phase === "error" ? "pending" : activeStep >= 6 ? "current" : "pending";

  els.reviewWorkflow.className = `review-workflow ${phase}`;
  els.reviewWorkflow.innerHTML = `
    <div class="workflow-status-bar">
      <div class="workflow-live">
        ${phase === "done" ? '<span class="workflow-check">✓</span>' : '<span class="workflow-dots"><i></i><i></i><i></i></span>'}
        <strong>${escapeHtml(statusText)}</strong>
      </div>
      <label class="workflow-switch">短信通知 <input type="checkbox" disabled><span></span></label>
    </div>
    <label class="workflow-auto"><input type="checkbox" checked disabled> 审查完成后自动生成修订建议</label>
    <div class="workflow-timeline">
      <section class="workflow-node ${workflowStatusClass(workflowStageStatus(0, activeStep, phase))}">
        <span class="workflow-icon"></span>
        <div>
          <h3>权利义务分析</h3>
          <p class="workflow-analysis">${escapeHtml(rightsText)}</p>
        </div>
      </section>
      <section class="workflow-node ${workflowStatusClass(workflowStageStatus(1, activeStep, phase))}">
        <span class="workflow-icon"></span>
        <div>
          <h3>知识库调用</h3>
          <ul class="workflow-knowledge">
            ${knowledgeBases.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}
          </ul>
          <small>${escapeHtml(ruleSelection)}</small>
        </div>
      </section>
      <section class="workflow-node ${workflowStatusClass(activeStep >= 2 || phase === "done" ? "current" : "pending")}">
        <span class="workflow-icon"></span>
        <div>
          <h3>合同审查</h3>
          <ul class="workflow-task-list">
            ${tasks.map((task) => `
              <li class="${workflowStatusClass(task.status)}">
                <span></span>
                <div><strong>${escapeHtml(task.label)}</strong><small>${escapeHtml(task.note || "")}</small></div>
              </li>
            `).join("")}
          </ul>
        </div>
      </section>
      <section class="workflow-node ${workflowStatusClass(resultStatus)}">
        <span class="workflow-icon"></span>
        <div>
          <h3>${phase === "done" ? "审查结果已生成" : "正在生成审查结果"}</h3>
          <p>${phase === "done" ? escapeHtml(data?.summary || "审查完成。") : "正在合并风险等级、原文定位、规则依据和可替换正文。"}</p>
        </div>
      </section>
    </div>
  `;
}

function filteredIssues() {
  return state.issues.filter((issue) => {
    if (state.filter === "summary") return false;
    if (state.filter === "risk") return issueMajorCategory(issue) === "risk";
    if (state.filter === "rule") return issueMajorCategory(issue) === "rule";
    if (state.filter === "common") return issueMajorCategory(issue) === "common";
    return true;
  });
}

function currentIssues() {
  return filteredIssues();
}

function currentIssue() {
  const issues = currentIssues();
  if (!issues.length) return null;
  state.activeIssueIndex = Math.max(0, Math.min(state.activeIssueIndex, issues.length - 1));
  return issues[state.activeIssueIndex];
}

function updateFilterButtons() {
  const counts = {
    summary: 1,
    risk: state.issues.filter((issue) => issueMajorCategory(issue) === "risk").length,
    common: state.issues.filter((issue) => issueMajorCategory(issue) === "common").length,
    rule: state.issues.filter((issue) => issueMajorCategory(issue) === "rule").length,
  };
  const labels = {
    summary: "审查结果",
    risk: "风险问题",
    common: "常见问题",
    rule: "规则库",
  };
  els.reviewFilterBar.querySelectorAll("[data-review-filter]").forEach((button) => {
    const key = button.dataset.reviewFilter;
    button.classList.toggle("active", key === state.filter);
    button.textContent = key === "summary" ? labels[key] : `${labels[key]} ${counts[key] || 0}`;
  });
}

function sourceEditorText() {
  if (officeEditor) return officeEditor.getText();
  const editor = document.getElementById("source-editor");
  return editor ? editor.innerText.trim() : state.workingText.trim();
}

function contractWordHtml(title = "合同当前版本", bodyHtml = "") {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>${escapeHtml(title)}</title>
  <style>
    @page { margin: 2.54cm 3.18cm; }
    body { margin: 0; color: #000; font-family: "SimSun", "宋体", serif; font-size: 12pt; line-height: 1.8; }
    .contract-doc { max-width: 780px; margin: 0 auto; }
    h1 { margin: 0 0 24pt; text-align: center; font-family: "SimHei", "黑体", sans-serif; font-size: 18pt; font-weight: 700; }
    h2 { margin: 18pt 0 8pt; font-family: "SimHei", "黑体", sans-serif; font-size: 14pt; font-weight: 700; }
    p { margin: 0 0 6pt; text-indent: 2em; }
    .clause-line, .list-line, .contract-blank { text-indent: 0; }
    table { width: 100%; border-collapse: collapse; }
    td, th { border: 1px solid #000; padding: 6pt; }
    mark, .risk-mark { background: transparent; color: inherit; }
  </style>
</head>
<body><div class="contract-doc">${bodyHtml}</div></body>
</html>`;
}

function setSourceMeta(text = state.workingText) {
  const label = state.fileName ? `${state.fileName}｜` : "";
  const markCount = state.reviewComplete ? `｜标记 ${state.issues.length} 处` : "";
  const preserve = state.templatePreserved ? "｜保真模板导出" : "";
  els.sourceMeta.textContent = text ? `${label}${text.length} 字${markCount}${preserve}` : "上传或粘贴合同后，原文会显示在这里。";
}

function ensureAnnotationLayer() {
  let layer = document.getElementById("source-annotation-layer");
  const container = officeEditor ? officeEditor.scrollWrap : els.reviewSource;
  if (!layer) {
    layer = document.createElement("div");
    layer.id = "source-annotation-layer";
    layer.className = "source-annotation-layer";
    container.appendChild(layer);
  } else if (layer.parentNode !== container) {
    container.appendChild(layer);
  }
  return layer;
}

function clearAnnotationLayer() {
  document.getElementById("source-annotation-layer")?.remove();
}

function textNodesUnder(root) {
  const nodes = [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      return node.nodeValue ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
    },
  });
  while (walker.nextNode()) nodes.push(walker.currentNode);
  return nodes;
}

function rangeForText(root, target) {
  const needle = String(target || "").trim();
  if (!needle) return null;
  const nodes = textNodesUnder(root);
  const joined = nodes.map((node) => node.nodeValue).join("");
  const index = joined.indexOf(needle);
  if (index < 0) return null;
  let offset = 0;
  let startNode = null;
  let endNode = null;
  let startOffset = 0;
  let endOffset = 0;
  for (const node of nodes) {
    const next = offset + node.nodeValue.length;
    if (!startNode && index <= next) {
      startNode = node;
      startOffset = Math.max(0, index - offset);
    }
    if (startNode && index + needle.length <= next) {
      endNode = node;
      endOffset = Math.max(0, index + needle.length - offset);
      break;
    }
    offset = next;
  }
  if (!startNode || !endNode) return null;
  const range = document.createRange();
  range.setStart(startNode, startOffset);
  range.setEnd(endNode, endOffset);
  return range;
}

function findSupplementAnchorTarget(text, issue) {
  const quote = String(issue.quote || "").trim();
  const quoteTarget = findNormalizedTarget(text, quote);
  if (quoteTarget) return quoteTarget;
  const key = issuePlacementKeyword(issue);
  if (key === "subject") return findSubjectTarget(text, issue);
  if (key === "payment") return findPaymentTarget(text, issue);
  if (key === "wording") return findWordingTarget(text, issue);
  if (key === "format" && /签署|签章|盖章|落款|签字|日期/.test(issueFullText(issue))) return "";
  return findKeywordTarget(text, issue) || exactQuoteTarget(text, issue);
}

function rangeForInsertion(root, issue) {
  const text = sourceEditorText();
  const anchor = findSupplementAnchorTarget(text, issue);
  const range = anchor ? rangeForText(root, anchor) : null;
  if (range) {
    range.collapse(false);
    return range;
  }
  const fallback = document.createRange();
  fallback.selectNodeContents(root);
  fallback.collapse(false);
  return fallback;
}

function renderAnnotationLayer() {
  const editor = officeEditor ? officeEditor.editor : document.getElementById("source-editor");
  if (!editor || !state.reviewComplete) {
    clearAnnotationLayer();
    return;
  }
  const layer = ensureAnnotationLayer();
  const scrollWrap = officeEditor ? officeEditor.scrollWrap : els.reviewSource;
  layer.innerHTML = "";
  layer.style.width = `${scrollWrap.scrollWidth}px`;
  layer.style.height = `${scrollWrap.scrollHeight}px`;
  const containerRect = scrollWrap.getBoundingClientRect();
  const text = sourceEditorText();
  state.issues.forEach((issue, index) => {
    const target = findIssueTarget(text, issue);
    const range = target ? rangeForText(editor, target) : rangeForInsertion(editor, issue);
    if (!range) return;
    const rects = Array.from(range.getClientRects()).filter((rect) => rect.width || rect.height);
    const issueIndex = index + 1;
    if (!rects.length) {
      const rect = range.getBoundingClientRect();
      rects.push(rect);
    }
    rects.slice(0, 6).forEach((rect, rectIndex) => {
      const marker = document.createElement("button");
      marker.type = "button";
      marker.className = `source-annotation-rect ${riskClass(issueLevel(issue))} ${issue.applied ? "resolved" : ""} ${target ? "" : "anchor"}`;
      marker.dataset.riskIndex = String(issueIndex);
      marker.title = issue.title || issueMajorLabel(issue);
      marker.style.left = `${rect.left - containerRect.left + scrollWrap.scrollLeft}px`;
      marker.style.top = `${rect.top - containerRect.top + scrollWrap.scrollTop}px`;
      marker.style.width = `${Math.max(rect.width, target ? 28 : 96)}px`;
      marker.style.height = `${Math.max(rect.height, 20)}px`;
      if (rectIndex === 0) marker.textContent = String(issueIndex);
      layer.appendChild(marker);
    });
  });
}

function contractTextToEditorHtml(text = "") {
  const lines = String(text || "").split(/\n/);
  let firstContent = true;
  return lines
    .map((line) => {
      const value = line.trimEnd();
      const compact = value.trim();
      if (!compact) return '<p class="contract-blank"><br></p>';
      if (firstContent && compact.length <= 60 && /合同|协议|承诺书|确认书|订单|补充协议/.test(compact)) {
        firstContent = false;
        return `<h1>${escapeHtml(compact)}</h1>`;
      }
      firstContent = false;
      if (/^第[一二三四五六七八九十百]+[章节篇]\s*/.test(compact) || /^[一二三四五六七八九十]+[、.．]\s*/.test(compact)) {
        return `<h2>${escapeHtml(compact)}</h2>`;
      }
      if (/^第[一二三四五六七八九十百]+条\s*/.test(compact) || /^\d+[、.．]\s*/.test(compact)) {
        return `<p class="clause-line">${escapeHtml(compact)}</p>`;
      }
      if (/^（[一二三四五六七八九十\d]+）/.test(compact) || /^\([一二三四五六七八九十\d]+\)/.test(compact)) {
        return `<p class="list-line">${escapeHtml(compact)}</p>`;
      }
      return `<p>${escapeHtml(compact)}</p>`;
    })
    .join("");
}

function renderEditableSource({ annotated = state.reviewComplete, replace = false } = {}) {
  const text = state.workingText || "";
  if (!text) {
    if (officeEditor) officeEditor.setHtml("");
    setSourceMeta("");
    clearAnnotationLayer();
    return;
  }
  if (officeEditor) {
    if (!officeEditor.getHtml().trim() || replace) {
      officeEditor.setHtml(officeEditor.textToContractHtml(text));
    }
  }
  setSourceMeta(text);
  window.requestAnimationFrame(() => {
    if (annotated) renderAnnotationLayer();
    else clearAnnotationLayer();
  });
}

function sourcePreview(file, text) {
  if (text) return `<div class="contract-source-editor muted online-office-editor" data-office-plugin="online-office" data-office-mode="word">${contractTextToEditorHtml(text.slice(0, 4000))}${text.length > 4000 ? "<p>...</p>" : ""}</div>`;
  if (file) return `<div class="notice">已上传：${escapeHtml(file.name)}，正在解析文件内容。</div>`;
  return `<div class="notice">上传或粘贴合同后，原文会显示在这里。</div>`;
}

function issueFullText(issue = {}) {
  return [
    issue.title,
    issue.problem,
    issue.suggestion,
    issue.category,
    issue.detail_category,
    issue.detailCategory,
    issue.source_rule,
    issue.sourceRule,
    issue.quote,
  ]
    .filter(Boolean)
    .join(" ");
}

function issueAmountMatches(value) {
  return String(value || "").match(/(?:人民币|RMB|￥|¥)?\s*\d[\d,]*(?:\.\d+)?\s*(?:元|万元|亿元|%|人民币)?|[零一二三四五六七八九十百千万亿]+(?:元|万元|亿元)/g) || [];
}

function quotedTerms(issue) {
  const terms = [];
  const text = issueFullText(issue);
  const reg = /[“"《「『']([^”"》」』']{1,50})[”"》」』']/g;
  let match;
  while ((match = reg.exec(text))) {
    const term = match[1].trim();
    if (term.length >= 2) terms.push(term);
  }
  return [...new Set(terms)];
}

function sentenceBounds(text, index) {
  const leftMarks = ["\n", "。", "；", ";", "！", "!", "？", "?"];
  const rightMarks = ["\n", "。", "；", ";", "！", "!", "？", "?"];
  let start = 0;
  for (const mark of leftMarks) {
    const found = text.lastIndexOf(mark, Math.max(0, index - 1));
    if (found >= start) start = found + mark.length;
  }
  let end = text.length;
  for (const mark of rightMarks) {
    const found = text.indexOf(mark, index);
    if (found >= 0) end = Math.min(end, found + mark.length);
  }
  while (start < end && /\s/.test(text[start])) start += 1;
  while (end > start && /\s/.test(text[end - 1])) end -= 1;
  return { start, end };
}

function sentenceAround(text, index) {
  if (index < 0) return "";
  const { start, end } = sentenceBounds(text, index);
  return text.slice(start, end).trim();
}

function locateCompact(value = "") {
  return String(value || "")
    .replace(/\s+/g, "")
    .replace(/[，。；;,.、：:【】\[\]（）()“”"'《》<>]/g, "");
}

function textSegments(text) {
  const segments = [];
  const reg = /[。！？；;.!?]\s*|\n+/g;
  let start = 0;
  let match;
  while ((match = reg.exec(text))) {
    const end = match.index + match[0].length;
    const raw = text.slice(start, end);
    const offset = raw.search(/\S/);
    const value = raw.trim();
    if (value) segments.push({ text: value, index: start + Math.max(offset, 0) });
    start = end;
  }
  const tail = text.slice(start);
  const offset = tail.search(/\S/);
  const value = tail.trim();
  if (value) segments.push({ text: value, index: start + Math.max(offset, 0) });
  return segments;
}

function findNormalizedTarget(text, target) {
  const raw = String(target || "").trim();
  if (!raw) return "";
  const exactIndex = text.indexOf(raw);
  if (exactIndex >= 0) return raw.length <= 220 ? raw : sentenceAround(text, exactIndex) || raw.slice(0, 220);
  const compactTarget = locateCompact(raw);
  if (compactTarget.length < 4) return "";
  const segments = textSegments(text);
  for (const segment of segments) {
    const compactSegment = locateCompact(segment.text);
    if (!compactSegment) continue;
    if (compactSegment.includes(compactTarget) || (compactTarget.includes(compactSegment) && compactSegment.length >= 6)) {
      return segment.text.slice(0, 260);
    }
  }
  const tokens = [...new Set((raw.match(/[\u4e00-\u9fa5]{2,8}|[A-Za-z0-9]{2,}/g) || []).filter((item) => !/^(合同|条款|问题|建议|明确|补充|应当|可以|需要|存在|未约定)$/.test(item)))].slice(0, 10);
  let best = { score: 0, text: "" };
  for (const segment of segments) {
    let score = 0;
    for (const token of tokens) {
      if (segment.text.includes(token)) score += token.length > 3 ? 2 : 1;
    }
    if (score > best.score) best = { score, text: segment.text };
  }
  return best.score >= Math.max(3, Math.ceil(tokens.length * 0.45)) ? best.text.slice(0, 260) : "";
}

function exactQuoteTarget(text, issue) {
  const targets = [issue.appliedText, issue.applied_text, issue.quote, issue.original_text, issue.originalText]
    .map((item) => String(item || "").trim())
    .filter((item) => item.length >= 2);
  for (const target of targets) {
    const normalized = findNormalizedTarget(text, target);
    if (normalized) return normalized;
  }
  return "";
}

function issuePlacementKeyword(issue) {
  const text = issueFullText(issue);
  if (/错别字|错字|字词|笔误|错写|用词|语病|表述错误|文字错误/.test(text)) return "wording";
  if (/主体|甲方|乙方|签约|名称|证照|授权|权限|用人单位|劳动者/.test(text)) return "subject";
  if (/付款|支付|价款|金额|费用|发票|结算|账期|租金|借款|利息/.test(text)) return "payment";
  if (/交付|交货|履行|服务|验收|整改|异议/.test(text)) return "delivery";
  if (/保密|秘密|隐私|个人信息|数据/.test(text)) return "confidentiality";
  if (/知识产权|著作权|专利|商标|源码|成果/.test(text)) return "ip";
  if (/违约|赔偿|责任|损失|解除|终止/.test(text)) return "liability";
  if (/争议|管辖|仲裁|诉讼/.test(text)) return "dispute";
  if (/签署|盖章|日期|附件|编号|标题|格式|排版/.test(text)) return "format";
  if (/风险|不合理|不可执行|无效|违规|违法|冲突|限制|免责|单方|过高|过低/.test(text)) return "risk";
  return "general";
}

function isSupplementIssue(issue) {
  if (issue.location_hint === "supplement" || issue.locationHint === "supplement") return true;
  const text = issueFullText(issue);
  if (/错别字|错字|字词|笔误|错写|用词|语病|表述错误|文字错误/.test(text)) return false;
  return /缺少|缺失|未约定|未明确|未载明|未包含|未写明|未体现|无.{0,8}约定|补充|新增|增加|应加入|应添加/.test(text);
}

function isDeleteIssue(issue) {
  const hint = String(issue.location_hint || issue.locationHint || issue.action || "").toLowerCase();
  if (hint === "delete" || hint === "remove") return true;
  const text = issueFullText(issue);
  return /删除|删去|移除|去除|不应保留|多余|重复|无效表述|最终解释权|连续空行|不必要空行/.test(text) && !replacementText(issue);
}

function isBlankLineIssue(issue) {
  return /连续空行|不必要空行|多余空行|空白行|排版空行/.test(issueFullText(issue));
}

function canApplyIssue(issue) {
  return isDeleteIssue(issue) || Boolean(replacementText(issue));
}

function issueApplyKind(issue) {
  if (isDeleteIssue(issue)) return "delete";
  const currentText = sourceEditorText();
  if (currentText && findIssueTarget(currentText, issue)) return "replace";
  return isSupplementIssue(issue) ? "insert" : "replace";
}

function issueActionLabel(issue) {
  const kind = issueApplyKind(issue);
  if (kind === "delete") return "删除";
  if (kind === "replace") return "修改";
  return "补充";
}

function findPaymentTarget(text, issue) {
  const quoteTarget = exactQuoteTarget(text, { quote: issue.quote || issue.original_text || issue.originalText });
  if (quoteTarget && /(付款|支付|价款|费用|合同金额|合同价款|结算|发票|款项|尾款|预付款|首付款|\d)/.test(quoteTarget)) {
    return quoteTarget;
  }
  const issueAmounts = [
    ...issueAmountMatches([issue.quote, issue.problem, issue.title].filter(Boolean).join(" ")),
    ...issueAmountMatches(issueFullText(issue)),
  ];
  for (const amount of issueAmounts) {
    const index = text.indexOf(amount);
    if (index >= 0) return sentenceAround(text, index) || amount;
  }
  const segments = textSegments(text);
  let best = { score: 0, text: "" };
  for (const segment of segments) {
    let score = 0;
    if (/(付款|支付|价款|费用|合同金额|合同价款|结算|发票|款项|尾款|预付款|首付款)/.test(segment.text)) score += 5;
    if (issueAmountMatches(segment.text).length) score += 6;
    if (/(节点|条件|期限|账期|逾期|比例|税费|开票)/.test(segment.text)) score += 2;
    if (score > best.score) best = { score, text: segment.text };
  }
  return best.score >= 6 ? best.text.slice(0, 260) : "";
}

function findSubjectTarget(text, issue) {
  const quoteTarget = exactQuoteTarget(text, { quote: issue.quote || issue.original_text || issue.originalText });
  if (quoteTarget && /甲方|乙方|主体|名称|住所|统一社会信用代码|法定代表人|联系人|证件/.test(quoteTarget)) return quoteTarget;
  const lines = [];
  let offset = 0;
  for (const raw of String(text || "").split(/\n/)) {
    const line = raw.trim();
    if (line && offset < 2500 && /甲方|乙方|主体|名称|住所|统一社会信用代码|法定代表人|联系人|联系电话|证件号码/.test(line)) {
      lines.push(line);
    }
    offset += raw.length + 1;
  }
  if (lines.length) return lines.slice(-1)[0].slice(0, 260);
  return findKeywordTarget(text, issue);
}

function findWordingTarget(text, issue) {
  for (const term of quotedTerms(issue)) {
    const index = text.indexOf(term);
    if (index >= 0) return sentenceAround(text, index) || term;
  }
  const exact = exactQuoteTarget(text, issue);
  if (exact) {
    const index = text.indexOf(exact);
    return sentenceAround(text, index) || exact;
  }
  return "";
}

function keywordTokens(issue) {
  const stop = /^(合同|条款|问题|建议|明确|约定|风险|审查|修改|补充|进行|存在|相关|可以|需要|应当|没有|未能|部分|内容)$/;
  return [...new Set((issueFullText(issue).match(/[\u4e00-\u9fa5]{2,8}/g) || []).filter((word) => !stop.test(word)).slice(0, 14))];
}

function findKeywordTarget(text, issue) {
  const key = issuePlacementKeyword(issue);
  const tokens = keywordTokens(issue);
  const segments = textSegments(text);
  let best = { score: 0, text: "" };
  for (const segment of segments) {
    let score = 0;
    for (const token of tokens) {
      if (segment.text.includes(token)) score += token.length > 3 ? 2 : 1;
    }
    if (key === "subject" && /甲方|乙方|签约主体|合同主体|名称|住所|统一社会信用代码|法定代表人/.test(segment.text)) score += 4;
    if (key === "liability" && /违约|赔偿|责任|损失|解除|终止/.test(segment.text)) score += 4;
    if (key === "delivery" && /交付|交货|履行|服务|验收|整改|异议/.test(segment.text)) score += 4;
    if (key === "confidentiality" && /保密|秘密|隐私|个人信息|数据/.test(segment.text)) score += 4;
    if (key === "ip" && /知识产权|著作权|专利|商标|源码|成果/.test(segment.text)) score += 4;
    if (key === "dispute" && /争议|管辖|仲裁|诉讼/.test(segment.text)) score += 4;
    if (key === "format" && /签署|盖章|日期|附件|编号|标题|格式|排版/.test(segment.text)) score += 4;
    if (score > best.score) best = { score, text: segment.text };
  }
  return best.score >= 3 ? best.text.slice(0, 260) : "";
}

function findIssueTarget(text, issue) {
  const applied = String(issue.appliedText || issue.applied_text || "").trim();
  if (applied && text.includes(applied)) return applied;
  if (isSupplementIssue(issue)) return findSupplementAnchorTarget(text, issue) || "";
  const key = issuePlacementKeyword(issue);
  if (key === "payment") return findPaymentTarget(text, issue) || exactQuoteTarget(text, issue);
  if (key === "wording") return findWordingTarget(text, issue) || exactQuoteTarget(text, issue);
  if (key === "subject") return findSubjectTarget(text, issue) || exactQuoteTarget(text, issue);
  const exact = exactQuoteTarget(text, issue);
  if (exact && exact.length <= 160) return exact;
  return findKeywordTarget(text, issue) || exact;
}

function findInsertionIndex(text, issue) {
  const key = issuePlacementKeyword(issue);
  if (key === "format" && /签署|签章|盖章|落款|签字|日期/.test(issueFullText(issue))) return text.length;
  if (isSupplementIssue(issue)) {
    const anchor = findSupplementAnchorTarget(text, issue);
    if (anchor) {
      const anchorIndex = text.indexOf(anchor);
      if (anchorIndex >= 0) return anchorIndex + anchor.length;
    }
  }
  const nearbyTarget =
    key === "payment"
      ? findPaymentTarget(text, issue)
      : key === "wording"
        ? findWordingTarget(text, issue)
        : findKeywordTarget(text, issue) || exactQuoteTarget(text, issue);
  if (nearbyTarget) {
    const index = text.indexOf(nearbyTarget);
    if (index >= 0) return key === "subject" ? index : index + nearbyTarget.length;
  }
  const patterns = {
    subject: [/甲方[：:]/, /乙方[：:]/, /签约主体/, /合同主体/, /主体/],
    payment: [/价款/, /付款/, /支付/, /费用/, /结算/],
    delivery: [/交付/, /履行/, /服务/, /验收/],
    confidentiality: [/保密/, /秘密/, /数据安全/, /个人信息/],
    ip: [/知识产权/, /成果归属/, /著作权/, /源码/],
    liability: [/违约/, /赔偿/, /责任/, /解除/, /终止/],
    dispute: [/争议/, /管辖/, /仲裁/, /诉讼/],
    format: [/签署/, /盖章/, /附件/, /日期/],
    general: [/一、/, /第一条/, /1[.、]/],
  };
  for (const pattern of patterns[key] || patterns.general) {
    const match = pattern.exec(text);
    if (!match) continue;
    if (key === "subject") return match.index;
    const { end } = sentenceBounds(text, match.index);
    return end || match.index;
  }
  const firstBreak = text.indexOf("\n\n");
  return firstBreak > -1 ? firstBreak : text.length;
}

function insertAnchorTokens(text, issues) {
  const insertions = [];
  const matchedTargets = new Set();
  issues.forEach((issue, index) => {
    const target = findIssueTarget(text, issue);
    const targetKey = target.slice(0, 120);
    if (target.length >= 6 && text.includes(target) && !matchedTargets.has(targetKey)) {
      matchedTargets.add(targetKey);
      return;
    }
    insertions.push({
      index: findInsertionIndex(text, issue),
      token: `[[__RISK_ANCHOR_${index + 1}__]]`,
    });
  });
  return insertions
    .sort((a, b) => b.index - a.index)
    .reduce((result, item) => `${result.slice(0, item.index)}\n${item.token}\n${result.slice(item.index)}`, text);
}

function issueMarkHtml(issue, index, content, isAnchor = false) {
  const appliedClass = issue.applied ? "resolved" : "";
  const anchorClass = isAnchor ? "anchor" : "";
  return `<mark class="risk-mark ${riskClass(issueLevel(issue))} ${appliedClass} ${anchorClass}" data-risk-index="${index + 1}">${content}</mark>`;
}

function highlightedSource(text, issues) {
  const textWithAnchors = insertAnchorTokens(text || "", issues);
  let html = escapeHtml(textWithAnchors);
  const seen = new Set();
  const snippets = issues
    .map((issue, index) => ({ issue, index, quote: findIssueTarget(text || "", issue) }))
    .sort((a, b) => b.quote.length - a.quote.length)
    .filter((item) => item.quote.length >= 6)
    .filter((item) => {
      const key = item.quote.slice(0, 120);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

  for (const item of snippets) {
    const escapedQuote = escapeHtml(item.quote.slice(0, 260));
    if (!escapedQuote || !html.includes(escapedQuote)) continue;
    html = html.replace(escapedQuote, issueMarkHtml(item.issue, item.index, escapedQuote));
  }
  issues.forEach((issue, index) => {
    const token = `[[__RISK_ANCHOR_${index + 1}__]]`;
    const label = issue.applied
      ? escapeHtml(issue.appliedText || replacementText(issue) || issue.title || "已补充")
      : `需补充条款：${escapeHtml(issue.title || issueMajorLabel(issue))}`;
    html = html.replace(token, issueMarkHtml(issue, index, label, true));
  });
  return html.replace(/\n/g, "<br>");
}

function startReviewProgress(file, text) {
  let step = 0;
  const render = () => {
    const activeStep = Math.min(step, reviewSteps.length - 1);
    renderReviewWorkflow({
      summary: "正在按权利义务、知识库、实质风险、文字符号和主体信息分层审查。",
      overallRisk: "审查中",
      reviewEngine: "大模型深度审查 + 本地规则复核",
      model: state.model,
      modelProvider: state.modelProvider,
    }, { phase: "running", activeStep });
    if (step === 0) {
      els.reviewIssues.innerHTML = `<div class="notice review-waiting">正在审查，完成后这里会直接展示风险点和修改意见。</div>`;
    }
    step = Math.min(step + 1, reviewSteps.length - 1);
  };
  window.clearInterval(reviewProgressTimer);
  els.reviewFilterBar.hidden = true;
  els.issueNav.hidden = true;
  if (els.acceptAllBtn) els.acceptAllBtn.disabled = true;
  clearAnnotationLayer();
  els.sourceMeta.textContent = file ? `正在解析：${file.name}` : `已读取 ${text.length} 个字符`;
  updateScoreboard({ overallRisk: "审查中", model: state.model, modelProvider: state.modelProvider }, []);
  renderReviewSummary({
    summary: "正在调用大模型深度审查，并同步匹配规则库、常见问题和格式风险，请稍候。",
    overallRisk: "审查中",
    reviewEngine: "大模型深度审查 + 本地规则复核",
    model: state.model,
    modelProvider: state.modelProvider,
  }, []);
  render();
  reviewProgressTimer = window.setInterval(render, 2400);
}

function stopReviewProgress() {
  window.clearInterval(reviewProgressTimer);
  reviewProgressTimer = null;
}

function renderReviewResult(data) {
  const issues = (data.issues || [])
    .map(normalizeIssueForReview)
    .slice()
    .sort((a, b) => riskRank(b.risk_level || b.riskLevel) - riskRank(a.risk_level || a.riskLevel));
  data.issues = issues;
  state.issues = issues;
  state.reviewData = data;
  state.reviewComplete = true;
  state.filter = "summary";
  updateScoreboard(data, issues);
  renderReviewSummary(data, issues);
  renderReviewWorkflow(data, { phase: "done" });
  updateFilterButtons();
  const contractText = data.contractText || state.workingText || els.reviewText.value.trim();
  const currentText = sourceEditorText() || contractText;
  state.originalText = currentText;
  state.workingText = currentText;
  els.reviewText.value = currentText;
  state.activeIssueIndex = 0;
  renderEditableSource({ annotated: true });
  renderIssueList();
  els.reviewBtn.textContent = "重新审查";
  if (els.exportReportBtn) els.exportReportBtn.disabled = false;
  saveState();
}

function renderIssueList() {
  if (!state.reviewData) {
    renderReviewWorkflow(null, { phase: state.workingText ? "idle" : "idle" });
    els.reviewFilterBar.hidden = true;
    els.issueNav.hidden = true;
    if (els.acceptAllBtn) els.acceptAllBtn.disabled = true;
    els.reviewIssues.innerHTML = `<div class="notice">审查完成后，右侧会按“审查结果、风险问题、常见问题、规则库”展示。</div>`;
    els.issueCounter.textContent = "0 / 0";
    els.prevIssueBtn.disabled = true;
    els.nextIssueBtn.disabled = true;
    return;
  }
  els.reviewFilterBar.hidden = false;
  els.issueNav.hidden = true;
  if (els.acceptAllBtn) {
    els.acceptAllBtn.disabled = !state.issues.some((issue) => !issue.applied && canApplyIssue(issue));
  }
  if (state.filter === "summary") {
    const data = state.reviewData || {};
    const counts = countsByRisk(state.issues);
    const dims = data.reviewReport?.counts || dimensionCounts(state.issues);
    const risk = data.overallRisk || data.overall_risk || (counts.high ? "高" : counts.mid ? "中" : counts.total ? "低" : "待审查");
    const score = ["待审查", "审查中"].includes(risk) ? "--" : reviewScore(state.issues);
    els.reviewIssues.innerHTML = `
      <article class="review-result-tab">
        <div class="result-score-line">
          <span class="tag ${riskClass(risk)}">整体风险：${escapeHtml(risk)}</span>
          <span class="tag neutral">评分：${score}</span>
          <span class="tag neutral">问题：${counts.total}</span>
        </div>
        <p>${escapeHtml(data.summary || "审查完成。")}</p>
        <small>内容 ${dims.content || 0} / 主体 ${dims.subject || 0} / 格式 ${dims.format || 0} / 规则 ${dims.rule || 0}；高 ${counts.high} / 中 ${counts.mid}。</small>
      </article>
      <div class="issue-list compact">
        ${state.issues.slice(0, 8).map((issue, index) => issueCardHtml(issue, state.issues.indexOf(issue) + 1, index)).join("")}
      </div>
    `;
    saveState();
    return;
  }
  const visibleIssues = filteredIssues();
  if (!visibleIssues.length) {
    els.reviewIssues.innerHTML = `<div class="notice">当前分类下暂无问题。</div>`;
    saveState();
    return;
  }
  els.reviewIssues.innerHTML = `
    <div class="issue-list compact">
      ${visibleIssues.map((issue, index) => issueCardHtml(issue, state.issues.indexOf(issue) + 1, index)).join("")}
    </div>
  `;
  saveState();
}

function issueCardHtml(issue, originalIndex, visibleIndex = 0) {
  const replacement = replacementText(issue);
  const basisItems = issueBasisItems(issue);
  const deleteTarget = isDeleteIssue(issue)
    ? (findIssueTarget(sourceEditorText(), issue) || issue.quote || issue.original_text || issue.originalText || "")
    : "";
  return `
    <article class="issue-card compact-card ${riskClass(issueLevel(issue))} ${issue.applied ? "resolved" : ""}" data-issue-index="${originalIndex}" data-visible-index="${visibleIndex}">
      <div class="issue-title">
        <strong><b>${originalIndex}</b>${escapeHtml(issue.title || "风险问题")}</strong>
        <span class="tag ${riskClass(issueLevel(issue))}">${escapeHtml(issueLevel(issue))}风险</span>
        <span class="tag detail">${escapeHtml(issueMajorLabel(issue))}</span>
      </div>
      <p>${escapeHtml(issue.problem || "")}</p>
      ${issue.suggestion ? `<small>${escapeHtml(issue.suggestion)}</small>` : ""}
      ${basisItems.length ? `<div class="issue-basis">${basisItems.map((item) => `<span>${escapeHtml(item)}</span>`).join("")}</div>` : ""}
      ${replacement ? `<div class="replacement-preview">${escapeHtml(replacement)}</div>` : ""}
      ${!replacement && deleteTarget ? `<div class="replacement-preview delete-preview">删除原文：${escapeHtml(String(deleteTarget).slice(0, 220))}</div>` : ""}
      <div class="actions compact-actions">
        <button class="secondary mini" type="button" data-locate-issue>定位原文</button>
        ${
          issue.applied
            ? '<button class="secondary mini" type="button" data-undo-suggestion>撤销</button>'
            : canApplyIssue(issue)
              ? `<button class="primary mini" type="button" data-apply-suggestion>${issueActionLabel(issue)}</button>`
              : ""
        }
      </div>
    </article>
  `;
}

function replacementText(issue) {
  return String(issue?.replacement_text || issue?.replacementText || issue?.suggested_text || "").trim();
}

function escapeRegExp(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function insertSuggestionText(text, issue, replacement) {
  const insertionIndex = findInsertionIndex(text, issue);
  const prefix = text.slice(0, insertionIndex).replace(/\s+$/, "");
  const suffix = text.slice(insertionIndex).replace(/^\s+/, "");
  return `${prefix}\n\n${replacement}\n\n${suffix}`.trim();
}

function locateIssue(animate = true) {
  const issue = currentIssue();
  locateSpecificIssue(issue, animate);
}

function locateSpecificIssue(issue, animate = true) {
  if (!issue) return;
  const originalIndex = state.issues.indexOf(issue) + 1;
  renderAnnotationLayer();
  const mark = els.reviewSource.querySelector(`#source-annotation-layer [data-risk-index="${originalIndex}"]`);
  if (!mark) return;
  mark.scrollIntoView({ behavior: animate ? "smooth" : "auto", block: "center" });
  mark.classList.add("focus");
  window.setTimeout(() => mark.classList.remove("focus"), 1400);
}

function replaceRangeWithText(range, replacement) {
  range.deleteContents();
  range.insertNode(document.createTextNode(replacement));
}

function insertSuggestionInEditor(editor, issue, replacement) {
  const range = rangeForInsertion(editor, issue);
  if (!range) return false;
  range.insertNode(document.createTextNode(`\n\n${replacement}\n\n`));
  return true;
}

function normalizeEditorSpacing(editor) {
  if (!editor) return false;
  const before = sourceEditorText();
  const cleaned = before
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  if (!cleaned || cleaned === before.trim()) return false;
  state.workingText = cleaned;
  els.reviewText.value = cleaned;
  renderEditableSource({ annotated: true, replace: true });
  return true;
}

function applyIssueSuggestion(issue, options = {}) {
  if (!issue) return false;
  if (issue.applied) return false;
  const quote = String(issue.quote || "").trim();
  const replacement = replacementText(issue);
  if (!canApplyIssue(issue)) return false;
  const editor = officeEditor ? officeEditor.editor : document.getElementById("source-editor");
  if (!editor) return false;
  const beforeText = sourceEditorText();
  let undoTarget = "";
  let appliedText = replacement;
  let applyMode = issueApplyKind(issue);
  if (applyMode === "delete") {
    const target = findIssueTarget(beforeText, issue) || quote;
    const range = target ? rangeForText(editor, target) : null;
    if (range) {
      range.deleteContents();
      undoTarget = target;
      appliedText = `已删除：${target}`;
    } else if (isBlankLineIssue(issue) && normalizeEditorSpacing(editor)) {
      appliedText = "已整理多余空行";
      applyMode = "delete-format";
    } else {
      return false;
    }
  } else if (applyMode === "insert") {
    if (!replacement || !insertSuggestionInEditor(editor, issue, replacement)) return false;
  } else {
    const current = beforeText;
    const target = findIssueTarget(current, issue) || quote;
    const range = target ? rangeForText(editor, target) : null;
    if (range) {
      replaceRangeWithText(range, replacement);
      undoTarget = target;
      applyMode = "replace";
    } else if (replacement && insertSuggestionInEditor(editor, issue, replacement)) {
      applyMode = "insert";
    } else {
      return false;
    }
  }
  state.workingText = sourceEditorText();
  issue.applied = true;
  issue.appliedText = appliedText;
  issue.undoText = beforeText;
  issue.undoTarget = undoTarget;
  issue.applyMode = applyMode;
  issue.appliedAt = new Date().toISOString();
  els.reviewText.value = state.workingText;
  renderAnnotationLayer();
  if (options.locate !== false) locateSpecificIssue(issue, false);
  els.reviewBtn.textContent = "重新审查";
  saveState();
  return true;
}

function applyCurrentSuggestion() {
  return applyIssueSuggestion(currentIssue());
}

function undoIssueSuggestion(issue, options = {}) {
  if (!issue || !issue.applied || !issue.undoText) return false;
  const current = sourceEditorText();
  const applied = String(issue.appliedText || replacementText(issue) || "").trim();
  const original = String(issue.undoTarget || "").trim();
  let nextText = "";
  if (issue.applyMode === "replace" && applied && original && current.includes(applied)) {
    nextText = current.replace(applied, original);
  } else if (issue.applyMode === "insert" && applied && current.includes(applied)) {
    const blockPattern = new RegExp(`\\n{0,2}${escapeRegExp(applied)}\\n{0,2}`);
    nextText = current.replace(blockPattern, "\n\n").replace(/\n{3,}/g, "\n\n").trim();
  } else {
    nextText = String(issue.undoText || "");
  }
  state.workingText = nextText;
  els.reviewText.value = state.workingText;
  issue.applied = false;
  issue.appliedText = "";
  issue.undoText = "";
  issue.undoTarget = "";
  issue.applyMode = "";
  issue.appliedAt = "";
  renderEditableSource({ annotated: true, replace: true });
  if (options.locate !== false) locateSpecificIssue(issue, false);
  saveState();
  return true;
}

function undoCurrentSuggestion() {
  return undoIssueSuggestion(currentIssue());
}

function applyAllSuggestions() {
  const targets = state.issues.filter((issue) => !issue.applied && canApplyIssue(issue));
  let appliedCount = 0;
  targets.forEach((issue) => {
    if (applyIssueSuggestion(issue, { locate: false })) appliedCount += 1;
  });
  renderEditableSource({ annotated: true, replace: true });
  renderIssueList();
  if (els.acceptAllBtn) {
    const originalText = els.acceptAllBtn.textContent;
    els.acceptAllBtn.textContent = appliedCount ? `已接受 ${appliedCount} 条` : "无可接受项";
    window.setTimeout(() => {
      els.acceptAllBtn.textContent = originalText;
      els.acceptAllBtn.disabled = !state.issues.some((issue) => !issue.applied && canApplyIssue(issue));
    }, 1400);
  }
}

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

async function downloadCorrectedText() {
  const text = sourceEditorText();
  if (!text) return;
  if (officeEditor) {
    const title = "合同当前版本";
    officeEditor.downloadDoc(title);
    return;
  }
  if (state.templatePreserved && state.templateDocxId) {
    const originalText = els.downloadCorrectedBtn.textContent;
    els.downloadCorrectedBtn.disabled = true;
    els.downloadCorrectedBtn.textContent = "保真导出中...";
    try {
      const res = await fetch(`${API_BASE}/api/contracts/export-docx`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          templateDocxId: state.templateDocxId,
          currentText: text,
          fileName: state.fileName || "合同当前版本.docx",
        }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || "保真导出失败");
      }
      const blob = await res.blob();
      const base = (state.fileName || "合同当前版本.docx").replace(/\.docx$/i, "").replace(/[\\/:*?"<>|]/g, "_");
      downloadBlob(blob, `${base}_修改版.docx`);
      return;
    } catch (error) {
      els.sourceMeta.textContent = `${error.message}，已回退为普通 Word 下载`;
    } finally {
      els.downloadCorrectedBtn.disabled = false;
      els.downloadCorrectedBtn.textContent = originalText;
    }
  }
  const editor = document.getElementById("source-editor");
  const html = contractWordHtml("合同当前版本", editor ? editor.innerHTML : contractTextToEditorHtml(text));
  const blob = new Blob([html], { type: "application/msword;charset=utf-8" });
  downloadBlob(blob, `合同当前版本_${new Date().toISOString().slice(0, 10)}.doc`);
}

async function copyCorrectedText() {
  const text = sourceEditorText();
  if (!text) return;
  if (officeEditor) {
    await officeEditor.copyContent();
  } else if (navigator.clipboard) {
    await navigator.clipboard.writeText(text);
  } else {
    window.prompt("复制当前原文", text);
  }
  els.copyCorrectedBtn.textContent = "已复制";
  window.setTimeout(() => {
    els.copyCorrectedBtn.textContent = "复制当前原文";
  }, 1400);
}

async function reviewContract() {
  state.workingText = sourceEditorText() || els.reviewText.value.trim();
  const text = state.workingText.trim();
  if (!text) {
    els.reviewIssues.innerHTML = `<div class="notice">请上传合同文件或粘贴合同文本。</div>`;
    return;
  }
  state.filter = "summary";
  state.activeIssueIndex = 0;
  updateFilterButtons();
  const originalText = els.reviewBtn.textContent;
  els.reviewBtn.disabled = true;
  els.reviewBtn.classList.add("loading");
  els.reviewBtn.textContent = "审查中...";
  startReviewProgress(state.fileName ? { name: state.fileName } : null, text);
  try {
    const data = await api("/api/contracts/review", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ contractText: text }),
    });
    if (state.fileName && !data.fileName) data.fileName = state.fileName;
    renderReviewResult(data);
  } catch (error) {
    els.reviewIssues.innerHTML = `<div class="notice">${escapeHtml(error.message)}</div>`;
    renderReviewSummary({ summary: error.message, overallRisk: "审查失败" }, []);
    renderReviewWorkflow({ summary: error.message, overallRisk: "审查失败" }, { phase: "error", activeStep: 3 });
  } finally {
    stopReviewProgress();
    els.reviewBtn.disabled = false;
    els.reviewBtn.classList.remove("loading");
    els.reviewBtn.textContent = state.reviewComplete ? "重新审查" : originalText;
  }
}

function openImportModal() {
  els.reviewFile.value = "";
  els.reviewImportText.value = "";
  els.reviewFileHint.textContent = "支持 Word、PDF、TXT";
  els.importModal.hidden = false;
  window.setTimeout(() => els.reviewFile.focus(), 0);
}

function closeImportModal() {
  els.importModal.hidden = true;
}

function resetReviewState(text = "", fileName = "", sourceHtml = "", options = {}) {
  state.workingText = text.trim();
  state.originalText = state.workingText;
  state.fileName = fileName;
  state.templateDocxId = options.templateDocxId || "";
  state.templatePreserved = Boolean(options.templatePreserved && state.templateDocxId);
  state.reviewComplete = false;
  state.reviewData = null;
  state.issues = [];
  state.filter = "summary";
  state.activeIssueIndex = 0;
  els.reviewText.value = state.workingText;
  updateScoreboard({}, []);
  renderReviewSummary(null, []);
  renderReviewWorkflow(null, { phase: "idle" });
  updateFilterButtons();
  if (officeEditor && sourceHtml) {
    clearAnnotationLayer();
    officeEditor.setHtml(sourceHtml);
    setSourceMeta(state.workingText);
  } else {
    renderEditableSource({ annotated: false, replace: true });
  }
  renderIssueList();
  els.reviewBtn.textContent = "开始审查";
  els.openImportBtn.classList.toggle("file-loaded", Boolean(state.workingText));
  els.openImportBtn.textContent = fileName ? `文件：${fileName}` : state.workingText ? "已导入文本合同" : "上传/粘贴合同";
  if (state.templatePreserved) els.sourceMeta.textContent = `${els.sourceMeta.textContent}｜保真模板导出`;
  saveState();
}

async function importContract() {
  const file = els.reviewFile.files[0];
  const text = els.reviewImportText.value.trim();
  if (!file && !text) {
    els.reviewFileHint.textContent = "请先选择文件或粘贴合同文本";
    return;
  }
  const originalText = els.confirmImportBtn.textContent;
  els.confirmImportBtn.disabled = true;
  els.confirmImportBtn.classList.add("loading");
  els.confirmImportBtn.textContent = "导入中...";
  try {
    let data;
    if (file) {
      const formData = new FormData();
      formData.append("contractFile", file);
      formData.append("contractText", text);
      data = await api("/api/contracts/parse", { method: "POST", body: formData });
    } else {
      data = { contractText: text, fileName: "" };
    }
    resetReviewState(data.contractText || text, data.fileName || file?.name || "", data.contractHtml || "", {
      templateDocxId: data.templateDocxId || "",
      templatePreserved: Boolean(data.templatePreserved),
    });
    closeImportModal();
  } catch (error) {
    els.reviewFileHint.textContent = error.message;
  } finally {
    els.confirmImportBtn.disabled = false;
    els.confirmImportBtn.classList.remove("loading");
    els.confirmImportBtn.textContent = originalText;
  }
}

function clearImport() {
  els.reviewFile.value = "";
  els.reviewImportText.value = "";
  els.reviewFileHint.textContent = "支持 Word、PDF、TXT";
}


/* ── 导出审查报告 ── */
function exportReviewReport() {
  if (!state.reviewData || !state.issues.length) return;
  const data = state.reviewData;
  const issues = state.issues;
  const counts = countsByRisk(issues);
  const dims = data.reviewReport?.counts || dimensionCounts(issues);
  const risk = data.overallRisk || data.overall_risk || "中";
  const score = reviewScore(issues);
  const riskColor = risk === "高" ? "#ff4d4f" : risk === "中" ? "#faad14" : "#52c41a";
  const now = new Date();
  const dateStr = now.getFullYear() + "-" + String(now.getMonth()+1).padStart(2,"0") + "-" + String(now.getDate()).padStart(2,"0");
  const modelLabel = currentModelLabel(data);

  const issueRows = issues.map(function(issue, i) {
    const level = issueLevel(issue);
    const lc = level === "高" ? "#ff4d4f" : level === "中" ? "#faad14" : "#52c41a";
    const basis = issueBasisItems(issue);
    const replacement = replacementText(issue);
    return '<tr><td style="text-align:center;font-weight:700;">' + (i+1) + '</td>'
      + '<td><span style="display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;color:#fff;background:' + lc + ';">' + level + '风险</span></td>'
      + '<td style="font-weight:600;">' + escapeHtml(issue.title || "风险问题") + '</td>'
      + '<td>' + escapeHtml(issue.problem || "") + '</td>'
      + '<td>' + escapeHtml(issue.suggestion || "—") + '</td>'
      + '<td style="font-size:11px;color:#666;">' + (basis.length ? basis.map(function(b){return escapeHtml(b);}).join("<br>") : "—") + '</td>'
      + '<td style="font-size:11px;">' + (replacement ? '<span style="color:#52c41a;">' + escapeHtml(replacement.slice(0,120)) + '</span>' : "—") + '</td></tr>';
  }).join("");

  const dimData = [
    { label: "实质内容", count: dims.content || 0 },
    { label: "签约主体", count: dims.subject || 0 },
    { label: "格式用语", count: dims.format || 0 },
    { label: "规则匹配", count: dims.rule || 0 },
  ];
  const dimRows = dimData.map(function(d) {
    var pct = Math.min(100, d.count * (100 / Math.max(1, counts.total)));
    return '<tr><td style="font-weight:600;">' + d.label + '</td><td style="text-align:center;">' + d.count + '</td>'
      + '<td><div style="height:12px;border-radius:6px;background:#f0f0f0;"><div style="height:12px;border-radius:6px;background:#1677ff;width:' + pct + '%"></div></div></td></tr>';
  }).join("");

  var replacementIssues = issues.filter(function(i){return replacementText(i);});
  var replRows = replacementIssues.map(function(issue, i) {
    var lc = issueLevel(issue)==="高"?"#ff4d4f":issueLevel(issue)==="中"?"#faad14":"#52c41a";
    return '<tr><td style="text-align:center;">' + (i+1) + '</td>'
      + '<td><span style="color:' + lc + ';font-weight:600;">' + issueLevel(issue) + '</span></td>'
      + '<td>' + escapeHtml(issue.title || "") + '</td>'
      + '<td style="font-size:10pt;">' + escapeHtml(replacementText(issue)) + '</td>'
      + '<td style="text-align:center;">' + (issue.applied ? "已应用" : "待处理") + '</td></tr>';
  }).join("");

  var reportHtml = '<!DOCTYPE html>\n<html><head><meta charset="utf-8">\n<title>合同审查报告</title>\n'
    + '<style>@page{margin:2cm 2.5cm;size:A4;}body{margin:0;font-family:"Microsoft YaHei","微软雅黑","SimSun",sans-serif;font-size:11pt;color:#333;line-height:1.7;}\n'
    + '.cover{text-align:center;padding:120px 40px 60px;page-break-after:always;}\n'
    + '.cover h1{font-size:28pt;color:#1a1a2e;margin:0 0 12px;}\n'
    + '.cover .subtitle{font-size:14pt;color:#666;margin:0 0 40px;}\n'
    + '.cover .meta-table{margin:0 auto;border-collapse:collapse;font-size:11pt;}\n'
    + '.cover .meta-table td{padding:8px 24px;text-align:left;}\n'
    + '.cover .meta-table td:first-child{color:#999;text-align:right;}\n'
    + '.cover .logo{font-size:18pt;font-weight:800;color:#1677ff;margin-bottom:40px;}\n'
    + 'h2{font-size:16pt;color:#1a1a2e;border-bottom:2px solid #1677ff;padding-bottom:6px;margin:32px 0 16px;}\n'
    + '.summary-card{background:#f6f8fa;border:1px solid #e8e8e8;border-radius:8px;padding:20px 24px;margin:16px 0;}\n'
    + '.score-grid{display:flex;gap:24px;margin:16px 0;}\n'
    + '.score-item{flex:1;text-align:center;padding:16px;border-radius:8px;border:1px solid #e8e8e8;}\n'
    + '.score-item .num{font-size:28pt;font-weight:800;}\n'
    + '.score-item .label{font-size:10pt;color:#999;}\n'
    + 'table{width:100%;border-collapse:collapse;margin:12px 0;font-size:10pt;}\n'
    + 'th{background:#f5f5f5;font-weight:600;}th,td{border:1px solid #e8e8e8;padding:8px 10px;text-align:left;vertical-align:top;}\n'
    + '.footer{text-align:center;color:#999;font-size:9pt;margin-top:40px;padding-top:16px;border-top:1px solid #e8e8e8;}\n</style>\n</head><body>\n\n'
    + '<div class="cover">\n<div class="logo">ContractAI</div>\n<h1>合同审查报告</h1>\n'
    + '<p class="subtitle">' + escapeHtml(state.fileName || "智能合同审查") + '</p>\n'
    + '<table class="meta-table">\n'
    + '<tr><td>审查日期</td><td>' + dateStr + '</td></tr>\n'
    + '<tr><td>综合风险</td><td style="color:' + riskColor + ';font-weight:700;">' + escapeHtml(risk) + '</td></tr>\n'
    + '<tr><td>合同评分</td><td style="font-weight:700;">' + score + ' / 100</td></tr>\n'
    + '<tr><td>风险问题</td><td>' + counts.total + ' 项（高 ' + counts.high + ' / 中 ' + counts.mid + ' / 低 ' + counts.low + '）</td></tr>\n'
    + '<tr><td>审查引擎</td><td>' + escapeHtml(modelLabel) + '</td></tr>\n'
    + '</table>\n</div>\n\n'
    + '<h2>一、审查摘要</h2>\n'
    + '<div class="summary-card"><p>' + escapeHtml(data.summary || "审查完成。") + '</p>\n'
    + '<div class="score-grid">\n'
    + '<div class="score-item"><div class="num" style="color:' + riskColor + ';">' + score + '</div><div class="label">合同评分</div></div>\n'
    + '<div class="score-item"><div class="num" style="color:' + riskColor + ';">' + escapeHtml(risk) + '</div><div class="label">综合风险</div></div>\n'
    + '<div class="score-item"><div class="num" style="color:' + (counts.high ? "#ff4d4f" : "#52c41a") + ';">' + (counts.high + counts.mid) + '</div><div class="label">风险问题</div></div>\n'
    + '<div class="score-item"><div class="num">' + counts.total + '</div><div class="label">审查要点</div></div>\n'
    + '</div></div>\n\n'
    + '<h2>二、维度分析</h2>\n'
    + '<table><thead><tr><th>审查维度</th><th style="width:80px;text-align:center;">问题数</th><th>分布</th></tr></thead>\n<tbody>' + dimRows + '</tbody></table>\n\n'
    + '<h2>三、风险问题逐条分析</h2>\n'
    + '<table><thead><tr><th style="width:40px;text-align:center;">#</th><th style="width:70px;">风险等级</th><th style="width:140px;">问题标题</th><th>问题描述</th><th>修改建议</th><th style="width:140px;">依据</th><th style="width:140px;">替换文本</th></tr></thead>\n<tbody>' + issueRows + '</tbody></table>\n\n'
    + '<h2>四、修改建议汇总</h2>\n'
    + (replacementIssues.length
      ? '<table><thead><tr><th style="width:40px;text-align:center;">#</th><th style="width:70px;">风险等级</th><th style="width:160px;">问题</th><th>建议替换文本</th><th style="width:60px;">状态</th></tr></thead>\n<tbody>' + replRows + '</tbody></table>\n'
      : '<p>暂无可直接替换的修改建议。</p>\n')
    + '<div class="footer"><p>本报告由 ContractAI 智能合同管理平台自动生成 · ' + dateStr + '</p>\n'
    + '<p>报告仅供参考，具体法律意见请咨询专业律师</p></div>\n\n</body></html>';

  var blob = new Blob([reportHtml], { type: "application/msword;charset=utf-8" });
  var base = (state.fileName || "合同").replace(/\.[^.]+$/, "").replace(/[\\/:*?"<>|]/g, "_");
  downloadBlob(blob, base + "_审查报告_" + dateStr + ".doc");
}

async function loadInitial() {
  officeEditor = new OfficeEditor("review-editor-container", {
    mode: "word",
    placeholder: "上传或粘贴合同后，原文会显示在这里并可直接编辑。",
    contractMode: true,
    onChange() {
      state.workingText = sourceEditorText();
      els.reviewText.value = state.workingText;
      setSourceMeta(state.workingText);
      if (state.reviewComplete) els.reviewBtn.textContent = "重新审查";
      saveState();
    },
  });
  officeEditor.scrollWrap.addEventListener("click", handleAnnotationClick);

  const health = await api("/api/health");
  const providerLabel = health.modelProvider && health.modelProvider !== "None" ? `${health.modelProvider} / ` : "";
  state.model = health.model || "";
  state.modelProvider = health.modelProvider || "";
  state.hasApiKey = Boolean(health.hasApiKey);
  els.serverState.textContent = health.hasApiKey ? `AI 已配置：${providerLabel}${health.model}，审查时实时校验` : "服务可用，未配置 API Key";
  els.serverState.className = `server-state ${health.hasApiKey ? "ok" : "warn"}`;
  if (restoreState()) {
    renderAnnotationLayer();
    return;
  }
  updateScoreboard({}, []);
  renderReviewSummary(null, []);
  renderReviewWorkflow(null, { phase: "idle" });
  renderIssueList();
}

function handleAnnotationClick(event) {
  const marker = event.target.closest(".source-annotation-rect");
  if (!marker) return;
  const issueIndex = Number(marker.dataset.riskIndex || 0) - 1;
  const issue = state.issues[issueIndex];
  if (!issue) return;
  state.filter = issueMajorCategory(issue);
  updateFilterButtons();
  const visibleIndex = currentIssues().indexOf(issue);
  if (visibleIndex >= 0) {
    state.activeIssueIndex = visibleIndex;
    renderIssueList();
  }
}

els.reviewFile.addEventListener("change", () => {
  const file = els.reviewFile.files[0];
  els.reviewFileHint.textContent = file ? `已选择：${file.name}` : "支持 Word、PDF、TXT，也可以直接粘贴文本";
  if (file) importContract();
});

els.reviewSource.addEventListener("click", handleAnnotationClick);

els.reviewSource.addEventListener("input", () => {
  if (officeEditor) return;
  state.workingText = sourceEditorText();
  els.reviewText.value = state.workingText;
  setSourceMeta(state.workingText);
  renderAnnotationLayer();
  if (state.reviewComplete) els.reviewBtn.textContent = "重新审查";
  saveState();
});

els.openImportBtn.addEventListener("click", openImportModal);
els.closeImportBtn.addEventListener("click", closeImportModal);
els.confirmImportBtn.addEventListener("click", importContract);
els.clearImportBtn.addEventListener("click", clearImport);
els.importModal.addEventListener("click", (event) => {
  if (event.target === els.importModal) closeImportModal();
});

els.reviewBtn.addEventListener("click", reviewContract);
els.acceptAllBtn?.addEventListener("click", applyAllSuggestions);

els.reviewFilterBar.addEventListener("click", (event) => {
  const button = event.target.closest("[data-review-filter]");
  if (!button) return;
  state.filter = button.dataset.reviewFilter;
  state.activeIssueIndex = 0;
  updateFilterButtons();
  renderIssueList();
  saveState();
});

els.reviewIssues.addEventListener("click", (event) => {
  const card = event.target.closest("[data-issue-index]");
  const issue = card ? state.issues[Number(card.dataset.issueIndex || 0) - 1] : currentIssue();
  if (event.target.closest("[data-apply-suggestion]")) {
    const button = event.target.closest("[data-apply-suggestion]");
    if (!button || button.disabled) return;
    if (applyIssueSuggestion(issue)) {
      renderIssueList();
    } else if (button) {
      const originalText = button.textContent;
      button.textContent = "未找到原文";
      window.setTimeout(() => {
        button.textContent = originalText;
      }, 1400);
    }
    return;
  }
  if (event.target.closest("[data-undo-suggestion]")) {
    if (undoIssueSuggestion(issue)) renderIssueList();
    return;
  }
  if (event.target.closest("[data-locate-issue]") || event.target.closest("[data-issue-index]")) {
    locateSpecificIssue(issue, true);
  }
});

els.prevIssueBtn.addEventListener("click", () => {
  state.activeIssueIndex = Math.max(0, state.activeIssueIndex - 1);
  renderIssueList();
  saveState();
});

els.nextIssueBtn.addEventListener("click", () => {
  state.activeIssueIndex = Math.min(currentIssues().length - 1, state.activeIssueIndex + 1);
  renderIssueList();
  saveState();
});

els.downloadCorrectedBtn.addEventListener("click", downloadCorrectedText);
els.exportReportBtn.addEventListener("click", exportReviewReport);
els.copyCorrectedBtn.addEventListener("click", copyCorrectedText);

loadInitial().catch((error) => {
  els.serverState.textContent = error.message;
  els.serverState.className = "server-state warn";
});
