const API_BASE = window.location.protocol === "file:" ? "http://localhost:5173" : "";

const state = {
  templates: [],
  builtInTemplates: [],
  customTemplates: [],
  selectedTemplateId: "",
  answers: {},
  intentSent: false,
  description: "",
  draftReady: false,
  editorDirty: false,
  messages: [],
  pendingQuestion: null,
  activeSourceTab: "all",
  allSuggestions: [],
  activeAnswerTab: "history",
};

const els = {
  serverState: document.getElementById("server-state"),
  templateList: document.getElementById("template-list"),
  templateModal: document.getElementById("template-modal"),
  openTemplateBtn: document.getElementById("open-template-btn"),
  closeTemplateBtn: document.getElementById("close-template-btn"),
  templatePickerLabel: document.getElementById("template-picker-label"),
  chat: document.getElementById("generate-chat"),
  composerSuggestions: document.getElementById("composer-suggestions"),
  intent: document.getElementById("generate-intent"),
  generateBtn: document.getElementById("generate-btn"),
  pauseGenerateBtn: document.getElementById("pause-generate-btn"),
  cancelGenerateBtn: document.getElementById("cancel-generate-btn"),
  generationCompletion: document.getElementById("generation-completion"),
  completionPercent: document.getElementById("completion-percent"),
  completionDetail: document.getElementById("completion-detail"),
  completionBar: document.getElementById("completion-bar"),
  contractEditorContainer: document.getElementById("contract-editor-container"),
  contractEditorTitle: document.getElementById("completion-title"),
  contractEditorMeta: document.getElementById("completion-detail"),
  copyContractBtn: document.getElementById("copy-contract-btn"),
  downloadContractBtn: document.getElementById("download-contract-btn"),
  regenerateContractBtn: document.getElementById("regenerate-contract-btn"),
  intentStatus: document.getElementById("intent-status"),
  templateStatus: document.getElementById("template-status"),
  fieldStatus: document.getElementById("field-status"),
  ruleStatus: document.getElementById("rule-status"),
  generationProgress: document.getElementById("generation-progress"),
  draftInsights: document.getElementById("draft-insights"),
  templateFieldPanel: document.getElementById("template-field-panel"),
  templateImportFile: document.getElementById("template-import-file"),
  templateImportText: document.getElementById("template-import-text"),
  importTemplateBtn: document.getElementById("import-template-btn"),
  templateImportStatus: document.getElementById("template-import-status"),
  answerSourceTabs: document.getElementById("answer-source-tabs"),
  answerPanel: document.getElementById("answer-panel"),
  answerTabs: document.getElementById("answer-tabs"),
  answerList: document.getElementById("answer-list"),
  answerInput: document.getElementById("answer-input"),
  answerSubmitBtn: document.getElementById("answer-submit-btn"),
};

const STORAGE_KEY = "contract-generation-workbench-v1";
let settingEditor = false;
let typingTimer = null;
let typingPaused = false;
let generationAbortController = null;
let officeEditor = null;
const initialGenerateMessage = "描述你要起草的合同，例如交易背景、双方、标的、金额、交付和付款安排。我会像对话助手一样逐项追问，左侧文档会实时生成并保留。";

const generationSteps = [
  { key: "intent", label: "意图识别" },
  { key: "template", label: "模板匹配" },
  { key: "knowledge", label: "资料检索" },
  { key: "fields", label: "要素补齐" },
  { key: "draft", label: "草稿生成" },
];

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

function mergeTemplateLists(builtIns = state.builtInTemplates, customs = state.customTemplates) {
  const seen = new Set();
  return [...builtIns, ...customs].filter((template) => {
    if (!template?.id || seen.has(template.id)) return false;
    seen.add(template.id);
    return true;
  });
}

function fieldKeyFromLabel(label = "") {
  const text = String(label || "").trim();
  if (/^(甲方|甲方名称|甲方全称|采购方|委托方|出租方|授权方|发包方|客户方|许可方|托运方|用人单位)$/.test(text)) return "partyA";
  if (/^(乙方|乙方名称|乙方全称|供应商|服务方|承租方|被许可方|承包方|经销方|承运方|劳动者|受托方)$/.test(text)) return "partyB";
  if (/金额|费用|价款|租金|报酬|工程款|许可费|运费|报名费用/.test(text)) return "amount";
  if (/付款|支付|缴费|结算|还款/.test(text)) return "payment";
  if (/期限|周期|工期|租期|服务期|订阅期限|许可期限/.test(text)) return "term";
  if (/交付|交货|运输|实施|施工|成果|服务成果/.test(text)) return "delivery";
  if (/验收|签收|考核|竣工/.test(text)) return "acceptance";
  if (/保密|数据|安全|隐私/.test(text)) return "confidentiality";
  if (/知识产权|权属|许可范围|模具/.test(text)) return "ipOwnership";
  if (/责任|违约|担保|侵权/.test(text)) return "liability";
  if (/范围|服务|内容|职责|授权区域|处理目的|加工要求|咨询范围/.test(text)) return "serviceScope";
  return text.replace(/[^\w\u4e00-\u9fa5]/g, "").slice(0, 24) || "field";
}

function extractFieldsFromTemplateText(text = "") {
  const fields = [];
  const seen = new Set();
  const re = /{{\s*([^{}]{1,40})\s*}}|【\s*([^【】]{1,40})\s*】/g;
  let match;
  while ((match = re.exec(String(text || "")))) {
    const label = String(match[1] || match[2] || "").trim();
    if (!label || seen.has(label)) continue;
    seen.add(label);
    fields.push({ key: fieldKeyFromLabel(label), label, question: `请填写${label}` });
  }
  return fields;
}

function parseCustomTemplateFields(text = "", templateText = "") {
  const lines = String(text || "")
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
  const fields = lines.map((line) => {
    const parts = line.split(/[|｜,，:：]/).map((part) => part.trim()).filter(Boolean);
    const label = parts[1] || parts[0];
    return { key: parts[1] ? parts[0] : fieldKeyFromLabel(label), label, question: `请填写${label}` };
  });
  const extracted = extractFieldsFromTemplateText(templateText);
  const seen = new Set();
  return [...fields, ...extracted].filter((field) => {
    const key = `${field.key}|${field.label}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function answerForTemplateLabel(label = "", template = selectedTemplate()) {
  const directKey = fieldKeyFromLabel(label);
  if (state.answers[directKey]) return state.answers[directKey];
  const matched = (template?.requiredFields || []).find((field) => field.label === label || field.key === label);
  if (matched && state.answers[matched.key]) return state.answers[matched.key];
  return state.answers[label] || "";
}

function fillTemplateText(templateText = "", template = selectedTemplate()) {
  const replaceOne = (rawLabel) => {
    const label = String(rawLabel || "").trim();
    return answerForTemplateLabel(label, template) || blank(label || "待补充");
  };
  return String(templateText || "")
    .replace(/{{\s*([^{}]{1,40})\s*}}/g, (_, label) => replaceOne(label))
    .replace(/【\s*([^【】]{1,40})\s*】/g, (_, label) => replaceOne(label));
}

function textToHtml(text = "") {
  const lines = String(text || "").split(/\n/);
  return lines
    .map((line) => {
      const value = line.trimEnd();
      if (!value.trim()) return "<p><br></p>";
      if (value.startsWith("# ")) return `<h1>${escapeHtml(value.slice(2))}</h1>`;
      if (value.startsWith("## ")) return `<h2>${escapeHtml(value.slice(3))}</h2>`;
      if (/^\d+\.\s+/.test(value)) return `<p class="clause-line">${escapeHtml(value)}</p>`;
      if (value.startsWith("- ")) return `<p class="list-line">• ${escapeHtml(value.slice(2))}</p>`;
      return `<p>${escapeHtml(value)}</p>`;
    })
    .join("");
}

function htmlToText(html = "") {
  const div = document.createElement("div");
  div.innerHTML = html;
  return div.innerText.trim();
}

function contractWordHtml(title = "合同内容", bodyHtml = "") {
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
    .clause-line, .list-line { text-indent: 0; }
    table { width: 100%; border-collapse: collapse; }
    td, th { border: 1px solid #000; padding: 6pt; }
  </style>
</head>
<body><div class="contract-doc">${bodyHtml}</div></body>
</html>`;
}

function stripGenerationProcess(text = "") {
  return String(text || "")
    .replace(/<think>[\s\S]*?<\/think>/gi, "")
    .replace(/```(?:json|markdown|md|text)?/gi, "")
    .replace(/```/g, "")
    .replace(/(?:^|\n)\s*#{0,6}\s*(?:思考过程|推理过程|分析过程|生成过程|规则命中|规则预检|生成说明|起草说明)[\s\S]*?(?=\n\s*#{1,6}\s*(?:合同|第一|一、|1[.、])|$)/g, "\n")
    .split("\n")
    .filter((line) => !/^\s*(?:思考|分析|推理|生成说明|起草说明|规则命中|规则预检|我将|我会先)[:：]/.test(line))
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function setEditorHtml(html) {
  settingEditor = true;
  if (officeEditor) officeEditor.setHtml(html || "");
  settingEditor = false;
}

function stopTypingDraft() {
  window.clearInterval(typingTimer);
  typingTimer = null;
  typingPaused = false;
  if (els.pauseGenerateBtn) {
    els.pauseGenerateBtn.hidden = true;
    els.pauseGenerateBtn.textContent = "暂停";
  }
}

function typeDraftIntoEditor(text, onDone) {
  const draft = stripGenerationProcess(text);
  stopTypingDraft();
  setEditorHtml("");
  let index = 0;
  const stepSize = Math.max(1, Math.ceil(draft.length / 900));
  if (els.pauseGenerateBtn) {
    els.pauseGenerateBtn.hidden = false;
    els.pauseGenerateBtn.textContent = "暂停";
  }
  typingTimer = window.setInterval(() => {
    if (typingPaused) return;
    index = Math.min(draft.length, index + stepSize);
    setEditorHtml(officeEditor.textToContractHtml(draft.slice(0, index)));
    const scrollWrap = els.contractEditorContainer.querySelector(".oe-scroll-wrap");
    if (scrollWrap) scrollWrap.scrollTop = scrollWrap.scrollHeight;
    if (index >= draft.length) {
      stopTypingDraft();
      if (onDone) onDone(draft);
    }
  }, 18);
}

function getEditorText() {
  return officeEditor ? officeEditor.getText() : "";
}

function saveState() {
  const payload = {
    selectedTemplateId: state.selectedTemplateId,
    customTemplates: state.customTemplates,
    answers: state.answers,
    intentSent: state.intentSent,
    description: state.description,
    draftReady: state.draftReady,
    editorDirty: state.editorDirty,
    messages: state.messages,
    pendingQuestion: state.pendingQuestion,
    editorHtml: officeEditor ? officeEditor.getHtml() : "",
    editorTitle: els.contractEditorTitle.textContent,
    editorMeta: els.contractEditorMeta.textContent,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function restoreState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return;
  try {
    const saved = JSON.parse(raw);
    state.customTemplates = Array.isArray(saved.customTemplates) ? saved.customTemplates : [];
    state.templates = mergeTemplateLists();
    state.selectedTemplateId = saved.selectedTemplateId || state.selectedTemplateId;
    state.answers = saved.answers || {};
    state.intentSent = Boolean(saved.intentSent);
    state.description = saved.description || "";
    state.draftReady = Boolean(saved.draftReady);
    state.editorDirty = Boolean(saved.editorDirty);
    state.messages = Array.isArray(saved.messages) ? saved.messages : [];
    state.pendingQuestion = saved.pendingQuestion || null;
    if (saved.editorTitle) els.contractEditorTitle.textContent = saved.editorTitle;
    if (saved.editorMeta) els.contractEditorMeta.textContent = saved.editorMeta;
    if (saved.editorHtml) setEditorHtml(saved.editorHtml);
  } catch {
    localStorage.removeItem(STORAGE_KEY);
  }
}

function renderTemplates() {
  state.templates = mergeTemplateLists();
  const currentTemplate = state.templates.find((template) => template.id === state.selectedTemplateId);
  if (els.templatePickerLabel) els.templatePickerLabel.textContent = currentTemplate?.name || "自动识别";
  els.templateList.innerHTML = state.templates
    .map(
      (template) => `
        <button class="template-pill ${template.id === state.selectedTemplateId ? "active" : ""}" type="button" data-template-id="${escapeHtml(template.id)}">
          <strong>${escapeHtml(template.name)}</strong>
          <small>${escapeHtml((template.keywords || template.outline || []).slice(0, 4).join(" / "))}</small>
          ${template.custom ? '<span class="tag neutral">自定义</span>' : ""}
        </button>
      `
    )
    .join("");
  const active = els.templateList.querySelector(".template-pill.active");
  if (active) active.scrollIntoView({ block: "nearest", inline: "center" });
  renderTemplateFieldPanel();
}

function renderMessages() {
  els.chat.innerHTML = state.messages
    .map((message) => `<div class="chat-message ${escapeHtml(message.role)}"><div class="chat-bubble">${message.html}</div></div>`)
    .join("");
  els.chat.scrollTop = els.chat.scrollHeight;
}

function renderComposerSuggestions(suggestions = [], sourceFilter) {
  if (!els.composerSuggestions) return;
  const filter = sourceFilter !== undefined ? sourceFilter : state.activeSourceTab;
  state.allSuggestions = suggestions || [];
  let filtered = state.allSuggestions;
  if (filter !== "all") {
    const sourceMap = { knowledge: "知识库推荐", history: "历史常用", model: "上下文/规则补全" };
    const targetSource = sourceMap[filter] || filter;
    filtered = state.allSuggestions.filter((item) => item.source === targetSource);
  }
  const list = filtered.slice(0, 5).filter((item) => item?.text);
  updateTabCounts(state.allSuggestions);
  if (!list.length) {
    els.composerSuggestions.hidden = true;
    els.composerSuggestions.innerHTML = "";
    return;
  }
  els.composerSuggestions.hidden = false;
  els.composerSuggestions.innerHTML = list
    .map(
      (suggestion) => `
        <button class="composer-suggestion-chip" type="button" data-suggestion="${escapeHtml(suggestion.text)}">
          <span>${escapeHtml(suggestion.text)}</span>
          <small>${escapeHtml(suggestion.source || "推荐")}</small>
        </button>
      `
    )
    .join("");
}

function updateTabCounts(suggestions = []) {
  if (!els.answerSourceTabs) return;
  const counts = { all: suggestions.length, knowledge: 0, history: 0, model: 0 };
  const sourceMap = { "知识库推荐": "knowledge", "历史常用": "history", "上下文/规则补全": "model" };
  for (const s of suggestions) {
    const key = sourceMap[s.source] || "model";
    counts[key]++;
  }
  els.answerSourceTabs.querySelectorAll(".answer-source-tab").forEach((tab) => {
    const source = tab.dataset.source;
    const count = counts[source] || 0;
    let countEl = tab.querySelector(".tab-count");
    if (count > 0) {
      if (!countEl) {
        countEl = document.createElement("span");
        countEl.className = "tab-count";
        tab.appendChild(countEl);
      }
      countEl.textContent = count;
    } else if (countEl) {
      countEl.remove();
    }
  });
}

function resetSourceTabs() {
  state.activeSourceTab = "all";
  if (els.answerSourceTabs) {
    els.answerSourceTabs.querySelectorAll(".answer-source-tab").forEach((t) => {
      t.classList.toggle("active", t.dataset.source === "all");
    });
  }
}

function clearComposerSuggestions() {
  renderComposerSuggestions([]);
}

function selectedTemplate() {
  return state.templates.find((template) => template.id === state.selectedTemplateId) || state.templates[0];
}

function templateCompletion(data = {}) {
  const template = data.template || selectedTemplate();
  const fields = template?.requiredFields || [];
  const missingKeys = new Set((data.missingFields || []).map((field) => field.key));
  const filled = fields.filter((field) => {
    if (missingKeys.has(field.key)) return false;
    return String(state.answers[field.key] || data.knownAnswers?.[field.key] || "").trim();
  }).length;
  const total = fields.length;
  const missing = data.missingFields?.length ?? Math.max(0, total - filled);
  const percent = total ? Math.max(0, Math.min(100, Math.round(((total - missing) / total) * 100))) : 0;
  return { template, filled: Math.max(0, total - missing), total, missing, percent };
}

function renderCompletionProgress(data = {}) {
  if (!els.generationCompletion) return;
  const { template, filled, total, missing, percent } = templateCompletion(data);
  els.completionPercent.textContent = `完成度 ${percent}%`;
  els.completionDetail.textContent = template
    ? `${template.name} · 已填 ${filled}/${total || 0} 项${missing ? ` · 还差 ${missing} 项` : " · 信息已补齐"}`
    : "选择或识别模板后会显示合同信息完成度";
  els.completionBar.style.width = `${percent}%`;
  els.generationCompletion.classList.toggle("complete", percent >= 100);
}

function renderTemplateFieldPanel() {
  if (els.templateFieldPanel) {
    els.templateFieldPanel.hidden = true;
    els.templateFieldPanel.innerHTML = "";
  }
  renderCompletionProgress();
}

function blank(label) {
  return `____（${label}）____`;
}

function inferAnswersFromText(text = "") {
  const inferred = {};
  const partyA = /(?:甲方|采购方|委托方|出租方|出借方|用人单位)[为是：:]?\s*([^，。；;\n]+)/.exec(text)?.[1];
  const partyB = /(?:乙方|供应商|服务方|承租方|借款方|劳动者)[为是：:]?\s*([^，。；;\n]+)/.exec(text)?.[1];
  const amount = /(?:金额|总价|价款|费用|租金|借款|薪资)[为是：:]?\s*([0-9,.]+万?元(?:\/[年月日季])?)/.exec(text)?.[1] || /([0-9,.]+万?元)/.exec(text)?.[1];
  const term = /(?:期限|周期|租期|服务期|合同期)[为是：:]?\s*([^，。；;\n]+)/.exec(text)?.[1];
  const payment = /(?:付款|支付|还款|结算)[为是：:]?\s*([^，。；;\n]+)/.exec(text)?.[1];
  const delivery = /(?:交付|交货|发货|交接)[为是：:]?\s*([^，。；;\n]+)/.exec(text)?.[1] || /([^，。；;\n]*(?:交付|交货|发货)[^，。；;\n]*)/.exec(text)?.[1];
  const subject = /(?:采购|购买|销售|服务|标的|产品|设备)[为是：:]?\s*([^，。；;\n]+)/.exec(text)?.[1];
  if (partyA) inferred.partyA = partyA.trim();
  if (partyB) inferred.partyB = partyB.trim();
  if (amount) inferred.amount = amount.trim();
  if (term) inferred.term = term.trim();
  if (payment) inferred.payment = payment.trim();
  if (delivery) inferred.delivery = delivery.trim();
  if (subject) inferred.subject = subject.trim();
  return inferred;
}

function mergeInferredAnswers(text) {
  const inferred = inferAnswersFromText(text);
  for (const [key, value] of Object.entries(inferred)) {
    if (value && !state.answers[key]) state.answers[key] = value;
  }
}

function buildLiveDraft(template = selectedTemplate()) {
  if (!template) return "";
  if (template.templateText) return fillTemplateText(template.templateText, template);
  const fields = template.requiredFields || [];
  const valueFor = (field) => state.answers[field.key] || blank(field.label);
  const fieldLines = fields.map((field) => `${field.label}：${valueFor(field)}`).join("\n");
  const outline = (template.outline || [])
    .map((item, index) => `${index + 1}. ${item}`)
    .join("\n");
  const partyA = state.answers.partyA || blank("甲方");
  const partyB = state.answers.partyB || blank("乙方");
  const amount = state.answers.amount || blank("合同金额");
  const delivery = state.answers.delivery || state.answers.term || blank("交付/履行安排");
  const payment = state.answers.payment || blank("付款/结算安排");
  const subject = state.answers.subject || state.answers.serviceScope || blank("合同标的");

  return `# ${template.name}\n\n甲方：${partyA}\n乙方：${partyB}\n\n## 一、合同要素\n${fieldLines}\n\n## 二、合同标的\n双方确认，本合同项下标的为：${subject}。\n\n## 三、价款与付款\n合同金额/计价方式为：${amount}。\n付款安排为：${payment}。\n\n## 四、交付、履行与验收\n交付或履行安排为：${delivery}。\n验收标准、异议期限和整改机制：${state.answers.acceptance || blank("验收标准")}。\n\n## 五、保密、知识产权与合规\n保密义务：${state.answers.confidentiality || blank("保密要求")}。\n知识产权归属：${state.answers.ipOwnership || blank("知识产权归属")}。\n\n## 六、违约责任\n违约责任：${state.answers.liability || blank("违约责任")}。\n\n## 七、争议解决\n争议解决方式：${state.answers.disputeResolution || "由双方协商确定，协商不成的按合同约定管辖处理。"}\n\n## 八、待完善条款结构\n${outline || "按所选模板结构补充。"}\n`;
}

function updateLivePreview(force = false) {
  const template = selectedTemplate();
  if (!template) return;
  if ((state.draftReady || state.editorDirty) && !force) return;
  els.contractEditorTitle.textContent = template.name;
  els.contractEditorMeta.textContent = "实时预览 · 空缺处已挖空";
  setEditorHtml(officeEditor ? officeEditor.textToContractHtml(buildLiveDraft(template)) : textToHtml(buildLiveDraft(template)));
  renderDraftInsights({ template, missingWarnings: (template.requiredFields || []).filter((field) => !state.answers[field.key]).map((field) => field.label) });
  renderCompletionProgress({ template, missingFields: (template.requiredFields || []).filter((field) => !state.answers[field.key]) });
  updateCommandStrip({ template, missingFields: (template.requiredFields || []).filter((field) => !state.answers[field.key]) });
  if (!document.activeElement?.closest?.("#template-field-panel")) renderTemplateFieldPanel();
  saveState();
}

function addMessage(role, html, persist = true) {
  const item = document.createElement("div");
  item.className = `chat-message ${role}`;
  item.innerHTML = `<div class="chat-bubble">${html}</div>`;
  els.chat.appendChild(item);
  els.chat.scrollTop = els.chat.scrollHeight;
  if (persist) {
    state.messages.push({ role, html });
    saveState();
  }
  return item;
}

function renderGenerationProgress(activeKey = "intent", doneKeys = []) {
  const done = new Set(doneKeys);
  els.generationProgress.innerHTML = generationSteps
    .map((step) => {
      const stateClass = done.has(step.key) ? "done" : step.key === activeKey ? "active" : "";
      return `
        <div class="flow-step ${stateClass}">
          <i></i>
          <span>${escapeHtml(step.label)}</span>
        </div>
      `;
    })
    .join("");
}

function setGenerationBusy(isBusy) {
  if (els.cancelGenerateBtn) els.cancelGenerateBtn.hidden = !isBusy;
  if (!isBusy) generationAbortController = null;
}

function startNewContract() {
  generationAbortController?.abort();
  stopTypingDraft();
  setGenerationBusy(false);
  clearComposerSuggestions();
  resetSourceTabs();
  hideAnswerPanel();
  state.selectedTemplateId = state.templates[0]?.id || "";
  state.answers = {};
  state.intentSent = false;
  state.description = "";
  state.draftReady = false;
  state.editorDirty = false;
  state.messages = [{ role: "assistant", html: initialGenerateMessage }];
  state.pendingQuestion = null;
  els.intent.value = "";
  els.contractEditorTitle.textContent = "";
  els.contractEditorMeta.textContent = "输入业务意图后实时预览";
  setEditorHtml("");
  renderMessages();
  renderTemplates();
  renderGenerationProgress("intent");
  renderDraftInsights();
  updateCommandStrip();
  renderCompletionProgress({ template: selectedTemplate(), missingFields: selectedTemplate()?.requiredFields || [] });
  saveState();
}

function updateCommandStrip(data = {}) {
  const knownCount = Object.values(state.answers).filter((value) => String(value || "").trim()).length;
  const missingCount = data.missingFields?.length ?? null;
  els.intentStatus.textContent = state.intentSent ? "已识别" : "待输入";
  const confidence = data.matchConfidence ? ` ${data.matchConfidence}%` : "";
  els.templateStatus.textContent = `${data.template?.name || state.templates.find((item) => item.id === state.selectedTemplateId)?.name || "自动识别"}${confidence}`;
  els.fieldStatus.textContent = missingCount === null ? `${knownCount} 项` : missingCount ? `缺 ${missingCount} 项` : "已补齐";
  els.ruleStatus.textContent = data.appliedRules?.length ? `命中 ${data.appliedRules.length} 条` : data.status === "draft_ready" ? "已预检" : "待生成";
}

function renderDraftInsights(data = {}) {
  const appliedCount = data.appliedRules?.length || 0;
  const warnings = data.missingWarnings || [];
  const templateName = data.template?.name || state.templates.find((item) => item.id === state.selectedTemplateId)?.name || "自动识别";
  const confidence = data.matchConfidence ? `${data.matchConfidence}%` : "自动判断";
  const candidateText = (data.templateCandidates || [])
    .slice(0, 2)
    .map((item) => `${item.name}${item.confidence ? ` ${item.confidence}%` : ""}`)
    .join(" / ");
  els.draftInsights.innerHTML = `
    <article>
      <span>当前模板</span>
      <strong>${escapeHtml(templateName)}</strong>
    </article>
    <article>
      <span>匹配置信</span>
      <strong>${escapeHtml(confidence)}</strong>
    </article>
    <article>
      <span>规则命中</span>
      <strong>${appliedCount ? `${appliedCount} 条` : "待预检"}</strong>
    </article>
    <article>
      <span>待确认</span>
      <strong>${warnings.length ? `${warnings.length} 项` : "无"}</strong>
    </article>
    ${candidateText ? `<article class="wide"><span>候选模板</span><strong>${escapeHtml(candidateText)}</strong></article>` : ""}
  `;
}

function updateContractEditor(data = {}, options = {}) {
  state.draftReady = true;
  state.editorDirty = false;
  const title = data.title || data.template?.name || "合同草稿";
  const templateName = data.template?.name || "";
  els.contractEditorTitle.textContent = title;
  const doneMeta = `${templateName || "已生成草稿"}${data.usedFallback ? " · 本地兜底" : ""} · 已预检`;
  const draft = stripGenerationProcess(data.draft || "");
  if (options.animate) {
    els.contractEditorMeta.textContent = "逐字生成中...";
    typeDraftIntoEditor(draft, () => {
      els.contractEditorMeta.textContent = doneMeta;
      setGenerationBusy(false);
      saveState();
    });
  } else {
    stopTypingDraft();
    els.contractEditorMeta.textContent = doneMeta;
    setEditorHtml(textToHtml(draft));
  }
  renderDraftInsights(data);
  updateCommandStrip(data);
  renderTemplateFieldPanel();
  saveState();
}

function collectAnswers() {
  const changed = [];
  document.querySelectorAll("[data-generation-field]").forEach((input) => {
    const value = input.value.trim();
    if (!value) return;
    if (state.answers[input.dataset.generationField] !== value) {
      changed.push(`${input.dataset.generationLabel || input.dataset.generationField}：${value}`);
    }
    state.answers[input.dataset.generationField] = value;
  });
  updateLivePreview();
  return changed;
}

function renderQuestionDialog(data) {
  state.selectedTemplateId = data.template.id;
  state.answers = { ...state.answers, ...(data.knownAnswers || {}) };
  renderTemplates();
  renderGenerationProgress("fields", ["intent", "template", "knowledge"]);
  updateCommandStrip(data);
  renderDraftInsights(data);
  updateLivePreview();
  const item = data.question || data.questions?.[0];
  if (!item) return;
  state.pendingQuestion = item;
  state.allSuggestions = item.suggestions || [];
  renderCompletionProgress(data);
  const switchText =
    data.templateSwitched && data.requestedTemplate
      ? `<div class="notice">已按你的描述从"${escapeHtml(data.requestedTemplate.name)}"自动切换为"${escapeHtml(data.template.name)}"。</div>`
      : "";
  addMessage(
    "assistant",
    `
      <div class="dialog-title compact-question-title">
        <strong>${escapeHtml(item.label || "请补充关键信息")}</strong>
        <span class="tag mid">还差 ${data.missingFields.length} 项</span>
      </div>
      ${switchText}
      <div class="field-question">
        <strong>${escapeHtml(item.question)}</strong>
        <p>从下方选择一个答案，或直接输入自定义内容后点击确认。</p>
      </div>
    `
  );
  showAnswerPanel(item.suggestions || []);
  saveState();
}

function showAnswerPanel(suggestions) {
  if (!els.answerPanel) return;
  els.answerPanel.hidden = false;
  state.activeAnswerTab = "history";
  renderAnswerTabContent(suggestions);
  updateAnswerTabBadges(suggestions);
  if (els.answerInput) {
    els.answerInput.value = "";
    els.answerInput.placeholder = "输入自定义答案，或点击上方选项...";
  }
}

function hideAnswerPanel() {
  if (els.answerPanel) els.answerPanel.hidden = true;
  state.activeAnswerTab = "history";
}

function renderAnswerTabContent(suggestions) {
  if (!els.answerList) return;
  const sourceMap = { knowledge: "知识库推荐", history: "历史常用", model: "上下文/规则补全" };
  const targetSource = sourceMap[state.activeAnswerTab] || state.activeAnswerTab;
  const items = suggestions.filter((s) => s.source === targetSource);
  if (!items.length) {
    els.answerList.innerHTML = `<div class="answer-empty">暂无${state.activeAnswerTab === "knowledge" ? "知识库" : state.activeAnswerTab === "history" ? "历史" : "AI"}相关答案</div>`;
    return;
  }
  els.answerList.innerHTML = items
    .map(
      (item, i) =>
        `<button class="answer-item" type="button" data-answer-index="${i}" data-answer-text="${escapeHtml(item.text)}">${escapeHtml(item.text)}</button>`
    )
    .join("");
}

function updateAnswerTabBadges(suggestions) {
  if (!els.answerTabs) return;
  const counts = { knowledge: 0, history: 0, model: 0 };
  const sourceMap = { "知识库推荐": "knowledge", "历史常用": "history", "上下文/规则补全": "model" };
  for (const s of suggestions) {
    const key = sourceMap[s.source] || "model";
    counts[key]++;
  }
  els.answerTabs.querySelectorAll(".answer-tab").forEach((tab) => {
    const source = tab.dataset.source;
    const count = counts[source] || 0;
    let badge = tab.querySelector(".tab-badge");
    if (count > 0) {
      if (!badge) {
        badge = document.createElement("span");
        badge.className = "tab-badge";
        tab.appendChild(badge);
      }
      badge.textContent = count;
    } else if (badge) {
      badge.remove();
    }
  });
}

async function validatePendingAnswer(item, answer) {
  if (!item?.key || !String(answer || "").trim()) {
    return { accepted: false, reason: "请先补充有效内容。" };
  }
  const template = selectedTemplate();
  const field = (template?.requiredFields || []).find((entry) => entry.key === item.key) || item;
  const data = await api("/api/contracts/validate-answer", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      templateId: state.selectedTemplateId,
      customTemplate: template?.custom ? template : null,
      field,
      answer,
      answers: state.answers,
      description: state.description,
    }),
  });
  return {
    accepted: data.accepted !== false,
    normalizedValue: String(data.normalizedValue || data.normalized_value || answer || "").trim(),
    reason: data.reason || "",
    warnings: Array.isArray(data.warnings) ? data.warnings : [],
    suggestions: Array.isArray(data.suggestions) ? data.suggestions : [],
    usedAI: Boolean(data.usedAI),
  };
}

function renderValidationFeedback(validation, item, originalAnswer) {
  if (!validation) return;
  if (validation.accepted) {
    const normalized = validation.normalizedValue || originalAnswer;
    if (normalized !== originalAnswer || validation.warnings.length) {
      const warningText = validation.warnings.length
        ? `<p>${validation.warnings.map(escapeHtml).join("；")}</p>`
        : "";
      addMessage(
        "assistant",
        `<div class="notice">已校验并填入"${escapeHtml(item.label || item.key)}"：${escapeHtml(normalized)}</div>${warningText}`
      );
    }
    return;
  }
  const suggestions = validation.suggestions?.length
    ? `<div class="composer-inline-suggestions">${validation.suggestions.slice(0, 3).map((suggestion) => {
        const text = suggestion.text || suggestion;
        return `<button type="button" class="composer-suggestion-chip" data-suggestion="${escapeHtml(text)}"><span>${escapeHtml(text)}</span><small>${escapeHtml(suggestion.source || "推荐")}</small></button>`;
      }).join("")}</div>`
    : "";
  addMessage(
    "assistant",
    `<div class="notice">这项信息暂未写入合同：${escapeHtml(validation.reason || "内容与当前合同上下文不匹配，请重新补充。")}</div>${suggestions}`
  );
}

function renderDraft(data) {
  state.selectedTemplateId = data.template?.id || state.selectedTemplateId;
  state.answers = { ...state.answers, ...(data.answers || {}) };
  state.pendingQuestion = null;
  clearComposerSuggestions();
  hideAnswerPanel();
  renderTemplates();
  renderCompletionProgress({ ...data, missingFields: [] });
  updateContractEditor(data, { animate: true });
  renderGenerationProgress("draft", ["intent", "template", "knowledge", "fields", "draft"]);
  const switchText =
    data.templateSwitched && data.requestedTemplate
      ? `<div class="notice">已按你的描述从"${escapeHtml(data.requestedTemplate.name)}"自动切换为"${escapeHtml(data.template.name)}"。</div>`
      : "";
  addMessage(
    "assistant",
    `
      <div class="result-head">
        <div>
          <h3>${escapeHtml(data.title || data.template?.name || "合同草稿")}</h3>
          <p>合同已生成到右侧编辑区，已结合模板匹配、资料检索和规则库预检。</p>
        </div>
      </div>
      ${switchText}
      ${
        data.missingWarnings?.length
          ? `<div class="notice">仍建议确认：${data.missingWarnings.map(escapeHtml).join("、")}</div>`
          : ""
      }
    `
  );
}

async function requestGeneration(options = {}) {
  const changedAnswers = collectAnswers();
  const message = els.intent.value.trim();
  const button = els.generateBtn;
  const currentTemplate = selectedTemplate();
  const submitButton = options.sourceButton || button;
  let userMessageRendered = false;
  let messageForDescription = message;
  if (message && state.pendingQuestion?.key) {
    const pending = { ...state.pendingQuestion };
    addMessage("user", escapeHtml(message));
    userMessageRendered = true;
    els.intent.value = "";
    state.intentSent = true;
    const originalText = submitButton.textContent;
    submitButton.disabled = true;
    submitButton.classList.add("loading");
    submitButton.textContent = "校验中...";
    let validation;
    try {
      validation = await validatePendingAnswer(pending, message);
    } catch (error) {
      validation = {
        accepted: true,
        normalizedValue: message,
        warnings: [`暂未完成大模型校验，已按原回答填入：${error.message}`],
        usedAI: false,
      };
    } finally {
      submitButton.disabled = false;
      submitButton.classList.remove("loading");
      submitButton.textContent = originalText;
    }
    if (!validation.accepted) {
      renderValidationFeedback(validation, pending, message);
      renderComposerSuggestions(validation.suggestions?.length ? validation.suggestions : pending.suggestions || [], state.activeSourceTab);
      saveState();
      return;
    }
    const normalizedAnswer = validation.normalizedValue || message;
    state.answers[pending.key] = normalizedAnswer;
    state.pendingQuestion = null;
    messageForDescription = `${pending.label || pending.key}：${normalizedAnswer}`;
    renderValidationFeedback(validation, pending, message);
  }
  if (message) {
    clearComposerSuggestions();
    state.description = [state.description, messageForDescription].filter(Boolean).join("\n");
    updateLivePreview();
  }
  let description = state.description.trim();
  if (!description && options.forceDraft) {
    const template = selectedTemplate();
    const fieldText = (template?.requiredFields || [])
      .map((field) => `${field.label}：${state.answers[field.key] || ""}`)
      .filter((line) => !line.endsWith("："))
      .join("\n");
    description = `${template?.name || "合同模板"}\n${fieldText}`.trim();
    state.description = description;
  }
  if (!description) {
    addMessage("assistant", "请先输入合同生成意图或业务描述。");
    return;
  }
  if (message && !userMessageRendered) {
    addMessage("user", escapeHtml(message));
    els.intent.value = "";
    state.intentSent = true;
  } else if (changedAnswers.length) {
    addMessage("user", escapeHtml(changedAnswers.join("\n")));
  }
  const originalText = button.textContent;
  const submitOriginalText = submitButton.textContent;
  button.disabled = true;
  submitButton.disabled = true;
  submitButton.classList.add("loading");
  submitButton.textContent = options.regenerate ? "重新生成中..." : options.forceDraft ? "模板生成中..." : "正在识别...";
  generationAbortController = new AbortController();
  setGenerationBusy(true);
  renderGenerationProgress("knowledge", ["intent", "template"]);
  updateCommandStrip({ status: "matching" });
  try {
    const data = await api("/api/contracts/generate", {
      method: "POST",
      headers: { "content-type": "application/json" },
      signal: generationAbortController.signal,
      body: JSON.stringify({
        description,
        templateId: state.selectedTemplateId,
        customTemplate: currentTemplate?.custom ? currentTemplate : null,
        answers: state.answers,
        forceDraft: Boolean(options.forceDraft),
        existingDraft: options.regenerate ? getEditorText() : "",
      }),
    });
    if (data.status === "need_more_info") renderQuestionDialog(data);
    else renderDraft(data);
  } catch (error) {
    if (error.name === "AbortError") addMessage("assistant", "已取消本次生成。");
    else addMessage("assistant", escapeHtml(error.message));
  } finally {
    if (typingTimer) {
      generationAbortController = null;
      if (els.cancelGenerateBtn) els.cancelGenerateBtn.hidden = false;
    } else {
      setGenerationBusy(false);
    }
    button.disabled = false;
    submitButton.disabled = false;
    submitButton.classList.remove("loading");
    button.textContent = originalText;
    submitButton.textContent = submitOriginalText;
  }
}

async function copyContractContent() {
  if (!officeEditor) return;
  await officeEditor.copyContent();
  els.copyContractBtn.textContent = "已复制";
  window.setTimeout(() => {
    els.copyContractBtn.textContent = "复制";
  }, 1400);
}

function downloadContractContent() {
  if (!officeEditor) return;
  const title = (els.contractEditorTitle.textContent || "合同内容").replace(/[\\/:*?"<>|]/g, "_");
  officeEditor.downloadDoc(title);
}

function saveCustomTemplate() {
  return;
}

async function importTemplateFromContract() {
  const file = els.templateImportFile?.files?.[0] || null;
  const templateText = els.templateImportText?.value.trim() || "";
  if (!file && !templateText) {
    if (els.templateImportStatus) els.templateImportStatus.textContent = "请先上传合同文件或粘贴合同正文。";
    els.templateImportText?.focus();
    return;
  }
  const button = els.importTemplateBtn;
  const originalText = button?.textContent || "";
  if (button) {
    button.disabled = true;
    button.classList.add("loading");
    button.textContent = "正在抽取模板...";
  }
  if (els.templateImportStatus) els.templateImportStatus.textContent = "正在解析合同结构、抽取占位字段和常用条款...";
  try {
    const form = new FormData();
    if (file) form.append("file", file);
    if (templateText) form.append("templateText", templateText);
    const data = await api("/api/contract-templates/import", {
      method: "POST",
      body: form,
    });
    const template = data.template;
    state.builtInTemplates = [template, ...state.builtInTemplates.filter((item) => item.id !== template.id && item.name !== template.name)];
    state.customTemplates = state.customTemplates.filter((item) => item.id !== template.id && item.name !== template.name);
    state.templates = mergeTemplateLists();
    state.selectedTemplateId = template.id;
    state.draftReady = false;
    state.editorDirty = false;
    state.answers = {};
    renderTemplates();
    updateLivePreview(true);
    updateCommandStrip({ template, missingFields: template.requiredFields || [] });
    if (els.templateImportStatus) {
      els.templateImportStatus.textContent = data.usedAI
        ? `已通过 AI 生成模板"${template.name}"，字段 ${template.requiredFields?.length || 0} 项。`
        : `已用本地算法生成模板"${template.name}"，字段 ${template.requiredFields?.length || 0} 项。`;
    }
    addMessage("assistant", `已导入并保存模板"${escapeHtml(template.name)}"，可以直接开始按对话补全信息。`);
    saveState();
  } catch (error) {
    if (els.templateImportStatus) els.templateImportStatus.textContent = error.message;
  } finally {
    if (button) {
      button.disabled = false;
      button.classList.remove("loading");
      button.textContent = originalText;
    }
  }
}

function saveCustomTemplateLegacy() {
  const name = document.getElementById("custom-template-name")?.value.trim() || "自定义合同模板";
  const templateText = document.getElementById("custom-template-content")?.value.trim() || "";
  if (!templateText) {
    document.getElementById("custom-template-content")?.focus();
    return;
  }
  const fieldText = document.getElementById("custom-template-fields")?.value || "";
  const requiredFields = parseCustomTemplateFields(fieldText, templateText);
  const outline = (templateText.match(/第[一二三四五六七八九十]+章[^\n]{0,24}|[一二三四五六七八九十]+、[^\n]{0,24}/g) || ["合同主体", "合同内容", "费用付款", "权利义务", "违约责任", "争议解决"]).slice(0, 12);
  const template = {
    id: `custom_${Date.now()}`,
    name,
    keywords: [name, "自定义模板", ...outline.slice(0, 3)],
    outline,
    requiredFields,
    templateText,
    custom: true,
  };
  state.customTemplates = [template, ...state.customTemplates.filter((item) => item.name !== name)].slice(0, 30);
  state.templates = mergeTemplateLists();
  state.selectedTemplateId = template.id;
  state.draftReady = false;
  state.editorDirty = false;
  renderTemplates();
  updateLivePreview(true);
  updateCommandStrip({ template, missingFields: requiredFields.filter((field) => !state.answers[field.key]) });
  addMessage("assistant", `已保存自定义模板"${escapeHtml(name)}"，左侧字段可填写，右侧会实时预览。`);
  saveState();
}

async function loadInitial() {
  officeEditor = new OfficeEditor("contract-editor-container", {
    mode: "word",
    placeholder: "合同草稿会在这里生成，支持富文本修改。",
    contractMode: true,
    onChange() {
      if (!settingEditor) {
        state.editorDirty = true;
        saveState();
      }
    },
  });

  const [health, templates] = await Promise.all([api("/api/health"), api("/api/contract-templates")]);
  const providerLabel = health.modelProvider && health.modelProvider !== "None" ? `${health.modelProvider} / ` : "";
  els.serverState.textContent = health.hasApiKey ? `AI 已配置：${providerLabel}${health.model}，生成时实时校验` : "服务可用，未配置 API Key";
  els.serverState.className = `server-state ${health.hasApiKey ? "ok" : "warn"}`;
  state.builtInTemplates = templates.templates || [];
  state.templates = mergeTemplateLists();
  restoreState();
  state.templates = mergeTemplateLists();
  state.selectedTemplateId = state.selectedTemplateId || state.templates[0]?.id || "";
  renderTemplates();
  if (state.messages.length) renderMessages();
  if (state.pendingQuestion?.suggestions?.length) renderComposerSuggestions(state.pendingQuestion.suggestions);
  renderGenerationProgress("intent");
  renderDraftInsights();
  updateCommandStrip();
  if (!officeEditor.getHtml().trim()) updateLivePreview(true);
}

els.templateList.addEventListener("click", (event) => {
  const button = event.target.closest("[data-template-id]");
  if (!button) return;
  state.selectedTemplateId = button.dataset.templateId;
  state.draftReady = false;
  renderTemplates();
  updateCommandStrip();
  renderDraftInsights();
  updateLivePreview(true);
  els.templateModal.hidden = true;
});

els.openTemplateBtn.addEventListener("click", () => {
  els.templateModal.hidden = false;
});

els.closeTemplateBtn.addEventListener("click", () => {
  els.templateModal.hidden = true;
});

els.templateModal.addEventListener("click", (event) => {
  if (event.target === els.templateModal) els.templateModal.hidden = true;
});

els.generateBtn.addEventListener("click", () => requestGeneration());
els.pauseGenerateBtn?.addEventListener("click", () => {
  if (!typingTimer) return;
  typingPaused = !typingPaused;
  els.pauseGenerateBtn.textContent = typingPaused ? "继续" : "暂停";
});
els.cancelGenerateBtn?.addEventListener("click", () => {
  generationAbortController?.abort();
  stopTypingDraft();
  setGenerationBusy(false);
});
els.regenerateContractBtn.addEventListener("click", startNewContract);
els.importTemplateBtn?.addEventListener("click", importTemplateFromContract);
els.intent.addEventListener("input", () => {
  saveState();
});
els.intent.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" || event.shiftKey || event.isComposing) return;
  event.preventDefault();
  if (!els.generateBtn.disabled) requestGeneration();
});

els.composerSuggestions?.addEventListener("click", (event) => {
  const suggestion = event.target.closest("[data-suggestion]");
  if (!suggestion) return;
  els.intent.value = suggestion.dataset.suggestion || "";
  els.intent.focus();
  saveState();
});

els.answerTabs?.addEventListener("click", (event) => {
  const tab = event.target.closest(".answer-tab");
  if (!tab) return;
  state.activeAnswerTab = tab.dataset.source;
  els.answerTabs.querySelectorAll(".answer-tab").forEach((t) => t.classList.remove("active"));
  tab.classList.add("active");
  renderAnswerTabContent(state.allSuggestions);
});

els.answerList?.addEventListener("click", (event) => {
  const item = event.target.closest(".answer-item");
  if (!item) return;
  if (els.answerInput) els.answerInput.value = item.dataset.answerText || "";
  els.answerList.querySelectorAll(".answer-item").forEach((btn) => btn.classList.remove("selected"));
  item.classList.add("selected");
});

els.answerSubmitBtn?.addEventListener("click", () => {
  const value = els.answerInput?.value.trim() || "";
  if (!value) return;
  if (els.intent) els.intent.value = value;
  hideAnswerPanel();
  requestGeneration();
});

els.answerInput?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" || event.isComposing) return;
  event.preventDefault();
  els.answerSubmitBtn?.click();
});

els.templateFieldPanel.addEventListener("input", (event) => {
  const input = event.target.closest("[data-generation-field]");
  if (!input) return;
  state.answers[input.dataset.generationField] = input.value.trim();
  state.draftReady = false;
  state.editorDirty = false;
  updateLivePreview(true);
  saveState();
});

els.templateFieldPanel.addEventListener("click", (event) => {
  if (!event.target.closest("[data-template-generate]")) return;
  requestGeneration({ forceDraft: true, sourceButton: event.target.closest("[data-template-generate]") });
});

els.chat.addEventListener("click", async (event) => {
  const candidate = event.target.closest(".candidate-chip[data-template-id]");
  if (candidate) {
    state.selectedTemplateId = candidate.dataset.templateId;
    state.draftReady = false;
    renderTemplates();
    updateCommandStrip();
    renderDraftInsights();
    updateLivePreview(true);
    addMessage("user", `切换为${escapeHtml(candidate.querySelector("strong")?.textContent || "所选模板")}`);
    return;
  }
  const suggestion = event.target.closest("[data-suggestion]");
  if (suggestion) {
    els.intent.value = suggestion.dataset.suggestion || "";
    els.intent.focus();
    return;
  }
  if (event.target.closest("[data-submit-answers]")) {
    requestGeneration();
    return;
  }
});

els.chat.addEventListener("input", (event) => {
  const input = event.target.closest("[data-generation-field]");
  if (!input) return;
  const value = input.value.trim();
  if (value) state.answers[input.dataset.generationField] = value;
  state.draftReady = false;
  updateLivePreview();
  saveState();
});

els.copyContractBtn.addEventListener("click", copyContractContent);
els.downloadContractBtn.addEventListener("click", downloadContractContent);

els.answerSourceTabs?.addEventListener("click", (event) => {
  const tab = event.target.closest(".answer-source-tab");
  if (!tab) return;
  const source = tab.dataset.source;
  state.activeSourceTab = source;
  els.answerSourceTabs.querySelectorAll(".answer-source-tab").forEach((t) => t.classList.remove("active"));
  tab.classList.add("active");
  renderComposerSuggestions(state.allSuggestions, source);
});

loadInitial().catch((error) => {
  els.serverState.textContent = error.message;
  els.serverState.className = "server-state warn";
});
