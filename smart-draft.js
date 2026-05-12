// smart-draft.js — Template Library + AI Draft + Word Real-time Display

let sessionId = "";
let knowledgeMode = "industry";
let currentDraft = "";
let selectedTpl = null;
let aiMode = false;
let currentQuestion = null;
let isProcessing = false;
let attachedFiles = [];

const STEPS = ["intent","match","fields","draft","done"];
const STEP_LABELS = { intent:"分析合同意图...", match:"匹配合同模板...", fields:"识别关键条款...", draft:"生成合同草稿...", done:"完成" };

// ── Init ──
window.addEventListener("DOMContentLoaded", () => {
  const icons = { navSD:"smartDraft", navTF:"templateFill", navCR:"contractReview", navKB:"knowledge", icoStart:"sparkles", icoClip:"paperclip", icoSend:"send" };
  for (const [id, n] of Object.entries(icons)) { const el=document.getElementById(id); if(el&&window.AppIcons) el.innerHTML=AppIcons[n]||""; }
  fetchAIStatus();
  loadTemplates();
});

async function fetchAIStatus() {
  try {
    const d = await (await fetch("/api/health")).json();
    document.getElementById("aiModel").textContent = d.model || "unknown";
    document.getElementById("aiProvider").textContent = d.provider || "";
    document.getElementById("aiDot").style.background = d.apiKeyConfigured ? "var(--green)" : "var(--red)";
  } catch { document.getElementById("aiModel").textContent="离线"; document.getElementById("aiDot").style.background="var(--red)"; }
}

// ── Template Library ──
let allTemplates = [];

async function loadTemplates() {
  try {
    const res = await fetch("/api/smart-draft/templates");
    allTemplates = await res.json();
    renderTemplates(allTemplates);
  } catch { document.getElementById("tplGrid").innerHTML = '<div style="color:var(--muted);font-size:13px">加载模板失败</div>'; }
}

function renderTemplates(tpls) {
  const grid = document.getElementById("tplGrid");
  grid.innerHTML = tpls.map(t => `
    <div class="tpl-card ${selectedTpl?.id===t.id?'selected':''}" onclick="selectTpl('${t.id}')">
      <div class="tpl-name">${t.name}</div>
      <div class="tpl-desc">${t.description}</div>
      <span class="tpl-tag">${t.category||t.type}</span>
    </div>
  `).join("");
}

function selectTpl(id) {
  selectedTpl = allTemplates.find(t => t.id === id) || null;
  renderTemplates(allTemplates);
  // Switch to AI mode with hint
  document.getElementById("aiMode").style.display = "block";
  document.getElementById("tplMode").style.display = "none";
  document.querySelectorAll(".mode-tabs button").forEach(b=>b.classList.remove("active"));
  document.querySelectorAll(".mode-tabs button")[1].classList.add("active");
  aiMode = true;
  document.getElementById("aiDesc").placeholder = `已选择「${selectedTpl.name}」模板，请描述合同具体信息...`;
  document.getElementById("aiDesc").focus();
}

function switchMode(mode, btn) {
  aiMode = (mode === "ai");
  document.querySelectorAll(".mode-tabs button").forEach(b=>b.classList.remove("active"));
  btn.classList.add("active");
  document.getElementById("tplMode").style.display = aiMode ? "none" : "block";
  document.getElementById("aiMode").style.display = aiMode ? "block" : "none";
  if (aiMode) { selectedTpl = null; renderTemplates(allTemplates); document.getElementById("aiDesc").focus(); }
}

function filterTemplates() {
  const q = document.getElementById("tplSearch").value.toLowerCase();
  const cat = document.getElementById("tplCategory").value;
  const filtered = allTemplates.filter(t => {
    if (cat && t.type !== cat) return false;
    if (q && !t.name.toLowerCase().includes(q) && !t.description.toLowerCase().includes(q)) return false;
    return true;
  });
  renderTemplates(filtered);
}

function toggleKB(el) { el.classList.toggle("active"); }

// ── Start Draft ──
async function startDraft() {
  const desc = document.getElementById("aiDesc")?.value.trim() || "";
  const tplHint = selectedTpl ? `${selectedTpl.name}模板` : "";
  const fullDesc = [tplHint, desc].filter(Boolean).join("，");

  if (!fullDesc) return alert("请选择模板或输入需求描述");
  if (isProcessing) return;
  isProcessing = true;

  // Switch to draft view
  document.getElementById("startPanel").style.display = "none";
  document.getElementById("draftView").style.display = "flex";
  document.getElementById("statusTag").style.display = "inline-flex";
  document.getElementById("agentBar").style.display = "flex";

  showAgentStep("intent");

  try {
    const res = await fetch("/api/smart-draft/init", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ description: fullDesc, knowledgeMode, fileTexts: [] }),
    });
    const txt = await res.text();
    let data; try { data = JSON.parse(txt); } catch { throw new Error("服务器响应异常"); }
    if (data.error) throw new Error(data.error);

    sessionId = data.sessionId;
    showAgentStep("match");

    if (data.messages) renderMessages(data.messages);

    if (data.dialogue?.missingCount > 0) {
      showAgentStep("fields");
      currentQuestion = data.dialogue.currentQuestion;
      showGuideOptions();
    } else {
      showAgentStep("draft");
      showGenerateBtn();
    }
  } catch(e) {
    addMsg("assistant", "错误: " + e.message);
  }
  isProcessing = false;
}

// ── Agent Progress ──
function showAgentStep(step) {
  const idx = STEPS.indexOf(step);
  STEPS.forEach((s,i) => {
    const el = document.getElementById("as-"+s);
    if(!el) return;
    el.className = "s";
    if (i < idx) el.classList.add("done");
    else if (i === idx) el.classList.add("active");
  });
  document.getElementById("agentLabel").textContent = STEP_LABELS[step] || step;
}

// ── Chat ──
function addMsg(role, content) {
  const t = document.getElementById("chatThread");
  const d = document.createElement("div");
  d.className = `chat-msg ${role}`;
  d.innerHTML = `<div class="avatar">${role==="assistant"?"AI":"U"}</div><div class="bubble">${esc(content)}</div>`;
  t.appendChild(d); t.scrollTop = t.scrollHeight;
}

function renderMessages(msgs) {
  const t = document.getElementById("chatThread");
  t.innerHTML = "";
  (msgs||[]).forEach(m => {
    const d = document.createElement("div");
    d.className = `chat-msg ${m.role}`;
    d.innerHTML = `<div class="avatar">${m.role==="assistant"?"AI":"U"}</div><div class="bubble">${esc(m.content)}</div>`;
    t.appendChild(d);
  });
  t.scrollTop = t.scrollHeight;
}

// ── Guide Options ──
function showGuideOptions() {
  document.getElementById("guideBar").style.display = "flex";
  // Add question to chat
  if (currentQuestion) {
    const t = document.getElementById("chatThread");
    const d = document.createElement("div");
    d.className = "chat-msg assistant";
    d.innerHTML = `<div class="avatar">AI</div><div class="bubble">${esc(currentQuestion.question||"请补充信息")}</div>`;
    t.appendChild(d); t.scrollTop = t.scrollHeight;
  }
}

function showGenerateBtn() {
  const t = document.getElementById("chatThread");
  const d = document.createElement("div");
  d.className = "chat-msg assistant";
  d.innerHTML = `<div class="avatar">AI</div><div><div class="bubble">信息已齐全，可以生成草稿。</div><div style="padding-top:8px"><button class="btn btn-primary btn-sm" onclick="generateDraft()">生成合同草稿</button></div></div>`;
  t.appendChild(d); t.scrollTop = t.scrollHeight;
}

async function pickGuide(mode) {
  document.getElementById("guideBar").style.display = "none";
  if (mode === "ask") { document.getElementById("chatInput").focus(); return; }

  const labels = { kb: "从知识库检索", ai: "AI 自动推断" };
  addMsg("user", labels[mode]);

  if (mode === "ai") { showAgentStep("draft"); addMsg("assistant", "正在推断..."); }

  try {
    const res = await fetch("/api/smart-draft/guide", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, mode: mode === "kb" ? "kb_search" : "llm_infer" }),
    });
    const txt = await res.text();
    let data; try { data = JSON.parse(txt); } catch { throw new Error("服务器异常"); }
    if (data.messages) renderMessages(data.messages);
    if (data.dialogue?.missingCount > 0) {
      currentQuestion = data.dialogue.currentQuestion;
      showGuideOptions();
    } else {
      showAgentStep("draft");
      showGenerateBtn();
    }
  } catch(e) { addMsg("assistant", "错误: " + e.message); }
}

// ── Send Answer ──
async function sendAnswer() {
  const input = document.getElementById("chatInput");
  const val = input.value.trim();
  if (!val || !sessionId) return;
  input.value = "";
  document.getElementById("guideBar").style.display = "none";

  const field = currentQuestion?.hint || "补充信息";
  addMsg("user", val);

  try {
    const res = await fetch("/api/smart-draft/answer", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, field, value: val }),
    });
    const txt = await res.text();
    let data; try { data = JSON.parse(txt); } catch { throw new Error("服务器异常"); }
    if (data.messages) renderMessages(data.messages);
    if (data.dialogue?.missingCount > 0) {
      currentQuestion = data.dialogue.currentQuestion;
      showGuideOptions();
    } else {
      showAgentStep("draft");
      showGenerateBtn();
    }
  } catch(e) { addMsg("assistant", "错误: " + e.message); }
}

// ── Generate with line-by-line display ──
async function generateDraft() {
  if (!sessionId || isProcessing) return;
  isProcessing = true;
  showAgentStep("draft");
  addMsg("assistant", "正在生成合同草稿...");

  try {
    const res = await fetch("/api/smart-draft/generate", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, guidanceMode: "llm_infer" }),
    });
    const txt = await res.text();
    let data; try { data = JSON.parse(txt); } catch { throw new Error("服务器返回异常，请重试"); }
    if (data.error) throw new Error(data.error);

    currentDraft = typeof data.draft === "string" ? data.draft : data.draft?.content || JSON.stringify(data.draft);
    currentDraft = currentDraft.replace(/^```[\s\S]*?\n/gm, "").replace(/\n```$/gm, "");

    // Line-by-line display in Word area
    await typeDraftLineByLine(currentDraft);

    showAgentStep("done");
    if (data.messages) renderMessages(data.messages);
    addMsg("assistant", `草稿已生成${data.appliedRules?`，应用了 ${data.appliedRules.length} 条规则`:""}`);
    document.getElementById("statusTag").textContent = "已完成";
  } catch(e) {
    addMsg("assistant", "生成失败: " + e.message);
  }
  isProcessing = false;
}

async function typeDraftLineByLine(text) {
  const container = document.getElementById("docxContent");
  container.innerHTML = "";
  const lines = text.split("\n");
  let accumulated = "";

  for (let i = 0; i < lines.length; i++) {
    accumulated += (i > 0 ? "\n" : "") + lines[i];
    // Render as formatted HTML
    container.innerHTML = renderContractHTML(accumulated);
    // Scroll to bottom
    const area = document.getElementById("wordArea");
    area.scrollTop = area.scrollHeight;
    // Delay for visual effect
    await delay(20);
  }
}

function renderContractHTML(text) {
  // Convert markdown-like formatting to HTML
  return text
    .replace(/^# (.+)$/gm, '<h1 style="text-align:center;font-size:20px;margin:20px 0 12px">$1</h1>')
    .replace(/^## (.+)$/gm, '<h2 style="font-size:16px;margin:16px 0 8px;border-bottom:1px solid #eee;padding-bottom:4px">$1</h2>')
    .replace(/^### (.+)$/gm, '<h3 style="font-size:14px;margin:12px 0 6px">$1</h3>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\n/g, '<br>');
}

function delay(ms) { return new Promise(r => setTimeout(r, ms)); }

// ── Attach files ──
function handleAttachFiles(files) {
  for (const f of files) attachedFiles.push(f.name);
  addMsg("user", "已上传: " + attachedFiles.join(", "));
}

// ── Download / Copy ──
function downloadDraft() {
  if (!currentDraft) return;
  const b = new Blob([currentDraft], { type: "text/plain;charset=utf-8" });
  const a = document.createElement("a"); a.href = URL.createObjectURL(b);
  a.download = `合同草稿_${new Date().toISOString().slice(0,10)}.txt`; a.click();
}

function copyDraft() {
  if (!currentDraft) return;
  navigator.clipboard.writeText(currentDraft).then(() => alert("已复制"));
}

function esc(s) { return (s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\n/g,"<br>"); }
