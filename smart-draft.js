// smart-draft.js — Doubao-style: Big input → Split view → AI dialogue + Word

let sessionId = "";
let knowledgeMode = "industry";
let currentDraft = "";
let selectedTpl = null;
let currentQuestion = null;
let isProcessing = false;
let mainFiles = [];
let draftFiles = [];

const STEPS = ["intent","match","kb","draft","done"];
const STEP_MSG = { intent:"分析意图", match:"匹配模板", kb:"检索知识库", draft:"生成草稿", done:"完成" };

// ── Init ──
window.addEventListener("DOMContentLoaded", () => {
  const icons = { navSD:"smartDraft", navTF:"templateFill", navCR:"contractReview", navKB:"knowledge", icoClip:"paperclip", icoClip2:"paperclip", icoSend:"send" };
  for (const [id,n] of Object.entries(icons)) { const el=document.getElementById(id); if(el&&window.AppIcons) el.innerHTML=AppIcons[n]||""; }
  fetchStatus();
  loadQuickTpls();
  initModelTest();
});

async function fetchStatus() {
  try {
    const d = await (await fetch("/api/health")).json();
    document.getElementById("aiModel").textContent = d.model || "unknown";
    document.getElementById("aiProvider").textContent = d.provider || "";
    document.getElementById("aiDot").style.background = d.apiKeyConfigured ? "var(--green)" : "var(--red)";
  } catch { document.getElementById("aiModel").textContent="离线"; document.getElementById("aiDot").style.background="var(--red)"; }
}

function fillExample(text) {
  document.getElementById("mainInput").value = text;
  document.getElementById("mainInput").focus();
}

function toggleKB(el) { el.classList.toggle("on"); }

// ── Quick Templates ──
async function loadQuickTpls() {
  try {
    const tpls = await (await fetch("/api/smart-draft/templates")).json();
    document.getElementById("quickTpls").innerHTML = tpls.slice(0, 8).map(t =>
      `<span class="quick-tpl" onclick="toggleTpl(this,'${t.id}')" data-id="${t.id}">${t.name}</span>`
    ).join("");
  } catch {}
}

function toggleTpl(el, id) {
  document.querySelectorAll(".quick-tpl").forEach(e => e.classList.remove("active"));
  if (selectedTpl?.id === id) { selectedTpl = null; return; }
  el.classList.add("active");
  // Store tpl info (will be used in startDraft)
  selectedTpl = { id };
}

// ── Files ──
function addFiles(files) {
  for (const f of files) mainFiles.push({ name: f.name, file: f });
  renderAttached();
}
function renderAttached() {
  document.getElementById("attachedFiles").innerHTML = mainFiles.map((f,i) =>
    `<span class="attached-tag">${f.name} <span class="rm" onclick="removeFile(${i})">✕</span></span>`
  ).join("");
}
function removeFile(i) { mainFiles.splice(i,1); renderAttached(); }

function addDraftFiles(files) {
  for (const f of files) draftFiles.push({ name: f.name, file: f });
  document.getElementById("draftFiles").innerHTML = draftFiles.map(f =>
    `<span class="attached-tag">${f.name}</span>`
  ).join("");
}

// ── Start Draft ──
async function startDraft() {
  const desc = document.getElementById("mainInput").value.trim();
  const tplHint = selectedTpl ? `使用${selectedTpl.id}模板` : "";
  const fullDesc = [tplHint, desc].filter(Boolean).join("，");
  if (!fullDesc && mainFiles.length === 0) return alert("请输入需求或上传文件");
  if (isProcessing) return;
  isProcessing = true;

  document.getElementById("goBtn").disabled = true;
  document.getElementById("goBtn").textContent = "分析中...";

  // Switch to draft view
  document.getElementById("startPage").style.display = "none";
  document.getElementById("draftPage").style.display = "flex";
  document.getElementById("statusTag").style.display = "inline-flex";
  document.getElementById("agentStrip").style.display = "flex";

  showStep("intent");
  addBubble("assistant", "正在分析你的合同需求...");

  // Show user's input and files in chat
  if (desc) addMsg("user", desc);
  if (mainFiles.length > 0) addMsg("user", "已上传 " + mainFiles.map(f => f.name).join("、"));

  try {
    // Read file texts for API
    const fileTexts = [];
    for (const f of mainFiles) {
      try {
        const text = await f.file.text();
        fileTexts.push({ name: f.name, text: text.slice(0, 8000) });
      } catch { fileTexts.push({ name: f.name, text: "" }); }
    }

    // Step 1: Init session
    addBubble("assistant", "正在匹配合同模板...");
    const res = await fetch("/api/smart-draft/init", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ description: fullDesc, knowledgeMode, fileTexts }),
    });
    const txt = await res.text();
    let data; try { data = JSON.parse(txt); } catch { throw new Error("服务器响应异常"); }
    if (data.error) throw new Error(data.error);

    sessionId = data.sessionId;
    showStep("match");

    if (data.intent?.contractTypeCn) {
      document.getElementById("docTitle").textContent = data.intent.contractTypeCn;
      addBubble("assistant", `已识别为「${data.intent.contractTypeCn}」，匹配到 ${data.rulesCount || 0} 条规则`);
    }

    // Step 2: Immediately generate draft
    showStep("kb");
    addBubble("assistant", "正在检索知识库...");
    await delay(300);

    showStep("draft");
    addBubble("assistant", "正在生成合同草稿...");

    const genRes = await fetch("/api/smart-draft/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, guidanceMode: "llm_infer" }),
    });
    const genTxt = await genRes.text();
    let genData; try { genData = JSON.parse(genTxt); } catch { throw new Error("生成失败"); }
    if (genData.error) throw new Error(genData.error);

    currentDraft = typeof genData.draft === "string" ? genData.draft : genData.draft?.content || "";
    currentDraft = currentDraft.replace(/^```[\s\S]*?\n/gm, "").replace(/\n```$/gm, "");

    // Display draft line by line
    await typeLineByLine(currentDraft);

    showStep("done");
    addBubble("assistant", `合同草稿已生成${genData.appliedRules ? `，应用了 ${genData.appliedRules.length} 条规则` : ""}。你可以在右侧对话中补充信息，AI 会帮你完善合同。`);

    document.getElementById("statusTag").textContent = "已完成";
    document.getElementById("statusTag").className = "tag tag-green";

    // Show generate button for re-generation
    showGenBtn();

  } catch(e) {
    addBubble("assistant", "错误: " + e.message);
    showStep("done");
  }
  isProcessing = false;
  document.getElementById("goBtn").disabled = false;
  document.getElementById("goBtn").textContent = "开始起草";
}

// ── Agent Steps ──
function showStep(step) {
  const idx = STEPS.indexOf(step);
  STEPS.forEach((s,i) => {
    const el = document.getElementById("as-"+s);
    if(!el) return;
    el.className = "s";
    if (i < idx) el.classList.add("done");
    else if (i === idx) el.classList.add("active");
  });
  document.getElementById("agentMsg").textContent = STEP_MSG[step] || step;
}

// ── Chat ──
function addBubble(role, text) {
  const t = document.getElementById("chatThread");
  const d = document.createElement("div");
  d.className = `chat-msg ${role}`;
  d.innerHTML = `<div class="avatar">${role==="assistant"?"AI":"U"}</div><div class="bubble">${esc(text)}</div>`;
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

function appendMessages(msgs) {
  const t = document.getElementById("chatThread");
  (msgs||[]).forEach(m => {
    const d = document.createElement("div");
    d.className = `chat-msg ${m.role}`;
    d.innerHTML = `<div class="avatar">${m.role==="assistant"?"AI":"U"}</div><div class="bubble">${esc(m.content)}</div>`;
    t.appendChild(d);
  });
  t.scrollTop = t.scrollHeight;
}

// ── Options ──
function showOptRow() { document.getElementById("optRow").style.display = "flex"; }
function hideOptRow() { document.getElementById("optRow").style.display = "none"; }

async function pickOpt(mode) {
  hideOptRow();
  if (mode === "ask") { document.getElementById("chatInput").focus(); return; }

  const labels = { kb: "从知识库检索", ai: "AI 自动推断" };
  addBubble("user", labels[mode]);

  if (mode === "ai") { showStep("draft"); addBubble("assistant", "正在推断..."); }

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
      showOptRow();
      addBubble("assistant", currentQuestion?.question || "请继续补充");
    } else {
      showStep("draft");
      showGenBtn();
    }
  } catch(e) { addBubble("assistant", "错误: " + e.message); }
}

async function sendAnswer() {
  const input = document.getElementById("chatInput");
  const val = input.value.trim();
  if (!val || !sessionId) return;
  input.value = "";
  hideOptRow();

  const field = currentQuestion?.hint || "补充信息";
  addBubble("user", val);

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
      showOptRow();
      addBubble("assistant", currentQuestion?.question || "请继续补充");
    } else {
      showStep("draft");
      showGenBtn();
    }
  } catch(e) { addBubble("assistant", "错误: " + e.message); }
}

function showGenBtn() {
  const t = document.getElementById("chatThread");
  const d = document.createElement("div");
  d.className = "chat-msg assistant";
  d.innerHTML = `<div class="avatar">AI</div><div><div class="bubble">信息已齐全，可以生成。</div><div style="padding-top:8px"><button class="btn btn-primary btn-sm" onclick="generateDraft()">生成合同草稿</button></div></div>`;
  t.appendChild(d); t.scrollTop = t.scrollHeight;
}

// ── Generate: line by line ──
async function generateDraft() {
  if (!sessionId || isProcessing) return;
  isProcessing = true;
  showStep("draft");
  addBubble("assistant", "正在生成合同草稿...");

  try {
    const res = await fetch("/api/smart-draft/generate", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, guidanceMode: "llm_infer" }),
    });
    const txt = await res.text();
    let data; try { data = JSON.parse(txt); } catch { throw new Error("服务器返回异常"); }
    if (data.error) throw new Error(data.error);

    currentDraft = typeof data.draft === "string" ? data.draft : data.draft?.content || JSON.stringify(data.draft);
    currentDraft = currentDraft.replace(/^```[\s\S]*?\n/gm, "").replace(/\n```$/gm, "");

    // Line by line into Word area
    await typeLineByLine(currentDraft);

    showStep("done");
    if (data.messages) renderMessages(data.messages);
    addBubble("assistant", `草稿已生成${data.appliedRules?`（${data.appliedRules.length} 条规则）`:""}`);
    document.getElementById("statusTag").textContent = "已完成";
  } catch(e) {
    addBubble("assistant", "生成失败: " + e.message);
  }
  isProcessing = false;
}

async function typeLineByLine(text) {
  const el = document.getElementById("contractContent");
  el.innerHTML = "";
  const lines = text.split("\n");
  let acc = "";
  for (let i = 0; i < lines.length; i++) {
    acc += (i > 0 ? "\n" : "") + lines[i];
    el.innerHTML = renderDoc(acc) + '<span class="typing-cursor"></span>';
    document.getElementById("docBody").scrollTop = document.getElementById("docBody").scrollHeight;
    await new Promise(r => setTimeout(r, 15));
  }
  el.innerHTML = renderDoc(text); // final render without cursor
}

function renderDoc(text) {
  return text
    .replace(/^# (.+)$/gm, '<h1>$1</h1>')
    .replace(/^## (.+)$/gm, '<h2>$1</h2>')
    .replace(/^### (.+)$/gm, '<h3>$1</h3>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\n/g, '<br>');
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
function delay(ms) { return new Promise(r => setTimeout(r, ms)); }
