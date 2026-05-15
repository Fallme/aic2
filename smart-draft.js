// smart-draft.js — Smart Draft v3: Real-time generation + 3 reference answers

let sessionId = "";
let knowledgeMode = "industry";
let currentDraft = "";
let selectedTpl = null;
let currentQuestion = null;
let isProcessing = false;
let mainFiles = [];

const STEPS = ["intent","match","kb","draft","done"];
const STEP_MSG = { intent:"分析意图", match:"匹配模板", kb:"检索知识库", draft:"生成草稿", done:"完成" };

// ── Init ──
window.addEventListener("DOMContentLoaded", () => {
  const icons = { navSD:"smartDraft", navCR:"contractReview", navCC:"layers", navCT:"templateFill", navKB:"knowledge", icoClip:"paperclip", icoSend:"send" };
  for (const [id,n] of Object.entries(icons)) { const el=document.getElementById(id); if(el&&window.AppIcons) el.innerHTML=AppIcons[n]||""; }
  fetchStatus();
  initModelTest();
  // File input
  const fi = document.getElementById("mainFileInput");
  if (fi) fi.addEventListener("change", () => { if(fi.files[0]) addFiles(fi.files); fi.value=""; });
});

function fillExample(text) {
  document.getElementById("mainInput").value = text;
  document.getElementById("mainInput").focus();
}

function toggleKB(el) { el.classList.toggle("on"); }

async function fetchStatus() {
  try {
    const d = await (await fetch("/api/health")).json();
    document.getElementById("aiModel").textContent = d.model || "unknown";
    document.getElementById("aiProvider").textContent = d.provider || "";
    document.getElementById("aiDot").style.background = d.apiKeyConfigured ? "var(--green)" : "var(--red)";
  } catch { document.getElementById("aiModel").textContent="离线"; document.getElementById("aiDot").style.background="var(--red)"; }
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

  // Show user's input in chat
  if (desc) addMsg("user", desc);
  if (mainFiles.length > 0) addMsg("user", "已上传 " + mainFiles.map(f=>f.name).join("、"));

  // Show loading animation in Word area
  showLoadingAnimation("正在分析合同需求...");

  try {
    const fileTexts = [];
    for (const f of mainFiles) {
      try { const text = await f.file.text(); fileTexts.push({ name: f.name, text: text.slice(0, 8000) }); }
      catch { fileTexts.push({ name: f.name, text: "" }); }
    }

    // Step 1: Init with typing indicator
    addTypingIndicator();
    const res = await fetch("/api/smart-draft/init", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ description: fullDesc, knowledgeMode, fileTexts }),
    });
    const txt = await res.text();
    let data; try { data = JSON.parse(txt); } catch { throw new Error("服务器响应异常"); }
    if (data.error) throw new Error(data.error);

    sessionId = data.sessionId;
    showStep("match");

    if (data.intent?.contractTypeCn) {
      document.getElementById("docTitle").textContent = data.intent.contractTypeCn;
      addBubble("assistant", `已识别为「${data.intent.contractTypeCn}」，匹配到 ${data.rulesCount||0} 条规则`);
    }

    // Step 2: Knowledge retrieval
    showStep("kb");
    removeTypingIndicator();
    addTypingIndicator();
    showLoadingAnimation("正在检索知识库...");
    await delay(500);

    // Step 3: Generate draft
    showStep("draft");
    removeTypingIndicator();
    addTypingIndicator();
    showLoadingAnimation("正在生成合同草稿...");

    const genRes = await fetch("/api/smart-draft/generate", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, guidanceMode: "llm_infer" }),
    });
    const genTxt = await genRes.text();
    let genData; try { genData = JSON.parse(genTxt); } catch { throw new Error("生成失败，请重试"); }
    if (genData.error) throw new Error(genData.error);

    currentDraft = typeof genData.draft === "string" ? genData.draft : genData.draft?.content || "";
    currentDraft = currentDraft.replace(/^```[\s\S]*?\n/gm, "").replace(/\n```$/gm, "");

    // Display draft line by line at slower speed for readability
    await typeLineByLine(currentDraft, 60);

    showStep("done");
    addBubble("assistant", `合同草稿已生成。你可以在右侧对话中补充信息，AI 会帮你完善合同。`);

    document.getElementById("statusTag").textContent = "已完成";
    document.getElementById("statusTag").className = "tag tag-green";

  } catch(e) {
    addBubble("assistant", "错误: " + e.message);
    showStep("done");
    document.getElementById("docxContent").innerHTML = `<div style="padding:40px;text-align:center;color:var(--red)">${esc(e.message)}</div>`;
  }
  isProcessing = false;
  document.getElementById("goBtn").disabled = false;
  document.getElementById("goBtn").textContent = "开始起草";
}

// ── Loading Animation ──
function showLoadingAnimation(text) {
  const el = document.getElementById("docxContent");
  el.innerHTML = `<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;padding:60px">
    <div style="width:48px;height:48px;border:3px solid var(--gray5);border-top-color:var(--gold);border-radius:50%;animation:spin .8s linear infinite;margin-bottom:20px"></div>
    <div style="font-size:16px;color:var(--ink);font-weight:500;margin-bottom:8px">${esc(text)}</div>
    <div style="font-size:13px;color:var(--muted)">AI 正在工作，请稍候...</div>
    <div style="margin-top:16px;display:flex;gap:6px;align-items:center">
      <div class="typing-dots"><span></span><span></span><span></span></div>
      <span style="font-size:12px;color:var(--subtle)">思考中</span>
    </div>
  </div>`;
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
function addMsg(role, text) { addBubble(role, text); }
function addTypingIndicator() {
  const t = document.getElementById("chatThread");
  const d = document.createElement("div");
  d.className = "chat-msg assistant";
  d.id = "typingIndicator";
  d.innerHTML = `<div class="avatar">AI</div><div class="bubble"><div class="typing-dots"><span></span><span></span><span></span></div></div>`;
  t.appendChild(d); t.scrollTop = t.scrollHeight;
}

function removeTypingIndicator() {
  const el = document.getElementById("typingIndicator");
  if (el) el.remove();
}

function addBubble(role, text) {
  const t = document.getElementById("chatThread");
  const d = document.createElement("div");
  d.className = `chat-msg ${role}`;
  d.innerHTML = `<div class="avatar">${role==="assistant"?"AI":"U"}</div><div class="bubble">${esc(text)}</div>`;
  t.appendChild(d); t.scrollTop = t.scrollHeight;
}

function appendMessages(msgs) {
  const t = document.getElementById("chatThread");
  (msgs||[]).forEach(m => {
    // Skip if this message already exists (prevent duplicates)
    const existing = t.querySelectorAll(".bubble");
    const lastBubble = existing.length > 0 ? existing[existing.length-1].textContent : "";
    if (m.content && m.content === lastBubble) return;
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

// ── Three Reference Answers (shown in chat) ──
function showRefAnswers(question) {
  if (!question) return;
  const t = document.getElementById("chatThread");
  const d = document.createElement("div");
  d.className = "chat-msg assistant";
  const hint = question.hint || "信息";
  d.innerHTML = `<div class="avatar">AI</div><div>
    <div class="bubble" style="margin-bottom:8px">参考回答：</div>
    <div style="display:flex;gap:6px;flex-wrap:wrap">
      <button class="ref-btn" onclick="pickRef('${esc(hint)}','kb')">📚 知识库推荐</button>
      <button class="ref-btn" onclick="pickRef('${esc(hint)}','common')">💡 常用答案</button>
      <button class="ref-btn" onclick="pickRef('${esc(hint)}','ai')">🤖 AI 生成</button>
    </div>
  </div>`;
  t.appendChild(d); t.scrollTop = t.scrollHeight;
}

async function pickOpt(mode) {
  if (mode === "ask") { document.getElementById("chatInput").focus(); return; }
  addMsg("user", mode === "kb" ? "从知识库检索" : "AI 自动推断");
  addTypingIndicator();

  try {
    const res = await fetch("/api/smart-draft/guide", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, mode: mode === "kb" ? "kb_search" : "llm_infer" }),
    });
    const txt = await res.text();
    let data; try { data = JSON.parse(txt); } catch { throw new Error("服务器异常"); }
    removeTypingIndicator();
    if (data.messages) appendMessages(data.messages);
    if (data.dialogue?.missingCount > 0) {
      currentQuestion = data.dialogue.currentQuestion;
      showRefAnswers(currentQuestion);
      addBubble("assistant", currentQuestion?.question || "请继续补充");
    } else {
      addBubble("assistant", "信息已齐全，可以生成草稿。");
    }
  } catch(e) { removeTypingIndicator(); addBubble("assistant", "错误: " + e.message); }
}

async function pickRef(field, type) {
  const labels = { kb: "📚 知识库推荐", common: "💡 常用答案", ai: "🤖 AI 生成" };
  addMsg("user", labels[type] || "选择参考回答");
  addTypingIndicator();
  try {
    const res = await fetch("/api/smart-draft/guide", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, mode: type === "kb" ? "kb_search" : "llm_infer" }),
    });
    const txt = await res.text();
    let data; try { data = JSON.parse(txt); } catch { throw new Error("服务器异常"); }
    removeTypingIndicator();
    if (data.messages) appendMessages(data.messages);
    if (data.inferred) addBubble("assistant", `已填写「${field}」: ${data.inferred}`);
    if (data.dialogue?.missingCount > 0) {
      currentQuestion = data.dialogue.currentQuestion;
      showRefAnswers(currentQuestion);
      addBubble("assistant", currentQuestion?.question || "请继续补充");
    } else {
      addBubble("assistant", "信息已齐全，可以生成草稿。");
    }
  } catch(e) { addBubble("assistant", "错误: " + e.message); }
}

// ── Send Answer ──
async function sendAnswer() {
  const input = document.getElementById("chatInput");
  const val = input.value.trim();
  if (!val || !sessionId) return;
  input.value = "";
  document.getElementById("optRow").style.display = "none";

  const field = currentQuestion?.hint || "补充信息";
  addMsg("user", val);
  addTypingIndicator();

  try {
    const res = await fetch("/api/smart-draft/answer", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, field, value: val }),
    });
    const txt = await res.text();
    let data; try { data = JSON.parse(txt); } catch { throw new Error("服务器异常"); }
    removeTypingIndicator();
    if (data.messages) appendMessages(data.messages);
    if (data.dialogue?.missingCount > 0) {
      currentQuestion = data.dialogue.currentQuestion;
      showRefAnswers(currentQuestion);
      addBubble("assistant", currentQuestion?.question || "请继续补充");
    } else {
      addBubble("assistant", "信息已齐全，可以生成草稿。");
    }
  } catch(e) { addBubble("assistant", "错误: " + e.message); }
}

// ── Generate ──
async function generateDraft() {
  if (!sessionId || isProcessing) return;
  isProcessing = true;
  showStep("draft");
  addBubble("assistant", "正在生成合同草稿...");
  showLoadingAnimation("正在生成合同草稿...");

  try {
    const res = await fetch("/api/smart-draft/generate", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, guidanceMode: "llm_infer" }),
    });
    const txt = await res.text();
    let data; try { data = JSON.parse(txt); } catch { throw new Error("服务器返回异常"); }
    if (data.error) throw new Error(data.error);

    currentDraft = typeof data.draft === "string" ? data.draft : data.draft?.content || "";
    currentDraft = currentDraft.replace(/^```[\s\S]*?\n/gm, "").replace(/\n```$/gm, "");

    await typeLineByLine(currentDraft, 30);

    showStep("done");
    addBubble("assistant", `草稿已生成${data.appliedRules ? `，应用了 ${data.appliedRules.length} 条规则` : ""}`);
    document.getElementById("statusTag").textContent = "已完成";
    document.getElementById("statusTag").className = "tag tag-green";
  } catch(e) {
    addBubble("assistant", "生成失败: " + e.message);
    document.getElementById("docxContent").innerHTML = `<div style="padding:40px;text-align:center;color:var(--red)">${esc(e.message)}</div>`;
  }
  isProcessing = false;
}

// ── Line by line typing ──
async function typeLineByLine(text, speed) {
  const el = document.getElementById("docxContent");
  el.innerHTML = "";
  const lines = text.split("\n");
  let acc = "";
  for (let i = 0; i < lines.length; i++) {
    acc += (i > 0 ? "\n" : "") + lines[i];
    el.innerHTML = renderDoc(acc) + '<span class="typing-cursor"></span>';
    document.getElementById("docBody").scrollTop = document.getElementById("docBody").scrollHeight;
    await delay(speed || 30);
  }
  el.innerHTML = renderDoc(text);
}

function renderDoc(text) {
  return text
    .replace(/^# (.+)$/gm, '<h1 style="text-align:center;font-size:20px;margin:20px 0 12px;font-family:SimHei,sans-serif">$1</h1>')
    .replace(/^## (.+)$/gm, '<h2 style="font-size:16px;margin:16px 0 8px;border-bottom:1px solid #eee;padding-bottom:4px;font-family:SimHei,sans-serif">$1</h2>')
    .replace(/^### (.+)$/gm, '<h3 style="font-size:14px;margin:12px 0 6px;font-family:SimHei,sans-serif">$1</h3>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\n/g, '<br>');
}

function delay(ms) { return new Promise(r => setTimeout(r, ms)); }
function esc(s) { return (s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\n/g,"<br>"); }
