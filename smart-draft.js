// smart-draft.js — Smart Draft with progress + AI status + completion

const API_BASE = "";
let editor = null;
let sessionId = "";
let knowledgeMode = "industry";
let uploadedFiles = [];
let currentQuestion = null;
let allAnswers = {};
let currentDraft = "";
let isProcessing = false;

const PROGRESS_STEPS = [
  { key: "intent", label: "分析合同意图..." },
  { key: "extract", label: "提取参考数据..." },
  { key: "knowledge", label: "检索法律法规..." },
  { key: "generate", label: "生成合同草稿..." },
  { key: "iterate", label: "完善合同细节..." },
  { key: "done", label: "完成" },
];

// ── Init ──
window.addEventListener("DOMContentLoaded", () => {
  setIcons();
  checkAIStatus();
  // KB: 通用法规 is default active (set in HTML)
});

function setIcons() {
  const m = {
    navTF: "templateFill", navSD: "smartDraft", navCR: "contractReview", navKB: "knowledge",
    icoDL: "download", icoCP: "copy", icoStart: "sparkles", icoClip: "paperclip", icoSend: "send",
  };
  for (const [id, name] of Object.entries(m)) {
    const el = document.getElementById(id);
    if (el) el.innerHTML = AppIcons[name] || "";
  }
}

// ── AI Status ──
async function checkAIStatus() {
  try {
    const res = await fetch(`${API_BASE}/api/health`);
    const d = await res.json();
    document.getElementById("aiModel").textContent = d.model || "unknown";
    document.getElementById("aiProvider").textContent = d.provider || "";
    document.getElementById("aiDot").style.background = d.apiKeyConfigured ? "var(--green)" : "var(--red)";
  } catch {
    document.getElementById("aiModel").textContent = "离线";
    document.getElementById("aiDot").style.background = "var(--red)";
  }
}

// ── Knowledge Toggle ──
function toggleKB(el) {
  el.classList.toggle("active");
  const active = document.querySelectorAll(".kb-chip.active");
  knowledgeMode = active.length > 0 ? active[0].dataset.kb : "industry";
}

// ── Files ──
async function handleFiles(fileList) {
  for (const file of fileList) {
    const text = await readText(file);
    uploadedFiles.push({ name: file.name, text: text.slice(0, 10000), size: file.size });
    addFileChip(file.name);
  }
}
function readText(file) {
  return new Promise((resolve) => {
    if (file.name.match(/\.(docx|doc|pdf)$/i)) { resolve(`[File: ${file.name}]`); return; }
    const r = new FileReader(); r.onload = () => resolve(r.result || ""); r.onerror = () => resolve(""); r.readAsText(file);
  });
}
function addFileChip(name) {
  const chip = document.createElement("span");
  chip.className = "tag tag-accent"; chip.textContent = name; chip.style.cursor = "pointer";
  chip.onclick = () => chip.remove();
  document.getElementById("fileChips").appendChild(chip);
}

// ── Completion Bar ──
function setCompletion(pct, label) {
  document.getElementById("completionBar").style.display = "flex";
  document.getElementById("completionPct").textContent = pct + "%";
  document.getElementById("completionFill").style.width = pct + "%";
  document.getElementById("completionLabel").textContent = label || "";
}

// ── AI Progress ──
function showProgress() {
  document.getElementById("aiProgressWrap").style.display = "block";
}
function setStep(stepKey) {
  const idx = PROGRESS_STEPS.findIndex((s) => s.key === stepKey);
  PROGRESS_STEPS.forEach((s, i) => {
    const el = document.getElementById("sp-" + s.key);
    if (!el) return;
    el.className = "step";
    if (i < idx) el.classList.add("done");
    else if (i === idx) el.classList.add("active");
  });
  const step = PROGRESS_STEPS[idx];
  document.getElementById("aiProgressLabel").textContent = step ? step.label : "";
  // Update completion based on step
  const pct = Math.min(90, Math.round((idx / (PROGRESS_STEPS.length - 1)) * 100));
  setCompletion(pct, step ? step.label : "");
}

// ── Chat ──
function addMsg(role, content) {
  const thread = document.getElementById("chatThread");
  const div = document.createElement("div");
  div.className = `chat-msg ${role}`;
  div.innerHTML = `<div class="avatar">${role === "assistant" ? "AI" : "U"}</div><div class="bubble">${esc(content)}</div>`;
  thread.appendChild(div);
  thread.scrollTop = thread.scrollHeight;
}
function renderMessages(messages) {
  const thread = document.getElementById("chatThread");
  thread.innerHTML = "";
  (messages || []).forEach((m) => {
    const div = document.createElement("div");
    div.className = `chat-msg ${m.role}`;
    div.innerHTML = `<div class="avatar">${m.role === "assistant" ? "AI" : "U"}</div><div class="bubble">${esc(m.content)}</div>`;
    thread.appendChild(div);
  });
  thread.scrollTop = thread.scrollHeight;
}
function esc(s) { return (s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br>"); }

// ── Start ──
async function startDraft() {
  const desc = document.getElementById("descriptionInput").value.trim();
  if (!desc && uploadedFiles.length === 0) return alert("请输入需求或上传文件");
  if (isProcessing) return;
  isProcessing = true;

  document.getElementById("startView").style.display = "none";
  document.getElementById("editorView").style.display = "block";
  document.getElementById("editorActions").style.display = "flex";
  document.getElementById("sessionTag").style.display = "inline-flex";

  if (!editor) editor = new OfficeEditor(document.getElementById("editorContainer"));

  showProgress();
  setStep("intent");

  try {
    const res = await fetch(`${API_BASE}/api/smart-draft/init`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ description: desc, knowledgeMode, fileTexts: uploadedFiles.map((f) => ({ name: f.name, text: f.text })) }),
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    sessionId = data.sessionId;
    allAnswers = data.answers || {};

    if (data.dialogue?.missingCount > 0) {
      setStep("iterate");
      setCompletion(40, `已识别${data.rulesCount || 0}条规则，还需补充${data.dialogue.missingCount}项`);
    } else {
      setStep("generate");
      setCompletion(50, "信息齐全，准备生成");
    }

    if (data.messages) renderMessages(data.messages);
    if (data.extractedData?.items?.length > 0) addMsg("assistant", `已提取 ${data.extractedData.items.length} 条数据`);

    if (data.dialogue?.missingCount > 0) {
      currentQuestion = data.dialogue.currentQuestion;
      showGuideButtons();
    } else {
      showGenerateBtn();
    }
  } catch (e) {
    addMsg("assistant", `错误: ${e.message}`);
    setCompletion(0, "出错");
  }
  isProcessing = false;
}

// ── Guide Buttons ──
function showGuideButtons() {
  const thread = document.getElementById("chatThread");
  const div = document.createElement("div");
  div.className = "chat-msg assistant";
  div.innerHTML = `
    <div class="avatar">AI</div>
    <div>
      <div class="bubble">${esc(currentQuestion?.question || "请补充信息")}</div>
      <div class="guide-options">
        <button class="guide-btn" onclick="pickGuide('ask_user')"><span class="guide-icon">${AppIcons.send}</span>我来回答</button>
        <button class="guide-btn" onclick="pickGuide('kb_search')"><span class="guide-icon">${AppIcons.search}</span>知识库检索</button>
        <button class="guide-btn" onclick="pickGuide('llm_infer')"><span class="guide-icon">${AppIcons.sparkles}</span>AI 推断</button>
      </div>
    </div>`;
  thread.appendChild(div);
  thread.scrollTop = thread.scrollHeight;
}

function showGenerateBtn() {
  const thread = document.getElementById("chatThread");
  const div = document.createElement("div");
  div.className = "chat-msg assistant";
  div.innerHTML = `<div class="avatar">AI</div><div><div class="bubble">信息已齐全。</div><div style="padding-top:8px"><button class="btn btn-primary btn-sm" onclick="generateDraft()"><span class="app-icon" style="width:14px;height:14px">${AppIcons.sparkles}</span> 生成合同草稿</button></div></div>`;
  thread.appendChild(div);
  thread.scrollTop = thread.scrollHeight;
}

async function pickGuide(mode) {
  document.querySelectorAll(".guide-options").forEach((el) => el.parentElement.removeChild(el));
  if (mode === "ask_user") { document.getElementById("chatInput").focus(); return; }
  addMsg("user", mode === "kb_search" ? "从知识库查找" : "AI 自动推断");
  if (mode === "llm_infer") { showProgress(); setStep("knowledge"); addMsg("assistant", "正在检索知识库..."); }

  try {
    const res = await fetch(`${API_BASE}/api/smart-draft/guide`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, mode }),
    });
    const data = await res.json();
    if (data.messages) renderMessages(data.messages);
    if (data.answers) allAnswers = data.answers;
    if (data.dialogue?.missingCount > 0) { currentQuestion = data.dialogue.currentQuestion; showGuideButtons(); }
    else showGenerateBtn();
  } catch (e) { addMsg("assistant", `错误: ${e.message}`); }
}

// ── Send Answer ──
async function sendAnswer() {
  const input = document.getElementById("chatInput");
  const value = input.value.trim();
  if (!value || !sessionId) return;
  input.value = "";
  const field = currentQuestion?.hint || "补充信息";
  addMsg("user", value);

  try {
    const res = await fetch(`${API_BASE}/api/smart-draft/answer`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, field, value }),
    });
    const data = await res.json();
    if (data.messages) renderMessages(data.messages);
    if (data.answers) allAnswers = data.answers;
    if (data.dialogue?.missingCount > 0) { currentQuestion = data.dialogue.currentQuestion; showGuideButtons(); }
    else showGenerateBtn();
  } catch (e) { addMsg("assistant", `错误: ${e.message}`); }
}

// ── Generate ──
async function generateDraft() {
  if (!sessionId || isProcessing) return;
  isProcessing = true;
  showProgress();
  setStep("generate");
  setCompletion(60, "正在生成合同草稿...");
  addMsg("assistant", "正在生成合同草稿，请稍候...");

  try {
    const res = await fetch(`${API_BASE}/api/smart-draft/generate`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, guidanceMode: "llm_infer" }),
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    currentDraft = typeof data.draft === "string" ? data.draft : data.draft?.content || JSON.stringify(data.draft);
    currentDraft = currentDraft.replace(/^```[\s\S]*?\n/gm, "").replace(/\n```$/gm, "");

    // Type draft line by line into editor
    if (editor) {
      const lines = currentDraft.split("\n");
      let accumulated = "";
      for (let i = 0; i < lines.length; i++) {
        accumulated += (i > 0 ? "\n" : "") + lines[i];
        editor.setContent(accumulated);
        // Small delay for visual effect
        if (i < lines.length - 1) await delay(30);
      }
    }

    setStep("done");
    setCompletion(100, "草稿已完成");
    if (data.messages) renderMessages(data.messages);
    addMsg("assistant", `草稿已生成${data.appliedRules ? `，应用了 ${data.appliedRules.length} 条规则` : ""}`);
    document.getElementById("sessionTag").textContent = "已完成";
  } catch (e) {
    addMsg("assistant", `生成失败: ${e.message}`);
    setCompletion(0, "生成失败");
  }
  isProcessing = false;
}

function delay(ms) { return new Promise((r) => setTimeout(r, ms)); }

// ── Download / Copy ──
function downloadDraft() {
  if (!currentDraft) return;
  const blob = new Blob([currentDraft], { type: "text/plain;charset=utf-8" });
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
  a.download = `合同草稿_${new Date().toISOString().slice(0, 10)}.txt`; a.click();
}
function copyDraft() {
  if (!currentDraft) return;
  navigator.clipboard.writeText(currentDraft).then(() => alert("已复制"));
}
