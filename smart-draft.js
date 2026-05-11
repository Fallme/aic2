// smart-draft.js — Smart Draft with new UI

const API_BASE = "";
let editor = null;
let sessionId = "";
let knowledgeMode = "industry";
let uploadedFiles = [];
let currentQuestion = null;
let allAnswers = {};
let currentDraft = "";
let isProcessing = false;

// ── Init ──
window.addEventListener("DOMContentLoaded", () => {
  // Set icons
  document.getElementById("navTemplateFill").innerHTML = AppIcons.templateFill;
  document.getElementById("navSmartDraft").innerHTML = AppIcons.smartDraft;
  document.getElementById("navReview").innerHTML = AppIcons.contractReview;
  document.getElementById("navKnowledge").innerHTML = AppIcons.knowledge;
  document.getElementById("iconDownload").innerHTML = AppIcons.download;
  document.getElementById("iconCopy").innerHTML = AppIcons.copy;
  document.getElementById("iconStartSparkle").innerHTML = AppIcons.sparkles;
  document.getElementById("iconPaperclip").innerHTML = AppIcons.paperclip;
  document.getElementById("iconSend").innerHTML = AppIcons.send;

  // Drag & drop on start view
  const sv = document.getElementById("startView");
  sv.addEventListener("dragover", (e) => { e.preventDefault(); sv.style.borderColor = "var(--blue)"; });
  sv.addEventListener("dragleave", () => { sv.style.borderColor = ""; });
  sv.addEventListener("drop", (e) => {
    e.preventDefault();
    sv.style.borderColor = "";
    handleFiles(e.dataTransfer.files);
  });
});

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
    if (file.name.match(/\.(docx|doc|pdf)$/i)) {
      resolve(`[File: ${file.name}]`);
      return;
    }
    const r = new FileReader();
    r.onload = () => resolve(r.result || "");
    r.onerror = () => resolve("");
    r.readAsText(file);
  });
}

function addFileChip(name) {
  const chip = document.createElement("span");
  chip.className = "tag tag-blue";
  chip.textContent = name;
  chip.style.cursor = "pointer";
  chip.onclick = () => chip.remove();
  document.getElementById("fileChips").appendChild(chip);
}

// ── Progress ──
function setStep(step) {
  const steps = ["intent", "extract", "knowledge", "generate", "iterate", "done"];
  const idx = steps.indexOf(step);
  steps.forEach((s, i) => {
    const el = document.getElementById("sp-" + s);
    if (!el) return;
    el.className = "step";
    if (i < idx) el.classList.add("done");
    else if (i === idx) el.classList.add("active");
  });
  document.getElementById("draftProgress").style.display = "flex";
}

// ── Chat ──
function addMsg(role, content) {
  const thread = document.getElementById("chatThread");
  const div = document.createElement("div");
  div.className = `chat-msg ${role}`;
  div.innerHTML = `
    <div class="avatar">${role === "assistant" ? "AI" : "U"}</div>
    <div class="bubble">${escapeHtml(content)}</div>
  `;
  thread.appendChild(div);
  thread.scrollTop = thread.scrollHeight;
}

function renderMessages(messages) {
  const thread = document.getElementById("chatThread");
  thread.innerHTML = "";
  (messages || []).forEach((m) => {
    const div = document.createElement("div");
    div.className = `chat-msg ${m.role}`;
    div.innerHTML = `
      <div class="avatar">${m.role === "assistant" ? "AI" : "U"}</div>
      <div class="bubble">${escapeHtml(m.content)}</div>
    `;
    thread.appendChild(div);
  });
  thread.scrollTop = thread.scrollHeight;
}

function escapeHtml(str) {
  return (str || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br>");
}

// ── Start Draft ──
async function startDraft() {
  const desc = document.getElementById("descriptionInput").value.trim();
  if (!desc && uploadedFiles.length === 0) return alert("请输入需求或上传文件");
  if (isProcessing) return;
  isProcessing = true;

  // Switch views
  document.getElementById("startView").style.display = "none";
  document.getElementById("editorView").style.display = "block";
  document.getElementById("editorActions").style.display = "flex";
  document.getElementById("sessionTag").style.display = "inline-flex";

  if (!editor) editor = new OfficeEditor(document.getElementById("editorContainer"));

  setStep("intent");
  addMsg("assistant", "正在分析合同意图...");

  try {
    const res = await fetch(`${API_BASE}/api/smart-draft/init`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        description: desc,
        knowledgeMode,
        fileTexts: uploadedFiles.map((f) => ({ name: f.name, text: f.text })),
      }),
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    sessionId = data.sessionId;
    allAnswers = data.answers || {};
    setStep(data.dialogue?.missingCount > 0 ? "iterate" : "generate");

    if (data.messages) renderMessages(data.messages);

    if (data.extractedData?.items?.length > 0) {
      addMsg("assistant", `已提取 ${data.extractedData.items.length} 条数据`);
    }

    if (data.dialogue?.missingCount > 0) {
      currentQuestion = data.dialogue.currentQuestion;
      showGuideButtons();
    } else {
      showGenerateBtn();
    }
  } catch (e) {
    addMsg("assistant", `错误: ${e.message}`);
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
      <div class="bubble">${escapeHtml(currentQuestion?.question || "请补充信息")}</div>
      <div class="guide-options">
        <button class="guide-btn" onclick="pickGuide('ask_user')">
          <span class="guide-icon">${AppIcons.send}</span>我来回答
        </button>
        <button class="guide-btn" onclick="pickGuide('kb_search')">
          <span class="guide-icon">${AppIcons.search}</span>知识库检索
        </button>
        <button class="guide-btn" onclick="pickGuide('llm_infer')">
          <span class="guide-icon">${AppIcons.sparkles}</span>AI 推断
        </button>
      </div>
    </div>
  `;
  thread.appendChild(div);
  thread.scrollTop = thread.scrollHeight;
}

function showGenerateBtn() {
  const thread = document.getElementById("chatThread");
  const div = document.createElement("div");
  div.className = "chat-msg assistant";
  div.innerHTML = `
    <div class="avatar">AI</div>
    <div>
      <div class="bubble">信息已齐全，可以生成草稿。</div>
      <div style="padding-top:8px">
        <button class="btn btn-primary btn-sm" onclick="generateDraft()">
          <span class="app-icon" style="width:14px;height:14px">${AppIcons.sparkles}</span> 生成合同草稿
        </button>
      </div>
    </div>
  `;
  thread.appendChild(div);
  thread.scrollTop = thread.scrollHeight;
}

async function pickGuide(mode) {
  // Remove guide buttons
  document.querySelectorAll(".guide-options").forEach((el) => el.parentElement.removeChild(el));

  if (mode === "ask_user") {
    document.getElementById("chatInput").focus();
    return;
  }

  addMsg("user", mode === "kb_search" ? "从知识库查找" : "AI 自动推断");
  if (mode === "llm_infer") addMsg("assistant", "正在推断...");

  try {
    const res = await fetch(`${API_BASE}/api/smart-draft/guide`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, mode }),
    });
    const data = await res.json();
    if (data.messages) renderMessages(data.messages);
    if (data.answers) allAnswers = data.answers;

    if (data.dialogue?.missingCount > 0) {
      currentQuestion = data.dialogue.currentQuestion;
      showGuideButtons();
    } else {
      showGenerateBtn();
    }
  } catch (e) {
    addMsg("assistant", `错误: ${e.message}`);
  }
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
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, field, value }),
    });
    const data = await res.json();
    if (data.messages) renderMessages(data.messages);
    if (data.answers) allAnswers = data.answers;

    if (data.dialogue?.missingCount > 0) {
      currentQuestion = data.dialogue.currentQuestion;
      showGuideButtons();
    } else {
      showGenerateBtn();
    }
  } catch (e) {
    addMsg("assistant", `错误: ${e.message}`);
  }
}

// ── Generate ──
async function generateDraft() {
  if (!sessionId || isProcessing) return;
  isProcessing = true;
  setStep("generate");
  addMsg("assistant", "正在生成合同草稿...");

  try {
    const res = await fetch(`${API_BASE}/api/smart-draft/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, guidanceMode: "llm_infer" }),
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    currentDraft = typeof data.draft === "string" ? data.draft : data.draft?.content || JSON.stringify(data.draft);
    currentDraft = currentDraft.replace(/^```[\s\S]*?\n/gm, "").replace(/\n```$/gm, "");

    if (editor) editor.setContent(currentDraft);
    setStep("done");
    if (data.messages) renderMessages(data.messages);
    addMsg("assistant", `草稿已生成${data.appliedRules ? `，应用了 ${data.appliedRules.length} 条规则` : ""}`);
    document.getElementById("sessionTag").textContent = "已完成";
  } catch (e) {
    addMsg("assistant", `生成失败: ${e.message}`);
  }
  isProcessing = false;
}

// ── Download / Copy ──
function downloadDraft() {
  if (!currentDraft) return;
  const blob = new Blob([currentDraft], { type: "text/plain;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `合同草稿_${new Date().toISOString().slice(0, 10)}.txt`;
  a.click();
}

function copyDraft() {
  if (!currentDraft) return;
  navigator.clipboard.writeText(currentDraft).then(() => alert("已复制"));
}
