// smart-draft.js — Smart Draft Frontend Orchestration
// Multi-Agent contract generation with iterative dialogue

const API_BASE = "";
let editor = null;
let sessionId = "";
let knowledgeMode = "industry";
let guidanceMode = "ask_user";
let uploadedFiles = []; // [{name, text}]
let currentQuestion = null;
let allAnswers = {};
let currentDraft = "";
let isProcessing = false;

// ── Initialization ──
window.addEventListener("DOMContentLoaded", () => {
  // Init drag & drop
  const zone = document.getElementById("uploadZone");
  zone.addEventListener("dragover", (e) => { e.preventDefault(); zone.classList.add("dragover"); });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("dragover");
    handleFiles(e.dataTransfer.files);
  });
});

// ── File Handling ──
async function handleFiles(fileList) {
  for (const file of fileList) {
    const text = await readFileAsText(file);
    uploadedFiles.push({ name: file.name, text: text.slice(0, 10000), size: file.size });
    renderFileList();
  }
}

function readFileAsText(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result || "");
    reader.onerror = () => resolve("");
    // For .docx/.doc/.pdf, read as text won't work well, but we still capture filename
    if (file.name.match(/\.(docx|doc|pdf)$/i)) {
      resolve(`[文件: ${file.name}, 大小: ${(file.size / 1024).toFixed(1)}KB — 需要服务端解析]`);
    } else {
      reader.readAsText(file);
    }
  });
}

function renderFileList() {
  const ul = document.getElementById("fileList");
  ul.innerHTML = uploadedFiles.map((f, i) =>
    `<li><span class="name">📄 ${f.name}</span><span style="color:#999;font-size:11px;margin:0 8px;">${(f.size / 1024).toFixed(0)}KB</span><span class="remove" onclick="removeFile(${i})">×</span></li>`
  ).join("");
}

function removeFile(i) {
  uploadedFiles.splice(i, 1);
  renderFileList();
}

// ── Knowledge Mode ──
function setKnowledgeMode(mode, btn) {
  knowledgeMode = mode;
  document.querySelectorAll(".knowledge-toggle button").forEach((b) => b.classList.remove("active"));
  btn.classList.add("active");
}

// ── Agent Progress ──
function updateProgress(step) {
  const steps = ["intent", "extract", "knowledge", "generate", "iterate"];
  const labels = {
    intent: "意图分析中...",
    extract: "数据提取中...",
    knowledge: "知识检索中...",
    generate: "草稿生成中...",
    iterate: "迭代完善中...",
    done: "完成",
    idle: "等待开始",
  };

  steps.forEach((s, i) => {
    const el = document.getElementById("step-" + s);
    const currentIdx = steps.indexOf(step);
    el.className = "step";
    if (i < currentIdx) el.classList.add("done");
    else if (i === currentIdx) el.classList.add("active");
  });

  document.getElementById("stepLabel").textContent = labels[step] || step;
}

// ── Chat Messages ──
function renderMessages(messages) {
  const container = document.getElementById("msgList");
  container.innerHTML = messages.map((m) =>
    `<div class="msg ${m.role}">${escapeHtml(m.content)}</div>`
  ).join("");
  container.scrollTop = container.scrollHeight;
}

function addMessage(role, content) {
  const container = document.getElementById("msgList");
  const div = document.createElement("div");
  div.className = `msg ${role}`;
  div.textContent = content;
  container.appendChild(div);
  container.scrollTop = container.scrollHeight;
}

function escapeHtml(str) {
  return (str || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ── Main Flow ──
async function startSmartDraft() {
  const desc = document.getElementById("descriptionInput").value.trim();
  if (!desc && uploadedFiles.length === 0) {
    alert("请输入合同描述或上传参考文件");
    return;
  }
  if (isProcessing) return;
  isProcessing = true;

  // Switch to editor view
  document.getElementById("startPanel").style.display = "none";
  document.getElementById("editorContainer").style.display = "block";

  // Init editor
  if (!editor) {
    editor = new OfficeEditor(document.getElementById("officeEditor"));
  }

  updateProgress("intent");
  addMessage("assistant", "🔄 正在分析合同意图...");

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

    updateProgress(data.dialogue?.missingCount > 0 ? "iterate" : "generate");

    // Render messages
    if (data.messages) renderMessages(data.messages);

    // If there's extracted data, show summary
    if (data.extractedData && data.extractedData.items && data.extractedData.items.length > 0) {
      addMessage("assistant", `📊 已提取 ${data.extractedData.items.length} 条数据项`);
    }

    // Handle dialogue
    if (data.dialogue && data.dialogue.missingCount > 0) {
      currentQuestion = data.dialogue.currentQuestion;
      showGuideOptions();
    } else {
      // All info ready, show generate button
      showGenerateButton();
    }
  } catch (e) {
    addMessage("assistant", `❌ 错误: ${e.message}`);
  }

  isProcessing = false;
}

// ── Guide Options ──
function showGuideOptions() {
  document.getElementById("guideSection").style.display = "block";
  document.getElementById("composerArea").style.display = "none";
  document.getElementById("generateSection").style.display = "none";
}

function showComposer() {
  document.getElementById("guideSection").style.display = "none";
  document.getElementById("composerArea").style.display = "flex";
  document.getElementById("generateSection").style.display = "none";
  document.getElementById("composerInput").focus();
}

function showGenerateButton() {
  document.getElementById("guideSection").style.display = "none";
  document.getElementById("composerArea").style.display = "none";
  document.getElementById("generateSection").style.display = "block";
  currentQuestion = null;
}

async function selectGuide(mode) {
  guidanceMode = mode;

  if (mode === "ask_user") {
    showComposer();
  } else if (mode === "kb_search") {
    addMessage("user", "📚 从知识库查找");
    try {
      const res = await fetch(`${API_BASE}/api/smart-draft/guide`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sessionId, mode: "kb_search" }),
      });
      const data = await res.json();
      if (data.messages) renderMessages(data.messages);

      if (data.snippets && data.snippets.length > 0) {
        addMessage("assistant", "以下是知识库中的相关内容，请选择或输入补充信息：");
      }
      showComposer();
    } catch (e) {
      addMessage("assistant", `❌ 检索失败: ${e.message}`);
    }
  } else if (mode === "llm_infer") {
    addMessage("user", "🤖 AI 自动推断");
    addMessage("assistant", "🔄 正在推断...");
    try {
      const res = await fetch(`${API_BASE}/api/smart-draft/guide`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sessionId, mode: "llm_infer" }),
      });
      const data = await res.json();
      if (data.messages) renderMessages(data.messages);
      if (data.answers) allAnswers = data.answers;

      if (data.dialogue && data.dialogue.missingCount > 0) {
        currentQuestion = data.dialogue.currentQuestion;
        showGuideOptions();
      } else {
        showGenerateButton();
      }
    } catch (e) {
      addMessage("assistant", `❌ 推断失败: ${e.message}`);
    }
  }
}

// ── Send Answer ──
async function sendAnswer() {
  const input = document.getElementById("composerInput");
  const value = input.value.trim();
  if (!value || !sessionId) return;
  input.value = "";

  const field = currentQuestion ? currentQuestion.hint : "补充信息";

  addMessage("user", `${field}: ${value}`);

  try {
    const res = await fetch(`${API_BASE}/api/smart-draft/answer`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, field, value }),
    });
    const data = await res.json();
    if (data.messages) renderMessages(data.messages);
    if (data.answers) allAnswers = data.answers;

    if (data.dialogue && data.dialogue.missingCount > 0) {
      currentQuestion = data.dialogue.currentQuestion;
      showGuideOptions();
    } else {
      showGenerateButton();
    }
  } catch (e) {
    addMessage("assistant", `❌ 错误: ${e.message}`);
  }
}

// ── Generate Draft ──
async function generateDraft() {
  if (!sessionId || isProcessing) return;
  isProcessing = true;
  updateProgress("generate");
  addMessage("assistant", "✨ 正在生成合同草稿...");

  try {
    const res = await fetch(`${API_BASE}/api/smart-draft/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, guidanceMode }),
    });
    const data = await res.json();

    if (data.error) throw new Error(data.error);

    currentDraft = data.draft || "";

    // Strip markdown code fences
    currentDraft = currentDraft.replace(/^```[\s\S]*?\n/gm, "").replace(/\n```$/gm, "");

    // Display in editor
    if (editor) {
      editor.setContent(currentDraft);
    }

    updateProgress("done");
    if (data.messages) renderMessages(data.messages);

    addMessage("assistant", `✅ 草稿已生成！${data.appliedRules ? `（应用了 ${data.appliedRules.length} 条规则）` : ""}`);

    document.getElementById("generateSection").style.display = "none";
  } catch (e) {
    addMessage("assistant", `❌ 生成失败: ${e.message}`);
    updateProgress("idle");
  }

  isProcessing = false;
}

// ── Download / Copy ──
function downloadDraft() {
  if (!currentDraft) return alert("暂无草稿内容");
  const blob = new Blob([currentDraft], { type: "text/plain;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `合同草稿_${new Date().toLocaleDateString("zh-CN").replace(/\//g, "-")}.txt`;
  a.click();
}

function copyDraft() {
  if (!currentDraft) return alert("暂无草稿内容");
  navigator.clipboard.writeText(currentDraft).then(() => {
    alert("已复制到剪贴板");
  });
}
