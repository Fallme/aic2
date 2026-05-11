// template-fill.js — Template Fill Feature Logic

let editor = null;
let templateData = null; // parsed template from server
let filledValues = {};
let docxId = "";
let attachFiles = [];

// ── Init icons ──
window.addEventListener("DOMContentLoaded", () => {
  // Set nav icons
  document.getElementById("navTemplateFill").innerHTML = AppIcons.templateFill;
  document.getElementById("navSmartDraft").innerHTML = AppIcons.smartDraft;
  document.getElementById("navReview").innerHTML = AppIcons.contractReview;
  document.getElementById("navKnowledge").innerHTML = AppIcons.knowledge;
  document.getElementById("iconDownload").innerHTML = AppIcons.download;
  document.getElementById("iconUploadBig").innerHTML = AppIcons.upload;
  document.getElementById("iconUploadBtn").innerHTML = AppIcons.upload;
  document.getElementById("iconSparkle").innerHTML = AppIcons.sparkles;
  document.getElementById("iconPaperclip").innerHTML = AppIcons.paperclip;
  document.getElementById("iconSend").innerHTML = AppIcons.send;

  // Init editor
  editor = new OfficeEditor(document.getElementById("editorContainer"));
});

// ── Upload Template ──
async function handleUpload(file) {
  if (!file) return;
  addChat("assistant", `正在上传并解析「${file.name}」...`);

  const formData = new FormData();
  formData.append("file", file);

  try {
    const res = await fetch("/api/templates/upload", { method: "POST", body: formData });
    const data = await res.json();

    if (data.error) throw new Error(data.error);

    templateData = data;
    docxId = data.docxId;
    filledValues = {};

    addChat("assistant", `已解析模板，发现 ${data.fields ? data.fields.length : 0} 个可填写字段。`);

    // Switch to fill state
    document.getElementById("uploadState").style.display = "none";
    document.getElementById("fillState").style.display = "flex";
    document.getElementById("pageTitle").textContent = file.name;

    renderFields(data.fields || []);

    // Load preview in editor
    if (docxId) {
      try {
        const pres = await fetch(`/api/templates/parse`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ docxId }),
        });
        const presData = await pres.json();
        if (presData.html && editor) {
          editor.setContent(presData.html);
        }
      } catch {}
    }
  } catch (e) {
    addChat("assistant", `上传失败: ${e.message}`);
  }
}

// ── Render Fields ──
function renderFields(fields) {
  const container = document.getElementById("fieldsList");
  container.innerHTML = "";

  fields.forEach((field, i) => {
    const div = document.createElement("div");
    div.className = "field-group";
    div.innerHTML = `
      <div class="field-label">
        ${field.importance === "required" ? '<span class="required">*</span>' : ""}
        ${field.label || field.placeholder}
        ${field.fieldType ? `<span class="tag tag-blue" style="font-size:10px">${field.fieldType}</span>` : ""}
      </div>
      <input class="field-input" id="field_${i}"
        placeholder="${field.placeholder || field.label || ""}"
        value="${filledValues[field.placeholder] || ""}"
        oninput="updateField('${field.placeholder.replace(/'/g, "\\'")}', this.value, ${i})">
    `;
    container.appendChild(div);
  });

  updateProgress();
}

// ── Update Field Value ──
function updateField(placeholder, value, index) {
  filledValues[placeholder] = value;
  const input = document.getElementById(`field_${index}`);
  if (input) {
    input.classList.toggle("filled", value.length > 0);
  }
  updateProgress();
}

// ── Update Progress ──
function updateProgress() {
  const fields = templateData?.fields || [];
  const total = fields.length;
  const filled = fields.filter((f) => filledValues[f.placeholder]).length;
  document.getElementById("fillCount").textContent = `${filled}/${total}`;
  document.getElementById("fillBar").style.width = total > 0 ? (filled / total * 100) + "%" : "0%";
}

// ── AI Auto Fill ──
async function aiAutoFill() {
  addChat("assistant", "正在分析合同信息，自动填充字段...");

  const allInfo = Object.entries(filledValues).filter(([, v]) => v).map(([k, v]) => `${k}: ${v}`).join("\n");

  const prompt = `你是一个合同填写助手。请根据以下信息，为合同模板的每个空字段填写合理的值。

已有信息:
${allInfo || "（暂无）"}

需要填写的字段:
${(templateData?.fields || []).map((f) => `- ${f.label || f.placeholder} (${f.fieldType || "未知类型"})`).join("\n")}

请返回JSON格式，key为字段placeholder，value为填写内容:
{"字段名1": "填写内容1", "字段名2": "填写内容2"}`;

  try {
    const res = await fetch("/api/templates/ai-analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ docxId, prompt, fields: templateData?.fields }),
    });
    const data = await res.json();

    if (data.values) {
      Object.assign(filledValues, data.values);
      // Update inputs
      const fields = templateData?.fields || [];
      fields.forEach((f, i) => {
        const input = document.getElementById(`field_${i}`);
        if (input && filledValues[f.placeholder]) {
          input.value = filledValues[f.placeholder];
          input.classList.add("filled");
        }
      });
      updateProgress();
      addChat("assistant", `已自动填充 ${Object.keys(data.values).length} 个字段。请检查并补充剩余信息。`);
    }
  } catch (e) {
    addChat("assistant", `AI 填充失败: ${e.message}`);
  }
}

// ── Download Filled ──
async function downloadFilled() {
  if (!docxId) return alert("请先上传模板");

  try {
    const res = await fetch("/api/templates/fill", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ docxId, values: filledValues }),
    });

    if (!res.ok) throw new Error("下载失败");
    const blob = await res.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `已填写_${templateData?.fileName || "合同"}.docx`;
    a.click();
    addChat("assistant", "已下载填充后的 Word 文档。");
  } catch (e) {
    addChat("assistant", `下载失败: ${e.message}`);
  }
}

// ── Chat ──
function sendChat() {
  const input = document.getElementById("chatInput");
  const text = input.value.trim();
  if (!text) return;
  input.value = "";

  addChat("user", text);

  // Simple pattern matching for field fill
  const fields = templateData?.fields || [];
  let matched = 0;
  for (const field of fields) {
    const label = (field.label || field.placeholder || "").toLowerCase();
    for (const [key, value] of getFieldKeywords(field)) {
      if (text.includes(key) && !filledValues[field.placeholder]) {
        filledValues[field.placeholder] = value;
        matched++;
        const idx = fields.indexOf(field);
        const input = document.getElementById(`field_${idx}`);
        if (input) { input.value = value; input.classList.add("filled"); }
      }
    }
  }
  updateProgress();

  if (matched > 0) {
    addChat("assistant", `已根据你的输入自动填写了 ${matched} 个字段。还有 ${fields.length - Object.keys(filledValues).length} 个字段待填写。`);
  } else {
    addChat("assistant", "我已记录你的信息。请继续提供合同相关细节，或点击"AI 智能填充"自动生成。");
  }
}

function getFieldKeywords(field) {
  const label = (field.label || field.placeholder || "").toLowerCase();
  const pairs = [];
  if (label.includes("甲方")) pairs.push(["甲方", ""], ["买方", ""], ["委托", ""]);
  if (label.includes("乙方")) pairs.push(["乙方", ""], ["卖方", ""], ["受托", ""]);
  if (label.includes("金额") || label.includes("总价")) pairs.push(["总价", ""], ["金额", ""], ["预算", ""]);
  if (label.includes("日期") || label.includes("时间")) pairs.push(["日期", ""], ["期限", ""]);
  // Extract value from text for these pairs
  return pairs.map(([key]) => {
    const match = new RegExp(`${key}[是为：:]?\\s*([^，,。.、\\s]+)`).exec(text);
    return [key, match ? match[1] : ""];
  });
}

function handleAttach(files) {
  for (const file of files) {
    attachFiles.push({ name: file.name, size: file.size });
    addFileChip(file.name);
  }
}

function addFileChip(name) {
  const chip = document.createElement("span");
  chip.className = "tag tag-blue";
  chip.textContent = name;
  chip.style.cursor = "pointer";
  chip.onclick = () => chip.remove();
  document.getElementById("fileChipArea").appendChild(chip);
}

function addChat(role, content) {
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

function escapeHtml(str) {
  return (str || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br>");
}
