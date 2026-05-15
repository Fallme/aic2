// model-test.js — AI Model connectivity test
function initModelTest() {
  const modelInfo = document.getElementById("aiModel");
  if (!modelInfo) return;

  // Create modal
  const modal = document.createElement("div");
  modal.className = "test-modal";
  modal.id = "testModal";
  modal.innerHTML = `
    <button class="close-btn" onclick="document.getElementById('testModal').style.display='none'">✕</button>
    <h4>AI 模型连接测试</h4>
    <div class="result" id="testResult"><span class="spinner"></span>正在测试连接...</div>
  `;
  document.body.appendChild(modal);

  // Click handler
  modelInfo.addEventListener("click", (e) => {
    e.stopPropagation();
    runModelTest();
  });

  // Close on outside click
  document.addEventListener("click", (e) => {
    if (!modal.contains(e.target) && e.target !== modelInfo) {
      modal.style.display = "none";
    }
  });
}

async function runModelTest() {
  const modal = document.getElementById("testModal");
  const result = document.getElementById("testResult");
  modal.style.display = "block";
  result.className = "result";
  result.innerHTML = '<span class="spinner"></span>正在测试连接...';

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 5000);
  const start = Date.now();

  try {
    const res = await fetch("/api/health", { signal: controller.signal });
    clearTimeout(timeout);
    const latency = Date.now() - start;
    const data = await res.json();

    result.className = "result ok";
    result.innerHTML = `
      <div style="margin-bottom:4px"><strong>连接成功</strong></div>
      <div>模型: <span class="latency">${data.model || "unknown"}</span></div>
      <div>提供商: ${data.provider || data.modelProvider || "unknown"}</div>
      <div>延迟: <span class="latency">${latency}ms</span></div>
      <div>API Key: ${data.apiKeyConfigured ? "已配置" : "未配置"}</div>
    `;
  } catch (e) {
    clearTimeout(timeout);
    result.className = "result fail";
    if (e.name === "AbortError") {
      result.innerHTML = `<div><strong>连接超时</strong>（超过5000ms）</div><div>请检查网络或API配置</div>`;
    } else {
      result.innerHTML = `<div><strong>连接失败</strong></div><div>${e.message}</div>`;
    }
  }
}
