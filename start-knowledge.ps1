param(
  [string]$ApiKey = $env:DASHSCOPE_API_KEY,
  [string]$BaseUrl = "https://dashscope.aliyuncs.com/compatible-mode/v1",
  [string]$Model = "qwen-plus",
  [int]$Port = 5173
)

if (-not $ApiKey) {
  Write-Host "请通过 -ApiKey 或 DASHSCOPE_API_KEY 配置 DashScope API Key。未配置时仍可上传归档，但 AI 提取会使用本地启发式兜底。"
}

$env:DASHSCOPE_API_KEY = $ApiKey
$env:DASHSCOPE_BASE_URL = $BaseUrl
$env:DASHSCOPE_MODEL = $Model
$env:PORT = [string]$Port

node server.js
