# 合同知识库与规则库模块

## 启动

推荐用服务端环境变量保存 DashScope Key，不要写进前端代码。

```powershell
cd E:\CODE\copenc\saascontract
.\start-knowledge.ps1 -ApiKey "你的DashScope API Key"
```

也可以使用通用 Node 启动方式：

```bash
npm start
```

启动后访问：

```text
http://localhost:5173/knowledge.html
```

如果不传 `-ApiKey`，模块仍可上传归档文档，但 AI 提取会走本地启发式兜底，效果弱于大模型。

## 功能

- 批量上传：支持 Word `.docx`、PDF、TXT、JSON、CSV、HTML 等文件。
- 分类归档：按资料类型、业务领域、合同类型、保密级别保存。
- 知识库：保存文档摘要、解析状态、切片数量和来源信息。
- 规则提取：调用 DashScope OpenAI 兼容接口，从选中文档中抽取候选规则。
- 规则分类：按“四大维度”归类，包括通用必备条款、软件外包行业条款、软硬件集成总包行业条款、制造业工业条款。
- 规则审核：候选规则默认待审核，可在页面上通过启用或驳回。

## 后端接口

- `GET /api/health`：服务状态与模型配置。
- `GET /api/dimensions`：四大维度参考库。
- `GET /api/documents`：文档归档列表。
- `POST /api/documents`：批量上传文档。
- `POST /api/extract-rules`：从选中文档提取规则。
- `GET /api/rules`：规则列表。
- `POST /api/rules/review`：审核规则。

## 文件说明

- `server.js`：本地后端服务，负责上传、解析、存储和调用 DashScope。
- `knowledge.html`：知识库与规则库页面。
- `knowledge.js`：前端交互逻辑。
- `knowledge.css`：模块样式。
- `contract-knowledge-dimensions.json`：四大维度参考库。
- `data/kb-store.json`：本地文档与规则元数据。
- `data/uploads/`：上传文件目录，已加入 `.gitignore`。

公网部署见 `DEPLOYMENT.md`。
