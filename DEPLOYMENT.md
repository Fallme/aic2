# 公网部署与 Git 托管

## 当前部署（Render）

| 项目 | 值 |
|---|---|
| 平台 | Render |
| 服务 ID | `srv-d7uptlv7f7vs73csve40` |
| 公网地址 | https://aic2.onrender.com |
| GitHub 仓库 | https://github.com/Fallme/aic2 |
| 触发方式 | push 到 `main` 自动部署，或手动 POST Deploy Hook |

手动触发重新部署（Key 存放在 `.env` 的 `RENDER_DEPLOY_HOOK` 中）：

```bash
source .env && curl -X POST "$RENDER_DEPLOY_HOOK"
```

> ⚠️ Deploy Hook Key 仅限内部使用，不要提交到公开仓库。

## 推荐方案

给领导试用，建议使用以下组合：

- 代码托管：GitHub、Gitee 或 GitLab
- 公网部署：Render、Railway、Fly.io、阿里云 ECS、腾讯云 CVM
- 临时演示：优先用云服务器或 Render/Railway；正式试点再接数据库和对象存储

当前模块是 Node.js 单服务，无第三方依赖，部署入口是：

```bash
npm start
```

访问路径：

```text
https://你的域名或部署地址/knowledge.html
```

## 必须配置的环境变量

不要把 API Key 写进代码仓库。部署平台里配置环境变量：

```text
DASHSCOPE_API_KEY=你的DashScope Key
DASHSCOPE_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
DASHSCOPE_MODEL=qwen-plus
PORT=平台自动分配或5173
DATABASE_URL=PostgreSQL连接串，建议正式试用配置
DATABASE_SSL=true
```

## Git 托管步骤

```bash
git init
git add .
git commit -m "Add contract knowledge rules module"
```

然后在 GitHub/Gitee 创建一个私有仓库，再执行平台给出的命令，例如：

```bash
git remote add origin https://github.com/你的账号/你的仓库.git
git branch -M main
git push -u origin main
```

注意：`data/`、`.env`、`node_modules/` 已加入 `.gitignore`，不会提交上传文档、试用数据或密钥。

## Render / Railway 类平台部署

1. 连接 GitHub/Gitee/GitLab 仓库。
2. 选择 Web Service / Node.js。
3. Build Command 留空或使用 `npm install`。
4. Start Command 填写：

```bash
npm start
```

5. 在 Environment Variables 里配置 DashScope 相关环境变量。
6. 部署完成后打开平台分配的公网 URL。

本项目已提供 `render.yaml`，在 Render 里也可以选择 Blueprint 方式部署。部署时必须手动填写 `DASHSCOPE_API_KEY`，不要把 Key 提交到 Git。

## 云服务器部署

服务器安装 Node.js 20+ 后：

```bash
git clone 你的仓库地址
cd 仓库目录
export DASHSCOPE_API_KEY="你的DashScope Key"
export DASHSCOPE_BASE_URL="https://dashscope.aliyuncs.com/compatible-mode/v1"
export DASHSCOPE_MODEL="qwen-plus"
export PORT=5173
npm start
```

生产环境建议用 Nginx 反向代理和 HTTPS。

## 当前版本限制

- 如果没有配置 `DATABASE_URL`，上传文件和规则数据会保存在本地 `data/` 目录。
- 配置 `DATABASE_URL` 后，服务会自动创建 `app_store` 表，并把规则库、文档归档信息和合同生成记忆保存到 PostgreSQL；重启、刷新、重新部署后仍会保留。
- 首次启用数据库时，如果本地 `data/kb-store.json` 已有数据，会自动迁移到数据库。
- 上传原始文件仍建议后续接对象存储；当前数据库会保存解析后的文档文本和规则数据。
- 还没有登录、权限、审计、文件大小限制和租户隔离，不建议直接公开给不受控用户。
