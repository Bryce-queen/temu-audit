# Temu Listing AI Audit

面向 **Temu 卖家**的「商品上架 AI 审计」Web 服务：上传商品信息 / 素材后，系统结合内置合规知识库（RAG）与云端大模型，自动检查上架合规性并生成可付费下载的**诊断报告**。

> 技术栈：Python · Flask · ChromaDB（向量检索）· 智谱 / SiliconFlow（大模型）· Stripe（支付）｜ 前端：原生 `index.html` + `widget.js`

---

## ✨ 功能特性

- **AI 合规审计**：基于 `knowledge/`、`knowledge_base/` 中的 Temu 上架规范与合规规则，用 ChromaDB 做向量召回 + 大模型生成审计结论。
- **多模型支持**：通过 `ZHIPU_API_KEY` / `SILICONFLOW_API_KEY` 接入智谱、SiliconFlow 等云端模型。
- **付费报告**：集成 **Stripe** 支付，审计报告按次计费（默认 `$4.99`，由 `AUDIT_PRICE_CENTS` 控制）。
- **后台管理**：`/orders`、`/upload`、`/chat` 等私有接口由 `ADMIN_TOKEN` 保护（Token 缺失时自动放开，便于本地调试）。
- **Coze 挂件**（可选）：配置 `COZE_BOT_ID` / `COZE_TOKEN` 后注入对话挂件；留空则不注入。
- **可嵌入组件**：对外提供 `/widget.js`，可嵌入第三方页面发起审计。
- **一键部署**：内置 Docker（`docker-compose.yml` + `deploy.sh`）、PythonAnywhere（`pa_wsgi.py`）、Render（`render.yaml`）三种方案。

---

## 🛠 技术栈

| 领域 | 方案 |
|------|------|
| Web 框架 | Flask 3.x |
| 向量检索 | ChromaDB（`chroma_db/`） |
| 大模型 | 智谱 AI / SiliconFlow（HTTP 调用） |
| 支付 | Stripe |
| 文档解析 | PyPDF2 / python-docx / openpyxl |
| 部署 | Docker / PythonAnywhere / Render |
| 服务进程 | gunicorn（生产）/ Flask dev（本地） |

---

## 📁 项目结构

```
temu-audit/
├── app.py                  # Flask 主应用（路由 + 业务逻辑）
├── pa_wsgi.py              # PythonAnywhere WSGI 入口
├── requirements.txt        # Python 依赖
├── .env.example            # 环境变量模板
├── Dockerfile
├── docker-compose.yml
├── deploy.sh               # 一键 Docker 部署脚本
├── render.yaml             # Render 部署配置
├── PA_DEPLOY.md            # PythonAnywhere 部署指南
├── DEPLOY_MBA_SD.md        # 另一份部署说明
├── index.html              # 前端页面
├── static/widget.js        # 可嵌入审计组件
├── knowledge/              # 知识库（Temu 上架指南）
│   └── temu_listing_guide.txt
├── knowledge_base/         # 知识库（合规规则）
│   └── temu_compliance_rules.txt
└── chroma_db/              # ChromaDB 向量库（运行时生成）
```

---

## ⚙️ 环境变量

复制 `.env.example` 为 `.env` 并按需填写：

| 变量 | 说明 |
|------|------|
| `STRIPE_SECRET_KEY` / `STRIPE_PUBLISHABLE_KEY` / `STRIPE_WEBHOOK_SECRET` | Stripe 支付密钥（测试用 `sk_test_*`） |
| `ADMIN_TOKEN` | 后台访问令牌；设置后私有接口需带 `?admin_token=` 或 `X-Admin-Token` 头 |
| `SECRET_KEY` | Flask session 密钥 |
| `ZHIPU_API_KEY` | 智谱大模型 Key |
| `SILICONFLOW_API_KEY` | SiliconFlow Key |
| `COZE_BOT_ID` / `COZE_TOKEN` | 可选 Coze 挂件 |
| `AUDIT_PRICE_CENTS` / `AUDIT_PRICE_CURRENCY` | 报告价格（分），默认 `499` / `usd` |
| `PUBLIC_DOMAIN` | 公网域名，用于 Stripe 回调等，默认 `http://localhost:8080` |

> ⚠️ 真实密钥**不要**写入代码或提交到仓库；通过 `.env` 或部署平台的环境变量注入。

---

## 🚀 本地运行

```bash
# 1. 安装依赖（建议使用虚拟环境）
pip install -r requirements.txt

# 2. 配置环境变量
cp .env.example .env
#   编辑 .env，填入 Stripe / 模型 Key（本地调试 ADMIN_TOKEN 可留空）

# 3. 启动
python app.py
#   或生产式：gunicorn app:app --bind 0.0.0.0:8080 --workers 2 --threads 4
```

访问：

- 首页 / 审计页：http://localhost:8080/audit
- 后台订单：http://localhost:8080/orders
- 健康检查：http://localhost:8080/healthz

---

## 🐳 Docker 一键部署

```bash
bash deploy.sh          # 首次会生成默认 .env，按提示填入真实 Key 后重跑
docker compose up -d    # 启动
docker compose logs -f  # 查看日志
docker compose down     # 停止
```

---

## ☁️ 其他部署方式

- **PythonAnywhere**（免费、无需绑卡）：参见 `PA_DEPLOY.md`，WSGI 文件使用 `pa_wsgi.py`。
- **Render**：连接仓库后使用 `render.yaml`，自动 `pip install` 并以 `gunicorn` 启动。

---

## 📝 备注

- 私有接口（`/orders`、`/upload`、`/chat` 等）在未设置 `ADMIN_TOKEN` 时默认不鉴权；**生产环境务必设置强随机 `ADMIN_TOKEN` 与 `SECRET_KEY`**。
- `chroma_db/` 与 `chat.db`、`server.log` 为运行时产物，已在 `.gitignore` 中忽略，无需提交。
