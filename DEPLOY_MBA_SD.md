# MBA & Steam Deck 模型部署指南

## MacBook Air M1（8GB）

### 安装 Ollama
```bash
# 去 https://ollama.com/download 下载 macOS 版，拖进 Applications 即可
# 终端验证：
ollama --version
```

### 拉取模型（约 5.8 GB）
```bash
ollama pull qwen3:4b           # 2.5 GB，主力聊天
ollama pull nomic-embed-text   # 274 MB，知识库嵌入
ollama pull openbmb/minicpm-v4:q4_0  # 3.0 GB，图片识别
```

### 安装 Open WebUI（Docker）
```bash
# 先装 Docker Desktop: https://www.docker.com/products/docker-desktop/
# 然后：
docker run -d -p 3000:8080 \
  --add-host=host.docker.internal:host-gateway \
  -v open-webui:/app/backend/data \
  --name open-webui \
  --restart always \
  ghcr.io/open-webui/open-webui:main
```

访问 `http://localhost:3000`，创建管理员账号，在设置里连接 `http://host.docker.internal:11434`。

---

## Steam Deck（16GB / Zen2）

### 安装 Ollama
```bash
# SteamOS 桌面模式 (Arch Linux)：
curl -fsSL https://ollama.com/install.sh | sh

# Windows 11：
# 下载安装包 https://ollama.com/download/windows
```

### 拉取模型（约 5.8 GB）
```bash
ollama pull qwen3:4b
ollama pull nomic-embed-text
ollama pull openbmb/minicpm-v4:q4_0
```

> Steam Deck CPU 纯跑，4B 模型约 5-8 token/s，聊天够用。

### 安装 Open WebUI
```bash
# SteamOS（Docker）：
docker run -d -p 3000:8080 \
  --add-host=host.docker.internal:host-gateway \
  -v open-webui:/app/backend/data \
  --name open-webui \
  --restart always \
  ghcr.io/open-webui/open-webui:main

# Windows 11（Docker Desktop 同理）
```

---

## 卸载模型
```bash
ollama rm qwen3:4b          # 删单个
ollama list                  # 看还剩哪些
```
