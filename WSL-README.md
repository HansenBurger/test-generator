# WSL 开发环境指南

## 快速启动

```bash
./start-wsl.sh          # 正常启动前后端
./start-wsl.sh --dev    # 调试模式（后端热重载 + 前端自动打开浏览器）
```

## 首次启动说明

**重要**：首次启动可能需要 20-30 秒，原因如下：

1. **openai 包导入慢**：WSL2 访问 Windows 文件系统（`/mnt/c/`）性能较差，`import openai` 需要约 11 秒
2. **LibreOffice 初始化**：如果未安装 LibreOffice，系统会等待 10 秒后超时

这是正常现象，请耐心等待后端健康检查通过。

## 性能优化建议

### 方案 1：将项目复制到 WSL 原生文件系统（推荐）

```bash
# 将项目复制到 WSL home 目录
cp -r /mnt/c/Users/davan/project/test-generator ~/test-generator
cd ~/test-generator
./start-wsl.sh
```

在 Linux 原生文件系统上，启动时间可减少到 5-10 秒。

### 方案 2：安装 LibreOffice（可选，用于 .doc 文件转换）

```bash
sudo apt update
sudo apt install -y libreoffice
```

安装后，`.doc` 文件转换功能将可用，且不会因等待 LibreOffice 启动而超时。

## 启动命令说明

| 命令 | 说明 |
|------|------|
| `./start-wsl.sh` | 正常启动前后端 |
| `./start-wsl.sh --dev` | 调试模式：后端热重载 + 前端自动打开浏览器 |
| `./start-wsl.sh --setup` | 仅安装依赖，不启动服务 |
| `./start-wsl.sh --backend` | 仅启动后端 |
| `./start-wsl.sh --frontend` | 仅启动前端 |

## 环境变量配置

首次运行会自动创建 `.env` 文件（从 `.env.example` 复制）。请编辑 `.env` 文件并填写：

- `DASHSCOPE_API_KEY`：必填，百炼 API 密钥
- `DATABASE_URL`：可选，MySQL 连接字符串（默认使用 SQLite）
- `BACKEND_PORT`：可选，后端端口（默认 8001）
- `FRONTEND_PORT`：可选，前端端口（默认 3000）

## 访问地址

启动成功后：

- 前端：http://localhost:3000
- 后端：http://localhost:8001
- API 健康检查：http://localhost:8001/api/health

## 常见问题

### Q: 前端出现 "ECONNREFUSED 127.0.0.1:8001" 错误

**A**: 后端还在启动中，请等待 20-30 秒后刷新页面。如果持续出现，检查后端日志。

### Q: 启动非常慢

**A**: 这是 WSL2 跨文件系统的已知性能问题。建议将项目复制到 WSL 原生文件系统（如 `~/projects/`）。

### Q: LibreOffice 相关警告

**A**: 如果不需要处理 `.doc` 文件（只需要 `.docx`），可以忽略此警告。如需支持 `.doc`，请安装 LibreOffice。

## 调试技巧

### 查看后端日志

后端日志会直接输出在终端，包括：

- FastAPI 应用启动信息
- 数据库初始化状态
- LibreOffice 守护进程状态
- API 请求日志

### 使用调试模式

```bash
./start-wsl.sh --dev
```

调试模式会：

1. 后端启用热重载（代码修改后自动重启）
2. 前端自动打开浏览器
3. 日志级别设为 debug

### 手动启动后端进行调试

```bash
cd backend
source .venv/bin/activate
export PYTHONUNBUFFERED=1
python main.py
```

## 停止服务

按 `Ctrl+C` 即可优雅停止所有服务。
