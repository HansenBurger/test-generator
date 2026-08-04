# 内网离线部署指南（Docker Compose v1/v2）

本文适用于"外网构建镜像 → 内网导入部署"的场景。

## 部署方式对比

| 方式 | 适用场景 | 传输大小 | 内网操作 |
|------|---------|---------|---------|
| **全量部署** | 首次部署 / 前后端同时更新 | ~数GB | `import-images` |
| **前端 dist 热更新** | 仅前端代码变更（最快） | ~436KB | `update-dist` |
| **前端镜像更新** | 仅前端代码变更 | ~30MB | `import-frontend` |
| **后端应用镜像更新** | 仅后端代码变更 | ~几MB | `import-backend` |
| **后端代码热更新** | 仅后端代码变更（volume 模式） | ~几MB | `update-backend-code` |
| **后端基础镜像更新** | Python 依赖变更（极少） | ~1-2GB | `import-backend-base` |

### 镜像架构说明

```
┌─────────────────────────────────────────────────────────────┐
│  前端镜像 (test-generator-frontend:latest) ~30MB            │
│  └── nginx + 静态文件                                        │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│  后端应用镜像 (test-generator-backend:latest) ~几MB          │
│  └── 应用代码 (FROM test-generator-backend-base)             │
└─────────────────────────────────────────────────────────────┘
         │ 基于
         ▼
┌─────────────────────────────────────────────────────────────┐
│  后端基础镜像 (test-generator-backend-base:latest) ~1-2GB    │
│  └── python:3.10-slim + LibreOffice + pip packages           │
│  └── 很少变化，仅依赖更新时需要重建                            │
└─────────────────────────────────────────────────────────────┘
```

**优化要点：**
- LibreOffice 体积大（~1-2GB），但很少变化
- 后端代码变更时，只需重建/导出应用镜像（~几MB），无需重新打包 LibreOffice
- 前端更新时，可用 `export-dist` 直接打包静态文件（~436KB），跳过镜像构建

---

## 方式一：全量部署（首次部署推荐）

### 1. 外网侧：构建并导出镜像

```bash
./deploy.sh export-images
```

产物：`test-generator-images.tar`（包含前后端镜像 + 基础镜像）

将以下文件拷贝到内网服务器：
- `test-generator-images.tar`
- `docker-compose.yml`
- `.env`（或基于 `.env.example` 生成）
- `deploy.sh`

### 2. 内网侧：导入镜像并启动

```bash
./deploy.sh import-images
```

---

## 方式二：前端 dist 热更新（最快，推荐日常前端优化）

无需构建 Docker 镜像，仅更新前端静态文件。

### 1. 外网侧：构建 dist 并打包

```bash
./deploy.sh export-dist
```

产物：`frontend-dist.tar.gz`（约 436KB）

### 2. 内网侧：解压并热更新

```bash
# 拷贝 frontend-dist.tar.gz 到项目根目录后执行
./deploy.sh update-dist
```

> **原理：** 通过 `docker cp` 直接将新的 dist 文件注入运行中的 nginx 容器，然后 reload nginx。
> 无需停止服务、无需重建镜像，秒级完成。

---

## 方式三：前端镜像增量更新（仅改前端代码）

### 1. 外网侧：构建并导出前端镜像

```bash
./deploy.sh export-frontend
```

产物：`test-generator-frontend.tar`（约 30MB）

### 2. 内网侧：导入并更新

```bash
./deploy.sh import-frontend
```

---

## 方式四：后端应用镜像增量更新（仅改后端代码）

**前提：** 内网已有 `test-generator-backend-base` 基础镜像。

### 1. 外网侧：构建并导出后端应用镜像

```bash
./deploy.sh export-backend
```

产物：`test-generator-backend.tar`（约几MB，不含 LibreOffice）

### 2. 内网侧：导入并更新

```bash
./deploy.sh import-backend
```

> 仅重建应用层镜像，LibreOffice 保持不变。

---

## 方式五：后端代码热更新（volume 模式，零镜像操作）

适用于 `docker-compose.yml` 中已配置 `volumes: - ./backend:/app` 的场景。

### 1. 外网侧：打包后端代码

```bash
./deploy.sh export-backend-code
```

产物：`backend-code.tar.gz`（约几MB，仅代码）

### 2. 内网侧：解压并重启

```bash
./deploy.sh update-backend-code
```

> **原理：** docker-compose 通过 volume mount 将 `./backend` 挂载到容器内 `/app`。
> 更新代码后只需 `docker compose restart backend`，无需重建镜像。

> **数据沿用：** 更新时旧 `backend/` 会先备份到 `backend.backup.<时间戳>/`，
> 代码包不含 `backend/data`（数据库与解析缓存不随包传输），脚本会自动把旧目录的
> `data/` 沿用回新 `backend/`，无需手动搬运。

---

## 方式六：后端基础镜像更新（仅依赖变更时）

**极少使用**，仅在 `requirements.txt` 或系统依赖变更时需要。

### 1. 外网侧：构建并导出基础镜像

```bash
./deploy.sh export-backend-base
```

产物：`test-generator-backend-base.tar`（约 1-2GB）

### 2. 内网侧：导入基础镜像

```bash
./deploy.sh import-backend-base
```

> 导入基础镜像后，还需要重新构建/导入应用镜像。

---

## 组合：代码级增量更新（前后端都只改代码、无新增依赖时推荐）

不需要更新全量镜像/基础镜像，组合 **方式五 + 方式二** 即可（两个包合计约几MB）：

### 1. 外网侧：打包

```bash
./deploy.sh export-backend-code   # 产物 backend-code.tar.gz
./deploy.sh export-dist           # 产物 frontend-dist.tar.gz
```

### 2. 内网侧：更新（把两个包拷到项目根目录后）

```bash
# 0) 仅从旧版本升级且 MySQL 账号无 ALTER 权限时需要：
#    先请 DBA 执行补列 SQL（见"数据库权限与表结构升级"），有 ALTER 权限则跳过
# 1) 后端：自动备份旧代码、自动沿用 data 目录、自动重启容器
./deploy.sh update-backend-code
# 2) 前端：热更新 dist
./deploy.sh update-dist
# 3) 解析逻辑有变化时：软失效解析缓存，让重新上传走新解析（可选）
./clear_cache.sh
```

> `clear_cache.sh` 在无宿主机 Python 依赖时会自动改用 backend 容器执行；
> 只打无效标志，不删数据。

### 3. 验证

```bash
./deploy.sh status
```

---

## 环境配置 (.env)

将 `.env.example` 复制为 `.env` 并按需修改：

```bash
cp .env.example .env
```

**最少需要设置：**

```
DASHSCOPE_API_KEY=你的密钥
```

**可选配置：**

```
# 大模型 API 访问地址（可选）
DASHSCOPE_API_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1

# MySQL 连接（二选一）
DATABASE_URL=mysql+pymysql://user:password@host:3306/test_generator?charset=utf8mb4
# 或
DB_HOST=127.0.0.1
DB_PORT=3306
DB_USER=root
DB_PASSWORD=your_password
DB_NAME=test_generator

# 存储路径（可选）
DATA_DIR=/app/data
PARSED_DIR=/app/data/parsed
GENERATION_DIR=/app/data/generation
XMIND_DIR=/app/data/xmind

# 端口（可选）
BACKEND_PORT=8001
FRONTEND_PORT=3000

# 数据卷（可选）
DATA_VOLUME=backend_data
```

### 数据库权限与表结构升级（MySQL）

应用**不需要 DROP（删表）权限**。首次启动时 SQLAlchemy 会自动创建缺失的表
（需要 CREATE 权限）；版本升级新增的列会自动 `ALTER TABLE ... ADD COLUMN`
（需要 ALTER 权限）。

按你的数据库账号权限，分三种情况：

| 情况 | 需要做什么 |
| --- | --- |
| 全新库（还没建过表），账号有 CREATE 权限 | 不用管，启动时自动建表（已含全部新列） |
| 旧版本已建过表，账号有 ALTER 权限 | 不用管，升级后首次启动自动补列 |
| 旧版本已建过表，账号只有增删改查权限 | 需要请 DBA 手动执行一次下方 SQL |

若账号无 ALTER 权限，启动日志会打印 WARNING 及需要手动执行的 SQL。
**注意：补列完成前不要升级到新版本**，否则 XMind 解析接口会因缺少
`is_invalid` 列而报错。手动补列 SQL：

```sql
ALTER TABLE parse_records ADD COLUMN is_invalid BOOLEAN;
ALTER TABLE parse_records ADD COLUMN invalidated_at DATETIME;
```

（历史上还可能需要的模型配置补列，若库是从更早版本升级且未执行过：
`ALTER TABLE model_config ADD COLUMN temperature FLOAT;`
`ALTER TABLE model_config ADD COLUMN model_mode VARCHAR(16);`）

---

## 验证与访问

查看状态：

```bash
./deploy.sh status
```

默认访问地址：
- 前端：`http://<内网IP>:3000`
- 后端：`http://<内网IP>:8001`

---

## 常见问题

1) **数据库无法连接**
   请确认 `DATABASE_URL` 或 `DB_HOST/DB_*` 正确，并确保内网数据库允许该服务器访问。

2) **XMind 文件存储路径想改**
   配置 `DATA_DIR` 或 `PARSED_DIR/GENERATION_DIR/XMIND_DIR` 即可。

3) **端口冲突**
   修改 `.env` 中 `BACKEND_PORT` / `FRONTEND_PORT`。

4) **前端热更新后页面没有变化**
   浏览器可能缓存了旧文件，尝试 `Ctrl+Shift+R`（强制刷新）或清除浏览器缓存。

5) **后端热更新后代码没有生效**
   确保 `docker-compose.yml` 中后端服务配置了 `volumes: - ./backend:/app`，否则需要走镜像更新流程。

6) **deploy.sh 交互式菜单**
   直接运行 `./deploy.sh`（不带参数）可进入交互式菜单，查看所有可用操作。
   输入 `99` 可查看镜像架构说明。
