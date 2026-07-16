# 内网离线部署指南（Docker Compose v1/v2）

本文适用于"外网构建镜像 → 内网导入部署"的场景。

## 部署方式对比

| 方式 | 适用场景 | 传输大小 | 内网操作 |
|------|---------|---------|---------|
| **全量部署** | 首次部署 / 前后端同时更新 | ~数GB | `import-images` |
| **前端镜像更新** | 仅前端代码变更 | ~30MB | `import-frontend` |
| **后端镜像更新** | 仅后端代码变更 | ~数GB | `import-backend` |
| **前端 dist 热更新** | 仅前端代码变更（最快） | ~1-5MB | `update-dist` |

---

## 方式一：全量部署（首次部署推荐）

### 1. 外网侧：构建并导出镜像

```bash
./deploy.sh export-images
```

产物：`test-generator-images.tar`（包含前后端镜像）

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

## 方式二：前端镜像增量更新（仅改前端代码）

### 1. 外网侧：构建并导出前端镜像

```bash
./deploy.sh export-frontend
```

产物：`test-generator-frontend.tar`（约 30MB）

### 2. 内网侧：导入并更新

```bash
# 拷贝 tar 文件后执行
./deploy.sh import-frontend
```

> 仅重建前端容器，后端服务不受影响。

---

## 方式三：后端镜像增量更新（仅改后端代码）

### 1. 外网侧：构建并导出后端镜像

```bash
./deploy.sh export-backend
```

产物：`test-generator-backend.tar`

### 2. 内网侧：导入并更新

```bash
# 拷贝 tar 文件后执行
./deploy.sh import-backend
```

---

## 方式四：前端 dist 热更新（最快，推荐日常前端优化）

无需构建 Docker 镜像，仅更新前端静态文件。

### 1. 外网侧：构建 dist 并打包

```bash
./deploy.sh export-dist
```

产物：`frontend-dist.tar.gz`（约 1-5MB）

### 2. 内网侧：解压并热更新

```bash
# 拷贝 frontend-dist.tar.gz 到项目 frontend/ 目录后执行
./deploy.sh update-dist
```

> **原理：** 通过 `docker cp` 直接将新的 dist 文件注入运行中的 nginx 容器，然后 reload nginx。
> 无需停止服务、无需重建镜像，秒级完成。

### 手动方式（如果不用 deploy.sh）

```bash
# 1. 解压
tar -xzf frontend-dist.tar.gz -C frontend/

# 2. 注入到运行中的容器
docker exec test-generator-frontend rm -rf /usr/share/nginx/html/assets
docker cp frontend/dist/. test-generator-frontend:/usr/share/nginx/html/
docker exec test-generator-frontend nginx -s reload
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

5) **deploy.sh 交互式菜单**
   直接运行 `./deploy.sh`（不带参数）可进入交互式菜单，查看所有可用操作。
