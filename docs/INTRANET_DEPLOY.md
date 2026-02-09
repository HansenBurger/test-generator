# 内网离线部署指南（Docker Compose v1/v2）

本文适用于“外网构建镜像 → 内网导入部署”的场景。

## 一、外网侧：构建并导出镜像

在外网机器（可访问依赖与镜像源）执行：

```bash
./deploy.sh export-images
```

产物：
- `test-generator-images.tar`（镜像包）

将 `test-generator-images.tar` 以及项目根目录下的以下文件拷贝到内网服务器：
- `docker-compose.yml`
- `.env`（或基于 `.env.example` 生成）

## 二、内网侧：导入镜像并启动

### 1. 准备 .env

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
# 注意：如果配置为空字符串，会自动回退到默认目录
# 默认目录：/home/burger/project/test-generator/backend/data
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

### 2. 导入镜像

```bash
docker load -i test-generator-images.tar
```

### 3. 启动服务

#### Docker Compose v2

```bash
docker compose up -d --no-build
```

#### Docker Compose v1

```bash
docker-compose up -d --no-build
```

> `--no-build` 用于内网禁止拉取依赖的场景。

## 三、验证与访问

查看状态：

```bash
docker compose ps
```

或：

```bash
docker-compose ps
```

默认访问地址：
- 前端：`http://<内网IP>:3000`
- 后端：`http://<内网IP>:8001`

如果你在 `.env` 中设置了端口，请替换上述端口。

## 四、日志查看

#### Docker Compose v2

```bash
docker compose logs -f
```

#### Docker Compose v1

```bash
docker-compose logs -f
```

## 五、常见问题

1) **数据库无法连接**  
请确认 `DATABASE_URL` 或 `DB_HOST/DB_*` 正确，并确保内网数据库允许该服务器访问。

2) **XMind 文件存储路径想改**  
配置 `DATA_DIR` 或 `PARSED_DIR/GENERATION_DIR/XMIND_DIR` 即可。

3) **端口冲突**  
修改 `.env` 中 `BACKEND_PORT` / `FRONTEND_PORT`。
