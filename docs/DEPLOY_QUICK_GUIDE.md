# 快速部署指南

## 核心原理

后端镜像采用分层架构：
- **基础镜像** (~1.8GB): LibreOffice + Python + 系统依赖
- **应用镜像** (基于基础镜像): 仅包含应用代码

⚠️ **重要**: `docker save` 会导出所有依赖层，所以应用镜像导出后仍然是 ~1.8GB。

## 推荐的增量更新方式

### 前端更新 (~436KB)

```bash
# 外网
./deploy.sh export-dist          # 产出: frontend-dist.tar.gz

# 传输到内网
scp frontend-dist.tar.gz user@内网:/path/to/project/

# 内网
./deploy.sh update-dist          # 热更新，无需重启容器
```

### 后端更新 (~72KB)

```bash
# 外网
./deploy.sh export-backend-code  # 产出: backend-code.tar.gz

# 传输到内网
scp backend-code.tar.gz user@内网:/path/to/project/

# 内网
./deploy.sh update-backend-code  # 热更新（依赖基础镜像已存在）
```

## 首次部署

### 外网构建并导出

```bash
# 1. 导出基础镜像（仅首次需要）
./deploy.sh export-backend-base  # 产出: test-generator-backend-base.tar (~1.8GB)

# 2. 导出前端 dist
./deploy.sh export-dist          # 产出: frontend-dist.tar.gz (~436KB)

# 3. 传输到内网
scp *.tar *.tar.gz user@内网:/path/to/project/
```

### 内网导入并启动

```bash
# 1. 导入基础镜像
./deploy.sh import-backend-base  # 导入 ~1.8GB 基础镜像

# 2. 构建应用镜像（基于已导入的基础镜像）
./deploy.sh build-backend-app

# 3. 更新前端
./deploy.sh update-dist

# 4. 启动服务
./deploy.sh start
```

## 日常更新

### 仅前端代码变更
```bash
# 外网
./deploy.sh export-dist

# 内网
./deploy.sh update-dist
```

### 仅后端代码变更
```bash
# 外网
./deploy.sh export-backend-code

# 内网
./deploy.sh update-backend-code
```

### Python 依赖变更 (requirements.txt)
```bash
# 外网
./deploy.sh export-backend-base  # 重新构建基础镜像

# 内网
./deploy.sh import-backend-base
./deploy.sh build-backend-app
./deploy.sh restart
```

## 镜像大小对比

| 组件 | 大小 | 说明 |
|------|------|------|
| 基础镜像 (test-generator-backend-base) | ~1.8GB | 包含 LibreOffice、Python、系统依赖 |
| 应用镜像 (test-generator-backend) | ~1.8GB | 基于基础镜像 + 代码（导出时包含所有层） |
| 后端代码包 (backend-code.tar.gz) | ~72KB | 仅代码，用于热更新 |
| 前端 dist 包 (frontend-dist.tar.gz) | ~436KB | 构建后的静态文件 |
| 前端镜像 (test-generator-frontend) | ~30MB | nginx + 静态文件 |

## 常见问题

**Q: 为什么 export-backend 导出的文件还是 1.8GB？**  
A: Docker 镜像是分层存储的，`docker save` 会导出所有层（包括依赖的基础镜像层）。这是 Docker 的设计，无法避免。请使用 `export-backend-code` 进行增量更新。

**Q: 内网没有基础镜像怎么办？**  
A: 首次部署需要先执行 `export-backend-base` 导出基础镜像（~1.8GB），内网执行 `import-backend-base` 导入。之后只需传代码包即可。

**Q: 如何验证内网是否已有基础镜像？**  
A: 在内网执行 `docker images | grep test-generator-backend-base`，如果有输出说明已存在。

**Q: 热更新后代码没生效？**  
A: 确保 `docker-compose.yml` 中后端配置了 `volumes: - ./backend:/app`，代码会通过挂载自动同步。

**Q: 前端热更新后页面没变化？**  
A: 浏览器缓存，按 `Ctrl+Shift+R` 强制刷新。
