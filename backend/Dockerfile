# 后端应用镜像 Dockerfile
# 基于 test-generator-backend-base 镜像
# 仅包含应用代码，体积小巧（几MB）
# 代码变更时只需重建此镜像

# 使用基础镜像（包含 LibreOffice 和 Python 依赖）
FROM test-generator-backend-base:latest

# 设置工作目录
WORKDIR /app

# 复制应用代码
# 注意：docker-compose 会通过 volumes 挂载覆盖此目录
# 这里保留 COPY 是为了支持独立构建（不使用 compose 的场景）
COPY . .

# 暴露端口
EXPOSE 8001

# 启动命令
CMD ["python", "main.py"]
