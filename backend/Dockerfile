# 后端 Dockerfile
FROM python:3.10-slim

# 设置工作目录
WORKDIR /app

# 设置环境变量
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    LIBREOFFICE_HOME=/usr/lib/libreoffice \
    HOME=/tmp

# 配置apt使用阿里云镜像源
RUN if [ -f /etc/apt/sources.list ]; then \
        sed -i 's|deb.debian.org|mirrors.aliyun.com|g' /etc/apt/sources.list; \
        sed -i 's|security.debian.org|mirrors.aliyun.com|g' /etc/apt/sources.list; \
    fi && \
    if [ -d /etc/apt/sources.list.d ]; then \
        find /etc/apt/sources.list.d -name "*.sources" -exec sed -i 's|deb.debian.org|mirrors.aliyun.com|g' {} \; 2>/dev/null || true; \
        find /etc/apt/sources.list.d -name "*.list" -exec sed -i 's|deb.debian.org|mirrors.aliyun.com|g' {} \; 2>/dev/null || true; \
    fi

# 安装系统依赖（优先使用antiword + pandoc，LibreOffice作为备选）
RUN apt-get update && apt-get install -y \
    gcc \
    antiword \
    pandoc \
    libreoffice \
    libreoffice-writer \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/* \
    && mkdir -p /tmp/.config /tmp/.cache

# 配置pip使用阿里云镜像源
RUN pip config set global.index-url https://mirrors.aliyun.com/pypi/simple/ && \
    pip config set install.trusted-host mirrors.aliyun.com

# 复制requirements文件
COPY requirements.txt .

# 安装Python依赖
RUN pip install --upgrade pip && \
    pip install --only-binary :all: pydantic && \
    pip install -r requirements.txt --no-cache-dir

# 复制应用代码
COPY . .

# 暴露端口
EXPOSE 8001

# 启动命令
CMD ["python", "main.py"]

