#!/bin/bash

# 测试大纲生成器部署脚本
# Docker镜像加速地址: https://5f4mc5ba.mirror.aliyuncs.com

set -e

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# 打印带颜色的消息
print_info() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

print_warn() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

print_step() {
    echo -e "${CYAN}[STEP]${NC} $1"
}

# 检查Docker是否安装
check_docker() {
    if ! command -v docker &> /dev/null; then
        print_error "Docker未安装，请先安装Docker"
        exit 1
    fi
    print_info "Docker已安装: $(docker --version)"
}

# 检测Docker Compose命令（支持V1和V2）
detect_compose_cmd() {
    if docker compose version &> /dev/null; then
        echo "docker compose"
    elif command -v docker-compose &> /dev/null; then
        echo "docker-compose"
    else
        print_error "Docker Compose未安装，请先安装Docker Compose"
        exit 1
    fi
}

# 检查Docker Compose是否安装
check_docker_compose() {
    COMPOSE_CMD=$(detect_compose_cmd)
    if [ "$COMPOSE_CMD" = "docker compose" ]; then
        print_info "Docker Compose已安装: $(docker compose version)"
    else
        print_info "Docker Compose已安装: $(docker-compose --version)"
    fi
}

# 配置Docker镜像加速
configure_docker_mirror() {
    print_info "配置Docker镜像加速..."
    
    DOCKER_DAEMON_JSON="/etc/docker/daemon.json"
    MIRROR_URL="https://5f4mc5ba.mirror.aliyuncs.com"
    
    # 检查是否已配置
    if [ -f "$DOCKER_DAEMON_JSON" ]; then
        if grep -q "$MIRROR_URL" "$DOCKER_DAEMON_JSON" 2>/dev/null; then
            print_info "Docker镜像加速已配置"
            return
        fi
    fi
    
    # 创建或更新daemon.json
    if [ ! -f "$DOCKER_DAEMON_JSON" ]; then
        sudo mkdir -p /etc/docker
        echo "{}" | sudo tee "$DOCKER_DAEMON_JSON" > /dev/null
    fi
    
    # 备份原配置
    sudo cp "$DOCKER_DAEMON_JSON" "${DOCKER_DAEMON_JSON}.bak.$(date +%Y%m%d_%H%M%S)"
    
    # 添加镜像加速配置
    sudo python3 << EOF
import json
import sys

try:
    with open('$DOCKER_DAEMON_JSON', 'r') as f:
        config = json.load(f)
except:
    config = {}

if 'registry-mirrors' not in config:
    config['registry-mirrors'] = []

if '$MIRROR_URL' not in config['registry-mirrors']:
    config['registry-mirrors'].append('$MIRROR_URL')

with open('$DOCKER_DAEMON_JSON', 'w') as f:
    json.dump(config, f, indent=2, ensure_ascii=False)

print("配置已更新")
EOF
    
    print_info "Docker镜像加速配置已添加，需要重启Docker服务"
    print_warn "请运行以下命令重启Docker: sudo systemctl restart docker"
}

# 检查 DASHSCOPE_API_KEY 是否已配置（.env 或环境变量）
check_dashscope_key() {
    local key_value=""
    if [ -f ".env" ]; then
        key_value=$(grep -E "^DASHSCOPE_API_KEY=.+" .env 2>/dev/null | cut -d= -f2- | tr -d "\"\\'")
    fi
    if [ -z "$key_value" ]; then
        key_value="${DASHSCOPE_API_KEY:-}"
    fi
    if [ -z "$key_value" ]; then
        print_error "未配置 DASHSCOPE_API_KEY，后端将无法启动"
        echo ""
        echo "请按以下步骤配置："
        echo "  1. 复制示例: cp .env.example .env"
        echo "  2. 编辑 .env，填入阿里云百炼 API Key: DASHSCOPE_API_KEY=sk-你的密钥"
        echo "  或导出环境变量: export DASHSCOPE_API_KEY=sk-你的密钥"
        echo ""
        return 1
    fi
    return 0
}

# 读取 .env 或环境变量的配置值
get_env_value() {
    local key="$1"
    local value=""
    if [ -f ".env" ]; then
        value=$(grep -E "^${key}=" .env 2>/dev/null | tail -n 1 | cut -d= -f2- | tr -d "\"\\'")
    fi
    if [ -z "$value" ]; then
        value="${!key:-}"
    fi
    echo "$value"
}

# 构建镜像
build_images() {
    print_info "开始构建Docker镜像..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD build
    print_info "镜像构建完成"
}

# 仅构建前端镜像
build_frontend() {
    print_info "开始构建前端镜像..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD build frontend
    print_info "前端镜像构建完成"
}

# 仅构建后端镜像
build_backend() {
    print_info "开始构建后端镜像..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD build backend
    print_info "后端镜像构建完成"
}

# 启动服务
start_services() {
    if ! check_dashscope_key; then
        exit 1
    fi
    print_info "启动服务..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD up -d
    print_info "服务启动完成"
}

# 启动服务（不构建镜像，适用于内网）
start_services_no_build() {
    if ! check_dashscope_key; then
        exit 1
    fi
    print_info "启动服务（不构建镜像）..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD up -d --no-build
    print_info "服务启动完成"
}

# 导出镜像（外网构建后打包 - 全量）
export_images() {
    build_images
    print_info "导出全量镜像到本地文件..."
    docker save -o test-generator-images.tar test-generator-backend:latest test-generator-frontend:latest
    local size=$(du -h test-generator-images.tar | cut -f1)
    print_info "导出完成：test-generator-images.tar (${size})"
}

# 仅导出前端镜像（适用于前端代码变更）
export_frontend() {
    print_step "构建前端镜像..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD build frontend
    print_step "导出前端镜像..."
    docker save -o test-generator-frontend.tar test-generator-frontend:latest
    local size=$(du -h test-generator-frontend.tar | cut -f1)
    print_info "前端镜像导出完成：test-generator-frontend.tar (${size})"
    echo ""
    print_info "将以下文件拷贝到内网服务器后执行:"
    echo "  scp test-generator-frontend.tar <user>@<host>:<path>/"
    echo "  # 在内网服务器上："
    echo "  ./deploy.sh import-frontend"
}

# 仅导出后端镜像（适用于后端代码变更）
export_backend() {
    print_step "构建后端镜像..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD build backend
    print_step "导出后端镜像..."
    docker save -o test-generator-backend.tar test-generator-backend:latest
    local size=$(du -h test-generator-backend.tar | cut -f1)
    print_info "后端镜像导出完成：test-generator-backend.tar (${size})"
    echo ""
    print_info "将以下文件拷贝到内网服务器后执行:"
    echo "  scp test-generator-backend.tar <user>@<host>:<path>/"
    echo "  # 在内网服务器上："
    echo "  ./deploy.sh import-backend"
}

# 构建前端 dist 并打包（不构建镜像，适用于挂载模式的热更新）
export_dist() {
    print_step "检查前端依赖..."
    if [ ! -d "frontend/node_modules" ]; then
        print_info "安装前端依赖..."
        (cd frontend && npm ci --prefer-offline --no-audit)
    fi
    print_step "构建前端 dist..."
    (cd frontend && npm run build)
    print_step "打包 dist..."
    tar -czf frontend-dist.tar.gz -C frontend dist
    local size=$(du -h frontend-dist.tar.gz | cut -f1)
    print_info "前端 dist 打包完成：frontend-dist.tar.gz (${size})"
    echo ""
    print_info "将以下文件拷贝到内网服务器后执行:"
    echo "  scp frontend-dist.tar.gz <user>@<host>:<path>/frontend/"
    echo "  # 在内网服务器上："
    echo "  ./deploy.sh update-dist"
}

# 导入镜像（内网离线部署 - 全量）
import_images() {
    if [ ! -f "test-generator-images.tar" ]; then
        print_error "未找到 test-generator-images.tar，请先拷贝镜像包到当前目录"
        exit 1
    fi
    print_info "导入全量镜像..."
    docker load -i test-generator-images.tar
    print_info "导入完成"
}

# 仅导入前端镜像并重启前端服务
import_frontend() {
    if [ ! -f "test-generator-frontend.tar" ]; then
        print_error "未找到 test-generator-frontend.tar，请先拷贝前端镜像包到当前目录"
        exit 1
    fi
    print_step "导入前端镜像..."
    docker load -i test-generator-frontend.tar
    print_step "重启前端容器..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD up -d --no-deps --force-recreate frontend
    print_info "前端更新完成！"
}

# 仅导入后端镜像并重启后端服务
import_backend() {
    if [ ! -f "test-generator-backend.tar" ]; then
        print_error "未找到 test-generator-backend.tar，请先拷贝后端镜像包到当前目录"
        exit 1
    fi
    print_step "导入后端镜像..."
    docker load -i test-generator-backend.tar
    print_step "重启后端容器..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD up -d --no-deps --force-recreate backend
    print_info "后端更新完成！"
}

# 更新前端 dist（挂载模式，零镜像操作）
update_dist() {
    if [ ! -f "frontend-dist.tar.gz" ] && [ ! -d "frontend/dist" ]; then
        print_error "未找到 frontend-dist.tar.gz 或 frontend/dist，请先构建或拷贝"
        exit 1
    fi

    if [ -f "frontend-dist.tar.gz" ]; then
        print_step "解压 dist..."
        rm -rf frontend/dist
        tar -xzf frontend-dist.tar.gz -C frontend
    fi

    print_step "将 dist 同步到前端容器..."
    # 如果容器正在运行，用 docker cp 直接更新
    if docker ps --format '{{.Names}}' | grep -q 'test-generator-frontend'; then
        # 先清除旧的静态资源
        docker exec test-generator-frontend rm -rf /usr/share/nginx/html/assets
        # 拷入新文件
        docker cp frontend/dist/. test-generator-frontend:/usr/share/nginx/html/
        # 重载 nginx 配置
        docker exec test-generator-frontend nginx -s reload
        print_info "前端 dist 热更新完成！（无需重建镜像）"
    else
        print_warn "前端容器未运行，请先启动服务"
        print_info "  ./deploy.sh start"
    fi
}

# 停止服务
stop_services() {
    print_info "停止服务..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD down
    print_info "服务已停止"
}

# 查看日志
view_logs() {
    print_info "查看服务日志..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD logs -f
}

# 查看状态
view_status() {
    print_info "服务状态:"
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD ps
}

# 重启服务
restart_services() {
    print_info "重启服务..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD restart
    print_info "服务重启完成"
}

# 清理
clean() {
    print_warn "这将删除所有容器、镜像和数据卷，确定要继续吗？(y/N)"
    read -r response
    if [[ "$response" =~ ^([yY][eE][sS]|[yY])$ ]]; then
        print_info "清理中..."
        COMPOSE_CMD=$(detect_compose_cmd)
        $COMPOSE_CMD down -v --rmi all
        print_info "清理完成"
    else
        print_info "已取消清理"
    fi
}

# 打印部署地址
print_addresses() {
    FRONTEND_PORT=$(get_env_value "FRONTEND_PORT")
    BACKEND_PORT=$(get_env_value "BACKEND_PORT")
    print_info "前端访问地址: http://localhost:${FRONTEND_PORT:-3000}"
    print_info "后端API地址: http://localhost:${BACKEND_PORT:-8001}"
}

# 主菜单
show_menu() {
    echo ""
    echo "=========================================="
    echo "  测试大纲生成器 - Docker部署脚本"
    echo "=========================================="
    echo ""
    echo "  ── 基础操作 ──"
    echo "1.  配置Docker镜像加速"
    echo "2.  构建镜像（全部）"
    echo "3.  启动服务"
    echo "4.  停止服务"
    echo "5.  重启服务"
    echo "6.  查看状态"
    echo "7.  查看日志"
    echo "8.  一键部署（配置+构建+启动）"
    echo "9.  清理所有（容器+镜像+数据卷）"
    echo ""
    echo "  ── 全量镜像导出/导入 ──"
    echo "10. 导出全量镜像（外网构建后打包）"
    echo "11. 导入全量镜像并启动（内网部署）"
    echo ""
    echo "  ── 增量更新（推荐）──"
    echo "12. 仅构建并导出前端镜像"
    echo "13. 仅构建并导出后端镜像"
    echo "14. 导入前端镜像并更新（内网）"
    echo "15. 导入后端镜像并更新（内网）"
    echo ""
    echo "  ── 前端热更新（最快）──"
    echo "16. 构建并打包前端 dist"
    echo "17. 更新前端 dist（内网，无需镜像）"
    echo ""
    echo "0.  退出"
    echo "=========================================="
    echo -n "请选择操作: "
}

# 主函数
main() {
    # 检查基本环境
    check_docker
    check_docker_compose
    
    # 如果提供了参数，直接执行对应操作
    case "${1:-}" in
        build)
            build_images
            ;;
        build-frontend)
            build_frontend
            ;;
        build-backend)
            build_backend
            ;;
        start)
            start_services
            ;;
        stop)
            stop_services
            ;;
        restart)
            restart_services
            ;;
        status)
            view_status
            ;;
        logs)
            view_logs
            ;;
        clean)
            clean
            ;;
        deploy)
            if ! check_dashscope_key; then exit 1; fi
            configure_docker_mirror
            build_images
            start_services
            view_status
            print_info "部署完成！"
            print_addresses
            ;;
        export-images)
            export_images
            ;;
        import-images)
            import_images
            start_services_no_build
            view_status
            print_info "内网部署完成！"
            print_addresses
            ;;
        # ── 增量更新命令 ──
        export-frontend)
            export_frontend
            ;;
        export-backend)
            export_backend
            ;;
        import-frontend)
            import_frontend
            ;;
        import-backend)
            import_backend
            ;;
        # ── 前端热更新命令 ──
        export-dist)
            export_dist
            ;;
        update-dist)
            update_dist
            ;;
        *)
            # 交互式菜单
            while true; do
                show_menu
                read -r choice
                case $choice in
                    1)
                        configure_docker_mirror
                        ;;
                    2)
                        build_images
                        ;;
                    3)
                        start_services
                        ;;
                    4)
                        stop_services
                        ;;
                    5)
                        restart_services
                        ;;
                    6)
                        view_status
                        ;;
                    7)
                        view_logs
                        ;;
                    8)
                        if ! check_dashscope_key; then continue; fi
                        configure_docker_mirror
                        build_images
                        start_services
                        view_status
                        print_info "部署完成！"
                        print_addresses
                        ;;
                    9)
                        clean
                        ;;
                    10)
                        export_images
                        ;;
                    11)
                        import_images
                        start_services_no_build
                        view_status
                        print_info "内网部署完成！"
                        print_addresses
                        ;;
                    12)
                        export_frontend
                        ;;
                    13)
                        export_backend
                        ;;
                    14)
                        import_frontend
                        ;;
                    15)
                        import_backend
                        ;;
                    16)
                        export_dist
                        ;;
                    17)
                        update_dist
                        ;;
                    0)
                        print_info "退出"
                        exit 0
                        ;;
                    *)
                        print_error "无效选择，请重新输入"
                        ;;
                esac
                echo ""
            done
            ;;
    esac
}

# 执行主函数
main "$@"
