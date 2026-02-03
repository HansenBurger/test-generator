#!/bin/bash

# 测试大纲生成器部署脚本
# Docker镜像加速地址: https://5f4mc5ba.mirror.aliyuncs.com

set -e

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
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

# 构建镜像
build_images() {
    print_info "开始构建Docker镜像..."
    COMPOSE_CMD=$(detect_compose_cmd)
    $COMPOSE_CMD build
    print_info "镜像构建完成"
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

# 主菜单
show_menu() {
    echo ""
    echo "=========================================="
    echo "  测试大纲生成器 - Docker部署脚本"
    echo "=========================================="
    echo "1. 配置Docker镜像加速"
    echo "2. 构建镜像"
    echo "3. 启动服务"
    echo "4. 停止服务"
    echo "5. 重启服务"
    echo "6. 查看状态"
    echo "7. 查看日志"
    echo "8. 一键部署（配置+构建+启动）"
    echo "9. 清理所有（容器+镜像+数据卷）"
    echo "0. 退出"
    echo "=========================================="
    echo -n "请选择操作 [0-9]: "
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
            print_info "前端访问地址: http://localhost:3000"
            print_info "后端API地址: http://localhost:8001"
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
                        print_info "前端访问地址: http://localhost:3000"
                        print_info "后端API地址: http://localhost:8001"
                        ;;
                    9)
                        clean
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

