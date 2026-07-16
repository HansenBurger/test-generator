#!/bin/bash
# ============================================================
# WSL 开发环境启动脚本
# 用于在 WSL 内启动并调试 test-generator 前后端
#
# 用法:
#   ./start-wsl.sh          # 正常启动前后端
#   ./start-wsl.sh --dev    # 调试模式：后端开启热重载，前端自动打开浏览器
#   ./start-wsl.sh --backend  # 仅启动后端
#   ./start-wsl.sh --frontend # 仅启动前端
#   ./start-wsl.sh --setup    # 仅安装依赖（不启动服务）
# ============================================================

set -e

# ---------- 颜色定义 ----------
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

info()  { echo -e "${CYAN}[INFO]${NC}  $*"; }
ok()    { echo -e "${GREEN}[OK]${NC}    $*"; }
warn()  { echo -e "${YELLOW}[WARN]${NC}  $*"; }
err()   { echo -e "${RED}[ERROR]${NC} $*"; }

# ---------- 项目路径 ----------
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
BACKEND_DIR="$SCRIPT_DIR/backend"
FRONTEND_DIR="$SCRIPT_DIR/frontend"
VENV_DIR="$BACKEND_DIR/.venv"

# ---------- 端口配置（可通过环境变量覆盖） ----------
BACKEND_PORT="${BACKEND_PORT:-8001}"
FRONTEND_PORT="${FRONTEND_PORT:-3000}"

# ---------- 全局进程管理 ----------
PIDS=()

cleanup() {
    echo ""
    info "正在停止服务..."
    for pid in "${PIDS[@]}"; do
        if kill -0 "$pid" 2>/dev/null; then
            info "停止进程 $pid"
            kill -TERM "$pid" 2>/dev/null || true
        fi
    done
    # 等待子进程退出
    wait 2>/dev/null
    ok "所有服务已停止"
    exit 0
}

trap cleanup SIGINT SIGTERM

# ---------- 依赖检查 ----------
check_deps() {
    local missing=0

    if ! command -v python3 &>/dev/null; then
        err "未找到 python3，请先安装 Python 3.10+"
        missing=1
    else
        info "Python: $(python3 --version 2>&1)"
    fi

    if ! command -v node &>/dev/null; then
        err "未找到 node，请先安装 Node.js 18+"
        missing=1
    else
        info "Node:   $(node --version 2>&1)"
    fi

    if ! command -v npm &>/dev/null; then
        err "未找到 npm"
        missing=1
    else
        info "npm:    $(npm --version 2>&1)"
    fi

    # 可选依赖：LibreOffice（用于 .doc 文件转换）
    if ! command -v soffice &>/dev/null; then
        warn "未安装 LibreOffice，.doc 文件转换功能将不可用"
        warn "如需支持 .doc，可运行: sudo apt install libreoffice"
    else
        info "LibreOffice: $(soffice --version 2>&1 || echo '已安装')"
    fi

    # 性能提示：WSL2 访问 /mnt/c/ 较慢
    if [[ "$SCRIPT_DIR" == /mnt/c/* ]] || [[ "$SCRIPT_DIR" == /mnt/d/* ]]; then
        echo ""
        warn "检测到项目在 Windows 文件系统（/mnt/c/ 或 /mnt/d/）上"
        warn "WSL2 跨文件系统访问性能较差，首次启动可能需要 30+ 秒"
        warn "建议：将项目复制到 WSL 原生文件系统（如 ~/projects/）以获得更好性能"
        echo ""
    fi

    if [ "$missing" -eq 1 ]; then
        exit 1
    fi
}

# ---------- 环境配置 ----------
setup_env() {
    # 复制 .env.example 到 .env（如果不存在）
    if [ ! -f "$SCRIPT_DIR/.env" ] && [ -f "$SCRIPT_DIR/.env.example" ]; then
        info "从 .env.example 创建 .env 文件"
        cp "$SCRIPT_DIR/.env.example" "$SCRIPT_DIR/.env"
        warn "请编辑 .env 文件，填写 DASHSCOPE_API_KEY 等必要配置"
    fi
}

# ---------- 后端依赖安装 ----------
setup_backend() {
    info "配置后端环境..."
    cd "$BACKEND_DIR"

    if [ ! -d "$VENV_DIR" ]; then
        info "创建 Python 虚拟环境..."
        python3 -m venv "$VENV_DIR"
        ok "虚拟环境创建完成"
    fi

    # 激活虚拟环境
    # shellcheck disable=SC1091
    source "$VENV_DIR/bin/activate"

    info "检查后端依赖..."
    pip install --upgrade pip -q 2>/dev/null || true
    pip install -r requirements.txt -q 2>/dev/null || {
        err "安装后端依赖失败"
        exit 1
    }
    # 安装调试工具
    pip install debugpy -q 2>/dev/null || true

    ok "后端依赖就绪"
}

# ---------- 前端依赖安装 ----------
setup_frontend() {
    cd "$FRONTEND_DIR"

    if [ ! -d "node_modules" ]; then
        info "安装前端依赖..."
        npm install || {
            err "安装前端依赖失败"
            exit 1
        }
        ok "前端依赖安装完成"
    else
        ok "前端依赖已就绪"
    fi
}

# ---------- 加载环境变量 ----------
load_env_file() {
    if [ -f "$SCRIPT_DIR/.env" ]; then
        info "加载 .env 配置"
        # 逐行读取，避免 source 命令的兼容性问题
        while IFS='=' read -r key value; do
            # 跳过注释和空行
            [[ "$key" =~ ^#.*$ ]] && continue
            [[ -z "$key" ]] && continue
            # 移除可能的引号
            value="${value%\"}"
            value="${value#\"}"
            value="${value%\'}"
            value="${value#\'}"
            # 导出变量
            export "$key=$value"
        done < "$SCRIPT_DIR/.env"
    fi
}

# ---------- 健康检查 ----------
wait_for_backend() {
    local max_attempts=35
    local attempt=1
    
    info "等待后端启动..."
    info "首次启动可能需要 20-30 秒（依赖加载 + LibreOffice 初始化）"
    
    while [ $attempt -le $max_attempts ]; do
        if curl -s "http://localhost:$BACKEND_PORT/api/health" > /dev/null 2>&1; then
            ok "后端已就绪（耗时 ${attempt} 秒）"
            return 0
        fi
        
        # 检查后端进程是否还在运行
        local backend_running=false
        for pid in "${PIDS[@]}"; do
            if kill -0 "$pid" 2>/dev/null; then
                backend_running=true
                break
            fi
        done
        
        if [ "$backend_running" = false ]; then
            err "后端进程已退出，请检查上方日志"
            return 1
        fi
        
        # 每 5 秒显示进度
        if [ $((attempt % 5)) -eq 0 ]; then
            info "仍在等待后端启动... (${attempt}/${max_attempts}s)"
        fi
        
        sleep 1
        attempt=$((attempt + 1))
    done
    
    warn "后端启动超时（${max_attempts}秒），仍将继续启动前端"
    warn "如果前端出现代理错误，请稍后刷新页面重试"
    return 0
}

# ---------- 启动后端 ----------
start_backend() {
    local dev_mode="$1"
    cd "$BACKEND_DIR"

    # shellcheck disable=SC1091
    source "$VENV_DIR/bin/activate"

    load_env_file

    export BACKEND_PORT
    # 强制 Python 不使用缓冲，确保日志实时输出
    export PYTHONUNBUFFERED=1

    echo ""
    echo "============================================================"
    echo "  后端日志输出"
    echo "============================================================"

    if [ "$dev_mode" = "true" ]; then
        info "启动后端（调试模式 - 热重载）..."
        python -m uvicorn main:app \
            --host 0.0.0.0 \
            --port "$BACKEND_PORT" \
            --reload \
            --log-level debug &
    else
        info "启动后端（端口 $BACKEND_PORT）..."
        python -m uvicorn main:app \
            --host 0.0.0.0 \
            --port "$BACKEND_PORT" \
            --log-level info &
    fi

    PIDS+=($!)
    info "后端 PID: ${PIDS[-1]}"
}

# ---------- 启动前端 ----------
start_frontend() {
    local dev_mode="$1"
    cd "$FRONTEND_DIR"

    export FRONTEND_PORT

    echo ""
    echo "============================================================"
    echo "  前端日志输出"
    echo "============================================================"

    if [ "$dev_mode" = "true" ]; then
        info "启动前端（调试模式 - 自动打开浏览器）..."
        npm run dev -- --open --port "$FRONTEND_PORT" &
    else
        info "启动前端（端口 $FRONTEND_PORT）..."
        npm run dev -- --port "$FRONTEND_PORT" &
    fi

    PIDS+=($!)
    info "前端 PID: ${PIDS[-1]}"
}

# ---------- 打印状态 ----------
print_status() {
    echo ""
    echo "============================================================"
    echo -e "  ${GREEN}test-generator 已启动${NC}"
    echo "============================================================"
    echo -e "  前端:  ${CYAN}http://localhost:$FRONTEND_PORT${NC}"
    echo -e "  后端:  ${CYAN}http://localhost:$BACKEND_PORT${NC}"
    echo -e "  API:   ${CYAN}http://localhost:$BACKEND_PORT/api/health${NC}"
    echo ""
    echo -e "  按 ${YELLOW}Ctrl+C${NC} 停止所有服务"
    echo "============================================================"
    echo ""
}

# ---------- 主流程 ----------
main() {
    local mode="${1:-all}"

    echo ""
    echo "============================================================"
    echo "  test-generator WSL 开发环境"
    echo "============================================================"
    echo ""

    check_deps
    setup_env

    case "$mode" in
        --dev)
            setup_backend
            setup_frontend
            echo ""
            start_backend "true"
            wait_for_backend || {
                err "后端启动失败，终止启动"
                cleanup
                exit 1
            }
            start_frontend "true"
            print_status
            ;;
        --backend)
            setup_backend
            echo ""
            start_backend "false"
            echo ""
            info "仅后端运行中，按 Ctrl+C 停止"
            echo ""
            ;;
        --frontend)
            setup_frontend
            echo ""
            start_frontend "false"
            echo ""
            info "仅前端运行中，按 Ctrl+C 停止"
            echo ""
            ;;
        --setup)
            setup_backend
            setup_frontend
            echo ""
            ok "依赖安装完成，可使用以下命令启动："
            echo "  ./start-wsl.sh         # 正常启动"
            echo "  ./start-wsl.sh --dev   # 调试模式"
            exit 0
            ;;
        *)
            setup_backend
            setup_frontend
            echo ""
            start_backend "false"
            wait_for_backend || {
                err "后端启动失败，终止启动"
                cleanup
                exit 1
            }
            start_frontend "false"
            print_status
            ;;
    esac

    # 等待所有子进程
    wait
}

main "$@"
