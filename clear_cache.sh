#!/bin/bash
# 清除解析缓存脚本
#
# 默认（软失效）：把数据库中的解析记录标记为无效（is_invalid），
#   不删除 parsed_*.json 缓存文件与数据库记录，历史可追溯；
#   标记后哈希复用/同版本覆盖/历史版本列表均不再命中，重新上传即重新解析。
#
# --purge（物理删除，旧行为）：删除 parsed_*.json 文件，并可选 DELETE 数据库记录。

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DATA_DIR="$SCRIPT_DIR/backend/data"
PARSED_DIR="$DATA_DIR/parsed"
DB_FILE="$DATA_DIR/app.db"
VENV_PYTHON="$SCRIPT_DIR/backend/.venv/bin/python"

MODE="soft"
if [ "$1" = "--purge" ]; then
    MODE="purge"
fi

echo "========================================"
echo "  测试用例生成器 - 缓存清理工具"
echo "========================================"
echo ""

# 统计缓存文件数量（仅用于展示）
FILE_COUNT=0
if [ -d "$PARSED_DIR" ]; then
    FILE_COUNT=$(find "$PARSED_DIR" -name "parsed_*.json" -type f 2>/dev/null | wc -l)
fi

if [ "$MODE" = "soft" ]; then
    echo "模式：软失效（打无效标志，保留数据）"
    echo ""
    echo "当前解析缓存文件：$FILE_COUNT 个（将被保留）"
    echo ""
    read -p "是否将所有解析记录标记为无效？(y/N): " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        if [ -x "$VENV_PYTHON" ]; then
            PYTHON_BIN="$VENV_PYTHON"
        else
            PYTHON_BIN="python3"
        fi
        "$PYTHON_BIN" "$SCRIPT_DIR/scripts/invalidate_parse_cache.py"
        echo ""
        echo "清理完成！重新上传 XMind 文件将会重新解析。"
    else
        echo ""
        echo "已取消操作"
    fi
    exit 0
fi

# ---------- --purge：物理删除（旧行为） ----------
echo "模式：物理删除（--purge，数据不可恢复）"
echo ""

if [ ! -d "$PARSED_DIR" ] || [ "$FILE_COUNT" -eq 0 ]; then
    echo "✓ 没有解析缓存文件"
else
    echo "发现 $FILE_COUNT 个解析缓存文件："
    echo ""
    find "$PARSED_DIR" -name "parsed_*.json" -type f -exec basename {} \; | while read f; do
        echo "  - $f"
    done
    echo ""
fi

read -p "是否物理删除所有解析缓存文件？(y/N): " -n 1 -r
echo ""

if [[ $REPLY =~ ^[Yy]$ ]]; then
    if [ "$FILE_COUNT" -gt 0 ]; then
        find "$PARSED_DIR" -name "parsed_*.json" -type f -delete
        echo ""
        echo "✓ 已删除 $FILE_COUNT 个解析缓存文件"
    fi

    read -p "是否同时物理删除数据库中的解析记录？(y/N): " -n 1 -r
    echo ""

    if [[ $REPLY =~ ^[Yy]$ ]] && [ -f "$DB_FILE" ]; then
        BACKUP_FILE="${DB_FILE}.backup.$(date +%Y%m%d_%H%M%S)"
        cp "$DB_FILE" "$BACKUP_FILE"
        echo "✓ 已备份数据库到: $(basename "$BACKUP_FILE")"

        sqlite3 "$DB_FILE" "DELETE FROM parse_records;" 2>/dev/null || echo "⚠ 无法清理数据库记录"
        echo "✓ 已清理数据库解析记录"
    fi

    echo ""
    echo "清理完成！重新上传 XMind 文件将会重新解析。"
else
    echo ""
    echo "已取消操作"
fi
