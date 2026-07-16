#!/bin/bash
# 清除解析缓存脚本
# 用途：删除已解析的 XMind JSON 缓存，下次上传会重新解析

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DATA_DIR="$SCRIPT_DIR/backend/data"
PARSED_DIR="$DATA_DIR/parsed"
DB_FILE="$DATA_DIR/app.db"

echo "========================================"
echo "  测试用例生成器 - 缓存清理工具"
echo "========================================"
echo ""

# 检查是否存在缓存目录
if [ ! -d "$PARSED_DIR" ]; then
    echo "✓ 解析缓存目录不存在，无需清理"
    exit 0
fi

# 统计文件数量
FILE_COUNT=$(find "$PARSED_DIR" -name "parsed_*.json" -type f 2>/dev/null | wc -l)

if [ "$FILE_COUNT" -eq 0 ]; then
    echo "✓ 没有解析缓存文件"
    exit 0
fi

echo "发现 $FILE_COUNT 个解析缓存文件："
echo ""
find "$PARSED_DIR" -name "parsed_*.json" -type f -exec basename {} \; | while read f; do
    echo "  - $f"
done
echo ""

# 询问是否清除
read -p "是否清除所有解析缓存？(y/N): " -n 1 -r
echo ""

if [[ $REPLY =~ ^[Yy]$ ]]; then
    # 删除所有解析缓存
    find "$PARSED_DIR" -name "parsed_*.json" -type f -delete
    
    echo ""
    echo "✓ 已清除 $FILE_COUNT 个解析缓存文件"
    
    # 可选：清理数据库中的记录
    read -p "是否同时清理数据库中的解析记录？(y/N): " -n 1 -r
    echo ""
    
    if [[ $REPLY =~ ^[Yy]$ ]] && [ -f "$DB_FILE" ]; then
        # 备份数据库
        BACKUP_FILE="${DB_FILE}.backup.$(date +%Y%m%d_%H%M%S)"
        cp "$DB_FILE" "$BACKUP_FILE"
        echo "✓ 已备份数据库到: $(basename "$BACKUP_FILE")"
        
        # 清理解析记录
        sqlite3 "$DB_FILE" "DELETE FROM parse_records;" 2>/dev/null || echo "⚠ 无法清理数据库记录"
        echo "✓ 已清理数据库解析记录"
    fi
    
    echo ""
    echo "清理完成！重新上传 XMind 文件将会重新解析。"
else
    echo ""
    echo "已取消操作"
fi
