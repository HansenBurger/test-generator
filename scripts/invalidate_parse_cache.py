"""将所有解析记录标记为失效（软删除，替代物理删除）

- 失效后哈希缓存复用、同版本号覆盖、历史版本列表均不再命中这些记录，
  重新上传 XMind 会重新解析并生成新记录
- 旧记录与 parsed_*.json 缓存文件保留，便于追溯
- 首次执行会自动为旧库补齐 is_invalid / invalidated_at 列

用法：
  backend/.venv/bin/python scripts/invalidate_parse_cache.py
"""
import os
import sys

BACKEND_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "backend")
sys.path.insert(0, BACKEND_DIR)

from app.db import init_db  # noqa: E402
from app.db import repository  # noqa: E402


def main() -> int:
    # 确保表结构含软失效列（旧库自动补齐）
    init_db()
    count = repository.invalidate_all_parse_records()
    if count == 0:
        print("✓ 没有可失效的解析记录（可能此前已全部失效）")
    else:
        print(f"✓ 已将 {count} 条解析记录标记为失效（数据保留，不再参与缓存复用）")
    print("提示：重新上传 XMind 文件将会重新解析；如需物理删除请使用 clear_cache.sh --purge")
    return 0


if __name__ == "__main__":
    sys.exit(main())
