"""
测试解析器，输出JSON结果
"""
import json
import sys
from pathlib import Path

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

from app.services.doc_parser import DocumentParser
from app.models.schemas import ParsedDocument

def main():
    # 文档路径
    doc_path = Path(__file__).parent.parent / "tmp" / "新一代信贷系统建设项目_管理特色互联网贷款账单用例_贷款核算组_V0.9_20251028.docx"
    
    if not doc_path.exists():
        print(f"文件不存在: {doc_path}")
        return
    
    try:
        # 解析文档
        parser = DocumentParser(str(doc_path))
        parsed_doc = parser.parse()
        
        # 转换为JSON
        json_str = parsed_doc.model_dump_json(indent=2, ensure_ascii=False)
        
        print("=" * 80)
        print("解析结果（JSON格式）：")
        print("=" * 80)
        print(json_str)
        print("=" * 80)
        
        # 也保存到文件
        output_path = Path(__file__).parent.parent / "tmp" / "parsed_result.json"
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(json_str)
        print(f"\n结果已保存到: {output_path}")
        
    except Exception as e:
        print(f"解析失败: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

