"""
测试XMind生成
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from app.services.doc_parser import DocumentParser
from app.services.xmind_generator import XMindGenerator

def main():
    # 文档路径
    doc_path = Path(__file__).parent.parent / "tmp" / "新一代信贷系统建设项目_管理特色互联网贷款账单用例_贷款核算组_V0.9_20251028.docx"
    
    if not doc_path.exists():
        print(f"文件不存在: {doc_path}")
        return
    
    try:
        # 解析文档
        print("正在解析文档...")
        parser = DocumentParser(str(doc_path))
        parsed_doc = parser.parse()
        
        # 打印解析结果摘要
        print("\n解析结果摘要：")
        print(f"  版本: {parsed_doc.version}")
        print(f"  用例名称: {parsed_doc.requirement_info.case_name}")
        print(f"  客户: {parsed_doc.requirement_info.customer}")
        print(f"  活动数量: {len(parsed_doc.activities)}")
        if parsed_doc.activities:
            for i, activity in enumerate(parsed_doc.activities):
                print(f"    活动{i+1}: {activity.name}")
                print(f"      组件数量: {len(activity.components)}")
                for j, component in enumerate(activity.components):
                    print(f"        组件{j+1}: {component.name}")
                    print(f"          任务数量: {len(component.tasks)}")
                    for k, task in enumerate(component.tasks):
                        print(f"            任务{k+1}: {task.name}")
                        print(f"              步骤数量: {len(task.steps)}")
        
        # 生成XMind
        print("\n正在生成XMind文件...")
        generator = XMindGenerator(parsed_doc)
        xmind_bytes = generator.generate()
        
        # 保存文件
        output_path = Path(__file__).parent.parent / "tmp" / "test_output.xmind"
        with open(output_path, 'wb') as f:
            f.write(xmind_bytes)
        
        print(f"\nXMind文件已生成: {output_path}")
        print(f"文件大小: {len(xmind_bytes)} 字节")
        
        # 检查根节点标题
        root_title = generator._build_root_title()
        print(f"\n根节点标题: {root_title}")
        
    except Exception as e:
        print(f"处理失败: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

