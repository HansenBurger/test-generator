"""
测试发放贷款文档解析
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from app.services.doc_parser import DocumentParser

def main():
    doc_path = Path(__file__).parent.parent / "tmp" / "新一代信贷系统建设项目_发放贷款用例_贷款核算组_V0.9_20251113(1).docx"
    
    if not doc_path.exists():
        print(f"文件不存在: {doc_path}")
        return
    
    try:
        parser = DocumentParser(str(doc_path))
        parsed_doc = parser.parse()
        
        print("=" * 80)
        print("解析结果：")
        print("=" * 80)
        print(f"版本: {parsed_doc.version}")
        print(f"用例名称: {parsed_doc.requirement_info.case_name}")
        print(f"活动数量: {len(parsed_doc.activities)}")
        
        if parsed_doc.activities:
            for activity in parsed_doc.activities:
                print(f"\n活动: {activity.name}")
                print(f"  组件数量: {len(activity.components)}")
                for component in activity.components:
                    print(f"  组件: {component.name}")
                    print(f"    任务数量: {len(component.tasks)}")
                    for task in component.tasks:
                        print(f"    任务: {task.name}")
                        print(f"      步骤数量: {len(task.steps)}")
                        for step in task.steps:
                            print(f"      步骤: {step.name}")
                            print(f"        输入要素数: {len(step.input_elements)}")
                            print(f"        输出要素数: {len(step.output_elements)}")
        
        # 保存JSON
        json_str = parsed_doc.model_dump_json(indent=2, ensure_ascii=False)
        output_path = Path(__file__).parent.parent / "tmp" / "loan_parsed_result.json"
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(json_str)
        print(f"\n结果已保存到: {output_path}")
        
    except Exception as e:
        print(f"解析失败: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

