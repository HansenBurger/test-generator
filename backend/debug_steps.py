"""
调试步骤提取逻辑
"""
import sys
from pathlib import Path
from docx import Document
import re

sys.path.insert(0, str(Path(__file__).parent))

def is_heading(para, level):
    style_name = para.style.name
    if style_name.startswith('Heading'):
        try:
            heading_level = int(style_name.replace('Heading ', ''))
            return heading_level == level
        except:
            return False
    return False

def main():
    doc_path = Path(__file__).parent.parent / "tmp" / "新一代信贷系统建设项目_发放贷款用例_贷款核算组_V0.9_20251113(1).docx"
    
    doc = Document(str(doc_path))
    paragraphs = [p for p in doc.paragraphs]
    
    # 查找任务"发放贷款"的位置
    task_index = -1
    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if "发放贷款" in text and "（A阶段、B阶段）" in text and is_heading(para, 3):
            task_index = i
            print(f"找到任务位置: {i}")
            break
    
    if task_index == -1:
        print("未找到任务")
        return
    
    # 模拟步骤提取逻辑
    exclude_keywords = ["任务规则说明", "输入输出", "业务流程", "业务规则",
                       "页面控制", "数据验证", "前置条件", "后置条件",
                       "任务-业务步骤/功能清单", "业务步骤/功能描述", "规则说明",
                       "错误处理", "权限说明", "用户操作注释"]
    
    steps = []
    start_index = task_index + 1
    
    print(f"\n从索引 {start_index} 开始查找步骤 (范围: {start_index} 到 {min(start_index + 200, len(paragraphs))}):")
    print("=" * 80)
    
    checked_count = 0
    for i in range(start_index, min(start_index + 200, len(paragraphs))):
        para = paragraphs[i]
        text = para.text.strip()
        checked_count += 1
        
        # 如果遇到下一个三级标题（新的任务），停止搜索
        if is_heading(para, 3) and "发放贷款" not in text:
            print(f"[{i:3d}] 遇到新任务，停止: {text[:50]}")
            break
        
        # 如果遇到二级标题（新的组件），停止搜索
        if is_heading(para, 2):
            print(f"[{i:3d}] 遇到新组件，停止: {text[:50]}")
            break
        
        # 如果遇到一级标题，停止搜索
        if is_heading(para, 1):
            print(f"[{i:3d}] 遇到一级标题，停止: {text[:50]}")
            break
        
        # 检查是否是四级标题（步骤名称）
        if is_heading(para, 4):
            match = re.match(r"(.+?)\*?（A阶段、B阶段）", text)
            if match:
                step_name = match.group(1).strip()
                print(f"[{i:3d}] 发现四级标题: {step_name} (排除关键词: {step_name in exclude_keywords})")
                if step_name not in exclude_keywords:
                    print(f"[{i:3d}] ✓ 找到步骤: {step_name}")
                    steps.append(step_name)
                else:
                    print(f"[{i:3d}] ✗ 跳过（在排除列表中）")
    
    print(f"\n总共检查了 {checked_count} 个段落")
    print(f"总共找到 {len(steps)} 个步骤:")
    for i, step in enumerate(steps):
        print(f"  {i+1}. {step}")

if __name__ == "__main__":
    main()

