"""
Word文档解析服务 - 银行需求文档专用解析器
"""
import re
from typing import List, Optional, Dict, Tuple
from docx import Document
from docx.document import Document as DocumentType
from docx.table import Table
from docx.text.paragraph import Paragraph

from app.models.schemas import (
    ParsedDocument, RequirementInfo, ActivityInfo, ComponentInfo,
    TaskInfo, StepInfo, InputElement, OutputElement, FunctionInfo
)


class DocumentParser:
    """文档解析器 - 针对银行需求文档格式"""
    
    def __init__(self, doc_path: str):
        self.doc = Document(doc_path)
        self.paragraphs = [p for p in self.doc.paragraphs]
        self.tables = self.doc.tables
        self.used_tables = set()  # 记录已使用的表格索引，避免重复使用
    
    def parse(self) -> ParsedDocument:
        """解析文档主方法"""
        # 1. 识别文档类型
        doc_type = self._identify_document_type()
        
        if doc_type == "modeling":
            return self._parse_modeling_document()
        elif doc_type == "non_modeling":
            return self._parse_non_modeling_document()
        else:
            raise ValueError("无法识别文档类型：未找到'用例版本控制信息'表或'文件受控信息'表")
    
    def _identify_document_type(self) -> Optional[str]:
        """识别文档类型：建模需求或非建模需求"""
        # 优先查找"用例版本控制信息"（建模需求特征）
        for para in self.paragraphs[:100]:
            text = para.text.strip()
            if "用例版本控制信息" in text:
                # 检查是否有包含"版本"字段的表格
                for table in self.tables:
                    if len(table.rows) < 1:
                        continue
                    header_row = table.rows[0]
                    header_text = ' '.join([cell.text.strip() for cell in header_row.cells])
                    if "版本" in header_text:
                        return "modeling"
        
        # 查找"文件受控信息"（非建模需求特征）
        for para in self.paragraphs[:100]:
            text = para.text.strip()
            if "文件受控信息" in text:
                # 检查是否有包含"文件编号"或"文件名称"字段的表格
                for table in self.tables:
                    if len(table.rows) < 1:
                        continue
                    header_row = table.rows[0]
                    header_text = ' '.join([cell.text.strip() for cell in header_row.cells])
                    if "文件编号" in header_text or "文件名称" in header_text:
                        return "non_modeling"
        
        return None
    
    def _parse_modeling_document(self) -> ParsedDocument:
        """解析建模需求文档"""
        # 1. 提取版本编号
        version = self._extract_version()
        if not version:
            raise ValueError("无法提取版本信息：未找到'用例版本控制信息'表或表中无版本数据")
        
        # 2. 提取需求用例基本信息
        requirement_info = self._extract_requirement_info()
        if not requirement_info.case_name:
            raise ValueError("无法提取需求基本信息：未找到'需求用例概述'表或表中无用例名称")
        
        # 3. 提取活动名称（从任务设计部分）
        activity_name = self._extract_activity_name()
        
        # 4. 提取组件、任务、步骤信息（从任务规则说明部分）
        components = self._extract_all_components()
        
        # 构建活动信息（组件信息放在活动下，但XMind生成时会分别处理）
        activity = ActivityInfo(
            name=activity_name or "",
            components=components
        )
        
        return ParsedDocument(
            version=version,
            requirement_info=requirement_info,
            activities=[activity] if activity_name or components else [],
            document_type="modeling"
        )
    
    def _parse_non_modeling_document(self) -> ParsedDocument:
        """解析非建模需求文档"""
        # 1. 提取文件编号和文件名称
        file_number, file_name = self._extract_file_controlled_info()
        
        # 2. 提取需求名称
        requirement_name = self._extract_requirement_name(file_name)
        
        # 3. 提取设计者
        designer = self._extract_designer()
        
        # 4. 提取功能列表
        functions = self._extract_functions()
        
        if not functions:
            raise ValueError("无法提取功能列表：未找到'功能清单'表或表中无功能数据")
        
        # 构建需求基本信息（使用需求名称作为用例名称）
        requirement_info = RequirementInfo(case_name=requirement_name or "")
        
        return ParsedDocument(
            version=file_number or "",  # 使用文件编号作为版本
            requirement_info=requirement_info,
            activities=[],
            document_type="non_modeling",
            file_number=file_number,
            file_name=file_name,
            requirement_name=requirement_name,
            designer=designer,
            functions=functions
        )
    
    def _validate_document(self) -> bool:
        """验证文档是否包含用例版本控制信息表"""
        # 检查前50个段落中是否包含"用例版本控制信息"
        for para in self.paragraphs[:50]:
            text = para.text.strip()
            if "用例版本控制信息" in text:
                # 检查是否有包含"版本"字段的表格
                for table in self.tables:
                    if len(table.rows) < 1:
                        continue
                    header_row = table.rows[0]
                    header_text = ' '.join([cell.text.strip() for cell in header_row.cells])
                    if "版本" in header_text:
                        return True
        return False
    
    def _extract_version(self) -> str:
        """从用例版本控制信息表提取版本编号"""
        for table in self.tables:
            if len(table.rows) < 2:  # 至少要有表头和数据行
                continue
            
            # 检查表头是否包含"版本"字段
            header_row = table.rows[0]
            header_text = ' '.join([cell.text.strip() for cell in header_row.cells])
            
            if "版本" in header_text:
                # 找到版本列的索引（通常是第一列）
                version_col_idx = 0
                for idx, cell in enumerate(header_row.cells):
                    if "版本" in cell.text.strip():
                        version_col_idx = idx
                        break
                
                # 取最后非空行的版本列值
                for row in reversed(table.rows[1:]):  # 跳过表头
                    if len(row.cells) > version_col_idx:
                        version = row.cells[version_col_idx].text.strip()
                        if version:
                            return version
        return ""
    
    def _extract_requirement_info(self) -> RequirementInfo:
        """提取需求用例基本信息"""
        info = RequirementInfo(case_name="")
        
        # 查找包含"需求用例概述"的段落，然后查找下面的表格
        found_overview = False
        for i, para in enumerate(self.paragraphs):
            text = para.text.strip()
            if "需求用例概述" in text and "（A阶段）" in text:
                found_overview = True
                # 查找后续的表格
                for table in self.tables:
                    if len(table.rows) < 1:
                        continue
                    
                    # 检查表格是否包含"用例名称"
                    first_row_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
                    if "用例名称" in first_row_text:
                        # 解析表格内容
                        self._parse_requirement_table(table, info)
                        break
                break
        
        return info
    
    def _parse_requirement_table(self, table: Table, info: RequirementInfo):
        """解析需求用例概述表格"""
        if len(table.rows) < 1:
            return
        
        # 获取表头
        header_row = table.rows[0]
        headers = [cell.text.strip() for cell in header_row.cells]
        
        # 特殊处理：如果表头中"用例名称"后面直接跟着值（如：['用例名称', '管理特色互联网贷款账单', ...]）
        case_name_idx = self._find_column_index(headers, ["用例名称"])
        if case_name_idx >= 0 and case_name_idx + 1 < len(headers):
            # 检查用例名称列后面的列是否是值（不是其他字段名）
            next_cell = headers[case_name_idx + 1]
            if next_cell and next_cell != '/' and "用例" not in next_cell and "名称" not in next_cell:
                # 这可能是用例名称的值
                if not info.case_name:
                    info.case_name = next_cell
        
        # 查找各字段的列索引
        channel_idx = self._find_column_index(headers, ["渠道（C）", "渠道"])
        product_idx = self._find_column_index(headers, ["产品（P）", "P产品（P）", "产品"])
        customer_idx = self._find_column_index(headers, ["客户（C）", "客户"])
        partner_idx = self._find_column_index(headers, ["合作方（P）", "P合作方（P）", "合作方"])
        
        # 解析数据行（可能是横向布局：第一行是表头，第二行是值）
        if len(table.rows) >= 2:
            value_row = table.rows[1]
            values = [cell.text.strip() for cell in value_row.cells]
            
            if channel_idx >= 0 and channel_idx < len(values):
                value = values[channel_idx]
                if value and value != '/':
                    info.channel = value
            
            if product_idx >= 0 and product_idx < len(values):
                value = values[product_idx]
                if value and value != '/':
                    info.product = value
            
            if customer_idx >= 0 and customer_idx < len(values):
                value = values[customer_idx]
                if value and value != '/':
                    info.customer = value
            
            if partner_idx >= 0 and partner_idx < len(values):
                value = values[partner_idx]
                if value and value != '/':
                    info.partner = value
        
        # 也尝试纵向布局：第一列是键，第二列是值
        for row in table.rows[1:]:
            if len(row.cells) >= 2:
                key = row.cells[0].text.strip()
                value = row.cells[1].text.strip()
                
                if value and value != '/':
                    if "用例名称" in key and not info.case_name:
                        info.case_name = value
                    elif ("渠道" in key and "（C）" in key) and not info.channel:
                        info.channel = value
                    elif (("产品" in key and "（P）" in key) or "P产品（P）" in key) and not info.product:
                        info.product = value
                    elif ("客户" in key and "（C）" in key) and not info.customer:
                        info.customer = value
                    elif (("合作方" in key and "（P）" in key) or "P合作方（P）" in key) and not info.partner:
                        info.partner = value
    
    def _extract_activity_name(self) -> Optional[str]:
        """提取活动名称：从'# 任务设计*（A阶段）'部分提取第一个子标题"""
        for i, para in enumerate(self.paragraphs):
            text = para.text.strip()
            
            # 查找"任务设计*（A阶段）"标题（一级标题）
            if "任务设计" in text and "（A阶段）" in text:
                # 检查是否是标题样式（Heading 1）
                if self._is_heading(para, level=1):
                    # 查找下一个二级标题（##级别）
                    for j in range(i + 1, min(i + 50, len(self.paragraphs))):
                        next_para = self.paragraphs[j]
                        next_text = next_para.text.strip()
                        
                        # 如果遇到下一个一级标题，停止搜索
                        if self._is_heading(next_para, level=1) and "任务设计" not in next_text:
                            break
                        
                        # 检查是否是二级标题且包含"（A阶段）"
                        if self._is_heading(next_para, level=2):
                            # 匹配模式：活动名称*（A阶段）
                            match = re.match(r"(.+?)\*?（A阶段）", next_text)
                            if match:
                                activity_name = match.group(1).strip()
                                # 排除特定关键词
                                exclude_keywords = ["需求用例概述", "活动任务图", "业务流程图", 
                                                   "任务设计", "业务步骤/功能描述", "规则说明",
                                                   "任务清单", "任务流程图", "流程描述"]
                                if activity_name not in exclude_keywords:
                                    return activity_name
        
        return None
    
    def _extract_all_components(self) -> List[ComponentInfo]:
        """提取所有组件、任务、步骤信息：从'# 任务规则说明*（A阶段、B阶段）'部分提取"""
        components = []
        exclude_keywords = ["任务规则说明", "输入输出", "业务流程", "业务规则",
                           "页面控制", "数据验证", "前置条件", "后置条件",
                           "任务-业务步骤/功能清单", "业务步骤/功能描述", "规则说明",
                           "错误处理", "权限说明", "用户操作注释"]
        
        for i, para in enumerate(self.paragraphs):
            text = para.text.strip()
            
            # 查找"任务规则说明*（A阶段、B阶段）"标题（一级标题）
            if "任务规则说明" in text and "（A阶段、B阶段）" in text:
                if self._is_heading(para, level=1):
                    # 首先找到搜索的结束位置（下一个一级标题）
                    end_index = len(self.paragraphs)  # 默认到文档末尾
                    
                    for j in range(i + 1, len(self.paragraphs)):
                        next_para = self.paragraphs[j]
                        next_text = next_para.text.strip()
                        
                        # 如果遇到下一个一级标题，停止搜索
                        if self._is_heading(next_para, level=1) and "任务规则说明" not in next_text:
                            end_index = j
                            break
                    
                    # 在确定的范围内查找所有组件名称（二级标题：##级别）
                    for j in range(i + 1, end_index):
                        next_para = self.paragraphs[j]
                        next_text = next_para.text.strip()
                        
                        # 检查是否是二级标题（组件名称）
                        if self._is_heading(next_para, level=2):
                            match = re.match(r"(.+?)\*?（A阶段、B阶段）", next_text)
                            if match:
                                component_name = match.group(1).strip()
                                if component_name not in exclude_keywords:
                                    # 提取该组件下的所有任务
                                    tasks = self._extract_tasks(j + 1, component_name, exclude_keywords)
                                    component = ComponentInfo(name=component_name, tasks=tasks)
                                    components.append(component)
        
        return components
    
    def _extract_tasks(self, start_index: int, component_name: str, exclude_keywords: List[str]) -> List[TaskInfo]:
        """提取任务列表（从组件名称后开始）
        
        使用动态边界检测，自动找到下一个组件或一级标题作为结束位置
        """
        tasks = []
        current_task = None
        
        # 首先找到搜索的结束位置（下一个组件或一级标题）
        end_index = len(self.paragraphs)  # 默认到文档末尾
        
        for i in range(start_index + 1, len(self.paragraphs)):
            para = self.paragraphs[i]
            text = para.text.strip()
            
            # 如果遇到下一个二级标题（新的组件），停止搜索
            if self._is_heading(para, level=2) and component_name not in text:
                end_index = i
                break
            
            # 如果遇到一级标题，停止搜索
            if self._is_heading(para, level=1):
                end_index = i
                break
        
        # 在确定的范围内搜索任务
        for i in range(start_index, end_index):
            para = self.paragraphs[i]
            text = para.text.strip()
            
            # 检查是否是三级标题（任务名称）
            if self._is_heading(para, level=3):
                match = re.match(r"(.+?)\*?（A阶段、B阶段）", text)
                if match:
                    task_name = match.group(1).strip()
                    if task_name not in exclude_keywords:
                        # 保存上一个任务
                        if current_task:
                            tasks.append(current_task)
                        
                        # 创建新任务
                        current_task = TaskInfo(name=task_name, steps=[])
                        # 提取该任务下的所有步骤
                        steps = self._extract_steps(i + 1, task_name, exclude_keywords)
                        current_task.steps = steps
        
        # 添加最后一个任务
        if current_task:
            tasks.append(current_task)
        
        return tasks
    
    def _extract_steps(self, start_index: int, task_name: str, exclude_keywords: List[str]) -> List[StepInfo]:
        """提取步骤列表（从任务名称后开始）
        
        使用动态边界检测，自动找到下一个任务/组件/一级标题作为结束位置
        这样可以适应任意长度的文档，不需要固定搜索范围
        """
        steps = []
        
        # 首先找到搜索的结束位置（下一个任务、组件或一级标题）
        end_index = len(self.paragraphs)  # 默认到文档末尾
        
        for i in range(start_index + 1, len(self.paragraphs)):
            para = self.paragraphs[i]
            text = para.text.strip()
            
            # 如果遇到下一个三级标题（新的任务），停止搜索
            if self._is_heading(para, level=3) and task_name not in text:
                end_index = i
                break
            
            # 如果遇到二级标题（新的组件），停止搜索
            if self._is_heading(para, level=2):
                end_index = i
                break
            
            # 如果遇到一级标题，停止搜索
            if self._is_heading(para, level=1):
                end_index = i
                break
        
        # 在确定的范围内搜索步骤
        for i in range(start_index, end_index):
            para = self.paragraphs[i]
            text = para.text.strip()
            
            # 检查是否是四级标题（步骤名称）
            if self._is_heading(para, level=4):
                match = re.match(r"(.+?)\*?（A阶段、B阶段）", text)
                if match:
                    step_name = match.group(1).strip()
                    if step_name not in exclude_keywords:
                        # 提取输入输出要素
                        input_elements, output_elements = self._extract_input_output_elements(i + 1)
                        step = StepInfo(
                            name=step_name,
                            input_elements=input_elements,
                            output_elements=output_elements
                        )
                        steps.append(step)
        
        return steps
    
    def _extract_input_output_elements(self, start_index: int) -> Tuple[List[InputElement], List[OutputElement]]:
        """提取输入输出要素"""
        input_elements = []
        output_elements = []
        
        # 查找"输入输出*（A阶段、B阶段）"标题（五级标题：#####级别）
        input_output_found = False
        input_output_index = -1
        
        for i in range(start_index, min(start_index + 50, len(self.paragraphs))):
            para = self.paragraphs[i]
            text = para.text.strip()
            
            if "输入输出" in text and "（A阶段、B阶段）" in text:
                if self._is_heading(para, level=5):
                    input_output_found = True
                    input_output_index = i
                    break
        
        if not input_output_found:
            return input_elements, output_elements
        
        # 查找输入输出标题后的表格
        # 查找输入输出标题后，下一个步骤或任务之前的范围
        end_index = min(input_output_index + 100, len(self.paragraphs))
        
        # 查找结束位置（下一个四级标题、三级标题、二级标题或一级标题）
        for i in range(input_output_index + 1, end_index):
            para = self.paragraphs[i]
            if (self._is_heading(para, level=4) or 
                self._is_heading(para, level=3) or 
                self._is_heading(para, level=2) or 
                self._is_heading(para, level=1)):
                end_index = i
                break
        
        # 在输入输出标题和结束位置之间查找"输入要素"和"输出要素"文本
        # 然后查找这些文本后第一个未使用的匹配表格
        
        found_input_text = False
        found_output_text = False
        
        for i in range(input_output_index + 1, end_index):
            para = self.paragraphs[i]
            text = para.text.strip()
            
            # 查找"输入要素"文本
            if ("输入要素" in text or ("输入" in text and "要素" in text)) and not found_input_text:
                found_input_text = True
                # 查找第一个未使用的输入要素表
                for table_idx, table in enumerate(self.tables):
                    if table_idx in self.used_tables:
                        continue
                    if len(table.rows) < 2:
                        continue
                    first_row_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
                    if "输入" in first_row_text and "字段名称" in first_row_text:
                        parsed = self._parse_input_table(table)
                        if parsed:  # 如果解析到数据，使用这个表格
                            input_elements = parsed
                            self.used_tables.add(table_idx)
                            break
            
            # 查找"输出要素"文本
            if ("输出要素" in text or ("输出" in text and "要素" in text)) and not found_output_text:
                found_output_text = True
                # 查找第一个未使用的输出要素表
                # 输出要素表的特征：包含"字段名称"和"类型"，但不包含"是否必输"（输入要素表才有）
                for table_idx, table in enumerate(self.tables):
                    if table_idx in self.used_tables:
                        continue
                    if len(table.rows) < 2:
                        continue
                    first_row_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
                    # 输出要素表：包含"字段名称"和"类型"，且不包含"是否必输"和"数据来源"
                    if ("字段名称" in first_row_text and "类型" in first_row_text and 
                        "是否必输" not in first_row_text and "数据来源" not in first_row_text):
                        parsed = self._parse_output_table(table)
                        if parsed:  # 如果解析到数据，使用这个表格
                            output_elements = parsed
                            self.used_tables.add(table_idx)
                            break
        
        return input_elements, output_elements
    
    def _parse_input_table(self, table: Table) -> List[InputElement]:
        """解析输入要素表"""
        elements = []
        
        if len(table.rows) < 2:
            return elements
        
        # 获取表头
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        
        # 查找各列索引
        index_idx = 0
        name_idx = self._find_column_index(headers, ["字段名称", "名称"])
        required_idx = self._find_column_index(headers, ["是否必输", "必输"])
        format_idx = self._find_column_index(headers, ["字段格式", "格式"])
        limit_idx = self._find_column_index(headers, ["输入限制", "限制"])
        desc_idx = self._find_column_index(headers, ["说明", "描述"])
        
        if name_idx == -1:
            return elements
        
        # 解析数据行
        for row in table.rows[1:]:
            if len(row.cells) < name_idx + 1:
                continue
            
            cells = [cell.text.strip() for cell in row.cells]
            
            # 跳过空行
            if name_idx < len(cells) and not cells[name_idx]:
                continue
            
            try:
                index = int(cells[index_idx]) if index_idx < len(cells) and cells[index_idx] else len(elements) + 1
            except:
                index = len(elements) + 1
            
            element = InputElement(
                index=index,
                field_name=cells[name_idx] if name_idx < len(cells) else "",
                required=cells[required_idx] if required_idx < len(cells) and required_idx != -1 else "否",
                field_format=cells[format_idx] if format_idx < len(cells) and format_idx != -1 else None,
                input_limit=cells[limit_idx] if limit_idx < len(cells) and limit_idx != -1 else None,
                description=cells[desc_idx] if desc_idx < len(cells) and desc_idx != -1 else None
            )
            elements.append(element)
        
        return elements
    
    def _parse_output_table(self, table: Table) -> List[OutputElement]:
        """解析输出要素表"""
        elements = []
        
        if len(table.rows) < 2:
            return elements
        
        # 获取表头
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        
        # 查找各列索引
        index_idx = 0
        name_idx = self._find_column_index(headers, ["字段名称", "名称"])
        type_idx = self._find_column_index(headers, ["类型", "字段类型"])
        desc_idx = self._find_column_index(headers, ["说明", "描述"])
        
        if name_idx == -1:
            return elements
        
        # 解析数据行
        for row in table.rows[1:]:
            if len(row.cells) < name_idx + 1:
                continue
            
            cells = [cell.text.strip() for cell in row.cells]
            
            # 跳过空行
            if name_idx < len(cells) and not cells[name_idx]:
                continue
            
            try:
                index = int(cells[index_idx]) if index_idx < len(cells) and cells[index_idx] else len(elements) + 1
            except:
                index = len(elements) + 1
            
            element = OutputElement(
                index=index,
                field_name=cells[name_idx] if name_idx < len(cells) else "",
                field_type=cells[type_idx] if type_idx < len(cells) and type_idx != -1 else None,
                description=cells[desc_idx] if desc_idx < len(cells) and desc_idx != -1 else None
            )
            elements.append(element)
        
        return elements
    
    def _is_heading(self, para: Paragraph, level: int) -> bool:
        """判断段落是否是指定级别的标题"""
        style_name = para.style.name
        # 检查样式名称（Heading 1, Heading 2等）
        if style_name.startswith('Heading'):
            try:
                heading_level = int(style_name.replace('Heading ', ''))
                return heading_level == level
            except:
                return False
        return False
    
    def _find_column_index(self, headers: List[str], keywords: List[str]) -> int:
        """查找包含关键词的列索引"""
        for i, header in enumerate(headers):
            for keyword in keywords:
                if keyword in header:
                    return i
        return -1
    
    # ========== 非建模需求解析方法 ==========
    
    def _extract_file_controlled_info(self) -> Tuple[Optional[str], Optional[str]]:
        """从文件受控信息表提取文件编号和文件名称"""
        file_number = None
        file_name = None
        
        for para in self.paragraphs[:100]:
            text = para.text.strip()
            if "文件受控信息" in text:
                # 查找后续的表格
                for table in self.tables:
                    if len(table.rows) < 1:
                        continue
                    
                    header_row = table.rows[0]
                    headers = [cell.text.strip() for cell in header_row.cells]
                    
                    # 查找文件编号和文件名称列
                    file_number_idx = self._find_column_index(headers, ["文件编号"])
                    file_name_idx = self._find_column_index(headers, ["文件名称"])
                    
                    if file_number_idx >= 0 or file_name_idx >= 0:
                        # 解析数据行（可能是横向布局：第一行是表头，第二行是值）
                        if len(table.rows) >= 2:
                            value_row = table.rows[1]
                            values = [cell.text.strip() for cell in value_row.cells]
                            
                            if file_number_idx >= 0 and file_number_idx < len(values):
                                file_number = values[file_number_idx]
                            
                            if file_name_idx >= 0 and file_name_idx < len(values):
                                file_name = values[file_name_idx]
                        
                        # 也尝试纵向布局：第一列是键，第二列是值
                        for row in table.rows[1:]:
                            if len(row.cells) >= 2:
                                key = row.cells[0].text.strip()
                                value = row.cells[1].text.strip()
                                
                                if value and value != '/':
                                    if "文件编号" in key and not file_number:
                                        file_number = value
                                    elif "文件名称" in key and not file_name:
                                        file_name = value
                        
                        if file_number or file_name:
                            return file_number, file_name
        
        return file_number, file_name
    
    def _extract_requirement_name(self, file_name: Optional[str]) -> Optional[str]:
        """提取需求名称
        方案一（主）：从文件名称中提取核心功能名
        方案二（备）：从功能清单第一项提取
        """
        # 方案一：从文件名称提取
        if file_name:
            # 正则模式：大信贷系统(.+?)业务需求说明书
            match = re.search(r"大信贷系统(.+?)业务需求说明书", file_name)
            if match:
                requirement_name = match.group(1).strip()
                # 清理处理：去除括号内容
                requirement_name = re.sub(r"（[^）]*）", "", requirement_name)
                requirement_name = re.sub(r"\([^)]*\)", "", requirement_name)
                requirement_name = requirement_name.strip()
                if requirement_name:
                    return requirement_name
        
        # 方案二：从功能清单第一项提取
        functions = self._extract_function_list()
        if functions:
            return functions[0]
        
        return None
    
    def _extract_designer(self) -> Optional[str]:
        """从文件信息表提取设计者（作者）"""
        for para in self.paragraphs[:100]:
            text = para.text.strip()
            if "文件信息" in text:
                # 查找后续的表格
                for table in self.tables:
                    if len(table.rows) < 1:
                        continue
                    
                    header_row = table.rows[0]
                    headers = [cell.text.strip() for cell in header_row.cells]
                    
                    # 查找作者列
                    author_idx = self._find_column_index(headers, ["作者"])
                    
                    if author_idx >= 0:
                        # 解析数据行
                        if len(table.rows) >= 2:
                            value_row = table.rows[1]
                            values = [cell.text.strip() for cell in value_row.cells]
                            
                            if author_idx < len(values):
                                designer = values[author_idx]
                                if designer and designer != '/':
                                    return designer
                        
                        # 也尝试纵向布局
                        for row in table.rows[1:]:
                            if len(row.cells) >= 2:
                                key = row.cells[0].text.strip()
                                value = row.cells[1].text.strip()
                                
                                if "作者" in key and value and value != '/':
                                    return value
        
        return None
    
    def _extract_function_list(self) -> List[str]:
        """提取功能清单（仅功能名称列表）"""
        functions = []
        
        # 查找"5.1 功能清单"章节
        for i, para in enumerate(self.paragraphs):
            text = para.text.strip()
            
            # 查找"5 功能*（A阶段）"或"5.1 功能清单"
            if ("功能" in text and "（A阶段）" in text) or "功能清单" in text:
                # 查找后续的表格
                for table in self.tables:
                    if len(table.rows) < 2:
                        continue
                    
                    header_row = table.rows[0]
                    headers = [cell.text.strip() for cell in header_row.cells]
                    
                    # 查找"业务功能名称"列
                    function_name_idx = self._find_column_index(headers, ["业务功能名称", "功能名称"])
                    
                    if function_name_idx >= 0:
                        # 解析数据行
                        for row in table.rows[1:]:
                            if len(row.cells) > function_name_idx:
                                function_name = row.cells[function_name_idx].text.strip()
                                if function_name and function_name not in functions:
                                    functions.append(function_name)
                        
                        if functions:
                            return functions
        
        return functions
    
    def _extract_functions(self) -> List[FunctionInfo]:
        """提取功能列表（包含输入输出要素）"""
        functions = []
        
        # 1. 先提取功能名称列表
        function_names = self._extract_function_list()
        
        if not function_names:
            return functions
        
        # 2. 为每个功能提取详细输入输出要素
        for function_name in function_names:
            input_elements, output_elements = self._extract_function_input_output(function_name)
            
            function = FunctionInfo(
                name=function_name,
                input_elements=input_elements,
                output_elements=output_elements
            )
            functions.append(function)
        
        return functions
    
    def _extract_function_input_output(self, function_name: str) -> Tuple[List[InputElement], List[OutputElement]]:
        """提取指定功能的输入输出要素"""
        input_elements = []
        output_elements = []
        
        # 查找"5.2 功能说明"下的对应功能章节
        function_section_index = -1
        
        for i, para in enumerate(self.paragraphs):
            text = para.text.strip()
            
            # 查找"5.2 功能说明"或"功能说明"
            if "功能说明" in text and ("5.2" in text or "（A阶段）" in text):
                # 在该章节下查找功能名称（三级标题）
                for j in range(i + 1, min(i + 200, len(self.paragraphs))):
                    next_para = self.paragraphs[j]
                    next_text = next_para.text.strip()
                    
                    # 如果遇到下一个一级或二级标题，停止搜索
                    if self._is_heading(next_para, level=1) or self._is_heading(next_para, level=2):
                        if "功能说明" not in next_text:
                            break
                    
                    # 检查是否是三级标题且匹配功能名称
                    if self._is_heading(next_para, level=3):
                        # 模糊匹配功能名称（允许标点符号差异）
                        cleaned_title = re.sub(r"[^\w\u4e00-\u9fa5]", "", next_text)
                        cleaned_function = re.sub(r"[^\w\u4e00-\u9fa5]", "", function_name)
                        
                        if cleaned_function in cleaned_title or cleaned_title in cleaned_function:
                            function_section_index = j
                            break
                
                if function_section_index >= 0:
                    break
        
        if function_section_index < 0:
            return input_elements, output_elements
        
        # 在功能章节内查找输入输出要素
        # 查找结束位置（下一个三级标题、二级标题或一级标题）
        end_index = len(self.paragraphs)
        for i in range(function_section_index + 1, len(self.paragraphs)):
            para = self.paragraphs[i]
            if (self._is_heading(para, level=3) or 
                self._is_heading(para, level=2) or 
                self._is_heading(para, level=1)):
                end_index = i
                break
        
        # 在功能章节内查找"输入要素"和"输出要素"
        found_input_text = False
        found_output_text = False
        
        for i in range(function_section_index + 1, end_index):
            para = self.paragraphs[i]
            text = para.text.strip()
            
            # 查找"输入要素"文本
            if ("输入要素" in text or ("输入" in text and "要素" in text)) and not found_input_text:
                found_input_text = True
                # 查找第一个未使用的输入要素表
                for table_idx, table in enumerate(self.tables):
                    if table_idx in self.used_tables:
                        continue
                    if len(table.rows) < 2:
                        continue
                    first_row_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
                    if "输入" in first_row_text and "字段名称" in first_row_text:
                        parsed = self._parse_input_table(table)
                        if parsed:
                            input_elements = parsed
                            self.used_tables.add(table_idx)
                            break
            
            # 查找"输出要素"文本
            if ("输出要素" in text or ("输出" in text and "要素" in text)) and not found_output_text:
                found_output_text = True
                # 查找第一个未使用的输出要素表
                for table_idx, table in enumerate(self.tables):
                    if table_idx in self.used_tables:
                        continue
                    if len(table.rows) < 2:
                        continue
                    first_row_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
                    if ("字段名称" in first_row_text and "类型" in first_row_text and 
                        "是否必输" not in first_row_text and "数据来源" not in first_row_text):
                        parsed = self._parse_output_table(table)
                        if parsed:
                            output_elements = parsed
                            self.used_tables.add(table_idx)
                            break
        
        return input_elements, output_elements
