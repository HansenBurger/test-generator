"""
Word文档解析服务 - 银行需求文档专用解析器
"""
import re
import os
import tempfile
from pathlib import Path
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
        self._temp_docx_path = None  # 用于存储临时转换的 .docx 文件路径
        actual_doc_path = self._handle_doc_file(doc_path)
        
        try:
            self.doc = Document(actual_doc_path)
            self.paragraphs = [p for p in self.doc.paragraphs]
            self.tables = self.doc.tables
            self.used_tables = set()  # 记录已使用的表格索引，避免重复使用
        except Exception as e:
            # 清理临时文件
            self._cleanup_temp_file()
            raise
    
    def _handle_doc_file(self, doc_path: str) -> str:
        """处理 .doc 文件，如果是 .doc 格式则转换为 .docx"""
        doc_path_obj = Path(doc_path)
        
        # 如果已经是 .docx 格式，直接返回
        if doc_path_obj.suffix.lower() == '.docx':
            return doc_path
        
        # 如果是 .doc 格式，需要转换
        if doc_path_obj.suffix.lower() == '.doc':
            return self._convert_doc_to_docx(doc_path)
        
        # 其他格式，尝试直接打开（可能会失败）
        return doc_path
    
    def _convert_doc_to_docx(self, doc_path: str) -> str:
        """将 .doc 文件转换为 .docx 格式（使用 Windows COM 接口）"""
        try:
            import win32com.client
        except ImportError:
            raise ValueError(
                "无法处理 .doc 格式文件：需要安装 pywin32 库。"
                "请运行: pip install pywin32"
            )
        
        try:
            # 创建临时 .docx 文件
            temp_dir = tempfile.gettempdir()
            # 使用安全的文件名（移除特殊字符，避免路径问题）
            safe_filename = re.sub(r'[<>:"/\\|?*]', '_', os.path.basename(doc_path))
            temp_docx_path = os.path.join(
                temp_dir,
                f"converted_{safe_filename}.docx"
            )
            
            # 使用 Word COM 接口转换
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
            
            try:
                # 打开 .doc 文件
                doc = word_app.Documents.Open(os.path.abspath(doc_path))
                
                # 保存为 .docx 格式
                doc.SaveAs2(
                    FileName=os.path.abspath(temp_docx_path),
                    FileFormat=16  # wdFormatXMLDocument = 16 (.docx)
                )
                
                doc.Close()
                word_app.Quit()
                
                # 保存临时文件路径，用于后续清理
                self._temp_docx_path = temp_docx_path
                
                return temp_docx_path
                
            except Exception as e:
                try:
                    word_app.Quit()
                except:
                    pass
                raise ValueError(
                    f"无法将 .doc 文件转换为 .docx 格式：{str(e)}。"
                    "请确保已安装 Microsoft Word，或手动将文件转换为 .docx 格式。"
                )
        except Exception as e:
            if isinstance(e, ValueError):
                raise
            raise ValueError(
                f"无法处理 .doc 格式文件：{str(e)}。"
                "请确保已安装 Microsoft Word，或手动将文件转换为 .docx 格式。"
            )
    
    def _cleanup_temp_file(self):
        """清理临时转换的 .docx 文件"""
        if self._temp_docx_path and os.path.exists(self._temp_docx_path):
            try:
                os.unlink(self._temp_docx_path)
            except:
                pass
            self._temp_docx_path = None
    
    def __del__(self):
        """析构函数，清理临时文件"""
        self._cleanup_temp_file()
    
    def parse(self) -> ParsedDocument:
        """解析文档主方法"""
        # 1. 识别文档类型
        doc_type = self._identify_document_type()
        
        if doc_type == "modeling":
            return self._parse_modeling_document()
        elif doc_type == "non_modeling":
            return self._parse_non_modeling_document()
        else:
            raise ValueError("无法识别文档类型：未找到'用例版本控制信息'表或'文件受控信息'/'文档受控信息'表")
    
    def _identify_document_type(self) -> Optional[str]:
        """识别文档类型：建模需求或非建模需求"""
        # 优先级1：查找"用例版本控制信息"（建模需求的明确标识）
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
        
        # 优先级2：查找"文件受控信息"或"文档受控信息"（非建模需求的明确标识）
        # 同时检查是否有"功能清单"（非建模需求的另一个特征）
        has_file_control = False
        has_function_list = False
        
        for para in self.paragraphs[:100]:
            text = para.text.strip()
            if "文件受控信息" in text or "文档受控信息" in text:
                has_file_control = True
            if "功能清单" in text:
                has_function_list = True
        
        # 检查表格
        for table in self.tables:
            if len(table.rows) < 1:
                continue
            header_row = table.rows[0]
            header_text = ' '.join([cell.text.strip() for cell in header_row.cells])
            
            if ("文件编号" in header_text or "文件名称" in header_text or 
                "文档受控信息" in header_text):
                has_file_control = True
            
            if "业务功能名称" in header_text or "功能名称" in header_text:
                has_function_list = True
        
        # 如果有文件受控信息或功能清单，识别为非建模需求
        if has_file_control or has_function_list:
            return "non_modeling"
        
        # 优先级3：查找"版本控制信息"（不带"用例"前缀，可能是建模需求）
        # 但需要更严格的判断：必须同时有"需求用例概述"
        has_version_control = False
        has_requirement_overview = False
        
        for para in self.paragraphs[:100]:
            text = para.text.strip()
            if "版本控制信息" in text and "用例" not in text:
                has_version_control = True
            if "需求用例概述" in text:
                has_requirement_overview = True
        
        # 检查是否有包含"版本"字段的表格
        for table in self.tables:
            if len(table.rows) < 1:
                continue
            header_row = table.rows[0]
            header_text = ' '.join([cell.text.strip() for cell in header_row.cells])
            if "版本" in header_text and has_version_control:
                # 如果同时有"需求用例概述"，才识别为建模需求
                if has_requirement_overview:
                    return "modeling"
        
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
        """解析输入要素表（增强版）"""
        elements = []
        
        if len(table.rows) < 2:
            return elements
        
        # 获取表头
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        
        # 使用模糊匹配查找各列索引
        index_idx = 0
        name_idx = self._fuzzy_find_column_index(headers, ["字段名称", "名称", "字段名"])
        required_idx = self._fuzzy_find_column_index(headers, ["是否必输", "是否必填", "必输", "必填"])
        type_idx = self._fuzzy_find_column_index(headers, ["类型", "字段类型"])
        precision_idx = self._fuzzy_find_column_index(headers, ["精度"])
        format_idx = self._fuzzy_find_column_index(headers, ["字段格式", "格式", "输入格式"])
        limit_idx = self._fuzzy_find_column_index(headers, ["输入限制", "限制", "数据字典"])
        desc_idx = self._fuzzy_find_column_index(headers, ["说明", "描述", "备注"])
        
        if name_idx == -1:
            return elements
        
        # 解析数据行
        for i, row in enumerate(table.rows[1:]):
            if len(row.cells) < name_idx + 1:
                continue
            
            cells = [cell.text.strip() for cell in row.cells]
            
            # 跳过空行
            if name_idx < len(cells) and not cells[name_idx]:
                continue
            
            try:
                index = int(cells[index_idx]) if index_idx < len(cells) and cells[index_idx] else i + 1
            except:
                index = i + 1
            
            element = InputElement(
                index=index,
                field_name=cells[name_idx] if name_idx < len(cells) else "",
                required=cells[required_idx] if required_idx != -1 and required_idx < len(cells) and cells[required_idx] else "否",
                field_type=cells[type_idx] if type_idx != -1 and type_idx < len(cells) and cells[type_idx] else None,
                precision=cells[precision_idx] if precision_idx != -1 and precision_idx < len(cells) and cells[precision_idx] else None,
                field_format=cells[format_idx] if format_idx != -1 and format_idx < len(cells) and cells[format_idx] else None,
                input_limit=cells[limit_idx] if limit_idx != -1 and limit_idx < len(cells) and cells[limit_idx] else None,
                description=cells[desc_idx] if desc_idx != -1 and desc_idx < len(cells) and cells[desc_idx] else None
            )
            elements.append(element)
        
        return elements
    
    def _parse_output_table(self, table: Table) -> List[OutputElement]:
        """解析输出要素表（增强版）"""
        elements = []
        
        if len(table.rows) < 2:
            return elements
        
        # 获取表头
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        
        # 使用模糊匹配查找各列索引
        index_idx = 0
        name_idx = self._fuzzy_find_column_index(headers, ["字段名称", "名称", "字段名"])
        type_idx = self._fuzzy_find_column_index(headers, ["类型", "字段类型"])
        precision_idx = self._fuzzy_find_column_index(headers, ["精度"])
        format_idx = self._fuzzy_find_column_index(headers, ["字段格式", "格式"])
        desc_idx = self._fuzzy_find_column_index(headers, ["说明", "描述", "备注"])
        
        if name_idx == -1:
            return elements
        
        # 解析数据行
        for i, row in enumerate(table.rows[1:]):
            if len(row.cells) < name_idx + 1:
                continue
            
            cells = [cell.text.strip() for cell in row.cells]
            
            # 跳过空行
            if name_idx < len(cells) and not cells[name_idx]:
                continue
            
            try:
                index = int(cells[index_idx]) if index_idx < len(cells) and cells[index_idx] else i + 1
            except:
                index = i + 1
            
            element = OutputElement(
                index=index,
                field_name=cells[name_idx] if name_idx < len(cells) else "",
                field_type=cells[type_idx] if type_idx != -1 and type_idx < len(cells) and cells[type_idx] else None,
                precision=cells[precision_idx] if precision_idx != -1 and precision_idx < len(cells) and cells[precision_idx] else None,
                field_format=cells[format_idx] if format_idx != -1 and format_idx < len(cells) and cells[format_idx] else None,
                description=cells[desc_idx] if desc_idx != -1 and desc_idx < len(cells) and cells[desc_idx] else None
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
    
    def _fuzzy_find_column_index(self, headers: List[str], keywords: List[str]) -> int:
        """模糊查找列索引
        - 先尝试完全匹配
        - 再尝试部分包含
        - 最后尝试相似匹配
        """
        # 1. 完全匹配
        for i, header in enumerate(headers):
            for keyword in keywords:
                if keyword == header or keyword in header:
                    return i
        
        # 2. 部分包含（去除空格和标点）
        for i, header in enumerate(headers):
            cleaned_header = re.sub(r'[^\w\u4e00-\u9fa5]', '', header)
            for keyword in keywords:
                cleaned_keyword = re.sub(r'[^\w\u4e00-\u9fa5]', '', keyword)
                if cleaned_keyword in cleaned_header or cleaned_header in cleaned_keyword:
                    return i
        
        return -1
    
    def _is_input_table(self, header_text: str) -> bool:
        """判断是否为输入要素表"""
        # 必须包含字段名称
        if "字段名称" not in header_text:
            return False
        
        # 优先判断：如果包含"是否必输"或"数据来源"，很可能是输入表
        if "是否必输" in header_text or "数据来源" in header_text:
            return True
        
        # 备选判断：包含"输入限制"且不包含"输出"
        if "输入限制" in header_text and "输出" not in header_text:
            return True
        
        # 排除输出表：如果包含"输出限制"或"输出"关键词，不是输入表
        if "输出限制" in header_text or ("输出" in header_text and "类型" in header_text):
            return False
        
        return False
    
    def _is_output_table(self, header_text: str) -> bool:
        """判断是否为输出要素表"""
        # 必须包含字段名称
        if "字段名称" not in header_text:
            return False
        
        # 排除输入表特征：如果包含"是否必输"和"数据来源"，不是输出表
        if "是否必输" in header_text and "数据来源" in header_text:
            return False
        
        # 如果包含"是否必输"或"数据来源"，不是输出表
        if "是否必输" in header_text or "数据来源" in header_text:
            return False
        
        # 包含"类型"且包含"输出限制"或"输出"，是输出表
        if "类型" in header_text and ("输出限制" in header_text or "输出" in header_text):
            return True
        
        # 包含"类型"且不包含"是否必输"和"数据来源"，可能是输出表
        # 注意：即使有"输入限制"列，只要没有"是否必输"和"数据来源"，也可能是输出表
        if "类型" in header_text:
            return True
        
        return False
    
    def _search_tables_near_marker(self, marker_index: int, is_input: bool, max_distance: int = 20) -> List:
        """在标记附近搜索表格"""
        elements = []
        start_idx = max(0, marker_index - max_distance)
        end_idx = min(len(self.paragraphs), marker_index + max_distance)
        
        # 在范围内查找表格（通过检查段落和表格的关联）
        # 由于python-docx无法直接关联段落和表格，我们采用顺序查找策略
        # 找到标记后，按顺序查找后续的未使用表格
        for table_idx, table in enumerate(self.tables):
            if table_idx in self.used_tables:
                continue
            if len(table.rows) < 2:
                continue
            
            header_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
            
            if is_input and self._is_input_table(header_text):
                parsed = self._parse_input_table(table)
                if parsed:
                    elements = parsed
                    self.used_tables.add(table_idx)
                    break
            elif not is_input and self._is_output_table(header_text):
                parsed = self._parse_output_table(table)
                if parsed:
                    elements = parsed
                    self.used_tables.add(table_idx)
                    break
        
        return elements
    
    def _search_tables_in_range(self, start: int, end: int, is_input: bool, allow_used: bool = False) -> List:
        """在指定范围内搜索表格"""
        elements = []
        
        # 按顺序查找表格
        for table_idx, table in enumerate(self.tables):
            if not allow_used and table_idx in self.used_tables:
                continue
            if len(table.rows) < 2:
                continue
            
            header_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
            
            if is_input and self._is_input_table(header_text):
                parsed = self._parse_input_table(table)
                if parsed:
                    elements = parsed
                    if not allow_used:
                        self.used_tables.add(table_idx)
                    break
            elif not is_input and self._is_output_table(header_text):
                parsed = self._parse_output_table(table)
                if parsed:
                    elements = parsed
                    if not allow_used:
                        self.used_tables.add(table_idx)
                    break
        
        return elements
    
    def _search_all_unused_tables(self, is_input: bool) -> List:
        """遍历所有未使用的表格"""
        elements = []
        
        for table_idx, table in enumerate(self.tables):
            if table_idx in self.used_tables:
                continue
            if len(table.rows) < 2:
                continue
            
            header_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
            
            if is_input and self._is_input_table(header_text):
                parsed = self._parse_input_table(table)
                if parsed:
                    elements = parsed
                    self.used_tables.add(table_idx)
                    break
            elif not is_input and self._is_output_table(header_text):
                parsed = self._parse_output_table(table)
                if parsed:
                    elements = parsed
                    self.used_tables.add(table_idx)
                    break
        
        return elements
    
    def _find_nearest_table_after_marker(self, marker_index: int, is_input: bool) -> List:
        """查找标记后最近的表格（即使已使用）
        
        这对于后面的功能很重要，因为它们的表格可能在已使用的表格之后
        我们需要找到标记后最近的一个匹配表格
        """
        elements = []
        
        # 由于python-docx无法直接关联段落和表格，我们采用策略：
        # 1. 先查找未使用的表格
        # 2. 如果没找到，查找所有表格中第一个匹配的（按表格索引顺序）
        
        # 先尝试未使用的表格
        for table_idx, table in enumerate(self.tables):
            if table_idx in self.used_tables:
                continue
            if len(table.rows) < 2:
                continue
            
            header_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
            
            if is_input and self._is_input_table(header_text):
                parsed = self._parse_input_table(table)
                if parsed:
                    elements = parsed
                    self.used_tables.add(table_idx)
                    return elements
        
        # 如果没找到未使用的，查找所有表格（包括已使用的）
        # 这对于"优惠利息查询"等后面功能很重要
        for table_idx, table in enumerate(self.tables):
            if len(table.rows) < 2:
                continue
            
            header_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
            
            if is_input and self._is_input_table(header_text):
                parsed = self._parse_input_table(table)
                if parsed:
                    # 检查表格内容是否可能属于当前功能
                    # 简单策略：如果表格有数据，就使用它
                    elements = parsed
                    if table_idx not in self.used_tables:
                        self.used_tables.add(table_idx)
                    return elements
            elif not is_input and self._is_output_table(header_text):
                parsed = self._parse_output_table(table)
                if parsed:
                    elements = parsed
                    if table_idx not in self.used_tables:
                        self.used_tables.add(table_idx)
                    return elements
        
        return elements
    
    # ========== 非建模需求解析方法 ==========
    
    def _extract_file_controlled_info(self) -> Tuple[Optional[str], Optional[str]]:
        """从文件受控信息表或文档受控信息表提取文件编号和文件名称"""
        file_number = None
        file_name = None
        
        # 查找包含"文件受控信息"或"文档受控信息"的段落或表格
        for para in self.paragraphs[:100]:
            text = para.text.strip()
            if "文件受控信息" in text or "文档受控信息" in text:
                # 查找后续的表格
                for table in self.tables:
                    if len(table.rows) < 2:
                        continue
                    
                    header_row = table.rows[0]
                    headers = [cell.text.strip() for cell in header_row.cells]
                    header_text = ' '.join(headers)
                    
                    # 检查表格是否包含"文档受控信息"（可能是表头就是"文档受控信息"）
                    if "文档受控信息" in header_text:
                        # 特殊格式：行1可能是 ['文件编号', '值', '文件名称', '值']
                        if len(table.rows) >= 2:
                            data_row = table.rows[1]
                            cells = [cell.text.strip() for cell in data_row.cells]
                            
                            # 查找"文件编号"和"文件名称"的位置
                            for i, cell_text in enumerate(cells):
                                if "文件编号" in cell_text and i + 1 < len(cells):
                                    # 下一个单元格是文件编号的值
                                    file_number = cells[i + 1] if cells[i + 1] and cells[i + 1] != '/' else None
                                elif "文件名称" in cell_text and i + 1 < len(cells):
                                    # 下一个单元格是文件名称的值
                                    file_name = cells[i + 1] if cells[i + 1] and cells[i + 1] != '/' else None
                            
                            if file_number or file_name:
                                return file_number, file_name
                        
                        # 也尝试纵向布局：第一列是键，第二列是值
                        for row in table.rows[1:]:
                            if len(row.cells) >= 2:
                                key = row.cells[0].text.strip()
                                value = row.cells[1].text.strip()
                                
                                if value and value != '/':
                                    if ("文件编号" in key or ("编号" in key and "文件" in key)) and not file_number:
                                        file_number = value
                                    elif ("文件名称" in key or ("名称" in key and "文件" in key)) and not file_name:
                                        file_name = value
                        
                        if file_number or file_name:
                            return file_number, file_name
                    
                    # 查找文件编号和文件名称列（标准表格格式）
                    file_number_idx = self._find_column_index(headers, ["文件编号"])
                    file_name_idx = self._find_column_index(headers, ["文件名称"])
                    
                    if file_number_idx >= 0 or file_name_idx >= 0:
                        # 解析数据行（可能是横向布局：第一行是表头，第二行是值）
                        if len(table.rows) >= 2:
                            value_row = table.rows[1]
                            values = [cell.text.strip() for cell in value_row.cells]
                            
                            if file_number_idx >= 0 and file_number_idx < len(values):
                                file_number = values[file_number_idx] if values[file_number_idx] and values[file_number_idx] != '/' else None
                            
                            if file_name_idx >= 0 and file_name_idx < len(values):
                                file_name = values[file_name_idx] if values[file_name_idx] and values[file_name_idx] != '/' else None
                        
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
        
        # 也尝试直接从表格中查找（不依赖段落文本）
        for table in self.tables:
            if len(table.rows) < 2:
                continue
            
            header_row = table.rows[0]
            headers = [cell.text.strip() for cell in header_row.cells]
            header_text = ' '.join(headers)
            
            # 检查是否是文档受控信息表
            if "文档受控信息" in header_text:
                # 特殊格式处理
                if len(table.rows) >= 2:
                    data_row = table.rows[1]
                    cells = [cell.text.strip() for cell in data_row.cells]
                    
                    for i, cell_text in enumerate(cells):
                        if "文件编号" in cell_text and i + 1 < len(cells):
                            file_number = cells[i + 1] if cells[i + 1] and cells[i + 1] != '/' else None
                        elif "文件名称" in cell_text and i + 1 < len(cells):
                            file_name = cells[i + 1] if cells[i + 1] and cells[i + 1] != '/' else None
                
                # 也尝试纵向布局
                for row in table.rows[1:]:
                    if len(row.cells) >= 2:
                        key = row.cells[0].text.strip()
                        value = row.cells[1].text.strip()
                        
                        if value and value != '/':
                            if ("文件编号" in key or ("编号" in key and "文件" in key)) and not file_number:
                                file_number = value
                            elif ("文件名称" in key or ("名称" in key and "文件" in key)) and not file_name:
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
            # 清理换行符和多余空格
            file_name = file_name.replace('\n', '').replace('\r', '')
            file_name = re.sub(r'\s+', '', file_name)
            
            # 尝试多种正则模式匹配
            patterns = [
                r"大信贷系统(.+?)业务需求说明书",  # 标准格式
                r"大信贷系统详细业务-(.+?)需求说明书",  # 详细业务格式
                r"大信贷系统(.+?)需求说明书",  # 简化格式
            ]
            
            for pattern in patterns:
                match = re.search(pattern, file_name)
                if match:
                    requirement_name = match.group(1).strip()
                    # 清理处理：去除括号内容（但保留功能名称中的括号，如"贷款当日冲正（前台）"）
                    # 只去除说明性的括号，如"（优化对客服务类）"
                    # 如果需求名称本身包含括号，应该保留
                    if requirement_name:
                        return requirement_name
        
        # 方案二：从功能清单第一项提取
        functions = self._extract_function_list()
        if functions:
            return functions[0]
        
        return None
    
    def _extract_designer(self) -> Optional[str]:
        """提取设计者（作者）
        
        注意：设计者字段用于填写测试人员，不需要从文档中提取，直接返回None
        """
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
        
        # 2. 性能优化：一次性查找"功能说明"部分起始位置
        function_section_start = -1
        for i, para in enumerate(self.paragraphs):
            text = para.text.strip()
            if "功能说明" in text and ("5.2" in text or "（A阶段）" in text):
                function_section_start = i
                break
        
        # 3. 性能优化：一次性建立功能名称到段落索引的映射
        function_name_to_index = {}
        search_start = max(function_section_start if function_section_start >= 0 else 0, 100)
        search_end = len(self.paragraphs)
        
        # 为每个功能名称建立索引映射（使用与原方法相同的逻辑）
        for function_name in function_names:
            cleaned_function = re.sub(r"[^\w\u4e00-\u9fa5]", "", function_name)
            found = False
            
            # 优先查找：在功能说明部分内精确匹配功能名称的段落
            if function_section_start >= 0:
                # 在功能说明部分内查找（跳过目录部分，通常目录在前100个段落）
                for i in range(max(function_section_start, 100), search_end):
                    para = self.paragraphs[i]
                    text = para.text.strip()
                    
                    # 精确匹配：段落文本就是功能名称（可能带编号）
                    if function_name == text or (function_name in text and len(text) <= len(function_name) + 10):
                        # 排除目录和编号行
                        if "目录" not in text and not re.match(r'^\d+\.\d+', text):
                            function_name_to_index[function_name] = i
                            found = True
                            break
            
            # 如果没找到精确匹配，使用模糊匹配（优先在功能说明部分内）
            if not found:
                if function_section_start >= 0:
                    # 先在功能说明部分内查找
                    for i in range(max(function_section_start, 100), search_end):
                        para = self.paragraphs[i]
                        text = para.text.strip()
                        
                        # 检查是否匹配功能名称（可能是标题或普通段落）
                        cleaned_text = re.sub(r"[^\w\u4e00-\u9fa5]", "", text)
                        
                        # 匹配逻辑：功能名称完全匹配，或者功能名称包含在文本中
                        if (cleaned_function in cleaned_text or cleaned_text in cleaned_function) and len(cleaned_text) >= len(cleaned_function) * 0.7:
                            # 确保不是在目录或其他不相关的地方
                            if ("功能" in text or function_name in text) and "目录" not in text:
                                # 排除目录行（通常包含页码）
                                if not re.match(r'^\d+\.\d+', text) or len(text) > 50:
                                    function_name_to_index[function_name] = i
                                    found = True
                                    break
                
                # 如果还没找到，在整个文档中查找
                if not found:
                    for i in range(search_start, search_end):
                        para = self.paragraphs[i]
                        text = para.text.strip()
                        
                        cleaned_text = re.sub(r"[^\w\u4e00-\u9fa5]", "", text)
                        
                        if (cleaned_function in cleaned_text or cleaned_text in cleaned_function) and len(cleaned_text) >= len(cleaned_function) * 0.7:
                            if ("功能" in text or function_name in text) and "目录" not in text:
                                # 排除目录行
                                if not re.match(r'^\d+\.\d+', text) or len(text) > 50:
                                    function_name_to_index[function_name] = i
                                    break
        
        # 4. 为每个功能提取详细输入输出要素（使用缓存的索引，失败时回退到原方法）
        for function_name in function_names:
            function_index = function_name_to_index.get(function_name, -1)
            
            # 如果优化方法找到了索引，使用优化版本；否则回退到原方法
            if function_index >= 0:
                input_elements, output_elements = self._extract_function_input_output_optimized(
                    function_name, function_index, function_section_start
                )
            else:
                # 回退到原来的方法，确保兼容性
                input_elements, output_elements = self._extract_function_input_output(function_name)
            
            function = FunctionInfo(
                name=function_name,
                input_elements=input_elements,
                output_elements=output_elements
            )
            functions.append(function)
        
        return functions
    
    def _extract_function_input_output_optimized(self, function_name: str, function_section_index: int, function_section_start: int) -> Tuple[List[InputElement], List[OutputElement]]:
        """提取指定功能的输入输出要素（优化版本，使用预计算的索引）"""
        input_elements = []
        output_elements = []
        
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
        
        # 扩大搜索范围：从功能章节开始，向后搜索
        search_start = max(0, function_section_index - 10)
        search_end = min(len(self.paragraphs), function_section_index + 200)
        
        # 在功能章节内查找"输入要素"和"输出要素"标记
        input_markers = ["输入要素", "输入要素：", "输入输出要素", "输入要素表"]
        output_markers = ["输出要素", "输出要素：", "输出要素表"]
        
        found_input_marker = False
        found_output_marker = False
        input_marker_index = -1
        output_marker_index = -1
        
        # 先找到"输入要素"和"输出要素"文本的位置，并检查是否"不涉及"
        input_not_involved = False
        output_not_involved = False
        found_input_output_section = False
        
        # 先查找"输入输出说明"标记，确保我们在正确的章节内
        for i in range(search_start, search_end):
            para = self.paragraphs[i]
            text = para.text.strip()
            
            # 查找"输入输出说明"或"输入输出要素"
            if "输入输出说明" in text or ("输入输出要素" in text and "：" in text):
                found_input_output_section = True
                search_start_io = i
                break
        
        # 如果找到了"输入输出说明"章节，在该章节内查找
        if found_input_output_section:
            for i in range(search_start_io, min(search_start_io + 20, search_end)):
                para = self.paragraphs[i]
                text = para.text.strip()
                
                # 查找"输入要素"标记
                if not found_input_marker:
                    for marker in input_markers:
                        if marker in text and ("输入要素" in marker or marker == "输入要素："):
                            found_input_marker = True
                            input_marker_index = i
                            # 检查后续段落（最多5行）是否包含"不涉及"
                            for j in range(i + 1, min(i + 6, search_end)):
                                next_text = self.paragraphs[j].text.strip()
                                if not next_text:
                                    continue
                                if any(m in next_text for m in output_markers):
                                    break
                                if re.match(r'^[一二三四五六七八九十]+、', next_text):
                                    break
                                if "不涉及" in next_text:
                                    input_not_involved = True
                                    break
                            break
                
                # 查找"输出要素"标记
                if not found_output_marker:
                    for marker in output_markers:
                        if marker in text and ("输出要素" in marker or marker == "输出要素："):
                            found_output_marker = True
                            output_marker_index = i
                            # 检查后续段落（最多5行）是否包含"不涉及"
                            for j in range(i + 1, min(i + 6, search_end)):
                                next_text = self.paragraphs[j].text.strip()
                                if not next_text:
                                    continue
                                if re.match(r'^[一二三四五六七八九十]+、', next_text):
                                    break
                                if "不涉及" in next_text:
                                    output_not_involved = True
                                    break
                            break
                
                if found_input_marker and found_output_marker:
                    break
        else:
            # 如果没有找到"输入输出说明"，使用原来的逻辑
            for i in range(search_start, search_end):
                para = self.paragraphs[i]
                text = para.text.strip()
                
                # 查找"输入要素"标记
                if not found_input_marker:
                    for marker in input_markers:
                        if marker in text:
                            found_input_marker = True
                            input_marker_index = i
                            # 检查后续段落是否包含"不涉及"
                            for j in range(i + 1, min(i + 4, search_end)):
                                next_text = self.paragraphs[j].text.strip()
                                if any(m in next_text for m in output_markers):
                                    break
                                if "不涉及" in next_text:
                                    input_not_involved = True
                                    break
                            break
                
                # 查找"输出要素"标记
                if not found_output_marker:
                    for marker in output_markers:
                        if marker in text:
                            found_output_marker = True
                            output_marker_index = i
                            # 检查后续段落是否包含"不涉及"
                            for j in range(i + 1, min(i + 4, search_end)):
                                next_text = self.paragraphs[j].text.strip()
                                if re.match(r'^[一二三四五六七八九十]+、', next_text):
                                    break
                                if "不涉及" in next_text:
                                    output_not_involved = True
                                    break
                            break
                
                if found_input_marker and found_output_marker:
                    break
        
        # 智能表格定位：按顺序查找未使用的表格
        if found_input_marker and not input_not_involved:
            input_elements = self._search_all_unused_tables(is_input=True)
            if not input_elements and input_marker_index >= 0:
                input_elements = self._find_nearest_table_after_marker(
                    input_marker_index, is_input=True
                )
        elif input_not_involved:
            input_elements = []
        
        if found_output_marker and not output_not_involved:
            # 优先查找标记后最近的输出要素表
            if output_marker_index >= 0:
                output_elements = self._find_nearest_table_after_marker(
                    output_marker_index, is_input=False
                )
            # 如果没找到，再查找所有未使用的表格
            if not output_elements:
                output_elements = self._search_all_unused_tables(is_input=False)
        elif output_not_involved:
            output_elements = []
        
        return input_elements, output_elements
    
    def _extract_function_input_output(self, function_name: str) -> Tuple[List[InputElement], List[OutputElement]]:
        """提取指定功能的输入输出要素"""
        input_elements = []
        output_elements = []
        
        # 查找功能名称所在位置（可能在标题中，也可能在普通段落中）
        # 优先查找功能说明部分（5.2）下的功能名称
        function_section_index = -1
        
        # 清理功能名称用于匹配（去除标点符号）
        cleaned_function = re.sub(r"[^\w\u4e00-\u9fa5]", "", function_name)
        
        # 先查找"功能说明"部分
        function_section_start = -1
        for i, para in enumerate(self.paragraphs):
            text = para.text.strip()
            if "功能说明" in text and ("5.2" in text or "（A阶段）" in text):
                function_section_start = i
                break
        
        # 在功能说明部分查找功能名称（优先匹配精确的功能名称段落）
        search_start = function_section_start if function_section_start >= 0 else 0
        search_end = len(self.paragraphs)
        
        # 优先查找：在功能说明部分内精确匹配功能名称的段落
        if function_section_start >= 0:
            # 在功能说明部分内查找（跳过目录部分，通常目录在前100个段落）
            for i in range(max(function_section_start, 100), search_end):
                para = self.paragraphs[i]
                text = para.text.strip()
                
                # 精确匹配：段落文本就是功能名称（可能带编号）
                if function_name == text or (function_name in text and len(text) <= len(function_name) + 10):
                    # 排除目录和编号行
                    if "目录" not in text and not re.match(r'^\d+\.\d+', text):
                        function_section_index = i
                        break
        
        # 如果没找到精确匹配，使用模糊匹配（优先在功能说明部分内）
        if function_section_index < 0:
            # 先在功能说明部分内查找
            if function_section_start >= 0:
                for i in range(max(function_section_start, 100), search_end):
                    para = self.paragraphs[i]
                    text = para.text.strip()
                    
                    # 检查是否匹配功能名称（可能是标题或普通段落）
                    cleaned_text = re.sub(r"[^\w\u4e00-\u9fa5]", "", text)
                    
                    # 匹配逻辑：功能名称完全匹配，或者功能名称包含在文本中
                    if (cleaned_function in cleaned_text or cleaned_text in cleaned_function) and len(cleaned_text) >= len(cleaned_function) * 0.7:
                        # 确保不是在目录或其他不相关的地方
                        if ("功能" in text or function_name in text) and "目录" not in text:
                            # 排除目录行（通常包含页码）
                            if not re.match(r'^\d+\.\d+', text) or len(text) > 50:
                                function_section_index = i
                                break
            
            # 如果还没找到，在整个文档中查找
            if function_section_index < 0:
                for i in range(search_start, search_end):
                    para = self.paragraphs[i]
                    text = para.text.strip()
                    
                    cleaned_text = re.sub(r"[^\w\u4e00-\u9fa5]", "", text)
                    
                    if (cleaned_function in cleaned_text or cleaned_text in cleaned_function) and len(cleaned_text) >= len(cleaned_function) * 0.7:
                        if ("功能" in text or function_name in text) and "目录" not in text:
                            # 排除目录行
                            if not re.match(r'^\d+\.\d+', text) or len(text) > 50:
                                function_section_index = i
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
        
        # 扩大搜索范围：从功能章节开始，向后搜索
        search_start = max(0, function_section_index - 10)
        search_end = min(len(self.paragraphs), function_section_index + 200)
        
        # 在功能章节内查找"输入要素"和"输出要素"标记
        input_markers = ["输入要素", "输入要素：", "输入输出要素", "输入要素表"]
        output_markers = ["输出要素", "输出要素：", "输出要素表"]
        
        found_input_marker = False
        found_output_marker = False
        input_marker_index = -1
        output_marker_index = -1
        
        # 先找到"输入要素"和"输出要素"文本的位置，并检查是否"不涉及"
        # 需要找到该功能章节内的第一个"输入输出要素"标记（在"输入输出说明"下）
        input_not_involved = False
        output_not_involved = False
        found_input_output_section = False
        
        # 先查找"输入输出说明"标记，确保我们在正确的章节内
        for i in range(search_start, search_end):
            para = self.paragraphs[i]
            text = para.text.strip()
            
            # 查找"输入输出说明"或"输入输出要素"
            if "输入输出说明" in text or ("输入输出要素" in text and "：" in text):
                found_input_output_section = True
                # 从"输入输出说明"开始查找输入输出要素标记
                search_start_io = i
                break
        
        # 如果找到了"输入输出说明"章节，在该章节内查找
        if found_input_output_section:
            for i in range(search_start_io, min(search_start_io + 20, search_end)):
                para = self.paragraphs[i]
                text = para.text.strip()
                
                # 查找"输入要素"标记（必须在"输入输出说明"章节内）
                if not found_input_marker:
                    for marker in input_markers:
                        if marker in text and ("输入要素" in marker or marker == "输入要素："):
                            found_input_marker = True
                            input_marker_index = i
                            # 检查后续段落（最多5行）是否包含"不涉及"
                            # 需要跳过空行，直到找到下一个标记或"不涉及"
                            for j in range(i + 1, min(i + 6, search_end)):
                                next_text = self.paragraphs[j].text.strip()
                                # 跳过空行
                                if not next_text:
                                    continue
                                # 如果遇到下一个标记（如"输出要素"），停止检查
                                if any(m in next_text for m in output_markers):
                                    break
                                # 如果遇到下一个章节标记（如"三、"），停止检查
                                if re.match(r'^[一二三四五六七八九十]+、', next_text):
                                    break
                                # 检查是否包含"不涉及"
                                if "不涉及" in next_text:
                                    input_not_involved = True
                                    break
                            break
                
                # 查找"输出要素"标记（必须在"输入输出说明"章节内）
                if not found_output_marker:
                    for marker in output_markers:
                        if marker in text and ("输出要素" in marker or marker == "输出要素："):
                            found_output_marker = True
                            output_marker_index = i
                            # 检查后续段落（最多5行）是否包含"不涉及"
                            # 需要跳过空行，直到找到下一个标记或"不涉及"
                            for j in range(i + 1, min(i + 6, search_end)):
                                next_text = self.paragraphs[j].text.strip()
                                # 跳过空行
                                if not next_text:
                                    continue
                                # 如果遇到下一个章节标记（如"三、"），停止检查
                                if re.match(r'^[一二三四五六七八九十]+、', next_text):
                                    break
                                # 检查是否包含"不涉及"
                                if "不涉及" in next_text:
                                    output_not_involved = True
                                    break
                            break
                
                if found_input_marker and found_output_marker:
                    break
        else:
            # 如果没有找到"输入输出说明"，使用原来的逻辑
            for i in range(search_start, search_end):
                para = self.paragraphs[i]
                text = para.text.strip()
                
                # 查找"输入要素"标记
                if not found_input_marker:
                    for marker in input_markers:
                        if marker in text:
                            found_input_marker = True
                            input_marker_index = i
                            # 检查后续段落是否包含"不涉及"
                            for j in range(i + 1, min(i + 4, search_end)):
                                next_text = self.paragraphs[j].text.strip()
                                if any(m in next_text for m in output_markers):
                                    break
                                if "不涉及" in next_text:
                                    input_not_involved = True
                                    break
                            break
                
                # 查找"输出要素"标记
                if not found_output_marker:
                    for marker in output_markers:
                        if marker in text:
                            found_output_marker = True
                            output_marker_index = i
                            # 检查后续段落是否包含"不涉及"
                            for j in range(i + 1, min(i + 4, search_end)):
                                next_text = self.paragraphs[j].text.strip()
                                if re.match(r'^[一二三四五六七八九十]+、', next_text):
                                    break
                                if "不涉及" in next_text:
                                    output_not_involved = True
                                    break
                            break
                
                if found_input_marker and found_output_marker:
                    break
        
        # 智能表格定位：按顺序查找未使用的表格
        # 查找输入要素表（只有在不是"不涉及"的情况下才查找）
        if found_input_marker and not input_not_involved:
            # 方法1：直接按顺序查找第一个未使用的输入要素表
            input_elements = self._search_all_unused_tables(is_input=True)
            
            # 方法2：如果没找到，尝试查找标记后最近的一个输入要素表（即使已使用）
            # 这对于"优惠利息查询"等后面功能很重要，因为它们的表格可能在已使用的表格之后
            if not input_elements and input_marker_index >= 0:
                # 查找标记后最近的输入要素表
                input_elements = self._find_nearest_table_after_marker(
                    input_marker_index, is_input=True
                )
        elif input_not_involved:
            # 如果标记为"不涉及"，不提取输入要素
            input_elements = []
        
        # 查找输出要素表（只有在不是"不涉及"的情况下才查找）
        if found_output_marker and not output_not_involved:
            # 方法1：直接按顺序查找第一个未使用的输出要素表
            output_elements = self._search_all_unused_tables(is_input=False)
            
            # 方法2：如果没找到，尝试查找标记后最近的一个输出要素表（即使已使用）
            if not output_elements and output_marker_index >= 0:
                # 查找标记后最近的输出要素表
                output_elements = self._find_nearest_table_after_marker(
                    output_marker_index, is_input=False
                )
        elif output_not_involved:
            # 如果标记为"不涉及"，不提取输出要素
            output_elements = []
        
        return input_elements, output_elements
