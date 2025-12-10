"""
数据模型定义
"""
from typing import List, Optional, Dict, Any
from pydantic import BaseModel


class RequirementInfo(BaseModel):
    """需求用例基本信息"""
    case_name: str  # 用例名称
    channel: Optional[str] = None  # 渠道（C）
    product: Optional[str] = None  # 产品（P）
    customer: Optional[str] = None  # 客户（C）
    partner: Optional[str] = None  # 合作方（P）


class InputElement(BaseModel):
    """输入要素"""
    index: int  # 序号
    field_name: str  # 字段名称
    required: str  # 是否必输（是/否）
    field_type: Optional[str] = None  # 类型
    precision: Optional[str] = None  # 精度
    field_format: Optional[str] = None  # 字段格式
    input_limit: Optional[str] = None  # 输入限制
    description: Optional[str] = None  # 说明


class OutputElement(BaseModel):
    """输出要素"""
    index: int  # 序号
    field_name: str  # 字段名称
    field_type: Optional[str] = None  # 类型
    precision: Optional[str] = None  # 精度
    field_format: Optional[str] = None  # 字段格式
    description: Optional[str] = None  # 说明


class StepInfo(BaseModel):
    """步骤信息"""
    name: str  # 步骤名称
    input_elements: List[InputElement] = []  # 输入要素
    output_elements: List[OutputElement] = []  # 输出要素


class TaskInfo(BaseModel):
    """任务信息"""
    name: str  # 任务名称
    steps: List[StepInfo] = []  # 步骤列表


class ComponentInfo(BaseModel):
    """组件信息"""
    name: str  # 组件名称
    tasks: List[TaskInfo] = []  # 任务列表


class ActivityInfo(BaseModel):
    """活动信息"""
    name: str  # 活动名称
    components: List[ComponentInfo] = []  # 组件列表


class FunctionInfo(BaseModel):
    """功能信息（用于非建模需求）"""
    name: str  # 功能名称
    input_elements: List[InputElement] = []  # 输入要素
    output_elements: List[OutputElement] = []  # 输出要素


class ParsedDocument(BaseModel):
    """解析后的文档数据"""
    version: str  # 版本编号
    requirement_info: RequirementInfo  # 需求用例基本信息
    activities: List[ActivityInfo] = []  # 活动列表（建模需求）
    document_number: Optional[str] = None  # 需求说明书编号
    case_number: Optional[str] = None  # 需求用例编号
    # 非建模需求相关字段
    document_type: Optional[str] = None  # 文档类型："modeling" 或 "non_modeling"
    file_number: Optional[str] = None  # 文件编号
    file_name: Optional[str] = None  # 文件名称
    requirement_name: Optional[str] = None  # 需求名称
    designer: Optional[str] = None  # 设计者
    functions: List[FunctionInfo] = []  # 功能列表（非建模需求）


class ParseResponse(BaseModel):
    """解析响应"""
    success: bool
    message: str
    data: Optional[ParsedDocument] = None


class GenerateOutlineRequest(BaseModel):
    """生成大纲请求"""
    parsed_data: ParsedDocument

