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


class TaskCreateResponse(BaseModel):
    """任务创建响应"""
    success: bool
    task_id: str
    message: str


class TaskStatusResponse(BaseModel):
    """任务状态响应"""
    task_id: str
    status: str  # pending, processing, completed, failed
    filename: str
    progress: float  # 0.0 - 1.0
    created_at: str
    started_at: Optional[str] = None
    completed_at: Optional[str] = None
    error: Optional[str] = None
    result: Optional[ParseResponse] = None


class TestPoint(BaseModel):
    """测试点"""
    point_id: str
    point_type: str  # process / rule
    subtype: Optional[str] = None  # positive / negative
    priority: Optional[int] = None  # 1/2/3
    text: str
    context: Optional[str] = None


class ParsedXmindDocument(BaseModel):
    """XMind解析结果"""
    parse_id: str
    requirement_name: str
    document_type: str  # modeling / non_modeling
    test_points: List[TestPoint] = []
    stats: Dict[str, Any] = {}


class ParseXmindResponse(BaseModel):
    """XMind解析响应"""
    success: bool
    message: str
    data: Optional[ParsedXmindDocument] = None


class TestCase(BaseModel):
    """生成的测试用例"""
    case_id: str
    point_id: str
    point_type: str
    subtype: Optional[str] = None
    priority: Optional[int] = None
    text: str
    preconditions: List[str] = []
    steps: List[str] = []
    expected_results: List[str] = []


class PreviewGenerateRequest(BaseModel):
    """预生成请求"""
    parse_id: str
    count: Optional[int] = None


class PreviewGenerateResponse(BaseModel):
    """预生成响应"""
    success: bool
    message: str
    preview_id: Optional[str] = None
    cases: List[TestCase] = []


class ConfirmPreviewRequest(BaseModel):
    """确认预生成请求"""
    preview_id: str
    strategy: Optional[str] = "standard"


class BulkGenerateRequest(BaseModel):
    """批量生成请求"""
    parse_id: str
    strategy: Optional[str] = "standard"


class GenerationStatusResponse(BaseModel):
    """生成任务状态响应"""
    task_id: str
    status: str
    progress: float
    total: int
    completed: int
    failed: int
    logs: List[str] = []
    cases: List[TestCase] = []
    token_usage: int = 0
    error: Optional[str] = None


class RetryGenerationRequest(BaseModel):
    """重新生成请求"""
    task_id: str
    strategy: Optional[str] = "standard"


class ExportCasesRequest(BaseModel):
    """导出测试用例请求"""
    requirement_name: str
    cases: List[TestCase] = []

