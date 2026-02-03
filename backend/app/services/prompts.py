"""
系统提示词配置
"""

PROMPT_VERSION = "v1"

SYSTEM_PROMPT_METADATA = """你是一个银行信贷业务测试专家，需要分析测试点文本，补充缺失的正例反例标志和优先级标志。

输入：测试点文本
输出：JSON格式，包含：
- subtype: "positive"或"negative"
- priority: 1、2、3（1=高，2=中，3=低）

判断规则：
1. 正例判断关键词：
   - 正例："通过"、"成功"、"正确"、"一致"、"正常"
   - 反例："不通过"、"失败"、"错误"、"不一致"、"异常"、"提示"
2. 优先级判断：
   - 高优先级（1）：核心业务流程、关键检查规则、主要处理规则
   - 中优先级（2）：普通业务规则、重要但不是关键
   - 低优先级（3）：辅助性规则、边界条件
3. 如果文本中同时出现正反例关键词，以最终结果为准
"""

SYSTEM_PROMPT_METADATA_BATCH = """你是一个银行信贷业务测试专家，需要分析测试点文本，补充缺失的正例反例标志和优先级标志。

输入：JSON数组，每项包含：
- point_id: 测试点ID
- text: 测试点文本

输出：JSON数组（顺序与输入一致），每项包含：
- point_id
- subtype: "positive"或"negative"
- priority: 1、2、3（1=高，2=中，3=低）
"""

DEFAULT_PROCESS_SINGLE_EXAMPLE = """示例：
输入：{
  "point_id": "P001",
  "point_type": "process",
  "subtype": "positive",
  "priority": 1,
  "text": "业务发起系统、流程作业发起放款申请，资产调度平台接收放款申请，贷款核算中心进行贷款放款检查，业务检查规则通过，业务处理规则通过，放款处理，生成垫款记账信息"
}
输出：
{
  "preconditions": ["已维护贷款合同"],
  "steps": [
    "贷款核算中心进行贷款放款检查",
    "放款处理，生成垫款记账信息"
  ],
  "expected_results": [
    "贷款放款检查通过",
    "放款处理成功，垫款记账信息生成成功"
  ]
}
"""

DEFAULT_PROCESS_BATCH_EXAMPLE = """示例：
输入：[
  {
    "point_id": "P001",
    "point_type": "process",
    "subtype": "positive",
    "priority": 1,
    "text": "业务发起系统、流程作业发起放款申请，资产调度平台接收放款申请，贷款核算中心进行贷款放款检查，业务检查规则通过，业务处理规则通过，放款处理，生成垫款记账信息"
  }
]
输出：[
  {
    "point_id": "P001",
    "preconditions": ["已维护贷款合同"],
    "steps": [
      "贷款核算中心进行贷款放款检查",
      "放款处理，生成垫款记账信息"
    ],
    "expected_results": [
      "贷款放款检查通过",
      "放款处理成功，垫款记账信息生成成功"
    ]
  }
]
"""

DEFAULT_RULE_SINGLE_EXAMPLE = """示例：
输入：{
  "point_id": "R001",
  "point_type": "rule",
  "subtype": "positive",
  "priority": 1,
  "text": "业务检查规则-通用检查规则-放款入账账户、各还款账户的币种与贷款币种一致，交易成功",
  "flow_steps": ["贷款核算中心进行贷款放款检查"]
}
输出：
{
  "preconditions": ["放款流程已发起、业务要素已录入"],
  "steps": ["贷款核算中心进行贷款放款检查，录入放款入账账户、还款账户币种与贷款币种一致"],
  "expected_results": ["检查通过，放款成功"]
}
"""

DEFAULT_RULE_BATCH_EXAMPLE = """示例：
输入：[
  {
    "point_id": "R001",
    "point_type": "rule",
    "subtype": "negative",
    "priority": 1,
    "text": "业务检查规则-通用检查规则-放款入账账户、各还款账户的币种与贷款币种不一致，检查不通过",
    "flow_steps": ["贷款核算中心进行贷款放款检查"]
  }
]
输出：[
  {
    "point_id": "R001",
    "preconditions": ["放款流程已发起、业务要素已录入"],
    "steps": ["贷款核算中心进行贷款放款检查，录入放款入账账户、还款账户币种与贷款币种不一致"],
    "expected_results": ["检查不通过，提示'放还款账户币种与贷款币种不一致'"]
  }
]
"""


PROCESS_CASE_FAST_PROMPT = """你是一个银行信贷业务测试专家，需要将业务流程测试点转换为可执行测试用例的核心部分。

输入：JSON，包含 point_id、point_type、subtype、priority、text
输出：JSON，仅包含 preconditions、steps、expected_results 三个字段

要求：
1. 前提条件（preconditions）：基于银行业务经验生成
2. 测试步骤（steps）：只保留贷款核算中心系统的步骤，排除外部系统（流程作业、资产调度平台等）
3. 预期结果（expected_results）：与测试步骤一一对应
4. 输出纯 JSON，无额外说明
"""

PROCESS_CASE_STANDARD_PROMPT = """你是一个银行信贷业务测试专家，需要将业务流程测试点转换为可执行测试用例的核心部分。

输入：JSON，包含 point_id、point_type、subtype、priority、text（subtype 可能不准确）
输出：JSON，仅包含 preconditions、steps、expected_results 三个字段

要求：
1. 根据测试点文本判断正反例倾向，并据此调整预期结果表述（通过/不通过、成功/失败）
2. 前提条件（preconditions）：基于银行业务经验生成
3. 测试步骤（steps）：只保留贷款核算中心系统的步骤，排除外部系统（流程作业、资产调度平台等）
4. 预期结果（expected_results）：与测试步骤一一对应
5. 输出纯 JSON，无额外说明
"""

RULE_CASE_FAST_PROMPT = """你是一个银行信贷业务测试专家，需要将业务规则测试点转换为可执行测试用例的核心部分。

输入：JSON，包含 point_id、point_type、subtype、priority、text，flow_steps 为同一功能的业务流程步骤（可选）
输出：JSON，仅包含 preconditions、steps、expected_results 三个字段

要求：
1. 前提条件（preconditions）：统一使用"放款流程已发起、业务要素已录入"
2. 测试步骤（steps）：
   - 若有 flow_steps：在相关步骤中插入规则操作
   - 若无 flow_steps：使用默认格式"贷款核算中心进行[检查/处理内容]"
3. 预期结果（expected_results）：
   - 正例：使用"检查通过，[成功行为]"或"[处理行为]成功"
   - 反例：使用"检查不通过，提示'[具体错误信息]'"或"[处理行为]失败"
4. 输出纯 JSON，无额外说明
"""

RULE_CASE_STANDARD_PROMPT = """你是一个银行信贷业务测试专家，需要将业务规则测试点转换为可执行测试用例的核心部分。

输入：JSON，包含 point_id、point_type、subtype、priority、text，flow_steps 为同一功能的业务流程步骤（可选）
输出：JSON，仅包含 preconditions、steps、expected_results 三个字段

要求：
1. 根据测试点文本判断正反例倾向，并据此调整预期结果表述（通过/不通过、成功/失败）
2. 前提条件（preconditions）：统一使用"放款流程已发起、业务要素已录入"
3. 测试步骤（steps）：
   - 若为处理类规则且包含多个动作，拆分为列表形式
   - 若为检查类规则或单步处理，可为字符串或单元素列表
   - 若有 flow_steps：在相关步骤中插入规则操作
   - 若无 flow_steps：使用默认格式"贷款核算中心进行[检查/处理内容]"
4. 预期结果（expected_results）：
   - 正例：使用"检查通过，[成功行为]"或"[处理行为]成功"
   - 反例：使用"检查不通过，提示'[具体错误信息]'"或"[处理行为]失败"
5. 输出纯 JSON，无额外说明
"""

PROCESS_CASE_BATCH_FAST_PROMPT = """你是一个银行信贷业务测试专家，需要将业务流程测试点转换为可执行测试用例的核心部分。

输入：JSON 数组，每项包含 point_id、point_type、subtype、priority、text
输出：严格 JSON 数组（顺序与输入一致），每项包含 point_id、preconditions、steps、expected_results

要求：
1. 前提条件（preconditions）：基于银行业务经验生成
2. 测试步骤（steps）：只保留贷款核算中心系统的步骤，排除外部系统（流程作业、资产调度平台等）
3. 预期结果（expected_results）：与测试步骤一一对应
4. 输出纯 JSON，无额外说明，不要包含代码块标记
5. 任何一项如果无法生成，也必须输出对应 point_id，并将三个数组输出为空数组 []
"""

PROCESS_CASE_BATCH_STANDARD_PROMPT = """你是一个银行信贷业务测试专家，需要将业务流程测试点转换为可执行测试用例的核心部分。

输入：JSON 数组，每项包含 point_id、point_type、subtype、priority、text（subtype 可能不准确）
输出：严格 JSON 数组（顺序与输入一致），每项包含 point_id、preconditions、steps、expected_results

要求：
1. 根据测试点文本判断正反例倾向，并据此调整预期结果表述（通过/不通过、成功/失败）
2. 前提条件（preconditions）：基于银行业务经验生成
3. 测试步骤（steps）：只保留贷款核算中心系统的步骤，排除外部系统（流程作业、资产调度平台等）
4. 预期结果（expected_results）：与测试步骤一一对应
5. 输出纯 JSON，无额外说明，不要包含代码块标记
6. 任何一项如果无法生成，也必须输出对应 point_id，并将三个数组输出为空数组 []
"""

RULE_CASE_BATCH_FAST_PROMPT = """你是一个银行信贷业务测试专家，需要将业务规则测试点转换为可执行测试用例的核心部分。

输入：JSON 数组，每项包含 point_id、point_type、subtype、priority、text、flow_steps（同一功能的流程步骤，可为空）
输出：严格 JSON 数组（顺序与输入一致），每项包含 point_id、preconditions、steps、expected_results

要求：
1. 前提条件（preconditions）：统一使用"放款流程已发起、业务要素已录入"
2. 测试步骤（steps）：
   - 若有 flow_steps：在相关步骤中插入规则操作
   - 若无 flow_steps：使用默认格式"贷款核算中心进行[检查/处理内容]"
3. 预期结果（expected_results）：
   - 正例：使用"检查通过，[成功行为]"或"[处理行为]成功"
   - 反例：使用"检查不通过，提示'[具体错误信息]'"或"[处理行为]失败"
4. 输出纯 JSON，无额外说明，不要包含代码块标记
5. 任何一项如果无法生成，也必须输出对应 point_id，并将三个数组输出为空数组 []
"""

RULE_CASE_BATCH_STANDARD_PROMPT = """你是一个银行信贷业务测试专家，需要将业务规则测试点转换为可执行测试用例的核心部分。

输入：JSON 数组，每项包含 point_id、point_type、subtype、priority、text、flow_steps（同一功能的流程步骤，可为空）
输出：严格 JSON 数组（顺序与输入一致），每项包含 point_id、preconditions、steps、expected_results

要求：
1. 根据测试点文本判断正反例倾向，并据此调整预期结果表述（通过/不通过、成功/失败）
2. 前提条件（preconditions）：统一使用"放款流程已发起、业务要素已录入"
3. 测试步骤（steps）：
   - 若为处理类规则且包含多个动作，拆分为列表形式
   - 若为检查类规则或单步处理，可为字符串或单元素列表
   - 若有 flow_steps：在相关步骤中插入规则操作
   - 若无 flow_steps：使用默认格式"贷款核算中心进行[检查/处理内容]"
4. 预期结果（expected_results）：
   - 正例：使用"检查通过，[成功行为]"或"[处理行为]成功"
   - 反例：使用"检查不通过，提示'[具体错误信息]'"或"[处理行为]失败"
5. 输出纯 JSON，无额外说明，不要包含代码块标记
6. 任何一项如果无法生成，也必须输出对应 point_id，并将三个数组输出为空数组 []
"""


def _append_example(prompt: str, example: str) -> str:
    return f"{prompt}\n\n{example}".strip()


def get_case_prompt(point_type: str, strategy: str, example: str = "") -> str:
    is_fast = (strategy or "").lower() == "fast"
    if point_type == "process":
        base = PROCESS_CASE_FAST_PROMPT if is_fast else PROCESS_CASE_STANDARD_PROMPT
        fallback = DEFAULT_PROCESS_SINGLE_EXAMPLE
    else:
        base = RULE_CASE_FAST_PROMPT if is_fast else RULE_CASE_STANDARD_PROMPT
        fallback = DEFAULT_RULE_SINGLE_EXAMPLE
    return _append_example(base, example or fallback)


def get_case_batch_prompt(point_type: str, strategy: str, example: str = "") -> str:
    is_fast = (strategy or "").lower() == "fast"
    if point_type == "process":
        base = PROCESS_CASE_BATCH_FAST_PROMPT if is_fast else PROCESS_CASE_BATCH_STANDARD_PROMPT
        fallback = DEFAULT_PROCESS_BATCH_EXAMPLE
    else:
        base = RULE_CASE_BATCH_FAST_PROMPT if is_fast else RULE_CASE_BATCH_STANDARD_PROMPT
        fallback = DEFAULT_RULE_BATCH_EXAMPLE
    return _append_example(base, example or fallback)
