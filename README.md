# 测试大纲生成器

自动化工具，将Word格式需求文档转换为XMind格式的测试大纲。

## 技术栈

- 前端：Vue 3 + Element Plus + Vite
- 后端：Python 3 + FastAPI
- 文档解析：python-docx
- XMind生成：直接生成XMind XML格式

## 快速开始

### 使用启动脚本（推荐）

运行 `start.bat` 一键启动所有服务。

### 手动启动

**后端：**

```bash
cd backend
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

**前端：**

```bash
cd frontend
npm install
npm run dev
```

## 功能特性

- Word文档解析和验证
- 自动提取需求用例基本信息
- 自动提取活动、组件、任务、步骤信息
- 支持输入输出要素提取和格式化
- 生成标准化XMind测试大纲（逻辑图向右）
- 解析结果预览确认
- 支持多文件批量处理（最多5个文件）

## 使用说明

1. 启动服务后，打开前端页面
2. 上传Word格式的需求文档（最多5个文件）
3. 点击"解析并生成测试大纲"按钮
4. 在预览弹窗中核对解析结果
5. 点击"确认并生成"生成XMind文件
6. 可通过"调试信息"按钮查看详细JSON数据

## 注意事项

- 文档需包含"用例版本控制信息"表
- 文档格式需符合规范
- 生成的XMind文件可在XMind软件中打开和编辑
