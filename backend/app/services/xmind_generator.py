"""
XMind文件生成服务 - 银行需求文档专用生成器
"""
import io
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import List
from app.models.schemas import ParsedDocument, ActivityInfo, ComponentInfo, TaskInfo, StepInfo


class XMindGenerator:
    """XMind文件生成器（直接生成XMind XML格式）"""
    
    def __init__(self, parsed_doc: ParsedDocument):
        self.parsed_doc = parsed_doc
    
    def generate(self) -> bytes:
        """生成XMind文件并返回字节流"""
        # 创建内存中的ZIP文件
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # 创建content.xml
            content_xml = self._create_content_xml()
            zip_file.writestr('content.xml', content_xml.encode('utf-8'))
            
            # 创建meta.xml
            meta_xml = self._create_meta_xml()
            zip_file.writestr('META-INF/manifest.xml', meta_xml.encode('utf-8'))
            
            # 创建styles.xml（可选，用于样式）
            styles_xml = self._create_styles_xml()
            zip_file.writestr('styles.xml', styles_xml.encode('utf-8'))
        
        zip_buffer.seek(0)
        return zip_buffer.getvalue()
    
    def _create_content_xml(self) -> str:
        """创建content.xml内容"""
        # 创建XML根元素
        root = ET.Element('xmap-content', {
            'xmlns': 'urn:xmind:xmap:xmlns:content:2.0',
            'xmlns:fo': 'http://www.w3.org/1999/XSL/Format',
            'version': '2.0'
        })
        
        import uuid
        sheet_id = uuid.uuid4().hex[:26]
        sheet = ET.SubElement(root, 'sheet', {'id': sheet_id})
        topic_id = uuid.uuid4().hex[:26]
        
        # 设置主题结构为逻辑图向右 - 作为topic元素的属性
        topic = ET.SubElement(sheet, 'topic', {
            'id': topic_id,
            'structure-class': 'org.xmind.ui.logic.right'
        })
        
        # 设置根节点标题：需求用例名称-版本号
        title = ET.SubElement(topic, 'title')
        root_title = self._build_root_title()
        title.text = str(root_title) if root_title else "测试大纲"
        
        # 创建children容器
        children = ET.SubElement(topic, 'children')
        topics_container = ET.SubElement(children, 'topics', {'type': 'attached'})
        
        # 根据文档类型生成不同的结构
        if self.parsed_doc.document_type == "non_modeling":
            self._add_non_modeling_structure(topics_container)
        else:
            # 建模需求结构
            # 1. 添加基础信息（固定节点）
            basic_info_topic = self._create_topic_element(topics_container, '基础信息')
            self._add_basic_info(basic_info_topic)
            
            # 2. 添加活动名称（如果有）
            if self.parsed_doc.activities:
                for activity in self.parsed_doc.activities:
                    if activity and activity.name:
                        activity_topic = self._create_topic_element(topics_container, activity.name)
                        self._add_activity_fixed_nodes(activity_topic)
            
            # 3. 添加组件名称（如果有）
            if self.parsed_doc.activities:
                for activity in self.parsed_doc.activities:
                    if activity and activity.components:
                        for component in activity.components:
                            if component and component.name:
                                component_topic = self._create_topic_element(topics_container, component.name)
                                self._add_component(component_topic, component)
        
        # 转换为XML字符串
        ET.indent(root, space='  ')
        xml_str = ET.tostring(root, encoding='unicode', xml_declaration=True)
        return xml_str
    
    def _create_meta_xml(self) -> str:
        """创建manifest.xml内容"""
        manifest = ET.Element('manifest', {
            'xmlns': 'urn:xmind:xmap:xmlns:manifest:1.0'
        })
        
        file_entry = ET.SubElement(manifest, 'file-entry', {
            'full-path': 'content.xml',
            'media-type': 'text/xml'
        })
        
        file_entry2 = ET.SubElement(manifest, 'file-entry', {
            'full-path': 'styles.xml',
            'media-type': 'text/xml'
        })
        
        file_entry3 = ET.SubElement(manifest, 'file-entry', {
            'full-path': 'META-INF/',
            'media-type': ''
        })
        
        ET.indent(manifest, space='  ')
        xml_str = ET.tostring(manifest, encoding='unicode', xml_declaration=True)
        return xml_str
    
    def _create_styles_xml(self) -> str:
        """创建styles.xml内容"""
        styles = ET.Element('xmap-styles', {
            'xmlns': 'urn:xmind:xmap:xmlns:style:2.0',
            'version': '2.0'
        })
        
        # 添加逻辑图向右的样式定义（如果需要）
        # 某些XMind版本可能需要显式定义样式
        
        ET.indent(styles, space='  ')
        xml_str = ET.tostring(styles, encoding='unicode', xml_declaration=True)
        return xml_str
    
    def _create_topic_element(self, parent, title_text: str) -> ET.Element:
        """创建主题元素
        
        parent应该是topics容器（type='attached'）
        返回创建的topic元素
        """
        import uuid
        topic_elem = ET.SubElement(parent, 'topic', {'id': uuid.uuid4().hex[:26]})
        title = ET.SubElement(topic_elem, 'title')
        # 确保文本不为None
        title.text = str(title_text) if title_text else ""
        return topic_elem
    
    def _build_root_title(self) -> str:
        """构建根节点标题"""
        # 非建模需求：需求说明书编号-需求名称
        if self.parsed_doc.document_type == "non_modeling":
            file_number = self.parsed_doc.file_number or ""
            requirement_name = self.parsed_doc.requirement_name or ""
            
            if file_number and requirement_name:
                return f"{file_number}-{requirement_name}"
            elif requirement_name:
                return requirement_name
            elif file_number:
                return f"{file_number}-测试大纲"
            else:
                return "测试大纲"
        
        # 建模需求：需求用例名称-版本号
        case_name = self.parsed_doc.requirement_info.case_name or ""
        version = self.parsed_doc.version or ""
        
        if case_name and version:
            return f"{case_name}-{version}"
        elif case_name:
            return case_name
        elif version:
            return f"测试大纲-{version}"
        else:
            return "测试大纲"
    
    def _add_basic_info(self, parent_topic: ET.Element):
        """添加基础信息节点"""
        # 非建模需求：只显示设计者
        if self.parsed_doc.document_type == "non_modeling":
            # 创建children和topics容器
            children = ET.SubElement(parent_topic, 'children')
            topics_container = ET.SubElement(children, 'topics', {'type': 'attached'})
            
            designer = self.parsed_doc.designer or ""
            designer_text = f"设计者：{designer}" if designer else "设计者："
            self._create_topic_element(topics_container, designer_text)
            return
        
        # 建模需求：显示客户、产品、渠道、合作方、设计者
        req_info = self.parsed_doc.requirement_info
        
        # 创建children和topics容器
        children = ET.SubElement(parent_topic, 'children')
        topics_container = ET.SubElement(children, 'topics', {'type': 'attached'})
        
        # 按固定顺序添加：客户、产品、渠道、合作方、设计者
        customer_text = f"客户（C）：{req_info.customer}" if req_info and req_info.customer else "客户（C）："
        self._create_topic_element(topics_container, customer_text)
        
        product_text = f"产品（P）：{req_info.product}" if req_info and req_info.product else "产品（P）："
        self._create_topic_element(topics_container, product_text)
        
        channel_text = f"渠道（C）：{req_info.channel}" if req_info and req_info.channel else "渠道（C）："
        self._create_topic_element(topics_container, channel_text)
        
        partner_text = f"合作方（P）：{req_info.partner}" if req_info and req_info.partner else "合作方（P）："
        self._create_topic_element(topics_container, partner_text)
        
        self._create_topic_element(topics_container, "设计者：")
    
    def _add_activity_fixed_nodes(self, parent_topic: ET.Element):
        """添加活动节点的固定子节点（业务流程、业务规则等）"""
        # 创建children和topics容器
        children = ET.SubElement(parent_topic, 'children')
        topics_container = ET.SubElement(children, 'topics', {'type': 'attached'})
        
        # 添加固定子节点
        for title in ["业务流程", "业务规则", "页面控制", "数据验证"]:
            self._create_topic_element(topics_container, title)
    
    def _add_component(self, parent_topic: ET.Element, component: ComponentInfo):
        """添加组件节点"""
        if not component:
            return
        
        # 创建children和topics容器
        children = ET.SubElement(parent_topic, 'children')
        topics_container = ET.SubElement(children, 'topics', {'type': 'attached'})
        
        # 添加任务
        if component.tasks:
            for task in component.tasks:
                if task and task.name:
                    task_topic = self._create_topic_element(topics_container, task.name)
                    self._add_task(task_topic, task)
    
    def _add_task(self, parent_topic: ET.Element, task: TaskInfo):
        """添加任务节点"""
        if not task:
            return
        
        # 创建children和topics容器
        children = ET.SubElement(parent_topic, 'children')
        topics_container = ET.SubElement(children, 'topics', {'type': 'attached'})
        
        # 添加步骤
        if task.steps:
            for step in task.steps:
                if step and step.name:
                    step_topic = self._create_topic_element(topics_container, step.name)
                    self._add_step(step_topic, step)
    
    def _add_step(self, parent_topic: ET.Element, step: StepInfo):
        """添加步骤节点"""
        if not step:
            return
        
        # 创建children和topics容器
        children = ET.SubElement(parent_topic, 'children')
        topics_container = ET.SubElement(children, 'topics', {'type': 'attached'})
        
        # 添加固定子节点：业务流程、业务规则、页面控制、数据验证
        page_control_topic = None
        for title in ["业务流程", "业务规则", "页面控制", "数据验证"]:
            topic_elem = self._create_topic_element(topics_container, title)
            if title == "页面控制":
                page_control_topic = topic_elem
        
        # 将输入输出要素添加到页面控制节点下
        if page_control_topic:
            # 为页面控制创建children和topics容器
            page_control_children = ET.SubElement(page_control_topic, 'children')
            page_control_topics = ET.SubElement(page_control_children, 'topics', {'type': 'attached'})
            
            # 添加输入要素（节点名称前加"输入-"）
            if step.input_elements:
                for elem in step.input_elements:
                    if elem:
                        input_text = self._format_input_element(elem)
                        self._create_topic_element(page_control_topics, f"输入-{input_text}")
            
            # 添加输出要素（节点名称前加"输出-"）
            if step.output_elements:
                for elem in step.output_elements:
                    if elem:
                        output_text = self._format_output_element(elem)
                        self._create_topic_element(page_control_topics, f"输出-{output_text}")
    
    def _format_input_element(self, elem) -> str:
        """格式化输入要素节点文本
        
        规则：
        1. 如果说明不为空：字段名称-是否必输-说明
        2. 如果说明为空且字段格式非下拉框：字段名称-是否必输
        3. 如果字段格式为下拉框：字段名称-是否必输；下拉选项包括：输入限制
        """
        if not elem:
            return ""
        
        field_name = str(elem.field_name) if elem.field_name else ""
        required = str(elem.required) if elem.required else "否"
        required_text = "必输" if required == "是" else "非必输"
        
        # 判断是否为下拉框类型
        field_format = str(elem.field_format) if elem.field_format else ""
        input_limit = str(elem.input_limit) if elem.input_limit else ""
        is_dropdown = "下拉" in field_format or "下拉" in input_limit
        
        # 规则3：下拉框类型
        if is_dropdown and input_limit:
            return f"{field_name}-{required_text}；下拉选项包括：{input_limit}"
        
        # 规则1：有说明
        description = str(elem.description) if elem.description else ""
        if description:
            return f"{field_name}-{required_text}-{description}"
        
        # 规则2：无说明且非下拉框
        return f"{field_name}-{required_text}"
    
    def _format_output_element(self, elem) -> str:
        """格式化输出要素节点文本：字段名称-类型-说明"""
        if not elem:
            return ""
        
        parts = [str(elem.field_name) if elem.field_name else ""]
        
        if elem.field_type:
            parts.append(str(elem.field_type))
        
        if elem.description:
            parts.append(str(elem.description))
        
        return "-".join(parts)
    
    # ========== 非建模需求结构生成方法 ==========
    
    def _add_non_modeling_structure(self, topics_container: ET.Element):
        """添加非建模需求的结构"""
        # 1. 添加基础信息（固定节点）
        basic_info_topic = self._create_topic_element(topics_container, '基础信息')
        self._add_basic_info(basic_info_topic)
        
        # 2. 添加需求名称节点（如果有）
        requirement_name = self.parsed_doc.requirement_name
        if requirement_name:
            requirement_topic = self._create_topic_element(topics_container, requirement_name)
            self._add_activity_fixed_nodes(requirement_topic)
        
        # 3. 添加功能节点
        if self.parsed_doc.functions:
            for function in self.parsed_doc.functions:
                if function and function.name:
                    function_topic = self._create_topic_element(topics_container, function.name)
                    self._add_function(function_topic, function)
    
    def _add_function(self, parent_topic: ET.Element, function):
        """添加功能节点（非建模需求）"""
        if not function:
            return
        
        # 创建children和topics容器
        children = ET.SubElement(parent_topic, 'children')
        topics_container = ET.SubElement(children, 'topics', {'type': 'attached'})
        
        # 添加固定子节点：业务流程、业务规则、页面控制、数据验证
        page_control_topic = None
        for title in ["业务流程", "业务规则", "页面控制", "数据验证"]:
            topic_elem = self._create_topic_element(topics_container, title)
            if title == "页面控制":
                page_control_topic = topic_elem
        
        # 将输入输出要素添加到页面控制节点下
        if page_control_topic:
            # 为页面控制创建children和topics容器
            page_control_children = ET.SubElement(page_control_topic, 'children')
            page_control_topics = ET.SubElement(page_control_children, 'topics', {'type': 'attached'})
            
            # 添加输入要素（按序号排序）
            if function.input_elements:
                sorted_inputs = sorted(function.input_elements, key=lambda x: x.index)
                for elem in sorted_inputs:
                    if elem:
                        input_text = self._format_input_element(elem)
                        self._create_topic_element(page_control_topics, input_text)
            
            # 添加输出要素（按序号排序）
            if function.output_elements:
                sorted_outputs = sorted(function.output_elements, key=lambda x: x.index)
                for elem in sorted_outputs:
                    if elem:
                        output_text = self._format_output_element(elem)
                        self._create_topic_element(page_control_topics, output_text)
