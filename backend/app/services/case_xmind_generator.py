"""
测试用例 XMind 生成器
"""
import io
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import List

from app.models.schemas import TestCase
from app.utils.logger import generator_logger


class CaseXMindGenerator:
    """将测试用例生成 XMind 文件"""

    def __init__(self, requirement_name: str, cases: List[TestCase]):
        self.requirement_name = requirement_name or "测试用例"
        self.cases = cases or []

    def generate(self) -> bytes:
        import time
        start_time = time.time()

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            content_xml = self._create_content_xml()
            zip_file.writestr("content.xml", content_xml.encode("utf-8"))
            meta_xml = self._create_meta_xml()
            zip_file.writestr("META-INF/manifest.xml", meta_xml.encode("utf-8"))
            styles_xml = self._create_styles_xml()
            zip_file.writestr("styles.xml", styles_xml.encode("utf-8"))

        zip_buffer.seek(0)
        result = zip_buffer.getvalue()
        elapsed_time = time.time() - start_time
        generator_logger.info(
            f"测试用例XMind生成成功 - 文件大小: {len(result) / 1024:.2f}KB, 耗时: {elapsed_time:.3f}秒"
        )
        return result

    def _create_content_xml(self) -> str:
        root = ET.Element("xmap-content", {
            "xmlns": "urn:xmind:xmap:xmlns:content:2.0",
            "xmlns:fo": "http://www.w3.org/1999/XSL/Format",
            "version": "2.0"
        })

        import uuid
        sheet_id = uuid.uuid4().hex[:26]
        sheet = ET.SubElement(root, "sheet", {"id": sheet_id})
        topic_id = uuid.uuid4().hex[:26]
        topic = ET.SubElement(sheet, "topic", {
            "id": topic_id,
            "structure-class": "org.xmind.ui.logic.right"
        })

        title = ET.SubElement(topic, "title")
        title.text = self.requirement_name

        children = ET.SubElement(topic, "children")
        topics_container = ET.SubElement(children, "topics", {"type": "attached"})

        for case in self.cases:
            case_topic = self._create_topic(topics_container, case.text or case.case_id)
            self._add_case_nodes(case_topic, case)

        ET.indent(root, space="  ")
        return ET.tostring(root, encoding="unicode", xml_declaration=True)

    def _add_case_nodes(self, parent_topic: ET.Element, case: TestCase):
        children = ET.SubElement(parent_topic, "children")
        topics_container = ET.SubElement(children, "topics", {"type": "attached"})

        self._create_topic(topics_container, f"用例编号：{case.case_id}")
        if case.priority:
            self._create_topic(topics_container, f"优先级：{case.priority}")
        if case.point_type:
            self._create_topic(topics_container, f"测试点类型：{case.point_type}")
        if case.subtype:
            self._create_topic(topics_container, f"子类型：{case.subtype}")

        pre_topic = self._create_topic(topics_container, "前提条件")
        self._add_list_children(pre_topic, case.preconditions)

        step_topic = self._create_topic(topics_container, "测试步骤")
        self._add_list_children(step_topic, case.steps)

        expected_topic = self._create_topic(topics_container, "预期结果")
        self._add_list_children(expected_topic, case.expected_results)

    def _add_list_children(self, parent_topic: ET.Element, items: List[str]):
        if not items:
            return
        children = ET.SubElement(parent_topic, "children")
        topics_container = ET.SubElement(children, "topics", {"type": "attached"})
        for item in items:
            self._create_topic(topics_container, item)

    def _create_topic(self, parent, title_text: str) -> ET.Element:
        import uuid
        topic_elem = ET.SubElement(parent, "topic", {"id": uuid.uuid4().hex[:26]})
        title = ET.SubElement(topic_elem, "title")
        title.text = title_text or ""
        return topic_elem

    def _create_meta_xml(self) -> str:
        manifest = ET.Element("manifest", {
            "xmlns": "urn:xmind:xmap:xmlns:manifest:1.0"
        })
        ET.SubElement(manifest, "file-entry", {"full-path": "content.xml", "media-type": "text/xml"})
        ET.SubElement(manifest, "file-entry", {"full-path": "styles.xml", "media-type": "text/xml"})
        ET.SubElement(manifest, "file-entry", {"full-path": "META-INF/", "media-type": ""})
        ET.indent(manifest, space="  ")
        return ET.tostring(manifest, encoding="unicode", xml_declaration=True)

    def _create_styles_xml(self) -> str:
        styles = ET.Element("xmap-styles", {
            "xmlns": "urn:xmind:xmap:xmlns:style:2.0",
            "version": "2.0"
        })
        ET.indent(styles, space="  ")
        return ET.tostring(styles, encoding="unicode", xml_declaration=True)
