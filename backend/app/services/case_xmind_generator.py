"""
测试用例 XMind 生成器（按建模/非建模规则）
"""
import io
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import Dict, List, Optional

from app.models.schemas import ParsedXmindDocument, TestCase, TestPoint
from app.utils.logger import generator_logger


class CaseXMindGenerator:
    """将测试用例生成 XMind 文件"""

    def __init__(self, parsed: ParsedXmindDocument, cases: List[TestCase]):
        self.parsed = parsed
        self.cases = cases or []
        self._case_map = {case.point_id: case for case in self.cases if case.point_id}
        self._point_map = {point.point_id: point for point in (self.parsed.test_points or []) if point.point_id}
        self._attached_case_ids: set[str] = set()
        self._points = self._build_points()

    def _build_points(self) -> List[TestPoint]:
        points = list(self.parsed.test_points or [])
        existing_ids = {p.point_id for p in points if p.point_id}
        for case in self.cases:
            if not case.point_id or case.point_id in existing_ids:
                continue
            context = self._extract_context(case.text)
            point_type = case.point_type or ("process" if "业务流程" in (context or "") else "rule")
            points.append(TestPoint(
                point_id=case.point_id,
                point_type=point_type,
                subtype=case.subtype,
                priority=case.priority,
                text=case.text or "",
                context=context
            ))
            existing_ids.add(case.point_id)
        return points

    def _extract_context(self, text: Optional[str]) -> str:
        if not text:
            return ""
        if " - " in text:
            return text.split(" - ", 1)[0].strip()
        return text.strip()

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
        title.text = self._build_root_title()

        children = ET.SubElement(topic, "children")
        topics_container = ET.SubElement(children, "topics", {"type": "attached"})

        if self.parsed.document_type == "modeling":
            self._build_modeling(topics_container)
        else:
            self._build_non_modeling(topics_container)

        ET.indent(root, space="  ")
        return ET.tostring(root, encoding="unicode", xml_declaration=True)

    def _build_root_title(self) -> str:
        doc_num = self.parsed.document_number or "需求说明书编号"
        requirement = self.parsed.requirement_name or "测试用例"
        return f"{doc_num}-{requirement}"

    def _build_basic_info(self, parent: ET.Element):
        basic = self._create_topic(parent, "基础信息")
        children = ET.SubElement(basic, "children")
        topics = ET.SubElement(children, "topics", {"type": "attached"})
        if self.parsed.document_type == "modeling":
            customer = self.parsed.customer or "/"
            product = self.parsed.product or "/"
            channel = self.parsed.channel or "/"
            partner = self.parsed.partner or "/"
            designer = self.parsed.designer or "/"
            self._create_topic(topics, f"客户（C）：{customer}")
            self._create_topic(topics, f"产品（P）：{product}")
            self._create_topic(topics, f"渠道（C）：{channel}")
            self._create_topic(topics, f"合作方（P）：{partner}")
            self._create_topic(topics, f"设计者：{designer}")
        else:
            designer = self.parsed.designer or "/"
            self._create_topic(topics, f"设计者：{designer}")

    def _build_modeling(self, parent: ET.Element):
        self._build_basic_info(parent)
        activity_names, component_map = self._collect_modeling_structure()

        for activity in activity_names:
            activity_topic = self._create_topic(parent, activity)
            self._add_fixed_leaf_nodes(activity_topic)
            self._append_points(activity_topic, activity, None, None)

        for component, task_map in component_map.items():
            component_topic = self._create_topic(parent, component)
            children = ET.SubElement(component_topic, "children")
            task_container = ET.SubElement(children, "topics", {"type": "attached"})
            for task, step_list in task_map.items():
                task_topic = self._create_topic(task_container, task)
                if step_list and None in step_list:
                    self._add_fixed_leaf_nodes(task_topic)
                    self._append_points(task_topic, component, task, None)
                concrete_steps = [s for s in step_list if s]
                if concrete_steps:
                    task_children = ET.SubElement(task_topic, "children")
                    step_container = ET.SubElement(task_children, "topics", {"type": "attached"})
                    for step in concrete_steps:
                        step_topic = self._create_topic(step_container, step)
                        self._add_fixed_leaf_nodes(step_topic)
                        self._append_points(step_topic, component, task, step)

    def _build_non_modeling(self, parent: ET.Element):
        self._build_basic_info(parent)
        top_nodes = self._collect_non_modeling_nodes()
        for top in top_nodes:
            top_topic = self._create_topic(parent, top)
            self._add_fixed_leaf_nodes(top_topic)
            self._append_points(top_topic, top, None, None)

    def _collect_modeling_structure(self) -> tuple[List[str], Dict[str, Dict[str, List[str]]]]:
        activities: List[str] = []
        components: Dict[str, Dict[str, List[str]]] = {}
        for point in self._points:
            segments = self._split_context(point.context)
            segments = self._strip_root(segments)
            if not segments:
                continue
            if segments[-1] in ("业务流程", "业务规则"):
                segments = segments[:-1]
            if not segments:
                continue
            if len(segments) == 1:
                if segments[0] not in activities:
                    activities.append(segments[0])
            elif len(segments) == 2:
                component, task = segments[0], segments[1]
                task_map = components.setdefault(component, {})
                step_list = task_map.setdefault(task, [])
                if None not in step_list:
                    step_list.append(None)
            elif len(segments) >= 3:
                component, task, step = segments[0], segments[1], segments[2]
                task_map = components.setdefault(component, {})
                step_list = task_map.setdefault(task, [])
                if step not in step_list:
                    step_list.append(step)
        return activities, components

    def _collect_non_modeling_nodes(self) -> List[str]:
        nodes: List[str] = []
        requirement = self.parsed.requirement_name or "测试用例"
        for point in self._points:
            segments = self._split_context(point.context)
            segments = self._strip_root(segments)
            if not segments:
                top = requirement
            else:
                if segments[0] == requirement and len(segments) >= 2:
                    top = segments[1]
                else:
                    top = segments[0]
            if top not in nodes:
                nodes.append(top)
        if requirement not in nodes:
            nodes.insert(0, requirement)
        return nodes

    def _append_points(
        self,
        parent: ET.Element,
        top: str,
        task: Optional[str],
        step: Optional[str]
    ):
        added = 0
        for point in self._points:
            case = self._case_map.get(point.point_id)
            if not case:
                continue
            if not point.context:
                continue
            if not self._match_point(point, top, task, step, point.context):
                continue
            branch = self._find_leaf_node(parent, point.point_type)
            if branch is None:
                continue
            if self._add_case_node(branch, point, case):
                added += 1


    def _match_point(
        self,
        point: TestPoint,
        top: str,
        task: Optional[str],
        step: Optional[str],
        context_override: Optional[str] = None
    ) -> bool:
        segments = self._split_context(context_override or point.context)
        segments = self._strip_root(segments)
        if not segments:
            return False
        if segments[-1] in ("业务流程", "业务规则"):
            leaf = segments[-1]
            path = segments[:-1]
        else:
            leaf = ""
            path = segments
        if point.point_type == "process" and leaf != "业务流程":
            return False
        if point.point_type == "rule" and leaf != "业务规则":
            return False
        if self.parsed.document_type == "modeling":
            if step:
                return self._contains_sequence(path, [top, task or "", step])
            if task:
                return self._contains_sequence(path, [top, task])
            return len(path) == 1 and path[0] == top
        requirement = self.parsed.requirement_name or "测试用例"
        if path and path[0] == requirement and len(path) >= 2:
            return path[1] == top
        return path and path[0] == top or top in path

    def _find_leaf_node(self, parent: ET.Element, point_type: str) -> Optional[ET.Element]:
        children = parent.find("children")
        if children is None:
            return None
        topics = children.find("topics")
        if topics is None:
            return None
        target = "业务流程" if point_type == "process" else "业务规则"
        for topic in topics.findall("topic"):
            title = topic.find("title")
            if title is not None and title.text == target:
                return topic
        return None

    def _add_fixed_leaf_nodes(self, parent: ET.Element):
        topics_container = self._ensure_child_topics(parent)
        for title in ["业务流程", "业务规则", "页面控制", "数据验证"]:
            self._create_topic(topics_container, title)

    def _add_case_node(self, parent: ET.Element, point: TestPoint, case: TestCase) -> bool:
        if not case:
            return False
        if case.case_id in self._attached_case_ids:
            return False
        case_title = self._strip_context_from_title(case.text or point.text or point.point_id)
        case_topic = self._create_child_topic(parent, case_title, point.priority)
        pre_content = self._create_child_topic(case_topic, self._join_with_index(case.preconditions))
        steps_content = self._create_child_topic(pre_content, self._join_with_index(case.steps))
        expected_text = self._join_with_index(case.expected_results)
        if point.subtype == "negative" and expected_text:
            expected_text = f"❌{expected_text}"
        self._create_child_topic(steps_content, expected_text)
        self._attached_case_ids.add(case.case_id)
        return True

    def _add_single_child(self, parent: ET.Element, text: str):
        if text is None:
            text = ""
        children = ET.SubElement(parent, "children")
        topics_container = ET.SubElement(children, "topics", {"type": "attached"})
        self._create_topic(topics_container, text)

    def _join_lines(self, items: List[str]) -> str:
        return "\n".join([i for i in items if i]) if items else ""

    def _join_with_index(self, items: Optional[List[str]]) -> str:
        if not items:
            return ""
        cleaned = [i for i in items if i]
        if len(cleaned) <= 1:
            return cleaned[0] if cleaned else ""
        return "\n".join([f"{idx}、{val}" for idx, val in enumerate(cleaned, start=1)])

    def _split_context(self, context: Optional[str]) -> List[str]:
        if not context:
            return []
        return [seg.strip() for seg in context.split(" / ") if seg.strip()]

    def _strip_context_from_title(self, text: Optional[str]) -> str:
        if not text:
            return ""
        if " - " in text:
            return text.split(" - ", 1)[-1].strip()
        return text.strip()

    def _ensure_child_topics(self, parent: ET.Element) -> ET.Element:
        children = parent.find("children")
        if children is None:
            children = ET.SubElement(parent, "children")
        topics = children.find("topics")
        if topics is None:
            topics = ET.SubElement(children, "topics", {"type": "attached"})
        return topics

    def _create_child_topic(self, parent: ET.Element, title_text: str, priority: Optional[int] = None) -> ET.Element:
        topics = self._ensure_child_topics(parent)
        return self._create_topic(topics, title_text, priority)

    def _contains_sequence(self, path: List[str], target: List[str]) -> bool:
        if not target:
            return False
        if len(target) > len(path):
            return False
        for i in range(0, len(path) - len(target) + 1):
            if path[i:i + len(target)] == target:
                return True
        return False

    def _fuzzy_match(
        self,
        text: Optional[str],
        point_type: str,
        top: str,
        task: Optional[str],
        step: Optional[str]
    ) -> bool:
        if not text:
            return False
        if point_type == "process" and "业务流程" not in text:
            return False
        if point_type == "rule" and "业务规则" not in text:
            return False
        if top and top not in text:
            return False
        if task and task not in text:
            return False
        if step and step not in text:
            return False
        return True

    def _strip_root(self, segments: List[str]) -> List[str]:
        if not segments:
            return segments
        root = segments[0]
        requirement = self.parsed.requirement_name or ""
        if requirement and requirement in root and ("需求项目编号" in root or "需求说明书编号" in root):
            return segments[1:]
        if root == requirement:
            return segments[1:]
        return segments

    def _create_topic(self, parent, title_text: str, priority: Optional[int] = None) -> ET.Element:
        import uuid
        topic_elem = ET.SubElement(parent, "topic", {"id": uuid.uuid4().hex[:26]})
        title = ET.SubElement(topic_elem, "title")
        title.text = title_text or ""
        if priority in (1, 2, 3):
            marker_refs = ET.SubElement(topic_elem, "marker-refs")
            ET.SubElement(marker_refs, "marker-ref", {"marker-id": f"priority-{priority}"})
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
