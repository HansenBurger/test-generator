"""
XMind 文件解析服务
将 XMind 测试大纲解析为测试点 JSON
"""
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Optional, Tuple
from uuid import uuid4

from app.models.schemas import TestPoint, ParsedXmindDocument
from app.utils.logger import parser_logger


class XMindParser:
    """XMind 解析器"""

    _NS = {"x": "urn:xmind:xmap:xmlns:content:2.0"}

    def __init__(self, xmind_path: str):
        self.xmind_path = xmind_path

    def parse(self) -> ParsedXmindDocument:
        content_xml = self._read_content_xml()
        root = ET.fromstring(content_xml)

        sheet = root.find("x:sheet", self._NS)
        if sheet is None:
            raise ValueError("XMind文件解析失败：未找到sheet节点")

        root_topic = sheet.find("x:topic", self._NS)
        if root_topic is None:
            raise ValueError("XMind文件解析失败：未找到根主题节点")

        root_title = self._get_title(root_topic)
        document_number, requirement_name = self._parse_root_title(root_title)
        requirement_name = requirement_name or "测试大纲"

        titles = set()
        test_points: List[TestPoint] = []
        self._traverse_topics(root_topic, [], titles, test_points)

        document_type = "non_modeling" if "功能" in titles else "modeling"

        # 解析阶段保留所有测试点（包括 priority=3）
        # 生成阶段是否参与生成由后续逻辑控制
        for idx, point in enumerate(test_points, start=1):
            point.point_id = f"TP{idx:03d}"

        stats = self._build_stats(test_points)
        basic_info = self._extract_basic_info(root_topic)

        parse_id = uuid4().hex
        return ParsedXmindDocument(
            parse_id=parse_id,
            requirement_name=requirement_name,
            document_type=document_type,
            document_number=document_number,
            customer=basic_info.get("customer"),
            product=basic_info.get("product"),
            channel=basic_info.get("channel"),
            partner=basic_info.get("partner"),
            designer=basic_info.get("designer"),
            test_points=test_points,
            stats=stats
        )

    def _read_content_xml(self) -> str:
        try:
            with zipfile.ZipFile(self.xmind_path, "r") as zf:
                with zf.open("content.xml") as content:
                    return content.read().decode("utf-8")
        except KeyError:
            raise ValueError("XMind文件缺少content.xml")
        except Exception as exc:
            raise ValueError(f"XMind文件读取失败：{str(exc)}")

    def _get_title(self, topic: ET.Element) -> str:
        title_elem = topic.find("x:title", self._NS)
        if title_elem is None or title_elem.text is None:
            return ""
        return title_elem.text.strip()

    def _get_children(self, topic: ET.Element) -> List[ET.Element]:
        children = topic.find("x:children", self._NS)
        if children is None:
            return []
        topics = children.find("x:topics", self._NS)
        if topics is None:
            return []
        return topics.findall("x:topic", self._NS)

    def _traverse_topics(
        self,
        topic: ET.Element,
        path_titles: List[str],
        titles: set,
        points: List[TestPoint]
    ):
        title = self._get_title(topic)
        if title:
            titles.add(title)
        current_path = path_titles + ([title] if title else [])

        if title in ("业务流程", "业务规则"):
            for child in self._get_children(topic):
                child_title = self._get_title(child)
                if not child_title:
                    continue
                priority, cleaned_title = self._parse_priority(child, child_title)
                subtype = self._detect_subtype(cleaned_title)
                context = self._build_context(current_path)
                point_text = self._build_point_text(context, cleaned_title)
                point_type = "process" if title == "业务流程" else "rule"
                points.append(
                    TestPoint(
                        point_id="",
                        point_type=point_type,
                        subtype=subtype,
                        priority=priority,
                        text=point_text,
                        context=context
                    )
                )
            return

        for child in self._get_children(topic):
            self._traverse_topics(child, current_path, titles, points)

    def _build_context(self, path_titles: List[str]) -> str:
        context_parts = [t for t in path_titles if t and t not in ("基础信息",)]
        return " / ".join(context_parts).strip()

    def _build_point_text(self, context: str, title: str) -> str:
        if context:
            return f"{context} - {title}"
        return title

    def _parse_root_title(self, title: str) -> Tuple[Optional[str], str]:
        if not title:
            return None, ""
        if "-" in title:
            left, right = title.split("-", 1)
            return left.strip() or None, right.strip()
        return None, title.strip()

    def _extract_basic_info(self, root_topic: ET.Element) -> dict:
        info = {
            "customer": None,
            "product": None,
            "channel": None,
            "partner": None,
            "designer": None
        }
        for child in self._get_children(root_topic):
            if self._get_title(child) != "基础信息":
                continue
            for item in self._get_children(child):
                title = self._get_title(item)
                if not title:
                    continue
                value = ""
                if "：" in title:
                    _, value = title.split("：", 1)
                elif ":" in title:
                    _, value = title.split(":", 1)
                value = value.strip()
                if title.startswith("客户"):
                    info["customer"] = value or None
                elif title.startswith("产品"):
                    info["product"] = value or None
                elif title.startswith("渠道"):
                    info["channel"] = value or None
                elif title.startswith("合作方"):
                    info["partner"] = value or None
                elif title.startswith("设计者"):
                    info["designer"] = value or None
            break
        return info

    def _parse_priority(self, topic: ET.Element, title: str) -> Tuple[Optional[int], str]:
        marker_priority = self._get_marker_priority(topic)
        if marker_priority in (1, 2, 3):
            return marker_priority, title.strip()

        pattern = r"^\s*[（(]?([123])[)）).、]\s*"
        match = re.match(pattern, title)
        if match:
            priority = int(match.group(1))
            cleaned = re.sub(pattern, "", title).strip()
            return priority, cleaned

        pattern_simple = r"^\s*([123])[\.\、]\s*"
        match = re.match(pattern_simple, title)
        if match:
            priority = int(match.group(1))
            cleaned = re.sub(pattern_simple, "", title).strip()
            return priority, cleaned

        return None, title.strip()

    def _get_marker_priority(self, topic: ET.Element) -> Optional[int]:
        # XMind8 Update9: marker-ref 可能在 topic 子节点中嵌套
        for elem in topic.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if tag != "marker-ref":
                continue
            marker_id = elem.attrib.get("marker-id", "")
            match = re.match(r"priority-(\d)", marker_id)
            if match:
                return int(match.group(1))

        # 兼容 markers 属性（逗号分隔）
        marker_attr = topic.attrib.get("markers", "")
        if marker_attr:
            for marker_id in [m.strip() for m in marker_attr.split(",") if m.strip()]:
                match = re.match(r"priority-(\d)", marker_id)
                if match:
                    return int(match.group(1))

        return None

    def _detect_subtype(self, text: str) -> Optional[str]:
        positive_keywords = ["通过", "成功", "正确", "一致", "正常"]
        negative_keywords = ["不通过", "失败", "错误", "不一致", "异常", "提示"]

        last_pos = -1
        subtype = None
        for word in positive_keywords:
            idx = text.rfind(word)
            if idx > last_pos:
                last_pos = idx
                subtype = "positive"
        for word in negative_keywords:
            idx = text.rfind(word)
            if idx > last_pos:
                last_pos = idx
                subtype = "negative"

        return subtype

    def _build_stats(self, points: List[TestPoint]) -> dict:
        stats = {
            "total": len(points),
            "by_type": {"process": 0, "rule": 0},
            "by_priority": {"1": 0, "2": 0, "3": 0, "unknown": 0},
            "by_subtype": {"positive": 0, "negative": 0, "unknown": 0}
        }

        for point in points:
            stats["by_type"][point.point_type] = stats["by_type"].get(point.point_type, 0) + 1
            if point.priority in (1, 2, 3):
                stats["by_priority"][str(point.priority)] += 1
            else:
                stats["by_priority"]["unknown"] += 1
            if point.subtype in ("positive", "negative"):
                stats["by_subtype"][point.subtype] += 1
            else:
                stats["by_subtype"]["unknown"] += 1

        return stats
