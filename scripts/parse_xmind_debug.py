"""
解析XMind内容并打印优先级标记与标题
用法：
  python scripts/parse_xmind_debug.py "path_to_xmind"
"""
import sys
import zipfile
import xml.etree.ElementTree as ET


NS = {"x": "urn:xmind:xmap:xmlns:content:2.0"}


def read_content_xml(xmind_path: str) -> str:
    with zipfile.ZipFile(xmind_path, "r") as zf:
        with zf.open("content.xml") as content:
            return content.read().decode("utf-8")


def get_title(topic: ET.Element) -> str:
    title_elem = topic.find("x:title", NS)
    if title_elem is None or title_elem.text is None:
        return ""
    return title_elem.text.strip()


def get_children(topic: ET.Element):
    children = topic.find("x:children", NS)
    if children is None:
        return []
    topics = children.find("x:topics", NS)
    if topics is None:
        return []
    return topics.findall("x:topic", NS)


def _local_name(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag


def get_marker_ids(topic: ET.Element):
    marker_ids = []
    for elem in topic.iter():
        if _local_name(elem.tag) == "marker-ref":
            marker_id = elem.attrib.get("marker-id", "")
            if marker_id:
                marker_ids.append(marker_id)
    # 兼容旧格式：markers属性可能直接挂在topic上
    marker_attr = topic.attrib.get("markers", "")
    if marker_attr:
        marker_ids.extend([m.strip() for m in marker_attr.split(",") if m.strip()])
    return marker_ids


def traverse(topic: ET.Element, path, stats):
    title = get_title(topic)
    current_path = path + ([title] if title else [])
    marker_ids = get_marker_ids(topic)
    stats["topics"] += 1
    if marker_ids:
        stats["marker_topics"] += 1
        for marker_id in marker_ids:
            stats["marker_ids"].add(marker_id)
        print("PATH:", " / ".join([p for p in current_path if p]))
        print("TITLE:", title)
        print("MARKERS:", marker_ids)
        print("-" * 60)

    for child in get_children(topic):
        traverse(child, current_path, stats)


def main():
    if len(sys.argv) < 2:
        print("请提供XMind文件路径")
        sys.exit(1)
    xmind_path = sys.argv[1]
    xml_text = read_content_xml(xmind_path)
    root = ET.fromstring(xml_text)
    sheet = root.find("x:sheet", NS)
    if sheet is None:
        print("未找到sheet节点")
        sys.exit(1)
    root_topic = sheet.find("x:topic", NS)
    if root_topic is None:
        print("未找到根主题节点")
        sys.exit(1)
    stats = {"topics": 0, "marker_topics": 0, "marker_ids": set()}
    traverse(root_topic, [], stats)
    print("TOTAL_TOPICS:", stats["topics"])
    print("MARKER_TOPICS:", stats["marker_topics"])
    print("MARKER_IDS:", sorted(stats["marker_ids"]))


if __name__ == "__main__":
    main()
