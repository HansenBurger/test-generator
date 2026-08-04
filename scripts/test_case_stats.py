"""用例统计解析回归验证脚本

验证点：
1. 规则简称（带"联系"标注节点）下的用例携带 rule_alias 归属标记，统计时与功能步骤分开
2. 规则简称下的 前提/步骤/预期 三层链不被截断
3. 优先级子->父传播：子节点（含前提/步骤/预期）标注优先级时以子节点为准，未标注才回退父节点
4. 父节点不窃取子孙节点的优先级 marker（兄弟节点优先级独立）
5. 并行多链（manual_case）优先级取链上最深标注
6. stats 新增 by_source（function_step / rule_alias）分开统计

用法：
  backend/.venv/bin/python scripts/test_case_stats.py
"""
import os
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

BACKEND_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "backend")
sys.path.insert(0, BACKEND_DIR)

from app.services.xmind_parser import XMindParser  # noqa: E402

NS = "urn:xmind:xmap:xmlns:content:2.0"
ET.register_namespace("", NS)


def topic(title, markers=None, children=None):
    elem = ET.Element(f"{{{NS}}}topic")
    t = ET.SubElement(elem, f"{{{NS}}}title")
    t.text = title
    if markers:
        refs = ET.SubElement(elem, f"{{{NS}}}marker-refs")
        for m in markers:
            ET.SubElement(refs, f"{{{NS}}}marker-ref", {"marker-id": m})
    if children:
        ch = ET.SubElement(elem, f"{{{NS}}}children")
        topics = ET.SubElement(ch, f"{{{NS}}}topics", {"type": "attached"})
        for c in children:
            topics.append(c)
    return elem


def build_fixture() -> str:
    root = topic("REQ001-用例统计验证需求", children=[
        topic("基础信息", children=[topic("客户：测试客户")]),
        topic("组件A", children=[
            topic("业务规则", children=[
                # 场景1/2：父 priority-2；子 priority-1 覆盖；另一子无标注则继承
                topic("录入金额", markers=["priority-2"], children=[
                    topic("金额合法通过", markers=["priority-1"]),
                    topic("金额为空提示"),
                ]),
                # 场景3：父无标注，子各自标注，父不得窃取
                topic("校验余额", children=[
                    topic("余额充足成功", markers=["priority-3"]),
                    topic("余额不足失败", markers=["priority-1"]),
                ]),
                # 场景4：规则简称（联系标注）+ 前提/步骤/预期链
                topic("R1", markers=["c_symbol_contact"], children=[
                    topic("已登录系统", markers=["priority-3"], children=[
                        topic("提交申请", children=[
                            topic("申请成功", markers=["priority-1"]),
                        ]),
                    ]),
                    topic("未登录系统", children=[
                        topic("提交申请被拦截", markers=["priority-2"], children=[
                            topic("提示先登录"),
                        ]),
                    ]),
                    topic("规则R1补充校验通过"),
                ]),
                # 场景5：自动化标记
                topic("自动复核", markers=["task-start"], children=[
                    topic("复核一致"),
                ]),
                # 场景6：并行多链（manual_case）
                topic("双链校验", children=[
                    topic("前提X", markers=["priority-3"], children=[
                        topic("步骤X", children=[
                            topic("预期X", markers=["priority-1"]),
                        ]),
                    ]),
                    topic("前提Y", children=[
                        topic("步骤Y", markers=["priority-2"], children=[
                            topic("预期Y"),
                        ]),
                    ]),
                ]),
            ]),
        ]),
    ])

    content = ET.Element(f"{{{NS}}}xmap-content", {"version": "2.0"})
    sheet = ET.SubElement(content, f"{{{NS}}}sheet", {"id": "s1"})
    sheet.append(root)
    st = ET.SubElement(sheet, f"{{{NS}}}title")
    st.text = "画布 1"

    xml_text = ET.tostring(content, encoding="unicode", xml_declaration=False)
    xml_text = '<?xml version="1.0" encoding="UTF-8" standalone="no"?>' + xml_text

    fd, path = tempfile.mkstemp(suffix=".xmind")
    os.close(fd)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("content.xml", xml_text)
    return path


def main() -> int:
    fixture = build_fixture()
    failures = []

    def check(name, cond, detail=""):
        print(f"[{'PASS' if cond else 'FAIL'}] {name} {detail}")
        if not cond:
            failures.append(name)

    try:
        doc = XMindParser(fixture).parse()
        pts = doc.test_points

        def find(kw):
            return [p for p in pts if kw in p.text]

        print(f"stats={doc.stats}\n")

        m = find("金额合法通过")
        check("子节点优先级覆盖父节点", len(m) == 1 and m[0].priority == 1, f"got={[x.priority for x in m]}")
        m = find("金额为空提示")
        check("子节点未标注继承父节点", len(m) == 1 and m[0].priority == 2, f"got={[x.priority for x in m]}")
        m = find("余额充足成功")
        check("兄弟节点优先级独立(充足=3)", len(m) == 1 and m[0].priority == 3, f"got={[x.priority for x in m]}")
        m = find("余额不足失败")
        check("兄弟节点优先级独立(不足=1)", len(m) == 1 and m[0].priority == 1, f"got={[x.priority for x in m]}")

        alias_pts = [p for p in pts if p.rule_alias == "R1"]
        check("R1 下 3 条用例标记 rule_alias", len(alias_pts) == 3, f"got={len(alias_pts)}")
        check("功能步骤用例 rule_alias 为空", all(p.rule_alias is None for p in pts if p.rule_alias != "R1"))
        m = find("申请成功")
        check("alias链式: 预期(1)覆盖前提(3) 且未截断",
              len(m) == 1 and m[0].priority == 1 and m[0].rule_alias == "R1",
              f"got={[(x.priority, x.text) for x in m]}")
        m = find("提示先登录")
        check("alias链式: 仅步骤标注(2) 且未截断", len(m) == 1 and m[0].priority == 2,
              f"got={[(x.priority, x.text) for x in m]}")
        m = find("规则R1补充校验通过")
        check("alias 下无标注用例 priority=None",
              len(m) == 1 and m[0].priority is None and m[0].rule_alias == "R1",
              f"got={[x.priority for x in m]}")

        m = find("复核一致")
        check("task-start 自动化标记", len(m) == 1 and m[0].is_automated)

        mx = [p for p in pts if p.manual_case and "双链校验" in p.text and p.preconditions == ["前提X"]]
        my = [p for p in pts if p.manual_case and "双链校验" in p.text and p.preconditions == ["前提Y"]]
        check("多链X: 预期(1)覆盖前提(3)", len(mx) == 1 and mx[0].priority == 1,
              f"got={[(x.priority, x.preconditions) for x in mx]}")
        check("多链Y: 仅步骤标注(2)", len(my) == 1 and my[0].priority == 2,
              f"got={[(x.priority, x.preconditions) for x in my]}")

        check("by_source 分开统计",
              doc.stats["by_source"] == {"function_step": len(pts) - 3, "rule_alias": 3},
              f"got={doc.stats['by_source']}")
        check("优先级分布",
              doc.stats["by_priority"] == {"1": 4, "2": 3, "3": 1, "unknown": 2},
              f"got={doc.stats['by_priority']}")

        print()
        print("FAILURES:", failures if failures else "无")
        return 1 if failures else 0
    finally:
        if os.path.exists(fixture):
            os.unlink(fixture)


if __name__ == "__main__":
    sys.exit(main())
