#!/usr/bin/env python3
"""Validate the generated SAN offline report against source command outputs."""
from __future__ import annotations

import re
import sys
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
REPORT = ROOT / "输出结果.xlsx"
TARGET_FILE = ROOT / "wwn_with_vsan.txt"
FABRICS = {
    "Fabric A": {
        "vsan": "5",
        "fcalias": ROOT / "Fabric A Command Result" / "show_fcalias.txt",
        "zoneset": ROOT / "Fabric A Command Result" / "show_zoneset_active.txt",
    },
    "Fabric B": {
        "vsan": "6",
        "fcalias": ROOT / "Fabric B Command Result" / "show_fcalias.txt",
        "zoneset": ROOT / "Fabric B Command Result" / "show_zoneset_active.txt",
    },
}

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
EXPECTED_HEADER = ["Fabric", "VSAN", "WWN", "Fcalias", "Zone", "Zone Member", "说明"]
EXTERNAL_NOTE = "本次不在提交的WWN列表中，但是和本次下线的WWN有Zone业务"
WWN_RE = re.compile(r"^[0-9a-f]{2}(?::[0-9a-f]{2}){7}$", re.I)
FCALIAS_HEADER_RE = re.compile(r"^fcalias\s+name\s+(\S+)\s+vsan\s+(\d+)", re.I)
ZONE_HEADER_RE = re.compile(r"^\s*zone\s+name\s+(\S+)\s+vsan\s+(\d+)", re.I)
PWWN_LINE_RE = re.compile(r"^\s*pwwn\s+([0-9a-f:]{23})(?:\s+\[([^\]]+)\])?", re.I)
FCID_PWWN_RE = re.compile(r"\[pwwn\s+([0-9a-f:]{23})\](?:\s+\[([^\]]+)\])?", re.I)
CELL_REF_RE = re.compile(r"([A-Z]+)(\d+)")


@dataclass(frozen=True)
class Member:
    raw: str
    wwn: str | None = None


@dataclass
class Zone:
    fabric: str
    vsan: str
    name: str
    members: list[Member] = field(default_factory=list)


def read_text(path: Path) -> str:
    for encoding in ("utf-8", "gb18030", "gbk"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def normalize_wwn(wwn: str) -> str:
    return wwn.strip().lower()


def load_targets() -> list[tuple[str, str]]:
    targets: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for lineno, line in enumerate(read_text(TARGET_FILE).splitlines(), start=1):
        line = line.strip()
        if not line:
            continue
        try:
            wwn, vsan = line.rsplit("_", 1)
        except ValueError as exc:
            raise AssertionError(f"下线清单第 {lineno} 行缺少下划线分隔符: {line}") from exc
        wwn = normalize_wwn(wwn)
        assert WWN_RE.match(wwn), f"下线清单第 {lineno} 行 WWN 格式错误: {wwn}"
        assert vsan in {"5", "6"}, f"下线清单第 {lineno} 行 VSAN 不是 5/6: {vsan}"
        item = (wwn, vsan)
        assert item not in seen, f"下线清单存在重复项: {line}"
        seen.add(item)
        targets.append(item)
    return targets


def parse_fcalias(path: Path) -> dict[tuple[str, str], list[str]]:
    aliases: dict[tuple[str, str], list[str]] = defaultdict(list)
    current_alias = None
    current_vsan = None
    for line in read_text(path).splitlines():
        header = FCALIAS_HEADER_RE.match(line.strip())
        if header:
            current_alias, current_vsan = header.groups()
            continue
        member = PWWN_LINE_RE.match(line)
        if current_alias and current_vsan and member:
            key = (normalize_wwn(member.group(1)), current_vsan)
            if current_alias not in aliases[key]:
                aliases[key].append(current_alias)
    return dict(aliases)


def parse_member(line: str) -> Member | None:
    member = PWWN_LINE_RE.match(line)
    if member:
        return Member(raw=line.rstrip(), wwn=normalize_wwn(member.group(1)))
    member = FCID_PWWN_RE.search(line)
    if member:
        return Member(raw=line.rstrip(), wwn=normalize_wwn(member.group(1)))
    return None


def deduplicate_zones(zones: list[Zone]) -> list[Zone]:
    unique_zones: list[Zone] = []
    seen: set[tuple[str, str, str, tuple[str, ...]]] = set()
    for zone in zones:
        key = (zone.fabric, zone.vsan, zone.name, tuple(member.raw for member in zone.members))
        if key in seen:
            continue
        seen.add(key)
        unique_zones.append(zone)
    return unique_zones


def parse_zones(path: Path, fabric: str) -> list[Zone]:
    zones: list[Zone] = []
    current: Zone | None = None
    for line in read_text(path).splitlines():
        header = ZONE_HEADER_RE.match(line)
        if header:
            if current:
                zones.append(current)
            name, vsan = header.groups()
            current = Zone(fabric=fabric, vsan=vsan, name=name)
            continue
        if current:
            member = parse_member(line)
            if member:
                current.members.append(member)
    if current:
        zones.append(current)
    return deduplicate_zones(zones)


def cell_col(cell_ref: str) -> int:
    letters = CELL_REF_RE.match(cell_ref).group(1)  # type: ignore[union-attr]
    value = 0
    for char in letters:
        value = value * 26 + ord(char) - 64
    return value


def read_report_rows() -> list[list[tuple[str, str]]]:
    with zipfile.ZipFile(REPORT) as zf:
        assert zf.testzip() is None, "输出结果.xlsx ZIP 完整性校验失败"
        worksheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
    rows: list[list[tuple[str, str]]] = []
    for row in worksheet.findall(".//x:row", NS):
        values = [('', '') for _ in range(7)]
        for cell in row.findall("x:c", NS):
            col = cell_col(cell.attrib["r"])
            if col > 7:
                continue
            text_node = cell.find("x:is/x:t", NS)
            text = text_node.text if text_node is not None and text_node.text is not None else ""
            values[col - 1] = (text, cell.attrib.get("s", "0"))
        rows.append(values)
    return rows


def flatten_report(rows: list[list[tuple[str, str]]]) -> dict[tuple[str, str, str], list[tuple[str, str, str]]]:
    flattened: dict[tuple[str, str, str], list[tuple[str, str, str]]] = defaultdict(list)
    current_wwn = current_vsan = current_zone = ""
    for row in rows[1:]:
        texts = [cell[0] for cell in row]
        styles = [cell[1] for cell in row]
        fabric, vsan, wwn, _alias, zone, member, note = texts
        if wwn:
            current_wwn, current_vsan = wwn, vsan
            expected_fabric = "Fabric A" if vsan == "5" else "Fabric B"
            assert fabric == expected_fabric, f"{wwn}_{vsan} Fabric 映射错误: {fabric}"
        if zone:
            current_zone = zone
        assert current_wwn and current_vsan and current_zone, f"报告存在无法归属上下文的行: {texts}"
        if note:
            assert note == EXTERNAL_NOTE, f"说明列内容不符合预期: {note}"
            assert all(style == "2" for style in styles), f"标黄说明行没有整行使用黄色样式: {texts}"
        else:
            assert all(style != "2" for style in styles), f"非外部成员行不应标黄: {texts}"
        flattened[(current_wwn, current_vsan, current_zone)].append((member, note, styles[5]))
    return dict(flattened)


def main() -> int:
    targets = load_targets()
    target_set = set(targets)
    aliases: dict[tuple[str, str], list[str]] = {}
    zones: list[Zone] = []
    for fabric, paths in FABRICS.items():
        aliases.update(parse_fcalias(paths["fcalias"]))
        zones.extend(parse_zones(paths["zoneset"], fabric))

    expected: dict[tuple[str, str, str], list[tuple[str, str]]] = {}
    matched_targets: set[tuple[str, str]] = set()
    expected_external_rows = 0
    for zone in zones:
        for member in zone.members:
            target = (member.wwn or "", zone.vsan)
            if target not in target_set:
                continue
            matched_targets.add(target)
            key = (target[0], target[1], zone.name)
            expected[key] = []
            for zone_member in zone.members:
                is_external = bool(zone_member.wwn and (zone_member.wwn, zone.vsan) not in target_set)
                note = EXTERNAL_NOTE if is_external else ""
                if is_external:
                    expected_external_rows += 1
                expected[key].append((zone_member.raw, note))

    rows = read_report_rows()
    header = [cell[0] for cell in rows[0]]
    assert header == EXPECTED_HEADER, f"表头不符合预期: {header}"
    actual = flatten_report(rows)

    assert set(actual) == set(expected), (
        f"目标WWN/VSAN/Zone 组合不一致: 缺失 {len(set(expected) - set(actual))} 个, "
        f"多出 {len(set(actual) - set(expected))} 个"
    )
    for key, expected_members in expected.items():
        actual_members = [(member, note) for member, note, _style in actual[key]]
        assert actual_members == expected_members, f"{key} 的 Zone Member 列表与源文件不一致"

    actual_external_rows = sum(1 for members in actual.values() for _member, note, _style in members if note)
    assert actual_external_rows == expected_external_rows, "标黄外部成员行数与源文件计算结果不一致"

    print("测试通过：输出结果.xlsx 与源文件解析结果一致")
    print(f"- 下线清单 WWN 数量：{len(targets)}")
    print(f"- 匹配到 active Zone 的 WWN 数量：{len(matched_targets)}")
    print(f"- 目标 WWN/VSAN/Zone 组合数量：{len(expected)}")
    print(f"- Excel 数据行数：{len(rows) - 1}")
    print(f"- 标黄非本次下线 Zone Member 行数：{actual_external_rows}")
    print("- 表头、Fabric/VSAN 映射、Zone Member 明细、说明列和黄色样式均校验通过")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except AssertionError as exc:
        print(f"测试失败：{exc}", file=sys.stderr)
        raise SystemExit(1)
