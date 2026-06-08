#!/usr/bin/env python3
"""Generate SAN offline impact report from NX-OS command outputs.

Inputs:
  - wwn_with_vsan.txt
  - Fabric A Command Result/show_fcalias.txt
  - Fabric A Command Result/show_zoneset_active.txt
  - Fabric B Command Result/show_fcalias.txt
  - Fabric B Command Result/show_zoneset_active.txt

Outputs:
  - 输出结果.xlsx
  - 编码思路.txt
"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parent
TARGET_FILE = ROOT / "wwn_with_vsan.txt"
OUTPUT_XLSX = ROOT / "输出结果.xlsx"
IDEA_FILE = ROOT / "编码思路.txt"

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

FCALIAS_HEADER_RE = re.compile(r"^fcalias\s+name\s+(\S+)\s+vsan\s+(\d+)", re.I)
ZONE_HEADER_RE = re.compile(r"^\s*zone\s+name\s+(\S+)\s+vsan\s+(\d+)", re.I)
PWWN_LINE_RE = re.compile(r"^\s*pwwn\s+([0-9a-f:]{23})(?:\s+\[([^\]]+)\])?", re.I)
FCID_PWWN_RE = re.compile(r"\[pwwn\s+([0-9a-f:]{23})\](?:\s+\[([^\]]+)\])?", re.I)


@dataclass
class Member:
    raw: str
    wwn: str | None = None
    alias: str | None = None


@dataclass
class Zone:
    fabric: str
    vsan: str
    name: str
    members: list[Member] = field(default_factory=list)


def read_text(path: Path) -> str:
    """Read text files that may be UTF-8 or GBK encoded."""
    for encoding in ("utf-8", "gb18030", "gbk"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def normalize_wwn(value: str) -> str:
    return value.strip().lower()


def load_targets(path: Path) -> list[tuple[str, str]]:
    targets: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for line in read_text(path).splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "_" not in line:
            continue
        wwn, vsan = line.rsplit("_", 1)
        item = (normalize_wwn(wwn), vsan.strip())
        if item not in seen:
            seen.add(item)
            targets.append(item)
    return targets


def parse_fcalias(path: Path) -> dict[tuple[str, str], list[str]]:
    alias_by_wwn: dict[tuple[str, str], list[str]] = {}
    current_alias: str | None = None
    current_vsan: str | None = None
    for line in read_text(path).splitlines():
        header = FCALIAS_HEADER_RE.match(line.strip())
        if header:
            current_alias, current_vsan = header.groups()
            continue
        if current_alias and current_vsan:
            pwwn = PWWN_LINE_RE.match(line)
            if pwwn:
                key = (normalize_wwn(pwwn.group(1)), current_vsan)
                alias_by_wwn.setdefault(key, [])
                if current_alias not in alias_by_wwn[key]:
                    alias_by_wwn[key].append(current_alias)
    return alias_by_wwn


def parse_member(line: str) -> Member | None:
    pwwn = PWWN_LINE_RE.match(line)
    if pwwn:
        return Member(raw=line.rstrip(), wwn=normalize_wwn(pwwn.group(1)), alias=pwwn.group(2))
    fcid = FCID_PWWN_RE.search(line)
    if fcid:
        return Member(raw=line.rstrip(), wwn=normalize_wwn(fcid.group(1)), alias=fcid.group(2))
    return None


def deduplicate_zones(zones: list[Zone]) -> list[Zone]:
    """Drop repeated zone blocks caused by duplicated command output sections."""
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
            zone_name, vsan = header.groups()
            current = Zone(fabric=fabric, vsan=vsan, name=zone_name)
            continue
        if current:
            member = parse_member(line)
            if member:
                current.members.append(member)
    if current:
        zones.append(current)
    return deduplicate_zones(zones)


def build_rows() -> tuple[list[list[str]], set[int], dict[str, int]]:
    targets = load_targets(TARGET_FILE)
    target_set = set(targets)

    aliases: dict[tuple[str, str], list[str]] = {}
    zones: list[Zone] = []
    for fabric, paths in FABRICS.items():
        aliases.update(parse_fcalias(paths["fcalias"]))
        zones.extend(parse_zones(paths["zoneset"], fabric))

    zones_by_target: dict[tuple[str, str], list[Zone]] = {target: [] for target in targets}
    zone_alias_by_target: dict[tuple[str, str], list[str]] = {target: [] for target in targets}
    for zone in zones:
        for member in zone.members:
            if member.wwn:
                key = (member.wwn, zone.vsan)
                if key in zones_by_target:
                    zones_by_target[key].append(zone)
                    if member.alias and member.alias not in zone_alias_by_target[key]:
                        zone_alias_by_target[key].append(member.alias)

    rows = [["Fabric", "VSAN", "WWN", "Fcalias", "Zone", "Zone Member", "说明"]]
    yellow_rows: set[int] = set()
    stats = {"targets": len(targets), "matched_targets": 0, "zones": 0, "yellow_members": 0}

    for wwn, vsan in targets:
        target_aliases = aliases.get((wwn, vsan), []) or zone_alias_by_target.get((wwn, vsan), [])
        alias_text = ", ".join(target_aliases)
        fabric = "Fabric A" if vsan == "5" else "Fabric B" if vsan == "6" else ""
        target_zones = zones_by_target.get((wwn, vsan), [])
        if target_zones:
            stats["matched_targets"] += 1
        if not target_zones:
            rows.append([fabric, vsan, wwn, alias_text, "未在zoneset active中找到", "", ""])
            continue
        first_target_row = True
        for zone in target_zones:
            stats["zones"] += 1
            first_member = True
            for member in zone.members or [Member(raw="")]:
                member_is_external = bool(member.wwn and (member.wwn, zone.vsan) not in target_set)
                note = "本次不在提交的WWN列表中，但是和本次下线的WWN有Zone业务" if member_is_external else ""
                row = [
                    fabric if first_target_row and first_member else "",
                    vsan if first_target_row and first_member else "",
                    wwn if first_target_row and first_member else "",
                    alias_text if first_target_row and first_member else "",
                    zone.name if first_member else "",
                    member.raw,
                    note,
                ]
                rows.append(row)
                excel_row_number = len(rows)
                if member_is_external:
                    yellow_rows.add(excel_row_number)
                    stats["yellow_members"] += 1
                first_member = False
            first_target_row = False
    return rows, yellow_rows, stats


def column_name(index: int) -> str:
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def write_xlsx(path: Path, rows: list[list[str]], yellow_rows: set[int]) -> None:
    # Minimal XLSX writer using inline strings, enough for this report and no third-party dependencies.
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    sheet_rows = []
    for r_idx, row in enumerate(rows, start=1):
        cells = []
        for c_idx, value in enumerate(row, start=1):
            ref = f"{column_name(c_idx)}{r_idx}"
            style = "1" if r_idx == 1 else "2" if r_idx in yellow_rows else "0"
            text = escape(str(value))
            cells.append(f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{text}</t></is></c>')
        sheet_rows.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
    dimension = f"A1:{column_name(max(len(r) for r in rows))}{len(rows)}"
    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dimension}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="2" width="8" customWidth="1"/><col min="3" max="3" width="26" customWidth="1"/><col min="4" max="4" width="36" customWidth="1"/><col min="5" max="5" width="70" customWidth="1"/><col min="6" max="6" width="72" customWidth="1"/><col min="7" max="7" width="58" customWidth="1"/></cols>
  <sheetData>{''.join(sheet_rows)}</sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>'''
    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'''
    workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="下线影响分析" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>'''
    app = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application></Properties>'''
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("xl/styles.xml", styles_xml)
        zf.writestr("docProps/core.xml", core)
        zf.writestr("docProps/app.xml", app)


def write_idea_file(path: Path, stats: dict[str, int]) -> None:
    content = f"""编码思路记录

1. 读取 wwn_with_vsan.txt，将每行按最后一个下划线拆分为 WWN 和 VSAN，并统一把 WWN 转为小写；VSAN 5 归入 Fabric A，VSAN 6 归入 Fabric B。
2. 分别解析两个 Fabric 的 show_fcalias.txt：识别 `fcalias name <别名> vsan <编号>` 块，以及块内的 `pwwn <WWN>`，建立 `(WWN, VSAN) -> fcalias` 映射。
3. 分别解析两个 Fabric 的 show_zoneset_active.txt：识别 `zone name <Zone> vsan <编号>`，并收集 Zone 内直接 `pwwn ...` 以及 active 输出中 `fcid ... [pwwn ...]` 形式的成员。
4. 对解析出的 Zone 块按 Fabric、VSAN、Zone 名称和完整成员列表去重，避免同一命令输出重复粘贴时导致结果重复。
5. 遍历下线清单中的每一个 `(WWN, VSAN)`，查找其 fcalias、所在 Zone，以及该 Zone 的当前全部成员。
6. 输出 Excel 时，每个目标 WWN 按 Zone 展开；Zone 成员中如果带有 WWN 且该 `(成员WWN, VSAN)` 不在本次下线清单中，则整行标黄，并在“说明”列写入“本次不在提交的WWN列表中，但是和本次下线的WWN有Zone业务”。
7. 对未在 zoneset active 中找到的下线 WWN，在 Zone 列标注“未在zoneset active中找到”，便于后续人工复核。

本次统计
- 下线清单 WWN 数量：{stats['targets']}
- 匹配到 active Zone 的 WWN 数量：{stats['matched_targets']}
- 输出的目标 WWN/Zone 关联数量：{stats['zones']}
- 标黄的非本次下线 Zone Member 行数：{stats['yellow_members']}
"""
    path.write_text(content, encoding="utf-8")


def main() -> None:
    rows, yellow_rows, stats = build_rows()
    write_xlsx(OUTPUT_XLSX, rows, yellow_rows)
    write_idea_file(IDEA_FILE, stats)
    print(f"Wrote {OUTPUT_XLSX.name}: {len(rows) - 1} data rows")
    print(f"Wrote {IDEA_FILE.name}")
    print(stats)


if __name__ == "__main__":
    main()
