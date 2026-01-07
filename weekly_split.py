#!/usr/bin/env python3
"""
Generate weekly Excel summaries split by project and task.

Reads `時間軌跡.csv`, groups entries by Monday–Sunday weeks, aggregates total
minutes for each project (`項目名稱`) and task (`任務名稱`), then writes the
results to `total_review.xlsx` with four worksheets:

1. `projects_long`  — long-form weekly totals for each project
2. `tasks_long`     — long-form weekly totals for each task
3. `projects_wide`  — weekly rows with projects as columns (minutes)
4. `tasks_wide`     — weekly rows with tasks as columns (minutes)

The workbook is authored without third-party libraries for portability.
"""

from __future__ import annotations

import csv
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple
from xml.sax.saxutils import escape
import zipfile


INPUT_FILE = Path("時間軌跡.csv")
OUTPUT_FILE = Path("total_review.xlsx")


@dataclass
class WeeklyBucket:
    start: date
    end: date
    by_project: Dict[str, int] = field(default_factory=lambda: defaultdict(int))
    by_task: Dict[str, int] = field(default_factory=lambda: defaultdict(int))

    @property
    def label(self) -> str:
        return f"{self.start:%Y/%m/%d} ~ {self.end:%Y/%m/%d}"


def iter_rows(path: Path) -> Iterable[dict]:
    """Yield rows from the CSV, automatically stripping UTF-8 BOM if present."""
    if not path.exists():
        raise FileNotFoundError(f"找不到檔案 {path}")

    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            yield row


def parse_minutes(value: str) -> int:
    try:
        return int(value)
    except ValueError as exc:
        raise ValueError(f"無法解析持續時間（分鐘）數值：{value!r}") from exc


def week_bounds(day: date) -> tuple[date, date]:
    """Return the Monday (inclusive) and Sunday (inclusive) for the given day."""
    start = day - timedelta(days=day.weekday())
    end = start + timedelta(days=6)
    return start, end


def build_buckets(rows: Iterable[dict]) -> Dict[date, WeeklyBucket]:
    buckets: Dict[date, WeeklyBucket] = {}

    for row in rows:
        try:
            day = datetime.strptime(row["開始日期"], "%Y/%m/%d").date()
        except (KeyError, ValueError) as exc:
            raise ValueError(f"無法解析開始日期：{row}") from exc

        minutes = parse_minutes(row["持續時間（分鐘）"])
        project_name = row.get("項目名稱", "").strip() or "未指定項目"
        task_name = row.get("任務名稱", "").strip() or "未指定任務"

        start, end = week_bounds(day)
        bucket = buckets.setdefault(start, WeeklyBucket(start=start, end=end))
        bucket.by_project[project_name] += minutes
        bucket.by_task[task_name] += minutes

    return buckets


def build_long_rows(
    buckets: Dict[date, WeeklyBucket]
) -> Tuple[List[List], List[List]]:
    project_rows: List[List] = []
    task_rows: List[List] = []

    for _, bucket in sorted(buckets.items()):
        for name, minutes in sorted(
            bucket.by_project.items(), key=lambda item: item[1], reverse=True
        ):
            hours = round(minutes / 60, 2)
            project_rows.append([bucket.label, name, minutes, hours])

        for name, minutes in sorted(
            bucket.by_task.items(), key=lambda item: item[1], reverse=True
        ):
            hours = round(minutes / 60, 2)
            task_rows.append([bucket.label, name, minutes, hours])

    return project_rows, task_rows


def build_wide_rows(
    buckets: Dict[date, WeeklyBucket]
) -> Tuple[Tuple[List[str], List[List]], Tuple[List[str], List[List]]]:
    sorted_buckets = [bucket for _, bucket in sorted(buckets.items())]

    project_names = sorted(
        {name for bucket in sorted_buckets for name in bucket.by_project}
    )
    task_names = sorted({name for bucket in sorted_buckets for name in bucket.by_task})

    project_header: List[str] = ["週期"] + project_names
    task_header: List[str] = ["週期"] + task_names

    project_rows: List[List] = []
    task_rows: List[List] = []

    for bucket in sorted_buckets:
        proj_row: List = [bucket.label] + [
            bucket.by_project.get(name, 0) for name in project_names
        ]
        task_row: List = [bucket.label] + [
            bucket.by_task.get(name, 0) for name in task_names
        ]
        project_rows.append(proj_row)
        task_rows.append(task_row)

    return (project_header, project_rows), (task_header, task_rows)


def column_letter(index: int) -> str:
    """Convert 1-based column index to Excel column letters."""
    letters: List[str] = []
    while index:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def format_cell(cell_ref: str, value) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return f'<c r="{cell_ref}"><v>{value}</v></c>'

    text = escape(str(value))
    return (
        f'<c r="{cell_ref}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'
    )


def sheet_xml(header: Sequence, rows: Sequence[Sequence]) -> str:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<sheetData>",
    ]

    all_rows = [header] + list(rows)
    for row_idx, row in enumerate(all_rows, start=1):
        cells = []
        for col_idx, value in enumerate(row, start=1):
            cell_ref = f"{column_letter(col_idx)}{row_idx}"
            cell_xml = format_cell(cell_ref, value)
            if cell_xml:
                cells.append(cell_xml)
        lines.append(f'<row r="{row_idx}">' + "".join(cells) + "</row>")

    lines.extend(["</sheetData>", "</worksheet>"])
    return "\n".join(lines)


def workbook_xml(sheet_names: Sequence[str]) -> str:
    sheets = [
        f'<sheet name="{escape(name)}" sheetId="{idx}" r:id="rId{idx}"/>'
        for idx, name in enumerate(sheet_names, start=1)
    ]
    return "\n".join(
        [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            "<sheets>",
            *sheets,
            "</sheets>",
            "</workbook>",
        ]
    )


def workbook_rels_xml(sheet_count: int) -> str:
    relationships = [
        (
            f'<Relationship Id="rId{idx}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{idx}.xml"/>'
        )
        for idx in range(1, sheet_count + 1)
    ]
    relationships.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    return "\n".join(
        [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            *relationships,
            "</Relationships>",
        ]
    )


def styles_xml() -> str:
    return "\n".join(
        [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>',
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>',
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
            "</styleSheet>",
        ]
    )


def content_types_xml(sheet_count: int) -> str:
    sheet_overrides = [
        (
            f'<Override PartName="/xl/worksheets/sheet{idx}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
        for idx in range(1, sheet_count + 1)
    ]
    return "\n".join(
        [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '<Default Extension="xml" ContentType="application/xml"/>',
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            *sheet_overrides,
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
            "</Types>",
        ]
    )


def rels_xml() -> str:
    return "\n".join(
        [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
            "</Relationships>",
        ]
    )


def app_xml(sheet_names: Sequence[str]) -> str:
    vector_items = [f"<vt:lpstr>{escape(name)}</vt:lpstr>" for name in sheet_names]
    return "\n".join(
        [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
            "<Application>Python</Application>",
            "<DocSecurity>0</DocSecurity>",
            "<ScaleCrop>false</ScaleCrop>",
            "<HeadingPairs>",
            '<vt:vector size="2" baseType="variant">',
            "<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>",
            f"<vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant>",
            "</vt:vector>",
            "</HeadingPairs>",
            "<TitlesOfParts>",
            f'<vt:vector size="{len(sheet_names)}" baseType="lpstr">',
            *vector_items,
            "</vt:vector>",
            "</TitlesOfParts>",
            "</Properties>",
        ]
    )


def core_xml() -> str:
    now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    return "\n".join(
        [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
            "<dc:creator>weekly_split</dc:creator>",
            "<cp:lastModifiedBy>weekly_split</cp:lastModifiedBy>",
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>',
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>',
            "</cp:coreProperties>",
        ]
    )


def write_xlsx(
    path: Path, sheets: Sequence[Tuple[str, Sequence, Sequence[Sequence]]]
) -> None:
    if not path.parent.exists():
        path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        sheet_names = [name for name, _, _ in sheets]

        for idx, (_, header, rows) in enumerate(sheets, start=1):
            xml_content = sheet_xml(header, rows)
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", xml_content)

        zf.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets)))
        zf.writestr("xl/styles.xml", styles_xml())
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheets)))
        zf.writestr("_rels/.rels", rels_xml())
        zf.writestr("docProps/app.xml", app_xml(sheet_names))
        zf.writestr("docProps/core.xml", core_xml())


def main() -> None:
    buckets = build_buckets(iter_rows(INPUT_FILE))

    if not buckets:
        print("資料為空，未建立任何輸出。")
        return

    project_long, task_long = build_long_rows(buckets)
    (project_wide_header, project_wide_rows), (task_wide_header, task_wide_rows) = (
        build_wide_rows(buckets)
    )

    sheets = [
        ("projects_long", ["週期", "項目名稱", "總分鐘", "總小時"], project_long),
        ("tasks_long", ["週期", "任務名稱", "總分鐘", "總小時"], task_long),
        ("projects_wide", project_wide_header, project_wide_rows),
        ("tasks_wide", task_wide_header, task_wide_rows),
    ]

    write_xlsx(OUTPUT_FILE, sheets)

    print(
        "已建立 total_review Excel："
        f"{OUTPUT_FILE}（projects_long {len(project_long)} 筆, tasks_long {len(task_long)} 筆）"
    )


if __name__ == "__main__":
    main()
