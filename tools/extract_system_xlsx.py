from __future__ import annotations

import json
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"a": MAIN_NS}


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for item in root.findall("a:si", NS):
        pieces = [node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t")]
        values.append("".join(pieces))
    return values


def read_rows(workbook_path: Path) -> list[list[str]]:
    with zipfile.ZipFile(workbook_path) as archive:
        shared = read_shared_strings(archive)
        sheet_xml = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        sheet_data = sheet_xml.find("a:sheetData", NS)
        if sheet_data is None:
            return []

        rows: list[list[str]] = []
        for row in sheet_data.findall("a:row", NS):
            current: list[str] = []
            for cell in row.findall("a:c", NS):
                cell_type = cell.attrib.get("t")
                value_node = cell.find("a:v", NS)
                if value_node is None:
                    current.append("")
                    continue
                raw_value = value_node.text or ""
                current.append(shared[int(raw_value)] if cell_type == "s" else raw_value)
            rows.append(current)
        return rows


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    workbook_path = root / "system.xlsx"
    output_path = root / "src" / "data" / "defaultWords.js"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    words = []
    for row in read_rows(workbook_path):
        if len(row) < 3:
            continue

        word_id, english, japanese = (part.strip() for part in row[:3])
        if not english or not japanese:
            continue

        words.append(
            {
                "id": word_id or str(len(words) + 1),
                "english": english,
                "japanese": japanese,
            }
        )

    module_text = (
        "export const defaultWords = "
        + json.dumps(words, ensure_ascii=False, indent=2)
        + ";\n\n"
        + "export const defaultDatasetMeta = "
        + json.dumps(
            {
                "id": "default",
                "name": "既定単語帳",
                "source": "system.xlsx",
                "count": len(words),
            },
            ensure_ascii=False,
            indent=2,
        )
        + ";\n"
    )
    output_path.write_text(module_text, encoding="utf-8")


if __name__ == "__main__":
    main()
