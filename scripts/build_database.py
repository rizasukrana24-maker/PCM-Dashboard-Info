import argparse
import calendar
import json
import re
import sqlite3
import subprocess
import zipfile
from pathlib import Path
from xml.etree import ElementTree


ROOT = Path(__file__).resolve().parent.parent
SCHEMA_PATH = ROOT / "database" / "schema.sql"
DEFAULT_DB_PATH = ROOT / "database" / "dashboard.sqlite"
DEFAULT_XLSX_PATH = Path(r"c:\RIZASKRN\PCM-IT\LEARNING\Gerakan\PTMS-Februari-2026.xlsx")

XML_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "office": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "package": "http://schemas.openxmlformats.org/package/2006/relationships",
}
MONTH_NAME_TO_NUMBER = {
    "januari": 1,
    "februari": 2,
    "maret": 3,
    "april": 4,
    "mei": 5,
    "juni": 6,
    "juli": 7,
    "agustus": 8,
    "september": 9,
    "oktober": 10,
    "november": 11,
    "desember": 12,
}
NUMBER_TO_MONTH_NAME = {value: key.title() for key, value in MONTH_NAME_TO_NUMBER.items()}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build atau append SQLite database dashboard dari file project atau workbook XLSX."
    )
    parser.add_argument(
        "--source",
        choices=("xlsx", "project"),
        default="xlsx",
        help="Sumber seed database. Default: xlsx",
    )
    parser.add_argument(
        "--mode",
        choices=("append", "rebuild"),
        default="append",
        help="append untuk tambah/update periode, rebuild untuk membuat ulang database dari nol.",
    )
    parser.add_argument(
        "--xlsx",
        nargs="+",
        default=[str(DEFAULT_XLSX_PATH)],
        help="Satu atau lebih path file XLSX saat source=xlsx.",
    )
    parser.add_argument(
        "--period",
        default="",
        help="Periode target dalam format YYYY-MM. Hanya untuk satu file XLSX.",
    )
    parser.add_argument(
        "--db",
        default=str(DEFAULT_DB_PATH),
        help="Path output file SQLite.",
    )
    return parser.parse_args()


def run_node(script: str) -> str:
    result = subprocess.run(
        ["node", "-e", script],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=True,
    )
    return result.stdout


def load_manual_dashboard_data() -> list[dict]:
    script = r"""
const fs = require("fs");
const text = fs.readFileSync("app.js", "utf8");
const match = text.match(/const manualDashboardData = (\[[\s\S]*?\n\]);/);
if (!match) {
  throw new Error("manualDashboardData tidak ditemukan di app.js");
}
const data = Function(`return (${match[1]})`)();
process.stdout.write(JSON.stringify(data));
"""
    return json.loads(run_node(script))


def load_recap_tables() -> list[dict]:
    script = r"""
const fs = require("fs");
const text = fs.readFileSync("excel-table-data.js", "utf8");
const data = Function("const window = {}; " + text + "; return window.excelRecapTablesData;")();
process.stdout.write(JSON.stringify(data));
"""
    return json.loads(run_node(script))


def column_ref_to_index(reference: str) -> int:
    letters = "".join(ch for ch in reference if ch.isalpha()).upper()
    index = 0
    for letter in letters:
        index = (index * 26) + (ord(letter) - ord("A") + 1)
    return index - 1


def normalize_row(values_by_index: dict[int, str]) -> list[str]:
    if not values_by_index:
        return []
    max_index = max(values_by_index)
    row = [""] * (max_index + 1)
    for index, value in values_by_index.items():
        row[index] = value
    while row and row[-1] == "":
        row.pop()
    return row


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    path = "xl/sharedStrings.xml"
    if path not in archive.namelist():
        return []

    root = ElementTree.fromstring(archive.read(path))
    return [
        "".join(
            text_node.text or ""
            for text_node in item.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
        )
        for item in root.findall("main:si", XML_NS)
    ]


def read_cell_value(cell: ElementTree.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    value_node = cell.find("main:v", XML_NS)
    inline_node = cell.find("main:is", XML_NS)

    if cell_type == "s" and value_node is not None and value_node.text is not None:
        return shared_strings[int(value_node.text)]

    if cell_type == "inlineStr" and inline_node is not None:
        return "".join(
            text_node.text or ""
            for text_node in inline_node.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
        )

    if value_node is not None and value_node.text is not None:
        return value_node.text

    return ""


def load_first_sheet_rows(xlsx_path: Path) -> dict[int, list[str]]:
    with zipfile.ZipFile(xlsx_path) as archive:
        shared_strings = read_shared_strings(archive)
        workbook_root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        relations_root = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        sheet_relations = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relations_root.findall("package:Relationship", XML_NS)
        }

        first_sheet = workbook_root.find("main:sheets/main:sheet", XML_NS)
        if first_sheet is None:
            raise ValueError("Workbook tidak memiliki sheet.")

        relation_id = first_sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = sheet_relations[relation_id]
        sheet_path = target if target.startswith("xl/") else f"xl/{target}"
        sheet_root = ElementTree.fromstring(archive.read(sheet_path))

    rows = {}
    for row_node in sheet_root.findall("main:sheetData/main:row", XML_NS):
        row_index = int(row_node.attrib.get("r", "0"))
        values_by_index = {}
        for cell_node in row_node.findall("main:c", XML_NS):
            reference = cell_node.attrib.get("r", "")
            column_index = column_ref_to_index(reference)
            values_by_index[column_index] = read_cell_value(cell_node, shared_strings).strip()
        row = normalize_row(values_by_index)
        if any(cell.strip() for cell in row):
            rows[row_index] = row
    return rows


def parse_int(value) -> int:
    digits = "".join(ch for ch in str(value) if ch.isdigit())
    return int(digits) if digits else 0


def derive_period_from_filename(xlsx_path: Path) -> str:
    normalized_name = xlsx_path.stem.lower()
    match = re.search(
        r"(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember)[\s_-]*(\d{4})",
        normalized_name,
    )
    if not match:
        raise ValueError(
            "Period tidak diberikan dan tidak bisa diturunkan dari nama file. Gunakan --period YYYY-MM."
        )

    month_number = MONTH_NAME_TO_NUMBER[match.group(1)]
    year = int(match.group(2))
    return f"{year:04d}-{month_number:02d}"


def month_label_from_period(period: str) -> str:
    _, month = period.split("-")
    return NUMBER_TO_MONTH_NAME[int(month)]


def category_from_title(title: str) -> str:
    return "luar_negeri" if "Luar Negeri" in title else "dalam_negeri"


def extract_dashboard_rows_from_workbook(rows: dict[int, list[str]], period: str) -> list[dict]:
    header_row = rows.get(14) or rows.get(8) or rows.get(2)
    if not header_row:
        raise ValueError("Header gerakan harian tidak ditemukan di workbook.")

    day_numbers = [parse_int(item) for item in header_row[1:] if item and item not in {"Total", "Rata-Rata"}]
    if not day_numbers:
        raise ValueError("Kolom tanggal harian tidak ditemukan di workbook.")

    south_domestic = rows.get(9)
    south_foreign = rows.get(10)
    south_total = rows.get(11)
    north_domestic = rows.get(15)
    north_foreign = rows.get(16)
    north_total = rows.get(17)

    required_rows = [
        south_domestic,
        south_foreign,
        south_total,
        north_domestic,
        north_foreign,
        north_total,
    ]
    if any(not row for row in required_rows):
        raise ValueError("Blok gerakan utara/selatan tidak lengkap di workbook.")

    max_day = calendar.monthrange(*map(int, period.split("-")))[1]
    valid_days = [
        day
        for day in day_numbers
        if 1 <= day <= max_day
        and all(len(row) > day for row in required_rows)
        and any(str(row[day]).strip() for row in required_rows)
    ]

    output = []
    for day in valid_days:
        output.append(
            {
                "tanggal": f"{period}-{day:02d}",
                "kantor": "utara",
                "idr": parse_int(north_domestic[day]),
                "usd": parse_int(north_foreign[day]),
                "total": parse_int(north_total[day]),
            }
        )
        output.append(
            {
                "tanggal": f"{period}-{day:02d}",
                "kantor": "selatan",
                "idr": parse_int(south_domestic[day]),
                "usd": parse_int(south_foreign[day]),
                "total": parse_int(south_total[day]),
            }
        )

    return output


def iter_report_blocks(rows: dict[int, list[str]]) -> list[tuple[str, list[str], list[list[str]]]]:
    ordered_row_numbers = sorted(rows)
    title_row_numbers = [row_number for row_number in ordered_row_numbers if rows[row_number][0].startswith("Rekapitulasi")]
    blocks = []

    for index, title_row_number in enumerate(title_row_numbers):
        next_title_row = title_row_numbers[index + 1] if index + 1 < len(title_row_numbers) else None
        header_row_number = title_row_number + 1
        data_start = title_row_number + 2

        title = rows[title_row_number][0]
        headers = rows.get(header_row_number, [])
        data_rows = []
        for row_number in ordered_row_numbers:
            if row_number < data_start:
                continue
            if next_title_row is not None and row_number >= next_title_row:
                break
            data_rows.append(rows[row_number])

        blocks.append((title, headers, data_rows))

    return blocks


def select_agent_rows_for_period(title: str, headers: list[str], data_rows: list[list[str]], period: str) -> dict:
    month_label = month_label_from_period(period)
    movement_header = f"Gerakan - {month_label}"
    revenue_header = f"Est Pendapatan - {month_label}"
    movement_index = headers.index(movement_header)
    revenue_index = headers.index(revenue_header)
    pair_index = max(0, (movement_index - 2) // 2)

    output_rows = []
    for row in data_rows:
        if not row:
            continue

        name = row[1].strip() if len(row) > 1 else ""
        movement_value = row[movement_index].strip() if len(row) > movement_index else ""
        revenue_value = row[revenue_index].strip() if len(row) > revenue_index else ""

        if name:
            output_rows.append([row[0].strip(), name, movement_value or "0", revenue_value or "0"])
            continue

        if len(headers) <= 4:
            total_movement = row[movement_index].strip() if len(row) > movement_index else ""
            total_revenue = row[revenue_index].strip() if len(row) > revenue_index else ""
            output_rows.append(["TOTAL", "", total_movement or "0", total_revenue or "0"])
            continue

        if len(row) > revenue_index and (row[movement_index].strip() or row[revenue_index].strip()):
            total_movement = row[movement_index].strip()
            total_revenue = row[revenue_index].strip()
            output_rows.append(["TOTAL", "", total_movement or "0", total_revenue or "0"])
            continue

        if len(row) >= (pair_index * 2) + 2:
            total_movement = row[pair_index * 2].strip()
            total_revenue = row[(pair_index * 2) + 1].strip()
            output_rows.append(["TOTAL", "", total_movement or "0", total_revenue or "0"])

    return {
        "title": f"{title.split(' Data PTMS ')[0]} {month_label} {period[:4]}",
        "headers": ["No", "Agen", movement_header, revenue_header],
        "rows": output_rows,
    }


def select_jetty_rows_for_period(title: str, headers: list[str], data_rows: list[list[str]], period: str) -> dict:
    month_label = month_label_from_period(period)
    month_headers = headers[3:]
    month_index = month_headers.index(month_label)
    data_index = headers.index(month_label)
    compact_total_index = data_index - 3

    output_rows = []
    for row in data_rows:
        if not row:
            continue

        jetty_name = row[1].strip() if len(row) > 1 else ""
        if jetty_name:
            region = row[2].strip() if len(row) > 2 else ""
            movements = row[data_index].strip() if len(row) > data_index else ""
            output_rows.append([row[0].strip(), jetty_name, region, movements or "0"])
            continue

        if len(headers) <= 4:
            total_movement = row[data_index].strip() if len(row) > data_index else ""
            output_rows.append(["TOTAL", "", "", total_movement or "0"])
            continue

        if len(row) > data_index and row[data_index].strip():
            total_movement = row[data_index].strip()
            output_rows.append(["TOTAL", "", "", total_movement or "0"])
            continue

        if len(row) > compact_total_index:
            total_movement = row[compact_total_index].strip()
            output_rows.append(["TOTAL", "", "", total_movement or "0"])

    return {
        "title": f"{title.split(' Data PTMS ')[0]} {month_label} {period[:4]}",
        "headers": ["No", "Jetty", "Wilayah", month_label],
        "rows": output_rows,
    }


def extract_recap_sections_from_workbook(rows: dict[int, list[str]], period: str) -> list[dict]:
    sections = []
    for title, headers, data_rows in iter_report_blocks(rows):
        if len(headers) < 2:
            continue
        second_header = headers[1].strip()
        if second_header == "Agen":
            sections.append(select_agent_rows_for_period(title, headers, data_rows, period))
        elif second_header == "Jetty":
            sections.append(select_jetty_rows_for_period(title, headers, data_rows, period))
    return sections


def load_workbook_seed(xlsx_path: Path, period: str) -> tuple[list[dict], list[dict], dict[str, str]]:
    rows = load_first_sheet_rows(xlsx_path)
    dashboard_rows = extract_dashboard_rows_from_workbook(rows, period)
    recap_sections = extract_recap_sections_from_workbook(rows, period)
    metadata = {
        "source_period": period,
        "source_origin": "xlsx_import",
        "source_file": str(xlsx_path),
    }
    return dashboard_rows, recap_sections, metadata


def seed_daily_movements(cursor: sqlite3.Cursor, rows: list[dict]) -> None:
    payload = [
        (
            row["tanggal"],
            row["tanggal"][:7],
            row["kantor"],
            int(row["idr"]),
            int(row["usd"]),
            int(row["total"]),
        )
        for row in rows
    ]
    cursor.executemany(
        """
        INSERT INTO daily_movements (
          movement_date,
          period,
          office,
          domestic_count,
          foreign_count,
          total_count
        )
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        payload,
    )


def seed_agent_reports(cursor: sqlite3.Cursor, sections: list[dict], period: str) -> None:
    inserts = []
    for section in sections:
        headers = section.get("headers", [])
        if len(headers) < 2 or headers[1] != "Agen":
            continue

        category = category_from_title(section["title"])
        for index, row in enumerate(section.get("rows", []), start=1):
            is_total = str(row[0]).strip().upper() == "TOTAL"
            row_no = None if is_total else parse_int(row[0])
            inserts.append(
                (
                    period,
                    category,
                    index,
                    row_no,
                    str(row[1]).strip(),
                    parse_int(row[2]),
                    parse_int(row[3]),
                    str(row[3]).strip(),
                    1 if is_total else 0,
                    section["title"],
                )
            )

    cursor.executemany(
        """
        INSERT INTO agent_reports (
          period,
          category,
          display_order,
          row_no,
          agent_name,
          movements,
          estimated_revenue,
          estimated_revenue_raw,
          is_total,
          source_title
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        inserts,
    )


def seed_jetty_reports(cursor: sqlite3.Cursor, sections: list[dict], period: str) -> None:
    inserts = []
    for section in sections:
        headers = section.get("headers", [])
        if len(headers) < 2 or headers[1] != "Jetty":
            continue

        category = category_from_title(section["title"])
        for index, row in enumerate(section.get("rows", []), start=1):
            is_total = str(row[0]).strip().upper() == "TOTAL"
            row_no = None if is_total else parse_int(row[0])
            inserts.append(
                (
                    period,
                    category,
                    index,
                    row_no,
                    str(row[1]).strip(),
                    str(row[2]).strip(),
                    parse_int(row[3]),
                    1 if is_total else 0,
                    section["title"],
                )
            )

    cursor.executemany(
        """
        INSERT INTO jetty_reports (
          period,
          category,
          display_order,
          row_no,
          jetty_name,
          region_name,
          movements,
          is_total,
          source_title
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        inserts,
    )


def upsert_metadata(cursor: sqlite3.Cursor, metadata: dict[str, str], mode: str) -> None:
    values = {
        "last_import_period": metadata["source_period"],
        "last_source_origin": metadata["source_origin"],
        "last_source_file": metadata.get("source_file", ""),
        "last_import_mode": mode,
    }
    for key, value in values.items():
        cursor.execute("INSERT OR REPLACE INTO app_metadata (key, value) VALUES (?, ?)", (key, value))


def replace_period_data(cursor: sqlite3.Cursor, period: str) -> None:
    cursor.execute("DELETE FROM daily_movements WHERE period = ?", (period,))
    cursor.execute("DELETE FROM agent_reports WHERE period = ?", (period,))
    cursor.execute("DELETE FROM jetty_reports WHERE period = ?", (period,))
    cursor.execute("DELETE FROM import_sources WHERE period = ?", (period,))


def import_seed(
    connection: sqlite3.Connection,
    dashboard_rows: list[dict],
    recap_sections: list[dict],
    metadata: dict[str, str],
    mode: str,
) -> None:
    cursor = connection.cursor()
    period = metadata["source_period"]
    replace_period_data(cursor, period)
    seed_daily_movements(cursor, dashboard_rows)
    seed_agent_reports(cursor, recap_sections, period)
    seed_jetty_reports(cursor, recap_sections, period)
    cursor.execute(
        """
        INSERT INTO import_sources (period, source_origin, source_file)
        VALUES (?, ?, ?)
        """,
        (period, metadata["source_origin"], metadata.get("source_file", "")),
    )
    upsert_metadata(cursor, metadata, mode)
    connection.commit()


def open_database(db_path: Path, mode: str) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    if mode == "rebuild" and db_path.exists():
        db_path.unlink()
    connection = sqlite3.connect(db_path)
    connection.executescript(SCHEMA_PATH.read_text(encoding="utf-8"))
    return connection


def load_project_seed() -> tuple[list[dict], list[dict], dict[str, str]]:
    dashboard_rows = load_manual_dashboard_data()
    recap_payload = load_recap_tables()
    if recap_payload and isinstance(recap_payload[0], dict) and "sections" in recap_payload[0]:
        recap_sections = recap_payload[0]["sections"]
    else:
        recap_sections = recap_payload

    period = dashboard_rows[0]["tanggal"][:7]
    metadata = {
        "source_period": period,
        "source_origin": "manual_project_seed",
        "source_file": "app.js + excel-table-data.js",
    }
    return dashboard_rows, recap_sections, metadata


def resolve_workbook_jobs(args: argparse.Namespace) -> list[tuple[Path, str]]:
    workbook_paths = [Path(item) for item in args.xlsx]
    if args.period and len(workbook_paths) > 1:
        raise ValueError("--period hanya boleh dipakai saat import satu file XLSX.")

    jobs = []
    for workbook_path in workbook_paths:
        if not workbook_path.exists():
            raise FileNotFoundError(f"File workbook tidak ditemukan: {workbook_path}")
        period = args.period or derive_period_from_filename(workbook_path)
        jobs.append((workbook_path, period))
    return jobs


def main() -> None:
    args = parse_args()
    db_path = Path(args.db)

    connection = open_database(db_path, args.mode)
    try:
        if args.source == "project":
            dashboard_rows, recap_sections, metadata = load_project_seed()
            import_seed(connection, dashboard_rows, recap_sections, metadata, args.mode)
        else:
            for workbook_path, period in resolve_workbook_jobs(args):
                dashboard_rows, recap_sections, metadata = load_workbook_seed(workbook_path, period)
                import_seed(connection, dashboard_rows, recap_sections, metadata, args.mode)
    finally:
        connection.close()


if __name__ == "__main__":
    main()
