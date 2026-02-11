import csv
import re
from xml.sax.saxutils import escape
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from statistics import median

REPO_ROOT = Path(__file__).resolve().parents[1]
XLSX_PATH = REPO_ROOT / "Grad Program Exit Survey Data 2024.xlsx"
TABLE_CSV = REPO_ROOT / "outputs/tables/course_ranking_2024.csv"
TABLE_MD = REPO_ROOT / "outputs/tables/course_ranking_2024.md"
FIG_SVG = REPO_ROOT / "outputs/figures/course_ranking_2024.svg"
REPORT_PATH = REPO_ROOT / "REPORT.md"

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def col_to_idx(col: str) -> int:
    n = 0
    for c in col:
        n = n * 26 + (ord(c) - 64)
    return n - 1


def load_sheet_rows(xlsx_path: Path):
    with zipfile.ZipFile(xlsx_path) as zf:
        shared = []
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        for si in root.findall("a:si", NS):
            t = si.find("a:t", NS)
            if t is not None:
                shared.append((t.text or "").strip())
            else:
                text = "".join(
                    (r.find("a:t", NS).text or "")
                    for r in si.findall("a:r", NS)
                    if r.find("a:t", NS) is not None
                )
                shared.append(text.strip())

        sh = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        rows = []
        for row in sh.findall(".//a:sheetData/a:row", NS):
            vals = {}
            for c in row.findall("a:c", NS):
                ref = c.attrib.get("r", "")
                m = re.match(r"([A-Z]+)(\d+)", ref)
                if not m:
                    continue
                col = col_to_idx(m.group(1))
                typ = c.attrib.get("t")
                v = c.find("a:v", NS)
                value = ""
                if v is not None and v.text is not None:
                    value = shared[int(v.text)] if typ == "s" else v.text
                vals[col] = value.strip()
            rows.append(vals)
    return rows


def parse_course_name(question_text: str) -> str:
    course = question_text.split(" - ")[-1].strip()
    course = re.sub(r"\s+", " ", course)
    replacements = {
        "Data Analytics": "Advanced Data Analytics",
        "Financial Audit": "Financial Auditing",
        "Financial Theory & Research I": "Financial Accounting Theory & Research I",
    }
    for old, new in replacements.items():
        course = course.replace(old, new)
    return course


def top2_pct(ranks):
    return round(100.0 * sum(1 for r in ranks if r <= 2) / len(ranks), 1) if ranks else 0.0


def validate_svg_safe(path: Path):
    content = path.read_text(encoding="utf-8")
    forbidden = ["<image", "data:image", "xlink:href", "href="]
    found = [token for token in forbidden if token in content]
    if found:
        raise ValueError(f"SVG contains forbidden embedded image references: {found}")

    # Also ensure XML is well-formed and parseable by strict renderers
    ET.fromstring(content)


def write_svg(rows):
    width = 1200
    left = 360
    top = 80
    row_h = 48
    chart_w = 760
    height = top + row_h * len(rows) + 70

    lines = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">',
        '<rect width="100%" height="100%" fill="white"/>',
        f'<text x="24" y="34" font-size="24" font-family="Arial" font-weight="bold">{escape("MAcc Core Course Ranking (2024)")}</text>',
        f'<text x="24" y="56" font-size="14" font-family="Arial" fill="#444">{escape("Metric: Average preference score (Rank 1=8 ... Rank 8=1)")}</text>',
    ]

    for t in range(0, 9):
        x = left + (chart_w * t / 8)
        lines.append(f'<line x1="{x:.1f}" y1="{top-18}" x2="{x:.1f}" y2="{height-30}" stroke="#e5e7eb" stroke-width="1"/>')
        lines.append(f'<text x="{x:.1f}" y="{height-10}" text-anchor="middle" font-size="12" font-family="Arial" fill="#555">{t}</text>')

    for i, row in enumerate(rows):
        y = top + i * row_h
        score = float(row["Average preference score (8=best)"])
        bar_w = chart_w * score / 8.0
        lines.append(f'<text x="24" y="{y+24}" font-size="14" font-family="Arial">{escape(row["Course"])}</text>')
        lines.append(f'<rect x="{left}" y="{y+8}" width="{bar_w:.1f}" height="24" fill="#2563eb"/>')
        lines.append(
            f'<text x="{left+bar_w+8:.1f}" y="{y+24}" font-size="13" font-family="Arial" fill="#111">'
            f'{row["Average preference score (8=best)"]} (n={row["n"]})</text>'
        )

    lines.append("</svg>")
    FIG_SVG.parent.mkdir(parents=True, exist_ok=True)
    FIG_SVG.write_text("\n".join(lines), encoding="utf-8")
    validate_svg_safe(FIG_SVG)


def main():
    rows = load_sheet_rows(XLSX_PATH)
    header_keys = [rows[0].get(i, "") for i in range(max(rows[0].keys()) + 1)]
    header_labels = [rows[1].get(i, "") for i in range(max(rows[1].keys()) + 1)]

    q35_cols = [i for i, k in enumerate(header_keys) if k.startswith("Q35_")]
    courses = {i: parse_course_name(header_labels[i]) for i in q35_cols}

    rank_data = {c: [] for c in q35_cols}
    for row in rows[3:]:
        for c in q35_cols:
            raw = row.get(c, "")
            if raw in {"", "NA", "N/A", "null", "NULL"}:
                continue
            if raw.isdigit():
                rank = int(raw)
                if 1 <= rank <= 8:
                    rank_data[c].append(rank)

    output_rows = []
    for c in q35_cols:
        ranks = rank_data[c]
        if not ranks:
            continue
        pref_scores = [9 - r for r in ranks]
        output_rows.append(
            {
                "Course": courses[c],
                "n": len(ranks),
                "Average preference score (8=best)": round(sum(pref_scores) / len(pref_scores), 2),
                "Median rank (1=best)": float(median(ranks)),
                "% favorable (rank 1-2)": top2_pct(ranks),
            }
        )

    output_rows.sort(
        key=lambda r: (r["Average preference score (8=best)"], r["% favorable (rank 1-2)"]),
        reverse=True,
    )
    for i, row in enumerate(output_rows, 1):
        row["Rank"] = i

    columns = [
        "Rank",
        "Course",
        "n",
        "Average preference score (8=best)",
        "Median rank (1=best)",
        "% favorable (rank 1-2)",
    ]

    TABLE_CSV.parent.mkdir(parents=True, exist_ok=True)
    with TABLE_CSV.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        writer.writerows(output_rows)

    with TABLE_MD.open("w", encoding="utf-8") as f:
        f.write("# Course ranking (2024)\n\n")
        f.write("Chosen metric: **Q35 core-course rank order** (most direct preference field). Converted ranks to preference score using **1→8, 2→7, ..., 8→1**.\n\n")
        f.write("| " + " | ".join(columns) + " |\n")
        f.write("|" + "---|" * len(columns) + "\n")
        for row in output_rows:
            f.write("| " + " | ".join(str(row[c]) for c in columns) + " |\n")

    write_svg(output_rows)

    top, bottom = output_rows[0], output_rows[-1]
    with REPORT_PATH.open("w", encoding="utf-8") as f:
        f.write("# 2024 Graduate Program Exit Survey: Course Ranking\n\n")
        f.write("- Selected `Q35_*` as the rating/preference field because it explicitly asks students to rank core courses from most to least beneficial; converted ordinal ranks to a preference score (Rank 1=8 ... Rank 8=1).\n")
        f.write(f"- Top ranked course: **{top['Course']}** (avg preference score {top['Average preference score (8=best)']}, median rank {top['Median rank (1=best)']}, top-2 box {top['% favorable (rank 1-2)']}%, n={top['n']}).\n")
        f.write(f"- Bottom ranked course: **{bottom['Course']}** (avg preference score {bottom['Average preference score (8=best)']}, median rank {bottom['Median rank (1=best)']}, top-2 box {bottom['% favorable (rank 1-2)']}%, n={bottom['n']}).\n")
        f.write("- Caveat: rankings come from ordinal drag-and-drop responses, so differences between adjacent average scores should be interpreted directionally rather than as interval-scale gaps.\n")


if __name__ == "__main__":
    main()
