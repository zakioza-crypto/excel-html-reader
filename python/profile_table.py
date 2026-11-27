#!/usr/bin/env python3
"""表データ（xlsx / csv / tsv / json）のプロファイルを作成するCLI。

Web版「ExcelをHTMLで読むツール」の概要ダッシュボードと同じ考え方で、
シートごとの行数・列数、列ごとの型・空欄率・ユニーク数・最小/最大/平均・
最頻値、要注意ポイント（空欄・重複行・型混在・外れ値）を集計する。

使い方:
    python3 profile_table.py input.xlsx [--format json|html] [--out FILE]

依存: xlsx の読み込みに openpyxl が必要（csv/tsv/json は標準ライブラリのみ）。
"""
from __future__ import annotations

import argparse
import csv
import html
import io
import json
import math
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path


# ---------- 読み込み ----------

def load_tables(path: Path) -> list[dict]:
    """ファイルを読み、[{"name": str, "header": [str], "rows": [[cell]]}] を返す。"""
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return _load_xlsx(path)
    if suffix in (".csv", ".tsv"):
        return [_load_delimited(path, "\t" if suffix == ".tsv" else ",")]
    if suffix == ".json":
        return [_load_json(path)]
    raise ValueError(f"未対応の拡張子です: {suffix}（xlsx / csv / tsv / json のみ）")


def _load_xlsx(path: Path) -> list[dict]:
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise RuntimeError("xlsxの読み込みには openpyxl が必要です: pip install openpyxl")
    wb = load_workbook(path, read_only=True, data_only=True)
    tables = []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        rows = [r for r in rows if any(c is not None and str(c).strip() != "" for c in r)]
        if not rows:
            tables.append({"name": ws.title, "header": [], "rows": []})
            continue
        header = ["" if c is None else str(c) for c in rows[0]]
        tables.append({"name": ws.title, "header": header, "rows": rows[1:]})
    return tables


def _decode_text(path: Path) -> str:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp932"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _load_delimited(path: Path, delimiter: str) -> dict:
    text = _decode_text(path)
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows = [r for r in reader if any(c.strip() != "" for c in r)]
    if not rows:
        return {"name": path.stem, "header": [], "rows": []}
    return {"name": path.stem, "header": rows[0], "rows": rows[1:]}


def _load_json(path: Path) -> dict:
    data = json.loads(_decode_text(path))
    if not isinstance(data, list) or not all(isinstance(d, dict) for d in data):
        raise ValueError("JSONは「オブジェクトの配列」形式のみ対応しています")
    header: list[str] = []
    for d in data:
        for k in d:
            if k not in header:
                header.append(k)
    rows = []
    for d in data:
        row = []
        for k in header:
            v = d.get(k)
            if isinstance(v, (dict, list)):
                v = json.dumps(v, ensure_ascii=False)
            row.append(v)
        rows.append(row)
    return {"name": path.stem, "header": header, "rows": rows}


# ---------- 型判定と統計 ----------

def _cell_kind(v) -> str:
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return "blank"
    if isinstance(v, bool):
        return "string"
    if isinstance(v, (int, float)):
        return "number" if not (isinstance(v, float) and math.isnan(v)) else "blank"
    if isinstance(v, (datetime, date)):
        return "date"
    s = str(v).strip()
    try:
        float(s.replace(",", ""))
        return "number"
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            datetime.strptime(s, fmt)
            return "date"
        except ValueError:
            continue
    return "string"


def _to_number(v) -> float | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(",", ""))
    except (ValueError, TypeError):
        return None


def _percentile(sorted_vals: list[float], p: float) -> float:
    if not sorted_vals:
        return float("nan")
    k = (len(sorted_vals) - 1) * p
    f = math.floor(k)
    c = math.ceil(k)
    if f == c:
        return sorted_vals[int(k)]
    return sorted_vals[f] * (c - k) + sorted_vals[c] * (k - f)


def profile_column(values: list) -> dict:
    kinds = [_cell_kind(v) for v in values]
    total = len(values)
    blanks = kinds.count("blank")
    non_blank_kinds = [k for k in kinds if k != "blank"]
    kind_counts = Counter(non_blank_kinds)
    if not kind_counts:
        col_type = "empty"
        minority = 0
    elif len(kind_counts) == 1:
        col_type = next(iter(kind_counts))
        minority = 0
    else:
        col_type = "mixed"
        majority = kind_counts.most_common(1)[0][0]
        minority = sum(c for k, c in kind_counts.items() if k != majority)

    display = ["" if _cell_kind(v) == "blank" else str(v) for v in values]
    uniques = len({d for d in display if d != ""})
    top = Counter(d for d in display if d != "").most_common(3)

    result = {
        "type": col_type,
        "rows": total,
        "blank_count": blanks,
        "blank_rate": round(blanks / total, 4) if total else 0.0,
        "unique_count": uniques,
        "minority_type_cells": minority,
        "top_values": [{"value": v, "count": c} for v, c in top],
    }

    numbers = sorted(n for v in values if (n := _to_number(v)) is not None)
    if numbers:
        q1 = _percentile(numbers, 0.25)
        q3 = _percentile(numbers, 0.75)
        iqr = q3 - q1
        lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        outliers = sum(1 for n in numbers if n < lo or n > hi)
        result["numeric"] = {
            "min": numbers[0],
            "max": numbers[-1],
            "avg": round(sum(numbers) / len(numbers), 4),
            "outlier_count": outliers,
        }
    return result


def profile_table(table: dict) -> dict:
    header, rows = table["header"], table["rows"]
    width = len(header)
    norm_rows = [(r + [None] * width)[:width] for r in rows]
    columns = []
    for i in range(width):
        col = profile_column([r[i] for r in norm_rows])
        col["name"] = header[i] if header[i] else f"(列{i + 1})"
        columns.append(col)
    dup_rows = len(norm_rows) - len({tuple("" if _cell_kind(c) == "blank" else str(c) for c in r) for r in norm_rows})
    warnings = {
        "blank_cells": sum(c["blank_count"] for c in columns),
        "duplicate_rows": dup_rows,
        "minority_type_cells": sum(c["minority_type_cells"] for c in columns),
        "outliers": sum(c.get("numeric", {}).get("outlier_count", 0) for c in columns),
    }
    return {
        "name": table["name"],
        "row_count": len(norm_rows),
        "column_count": width,
        "columns": columns,
        "warnings": warnings,
    }


def profile_file(path: str | Path) -> dict:
    path = Path(path)
    tables = load_tables(path)
    return {"file": path.name, "sheets": [profile_table(t) for t in tables]}


# ---------- HTML出力 ----------

def render_html(profile: dict) -> str:
    parts = [
        "<!DOCTYPE html><html lang='ja'><head><meta charset='utf-8'>",
        f"<title>プロファイル: {html.escape(profile['file'])}</title>",
        "<style>body{font-family:sans-serif;margin:2em;line-height:1.6}"
        "table{border-collapse:collapse;margin:1em 0}th,td{border:1px solid #ccc;padding:4px 8px}"
        ".warn{color:#b00;font-weight:bold}</style></head><body>",
        f"<h1>{html.escape(profile['file'])} のプロファイル</h1>",
    ]
    for sheet in profile["sheets"]:
        w = sheet["warnings"]
        parts.append(f"<h2>{html.escape(sheet['name'])}（{sheet['row_count']}行 × {sheet['column_count']}列）</h2>")
        parts.append(
            "<p>要注意: 空欄セル <span class='warn'>%d</span> 件 / 重複行 <span class='warn'>%d</span> 件"
            " / 型混在の少数派セル <span class='warn'>%d</span> 件 / 外れ値 <span class='warn'>%d</span> 件</p>"
            % (w["blank_cells"], w["duplicate_rows"], w["minority_type_cells"], w["outliers"])
        )
        parts.append("<table><tr><th>列名</th><th>型</th><th>空欄</th><th>ユニーク数</th><th>最小</th><th>最大</th><th>平均</th><th>最頻値</th></tr>")
        for c in sheet["columns"]:
            num = c.get("numeric", {})
            top = "、".join(f"{html.escape(t['value'])}（{t['count']}）" for t in c["top_values"])
            parts.append(
                "<tr><td>%s</td><td>%s</td><td>%d（%.0f%%）</td><td>%d</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>"
                % (
                    html.escape(c["name"]), c["type"], c["blank_count"], c["blank_rate"] * 100,
                    c["unique_count"],
                    num.get("min", "-"), num.get("max", "-"), num.get("avg", "-"), top,
                )
            )
        parts.append("</table>")
    parts.append("</body></html>")
    return "\n".join(parts)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="表データのプロファイルを作成します")
    ap.add_argument("input", help="入力ファイル（xlsx / csv / tsv / json）")
    ap.add_argument("--format", choices=("json", "html"), default="json")
    ap.add_argument("--out", help="出力先ファイル（省略時は標準出力）")
    args = ap.parse_args(argv)

    profile = profile_file(args.input)
    text = json.dumps(profile, ensure_ascii=False, indent=2) if args.format == "json" else render_html(profile)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
        print(f"書き出しました: {args.out}", file=sys.stderr)
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
