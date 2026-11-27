#!/usr/bin/env python3
"""表データ（xlsx / csv / tsv / json）のプロファイルを作成するCLI。

Web版「ExcelをHTMLで読むツール」の概要ダッシュボードと同じ考え方で、
シートごとの行数・列数、列ごとの型・空欄率・ユニーク数・最小/最大/平均・
最頻値、要注意ポイント（空欄・重複行・型混在・外れ値）を集計する。

使い方:
    python3 profile_table.py input.xlsx [--format json|html] [--out FILE]
    python3 profile_table.py 資料1.xlsx 資料2.csv ... [--format json|html] [--out FILE]
    python3 profile_table.py 資料フォルダ/ [--format json|html] [--out FILE]

複数ファイル（またはフォルダ）の場合は、全ファイル分のプロファイルをまとめた
1つのJSON/HTMLを出力する。1ファイルだけの場合は従来どおり、そのファイル
1件分のプロファイルだけを出力する。読めないファイルが混ざっていても、
そのファイルは error 付きで記録して残りの処理を続ける。

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

SUPPORTED_SUFFIXES = {".xlsx", ".csv", ".tsv", ".json"}


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
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise ValueError(f"xlsxファイルを開けませんでした（壊れているか、未対応の形式です）: {e}") from e
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


# ---------- 複数ファイル ----------

def expand_inputs(inputs: list[str | Path]) -> list[Path]:
    """入力指定をファイル一覧に展開する。フォルダは対応形式のファイルを名前順に拾う。"""
    files: list[Path] = []
    for raw in inputs:
        path = Path(raw)
        if path.is_dir():
            files.extend(sorted(p for p in path.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_SUFFIXES))
        else:
            files.append(path)
    return files


def summarize_profiles(profiles: list[dict]) -> dict:
    """複数ファイル分のプロファイルの合計サマリ。error 付きの項目は件数だけ数える。"""
    ok = [p for p in profiles if "error" not in p]
    total_warnings = {"blank_cells": 0, "duplicate_rows": 0, "minority_type_cells": 0, "outliers": 0}
    total_rows = 0
    total_sheets = 0
    for p in ok:
        total_sheets += len(p["sheets"])
        for sheet in p["sheets"]:
            total_rows += sheet["row_count"]
            for k in total_warnings:
                total_warnings[k] += sheet["warnings"].get(k, 0)
    return {
        "file_count": len(profiles),
        "ok_count": len(ok),
        "error_count": len(profiles) - len(ok),
        "total_sheets": total_sheets,
        "total_rows": total_rows,
        "total_warnings": total_warnings,
    }


def profile_files(paths: list[str | Path]) -> list[dict]:
    """複数ファイルをプロファイル化する。読めないファイルは error 付きで残す。"""
    profiles: list[dict] = []
    for raw in paths:
        path = Path(raw)
        try:
            profiles.append(profile_file(path))
        except Exception as e:  # openpyxlのBadZipFileなど、読み込み系の例外は種類を問わず記録する
            profiles.append({"file": path.name, "error": str(e)})
    return profiles


def render_combined_html(profiles: list[dict], summary: dict) -> str:
    """複数ファイル分のプロファイルを1つのHTML（一覧＋各ファイルの詳細）にまとめる。"""
    w = summary["total_warnings"]
    parts = [
        "<!DOCTYPE html><html lang='ja'><head><meta charset='utf-8'>",
        "<title>プロファイル（一括）</title>",
        "<style>body{font-family:sans-serif;margin:2em;line-height:1.6}"
        "table{border-collapse:collapse;margin:1em 0}th,td{border:1px solid #ccc;padding:4px 8px}"
        ".warn{color:#b00;font-weight:bold}.err{color:#b00}</style></head><body>",
        "<h1>プロファイル（一括）</h1>",
        f"<p>対象 {summary['file_count']} 件（成功 {summary['ok_count']} 件 / 失敗 {summary['error_count']} 件）・"
        f"合計 {summary['total_sheets']} シート / {summary['total_rows']} 行</p>",
        "<p>要注意（合計）: 空欄セル <span class='warn'>%d</span> 件 / 重複行 <span class='warn'>%d</span> 件"
        " / 型混在の少数派セル <span class='warn'>%d</span> 件 / 外れ値 <span class='warn'>%d</span> 件</p>"
        % (w["blank_cells"], w["duplicate_rows"], w["minority_type_cells"], w["outliers"]),
        "<h2>ファイル一覧</h2><table><tr><th>ファイル</th><th>シート数</th><th>行数</th>"
        "<th>空欄</th><th>重複行</th><th>型混在</th><th>外れ値</th></tr>",
    ]
    for p in profiles:
        if "error" in p:
            parts.append(
                "<tr><td>%s</td><td colspan='6' class='err'>読めませんでした: %s</td></tr>"
                % (html.escape(p["file"]), html.escape(p["error"]))
            )
            continue
        rows = sum(s["row_count"] for s in p["sheets"])
        warns = {k: sum(s["warnings"].get(k, 0) for s in p["sheets"]) for k in w}
        parts.append(
            "<tr><td>%s</td><td>%d</td><td>%d</td><td>%d</td><td>%d</td><td>%d</td><td>%d</td></tr>"
            % (
                html.escape(p["file"]), len(p["sheets"]), rows,
                warns["blank_cells"], warns["duplicate_rows"],
                warns["minority_type_cells"], warns["outliers"],
            )
        )
    parts.append("</table>")
    for p in profiles:
        if "error" in p:
            continue
        # 既存の単体HTMLから body 部分だけ取り出して1ファイル分として埋め込む
        single = render_html(p)
        body = single.split("<body>", 1)[1].rsplit("</body>", 1)[0]
        parts.append(body)
    parts.append("</body></html>")
    return "\n".join(parts)


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
    ap.add_argument("input", nargs="+", help="入力ファイル（xlsx / csv / tsv / json）またはフォルダ。複数指定可")
    ap.add_argument("--format", choices=("json", "html"), default="json")
    ap.add_argument("--out", help="出力先ファイル（省略時は標準出力）")
    args = ap.parse_args(argv)

    files = expand_inputs(args.input)
    if not files:
        print("対応するファイルが見つかりませんでした（xlsx / csv / tsv / json）", file=sys.stderr)
        return 1
    profiles = profile_files(files)
    if len(profiles) == 1 and "error" not in profiles[0]:
        profile = profiles[0]
        text = json.dumps(profile, ensure_ascii=False, indent=2) if args.format == "json" else render_html(profile)
    else:
        summary = summarize_profiles(profiles)
        if args.format == "json":
            text = json.dumps({"files": profiles, "summary": summary}, ensure_ascii=False, indent=2)
        else:
            text = render_combined_html(profiles, summary)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
        print(f"書き出しました: {args.out}", file=sys.stderr)
    else:
        try:
            print(text)
            sys.stdout.flush()
        except BrokenPipeError:
            # パイプの読み手が先に閉じた場合（例: | head）は静かに終了する
            try:
                sys.stdout.close()
            except Exception:
                pass
            return 0
    if all("error" in p for p in profiles):
        for p in profiles:
            print(f"読めませんでした: {p['file']}: {p['error']}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
