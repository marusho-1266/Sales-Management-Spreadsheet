#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Joom.csv を見やすい形式に変換するスクリプト
- 改行を含むDescriptionを1行にまとめる
- 主要列のみに絞った簡易CSVを出力
- サマリ付きMarkdown一覧も出力
"""

import csv
import os
from pathlib import Path

# 列名 → 0-based index（ヘッダー解析で取得するが、固定マッピングも用意）
KEY_COLS = {
    "Product ID": 0,
    "Product SKU": 2,
    "Name": 3,
    "Description": 4,
    "Brand": 8,
    "State": 10,
    "Product Main Image URL": 12,
    "Date Product Uploaded": 13,
    "Variant SKU": 36,
    "Shipping Weight (kg)": 38,
    "Shipping Length (cm)": 39,
    "Shipping Width (cm)": 40,
    "Shipping Height (cm)": 41,
    "Color": 44,
    "Size": 45,
    "Price (without VAT)": 46,
    "Currency": 50,
    "Inventory (default warehouse)": 51,
    "Status": 56,
    "Date Variant Uploaded": 57,
}

DESC_MAX = 120  # 説明文の最大表示文字数


def flatten_desc(s: str) -> str:
    if not s or not isinstance(s, str):
        return ""
    t = " ".join(s.split())
    return (t[:DESC_MAX] + "…") if len(t) > DESC_MAX else t


def parse_joom_csv(path: str) -> tuple[list[str], list[list]]:
    with open(path, "r", encoding="utf-8", newline="") as f:
        r = csv.reader(f)
        header = next(r)
        rows = list(r)
    return header, rows


def build_col_index(header: list[str]) -> dict[str, int]:
    return {h: i for i, h in enumerate(header)}


def extract(row: list, idx: dict, col: str, default: str = "") -> str:
    i = idx.get(col, KEY_COLS.get(col, -1))
    if i < 0 or i >= len(row):
        return default
    v = (row[i] or "").strip()
    return v if v else default


def main():
    base = Path(__file__).resolve().parent.parent
    src = base / "doc" / "Joom.csv"
    out_dir = base / "doc"

    if not src.exists():
        print(f"NotFound: {src}")
        return 1

    header, rows = parse_joom_csv(str(src))
    idx = build_col_index(header)

    # 簡易CSV用ヘッダー（日本語）
    simple_header = [
        "商品ID",
        "商品SKU",
        "商品名",
        "状態",
        "ブランド",
        "価格(税抜)",
        "通貨",
        "在庫",
        "バリアントSKU",
        "色",
        "サイズ",
        "重量(kg)",
        "出品日",
        "説明(要約)",
    ]

    simple_rows = []
    for row in rows:
        if len(row) < 2:
            continue
        desc = extract(row, idx, "Description")
        simple_rows.append([
            extract(row, idx, "Product ID"),
            extract(row, idx, "Product SKU"),
            extract(row, idx, "Name"),
            extract(row, idx, "State"),
            extract(row, idx, "Brand"),
            extract(row, idx, "Price (without VAT)"),
            extract(row, idx, "Currency"),
            extract(row, idx, "Inventory (default warehouse)"),
            extract(row, idx, "Variant SKU"),
            extract(row, idx, "Color"),
            extract(row, idx, "Size"),
            extract(row, idx, "Shipping Weight (kg)"),
            extract(row, idx, "Date Product Uploaded"),
            flatten_desc(desc),
        ])

    # --- 簡易CSV（UTF-8 BOM：Excelで文字化けしにくい）
    out_csv = out_dir / "Joom_一覧.csv"
    with open(out_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(simple_header)
        w.writerows(simple_rows)
    print(f"Created: {out_csv} ({len(simple_rows)} 件)")

    # --- Markdown一覧（サマリ + 表、長いので先頭50件＋件数）
    out_md = out_dir / "Joom_一覧.md"
    state_count = {}
    for r in simple_rows:
        s = r[3] or "(空)"
        state_count[s] = state_count.get(s, 0) + 1

    def short_date(s: str) -> str:
        if not s or len(s) < 10:
            return s
        return s[:10]  # YYYY-MM-DD

    table_cap = 50
    show_count = min(len(simple_rows), table_cap)
    section_title = "一覧（全件）" if len(simple_rows) <= table_cap else f"一覧（先頭{table_cap}件）"

    md_lines = [
        "# Joom 出品商品 一覧（見やすく成型）",
        "",
        f"元ファイル: `Joom.csv` を解析し、主要項目のみ抽出しました。",
        "",
        "## サマリ",
        f"- **総商品数（バリアント）**: {len(simple_rows)} 件",
        "- **状態別**: " + ", ".join(f"{k}: {v}件" for k, v in sorted(state_count.items())),
        "",
        f"## {section_title}",
        "",
        "| 商品SKU | 商品名 | 状態 | ブランド | 価格 | 通貨 | 在庫 | 出品日 |",
        "|---------|--------|------|----------|------|------|------|--------|",
    ]

    for r in simple_rows[:show_count]:
        name = (r[2][:30] + "…") if len(r[2]) > 30 else r[2]
        md_lines.append(f"| {r[1]} | {name} | {r[3]} | {r[4]} | {r[5]} | {r[6]} | {r[7]} | {short_date(r[12])} |")

    if len(simple_rows) > 50:
        md_lines.append("")
        md_lines.append(f"※ 全 {len(simple_rows)} 件の詳細は `Joom_一覧.csv` を参照してください。")

    md_lines.extend([
        "",
        "## 出力ファイル",
        "- **Joom_一覧.csv** … 主要列のみのCSV（1行1商品、説明は要約）。Excel等で開いて編集・フィルタしやすい形式。",
        "- **Joom_一覧.md** … 本ドキュメント。サマリと一覧表。",
        "- 再生成: `python scripts/format_joom_csv.py` で `doc/Joom.csv` から再作成できます。",
        "",
    ])

    with open(out_md, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))
    print(f"Created: {out_md}")

    return 0


if __name__ == "__main__":
    exit(main())
