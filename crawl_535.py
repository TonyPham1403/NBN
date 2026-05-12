import argparse
import itertools
import importlib
import os
from collections import Counter

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment


URL = "https://www.vietlott.vn/vi/trung-thuong/ket-qua-trung-thuong/winning-number-535#top"


def build_headers():
    return {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": (
            "text/html,application/xhtml+xml,application/xml;"
            "q=0.9,image/avif,image/webp,*/*;q=0.8"
        ),
        "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "Referer": "https://www.vietlott.vn/",
        "Upgrade-Insecure-Requests": "1",
    }


def fetch_html(url: str):
    headers = build_headers()

    resp = requests.get(url, timeout=30, headers=headers)
    if resp.status_code == 403:
        try:
            curl_requests = importlib.import_module("curl_cffi.requests")

            resp = curl_requests.get(
                url,
                timeout=30,
                headers=headers,
                impersonate="chrome124",
            )
        except ImportError as exc:
            raise RuntimeError(
                "Vietlott dang chan requests thuong. Hay cai curl_cffi de fake browser request."
            ) from exc

    resp.raise_for_status()
    return resp.text


def parse_args():
    parser = argparse.ArgumentParser(description="Auto update 535.xlsm from Vietlott latest draw.")
    parser.add_argument("--file", default="535.xlsm", help="Path to XLSM file (relative to repo root).")
    parser.add_argument("--url", default=URL, help="Vietlott URL.")
    return parser.parse_args()


def fetch_latest_record(url: str):
    html = fetch_html(url)

    soup = BeautifulSoup(html, "html.parser")
    row = soup.select_one("#divResultContent table tbody tr")
    if row is None:
        raise RuntimeError("Khong tim thay dong ket qua tren trang Vietlott.")

    cols = row.select("td")
    if len(cols) < 3:
        raise RuntimeError("Khong doc duoc du lieu cot ngay/id/result.")

    date_text = cols[0].get_text(strip=True)
    id_text = cols[1].get_text(strip=True)
    id_num = int(id_text)

    spans = cols[2].select("span")
    values: list[str] = []
    for span in spans:
        txt = span.get_text(strip=True)
        if not txt or txt == "|":
            continue
        if txt.isdigit():
            values.append(str(int(txt)))

    if len(values) < 6:
        raise RuntimeError(f"Khong du 6 so de tao ket qua: {values}")

    nums = values[:5]
    special = values[5]
    result = ",".join(nums) + "|" + special

    return {"date": date_text, "id": id_num, "result": result}


def load_existing_rows(file_path: str):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Khong tim thay file: {file_path}")

    wb = load_workbook(file_path, keep_vba=True)
    ws = wb.active
    rows = []
    for r in range(2, ws.max_row + 1):
        date_val = ws.cell(r, 1).value
        id_val = ws.cell(r, 2).value
        result_val = ws.cell(r, 3).value
        if id_val is None:
            continue
        id_digits = "".join(ch for ch in str(id_val) if ch.isdigit())
        if not id_digits:
            continue
        rows.append(
            {
                "date": str(date_val).strip() if date_val is not None else "",
                "id": int(id_digits),
                "result": str(result_val).strip() if result_val is not None else "",
            }
        )
    return wb, ws, rows


def save_rows_to_workbook(wb, ws, rows, file_path: str):
    df = pd.DataFrame(rows, columns=["date", "id", "result"])
    df = df.drop_duplicates(subset="id").sort_values(by="id")
    df["id"] = df["id"].astype(int).astype(str).str.zfill(5)

    for r in range(2, ws.max_row + 1):
        ws.cell(r, 1).value = None
        ws.cell(r, 2).value = None
        ws.cell(r, 3).value = None

    for i, row in df.iterrows():
        ws.cell(row=i + 2, column=1, value=row["date"])
        ws.cell(row=i + 2, column=2, value=row["id"])
        ws.cell(row=i + 2, column=3, value=row["result"])

    ws.freeze_panes = "A2"
    center = Alignment(horizontal="center")
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 1).alignment = center
        ws.cell(r, 2).alignment = center

    for r in range(2, ws.max_row + 1):
        cell = ws[f"C{r}"]
        if cell.value and "|" in str(cell.value):
            left, right = str(cell.value).split("|", 1)
            rich = CellRichText()
            rich.append(TextBlock(InlineFont(rFont="Calibri", sz=11), left + "|"))
            rich.append(
                TextBlock(
                    InlineFont(rFont="Calibri", sz=16, b=True, color="FF006400"),
                    right,
                )
            )
            cell.value = rich

    rebuild_stats_sheets(wb, ws)
    wb.save(file_path)
    print("Saved rows:", len(df))


def rebuild_stats_sheets(wb, ws):
    main_numbers: list[list[int]] = []
    special_numbers: list[int] = []

    for r in range(2, ws.max_row + 1):
        val = ws[f"C{r}"].value
        if not val:
            continue
        text = str(val)
        if "|" not in text:
            continue
        left, right = text.split("|", 1)
        nums = [int(x) for x in left.split(",") if str(x).strip().isdigit()]
        if len(nums) != 5:
            continue
        if not str(right).strip().isdigit():
            continue
        main_numbers.append(nums)
        special_numbers.append(int(right))

    combo_counts = {k: Counter() for k in range(1, 6)}
    for nums in main_numbers:
        nums = sorted(nums)
        for k in range(1, 6):
            for c in itertools.combinations(nums, k):
                combo_counts[k][c] += 1

    special_counter = Counter(special_numbers)

    for name in ["combo_1", "combo_2", "combo_3", "combo_4", "combo_5", "special_freq"]:
        if name in wb.sheetnames:
            del wb[name]

    for k in range(1, 6):
        ws2 = wb.create_sheet(f"combo_{k}")
        ws2.append(["combo", "appear"])
        rows = [
            (",".join(map(str, c)), cnt)
            for c, cnt in combo_counts[k].items()
            if cnt >= 2
        ]
        rows.sort(key=lambda x: x[1], reverse=True)
        for row in rows:
            ws2.append(row)

    ws2 = wb.create_sheet("special_freq")
    ws2.append(["special", "count"])
    rows = sorted(special_counter.items(), key=lambda x: x[1], reverse=True)
    for row in rows:
        ws2.append(row)


def main():
    args = parse_args()
    wb, ws, existing_rows = load_existing_rows(args.file)
    latest = fetch_latest_record(args.url)

    existing_ids = {row["id"] for row in existing_rows}
    if latest["id"] in existing_ids:
        print(f"No new draw. Latest ID {latest['id']} da ton tai.")
        return 0

    print(f"Append new draw: {latest['date']} {latest['id']} {latest['result']}")
    existing_rows.append(latest)
    save_rows_to_workbook(wb, ws, existing_rows, args.file)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

