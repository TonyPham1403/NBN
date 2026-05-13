"""
Dong bo tat ca ky xo so 535 (Max 3D+ / Lotto 21h) moi hon data.json hien co.
Chi lay tu trang xskt xslotto-5-35. Ghi truc tiep proj/data.json,
giu dung hang cuoi result rong nhu convert_data_to_json.js.
"""
from __future__ import annotations

import argparse
import json
import os
import re
from datetime import datetime, timedelta

import requests
from bs4 import BeautifulSoup


DEFAULT_XSKT_535_URL = "https://xskt.com.vn/xslotto-5-35"


def _headers(referer: str | None) -> dict[str, str]:
    h: dict[str, str] = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
        ),
        "Accept": (
            "text/html,application/xhtml+xml,application/xml;q=0.9,"
            "image/avif,image/webp,image/apng,*/*;q=0.8"
        ),
        "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
        "Cache-Control": "max-age=0",
        "Sec-Ch-Ua": '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
        "Sec-Ch-Ua-Mobile": "?0",
        "Sec-Ch-Ua-Platform": '"Windows"',
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin" if referer else "none",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
    }
    if referer:
        h["Referer"] = referer
    return h


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Dong bo ky moi vao data.json tu xskt xslotto-5-35 (khong sua xlsm).",
    )
    p.add_argument(
        "--json",
        default=os.environ.get("VIETLOTT_535_DATA_JSON", "data.json"),
        help="Duong toi data.json (mac dinh data.json trong cwd).",
    )
    p.add_argument(
        "--url",
        default=os.environ.get("VIETLOTT_535_XSKT_535_URL", "").strip() or DEFAULT_XSKT_535_URL,
        help="URL trang ket qua Lotto 5/35 tren xskt (mac dinh xslotto-5-35).",
    )
    return p.parse_args()


def _result_meaningful(result: str) -> bool:
    """Giong convert_data_to_json.js: result co ky tu so/khong phai dau phan cach."""
    return bool(re.sub(r"[,|\s]", "", str(result or "").strip()))


def _parse_id_int(id_raw: str) -> int | None:
    m = re.search(r"(\d+)$", str(id_raw or "").strip())
    if not m:
        return None
    return int(m.group(1), 10)


def normalize_date_text(value: str) -> str:
    text = str(value or "").strip()
    parts = [p.strip() for p in re.split(r"[/\-.]", text) if p.strip()]
    if len(parts) != 3:
        return ""
    dd, mm, yyyy = parts[0], parts[1], parts[2]
    if len(yyyy) == 2:
        yyyy = "20" + yyyy
    if not re.fullmatch(r"\d{1,2}", dd) or not re.fullmatch(r"\d{1,2}", mm) or not re.fullmatch(r"\d{4}", yyyy):
        return ""
    return f"{int(dd):02d}/{int(mm):02d}/{yyyy}"


def add_one_day(date_text: str) -> str:
    n = normalize_date_text(date_text)
    if not n:
        return ""
    dd, mm, yyyy = (int(x) for x in n.split("/"))
    dt = datetime(yyyy, mm, dd) + timedelta(days=1)
    return dt.strftime("%d/%m/%Y")


def compute_trailing_blank_date(rows: list[dict]) -> str:
    """Giong computeTrailingBlankDate trong convert_data_to_json.js."""
    if not rows:
        return ""
    last = rows[-1]
    last_date = normalize_date_text(str(last.get("date", "")))
    if not last_date:
        return ""
    if len(rows) < 2:
        return last_date
    prev = rows[-2]
    prev_date = normalize_date_text(str(prev.get("date", "")))
    if not prev_date:
        return last_date
    if prev_date == last_date:
        return add_one_day(last_date) or last_date
    return last_date


def next_id_string_after_rows(rows: list[dict]) -> str:
    """Tang id theo do rong chu so nhu convert_data_to_json.js."""
    for i in range(len(rows) - 1, -1, -1):
        id_raw = str(rows[i].get("id") or "").strip()
        if not id_raw:
            continue
        m = re.search(r"(\d+)$", id_raw)
        if not m:
            continue
        width = len(m.group(1))
        n = int(m.group(1), 10) + 1
        return str(n).zfill(width)
    return "00001"


def read_real_rows_from_data_json(path: str) -> tuple[list[dict], int]:
    """
    Doc cac dong co result day du; tra ve list {date, id:int, result} sap xep theo id,
    va max_id (0 neu rong).
    """
    if not os.path.exists(path):
        return [], 0
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise ValueError(f"{path} khong phai mang JSON.")
    out: list[dict] = []
    for row in data:
        if not isinstance(row, dict):
            continue
        result = str(row.get("result") or "").strip()
        if not _result_meaningful(result):
            continue
        rid = _parse_id_int(str(row.get("id") or ""))
        if rid is None:
            continue
        out.append(
            {
                "date": str(row.get("date") or "").strip(),
                "id": rid,
                "result": result,
            }
        )
    out.sort(key=lambda r: r["id"])
    max_id = out[-1]["id"] if out else 0
    return out, max_id


def rows_to_json_file_shape(rows: list[dict]) -> list[dict]:
    """id thanh chuoi 5 so nhu data.json."""
    return [
        {
            "date": r["date"],
            "id": f"{r['id']:05d}",
            "result": r["result"],
        }
        for r in rows
    ]


def append_trailing_placeholder(rows_for_json: list[dict]) -> list[dict]:
    """Them dong cuoi result rong — giong cuoi convert_data_to_json.js."""
    out = list(rows_for_json)
    next_id = next_id_string_after_rows(out)
    next_date = compute_trailing_blank_date(out)
    out.append({"date": next_date, "id": next_id, "result": ""})
    return out


def write_data_json(path: str, rows_with_trailing: list[dict]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rows_with_trailing, f, ensure_ascii=False, indent=2)
        f.write("\n")


def _fetch_all_from_xskt_535(page_url: str) -> list[dict]:
    ua = _headers(None)["User-Agent"]
    resp = requests.get(page_url, timeout=60, headers={"User-Agent": ua})
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    out: list[dict] = []
    for table in soup.select("table.result"):
        kmt = table.select_one("td.kmt")
        if not kmt:
            continue
        mid = re.search(r"#(\d{5})", kmt.get_text())
        if not mid:
            continue
        draw_id = int(mid.group(1))
        em = table.select_one("td.megaresult em")
        if not em:
            continue
        # Lay 5 so + so dac biet theo thu tu trong <em> (get_text).
        # Khong dung replace(tung span): neu mega trung mot so trong 5 so
        # (vd 01 ... <span>01</span>) thi replace xoa nham ca so chinh.
        parts = [
            p for p in re.split(r"\s+", em.get_text(" ", strip=True)) if p.isdigit()
        ]
        if len(parts) < 6:
            continue
        main_nums = parts[:5]
        special = parts[-1]
        if len(main_nums) != 5 or not special.isdigit():
            continue

        date_text = ""
        link = kmt.select_one('a[href*="ngay-"]')
        href = link.get("href") if link else ""
        if href:
            dm = re.search(r"ngay-(\d+)-(\d+)-(\d{4})", href)
            if dm:
                da, mo, ye = int(dm.group(1)), int(dm.group(2)), dm.group(3)
                date_text = f"{da:02d}/{mo:02d}/{ye}"
        if not date_text:
            continue
        result = ",".join(str(int(x)) for x in main_nums) + "|" + str(int(special))
        out.append({"date": date_text, "id": draw_id, "result": result})
    if not out:
        raise RuntimeError("Khong tim thay bang Lotto 5/35 tren xskt xslotto-5-35.")
    return out


def _fetch_new_rows_xskt(min_id_exclusive: int, page_url: str) -> list[dict]:
    try:
        rows = _fetch_all_from_xskt_535(page_url)
    except Exception as e:
        print(
            f"Canh bao: khong lay duoc ky moi tu xskt (se thu lai o lan sau). {e}",
            flush=True,
        )
        return []
    merged: dict[int, dict] = {}
    for r in rows:
        if r["id"] <= min_id_exclusive:
            continue
        merged[r["id"]] = dict(r)
    return sorted(merged.values(), key=lambda r: r["id"])


def main() -> int:
    args = parse_args()
    json_path = args.json
    if not os.path.isabs(json_path):
        json_path = os.path.join(os.getcwd(), json_path)

    existing, max_id = read_real_rows_from_data_json(json_path)

    new_rows = _fetch_new_rows_xskt(max_id, args.url)
    if not new_rows:
        print(f"Khong co ky moi (max_id hien tai: {max_id:05d}).")
        return 0

    by_id = {r["id"]: dict(r) for r in existing}
    for r in new_rows:
        by_id[r["id"]] = dict(r)
    combined = sorted(by_id.values(), key=lambda r: r["id"])
    shaped = rows_to_json_file_shape(combined)
    final_out = append_trailing_placeholder(shaped)
    write_data_json(json_path, final_out)
    ids_added = ", ".join(f"{r['id']:05d}" for r in new_rows)
    print(f"Them {len(new_rows)} ky: {ids_added}. Ghi {json_path} ({len(final_out)} dong, co hang cuoi result rong).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
