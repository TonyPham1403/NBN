"""
Dong bo tat ca ky xo so Lotto 5/35 moi hon data.json hien co.

Thu tu nguon:
  1) Live JS Minh Chinh (cap nhat nhanh sau khi quay):
     https://www.minhchinh.com/livekqxs/xstt/js/LT535.js
  2) HTML truc tiep + trang theo ngay (fallback):
     https://www.minhchinh.com/truc-tiep-xo-so-tu-chon-lotto-535.html
Vietlott chinh thuc bi Cloudflare chan bot — khong dung duoc bang requests.
xskt (ben thu 3) bo: cham + de sai.

Moi lan lay: hard-refresh (no-cache + cache-bust query) de tranh HTML cu.
Ghi truc tiep proj/data.json, giu hang cuoi result rong nhu convert_data_to_json.js.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import time
from datetime import datetime, timedelta
from urllib.parse import urlparse, urlunparse, parse_qsl, urlencode

import requests
from bs4 import BeautifulSoup


DEFAULT_MINHCHINH_URL = (
    "https://www.minhchinh.com/truc-tiep-xo-so-tu-chon-lotto-535.html"
)
DEFAULT_LIVE_JS_URL = "https://www.minhchinh.com/livekqxs/xstt/js/LT535.js"
DAY_PAGE_TMPL = (
    "https://www.minhchinh.com/xs-lotto-535-ket-qua-lotto-535-ngay-{dd}-{mm}-{yyyy}.html"
)


def _headers(referer: str | None = None) -> dict[str, str]:
    """Header trinh duyet + hard-refresh (no-cache)."""
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
        # Hard refresh: ep proxy/CDN khong tra ban cache.
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache",
        "Expires": "0",
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


def _cache_bust_url(url: str) -> str:
    """Them ?_=<ms> (hoac &_=) de moi request la URL moi — tranh cache CDN/trinh duyet."""
    parts = urlparse(url)
    q = dict(parse_qsl(parts.query, keep_blank_values=True))
    q["_"] = str(int(time.time() * 1000))
    return urlunparse(parts._replace(query=urlencode(q)))


def fetch_html(url: str, referer: str | None = None, timeout: int = 60) -> str:
    bust = _cache_bust_url(url)
    resp = requests.get(bust, timeout=timeout, headers=_headers(referer))
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Dong bo ky moi vao data.json tu Minh Chinh Lotto 5/35.",
    )
    p.add_argument(
        "--json",
        default=os.environ.get("VIETLOTT_535_DATA_JSON", "data.json"),
        help="Duong toi data.json (mac dinh data.json trong cwd).",
    )
    p.add_argument(
        "--url",
        default=(
            os.environ.get("VIETLOTT_535_MINHCHINH_URL", "").strip()
            or DEFAULT_MINHCHINH_URL
        ),
        help="URL trang HTML truc tiep Lotto 5/35 Minh Chinh (fallback).",
    )
    p.add_argument(
        "--live-js",
        default=(
            os.environ.get("VIETLOTT_535_LIVE_JS_URL", "").strip()
            or DEFAULT_LIVE_JS_URL
        ),
        help="URL live JS LT535 (uu tien — cap nhat nhanh hon HTML).",
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


def _nums_from_box(box) -> tuple[list[int], int] | None:
    """5 so chinh (.ball_lotto) + 1 so DB (.ball_lotto_2 hoac .special)."""
    if box is None:
        return None
    mains = [
        int(s.get_text(strip=True))
        for s in box.select("span.ball_lotto")
        if re.fullmatch(r"\d{1,2}", s.get_text(strip=True) or "")
        and "ball_lotto_2" not in (s.get("class") or [])
        and "special" not in (s.get("class") or [])
    ]
    special_el = box.select_one("span.ball_lotto_2, span.special")
    if not special_el:
        # Fallback: 6 so theo thu tu trong box
        parts = [
            int(p)
            for p in re.findall(r"\d{1,2}", box.get_text(" ", strip=True))
        ]
        if len(parts) < 6:
            return None
        return parts[:5], parts[5]
    sp_txt = special_el.get_text(strip=True)
    if not re.fullmatch(r"\d{1,2}", sp_txt or ""):
        return None
    if len(mains) < 5:
        return None
    return mains[:5], int(sp_txt)


def _format_result(mains: list[int], special: int) -> str:
    return ",".join(str(int(x)) for x in mains) + "|" + str(int(special))


def _parse_ky_date_session(text: str) -> tuple[int | None, str, str]:
    """
    'Kết quả QSMT kỳ #964 ngày 23/10/2026 - Lúc 21:00'
    → (964, '23/10/2026', '21h')
    """
    t = " ".join(str(text or "").split())
    mid = re.search(r"#\s*(\d{1,5})", t)
    draw_id = int(mid.group(1)) if mid else None
    dm = re.search(r"(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4})", t)
    date_text = normalize_date_text(dm.group(1)) if dm else ""
    session = ""
    sm = re.search(r"(\d{1,2})\s*:\s*\d{2}", t)
    if sm:
        hour = int(sm.group(1))
        session = "21h" if hour >= 18 else "13h"
    elif re.search(r"\b21h\b", t, re.I):
        session = "21h"
    elif re.search(r"\b13h\b", t, re.I):
        session = "13h"
    return draw_id, date_text, session


def parse_draws_from_minhchinh_html(html: str) -> list[dict]:
    """
    Parse moi khoi .content_vl co .ngay + .box_ketqua (trang truc tiep / trang theo ngay).
    Tra ve list {date, id, result} — moi id 1 lan (uu tien lan gap dau = moi hon).
    """
    soup = BeautifulSoup(html, "html.parser")
    out: list[dict] = []
    seen: set[int] = set()
    for block in soup.select("div.content_vl"):
        ngay = block.select_one("div.ngay")
        box = block.select_one("div.box_ketqua")
        if not ngay or not box:
            continue
        draw_id, date_text, _sess = _parse_ky_date_session(ngay.get_text(" ", strip=True))
        if draw_id is None or not date_text:
            continue
        parsed = _nums_from_box(box)
        if not parsed:
            continue
        mains, special = parsed
        if draw_id in seen:
            continue
        seen.add(draw_id)
        out.append(
            {
                "date": date_text,
                "id": draw_id,
                "result": _format_result(mains, special),
            }
        )
    return out


def parse_history_table_minhchinh(html: str, latest_id: int | None) -> list[dict]:
    """
    Bang '15 ky gan nhat': khong co #ky — gan id = latest_id, latest_id-1, ...
    (hang moi nhat truoc). Chi dung khi da biet latest_id tu khoi .ngay.
    """
    if latest_id is None:
        return []
    soup = BeautifulSoup(html, "html.parser")
    out: list[dict] = []
    for table in soup.select("table"):
        headers = " ".join(
            th.get_text(" ", strip=True).lower() for th in table.select("th")
        )
        if "kết quả" not in headers and "ket qua" not in headers:
            continue
        if "ngày" not in headers and "ngay" not in headers:
            continue
        rows = table.select("tr")[1:]
        for i, tr in enumerate(rows):
            tds = tr.find_all("td")
            if len(tds) < 2:
                continue
            date_cell = tds[0].get_text(" ", strip=True)
            dm = re.search(r"(\d{1,2}[/\-.]\d{1,2}[/\-.](?:\d{2}|\d{4}))", date_cell)
            if not dm:
                continue
            date_text = normalize_date_text(dm.group(1))
            if not date_text:
                continue
            # Prefer mini-balls trong o ket qua
            box_like = tds[1]
            mains = [
                int(s.get_text(strip=True))
                for s in box_like.select("span.mini-ball.ball_lotto, span.ball_lotto")
                if re.fullmatch(r"\d{1,2}", s.get_text(strip=True) or "")
                and "special" not in (s.get("class") or [])
            ]
            sp_el = box_like.select_one("span.special, span.ball_lotto_2")
            if sp_el and re.fullmatch(r"\d{1,2}", sp_el.get_text(strip=True) or ""):
                special = int(sp_el.get_text(strip=True))
            else:
                parts = [int(x) for x in re.findall(r"\d{1,2}", box_like.get_text(" ", strip=True))]
                if len(parts) < 6:
                    continue
                mains, special = parts[:5], parts[5]
            if len(mains) < 5:
                continue
            draw_id = latest_id - i
            if draw_id < 1:
                break
            out.append(
                {
                    "date": date_text,
                    "id": draw_id,
                    "result": _format_result(mains[:5], special),
                }
            )
        if out:
            break
    return out


def _day_page_url(date_text: str) -> str:
    n = normalize_date_text(date_text)
    dd, mm, yyyy = n.split("/")
    return DAY_PAGE_TMPL.format(dd=dd, mm=mm, yyyy=yyyy)


def parse_live_lt535_js(text: str) -> dict | None:
    """
    Parse livekqxs/xstt/js/LT535.js:
      xsdt[11]={..., "period":973, "date":"28-10-2026", "time":"13",
                "numbers":["02","17","20","24","32","06"], ...}
    Tra ve {date, id, result} neu co du 6 so; None neu chua quay / thieu data.
    """
    raw = str(text or "")
    # Chi lay object trong xsdt[11]={...}
    m = re.search(r"xsdt\s*\[\s*11\s*\]\s*=\s*(\{.*?\})\s*;?", raw, re.S)
    blob = m.group(1) if m else raw
    pm = re.search(r'"period"\s*:\s*(\d+)', blob)
    dm = re.search(r'"date"\s*:\s*"([^"]+)"', blob)
    nm = re.search(r'"numbers"\s*:\s*\[(.*?)\]', blob, re.S)
    if not pm or not dm or not nm:
        return None
    draw_id = int(pm.group(1))
    date_text = normalize_date_text(dm.group(1).replace("-", "/"))
    nums = [int(x) for x in re.findall(r'"?(\d{1,2})"?', nm.group(1))]
    if draw_id < 1 or not date_text or len(nums) < 6:
        return None
    return {
        "date": date_text,
        "id": draw_id,
        "result": _format_result(nums[:5], nums[5]),
    }


def fetch_live_lt535_row(live_js_url: str, referer: str | None = None) -> dict | None:
    """Hard-refresh live JS — thuong co ky moi truoc khi HTML truc tiep cap nhat."""
    try:
        html = fetch_html(live_js_url, referer=referer or DEFAULT_MINHCHINH_URL)
    except Exception as e:
        print(f"Canh bao: live JS loi ({e}).", flush=True)
        return None
    return parse_live_lt535_js(html)


def _fetch_all_from_minhchinh(page_url: str, min_id_exclusive: int) -> list[dict]:
    """
    Hard-refresh trang truc tiep + (neu thieu) trang theo ngay lui dan
    den khi gom du moi ky > min_id_exclusive.
    """
    html = fetch_html(page_url)
    by_id: dict[int, dict] = {}

    def absorb(rows: list[dict]) -> None:
        for r in rows:
            rid = int(r["id"])
            if rid <= min_id_exclusive:
                continue
            by_id[rid] = dict(r)

    primary = parse_draws_from_minhchinh_html(html)
    absorb(primary)
    latest_id = max((r["id"] for r in primary), default=None)
    if latest_id is None:
        # Fallback: #ky trong HTML + bang 15 ky
        m = re.search(r"k[ỳy]\s*<[^>]*>\s*#\s*(\d+)|#\s*(\d{3,5})", html, re.I)
        if m:
            latest_id = int(m.group(1) or m.group(2))
    absorb(parse_history_table_minhchinh(html, latest_id))

    if latest_id is None:
        raise RuntimeError("Khong tim thay #ky tren trang Minh Chinh.")

    # Con thieu ky (lech > ~15) → doi trang theo ngay, lui toi da 40 ngay.
    missing = [i for i in range(min_id_exclusive + 1, latest_id + 1) if i not in by_id]
    if missing:
        # Bat dau tu ngay cua ky moi nhat (neu co), khong thi hom nay.
        seed_date = ""
        if latest_id in by_id:
            seed_date = by_id[latest_id]["date"]
        if not seed_date:
            # Tim ngay trong .ngay gan #latest
            soup = BeautifulSoup(html, "html.parser")
            for ngay in soup.select("div.ngay"):
                did, dtxt, _ = _parse_ky_date_session(ngay.get_text(" ", strip=True))
                if did == latest_id and dtxt:
                    seed_date = dtxt
                    break
        if not seed_date:
            seed_date = datetime.now().strftime("%d/%m/%Y")

        cur = normalize_date_text(seed_date)
        referer = page_url
        for _ in range(40):
            if not missing:
                break
            if not cur:
                break
            day_url = _day_page_url(cur)
            try:
                day_html = fetch_html(day_url, referer=referer)
            except Exception:
                # Het trang / 404 → lui 1 ngay
                cur = (
                    datetime.strptime(cur, "%d/%m/%Y") - timedelta(days=1)
                ).strftime("%d/%m/%Y")
                continue
            absorb(parse_draws_from_minhchinh_html(day_html))
            missing = [i for i in range(min_id_exclusive + 1, latest_id + 1) if i not in by_id]
            referer = day_url
            cur = (
                datetime.strptime(cur, "%d/%m/%Y") - timedelta(days=1)
            ).strftime("%d/%m/%Y")
            time.sleep(0.15)

    if not by_id and latest_id > min_id_exclusive:
        raise RuntimeError(
            f"Co ky moi (#{latest_id}) nhung khong parse duoc so tren Minh Chinh."
        )
    return sorted(by_id.values(), key=lambda r: r["id"])


def _fetch_new_rows(
    min_id_exclusive: int,
    page_url: str,
    live_js_url: str,
) -> tuple[list[dict], str]:
    """
    Uu tien live JS; thieu / loi → HTML Minh Chinh.
    Tra (rows, source_label_for_log).
    """
    by_id: dict[int, dict] = {}
    sources: list[str] = []

    live = fetch_live_lt535_row(live_js_url, referer=page_url)
    if live and live["id"] > min_id_exclusive:
        by_id[live["id"]] = dict(live)
        sources.append("live-js")
        print(
            f"Live JS: ky #{live['id']:05d} {live['date']} {live['result']}",
            flush=True,
        )
    elif live:
        print(
            f"Live JS: ky #{live['id']:05d} (khong moi hon max {min_id_exclusive:05d}).",
            flush=True,
        )
    else:
        print("Live JS: chua co numbers / khong parse duoc.", flush=True)

    # HTML fallback: luon thu neu con thieu ky (lech >1) hoac live chua co.
    need_html = False
    if not by_id:
        need_html = True
    else:
        latest = max(by_id)
        missing = [
            i for i in range(min_id_exclusive + 1, latest + 1) if i not in by_id
        ]
        need_html = bool(missing)

    if need_html:
        try:
            html_rows = _fetch_all_from_minhchinh(page_url, min_id_exclusive)
            for r in html_rows:
                rid = int(r["id"])
                if rid <= min_id_exclusive:
                    continue
                # Live JS uu tien hon HTML neu trung id.
                if rid not in by_id:
                    by_id[rid] = dict(r)
            if html_rows:
                sources.append("html-minhchinh")
        except Exception as e:
            print(
                f"Canh bao: HTML Minh Chinh loi (se thu lai o lan sau). {e}",
                flush=True,
            )

    rows = sorted(by_id.values(), key=lambda r: r["id"])
    label = "+".join(sources) if sources else live_js_url
    return rows, label


def main() -> int:
    args = parse_args()
    json_path = args.json
    if not os.path.isabs(json_path):
        json_path = os.path.join(os.getcwd(), json_path)

    existing, max_id = read_real_rows_from_data_json(json_path)

    new_rows, src_label = _fetch_new_rows(max_id, args.url, args.live_js)
    if not new_rows:
        print(f"Khong co ky moi (max_id hien tai: {max_id:05d}).")
        print(args.live_js, flush=True)
        print(args.url, flush=True)
        return 0

    by_id = {r["id"]: dict(r) for r in existing}
    for r in new_rows:
        by_id[r["id"]] = dict(r)
    combined = sorted(by_id.values(), key=lambda r: r["id"])
    shaped = rows_to_json_file_shape(combined)
    final_out = append_trailing_placeholder(shaped)
    write_data_json(json_path, final_out)
    ids_added = ", ".join(f"{r['id']:05d}" for r in new_rows)
    print(
        f"Them {len(new_rows)} ky: {ids_added}. "
        f"Ghi {json_path} ({len(final_out)} dong, co hang cuoi result rong). "
        f"Nguon: {src_label}."
    )
    print(args.live_js, flush=True)
    print(args.url, flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
