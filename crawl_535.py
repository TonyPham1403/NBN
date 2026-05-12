"""
Dong bo tat ca ky xo so 535 (Max 3D+ / Lotto 21h) moi hon data.json hien co.
Ghi truc tiep proj/data.json, giu dung hang cuoi result rong nhu convert_data_to_json.js.
"""
from __future__ import annotations

import argparse
import json
import os
import re
from datetime import datetime, timedelta

import requests
from bs4 import BeautifulSoup


URL = "https://www.vietlott.vn/vi/trung-thuong/ket-qua-trung-thuong/winning-number-535#top"
HOME = "https://www.vietlott.vn/vi/"
AJAX_URL = (
    "https://www.vietlott.vn/ajaxpro/Vietlott.PlugIn.WebParts.Game535CompareWebPart,"
    "Vietlott.PlugIn.WebParts.ashx"
)
DEFAULT_AJAX_KEY = "64bdd318"
IMPERSONATE = "chrome124"
DEFAULT_JSONL_URLS = (
    "https://raw.githubusercontent.com/vietvudanh/vietlott-data/master/data/power535.jsonl",
    "https://cdn.jsdelivr.net/gh/vietvudanh/vietlott-data@master/data/power535.jsonl",
)
DEFAULT_XSKT_VIETLOTT_URL = "https://xskt.com.vn/ket-qua-vietlott"


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


def _ajax_key_from_page_html(html: str) -> str | None:
    m = re.search(
        r"Game535CompareWebPart\.ServerSideDrawResult\([^,]+,\s*'([0-9a-f]{8})'",
        html,
        re.I,
    )
    return m.group(1) if m else None


def _ajax_request_body(ajax_key: str, page_index: int = 0) -> dict:
    return {
        "ORenderInfo": {
            "ExtraParam1": "",
            "ExtraParam2": "",
            "ExtraParam3": "",
            "FullPageAlias": "",
            "IsPageDesign": False,
            "OrgPageAlias": "",
            "PageAlias": "",
            "RefKey": "",
            "SiteAlias": "main.vi",
            "SiteId": "main.frontend.vi",
            "SiteLang": "vi",
            "SiteName": "Vietlott",
            "SiteURL": "",
            "System": 1,
            "UserSessionId": "",
            "WebPage": "",
        },
        "Key": ajax_key,
        "GameDrawId": "",
        "ArrayNumbers": [["" for _ in range(35)] for _ in range(5)],
        "CheckMulti": False,
        "PageIndex": page_index,
    }


def _ajaxpro_headers() -> dict[str, str]:
    return {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
        ),
        "Accept": "*/*",
        "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
        "Content-Type": "text/plain; charset=utf-8",
        "X-AjaxPro-Method": "ServerSideDrawResult",
        "Origin": "https://www.vietlott.vn",
        "Referer": "https://www.vietlott.vn/vi/trung-thuong/ket-qua-trung-thuong/winning-number-535",
    }


def _fetch_html_via_ajaxpro(page_url: str, ajax_key: str, page_index: int = 0) -> str:
    page_url = page_url.split("#", 1)[0]
    try:
        from curl_cffi import requests as curl_requests

        session = curl_requests.Session()
        session.get(HOME, impersonate=IMPERSONATE, timeout=30, headers=_headers(None))
        win_html = None
        try:
            win = session.get(
                page_url,
                impersonate=IMPERSONATE,
                timeout=30,
                headers=_headers(HOME),
            )
            if win.ok:
                win_html = win.text
        except Exception:
            win_html = None
        key = (_ajax_key_from_page_html(win_html) if win_html else None) or ajax_key
        body = _ajax_request_body(key, page_index)
        resp = session.post(
            AJAX_URL,
            data=json.dumps(body, ensure_ascii=False),
            headers=_ajaxpro_headers(),
            impersonate=IMPERSONATE,
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        val = data.get("value") or {}
        if val.get("Error"):
            msg = val.get("InfoMessage") or str(val)
            raise RuntimeError(f"AjaxPro loi: {msg}")
        html = val.get("HtmlContent") or ""
        if not html.strip():
            raise RuntimeError("AjaxPro tra ve HtmlContent rong.")
        return html
    except ImportError:
        session = requests.Session()
        session.get(HOME, timeout=30, headers=_headers(None))
        win_html = None
        try:
            win = session.get(page_url, timeout=30, headers=_headers(HOME))
            if win.ok:
                win_html = win.text
        except Exception:
            win_html = None
        key = (_ajax_key_from_page_html(win_html) if win_html else None) or ajax_key
        body = _ajax_request_body(key, page_index)
        resp = session.post(
            AJAX_URL,
            data=json.dumps(body, ensure_ascii=False),
            headers=_ajaxpro_headers(),
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        val = data.get("value") or {}
        if val.get("Error"):
            msg = val.get("InfoMessage") or str(val)
            raise RuntimeError(f"AjaxPro loi: {msg}")
        html = val.get("HtmlContent") or ""
        if not html.strip():
            raise RuntimeError("AjaxPro tra ve HtmlContent rong.")
        return html


def _fetch_page_html(page_url: str) -> str:
    page_url = page_url.split("#", 1)[0]
    try:
        from curl_cffi import requests as curl_requests

        session = curl_requests.Session()
        session.get(HOME, impersonate=IMPERSONATE, timeout=30, headers=_headers(None))
        resp = session.get(
            page_url,
            impersonate=IMPERSONATE,
            timeout=30,
            headers=_headers(HOME),
        )
        resp.raise_for_status()
        return resp.text
    except ImportError:
        session = requests.Session()
        session.get(HOME, timeout=30, headers=_headers(None))
        resp = session.get(page_url, timeout=30, headers=_headers(HOME))
        resp.raise_for_status()
        return resp.text


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Dong bo ky moi vao data.json (khong sua xlsm).",
    )
    p.add_argument(
        "--json",
        default=os.environ.get("VIETLOTT_535_DATA_JSON", "data.json"),
        help="Duong toi data.json (mac dinh data.json trong cwd).",
    )
    p.add_argument("--url", default=URL, help="URL trang ket qua Vietlott.")
    p.add_argument(
        "--ajax-key",
        default=os.environ.get("VIETLOTT_535_AJAX_KEY", DEFAULT_AJAX_KEY),
        help="AjaxPro key (hoac VIETLOTT_535_AJAX_KEY).",
    )
    p.add_argument(
        "--max-ajax-pages",
        type=int,
        default=int(os.environ.get("VIETLOTT_535_MAX_AJAX_PAGES", "15")),
        help="So trang AjaxPro toi da de lay du cac ky gan day.",
    )
    return p.parse_args()


def _github_actions() -> bool:
    return os.environ.get("GITHUB_ACTIONS", "").lower() == "true"


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


def _parse_row_element_to_record(row) -> dict:
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


def _parse_all_vietlott_table_rows(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    out: list[dict] = []
    for sel in (
        "#divResultContent table tbody tr",
        "table.table-hover tbody tr",
        "table tbody tr",
    ):
        found = soup.select(sel)
        if not found:
            continue
        for row in found:
            try:
                out.append(_parse_row_element_to_record(row))
            except Exception:
                continue
        break
    return out


def _mirror_jsonl_urls() -> list[str]:
    out: list[str] = []
    env_u = os.environ.get("VIETLOTT_535_JSONL_URL", "").strip()
    if env_u:
        out.append(env_u)
    for u in DEFAULT_JSONL_URLS:
        if u not in out:
            out.append(u)
    return out


def _record_from_jsonl_row(row: dict) -> dict:
    raw_date = str(row.get("date", "")).strip()
    dt = datetime.strptime(raw_date, "%Y-%m-%d")
    date_text = dt.strftime("%d/%m/%Y")
    id_num = int(str(row.get("id", "")).strip())
    nums = row.get("result")
    if not isinstance(nums, list) or len(nums) != 6:
        raise RuntimeError(f"JSONL result khong hop le: {nums!r}")
    main = ",".join(str(int(x)) for x in nums[:5])
    special = str(int(nums[5]))
    return {"date": date_text, "id": id_num, "result": f"{main}|{special}"}


def _fetch_all_from_mirror_jsonl() -> list[dict]:
    ua = _headers(None)["User-Agent"]
    last_err: Exception | None = None
    for src in _mirror_jsonl_urls():
        try:
            resp = requests.get(src, timeout=60, headers={"User-Agent": ua})
            resp.raise_for_status()
            rows: list[dict] = []
            for line in resp.text.splitlines():
                line = line.strip()
                if not line:
                    continue
                rows.append(_record_from_jsonl_row(json.loads(line)))
            if not rows:
                raise RuntimeError("JSONL rong.")
            return rows
        except Exception as e:
            last_err = e
    raise last_err if last_err else RuntimeError("Khong doc duoc JSONL mirror.")


def _xskt_vietlott_url() -> str:
    return os.environ.get("VIETLOTT_535_XSKT_URL", "").strip() or DEFAULT_XSKT_VIETLOTT_URL


def _fetch_all_from_xskt_vietlott() -> list[dict]:
    ua = _headers(None)["User-Agent"]
    resp = requests.get(_xskt_vietlott_url(), timeout=60, headers={"User-Agent": ua})
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    out: list[dict] = []
    for table in soup.select("table.result"):
        kmt = table.select_one("td.kmt")
        if not kmt or "21h" not in kmt.get_text():
            continue
        mid = re.search(r"#(\d{5})", kmt.get_text())
        if not mid:
            continue
        draw_id = int(mid.group(1))
        em = table.select_one("td.megaresult em")
        if not em:
            continue
        spans = em.find_all("span")
        main_text = em.get_text(" ", strip=True)
        for sp in spans:
            main_text = main_text.replace(sp.get_text(strip=True), "").strip()
        main_nums = [p for p in re.split(r"\s+", main_text.strip()) if p.isdigit()]
        if not spans:
            continue
        special = spans[-1].get_text(strip=True)
        if len(main_nums) != 5 or not special.isdigit():
            continue
        h2 = table.find_previous("h2")
        date_text = ""
        if h2:
            link = h2.select_one('a[href*="xslotto-21h/ngay-"]')
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
        raise RuntimeError("Khong tim thay bang Lotto 21h tren xskt.")
    return out


def _fetch_all_vietlott_ajax_pages(
    url: str,
    ajax_key: str,
    max_pages: int,
    min_id_exclusive: int,
) -> list[dict]:
    by_id: dict[int, dict] = {}
    prev_sig: tuple[int, ...] | None = None
    for page in range(max_pages):
        try:
            html = _fetch_html_via_ajaxpro(url, ajax_key, page)
        except Exception:
            break
        rows = _parse_all_vietlott_table_rows(html)
        if not rows:
            break
        if not any(r["id"] > min_id_exclusive for r in rows):
            break
        sig = tuple(sorted(r["id"] for r in rows))
        if sig == prev_sig:
            break
        prev_sig = sig
        for r in rows:
            if r["id"] > min_id_exclusive:
                by_id[r["id"]] = r
    return sorted(by_id.values(), key=lambda r: r["id"])


def _merge_remote_rows(
    min_id_exclusive: int,
    url: str,
    ajax_key: str,
    max_pages: int,
    skip_official: bool,
    skip_xskt: bool,
) -> list[dict]:
    """
    Gom tu nhieu nguon, chi giu id > min_id_exclusive.
    Thu tu ghi de: JSONL -> GET trang -> tung trang AjaxPro -> xskt (uu tien xskt neu trung id).
    """
    merged: dict[int, dict] = {}

    def put_many(rows: list[dict]) -> None:
        for r in rows:
            if r["id"] <= min_id_exclusive:
                continue
            merged[r["id"]] = dict(r)

    errors: list[str] = []

    try:
        put_many(_fetch_all_from_mirror_jsonl())
    except Exception as e:
        errors.append(f"jsonl: {e}")

    if not skip_official:
        try:
            put_many(_parse_all_vietlott_table_rows(_fetch_page_html(url)))
        except Exception as e:
            errors.append(f"GET: {e}")
        try:
            put_many(_fetch_all_vietlott_ajax_pages(url, ajax_key, max_pages, min_id_exclusive))
        except Exception as e:
            errors.append(f"AjaxPro: {e}")

    if not skip_xskt:
        try:
            put_many(_fetch_all_from_xskt_vietlott())
        except Exception as e:
            errors.append(f"xskt: {e}")

    out = sorted(merged.values(), key=lambda r: r["id"])
    if not out and errors:
        raise RuntimeError("Khong lay duoc ky moi. " + " | ".join(errors))
    return out


def main() -> int:
    args = parse_args()
    json_path = args.json
    if not os.path.isabs(json_path):
        json_path = os.path.join(os.getcwd(), json_path)

    existing, max_id = read_real_rows_from_data_json(json_path)

    skip_official = os.environ.get("VIETLOTT_535_SKIP_OFFICIAL", "").lower() in (
        "1",
        "true",
        "yes",
    )
    if _github_actions() and os.environ.get("VIETLOTT_535_TRY_OFFICIAL_ON_CI", "").lower() not in (
        "1",
        "true",
        "yes",
    ):
        skip_official = True

    skip_xskt = os.environ.get("VIETLOTT_535_SKIP_XSKT", "").lower() in (
        "1",
        "true",
        "yes",
    )

    new_rows = _merge_remote_rows(
        max_id,
        args.url,
        args.ajax_key,
        args.max_ajax_pages,
        skip_official=skip_official,
        skip_xskt=skip_xskt,
    )
    new_rows = [r for r in new_rows if r["id"] > max_id]
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
