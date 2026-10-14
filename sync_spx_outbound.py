"""
Sync SPX Outbound (High Value + DG Type 4) -> Google Sheet
Bản chuyển đổi từ Google Apps Script sang Python.

Yêu cầu:
    pip install requests gspread google-auth tzdata

Biến môi trường cần thiết:
    SPREADSHEET_ID              - ID của Google Sheet
    GOOGLE_SERVICE_ACCOUNT_JSON - đường dẫn tới file service-account.json
                                   (hoặc nội dung JSON trực tiếp, script tự nhận diện)

Luồng chạy:
    1. Gọi high_value=1 lần lượt với status 2, 10, 11, 4, 5
    2. Gọi dg_type=4   lần lượt với status 2, 10, 11, 4, 5
    3. Ghép dữ liệu, trùng to_number thì bỏ (giữ dòng lấy được trước)

Lưu ý múi giờ:
    Mọi thời gian hiển thị đều dùng múi giờ Asia/Ho_Chi_Minh (UTC+7),
    không phụ thuộc múi giờ của máy chạy (GitHub Actions runner mặc định UTC).
"""

import json
import os
import random
import time
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import gspread
import requests
from google.oauth2.service_account import Credentials

# ============================== CONFIG ==============================
TZ = ZoneInfo("Asia/Ho_Chi_Minh")   # múi giờ Việt Nam (UTC+7)

SPX_BASE_URL = "https://spx.shopee.vn/api/in-station/general_to/outbound/search"
SHEET_NAME = "HV pack"
COOKIE_SHEET_NAME = "Cookie"

PAGE_SIZE = 1000
MAX_PAGES = 7
SLEEP_SEC = 2.0                 # nghỉ giữa các page CÙNG status
STATUS_SLEEP_SEC = 3.0          # nghỉ giữa các lượt gọi khác nhau
MAX_RETRY = 10
MAX_BACKOFF_SEC = 30.0
RETRY_FAILED_STATUSES = True    # thử lại 1 vòng riêng cho tác vụ bị lỗi, sau khi chạy hết
DAYS_BACK = 6
STATUS_FILTERS = [2, 10, 11, 4, 5]

# Thứ tự chạy: HV trước, DG sau. Dòng trùng to_number ở lượt sau sẽ bị bỏ.
FILTER_JOBS = [
    {"name": "HV=1", "params": {"high_value": 1}},
    {"name": "DG=4", "params": {"dg_type": 4}},
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

HEADERS_TEMPLATE = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
    "Origin": "https://spx.shopee.vn",
    "Referer": "https://spx.shopee.vn/#/general-to-management",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "X-Requested-With": "XMLHttpRequest",
    "sec-ch-ua": '"Chromium";v="128", "Not;A=Brand";v="24", "Google Chrome";v="128"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "Connection": "keep-alive",
}

HEADER_ROW = [
    "TO Number", "Driver", "TO Pack", "High Value", "Dangerous Goods",
    "Packing Method", "Receiver Type", "Receiver", "Current Station",
    "Quantity", "Weight(kg)", "Direction", "Exception Tag", "Remark",
    "LH Trip Number", "Length(cm)", "Width(cm)", "Height(cm)", "Operator",
    "Status", "Sender", "Create time", "Complete time", "Offline Mode",
    "Staging Area", "Action", "DG Raw",
]


def log(msg):
    ts = datetime.now(TZ).strftime("%H:%M:%S")
    print(f"{ts}  {msg}", flush=True)


# ============================== GOOGLE SHEETS ==============================
def get_gspread_client():
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        raise RuntimeError("Thiếu biến môi trường GOOGLE_SERVICE_ACCOUNT_JSON.")

    if raw.startswith("{"):
        info = json.loads(raw)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(raw, scopes=SCOPES)

    return gspread.authorize(creds)


def get_spreadsheet(gc):
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    if not spreadsheet_id:
        raise RuntimeError("Thiếu biến môi trường SPREADSHEET_ID.")
    return gc.open_by_key(spreadsheet_id)


def get_or_create_sheet(ss, name, rows=1000, cols=30):
    try:
        return ss.worksheet(name)
    except gspread.WorksheetNotFound:
        return ss.add_worksheet(title=name, rows=rows, cols=cols)


def load_cookie(ss):
    try:
        sh = ss.worksheet(COOKIE_SHEET_NAME)
    except gspread.WorksheetNotFound:
        raise RuntimeError(f'Không tìm thấy sheet "{COOKIE_SHEET_NAME}".')
    cookie = str(sh.acell("B1").value or "").strip()
    if not cookie:
        raise RuntimeError(f"Thiếu cookie trong {COOKIE_SHEET_NAME}!B1.")
    return cookie


def reset_outbound_sheet_layout(sh):
    sh.clear()
    sh.update("A1", [HEADER_ROW])
    sh.freeze(rows=1)


def format_data_range(sh, last_row):
    if last_row <= 1:
        return
    sheet_id = sh.id
    requests_body = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": last_row,
                        "startColumnIndex": 10,  # cột K (Weight)
                        "endColumnIndex": 11,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {"type": "NUMBER", "pattern": "0.000"}
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": last_row,
                        "startColumnIndex": 21,  # cột V-W (Create/Complete time)
                        "endColumnIndex": 23,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "DATE_TIME",
                                "pattern": "dd/MM/yyyy HH:mm:ss",
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            },
        ]
    }
    sh.spreadsheet.batch_update(requests_body)


def set_last_update(ss):
    sh = ss.worksheet(COOKIE_SHEET_NAME)
    now_str = datetime.now(TZ).strftime("%d/%m/%Y %H:%M:%S")
    sh.update_acell("B4", now_str)


# ============================== SPX API ==============================
def get_last_n_days_time_range(days):
    end = datetime.now(timezone.utc)
    start = end - timedelta(days=days)
    return int(start.timestamp()), int(end.timestamp())


def fetch_outbound_with_retry(url, cookie, session):
    headers = dict(HEADERS_TEMPLATE)
    headers["Cookie"] = cookie

    last_resp = None
    for attempt in range(1, MAX_RETRY + 1):
        try:
            resp = session.get(url, headers=headers, timeout=30)
            last_resp = resp
            if resp.status_code < 400:
                return resp
            log(f"Retry lần {attempt} vì HTTP {resp.status_code}: {resp.text[:300]}")
            backoff = min(MAX_BACKOFF_SEC, 1.5 * (1.6 ** attempt)) + random.uniform(0, 0.8)
            time.sleep(backoff)
        except requests.RequestException as e:
            log(f"Exception retry lần {attempt}: {e}")
            time.sleep(1.5 * attempt)

    return last_resp


def fetch_outbound_list_filtered(
    page, count, ctime_start, ctime_end, complete_start, complete_end,
    extra_params, status_filter, cookie, session
):
    query = (
        f"pageno={page}&count={count}"
        f"&ctime={ctime_start},{ctime_end}"
        f"&complete_time={complete_start},{complete_end}"
        f"&status={status_filter}"
    )
    for k, v in extra_params.items():
        query += f"&{k}={v}"

    url = f"{SPX_BASE_URL}?{query}"
    log(f"Đang gọi {extra_params} | Status {status_filter} - page {page}")

    resp = fetch_outbound_with_retry(url, cookie, session)
    if resp is None:
        raise RuntimeError(f"Không nhận được phản hồi cho status {status_filter} page {page}")

    code = resp.status_code
    text = resp.text
    log(f"HTTP code {extra_params} | Status {status_filter} - page {page}: {code}")

    if code == 401:
        raise RuntimeError("401 Unauthorized - Cookie sai hoặc hết hạn")
    if code != 200:
        log(f"Response lỗi: {text[:500]}")
        raise RuntimeError(f"HTTP {code}: {text[:500]}")

    data = resp.json()
    if data.get("retcode") != 0:
        raise RuntimeError(
            f"retcode={data.get('retcode')}, message={data.get('message', '')}, "
            f"response={text[:500]}"
        )

    d = data.get("data") or {}
    return {"total": d.get("total", 0), "list": d.get("list", [])}


# ============================== MAPPING ==============================
def to_kg(w):
    try:
        return round(float(w) / 1000, 3)
    except (TypeError, ValueError):
        return ""


def to_date_str(sec):
    """Đổi epoch giây -> chuỗi dd/MM/yyyy HH:mm:ss theo giờ Việt Nam."""
    try:
        n = float(sec)
        if n > 0:
            return datetime.fromtimestamp(n, tz=TZ).strftime("%d/%m/%Y %H:%M:%S")
    except (TypeError, ValueError, OverflowError, OSError):
        pass
    return ""


def map_receiver_type(v):
    try:
        n = int(v)
    except (TypeError, ValueError):
        return v or ""
    return {1: "Station", 2: "Hub"}.get(n, str(n))


def map_high_value(v):
    try:
        n = int(v)
    except (TypeError, ValueError):
        return v or ""
    return {1: "Y", 2: "N"}.get(n, str(n))


def map_packing_method(v):
    try:
        n = int(v)
    except (TypeError, ValueError):
        return v or ""
    return "Individual Packing" if n == 0 else str(n)


def map_dg_from_dg_type(dg_type):
    if not isinstance(dg_type, list) or not dg_type:
        return ""
    labels = []
    for v in dg_type:
        try:
            n = int(v)
        except (TypeError, ValueError):
            continue
        if n == 1:
            continue
        labels.append({2: "DG Type A", 3: "DG Type B", 4: "DG Type C"}.get(n, f"DG Type {n}"))
    return ", ".join(labels)


def map_row(it):
    dg_type = it.get("dg_type") or []
    operator = str(it.get("assign_operator") or it.get("operator") or "").strip()
    return [
        it.get("to_number", ""),
        it.get("driver_name", ""),
        it.get("pack_name", ""),
        map_high_value(it.get("high_value")),
        map_dg_from_dg_type(dg_type),
        map_packing_method(it.get("packing_method")),
        map_receiver_type(it.get("receiver_type")),
        it.get("receiver") or it.get("dest_station_name") or "",
        it.get("current_station_name", ""),
        it.get("quantity", "") if it.get("quantity") is not None else "",
        to_kg(it.get("weight")),
        it.get("display_direction") or it.get("transfer_direction") or "",
        "",
        it.get("remark", ""),
        it.get("lh_trip_number", ""),
        it.get("length", "") if it.get("length") is not None else "",
        it.get("width", "") if it.get("width") is not None else "",
        it.get("height", "") if it.get("height") is not None else "",
        operator,
        it.get("status", ""),
        it.get("sender", ""),
        to_date_str(it.get("ctime")),
        to_date_str(it.get("complete_time")),
        it.get("initialize_source", "") if it.get("initialize_source") is not None else "",
        it.get("staging_area_id", "") if it.get("staging_area_id") is not None else "",
        "https://spx.shopee.vn/#/general-to-management",
        json.dumps(dg_type, ensure_ascii=False),
    ]


# ============================== MAIN SYNC LOGIC ==============================
def fetch_all_pages_for_status(
    job_name, extra_params, status_filter,
    ctime_start, ctime_end, complete_start, complete_end,
    cookie, session, all_rows, seen_to_numbers
):
    log(f"===== [{job_name}] Bắt đầu lấy Status = {status_filter} =====")
    fetched = 0
    total = 0
    failed = False

    for page in range(1, MAX_PAGES + 1):
        try:
            result = fetch_outbound_list_filtered(
                page, PAGE_SIZE, ctime_start, ctime_end, complete_start, complete_end,
                extra_params, status_filter, cookie, session,
            )
        except Exception as e:
            log(f"Lỗi [{job_name}] Status {status_filter} - page {page}: {e}")
            failed = True
            break

        lst = result.get("list", [])
        total = result.get("total") or total
        log(f"[{job_name}] Status {status_filter} | Page {page}: {len(lst)} | api_total={result.get('total', 0)}")

        if not lst:
            break

        fetched += len(lst)
        for it in lst:
            to_number = str(it.get("to_number") or "").strip()
            if not to_number or to_number not in seen_to_numbers:
                all_rows.append(map_row(it))
                if to_number:
                    seen_to_numbers[to_number] = True

        if len(lst) < PAGE_SIZE:
            break
        time.sleep(SLEEP_SEC)

    return {"fetched": fetched, "total": total, "failed": failed}


def sync_spx_outbound_hvY_to_ob():
    gc = get_gspread_client()
    ss = get_spreadsheet(gc)
    sh = get_or_create_sheet(ss, SHEET_NAME)
    cookie = load_cookie(ss)

    complete_start, complete_end = get_last_n_days_time_range(DAYS_BACK)
    ctime_start = complete_start - 90 * 24 * 3600
    ctime_end = complete_end

    log(f"ctime={ctime_start},{ctime_end}")
    log(f"complete_time={complete_start},{complete_end}")
    log(f"status={','.join(map(str, STATUS_FILTERS))}")
    log(f"jobs={[j['name'] for j in FILTER_JOBS]}")

    all_rows = []
    total_fetched = 0
    seen_to_numbers = {}
    failed_tasks = []   # danh sách (job, status) bị lỗi

    session = requests.Session()

    # Thứ tự: HV (5 status) -> DG (5 status)
    tasks = [(job, st) for job in FILTER_JOBS for st in STATUS_FILTERS]

    for i, (job, status_filter) in enumerate(tasks):
        r = fetch_all_pages_for_status(
            job["name"], job["params"], status_filter,
            ctime_start, ctime_end, complete_start, complete_end,
            cookie, session, all_rows, seen_to_numbers,
        )
        total_fetched += r["fetched"]
        if r["failed"]:
            failed_tasks.append((job, status_filter))

        if i < len(tasks) - 1:
            time.sleep(STATUS_SLEEP_SEC + random.uniform(0, 1.0))

    if RETRY_FAILED_STATUSES and failed_tasks:
        log(f"Sẽ thử lại: {[(j['name'], s) for j, s in failed_tasks]}")
        time.sleep(15 + random.uniform(0, 5))
        still_failed = []
        for i, (job, status_filter) in enumerate(failed_tasks):
            r = fetch_all_pages_for_status(
                job["name"], job["params"], status_filter,
                ctime_start, ctime_end, complete_start, complete_end,
                cookie, session, all_rows, seen_to_numbers,
            )
            total_fetched += r["fetched"]
            if r["failed"]:
                still_failed.append((job["name"], status_filter))
            if i < len(failed_tasks) - 1:
                time.sleep(STATUS_SLEEP_SEC + random.uniform(0, 1.0))
        if still_failed:
            log(f"Vẫn lỗi sau khi thử lại: {still_failed}")

    reset_outbound_sheet_layout(sh)

    if not all_rows:
        sh.update_acell("A2", "Không có dữ liệu")
        set_last_update(ss)
        log("Không có dữ liệu để ghi.")
        return

    # Ghi theo batch 500 dòng như bản Apps Script gốc
    total_written = 0
    for i in range(0, len(all_rows), 500):
        chunk = all_rows[i:i + 500]
        start_row = 2 + total_written
        end_row = start_row + len(chunk) - 1
        sh.update(f"A{start_row}:AA{end_row}", chunk, value_input_option="USER_ENTERED")
        total_written += len(chunk)

    format_data_range(sh, total_written + 1)
    set_last_update(ss)

    log(
        f"Done: fetched={total_fetched} | unique_written={total_written} | "
        f"jobs={','.join(j['name'] for j in FILTER_JOBS)} | "
        f"statuses={','.join(map(str, STATUS_FILTERS))}"
    )


if __name__ == "__main__":
    sync_spx_outbound_hvY_to_ob()
