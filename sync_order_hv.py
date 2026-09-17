"""
sync_order_hv.py
-----------------
Port của Apps Script `fetchSpxOrderTracking()` sang Python.
Dùng CHUNG Google Sheet ID và Service Account với script sync_spx_outbound.py
đã setup trước đó — chỉ khác payload gọi API và sheet đích ("Order HV").

Cách hoạt động:
  1. Đọc cookie mới nhất từ sheet "Cookie" ô B1 (giống hệt script cũ).
  2. Gọi API tracking_list/search với payload cố định (high_value=1,
     order_status="33,60,8,58", current_station_ids="1959"), phân trang
     tới khi hết dữ liệu hoặc chạm MAX_PAGES.
  3. Map field -> cột hiển thị y hệt COLUMNS trong Apps Script gốc.
  4. Xoá sạch sheet "Order HV" rồi ghi lại toàn bộ dữ liệu mới.

ENV cần có (đã có sẵn nếu bạn dùng chung repo/secrets với script cũ):
  SPREADSHEET_ID              - ID của Google Sheet (giống script cũ)
  GOOGLE_SERVICE_ACCOUNT_JSON - Nội dung JSON của service account (dạng chuỗi)
"""

import os
import json
import re
import time
from datetime import datetime, timezone, timedelta

import requests
import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# CONFIG — tương ứng CONFIG trong Apps Script
# ============================================================
CONFIG = {
    "COOKIE_SHEET": "Cookie",
    "COOKIE_CELL": "B1",
    "OUTPUT_SHEET": "Order HV",
    "API_URL": "https://spx.shopee.vn/api/fleet_order/order/tracking_list/search",

    "PAGE_SIZE": 1000,
    "MAX_PAGES": 5,
    "SLEEP_SEC": 2.0,      # SLEEP_MS = 2000
    "MAX_RETRY": 6,
}

# (key trong response, tên cột hiển thị) — y hệt COLUMNS gốc
COLUMNS = [
    ("shipment_id", "SPX Tracking Number"),
    ("sls_tracking_number", "SLS Tracking Number"),
    ("third_party_tracking_num", "3PL Tracking Number"),
    ("installation_tn", "Installation TN"),
    ("shopee_order_sn", "Order SN"),
    ("__current_to_number", "Current TO Number"),
    ("order_account", "Order Account"),
    ("cod_status", "COD Status"),
    ("order_direction", "Order Direction"),
    ("order_status", "Order Status"),
    ("current_station_name", "Current Station"),
    ("next_station_name", "Next Station"),
    ("station_name", "Destination Station"),
    ("return_from_station_name", "Return From Station"),
    ("return_dest_station_name", "Return Destination"),
    ("pickup_station_name", "Pickup Station"),
    ("sort_code_name", "Sort Code"),
    ("zone_name", "Zone Name"),
    ("driver_id", "Driver ID"),
    ("driver_name", "Driver Name"),
    ("shop_id", "Shop ID"),
    ("chargeable_weight", "Chargeable Weight"),
    ("high_value", "High Value"),
    ("mutual_check", "Co-Check"),
    ("cache_type", "Cache Type"),
    ("order_dg_type", "Specific DG Type"),
    ("create_time", "Created Time"),
    ("receive_time", "Received Time"),
    ("current_station_received_time", "Current Station Received Time"),
    ("delivering_time", "Delivering Time"),
    ("delivered_time", "Delivered Time"),
    ("pickup_time", "Pickup Time"),
    ("on_hold_times", "Total On Hold Times"),
]

TIME_FIELDS = {
    "create_time",
    "receive_time",
    "current_station_received_time",
    "delivering_time",
    "delivered_time",
    "pickup_time",
}

BANGKOK_TZ = timezone(timedelta(hours=7))  # Asia/Bangkok, không có DST


# ============================================================
# GOOGLE SHEETS HELPERS
# ============================================================
def get_gspread_client():
    """Xác thực Google Sheets bằng Service Account (giống script cũ)."""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds_raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not creds_raw:
        raise RuntimeError("Thiếu biến môi trường GOOGLE_SERVICE_ACCOUNT_JSON")

    creds_dict = json.loads(creds_raw)
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def get_or_create_worksheet(ss, name):
    try:
        return ss.worksheet(name)
    except gspread.WorksheetNotFound:
        return ss.add_worksheet(title=name, rows=1000, cols=len(COLUMNS) + 5)


# ============================================================
# SPX API HELPERS
# ============================================================
def fetch_spx_with_retry(url, cookie, payload):
    headers = {
        "Content-Type": "application/json",
        "Cookie": cookie,
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://spx.shopee.vn",
        "Referer": "https://spx.shopee.vn/#/orderTracking",
        "User-Agent": "Mozilla/5.0",
    }

    last_response = None
    for attempt in range(1, CONFIG["MAX_RETRY"] + 1):
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=30)
            last_response = response
            if response.status_code < 400:
                return response

            print(
                f"Retry lần {attempt} vì HTTP {response.status_code}: "
                f"{response.text[:300]}"
            )
            time.sleep(attempt)  # 1000ms * attempt

        except requests.RequestException as e:
            print(f"Exception retry lần {attempt}: {e}")
            time.sleep(attempt)

    if last_response is not None:
        return last_response
    # Nếu tất cả attempt đều raise exception (chưa có response nào)
    return requests.post(url, headers=headers, json=payload, timeout=30)


def extract_rows(data):
    """Tương ứng extractRows_ / findFirstObjectArray_ trong Apps Script."""
    candidates = [
        (data.get("data") or {}).get("list"),
        (data.get("data") or {}).get("orders"),
        (data.get("data") or {}).get("rows"),
        (data.get("data") or {}).get("data"),
        (data.get("data") or {}).get("result"),
        (data.get("result") or {}).get("list"),
        (data.get("result") or {}).get("orders"),
        (data.get("result") or {}).get("rows"),
        data.get("list"),
        data.get("rows"),
    ]

    for c in candidates:
        if isinstance(c, list):
            return c

    found = find_first_object_array(data)
    return found or []


def find_first_object_array(value):
    if isinstance(value, list):
        if value and isinstance(value[0], dict):
            return value
        return None

    if not isinstance(value, dict):
        return None

    for v in value.values():
        found = find_first_object_array(v)
        if found:
            return found

    return None


# ============================================================
# VALUE FORMATTING (tương ứng formatValue_ và các map*_ )
# ============================================================
def first_non_empty(values):
    for v in values:
        if v not in (None, ""):
            return v
    return ""


def format_unix_time(value):
    try:
        seconds = float(value)
    except (TypeError, ValueError):
        return ""
    if not seconds:
        return ""
    dt = datetime.fromtimestamp(seconds, tz=BANGKOK_TZ)
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def map_order_status(value):
    mapping = {
        8: "SOC_Received",
        33: "SOC_Packed",
        58: "Return_SOC_Received",
        60: "Return_SOC_Packed",
    }
    return mapping.get(_to_int_or_none(value), value)


def map_order_account(value):
    mapping = {
        2: "Marketplace",
        12: "NS Marketplace Standard",
    }
    return mapping.get(_to_int_or_none(value), value)


def map_cod_status(value):
    mapping = {
        0: "-",
        1: "Created",
    }
    return mapping.get(_to_int_or_none(value), value)


def _to_int_or_none(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


TO_PATTERN = re.compile(r"^(TO|IHT)\d+", re.IGNORECASE)


def find_value_by_pattern(obj, pattern):
    if obj is None:
        return ""

    if isinstance(obj, (str, int, float)):
        text = str(obj)
        return text if pattern.search(text) else ""

    if isinstance(obj, list):
        for item in obj:
            found = find_value_by_pattern(item, pattern)
            if found:
                return found
        return ""

    if isinstance(obj, dict):
        for v in obj.values():
            found = find_value_by_pattern(v, pattern)
            if found:
                return found

    return ""


def find_current_to_number(row):
    direct = first_non_empty([
        row.get("current_to_number"),
        row.get("current_to_no"),
        row.get("current_to_num"),
        row.get("current_to"),
        row.get("to_number"),
        row.get("to_no"),
        row.get("to_num"),
        row.get("transport_order_number"),
        row.get("transport_order_no"),
        row.get("transport_no"),
        row.get("current_transport_order_number"),
        row.get("current_transport_order_no"),
    ])
    if direct:
        return direct

    return find_value_by_pattern(row, TO_PATTERN) or ""


def format_value(key, value, row):
    if key == "__current_to_number":
        return find_current_to_number(row)

    if value is None:
        return ""

    if key in TIME_FIELDS:
        return format_unix_time(value)

    if key == "order_status":
        return map_order_status(value)

    if key == "order_account":
        return map_order_account(value)

    if key == "cod_status":
        return map_cod_status(value)

    if key in ("high_value", "mutual_check"):
        try:
            return "Yes" if int(value) == 1 else "No"
        except (TypeError, ValueError):
            return "No"

    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)

    return value


# ============================================================
# MAIN
# ============================================================
def fetch_spx_order_tracking():
    gc = get_gspread_client()
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    if not spreadsheet_id:
        raise RuntimeError("Thiếu biến môi trường SPREADSHEET_ID")

    ss = gc.open_by_key(spreadsheet_id)

    cookie_ws = ss.worksheet(CONFIG["COOKIE_SHEET"])
    cookie = str(cookie_ws.acell(CONFIG["COOKIE_CELL"]).value or "").strip()
    if not cookie:
        raise RuntimeError(f"Chưa có cookie ở {CONFIG['COOKIE_SHEET']}!{CONFIG['COOKIE_CELL']}")

    output_ws = get_or_create_worksheet(ss, CONFIG["OUTPUT_SHEET"])

    all_rows = []

    for page in range(1, CONFIG["MAX_PAGES"] + 1):
        payload = {
            "high_value": 1,
            "order_status": "33,60,8,58",
            "count": CONFIG["PAGE_SIZE"],
            "current_station_ids": "1959",
            "page_no": page,
        }

        print(f"Đang gọi page {page}: {json.dumps(payload)}")

        res = fetch_spx_with_retry(CONFIG["API_URL"], cookie, payload)
        code = res.status_code
        text = res.text

        print(f"HTTP code page {page}: {code}")

        if code < 200 or code >= 300:
            print(f"Response lỗi: {text[:500]}")
            raise RuntimeError(f"SPX trả lỗi HTTP {code} tại page {page}: {text[:500]}")

        data = res.json()

        if data.get("retcode") != 0:
            print(f"SPX trả lỗi: {text[:500]}")
            raise RuntimeError(f"SPX retcode lỗi tại page {page}: {text[:500]}")

        rows = extract_rows(data)

        if not rows:
            print(f"Page {page} không có dữ liệu, dừng.")
            break

        all_rows.extend(rows)

        print(f"Page {page} lấy được {len(rows)} dòng. Tổng hiện tại: {len(all_rows)}")

        if len(rows) < CONFIG["PAGE_SIZE"]:
            print("Page cuối vì số dòng nhỏ hơn PAGE_SIZE.")
            break

        time.sleep(CONFIG["SLEEP_SEC"])

    output_ws.clear()

    if not all_rows:
        output_ws.update("A1", [["Không có dữ liệu"]])
        return

    write_rows(output_ws, all_rows)

    print(f"Đã ghi {len(all_rows)} dòng vào sheet {CONFIG['OUTPUT_SHEET']}")


def write_rows(ws, rows):
    headers = [col[1] for col in COLUMNS]

    values = [
        [format_value(key, row.get(key), row) for key, _ in COLUMNS]
        for row in rows
    ]

    ws.update("A1", [headers] + values, value_input_option="USER_ENTERED")

    ws.freeze(rows=1)
    ws.format("A1:" + gspread.utils.rowcol_to_a1(1, len(headers)), {
        "textFormat": {"bold": True},
        "backgroundColor": {"red": 0.91, "green": 0.945, "blue": 0.996},
    })

    print(f"Da ghi {len(values)} dong x {len(headers)} cot.")


if __name__ == "__main__":
    fetch_spx_order_tracking()
