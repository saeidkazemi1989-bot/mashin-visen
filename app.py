# -*- coding: utf-8 -*-
"""
Mashin-Visen — داشبورد گزارشگاری تست عملکردی (خواندن از مسیر شبکه/سیستم)

قابلیت‌ها:
  * خواندن داده‌های تست از مسیر شبکه یا پوشه محلی (بدون ذخیره محلي)
  * رابط کاملاً فارسی (RTL)
  * تاریخ شمسی (ج_dims)
  * دسته‌بندی تست‌ها (LOGO / GAUGE / LCD) و نمایش Tolerance و رنج مجاز
  * فیلتر بازه زمانی، نوع محصول، وضعیت
  * روند روزانه تست‌ها (تعداد، OK، رد)
  * جدول قابل جستجو و فیلتر + دریل‌داون به جزئیات هر محصول
"""
import os
import json
import math
import datetime
from io import BytesIO
from pathlib import Path

from flask import (
    Flask, render_template, request, redirect, url_for,
    send_file, flash, abort
)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart, Reference
import jdatetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# مسیر داده‌ها: می‌تواند مسیر شبکه (\\서버\share) یا پوشه محلی باشد
# قابل تغییر توسط متغیر محیطی MASHIN_DATA_PATH یا فرم تنظیمات
DEFAULT_DATA_PATH = os.environ.get("MASHIN_DATA_PATH", "")
PRODUCTS_DIR = os.path.join(BASE_DIR, "products")  # فقط برای کش موقت گزارش اکسل، ذخیره اصلی نیست
os.makedirs(PRODUCTS_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "mashin-visen-dev-key")
app.jinja_env.globals.update(enumerate=enumerate, zip=zip, len=len)


# ---------- تاریخ شمسی ----------
def to_shamsi(ts, fmt="%Y/%m/%d %H:%M:%S"):
    """تبدیل timestamp یا ISO به تاریخ شمسی."""
    if ts is None or ts == "":
        return ""
    try:
        if isinstance(ts, (int, float)):
            dt = datetime.datetime.fromtimestamp(float(ts))
        else:
            dt = datetime.datetime.fromisoformat(str(ts))
        jdt = jdatetime.datetime.fromgregorian(datetime=dt)
        return jdt.strftime(fmt)
    except Exception:
        return str(ts)


def to_shamsi_date(ts):
    if ts is None or ts == "":
        return ""
    try:
        if isinstance(ts, (int, float)):
            dt = datetime.datetime.fromtimestamp(float(ts))
        else:
            dt = datetime.datetime.fromisoformat(str(ts))
        jdt = jdatetime.datetime.fromgregorian(datetime=dt)
        return jdt.strftime("%Y/%m/%d")
    except Exception:
        return str(ts)


def parse_shamsi_date(s):
    """تبدیل رشته شمسی YYYY/MM/DD به datetime گريغوری برای فیلتر."""
    if not s:
        return None
    try:
        parts = str(s).split("/")
        if len(parts) == 3:
            jdt = jdatetime.datetime(int(parts[0]), int(parts[1]), int(parts[2]))
            return jdt.togregorian()
    except Exception:
        return None
    return None


# ---------- helpers ----------
def _load_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _load_jsonl(path):
    rows = []
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        rows.append(json.loads(line))
                    except Exception:
                        pass
    except Exception:
        pass
    return rows


def discover_sessions(data_root):
    """اسکن مسیر داده برای یافتن پوشه‌های حاوی session.json."""
    sessions = []
    root = Path(data_root)
    if not root.exists():
        return sessions
    # جستجوی عمقی تا 3 لایه (محصول / نشست / فایل‌ها)
    for session_file in root.rglob("session.json"):
        try:
            pdir = session_file.parent
            session = _load_json(str(session_file))
            if session and isinstance(session, dict):
                session["_folder"] = str(pdir)
                sessions.append(session)
        except Exception:
            pass
    return sessions


def _calc_tolerance(expected, tolerance):
    if expected is None:
        return None, None
    try:
        exp = float(expected)
    except (TypeError, ValueError):
        return None, None
    if tolerance is None:
        return exp, exp
    try:
        tol = float(tolerance)
    except (TypeError, ValueError):
        return exp, exp
    return exp - tol, exp + tol


def _status_with_tolerance(observed, expected, tolerance):
    if observed is None or expected is None:
        return "NA"
    try:
        obs = float(observed)
        exp = float(expected)
    except (TypeError, ValueError):
        return "NA"
    if tolerance is None:
        return "OK" if abs(obs - exp) < 1e-9 else "OUT_OF_RANGE"
    try:
        tol = float(tolerance)
    except (TypeError, ValueError):
        return "OK" if abs(obs - exp) < 1e-9 else "OUT_OF_RANGE"
    low, high = _calc_tolerance(expected, tolerance)
    if low is None or high is None:
        return "NA"
    if low - 1e-9 <= obs <= high + 1e-9:
        return "OK"
    return "OUT_OF_RANGE"


def _summarize_results(results, plan_steps=None):
    total = len(results)
    pass_c = sum(1 for r in results if r.get("status") == "PASS")
    fail_c = sum(1 for r in results if r.get("status") == "FAIL")
    error_c = sum(1 for r in results if r.get("status") == "ERROR")
    rate = (pass_c / total * 100) if total else 0.0

    by_kind = {}
    examples_fa = []
    plan_lookup = {}
    if plan_steps:
        for s in plan_steps:
            key = (s.get("kind"), s.get("target"))
            plan_lookup[key] = s

    for r in results:
        k = r.get("test_kind") or "NA"
        b = by_kind.setdefault(k, {"total": 0, "pass": 0, "fail": 0, "error": 0, "items": []})
        b["total"] += 1
        s = r.get("status")
        if s == "PASS":
            b["pass"] += 1
        elif s == "FAIL":
            b["fail"] += 1
        elif s == "ERROR":
            b["error"] += 1

        expected = r.get("expected")
        observed = r.get("observed")
        tolerance = None
        pl_key = (r.get("test_kind"), r.get("target"))
        if pl_key in plan_lookup:
            tolerance = plan_lookup[pl_key].get("tolerance") or plan_lookup[pl_key].get("expected_tolerance")

        range_status = _status_with_tolerance(observed, expected, tolerance)
        low, high = _calc_tolerance(expected, tolerance)

        item = {
            "test_kind": k,
            "target": r.get("target"),
            "expected": expected,
            "observed": observed,
            "status": s,
            "range_status": range_status,
            "tolerance": tolerance,
            "range_low": low,
            "range_high": high,
            "technical_reason": r.get("technical_reason"),
            "explanation_fa": r.get("user_explanation_fa") or "",
            "started_at_fmt": to_shamsi(r.get("started_at")),
            "operator_confirmed": bool(r.get("operator_confirmed")),
            "confirmed_by": r.get("confirmed_by", ""),
        }
        b["items"].append(item)

        if len(examples_fa) < 6:
            examples_fa.append({
                "status": s,
                "target": r.get("target"),
                "explanation": r.get("user_explanation_fa") or "",
            })
    return {
        "total": total, "pass": pass_c, "fail": fail_c, "error": error_c,
        "pass_rate": rate, "by_kind": by_kind, "examples_fa": examples_fa,
    }


def build_product_summary(session, results, events, packets, test_plan, placement, errors, folder):
    ctx = session.get("product_context") or {}
    tp = session.get("test_plan") or {}
    outcome = session.get("inspection_outcome") or {}
    plan_steps = tp.get("steps") or [] if isinstance(tp, dict) else []

    s = _summarize_results(results, plan_steps)

    total_tests = outcome.get("result_count", s["total"])
    if outcome.get("result_count"):
        pass_tests = outcome.get("result_count", 0) - outcome.get("fail_count", 0) - outcome.get("error_count", 0)
        fail_tests = outcome.get("fail_count", 0)
        error_tests = outcome.get("error_count", 0)
        pass_rate = (pass_tests / total_tests * 100) if total_tests else 0.0
    else:
        pass_tests, fail_tests, error_tests, pass_rate = s["pass"], s["fail"], s["error"], s["pass_rate"]

    return {
        "session_id": session.get("session_id") or os.path.basename(folder),
        "folder": folder,
        "product_serial": session.get("product_serial"),
        "operator": session.get("operator"),
        "product_id": ctx.get("product_id"),
        "product_name": ctx.get("product_name"),
        "product_type": ctx.get("product_type") or ctx.get("product_name") or ctx.get("product_id") or session.get("product_type") or "نامشخص",
        "product_revision": ctx.get("product_revision"),
        "test_plan_id": ctx.get("test_plan_id") or tp.get("plan_id"),
        "test_plan_revision": ctx.get("test_plan_revision") or tp.get("plan_revision"),
        "test_plan_mode": ctx.get("test_plan_mode") or tp.get("mode"),
        "plan_name": tp.get("plan_name"),
        "plan_id": tp.get("plan_id"),
        "plan_revision": tp.get("plan_revision"),
        "total_steps": tp.get("total_steps", len(plan_steps)),
        "observed_steps": tp.get("observed_steps", len(plan_steps)),
        "status": session.get("status"),
        "started_at_iso": to_shamsi(session.get("started_at") or session.get("started_at_iso")),
        "started_at_raw": session.get("started_at") or session.get("started_at_iso"),
        "completed_at_iso": to_shamsi(session.get("completed_at") or session.get("completed_at_iso")),
        "duration_seconds": session.get("duration_seconds", 0.0),
        "packet_count": session.get("packet_count", len(packets)),
        "event_count": session.get("event_count", len(events)),
        "result_count": session.get("result_count", len(results)),
        "error_count": session.get("error_count", len(errors)),
        "outcome_code": outcome.get("code"),
        "outcome_accepted": bool(outcome.get("accepted")),
        "total_tests": total_tests,
        "pass_tests": pass_tests,
        "fail_tests": fail_tests,
        "error_tests": error_tests,
        "pass_rate": pass_rate,
        "by_kind": s["by_kind"],
        "examples_fa": s["examples_fa"],
    }


def build_product_detail(session, results, events, packets, test_plan, placement, errors, folder):
    product = build_product_summary(session, results, events, packets, test_plan, placement, errors, folder)

    for r in results:
        r["started_at_fmt"] = to_shamsi(r.get("started_at"))
    for e in events:
        e["occurred_at_fmt"] = to_shamsi(e.get("occurred_at"))
    for p in packets:
        p["received_at_fmt"] = to_shamsi(p.get("received_at"))

    product["results"] = results
    product["events"] = events
    product["packets"] = packets
    product["plan_steps"] = (test_plan.get("steps") or []) if isinstance(test_plan, dict) else []
    product["placement"] = placement
    product["errors"] = errors
    return product


def list_products_from_path(data_root):
    """لیست محصولات از مسیر داده."""
    products = []
    sessions = discover_sessions(data_root)
    for session in sessions:
        folder = session.get("_folder", "")
        results = _load_jsonl(os.path.join(folder, "results.jsonl"))
        events = _load_jsonl(os.path.join(folder, "events.jsonl"))
        packets = _load_jsonl(os.path.join(folder, "packets.jsonl"))
        test_plan = _load_json(os.path.join(folder, "test_plan_shadow.json")) or {}
        placement = _load_json(os.path.join(folder, "placement_shadow.json")) or {}
        errors = _load_json(os.path.join(folder, "errors.json")) or []
        products.append(build_product_summary(
            session, results, events, packets, test_plan, placement, errors, folder
        ))
    # مرتب‌سازی بر اساس تاریخ شروع (ن你知道 جدیدتر اول)
    products.sort(key=lambda p: p.get("started_at_raw") or "", reverse=True)
    return products


def get_product_from_path(data_root, session_id):
    sessions = discover_sessions(data_root)
    for session in sessions:
        sid = session.get("session_id") or os.path.basename(session.get("_folder", ""))
        if sid == session_id:
            folder = session.get("_folder", "")
            results = _load_jsonl(os.path.join(folder, "results.jsonl"))
            events = _load_jsonl(os.path.join(folder, "events.jsonl"))
            packets = _load_jsonl(os.path.join(folder, "packets.jsonl"))
            test_plan = _load_json(os.path.join(folder, "test_plan_shadow.json")) or {}
            placement = _load_json(os.path.join(folder, "placement_shadow.json")) or {}
            errors = _load_json(os.path.join(folder, "errors.json")) or []
            return build_product_detail(
                session, results, events, packets, test_plan, placement, errors, folder
            )
    return None


# ---------- excel report ----------
def _style_header(ws, row, ncols, fill="1D4ED8"):
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hfont = Font(bold=True, color="FFFFFF", size=11, name="Tahoma")
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.font = hfont
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border


def _cell(v):
    if v is None or isinstance(v, (str, int, float, bool)):
        return v
    return json.dumps(v, ensure_ascii=False)


def _write_sheet(ws, title, headers, rows, date_cols=()):
    ws.title = title[:31]
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=13, color="1D4ED8", name="Tahoma")
    hr = 3
    for j, h in enumerate(headers, start=1):
        ws.cell(row=hr, column=j, value=h)
    _style_header(ws, hr, len(headers))
    for i, row in enumerate(rows, start=hr + 1):
        for j, v in enumerate(row, start=1):
            cell = ws.cell(row=i, column=j, value=_cell(v))
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True,
                                       horizontal="right" if j in date_cols else "general")
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.alignment = Alignment(vertical="center", horizontal="center")
    for j in range(1, len(headers) + 1):
        maxlen = len(str(headers[j - 1]))
        for i in range(hr + 1, hr + 1 + len(rows)):
            v = ws.cell(row=i, column=j).value
            if v is not None:
                maxlen = max(maxlen, len(str(v)))
        ws.column_dimensions[get_column_letter(j)].width = min(max(maxlen + 2, 12), 60)
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)


def _range_bar(observed, low, high):
    """ساخت نوار بصری برای نشان دادن جایگاه مقدار مشاهدشده در محدوده مجاز.
    نمایش: |===|●|===|  که ● جایگاه مقدار است."""
    if observed is None or low is None or high is None:
        return "—"
    try:
        obs = float(observed)
        lo = float(low)
        hi = float(high)
    except (TypeError, ValueError):
        return "—"
    span = hi - lo
    if span <= 0:
        return "—"
    # позиции 0 تا 20 کاراکتر
    pos = int((obs - lo) / span * 20)
    pos = max(0, min(20, pos))
    bar = "─" * pos + "●" + "─" * (20 - pos)
    return f"[{bar}]"


def _status_fill(cell, status):
    """رنگ‌بندی خودکار سلول وضعیت."""
    if status == "PASS" or status == "OK":
        cell.fill = PatternFill("solid", fgColor="C6EFCE")  # سبز روشنی
        cell.font = Font(bold=True, color="006100", size=11, name="Tahoma")
    elif status == "FAIL" or status == "OUT_OF_RANGE":
        cell.fill = PatternFill("solid", fgColor="FFC7CE")  # قرمز روشنی
        cell.font = Font(bold=True, color="9C0006", size=11, name="Tahoma")
    else:
        cell.fill = PatternFill("solid", fgColor="FFEB9C")  # زرد
        cell.font = Font(bold=True, color="9C6500", size=11, name="Tahoma")


def build_excel_bytes(product):
    wb = Workbook()
    wb.remove(wb.active)

    # ========== Sheet 1: داشبورد خلاصه (بصری و در یک نگاه قابل فهم) ==========
    ws = wb.create_sheet("داشبورد خلاصه")
    ws.sheet_view.showGridLines = False

    # عنوان اصلی
    ws.merge_cells("A1:H1")
    c = ws.cell(row=1, column=1, value="گزارش تست محصول — ماشین.Visual (هوش مصنوعی)")
    c.font = Font(bold=True, size=18, color="1D4ED8", name="Tahoma")
    c.alignment = Alignment(horizontal="center", vertical="center")

    # اطلاعات محصول
    info_rows = [
        ["شناسه جلسه", product.get("session_id"), "نوع محصول", product.get("product_type")],
        ["نام محصول", product.get("product_name") or product.get("product_id"), "رِوِژن محصول", product.get("product_revision")],
        ["شمارهٔ سريال", product.get("product_serial"), "اپراتور", product.get("operator")],
        ["برنامه تست", product.get("plan_name"), "رِوِژن برنامه", product.get("plan_revision")],
        ["زمان شروع (شمسی)", product.get("started_at_iso"), "زمان اتمام (شمسی)", product.get("completed_at_iso")],
    ]
    r = 3
    for a, va, b, vb in info_rows:
        ca = ws.cell(row=r, column=1, value=a)
        ca.font = Font(bold=True, size=11, name="Tahoma", color="475569")
        ca.fill = PatternFill("solid", fgColor="F1F5F9")
        ca.border = Border(bottom=Side(style="thin", color="E2E8F0"))
        cva = ws.cell(row=r, column=2, value=va)
        cva.font = Font(size=11, name="Tahoma")
        cva.border = Border(bottom=Side(style="thin", color="E2E8F0"))
        cb = ws.cell(row=r, column=3, value=b)
        cb.font = Font(bold=True, size=11, name="Tahoma", color="475569")
        cb.fill = PatternFill("solid", fgColor="F1F5F9")
        cb.border = Border(bottom=Side(style="thin", color="E2E8F0"))
        cvb = ws.cell(row=r, column=4, value=vb)
        cvb.font = Font(size=11, name="Tahoma")
        cvb.border = Border(bottom=Side(style="thin", color="E2E8F0"))
        r += 1

    # وضعیت نهایی با رنگ بزرگ
    r += 1
    ws.merge_cells(f"A{r}:H{r}")
    cc = ws.cell(row=r, column=1)
    cc.value = f"وضعیت نهایی: {'تاییدشده (PASS)' if product.get('outcome_accepted') else 'ردشده/خطا (FAIL)'}"
    cc.font = Font(bold=True, size=16, color="FFFFFF", name="Tahoma")
    cc.fill = PatternFill("solid", fgColor="006100" if product.get("outcome_accepted") else "9C0006")
    cc.alignment = Alignment(horizontal="center", vertical="center")
    cc.border = Border(left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"),
                       top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"))

    # شاخص‌های کلیدی با کارت رنگی
    r += 2
    ws.merge_cells(f"A{r}:H{r}")
    ch = ws.cell(row=r, column=1, value="شاخص‌های کلیدی")
    ch.font = Font(bold=True, size=14, color="1D4ED8", name="Tahoma")

    kpis = [
        ("تست‌های تاییدشده", product.get("pass_tests", 0), "C6EFCE", "006100"),
        ("تست‌های خارج از رنج", product.get("fail_tests", 0), "FFC7CE", "9C0006"),
        ("خطاهای سیستمی", product.get("error_tests", 0), "FFEB9C", "9C6500"),
        ("نرخ تایید", f'{product.get("pass_rate", 0):.1f}%', "E2EFDA", "375623"),
    ]
    ccol = 1
    for label, val, fill, fontc in kpis:
        cell = ws.cell(row=r + 1, column=ccol, value=label)
        cell.font = Font(bold=True, size=11, name="Tahoma", color="475569")
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.border = Border(left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
                             top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"))
        cell.alignment = Alignment(horizontal="center", vertical="center")
        vcell = ws.cell(row=r + 1, column=ccol + 1, value=val)
        vcell.font = Font(bold=True, size=16, name="Tahoma", color=fontc)
        vcell.fill = PatternFill("solid", fgColor=fill)
        vcell.border = Border(left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
                              top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"))
        vcell.alignment = Alignment(horizontal="center", vertical="center")
        ccol += 2

    # نمودارทortie (پاورپوينت-اشخانه) با openpyxl PieChart
    from openpyxl.chart import PieChart, Reference
    pie = PieChart()
    pie.title = "توزیع نتایج تست"
    data = Reference(ws, min_col=2, min_row=r + 1, max_row=r + 3, max_col=2)
    cats = Reference(ws, min_col=1, min_row=r + 1, max_row=r + 3, max_col=1)
    pie.add_data(data, titles_from_data=False)
    pie.set_categories(cats)
    pie.height = 9
    pie.width = 18
    ws.add_chart(pie, f"A{r + 4}")

    # ========== Sheet 2: نتایج تست (بصری با نوار محدوده) ==========
    res_headers = ["نوع تست", "هدف (Target)", "منتظرشده", "مشاهدشده", "حد پایین", "حد بالایی",
                   "نوار جایگاه عقربه", "وضعیت رنج", "وضعیت کلی", "تایید اپراتور", "شرح (فارسی)"]
    res_rows = []
    for x in (product.get("results") or []):
        bar = _range_bar(x.get("observed"), x.get("range_low"), x.get("range_high"))
        res_rows.append([
            x.get("test_kind"), x.get("target"), x.get("expected"), x.get("observed"),
            x.get("range_low"), x.get("range_high"),
            bar, x.get("range_status"), x.get("status"),
            "✅ تایید شده" if x.get("operator_confirmed") else "⏳ تایید نشده",
            x.get("explanation_fa"),
        ])
    ws2 = wb.create_sheet("نتایج تست")
    _write_sheet(ws2, "نتایج تست — جایگاه هر تست در محدوده مجاز", res_headers, res_rows)
    # رنگ‌بندی ستون‌های وضعیت
    for i, row in enumerate(res_rows, start=4):
        status_cell = ws2.cell(row=i, column=9)
        _status_fill(status_cell, row[8])
        range_cell = ws2.cell(row=i, column=8)
        _status_fill(range_cell, row[7])
        op_cell = ws2.cell(row=i, column=10)
        if "تایید شده" in str(row[9]):
            op_cell.fill = PatternFill("solid", fgColor="C6EFCE")
            op_cell.font = Font(bold=True, color="006100", size=10, name="Tahoma")
        else:
            op_cell.fill = PatternFill("solid", fgColor="FFEB9C")
            op_cell.font = Font(bold=True, color="9C6500", size=10, name="Tahoma")

    # ========== Sheet 3: مراحل برنامه تست ==========
    step_headers = ["ترتیب", "شناسه مرحله", "نوع", "هدف", "منتظرشده", "Tolerance", "اکشن", "مشاهدشده",
                    "زمان نسبت به شروع (ms)", "پکت منبع"]
    step_rows = [
        [s.get("order"), s.get("step_id"), s.get("kind"), s.get("target"), s.get("expected"),
         s.get("tolerance"), s.get("action"), "بله" if s.get("observed") else "خیر",
         f'{s.get("relative_ms", 0):.1f}', s.get("source_packet_id")]
        for s in (product.get("plan_steps") or [])
    ]
    _write_sheet(wb.create_sheet("مراحل برنامه"), "مراحل برنامه تست", step_headers, step_rows)

    # ========== Sheet 4: رویدادها ==========
    ev_headers = ["زمان (شمسی)", "نوع رویداد", "شناسه", "هدف", "منتظرشده", "مرحله تست", "پکت"]
    ev_rows = [
        [e.get("occurred_at_fmt") or to_shamsi(e.get("occurred_at")), e.get("kind"),
         e.get("identifier"), e.get("target"), e.get("expected"),
         ((e.get("details") or {}).get("test_plan") or {}).get("step_id", "") if e.get("details") else "",
         e.get("packet_id")]
        for e in (product.get("events") or [])
    ]
    _write_sheet(wb.create_sheet("رویدادها"), "رویدادها", ev_headers, ev_rows)

    # ========== Sheet 5: پکت‌های CAN ==========
    pk_headers = ["زمان دریافت (شمسی)", "مبدا IP:پورت", "مقصد IP:پورت", "پديو (hex)", "رابطه"]
    pk_rows = [
        [p.get("received_at_fmt") or to_shamsi(p.get("received_at")),
         f'{p.get("source_ip")}:{p.get("source_port")}',
         f'{p.get("destination_ip")}:{p.get("destination_port")}',
         p.get("payload_hex"), p.get("capture_interface")]
        for p in (product.get("packets") or [])
    ]
    _write_sheet(wb.create_sheet("پکت‌های CAN"), "پکت‌های CAN", pk_headers, pk_rows)

    # ========== Sheet 6: خطاها ==========
    _write_sheet(wb.create_sheet("خطاها"), "خطاها",
                 ["ردیف", "خطا"],
                 [[i + 1, e] for i, e in enumerate(product.get("errors") or [])])

    # ========== Sheet 7: جایگاه‌دهی تصویر ==========
    pl = product.get("placement") or {}
    counts = pl.get("counts") or {}
    pl_rows = [
        ["موتور estimate", pl.get("engine")],
        ["dx", pl.get("dx")],
        ["dy", pl.get("dy")],
        ["spreッド (px)", pl.get("spread_px")],
        ["تاییدشده", counts.get("trusted")],
        ["ردشده", counts.get("rejected")],
        ["presenting", counts.get("offered")],
        ["mediana زمان estimate (ms)", pl.get("estimate_ms_median")],
        ["حداکثر زمان estimate (ms)", pl.get("estimate_ms_max")],
        ["وضعیت قفل", "قفل‌شده" if pl.get("locked") else "باز"],
    ]
    _write_sheet(wb.create_sheet("جایگاه‌دهی"), "جایگاه‌دهی تصویر", ["مورد", "مقدار"], pl_rows)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------- routes ----------
@app.route("/", methods=["GET", "POST"])
def dashboard():
    # مسیر داده‌ها: اول از فرم POST، سپس از query string، سپس از session، در نهایت DEFAULT
    data_root = request.form.get("data_path") or request.args.get("data_path") or session.get("data_path") or DEFAULT_DATA_PATH

    # ذخیره در session تا در همه صفحات حفظ شود
    if data_root:
        session["data_path"] = data_root

    if not data_root:
        flash("لطفاً مسیر پوشه داده‌های تست (شبکه یا محلی) را وارد کنید.", "error")
        return render_template("dashboard.html", data_path="", products=[], kpis={},
                               donut_pct=0, donut_color="#dc2626",
                               donut_path="M 110 18 A 92 92 0 1 1 202 110 A 92 92 0 1 1 110 18",
                               recent=[], date_from="", date_to="", product_type_filter="",
                               status_filter="", search="")

    # پس از بارگذاری داده‌ها (POST)، به همان صفحه redirected می‌شویم تا data_path در URL بماند
    # و فیلترها و رفرش صفحه باعث از بین رفتن داده‌ها نشود
    if request.method == "POST":
        return redirect(url_for("dashboard", data_path=data_root))

    products = list_products_from_path(data_root)

    # فیلترها
    date_from = request.form.get("date_from") or request.args.get("date_from") or ""
    date_to = request.form.get("date_to") or request.args.get("date_to") or ""
    product_type_filter = request.form.get("product_type") or request.args.get("product_type") or ""
    status_filter = request.form.get("status") or request.args.get("status") or ""
    search = request.form.get("search") or request.args.get("search") or ""

    dt_from = parse_shamsi_date(date_from)
    dt_to = parse_shamsi_date(date_to)

    filtered = []
    for p in products:
        # بازه زمانی
        if dt_from or dt_to:
            pdt = parse_shamsi_date(p.get("started_at_iso")) or \
                  parse_shamsi_date(p.get("completed_at_iso"))
            if pdt:
                if dt_from and pdt < dt_from:
                    continue
                if dt_to and pdt > dt_to:
                    continue
        # نوع محصول
        if product_type_filter and product_type_filter not in str(p.get("product_type", "")):
            continue
        # وضعیت
        if status_filter:
            if status_filter == "PASS" and not p.get("outcome_accepted"):
                continue
            if status_filter == "FAIL" and p.get("outcome_accepted"):
                continue
        # جستجو
        if search:
            text = f"{p.get('session_id','')} {p.get('product_name','')} {p.get('product_id','')} {p.get('operator','')} {p.get('product_type','')}".lower()
            if search.lower() not in text:
                continue
        filtered.append(p)

    products = filtered

    # KPIها
    total_products = len(products)
    total_tests = sum(p["total_tests"] for p in products)
    pass_tests = sum(p["pass_tests"] for p in products)
    fail_tests = sum(p["fail_tests"] for p in products)
    error_tests = sum(p["error_tests"] for p in products)
    pass_rate = (pass_tests / total_tests * 100) if total_tests else 0.0
    avg_duration = (sum(p["duration_seconds"] for p in products) / total_products) if total_products else 0.0

    by_kind = {}
    for p in products:
        for k, v in (p.get("by_kind") or {}).items():
            b = by_kind.setdefault(k, {"total": 0, "pass": 0, "fail": 0})
            b["total"] += v["total"]; b["pass"] += v["pass"]; b["fail"] += v["fail"]

    # روند روزانه (بر اساس تاریخ شمسی شروع)
    daily = {}
    for p in products:
        d = (p.get("started_at_iso") or "").split(" ")[0]
        if d:
            daily[d] = daily.get(d, {"total": 0, "pass": 0, "fail": 0})
            daily[d]["total"] += p["total_tests"]
            daily[d]["pass"] += p["pass_tests"]
            daily[d]["fail"] += p["fail_tests"]
    daily_list = [{"day": d, **c} for d, c in sorted(daily.items())]

    # نمودار دونت
    pct = pass_rate
    color = "#059669" if pct >= 95 else "#dc2626"
    if pct >= 100:
        donut_path = "M 110 18 A 92 92 0 1 1 202 110 A 92 92 0 1 1 110 18"
    else:
        large = 1 if pct > 50 else 0
        theta = math.radians(-90 + pct / 100.0 * 360.0)
        x = 110 + 92 * math.cos(theta)
        y = 110 + 92 * math.sin(theta)
        donut_path = f"M 110 18 A 92 92 0 {large} 1 {x:.1f} {y:.1f}"

    recent = products[:10]

    return render_template(
        "dashboard.html",
        data_path=data_root,
        products=products,
        kpis={
            "total_products": total_products, "pass_rate": pass_rate,
            "total_tests": total_tests, "pass_tests": pass_tests,
            "fail_tests": fail_tests, "error_tests": error_tests,
            "avg_duration": avg_duration, "by_kind": by_kind, "daily": daily_list,
        },
        donut_pct=pct, donut_color=color, donut_path=donut_path,
        recent=recent,
        date_from=date_from, date_to=date_to,
        product_type_filter=product_type_filter, status_filter=status_filter, search=search,
    )


@app.route("/products")
def product_list():
    data_root = request.args.get("data_path") or session.get("data_path") or DEFAULT_DATA_PATH
    # حفظ در session برای صفحه‌های بعدی
    if data_root:
        session["data_path"] = data_root
    products = list_products_from_path(data_root) if data_root else []
    return render_template("product_list.html", products=products, data_path=data_root)


@app.route("/products/<pid>")
def product_detail(pid):
    data_root = request.args.get("data_path") or session.get("data_path") or DEFAULT_DATA_PATH
    if data_root:
        session["data_path"] = data_root
    product = get_product_from_path(data_root, pid) if data_root else None
    if product is None:
        abort(404)
    return render_template("product_detail.html", product=product, data_path=data_root)


@app.route("/products/<pid>/download/<filename>")
def download_file(pid, filename):
    data_root = request.args.get("data_path") or session.get("data_path") or DEFAULT_DATA_PATH
    if data_root:
        session["data_path"] = data_root
    sessions = discover_sessions(data_root) if data_root else []
    target_dir = None
    for sess in sessions:
        sid = sess.get("session_id") or os.path.basename(sess.get("_folder", ""))
        if sid == pid:
            target_dir = sess.get("_folder", "")
            break
    if not target_dir:
        abort(404)
    path = os.path.join(target_dir, filename)
    if not os.path.isfile(path):
        abort(404)
    return send_file(path, as_attachment=True)


@app.route("/products/<pid>/excel")
def download_excel(pid):
    data_root = request.args.get("data_path") or session.get("data_path") or DEFAULT_DATA_PATH
    if data_root:
        session["data_path"] = data_root
    product = get_product_from_path(data_root, pid) if data_root else None
    if product is None:
        abort(404)
    xlsx = build_excel_bytes(product)
    safe_name = f"گزارش_{product.get('session_id','product')}.xlsx"
    return send_file(
        xlsx,
        as_attachment=True,
        download_name=safe_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/products/<pid>/confirm/<result_id>", methods=["POST"])
def confirm_result(pid, result_id):
    """تایید یک نتیجه تست توسط اپراتور (ذخیره در همان پوشه منبع)."""
    data_root = request.form.get("data_path") or request.args.get("data_path") or DEFAULT_DATA_PATH
    sessions = discover_sessions(data_root) if data_root else []
    target_dir = None
    for session in sessions:
        sid = session.get("session_id") or os.path.basename(session.get("_folder", ""))
        if sid == pid:
            target_dir = session.get("_folder", "")
            break
    if not target_dir:
        abort(404)

    results_path = os.path.join(target_dir, "results.jsonl")
    rows = _load_jsonl(results_path)
    updated = False
    for r in rows:
        if r.get("result_id") == result_id:
            r["operator_confirmed"] = True
            r["confirmed_by"] = request.form.get("operator_name", "اپراتور")
            r["confirmed_at"] = datetime.datetime.now().isoformat()
            updated = True
            break
    if updated:
        with open(results_path, "w", encoding="utf-8") as f:
            for r in rows:
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
        flash("نتایج تست توسط اپراتور تایید شد.", "success")
    else:
        flash("نتایج تست یافت نشد.", "error")
    return redirect(url_for("product_detail", pid=pid, data_path=data_root))


@app.errorhandler(404)
def _404(e):
    return render_template("base.html", title="یافت نشد"), 404


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
