# -*- coding: utf-8 -*-
"""
Mashin-Visen — داشبورد گزارشگاری تست عملکردی لامپ‌های شاسوی توسط ماشین ویژن (AI).

قابلیت‌ها:
  * خواندن فایل‌های خروجی ماشین ویژن (session, results, events, packets, plan, placement, errors)
  * داشبورد مركزی با شاخص‌ها و نمودار (بدون وابستگی به CDN)/
  * ثبت هر محصول تست‌شده در پوشه مجزا تحت products/<شناسه>/
  * تولید گزارش اکسل چند صفحه‌ای برای هر محصول
"""
import os
import json
import math
import shutil
import datetime
from io import BytesIO

from flask import (
    Flask, render_template, request, redirect, url_for,
    send_file, flash, abort
)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRODUCTS_DIR = os.path.join(BASE_DIR, "products")
os.makedirs(PRODUCTS_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "mashin-visen-dev-key")
app.jinja_env.globals.update(enumerate=enumerate, zip=zip, len=len)

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


def fmt_ts(ts, fmt="%Y/%m/%d %H:%M:%S"):
    if ts is None or ts == "":
        return ""
    try:
        if isinstance(ts, (int, float)):
            dt = datetime.datetime.fromtimestamp(float(ts))
        else:
            dt = datetime.datetime.fromisoformat(str(ts))
        return dt.strftime(fmt)
    except Exception:
        return str(ts)


def fmt_date(ts):
    return fmt_ts(ts, "%Y/%m/%d")


def _summarize_results(results):
    total = len(results)
    pass_c = sum(1 for r in results if r.get("status") == "PASS")
    fail_c = sum(1 for r in results if r.get("status") == "FAIL")
    error_c = sum(1 for r in results if r.get("status") == "ERROR")
    rate = (pass_c / total * 100) if total else 0.0

    by_kind = {}
    examples_fa = []
    for r in results:
        k = r.get("test_kind") or "NA"
        b = by_kind.setdefault(k, {"total": 0, "pass": 0, "fail": 0, "error": 0})
        b["total"] += 1
        s = r.get("status")
        if s == "PASS":
            b["pass"] += 1
        elif s == "FAIL":
            b["fail"] += 1
        elif s == "ERROR":
            b["error"] += 1
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


def _detail_from_parts(session, results, events, packets, test_plan, placement, errors):
    """ساخت دیکشنر کامل یک محصول از بخش‌های داده‌ای."""
    ctx = session.get("product_context") or {}
    tp = session.get("test_plan") or {}
    outcome = session.get("inspection_outcome") or {}
    s = _summarize_results(results)

    plan_steps = tp.get("steps") or [] if isinstance(tp, dict) else []
    for r in results:
        r["started_at_fmt"] = fmt_ts(r.get("started_at"))
    for e in events:
        e["occurred_at_fmt"] = fmt_ts(e.get("occurred_at"))
    for p in packets:
        p["received_at_fmt"] = fmt_ts(p.get("received_at"))

    total_tests = outcome.get("result_count", s["total"])
    if outcome.get("result_count"):
        pass_tests = outcome.get("result_count", 0) - outcome.get("fail_count", 0) - outcome.get("error_count", 0)
        fail_tests = outcome.get("fail_count", 0)
        error_tests = outcome.get("error_count", 0)
        pass_rate = (pass_tests / total_tests * 100) if total_tests else 0.0
    else:
        pass_tests, fail_tests, error_tests, pass_rate = s["pass"], s["fail"], s["error"], s["pass_rate"]

    return {
        "session_id": session.get("session_id"),
        "product_serial": session.get("product_serial"),
        "operator": session.get("operator"),
        "product_id": ctx.get("product_id"),
        "product_name": ctx.get("product_name"),
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
        "started_at_iso": session.get("started_at_iso") or fmt_ts(session.get("started_at")),
        "completed_at_iso": session.get("completed_at_iso") or fmt_ts(session.get("completed_at")),
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
        "results": results,
        "events": events,
        "packets": packets,
        "plan_steps": plan_steps,
        "placement": placement,
        "errors": errors,
    }


def _summary_from_session(session, results):
    """دیکشنر خلاصه برای داشبورد/لیست."""
    ctx = session.get("product_context") or {}
    tp = session.get("test_plan") or {}
    outcome = session.get("inspection_outcome") or {}
    s = _summarize_results(results)
    total_tests = outcome.get("result_count", s["total"])
    if outcome.get("result_count"):
        pass_tests = outcome.get("result_count", 0) - outcome.get("fail_count", 0) - outcome.get("error_count", 0)
        fail_tests = outcome.get("fail_count", 0)
        error_tests = outcome.get("error_count", 0)
        rate = (pass_tests / total_tests * 100) if total_tests else 0.0
    else:
        pass_tests, fail_tests, error_tests, rate = s["pass"], s["fail"], s["error"], s["pass_rate"]
    return {
        "session_id": session.get("session_id"),
        "product_id": ctx.get("product_id"),
        "product_name": ctx.get("product_name"),
        "product_revision": ctx.get("product_revision"),
        "operator": session.get("operator"),
        "plan_name": tp.get("plan_name"),
        "plan_id": tp.get("plan_id"),
        "plan_revision": tp.get("plan_revision"),
        "status": session.get("status"),
        "started_at_iso": session.get("started_at_iso") or fmt_ts(session.get("started_at")),
        "completed_at_iso": session.get("completed_at_iso") or fmt_ts(session.get("completed_at")),
        "duration_seconds": session.get("duration_seconds", 0.0),
        "packet_count": session.get("packet_count", 0),
        "event_count": session.get("event_count", 0),
        "total_tests": total_tests,
        "pass_tests": pass_tests,
        "fail_tests": fail_tests,
        "error_tests": error_tests,
        "pass_rate": rate,
        "outcome_accepted": bool(outcome.get("accepted", (fail_tests + error_tests) == 0)),
        "by_kind": s["by_kind"],
    }


def _build_current():
    """جلسه تست جاری از فایل‌های ریشه پروژه."""
    session = _load_json(os.path.join(BASE_DIR, "session.json")) or {}
    return {
        "session": session,
        "results": _load_jsonl(os.path.join(BASE_DIR, "results.jsonl")),
        "events": _load_jsonl(os.path.join(BASE_DIR, "events.jsonl")),
        "packets": _load_jsonl(os.path.join(BASE_DIR, "packets.jsonl")),
        "test_plan": _load_json(os.path.join(BASE_DIR, "test_plan_shadow.json")) or {},
        "placement": _load_json(os.path.join(BASE_DIR, "placement_shadow.json")) or {},
        "errors": _load_json(os.path.join(BASE_DIR, "errors.json")) or [],
        "report_status": _load_json(os.path.join(BASE_DIR, "report_status.json")) or {},
        "summary": _summary_from_session(
            session, _load_jsonl(os.path.join(BASE_DIR, "results.jsonl"))
        ),
    }


def list_products():
    products = []
    if not os.path.isdir(PRODUCTS_DIR):
        return products
    for name in sorted(os.listdir(PRODUCTS_DIR)):
        pdir = os.path.join(PRODUCTS_DIR, name)
        if not os.path.isdir(pdir):
            continue
        session = _load_json(os.path.join(pdir, "session.json")) or {}
        results = _load_jsonl(os.path.join(pdir, "results.jsonl"))
        products.append(_summary_from_session(session, results))
    products.sort(key=lambda p: p.get("started_at_iso") or "", reverse=True)
    return products


def get_product(pid):
    pdir = os.path.join(PRODUCTS_DIR, pid)
    if not os.path.isdir(pdir):
        for name in os.listdir(PRODUCTS_DIR):
            cand = os.path.join(PRODUCTS_DIR, name)
            if os.path.isdir(cand):
                sess = _load_json(os.path.join(cand, "session.json")) or {}
                if sess.get("session_id") == pid:
                    pdir = cand
                    break
        else:
            return None
    session = _load_json(os.path.join(pdir, "session.json")) or {}
    return _detail_from_parts(
        session,
        _load_jsonl(os.path.join(pdir, "results.jsonl")),
        _load_jsonl(os.path.join(pdir, "events.jsonl")),
        _load_jsonl(os.path.join(pdir, "packets.jsonl")),
        _load_json(os.path.join(pdir, "test_plan_shadow.json")) or {},
        _load_json(os.path.join(pdir, "placement_shadow.json")) or {},
        _load_json(os.path.join(pdir, "errors.json")) or [],
    )


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


def build_excel_bytes(product):
    wb = Workbook()
    wb.remove(wb.active)

    # Sheet 1: Summary
    ws = wb.create_sheet("خلاصه")
    ws.cell(row=1, column=1, value="گزارش تست محصول — ماشین ویژن").font = Font(
        bold=True, size=15, color="1D4ED8", name="Tahoma")
    summary_rows = [
        ["شناسه جلسه", product.get("session_id")],
        ["محصول", product.get("product_name") or product.get("product_id")],
        ["رِوِژن محصول", product.get("product_revision")],
        ["شمارهٔ سريال", product.get("product_serial")],
        ["اپراتور", product.get("operator")],
        ["برنامه تست", product.get("plan_name")],
        ["شناسه برنامه", product.get("plan_id")],
        ["رِوِژن برنامه", product.get("plan_revision")],
        ["حالت برنامه", product.get("test_plan_mode")],
        ["وضعیت نهایی", "تاییدشده (PASS)" if product.get("outcome_accepted") else "ردشده/خطا (FAIL)"],
        ["تعداد کل تست‌ها", product.get("total_tests")],
        ["تست‌های تاییدشده", product.get("pass_tests")],
        ["تست‌های ردشده", product.get("fail_tests")],
        ["خطاهای سیستمی", product.get("error_tests")],
        ["نرخ تایید", f'{product.get("pass_rate", 0):.1f}%'],
        ["زمان شروع", product.get("started_at_iso")],
        ["زمان اتمام", product.get("completed_at_iso")],
        ["مدت تست (ثانیه)", f'{product.get("duration_seconds", 0):.2f}'],
        ["تعداد پکت CAN", product.get("packet_count")],
        ["تعداد رویداد", product.get("event_count")],
    ]
    r = 3
    for k, v in summary_rows:
        ck = ws.cell(row=r, column=1, value=_cell(k))
        ck.font = Font(bold=True, size=11, name="Tahoma", color="475569")
        ck.fill = PatternFill("solid", fgColor="F1F5F9")
        cv = ws.cell(row=r, column=2, value=_cell(v))
        cv.font = Font(size=11, name="Tahoma")
        for cell in (ck, cv):
            cell.border = Border(bottom=Side(style="thin", color="E2E8F0"))
        r += 1
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 60

    # Sheet 2: Test Results
    res_headers = ["نوع تست", "هدف (target)", "منتظرشده", "مشاهدشده", "وضعیت",
                   "دلیل فنی", "شرح (فارسی)", "زمان شروع"]
    res_rows = [
        [x.get("test_kind"), x.get("target"), x.get("expected"), x.get("observed"),
         x.get("status"), x.get("technical_reason"), x.get("user_explanation_fa"),
         x.get("started_at_fmt") or fmt_ts(x.get("started_at"))]
        for x in (product.get("results") or [])
    ]
    _write_sheet(wb.create_sheet("نتایج تست"), "نتایج تست", res_headers, res_rows, date_cols=(8,))

    # Sheet 3: Test Plan Steps
    step_headers = ["ترتیب", "شناسه مرحله", "نوع", "هدف", "اکشن", "مشاهدشده",
                    "زمان نسبت به شروع (ms)", "پکت منبع"]
    step_rows = [
        [s.get("order"), s.get("step_id"), s.get("kind"), s.get("target"), s.get("action"),
         "بله" if s.get("observed") else "خیر", f'{s.get("relative_ms", 0):.1f}',
         s.get("source_packet_id")]
        for s in (product.get("plan_steps") or [])
    ]
    _write_sheet(wb.create_sheet("مراحل برنامه"), "مراحل برنامه", step_headers, step_rows)

    # Sheet 4: Events
    ev_headers = ["زمان", "نوع رویداد", "شناسه", "هدف", "منتظرشده", "مرحله تست", "پکت"]
    ev_rows = [
        [e.get("occurred_at_fmt") or fmt_ts(e.get("occurred_at")), e.get("kind"),
         e.get("identifier"), e.get("target"), e.get("expected"),
         ((e.get("details") or {}).get("test_plan") or {}).get("step_id", "") if e.get("details") else "",
         e.get("packet_id")]
        for e in (product.get("events") or [])
    ]
    _write_sheet(wb.create_sheet("رویدادها"), "رویدادها", ev_headers, ev_rows, date_cols=(1,))

    # Sheet 5: CAN Packets
    pk_headers = ["زمان دریافت", "مبدا IP:پورت", "مقصد IP:پورت", "پديو (hex)", "رابطه"]
    pk_rows = [
        [p.get("received_at_fmt") or fmt_ts(p.get("received_at")),
         f'{p.get("source_ip")}:{p.get("source_port")}',
         f'{p.get("destination_ip")}:{p.get("destination_port")}',
         p.get("payload_hex"), p.get("capture_interface")]
        for p in (product.get("packets") or [])
    ]
    _write_sheet(wb.create_sheet("پکت‌های CAN"), "پکت‌های CAN", pk_headers, pk_rows, date_cols=(1,))

    # Sheet 6: Errors
    _write_sheet(wb.create_sheet("خطاها"), "خطاها",
                 ["ردیف", "خطا"],
                 [[i + 1, e] for i, e in enumerate(product.get("errors") or [])])

    # Sheet 7: Placement
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
    _write_sheet(wb.create_sheet("جایگاه‌دهی"), "جایگاه‌دهی", ["مورد", "مقدار"], pl_rows)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------- routes ----------
@app.route("/")
def dashboard():
    current = _build_current()
    products = list_products()

    # KPIها: محصولات بوفـsed + جلسه جاری (اگر بوفـsed نشده باشد)
    all_items = list(products)
    cur_sum = current["summary"]
    if not current["session"].get("session_id") or not any(
            p["session_id"] == cur_sum["session_id"] for p in products):
        all_items.append(cur_sum)

    total_products = len(all_items)
    total_tests = sum(p["total_tests"] for p in all_items)
    pass_tests = sum(p["pass_tests"] for p in all_items)
    fail_tests = sum(p["fail_tests"] for p in all_items)
    error_tests = sum(p["error_tests"] for p in all_items)
    pass_rate = (pass_tests / total_tests * 100) if total_tests else 0.0
    avg_duration = (sum(p["duration_seconds"] for p in all_items) / total_products) if total_products else 0.0

    by_kind = {}
    for p in all_items:
        for k, v in (p.get("by_kind") or {}).items():
            b = by_kind.setdefault(k, {"total": 0, "pass": 0, "fail": 0})
            b["total"] += v["total"]; b["pass"] += v["pass"]; b["fail"] += v["fail"]

    daily = {}
    for r in (current["results"] or []):
        d = fmt_date(r.get("started_at"))
        daily[d] = daily.get(d, 0) + 1
    if not daily:
        for p in products:
            d = (p.get("started_at_iso") or "").split(" ")[0]
            if d:
                daily[d] = daily.get(d, 0) + 1
    daily_list = [{"day": d, "count": c} for d, c in sorted(daily.items())]

    pct = pass_rate
    sweep = pct / 100.0 * 360.0
    color = "#059669" if pct >= 95 else "#dc2626"
    if pct >= 100:
        donut_path = "M 110 18 A 92 92 0 1 1 202 110 A 92 92 0 1 1 110 18"
    else:
        large = 1 if pct > 50 else 0
        theta = math.radians(-90 + sweep)
        x = 110 + 92 * math.cos(theta)
        y = 110 + 92 * math.sin(theta)
        donut_path = f"M 110 18 A 92 92 0 {large} 1 {x:.1f} {y:.1f}"

    return render_template(
        "dashboard.html",
        current=cur_sum,
        kpis={
            "total_products": total_products, "pass_rate": pass_rate,
            "total_tests": total_tests, "pass_tests": pass_tests,
            "fail_tests": fail_tests, "error_tests": error_tests,
            "avg_duration": avg_duration, "by_kind": by_kind, "daily": daily_list,
        },
        donut_pct=pct, donut_color=color, donut_path=donut_path,
        recent=products[:8],
    )


@app.route("/products")
def product_list():
    return render_template("product_list.html", products=list_products())


@app.route("/products/<pid>")
def product_detail(pid):
    product = get_product(pid)
    if product is None:
        current = _build_current()
        if current["session"].get("session_id") == pid:
            product = _detail_from_parts(
                current["session"], current["results"], current["events"],
                current["packets"], current["test_plan"], current["placement"],
                current["errors"],
            )
    if product is None:
        abort(404)
    return render_template("product_detail.html", product=product)


@app.route("/products/create", methods=["GET", "POST"])
def create_product():
    current = _build_current()
    if request.method == "POST":
        product_id = request.form.get("product_id", "").strip()
        product_name = request.form.get("product_name", "").strip()
        product_revision = request.form.get("product_revision", "00").strip()
        product_serial = request.form.get("product_serial", "").strip()
        operator = request.form.get("operator", "").strip()
        plan_id = request.form.get("plan_id", "").strip()
        plan_name = request.form.get("plan_name", "").strip()
        plan_revision = request.form.get("plan_revision", "").strip()
        notes = request.form.get("notes", "").strip()

        if not (product_id and product_serial and operator):
            flash("لطفاً فیلدهای الزامی (محصول، سريال، اپراتور) را پر کنید.", "error")
            return render_template("create_product.html", current=current)

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder = f"{product_id}_{product_serial}_{ts}"
        pdir = os.path.join(PRODUCTS_DIR, folder)
        os.makedirs(pdir, exist_ok=True)

        session = {
            "session_id": folder,
            "product_serial": product_serial,
            "operator": operator,
            "notes": notes,
            "status": "CREATED",
            "started_at": datetime.datetime.now().timestamp(),
            "started_at_iso": datetime.datetime.now().isoformat(timespec="seconds"),
            "product_context": {
                "product_id": product_id, "product_name": product_name,
                "product_revision": product_revision,
                "recipe_status": "DRAFT", "test_plan_id": plan_id,
                "test_plan_revision": plan_revision, "test_plan_mode": "MANUAL",
            },
            "test_plan": {
                "plan_id": plan_id, "plan_name": plan_name,
                "plan_revision": plan_revision, "total_steps": 0,
                "observed_steps": 0,
            },
            "inspection_outcome": {
                "code": "PENDING", "accepted": None, "rejected": False,
                "result_count": 0, "fail_count": 0, "error_count": 0,
            },
        }
        with open(os.path.join(pdir, "session.json"), "w", encoding="utf-8") as f:
            json.dump(session, f, ensure_ascii=False, indent=2)
        with open(os.path.join(pdir, "results.jsonl"), "w", encoding="utf-8") as f:
            pass

        prod = _detail_from_parts(session, [], [], [], {}, {}, [])
        xlsx = build_excel_bytes(prod)
        with open(os.path.join(pdir, "report.xlsx"), "wb") as f:
            f.write(xlsx.getvalue())

        flash(f"پوشه محصول '{folder}' ساخته شد و گزارش اولیه ذخیره شد.", "success")
        return redirect(url_for("product_detail", pid=folder))

    return render_template("create_product.html", current=current["summary"])


@app.route("/products/create-from-current", methods=["POST"])
def create_from_current():
    current = _build_current()
    sid = current["session"].get("session_id")
    if not sid:
        flash("جلسه تست جاری یافت نشد.", "error")
        return redirect(url_for("create_product"))
    target_dir = os.path.join(PRODUCTS_DIR, sid)
    os.makedirs(target_dir, exist_ok=True)
    already = os.path.exists(os.path.join(target_dir, "session.json"))
    for fname in ["session.json", "results.jsonl", "events.jsonl", "packets.jsonl",
                  "test_plan_shadow.json", "placement_shadow.json", "errors.json",
                  "report_status.json"]:
        src = os.path.join(BASE_DIR, fname)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(target_dir, fname))

    prod = get_product(sid)
    xlsx = build_excel_bytes(prod)
    with open(os.path.join(target_dir, "report.xlsx"), "wb") as f:
        f.write(xlsx.getvalue())

    if already:
        flash(f"جلسه تست '{sid}' به‌روزرسانی شد و گزارش اکسل بازسازی شد.", "info")
    else:
        flash(f"جلسه تست '{sid}' در پوشه محصول ذخیره شد و گزارش اکسل تولید شد.", "success")
    return redirect(url_for("product_detail", pid=sid))


@app.route("/products/<pid>/download/<filename>")
def download_file(pid, filename):
    pdir = os.path.join(PRODUCTS_DIR, pid)
    if not os.path.isdir(pdir):
        abort(404)
    path = os.path.join(pdir, filename)
    if not os.path.isfile(path):
        abort(404)
    return send_file(path, as_attachment=True)


@app.errorhandler(404)
def _404(e):
    return render_template("base.html", title="یافت نشد"), 404


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
