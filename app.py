import os
import json
import time
import datetime
import requests
import re
import tempfile
from urllib.parse import urlparse
from bs4 import BeautifulSoup
from flask import Flask, render_template, request, jsonify, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ── Storage: use /tmp on Railway (ephemeral), or local data/ folder ──
DATA_DIR  = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(__file__), "data"))
os.makedirs(DATA_DIR, exist_ok=True)
DATA_FILE  = os.path.join(DATA_DIR, "history.json")
EXCEL_FILE = os.path.join(DATA_DIR, "traffic_report.xlsx")

# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def load_history():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_history(history):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

def clean_url(url: str) -> str:
    url = url.strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    return url

def get_domain(url: str) -> str:
    parsed = urlparse(clean_url(url))
    return parsed.netloc.replace("www.", "")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

# ─────────────────────────────────────────────
# Parsers
# ─────────────────────────────────────────────

def fetch_similarweb_public(domain: str) -> dict:
    result = {
        "source": "SimilarWeb (public)",
        "monthly_visits": None,
        "bounce_rate": None,
    }
    try:
        resp = requests.get(
            f"https://www.similarweb.com/website/{domain}/",
            headers=HEADERS, timeout=12
        )
        if resp.status_code != 200:
            result["error"] = f"HTTP {resp.status_code}"
            return result
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup.find_all(["span", "p", "div"]):
            text = tag.get_text(strip=True)
            if re.search(r'\d[\d,\.]+[KkMmBb]?\s*(visits|визитов)', text, re.I):
                m = re.search(r'([\d,\.]+[KkMmBb]?)', text)
                if m:
                    result["monthly_visits"] = m.group(1)
                    break
        for tag in soup.find_all(["span", "p", "div"]):
            text = tag.get_text(strip=True)
            if re.search(r'bounce\s*rate', text, re.I):
                m = re.search(r'(\d+[\.,]?\d*)\s*%', text)
                if m:
                    result["bounce_rate"] = m.group(1) + "%"
                    break
    except requests.exceptions.RequestException as e:
        result["error"] = str(e)
    return result

def fetch_pr_cy(domain: str) -> dict:
    result = {"source": "PR-CY", "traffic_day": None, "traffic_month": None}
    try:
        resp = requests.get(
            f"https://pr-cy.ru/site/{domain}/",
            headers=HEADERS, timeout=12
        )
        if resp.status_code != 200:
            result["error"] = f"HTTP {resp.status_code}"
            return result
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup.find_all(["td", "span", "div", "p"]):
            text = tag.get_text(strip=True)
            if "в день" in text.lower():
                m = re.search(r'([\d\s]+)', text)
                if m:
                    result["traffic_day"] = m.group(1).strip()
            if "в месяц" in text.lower():
                m = re.search(r'([\d\s]+)', text)
                if m:
                    result["traffic_month"] = m.group(1).strip()
    except requests.exceptions.RequestException as e:
        result["error"] = str(e)
    return result

def fetch_site_meta(url: str) -> dict:
    result = {
        "title": None, "description": None,
        "status_code": None, "server": None,
        "response_time_ms": None,
        "https": url.startswith("https"),
    }
    try:
        start = time.time()
        resp = requests.get(url, headers=HEADERS, timeout=12, allow_redirects=True)
        elapsed = round((time.time() - start) * 1000)
        result["status_code"] = resp.status_code
        result["response_time_ms"] = elapsed
        result["server"] = resp.headers.get("Server", "—")
        soup = BeautifulSoup(resp.text, "html.parser")
        t = soup.find("title")
        if t:
            result["title"] = t.get_text(strip=True)[:120]
        md = soup.find("meta", attrs={"name": "description"})
        if md:
            result["description"] = md.get("content", "")[:200]
    except requests.exceptions.RequestException as e:
        result["error"] = str(e)
    return result

def analyze_site(url: str) -> dict:
    url       = clean_url(url)
    domain    = get_domain(url)
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    meta       = fetch_site_meta(url)
    similarweb = fetch_similarweb_public(domain)
    pr_cy      = fetch_pr_cy(domain)
    return {
        "url": url, "domain": domain, "timestamp": timestamp,
        "meta": meta, "similarweb": similarweb, "pr_cy": pr_cy,
        "title":            meta.get("title"),
        "status_code":      meta.get("status_code"),
        "response_time_ms": meta.get("response_time_ms"),
        "monthly_visits": (
            similarweb.get("monthly_visits") or
            pr_cy.get("traffic_month") or "Нет данных"
        ),
        "bounce_rate": similarweb.get("bounce_rate") or "—",
        "traffic_day": pr_cy.get("traffic_day") or "—",
    }

# ─────────────────────────────────────────────
# Excel builder
# ─────────────────────────────────────────────

def build_excel(history: list, filepath: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Traffic Report"

    hf   = Font(bold=True, color="FFFFFF", size=11)
    hfil = PatternFill("solid", fgColor="2563EB")
    ctr  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    lft  = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin = Side(style="thin", color="D1D5DB")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt  = PatternFill("solid", fgColor="EFF6FF")

    cols = [
        ("Дата проверки", 20), ("URL / Домен", 30), ("Заголовок сайта", 35),
        ("Статус HTTP", 13),   ("Время ответа (мс)", 18), ("Визитов/месяц", 18),
        ("Показатель отказов", 20), ("Трафик/день (PR-CY)", 20), ("Источник данных", 20),
    ]
    for ci, (name, w) in enumerate(cols, 1):
        c = ws.cell(row=1, column=ci, value=name)
        c.font = hf; c.fill = hfil; c.alignment = ctr; c.border = brd
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 30

    for ri, rec in enumerate(history, 2):
        fill = alt if ri % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        vals = [
            rec.get("timestamp", ""),
            rec.get("domain", rec.get("url", "")),
            rec.get("title") or "—",
            rec.get("status_code") or "—",
            rec.get("response_time_ms") or "—",
            rec.get("monthly_visits") or "—",
            rec.get("bounce_rate") or "—",
            rec.get("traffic_day") or "—",
            rec.get("similarweb", {}).get("source", "—"),
        ]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=v)
            c.alignment = lft if ci in (2, 3) else ctr
            c.fill = fill; c.border = brd
        ws.row_dimensions[ri].height = 22

    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("Сводка")
    ws2["A1"] = "Всего проверок:"; ws2["B1"] = len(history)
    ws2["A2"] = "Обновлено:";      ws2["B2"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws2["A1"].font = Font(bold=True); ws2["A2"].font = Font(bold=True)
    ws2.column_dimensions["A"].width = 25
    ws2.column_dimensions["B"].width = 25

    wb.save(filepath)

# ─────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/analyze", methods=["POST"])
def analyze():
    data = request.get_json()
    urls = data.get("urls", [])
    if not urls:
        return jsonify({"error": "Не указаны URL"}), 400

    history = load_history()
    results = []
    for url in urls:
        url = url.strip()
        if not url:
            continue
        try:
            rec = analyze_site(url)
            history.append(rec)
            results.append(rec)
        except Exception as e:
            results.append({
                "url": url, "error": str(e),
                "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

    save_history(history)
    # Rebuild Excel in background path
    try:
        build_excel(history, EXCEL_FILE)
    except Exception:
        pass

    return jsonify({"results": results, "total_history": len(history)})

@app.route("/history")
def get_history():
    return jsonify(load_history())

@app.route("/export")
def export():
    history = load_history()
    if not history:
        return jsonify({"error": "Нет данных для экспорта"}), 400
    # Build into a temp file so concurrent requests don't conflict
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    build_excel(history, tmp.name)
    return send_file(
        tmp.name,
        as_attachment=True,
        download_name="traffic_report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/clear", methods=["POST"])
def clear():
    save_history([])
    if os.path.exists(EXCEL_FILE):
        os.remove(EXCEL_FILE)
    return jsonify({"ok": True})

@app.route("/health")
def health():
    return jsonify({"status": "ok", "time": datetime.datetime.now().isoformat()})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
