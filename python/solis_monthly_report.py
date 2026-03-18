#!/usr/bin/env python3
"""
Solis Cloud – Monthly Solar Panel Report (v4)
==============================================
Uses inverterDetail for current month/year data + live PV.
Uses inverterAll for per-year totals.

Report layout:
  1. Summary tiles: this month's figures
  2. Table: this year + all previous years with % change

Schedule: 0 8 1 * * /usr/bin/python3 /path/to/solis_monthly_report.py
Requirements: Python 3.8+ (stdlib only)
"""

import hashlib, hmac, base64, json, time, datetime, calendar
import smtplib, os, sys, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from http.client import HTTPSConnection
from urllib.parse import urlparse

SSL_CTX = ssl.create_default_context()
try:
    import certifi
    SSL_CTX.load_verify_locations(certifi.where())
except Exception:
    SSL_CTX.check_hostname = False
    SSL_CTX.verify_mode = ssl.CERT_NONE

CONFIG = {
    "api_url":        os.getenv("SOLIS_API_URL"),
    "api_key_id":     os.getenv("SOLIS_KEY_ID"),
    "api_key_secret": os.getenv("SOLIS_KEY_SECRET"),
    "smtp_server":    os.getenv("SMTP_SERVER"),
    "smtp_port":      int(os.getenv("SMTP_PORT")),
    "smtp_user":      os.getenv("SMTP_USER"),
    "smtp_password":  os.getenv("SMTP_PASSWORD"),
    "email_to":       os.getenv("EMAIL_TO"),
    "timezone_offset": int(os.getenv("TZ_OFFSET", "0")),
    "export_rate_gbp": float(os.getenv("EXPORT_RATE_GBP", "0.15")),
}

# ── API Client ──
class SolisAPI:
    def __init__(self, api_url, key_id, key_secret):
        p = urlparse(api_url); self.host = p.hostname; self.port = p.port or 13333
        self.key_id = key_id; self.key_secret = key_secret
    def _sign(self, bj, path):
        md5 = base64.b64encode(hashlib.md5(bj.encode()).digest()).decode()
        ct = "application/json"; dt = time.strftime("%a, %d %b %Y %H:%M:%S GMT", time.gmtime())
        sig = base64.b64encode(hmac.new(self.key_secret.encode(),
            f"POST\n{md5}\n{ct}\n{dt}\n{path}".encode(), hashlib.sha1).digest()).decode()
        return {"Content-Type":ct,"Content-MD5":md5,"Date":dt,"Authorization":f"API {self.key_id}:{sig}"}
    def _post(self, path, body):
        bj = json.dumps(body); h = self._sign(bj, path)
        c = HTTPSConnection(self.host, self.port, timeout=30, context=SSL_CTX)
        try:
            c.request("POST", path, body=bj, headers=h)
            return json.loads(c.getresponse().read().decode())
        finally:
            c.close()
    def get_station_list(self):
        return self._post("/v1/api/userStationList", {"pageNo":1,"pageSize":20})
    def get_inverter_list(self, sid):
        return self._post("/v1/api/inverterList", {"stationId":sid,"pageNo":1,"pageSize":20})
    def get_inverter_detail(self, iid):
        return self._post("/v1/api/inverterDetail", {"id":iid})
    def get_inverter_all(self, sn, money):
        return self._post("/v1/api/inverterAll", {"sn":sn,"money":money})

# ── Helpers ──
def safe_float(v):
    try: return float(v) if v is not None else None
    except: return None

def collect_data(api):
    results = []
    sr = api.get_station_list()
    stns = (sr.get("data") or {}).get("page",{}).get("records",[])
    if not stns:
        print("  ⚠  No stations found.", file=sys.stderr); return results
    print(f"Found {len(stns)} station(s)")

    for stn in stns:
        sid = str(stn.get("id",""))
        sname = stn.get("stationName", stn.get("sno", f"Station {sid}"))
        print(f"\n  Station: {sname}")

        ir = api.get_inverter_list(sid)
        invs = (ir.get("data") or {}).get("page",{}).get("records",[])
        if not invs:
            print("    ⚠  No inverters"); continue

        for inv in invs:
            iid = str(inv.get("id",""))
            isn = inv.get("sn","")
            print(f"    Inverter: {isn}")

            # ── inverterDetail: this month, this year, live PV ──
            print(f"      Fetching inverterDetail ...")
            det = (api.get_inverter_detail(iid).get("data") or {})
            rate = CONFIG["export_rate_gbp"]

            this_month = {
                "yield": safe_float(det.get("eMonth")),
                "import": safe_float(det.get("gridPurchasedMonthEnergy")),
                "export": safe_float(det.get("gridSellMonthEnergy")),
            }
            exp = this_month["export"]
            this_month["earnings"] = round(exp * rate, 2) if exp is not None else None

            this_year_detail = {
                "yield": safe_float(det.get("eYear")),
                "import": safe_float(det.get("gridPurchasedYearEnergy")),
                "export": safe_float(det.get("gridSellYearEnergy")),
            }
            ey = this_year_detail["export"]
            this_year_detail["earnings"] = round(ey * rate, 2) if ey is not None else None

            live_pv = {
                "uPv1": safe_float(det.get("uPv1")),
                "iPv1": safe_float(det.get("iPv1")),
                "uPv2": safe_float(det.get("uPv2")),
                "iPv2": safe_float(det.get("iPv2")),
            }

            print(f"        This month: yield={this_month['yield']}kWh, "
                  f"import={this_month['import']}kWh, export={this_month['export']}kWh")
            print(f"        This year:  yield={this_year_detail['yield']}kWh")
            print(f"        PV1: {live_pv['uPv1']}V/{live_pv['iPv1']}A  "
                  f"PV2: {live_pv['uPv2']}V/{live_pv['iPv2']}A")

            # ── inverterAll: per-year totals ──
            print(f"      Fetching inverterAll (yearly totals) ...")
            yearly = {}
            try:
                resp = api.get_inverter_all(isn, str(datetime.date.today().year))
                records = resp.get("data") or []
                if isinstance(records, list):
                    for rec in records:
                        y = rec.get("year")
                        if y is None: continue
                        y = int(y)
                        energy = safe_float(rec.get("energy"))
                        epec = safe_float(rec.get("energyPec"))
                        if energy is not None and epec is not None and epec < 1:
                            yield_kwh = round(energy * epec * 1000, 2)
                        else:
                            yield_kwh = round(energy, 2) if energy else None
                        gi = safe_float(rec.get("gridPurchasedEnergy"))
                        ge = safe_float(rec.get("gridSellEnergy"))
                        yearly[y] = {
                            "yield": yield_kwh,
                            "import": round(gi, 2) if gi else None,
                            "export": round(ge, 2) if ge else None,
                            "earnings": round(ge * rate, 2) if ge else None,
                        }
                        print(f"        {y}: yield={yield_kwh}kWh, import={gi}kWh, export={ge}kWh")
            except Exception as e:
                print(f"        ERROR: {e}")

            # Override current year in yearly with the more accurate inverterDetail data
            current_year = datetime.date.today().year
            yearly[current_year] = this_year_detail

            results.append({
                "station_name": sname,
                "inverter_sn": isn,
                "this_month": this_month,
                "yearly": yearly,
                "live_pv": live_pv,
            })

    return results

# ── HTML Report ──
def fmt(v, u=""):
    if v is None: return "—"
    if u.strip() == "£": return f"£{v:.2f}"
    return f"{v}{u}"

def build_html_report(data):
    now = datetime.datetime.now()
    month_name = now.strftime("%B %Y")
    inv_html = ""

    for blk in data:
        sn = blk["inverter_sn"]
        station = blk["station_name"]
        cm = blk["this_month"]
        yearly = blk["yearly"]
        live_pv = blk.get("live_pv", {})

        # ── 1. This Month tiles ──
        TILES = [
            ("Monthly Yield", cm["yield"], "kWh", "#e67e00", "#fef7e0"),
            ("Monthly Export Earnings", cm["earnings"], "£", "#2e7d32", "#e8f5e9"),
            ("Monthly Import", cm["import"], "kWh", "#c2255c", "#fce4ec"),
            ("Monthly Export", cm["export"], "kWh", "#6a1b9a", "#f3e5f5"),
            ("DC Voltage PV1", live_pv.get("uPv1"), "V", "#1a73e8", "#e8f0fe"),
            ("DC Voltage PV2", live_pv.get("uPv2"), "V", "#1557b0", "#e8f0fe"),
            ("DC Current PV1", live_pv.get("iPv1"), "A", "#0d8050", "#e6f4ea"),
            ("DC Current PV2", live_pv.get("iPv2"), "A", "#0a6b40", "#e6f4ea"),
        ]
        cards = ""
        for label, val, unit, col, bg in TILES:
            if unit == "£" and val is not None:
                dv = f"£{val:.2f}"; du = "&nbsp;"
            elif "DC" in label:
                dv = fmt(val); du = "Live" if val is not None else ""
            else:
                dv = fmt(val); du = unit if val is not None else ""
            cards += f'<div style="display:inline-block;width:190px;min-height:90px;padding:14px 8px;margin:5px;background:{bg};border-radius:10px;text-align:center;vertical-align:top"><div style="font-size:10px;color:#666;margin-bottom:3px">{label}</div><div style="font-size:22px;font-weight:700;color:{col}">{dv}</div><div style="font-size:11px;color:#888">{du}</div></div>'

        # ── PV snapshot ──
        pv_note = ""
        if live_pv.get("uPv1") is not None:
            pv_note = f'''<div style="margin:12px 0;padding:12px 16px;background:#f0f7ff;border-radius:8px;font-size:12px;color:#555">
                <strong>📡 Live PV Snapshot</strong> (current reading)<br>
                PV1: {live_pv.get("uPv1",0)}V / {live_pv.get("iPv1",0)}A &nbsp;&nbsp;·&nbsp;&nbsp;
                PV2: {live_pv.get("uPv2",0)}V / {live_pv.get("iPv2",0)}A
            </div>'''

        # ── 2. Yearly totals table ──
        sorted_years = sorted(yearly.keys(), reverse=True)
        current_year = datetime.date.today().year

        METRICS = [
            ("Yield", "yield", "kWh"),
            ("Import", "import", "kWh"),
            ("Export", "export", "kWh"),
            ("Export Earnings", "earnings", "£"),
        ]

        # Table header row
        yr_headers = ""
        for y in sorted_years:
            highlight = "font-weight:700" if y == current_year else ""
            label = "This Year" if y == current_year else str(y)
            yr_headers += f'<th style="padding:8px 12px;text-align:right;{highlight}">{label}</th>'

        # Table body rows
        yr_rows = ""
        for label, key, unit in METRICS:
            yr_rows += f'<tr><td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:500">{label}</td>'
            for y in sorted_years:
                val = yearly.get(y, {}).get(key)
                style = "font-weight:700" if y == current_year else ""
                yr_rows += f'<td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;{style}">{fmt(val, f" {unit}")}</td>'
            yr_rows += '</tr>'

        inv_html += f'''<div style="margin-bottom:36px">
            <h2 style="font-size:16px;color:#444;margin:20px 0 6px">🔌 {sn} <span style="font-weight:400;color:#999">({station})</span></h2>

            <h3 style="font-size:14px;color:#555;margin:16px 0 8px">📅 This Month — {month_name}</h3>
            <div style="text-align:center;padding:4px 0">{cards}</div>
            {pv_note}

            <h3 style="font-size:14px;color:#555;margin:24px 0 8px">📊 Yearly Totals</h3>
            <table style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr style="background:#f5f5f5">
                <th style="padding:8px 12px;text-align:left">Metric</th>{yr_headers}
            </tr></thead><tbody>{yr_rows}</tbody></table>
        </div>'''

    return f'''<!DOCTYPE html><html><head><meta charset="utf-8"></head>
    <body style="margin:0;padding:20px;font-family:'Segoe UI',Roboto,Arial,sans-serif;background:#f5f5f5;color:#333">
    <div style="max-width:1000px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08)">
        <div style="background:linear-gradient(135deg,#1a73e8,#0d47a1);padding:28px 32px;color:#fff">
            <h1 style="margin:0;font-size:24px">☀️ Monthly Solar Panel Report</h1>
            <p style="margin:6px 0 0;opacity:0.9;font-size:15px">{month_name}</p></div>
        <div style="padding:16px 28px 28px">{inv_html}</div>
        <div style="padding:14px 28px;background:#fafafa;font-size:11px;color:#aaa;border-top:1px solid #eee;text-align:center">
            Generated via Solis Cloud API · {now.strftime("%Y-%m-%d %H:%M")}</div>
    </div></body></html>'''

# ── Email ──
def send_email(subject, html_body, cfg):
    msg = MIMEMultipart("alternative")
    msg["Subject"], msg["From"], msg["To"] = subject, cfg["smtp_user"], cfg["email_to"]
    msg.attach(MIMEText("View in an HTML-capable email client.", "plain"))
    msg.attach(MIMEText(html_body, "html"))
    print(f"\nSending to {cfg['email_to']} ...")
    with smtplib.SMTP(cfg["smtp_server"], cfg["smtp_port"]) as s:
        s.ehlo(); s.starttls(); s.ehlo()
        s.login(cfg["smtp_user"], cfg["smtp_password"])
        s.sendmail(cfg["smtp_user"], cfg["email_to"], msg.as_string())
    print("  ✓ Email sent!")

# ── Main ──
def main():
    mn = datetime.datetime.now().strftime("%B %Y")
    print(f"{'═'*45}\n  Solis Monthly Report – {mn}\n{'═'*45}\n")

    api = SolisAPI(CONFIG["api_url"], CONFIG["api_key_id"], CONFIG["api_key_secret"])
    data = collect_data(api)
    if not data:
        print("\n✘ No data.", file=sys.stderr); sys.exit(1)

    html = build_html_report(data)
    path = f"solar_report_{datetime.date.today().strftime('%Y_%m')}.html"
    with open(path, "w") as f: f.write(html)
    print(f"\n  ✓ Saved to {path}")

    if CONFIG["smtp_password"]:
        send_email(f"☀️ Solar Report – {mn}", html, CONFIG)
    else:
        print("  ⚠  No SMTP password – email skipped.")

if __name__ == "__main__":
    main()
