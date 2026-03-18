#!/usr/bin/env python3
"""
Solar Dashboard — Octopus Energy + Solis Cloud
================================================
Combines Octopus export metering with Solis inverter data
into a single HTML dashboard, optionally emailed.

All credentials are in the CONFIG dict below.
Schedule: 0 8 1 * * /usr/bin/python3 /path/to/solar_dashboard.py
"""

import hashlib, hmac, base64, json, time, os, sys, ssl
import smtplib, urllib.request, urllib.error
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from http.client import HTTPSConnection
from urllib.parse import urlparse
from datetime import datetime, date
from collections import defaultdict

SSL_CTX = ssl.create_default_context()
try:
    import certifi
    SSL_CTX.load_verify_locations(certifi.where())
except Exception:
    SSL_CTX.check_hostname = False
    SSL_CTX.verify_mode = ssl.CERT_NONE

# ═══════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════

CONFIG = {
    # Octopus Energy
    "octopus_api_key":    os.getenv("OCTOPUS_API_KEY"),
    "octopus_export_mpan": os.getenv("OCTOPUS_EXPORT_MPAN"),
    "octopus_export_serial": os.getenv("OCTOPUS_EXPORT_SERIAL"),

    # Solis Cloud
    "solis_api_url":      os.getenv("SOLIS_API_URL"),
    "solis_key_id":       os.getenv("SOLIS_KEY_ID"),
    "solis_key_secret":   os.getenv("SOLIS_KEY_SECRET"),

    # Email (leave smtp_password blank to skip emailing)
    "smtp_server":   os.getenv("SMTP_SERVER"),
    "smtp_port":     int(os.getenv("SMTP_PORT")),
    "smtp_user":      os.getenv("SMTP_USER"),
    "smtp_password":  os.getenv("SMTP_PASSWORD"),
    "email_to":       os.getenv("EMAIL_TO"),

    # Rates
    "export_rate_gbp": float(os.getenv("EXPORT_RATE_GBP", "0.15")),
}


# ═══════════════════════════════════════════════════════════════════════
# OCTOPUS ENERGY API
# ═══════════════════════════════════════════════════════════════════════

def fetch_octopus_data(cfg):
    """Fetch all available export data from Octopus Energy."""
    api_key = cfg["octopus_api_key"]
    mpan = cfg["octopus_export_mpan"]
    serial = cfg["octopus_export_serial"]

    all_results = []
    url = (
        f"https://api.octopus.energy/v1/electricity-meter-points/{mpan}"
        f"/meters/{serial}/consumption/"
        f"?period_from=2015-01-01T00:00:00Z"
        f"&period_to={datetime.now().strftime('%Y-%m-%dT23:59:59Z')}"
        f"&page_size=25000&order_by=period"
    )
    credentials = base64.b64encode(f"{api_key}:".encode()).decode()

    while url:
        print(f"  Octopus: fetching ({len(all_results)} records so far)...", flush=True)
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Basic {credentials}")
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            print(f"  Octopus API Error {e.code}: {e.reason}")
            return None
        except (urllib.error.URLError, TimeoutError) as e:
            print(f"  Octopus connection error: {e}")
            return None

        all_results.extend(data.get("results", []))
        url = data.get("next")

    print(f"  Octopus: ✓ {len(all_results)} readings")
    return all_results


def process_octopus_data(results):
    current_year = datetime.now().year
    monthly = defaultdict(float)
    yearly = defaultdict(float)
    daily = defaultdict(float)

    for r in results:
        dt = datetime.fromisoformat(r["interval_start"].replace("Z", "+00:00"))
        kwh = r["consumption"]
        daily[dt.date()] += kwh
        if dt.year == current_year:
            monthly[dt.month] += kwh
        else:
            yearly[dt.year] += kwh

    best_day = max(daily.items(), key=lambda x: x[1]) if daily else (None, 0)
    return {
        "current_year": current_year,
        "monthly": dict(monthly),
        "yearly": dict(sorted(yearly.items(), reverse=True)),
        "best_day_date": best_day[0],
        "best_day_kwh": best_day[1],
    }


# ═══════════════════════════════════════════════════════════════════════
# SOLIS CLOUD API
# ═══════════════════════════════════════════════════════════════════════

class SolisAPI:
    def __init__(self, api_url, key_id, key_secret):
        p = urlparse(api_url)
        self.host = p.hostname
        self.port = p.port or 13333
        self.key_id = key_id
        self.key_secret = key_secret

    def _sign(self, bj, path):
        md5 = base64.b64encode(hashlib.md5(bj.encode()).digest()).decode()
        ct = "application/json"
        dt = time.strftime("%a, %d %b %Y %H:%M:%S GMT", time.gmtime())
        sig = base64.b64encode(hmac.new(
            self.key_secret.encode(),
            f"POST\n{md5}\n{ct}\n{dt}\n{path}".encode(),
            hashlib.sha1
        ).digest()).decode()
        return {"Content-Type": ct, "Content-MD5": md5, "Date": dt,
                "Authorization": f"API {self.key_id}:{sig}"}

    def _post(self, path, body):
        bj = json.dumps(body)
        h = self._sign(bj, path)
        c = HTTPSConnection(self.host, self.port, timeout=30, context=SSL_CTX)
        try:
            c.request("POST", path, body=bj, headers=h)
            return json.loads(c.getresponse().read().decode())
        finally:
            c.close()

    def get_station_list(self):
        return self._post("/v1/api/userStationList", {"pageNo": 1, "pageSize": 20})

    def get_inverter_list(self, sid):
        return self._post("/v1/api/inverterList", {"stationId": sid, "pageNo": 1, "pageSize": 20})

    def get_inverter_detail(self, iid):
        return self._post("/v1/api/inverterDetail", {"id": iid})

    def get_inverter_all(self, sn, money):
        return self._post("/v1/api/inverterAll", {"sn": sn, "money": money})


def safe_float(v):
    try:
        return float(v) if v is not None else None
    except (ValueError, TypeError):
        return None


def fetch_solis_data(cfg):
    api = SolisAPI(cfg["solis_api_url"], cfg["solis_key_id"], cfg["solis_key_secret"])
    rate = cfg["export_rate_gbp"]

    sr = api.get_station_list()
    stns = (sr.get("data") or {}).get("page", {}).get("records", [])
    if not stns:
        print("  Solis: no stations found.")
        return None

    results = []
    for stn in stns:
        sid = str(stn.get("id", ""))
        sname = stn.get("stationName", stn.get("sno", f"Station {sid}"))
        print(f"  Solis: station '{sname}'")

        ir = api.get_inverter_list(sid)
        invs = (ir.get("data") or {}).get("page", {}).get("records", [])
        if not invs:
            continue

        for inv in invs:
            iid = str(inv.get("id", ""))
            isn = inv.get("sn", "")
            print(f"    Inverter: {isn}")

            det = (api.get_inverter_detail(iid).get("data") or {})

            # This month
            this_month = {
                "yield": safe_float(det.get("eMonth")),
                "import": safe_float(det.get("gridPurchasedMonthEnergy")),
                "export": safe_float(det.get("gridSellMonthEnergy")),
            }
            exp = this_month["export"]
            this_month["earnings"] = round(exp * rate, 2) if exp is not None else None

            # This year
            this_year = {
                "yield": safe_float(det.get("eYear")),
                "import": safe_float(det.get("gridPurchasedYearEnergy")),
                "export": safe_float(det.get("gridSellYearEnergy")),
            }
            ey = this_year["export"]
            this_year["earnings"] = round(ey * rate, 2) if ey is not None else None

            # Today's activity
            today = {
                "power_kw": safe_float(det.get("pac")),
                "power_pec": det.get("pacPec", "1"),
                "yield": safe_float(det.get("eToday")),
                "export": safe_float(det.get("gridSellTodayEnergy")),
                "import": safe_float(det.get("gridPurchasedTodayEnergy")),
                "family_load": safe_float(det.get("familyLoadPower")),
                "family_load_pec": det.get("familyLoadPowerPec", "1"),
            }
            te = today["export"]
            today["earnings"] = round(te * rate, 2) if te is not None else None

            # Live PV strings (only meaningful during generation)
            live_pv = {
                "uPv1": safe_float(det.get("uPv1")),
                "iPv1": safe_float(det.get("iPv1")),
                "uPv2": safe_float(det.get("uPv2")),
                "iPv2": safe_float(det.get("iPv2")),
            }

            # Per-year totals from inverterAll
            yearly = {}
            try:
                resp = api.get_inverter_all(isn, str(date.today().year))
                records = resp.get("data") or []
                if isinstance(records, list):
                    for rec in records:
                        y = rec.get("year")
                        if y is None:
                            continue
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
            except Exception as e:
                print(f"    inverterAll error: {e}")

            yearly[date.today().year] = this_year

            results.append({
                "station_name": sname,
                "inverter_sn": isn,
                "this_month": this_month,
                "today": today,
                "live_pv": live_pv,
                "yearly": yearly,
            })

    print(f"  Solis: ✓ {len(results)} inverter(s)")
    return results


# ═══════════════════════════════════════════════════════════════════════
# HTML DASHBOARD
# ═══════════════════════════════════════════════════════════════════════

def fmt(v, unit=""):
    if v is None:
        return "—"
    if unit == "£":
        return f"£{v:,.2f}"
    return f"{v:,.2f} {unit}".strip()


def apply_pec(val, pec):
    """Apply Solis power multiplier (pacPec). '1'=kW, '2'=W, etc."""
    if val is None:
        return None, "kW"
    try:
        pec = str(pec)
        if pec == "2":
            return val / 1000.0, "kW"
    except (ValueError, TypeError):
        pass
    return val, "kW"


def generate_dashboard(octopus, solis, cfg):
    now = datetime.now()
    current_year = now.year
    rate = cfg["export_rate_gbp"]
    month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                   "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    month_label = now.strftime("%B %Y")

    # ── Octopus stats ──
    oct_monthly_kwh = [round(octopus["monthly"].get(m, 0), 2) for m in range(1, 13)]
    oct_monthly_earn = [round(v * rate, 2) for v in oct_monthly_kwh]
    oct_ytd = sum(oct_monthly_kwh)
    oct_ytd_earn = oct_ytd * rate
    oct_prev_kwh = sum(octopus["yearly"].values())
    oct_grand_kwh = oct_ytd + oct_prev_kwh
    oct_grand_earn = oct_grand_kwh * rate
    best_kwh = octopus["best_day_kwh"]
    best_earn = best_kwh * rate
    best_label = octopus["best_day_date"].strftime('%A %d %b %Y') if octopus["best_day_date"] else 'N/A'

    # ── Solis stats (first inverter) ──
    sol = solis[0] if solis else None
    sol_month_yield = fmt(sol["this_month"]["yield"], "kWh") if sol else "—"
    sol_month_import = fmt(sol["this_month"]["import"], "kWh") if sol else "—"
    sol_month_export = fmt(sol["this_month"]["export"], "kWh") if sol else "—"
    sol_month_earn = fmt(sol["this_month"]["earnings"], "£") if sol else "—"

    # Today's activity
    today_html = ""
    if sol:
        td = sol["today"]
        pac_val, pac_unit = apply_pec(td["power_kw"], td["power_pec"])
        load_val, load_unit = apply_pec(td["family_load"], td["family_load_pec"])
        generating = pac_val is not None and pac_val > 0

        status_dot = "●" if generating else "○"
        status_text = "Generating" if generating else "Idle"
        status_color = "var(--green)" if generating else "var(--muted)"

        today_html = f"""
  <div class="card">
    <h2>Today's Activity <span style="float:right;font-size:0.8rem;color:{status_color}">{status_dot} {status_text}</span></h2>
    <div class="metric-grid">
      <div class="metric">
        <span class="metric-label">Current Output</span>
        <span class="metric-val">{fmt(pac_val, pac_unit) if pac_val else '—'}</span>
      </div>
      <div class="metric">
        <span class="metric-label">Home Consumption</span>
        <span class="metric-val">{fmt(load_val, load_unit) if load_val else '—'}</span>
      </div>
      <div class="metric">
        <span class="metric-label">Generated Today</span>
        <span class="metric-val">{fmt(td['yield'], 'kWh')}</span>
      </div>
      <div class="metric">
        <span class="metric-label">Exported Today</span>
        <span class="metric-val">{fmt(td['export'], 'kWh')}</span>
      </div>
      <div class="metric">
        <span class="metric-label">Imported Today</span>
        <span class="metric-val">{fmt(td['import'], 'kWh')}</span>
      </div>
      <div class="metric">
        <span class="metric-label">Export Earnings Today</span>
        <span class="metric-val">{fmt(td['earnings'], '£')}</span>
      </div>
    </div>"""

        # PV string detail — only show if panels are active
        pv = sol["live_pv"]
        has_pv = any(v and v > 0 for v in pv.values())
        if has_pv:
            today_html += f"""
    <div class="pv-bar">
      <span>PV1: {pv['uPv1']:.0f}V / {pv['iPv1']:.1f}A</span>
      <span>PV2: {pv['uPv2']:.0f}V / {pv['iPv2']:.1f}A</span>
    </div>"""

        today_html += "\n  </div>"

    # Solis yearly table
    solis_yearly_html = ""
    if sol and sol["yearly"]:
        sorted_years = sorted(sol["yearly"].keys(), reverse=True)
        rows = ""
        for y in sorted_years:
            d = sol["yearly"][y]
            yr_label = f"{y} (YTD)" if y == current_year else str(y)
            rows += f'        <tr><td>{yr_label}</td><td class="mono">{fmt(d.get("yield"), "kWh")}</td><td class="mono">{fmt(d.get("import"), "kWh")}</td><td class="mono">{fmt(d.get("export"), "kWh")}</td><td class="mono">{fmt(d.get("earnings"), "£")}</td></tr>\n'

        solis_yearly_html = f"""
  <div class="card">
    <h2>Solis — Yearly Totals</h2>
    <table>
      <thead>
        <tr><th>Year</th><th>Generation</th><th>Import</th><th>Export</th><th>Export Earnings</th></tr>
      </thead>
      <tbody>
{rows}      </tbody>
    </table>
  </div>"""

    # ── Full HTML ──
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Solar Dashboard — {month_label}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root {{
    --bg: #f4f5f7;
    --surface: #ffffff;
    --border: #e2e4e9;
    --text: #1a1d23;
    --text-secondary: #5f6577;
    --muted: #9ca0ad;
    --green: #0d9668;
    --green-bg: #ecfdf3;
    --amber: #b45309;
    --amber-bg: #fffbeb;
    --blue: #2563eb;
    --blue-bg: #eff6ff;
    --rose: #be123c;
    --rose-bg: #fff1f2;
    --purple: #7c3aed;
    --radius: 10px;
    --shadow: 0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04);
  }}
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{
    font-family: 'Inter', -apple-system, sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    padding: 2rem;
    -webkit-font-smoothing: antialiased;
  }}
  .wrap {{ max-width: 960px; margin: 0 auto; }}

  .header {{
    margin-bottom: 2rem;
  }}
  .header h1 {{
    font-size: 1.5rem;
    font-weight: 700;
    letter-spacing: -0.02em;
    color: var(--text);
  }}
  .header p {{
    color: var(--text-secondary);
    font-size: 0.85rem;
    margin-top: 0.25rem;
  }}

  .section-label {{
    font-size: 0.65rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: var(--muted);
    margin: 2rem 0 0.75rem;
  }}

  .kpis {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 0.75rem;
    margin-bottom: 0.5rem;
  }}
  .kpi {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.1rem 1.2rem;
    box-shadow: var(--shadow);
  }}
  .kpi .label {{
    font-size: 0.68rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: var(--text-secondary);
    margin-bottom: 0.35rem;
  }}
  .kpi .val {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.4rem;
    font-weight: 600;
  }}
  .kpi .sub {{
    font-size: 0.78rem;
    color: var(--muted);
    margin-top: 0.2rem;
  }}
  .green {{ color: var(--green); }}
  .amber {{ color: var(--amber); }}
  .blue {{ color: var(--blue); }}
  .rose {{ color: var(--rose); }}
  .purple {{ color: var(--purple); }}

  .card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.4rem;
    margin-bottom: 1rem;
    box-shadow: var(--shadow);
  }}
  .card h2 {{
    font-size: 0.8rem;
    font-weight: 600;
    color: var(--text-secondary);
    margin-bottom: 1rem;
    text-transform: uppercase;
    letter-spacing: 0.04em;
  }}

  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.85rem;
  }}
  th {{
    text-align: left;
    font-weight: 600;
    color: var(--muted);
    font-size: 0.68rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    padding: 0.55rem 0.8rem;
    border-bottom: 2px solid var(--border);
  }}
  td {{
    padding: 0.6rem 0.8rem;
    border-bottom: 1px solid var(--border);
  }}
  td.mono {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.82rem;
  }}
  tr:last-child td {{ border-bottom: none; }}
  tr.total td {{
    font-weight: 700;
    border-top: 2px solid var(--border);
    padding-top: 0.8rem;
  }}

  .metric-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
    gap: 0.6rem;
  }}
  .metric {{
    display: flex;
    flex-direction: column;
    gap: 0.2rem;
    padding: 0.7rem 0.8rem;
    background: var(--bg);
    border-radius: 8px;
  }}
  .metric-label {{
    font-size: 0.68rem;
    font-weight: 500;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 0.04em;
  }}
  .metric-val {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.95rem;
    font-weight: 500;
    color: var(--text);
  }}

  .pv-bar {{
    display: flex;
    gap: 1.5rem;
    margin-top: 0.8rem;
    padding: 0.5rem 0.8rem;
    background: var(--green-bg);
    border-radius: 6px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.78rem;
    color: var(--green);
  }}

  .footer {{
    text-align: center;
    margin-top: 2rem;
    padding-top: 1rem;
    border-top: 1px solid var(--border);
    font-size: 0.7rem;
    color: var(--muted);
  }}
</style>
</head>
<body>
<div class="wrap">

  <div class="header">
    <h1>Solar Dashboard</h1>
    <p>{month_label} · Export rate: £{rate:.2f}/kWh</p>
  </div>

  <!-- ── Solis Solar Panels ── -->
  <div class="section-label">Solis Solar Panels</div>
  <div class="kpis">
    <div class="kpi">
      <div class="label">Monthly Generation</div>
      <div class="val green">{sol_month_yield}</div>
    </div>
    <div class="kpi">
      <div class="label">Monthly Export</div>
      <div class="val blue">{sol_month_export}</div>
    </div>
    <div class="kpi">
      <div class="label">Monthly Import</div>
      <div class="val rose">{sol_month_import}</div>
    </div>
    <div class="kpi">
      <div class="label">Export Earnings</div>
      <div class="val amber">{sol_month_earn}</div>
    </div>
  </div>

  <!-- ── Today's Activity ── -->
{today_html}

  <!-- ── Octopus Metered Export ── -->
  <div class="section-label">Octopus Metered Export</div>
  <div class="kpis">
    <div class="kpi">
      <div class="label">{current_year} Year to Date</div>
      <div class="val green">{oct_ytd:,.1f} kWh</div>
      <div class="sub">£{oct_ytd_earn:,.2f} earned</div>
    </div>
    <div class="kpi">
      <div class="label">All-Time Total</div>
      <div class="val blue">{oct_grand_kwh:,.1f} kWh</div>
      <div class="sub">£{oct_grand_earn:,.2f} earned</div>
    </div>
    <div class="kpi">
      <div class="label">Best Ever Day</div>
      <div class="val green">{best_kwh:,.2f} kWh</div>
      <div class="sub">£{best_earn:,.2f} · {best_label}</div>
    </div>
  </div>

  <!-- ── Octopus Monthly ── -->
  <div class="card">
    <h2>Octopus — {current_year} Monthly Export</h2>
    <table>
      <thead>
        <tr><th>Month</th><th>Exported (kWh)</th><th>Earned (£)</th></tr>
      </thead>
      <tbody>
"""

    for i in range(12):
        kwh = oct_monthly_kwh[i]
        if kwh > 0:
            earn = oct_monthly_earn[i]
            html += f'        <tr><td>{month_names[i]}</td><td class="mono">{kwh:,.2f}</td><td class="mono">£{earn:,.2f}</td></tr>\n'

    html += f"""        <tr class="total"><td>Total</td><td class="mono green">{oct_ytd:,.2f}</td><td class="mono green">£{oct_ytd_earn:,.2f}</td></tr>
      </tbody>
    </table>
  </div>

  <!-- ── Octopus Previous Years ── -->
  <div class="card">
    <h2>Octopus — Previous Years</h2>
    <table>
      <thead>
        <tr><th>Year</th><th>Exported (kWh)</th><th>Earned (£)</th></tr>
      </thead>
      <tbody>
"""

    for year, kwh in octopus["yearly"].items():
        earn = kwh * rate
        html += f'        <tr><td>{year}</td><td class="mono">{kwh:,.2f}</td><td class="mono">£{earn:,.2f}</td></tr>\n'

    if octopus["yearly"]:
        html += f'        <tr class="total"><td>Total</td><td class="mono blue">{oct_prev_kwh:,.2f}</td><td class="mono blue">£{oct_prev_kwh * rate:,.2f}</td></tr>\n'
    else:
        html += '        <tr><td colspan="3" style="color:var(--muted)">No previous year data found</td></tr>\n'

    html += f"""      </tbody>
    </table>
  </div>

  <!-- ── Solis Yearly ── -->
{solis_yearly_html}

  <div class="footer">Generated {now.strftime('%d %b %Y at %H:%M')} · Octopus Energy API + Solis Cloud API</div>
</div>
</body>
</html>"""
    return html


# ═══════════════════════════════════════════════════════════════════════
# EMAIL
# ═══════════════════════════════════════════════════════════════════════

def send_email(subject, html_body, cfg):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = cfg["smtp_user"]
    msg["To"] = cfg["email_to"]
    msg.attach(MIMEText("View in an HTML-capable email client.", "plain"))
    msg.attach(MIMEText(html_body, "html"))
    print(f"\nSending to {cfg['email_to']} ...", flush=True)
    with smtplib.SMTP(cfg["smtp_server"], cfg["smtp_port"]) as s:
        s.ehlo()
        s.starttls()
        s.ehlo()
        s.login(cfg["smtp_user"], cfg["smtp_password"])
        s.sendmail(cfg["smtp_user"], cfg["email_to"], msg.as_string())
    print("  ✓ Email sent!")


# ═══════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════

def main():
    month_label = datetime.now().strftime("%B %Y")
    print(f"{'═' * 50}")
    print(f"  Solar Dashboard — {month_label}")
    print(f"{'═' * 50}\n")

    print("[1/2] Octopus Energy export data", flush=True)
    oct_raw = fetch_octopus_data(CONFIG)
    if not oct_raw:
        print("  ✘ No Octopus data — check credentials.")
        sys.exit(1)
    octopus = process_octopus_data(oct_raw)

    print("\n[2/2] Solis Cloud inverter data", flush=True)
    solis = fetch_solis_data(CONFIG)
    if not solis:
        print("  ⚠ No Solis data — dashboard will show Octopus only.")
        solis = []

    print("\nGenerating dashboard...", flush=True)
    html = generate_dashboard(octopus, solis, CONFIG)

    output_file = f"solar_dashboard_{datetime.now().strftime('%Y_%m')}.html"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  ✓ Saved: {output_file}")

    if CONFIG["smtp_password"]:
        send_email(f"☀️ Solar Dashboard — {month_label}", html, CONFIG)
    else:
        print("  ⚠ No SMTP password — email skipped.")

    print("\nDone.\n")


if __name__ == "__main__":
    main()
