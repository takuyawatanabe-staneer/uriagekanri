#!/usr/bin/env python3
"""売上管理ダッシュボード - Google Sheets連携"""

import json
import os
import urllib.request
import urllib.parse
from http.server import HTTPServer, SimpleHTTPRequestHandler

SHEET_ID = "17FfNTDve-vzL3n_byDV-RlncGvpQyrgdyzA5y6R29f0"
SHEET_NAME = "売上計画"
TOKEN_PATH = os.path.join(os.path.dirname(__file__), "..", "system", "token.json")
PORT = 8080


def refresh_token():
    """OAuth2トークンをリフレッシュして新しいアクセストークンを返す"""
    with open(TOKEN_PATH) as f:
        token_data = json.load(f)

    data = urllib.parse.urlencode({
        "client_id": token_data["client_id"],
        "client_secret": token_data["client_secret"],
        "refresh_token": token_data["refresh_token"],
        "grant_type": "refresh_token",
    }).encode()

    req = urllib.request.Request("https://oauth2.googleapis.com/token", data=data)
    resp = urllib.request.urlopen(req)
    return json.loads(resp.read())["access_token"]


def fetch_sheet_data(access_token):
    """スプレッドシートからデータを取得"""
    encoded_name = urllib.parse.quote(SHEET_NAME)
    url = (
        f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}"
        f"/values/{encoded_name}!A1:AX50"
    )
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {access_token}"})
    resp = urllib.request.urlopen(req)
    return json.loads(resp.read()).get("values", [])


def parse_yen(val):
    """¥1,234 形式の文字列を数値に変換"""
    if not val or val.strip() == "":
        return 0
    cleaned = val.replace("¥", "").replace(",", "").strip()
    try:
        return int(cleaned)
    except ValueError:
        try:
            return float(cleaned)
        except ValueError:
            return 0


def build_dashboard_data(rows):
    """スプレッドシートのデータをダッシュボード用JSONに変換"""
    if len(rows) < 2:
        return {}

    months = rows[1][2:]  # Row 2: headers (2024/01, 2024/02, ...)

    data = {
        "months": months,
        "入社売上_確定": {},
        "入社売上_ヨミ": {},
        "入社売上合計": [],
        "入金_確定_税込": {},
        "入金_ヨミ_税込": {},
        "入金合計_確定_税込": [],
        "入金合計_確定ヨミ_税込": [],
        "その他_入金": {},
    }

    for row in rows:
        if not row:
            continue
        label = row[0] if len(row) > 0 else ""
        person = row[1] if len(row) > 1 else ""
        values = [parse_yen(v) for v in row[2:]]

        # Pad values to match months length
        values += [0] * (len(months) - len(values))

        if label == "入社売上(確定)":
            data["入社売上_確定"][person] = values
        elif label == "入社売上(ヨミ)":
            data["入社売上_ヨミ"][person] = values
        elif label == "入社売上合計(確定+ヨミ)":
            data["入社売上合計"] = values
        elif label == "採用支援 ※税込" and person in ("ダイナトレック",):
            data["入社売上_確定"][f"採用支援({person})"] = values
        elif label == "CA支援 ※税込" and person in ("Ms. Engineer",):
            data["入社売上_確定"][f"CA支援({person})"] = values
        elif label == "入金(確定)※税込":
            data["入金_確定_税込"][person] = values
        elif label == "入金(ヨミ)※税込":
            data["入金_ヨミ_税込"][person] = values
        elif label == "入金合計(確定)※税込":
            data["入金合計_確定_税込"] = values
        elif label == "入金合計(確定＋ヨミ)※税込":
            data["入金合計_確定ヨミ_税込"] = values
        elif label == "サーカスRA":
            data["その他_入金"][f"サーカスRA({person})"] = values
        elif label == "採用支援 ※税込" and person in ("ダイナ",):
            data["その他_入金"][f"採用支援({person})"] = values
        elif label == "CA支援 ※税込" and person in ("Ms. Eng",):
            data["その他_入金"][f"CA支援({person})"] = values
        elif label == "研修※税込":
            data["その他_入金"][f"研修({person})"] = values

    return data


def generate_html(data):
    """ダッシュボードHTMLを生成"""
    json_data = json.dumps(data, ensure_ascii=False)
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>売上管理ダッシュボード</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', sans-serif; background: #0f172a; color: #e2e8f0; padding: 20px; }}
h1 {{ text-align: center; font-size: 1.8rem; margin-bottom: 8px; color: #38bdf8; }}
.subtitle {{ text-align: center; color: #64748b; margin-bottom: 24px; font-size: 0.9rem; }}
.kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 24px; }}
.kpi-card {{ background: #1e293b; border-radius: 12px; padding: 20px; border: 1px solid #334155; }}
.kpi-card .label {{ font-size: 0.8rem; color: #94a3b8; margin-bottom: 4px; }}
.kpi-card .value {{ font-size: 1.6rem; font-weight: bold; }}
.kpi-card .value.blue {{ color: #38bdf8; }}
.kpi-card .value.green {{ color: #4ade80; }}
.kpi-card .value.purple {{ color: #a78bfa; }}
.kpi-card .value.orange {{ color: #fb923c; }}
.chart-container {{ background: #1e293b; border-radius: 12px; padding: 20px; margin-bottom: 24px; border: 1px solid #334155; }}
.chart-container h2 {{ font-size: 1.1rem; margin-bottom: 16px; color: #cbd5e1; }}
.chart-wrapper {{ position: relative; height: 350px; }}
.detail-table {{ width: 100%; border-collapse: collapse; margin-top: 12px; font-size: 0.85rem; }}
.detail-table th, .detail-table td {{ padding: 8px 12px; text-align: right; border-bottom: 1px solid #334155; }}
.detail-table th {{ color: #94a3b8; font-weight: normal; text-align: left; position: sticky; top: 0; background: #1e293b; }}
.detail-table td:first-child, .detail-table th:first-child {{ text-align: left; position: sticky; left: 0; background: #1e293b; z-index: 1; }}
.detail-table tr:hover td {{ background: #253349; }}
.table-scroll {{ overflow-x: auto; max-height: 300px; overflow-y: auto; }}
.tabs {{ display: flex; gap: 8px; margin-bottom: 16px; }}
.tab {{ padding: 8px 16px; border-radius: 8px; cursor: pointer; background: #334155; color: #94a3b8; border: none; font-size: 0.85rem; }}
.tab.active {{ background: #38bdf8; color: #0f172a; font-weight: bold; }}
.section {{ margin-bottom: 32px; }}
</style>
</head>
<body>
<h1>売上管理ダッシュボード</h1>
<p class="subtitle">データソース: Google Sheets（売上計画） | 最終更新: <span id="updateTime"></span></p>

<div class="kpi-grid" id="kpiGrid"></div>

<div class="section">
  <div class="chart-container">
    <h2>入社売上合計（確定+ヨミ） - 月次推移</h2>
    <div class="chart-wrapper"><canvas id="salesChart"></canvas></div>
  </div>
  <div class="chart-container">
    <h2>入社売上 詳細（担当者別）</h2>
    <div class="tabs" id="salesTabs"></div>
    <div class="table-scroll" id="salesDetail"></div>
  </div>
</div>

<div class="section">
  <div class="chart-container">
    <h2>入金合計（税込） - 月次推移</h2>
    <div class="chart-wrapper"><canvas id="paymentChart"></canvas></div>
  </div>
  <div class="chart-container">
    <h2>入金 詳細（担当者別）</h2>
    <div class="tabs" id="paymentTabs"></div>
    <div class="table-scroll" id="paymentDetail"></div>
  </div>
</div>

<script>
const D = {json_data};

document.getElementById('updateTime').textContent = new Date().toLocaleString('ja-JP');

// Utility
const fmt = v => v === 0 ? '¥0' : '¥' + Math.round(v).toLocaleString('ja-JP');
const COLORS = ['#38bdf8','#4ade80','#fb923c','#a78bfa','#f472b6','#facc15','#34d399','#f87171'];

// Find current month index
const now = new Date();
const currentMonthStr = now.getFullYear() + '/' + String(now.getMonth()+1).padStart(2,'0');
const currentIdx = D.months.indexOf(currentMonthStr);
const prevIdx = currentIdx > 0 ? currentIdx - 1 : 0;

// Filter to show only months with data (non-zero)
function findLastNonZero(arr) {{
    for (let i = arr.length - 1; i >= 0; i--) if (arr[i] !== 0) return i;
    return -1;
}}

const salesLastIdx = Math.max(findLastNonZero(D['入社売上合計']), currentIdx >= 0 ? currentIdx + 6 : 30);
const payLastIdx = Math.max(findLastNonZero(D['入金合計_確定_税込']), findLastNonZero(D['入金合計_確定ヨミ_税込']), currentIdx >= 0 ? currentIdx + 6 : 30);
const displayEnd = Math.min(Math.max(salesLastIdx, payLastIdx) + 1, D.months.length);
const displayStart = Math.max(0, displayEnd - 24); // Show last 24 months max

const months = D.months.slice(displayStart, displayEnd);
const slice = arr => (arr || []).slice(displayStart, displayEnd);

// KPI Cards
const salesTotal = slice(D['入社売上合計']);
const payTotal = slice(D['入金合計_確定_税込']);
const payTotalYomi = slice(D['入金合計_確定ヨミ_税込']);

const curSales = currentIdx >= displayStart ? salesTotal[currentIdx - displayStart] || 0 : 0;
const prevSales = prevIdx >= displayStart ? salesTotal[prevIdx - displayStart] || 0 : 0;
const curPay = currentIdx >= displayStart ? payTotal[currentIdx - displayStart] || 0 : 0;
const prevPay = prevIdx >= displayStart ? payTotal[prevIdx - displayStart] || 0 : 0;
const curPayYomi = currentIdx >= displayStart ? payTotalYomi[currentIdx - displayStart] || 0 : 0;

// Cumulative for current FY (April start)
const fyStart = D.months.findIndex(m => {{
    const y = parseInt(m.split('/')[0]);
    const mo = parseInt(m.split('/')[1]);
    return (mo === 4 && y === (now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1));
}});
let fySalesSum = 0, fyPaySum = 0;
if (fyStart >= 0 && currentIdx >= fyStart) {{
    for (let i = fyStart; i <= currentIdx; i++) {{
        fySalesSum += (D['入社売上合計'][i] || 0);
        fyPaySum += (D['入金合計_確定_税込'][i] || 0);
    }}
}}

document.getElementById('kpiGrid').innerHTML = `
  <div class="kpi-card"><div class="label">今月 入社売上合計</div><div class="value blue">${{fmt(curSales)}}</div></div>
  <div class="kpi-card"><div class="label">前月 入社売上合計</div><div class="value blue">${{fmt(prevSales)}}</div></div>
  <div class="kpi-card"><div class="label">今月 入金合計(確定)</div><div class="value green">${{fmt(curPay)}}</div></div>
  <div class="kpi-card"><div class="label">今月 入金合計(確定+ヨミ)</div><div class="value purple">${{fmt(curPayYomi)}}</div></div>
  <div class="kpi-card"><div class="label">今期累計 入社売上</div><div class="value orange">${{fmt(fySalesSum)}}</div></div>
  <div class="kpi-card"><div class="label">今期累計 入金(確定)</div><div class="value green">${{fmt(fyPaySum)}}</div></div>
`;

// Sales Chart
new Chart(document.getElementById('salesChart'), {{
  type: 'bar',
  data: {{
    labels: months,
    datasets: [
      {{
        label: '入社売上合計(確定+ヨミ)',
        data: salesTotal,
        backgroundColor: 'rgba(56,189,248,0.6)',
        borderColor: '#38bdf8',
        borderWidth: 1,
        borderRadius: 4,
      }}
    ]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ labels: {{ color: '#94a3b8' }} }},
      tooltip: {{ callbacks: {{ label: ctx => fmt(ctx.raw) }} }}
    }},
    scales: {{
      x: {{ ticks: {{ color: '#64748b', maxRotation: 45 }}, grid: {{ color: '#1e293b' }} }},
      y: {{ ticks: {{ color: '#64748b', callback: v => fmt(v) }}, grid: {{ color: '#334155' }} }}
    }}
  }}
}});

// Payment Chart
new Chart(document.getElementById('paymentChart'), {{
  type: 'bar',
  data: {{
    labels: months,
    datasets: [
      {{
        label: '入金合計(確定)税込',
        data: payTotal,
        backgroundColor: 'rgba(74,222,128,0.6)',
        borderColor: '#4ade80',
        borderWidth: 1,
        borderRadius: 4,
      }},
      {{
        label: '入金合計(確定+ヨミ)税込',
        data: payTotalYomi,
        backgroundColor: 'rgba(167,139,250,0.3)',
        borderColor: '#a78bfa',
        borderWidth: 1,
        borderRadius: 4,
      }}
    ]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ labels: {{ color: '#94a3b8' }} }},
      tooltip: {{ callbacks: {{ label: ctx => ctx.dataset.label + ': ' + fmt(ctx.raw) }} }}
    }},
    scales: {{
      x: {{ ticks: {{ color: '#64748b', maxRotation: 45 }}, grid: {{ color: '#1e293b' }} }},
      y: {{ ticks: {{ color: '#64748b', callback: v => fmt(v) }}, grid: {{ color: '#334155' }} }}
    }}
  }}
}});

// Detail Tables
function renderTable(containerId, tabsId, sections) {{
  const tabsEl = document.getElementById(tabsId);
  const containerEl = document.getElementById(containerId);
  let activeTab = 0;

  function render(idx) {{
    activeTab = idx;
    const sec = sections[idx];
    tabsEl.innerHTML = sections.map((s, i) =>
      `<button class="tab ${{i === idx ? 'active' : ''}}" onclick="window.__switchTab_${{containerId}}(${{i}})">${{s.name}}</button>`
    ).join('');

    let html = '<table class="detail-table"><thead><tr><th>担当者</th>';
    months.forEach(m => html += `<th>${{m}}</th>`);
    html += '<th>合計</th></tr></thead><tbody>';

    const entries = Object.entries(sec.data);
    entries.forEach(([person, vals]) => {{
      const sliced = slice(vals);
      const total = sliced.reduce((a, b) => a + b, 0);
      html += `<tr><td>${{person}}</td>`;
      sliced.forEach(v => html += `<td>${{fmt(v)}}</td>`);
      html += `<td style="font-weight:bold">${{fmt(total)}}</td></tr>`;
    }});

    // Total row
    if (entries.length > 1) {{
      html += '<tr style="border-top:2px solid #64748b;font-weight:bold"><td>合計</td>';
      for (let i = 0; i < months.length; i++) {{
        const sum = entries.reduce((s, [, vals]) => s + (slice(vals)[i] || 0), 0);
        html += `<td>${{fmt(sum)}}</td>`;
      }}
      const grandTotal = entries.reduce((s, [, vals]) => s + slice(vals).reduce((a, b) => a + b, 0), 0);
      html += `<td>${{fmt(grandTotal)}}</td></tr>`;
    }}

    html += '</tbody></table>';
    containerEl.innerHTML = html;
  }}

  window[`__switchTab_${{containerId}}`] = render;
  render(0);
}}

renderTable('salesDetail', 'salesTabs', [
  {{ name: '確定', data: D['入社売上_確定'] }},
  {{ name: 'ヨミ', data: D['入社売上_ヨミ'] }},
]);

renderTable('paymentDetail', 'paymentTabs', [
  {{ name: '確定(税込)', data: D['入金_確定_税込'] }},
  {{ name: 'ヨミ(税込)', data: D['入金_ヨミ_税込'] }},
  {{ name: 'その他', data: D['その他_入金'] }},
]);
</script>
</body>
</html>"""


class DashboardHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/" or self.path == "/index.html":
            try:
                token = refresh_token()
                rows = fetch_sheet_data(token)
                data = build_dashboard_data(rows)
                html = generate_html(data)
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(html.encode("utf-8"))
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "text/plain; charset=utf-8")
                self.end_headers()
                self.wfile.write(f"エラー: {e}".encode("utf-8"))
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        print(f"[{self.log_date_time_string()}] {format % args}")


if __name__ == "__main__":
    server = HTTPServer(("localhost", PORT), DashboardHandler)
    print(f"ダッシュボードを起動しました: http://localhost:{PORT}")
    print("終了するには Ctrl+C を押してください")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nサーバーを停止しました")
        server.server_close()
