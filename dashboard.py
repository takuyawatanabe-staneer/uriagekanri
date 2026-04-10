#!/usr/bin/env python3
"""売上管理ダッシュボード - ローカル開発用サーバー
Netlify Functionsのロジックをローカルで再現"""

import json
import os
import urllib.request
import urllib.parse
from http.server import HTTPServer, SimpleHTTPRequestHandler

SHEET_ID = "17FfNTDve-vzL3n_byDV-RlncGvpQyrgdyzA5y6R29f0"
TOKEN_PATH = os.path.join(os.path.dirname(__file__), "..", "system", "token.json")
PUBLIC_DIR = os.path.join(os.path.dirname(__file__), "public")
PORT = 8080


def refresh_token():
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


def fetch_sheet(token, sheet_name, range_str):
    encoded = urllib.parse.quote(sheet_name)
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{encoded}!{range_str}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    resp = urllib.request.urlopen(req)
    return json.loads(resp.read()).get("values", [])


def parse_yen(val):
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


def get_sales_data(token):
    rows = fetch_sheet(token, "売上計画", "A1:AX50")
    if len(rows) < 2:
        return {}

    months = rows[1][2:]
    start_idx = months.index("2025/08")
    end_idx = months.index("2026/08")
    target_months = months[start_idx:end_idx + 1]

    target_rows = [
        {"row": 3, "label": "入社売上(確定)", "person": "渡邉", "editable": True},
        {"row": 4, "label": "入社売上(確定)", "person": "片桐", "editable": True},
        {"row": 5, "label": "採用支援 ※税込", "person": "ダイナトレック", "editable": True},
        {"row": 6, "label": "CA支援 ※税込", "person": "Ms. Engineer", "editable": True},
        {"row": 11, "label": "入社売上合計(確定+ヨミ)", "person": "", "editable": False},
        {"row": 13, "label": "成約金額(確定)", "person": "渡邉", "editable": True},
        {"row": 14, "label": "成約金額(確定)", "person": "片桐", "editable": True},
        {"row": 20, "label": "入金(確定)※税抜", "person": "渡邉", "editable": True},
        {"row": 21, "label": "入金(確定)※税抜", "person": "片桐", "editable": True},
        {"row": 23, "label": "入金(確定)※税込", "person": "渡邉", "editable": True},
        {"row": 24, "label": "入金(確定)※税込", "person": "片桐", "editable": True},
        {"row": 25, "label": "サーカスRA", "person": "師玉", "editable": True},
        {"row": 26, "label": "採用支援 ※税込", "person": "ダイナ", "editable": True},
        {"row": 27, "label": "CA支援 ※税込", "person": "Ms. Eng", "editable": True},
        {"row": 28, "label": "入金合計(確定)※税込", "person": "", "editable": False},
        {"row": 29, "label": "固定費", "person": "", "editable": True},
        {"row": 30, "label": "残利益", "person": "", "editable": False},
    ]

    sales_data = []
    for t in target_rows:
        row_data = rows[t["row"] - 1] if t["row"] - 1 < len(rows) else []
        values = [parse_yen(v) for v in row_data[2:]]
        while len(values) < len(months):
            values.append(0)
        sales_data.append({
            "label": t["label"],
            "person": t["person"],
            "editable": t["editable"],
            "sheetRow": t["row"],
            "values": values[start_idx:end_idx + 1],
        })

    return {"months": target_months, "rows": sales_data, "colOffset": start_idx + 2}


def get_seiyaku_data(token):
    rows = fetch_sheet(token, "成約Data", "A1:U100")
    if len(rows) < 2:
        return {"headers": [], "rows": []}

    headers = rows[1]
    data_rows = []
    for i in range(2, len(rows)):
        r = rows[i]
        if not r or len(r) < 3 or not r[2] or r[2].strip() == "":
            continue
        data_rows.append({"sheetRow": i + 1, "cells": r})

    return {"headers": headers, "rows": data_rows}


def update_cells(token, updates):
    data_list = []
    for u in updates:
        col = u["col"]
        if col < 26:
            col_letter = chr(65 + col)
        else:
            col_letter = "A" + chr(65 + col - 26)
        data_list.append({
            "range": f"{u['sheet']}!{col_letter}{u['row']}",
            "values": [[u["value"]]],
        })

    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values:batchUpdate"
    body = json.dumps({
        "valueInputOption": "USER_ENTERED",
        "data": data_list,
    }).encode()
    req = urllib.request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    })
    resp = urllib.request.urlopen(req)
    return json.loads(resp.read())


class DashboardHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path.startswith("/api/sheets"):
            self._handle_api_get()
        elif self.path == "/" or self.path == "/index.html":
            self._serve_file(os.path.join(PUBLIC_DIR, "index.html"), "text/html")
        else:
            file_path = os.path.join(PUBLIC_DIR, self.path.lstrip("/"))
            if os.path.exists(file_path):
                self._serve_file(file_path)
            else:
                self.send_response(404)
                self.end_headers()

    def do_POST(self):
        if self.path.startswith("/api/sheets"):
            self._handle_api_post()
        else:
            self.send_response(404)
            self.end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors_headers()
        self.end_headers()

    def _cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _handle_api_get(self):
        try:
            token = refresh_token()
            parsed = urllib.parse.urlparse(self.path)
            params = urllib.parse.parse_qs(parsed.query)
            action = params.get("action", ["sales"])[0]

            if action == "seiyaku":
                data = get_seiyaku_data(token)
            elif action == "all":
                sales = get_sales_data(token)
                seiyaku = get_seiyaku_data(token)
                data = {"sales": sales, "seiyaku": seiyaku}
            else:
                data = get_sales_data(token)

            self._json_response(200, data)
        except Exception as e:
            self._json_response(500, {"error": str(e)})

    def _handle_api_post(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length))
            token = refresh_token()
            result = update_cells(token, body["updates"])
            self._json_response(200, {"ok": True, "result": result})
        except Exception as e:
            self._json_response(500, {"error": str(e)})

    def _json_response(self, status, data):
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self._cors_headers()
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))

    def _serve_file(self, path, content_type="text/html"):
        with open(path, "rb") as f:
            content = f.read()
        self.send_response(200)
        self.send_header("Content-Type", f"{content_type}; charset=utf-8")
        self.end_headers()
        self.wfile.write(content)

    def log_message(self, format, *args):
        print(f"[{self.log_date_time_string()}] {format % args}")


if __name__ == "__main__":
    server = HTTPServer(("localhost", PORT), DashboardHandler)
    print(f"ダッシュボードを起動: http://localhost:{PORT}")
    print("Ctrl+C で停止")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n停止しました")
        server.server_close()
