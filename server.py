# 房屋收支帳本機伺服器：提供網站並把新增的收支寫回 HOUSE.xlsx
# 用法：python server.py          （預設 http://127.0.0.1:8765 並自動開啟瀏覽器）
#       python server.py 9000 --no-browser
import datetime
import json
import shutil
import sys
import urllib.request
import webbrowser
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "lib"))  # 專案內附的 openpyxl
try:
    import openpyxl
except ImportError as e:
    print(f"無法載入 openpyxl（{e}）")
    print(f"使用的 Python：{sys.executable}")
    print("請確認資料夾內的 lib 子資料夾仍存在，或執行：python -m pip install openpyxl")
    input("按 Enter 結束…")
    sys.exit(1)

from update_data import HERE, XLSX, read_rows, write_data_js

BACKUP_DIR = HERE / "backups"


def clean_entry(data):
    """檢查並整理一筆輸入，回傳 (kind, amount, desc, date, memo)；不合法時拋出 ValueError。"""
    kind = data.get("kind")
    if kind not in ("in", "out", "usd"):
        raise ValueError("類型錯誤")
    try:
        amount = float(data.get("amount"))
    except (TypeError, ValueError):
        raise ValueError("金額格式錯誤")
    if amount == 0 or amount != amount or abs(amount) > 1e9 or (kind != "usd" and amount < 0):
        raise ValueError("金額需大於 0")
    desc = str(data.get("desc") or "").strip()[:200]
    if not desc:
        raise ValueError("請填寫說明")
    date = str(data.get("date") or "").strip() or None
    if date:
        datetime.datetime.strptime(date, "%Y-%m-%d")  # 格式錯誤會拋出 ValueError
    memo = str(data.get("memo") or "").strip()[:200] or None
    return kind, amount, desc, date, memo


def add_entry(kind, amount, desc, date=None, memo=None):
    """在「結餘」列上方新增一筆，回傳寫入的列號。"""
    wb = openpyxl.load_workbook(XLSX)
    ws = wb.worksheets[0]

    total = next((r for r in range(2, ws.max_row + 1) if ws.cell(r, 3).value == "結餘"), None)
    if total is None:
        raise ValueError("找不到「結餘」列，無法判斷新增位置")

    last = 1
    for r in range(2, total):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, 9)):
            last = r
    target = last + 1
    if target >= total:  # 沒有空列，插入一列
        ws.insert_rows(total)
        target, total = total, total + 1

    if date:
        cell = ws.cell(target, 3, datetime.datetime.strptime(date, "%Y-%m-%d"))
        cell.number_format = "mm-dd-yy"
    col = {"usd": 4, "in": 5, "out": 6}[kind]
    ws.cell(target, col, amount)
    ws.cell(target, 7, desc)
    if memo:
        ws.cell(target, 8, memo)

    # 讓結餘公式涵蓋到最後一筆資料
    end = total - 1
    for c in "DEF":
        ws[f"{c}{total}"] = f"=SUM({c}2:{c}{end})"
    ws[f"G{total}"] = f"=E{total}-F{total}"

    with open(XLSX, "r+b"):  # 檔案被 Excel 鎖定時會在此拋出 PermissionError，避免留下多餘備份
        pass
    BACKUP_DIR.mkdir(exist_ok=True)
    shutil.copy2(XLSX, BACKUP_DIR / f"HOUSE-{datetime.datetime.now():%Y%m%d-%H%M%S-%f}.xlsx")
    wb.save(XLSX)
    write_data_js()
    return target


class Handler(SimpleHTTPRequestHandler):
    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path.split("?")[0] == "/api/rows":
            try:
                rows, source = read_rows()
                return self._json(200, {"rows": rows, "source": source})
            except Exception as e:
                return self._json(500, {"error": str(e)})
        return super().do_GET()

    def do_POST(self):
        if self.path != "/api/add":
            return self._json(404, {"error": "not found"})
        try:
            data = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or b"{}")
            row = add_entry(*clean_entry(data))
        except PermissionError:
            return self._json(409, {"error": "HOUSE.xlsx 正在被其他程式（例如 Excel）開啟，請關閉後再試。"})
        except (ValueError, TypeError) as e:
            return self._json(400, {"error": str(e)})
        except Exception as e:
            return self._json(500, {"error": f"寫入失敗：{e}"})
        rows, source = read_rows()
        return self._json(200, {"row": row, "rows": rows, "source": source})

    def log_message(self, fmt, *args):
        line = fmt % args
        if "/api/" in line or "code 5" in line:
            super().log_message("%s", line)

    def log_error(self, fmt, *args):
        pass  # 例如瀏覽器要求 favicon.ico 的 404，不需要顯示


def already_running(port):
    """該埠是否已有本程式在執行。"""
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/rows", timeout=2) as r:
            return r.status == 200
    except Exception:
        return False


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    port = int(args[0]) if args else 8765
    open_browser = "--no-browser" not in sys.argv
    if already_running(port):
        url = f"http://127.0.0.1:{port}/"
        print(f"房屋收支帳已在執行：{url}", flush=True)
        if open_browser:
            webbrowser.open(url)
        sys.exit(0)
    try:
        server = ThreadingHTTPServer(("127.0.0.1", port), partial(Handler, directory=str(HERE)))
    except OSError:
        server = ThreadingHTTPServer(("127.0.0.1", 0), partial(Handler, directory=str(HERE)))
        port = server.server_address[1]
    url = f"http://127.0.0.1:{port}/"
    print(f"房屋收支帳：{url}", flush=True)
    print("新增收支時請保持此視窗開啟；關閉此視窗或按 Ctrl+C 即結束。", flush=True)
    if open_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
