# 從 HOUSE.xlsx 重新產生 data.js（修改試算表後執行：python update_data.py）
import datetime
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "lib"))  # 使用專案內附的 openpyxl，不需另外安裝

import openpyxl  # noqa: E402

XLSX = HERE / "HOUSE.xlsx"


def read_rows():
    """回傳 (rows, source)。rows 每列為 [列號, A..H]，讀到「結餘」列為止。"""
    ws = openpyxl.load_workbook(XLSX).worksheets[0]
    rows = []
    for r in ws.iter_rows(min_row=2, max_col=8):
        v = [c.value for c in r]
        if v[2] == "結餘":
            break
        if all(x is None for x in v):
            continue
        v = [x.strftime("%Y-%m-%d") if isinstance(x, datetime.datetime) else x for x in v]
        rows.append([r[0].row] + v)
    stamp = datetime.datetime.fromtimestamp(XLSX.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
    return rows, f"HOUSE.xlsx（{stamp}）"


def write_data_js():
    rows, source = read_rows()
    (HERE / "data.js").write_text(
        "// 由 update_data.py 從 HOUSE.xlsx 產生，請勿手動編輯\n"
        f"window.HOUSE_SOURCE = {json.dumps(source, ensure_ascii=False)};\n"
        f"window.HOUSE_ROWS = {json.dumps(rows, ensure_ascii=False)};\n",
        encoding="utf-8",
    )
    return len(rows)


if __name__ == "__main__":
    print(f"已更新 data.js：{write_data_js()} 列")
