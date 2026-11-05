# 由 GitHub Actions 呼叫：從環境變數讀取一筆收支並寫入 HOUSE.xlsx
# 環境變數：KIND (in/out/usd)、AMOUNT、DESC、DATE (YYYY-MM-DD，可空)、MEMO（可空）
import os
import sys

from server import add_entry, clean_entry

try:
    entry = clean_entry({k.lower(): os.environ.get(k) for k in ("KIND", "AMOUNT", "DESC", "DATE", "MEMO")})
except ValueError as e:
    print(f"::error::輸入不合法：{e}")
    sys.exit(1)

row = add_entry(*entry)
print(f"已寫入 HOUSE.xlsx 第 {row} 列：{entry}")
gh_out = os.environ.get("GITHUB_OUTPUT")
if gh_out:
    with open(gh_out, "a", encoding="utf-8") as f:
        f.write(f"row={row}\n")
