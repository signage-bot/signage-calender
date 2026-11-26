"""
共有Excel → data/events.json 変換スクリプト
GitHub Actions から定時実行される。

Excelの想定（1枚目のシート、1行目が見出し）:
  開始日 | 終了日 | 現場名 | 業者名 | 作業内容
  ※終了日が空欄なら開始日と同じ日とみなす
"""
import io, json, os, sys
from datetime import datetime, timedelta, timezone
import pandas as pd
import requests

EXCEL_URL = os.environ.get("EXCEL_URL", "")       # GitHubのSecretsに登録
LOCAL_FILE = os.environ.get("LOCAL_FILE", "")     # ローカルテスト用
OUT = "data/events.json"

COL_START, COL_END = "開始日", "終了日"
COL_SITE, COL_COMPANY, COL_DETAIL = "現場名", "業者名", "作業内容"

def read_excel():
    if LOCAL_FILE:
        return pd.read_excel(LOCAL_FILE)
    if not EXCEL_URL:
        sys.exit("EXCEL_URL が未設定です")
    url = EXCEL_URL + ("&" if "?" in EXCEL_URL else "?") + "download=1"
    r = requests.get(url, timeout=60, allow_redirects=True)
    r.raise_for_status()
    if not r.content.startswith(b"PK"):   # xlsxはZIP形式。ログイン画面(HTML)が返ったら止める
        sys.exit("Excelを取得できませんでした（ログインを求められている可能性があります）")
    return pd.read_excel(io.BytesIO(r.content))

def to_date(v):
    if pd.isna(v) or v == "":
        return None
    return pd.to_datetime(v).date()

def main():
    df = read_excel()
    events = []
    for _, row in df.iterrows():
        start = to_date(row.get(COL_START))
        if not start:
            continue
        end = to_date(row.get(COL_END)) or start
        site = "" if pd.isna(row.get(COL_SITE)) else str(row.get(COL_SITE))
        company = "" if pd.isna(row.get(COL_COMPANY)) else str(row.get(COL_COMPANY))
        detail = "" if pd.isna(row.get(COL_DETAIL)) else str(row.get(COL_DETAIL))
        events.append({
            "title": " ".join(x for x in [site, company] if x) or detail,
            "start": start.isoformat(),
            "end": (end + timedelta(days=1)).isoformat(),  # カレンダー表示用（終日は翌日指定）
            "end_display": end.isoformat(),
            "site": site, "company": company, "detail": detail,
        })
    jst = timezone(timedelta(hours=9))
    out = {"updated": datetime.now(jst).strftime("%Y/%m/%d %H:%M"), "events": events}
    os.makedirs("data", exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"{len(events)} 件を書き出しました")

if __name__ == "__main__":
    main()
