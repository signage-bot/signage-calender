"""
Forms_予定表.xlsx → data/events.json 変換スクリプト（GitHub Actionsで定時実行）

読み込むもの
  ・tbl_予定（予定シート）… 日付・開始・終了・現場・区分・業者名・内容・備考・状態
  ・tbl_現場（設定シート）… 現場名ごとの文字色・塗り色・表示順
  ・tbl_設定（設定シート）… 表示開始・表示終了
状態が「有効」の行だけ出力します。
"""
import io, json, os, sys
from datetime import datetime, date, time, timedelta, timezone
import requests
from openpyxl import load_workbook

EXCEL_URL = os.environ.get("EXCEL_URL", "")
LOCAL_FILE = os.environ.get("LOCAL_FILE", "")
OUT = "data/events.json"


def open_book():
    if LOCAL_FILE:
        return load_workbook(LOCAL_FILE, data_only=True)
    if not EXCEL_URL:
        sys.exit("EXCEL_URL が未設定です")
    url = EXCEL_URL + ("&" if "?" in EXCEL_URL else "?") + "download=1"
    r = requests.get(url, timeout=60, allow_redirects=True)
    r.raise_for_status()
    if not r.content.startswith(b"PK"):
        sys.exit("Excelを取得できませんでした（ログインを求められている可能性があります）")
    return load_workbook(io.BytesIO(r.content), data_only=True)


def read_table(wb, name):
    """テーブル名で読み込み、[{列名: 値}, ...] を返す（シートの位置が変わっても動く）"""
    for ws in wb.worksheets:
        if name in ws.tables:
            rows = list(ws[ws.tables[name].ref])
            head = [str(c.value).strip() if c.value is not None else "" for c in rows[0]]
            return [dict(zip(head, [c.value for c in r])) for r in rows[1:]]
    sys.exit(f"テーブル {name} が見つかりません")


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str) and v.strip():
        return datetime.strptime(v.strip().replace("-", "/"), "%Y/%m/%d").date()
    return None


def to_time(v):
    if isinstance(v, datetime):
        return v.time()
    if isinstance(v, time):
        return v
    if isinstance(v, timedelta):
        return (datetime.min + v).time()
    if isinstance(v, (int, float)):            # Excelのシリアル値（0.375 = 9:00）
        return (datetime.min + timedelta(days=float(v))).time().replace(microsecond=0)
    if isinstance(v, str) and v.strip():
        return datetime.strptime(v.strip(), "%H:%M").time()
    return None


def s(v):
    return "" if v is None else str(v).strip()


def main():
    wb = open_book()

    # 現場ごとの色
    sites = {}
    for r in read_table(wb, "tbl_現場"):
        name = s(r.get("現場名"))
        if name:
            sites[name] = {"text": s(r.get("文字色")) or "#333333",
                           "bg": s(r.get("塗り色")) or "#EEEEEE",
                           "order": r.get("表示順") or 99}

    # 表示する時間帯
    conf = {s(r.get("キー")): r.get("値") for r in read_table(wb, "tbl_設定")}
    t_from = to_time(conf.get("表示開始")) or time(8, 0)
    t_to = to_time(conf.get("表示終了")) or time(17, 0)

    events = []
    for r in read_table(wb, "tbl_予定"):
        if s(r.get("状態")) != "有効":
            continue
        d = to_date(r.get("日付"))
        if not d:
            continue
        st = to_time(r.get("開始")) or time(8, 0)
        en = to_time(r.get("終了")) or (datetime.combine(d, st) + timedelta(hours=1)).time()
        site, company = s(r.get("現場")), s(r.get("業者名"))
        color = sites.get(site, {"text": "#333333", "bg": "#EEEEEE"})
        events.append({
            "start": datetime.combine(d, st).strftime("%Y-%m-%dT%H:%M"),
            "end": datetime.combine(d, en).strftime("%Y-%m-%dT%H:%M"),
            "site": site,
            "kubun": s(r.get("区分")),
            "company": company,
            "detail": s(r.get("内容")),
            "note": s(r.get("備考")),
            "no": s(r.get("管理番号")),
            "textColor": color["text"],
            "bgColor": color["bg"],
        })

    jst = timezone(timedelta(hours=9))
    out = {
        "updated": datetime.now(jst).strftime("%Y/%m/%d %H:%M"),
        "slotMin": t_from.strftime("%H:%M:00"),
        "slotMax": t_to.strftime("%H:%M:00"),
        "sites": [{"name": k, **v} for k, v in sorted(sites.items(), key=lambda x: x[1]["order"])],
        "events": events,
    }
    os.makedirs("data", exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"{len(events)} 件を書き出しました")


if __name__ == "__main__":
    main()
