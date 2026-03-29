#!/usr/bin/env python3
"""
業績分析報告 LINE 推送腳本
讀取最新 CSV 資料，產生摘要報告，透過 LINE Messaging API 推送
"""

import calendar
import csv
import glob
import os
import json
import urllib.request
import urllib.error
from datetime import datetime
from pathlib import Path

# ─── 設定 ───
# 請填入你的 LINE Bot Channel Access Token 和你自己的 LINE User ID
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "line_config.json")

def load_config():
    if not os.path.exists(CONFIG_FILE):
        print("❌ 找不到 line_config.json，請先建立設定檔。")
        print("   執行: python3 notify.py --setup")
        return None
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)

def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
    print(f"✅ 設定已儲存到 {CONFIG_FILE}")

def setup():
    print("=== LINE Bot 設定 ===")
    print()
    token = input("請貼上 Channel Access Token: ").strip()
    user_id = input("請貼上你的 LINE User ID: ").strip()
    config = {
        "channel_access_token": token,
        "user_id": user_id
    }
    save_config(config)
    print()
    print("測試傳送中...")
    send_line_message(config, "✅ LINE Bot 設定成功！業績報告將會透過這裡推送。")
    print("請檢查 LINE 是否有收到測試訊息。")

def send_line_message(config, message):
    """透過 LINE Messaging API 推送訊息"""
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {config['channel_access_token']}"
    }
    # LINE 訊息長度限制 5000 字，若超過則分段
    chunks = split_message(message, 4900)
    for chunk in chunks:
        body = json.dumps({
            "to": config["user_id"],
            "messages": [{"type": "text", "text": chunk}]
        }).encode("utf-8")
        req = urllib.request.Request(url, data=body, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req) as resp:
                if resp.status == 200:
                    pass  # 成功
        except urllib.error.HTTPError as e:
            error_body = e.read().decode("utf-8")
            print(f"❌ LINE API 錯誤 ({e.code}): {error_body}")
            return False
    print("✅ LINE 訊息已發送")
    return True

def split_message(text, max_len):
    if len(text) <= max_len:
        return [text]
    chunks = []
    lines = text.split("\n")
    current = ""
    for line in lines:
        if len(current) + len(line) + 1 > max_len:
            chunks.append(current)
            current = line
        else:
            current = current + "\n" + line if current else line
    if current:
        chunks.append(current)
    return chunks

def parse_number(s):
    """解析可能含逗號的數字字串"""
    if not s:
        return 0
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return 0

def find_latest_csv():
    """找到最新的 CSV 檔案"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    csvs = glob.glob(os.path.join(base_dir, "*.csv"))
    if not csvs:
        return None
    return max(csvs, key=os.path.getmtime)

def read_csv(filepath):
    """讀取 CSV 並回傳結構化資料"""
    rows = []
    with open(filepath, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows

def generate_report(rows, csv_filename):
    """產生業績摘要報告"""
    report_date = os.path.basename(csv_filename).replace(".csv", "")

    # 計算預估月底業績的比率
    try:
        dt = datetime.strptime(report_date, "%Y-%m-%d")
        days_in_month = calendar.monthrange(dt.year, dt.month)[1]
        day_of_month = dt.day
        projection_ratio = days_in_month / day_of_month if day_of_month > 0 else 1
    except ValueError:
        projection_ratio = 1

    # 按 app 分組統計
    apps = {}
    for row in rows:
        app = row.get("app_id", "unknown")
        if app not in apps:
            apps[app] = {"revenue": 0, "prev_revenue": 0, "people": 0, "active": 0, "refund": 0}
        revenue = parse_number(row.get("本期業績", "0"))
        prev_revenue = parse_number(row.get("上期業績", "0"))
        refund = parse_number(row.get("本期退費金額", "0"))
        apps[app]["revenue"] += revenue
        apps[app]["prev_revenue"] += prev_revenue
        apps[app]["refund"] += refund
        apps[app]["people"] += 1
        if revenue > 0:
            apps[app]["active"] += 1

    # 按組別統計
    groups = {}
    for row in rows:
        group = row.get("組別", "").strip()
        if not group:
            continue
        revenue = parse_number(row.get("本期業績", "0"))
        prev_revenue = parse_number(row.get("上期業績", "0"))
        if group not in groups:
            groups[group] = {"revenue": 0, "prev_revenue": 0, "members": []}
        groups[group]["revenue"] += revenue
        groups[group]["prev_revenue"] += prev_revenue
        name = row.get("姓名", "").strip()
        if revenue > 0:
            groups[group]["members"].append((name, revenue))

    # 個人 Top 10
    individuals = []
    for row in rows:
        name = row.get("姓名", "").strip()
        revenue = parse_number(row.get("本期業績", "0"))
        app = row.get("app_id", "")
        group = row.get("組別", "")
        if revenue > 0:
            individuals.append((name, app, group, revenue))
    individuals.sort(key=lambda x: x[3], reverse=True)

    # 全公司合計
    total_revenue = sum(a["revenue"] for a in apps.values())
    total_prev = sum(a["prev_revenue"] for a in apps.values())
    total_refund = sum(a["refund"] for a in apps.values())
    total_people = sum(a["people"] for a in apps.values())
    total_active = sum(a["active"] for a in apps.values())

    growth = ((total_revenue - total_prev) / total_prev * 100) if total_prev > 0 else 0

    lines = []
    lines.append(f"📊 業績分析報告 — {report_date}")
    lines.append("=" * 30)
    lines.append("")

    # 總覽
    projected_revenue = total_revenue * projection_ratio
    lines.append("【全公司總覽】")
    lines.append(f"  總業績：${total_revenue:,.0f}")
    lines.append(f"  預估月底：${projected_revenue:,.0f}")
    lines.append(f"  上期：${total_prev:,.0f} ({growth:+.1f}%)")
    lines.append(f"  退費：${total_refund:,.0f}")
    lines.append(f"  淨業績：${total_revenue - total_refund:,.0f}")
    lines.append(f"  人數：{total_people} 人（有業績 {total_active} 人）")
    lines.append("")

    # 各 App 業績
    lines.append("【各品牌業績】")
    for app, data in sorted(apps.items(), key=lambda x: x[1]["revenue"], reverse=True):
        g = ((data["revenue"] - data["prev_revenue"]) / data["prev_revenue"] * 100) if data["prev_revenue"] > 0 else 0
        g_str = f" ({g:+.1f}%)" if data["prev_revenue"] > 0 else ""
        proj = data["revenue"] * projection_ratio
        lines.append(f"  {app}: ${data['revenue']:,.0f}{g_str}")
        lines.append(f"    預估月底 ${proj:,.0f} | 有業績 {data['active']}/{data['people']} 人, 退費 ${data['refund']:,.0f}")
    lines.append("")

    # 組別排行（全部顯示）
    lines.append("【組別排行】")
    sorted_groups = sorted(groups.items(), key=lambda x: x[1]["revenue"], reverse=True)
    for i, (gname, gdata) in enumerate(sorted_groups, 1):
        g = ((gdata["revenue"] - gdata["prev_revenue"]) / gdata["prev_revenue"] * 100) if gdata["prev_revenue"] > 0 else 0
        g_str = f" ({g:+.1f}%)" if gdata["prev_revenue"] > 0 else ""
        proj = gdata["revenue"] * projection_ratio
        lines.append(f"  {i}. {gname}: ${gdata['revenue']:,.0f}{g_str}")
        lines.append(f"     預估月底 ${proj:,.0f}")
    lines.append("")

    # 個人 Top 10
    lines.append("【個人 Top 10】")
    for i, (name, app, group, rev) in enumerate(individuals[:10], 1):
        lines.append(f"  {i}. {name} ({app}): ${rev:,.0f}")
    lines.append("")

    # 需關注人員（業績大幅下降）
    alerts = []
    for row in rows:
        name = row.get("姓名", "").strip()
        rev = parse_number(row.get("本期業績", "0"))
        prev = parse_number(row.get("上期業績", "0"))
        if prev > 100000 and rev < prev * 0.5:
            drop = (1 - rev / prev) * 100
            alerts.append((name, rev, prev, drop))
    if alerts:
        alerts.sort(key=lambda x: x[3], reverse=True)
        lines.append("⚠️ 【需關注 — 業績大幅下降】")
        for name, rev, prev, drop in alerts[:5]:
            lines.append(f"  {name}: ${rev:,.0f} (↓{drop:.0f}%, 上期 ${prev:,.0f})")
        lines.append("")

    # 表現優異（大幅成長）
    stars = []
    for row in rows:
        name = row.get("姓名", "").strip()
        rev = parse_number(row.get("本期業績", "0"))
        prev = parse_number(row.get("上期業績", "0"))
        if rev > 100000 and prev > 0 and rev > prev * 1.5:
            growth_pct = (rev / prev - 1) * 100
            stars.append((name, rev, prev, growth_pct))
    if stars:
        stars.sort(key=lambda x: x[3], reverse=True)
        lines.append("🌟 【表現優異 — 業績大幅成長】")
        for name, rev, prev, g in stars[:5]:
            lines.append(f"  {name}: ${rev:,.0f} (↑{g:.0f}%, 上期 ${prev:,.0f})")
        lines.append("")

    lines.append(f"📅 報告產生時間：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    return "\n".join(lines)


def main():
    import sys

    if "--setup" in sys.argv:
        setup()
        return

    if "--test" in sys.argv:
        # 只產生報告不傳送
        csv_file = find_latest_csv()
        if not csv_file:
            print("❌ 找不到 CSV 檔案")
            return
        rows = read_csv(csv_file)
        report = generate_report(rows, csv_file)
        print(report)
        return

    # 正式執行：產生報告並傳送
    config = load_config()
    if not config:
        return

    csv_file = find_latest_csv()
    if not csv_file:
        print("❌ 找不到 CSV 檔案")
        return

    print(f"📂 讀取: {csv_file}")
    rows = read_csv(csv_file)
    report = generate_report(rows, csv_file)
    print(report)
    print()
    send_line_message(config, report)


if __name__ == "__main__":
    main()
