"""
stats.py — 裁罰趨勢統計報告產生器

讀取 penalty_history.json，產生季度/年度統計 HTML 報告。
可手動執行，或設定為每季自動產生。

用法：
    python stats.py          # 產生當年度統計
    python stats.py 2025     # 產生指定年度統計
"""

import json
import sys
from pathlib import Path
from datetime import datetime
from collections import Counter, defaultdict

# 終端機強制使用 UTF-8 輸出 (針對 Windows)
if sys.platform == "win32":
    if sys.stdout and sys.stdout.encoding != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if sys.stderr and sys.stderr.encoding != "utf-8":
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

SCRIPT_DIR = Path(__file__).parent
HISTORY_FILE = SCRIPT_DIR / "penalty_history.json"
REPORTS_DIR = SCRIPT_DIR / "reports"


def load_history() -> dict:
    if not HISTORY_FILE.exists():
        print("[ERROR] 找不到 penalty_history.json，請先執行 scraper.py 至少一次。")
        sys.exit(1)
    return json.loads(HISTORY_FILE.read_text(encoding="utf-8"))


def generate_stats(history: dict, year: int) -> dict:
    """從歷史資料中統計指定年度的裁罰趨勢"""
    quarterly = defaultdict(int)      # Q1~Q4 案件數
    entity_counts = Counter()         # 各機構被罰次數
    monthly = defaultdict(int)        # 每月案件數
    repeat_offenders = []             # 累犯（>=2 次）

    for entity, records in history.items():
        year_records = [r for r in records if r.get("date", "").startswith(str(year))]
        if not year_records:
            continue

        entity_counts[entity] += len(year_records)

        for r in year_records:
            date_str = r.get("date", "")
            if len(date_str) >= 7:
                month = int(date_str[5:7])
                monthly[month] += 1
                quarter = (month - 1) // 3 + 1
                quarterly[quarter] += 1

    # 累犯清單
    for entity, count in entity_counts.most_common():
        if count >= 2:
            repeat_offenders.append({"entity": entity, "count": count})

    total = sum(entity_counts.values())

    return {
        "year": year,
        "total": total,
        "quarterly": dict(quarterly),
        "monthly": dict(monthly),
        "top_entities": entity_counts.most_common(10),
        "repeat_offenders": repeat_offenders,
        "unique_entities": len(entity_counts),
    }


def render_bar(value: int, max_value: int, color: str = "#c62828") -> str:
    """產生簡易 CSS 長條"""
    if max_value == 0:
        return ""
    width = int(value / max_value * 200)
    return f'<span style="display:inline-block;width:{width}px;height:16px;background:{color};border-radius:3px;margin-right:8px;vertical-align:middle;"></span>'


def generate_html(stats: dict) -> str:
    year = stats["year"]
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 季度統計
    q_max = max(stats["quarterly"].values()) if stats["quarterly"] else 1
    quarterly_rows = ""
    for q in range(1, 5):
        count = stats["quarterly"].get(q, 0)
        bar = render_bar(count, q_max)
        quarterly_rows += f"<tr><td>Q{q}</td><td>{bar}<strong>{count}</strong> 件</td></tr>\n"

    # 月度統計
    m_max = max(stats["monthly"].values()) if stats["monthly"] else 1
    monthly_rows = ""
    for m in range(1, 13):
        count = stats["monthly"].get(m, 0)
        bar = render_bar(count, m_max, "#1565c0")
        monthly_rows += f"<tr><td>{m} 月</td><td>{bar}<strong>{count}</strong></td></tr>\n"

    # Top 10 機構
    top_rows = ""
    if stats["top_entities"]:
        t_max = stats["top_entities"][0][1]
        for idx, (entity, count) in enumerate(stats["top_entities"], 1):
            bar = render_bar(count, t_max, "#ff6f00")
            top_rows += f"<tr><td>{idx}</td><td>{entity}</td><td>{bar}<strong>{count}</strong></td></tr>\n"

    # 累犯清單
    repeat_html = ""
    if stats["repeat_offenders"]:
        repeat_items = ""
        for r in stats["repeat_offenders"]:
            repeat_items += f"<li><strong>{r['entity']}</strong> — 共 {r['count']} 次</li>\n"
        repeat_html = f"""
        <div style="background:#fff3e0;border-left:5px solid #ff6f00;padding:15px;margin:20px 0;border-radius:4px;">
            <h3 style="margin-top:0;color:#e65100;">累犯機構（年度內 2 次以上）</h3>
            <ul>{repeat_items}</ul>
        </div>"""

    return f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<title>裁罰趨勢統計 — {year} 年度</title>
<style>
  body {{ font-family: "Microsoft JhengHei", Arial, sans-serif; max-width: 900px; margin: 0 auto; padding: 20px; color: #333; }}
  h1 {{ color: #c62828; border-bottom: 2px solid #c62828; padding-bottom: 10px; }}
  h2 {{ color: #1565c0; margin-top: 30px; }}
  table {{ border-collapse: collapse; width: 100%; margin: 10px 0; }}
  td {{ padding: 6px 12px; border-bottom: 1px solid #eee; }}
  .summary-box {{ display: flex; gap: 20px; margin: 20px 0; }}
  .summary-card {{ flex: 1; text-align: center; padding: 20px; border-radius: 8px; }}
  .disclaimer {{ color: #999; font-size: 0.85em; margin-top: 40px; border-top: 1px solid #ddd; padding-top: 10px; }}
</style>
</head>
<body>
<h1>裁罰趨勢統計 — {year} 年度</h1>
<p>報告產生時間：{now}</p>

<div class="summary-box">
    <div class="summary-card" style="background:#fce4ec;">
        <div style="font-size:2.5em;font-weight:bold;color:#c62828;">{stats['total']}</div>
        <div>年度裁罰總件數</div>
    </div>
    <div class="summary-card" style="background:#e3f2fd;">
        <div style="font-size:2.5em;font-weight:bold;color:#1565c0;">{stats['unique_entities']}</div>
        <div>受罰機構數</div>
    </div>
    <div class="summary-card" style="background:#fff3e0;">
        <div style="font-size:2.5em;font-weight:bold;color:#e65100;">{len(stats['repeat_offenders'])}</div>
        <div>累犯機構數</div>
    </div>
</div>

<h2>季度分布</h2>
<table>{quarterly_rows}</table>

<h2>月度分布</h2>
<table>{monthly_rows}</table>

{repeat_html}

<h2>受罰次數排名（前 10）</h2>
<table>
<tr style="font-weight:bold;border-bottom:2px solid #333;"><td>#</td><td>機構名稱</td><td>裁罰次數</td></tr>
{top_rows}
</table>

<div class="disclaimer">本報告由系統自動產生，資料來源為金管會裁罰案件 RSS。統計數據僅供內部參考。</div>
</body>
</html>"""


def main():
    # 取得目標年度
    if len(sys.argv) > 1:
        try:
            year = int(sys.argv[1])
        except ValueError:
            print(f"[ERROR] 無效的年度參數：{sys.argv[1]}")
            sys.exit(1)
    else:
        year = datetime.now().year

    print(f"正在產生 {year} 年度裁罰趨勢統計...")

    history = load_history()
    stats = generate_stats(history, year)

    if stats["total"] == 0:
        print(f"  {year} 年度無任何裁罰紀錄。")
        sys.exit(0)

    REPORTS_DIR.mkdir(exist_ok=True)
    output_file = REPORTS_DIR / f"stats-{year}.html"
    html = generate_html(stats)
    output_file.write_text(html, encoding="utf-8")

    print(f"  統計報告已產生：{output_file}")
    print(f"  年度裁罰總數：{stats['total']} 件")
    print(f"  受罰機構數：{stats['unique_entities']}")
    print(f"  累犯機構數：{len(stats['repeat_offenders'])}")


if __name__ == "__main__":
    main()
