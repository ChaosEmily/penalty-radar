"""
setup.py — 裁罰案件追蹤器安裝腳本

執行內容：
1. 建立預設的 config.json
2. 檢查必填欄位 (Gmail, API Key)
3. 註冊 Windows 工作排程器（自動執行 scraper.py）
"""

import json
import subprocess
import sys
from pathlib import Path

# 終端機強制使用 UTF-8 輸出 (針對 Windows)
if sys.platform == "win32":
    if sys.stdout.encoding != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

SCRIPT_DIR  = Path(__file__).parent
CONFIG_FILE = SCRIPT_DIR / "config.json"

DEFAULT_CONFIG = {
    "gmail_user": "your_email@gmail.com",
    "gmail_app_password": "your_16_digit_app_password",
    "recipient_email": "risk-management@company.com",
    "email_strategy": "single_emails",
    "scrape_interval_minutes": 60,
    "max_new_per_run": 3,
    "anthropic_api_key": "",
    "company_type": "銀行",
    "rss_url": "https://www.fsc.gov.tw/RSS/Messages?serno=201202290003&language=chinese"
}

def ensure_config() -> dict:
    if not CONFIG_FILE.exists():
        print("[INFO] 找不到 config.json，正在建立預設設定檔...")
        CONFIG_FILE.write_text(json.dumps(DEFAULT_CONFIG, indent=4, ensure_ascii=False), encoding="utf-8")
        print(f"  預設設定檔已建立於：{CONFIG_FILE}")
        print("  請開啟 config.json 填寫您的 Gmail 帳號、應用程式密碼與 Anthropic API 金鑰後，再次執行本指令。")
        sys.exit(0)

    config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))

    missing = []
    if config.get("gmail_user") == "your_email@gmail.com" or not config.get("gmail_user"):
        missing.append("gmail_user")
    if config.get("gmail_app_password") == "your_16_digit_app_password" or not config.get("gmail_app_password"):
        missing.append("gmail_app_password")
    if not config.get("anthropic_api_key"):
        missing.append("anthropic_api_key")

    if missing:
        print("[ERROR] config.json 中有尚未填寫的必填項目：")
        for m in missing:
            print(f"  - {m}")
        print("請填寫後再次執行 setup.py。")
        sys.exit(1)

    return config

def register_scheduler(config: dict) -> None:
    interval  = config.get("scrape_interval_minutes", 60)
    task_name = "PenaltyRadar_Scraper"
    script    = str(SCRIPT_DIR / "scraper.py")
    python    = sys.executable

    if sys.platform == "win32":
        cmd = [
            "schtasks", "/create",
            "/tn", task_name,
            "/tr", f'"{python}" "{script}"',
            "/sc", "minute",
            "/mo", str(interval),
            "/f",
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            print(f"  Windows 工作排程器已註冊（每 {interval} 分鐘自動執行一次）")
            print(f"    任務名稱：{task_name}")
        else:
            print(f"  排程器註冊失敗（您可能需要以系統管理員身分執行 CMD）：")
            print(f"    {result.stderr.strip()}")
            print(f"    如果無法註冊，請手動執行：{python} {script}")
    else:
        print(f"  偵測到非 Windows 系統 ({sys.platform})，請手動設定 Crontab。")
        print(f"  請執行 `crontab -e` 並加入以下行：")
        print(f"  */{interval} * * * * \"{python}\" \"{script}\" >> \"{SCRIPT_DIR}/cron.log\" 2>&1")

def main():
    print("=" * 60)
    print("  裁罰案件追蹤器 (Penalty Radar) — 安裝程式")
    print("=" * 60)

    print("\n[1/2] 檢查設定檔 (config.json)...")
    config = ensure_config()
    print("  設定檔格式與必要欄位驗證通過。")

    print(f"\n[2/2] 設定自動化排程...")
    register_scheduler(config)

    print("\n" + "=" * 60)
    print("  安裝完成")
    print("=" * 60)
    print(f"""
系統已將 config.json 與您的信箱綁定。
Email 發送策略為：{config.get('email_strategy', 'single_emails')}
每次最多處理新案件數：{config.get('max_new_per_run', 3)}

接著可執行：
   python scraper.py      （立刻手動測試抓取與發信機制）
""")

if __name__ == "__main__":
    main()
