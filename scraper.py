"""
scraper.py — 裁罰案件追蹤器核心腳本 (RSS-to-Gmail 版)

執行內容：
1. 讀取 config.json 與 processed_penalties.json
2. 解析金管會裁罰案件 RSS，找出尚未處理的新裁罰
3. 針對每則新裁罰，爬取原文網頁偵測是否有 PDF/DOC 附件
4. 呼叫 Claude Sonnet 產生裁罰分析摘要與風險啟示
5. 寄送 Email 給風險管理部門
6. 寄送成功後，更新 processed_penalties.json 確保不重複發送
"""

import json
import os
import sys
import time
import re
from pathlib import Path
from datetime import datetime, timedelta

import feedparser
import requests
from bs4 import BeautifulSoup
from anthropic import Anthropic

import subprocess
import tempfile
import smtplib
from email.message import EmailMessage

# 終端機強制使用 UTF-8 輸出 (針對 Windows 排程環境)
if sys.platform == "win32":
    if sys.stdout and sys.stdout.encoding != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if sys.stderr and sys.stderr.encoding != "utf-8":
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

SCRIPT_DIR = Path(__file__).parent
CONFIG_FILE = SCRIPT_DIR / "config.json"
STATE_FILE = SCRIPT_DIR / "processed_penalties.json"
HISTORY_FILE = SCRIPT_DIR / "penalty_history.json"
PENDING_FILE = SCRIPT_DIR / "pending_digest.json"
REPORTS_DIR = SCRIPT_DIR / "reports"

# ============================================================================
# Core Utilities
# ============================================================================

def load_config() -> dict:
    if not CONFIG_FILE.exists():
        print("[ERROR] 找不到 config.json，請先執行 python setup.py 建立設定檔。")
        sys.exit(1)
    return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))

def load_state() -> list:
    if not STATE_FILE.exists():
        return []
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return []

def save_state(processed_urls: list):
    STATE_FILE.write_text(json.dumps(processed_urls, indent=4, ensure_ascii=False), encoding="utf-8")

def load_history() -> dict:
    """載入機構裁罰歷史 {entity_name: [{"date": ..., "link": ...}, ...]}"""
    if not HISTORY_FILE.exists():
        return {}
    try:
        return json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}

def save_history(history: dict):
    HISTORY_FILE.write_text(json.dumps(history, indent=4, ensure_ascii=False), encoding="utf-8")

def record_penalty(history: dict, entity: str, date: str, link: str):
    """記錄一筆裁罰到歷史"""
    if entity not in history:
        history[entity] = []
    history[entity].append({"date": date, "link": link})

def get_repeat_info(history: dict, entity: str) -> str:
    """查詢該機構過去的裁罰次數，回傳提示文字（空字串表示首次）"""
    records = history.get(entity, [])
    if not records:
        return ""
    return f"該機構近年已累計被裁罰 {len(records)} 次（本次為第 {len(records) + 1} 次）"

def load_pending() -> list:
    """載入待彙整寄送的佇列"""
    if not PENDING_FILE.exists():
        return []
    try:
        return json.loads(PENDING_FILE.read_text(encoding="utf-8"))
    except Exception:
        return []

def save_pending(items: list):
    PENDING_FILE.write_text(json.dumps(items, indent=4, ensure_ascii=False), encoding="utf-8")

def flush_pending_digest(config: dict, pending: list) -> bool:
    """檢查佇列中是否有超過 digest_hold_hours 的案件，若有則全部彙整寄出"""
    if not pending:
        return True

    hold_hours = config.get("digest_hold_hours", 24)
    oldest_ts = min(item.get("queued_at", "") for item in pending)
    if not oldest_ts:
        return True

    try:
        oldest_dt = datetime.strptime(oldest_ts, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return True

    hours_elapsed = (datetime.now() - oldest_dt).total_seconds() / 3600
    if hours_elapsed < hold_hours:
        print(f"  佇列中有 {len(pending)} 則非同業案件（最早入列 {oldest_ts}），尚未達 {hold_hours} 小時，暫不寄出。")
        return True

    print(f"  佇列中有 {len(pending)} 則非同業案件已超過 {hold_hours} 小時，開始彙整寄出...")
    success = dispatch_digest_email(config, pending)
    if success:
        save_pending([])
        print(f"  非同業彙整信寄送成功，佇列已清空。")
    return success

def check_for_attachments(url: str) -> bool:
    """爬取原始網頁，檢查是否含有特定副檔名的連結"""
    try:
        try:
            resp = requests.get(url, timeout=10)
        except requests.exceptions.SSLError:
            print(f"  [WARN] SSL 驗證失敗，改用不驗證模式：{url}")
            resp = requests.get(url, timeout=10, verify=False)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        target_exts = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.zip']
        for a_tag in soup.find_all('a', href=True):
            href = a_tag['href'].lower()
            if any(href.endswith(ext) for ext in target_exts):
                return True
        return False
    except Exception as e:
        print(f"  [WARN] 檢查附件失敗 ({url}): {e}")
        return False

# ============================================================================
# Windows Desktop Notification
# ============================================================================

def show_windows_toast(title: str, message: str):
    """在 Windows 10/11 顯示原生 Toast 桌面通知"""
    if sys.platform != "win32":
        return
    try:
        # PowerShell 腳本寫入暫存檔（UTF-8 BOM），避免中文編碼問題
        ps_script = f"""
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
$template = @"
<toast>
    <visual>
        <binding template="ToastGeneric">
            <text>{title}</text>
            <text>{message}</text>
        </binding>
    </visual>
</toast>
"@
$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml($template)
$toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("PenaltyRadar").Show($toast)
"""
        tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.ps1', delete=False, encoding='utf-8-sig')
        tmp.write(ps_script)
        tmp.close()

        subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", tmp.name],
            capture_output=True, timeout=10
        )
        os.unlink(tmp.name)
    except Exception as e:
        print(f"  [WARN] 桌面通知顯示失敗（不影響主流程）: {e}")

# ============================================================================
# AI Processing
# ============================================================================

def process_with_claude(item: dict, client: Anthropic, company_type: str = "") -> dict:
    """呼叫 Claude 分析裁罰案件，產生風險啟示摘要"""
    print(f"  正在讓 AI 分析：{item['title'][:60]}...")

    company_ctx = ""
    relevance_field = ""
    if company_type:
        company_ctx = f"\n本公司為「{company_type}」業者，請據此判斷此裁罰案件與本公司的關聯程度。"
        relevance_field = f"""    "relevance": "判斷被裁罰對象是否與本公司屬同一業別（{company_type}）。若是同業填 high；若為金融業但非同業填 medium；若為非金融業填 low","""

    prompt = f"""請扮演專業的台灣金融控股公司風險管理分析師。
我收到一則金管會的裁罰案件公告，請幫我分析此案件對本公司的風險啟示。{company_ctx}

標題：{item['title']}
發布時間：{item['published']}
原文連結：{item['link']}
公告內容：
{item['summary']}

請直接輸出一個格式正確的 JSON Object，不要包含任何其他文字或 Markdown 標記：
{{
    "penalized_entity": "被裁罰的機構或個人名稱",
    "penalty_amount": "罰鍰金額（如有），例如「新臺幣 600 萬元」；若無金額則填「未載明」",
    "violated_laws": "違反的法條（簡要列舉）",
    "violation_summary": "一段白話文摘要，說明被裁罰的原因與事實",
    "risk_implication": "用 HTML 格式列出風險啟示。必須使用 <ul><li> 條列，每個風險類型用 <strong> 粗體標示（如 <strong>作業風險</strong>）。不要使用 \\n 換行，只用 HTML 標籤。",
    "suggested_departments": "建議通知的部門（例如：稽核室、法遵部、資訊安全部等）",
{relevance_field}
    "checklist": "用 HTML 格式產出自我檢核清單。根據此裁罰的違規事由，列出本公司應檢視的具體項目。格式：<ol><li><strong>[檢核項目]</strong>：[具體檢視內容與標準]</li></ol>。至少 3 項，不超過 6 項。不要使用 \\n，只用 HTML 標籤。",
    "draft_subject": "【裁罰警示】{item['title'][:50]}...",
    "draft_body": "HTML 格式的內部通知草稿。嚴格依照以下模板填入內容，不可省略任何區塊：<p>各位同仁好，</p><p>金管會於[日期]公布裁罰案件，茲摘要通報如下，請各單位檢視自身作業流程。</p><h4>一、被裁罰對象</h4><p>[機構名稱]</p><h4>二、違規事實</h4><ul><li>[事實1]</li><li>[事實2]</li></ul><h4>三、裁罰依據與金額</h4><p>違反[法條]，裁處[金額]。</p><h4>四、風險啟示與自我檢視建議</h4><ul><li><strong>[風險類型1]</strong>：[具體建議]</li><li><strong>[風險類型2]</strong>：[具體建議]</li></ul><h4>五、建議行動</h4><ul><li>[部門A]：請於[期限]前完成[事項]</li><li>[部門B]：請檢視[事項]</li></ul><p>詳細裁罰書請見：<a href='原文連結'>原文連結</a></p><p>如有疑義請洽風險管理部。</p>  不要使用 \\n，只用 HTML 標籤排版。不要包含 <html><body> 標籤。"
}}
"""
    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4000,
            temperature=0.2,
            messages=[{"role": "user", "content": prompt}]
        )
        raw_text = response.content[0].text.strip()

        # 移除 markdown 程式碼區塊標記
        raw_text = re.sub(r'^```(?:json)?\s*', '', raw_text)
        raw_text = re.sub(r'\s*```$', '', raw_text)

        # 抓取 JSON 內容
        json_match = re.search(r'(\{.*\})', raw_text, re.DOTALL | re.MULTILINE)
        if json_match:
            full_content = json_match.group(1)
            last_brace = full_content.rfind('}')
            json_text = full_content[:last_brace+1]
        else:
            json_text = raw_text

        try:
            return json.loads(json_text)
        except json.JSONDecodeError:
            cleaned_text = json_text.replace('\n', ' ').replace('\r', '')
            try:
                return json.loads(cleaned_text)
            except Exception:
                # 逐欄位 regex 提取
                entity_m = re.search(r'"penalized_entity"\s*:\s*"(.*?)"(?=\s*[,}])', json_text, re.DOTALL)
                amount_m = re.search(r'"penalty_amount"\s*:\s*"(.*?)"(?=\s*[,}])', json_text, re.DOTALL)
                laws_m = re.search(r'"violated_laws"\s*:\s*"(.*?)"(?=\s*[,}])', json_text, re.DOTALL)
                summary_m = re.search(r'"violation_summary"\s*:\s*"(.*?)"(?=\s*[,}])', json_text, re.DOTALL)
                risk_m = re.search(r'"risk_implication"\s*:\s*"(.*?)"(?=\s*[,}])', json_text, re.DOTALL)
                dept_m = re.search(r'"suggested_departments"\s*:\s*"(.*?)"(?=\s*[,}])', json_text, re.DOTALL)
                relevance_m = re.search(r'"relevance"\s*:\s*"(.*?)"(?=\s*[,}])', json_text, re.DOTALL)
                checklist_m = re.search(r'"checklist"\s*:\s*"(.*?)"(?=\s*[,}])', json_text, re.DOTALL)
                subject_m = re.search(r'"draft_subject"\s*:\s*"(.*?)"(?=\s*[,}])', json_text, re.DOTALL)
                body_m = re.search(r'"draft_body"\s*:\s*"(.*?)"(?=\s*[,}"\s]*\})', json_text, re.DOTALL)

                if summary_m and subject_m:
                    return {
                        "penalized_entity": entity_m.group(1).strip() if entity_m else "未知",
                        "penalty_amount": amount_m.group(1).strip() if amount_m else "未載明",
                        "violated_laws": laws_m.group(1).strip() if laws_m else "未載明",
                        "violation_summary": summary_m.group(1).strip(),
                        "risk_implication": risk_m.group(1).strip() if risk_m else "",
                        "suggested_departments": dept_m.group(1).strip() if dept_m else "",
                        "relevance": relevance_m.group(1).strip() if relevance_m else "",
                        "checklist": checklist_m.group(1).strip() if checklist_m else "",
                        "draft_subject": subject_m.group(1).strip(),
                        "draft_body": body_m.group(1).strip() if body_m else summary_m.group(1).strip(),
                    }
                raise
    except Exception as e:
        print(f"  [ERROR] Claude API 呼叫失敗或 JSON 解析錯誤：{e}")
        return None

# ============================================================================
# Email Dispatcher
# ============================================================================

def smtp_connection(config: dict):
    """建立共用的 SMTP 連線（context manager）"""
    server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    server.login(config["gmail_user"], config["gmail_app_password"])
    return server

def send_smtp_email(config: dict, msg, server=None) -> bool:
    """底層共用的 SMTP 寄信邏輯。可接受既有連線或自行建立。"""
    try:
        if server:
            server.send_message(msg)
        else:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
                s.login(config["gmail_user"], config["gmail_app_password"])
                s.send_message(msg)
        return True
    except Exception as e:
        print(f"  [ERROR] 發送 Email 失敗：{e}")
        return False

def dispatch_single_emails(config: dict, results: list) -> bool:
    """每則裁罰案件獨立發送一封 Email（共用 SMTP 連線）"""
    all_success = True
    try:
        server = smtp_connection(config)
    except Exception as e:
        print(f"  [ERROR] SMTP 連線失敗：{e}")
        return False

    try:
        for res in results:
            ai = res['ai_output']
            # 將所有文字欄位中的 \n 轉換為 <br>（防止 AI 回傳純文字換行）
            for key in ('risk_implication', 'violation_summary', 'draft_body', 'checklist'):
                if key in ai and isinstance(ai[key], str):
                    ai[key] = ai[key].replace('\\n', '<br>').replace('\n', '<br>')
            # log 移至主旨替換後

            attachment_notice = ""
            if res.get('has_attachments'):
                attachment_notice = '<div style="background-color: #fff3cd; color: #856404; padding: 10px; border: 1px solid #ffeeba; border-radius: 4px; margin: 10px 0; font-weight: bold;">包含附件，請務必點擊原文連結詳閱裁罰書全文。</div>'

            html_content = f"""
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <div style="background-color: #fce4ec; border-left: 5px solid #c62828; padding: 15px; margin-bottom: 20px;">
                <h3 style="margin-top: 0; color: #c62828;">裁罰案件摘要</h3>
                <p><strong>{res['title']}</strong></p>
                <table style="border-collapse: collapse; width: 100%; margin: 10px 0;">
                    <tr><td style="padding: 4px 8px; font-weight: bold; width: 120px;">被裁罰對象</td><td style="padding: 4px 8px;">{ai.get('penalized_entity', '未知')}{"  <span style='background:#ff6f00;color:#fff;padding:1px 8px;border-radius:3px;font-size:0.8em;'>" + ai['repeat_info'] + "</span>" if ai.get('repeat_info') else ""}</td></tr>
                    <tr><td style="padding: 4px 8px; font-weight: bold;">罰鍰金額</td><td style="padding: 4px 8px;">{ai.get('penalty_amount', '未載明')}</td></tr>
                    <tr><td style="padding: 4px 8px; font-weight: bold;">違反法規</td><td style="padding: 4px 8px;">{ai.get('violated_laws', '未載明')}</td></tr>
                    <tr><td style="padding: 4px 8px; font-weight: bold;">建議通知</td><td style="padding: 4px 8px;">{ai.get('suggested_departments', '')}</td></tr>
                </table>
                <p><strong>違規事實：</strong>{ai.get('violation_summary', '')}</p>
                {attachment_notice}
            </div>

            <div style="background-color: #e3f2fd; border-left: 5px solid #1565c0; padding: 15px; margin-bottom: 20px;">
                <h3 style="margin-top: 0; color: #1565c0;">風險啟示與自我檢視建議</h3>
                {"<div style='display:inline-block; background:#c62828; color:#fff; padding:2px 10px; border-radius:4px; font-size:0.85em; margin-bottom:8px;'>同業案件</div>" if ai.get('relevance') == 'high' else ""}
                <p>{ai.get('risk_implication', '')}</p>
            </div>

            <div style="background-color: #f3e5f5; border-left: 5px solid #7b1fa2; padding: 15px; margin-bottom: 20px;">
                <h3 style="margin-top: 0; color: #7b1fa2;">自我檢核清單</h3>
                {ai.get('checklist', '<p>（未產生檢核項目）</p>')}
            </div>

            <hr/>
            <h3>內部通知草稿 (請確認後直接轉寄本信)</h3>
            <p style="color: #666;">建議主旨: {ai['draft_subject']}</p>
            <div style="border: 1px solid #ddd; padding: 15px; border-radius: 5px;">
                {ai['draft_body']}
            </div>

            <p style="color: #999; font-size: 0.85em; margin-top: 20px;">
                原文連結：<a href="{res['link']}">{res['link']}</a><br/>
                本摘要由 AI 輔助產生，不構成正式法律意見。請以主管機關公告為準。
            </p>
        </body>
        </html>
            """

            # 同業標記：relevance=high 時在主旨前加標籤
            subject = ai['draft_subject']
            relevance = ai.get('relevance', '')
            if relevance == 'high':
                subject = "【同業警示】" + subject.replace("【裁罰警示】", "")

            print(f"  正在發送 Email: {subject[:50]}...")

            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = config["gmail_user"]
            msg["To"] = config["recipient_email"]
            msg.set_content(html_content, subtype='html')

            if not send_smtp_email(config, msg, server):
                all_success = False
                break
    finally:
        server.quit()

    return all_success

def dispatch_digest_email(config: dict, results: list) -> bool:
    """彙整多則裁罰案件為一封摘要信"""
    print(f"  正在發送彙整信含 {len(results)} 則裁罰案件...")

    html_body = """<html><body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    <h2 style="color: #c62828;">裁罰案件追蹤 — 本期摘要</h2>
    <p>風管同仁您好，以下為最新的金管會裁罰案件摘要。請評估是否需要通知相關部門進行內部檢視。</p>
    """

    for idx, res in enumerate(results, 1):
        ai = res['ai_output']
        attachment_notice = ""
        if res.get('has_attachments'):
            attachment_notice = '<span style="color: #856404; font-weight: bold;">[ 含附件 ]</span>'

        html_body += f"""
        <div style="background-color: #fce4ec; border-left: 5px solid #c62828; padding: 12px 16px; margin: 16px 0; border-radius: 4px;">
            <strong>{idx}. {ai.get('penalized_entity', '未知')}</strong> — {ai.get('penalty_amount', '未載明')} {attachment_notice}<br/>
            <span style="color: #555;">{ai.get('violation_summary', '')}</span><br/>
            <span style="color: #1565c0;"><strong>風險啟示：</strong>{ai.get('risk_implication', '')}</span><br/>
            <a href="{res['link']}" style="color: #0056b3; font-size: 0.9em;">原文連結</a>
        </div>
        """

    html_body += """
    <p style="color: #999; font-size: 0.85em; margin-top: 20px; border-top: 1px solid #ddd; padding-top: 10px;">
    本摘要由 AI 輔助產生，不構成正式法律意見。請以主管機關公告為準。
    </p>
    </body></html>"""

    msg = EmailMessage()
    msg["Subject"] = f"【裁罰追蹤】本期共 {len(results)} 則新裁罰案件"
    msg["From"] = config["gmail_user"]
    msg["To"] = config["recipient_email"]
    msg.set_content(html_body, subtype='html')

    return send_smtp_email(config, msg)

# ============================================================================
# HTML Report
# ============================================================================

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<title>裁罰案件追蹤 — {month_label}</title>
<style>
  body {{ font-family: "Microsoft JhengHei", Arial, sans-serif; max-width: 900px; margin: 0 auto; padding: 20px; color: #333; }}
  h1 {{ color: #c62828; border-bottom: 2px solid #c62828; padding-bottom: 10px; }}
  .entry {{ padding: 12px 16px; margin: 16px 0; border-radius: 4px; background-color: #fce4ec; border-left: 5px solid #c62828; }}
  .entry h3 {{ margin: 0 0 8px 0; }}
  .attachment-warn {{ background-color: #fff3cd; color: #856404; padding: 8px 12px; border: 1px solid #ffeeba; border-radius: 4px; margin: 10px 0; font-weight: bold; font-size: 0.9em; }}
  .meta {{ color: #666; font-size: 0.9em; margin-bottom: 8px; }}
  .summary {{ margin: 8px 0; }}
  .risk {{ margin: 8px 0; padding: 8px; background: #e3f2fd; border-radius: 4px; }}
  .disclaimer {{ color: #999; font-size: 0.85em; margin-top: 40px; border-top: 1px solid #ddd; padding-top: 10px; }}
  a {{ color: #0056b3; }}
</style>
</head>
<body>
<h1>裁罰案件追蹤 — {month_label}</h1>
<p>本報告由系統自動產生，最後更新：{updated_at}</p>

<div id="entries-container">
<!-- ENTRIES -->
{entries_html}
<!-- /ENTRIES -->
</div>

<div class="disclaimer">本報告為 AI 輔助分析，不構成正式法律意見。請以主管機關裁罰書全文為準。</div>
</body>
</html>"""

ENTRY_TEMPLATE = """<div class="entry" data-published="{published}">
<h3>{penalized_entity} — {penalty_amount}</h3>
<div class="meta">發布時間：{published} ｜ <a href="{link}" target="_blank">原文連結</a></div>
{attachment_html}
<div class="summary"><strong>違規事實：</strong>{violation_summary}</div>
<div class="risk"><strong>風險啟示：</strong>{risk_implication}</div>
</div>"""


def append_to_html_report(results: list):
    """將本次結果追加到當月 HTML 報告。"""
    REPORTS_DIR.mkdir(exist_ok=True)
    now = datetime.now()
    month_key = now.strftime("%Y-%m")
    month_label = now.strftime("%Y 年 %m 月")
    report_file = REPORTS_DIR / f"{month_key}.html"

    all_entries = []

    # 讀取既有 entry
    if report_file.exists():
        try:
            old_html = report_file.read_text(encoding="utf-8")
            soup = BeautifulSoup(old_html, "html.parser")
            existing_divs = soup.find_all("div", class_="entry")
            for div in existing_divs:
                meta_text = div.find("div", class_="meta").get_text()
                date_match = re.search(r"發布時間：(\d{4}-\d{2}-\d{2})", meta_text)
                pub_date = date_match.group(1) if date_match else "0000-00-00"
                all_entries.append({"html": str(div), "pub_date": pub_date})
        except Exception as e:
            print(f"  [WARNING] 解析舊報告失敗，將建立新檔：{e}")

    # 加入新 entry
    for res in results:
        ai = res['ai_output']
        attachment_html = ""
        if res.get("has_attachments"):
            attachment_html = '<div class="attachment-warn">包含附件，請務必點擊原文連結詳閱裁罰書全文。</div>'

        pub_date = res.get("published", "0000-00-00")
        if len(pub_date) > 10:
            pub_date = pub_date[:10]

        entry_html = ENTRY_TEMPLATE.format(
            penalized_entity=ai.get("penalized_entity", "未知"),
            penalty_amount=ai.get("penalty_amount", "未載明"),
            published=pub_date,
            link=res["link"],
            attachment_html=attachment_html,
            violation_summary=ai.get("violation_summary", ""),
            risk_implication=ai.get("risk_implication", ""),
        )
        all_entries.append({"html": entry_html, "pub_date": pub_date})

    # 排序（由舊到新）
    all_entries.sort(key=lambda x: x["pub_date"])

    final_entries_html = "\n".join([e["html"] for e in all_entries])
    html = HTML_TEMPLATE.format(
        month_label=month_label,
        updated_at=now.strftime("%Y-%m-%d %H:%M"),
        entries_html=final_entries_html,
    )

    report_file.write_text(html, encoding="utf-8")
    print(f"  HTML 報告已更新：{report_file}")
    update_index_html()


def update_index_html():
    """重新產生 index.html，列出所有月份報告連結。"""
    report_files = sorted(REPORTS_DIR.glob("2*.html"), reverse=True)
    if not report_files:
        return

    links = ""
    for f in report_files:
        month_key = f.stem
        links += f'<li><a href="{f.name}">{month_key}</a></li>\n'

    index_html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<title>裁罰案件追蹤 — 歷史報告索引</title>
<style>
  body {{ font-family: "Microsoft JhengHei", Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }}
  h1 {{ color: #c62828; }}
  li {{ margin: 8px 0; font-size: 1.1em; }}
  a {{ color: #0056b3; }}
</style>
</head>
<body>
<h1>裁罰案件追蹤 — 歷史報告索引</h1>
<ul>
{links}</ul>
</body>
</html>"""

    (REPORTS_DIR / "index.html").write_text(index_html, encoding="utf-8")


def check_retention_reminder():
    """檢查超過 5 年保存期限的報告檔案。"""
    if not REPORTS_DIR.exists():
        return

    cutoff = datetime.now() - timedelta(days=5 * 365)
    expired = []
    for f in REPORTS_DIR.glob("2*.html"):
        try:
            file_month = datetime.strptime(f.stem, "%Y-%m")
            if file_month < cutoff:
                expired.append(f.name)
        except ValueError:
            continue

    if expired:
        print(f"\n  以下 {len(expired)} 份報告已超過 5 年保存期限，請評估是否歸檔：")
        for name in expired:
            print(f"     - {name}")

# ============================================================================
# Run History Logging
# ============================================================================

RUN_HISTORY_FILE = SCRIPT_DIR / "run_history.jsonl"

def log_run(rss_new: int = 0, ai_processed: int = 0, email_sent: bool = False, error: str = None):
    """每次執行結束時追加一行 JSON 到 run_history.jsonl"""
    record = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "rss_new": rss_new,
        "ai_processed": ai_processed,
        "email_sent": email_sent,
        "error": error,
    }
    try:
        with open(RUN_HISTORY_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception as e:
        print(f"  [WARN] 無法寫入運行紀錄: {e}")

# ============================================================================
# Main
# ============================================================================

def main():
    print("=" * 60)
    print(f"  Penalty Radar 啟動時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    config = load_config()
    processed_list = load_state()
    processed_urls = set(processed_list)
    penalty_history = load_history()

    check_retention_reminder()

    api_key = config.get("anthropic_api_key")
    if not api_key:
        log_run(error="Anthropic API Key 未設定")
        print("[ERROR] Anthropic API Key 未設定，請檢查 config.json。")
        sys.exit(1)

    # 1. 爬取 RSS
    print("[1/3] 正在檢查最新的裁罰案件...")
    rss_url = config.get("rss_url", "https://www.fsc.gov.tw/RSS/Messages?serno=201202290003&language=chinese")

    new_items = []
    skipped_items = []
    rss_fetch_ok = False

    try:
        resp = requests.get(rss_url, timeout=15, verify=False)
        # 重要：使用 resp.content (bytes) 而非 resp.text，避免編碼錯誤
        feed = feedparser.parse(resp.content)
        print(f"  RSS 來源取得 {len(feed.entries)} 筆裁罰案件")
        rss_fetch_ok = True

        for entry in feed.entries:
            link = entry.link
            if link in processed_urls:
                continue

            item = {
                "title": entry.title,
                "link": link,
                "published": getattr(entry, 'published', '未知時間'),
                "summary": getattr(entry, 'summary', ''),
            }

            if len(new_items) < config.get("max_new_per_run", 3):
                new_items.append(item)
            else:
                skipped_items.append(item)

    except Exception as e:
        print(f"  [ERROR] 讀取 RSS 失敗：{e}")
        log_run(error=f"RSS 讀取失敗: {e}")
        sys.exit(1)

    # 首次執行：將超出數量限制的舊案件標記為已處理
    if skipped_items:
        print(f"  將 {len(skipped_items)} 則舊案件標記為已處理（首次執行限制）")
        for item in skipped_items:
            processed_urls.add(item['link'])
        save_state(list(processed_urls)[-2000:])

    if not new_items:
        # 即使沒有新案件，仍檢查佇列是否到期
        strategy = config.get("email_strategy", "priority")
        if strategy == "priority":
            pending = load_pending()
            if pending:
                flush_pending_digest(config, pending)
        log_run(rss_new=0)
        print("  目前沒有新的裁罰案件需要處理。")
        sys.exit(0)

    # 2. AI 分析
    print(f"\n[2/3] 發現 {len(new_items)} 則新裁罰案件，開始進行 AI 分析...")
    client = Anthropic(api_key=api_key)
    company_type = config.get("company_type", "")
    results = []
    for item in new_items:
        item['has_attachments'] = check_for_attachments(item['link'])
        ai_data = process_with_claude(item, client, company_type)
        if ai_data:
            # 重複違規偵測
            entity = ai_data.get('penalized_entity', '未知')
            repeat_info = get_repeat_info(penalty_history, entity)
            if repeat_info:
                print(f"    [注意] {repeat_info}")
            ai_data['repeat_info'] = repeat_info

            item['ai_output'] = ai_data
            results.append(item)
        time.sleep(1)

    if not results:
        log_run(rss_new=len(new_items), ai_processed=0, error="AI 分析全部失敗")
        print("  未能成功產生任何 AI 摘要，跳過發信。")
        sys.exit(1)

    # 3. 發送 Email
    print("\n[3/3] 開始配送 Email 報告...")
    strategy = config.get("email_strategy", "priority")

    if strategy == "priority":
        # 分級寄送：同業即時，非同業延遲彙整
        high_results = [r for r in results if r['ai_output'].get('relevance') == 'high']
        deferred_results = [r for r in results if r['ai_output'].get('relevance') != 'high']

        success = True

        # 同業案件立即寄出
        if high_results:
            print(f"  同業案件 {len(high_results)} 則，立即寄出...")
            success = dispatch_single_emails(config, high_results)

        # 非同業案件加入佇列
        if deferred_results:
            pending = load_pending()
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for r in deferred_results:
                r['queued_at'] = now_str
                pending.append(r)
            save_pending(pending)
            print(f"  非同業案件 {len(deferred_results)} 則，已加入彙整佇列。")

        # 檢查佇列是否到期
        pending = load_pending()
        if pending:
            flush_result = flush_pending_digest(config, pending)
            if not flush_result:
                success = False

    elif strategy == "digest":
        success = dispatch_digest_email(config, results)
    else:
        success = dispatch_single_emails(config, results)

    if success:
        print("\n  處理完成！正在更新已處理清單...")
        for r in results:
            processed_urls.add(r['link'])
            # 記錄裁罰歷史（供重複違規偵測）
            entity = r['ai_output'].get('penalized_entity', '未知')
            pub_date = r.get('published', '')[:10]
            record_penalty(penalty_history, entity, pub_date, r['link'])
        save_state(list(processed_urls)[-2000:])
        save_history(penalty_history)
        append_to_html_report(results)
        log_run(rss_new=len(new_items), ai_processed=len(results), email_sent=True)

        # 桌面通知（僅同業即時案件觸發）
        immediate = [r for r in results if r['ai_output'].get('relevance') == 'high'] if strategy == "priority" else results
        if immediate:
            entities = [r['ai_output'].get('penalized_entity', '未知') for r in immediate]
            toast_msg = "、".join(entities[:3])
            if len(entities) > 3:
                toast_msg += f" 等 {len(entities)} 件"
            show_windows_toast("裁罰案件追蹤器", f"發現 {len(immediate)} 則同業裁罰：{toast_msg}")
    else:
        log_run(rss_new=len(new_items), ai_processed=len(results), error="Email 發送失敗")
        print("\n  發送過程中發生錯誤，狀態將不會更新，下次將重試。")
        sys.exit(1)

if __name__ == "__main__":
    main()
