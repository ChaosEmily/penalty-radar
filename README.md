# 裁罰案件追蹤器 (Penalty Radar)

專為金融機構風險管理部門設計的 AI 輔助工具。自動監測金管會裁罰案件 RSS，利用 Claude AI 分析違規事實與風險啟示，並自動寄送警示信給風管人員。

## 核心功能

- **自動監測**：追蹤金管會裁罰案件 RSS，即時掌握同業裁罰動態。
- **AI 風險分析**：使用 Claude Sonnet 4.6 自動提取被罰對象、罰鍰金額、違反法條、風險啟示。
- **同業篩選器**：依據 `company_type` 設定自動判斷裁罰案件與本公司的關聯程度，同業案件標註「同業警示」。
- **自我檢核清單**：AI 根據裁罰事由自動產出本公司應檢視的具體項目，可作為內部稽核工作底稿。
- **重複違規偵測**：自動比對歷史裁罰紀錄，若同一機構多次受罰則標註累犯次數。
- **內部通知草稿**：產生可直接轉寄給相關部門的 HTML 郵件草稿。
- **裁罰趨勢統計**：產生年度/季度統計報告，含裁罰分布、累犯排名，供呈報董事會。
- **桌面通知**：Windows 10/11 原生 Toast 通知，發現新裁罰時即時提醒。
- **HTML 歷史報告**：自動產生按月分類的裁罰追蹤報告。
- **附件偵測**：自動檢測公告是否含有裁罰書附件。

## 安裝步驟

1. **複製專案**：

   ```bash
   git clone https://github.com/your-username/penalty-radar.git
   cd penalty-radar
   ```

2. **建立虛擬環境**：

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Windows 請用 .venv\Scripts\activate
   ```

3. **安裝依賴**：

   ```bash
   pip install -r requirements.txt
   ```

4. **初始化設定**：

   - 執行腳本，首次執行會自動產生 `config.json`：

     ```bash
     python setup.py
     ```

   - 開啟 `config.json` 填寫 Gmail 帳號、應用程式密碼與 Anthropic API 金鑰後，再次執行 `setup.py`。
   - 該腳本會自動在 **Windows 工作排程器** 中建立定時任務（預設每 60 分鐘檢查一次）。

## 如何使用

### 手動執行

```bash
python scraper.py
```

### 設定說明 (config.json)

| 欄位 | 說明 |
|------|------|
| `gmail_user` | 發信用 Gmail 帳號 |
| `gmail_app_password` | Gmail 應用程式密碼 (非登入密碼) |
| `recipient_email` | 風管部門收件信箱 |
| `email_strategy` | `single_emails` (每則獨立) 或 `digest` (彙整一封) |
| `max_new_per_run` | 每次最多處理幾則新案件 (首次執行時，超出的舊案件會標記為已處理) |
| `anthropic_api_key` | Claude API 金鑰 |
| `company_type` | 本公司業別（如 `銀行`、`證券`、`保險`），用於同業篩選 |

## 檔案結構

- `scraper.py`: 核心抓取、AI 分析與發信邏輯。
- `stats.py`: 裁罰趨勢統計報告產生器（`python stats.py` 或 `python stats.py 2025`）。
- `setup.py`: 初始化設定與排程註冊。
- `config.json`: 系統設定（執行時自動產生，不入版控）。
- `processed_penalties.json`: 已處理的裁罰案件紀錄（不入版控）。
- `penalty_history.json`: 機構裁罰歷史（供重複違規偵測與統計，不入版控）。
- `reports/`: 自動產生的 HTML 月報與年度統計。

## 支援平台

- **Windows 10/11**：完整支援，含 Windows 工作排程器自動註冊。
- **Linux / macOS**：`scraper.py` 可正常執行，排程需手動設定 Crontab（`setup.py` 會提示指令）。

## 免責聲明

本工具產出之分析與草稿僅供內部參考，不代表正式法律意見。所有裁罰案件請以主管機關公告及裁罰書全文為準。
