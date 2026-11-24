<%
'=====================================================================
' 護理標準查詢助理  設定範本（config.sample.asp）
'   使用方式：複製本檔為 config.asp，再把下方 ★ 的值改成實際環境。
'   本檔只放「不含真實值」的範例，可以 commit；config.asp 不可 commit。
'=====================================================================

'---------- 0. 機構 ----------
Const SC_ORG_NAME  = "○○醫院護理部"           ' ★ 寫入 AI 系統提示詞的機構名稱
Const SC_HOME_URL  = "../index.asp"             ' ★ 未登入時導回、及右上角「回內網首頁」的網址

'---------- 1. 路徑設定 ----------
Const SC_DB        = "fpdb/S_coach.mdb"          ' 本系統資料庫（相對於網站根目錄；由 S_coach_setup.asp 自動建立）
Const SC_PROVIDER  = "Microsoft.Jet.OLEDB.4.0"   ' 若用 ACE，改 "Microsoft.ACE.OLEDB.12.0"
Const SC_CATALOG   = "X:\path\to\amend_work.mdb" ' ★ 護理標準系統的資料庫（讀 catalog 表；只讀不寫）
Const SC_PDF_ROOT  = "X:\path\to\paper"          ' ★ 標準 PDF 資料夾：底下是 CP\、CPT\、CR\… 各類型資料夾
Const SC_PDFTOTEXT = "X:\path\to\pdftotext.exe"  ' ★ xpdf 的 pdftotext.exe 位置
Const SC_TMP       = "paper/sc_tmp"              ' 轉檔暫存資料夾（IIS 帳號需可寫入，會自動建立）

'---------- 1b. 病房單位次專科標準（選用）----------
Const SC_SUS_DB       = ""                       ' ★ 次專科標準資料庫，例 "X:\path\to\D_sus.mdb"；留空 = 不使用
Const SC_SUS_PDF_ROOT = ""                       ' ★ 次專科標準 PDF 資料夾，例 "X:\path\to\paper\D_sus"

'---------- 2. 權限 ----------
Const SC_REQUIRE_LOGIN = True                    ' True = 沒有 worknb cookie 就導回內網首頁
Const SC_ADMINS = ",A000000,"                    ' ★ 管理者員工代號，前後加逗號，例 ",A000001,A000002,"

'---------- 3. AI 摘要 ----------
Const SC_AI_ENABLED  = False        ' False = 只顯示標準原文段落；True = 由 AI 依原文整理
'  SC_AI_PROVIDER：
'    "ollama"    院內自架 Ollama            → SC_AI_URL 填 "http://<主機>:11434/api/chat"
'    "openai"    OpenAI 相容格式            → SC_AI_URL 填該服務的 /chat/completions 網址
'    "anthropic" Claude API                → SC_AI_URL 填 "https://api.anthropic.com/v1/messages"
Const SC_AI_PROVIDER = "ollama"
Const SC_AI_URL      = "http://127.0.0.1:11434/api/chat"   ' ★ AI 服務網址
Const SC_MODEL       = "your-model:tag"     ' ★ 模型名稱（ollama pull 時的名稱）
Const SC_AI_THREADS  = 0            ' AI 最多用幾個 CPU 執行緒（0 = 自動）
Const SC_API_KEY     = ""           ' Ollama 不需要；雲端服務才填（不要 commit 真實金鑰）
Const SC_PROXY       = ""           ' 連外才需要
Const SC_AI_TIMEOUT  = 180          ' 等 AI 回答最多幾秒

'---------- 4. 檢索參數 ----------
Const SC_TOPK       = 6      ' 每題最多帶幾段原文
Const SC_PER_STD    = 2      ' 同一份標準最多幾段
Const SC_MIN_COVER  = 0.35   ' 問題關鍵詞命中率低於此值 → 視為「標準中找不到」，不呼叫 AI
Const SC_CHUNK_LEN  = 450    ' 每段原文約幾個字
%>
