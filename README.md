# GCP Agent Search & 建置整合標準作業程序 (SOP)

本專案提供 **Google Cloud Platform (GCP) Vertex AI Search (Discovery Engine)** 企業級智慧搜尋建置整合指南與完整原始碼，包含：

1. 📑 **[GCP_Vertex_AI_Search_SOP.pptx](GCP_Vertex_AI_Search_SOP.pptx)**：呈報長官專用之 10 步驟標準 SOP 簡報檔（支援 Windows PowerPoint 及 Google 簡報放映）。
2. 🌐 **[gcpsearchsop.html](gcpsearchsop.html)**：清新明亮藍色調之 SOP 互動式網頁（支援投影片模式與完整全覽模式，附帶一鍵複製指令與檢核點）。
3. 💻 **Python 後端與自訂前端**：使用服務帳戶憑證安全串接 Discovery Engine API，實測回應延遲低於 300ms。

## 架構設計與安全規範

```text
[ 前端瀏覽器 (HTML / JS) ]
           │
           │  (REST API: POST /api/search)
           ▼
[ 後端 Python (Flask + GCP SDK) ]
           │
           │  (使用 gen-lang-client-*.json 安全身分驗證)
           ▼
[ GCP Vertex AI Search (Discovery Engine) ]
```

> **資安特別說明**：
> GCP 服務帳戶的私密金鑰（JSON）**嚴禁直接暴露在前端 HTML 或瀏覽器 JavaScript 中**。前端透過呼叫後端 API，後端負責與 GCP 溝通並解析回傳結果，確保憑證安全。

---

## 專案結構

```text
gcp_search_sop/
├── GCP_Vertex_AI_Search_SOP.pptx  # 10 步驟標準 SOP 簡報檔 (可上傳 Google Drive / Windows PPT)
├── gcpsearchsop.html              # 清新明亮藍色系 SOP 互動式網頁 (支援簡報與全覽雙模式)
├── generate_presentation.py       # PPTX 簡報自動生成腳本 (python-pptx)
├── app.py                         # 後端 Flask API 主程式
├── requirements.txt               # Python 相依套件清單
├── .env.example                   # 環境變數範本檔
├── templates/
│   ├── index.html                 # 模式一：自訂 UI 搜尋介面 (Custom UI)
│   ├── gcpsearchsop.html          # SOP 標準作業程序頁面
│   └── widget.html                # 模式二：Google 官方小工具 (<gen-search-widget>)
└── static/
    ├── css/
    │   └── style.css              # 前端深色玻璃態質感樣式
    └── js/
        └── app.js                 # 前端搜尋互動邏輯
```

---

## 支援兩種運作模式

1. **模式一：自訂 UI 搜尋介面（推薦，自由度最高）**
   - 網址：`http://localhost:5000/`
   - 架構：純 HTML/CSS/JS 呼叫 Python 後端 `/api/search`，由後端使用服務帳戶金鑰連線 GCP，回傳結構化資料渲染。

2. **模式二：Google 官方現成小工具 (<gen-search-widget>)**
   - 網址：`http://localhost:5000/widget`
   - 架構：前端載入 Google 官方 `<gen-search-widget>` Web Component，後端提供 `/api/token` 動態產生 OAuth Access Token，並由 JavaScript 賦值：
     ```javascript
     const searchWidget = document.querySelector('gen-search-widget');
     searchWidget.authToken = "<由後端產生的短期 OAuth Token>";
     ```

---

## 快速啟動

### 1. 建立並啟動虛擬環境 (Virtualenv)

```bash
python3 -m venv venv
source venv/bin/activate
```

### 2. 安裝相依套件

```bash
pip install -r requirements.txt
```

### 3. 設定環境變數 (`.env`)

本專案已自動偵測並配置：
```ini
GCP_KEY_FILE=gen-lang-client-0687915907-5ba27441b36c.json
GCP_PROJECT_ID=gen-lang-client-0687915907
GCP_LOCATION=global
GCP_ENGINE_ID=pbsearch_1791792965002
GCP_DATA_STORE_ID=pbdatastore_1791793664793
PORT=5000
DEBUG=True
```

### 4. 啟動伺服器

```bash
python3 app.py
```

### 5. 開啟瀏覽器體驗

打開瀏覽器前往：
[http://localhost:5000](http://localhost:5000)

---

## API 介面規格

### 1. 執行搜尋 (`POST /api/search`)
- **請求參數 (JSON)**:
  ```json
  {
    "query": "半導體測試",
    "page_size": 10,
    "offset": 0
  }
  ```
- **回傳內容**:
  ```json
  {
    "status": "success",
    "query": "半導體測試",
    "total_size": 421,
    "latency_ms": 285.4,
    "results": [
      {
        "id": "0",
        "title": "Mission Mode Testing: Aligning Test Strategy ...",
        "html_title": "Mission Mode <b>Testing</b>: Aligning ...",
        "link": "https://www.semicontaiwan.org/zh/node/15611",
        "display_link": "www.semicontaiwan.org",
        "snippet": "行政資源中心副總經理...",
        "html_snippet": "行政資源中心副總經理, <b>日月光</b> ...",
        "thumbnail": "https://..."
      }
    ],
    "summary": "AI 生成摘要（若啟用）"
  }
  ```

### 2. 取得系統狀態 (`GET /api/info`)
- 回傳當前 GCP 連線專案、服務帳戶 Email、Engine ID 與健康狀態。
