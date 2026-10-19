import os
import time
import logging
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from dotenv import load_dotenv
import google.auth.transport.requests
from google.oauth2 import service_account
from google.cloud import discoveryengine_v1 as discoveryengine

# 載入環境變數
load_dotenv()

# 設定日誌
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

# 初始化 Flask
app = Flask(__name__, template_folder="templates", static_folder="static")
CORS(app)

# GCP 設定參數
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
KEY_FILE_NAME = os.getenv("GCP_KEY_FILE", "gen-lang-client-0687915907-5ba27441b36c.json")
KEY_FILE_PATH = os.path.join(BASE_DIR, KEY_FILE_NAME)

PROJECT_ID = os.getenv("GCP_PROJECT_ID", "gen-lang-client-0687915907")
LOCATION = os.getenv("GCP_LOCATION", "global")
ENGINE_ID = os.getenv("GCP_ENGINE_ID", "pbsearch_1791792965002")
DATA_STORE_ID = os.getenv("GCP_DATA_STORE_ID", "pbdatastore_1791793664793")
WIDGET_CONFIG_ID = os.getenv("GCP_WIDGET_CONFIG_ID", "b521b4db-4f6a-482a-aab4-0fc1d31a7edb")

# 初始化 GCP Client
credentials = None
search_client = None
service_account_email = None

try:
    if os.path.exists(KEY_FILE_PATH):
        credentials = service_account.Credentials.from_service_account_file(
            KEY_FILE_PATH,
            scopes=["https://www.googleapis.com/auth/cloud-platform"]
        )
        search_client = discoveryengine.SearchServiceClient(credentials=credentials)
        service_account_email = credentials.service_account_email
        logger.info(f"✅ 成功載入 GCP 憑證: {KEY_FILE_NAME}")
        logger.info(f"   服務帳戶: {service_account_email}")
        logger.info(f"   專案 ID: {PROJECT_ID}")
        logger.info(f"   Engine ID: {ENGINE_ID}")
    else:
        logger.error(f"❌ 找不到憑證檔案: {KEY_FILE_PATH}")
except Exception as e:
    logger.error(f"❌ 初始化 GCP Client 失敗: {str(e)}", exc_info=True)


def get_serving_config_path():
    """取得 Discovery Engine 的 Serving Config 資源路徑"""
    if ENGINE_ID:
        # 使用 App / Engine 搜尋配置
        return f"projects/{PROJECT_ID}/locations/{LOCATION}/collections/default_collection/engines/{ENGINE_ID}/servingConfigs/default_search"
    elif DATA_STORE_ID:
        # 使用 DataStore 搜尋配置
        return f"projects/{PROJECT_ID}/locations/{LOCATION}/collections/default_collection/dataStores/{DATA_STORE_ID}/servingConfigs/default_search"
    else:
        raise ValueError("必須指定 GCP_ENGINE_ID 或 GCP_DATA_STORE_ID")


def extract_thumbnail(pagemap):
    """從 pagemap 結構中嘗試提取圖片縮圖 URL"""
    if not pagemap:
        return None
    try:
        # 1. 檢查 cse_image
        if "cse_image" in pagemap and len(pagemap["cse_image"]) > 0:
            return pagemap["cse_image"][0].get("src")
        # 2. 檢查 cse_thumbnail
        if "cse_thumbnail" in pagemap and len(pagemap["cse_thumbnail"]) > 0:
            return pagemap["cse_thumbnail"][0].get("src")
    except Exception:
        pass
    return None


def get_oauth_token():
    """使用服務帳戶金鑰產生 Google OAuth2 Access Token (供前端小工具使用)"""
    if not credentials:
        return None
    try:
        auth_req = google.auth.transport.requests.Request()
        credentials.refresh(auth_req)
        return credentials.token
    except Exception as e:
        logger.error(f"❌ 產生 OAuth Token 失敗: {str(e)}", exc_info=True)
        return None


@app.route("/")
def index():
    """首頁（自訂 UI 搜尋介面）"""
    return render_template("index.html")


@app.route("/gcpsearchsop")
@app.route("/gcpsearchsop.html")
@app.route("/sop")
@app.route("/sop.html")
def gcpsearchsop_page():
    """GCP Vertex AI Search 建置與整合 SOP 標準作業程序頁面"""
    return render_template("gcpsearchsop.html")


@app.route("/widget")
@app.route("/widget.html")
def widget_page():
    """Google 官方現成小工具示範頁面 (使用 gen-search-widget)"""
    token = get_oauth_token()
    return render_template("widget.html", token=token, config_id=WIDGET_CONFIG_ID)


@app.route("/api/token", methods=["GET", "POST"])
def get_token():
    """API: 取得或更新供前端 gen-search-widget 使用的短期 OAuth Token"""
    token = get_oauth_token()
    if not token:
        return jsonify({
            "status": "error",
            "message": "無法從 GCP 憑證取得 OAuth Token"
        }), 500

    return jsonify({
        "status": "success",
        "token": token,
        "config_id": WIDGET_CONFIG_ID,
        "expires_in": 3600
    })


@app.route("/api/info", methods=["GET"])
def get_info():
    """取得當前 GCP 連線資訊狀態"""
    return jsonify({
        "status": "ready" if search_client is not None else "unconfigured",
        "project_id": PROJECT_ID,
        "engine_id": ENGINE_ID,
        "data_store_id": DATA_STORE_ID,
        "location": LOCATION,
        "service_account_email": service_account_email,
        "key_file": KEY_FILE_NAME
    })


@app.route("/api/search", methods=["GET", "POST"])
def search():
    """搜尋 API 路由 (支援 GET 與 POST)"""
    if not search_client:
        return jsonify({
            "status": "error",
            "message": "GCP SearchServiceClient 尚未正確初始化，請檢查憑證與設定"
        }), 500

    # 解析查詢參數
    if request.method == "POST":
        data = request.get_json(silent=True) or {}
        query = data.get("query") or data.get("q", "")
        page_size = int(data.get("page_size", 10))
        offset = int(data.get("offset", 0))
    else:
        query = request.args.get("query") or request.args.get("q", "")
        page_size = int(request.args.get("page_size", 10))
        offset = int(request.args.get("offset", 0))

    query = str(query).strip()
    if not query:
        return jsonify({
            "status": "error",
            "message": "搜尋字串不得為空"
        }), 400

    start_time = time.time()

    try:
        serving_config = get_serving_config_path()

        # 建立搜尋請求
        search_request = discoveryengine.SearchRequest(
            serving_config=serving_config,
            query=query,
            page_size=min(max(page_size, 1), 50),
            offset=max(offset, 0),
            # 智慧查詢擴展與拼字檢查
            query_expansion_spec=discoveryengine.SearchRequest.QueryExpansionSpec(
                condition=discoveryengine.SearchRequest.QueryExpansionSpec.Condition.AUTO,
            ),
            spell_correction_spec=discoveryengine.SearchRequest.SpellCorrectionSpec(
                mode=discoveryengine.SearchRequest.SpellCorrectionSpec.Mode.AUTO
            ),
            # 摘要生成 (AI Summary) 與摘要片段 (Snippet)
            content_search_spec=discoveryengine.SearchRequest.ContentSearchSpec(
                summary_spec=discoveryengine.SearchRequest.ContentSearchSpec.SummarySpec(
                    summary_result_count=5,
                    include_citations=True
                ),
                snippet_spec=discoveryengine.SearchRequest.ContentSearchSpec.SnippetSpec(
                    return_snippet=True
                )
            )
        )

        response = search_client.search(search_request)

        # 解析 AI 生成摘要 (若有)
        summary_text = ""
        summary_citations = []
        if hasattr(response, "summary") and response.summary:
            summary_text = getattr(response.summary, "summary_text", "")
            if hasattr(response.summary, "safety_attributes"):
                pass

        # 解析搜尋結果清單
        results = []
        for item in response.results:
            doc = item.document
            derived = dict(doc.derived_struct_data) if doc.derived_struct_data else {}
            
            # 提取標題與高亮
            title = derived.get("title", "")
            html_title = derived.get("htmlTitle", title)

            # 提取網址
            link = derived.get("link", "")
            display_link = derived.get("displayLink", "")
            formatted_url = derived.get("formattedUrl", link)

            # 提取 Snippet 片段
            snippet_text = ""
            html_snippet = ""
            snippets = derived.get("snippets", [])
            if snippets and len(snippets) > 0:
                first_snippet = dict(snippets[0])
                snippet_text = first_snippet.get("snippet", "")
                html_snippet = first_snippet.get("htmlSnippet", snippet_text)

            # 提取縮圖
            pagemap = derived.get("pagemap")
            thumbnail = extract_thumbnail(pagemap)

            results.append({
                "id": doc.id,
                "title": title,
                "html_title": html_title,
                "link": link,
                "display_link": display_link,
                "formatted_url": formatted_url,
                "snippet": snippet_text,
                "html_snippet": html_snippet,
                "thumbnail": thumbnail
            })

        latency_ms = round((time.time() - start_time) * 1000, 2)

        return jsonify({
            "status": "success",
            "query": query,
            "total_size": response.total_size,
            "results": results,
            "summary": summary_text,
            "latency_ms": latency_ms
        })

    except Exception as e:
        logger.error(f"搜尋發生例外錯誤: {str(e)}", exc_info=True)
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    debug = os.getenv("DEBUG", "True").lower() in ("true", "1")
    print(f"\n🚀 GCP Agent Search 服務啟動中: http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=debug)
