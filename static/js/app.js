/**
 * GCP Agent Search - Frontend Logic
 */

document.addEventListener("DOMContentLoaded", () => {
  // DOM 元素選取
  const searchForm = document.getElementById("search-form");
  const searchInput = document.getElementById("search-input");
  const clearSearchBtn = document.getElementById("clear-search-btn");
  const searchSubmitBtn = document.getElementById("search-submit-btn");
  const tagChips = document.querySelectorAll(".tag-chip");

  const resultsSection = document.getElementById("results-section");
  const resultsList = document.getElementById("results-list");
  const totalCountEl = document.getElementById("total-count");
  const latencyTimeEl = document.getElementById("latency-time");
  const currentQueryDisplay = document.getElementById("current-query-display");

  const aiSummaryCard = document.getElementById("ai-summary-card");
  const summaryTextContent = document.getElementById("summary-text-content");

  const loadingSection = document.getElementById("loading-section");
  const emptyStateSection = document.getElementById("empty-state-section");
  const errorStateSection = document.getElementById("error-state-section");
  const errorMessageText = document.getElementById("error-message-text");
  const retryBtn = document.getElementById("retry-btn");

  const paginationContainer = document.getElementById("pagination-container");
  const loadMoreBtn = document.getElementById("load-more-btn");

  const gcpStatusChip = document.getElementById("gcp-status-chip");
  const statusIndicator = document.getElementById("status-indicator");
  const statusText = document.getElementById("status-text");
  const statusInfoBtn = document.getElementById("status-info-btn");

  const infoModalOverlay = document.getElementById("info-modal-overlay");
  const modalCloseBtn = document.getElementById("modal-close-btn");
  const modalBodyContent = document.getElementById("modal-body-content");

  const toastNotification = document.getElementById("toast-notification");

  // 狀態管理
  let currentQuery = "";
  let currentOffset = 0;
  const pageSize = 10;
  let totalResults = 0;
  let isSearching = false;
  let gcpConfig = null;

  // 1. 初始化取得 GCP 系統資訊
  fetchGcpInfo();

  async function fetchGcpInfo() {
    try {
      const res = await fetch("/api/info");
      const data = await res.json();
      gcpConfig = data;

      if (data.status === "ready") {
        gcpStatusChip.className = "status-chip online";
        statusText.textContent = `已連線: ${data.project_id}`;
        gcpStatusChip.title = `連線成功！服務帳戶: ${data.service_account_email}`;
      } else {
        gcpStatusChip.className = "status-chip error";
        statusText.textContent = "未配置憑證";
        gcpStatusChip.title = "GCP Client 尚未正確初始化";
      }
    } catch (err) {
      console.error("無法取得 GCP 狀態:", err);
      gcpStatusChip.className = "status-chip error";
      statusText.textContent = "連線失敗";
    }
  }

  // 2. 顯示詳細資訊彈窗
  statusInfoBtn.addEventListener("click", () => {
    if (!gcpConfig) return;
    modalBodyContent.innerHTML = `
      <table class="info-table">
        <tbody>
          <tr>
            <td class="label">狀態</td>
            <td class="value"><strong style="color: ${gcpConfig.status === 'ready' ? '#34a853' : '#ea4335'}">${gcpConfig.status.toUpperCase()}</strong></td>
          </tr>
          <tr>
            <td class="label">專案 ID</td>
            <td class="value"><code>${gcpConfig.project_id}</code></td>
          </tr>
          <tr>
            <td class="label">Engine ID</td>
            <td class="value"><code>${gcpConfig.engine_id || '未設定'}</code></td>
          </tr>
          <tr>
            <td class="label">Data Store</td>
            <td class="value"><code>${gcpConfig.data_store_id || '未設定'}</code></td>
          </tr>
          <tr>
            <td class="label">位置 (Location)</td>
            <td class="value"><code>${gcpConfig.location}</code></td>
          </tr>
          <tr>
            <td class="label">服務帳戶金鑰</td>
            <td class="value"><code>${gcpConfig.key_file}</code></td>
          </tr>
          <tr>
            <td class="label">服務帳戶 Email</td>
            <td class="value"><code>${gcpConfig.service_account_email || '無'}</code></td>
          </tr>
        </tbody>
      </table>
      <p style="margin-top: 14px; font-size: 0.8rem; color: #9ca3af; line-height: 1.5;">
        🔒 <strong>安全設計提示：</strong> GCP 私密金鑰留在 Python 後端，前端僅透過安全 API 呼叫，不會暴露金鑰。
      </p>
    `;
    infoModalOverlay.classList.remove("hidden");
  });

  modalCloseBtn.addEventListener("click", () => {
    infoModalOverlay.classList.add("hidden");
  });

  infoModalOverlay.addEventListener("click", (e) => {
    if (e.target === infoModalOverlay) {
      infoModalOverlay.classList.add("hidden");
    }
  });

  // 3. 輸入框控制與清除按鈕
  searchInput.addEventListener("input", () => {
    if (searchInput.value.trim().length > 0) {
      clearSearchBtn.classList.remove("hidden");
    } else {
      clearSearchBtn.classList.add("hidden");
    }
  });

  clearSearchBtn.addEventListener("click", () => {
    searchInput.value = "";
    clearSearchBtn.classList.add("hidden");
    searchInput.focus();
  });

  // 4. 推薦關鍵字點擊
  tagChips.forEach(chip => {
    chip.addEventListener("click", () => {
      const query = chip.getAttribute("data-query");
      searchInput.value = query;
      clearSearchBtn.classList.remove("hidden");
      triggerSearch(query);
    });
  });

  // 5. 表單提交搜尋
  searchForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const query = searchInput.value.trim();
    if (query) {
      triggerSearch(query);
    }
  });

  retryBtn.addEventListener("click", () => {
    if (currentQuery) {
      triggerSearch(currentQuery);
    }
  });

  // 6. 載入更多結果
  loadMoreBtn.addEventListener("click", () => {
    if (isSearching) return;
    executeSearch(currentQuery, currentOffset, true);
  });

  function triggerSearch(query) {
    currentQuery = query;
    currentOffset = 0;
    executeSearch(query, 0, false);
  }

  // 7. 呼叫後端 API
  async function executeSearch(query, offset = 0, isAppend = false) {
    if (isSearching) return;
    isSearching = true;

    if (!isAppend) {
      // 全新搜尋，重置顯示狀態
      showLoading();
      resultsList.innerHTML = "";
      aiSummaryCard.classList.add("hidden");
      paginationContainer.classList.add("hidden");
    } else {
      loadMoreBtn.disabled = true;
      loadMoreBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> 載入中...';
    }

    try {
      const response = await fetch("/api/search", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          query: query,
          page_size: pageSize,
          offset: offset
        })
      });

      const data = await response.json();

      if (!response.ok || data.status === "error") {
        throw new Error(data.message || "搜尋發生錯誤");
      }

      renderSearchResults(data, isAppend);
    } catch (error) {
      console.error("搜尋失敗:", error);
      showError(error.message);
    } finally {
      isSearching = false;
      if (isAppend) {
        loadMoreBtn.disabled = false;
        loadMoreBtn.innerHTML = '<i class="fa-solid fa-spinner-third"></i> 載入更多結果';
      }
    }
  }

  // 8. 渲染搜尋結果
  function renderSearchResults(data, isAppend = false) {
    hideAllStates();

    const results = data.results || [];
    totalResults = data.total_size || 0;

    if (!isAppend && results.length === 0) {
      showEmpty();
      return;
    }

    // 顯示結果區塊
    resultsSection.classList.remove("hidden");
    totalCountEl.textContent = totalResults.toLocaleString();
    latencyTimeEl.textContent = data.latency_ms || 0;
    currentQueryDisplay.textContent = data.query;

    // 渲染 AI 摘要
    if (!isAppend) {
      if (data.summary && data.summary.trim().length > 0) {
        summaryTextContent.textContent = data.summary;
        aiSummaryCard.classList.remove("hidden");
      } else {
        aiSummaryCard.classList.add("hidden");
      }
    }

    // 渲染每一張結果卡片
    results.forEach(item => {
      const card = createResultCard(item);
      resultsList.appendChild(card);
    });

    // 計算新 offset
    currentOffset += results.length;

    // 判斷是否還有更多結果
    if (currentOffset < totalResults && results.length === pageSize) {
      paginationContainer.classList.remove("hidden");
    } else {
      paginationContainer.classList.add("hidden");
    }
  }

  // 9. 建立單一搜尋卡片 DOM
  function createResultCard(item) {
    const card = document.createElement("article");
    card.className = "result-card";

    // 網址與來源顯示
    const displayUrl = item.formatted_url || item.link || "#";
    const domain = item.display_link || getHostname(item.link);

    // 標題 (優先使用 GCP Discovery Engine 返回的高亮 HTML)
    const titleHtml = item.html_title || escapeHtml(item.title) || "無標題文件";

    // 內文 Snippet
    const snippetHtml = item.html_snippet || escapeHtml(item.snippet) || "尚無摘要內文可供預覽。";

    // 縮圖 HTML
    let thumbnailHtml = "";
    if (item.thumbnail) {
      thumbnailHtml = `
        <div class="result-thumbnail-wrapper">
          <img src="${escapeHtml(item.thumbnail)}" alt="縮圖" class="result-thumbnail" loading="lazy" onerror="this.parentElement.style.display='none'">
        </div>
      `;
    }

    card.innerHTML = `
      <div class="result-main">
        <div class="result-source-meta">
          <i class="fa-solid fa-globe source-icon"></i>
          <span class="source-domain">${escapeHtml(domain)}</span>
          <span class="source-breadcrumb">${escapeHtml(displayUrl)}</span>
        </div>
        
        <h2 class="result-title">
          <a href="${escapeHtml(item.link)}" target="_blank" rel="noopener noreferrer">${titleHtml}</a>
        </h2>
        
        <p class="result-snippet">${snippetHtml}</p>
        
        <div class="result-footer">
          <a href="${escapeHtml(item.link)}" target="_blank" rel="noopener noreferrer" class="card-action-btn">
            <i class="fa-solid fa-arrow-up-right-from-square"></i> 開啟連結
          </a>
          <button type="button" class="card-action-btn copy-link-btn" data-url="${escapeHtml(item.link)}">
            <i class="fa-regular fa-copy"></i> 複製網址
          </button>
        </div>
      </div>
      ${thumbnailHtml}
    `;

    // 綁定複製按鈕
    const copyBtn = card.querySelector(".copy-link-btn");
    copyBtn.addEventListener("click", () => {
      const urlToCopy = copyBtn.getAttribute("data-url");
      navigator.clipboard.writeText(urlToCopy).then(() => {
        showToast("已成功複製連結至剪貼簿！");
      }).catch(() => {
        showToast("複製連結失敗");
      });
    });

    return card;
  }

  // 10. 狀態切換輔助函式
  function showLoading() {
    hideAllStates();
    loadingSection.classList.remove("hidden");
  }

  function showEmpty() {
    hideAllStates();
    emptyStateSection.classList.remove("hidden");
  }

  function showError(msg) {
    hideAllStates();
    errorMessageText.textContent = msg || "發生未預期的搜尋錯誤，請稍後再試。";
    errorStateSection.classList.remove("hidden");
  }

  function hideAllStates() {
    loadingSection.classList.add("hidden");
    emptyStateSection.classList.add("hidden");
    errorStateSection.classList.add("hidden");
  }

  function getHostname(url) {
    try {
      return new URL(url).hostname;
    } catch {
      return "semicontaiwan.org";
    }
  }

  function escapeHtml(str) {
    if (!str) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function showToast(msg) {
    toastNotification.textContent = msg;
    toastNotification.classList.remove("hidden");
    setTimeout(() => {
      toastNotification.classList.add("hidden");
    }, 3000);
  }
});
