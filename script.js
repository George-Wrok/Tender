document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('scraper-form');
    const urlInput = document.getElementById('tender-url');
    const submitBtn = document.getElementById('submit-btn');
    const btnText = submitBtn.querySelector('.btn-text');
    const btnLoader = submitBtn.querySelector('.btn-loader');
    const statusBox = document.getElementById('status-message');

    // History Table Elements
    const historyHead = document.getElementById('history-head');
    const historyBody = document.getElementById('history-body');
    const sortSelect = document.getElementById('sort-select');

    // Global Data State 
    let currentHistoryData = [];

    // Google Apps Script Web App API URL
    const API_URL = 'https://script.google.com/macros/s/AKfycbxMextf-2rnS1Hygj2nd18hKwvU5rT_i-qfP7B0Q0xCkq8s8Z1gDO5_jGYF9_t2G0Pp/exec';

    // 頁面載入時先抓取一次歷史資料
    fetchHistory();

    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        const url = urlInput.value.trim();

        if (!url) {
            showStatus('請輸入有效的網址', 'error');
            return;
        }

        // 設置按鈕為載入狀態
        setLoadingState(true);
        hideStatus();

        try {
            // 發送請求到 Google Apps Script
            const response = await fetch(API_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    'url': url
                })
            });

            const result = await response.json();

            if (result.status === 'success') {
                showStatus('成功！資料已匯入您的試算表。', 'success');
                urlInput.value = ''; // 清空輸入框

                // 成功後，自動更新下方的歷史紀錄總表
                await fetchHistory();

            } else {
                showStatus(`錯誤：${result.message || '發生未知錯誤'}`, 'error');
            }

        } catch (error) {
            console.error('API 請求失敗:', error);

            // 若為跨域問題，提供備用說明
            if (error instanceof TypeError && error.message.includes('fetch')) {
                showStatus('無法連線至 API。可能是跨域資源共用 (CORS) 問題，請確認 GAS 腳本已重新部署為「所有人」皆可存取。', 'error');
            } else {
                showStatus(`錯誤：連線失敗或伺服器無回應`, 'error');
            }
        } finally {
            // 恢復按鈕狀態
            setLoadingState(false);
        }
    });

    function setLoadingState(isLoading) {
        if (isLoading) {
            submitBtn.disabled = true;
            btnText.textContent = '抓取中...';
            btnLoader.classList.remove('loader-hidden');
        } else {
            submitBtn.disabled = false;
            btnText.textContent = '一鍵抓取並寫入';
            btnLoader.classList.add('loader-hidden');
        }
    }

    function showStatus(message, type) {
        statusBox.textContent = message;
        statusBox.className = `status-box status-${type}`;
        statusBox.classList.remove('hidden');
    }

    function hideStatus() {
        statusBox.classList.add('hidden');
    }

    // 將 Google 試算表資料拉回來的邏輯
    async function fetchHistory() {
        const CACHE_KEY = 'tender_history_cache';

        // 1. 先嘗試從 LocalStorage 讀取快取，若有資料就立刻顯示 (達到秒開效果)
        const cachedData = localStorage.getItem(CACHE_KEY);
        if (cachedData) {
            try {
                const parsedData = JSON.parse(cachedData);
                if (parsedData && parsedData.length > 0) {
                    currentHistoryData = parsedData;
                    applySortingAndRender();

                    // 在標題旁加上一個小提示，表示這是快取資料，正在背景更新中
                    historyHead.parentElement.caption = createLoadingCaption();
                }
            } catch (e) {
                console.warn('快取解析失敗', e);
            }
        }

        try {
            // 2. 背景向 Google Apps Script 發送 GET 請求獲取最新資料
            const response = await fetch(API_URL);
            const result = await response.json();

            // 移除載入中的提示
            if (historyHead.parentElement.caption) {
                historyHead.parentElement.deleteCaption();
            }

            if (result.status === 'success' && result.data && result.data.length > 0) {
                // 將最新資料存入快取與全域狀態
                localStorage.setItem(CACHE_KEY, JSON.stringify(result.data));
                currentHistoryData = result.data;

                // 依據當前排序重繪畫面
                applySortingAndRender();
            } else if (!cachedData) {
                // 只有在沒有快取的情況下，才顯示「尚無資料」
                historyBody.innerHTML = '<tr><td colspan="100%" class="text-center loading-text">目前尚無任何資料</td></tr>';
            }
        } catch (error) {
            console.error('抓取歷史紀錄失敗:', error);

            // 移除載入中的提示
            if (historyHead.parentElement.caption) {
                historyHead.parentElement.deleteCaption();
            }

            if (!cachedData) {
                // 如果連快取都沒有，才顯示錯誤訊息
                historyBody.innerHTML = '<tr><td colspan="100%" class="text-center loading-text" style="color:var(--error-color)">無法載入雲端資料，請確認網路連線或 GAS 部署是否正確。</td></tr>';
            } else {
                // 如果有快取，可以選擇悄悄印出錯誤，或者顯示一個小提示
                console.warn('無法連線取得最新資料，目前顯示為快取內容。');
            }
        }
    }

    function createLoadingCaption() {
        // ... (原樣保留)
        const caption = document.createElement('caption');
        caption.innerHTML = '<span style="font-size: 0.8rem; color: var(--text-secondary);">正在背景同步最新資料...</span>';
        caption.style.textAlign = 'right';
        caption.style.marginBottom = '5px';
        return caption;
    }

    // 將民國年轉西元年 Date (例如: 115/01/21 -> Date object)
    function parseROCDate(dateStr) {
        if (!dateStr) return new Date(0); // 最舊的日期做 fallback
        // 日期常見格式：115/01/22 或是 115/01/21 - 115/02/19(預估)
        // 使用正則匹配出最前面的民國年/月/日
        const match = dateStr.match(/(\d{3})\/(\d{2})\/(\d{2})/);
        if (match) {
            const year = parseInt(match[1], 10) + 1911;
            const month = parseInt(match[2], 10) - 1;
            const day = parseInt(match[3], 10);
            return new Date(year, month, day);
        }
        return new Date(0);
    }

    // 負責處理排序並呼叫 render
    function applySortingAndRender() {
        if (currentHistoryData.length === 0) {
            renderHistoryTable([]);
            return;
        }

        const sortValue = sortSelect.value;
        let sortedData = [...currentHistoryData]; // 複製一份，避免改到原始順序

        // 處理 4 個精簡後的排序選項
        sortedData.sort((a, b) => {
            if (sortValue === 'post_date') {
                // 公告日 (預設：最舊到最新 -> 遞增)
                let dateA = parseROCDate(a['公告日'] || a['公告日期']);
                let dateB = parseROCDate(b['公告日'] || b['公告日期']);
                return dateA - dateB;
            } else if (sortValue === 'contract_date') {
                // 履約日期 (預設：最舊到最新 -> 遞增)
                let dateA = parseROCDate(a['履約起迄日期'] || a['履約日期']);
                let dateB = parseROCDate(b['履約起迄日期'] || b['履約日期']);
                return dateA - dateB;
            } else if (sortValue === 'agency_address') {
                // 機關地址 (預設：郵遞區號數字小到大 -> 遞增)
                let aMatch = (a['機關地址'] || '').match(/^\D*(\d{1,3})/);
                let bMatch = (b['機關地址'] || '').match(/^\D*(\d{1,3})/);
                let valA = aMatch ? parseInt(aMatch[1], 10) : 9999;
                let valB = bMatch ? parseInt(bMatch[1], 10) : 9999;
                return valA - valB;
            } else if (sortValue === 'vendor_address') {
                // 廠商地址 (預設：郵遞區號數字小到大 -> 遞增)
                let aMatch = (a['廠商地址'] || '').match(/^\D*(\d{1,3})/);
                let bMatch = (b['廠商地址'] || '').match(/^\D*(\d{1,3})/);
                let valA = aMatch ? parseInt(aMatch[1], 10) : 9999;
                let valB = bMatch ? parseInt(bMatch[1], 10) : 9999;
                return valA - valB;
            }

            return 0; // fallback
        });

        renderHistoryTable(sortedData);
    }

    // 監聽選單變化
    if (sortSelect) {
        sortSelect.addEventListener('change', () => {
            applySortingAndRender();
        });
    }

    // 渲染歷史紀錄表格
    function renderHistoryTable(dataArray) {
        if (dataArray.length === 0) {
            historyHead.innerHTML = '';
            historyBody.innerHTML = '<tr><td colspan="100%" class="text-center loading-text">目前尚無需要處理的資料</td></tr>';
            return;
        }

        // 第一筆資料的 Keys 當作標題
        const headers = Object.keys(dataArray[0]);

        // 生成 Thead
        let theadHtml = '<tr>';
        headers.forEach(header => {
            theadHtml += `<th>${header}</th>`;
        });
        theadHtml += '</tr>';
        historyHead.innerHTML = theadHtml;

        // 生成 Tbody
        let tbodyHtml = '';
        dataArray.forEach(row => {
            // ★ 新增邏輯：如果在「結束」欄位打勾 (值為 "TRUE" 或 true)，則不要顯示該筆資料
            const isFinished = row['結束'] === 'TRUE' || row['結束'] === true;
            if (isFinished) {
                return; // 跳過這一筆資料 (類似 continue)
            }

            tbodyHtml += '<tr>';
            headers.forEach(header => {
                let text = row[header] || '';

                // 處理特殊欄位顯示
                if (header === '結束') {
                    // 若是結束欄位，且值為 FALSE，則顯示一個可點擊的核取方塊
                    const tenderUrl = row['招標網站'] || row['招標網址'] || row['網址'] || row['連結網址'] || '';
                    tbodyHtml += `<td class="text-center">
                        <input type="checkbox" class="finish-checkbox" data-url="${tenderUrl}" title="標記為已結束">
                    </td>`;
                } else {
                    // 網址若為連結，可加上 a tag 處理
                    if (typeof text === 'string' && text.startsWith('http')) {
                        text = `<a href="${text}" target="_blank" style="color:var(--accent-color)">連結</a>`;
                    }
                    tbodyHtml += `<td>${text}</td>`;
                }
            });
            tbodyHtml += '</tr>';
        });

        // 如果過濾後沒有資料顯示，給予提示
        if (tbodyHtml === '') {
            tbodyHtml = '<tr><td colspan="100%" class="text-center loading-text">目前尚無需要處理的資料</td></tr>';
        }

        historyBody.innerHTML = tbodyHtml;

        // 重新綁定 Checkbox 的點擊事件
        bindCheckboxEvents();
    }

    // 綁定核取方塊的事件監聽器
    function bindCheckboxEvents() {
        const checkboxes = document.querySelectorAll('.finish-checkbox');
        checkboxes.forEach(box => {
            box.addEventListener('change', async function (e) {
                // 先防止預設打勾，等確認後再處理
                e.preventDefault();
                const isChecked = this.checked;
                const urlToUpdate = this.getAttribute('data-url');

                if (!urlToUpdate) {
                    alert('無法取得該筆標案網址，無法更新。');
                    this.checked = false;
                    return;
                }

                if (isChecked) {
                    // 彈出確認視窗防呆
                    const confirmFinish = confirm('您確定要將此標案標記為「已結束」嗎？\n(標記後將不會顯示在此清單中，且會同步至後台試算表)');

                    if (confirmFinish) {
                        try {
                            // 禁用該按鈕避免重複點擊
                            this.disabled = true;
                            // 顯示小提示框表示正在處理
                            showStatus('正在為您更新至雲端試算表...', 'success');

                            const response = await fetch(API_URL, {
                                method: 'POST',
                                headers: {
                                    'Content-Type': 'application/x-www-form-urlencoded',
                                },
                                body: new URLSearchParams({
                                    'action': 'update_status',
                                    'url': urlToUpdate
                                })
                            });

                            const result = await response.json();

                            if (result.status === 'success') {
                                showStatus('更新成功！', 'success');
                                // 成功後自動重抓資料，該筆就會消失了
                                await fetchHistory();
                            } else {
                                alert(`更新失敗: ${result.message}`);
                                this.checked = false;
                                this.disabled = false;
                            }
                        } catch (error) {
                            console.error('更新狀態失敗:', error);
                            alert('連線失敗，請稍後再試。');
                            this.checked = false;
                            this.disabled = false;
                        }
                    } else {
                        // 取消打勾狀態
                        this.checked = false;
                    }
                }
            });
        });
    }
});
