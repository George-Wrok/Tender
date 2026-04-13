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
    const categoryFilter = document.getElementById('category-filter');
    const agencyFilter = document.getElementById('agency-filter');
    const vendorFilter = document.getElementById('vendor-filter');

    // Global Data State 
    let currentHistoryData = [];
    let isScraping = false;
    let scrapingAborted = false;

    // Map State
    let map = null;
    let markersGroup = null;
    let userLocationMarker = null;
    let userCoords = null; // {lat, lng}
    let currentView = 'list'; // 'list' | 'map'

    // Google Apps Script Web App API URL
    const API_URL = 'https://script.google.com/macros/s/AKfycbxMextf-2rnS1Hygj2nd18hKwvU5rT_i-qfP7B0Q0xCkq8s8Z1gDO5_jGYF9_t2G0Pp/exec';

    // 頁面載入時先抓取一次歷史資料
    fetchHistory();

    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        // 如果正在抓取中，點擊按鈕代表「停止」
        if (isScraping) {
            scrapingAborted = true;
            setLoadingState(true, 0, true);
            return;
        }

        // 將輸入框內容按換行或空白切開，過濾掉空的字串
        const rawInput = urlInput.value.trim();
        if (!rawInput) {
            showStatus('請輸入有效的網址', 'error');
            return;
        }

        // Q1 先進行初步分析
        const allParsedUrls = rawInput.split(/[\n\r\s]+/).filter(u => u.trim().startsWith('http'));
        const uniqueInputs = [...new Set(allParsedUrls)]; // 移除輸入內容本身重複的部分

        if (uniqueInputs.length === 0) {
            showStatus('找不到有效的網址 (需以 http 開頭)', 'error');
            return;
        }

        // 過濾已存在於歷史紀錄中的網址 (防呆)
        const urlsToProcess = [];
        const duplicates = [];

        uniqueInputs.forEach(url => {
            const isDuplicate = currentHistoryData.some(row => {
                const rowUrl = row['招標網站'] || row['招標網址'] || row['網址'] || row['連結網址'] || '';
                return rowUrl === url;
            });
            if (isDuplicate) {
                duplicates.push(url);
            } else {
                urlsToProcess.push(url);
            }
        });

        if (urlsToProcess.length === 0 && duplicates.length > 0) {
            showStatus(`所有網址 (${duplicates.length} 筆) 皆已抓取過，請勿重複提交。`, 'error');
            return;
        }

        // 若有部分重複，給予提示
        if (duplicates.length > 0) {
            const confirmProcess = confirm(`發現 ${duplicates.length} 筆網址已在紀錄中，將自動過濾。\n剩餘 ${urlsToProcess.length} 筆新網址準備抓取。\n\n是否繼續？`);
            if (!confirmProcess) return;
        }

        // 設置狀態
        isScraping = true;
        scrapingAborted = false;
        setLoadingState(true, urlsToProcess.length);
        hideStatus();

        let successCount = 0;
        let skipCount = duplicates.length; // 預設跳過的筆數
        let errorCount = 0;

        for (let i = 0; i < urlsToProcess.length; i++) {
            // Q2: 檢查是否點擊了停止
            if (scrapingAborted) {
                console.log('使用者中斷抓取');
                break;
            }

            const url = urlsToProcess[i];
            const progress = `[${i + 1}/${urlsToProcess.length}]`;
            
            showStatus(`${progress} 正在處理：${url}`, 'success');

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
                    successCount++;
                    // 每次成功後，更新下方的歷史紀錄總表 (讓使用者看到進度)
                    await fetchHistory();
                } else if (result.status === 'captcha') {
                    errorCount++;
                    showStatus(`${progress} ⚠️ 偵測到驗證碼，請手動處理`, 'error');
                    alert(`【偵測到驗證碼】\n\n網址：${url}\n\n為避免後續全部失敗，建議先完成驗證再繼續。`);
                    if (confirm('是否要立刻開啟該網址進行驗證？')) {
                        window.open(url, '_blank');
                    }
                    break; // 偵測到驗證碼通常建議停止，否則後面都會失敗
                } else {
                    errorCount++;
                    console.error(`${progress} 錯誤：${result.message}`);
                }

            } catch (error) {
                errorCount++;
                console.error(`${progress} API 請求失敗:`, error);
            }

            // 如果還有下一筆，且尚未中斷，等待 10 秒
            if (i < urlsToProcess.length - 1 && !scrapingAborted) {
                for (let seconds = 10; seconds > 0; seconds--) {
                    if (scrapingAborted) break;
                    showStatus(`${progress} 成功，等待 ${seconds} 秒後處理下一筆...`, 'success');
                    await new Promise(resolve => setTimeout(resolve, 1000));
                }
            }
        }

        // 結束後的最終狀態
        isScraping = false;
        setLoadingState(false);

        if (scrapingAborted) {
            showStatus(`抓取已中斷。目前進度: 成功 ${successCount} 筆, 失敗 ${errorCount} 筆。`, 'error');
        } else {
            showStatus(`處理完成！成功: ${successCount} 筆, 跳過重複: ${skipCount} 筆, 失敗: ${errorCount} 筆。`, successCount > 0 ? 'success' : 'error');
        }
        
        if (successCount > 0) {
            urlInput.value = ''; // 清空輸入框
        }
    });

    // --- 地圖與切換按鈕事件 ---
    const viewListBtn = document.getElementById('view-list-btn');
    const viewMapBtn = document.getElementById('view-map-btn');
    const mapWrapper = document.getElementById('map-wrapper');
    const historyWrapper = document.querySelector('.history-wrapper');
    const mapControls = document.getElementById('map-controls');
    const listControls = document.querySelectorAll('.list-controls');
    const getLocationBtn = document.getElementById('get-location-btn');
    const distanceFilter = document.getElementById('distance-filter');

    viewListBtn.addEventListener('click', () => switchView('list'));
    viewMapBtn.addEventListener('click', () => switchView('map'));
    getLocationBtn.addEventListener('click', getUserLocation);
    distanceFilter.addEventListener('change', () => applySortingAndRender());

    function switchView(view) {
        currentView = view;
        if (view === 'list') {
            viewListBtn.classList.add('active');
            viewMapBtn.classList.remove('active');
            mapWrapper.classList.add('hidden');
            historyWrapper.classList.remove('hidden');
            mapControls.classList.add('hidden');
            listControls.forEach(el => el.classList.remove('hidden'));
        } else {
            viewListBtn.classList.remove('active');
            viewMapBtn.classList.add('active');
            mapWrapper.classList.remove('hidden');
            historyWrapper.classList.add('hidden');
            mapControls.classList.remove('hidden');
            listControls.forEach(el => el.classList.add('hidden'));
            
            // 延遲初始化地圖確保容器尺寸正確
            if (!map) {
                setTimeout(initMap, 100);
            } else {
                map.invalidateSize();
            }
        }
    }

    function initMap() {
        if (map) return;
        
        // 初始化地圖 (中心點設為台灣)
        map = L.map('map').setView([23.6, 121], 7);
        
        // 使用 CartoDB Dark Matter 樣式地圖 (符合深色主題)
        L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', {
            attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
            subdomains: 'abcd',
            maxZoom: 20
        }).addTo(map);

        markersGroup = L.layerGroup().addTo(map);
        
        // 渲染現有資料
        applySortingAndRender();
    }

    function renderMapMarkers(dataArray) {
        if (!map || !markersGroup) return;
        markersGroup.clearLayers();

        const agencyIcon = L.divIcon({
            className: 'custom-div-icon',
            html: "<div style='background-color:#ef4444; width:12px; height:12px; border-radius:50%; border:2px solid white;'></div>",
            iconSize: [12, 12],
            iconAnchor: [6, 6]
        });

        const vendorIcon = L.divIcon({
            className: 'custom-div-icon',
            html: "<div style='background-color:#10b981; width:12px; height:12px; border-radius:50%; border:2px solid white;'></div>",
            iconSize: [12, 12],
            iconAnchor: [6, 6]
        });

        dataArray.forEach(row => {
            const tenderName = row['標案名稱'] || '未命名標案';
            const tenderUrl = row['招標網站'] || row['招標網址'] || row['網址'] || '';
            const popupContent = `
                <div style="min-width:150px">
                    <strong>${tenderName}</strong><br>
                    <small>公告日: ${row['公告日'] || '-'}</small><hr style="margin:5px 0; opacity:0.2">
                    機關: ${row['機關名稱'] || '-'}<br>
                    得標廠商: ${row['得標廠商'] || '-'}<br>
                    <a href="${tenderUrl}" target="_blank" style="color:#3b82f6; text-decoration:none; font-weight:bold">開啟網址 ↗</a>
                </div>
            `;

            // 機關標記
            const aLat = parseFloat(String(row['機關地址緯度'] || '').trim());
            const aLng = parseFloat(String(row['機關地址經度'] || '').trim());
            
            if (!isNaN(aLat) && !isNaN(aLng) && aLat !== 0) {
                console.log(`正在標記機關: ${row['機關名稱']}, 座標: ${aLat}, ${aLng}`);
                L.marker([aLat, aLng], {icon: agencyIcon})
                    .bindPopup(popupContent + `<br><span style="color:#94a3b8">📍 機關地址: ${row['機關地址']}</span>`)
                    .addTo(markersGroup);
            }

            // 廠商標記
            const vLat = parseFloat(String(row['廠商地址緯度'] || '').trim());
            const vLng = parseFloat(String(row['廠商地址經度'] || '').trim());
            
            if (!isNaN(vLat) && !isNaN(vLng) && vLat !== 0) {
                console.log(`正在標記廠商: ${row['得標廠商']}, 座標: ${vLat}, ${vLng}`);
                L.marker([vLat, vLng], {icon: vendorIcon})
                    .bindPopup(popupContent + `<br><span style="color:#94a3b8">📍 廠商地址: ${row['廠商地址']}</span>`)
                    .addTo(markersGroup);
            }
        });

        // 如果只有少數點，自動調整縮放
        if (dataArray.length > 0 && dataArray.length < 50) {
            const allCoords = [];
            markersGroup.eachLayer(marker => allCoords.push(marker.getLatLng()));
            if (allCoords.length > 0) {
                const bounds = L.latLngBounds(allCoords);
                map.fitBounds(bounds, {padding: [50, 50]});
            }
        }
    }

    function getUserLocation() {
        if (!navigator.geolocation) {
            alert('您的瀏覽器不支援定位功能');
            return;
        }

        getLocationBtn.textContent = '定位中...';
        navigator.geolocation.getCurrentPosition((position) => {
            userCoords = {
                lat: position.coords.latitude,
                lng: position.coords.longitude
            };

            if (userLocationMarker) map.removeLayer(userLocationMarker);

            userLocationMarker = L.circleMarker([userCoords.lat, userCoords.lng], {
                radius: 8,
                fillColor: "#3b82f6",
                color: "#fff",
                weight: 2,
                opacity: 1,
                fillOpacity: 0.8
            }).addTo(map).bindPopup("<b>我在這裡</b>").openPopup();

            map.setView([userCoords.lat, userCoords.lng], 13);
            getLocationBtn.textContent = '📍 重新定位';
            
            // 觸發重新篩選 (因為 userCoords 改變了)
            applySortingAndRender();
            
        }, (error) => {
            console.error('定位失敗', error);
            alert('無法獲取您的位置，請檢查權限設定。');
            getLocationBtn.textContent = '📍 抓取我的定位';
        });
    }

    // 計算兩個經緯度之間的距離 (km)
    function getDistance(lat1, lon1, lat2, lon2) {
        const R = 6371; // 地球半徑
        const dLat = (lat2 - lat1) * Math.PI / 180;
        const dLon = (lon2 - lon1) * Math.PI / 180;
        const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
                  Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
                  Math.sin(dLon/2) * Math.sin(dLon/2);
        const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
        return R * c;
    }

    function setLoadingState(isLoading, totalCount = 0, isStopping = false) {
        if (isLoading) {
            // submitBtn.disabled = true; // 不再禁用，讓它可以被點擊觸發停止
            submitBtn.classList.add('btn-stop');
            
            if (isStopping) {
                submitBtn.disabled = true;
                btnText.textContent = '中斷中，請稍候...';
            } else {
                btnText.textContent = totalCount > 1 ? `停止批量抓取 (共 ${totalCount} 筆)...` : '停止抓取...';
            }
            btnLoader.classList.remove('loader-hidden');
        } else {
            submitBtn.disabled = false;
            submitBtn.classList.remove('btn-stop');
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

    // 負責處理篩選、排序並呼叫 render
    function applySortingAndRender() {
        if (currentHistoryData.length === 0) {
            renderHistoryTable([]);
            return;
        }

        // 1. 執行篩選 (Filter)
        const categoryVal = categoryFilter ? categoryFilter.value : 'all';
        const agencyVal = agencyFilter ? agencyFilter.value : 'all';
        const vendorVal = vendorFilter ? vendorFilter.value : 'all';

        let filteredData = currentHistoryData.filter(row => {
            let categoryMatch = true;
            let agencyMatch = true;
            let vendorMatch = true;

            const tName = row['標案名稱'] || '';
            const aAddr = row['機關地址'] || '';
            const vAddr = row['廠商地址'] || '';

            // 類別篩選 (關鍵字比對標案名稱)
            if (categoryVal !== 'all') {
                if (categoryVal === 'reinforce') {
                    categoryMatch = tName.includes('補強');
                } else if (categoryVal === 'road') {
                    categoryMatch = tName.includes('道路');
                } else if (categoryVal === 'repair') {
                    categoryMatch = tName.includes('修復');
                } else if (categoryVal === 'waterproof') {
                    categoryMatch = tName.includes('防水');
                } else if (categoryVal === 'other') {
                    categoryMatch = !(tName.includes('補強') || tName.includes('道路') || tName.includes('修復') || tName.includes('防水'));
                }
            }

            // 機關篩選
            if (agencyVal !== 'all') {
                if (agencyVal === 'taipei') {
                    agencyMatch = aAddr.includes('台北') || aAddr.includes('臺北');
                    // 排除新北市
                    if (aAddr.includes('新北')) agencyMatch = false;
                } else if (agencyVal === 'new_taipei') {
                    agencyMatch = aAddr.includes('新北');
                } else if (agencyVal === 'taoyuan') {
                    agencyMatch = aAddr.includes('桃園');
                } else if (agencyVal === 'hsinchu') {
                    agencyMatch = aAddr.includes('新竹');
                } else if (agencyVal === 'other') {
                    agencyMatch = !(aAddr.includes('台北') || aAddr.includes('臺北') || aAddr.includes('新北') || aAddr.includes('桃園') || aAddr.includes('新竹'));
                }
            }

            // 廠商篩選
            if (vendorVal !== 'all') {
                if (vendorVal === 'taipei') {
                    vendorMatch = vAddr.includes('台北') || vAddr.includes('臺北');
                    // 排除新北市
                    if (vAddr.includes('新北')) vendorMatch = false;
                } else if (vendorVal === 'new_taipei') {
                    vendorMatch = vAddr.includes('新北');
                } else if (vendorVal === 'taoyuan') {
                    vendorMatch = vAddr.includes('桃園');
                } else if (vendorVal === 'hsinchu') {
                    vendorMatch = vAddr.includes('新竹');
                } else if (vendorVal === 'other') {
                    vendorMatch = !(vAddr.includes('台北') || vAddr.includes('臺北') || vAddr.includes('新北') || vAddr.includes('桃園') || vAddr.includes('新竹'));
                }
            }

            return categoryMatch && agencyMatch && vendorMatch;
        });

        // 2. 執行距離篩選 (如果用戶有定位)
        const maxDist = distanceFilter ? distanceFilter.value : 'all';
        if (maxDist !== 'all' && userCoords) {
            filteredData = filteredData.filter(row => {
                const aLat = parseFloat(row['機關地址緯度']);
                const aLng = parseFloat(row['機關地址經度']);
                const vLat = parseFloat(row['廠商地址緯度']);
                const vLng = parseFloat(row['廠商地址經度']);
                
                let aDist = Infinity, vDist = Infinity;
                if (!isNaN(aLat)) aDist = getDistance(userCoords.lat, userCoords.lng, aLat, aLng);
                if (!isNaN(vLat)) vDist = getDistance(userCoords.lat, userCoords.lng, vLat, vLng);
                
                return aDist <= parseFloat(maxDist) || vDist <= parseFloat(maxDist);
            });
        }

        const sortValue = sortSelect.value;
        let sortedData = [...filteredData]; // 使用過濾後的資料進行排序

        // 若為 default 則不另外 sort，因為 records.reverse() 在後端已將最新排前面 (也就是抓取時間最新)
        if (sortValue !== 'default') {
            sortedData.sort((a, b) => {
                if (sortValue === 'post_date') {
                    // 公告日 (預設：最舊到最新 -> 遞增)
                    let dateA = parseROCDate(a['公告日'] || a['公告日期']);
                    let dateB = parseROCDate(b['公告日'] || b['公告日期']);
                    return dateA - dateB;
                } else if (sortValue === 'start_date') {
                    // 開始日 (預設：最舊到最新 -> 遞增)
                    let dateA = parseROCDate(a['開始日'] || a['履約起迄日期'] || a['履約日期']);
                    let dateB = parseROCDate(b['開始日'] || b['履約起迄日期'] || b['履約日期']);
                    return dateA - dateB;
                } else if (sortValue === 'end_date') {
                    // 結束日 (預設：最舊到最新 -> 遞增)
                    let dateA = parseROCDate(a['結束日'] || a['履約起迄日期'] || a['履約日期']);
                    let dateB = parseROCDate(b['結束日'] || b['履約起迄日期'] || b['履約日期']);
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
        }

        renderHistoryTable(sortedData);
        if (map) renderMapMarkers(sortedData);
    }

    // 監聽選單變化
    if (sortSelect) {
        sortSelect.addEventListener('change', () => {
            applySortingAndRender();
        });
    }

    if (categoryFilter) {
        categoryFilter.addEventListener('change', () => {
            applySortingAndRender();
        });
    }

    if (agencyFilter) {
        agencyFilter.addEventListener('change', () => {
            applySortingAndRender();
        });
    }

    if (vendorFilter) {
        vendorFilter.addEventListener('change', () => {
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
        const allHeaders = Object.keys(dataArray[0]);
        // ★ 新增邏輯：隱藏網址類的欄位，因為要合併到標案名稱中
        const urlFieldNames = ['招標網站', '招標網址', '網址', '連結網址', '原網址'];
        const headers = allHeaders.filter(h => !urlFieldNames.includes(h));

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

            // 判斷是否過期
            const endDateText = row['結束日'] || row['履約起迄日期'] || '';
            const endDate = parseROCDate(endDateText);
            const today = new Date();
            today.setHours(0, 0, 0, 0); // 只比較日期，不比較時間
            
            const isExpired = endDate.getTime() !== 0 && endDate < today;

            tbodyHtml += `<tr class="${isExpired ? 'row-expired' : ''}">`;
            headers.forEach(header => {
                let text = row[header] || '';
                const tenderUrl = row['招標網站'] || row['招標網址'] || row['網址'] || row['連結網址'] || '';

                // 處理特殊欄位顯示
                if (header === '結束') {
                    // 若是結束欄位，且值為 FALSE，則顯示一個可點擊的核取方塊
                    tbodyHtml += `<td class="text-center">
                        <input type="checkbox" class="finish-checkbox" data-url="${tenderUrl}" title="標記為已結束">
                    </td>`;
                } else if (header === '接觸') {
                    const val = typeof text === 'string' ? text.trim() : '';
                    tbodyHtml += `<td>
                        <select class="contact-select custom-select" data-url="${tenderUrl}" style="padding: 0.3rem 1.8rem 0.3rem 0.8rem; min-width: 40px; white-space: nowrap;">
                            <option value="">-</option>
                            <option value="凱" ${val === '凱' ? 'selected' : ''}>凱</option>
                            <option value="娟" ${val === '娟' ? 'selected' : ''}>娟</option>
                            <option value="喬" ${val === '喬' ? 'selected' : ''}>喬</option>
                        </select>
                    </td>`;
                } else if (header === '標案名稱') {
                    // ★ 新增邏輯：將網址點入標案名稱中
                    text = `<a href="${tenderUrl}" target="_blank" class="tender-link">${text}</a>`;
                    tbodyHtml += `<td>${text}</td>`;
                } else {
                    // 網址若為連結，可加上 a tag 處理 (雖然現在隱藏了，但保留邏輯以防萬一)
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

        // 重新綁定 Checkbox 與 Select 的點擊事件
        bindInteractiveEvents();
    }

    // 綁定動態元素的事件監聽器
    function bindInteractiveEvents() {
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

        const contactSelects = document.querySelectorAll('.contact-select');
        contactSelects.forEach(select => {
            select.addEventListener('change', async function () {
                const urlToUpdate = this.getAttribute('data-url');
                const contactValue = this.value;

                if (!urlToUpdate) {
                    alert('無法取得該筆標案網址，無法更新。');
                    return;
                }

                try {
                    // 禁用該選單避免重複操作
                    this.disabled = true;
                    showStatus('正在為您更新接觸人員至雲端試算表...', 'success');

                    const response = await fetch(API_URL, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/x-www-form-urlencoded',
                        },
                        body: new URLSearchParams({
                            'action': 'update_contact',
                            'url': urlToUpdate,
                            'contact': contactValue
                        })
                    });

                    const result = await response.json();

                    if (result.status === 'success') {
                        showStatus('接觸人員更新成功！', 'success');
                        // 成功後自動重抓資料，確保畫面與雲端一致
                        await fetchHistory();
                    } else {
                        alert(`更新失敗: ${result.message}`);
                        this.disabled = false;
                    }
                } catch (error) {
                    console.error('更新接觸失敗:', error);
                    alert('連線失敗，請稍後再試。');
                    this.disabled = false;
                }
            });
        });
    }
});
