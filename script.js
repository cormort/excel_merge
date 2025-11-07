document.addEventListener('DOMContentLoaded', () => {
    const tabContainer = document.querySelector('.tab-container');
    const tabBtns = tabContainer.querySelectorAll('.tab-btn');
    const tabPanes = tabContainer.querySelectorAll('.tab-pane');

    // 預先載入第一個說明頁面
    const initialTab = document.querySelector('.tab-btn[data-source]');
    if (initialTab) {
        loadTabContent(initialTab);
    }
    
    tabContainer.addEventListener('click', async (e) => {
        const clickedBtn = e.target.closest('.tab-btn');
        if (!clickedBtn) return;

        // 如果點擊的頁籤需要從外部載入內容
        if (clickedBtn.dataset.source) {
            await loadTabContent(clickedBtn);
        }

        const tabId = clickedBtn.dataset.tab;
        const targetPane = document.getElementById(tabId);

        if (targetPane) {
            // 移除所有按鈕和面板的 active 狀態
            tabBtns.forEach(btn => btn.classList.remove('active'));
            tabPanes.forEach(pane => pane.classList.remove('active'));

            // 為點擊的按鈕和對應的面板添加 active 狀態
            clickedBtn.classList.add('active');
            targetPane.classList.add('active');
        }
    });

    /**
     * 載入頁籤內容的函數
     * @param {HTMLElement} tabButton - 被點擊的按鈕
     */
    async function loadTabContent(tabButton) {
        const source = tabButton.dataset.source;
        const paneId = tabButton.dataset.tab;
        const targetPane = document.getElementById(paneId);

        // 如果內容尚未載入 (is-loaded 屬性用來避免重複載入)
        if (source && targetPane && !targetPane.hasAttribute('data-is-loaded')) {
            try {
                const response = await fetch(source);
                if (!response.ok) {
                    throw new Error(`無法載入 ${source}: ${response.statusText}`);
                }
                const content = await response.text();
                targetPane.innerHTML = content;
                targetPane.setAttribute('data-is-loaded', 'true'); // 標記為已載入
            } catch (error) {
                console.error('載入內容時發生錯誤:', error);
                targetPane.innerHTML = `<p style="color: red;">內容載入失敗。</p>`;
            }
        }
    }
});
