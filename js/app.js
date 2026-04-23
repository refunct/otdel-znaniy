// app.js
import { renderGuidesList, renderSections } from './guides.js';
import { renderTestsList, startTest } from './tests.js';

export const state = {
    guides: [],
    sections: [],
    tests: [],
    questions: [],
    notifications: [],
    dataLoaded: false,
    currentPage: 'guides',
    currentGuideId: null,
    currentTestId: null,
    isOnline: navigator.onLine,
    testInProgress: false
};

export const elements = {
    navTabs: document.getElementById('navTabs'),
    contentContainer: document.getElementById('contentContainer'),
    notificationsContainer: document.getElementById('notificationsContainer'),
    loader: document.getElementById('loader'),
    modalOverlay: document.getElementById('modalOverlay'),
    modalContent: document.getElementById('modalContent')
};

export function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

export function parseMediaList(str) {
    if (!str) return [];
    return str.split(',').map(s => s.trim()).filter(s => s.length > 0);
}

export function getFileName(path) {
    const parts = String(path).split('/');
    return parts[parts.length - 1] || 'Файл';
}

export function formatExcelDate(excelDate) {
    if (!excelDate) return '';
    if (typeof excelDate === 'string') return excelDate;
    const date = new Date((excelDate - 25569) * 86400 * 1000);
    return date.toLocaleDateString('ru-RU');
}

async function loadExcel(filename) {
    const response = await fetch(filename);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const buffer = await response.arrayBuffer();
    const data = new Uint8Array(buffer);
    const workbook = XLSX.read(data, { type: 'array' });
    const result = {};
    workbook.SheetNames.forEach(sheetName => {
        result[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    });
    return result;
}

async function loadData() {
    const timeout = (ms) => new Promise((_, reject) => setTimeout(() => reject(new Error('Timeout')), ms));
    try {
        const [guidesData, testsData] = await Promise.race([
            Promise.all([
                loadExcel('docs/guide.xlsx'),
                loadExcel('docs/tests.xlsx')
            ]),
            timeout(15000)
        ]);
        state.guides = (guidesData.guides || []).map(g => ({ ...g, id: String(g.id) }));
        state.sections = (guidesData.sections || []).map(s => ({ ...s, id: String(s.id), guide_id: String(s.guide_id) }));
        state.tests = (testsData.tests || []).map(t => ({ ...t, id: String(t.id) }));
        state.questions = (testsData.questions || []).map(q => ({ ...q, id: String(q.id), test_id: String(q.test_id) }));
    } catch (error) {
        console.error('Load error:', error);
        showRetryButton();
        return;
    }

    try {
        const notificationsData = await Promise.race([
            loadExcel('docs/notifications.xlsx'),
            timeout(5000)
        ]);
        if (notificationsData && notificationsData.notifications) {
            const closed = JSON.parse(localStorage.getItem('closedNotifications') || '[]');
            state.notifications = notificationsData.notifications
                .map(row => {
                    const normalized = {};
                    Object.keys(row).forEach(key => { normalized[key.trim().toLowerCase()] = row[key]; });
                    return {
                        id: String(normalized.id || ''),
                        message: normalized.message || '',
                        active: normalized.active,
                        type: (normalized.type || 'info').toString().trim().toLowerCase()
                    };
                })
                .filter(n => n.active == 1 && !closed.includes(n.id));
        } else {
            state.notifications = [];
        }
    } catch (err) {
        console.warn('Notifications load skipped:', err);
        state.notifications = [];
    }

    state.dataLoaded = true;
    elements.loader.style.display = 'none';
    renderNotifications();
    handleRouting();
}

function showRetryButton() {
    elements.loader.innerHTML = `
        <div class="empty-state">
            ❌ Не удалось загрузить данные<br>
            <button id="retryLoadBtn" class="nav-test-btn primary" style="margin-top:20px;">Повторить</button>
        </div>
    `;
    document.getElementById('retryLoadBtn').addEventListener('click', () => {
        elements.loader.innerHTML = 'Загрузка данных...';
        elements.loader.style.display = 'block';
        loadData();
    });
}

function parseHash() {
    const hash = window.location.hash.slice(1) || 'guides';
    const parts = hash.split('/');
    return { page: parts[0], id: parts[1] || null };
}

function updateActiveTab(page) {
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.page === page);
    });
}

export function navigateTo(page, id = null) {
    let hash = page;
    if (id) hash += '/' + id;
    window.location.hash = hash;
}

function handleRouting() {
    if (!state.dataLoaded) return;
    const { page, id } = parseHash();
    state.currentPage = page;
    updateActiveTab(page);

    if (page === 'guides') {
        state.currentGuideId = id;
        id ? renderSections() : renderGuidesList();
    } else if (page === 'tests') {
        state.currentTestId = id;
        id ? startTest(id) : renderTestsList();
    }
}

export function showModal(message, onConfirm, onCancel) {
    elements.modalContent.innerHTML = `
        <p>${escapeHtml(message)}</p>
        <div class="modal-buttons">
            <button class="modal-btn cancel" id="modalCancel">Отмена</button>
            <button class="modal-btn confirm" id="modalConfirm">Да</button>
        </div>
    `;
    elements.modalOverlay.style.display = 'flex';
    document.getElementById('modalCancel').onclick = () => {
        elements.modalOverlay.style.display = 'none';
        if (onCancel) onCancel();
    };
    document.getElementById('modalConfirm').onclick = () => {
        elements.modalOverlay.style.display = 'none';
        if (onConfirm) onConfirm();
    };
    elements.modalOverlay.onclick = (e) => {
        if (e.target === elements.modalOverlay) {
            elements.modalOverlay.style.display = 'none';
            if (onCancel) onCancel();
        }
    };
}

// PWA: баннер установки
let deferredPrompt = null;

function setupPWA() {
    const banner = document.getElementById('installBanner');
    if (!banner) return;
    showInstallBanner(banner, false);

    window.addEventListener('beforeinstallprompt', (e) => {
        e.preventDefault();
        deferredPrompt = e;
    });

    window.addEventListener('appinstalled', () => {
        showInstallBanner(banner, true);
        deferredPrompt = null;
    });
}

function showInstallBanner(banner, installed) {
    banner.style.display = 'block';
    banner.className = 'install-banner';
    if (installed) {
        banner.innerHTML = `
            <div class="install-banner-message">✅ Приложение установлено. Спасибо!</div>
            <button class="install-banner-btn" id="installBannerBtn" disabled>Готово</button>
        `;
    } else {
        banner.innerHTML = `
            <div class="install-banner-message">📱 Установите приложение, чтобы пользоваться справочником даже без интернета.</div>
            <button class="install-banner-btn" id="installBannerBtn">Установить</button>
        `;
    }
}

function openLightbox(src) {
    elements.modalContent.innerHTML = `
        <div style="position:relative; max-width:90vw; max-height:90vh;">
            <img src="${escapeHtml(src)}" style="width:100%; height:auto; border-radius:var(--radius);">
            <button class="lightbox-close" style="position:absolute; top:-15px; right:-15px; background:var(--primary); color:white; border:none; border-radius:50%; width:40px; height:40px; cursor:pointer; font-size:24px; line-height:1; box-shadow:var(--shadow);">×</button>
        </div>
    `;
    elements.modalOverlay.style.display = 'flex';
    const closeBtn = elements.modalContent.querySelector('.lightbox-close');
    if (closeBtn) closeBtn.onclick = () => elements.modalOverlay.style.display = 'none';
    elements.modalOverlay.onclick = (e) => {
        if (e.target === elements.modalOverlay) elements.modalOverlay.style.display = 'none';
    };
}

function renderNotifications() {
    const container = elements.notificationsContainer;
    if (!container) return;

    const installBanner = document.getElementById('installBanner');
    const bannerHTML = installBanner ? installBanner.outerHTML : '';

    let html = bannerHTML;
    state.notifications.forEach(n => {
        const type = String(n.type || '').trim().toLowerCase();
        let bg = '#d1ecf1', border = '#0c5460', color = '#0c5460';
        if (type === 'warning') { bg = '#fff3cd'; border = '#856404'; color = '#856404'; }
        else if (type === 'error') { bg = '#f8d7da'; border = '#721c24'; color = '#721c24'; }
        html += `
            <div class="notification-item" data-id="${n.id}" style="background:${bg}; border-left:4px solid ${border}; color:${color}; margin-bottom:8px; padding:12px 16px; border-radius:0 var(--radius-sm) var(--radius-sm) 0; display:flex; align-items:center; justify-content:space-between; transition:all 0.3s ease;">
                <div class="notification-message" style="flex:1;">${n.message}</div>
                <button class="notification-close" style="background:none; border:none; font-size:1.5rem; cursor:pointer; color:${color}; padding:0 0 0 12px; line-height:1;">&times;</button>
            </div>
        `;
    });
    container.innerHTML = html;
    container.style.display = 'block';
}

function closeNotification(id) {
    const closed = JSON.parse(localStorage.getItem('closedNotifications') || '[]');
    if (!closed.includes(id)) {
        closed.push(id);
        localStorage.setItem('closedNotifications', JSON.stringify(closed));
    }
    state.notifications = state.notifications.filter(n => n.id !== id);
    const item = document.querySelector(`.notification-item[data-id="${id}"]`);
    if (item) {
        item.style.opacity = '0';
        item.style.maxHeight = '0';
        item.style.padding = '0 16px';
        item.style.marginBottom = '0';
        item.style.borderLeftWidth = '0';
        setTimeout(renderNotifications, 300);
    } else renderNotifications();
}

async function init() {
    setupPWA();
    window.addEventListener('hashchange', handleRouting);
    elements.navTabs.addEventListener('click', (e) => {
        const btn = e.target.closest('.nav-btn');
        if (!btn) return;
        const page = btn.dataset.page;
        if (page) navigateTo(page);
    });
    document.addEventListener('click', (e) => {
        const img = e.target.closest('[data-lightbox]');
        if (!img) return;
        e.preventDefault();
        openLightbox(img.src);
    });

    if (elements.notificationsContainer) {
        elements.notificationsContainer.addEventListener('click', (e) => {
            const closeBtn = e.target.closest('.notification-close');
            if (closeBtn) {
                const item = closeBtn.closest('.notification-item');
                if (item) closeNotification(item.dataset.id);
                return;
            }
            if (e.target.id === 'installBannerBtn') {
                if (deferredPrompt) {
                    deferredPrompt.prompt();
                    deferredPrompt.userChoice.then(() => {
                        deferredPrompt = null;
                        const banner = document.getElementById('installBanner');
                        if (banner) showInstallBanner(banner, false);
                    });
                } else {
                    alert('Чтобы установить приложение, используйте меню браузера: "Добавить на главный экран" или "Установить приложение".');
                }
            }
        });
    }

    await loadData();
}

init();