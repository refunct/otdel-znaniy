// app.js
import { renderGuidesList, renderSections } from './guides.js';
import { renderTestsList, startTest } from './tests.js';

// Глобальное состояние
export const state = {
    guides: [],
    sections: [],
    tests: [],
    questions: [],
    notifications: [],   // новый массив
    dataLoaded: false,
    currentPage: 'guides',
    currentGuideId: null,
    currentTestId: null,
    isOnline: navigator.onLine,
    testInProgress: false
};
// DOM элементы
export const elements = {
    navTabs: document.getElementById('navTabs'),
    contentContainer: document.getElementById('contentContainer'),
    notificationsContainer: document.getElementById('notificationsContainer'), // новая строка
    loader: document.getElementById('loader'),
    modalOverlay: document.getElementById('modalOverlay'),
    modalContent: document.getElementById('modalContent')
};

// Вспомогательные функции
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

// Загрузка Excel через fetch
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

// Загрузка всех данных с таймаутом и обработкой ошибок
async function loadData() {
    const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error('Timeout')), 15000)
    );
    try {
        const [guidesData, testsData, notificationsData] = await Promise.race([
            Promise.all([
                loadExcel('docs/guide.xlsx'),
                loadExcel('docs/tests.xlsx'),
                loadExcel('docs/notifications.xlsx').catch(() => null) // игнорируем ошибку
            ]),
            timeoutPromise
        ]);

        // ... существующая обработка guides и tests ...

        // Обработка уведомлений
        if (notificationsData) {
            state.notifications = (notificationsData.notifications || [])
                .map(n => ({ ...n, id: String(n.id) }))
                .filter(n => {
                    // Только активные (active == 1) и не скрытые ранее
                    if (n.active !== 1) return false;
                    const closed = JSON.parse(localStorage.getItem('closedNotifications') || '[]');
                    return !closed.includes(n.id);
                });
        } else {
            state.notifications = [];
        }

        state.dataLoaded = true;
        elements.loader.style.display = 'none';
        renderNotifications(); // показываем уведомления после загрузки
        handleRouting();
    } catch (error) {
        console.error('Load error:', error);
        showRetryButton();
    }
}

function renderNotifications() {
    const container = elements.notificationsContainer;
    if (!container) return;

    const notifications = state.notifications;
    if (!notifications.length) {
        container.innerHTML = '';
        container.style.display = 'none';
        return;
    }

    let html = '';
    notifications.forEach(n => {
        let bgColor = '#d1ecf1', borderColor = '#0c5460', textColor = '#0c5460'; // info
        if (n.type === 'warning') {
            bgColor = '#fff3cd'; borderColor = '#856404'; textColor = '#856404';
        } else if (n.type === 'error') {
            bgColor = '#f8d7da'; borderColor = '#721c24'; textColor = '#721c24';
        }
        html += `
            <div class="notification-item" data-id="${n.id}" style="background:${bgColor}; border-left:4px solid ${borderColor}; color:${textColor}; margin-bottom:8px; padding:12px 16px; border-radius:0 var(--radius-sm) var(--radius-sm) 0; display:flex; align-items:center; justify-content:space-between; opacity:1; max-height:200px; transition:all 0.3s ease;">
                <div class="notification-message" style="flex:1;">${n.message}</div>
                <button class="notification-close" style="background:none; border:none; font-size:1.5rem; cursor:pointer; color:${textColor}; padding:0 0 0 12px; line-height:1;">&times;</button>
            </div>
        `;
    });

    container.innerHTML = html;
    container.style.display = 'block';
}

function closeNotification(id) {
    // Сохраняем ID в localStorage
    const closed = JSON.parse(localStorage.getItem('closedNotifications') || '[]');
    if (!closed.includes(id)) {
        closed.push(id);
        localStorage.setItem('closedNotifications', JSON.stringify(closed));
    }
    // Удаляем из state.notifications
    state.notifications = state.notifications.filter(n => n.id !== id);
    // Плавно скрываем плашку
    const item = document.querySelector(`.notification-item[data-id="${id}"]`);
    if (item) {
        item.style.opacity = '0';
        item.style.maxHeight = '0';
        item.style.padding = '0 16px';
        item.style.marginBottom = '0';
        item.style.borderLeftWidth = '0';
        setTimeout(() => {
            renderNotifications(); // обновим список (удалит элемент или скроет контейнер)
        }, 300);
    } else {
        renderNotifications();
    }
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

// Роутинг
function parseHash() {
    const hash = window.location.hash.slice(1) || 'guides';
    const parts = hash.split('/');
    const page = parts[0];
    const id = parts[1] || null;
    return { page, id };
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
        if (id) {
            state.currentGuideId = id;
            renderSections();
        } else {
            state.currentGuideId = null;
            renderGuidesList();
        }
    } else if (page === 'tests') {
        if (id) {
            state.currentTestId = id;
            startTest(id);
        } else {
            state.currentTestId = null;
            renderTestsList();
        }
    }
}

// Универсальное модальное окно
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

// PWA установка (без Service Worker)
let deferredPrompt;
let installPromptShown = false;

function setupPWA() {
    window.addEventListener('beforeinstallprompt', (e) => {
        e.preventDefault();
        deferredPrompt = e;
        if (!installPromptShown && state.dataLoaded) {
            setTimeout(showInstallPrompt, 3000);
        }
    });
}

function showInstallPrompt() {
    if (!deferredPrompt) return;
    installPromptShown = true;
    elements.modalContent.innerHTML = `
        <h3 style="margin-bottom:16px;">📱 Установить приложение</h3>
        <p>Добавьте "Отдел знаний" на главный экран для быстрого доступа</p>
        <div class="modal-buttons">
            <button class="modal-btn cancel" id="modalCancel">Закрыть</button>
            <button class="modal-btn confirm" id="modalInstall">Установить</button>
        </div>
    `;
    elements.modalOverlay.style.display = 'flex';
    document.getElementById('modalCancel').onclick = () => {
        elements.modalOverlay.style.display = 'none';
        installPromptShown = false;
    };
    document.getElementById('modalInstall').onclick = async () => {
        deferredPrompt.prompt();
        const { outcome } = await deferredPrompt.userChoice;
        deferredPrompt = null;
        elements.modalOverlay.style.display = 'none';
        if (outcome === 'accepted') installPromptShown = true;
    };
}

function openLightbox(src) {
    elements.modalContent.innerHTML = `
        <div style="position: relative; max-width: 90vw; max-height: 90vh;">
            <img src="${escapeHtml(src)}" style="width: 100%; height: auto; display: block; border-radius: var(--radius);">
            <button class="lightbox-close" style="position: absolute; top: -15px; right: -15px; background: var(--primary); color: white; border: none; border-radius: 50%; width: 40px; height: 40px; cursor: pointer; font-size: 24px; line-height: 1; box-shadow: var(--shadow);">×</button>
        </div>
    `;
    elements.modalOverlay.style.display = 'flex';
    const closeBtn = elements.modalContent.querySelector('.lightbox-close');
    closeBtn.onclick = () => elements.modalOverlay.style.display = 'none';
    elements.modalOverlay.onclick = (e) => {
        if (e.target === elements.modalOverlay) {
            elements.modalOverlay.style.display = 'none';
        }
    };
}

// Инициализация
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
    elements.notificationsContainer.addEventListener('click', (e) => {
        const closeBtn = e.target.closest('.notification-close');
        if (!closeBtn) return;
        const item = closeBtn.closest('.notification-item');
        if (!item) return;
        const id = item.dataset.id;
        closeNotification(id);
    });
    await loadData();
}

init();