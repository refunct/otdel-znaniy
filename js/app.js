// app.js
import { renderGuidesList, renderSections } from './guides.js';
import { renderTestsList, startTest } from './tests.js';

// Глобальное состояние
export const state = {
    guides: [],
    sections: [],
    tests: [],
    questions: [],
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

// Загрузка Excel
async function loadExcel(filename) {
    return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        xhr.open('GET', filename, true);
        xhr.responseType = 'arraybuffer';
        xhr.onload = () => {
            if (xhr.status === 200) {
                try {
                    const data = new Uint8Array(xhr.response);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const result = {};
                    workbook.SheetNames.forEach(sheetName => {
                        result[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                    });
                    resolve(result);
                } catch (e) { reject(e); }
            } else { reject(new Error(`HTTP ${xhr.status}`)); }
        };
        xhr.onerror = () => reject(new Error('Network error'));
        xhr.send();
    });
}

async function loadData() {
    try {
        const [guidesData, testsData] = await Promise.all([
            loadExcel('docs/guide.xlsx'),
            loadExcel('docs/tests.xlsx')
        ]);
        state.guides = (guidesData.guides || []).map(g => ({ ...g, id: String(g.id) }));
        state.sections = (guidesData.sections || []).map(s => ({ ...s, id: String(s.id), guide_id: String(s.guide_id) }));
        state.tests = (testsData.tests || []).map(t => ({ ...t, id: String(t.id) }));
        state.questions = (testsData.questions || []).map(q => ({ ...q, id: String(q.id), test_id: String(q.test_id) }));
        state.dataLoaded = true;
        elements.loader.style.display = 'none';
        handleRouting();
    } catch (error) {
        console.error(error);
        elements.loader.innerHTML = '<div class="empty-state">❌ Ошибка загрузки данных</div>';
    }
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
            state.currentTestId = String(id);
            startTest(String(id));
        } else {
            state.currentTestId = null;
            renderTestsList();
        }
    }
}

// PWA установка
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
        <p>Добавьте "Отдел знаний" на главный экран для быстрого доступа и работы без интернета</p>
        <div class="modal-buttons">
            <button class="modal-btn cancel" id="modalCancel">Закрыть</button>
            <button class="modal-btn confirm" id="modalInstall">Установить</button>
        </div>
    `;
    elements.modalOverlay.style.display = 'flex';
    document.getElementById('modalCancel').onclick = () => {
        elements.modalOverlay.style.display = 'none';
        installPromptShown = false; // Появится при следующем обновлении
    };
    document.getElementById('modalInstall').onclick = async () => {
        deferredPrompt.prompt();
        const { outcome } = await deferredPrompt.userChoice;
        deferredPrompt = null;
        elements.modalOverlay.style.display = 'none';
        if (outcome === 'accepted') installPromptShown = true; // больше не показывать
    };
}

// Service Worker
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('js/sw.js').catch(console.warn);
}

// Инициализация
async function init() {
    setupPWA();
    window.addEventListener('hashchange', handleRouting);
    await loadData();
    elements.navTabs.addEventListener('click', (e) => {
        const btn = e.target.closest('.nav-btn');
        if (!btn) return;
        const page = btn.dataset.page;
        if (page) navigateTo(page);
    });
}

init();