// app.js
(function() {
    'use strict';

    // Состояние приложения
    const state = {
        guides: [],
        sections: [],
        tests: [],
        questions: [],
        dataLoaded: false,
        currentPage: 'guides',
        currentGuideId: null,
        currentTestId: null,
        currentQuestionIndex: 0,
        testAnswers: {},
        testStartTime: null,
        testQuestions: [], // было упущено при инициализации
        searchQuery: '',
        isOnline: navigator.onLine
    };

    // DOM элементы
    const elements = {
        header: document.getElementById('appHeader'),
        navTabs: document.getElementById('navTabs'),
        searchWrapper: document.getElementById('searchWrapper'),
        searchInput: document.getElementById('searchInput'),
        contentContainer: document.getElementById('contentContainer'),
        loader: document.getElementById('loader'),
        modalOverlay: document.getElementById('modalOverlay'),
        modalContent: document.getElementById('modalContent')
    };

    let timerInterval = null;

    // Вспомогательные функции
    function formatExcelDate(excelDate) {
        if (!excelDate) return '';
        if (typeof excelDate === 'string') return excelDate;
        const date = new Date((excelDate - 25569) * 86400 * 1000);
        return date.toLocaleDateString('ru-RU');
    }

    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function parseMediaList(str) {
        if (!str) return [];
        return str.split(',').map(s => s.trim()).filter(s => s.length > 0);
    }

    function getFileName(path) {
        const parts = String(path).split('/');
        return parts[parts.length - 1] || 'Файл';
    }

    function processContent(content) {
        return content || '';
    }

    function shuffleArray(arr) {
        const a = [...arr];
        for (let i = a.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [a[i], a[j]] = [a[j], a[i]];
        }
        return a;
    }

    function shuffleAnswers(question) {
        if (question.text_answer) return [];
        const answers = [];
        for (let i = 1; i <= 6; i++) {
            const answer = question[`answer${i}`];
            if (answer !== undefined && answer !== null && String(answer).trim() !== '') {
                answers.push(String(answer).trim());
            }
        }
        return shuffleArray(answers);
    }

    // Инициализация
    async function init() {
        // Проверяем наличие иконок и генерируем только если их нет
        async function checkAndGenerateIcons() {
            const manifestLink = document.querySelector('link[rel="manifest"]');
            if (!manifestLink) return;
            
            // Пробуем загрузить существующие иконки
            const icon192exists = await fetch('./img/icon-192.png').then(r => r.ok).catch(() => false);
            const icon512exists = await fetch('./img/icon-512.png').then(r => r.ok).catch(() => false);
            
            // Если обе есть - используем их
            if (icon192exists && icon512exists) {
                manifestLink.href = 'manifest.json';
                return;
            }
            
            // Если нет - генерируем
            function generateIcon(size) {
                const canvas = document.createElement('canvas');
                canvas.width = size;
                canvas.height = size;
                const ctx = canvas.getContext('2d');
                
                ctx.fillStyle = '#2c3e50';
                ctx.fillRect(0, 0, size, size);
                
                ctx.fillStyle = 'white';
                ctx.font = `bold ${size * 0.4}px Arial`;
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.fillText('📚', size/2, size/2);
                
                return canvas.toDataURL();
            }
            
            const manifest = {
                name: "Отдел знаний",
                short_name: "Знания",
                start_url: ".",
                display: "standalone",
                background_color: "#f5f7fa",
                theme_color: "#2c3e50",
                icons: [
                    { src: generateIcon(192), sizes: "192x192", type: "image/png" },
                    { src: generateIcon(512), sizes: "512x512", type: "image/png" }
                ]
            };
            
            const manifestBlob = new Blob([JSON.stringify(manifest)], { type: 'application/json' });
            const manifestURL = URL.createObjectURL(manifestBlob);
            manifestLink.href = manifestURL;
        }
        
        await checkAndGenerateIcons();
        
        registerServiceWorker();
        setupEventListeners();
        await loadData();
        addInstallButton();
        handleRouting();
        window.addEventListener('popstate', handleRouting);
        updateOnlineStatus();
    }

    function updateOnlineStatus() {
        state.isOnline = navigator.onLine;
        window.addEventListener('online', () => { state.isOnline = true; });
        window.addEventListener('offline', () => { state.isOnline = false; });
    }

    function registerServiceWorker() {
        if ('serviceWorker' in navigator) {
            navigator.serviceWorker.register('sw.js').catch(console.warn);
        }
    }

    function setupEventListeners() {
        elements.navTabs.addEventListener('click', (e) => {
            const btn = e.target.closest('.nav-btn');
            if (!btn) return;
            const page = btn.dataset.page;
            if (page) {
                navigateTo(page);
            }
        });

        elements.searchInput.addEventListener('input', (e) => {
            state.searchQuery = e.target.value.trim().toLowerCase();
            if (state.currentPage === 'guides' && !state.currentGuideId) {
                renderGuidesList();
            } else if (state.currentPage === 'guides' && state.currentGuideId) {
                renderSections();
            }
        });
    }

    async function loadData() {
        try {
            const [guidesData, testsData] = await Promise.all([
                loadExcel('guide.xlsx'),
                loadExcel('tests.xlsx')
            ]);

            if (guidesData) {
                state.guides = (guidesData.guides || []).map(g => ({ ...g, id: String(g.id) }));
                state.sections = (guidesData.sections || []).map(s => ({ ...s, id: String(s.id), guide_id: String(s.guide_id) }));
            }

            if (testsData) {
                state.tests = (testsData.tests || []).map(t => ({ ...t, id: String(t.id) }));
                state.questions = (testsData.questions || []).map(q => ({ ...q, id: String(q.id), test_id: String(q.test_id) }));
            }

            state.dataLoaded = true;
            elements.searchInput.disabled = false;
            hideLoader();
        } catch (error) {
            console.error('Ошибка загрузки данных:', error);
            showError('Не удалось загрузить данные. Проверьте наличие файлов guide.xlsx и tests.xlsx');
        }
    }

    function loadExcel(filename) {
        return new Promise((resolve, reject) => {
            const xhr = new XMLHttpRequest();
            xhr.open('GET', filename, true);
            xhr.responseType = 'arraybuffer';
            
            xhr.onload = function() {
                if (xhr.status === 200) {
                    try {
                        const data = new Uint8Array(xhr.response);
                        const workbook = XLSX.read(data, { type: 'array' });
                        const result = {};
                        
                        workbook.SheetNames.forEach(sheetName => {
                            const sheet = workbook.Sheets[sheetName];
                            const jsonData = XLSX.utils.sheet_to_json(sheet);
                            result[sheetName] = jsonData;
                        });
                        
                        resolve(result);
                    } catch (e) {
                        reject(e);
                    }
                } else {
                    reject(new Error(`HTTP ${xhr.status}`));
                }
            };
            
            xhr.onerror = () => reject(new Error('Network error'));
            xhr.send();
        });
    }

    function hideLoader() {
        elements.loader.style.display = 'none';
    }

    function showError(message) {
        elements.loader.innerHTML = `<div class="empty-state">❌ ${message}</div>`;
    }

    function navigateTo(page, id = null) {
        const url = new URL(window.location);
        url.searchParams.set('page', page);
        if (id) {
            url.searchParams.set('id', String(id));
        } else {
            url.searchParams.delete('id');
        }
        window.history.pushState({}, '', url);
        
        updateActiveTab(page);
        handleRouting();
    }

    function updateActiveTab(page) {
        document.querySelectorAll('.nav-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.page === page);
        });
        
        elements.searchWrapper.style.display = (page === 'guides') ? 'block' : 'none';
        if (page !== 'guides') {
            elements.searchInput.value = '';
            state.searchQuery = '';
        }
    }

    function handleRouting() {
        const params = new URLSearchParams(window.location.search);
        const page = params.get('page') || 'guides';
        const id = params.get('id');
        
        state.currentPage = page;
        updateActiveTab(page);
        
        if (timerInterval) {
            clearInterval(timerInterval);
            timerInterval = null;
        }
        
        if (page === 'guides') {
            if (id) {
                state.currentGuideId = String(id);
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

    function renderGuidesList() {
        const filteredGuides = state.guides;
        
        if (filteredGuides.length === 0) {
            elements.contentContainer.innerHTML = '<div class="empty-state">Справочники не найдены</div>';
            return;
        }
        
        const html = `
            <div class="guides-grid">
                ${filteredGuides.map(guide => `
                    <div class="guide-card" data-guide-id="${guide.id}">
                        ${guide.image ? `<img src="${escapeHtml(guide.image)}" class="card-image" alt="${escapeHtml(guide.title)}" loading="lazy">` : 
                        `<div class="card-image"></div>`}
                        <div class="card-content">
                            <h3 class="card-title">${escapeHtml(guide.title || 'Без названия')}</h3>
                            <div class="card-meta">
                                ${guide.author ? `<span>${escapeHtml(guide.author)}</span>` : ''}
                                ${guide.date ? `<span>${formatExcelDate(guide.date)}</span>` : ''}
                            </div>
                        </div>
                    </div>
                `).join('')}
            </div>
        `;
        
        elements.contentContainer.innerHTML = html;
        
        document.querySelectorAll('.guide-card').forEach(card => {
            card.addEventListener('click', () => {
                navigateTo('guides', card.dataset.guideId);
            });
        });
    }

    function renderSections() {
        const guide = state.guides.find(g => String(g.id) === String(state.currentGuideId));
        if (!guide) {
            navigateTo('guides');
            return;
        }
        
        let sections = state.sections.filter(s => String(s.guide_id) === String(state.currentGuideId));
        
        if (state.searchQuery) {
            sections = sections.filter(s => 
                (s.title || '').toLowerCase().includes(state.searchQuery)
            );
        }
        
        const html = `
            <div class="sections-container">
                <button class="back-button" id="backToGuides">← Назад к справочникам</button>
                <h2 style="margin-bottom: 24px; color: var(--primary);">${escapeHtml(guide.title || 'Справочник')}</h2>
                ${sections.length === 0 ? '<div class="empty-state">Разделы не найдены</div>' : 
                    sections.map(section => renderSection(section)).join('')}
            </div>
        `;
        
        elements.contentContainer.innerHTML = html;
        
        document.getElementById('backToGuides')?.addEventListener('click', () => {
            navigateTo('guides');
        });
    }

    function renderSection(section) {
        const images = parseMediaList(section.images);
        const files = parseMediaList(section.files);
        const videos = parseMediaList(section.videos);
        
        return `
            <div class="section-card">
                <h3 class="section-title">${escapeHtml(section.title || 'Без названия')}</h3>
                <div class="section-content">${processContent(section.content || '')}</div>
                
                ${images.length > 0 ? `
                    <div class="media-section">
                        <div class="media-title">Изображения</div>
                        <div class="images-grid">
                            ${images.map(img => state.isOnline ? 
                                `<img src="${escapeHtml(img)}" class="section-image" alt="Изображение" loading="lazy" onerror="if(this.parentElement) this.parentElement.innerHTML='<div class=\\'offline-placeholder\\'>Не удалось загрузить изображение</div>'">` :
                                '<div class="offline-placeholder">Изображение недоступно в офлайн-режиме</div>'
                            ).join('')}
                        </div>
                    </div>
                ` : ''}
                
                ${files.length > 0 ? `
                    <div class="media-section">
                        <div class="media-title">Файлы для скачивания</div>
                        <ul class="files-list">
                            ${files.map(file => state.isOnline ?
                                `<li class="file-item"><a href="${escapeHtml(file)}" class="file-link" download>📄 ${getFileName(file)}</a></li>` :
                                '<li class="file-item"><span class="offline-placeholder" style="display: inline-block; padding: 8px;">Файл недоступен в офлайн-режиме</span></li>'
                            ).join('')}
                        </ul>
                    </div>
                ` : ''}
                
                ${videos.length > 0 ? `
                    <div class="media-section">
                        <div class="media-title">Видео</div>
                        <ul class="videos-list">
                            ${videos.map(video => state.isOnline ?
                                `<li class="video-item"><a href="${escapeHtml(video)}" class="video-link" target="_blank">🎬 ${getFileName(video)}</a></li>` :
                                '<li class="video-item"><span class="offline-placeholder" style="display: inline-block; padding: 8px;">Видео недоступно в офлайн-режиме</span></li>'
                            ).join('')}
                        </ul>
                    </div>
                ` : ''}
            </div>
        `;
    }

    function renderTestsList() {
        if (state.tests.length === 0) {
            elements.contentContainer.innerHTML = '<div class="empty-state">Тесты не найдены</div>';
            return;
        }
        
        const html = `
            <div class="tests-grid">
                ${state.tests.map(test => `
                    <div class="test-card" data-test-id="${test.id}">
                        ${test.image ? `<img src="${escapeHtml(test.image)}" class="card-image" alt="${escapeHtml(test.title)}" loading="lazy">` : 
                        `<div class="card-image"></div>`}
                        <div class="card-content">
                            <h3 class="card-title">${escapeHtml(test.title || 'Без названия')}</h3>
                            <div class="card-meta">
                                ${test.author ? `<span>${escapeHtml(test.author)}</span>` : ''}
                                ${test.date ? `<span>${formatExcelDate(test.date)}</span>` : ''}
                                ${test.time_limit && Number(test.time_limit) > 0 ? `<span>⏱ ${test.time_limit} мин</span>` : ''}
                            </div>
                        </div>
                    </div>
                `).join('')}
            </div>
        `;
        
        elements.contentContainer.innerHTML = html;
        
        document.querySelectorAll('.test-card').forEach(card => {
            card.addEventListener('click', () => {
                navigateTo('tests', card.dataset.testId);
            });
        });
    }

    function startTest(testId) {
        const test = state.tests.find(t => String(t.id) === String(testId));
        const questions = state.questions.filter(q => String(q.test_id) === String(testId));
        
        if (!test || questions.length === 0) {
            navigateTo('tests');
            return;
        }
        
        if (timerInterval) {
            clearInterval(timerInterval);
            timerInterval = null;
        }
        
        state.currentTestId = String(testId);
        state.currentQuestionIndex = 0;
        state.testAnswers = {};
        state.testStartTime = Date.now();
        state.testQuestions = questions.map(q => ({
            ...q,
            shuffledAnswers: shuffleAnswers(q)
        }));
        
        renderTestQuestion();
    }

    function startTimer(seconds) {
        if (timerInterval) clearInterval(timerInterval);
        
        const timerElement = document.getElementById('timerDisplay');
        if (!timerElement) return;
        
        let remaining = seconds;
        
        const updateTimerDisplay = () => {
            const mins = Math.floor(remaining / 60);
            const secs = remaining % 60;
            timerElement.textContent = `${mins}:${secs.toString().padStart(2, '0')}`;
            
            timerElement.classList.remove('warning', 'danger');
            if (remaining <= 60) {
                timerElement.classList.add('danger');
            } else if (remaining <= 120) {
                timerElement.classList.add('warning');
            }
        };
        
        updateTimerDisplay();
        
        timerInterval = setInterval(() => {
            remaining--;
            updateTimerDisplay();
            
            if (remaining <= 0) {
                clearInterval(timerInterval);
                timerInterval = null;
                finishTest(true);
            }
        }, 1000);
    }

    function renderTestQuestion() {
        const test = state.tests.find(t => String(t.id) === String(state.currentTestId));
        const questions = state.testQuestions;
        
        if (!test || !questions.length) {
            navigateTo('tests');
            return;
        }
        
        const currentQ = questions[state.currentQuestionIndex];
        const savedAnswer = state.testAnswers[currentQ.id];
        
        const isTextAnswer = !!currentQ.text_answer;
        const answers = isTextAnswer ? [] : currentQ.shuffledAnswers;
        
        const progress = `${state.currentQuestionIndex + 1} / ${questions.length}`;
        
        const html = `
            <div class="test-container">
                <button class="back-button" id="exitTest">← К списку тестов</button>
                
                <div class="test-header">
                    <span class="test-progress">${escapeHtml(test.title)} — ${progress}</span>
                    ${test.time_limit && Number(test.time_limit) > 0 ? `<span class="timer" id="timerDisplay"></span>` : ''}
                </div>
                
                <div class="question-card">
                    <div class="question-text">${escapeHtml(currentQ.question || '')}</div>
                    
                    ${isTextAnswer ? `
                        <input type="text" class="text-answer-input" id="textAnswer" 
                               placeholder="Введите ваш ответ..." 
                               value="${escapeHtml(savedAnswer || '')}">
                    ` : `
                        <div class="answers-list">
                            ${answers.map((answer) => `
                                <div class="answer-item ${savedAnswer === answer ? 'selected' : ''}" data-answer="${escapeHtml(answer)}">
                                    <span class="answer-radio"></span>
                                    <span class="answer-text">${escapeHtml(answer)}</span>
                                </div>
                            `).join('')}
                        </div>
                    `}
                </div>
                
                <div class="test-navigation">
                    ${state.currentQuestionIndex > 0 ? 
                        '<button class="nav-test-btn" id="prevQuestion">← Назад</button>' : 
                        '<div></div>'}
                    
                    ${state.currentQuestionIndex < questions.length - 1 ?
                        '<button class="nav-test-btn primary" id="nextQuestion">Далее →</button>' :
                        '<button class="nav-test-btn success" id="finishTest">Завершить тест</button>'}
                </div>
            </div>
        `;
        
        elements.contentContainer.innerHTML = html;
        
        // Запуск таймера если нужно
        if (test.time_limit && Number(test.time_limit) > 0) {
            const elapsed = Math.floor((Date.now() - state.testStartTime) / 1000);
            const total = Number(test.time_limit) * 60;
            const remaining = Math.max(0, total - elapsed);
            
            if (remaining > 0) {
                startTimer(remaining);
            } else {
                finishTest(true);
                return;
            }
        }
        
        document.getElementById('exitTest')?.addEventListener('click', () => {
            showConfirmModal('Вы уверены? Прогресс теста будет потерян.', () => {
                if (timerInterval) {
                    clearInterval(timerInterval);
                    timerInterval = null;
                }
                navigateTo('tests');
            });
        });
        
        if (!isTextAnswer) {
            document.querySelectorAll('.answer-item').forEach(item => {
                item.addEventListener('click', () => {
                    const answer = item.dataset.answer;
                    saveAnswer(currentQ.id, answer);
                    
                    document.querySelectorAll('.answer-item').forEach(i => {
                        i.classList.toggle('selected', i.dataset.answer === answer);
                    });
                });
            });
        } else {
            const input = document.getElementById('textAnswer');
            input?.addEventListener('input', (e) => {
                saveAnswer(currentQ.id, e.target.value);
            });
        }
        
        document.getElementById('prevQuestion')?.addEventListener('click', () => {
            if (state.currentQuestionIndex > 0) {
                state.currentQuestionIndex--;
                renderTestQuestion();
            }
        });
        
        document.getElementById('nextQuestion')?.addEventListener('click', () => {
            if (state.currentQuestionIndex < questions.length - 1) {
                state.currentQuestionIndex++;
                renderTestQuestion();
            }
        });
        
        document.getElementById('finishTest')?.addEventListener('click', () => {
            showConfirmModal('Завершить тест и посмотреть результат?', () => finishTest(false));
        });
    }

    function saveAnswer(questionId, answer) {
        state.testAnswers[questionId] = answer;
    }

    function finishTest(isTimeout) {
        if (timerInterval) {
            clearInterval(timerInterval);
            timerInterval = null;
        }
        
        const questions = state.testQuestions;
        let correct = 0;
        
        questions.forEach(q => {
            const userAnswer = String(state.testAnswers[q.id] || '').trim();
            const correctAnswer = String(q.answer1 || '').trim();
            
            // Пропускаем, если ответ не выбран
            if (userAnswer === '') {
                return;
            }
            
            if (q.text_answer) {
                if (userAnswer.toLowerCase() === correctAnswer.toLowerCase()) {
                    correct++;
                }
            } else {
                if (userAnswer.toLowerCase() === correctAnswer.toLowerCase()) {
                    correct++;
                }
            }
        });
        
        const percent = questions.length > 0 ? Math.round((correct / questions.length) * 100) : 0;
        
        let percentClass = 'bad';
        if (percent >= 70) percentClass = 'good';
        else if (percent >= 40) percentClass = 'medium';
        
        const html = `
            <div class="test-container">
                <button class="back-button" id="backToTests">← К списку тестов</button>
                
                <div class="result-card">
                    <h2>Тест завершен${isTimeout ? ' (время вышло)' : ''}</h2>
                    <div class="result-percent ${percentClass}">${percent}%</div>
                    <div class="result-details">
                        Правильных ответов: ${correct} из ${questions.length}
                    </div>
                    <button class="nav-test-btn primary" id="retakeTest">Пройти заново</button>
                </div>
            </div>
        `;
        
        elements.contentContainer.innerHTML = html;
        
        document.getElementById('backToTests')?.addEventListener('click', () => {
            navigateTo('tests');
        });
        
        document.getElementById('retakeTest')?.addEventListener('click', () => {
            startTest(state.currentTestId);
        });
    }

    function showConfirmModal(message, onConfirm) {
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
        };
        
        document.getElementById('modalConfirm').onclick = () => {
            elements.modalOverlay.style.display = 'none';
            onConfirm();
        };
        
        elements.modalOverlay.onclick = (e) => {
            if (e.target === elements.modalOverlay) {
                elements.modalOverlay.style.display = 'none';
            }
        };
    }

    init();
})();

// PWA установка
let deferredPrompt;
let installPromptShown = false;

window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    
    // Показываем не сразу, а через 3 секунды после загрузки
    setTimeout(() => {
        if (deferredPrompt && !installPromptShown && state.dataLoaded) {
            showInstallPrompt();
        }
    }, 3000);
});

function showInstallPrompt() {
    if (!deferredPrompt) return;
    installPromptShown = true;
    
    elements.modalContent.innerHTML = `
        <h3 style="margin-bottom: 16px; font-size: 1.5rem;">📱 Установить приложение</h3>
        <p style="margin-bottom: 20px;">Добавьте "Отдел знаний" на главный экран для быстрого доступа и работы без интернета</p>
        <div class="modal-buttons">
            <button class="modal-btn cancel" id="modalCancel">Позже</button>
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
        
        if (outcome === 'accepted') {
            console.log('PWA установлено');
        }
    };
}

// Кнопка установки в интерфейсе (опционально)
function addInstallButton() {
    const header = elements.header;
    const installBtn = document.createElement('button');
    installBtn.id = 'installBtn';
    installBtn.className = 'install-btn';
    installBtn.innerHTML = '📱 Установить';
    installBtn.style.display = 'none';
    installBtn.onclick = () => {
        if (deferredPrompt) {
            showInstallPrompt();
        } else {
            alert('Установка недоступна. Возможно, приложение уже установлено.');
        }
    };
    
    header.appendChild(installBtn);
    
    window.addEventListener('beforeinstallprompt', () => {
        installBtn.style.display = 'block';
    });
    
    window.addEventListener('appinstalled', () => {
        installBtn.style.display = 'none';
        deferredPrompt = null;
    });
}