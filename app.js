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

    // Инициализация
    async function init() {
        registerServiceWorker();
        setupEventListeners();
        await loadData();
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
                state.guides = guidesData.guides || [];
                state.sections = guidesData.sections || [];
            }

            if (testsData) {
                state.tests = testsData.tests || [];
                state.questions = testsData.questions || [];
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
                            
                            if (sheetName === 'guides' || sheetName === 'tests') {
                                result[sheetName] = jsonData.map(row => ({
                                    ...row,
                                    id: String(row.id || '')
                                }));
                            } else if (sheetName === 'sections') {
                                result[sheetName] = jsonData.map(row => ({
                                    ...row,
                                    id: String(row.id || ''),
                                    guide_id: String(row.guide_id || '')
                                }));
                            } else if (sheetName === 'questions') {
                                result[sheetName] = jsonData.map(row => ({
                                    ...row,
                                    id: String(row.id || ''),
                                    test_id: String(row.test_id || '')
                                }));
                            }
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

    function showLoader() {
        elements.loader.style.display = 'block';
        elements.contentContainer.innerHTML = '';
    }

    function showError(message) {
        elements.loader.innerHTML = `<div class="empty-state">❌ ${message}</div>`;
    }

    function navigateTo(page, id = null) {
        const url = new URL(window.location);
        url.searchParams.set('page', page);
        if (id) {
            url.searchParams.set('id', id);
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
                    ${guide.image ? `<img src="${guide.image}" class="card-image" alt="${escapeHtml(guide.title)}" loading="lazy" onerror="this.src='data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'300\' height=\'180\' viewBox=\'0 0 300 180\'%3E%3Crect width=\'300\' height=\'180\' fill=\'%23e9ecef\'/%3E%3Ctext x=\'150\' y=\'90\' text-anchor=\'middle\' fill=\'%2395a5a6\' font-size=\'14\'%3EНет изображения%3C/text%3E%3C/svg%3E'">` : 
                    `<div class="card-image" style="background: #e9ecef; display: flex; align-items: center; justify-content: center; color: #95a5a6;">Нет изображения</div>`}
                    <div class="card-content">
                        <h3 class="card-title">${escapeHtml(guide.title || 'Без названия')}</h3>
                        <div class="card-meta">
                            ${guide.author ? `<span>${escapeHtml(guide.author)}</span>` : ''}
                            ${guide.date ? `<span>${escapeHtml(guide.date)}</span>` : ''}
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
        const guide = state.guides.find(g => g.id === state.currentGuideId);
        if (!guide) {
            navigateTo('guides');
            return;
        }
        
        let sections = state.sections.filter(s => s.guide_id === state.currentGuideId);
        
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
                            `<img src="${escapeHtml(img)}" class="section-image" alt="Изображение" loading="lazy" onerror="this.parentElement.innerHTML='<div class=\\'offline-placeholder\\'>Не удалось загрузить изображение</div>'">` :
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

    function parseMediaList(str) {
        if (!str) return [];
        return str.split(',').map(s => s.trim()).filter(s => s.length > 0);
    }

    function getFileName(path) {
        const parts = path.split('/');
        return parts[parts.length - 1] || 'Файл';
    }

    function processContent(content) {
        if (!content) return '';
        // q тег уже обрабатывается CSS, просто возвращаем как есть
        return content;
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
                    ${test.image ? `<img src="${test.image}" class="card-image" alt="${escapeHtml(test.title)}" loading="lazy" onerror="this.src='data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'300\' height=\'180\' viewBox=\'0 0 300 180\'%3E%3Crect width=\'300\' height=\'180\' fill=\'%23e9ecef\'/%3E%3Ctext x=\'150\' y=\'90\' text-anchor=\'middle\' fill=\'%2395a5a6\' font-size=\'14\'%3EНет изображения%3C/text%3E%3C/svg%3E'">` : 
                    `<div class="card-image" style="background: #e9ecef; display: flex; align-items: center; justify-content: center; color: #95a5a6;">Нет изображения</div>`}
                    <div class="card-content">
                        <h3 class="card-title">${escapeHtml(test.title || 'Без названия')}</h3>
                        <div class="card-meta">
                            ${test.author ? `<span>${escapeHtml(test.author)}</span>` : ''}
                            ${test.date ? `<span>${escapeHtml(test.date)}</span>` : ''}
                            ${test.time_limit && test.time_limit > 0 ? `<span>⏱ ${test.time_limit} мин</span>` : ''}
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
        const test = state.tests.find(t => t.id === testId);
        const questions = state.questions.filter(q => q.test_id === testId);
        
        if (!test || questions.length === 0) {
            navigateTo('tests');
            return;
        }
        
        state.currentTestId = testId;
        state.currentQuestionIndex = 0;
        state.testAnswers = {};
        state.testStartTime = Date.now();
        state.testQuestions = questions.map(q => ({
            ...q,
            shuffledAnswers: shuffleAnswers(q)
        }));
        
        renderTestQuestion();
        
        if (test.time_limit && test.time_limit > 0) {
            startTimer(test.time_limit * 60);
        }
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

    function shuffleArray(arr) {
        const a = [...arr];
        for (let i = a.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [a[i], a[j]] = [a[j], a[i]];
        }
        return a;
    }

    let timerInterval = null;

    function startTimer(seconds) {
        if (timerInterval) clearInterval(timerInterval);
        
        const timerElement = document.getElementById('timerDisplay');
        if (!timerElement) return;
        
        let remaining = seconds;
        
        timerInterval = setInterval(() => {
            remaining--;
            
            if (remaining <= 0) {
                clearInterval(timerInterval);
                timerInterval = null;
                finishTest(true);
                return;
            }
            
            const mins = Math.floor(remaining / 60);
            const secs = remaining % 60;
            timerElement.textContent = `${mins}:${secs.toString().padStart(2, '0')}`;
            
            timerElement.classList.remove('warning', 'danger');
            if (remaining <= 60) {
                timerElement.classList.add('danger');
            } else if (remaining <= 120) {
                timerElement.classList.add('warning');
            }
        }, 1000);
    }

    function renderTestQuestion() {
        const test = state.tests.find(t => t.id === state.currentTestId);
        const questions = state.testQuestions;
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
                    ${test.time_limit && test.time_limit > 0 ? `<span class="timer" id="timerDisplay">${test.time_limit}:00</span>` : ''}
                </div>
                
                <div class="question-card">
                    <div class="question-text">${escapeHtml(currentQ.question || '')}</div>
                    
                    ${isTextAnswer ? `
                        <input type="text" class="text-answer-input" id="textAnswer" 
                               placeholder="Введите ваш ответ..." 
                               value="${escapeHtml(savedAnswer || '')}">
                    ` : `
                        <div class="answers-list">
                            ${answers.map((answer, idx) => `
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
        
        document.getElementById('exitTest')?.addEventListener('click', () => {
            showConfirmModal('Вы уверены? Прогресс теста будет потерян.', () => {
                navigateTo('tests');
                if (timerInterval) {
                    clearInterval(timerInterval);
                    timerInterval = null;
                }
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
        
        if (test.time_limit && test.time_limit > 0) {
            const elapsed = Math.floor((Date.now() - state.testStartTime) / 1000);
            const remaining = test.time_limit * 60 - elapsed;
            if (remaining > 0) {
                startTimer(remaining);
            } else {
                finishTest(true);
            }
        }
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
            const userAnswer = state.testAnswers[q.id] || '';
            const correctAnswer = (q.answer1 || '').trim();
            
            if (q.text_answer) {
                if (userAnswer.trim().toLowerCase() === correctAnswer.toLowerCase()) {
                    correct++;
                }
            } else {
                if (userAnswer.trim().toLowerCase() === correctAnswer.toLowerCase()) {
                    correct++;
                }
            }
        });
        
        const percent = Math.round((correct / questions.length) * 100);
        
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

    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // Запуск
    init();
})();