// tests.js
import { state, elements, escapeHtml, navigateTo, formatExcelDate, showModal } from './app.js';

let timerInterval = null;
let testState = {
    questions: [],
    answers: {},
    startTime: null,
    currentIndex: 0,
    test: null
};

export function resetTestTimer() {
    if (timerInterval) {
        clearInterval(timerInterval);
        timerInterval = null;
    }
}

function shuffle(arr) {
    for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
}

function prepareQuestions(testId) {
    // Приводим к строке для надёжности
    return state.questions
        .filter(q => String(q.test_id) === String(testId))
        .map(q => {
            const answers = [];
            for (let i = 1; i <= 6; i++) {
                const ans = q[`answer${i}`];
                if (ans && String(ans).trim() !== '') {
                    answers.push(String(ans).trim());
                }
            }
            return {
                ...q,
                shuffledAnswers: q.text_answer ? [] : shuffle(answers)
            };
        });
}

export function renderTestsList() {
    if (!state.tests.length) {
        elements.contentContainer.innerHTML = '<div class="empty-state">Тесты не найдены</div>';
        return;
    }
    const html = `<div class="tests-grid">${state.tests.map(t => `
        <div class="test-card" data-test-id="${t.id}">
            ${t.image ? `<img src="${escapeHtml(t.image)}" class="card-image" loading="lazy">` : '<div class="card-image"></div>'}
            <div class="card-content">
                <h3 class="card-title">${escapeHtml(t.title)}</h3>
                <div class="card-meta">
                    ${t.author ? `<span>${escapeHtml(t.author)}</span>` : ''}
                    ${t.date ? `<span>${formatExcelDate(t.date)}</span>` : ''}
                    ${t.time_limit > 0 ? `<span>⏱ ${t.time_limit} мин</span>` : ''}
                </div>
            </div>
        </div>`).join('')}</div>`;
    elements.contentContainer.innerHTML = html;
    document.querySelectorAll('.test-card').forEach(c => {
        c.addEventListener('click', () => {
            const testId = c.dataset.testId;
            console.log('Opening test:', testId);
            navigateTo('tests', testId);
        });
    });
}

export function startTest(testId) {
    console.log('startTest called with', testId);
    const test = state.tests.find(t => String(t.id) === String(testId));
    if (!test) {
        console.error('Test not found for id', testId);
        return navigateTo('tests');
    }
    const questions = prepareQuestions(testId);
    if (!questions.length) {
        console.error('No questions for test', testId);
        return navigateTo('tests');
    }
    testState = {
        questions,
        answers: {},
        startTime: Date.now(),
        currentIndex: 0,
        test: test
    };
    if (timerInterval) clearInterval(timerInterval);
    renderQuestion();
}

function renderQuestion() {
    const { questions, currentIndex, answers, test } = testState;
    const q = questions[currentIndex];
    const saved = answers[q.id];
    const isText = !!q.text_answer;
    const html = `
        <div class="test-container">
            <button class="back-button" id="exitTest">← К тестам</button>
            <div class="test-header">
                <span>${escapeHtml(test.title)} — ${currentIndex+1}/${questions.length}</span>
                ${test.time_limit > 0 ? `<span class="timer" id="timerDisplay"></span>` : ''}
            </div>
            <div class="question-card">
                <div class="question-text">${escapeHtml(q.question)}</div>
                ${isText ? `<input type="text" class="text-answer-input" id="textAnswer" value="${escapeHtml(saved||'')}">` :
                `<div class="answers-list">${q.shuffledAnswers.map(a => `
                    <div class="answer-item ${saved===a?'selected':''}" data-answer="${escapeHtml(a)}">
                        <span class="answer-radio"></span><span class="answer-text">${escapeHtml(a)}</span>
                    </div>`).join('')}</div>`}
            </div>
            <div class="test-navigation">
                ${currentIndex>0?'<button class="nav-test-btn" id="prevQuestion">← Назад</button>':'<div></div>'}
                ${currentIndex<questions.length-1?
                    '<button class="nav-test-btn primary" id="nextQuestion">Далее →</button>':
                    '<button class="nav-test-btn success" id="finishTest">Завершить</button>'}
            </div>
        </div>`;
    elements.contentContainer.innerHTML = html;
    if (test.time_limit > 0) startTimer(test.time_limit * 60);
    document.getElementById('exitTest').addEventListener('click', () => {
        showModal('Прогресс теста будет потерян. Выйти?',
            () => {
                clearInterval(timerInterval);
                timerInterval = null;
                navigateTo('tests');
            },
            () => {}
        );
    });
    if (!isText) {
        document.querySelectorAll('.answer-item').forEach(el => el.addEventListener('click', () => {
            const ans = el.dataset.answer;
            testState.answers[q.id] = ans;
            document.querySelectorAll('.answer-item').forEach(e => e.classList.toggle('selected', e.dataset.answer===ans));
        }));
    } else {
        document.getElementById('textAnswer').addEventListener('input', e => testState.answers[q.id] = e.target.value);
    }
    document.getElementById('prevQuestion')?.addEventListener('click', () => { testState.currentIndex--; renderQuestion(); });
    document.getElementById('nextQuestion')?.addEventListener('click', () => { testState.currentIndex++; renderQuestion(); });
    document.getElementById('finishTest')?.addEventListener('click', () => finishTest(false));
}

function startTimer(seconds) {
    const el = document.getElementById('timerDisplay');
    if (!el) return;
    let remaining = Math.max(0, seconds - Math.floor((Date.now()-testState.startTime)/1000));
    const update = () => {
        const m = Math.floor(remaining/60), s = remaining%60;
        el.textContent = `${m}:${s.toString().padStart(2,'0')}`;
        el.classList.toggle('danger', remaining<=60);
        el.classList.toggle('warning', remaining<=120 && remaining>60);
        if (remaining<=0) { clearInterval(timerInterval); finishTest(true); }
    };
    update();
    timerInterval = setInterval(() => { remaining--; update(); }, 1000);
}

function finishTest(isTimeout) {
    clearInterval(timerInterval);
    const { questions, answers } = testState;
    let correct = 0;
    questions.forEach(q => {
        const user = (answers[q.id] || '').trim().toLowerCase();
        if (!user) return;
        const correctAns = (q.answer1 || '').trim().toLowerCase();
        if (q.text_answer ? user === correctAns : user === correctAns) correct++;
    });
    const percent = Math.round(correct / questions.length * 100) || 0;
    let cls = 'bad'; if (percent >= 70) cls = 'good'; else if (percent >= 40) cls = 'medium';
    elements.contentContainer.innerHTML = `
        <div class="test-container">
            <button class="back-button" id="backToTests">← К тестам</button>
            <div class="result-card">
                <h2>Тест завершён${isTimeout?' (время вышло)':''}</h2>
                <div class="result-percent ${cls}">${percent}%</div>
                <div>${correct} из ${questions.length}</div>
                <button class="nav-test-btn primary" id="retakeTest">Пройти заново</button>
            </div>
        </div>`;
    document.getElementById('backToTests').addEventListener('click', () => navigateTo('tests'));
    document.getElementById('retakeTest').addEventListener('click', () => startTest(testState.test.id));
}