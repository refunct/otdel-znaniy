
/* ===================== STATE ===================== */

const state = {
  data: null,
  page: 'guides', // 👈 ГЛАВНАЯ СТРАНИЦА = справочники
  id: null,
  testId: null,
  questions: [],
  index: 0,
  answers: {}
}

/* ===================== INIT ===================== */

window.addEventListener('DOMContentLoaded', init)

async function init() {
  bindEvents()
  await loadData()
  router(true)
}

/* ===================== EVENTS ===================== */

function bindEvents() {
  document.addEventListener('click', handleClick)
  document.addEventListener('input', handleInput)
  window.addEventListener('popstate', () => router(false))
}

function handleClick(e) {
  const t = e.target

  // Навигация
  if (t.dataset.nav) go(t.dataset.nav)
  if (t.dataset.guide) go('guide', t.dataset.guide)
  if (t.dataset.test) startTest(t.dataset.test)

  // тест
  if (t.id === 'nextBtn') nextQuestion()
  if (t.id === 'prevBtn') prevQuestion()

  // очистка кеша вручную (скрытая кнопка, если понадобится)
  if (t.id === 'clearCache') clearCache()
}

function handleInput(e) {
  if (e.target.id === 'searchInput') {
    search(e.target.value)
  }
}

/* ===================== ROUTER ===================== */

function go(page, id = null) {
  state.page = page
  state.id = id

  const url = id ? `?page=${page}&id=${id}` : `?page=${page}`
  history.pushState({}, '', url)

  render()
}

function router(shouldReplace = false) {
  const p = new URLSearchParams(location.search)

  state.page = p.get('page') || 'guides' // 👈 default = справочники
  state.id = p.get('id')

  render()
}

/* ===================== LOAD DATA ===================== */

async function loadData() {
  try {
    const guide = await fetch('guide.xlsx', { cache: "no-store" }).then(r => r.arrayBuffer())
    const test = await fetch('tests.xlsx', { cache: "no-store" }).then(r => r.arrayBuffer())

    const wb1 = XLSX.read(guide)
    const wb2 = XLSX.read(test)

    state.data = {
      guides: XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[0]]),
      sections: XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[1]]),
      tests: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]]),
      questions: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[1]])
    }

  } catch (e) {
    console.error(e)
    mount(`<div class="card">Ошибка загрузки данных</div>`)
  }
}

/* ===================== RENDER ===================== */

function render() {
  if (!state.data) return

  switch (state.page) {

    case 'guides':
      return renderGuides()

    case 'guide':
      return renderGuide()

    case 'tests':
      return renderTests()

    case 'test':
      return renderTest()

    case 'result':
      return renderResult()

    default:
      return renderGuides()
  }
}

/* ===================== UI MOUNT ===================== */

function mount(html) {
  const app = document.getElementById('app')
  if (!app) return
  app.innerHTML = html
}

/* ===================== HEADER UI FIX ===================== */

function header() {
  return `
    <div class="nav">
      <button data-nav="guides">Справочники</button>
      <button data-nav="tests">Тесты</button>
    </div>

    <input id="searchInput" placeholder="Поиск..." />
  `
}

/* ===================== GUIDES ===================== */

function renderGuides() {
  mount(`
    ${header()}

    ${state.data.guides.map(g => `
      <div class="card" data-guide="${g.id}">
        ${g['Название справочника']}
      </div>
    `).join('')}
  `)
}

/* ===================== GUIDE ===================== */

function renderGuide() {
  const items = state.data.sections.filter(
    s => String(s['id справочника']) === String(state.id)
  )

  mount(`
    ${header()}

    ${items.map(s => `
      <div class="card">
        <h3>${s['название раздела']}</h3>
        ${format(s['текст раздела'] || '')}
      </div>
    `).join('')}
  `)
}

/* ===================== TESTS ===================== */

function renderTests() {
  mount(`
    ${header()}

    ${state.data.tests.map(t => `
      <div class="card" data-test="${t.id}">
        ${t['название теста']}
      </div>
    `).join('')}
  `)
}

/* ===================== TEST ENGINE ===================== */

function startTest(id) {
  state.testId = id
  state.index = 0
  state.answers = {}

  state.questions = state.data.questions.filter(
    q => String(q['id теста']) === String(id)
  )

  go('test', id)
}

function renderTest() {
  const q = state.questions[state.index]
  if (!q) return go('result')

  let answers = [
    q['ответ 1'],
    q['ответ 2'],
    q['ответ 3'],
    q['ответ 4'],
    q['ответ 5'],
    q['ответ 6']
  ].filter(Boolean)

  answers.sort(() => Math.random() - 0.5)

  mount(`
    ${header()}

    <div class="card">
      <h3>${q['вопрос']}</h3>

      ${answers.map(a => `
        <button data-answer="${a}">${a}</button>
      `).join('')}

      ${q['текстовый ответ'] ? `
        <input id="textAnswer" placeholder="Ответ">
      ` : ''}

      <br><br>

      <button id="prevBtn">Назад</button>
      <button id="nextBtn">Далее</button>
    </div>
  `)
}

/* ===================== ANSWERS ===================== */

document.addEventListener('click', (e) => {
  if (e.target.dataset.answer) {
    state.answers[state.index] = e.target.dataset.answer
  }
})

function nextQuestion() {
  if (state.index < state.questions.length - 1) {
    state.index++
    renderTest()
  } else {
    go('result')
  }
}

function prevQuestion() {
  if (state.index > 0) {
    state.index--
    renderTest()
  }
}

/* ===================== RESULT ===================== */

function renderResult() {
  let correct = 0

  state.questions.forEach((q, i) => {
    const u = (state.answers[i] || '').toLowerCase().trim()
    const r = (q['ответ 1'] || '').toLowerCase().trim()
    if (u === r) correct++
  })

  const percent = Math.round((correct / state.questions.length) * 100)

  mount(`
    ${header()}

    <div class="card">
      <h2>Результат: ${percent}%</h2>
      <button data-nav="tests">К тестам</button>
    </div>
  `)
}

/* ===================== SEARCH ===================== */

function search(v) {
  if (!v) return render()

  const q = v.toLowerCase()

  const filtered = state.data.sections.filter(s =>
    (s['название раздела'] || '').toLowerCase().includes(q)
  )

  mount(`
    ${header()}

    ${filtered.map(s => `
      <div class="card">${s['название раздела']}</div>
    `).join('')}
  `)
}

/* ===================== FORMAT ===================== */

function format(t = '') {
  return t
    .replace(/<q>(.*?)<\/q>/g, `<div class="quote">$1</div>`)
    .replace(/<note>(.*?)<\/note>/g, `<div class="note">$1</div>`)
    .replace(/<warn>(.*?)<\/warn>/g, `<div class="warn">$1</div>`)
}

/* ===================== CACHE CONTROL ===================== */

function clearCache() {
  if ('caches' in window) {
    caches.keys().then(keys => {
      keys.forEach(k => caches.delete(k))
    })
  }
}

/* ===================== SW ===================== */

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js')
}