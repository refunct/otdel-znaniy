/* ===================== STATE ===================== */

const state = {
  data: null,
  page: 'home',
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
  router()
}

/* ===================== EVENTS ===================== */

function bindEvents() {
  document.addEventListener('click', handleClick)
  document.addEventListener('input', handleInput)
  window.addEventListener('popstate', router)
}

function handleClick(e) {
  const t = e.target

  if (t.dataset.page) {
    go(t.dataset.page)
  }

  if (t.dataset.guide) {
    go('guide', t.dataset.guide)
  }

  if (t.dataset.test) {
    startTest(t.dataset.test)
  }

  if (t.id === 'nextBtn') nextQuestion()
  if (t.id === 'prevBtn') prevQuestion()
}

/* ===================== INPUT ===================== */

function handleInput(e) {
  if (e.target.id === 'searchInput') {
    search(e.target.value)
  }
}

/* ===================== ROUTER ===================== */

function go(page, id = null) {
  const url = id ? `?page=${page}&id=${id}` : `?page=${page}`
  history.pushState({}, '', url)
  router()
}

function router() {
  const p = new URLSearchParams(location.search)

  state.page = p.get('page') || 'home'
  state.id = p.get('id')

  render()
}

/* ===================== DATA ===================== */

async function loadData() {
  try {
    const g = await fetch('guide.xlsx').then(r => r.arrayBuffer())
    const t = await fetch('tests.xlsx').then(r => r.arrayBuffer())

    const wb1 = XLSX.read(g)
    const wb2 = XLSX.read(t)

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
    case 'home':
      return mount(`<div class="card">Отдел знаний</div>`)

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
  }
}

function mount(html) {
  const app = document.getElementById('app')
  if (!app) return
  app.innerHTML = html
}

/* ===================== GUIDES ===================== */

function renderGuides() {
  mount(`
    ${state.data.guides.map(g => `
      <div class="card" data-guide="${g.id}">
        ${g['Название справочника']}
      </div>
    `).join('')}
  `)
}

function renderGuide() {
  const items = state.data.sections.filter(
    s => String(s['id справочника']) === String(state.id)
  )

  mount(`
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
    <div class="card">
      <h2>Результат: ${percent}%</h2>
      <button data-page="tests">К тестам</button>
    </div>
  `)
}

/* ===================== SEARCH ===================== */

function search(value) {
  if (!value) return render()

  const v = value.toLowerCase()

  const filtered = state.data.sections.filter(s =>
    (s['название раздела'] || '').toLowerCase().includes(v)
  )

  mount(`
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

/* ===================== SW ===================== */

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js')
}