/* ===================== STATE ===================== */

const state = {
  data: null,
  currentTestId: null,
  questions: [],
  answers: {},
  index: 0
}

let installPrompt = null

/* ===================== PWA INSTALL ===================== */

window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault()
  installPrompt = e

  const banner = document.getElementById('installBanner')
  if (banner && !localStorage.getItem('installed')) {
    banner.classList.remove('hidden')
  }
})

document.addEventListener('click', (e) => {
  if (e.target.id === 'installBtn') {
    if (!installPrompt) return
    installPrompt.prompt()
    installPrompt.userChoice.then(() => {
      localStorage.setItem('installed', '1')
    })
  }
})

/* ===================== LOAD EXCEL ===================== */

async function loadExcel() {
  try {
    const guideBuf = await fetch('guide.xlsx').then(r => r.arrayBuffer())
    const testBuf = await fetch('tests.xlsx').then(r => r.arrayBuffer())

    const wb1 = XLSX.read(guideBuf)
    const wb2 = XLSX.read(testBuf)

    state.data = {
      guides: XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[0]]),
      sections: XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[1]]),
      tests: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]]),
      questions: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[1]])
    }

  } catch (e) {
    console.error('Excel error:', e)

    const app = document.getElementById('app')
    if (app) {
      app.innerHTML = `<div class="card">Ошибка загрузки данных</div>`
    }
  }
}

/* ===================== TEXT FORMAT ===================== */

function formatText(t = "") {
  return t
    .replace(/<q>(.*?)<\/q>/g, '<div class="quote">$1</div>')
    .replace(/<note>(.*?)<\/note>/g, '<div class="note">$1</div>')
    .replace(/<warn>(.*?)<\/warn>/g, '<div class="warn">$1</div>')
}

/* ===================== ROUTER ===================== */

function go(page, id) {
  const url = id ? `?page=${page}&id=${id}` : `?page=${page}`
  history.pushState({}, '', url)
  render()
}

/* ===================== RENDER ===================== */

function render() {
  if (!state.data) return

  const p = new URLSearchParams(location.search)
  const page = p.get('page') || 'home'
  const id = p.get('id')

  const app = document.getElementById('app')
  if (!app) return

  app.innerHTML = ''

  /* HOME */
  if (page === 'home') {
    app.innerHTML = `<div class="card">Отдел знаний</div>`
  }

  /* GUIDES */
  if (page === 'guides') {
    state.data.guides.forEach(g => {
      app.innerHTML += `
        <div class="card" data-guide="${g.id}">
          ${g['Название справочника']}
        </div>
      `
    })
  }

  /* GUIDE */
  if (page === 'guide') {
    state.data.sections
      .filter(s => String(s['id справочника']) === String(id))
      .forEach(s => {
        app.innerHTML += `
          <div class="card">
            <h3>${s['название раздела']}</h3>
            ${formatText(s['текст раздела'] || '')}
          </div>
        `
      })
  }

  /* TESTS */
  if (page === 'tests') {
    state.data.tests.forEach(t => {
      app.innerHTML += `
        <div class="card" data-test="${t.id}">
          ${t['название теста']}
        </div>
      `
    })
  }

  if (page === 'test') renderTest()
  if (page === 'result') renderResult()
}

/* ===================== TEST ===================== */

function startTest(id) {
  state.currentTestId = id
  state.answers = {}
  state.index = 0

  state.questions = state.data.questions.filter(
    q => String(q['id теста']) === String(id)
  )

  go('test', id)
}

function renderTest() {
  const q = state.questions[state.index]
  const app = document.getElementById('app')

  if (!q) {
    go('result')
    return
  }

  let answers = [
    q['ответ 1'],
    q['ответ 2'],
    q['ответ 3'],
    q['ответ 4'],
    q['ответ 5'],
    q['ответ 6']
  ].filter(Boolean)

  answers.sort(() => Math.random() - 0.5)

  const progress = (state.index / state.questions.length) * 100

  app.innerHTML = `
    <div class="progress">
      <div style="width:${progress}%"></div>
    </div>

    <div class="card">
      <h3>${q['вопрос']}</h3>

      ${answers.map(a => `
        <button onclick="selectAnswer('${String(a).replace(/'/g,"\\'")}')">
          ${a}
        </button>
      `).join('')}

      ${q['текстовый ответ'] ? `
        <input placeholder="Введите ответ" oninput="textAnswer(this.value)">
      ` : ''}

      <br><br>

      <button onclick="prev()">Назад</button>
      <button onclick="next()">Далее</button>
    </div>
  `
}

/* ===================== ANSWERS ===================== */

function selectAnswer(v) {
  state.answers[state.index] = v
}

function textAnswer(v) {
  state.answers[state.index] = v
}

function next() {
  if (state.index < state.questions.length - 1) {
    state.index++
    renderTest()
  } else {
    go('result')
  }
}

function prev() {
  if (state.index > 0) {
    state.index--
    renderTest()
  }
}

/* ===================== RESULT ===================== */

function renderResult() {
  let correct = 0

  state.questions.forEach((q, i) => {
    const user = (state.answers[i] || '').toLowerCase().trim()
    const right = (q['ответ 1'] || '').toLowerCase().trim()

    if (user === right) correct++
  })

  const percent = Math.round((correct / state.questions.length) * 100)

  document.getElementById('app').innerHTML = `
    <div class="card">
      <h2>Результат: ${percent}%</h2>
      <button onclick="go('tests')">Назад к тестам</button>
    </div>
  `
}

/* ===================== SEARCH (FIXED) ===================== */

document.addEventListener('input', (e) => {
  if (e.target.id === 'searchInput') {
    search(e.target.value)
  }
})

function search(v) {
  if (!state.data) return
  if (!v) return render()

  const app = document.getElementById('app')
  const q = v.toLowerCase()

  app.innerHTML = ''

  state.data.sections
    .filter(s => (s['название раздела'] || '').toLowerCase().includes(q))
    .forEach(s => {
      app.innerHTML += `
        <div class="card">${s['название раздела']}</div>
      `
    })
}

/* ===================== EVENTS ===================== */

document.addEventListener('click', (e) => {
  if (e.target.dataset.guide) go('guide', e.target.dataset.guide)
  if (e.target.dataset.test) startTest(e.target.dataset.test)
})

/* ===================== INIT ===================== */

(async function init() {
  await loadExcel()
  render()
})()

/* ===================== SERVICE WORKER ===================== */

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js')
}