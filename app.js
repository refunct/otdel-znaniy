/* ===================== STATE ===================== */

const state = {
  data: null,
  currentTest: null,
  questions: [],
  answers: {},
  questionIndex: 0,
  showAnswers: false
}

let deferredPrompt = null

/* ===================== PWA INSTALL ===================== */

window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault()
  deferredPrompt = e

  const banner = document.getElementById('installBanner')
  if (banner && !localStorage.getItem('installed')) {
    banner.classList.remove('hidden')
  }
})

window.addEventListener('DOMContentLoaded', () => {
  const btn = document.getElementById('installBtn')

  if (btn) {
    btn.addEventListener('click', async () => {
      if (!deferredPrompt) return
      deferredPrompt.prompt()
      await deferredPrompt.userChoice
      localStorage.setItem('installed', '1')
    })
  }
})

/* ===================== LOAD EXCEL ===================== */

async function loadData() {
  try {
    const guideBuffer = await fetch('guide.xlsx').then(r => r.arrayBuffer())
    const testBuffer = await fetch('tests.xlsx').then(r => r.arrayBuffer())

    const wb1 = XLSX.read(guideBuffer)
    const wb2 = XLSX.read(testBuffer)

    state.data = {
      guides: XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[0]]),
      sections: XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[1]]),
      tests: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]]),
      questions: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[1]])
    }

  } catch (err) {
    console.error('Excel load error:', err)

    const app = document.getElementById('app')
    if (app) {
      app.innerHTML = `<div class="card">Ошибка загрузки данных</div>`
    }
  }
}

/* ===================== TEXT PARSER ===================== */

function parseText(text = "") {
  return text
    .replace(/<q>(.*?)<\/q>/g, '<div class="quote">$1</div>')
    .replace(/<note>(.*?)<\/note>/g, '<div class="note">$1</div>')
    .replace(/<warn>(.*?)<\/warn>/g, '<div class="warn">$1</div>')
}

/* ===================== ROUTER ===================== */

function navigate(page, id) {
  const url = id ? `?page=${page}&id=${id}` : `?page=${page}`
  history.pushState({}, '', url)
  render()
}

/* ===================== RENDER ===================== */

function render() {
  if (!state.data) return

  const params = new URLSearchParams(location.search)
  const page = params.get('page') || 'home'
  const id = params.get('id')

  const app = document.getElementById('app')
  if (!app) return

  app.innerHTML = ''
  app.classList.add('fade')

  /* HOME */
  if (page === 'home') {
    app.innerHTML = `<div class="card">Добро пожаловать в Отдел знаний</div>`
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

  /* GUIDE DETAIL */
  if (page === 'guide') {
    state.data.sections
      .filter(s => String(s['id справочника']) === String(id))
      .forEach(s => {
        app.innerHTML += `
          <div class="card">
            <h3>${s['название раздела']}</h3>
            ${parseText(s['текст раздела'] || '')}
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

/* ===================== TEST LOGIC ===================== */

function startTest(id) {
  state.currentTest = id
  state.answers = {}
  state.questionIndex = 0

  state.questions = state.data.questions.filter(
    q => String(q['id теста']) === String(id)
  )

  const meta = state.data.tests.find(t => String(t.id) === String(id))
  state.showAnswers = meta && Number(meta['показывать ответы']) === 1

  navigate('test', id)
}

function renderTest() {
  const app = document.getElementById('app')
  if (!app) return

  const q = state.questions[state.questionIndex]
  if (!q) {
    navigate('result')
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

  const progress = (state.questionIndex / state.questions.length) * 100

  app.innerHTML = `
    <div class="progress">
      <div style="width:${progress}%"></div>
    </div>

    <div class="card">
      <h3>${q['вопрос']}</h3>

      ${answers.map(a => `
        <div>
          <button onclick="answer('${String(a).replace(/'/g, "\\'")}')">
            ${a}
          </button>
        </div>
      `).join('')}

      ${q['текстовый ответ']
        ? `<input placeholder="Введите ответ" oninput="textAnswer(this.value)">`
        : ''
      }

      <br><br>

      <button onclick="prevQ()">Назад</button>
      <button onclick="nextQ()">Далее</button>
    </div>
  `
}

/* ===================== ANSWERS ===================== */

function answer(val) {
  state.answers[state.questionIndex] = val
}

function textAnswer(val) {
  state.answers[state.questionIndex] = val
}

function nextQ() {
  if (state.questionIndex < state.questions.length - 1) {
    state.questionIndex++
    renderTest()
  } else {
    navigate('result')
  }
}

function prevQ() {
  if (state.questionIndex > 0) {
    state.questionIndex--
    renderTest()
  }
}

/* ===================== RESULT ===================== */

function renderResult() {
  let correct = 0

  state.questions.forEach((q, i) => {
    const user = (state.answers[i] || '').toString().trim().toLowerCase()
    const right = (q['ответ 1'] || '').toString().trim().toLowerCase()

    if (user === right) correct++
  })

  const percent = Math.round((correct / state.questions.length) * 100)

  const app = document.getElementById('app')
  if (!app) return

  app.innerHTML = `
    <div class="card">
      <h2>Результат: ${percent}%</h2>
      <button onclick="navigate('tests')">К тестам</button>
    </div>
  `
}

/* ===================== SEARCH (FIXED) ===================== */

function searchQuery(value) {
  if (!state.data) return
  if (!value) return render()

  const v = value.toLowerCase()
  const app = document.getElementById('app')

  if (!app) return

  app.innerHTML = ''

  state.data.sections
    .filter(s =>
      (s['название раздела'] || '').toLowerCase().includes(v)
    )
    .forEach(s => {
      app.innerHTML += `
        <div class="card">
          ${s['название раздела']}
        </div>
      `
    })
}

/* ===================== EVENTS (FIXED DELEGATION) ===================== */

document.addEventListener('click', (e) => {
  const t = e.target

  if (t.dataset.nav) navigate(t.dataset.nav)
  if (t.dataset.guide) navigate('guide', t.dataset.guide)
  if (t.dataset.test) startTest(t.dataset.test)
})

document.addEventListener('input', (e) => {
  if (e.target.id === 'searchInput') {
    searchQuery(e.target.value)
  }
})

/* ===================== INIT ===================== */

(async () => {
  await loadData()
  render()
})()

/* ===================== SERVICE WORKER ===================== */

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js')
}