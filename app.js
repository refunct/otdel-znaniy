const state = {
  data: null,
  currentTest: null,
  questions: [],
  answers: {},
  questionIndex: 0,
  showAnswers: false
}

let deferredPrompt

// PWA install
window.addEventListener('beforeinstallprompt', e => {
  e.preventDefault()
  deferredPrompt = e
  if (!localStorage.getItem('installed')) {
    document.getElementById('installBanner').classList.remove('hidden')
  }
})

document.getElementById('installBtn').onclick = async () => {
  if (!deferredPrompt) return
  deferredPrompt.prompt()
  await deferredPrompt.userChoice
  localStorage.setItem('installed','1')
}

// Load Excel
async function loadData(){
  try{
    const g = await fetch('guide.xlsx').then(r=>r.arrayBuffer())
    const t = await fetch('tests.xlsx').then(r=>r.arrayBuffer())

    const wb = XLSX.read(g)
    const wb2 = XLSX.read(t)

    state.data = {
      guides: XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]),
      sections: XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[1]]),
      tests: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]]),
      questions: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[1]])
    }
  } catch(err){
    console.error(err)
    document.getElementById('app').innerHTML = '<div class="card">Ошибка загрузки данных</div>'
  }
}
// Text parser
function parseText(t=""){
  return t
    .replace(/<q>(.*?)<\/q>/g,'<div class="quote">$1</div>')
    .replace(/<note>(.*?)<\/note>/g,'<div class="note">$1</div>')
    .replace(/<warn>(.*?)<\/warn>/g,'<div class="warn">$1</div>')
}

// Router
function navigate(page,id){
  const url = id ? `?page=${page}&id=${id}` : `?page=${page}`
  history.pushState({},'',url)
  render()
}

// Render
function render(){
  if(!state.data) return

  const p = new URLSearchParams(location.search)
  const page = p.get('page') || 'home'
  const id = p.get('id')

  const app = document.getElementById('app')
  app.classList.add('fade')
  app.innerHTML = ''

  if(page==='home'){
    app.innerHTML = `<div class='card'>Добро пожаловать в Отдел знаний</div>`
  }

  if(page==='guides'){
    state.data.guides.forEach(g=>{
      app.innerHTML += `<div class='card' data-guide='${g.id}'>${g['Название справочника']}</div>`
    })
  }

  if(page==='guide'){
    state.data.sections
      .filter(s=>String(s['id справочника'])===String(id))
      .forEach(s=>{
        app.innerHTML += `<div class='card'><h3>${s['название раздела']}</h3>${parseText(s['текст раздела']||'')}</div>`
      })
  }

  if(page==='tests'){
    state.data.tests.forEach(t=>{
      app.innerHTML += `<div class='card' data-test='${t.id}'>${t['название теста']}</div>`
    })
  }

  if(page==='test') renderTest()
  if(page==='result') renderResult()
}
// Test start
function startTest(id){
  state.currentTest = id
  state.answers = {}
  state.questionIndex = 0

  state.questions = state.data.questions.filter(q=>String(q['id теста'])===String(id))

  const meta = state.data.tests.find(t=>String(t.id)===String(id))
  state.showAnswers = meta && Number(meta['показывать ответы'])===1

  navigate('test',id)
}
// Test render
function renderTest(){
  const q = state.questions[state.questionIndex]
  const app = document.getElementById('app')

  if(!q){ navigate('result'); return }

  let answers = [q['ответ 1'],q['ответ 2'],q['ответ 3'],q['ответ 4'],q['ответ 5'],q['ответ 6']].filter(Boolean)
  answers = answers.sort(()=>Math.random()-0.5)

  const progress = (state.questionIndex/state.questions.length)*100

  app.innerHTML = `
    <div class='progress'><div style='width:${progress}%'></div></div>
    <div class='card'>
      <h3>${q['вопрос']}</h3>
      ${answers.map(a=>
        `<div><button onclick="answer('${String(a).replace(/'/g,"\\'")}')">${a}</button></div>`
      ).join('')}

      ${q['текстовый ответ'] ? `<input placeholder='Введите ответ' oninput="textAnswer(this.value)">` : ''}

      <br><br>
      <button onclick="prevQ()">Назад</button>
      <button onclick="nextQ()">Далее</button>
    </div>
  `
}
function answer(a){ state.answers[state.questionIndex]=a }
function textAnswer(v){ state.answers[state.questionIndex]=v }

function nextQ(){
  if(state.questionIndex < state.questions.length-1){
    state.questionIndex++
    renderTest()
  } else navigate('result')
}

function prevQ(){
  if(state.questionIndex>0){
    state.questionIndex--
    renderTest()
  }
}
// Result
function renderResult(){
  let correct=0

  state.questions.forEach((q,i)=>{
    const u=(state.answers[i]||'').toString().toLowerCase().trim()
    const r=(q['ответ 1']||'').toString().toLowerCase().trim()
    if(u===r) correct++
  })

  const percent = Math.round((correct/state.questions.length)*100)

  document.getElementById('app').innerHTML = `
    <div class='card'>
      <h2>Результат: ${percent}%</h2>
      <button onclick="navigate('tests')">К тестам</button>
    </div>
  `
}
// Search
function search(q){
  if(!state.data) return
  if(!q){ render(); return }

  const v=q.toLowerCase()
  const app=document.getElementById('app')

  app.innerHTML=''

  state.data.sections
    .filter(s=>(s['название раздела']||'').toLowerCase().includes(v))
    .forEach(s=>{
      app.innerHTML += `<div class='card'>${s['название раздела']}</div>`
    })
}
// Events
window.addEventListener('click',e=>{
  const t=e.target

  if(t.dataset.nav) navigate(t.dataset.nav)
  if(t.dataset.guide) navigate('guide',t.dataset.guide)
  if(t.dataset.test) startTest(t.dataset.test)
})

document.getElementById('searchInput').addEventListener('input',e=>search(e.target.value))

// Init
(async()=>{
  await loadData()
  render()
})()

// SW
if('serviceWorker' in navigator){
  navigator.serviceWorker.register('sw.js')
}