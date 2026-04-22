const state = {
  data: null,
  page: "guides",
  id: null,
  testId: null,
  questions: [],
  index: 0,
  answers: {}
}

/* ================= INIT ================= */

window.addEventListener("DOMContentLoaded", init)

async function init(){
  bind()
  await load()
  route()
}

/* ================= EVENTS ================= */

function bind(){
  document.addEventListener("click", onClick)
  document.addEventListener("input", onInput)
  window.addEventListener("popstate", route)
}

function onClick(e){
  const t = e.target

  if(t.dataset.page) go(t.dataset.page)
  if(t.dataset.guide) go("guide", t.dataset.guide)
  if(t.dataset.test) startTest(t.dataset.test)

  if(t.id === "next") next()
  if(t.id === "prev") prev()

  if(t.dataset.answer){
    state.answers[state.index] = t.dataset.answer
  }
}

function onInput(e){
  if(e.target.id === "searchInput"){
    search(e.target.value)
  }
}

/* ================= ROUTER ================= */

function go(page,id=null){
  const url = id ? `?page=${page}&id=${id}` : `?page=${page}`
  history.pushState({}, "", url)

  state.page = page
  state.id = id

  render()
}

function route(){
  const p = new URLSearchParams(location.search)

  state.page = p.get("page") || "guides"
  state.id = p.get("id")

  render()
}

/* ================= DATA ================= */

async function load(){
  const g = await fetch("guide.xlsx").then(r=>r.arrayBuffer())
  const t = await fetch("tests.xlsx").then(r=>r.arrayBuffer())

  const wb1 = XLSX.read(g)
  const wb2 = XLSX.read(t)

  state.data = {
    guides: XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[0]]),
    sections: XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[1]]),
    tests: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]]),
    questions: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[1]])
  }
}

/* ================= RENDER ================= */

function render(){
  if(!state.data) return

  if(state.page === "guides") return renderGuides()
  if(state.page === "guide") return renderGuide()
  if(state.page === "tests") return renderTests()
  if(state.page === "test") return renderTest()
  if(state.page === "result") return result()
}

/* ================= UI ================= */

function app(html){
  document.getElementById("app").innerHTML = html
}

/* ================= GUIDES ================= */

function renderGuides(){
  app(`
    ${state.data.guides.map(g=>`
      <div class="card" data-guide="${g.id}">
        ${g["Название справочника"]}
      </div>
    `).join("")}
  `)
}

/* ================= GUIDE ================= */

function renderGuide(){
  const items = state.data.sections.filter(
    s => String(s["id справочника"]) === String(state.id)
  )

  app(`
    ${items.map(s=>`
      <div class="card">
        <b>${s["название раздела"]}</b><br><br>
        ${s["текст раздела"] || ""}
      </div>
    `).join("")}
  `)
}

/* ================= TESTS ================= */

function renderTests(){
  app(`
    ${state.data.tests.map(t=>`
      <div class="card" data-test="${t.id}">
        ${t["название теста"]}
      </div>
    `).join("")}
  `)
}

/* ================= TEST ================= */

function startTest(id){
  state.testId = id
  state.index = 0
  state.answers = {}

  state.questions = state.data.questions.filter(
    q => String(q["id теста"]) === String(id)
  )

  go("test",id)
}

function renderTest(){
  const q = state.questions[state.index]
  if(!q) return go("result")

  const answers = [
    q["ответ 1"],
    q["ответ 2"],
    q["ответ 3"],
    q["ответ 4"],
    q["ответ 5"],
    q["ответ 6"]
  ].filter(Boolean).sort(()=>Math.random()-0.5)

  const progress = (state.index/state.questions.length)*100

  app(`
    <div class="progress"><div style="width:${progress}%"></div></div>

    <div class="card">
      <b>${q["вопрос"]}</b><br><br>

      ${answers.map(a=>`
        <button class="answer" data-answer="${a}">${a}</button>
      `).join("")}

      <br>

      <button id="prev">Назад</button>
      <button id="next">Далее</button>
    </div>
  `)
}

function next(){
  if(state.index < state.questions.length-1){
    state.index++
    renderTest()
  } else go("result")
}

function prev(){
  if(state.index>0){
    state.index--
    renderTest()
  }
}

/* ================= RESULT ================= */

function result(){
  let correct = 0

  state.questions.forEach((q,i)=>{
    const u = (state.answers[i]||"").toLowerCase()
    const r = (q["ответ 1"]||"").toLowerCase()
    if(u===r) correct++
  })

  const percent = Math.round(correct/state.questions.length*100)

  app(`
    <div class="card">
      <h2>Результат: ${percent}%</h2>
      <button data-page="tests">Назад</button>
    </div>
  `)
}

/* ================= SEARCH ================= */

function search(v){
  if(!v) return render()

  const q = v.toLowerCase()

  const res = state.data.sections.filter(
    s => (s["название раздела"]||"").toLowerCase().includes(q)
  )

  app(`
    ${res.map(s=>`
      <div class="card">${s["название раздела"]}</div>
    `).join("")}
  `)
}

/* ================= SW ================= */

if("serviceWorker" in navigator){
  navigator.serviceWorker.register("sw.js")
}