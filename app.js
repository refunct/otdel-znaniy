const state = {
  data: null,
  currentTest: null,
  answers: {},
  questionIndex: 0,
  questions: [],
  showAnswers: false
}

let deferredPrompt

window.addEventListener('beforeinstallprompt', e => {
  e.preventDefault()
  deferredPrompt = e
  if (!localStorage.getItem('installed')) {
    document.getElementById('installBanner').classList.remove('hidden')
  }
})

document.getElementById('installBtn').onclick = () => {
  if (!deferredPrompt) return
  deferredPrompt.prompt()
  deferredPrompt.userChoice.then(() => {
    localStorage.setItem('installed', '1')
  })
}

async function loadData() {
  try {
    const g = await fetch('guide.xlsx').then(r => r.arrayBuffer())
    const t = await fetch('tests.xlsx').then(r => r.arrayBuffer())

    const wb = XLSX.read(g)
    const wb2 = XLSX.read(t)

    state.data = {
      guides: XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]),
      sections: XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[1]]),
      tests: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]]),
      questions: XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[1]])
    }
  } catch (e) {
    console.error('Ошибка загрузки данных:', e)
    document.getElementById('app').innerHTML = 
      '<div class="card">Ошибка загрузки данных. Проверьте файлы guide.xlsx и tests.xlsx</div>'
  }
}