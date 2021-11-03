const outputButton = document.getElementById('outputButton')
const yearInput = document.getElementById('year')
const monthInput = document.getElementById('month')
const now = new Date()
yearInput.placeholder = now.getFullYear()
monthInput.placeholder = now.getMonth() + 1
outputButton.addEventListener('click', () => {
  if (yearInput.value === "" || monthInput === "") {
    return
  }
  const year = Number(yearInput.value)
  const month = Number(monthInput.value)
  if (year < 4) {
    window.api.invalidYearAlert()
    return
  }
  if (0 < year && 0 < month && month <= 12 && 4 <= year) {
    console.log("start");
    window.api.rendExcelCalender(year, month)
    console.log("Done");
  }
})

yearInput.addEventListener('keyup', (e) => {
  if (0 < e.target.value && 0 < monthInput.value && monthInput.value <= 12 && e.target.value[0] != 0 && 4 <= e.target.value) {
    enableButton()
  } else {
    disabelButton()
  }
})

monthInput.addEventListener('keyup', (e) => {
  if (0 < e.target.value && e.target.value <= 12 && 0 < yearInput.value && yearInput.value[0] != 0 && 4 <= yearInput.value) {
    enableButton()
  } else {
    disabelButton()
  }
})

function changeColor(element, color) {
  element.style.background = color
}

function disabelButton() {
  changeColor(outputButton, "#70A5F5")
  outputButton.style.cursor = "default"
}

function enableButton() {
  changeColor(outputButton, "#0066FF")
  outputButton.style.cursor = "pointer"
}
