const excel4node = require('excel4node');

exports.rend = (year, month) => {
  const workBook = new excel4node.Workbook();
  const workSheet = workBook.addWorksheet('Sheet 1');
  writeWeekOfDaysToCalender(workSheet, new styleOfAllCells())
  const calender = makeCalenderArray(year, month)
  renderCalenderFromArray(workSheet, calender, year, month, workBook)
  return workBook
}

class styleOfAllCells {
  constructor() {
    this.border = {
      top: {
        style: "thick"
      },
      bottom: {
        style: "thick"
      },
      left: {
        style: "thick"
      },
      right: {
        style: "thick"
      }
    }
    this.alignment = {
      horizontal: 'center',
      vertical: 'center'
    }
  }
}

class holidaysCellsStyle extends styleOfAllCells {
  constructor(dayOfWeek) {
    let color = '#FFFFFF'
    super()
    if (dayOfWeek === 0) {
      color = '#FFCCFF'
    } else if (dayOfWeek === 6) {
      color = '#CCFFFF'
    }
    this.fill = {
      type: 'pattern',
      patternType: 'solid',
      fgColor: color
    }
  }
}

function renderCalenderFromArray(workSheet, array, y, m, workBook) {
  const stylePropsOfAllCells = new styleOfAllCells()
  const allCellsStyle = workBook.createStyle(stylePropsOfAllCells)
  array.forEach((week, s) => {
    week.forEach((day, o) => {
      if ((getDayOfWeek(y, m, day) === 0 || getDayOfWeek(y, m, day) === 6) && day != 0) {
        const dayOfWeek = getDayOfWeek(y, m, day)
        const stylePropsOfHolidayCells = new holidaysCellsStyle(dayOfWeek)
        const holidayCellsStyle = workBook.createStyle(stylePropsOfHolidayCells)
        writeNumberToCell(workSheet, ((s + 1) * 2) + 1, o + 1, day).style(holidayCellsStyle)
      } else {
        writeNumberToCell(workSheet, ((s + 1) * 2) + 1, o + 1, day).style(allCellsStyle)
      }
      if (day === 0) {
        workSheet.cell(((s + 1) * 2) + 1, o + 1)
          .string("")
      }
    })
  })

  for (let i = 0; i < array.length; i++) {
    for (let s = 1; s <= 7; s++) {
      workSheet.row((3 + (i * 2)) + 1).setHeight(80)
      writeStringToCell(workSheet, (3 + (i * 2)) + 1, s, "").style(allCellsStyle)
    }
  }
}


function writeNumberToCell(worksheet, x, y, number) {
  return worksheet.cell(x, y)
    .number(number)
}

function writeStringToCell(worksheet, x, y, string) {
  return worksheet.cell(x, y)
    .string(string)
}

function writeWeekOfDaysToCalender(workSheet, style) {
  const week = ['日', '月', '火', '水', '木', '金', '土']
  week.forEach((element, index) => {
    writeStringToCell(workSheet, 2, index + 1, element).style(style)
  })
}

function makeCalenderArray(y, m) {
  const monthTable = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
  if (((y % 4) == 0 && (y % 100) != 0) || (y % 400) == 0) {
    monthTable[1] = 29;
  }
  const dayOfWeek = getDayOfWeek(y, m, 1)
  const allWeeks = (dayOfWeek + monthTable[m - 1]) / 7
  const calenderFromWeeks = new Array(Math.ceil(allWeeks)).fill(0)
  calenderFromWeeks.forEach((_, i) => {
    calenderFromWeeks[i] = new Array(7).fill(0)
  })
  for (let i = dayOfWeek, day = 1; day <= monthTable[m - 1]; i++, day++) {
    const row = (i - (i % 7)) / 7
    const cell = (i - (7 * row))
    calenderFromWeeks[row][cell] = day
  }
  return calenderFromWeeks
}

function getDayOfWeek(y, m, d) {
  if (m <= 2) {
    y -= 1
    m += 12
  }
  let l
  const c = Math.floor(y / 100)
  if (1582 <= y) {
    l = ((5 * c) + Math.floor((c / 4)))
  } else if (4 <= y && y <= 1582) {
    l = ((6 * c) + 5)
  }
  y %= 100
  const dayOfWeek = (
    (d + Math.floor(((26 * (m + 1)) / 10)) + (y) + Math.floor((y / 4)) + (l)) % 7
  )
  if (dayOfWeek == 0) {
    return 6
  }
  return dayOfWeek - 1
}
