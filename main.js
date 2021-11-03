const electron = require('electron');
const excelCalender = require('./src/main/rendCalender')
const fs = require('fs');
const dialog = require('electron').dialog
const app = electron.app
const browserWindow = electron.BrowserWindow
const ipcMain = electron.ipcMain

let mainWindow = null

app.on('ready', () => {
  mainWindow = new browserWindow({
    width: 1280,
    height: 720,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: __dirname + "\\src\\main\\preload.js"
    },
    icon: __dirname + '/Icon.ico'
  })
  // mainWindow.openDevTools()
  mainWindow.setMenu(null)
  mainWindow.setIcon(__dirname + '/Icon.ico')
  mainWindow.loadFile(__dirname + "/src/renderer/main/index.html")
  mainWindow.on('close', () => {
    mainWindow = null
  })
})

ipcMain.on('rendExcelCalender', (e, arg) => {
  dialog.showSaveDialog(mainWindow, {
    title: "保存",
    defaultPath: `${process.env.USERPROFILE}\\Desktop\\`,
    buttonLabel: "保存",
    filters: [{
      name: 'XLSXファイル',
      extensions: ['xlsx']
    }, {
      name: 'All files',
      extensions: ['*']
    }]
  }).then((e) => {
    const filePath = e.filePath
    if (e.canceled) {
      return
    }
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath)
    }
    excelCalender.rend(arg.year, arg.month).write(filePath)
  }).catch(error => {
    if (error.errno === -4082) {
      dialog.showMessageBox(mainWindow, {
        type: "error",
        title: "エラー",
        message: `このファイルは他のプロセスが使用中です.\nCODE:${error.code}\nPATH:${error.path}`
      })
    }
    console.log(error);
  })
})

ipcMain.on('invalidYear', () => {
  dialog.showMessageBox(mainWindow, {
    title: ":)",
    message: "西暦３年以前のカレンダーを作成することは出来ません。タイムマシンがあるといいですね。"
  })
})
