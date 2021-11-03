const ipcRenderer = require('electron').ipcRenderer
const contextBridge = require('electron').contextBridge

contextBridge.exposeInMainWorld("api", {
  rendExcelCalender: (year, month) => {
    ipcRenderer.send('rendExcelCalender', {
      year: year,
      month: month
    })
  },
  invalidYearAlert: () => {
    ipcRenderer.send('invalidYear', "")
  }
})
