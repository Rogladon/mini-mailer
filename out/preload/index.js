"use strict";
const electron = require("electron");
const preload = require("@electron-toolkit/preload");
const api = {
  startMailing: (payload) => electron.ipcRenderer.invoke("start-mailing", payload),
  onMailProgress: (cb) => {
    const wrapper = (_, data) => cb(data);
    electron.ipcRenderer.on("mail-progress", wrapper);
    return () => electron.ipcRenderer.removeListener("mail-progress", wrapper);
  },
  getAccounts: () => electron.ipcRenderer.invoke("get-accounts")
};
if (process.contextIsolated) {
  try {
    electron.contextBridge.exposeInMainWorld("electron", preload.electronAPI);
    electron.contextBridge.exposeInMainWorld("electronAPI", api);
  } catch (error) {
    console.error(error);
  }
} else {
  window.electron = preload.electronAPI;
  window.api = api;
}
