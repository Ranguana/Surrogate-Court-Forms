const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("auth", {
  login: (password) => ipcRenderer.invoke("login", password),
  setupPassword: (password) => ipcRenderer.invoke("setup-password", password),
});
