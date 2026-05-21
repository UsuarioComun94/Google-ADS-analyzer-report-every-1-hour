const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("hma", {
  version: "0.2.0",
  runAction: (payload) => ipcRenderer.invoke("hma:run-action", payload),
  getBaseDir: () => ipcRenderer.invoke("hma:get-base-dir"),
  getSystemStatus: () => ipcRenderer.invoke("hma:get-system-status")
});
