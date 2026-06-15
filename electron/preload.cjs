const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("excelAI", {
    matchExcelFields(payload) {
        return ipcRenderer.invoke("excel:match-fields", payload);
    },
});
