const {contextBridge, ipcRenderer} = require("electron");

contextBridge.exposeInMainWorld("api", {
    uploadFiles: (files, percentage) => ipcRenderer.invoke("upload-files", {files, percentage}),
    getDownloadsPath: () => ipcRenderer.invoke("get-downloads-path"),
    showSaveDialog: (options) => ipcRenderer.invoke("show-save-dialog", options),
    saveFile: (sourcePath, destinationPath) => ipcRenderer.invoke("save-file", {sourcePath, destinationPath}),
});
