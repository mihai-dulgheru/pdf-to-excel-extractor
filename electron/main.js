const {app, BrowserWindow, ipcMain, dialog} = require("electron");
const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');
const path = require("path");

let mainWindow;

app.on("ready", () => {
    mainWindow = new BrowserWindow({
        width: 800, height: 600, webPreferences: {
            preload: path.join(__dirname, "preload.js"), nodeIntegration: false, contextIsolation: true,
        },
    });

    mainWindow.loadFile(path.join(__dirname, "renderer", "index.html")).then(() => {
        console.log("Main window loaded successfully.");
    }).catch((err) => {
        console.error("Failed to load main window:", err);
    });
});

ipcMain.handle("upload-files", async (event, {files, percentage}) => {
    const tempDir = path.join(app.getPath("temp"), "pdf-to-excel");
    if (!fs.existsSync(tempDir)) {
        fs.mkdirSync(tempDir, {recursive: true});
    }

    try {
        const inputPaths = files.map((file) => {
            const tempFilePath = path.join(tempDir, file.name);
            fs.writeFileSync(tempFilePath, Buffer.from(file.content));
            return tempFilePath;
        });

        const data = new FormData();
        inputPaths.forEach((filePath) => {
            data.append("files", fs.createReadStream(filePath));
        });
        data.append("percentage", percentage);

        const response = await axios.post("http://127.0.0.1:8000/process-invoices/", data, {
            headers: {
                'Content-Type': 'multipart/form-data', ...data.getHeaders()
            }, responseType: "arraybuffer",
        });

        const contentDisposition = response.headers["content-disposition"];
        const filename = contentDisposition ? contentDisposition.match(/filename="(.+?)"/)[1] : "Processed-Invoices.xlsx";

        const outputPath = path.join(tempDir, filename);
        fs.writeFileSync(outputPath, Buffer.from(response.data));

        inputPaths.forEach((filePath) => fs.unlinkSync(filePath));

        return outputPath;
    } catch (error) {
        console.error("Error:", error.message);
        console.error("Response:", error.response?.data);
        throw new Error(error.response?.data?.detail || "Failed to process files.");
    }
});

ipcMain.handle("get-downloads-path", async () => {
    return app.getPath("downloads");
});

ipcMain.handle("show-save-dialog", async (event, options) => {
    return await dialog.showSaveDialog(BrowserWindow.getFocusedWindow(), options);
});

ipcMain.handle("save-file", async (event, {sourcePath, destinationPath}) => {
    fs.copyFileSync(sourcePath, destinationPath);
});
