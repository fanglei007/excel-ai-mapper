import { app, BrowserWindow, ipcMain } from "electron";
import path from "node:path";
import { fileURLToPath } from "node:url";
import matchExcelFields from "./doubao.js";

const currentFilePath = fileURLToPath(import.meta.url);
const currentDirectory = path.dirname(currentFilePath);

function createWindow() {
    const mainWindow = new BrowserWindow({
        width: 1100,
        height: 760,
        minWidth: 800,
        minHeight: 600,
        webPreferences: {
            preload: path.join(currentDirectory, "preload.cjs"),
            contextIsolation: true,
            nodeIntegration: false,
        },
    });

    if (app.isPackaged) {
        mainWindow.loadFile(
            path.join(currentDirectory, "../dist/index.html"),
        );
    } else {
        mainWindow.loadURL("http://127.0.0.1:5173");
        mainWindow.webContents.openDevTools();
    }
}

ipcMain.handle("excel:match-fields", async (_event, payload) => {
    return matchExcelFields(payload);
});

app.whenReady().then(() => {
    createWindow();

    app.on("activate", () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on("window-all-closed", () => {
    if (process.platform !== "darwin") {
        app.quit();
    }
});
