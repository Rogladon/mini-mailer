"use strict";
const electron = require("electron");
require("dotenv/config");
const path = require("path");
const utils = require("@electron-toolkit/utils");
const nodemailer = require("nodemailer");
const XLSX = require("xlsx");
const path$1 = require("node:path");
const fs = require("node:fs/promises");
function _interopNamespaceDefault(e) {
  const n = Object.create(null, { [Symbol.toStringTag]: { value: "Module" } });
  if (e) {
    for (const k in e) {
      if (k !== "default") {
        const d = Object.getOwnPropertyDescriptor(e, k);
        Object.defineProperty(n, k, d.get ? d : {
          enumerable: true,
          get: () => e[k]
        });
      }
    }
  }
  n.default = e;
  return Object.freeze(n);
}
const XLSX__namespace = /* @__PURE__ */ _interopNamespaceDefault(XLSX);
const icon = path.join(__dirname, "./chunks/icon-BE0e6We9.png");
const EMAIL_RE = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
function extractEmail(raw) {
  if (typeof raw !== "string") return null;
  const m = raw.match(EMAIL_RE);
  return m?.[0]?.trim() ?? null;
}
const tpl = (s, vars) => s.replace(/\{\{(\w+)\}\}/g, (_, k) => vars[k] ?? "");
const rand = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
class InvalidEmailError extends Error {
  constructor() {
    super("Invalid email");
  }
}
function initMailer() {
  electron.ipcMain.handle(
    "start-mailing",
    async (e, {
      smtp,
      recipients,
      subjectTemplate,
      htmlTemplate,
      pauseMin,
      pauseMax
    }) => {
      const win = electron.BrowserWindow.fromWebContents(e.sender);
      const transport = nodemailer.createTransport({
        host: smtp.host,
        port: smtp.port,
        secure: smtp.secure,
        auth: { user: smtp.user, pass: smtp.pass }
      });
      const report = [];
      for (const r of recipients) {
        const vars = { name: r.name };
        let pause = true;
        try {
          if (!extractEmail(r.email)) throw new InvalidEmailError();
          await transport.sendMail({
            from: smtp.user,
            to: r.email,
            subject: tpl(subjectTemplate, vars),
            html: tpl(htmlTemplate, vars)
          });
          win.webContents.send("mail-progress", { ...r, status: "OK" });
          report.push({ ...r, status: "OK" });
        } catch (err) {
          if (err instanceof InvalidEmailError) pause = false;
          win.webContents.send("mail-progress", {
            ...r,
            status: "FAIL",
            error: err.message
          });
          report.push({ ...r, status: "FAIL", error: err.message });
        }
        if (pause)
          await new Promise((res) => setTimeout(res, rand(pauseMin, pauseMax)));
      }
      const wb = XLSX__namespace.utils.book_new();
      XLSX__namespace.utils.book_append_sheet(wb, XLSX__namespace.utils.json_to_sheet(report), "Report");
      const file = path$1.join(electron.app.getPath("desktop"), "report.xlsx");
      XLSX__namespace.writeFile(wb, file);
      return { file };
    }
  );
}
const cfgPath = electron.app.isPackaged ? path$1.join(path$1.dirname(electron.app.getPath("exe")), "accounts.json") : path$1.join(electron.app.getAppPath(), "/resources/accounts.json");
function registerAccountsLoader() {
  electron.ipcMain.handle("get-accounts", async () => {
    try {
      const raw = await fs.readFile(cfgPath, "utf-8");
      return JSON.parse(raw);
    } catch {
      return [];
    }
  });
}
function createWindow() {
  const mainWindow = new electron.BrowserWindow({
    width: 900,
    height: 670,
    show: false,
    autoHideMenuBar: true,
    ...process.platform === "linux" ? { icon } : {},
    webPreferences: {
      preload: path.join(__dirname, "../preload/index.js"),
      sandbox: false
    }
  });
  mainWindow.on("ready-to-show", () => {
    mainWindow.show();
  });
  mainWindow.webContents.setWindowOpenHandler((details) => {
    electron.shell.openExternal(details.url);
    return { action: "deny" };
  });
  if (utils.is.dev && process.env["ELECTRON_RENDERER_URL"]) {
    mainWindow.loadURL(process.env["ELECTRON_RENDERER_URL"]);
  } else {
    mainWindow.loadFile(path.join(__dirname, "../renderer/index.html"));
  }
}
electron.app.whenReady().then(() => {
  utils.electronApp.setAppUserModelId("com.electron");
  electron.app.on("browser-window-created", (_, window) => {
    utils.optimizer.watchWindowShortcuts(window);
  });
  electron.ipcMain.on("ping", () => console.log("pong"));
  createWindow();
  initMailer();
  registerAccountsLoader();
  electron.app.on("activate", function() {
    if (electron.BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});
electron.app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    electron.app.quit();
  }
});
