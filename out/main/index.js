"use strict";
const electron = require("electron");
require("dotenv/config");
const path = require("path");
const utils = require("@electron-toolkit/utils");
const nodemailer = require("nodemailer");
const XLSX = require("xlsx");
const fs = require("node:fs/promises");
const path$1 = require("node:path");
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
function generateReport(report) {
  const wb = XLSX__namespace.utils.book_new();
  const mappedReport = report.map((r) => ({
    "Организация": r.name,
    "Дата отправки": r.date ? formatDate(new Date(r.date)) : "",
    "Email": r.email,
    "Контакты": r.contacts,
    "Статус": parseStatus(r.status, r.error)
  }));
  const ws = XLSX__namespace.utils.json_to_sheet(mappedReport);
  ws["!cols"] = [
    { wch: 30 },
    // Организация — Ширина 30, перенос
    { wch: 15 },
    // Дата отправки — Автоподбор
    { wch: 25 },
    // Email — Автоподбор
    { wch: 30 },
    // Контакты — Ширина 30, перенос
    { wch: 15 }
    // Статус — Ширина 15, перенос
  ];
  Object.keys(ws).forEach((cell) => {
    if (cell[0] !== "!") {
      ws[cell].s = {
        alignment: {
          wrapText: true,
          // Это включит перенос текста
          vertical: "center",
          horizontal: "left"
        }
      };
    }
  });
  mappedReport.forEach((row, index) => {
    const excelRow = index + 2;
    const status = row["Статус"];
    const range = `A${excelRow}:E${excelRow}`;
    if (status.includes("Ошибка")) {
      ws[range] = { s: { fill: { fgColor: { rgb: "FFCCCC" } } } };
    } else if (status === "Отправлено") {
      ws[range] = { s: { fill: { fgColor: { rgb: "CCFFCC" } } } };
    }
  });
  XLSX__namespace.utils.book_append_sheet(wb, ws, "Отчет");
  const file = path.join(electron.app.getPath("desktop"), `отчет_рассылки_${formatDateForFileName()}.xlsx`);
  XLSX__namespace.writeFile(wb, file);
  return file;
}
const parseStatus = (status, error) => {
  switch (status) {
    case "OK":
      return "Отправлено";
    case "FAIL":
      return `Ошибка: ${error ?? "Неизвестная ошибка"}`;
    case "VALID":
      return "Требуется проверка";
    default:
      return "Неизвестный статус";
  }
};
const formatDate = (date) => {
  return date.toLocaleString("ru-RU", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
};
const formatDateForFileName = () => {
  const now = /* @__PURE__ */ new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  String(now.getSeconds()).padStart(2, "0");
  return `${year}-${month}-${day}_${hours}-${minutes}`;
};
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
          report.push({ ...r, status: "OK", date: /* @__PURE__ */ new Date() });
        } catch (err) {
          if (err instanceof InvalidEmailError) pause = false;
          win.webContents.send("mail-progress", {
            ...r,
            status: "FAIL",
            error: err.message
          });
          report.push({ ...r, status: "FAIL", error: err.message, date: /* @__PURE__ */ new Date() });
        }
        if (pause)
          await new Promise((res) => setTimeout(res, rand(pauseMin, pauseMax)));
      }
      const file = generateReport(report);
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
