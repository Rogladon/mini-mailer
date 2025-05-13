import {  ipcMain, BrowserWindow } from 'electron';
import nodemailer from 'nodemailer';
import { Account, Recipient, SendResult } from '../renderer/src/global';
import { extractEmail } from '../utils/email';
import { generateReport } from '../utils/reports';

// простой рендер {{var}}
const tpl = (s: string, vars: Record<string, string>) =>
  s.replace(/\{\{(\w+)\}\}/g, (_, k) => vars[ k ] ?? '');

const rand = (min: number, max: number) =>
  Math.floor(Math.random() * (max - min + 1)) + min;

class InvalidEmailError extends Error {
  constructor() {
    super('Invalid email');
  }
}

// регистрация IPC-хэндлера
export function initMailer() {
  ipcMain.handle(
    'start-mailing',
    async (
      e,
      {
        smtp,
        recipients,
        subjectTemplate,
        htmlTemplate,
        pauseMin,
        pauseMax,
      }: {
        smtp: Account;
        recipients: Recipient[];
        subjectTemplate: string;
        htmlTemplate: string;
        pauseMin: number;
        pauseMax: number;
      },
    ) => {
      const win = BrowserWindow.fromWebContents(e.sender)!;
      const transport = nodemailer.createTransport({
        host: smtp.host,
        port: smtp.port,
        secure: smtp.secure,
        auth: { user: smtp.user, pass: smtp.pass },
      });

      const report: SendResult[] = [];

      for (const r of recipients) {
        const vars = { name: r.name };
        let pause = true;
        try {
          if (!extractEmail(r.email)) throw new InvalidEmailError();
          await transport.sendMail({
            from: smtp.user,
            to: r.email,
            subject: tpl(subjectTemplate, vars),
            html: tpl(htmlTemplate, vars),
          });
          win.webContents.send('mail-progress', { ...r, status: 'OK' });
          report.push({ ...r, status: 'OK', date: new Date() });
        }  catch (err: any) {
          if(err instanceof InvalidEmailError) pause = false;
          win.webContents.send('mail-progress', {
            ...r,
            status: 'FAIL',
            error: err.message,
          });
          report.push({ ...r, status: 'FAIL', error: err.message, date: new Date() });
        }
        if (pause)
          await new Promise((res) => setTimeout(res, rand(pauseMin, pauseMax)));
      }

      // отчёт
      const file = generateReport(report);

      return { file };
    },
  );
}

