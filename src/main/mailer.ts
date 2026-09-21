import { ipcMain, BrowserWindow, app } from 'electron';
import nodemailer from 'nodemailer';
import { Account, FilePath, Recipient, SendResult } from '../renderer/src/global';
import { extractEmail } from '../utils/email';
import { generateReport } from '../utils/reports';
import mime from 'mime-types';
import path from 'node:path';
import fs from 'node:fs/promises';

const sentEmailsPath = path.join(app.getPath('userData'), 'sent-emails.json');
let mailingInProgress = false;

async function readSentEmails(): Promise<Set<string>> {
  try {
    const data: unknown = JSON.parse(await fs.readFile(sentEmailsPath, 'utf-8'));
    return Array.isArray(data)
      ? new Set(data.filter((value): value is string => typeof value === 'string'))
      : new Set();
  } catch {
    return new Set();
  }
}

async function saveSentEmails(sentEmails: Set<string>): Promise<void> {
  await fs.mkdir(path.dirname(sentEmailsPath), { recursive: true });
  const tempPath = `${sentEmailsPath}.tmp`;
  await fs.writeFile(tempPath, JSON.stringify([ ...sentEmails ], null, 2), 'utf-8');
  await fs.rename(tempPath, sentEmailsPath);
}

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

class DuplicateEmailError extends Error {
  constructor(message: string) {
    super(message);
  }
}

// регистрация IPC-хэндлера
export function initMailer() {
  ipcMain.handle('reset-sent-emails', async () => {
    if (mailingInProgress) throw new Error('Нельзя сбросить память во время рассылки');
    await saveSentEmails(new Set());
  });

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
        attachments,
        colsCopyNumbers,
        rows,
        vars
      }: {
        smtp: Account;
        recipients: Recipient[];
        subjectTemplate: string;
        htmlTemplate: string;
        pauseMin: number;
        pauseMax: number;
        attachments: FilePath[];
        colsCopyNumbers: number[];
        rows: any[];
        vars: { name: string; columnName: string }[];
      },
    ) => {
      if (mailingInProgress) throw new Error('Рассылка уже выполняется');
      mailingInProgress = true;

      return (async () => {
      const win = BrowserWindow.fromWebContents(e.sender)!;
      const transport = nodemailer.createTransport({
        host: smtp.host,
        port: smtp.port,
        secure: smtp.secure,
        auth: { user: smtp.user, pass: smtp.pass },
      });
      const report: SendResult[] = [];
      const sentEmails = await readSentEmails();
      const reservedEmails = new Set<string>();

      const formattedAttachments = attachments.map((file) => ({
        filename: file.name,
        path: file.path,
        contentType: mime.lookup(file.name) || 'application/octet-stream',
      }))

      for (const r of recipients) {
        const v = vars.reduce((acc, v) => {
          acc[ v.name ] = rows[ r.rowNumber - 1 ][ v.columnName ];
          return acc;
        }, {} as Record<string, string>)
        let pause = true;
        let emailReserved = false;
        try {
          const email = extractEmail(r.email);
          if (!email) throw new InvalidEmailError();
          if (sentEmails.has(email)) throw new DuplicateEmailError('Адрес уже был использован в предыдущей рассылке');
          if (reservedEmails.has(email)) throw new DuplicateEmailError('Дубликат адреса в текущей рассылке');
          reservedEmails.add(email);
          emailReserved = true;
          await transport.sendMail({
            from: smtp.user,
            to: email,
            subject: tpl(subjectTemplate, v),
            html: tpl(htmlTemplate, v),
            attachments: [ ...formattedAttachments, {
              filename: 'foroteh.png',
              path: path.join(app.getAppPath(), '/resources/foroteh.png'),
              cid: 'logo'
            }, {
              filename: 'image005.gif',
              path: path.join(app.getAppPath(), '/resources/image005.gif'),
              cid: 'image005'
            } ]
          });
          sentEmails.add(email);
          await saveSentEmails(sentEmails);
          win.webContents.send('mail-progress', { ...r, status: 'OK' });
          report.push({ ...r, status: 'OK', date: new Date() });
        } catch (err: any) {
          if (err instanceof InvalidEmailError) pause = false;
          if (err instanceof DuplicateEmailError) pause = false;
          const email = extractEmail(r.email);
          if (email && emailReserved) reservedEmails.delete(email);

          win.webContents.send('mail-progress', {
            ...r,
            status: err instanceof DuplicateEmailError ? 'DUBLICATE' : 'FAIL',
            error: err.message,
          });
          report.push({ ...r, status: err instanceof DuplicateEmailError ? 'DUBLICATE' : 'FAIL', error: err.message, date: new Date() });

        }
        if (pause)
          await new Promise((res) => setTimeout(res, rand(pauseMin, pauseMax)));
      }

      // отчёт
      const file = await generateReport(report, rows, colsCopyNumbers);

      return { file };
      })().finally(() => {
        mailingInProgress = false;
      });
    },
  );
}
