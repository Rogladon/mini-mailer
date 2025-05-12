import { ipcMain, app } from 'electron';
import fs from 'node:fs/promises';
import path from 'node:path';
import { Account } from '../renderer/src/global';

const cfgPath = app.isPackaged
  // prod: рядом с MiniMailer.exe
  ? path.join(path.dirname(app.getPath('exe')), 'accounts.json')
  // dev: корень проекта (где package.json)
  : path.join(app.getAppPath(), '/resources/accounts.json');

export function registerAccountsLoader() {
  ipcMain.handle('get-accounts', async (): Promise<Account[]> => {
    try {
      const raw = await fs.readFile(cfgPath, 'utf-8');
      return JSON.parse(raw) as Account[];
    } catch {
      return [];
    }
  });
}
