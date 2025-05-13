import { contextBridge, ipcRenderer } from 'electron'
import { electronAPI } from '@electron-toolkit/preload'
import type { Account, Recipient, SendResult } from '../renderer/src/global';

// Custom APIs for renderer
const api = {
  startMailing: (payload: {
    smtp: { user: string; pass: string };
    recipients: Recipient[];
    subjectTemplate: string;
    htmlTemplate: string;
    pauseMin: number;
    pauseMax: number;
  }) => ipcRenderer.invoke('start-mailing', payload),

  onMailProgress: (cb: (r: SendResult) => void) => {
    const wrapper = (_: unknown, data: SendResult) => cb(data);
    ipcRenderer.on('mail-progress', wrapper);
    return () => ipcRenderer.removeListener('mail-progress', wrapper); // <-- off
  },

  getAccounts: (): Promise<Account[]> => ipcRenderer.invoke('get-accounts'),

  selectFiles: () => ipcRenderer.invoke('dialog:openFile'),
};

// Use `contextBridge` APIs to expose Electron APIs to
// renderer only if context isolation is enabled, otherwise
// just add to the DOM global.
if (process.contextIsolated) {
  try {
    contextBridge.exposeInMainWorld('electron', electronAPI)
    contextBridge.exposeInMainWorld('electronAPI', api)
  } catch (error) {
    console.error(error)
  }
} else {
  // @ts-ignore (define in dts)
  window.electron = electronAPI
  // @ts-ignore (define in dts)
  window.api = api
}
