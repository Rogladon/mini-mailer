

export interface Account {
  label: string;
  user: string;
  pass: string,
  host: number,
  port: number,
  secure: boolean
};
export interface Recipient {
  name: string;
  email: string;
  rowNumber: number;
  contacts: string;
}
export interface SendResult extends Recipient {
  status: 'OK' | 'FAIL' | 'VALID';
  error?: string;
  date?: Date
}

export interface FilePath {
  name?: string
  path: string
}

interface ElectronAPI {
  startMailing(payload: { /* … */ }): Promise<{ file: string }>;
  onMailProgress(cb: (r: SendResult) => void): () => void;
  getAccounts(): Promise<Account[]>;
  selectFiles(): Promise<{ filePaths: string[] }>;
}
declare global { interface Window { electronAPI: ElectronAPI; } }
