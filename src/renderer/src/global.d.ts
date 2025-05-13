

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

interface ElectronAPI {
  startMailing(payload: { /* … */ }): Promise<{ file: string }>;
  onMailProgress(cb: (r: SendResult) => void): () => void;
  getAccounts(): Promise<Account[]>;                // NEW
}
declare global { interface Window { electronAPI: ElectronAPI; } }
