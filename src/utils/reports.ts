import path from "path";
import { SendResult } from "../renderer/src/global";
import * as XLSX from 'xlsx';
import { app } from "electron";


export function generateReport(report: SendResult[]) {
  const wb = XLSX.utils.book_new();

  // Маппинг данных в нужный формат
  const mappedReport = report.map((r) => ({
    'Организация': r.name,
    'Дата отправки': r.date ? formatDate(new Date(r.date)) : '',
    'Email': r.email,
    'Контакты': r.contacts,
    'Статус': parseStatus(r.status, r.error),
  }));

  // Создание листа
  const ws = XLSX.utils.json_to_sheet(mappedReport);

  // Настройка ширины и автопереноса текста
  ws[ '!cols' ] = [
    { wch: 30 },  // Организация — Ширина 30, перенос
    { wch: 15 },  // Дата отправки — Автоподбор
    { wch: 25 },  // Email — Автоподбор
    { wch: 30 },  // Контакты — Ширина 30, перенос
    { wch: 15 }   // Статус — Ширина 15, перенос
  ];

  // Автоперенос текста в ячейках
  Object.keys(ws).forEach((cell) => {
    if (cell[ 0 ] !== '!') {
      ws[ cell ].s = {
        alignment: {
          wrapText: true,  // Это включит перенос текста
          vertical: 'center',
          horizontal: 'left',
        }
      };
    }
  });

  // Цветовое выделение строк
  mappedReport.forEach((row, index) => {
    const excelRow = index + 2; // +2 потому что индекс в массиве и заголовок
    const status = row[ 'Статус' ];
    const range = `A${excelRow}:E${excelRow}`;
    if (status.includes('Ошибка')) {
      ws[ range ] = { s: { fill: { fgColor: { rgb: "FFCCCC" } } } }; // Красный
    } else if (status === 'Отправлено') {
      ws[ range ] = { s: { fill: { fgColor: { rgb: "CCFFCC" } } } }; // Зеленый
    }
  });

  // Добавление листа в книгу
  XLSX.utils.book_append_sheet(wb, ws, 'Отчет');

  // Сохранение файла
  const file = path.join(app.getPath('desktop'), `отчет_рассылки_${formatDateForFileName()}.xlsx`);
  XLSX.writeFile(wb, file);

  return file;
}

const parseStatus = (status: string, error?: string) => {
  switch (status) {
    case 'OK':
      return 'Отправлено';
    case 'FAIL':
      return `Ошибка: ${error ?? 'Неизвестная ошибка'}`;
    case 'VALID':
      return 'Требуется проверка';
    default:
      return 'Неизвестный статус';
  }
};


const formatDate = (date: Date): string => {
  return date.toLocaleString('ru-RU', {
    day: '2-digit',
    month: '2-digit',
    year: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  })
}

const formatDateForFileName = () => {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');

  // Итоговая строка без символов разделения
  return `${year}-${month}-${day}_${hours}-${minutes}`;
};
