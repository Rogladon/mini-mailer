import ExcelJS from 'exceljs';
import path from 'node:path';
import { app } from 'electron';

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

const formatDateForFileName = () => {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  return `${year}${month}${day}_${hours}${minutes}`;
};

export async function generateReport(report: any[]) {
  // Создаем новую книгу и лист
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Отчет');

  // Задаем ширину колонок
  sheet.columns = [
    { header: 'Организация', key: 'organization', width: 30 },
    { header: 'Дата отправки', key: 'date', width: 20 },
    { header: 'Email', key: 'email', width: 25 },
    { header: 'Контакты', key: 'contacts', width: 30 },
    { header: 'Статус', key: 'status', width: 35 },
  ];

  // Добавляем стили к заголовкам
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
  };
  headerRow.height = 25;

  const calculateRowHeight = (text: string, columnWidth: number) => {
    const maxWidth = columnWidth; // умножаем на 1.2 для учета ширины символов
    const lines = text.split('\n');
    let lineCount = 0;

    lines.forEach((line) => {
      lineCount += Math.ceil(line.length / maxWidth);
    });

    return Math.max(lineCount, 1) * 20; // 20px на каждую строку
  };

  // Заполняем данные
  report.forEach((r) => {
    const row = sheet.addRow({
      organization: r.name,
      date: r.date ? new Date(r.date).toLocaleString('ru-RU', {
        day: '2-digit',
        month: '2-digit',
        year: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
      }) : '',
      email: r.email,
      contacts: r.contacts,
      status: parseStatus(r.status, r.error),
    });

    // Автоперенос текста
    row.alignment = { wrapText: true, vertical: 'middle' };

    // 🟢 **Границы для всех ячеек в строке**
    row.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    // 🟢 **Цветовое выделение**
    if (r.status === 'FAIL') {
      row.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFCCCC' }, // Красный фон
        };
      });
    } else if (r.status === 'OK') {
      row.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFCCFFCC' }, // Зеленый фон
        };
      });
    }

    const orgHeight = calculateRowHeight(r.name, 30); // Ширина 'Организация'
    const contactHeight = calculateRowHeight(r.contacts, 30); // Ширина 'Контакты'

    // Берем максимальное значение из двух, чтобы строка не обрезалась
    row.height = Math.max(orgHeight, contactHeight, 20);
  });

  // Сохранение файла
  const timestamp = formatDateForFileName();
  const file = path.join(app.getPath('desktop'), `отчет_рассылки_${timestamp}.xlsx`);

  // Ждем, пока запишется файл
  await workbook.xlsx.writeFile(file);

  return file;
}
