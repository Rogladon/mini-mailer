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

export async function generateReport(report: any[], rows: any[], copyNumbers: number[]) {

  // Создаем новую книгу и лист
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Отчет');

  const headers = copyNumbers.map((index) => {
    if (index >= 0) {
      const header = Object.keys(rows[ 0 ])[ index ];
      return { header: header ?? `Колонка ${index + 1}`, key: `column_${index}`, width: 25 };
    } else if (index === -1) {
      return { header: 'Время отправки', key: 'sendTime', width: 20 };
    } else if (index === -2) {
      return { header: 'Статус отправки', key: 'sendStatus', width: 25 };
    }
    return ''
  });

  // Задаем ширину колонок
  sheet.columns = headers

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
  headerRow.height = 50;

  const calculateRowHeight = (text: string, columnWidth: number) => {
    if (!text || typeof text !== 'string' || !columnWidth) return 20;
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
    const originalRow = rows.find((row) => row.__rowNumber === r.rowNumber);

    // Массив данных для вставки в Excel
    const rowData = copyNumbers.map((index) => {
      if (index >= 0) {
        return originalRow ? originalRow[ Object.keys(originalRow)[ index ] ] : '';
      } else if (index === -1) {
        return r.date ? new Date(r.date).toLocaleString('ru-RU', {
          day: '2-digit',
          month: '2-digit',
          year: '2-digit',
          hour: '2-digit',
          minute: '2-digit',
        }) : '';
      } else if (index === -2) {
        return parseStatus(r.status);
      }
    });

    // Добавляем строку в Excel
    const row = sheet.addRow(rowData);

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
