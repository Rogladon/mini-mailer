import React, { useCallback, useEffect, useState } from 'react';
import {
  Box,
  Button,
  Flex,
  Heading,
  Input,
  NumberInput,
  NumberInputField,
  Progress,
  Table,
  Tbody,
  Td,
  Th,
  Thead,
  Tr,
  Textarea,
  VStack,
  Select,
  Alert,
  useToast,
} from '@chakra-ui/react';
import * as XLSX from 'xlsx';
import { Recipient, SendResult } from './global';
import type { Account, ElectronAPI, FilePath } from './global';
import { extractEmail, renderTemplate } from '../../utils/email';
import Preview from './components/Preview';

/* --------------------------------- helpers -------------------------------- */


/* -------------------------------------------------------------------------- */
const api = (window as any).electronAPI as ElectronAPI | undefined;

const customCols = [
  'Время отправки',
  'Статус отправки'
]

export default function App() {
  /* ------------------------------ ui state ------------------------------ */
  const [ tplMode, setTplMode ] = useState<'inline' | 'file'>('inline');
  const [ tplFileName, setTplFile ] = useState('');
  const [ subjectTpl, setSubjectTpl ] = useState(
    'Летняя программа «Город навыков» для школ и лагерей',
  );
  const [ htmlTpl, setHtmlTpl ] = useState('<p>Здравствуйте, {{name}}</p>');

  const [ accounts, setAccounts ] = useState<Account[]>([]);
  const [ selected, setSelected ] = useState<number | null>(null);

  const [ fileName, setFileName ] = useState('');
  const [ rows, setRows ] = useState<any[]>([]); // строки из xlsx
  const [ columns, setColumns ] = useState<string[]>([]);
  const [ nameCol, setNameCol ] = useState<string>('');
  const [ colsCopyNumbers, setCopyColsNumbers ] = useState<number[]>([])
  const [ emailCol, setEmailCol ] = useState<string>('');

  const [ recipients, setRecipients ] = useState<Recipient[]>([]);
  const [ results, setResults ] = useState<SendResult[]>([]);

  const [ sending, setSending ] = useState(false);
  const [ smtp, setSmtp ] = useState<Account>();
  const [ pause, setPause ] = useState({ min: 2000, max: 4000 });
  const [ attachments, setAttachments ] = useState<FilePath[]>([]);
  const [ isPreviewOpen, setIsPreviewOpen ] = useState(false);
  const [ previewTemplate, setTemplatePreview ] = useState<string>('')

  const toast = useToast();

  const total = recipients.length;
  const done = results.length;

  /* --------------------------- load accounts --------------------------- */
  useEffect(() => {
    api?.getAccounts().then(setAccounts);
  }, []);

  /* --------------------------- mail progress --------------------------- */
  useEffect(() => {
    if (!api) return;
    const off = api.onMailProgress((r) =>
      setResults((s) => [ ...s, r ]),
    );
    return off;
  }, [ api ]);

  /* --------------------------- file loaders --------------------------- */
  const loadHtml = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[ 0 ];
    if (!f) return;
    setTplFile(f.name);
    const reader = new FileReader();
    reader.onload = () => setHtmlTpl(reader.result as string);
    reader.readAsText(f, 'utf-8');
    setTplMode('file');
  };

  const guessColumn = (cols: string[], pattern: RegExp) =>
    cols.find((c) => pattern.test(c)) ?? cols[ 0 ] ?? '';

  const buildRecipients = useCallback(
    (
      data: any[],
      nameColumn: string,
      emailColumn: string,
    ): SendResult[] => {
      const valid: SendResult[] = [];

      data.forEach((row) => {
        if (row[ nameColumn ] == '' || row[ emailColumn ] == '') return;
        const nameVal = (row[ nameColumn ] ?? '').toString().trim();
        const emailRaw = row[ emailColumn ];
        const rowNumber = row[ '__rowNumber' ];
        const email = extractEmail(emailRaw);
        const rec : SendResult = {
          name: nameVal,
          email: email ?? "",
          rowNumber,
          contacts: emailRaw,
          status: email ? 'VALID' : 'FAIL',
          error: email ? undefined : 'Email не найден'
        };
        if (valid.find(p => p.email == email)) rec.status = 'DUBLICATE'
        valid.push(rec)
      });
      return valid;
    }, []
  );

  const loadXlsx = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[ 0 ];
      if (!file) return;
      setFileName(file.name);

      const reader = new FileReader();
      reader.onload = () => {
        const sheet = XLSX.read(reader.result as ArrayBuffer).Sheets;
        const rowsArray = XLSX.utils.sheet_to_json<any>(
          sheet[ Object.keys(sheet)[ 0 ] ],
          { defval: '' }
        ).map((row, index) => ({ ...row, __rowNumber: index + 1 }));
        const cols = Object.keys(rowsArray[ 0 ] ?? {});
        setRows(rowsArray);
        setColumns(cols);

        // auto‑guess mapping
        const guessedEmail = guessColumn(cols, /mail|contact|email/i);
        const guessedName = guessColumn(cols, /name|имя/i);
        setEmailCol(guessedEmail);
        setNameCol(guessedName);

        const recipients = buildRecipients(
          rowsArray,
          guessedName,
          guessedEmail,
        );
        setRecipients(recipients);
        setResults(recipients);
      };
      reader.readAsArrayBuffer(file);
    },
    [ buildRecipients ],
  );

  /* ------------ rebuild recipients when mapping changes ------------- */
  useEffect(() => {
    if (!rows.length || !nameCol || !emailCol) return;
    const recipients = buildRecipients(rows, nameCol, emailCol);
    setRecipients(recipients);
    setResults(recipients);
  }, [ rows, nameCol, emailCol, buildRecipients ]);

  const onAccount = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const i = +e.target.value;
    setSelected(i);
    const acc = accounts[ i ];
    if (acc) setSmtp(acc);
  };

  const start = async () => {
    if (!smtp || !smtp.user || !smtp.pass || !smtp.host || !smtp.port || !recipients.length) {
      toast({ status: 'warning', title: 'Укажите учётку и загрузите адреса' });
      return;
    }
    setSending(true);
    setResults([]); // не очищаем статические ошибки
    try {
      const { file } = await api!.startMailing({
        smtp,
        recipients,
        subjectTemplate: subjectTpl,
        htmlTemplate: htmlTpl,
        pauseMin: pause.min,
        pauseMax: pause.max,
        attachments,
        colsCopyNumbers,
        rows
      });
      toast({ status: 'success', title: `Готово, отчёт: ${file}` });
    } catch (errr) {
      console.error(errr);
      toast({ status: 'error', title: `Ошибка: ${errr}` });
    } finally {
      setSending(false);
    }
  };

  const handleFileSelect = async () => {
    const { filePaths } = await window.electronAPI.selectFiles();
    const files = filePaths.map((path) => ({
      name: path.split('/').pop(),
      path,
    }));
    setAttachments(files);
  };

  const handlePreview = () => {
    const vars = { name: recipients[ 0 ].name };
    const renderedHtml = renderTemplate(htmlTpl, vars);
    setTemplatePreview(renderedHtml)
    setIsPreviewOpen(true);
  };

  const handleSetCopyCol = (v: any, index?: number) => {
    const exelColumn = columns.findIndex(p => p === v);
    const customColumn = customCols.findIndex(p => p === v)
    if (index)
      setCopyColsNumbers((d) => {
        const newData = [ ...d ];
        newData[ index - 1 ] = exelColumn !== -1 ? exelColumn : -(customColumn + 1);
        return newData;
      });
    else
      setCopyColsNumbers(d => [ ...d, exelColumn !== -1 ? exelColumn : -(customColumn + 1) ])
  }

  const handleDeleteCopyCol = (index: number) => {
    setCopyColsNumbers(d => d.filter((_, i) => i !== index))
  }

  const handleAutoCopeCols = () => {
    setCopyColsNumbers([
      columns.findIndex(p => p == nameCol),
      columns.findIndex(p => p == emailCol),
      -2, -1
    ])
  }

  /* -------------------------------- render ------------------------------- */
  return (
    <Flex p={4} gap={6} wrap="wrap">
      <Preview isOpen={isPreviewOpen} onClose={() => setIsPreviewOpen(false)}
        previewContent={previewTemplate} />
      {/* ---------------------- left: settings ---------------------- */}
      <VStack w="340px" align="stretch" spacing={4}>
        <Heading size="md">Файл адресов</Heading>
        <Input type="file" accept=".xlsx" onChange={loadXlsx} />
        <Box fontSize="sm" color="gray.500">
          {fileName || 'файл не выбран'}
          {rows.length ? ` (строк: ${rows.length})` : ''}
        </Box>

        {/* mapping selects */}
        {columns.length > 0 && (
          <>
            <Heading size="sm" pt={2}>
              Настройка столбцов
            </Heading>
            <Select
              value={nameCol}
              onChange={(e) => setNameCol(e.target.value)}
              placeholder="Столбец с именем"
            >
              {columns.map((c) => (
                <option key={c}>{c}</option>
              ))}
            </Select>
            <Select
              value={emailCol}
              onChange={(e) => setEmailCol(e.target.value)}
              placeholder="Столбец с email / контактами"
            >
              {columns.map((c) => (
                <option key={c}>{c}</option>
              ))}
            </Select>
            <Alert status="info" fontSize="xs">
              Email извлекается по регулярному выражению, поэтому можно выбирать
              столбец с любыми контактами.
            </Alert>
            <Box fontSize="sm" color="gray.600">
              Найдено адресов: {recipients.length} | Ошибок: {results.filter(p => p.error).length}
            </Box>
          </>
        )}

        {columns.length > 0 && (
          <>
            <Heading size="sm" pt={2}>
              Столбцы в отчет
            </Heading>
            <Table>
              {colsCopyNumbers.map((p, i) => <Tr>
                <Td>
                  <Select
                    value={columns[ p ] ?? customCols[ -p - 1 ]}
                    onChange={(e) => handleSetCopyCol(e.target.value, i + 1)}
                    placeholder="Выберите столбец"
                  >
                    <option key={100}>Время отправки</option>
                    <option key={101}>Статус отправки</option>
                    {columns.map((c) => (
                      <option key={c}>{c}</option>
                    ))}
                  </Select>
                </Td>
                <Td><Button onClick={() => handleDeleteCopyCol(i)}>Delete</Button></Td>
              </Tr>
              )}
            </Table>
            <Select
              value={''}
              onChange={(e) => handleSetCopyCol(e.target.value)}
              placeholder="Выберите столбец"
            >
              <option key={100}>Время отправки</option>
              <option key={101}>Статус отправки</option>
              {columns.map((c) => (
                <option key={c}>{c}</option>
              ))}
            </Select>
            <Button onClick={handleAutoCopeCols}>Авто</Button>
          </>
        )}

        <Heading size="md">Учётная запись</Heading>
        <Select placeholder="Выберите ящик" onChange={onAccount} value={selected ?? ''}>
          {accounts.map((a, i) => (
            <option key={i} value={i}>
              {a.label}
            </option>
          ))}
        </Select>

        <Heading size="md" pt={4}>
          Тема письма
        </Heading>
        <Input value={subjectTpl} onChange={(e) => setSubjectTpl(e.target.value)} />

        <Heading size="md" pt={4}>
          Шаблон письма
        </Heading>
        <Select value={tplMode} onChange={(e) => setTplMode(e.target.value as any)}>
          <option value="inline">Ввести вручную</option>
          <option value="file">Загрузить HTML‑файл</option>
        </Select>
        {tplMode === 'inline' ? (
          <Textarea rows={4} value={htmlTpl} onChange={(e) => setHtmlTpl(e.target.value)} />
        ) : (
          <>
            <Input type="file" accept=".html,.htm" onChange={loadHtml} />
            <Box fontSize="sm" color="gray.500">
              {tplFileName || 'файл не выбран'}
            </Box>
          </>
        )}

        <Heading size="md" pt={4}>
          Вложения
        </Heading>
        <Button onClick={handleFileSelect} colorScheme="teal">
          Выбрать файлы
        </Button>

        <Box fontSize="sm" color="gray.500">
          {attachments.map((file) => (
            <div key={file.path}>{file.name}</div>
          ))}
        </Box>

        <Heading size="md" pt={4}>
          Пауза (мс)
        </Heading>
        <Flex gap={2}>
          <NumberInput
            value={pause.min}
            min={500}
            onChange={(_, v) => setPause({ ...pause, min: v })}
          >
            <NumberInputField />
          </NumberInput>
          <NumberInput
            value={pause.max}
            min={pause.min}
            onChange={(_, v) => setPause({ ...pause, max: v })}
          >
            <NumberInputField />
          </NumberInput>
        </Flex>

        <Button colorScheme="blue" onClick={handlePreview}>
          Предпросмотр
        </Button>

        <Button colorScheme="teal" isLoading={sending} isDisabled={!rows.length} onClick={start}>
          Отправить
        </Button>
        {sending && <Progress value={(done / total) * 100} size="sm" />}
      </VStack>

      {/* ---------------------- right: results ---------------------- */}
      <Box flex="1" minW="420px" maxH="80vh" overflow="auto">
        <Heading size="md" mb={2}>
          Результаты ({done}/{total})
        </Heading>
        <Table size="sm">
          <Thead>
            <Tr>
              <Th>Номер</Th>
              <Th>Имя (hover для полного)</Th>
              <Th>Email</Th>
              <Th>Статус</Th>
              <Th>Ошибка</Th>
            </Tr>
          </Thead>
          <Tbody>
            {results.map((r, i) => (
              <Tr key={i}>
                <Td>{r.rowNumber}</Td>
                <Td title={r.name}>{r.name.length > 15 ? `${r.name.slice(0, 15)}...` : r.name}</Td>
                <Td>{r.email || '—'}</Td>
                <Td color={getColor(r.status)}>{r.status}</Td>
                <Td>{r.error}</Td>
              </Tr>
            ))}
          </Tbody>
        </Table>
      </Box>
    </Flex>
  );
}

const getColor = (status: SendResult[ 'status' ]) => {
  switch (status) {
    case 'OK':
      return 'green.600';
    case 'FAIL':
      return 'red.600';
    case 'VALID':
      return 'green.600';
  }
};
