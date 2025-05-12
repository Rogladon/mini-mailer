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
import type { Account, ElectronAPI } from './global';
import { extractEmail } from '../../utils/email';

/* --------------------------------- helpers -------------------------------- */


/* -------------------------------------------------------------------------- */
const api = (window as any).electronAPI as ElectronAPI | undefined;

export default function App() {
  /* ------------------------------ ui state ------------------------------ */
  const [ tplMode, setTplMode ] = useState<'inline' | 'file'>('inline');
  const [ tplFileName, setTplFile ] = useState('');
  const [ subjectTpl, setSubjectTpl ] = useState(
    'Безвозмездное партнёрство: летняя программа «Город навыков» для школ и лагерей',
  );
  const [ htmlTpl, setHtmlTpl ] = useState('<p>Здравствуйте, {{name}}</p>');

  const [ accounts, setAccounts ] = useState<Account[]>([]);
  const [ selected, setSelected ] = useState<number | null>(null);

  const [ fileName, setFileName ] = useState('');
  const [ rows, setRows ] = useState<any[]>([]); // строки из xlsx
  const [ columns, setColumns ] = useState<string[]>([]);
  const [ nameCol, setNameCol ] = useState<string>('');
  const [ emailCol, setEmailCol ] = useState<string>('');

  const [ recipients, setRecipients ] = useState<Recipient[]>([]);
  const [ previewResult, setPreviewResult ] = useState<SendResult[]>([]); // ошибки «не найден email»
  const [ results, setResults ] = useState<SendResult[]>([]);

  const [ sending, setSending ] = useState(false);
  const [ smtp, setSmtp ] = useState<Account>();
  const [ pause, setPause ] = useState({ min: 2000, max: 4000 });
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
    ): { recipients: Recipient[]; errors: SendResult[] } => {
      const valid: Recipient[] = [];
      const errs: SendResult[] = [];

      data.forEach((row) => {
        const nameVal = (row[ nameColumn ] ?? '').toString().trim();
        const emailRaw = row[ emailColumn ];
        const rowNumber = row[ '__rowNumber' ];
        const email = extractEmail(emailRaw);
        valid.push({ name: nameVal, email: email ?? "", rowNumber });
        errs.push({
          rowNumber: rowNumber,
          name: nameVal,
          email: email ?? "",
          status: email ? 'VALID' : 'FAIL',
          error: email ? undefined : 'Не найден email',
        });
      });
      return { recipients: valid, errors: errs };
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

        const { recipients: valid, errors } = buildRecipients(
          rowsArray,
          guessedName,
          guessedEmail,
        );
        setRecipients(valid);
        setPreviewResult(errors);
        setResults(errors); // отображаем статические ошибки сразу
      };
      reader.readAsArrayBuffer(file);
    },
    [ buildRecipients ],
  );

  /* ------------ rebuild recipients when mapping changes ------------- */
  useEffect(() => {
    if (!rows.length || !nameCol || !emailCol) return;
    const { recipients: valid, errors } = buildRecipients(rows, nameCol, emailCol);
    setRecipients(valid);
    setPreviewResult(errors);
    setResults(errors);
  }, [ rows, nameCol, emailCol, buildRecipients ]);

  const onAccount = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const i = +e.target.value;
    setSelected(i);
    const acc = accounts[ i ];
    if (acc) setSmtp(acc);
  };

  const start = async () => {
    if (!smtp || !smtp.user || !smtp.pass || !smtp.host || !smtp.port || !smtp.secure || !recipients.length) {
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
      });
      toast({ status: 'success', title: `Готово, отчёт: ${file}` });
    } finally {
      setSending(false);
    }
  };

  /* -------------------------------- render ------------------------------- */
  return (
    <Flex p={4} gap={6} wrap="wrap">
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
              Найдено адресов: {recipients.length} | Ошибок: {previewResult.filter(p => p.error).length}
            </Box>
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

const getColor = (status: SendResult['status']) => {
  switch (status) {
    case 'OK':
      return 'green.600';
    case 'FAIL':
      return 'red.600';
    case 'VALID':
      return 'green.600';
  }
};
