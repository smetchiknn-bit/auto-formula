import { useState, useCallback } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// Версия приложения и история изменений
const APP_VERSION = '1.0.0';
const APP_DATE = '15.01.2026';

interface VersionEntry {
  version: string;
  date: string;
  changes: string[];
}

const VERSION_HISTORY: VersionEntry[] = [
  {
    version: '1.0.0',
    date: '15.01.2026',
    changes: [
      'Первый стабильный релиз',
      'Автоматическое определение последней заполненной строки',
      'Вставка значений СМР, ТМЦ, ВСЕГО в ячейки AS2–AS4',
      'Вставка формул СУММПРОИЗВ в ячейки AT2–AT4',
      'Красный цвет шрифта и полужирное начертание',
      'Установка курсора на ячейку AT4',
      'Автоматическое раскрытие скрытых колонок AS и AT',
      'Сохранение структуры файла',
    ],
  },
];

// SVG иконки вместо Font Awesome
const FileExcelIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
    <path d="M14 2H6C4.89543 2 4 2.89543 4 4V20C4 21.1046 4.89543 22 6 22H18C19.1046 22 20 21.1046 20 20V8L14 2Z" fill="currentColor" opacity="0.2"/>
    <path d="M14 2V8H20" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
    <path d="M8 13L10 17M10 13L8 17M13 13H16M13 17H16" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
    <rect x="4" y="2" width="16" height="20" rx="2" stroke="currentColor" strokeWidth="1.5" fill="none"/>
  </svg>
);

const FolderOpenIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2v11z"/>
    <path d="M2 10h20"/>
  </svg>
);

const CheckCircleIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
    <polyline points="22 4 12 14.01 9 11.01"/>
  </svg>
);

const ExclamationIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
    <line x1="12" y1="9" x2="12" y2="13"/>
    <line x1="12" y1="17" x2="12.01" y2="17"/>
  </svg>
);

const ListCheckIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M10 6h11"/><path d="M10 12h11"/><path d="M10 18h11"/>
    <polyline points="3 6 4 7 6 5"/><polyline points="3 12 4 13 6 11"/><polyline points="3 18 4 19 6 17"/>
  </svg>
);

const PaintBrushIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M18.37 2.63L14 7l-1.59-1.59a2 2 0 0 0-2.82 0L8 7l9 9 1.59-1.59a2 2 0 0 0 0-2.82L17 10l4.37-4.37a2.12 2.12 0 1 0-3-3z"/>
    <path d="M9 8c-2 3-4 3.5-7 4l8 10c2-1 6-5 6-7"/>
  </svg>
);

const EyeIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
    <circle cx="12" cy="12" r="3"/>
  </svg>
);

const MousePointerIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 3l7.07 16.97 2.51-7.39 7.39-2.51L3 3z"/>
    <path d="M13 13l6 6"/>
  </svg>
);

const ShieldIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
  </svg>
);

const SpinnerIcon = ({ className = '' }: { className?: string }) => (
  <svg className={`animate-spin ${className}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <circle cx="12" cy="12" r="10" strokeOpacity="0.25"/>
    <path d="M12 2a10 10 0 0 1 10 10" strokeLinecap="round"/>
  </svg>
);

const TimesIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
  </svg>
);

const TimesCircleIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="10"/>
    <line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>
  </svg>
);

const HashIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <line x1="4" y1="9" x2="20" y2="9"/><line x1="4" y1="15" x2="20" y2="15"/>
    <line x1="10" y1="3" x2="8" y2="21"/><line x1="16" y1="3" x2="14" y2="21"/>
  </svg>
);

const CodeBranchIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <line x1="6" y1="3" x2="6" y2="15"/>
    <circle cx="18" cy="6" r="3"/><circle cx="6" cy="18" r="3"/>
    <path d="M18 9a9 9 0 0 1-9 9"/>
  </svg>
);

const CheckIcon = ({ className = '' }: { className?: string }) => (
  <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <polyline points="20 6 9 17 4 12"/>
  </svg>
);

function App() {
  const [fileName, setFileName] = useState<string>('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [status, setStatus] = useState<string>('');
  const [error, setError] = useState<string>('');
  const [isSuccess, setIsSuccess] = useState(false);
  const [lastRowInfo, setLastRowInfo] = useState<number>(0);
  const [showVersionHistory, setShowVersionHistory] = useState(false);
  const [isDragging, setIsDragging] = useState(false);

  const processFile = useCallback(async (file: File) => {
    setIsProcessing(true);
    setStatus('Загрузка файла...');
    setError('');
    setIsSuccess(false);
    setLastRowInfo(0);

    try {
      const arrayBuffer = await file.arrayBuffer();

      setStatus('Обработка файла...');

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer as ArrayBuffer);

      const worksheet = workbook.getWorksheet('Корректировка 1');
      if (!worksheet) {
        throw new Error('Лист "Корректировка 1" не найден в файле. Убедитесь, что лист существует и имеет точное название.');
      }

      setStatus('Определение последней заполненной строки...');

      let lastRow = 0;
      worksheet.eachRow({ includeEmpty: false }, (_row: any, rowNumber: number) => {
        if (rowNumber > lastRow) {
          lastRow = rowNumber;
        }
      });

      if (lastRow < 6) {
        lastRow = 6;
      }

      setLastRowInfo(lastRow);
      setStatus(`Строка: ${lastRow}. Раскрытие колонок и вставка формул...`);

      const colAS = 45;
      const colAT = 46;

      // Раскрываем скрытые колонки AS и AT (если они скрыты)
      const colASObj = worksheet.getColumn(colAS);
      const colATObj = worksheet.getColumn(colAT);
      if (colASObj.hidden) {
        colASObj.hidden = false;
      }
      if (colATObj.hidden) {
        colATObj.hidden = false;
      }

      // Очистка ячеек
      const cellsToClean = [
        { row: 2, col: colAS },
        { row: 3, col: colAS },
        { row: 4, col: colAS },
        { row: 2, col: colAT },
        { row: 3, col: colAT },
        { row: 4, col: colAT },
      ];

      cellsToClean.forEach(({ row, col }) => {
        const cell = worksheet.getCell(row, col);
        cell.value = null;
      });

      // Вставка слов
      const cellAS2 = worksheet.getCell(2, colAS);
      cellAS2.value = 'СМР';
      cellAS2.font = { bold: true, color: { argb: 'FFFF0000' } };

      const cellAS3 = worksheet.getCell(3, colAS);
      cellAS3.value = 'ТМЦ';
      cellAS3.font = { bold: true, color: { argb: 'FFFF0000' } };

      const cellAS4 = worksheet.getCell(4, colAS);
      cellAS4.value = 'ВСЕГО';
      cellAS4.font = { bold: true, color: { argb: 'FFFF0000' } };

      // Формулы
      const formulaAT2 = `SUMPRODUCT($X6:$X${lastRow},R6:R${lastRow},SUBTOTAL(3,OFFSET($X$6:$X$${lastRow},ROW($X$6:$X$${lastRow})-ROW($X6),,1)))+SUMPRODUCT($Y6:$Y${lastRow},AA6:AA${lastRow},SUBTOTAL(3,OFFSET($Y$6:$Y$${lastRow},ROW($Y$6:$Y$${lastRow})-ROW($X6),,1)))`;

      const cellAT2 = worksheet.getCell(2, colAT);
      cellAT2.value = { formula: formulaAT2 };
      cellAT2.font = { bold: true, color: { argb: 'FFFF0000' } };

      const formulaAT3 = `SUMPRODUCT($X6:$X${lastRow},Q6:Q${lastRow},SUBTOTAL(3,OFFSET($X$6:$X$${lastRow},ROW($X$6:$X$${lastRow})-ROW($X6),,1)))+SUMPRODUCT($Y6:$Y${lastRow},Z6:Z${lastRow},SUBTOTAL(3,OFFSET($Y$6:$Y$${lastRow},ROW($Y$6:$Y$${lastRow})-ROW($X6),,1)))`;

      const cellAT3 = worksheet.getCell(3, colAT);
      cellAT3.value = { formula: formulaAT3 };
      cellAT3.font = { bold: true, color: { argb: 'FFFF0000' } };

      const formulaAT4 = 'AT2+AT3';

      const cellAT4 = worksheet.getCell(4, colAT);
      cellAT4.value = { formula: formulaAT4 };
      cellAT4.font = { bold: true, color: { argb: 'FFFF0000' } };

      // Курсор на AT4
      worksheet.views = [
        {
          state: 'normal' as const,
          activeCell: 'AT4',
          zoomScale: 100
        }
      ];

      setStatus('Сохранение файла...');

      const output = await workbook.xlsx.writeBuffer();
      const blob = new Blob([output], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });

      const outputFileName = file.name.replace(/\.xlsx?$/i, '') + '_Корректировка.xlsx';
      saveAs(blob, outputFileName);

      setStatus('');
      setIsSuccess(true);
      setIsProcessing(false);

    } catch (err) {
      console.error('Ошибка обработки:', err);
      setError(err instanceof Error ? err.message : 'Произошла неизвестная ошибка при обработке файла');
      setIsProcessing(false);
      setStatus('');
    }
  }, []);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (!file.name.match(/\.xlsx?$/i)) {
        setError('Пожалуйста, выберите файл Excel (.xlsx или .xlsm)');
        return;
      }
      setFileName(file.name);
      setIsSuccess(false);
      setError('');
      processFile(file);
    }
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      if (!file.name.match(/\.xlsx?$/i)) {
        setError('Пожалуйста, загрузите файл Excel (.xlsx или .xlsm)');
        return;
      }
      setFileName(file.name);
      setIsSuccess(false);
      setError('');
      processFile(file);
    }
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const resetApp = () => {
    setFileName('');
    setStatus('');
    setError('');
    setIsSuccess(false);
    setLastRowInfo(0);
  };

  return (
    <div className="min-h-screen flex flex-col relative overflow-hidden">
      {/* Плавающие математические символы */}
      <div className="fixed inset-0 pointer-events-none overflow-hidden z-0">
        <span className="float-symbol absolute top-[10%] left-[5%] text-6xl font-bold text-green-700 opacity-15">Σ</span>
        <span className="float-symbol-reverse absolute top-[20%] right-[8%] text-5xl font-bold text-green-800 opacity-10">₽</span>
        <span className="float-symbol-slow absolute top-[60%] left-[10%] text-7xl font-bold text-green-600 opacity-10">=</span>
        <span className="float-symbol absolute top-[75%] right-[15%] text-5xl font-bold text-green-700 opacity-12">Σ</span>
        <span className="float-symbol-reverse absolute top-[40%] left-[80%] text-4xl font-bold text-green-800 opacity-10">+</span>
        <span className="float-symbol-slow absolute top-[85%] left-[50%] text-6xl font-bold text-green-600 opacity-8">%</span>
        <span className="float-symbol absolute top-[15%] left-[45%] text-4xl font-bold text-green-700 opacity-10">÷</span>
        <span className="float-symbol-reverse absolute top-[50%] left-[30%] text-5xl font-bold text-green-800 opacity-8">∑</span>
        <span className="float-symbol-slow absolute top-[30%] right-[25%] text-4xl font-bold text-green-600 opacity-12">×</span>
        <span className="float-symbol absolute top-[70%] left-[70%] text-5xl font-bold text-green-700 opacity-10">₽</span>
      </div>

      {/* Основной контент */}
      <div className="flex-1 flex flex-col items-center justify-center p-4 md:p-8 relative z-10">
        {/* Заголовок с логотипами */}
        <div className="flex items-center justify-center gap-4 md:gap-6 mb-2">
          <img src="./logo1C.png" alt="1C" className="w-12 h-12 md:w-14 md:h-14 object-contain" />
          <div className="text-center">
            <h1 className="text-2xl md:text-4xl font-extrabold text-green-900 tracking-tight">
              ФОРМУЛЫ для КОРРЕКТИРОВКИ
            </h1>
            <p className="text-green-700 text-base md:text-lg font-medium mt-1">
              Автоматическая вставка формул в файл Excel
            </p>
          </div>
          <img src="./logoXLSX.png" alt="Excel" className="w-12 h-12 md:w-14 md:h-14 object-contain" />
        </div>

        {/* Основная карточка */}
        <div className="w-full max-w-2xl bg-white rounded-2xl border-2 border-green-800 p-6 md:p-8 mt-6">
          {/* Зона загрузки */}
          <div
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            className={`border-3 border-dashed rounded-xl p-8 text-center transition-all duration-200 mb-6 ${
              isDragging
                ? 'border-green-600 bg-green-50 scale-[1.02]'
                : 'border-green-400 hover:border-green-600 hover:bg-green-50/50'
            }`}
            style={{ borderWidth: '3px', minHeight: '280px' }}
          >
            <div className="mb-4">
              <div className="w-16 h-16 mx-auto bg-green-100 rounded-full flex items-center justify-center">
                <FileExcelIcon className="w-8 h-8 text-green-600" />
              </div>
            </div>
            <p className="text-green-900 text-lg mb-1 font-semibold">
              Укажите путь к файлу Корректировка
            </p>
            <p className="text-green-600 text-sm mb-5">
              Перетащите файл Excel сюда или нажмите кнопку
            </p>
            <div className="mt-8">
              <label className="inline-flex items-center bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-8 rounded-xl cursor-pointer transition-all">
                <FolderOpenIcon className="w-5 h-5 mr-2" />
                Выбрать файл
                <input
                  type="file"
                  accept=".xlsx,.xlsm"
                  onChange={handleFileChange}
                  className="hidden"
                />
              </label>
            </div>
          </div>

          {/* Имя файла */}
          {fileName && (
            <div className="bg-green-50 rounded-xl p-4 mb-4 flex items-center justify-between border border-green-200">
              <div className="flex items-center">
                <div className="w-10 h-10 bg-green-600 rounded-lg flex items-center justify-center mr-3">
                  <FileExcelIcon className="w-5 h-5 text-white" />
                </div>
                <div>
                  <p className="text-green-900 font-semibold text-sm">{fileName}</p>
                  <p className="text-green-600 text-xs">Файл для обработки</p>
                </div>
              </div>
              {!isProcessing && (
                <button
                  onClick={resetApp}
                  className="text-green-600 hover:text-green-800 transition-colors p-2"
                  title="Сбросить"
                >
                  <TimesCircleIcon className="w-5 h-5" />
                </button>
              )}
            </div>
          )}

          {/* Статус */}
          {isProcessing && (
            <div className="bg-green-100 rounded-xl p-4 mb-4 flex items-center border border-green-300">
              <SpinnerIcon className="w-5 h-5 text-green-600 mr-3" />
              <p className="text-green-800 font-medium">{status}</p>
            </div>
          )}

          {/* Информация о строке */}
          {lastRowInfo > 0 && (
            <div className="bg-green-50 rounded-xl p-3 mb-4 flex items-center border border-green-200">
              <HashIcon className="w-4 h-4 text-green-600 mr-3" />
              <p className="text-green-800 text-sm">
                Последняя заполненная строка: <strong className="text-green-900">{lastRowInfo}</strong>
              </p>
            </div>
          )}

          {/* Успех */}
          {isSuccess && (
            <div className="bg-green-100 rounded-xl p-4 mb-4 flex items-start border-2 border-green-400">
              <CheckCircleIcon className="w-6 h-6 text-green-600 mr-3 mt-0.5 shrink-0" />
              <div>
                <p className="text-green-900 font-bold">Файл успешно обработан!</p>
                <p className="text-green-700 text-sm mt-1">
                  Формулы вставлены. Файл скачан. Курсор на AT4.
                </p>
              </div>
            </div>
          )}

          {/* Ошибка */}
          {error && (
            <div className="bg-red-50 rounded-xl p-4 mb-4 flex items-start border-2 border-red-300">
              <ExclamationIcon className="w-6 h-6 text-red-500 mr-3 mt-0.5 shrink-0" />
              <div>
                <p className="text-red-800 font-bold">Ошибка</p>
                <p className="text-red-600 text-sm mt-1">{error}</p>
              </div>
            </div>
          )}

          {/* Описание действий */}
          <div className="bg-green-50/50 rounded-xl p-5 border border-green-200">
            <h3 className="text-green-900 font-bold mb-3 flex items-center">
              <ListCheckIcon className="w-5 h-5 text-green-600 mr-2" />
              Что будет сделано:
            </h3>
            <ul className="text-green-800 text-sm space-y-2.5">
              <li className="flex items-start">
                <span className="bg-green-600 text-white rounded px-2 py-0.5 text-xs font-mono mr-3 mt-0.5 shrink-0 font-bold">AS</span>
                <span>Ячейки AS2–AS4: значения <strong className="text-red-600">СМР</strong>, <strong className="text-red-600">ТМЦ</strong>, <strong className="text-red-600">ВСЕГО</strong></span>
              </li>
              <li className="flex items-start">
                <span className="bg-green-600 text-white rounded px-2 py-0.5 text-xs font-mono mr-3 mt-0.5 shrink-0 font-bold">AT</span>
                <span>Ячейки AT2–AT4: формулы с автоопределением последней строки</span>
              </li>
              <li className="flex items-start">
                <PaintBrushIcon className="w-4 h-4 text-yellow-600 mr-3 mt-0.5 shrink-0" />
                <span><strong className="text-red-600">Красный шрифт</strong> + <strong>полужирный</strong> для всех вставленных ячеек</span>
              </li>
              <li className="flex items-start">
                <EyeIcon className="w-4 h-4 text-green-600 mr-3 mt-0.5 shrink-0" />
                <span>Колонки <strong>AS</strong> и <strong>AT</strong> автоматически <strong>раскрываются</strong> (если были скрыты)</span>
              </li>
              <li className="flex items-start">
                <MousePointerIcon className="w-4 h-4 text-green-600 mr-3 mt-0.5 shrink-0" />
                <span>Курсор устанавливается на <strong>AT4</strong></span>
              </li>
              <li className="flex items-start">
                <ShieldIcon className="w-4 h-4 text-green-600 mr-3 mt-0.5 shrink-0" />
                <span>Остальная структура файла <strong>не нарушается</strong></span>
              </li>
            </ul>
          </div>
        </div>

        {/* Подпись под карточкой */}
        <div className="text-center mt-6 text-green-700/70 text-sm font-medium">
          <p>Лист: «Корректировка 1» | Ячейки: AS2–AS4, AT2–AT4</p>
        </div>
      </div>

      {/* Футер с кнопкой версии */}
      <footer className="relative z-10 py-4 px-4 text-center">
        <button
          onClick={() => setShowVersionHistory(!showVersionHistory)}
          className="text-green-600/40 hover:text-green-700 text-xs font-medium transition-colors underline underline-offset-2 decoration-dotted"
        >
          Версия {APP_VERSION} от {APP_DATE}
        </button>

        {/* Модальное окно истории версий */}
        {showVersionHistory && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" onClick={() => setShowVersionHistory(false)}>
            <div className="absolute inset-0 bg-black/30 backdrop-blur-sm"></div>
            <div
              className="relative bg-white rounded-2xl border-2 border-green-800 max-w-lg w-full max-h-[80vh] overflow-y-auto p-6"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="flex items-center justify-between mb-5">
                <h2 className="text-xl font-bold text-green-900 flex items-center">
                  <CodeBranchIcon className="w-5 h-5 text-green-600 mr-2" />
                  История изменений
                </h2>
                <button
                  onClick={() => setShowVersionHistory(false)}
                  className="text-green-600 hover:text-green-800 transition-colors"
                >
                  <TimesIcon className="w-5 h-5" />
                </button>
              </div>

              <div className="space-y-5">
                {VERSION_HISTORY.map((entry, idx) => (
                  <div key={idx} className="border-l-4 border-green-600 pl-4">
                    <div className="flex items-center gap-2 mb-2">
                      <span className="bg-green-600 text-white text-xs font-bold px-2 py-1 rounded">
                        v{entry.version}
                      </span>
                      <span className="text-green-600 text-xs font-medium">
                        {entry.date}
                      </span>
                      <span className="bg-green-100 text-green-800 text-xs px-2 py-0.5 rounded-full font-medium">
                        Стабильная
                      </span>
                    </div>
                    <ul className="space-y-1">
                      {entry.changes.map((change, cIdx) => (
                        <li key={cIdx} className="text-sm text-green-800 flex items-start">
                          <CheckIcon className="w-3 h-3 text-green-500 mr-2 mt-1 shrink-0" />
                          <span>{change}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </footer>
    </div>
  );
}

export default App;
