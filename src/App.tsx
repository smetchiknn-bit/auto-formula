import { useState, useCallback } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

function App() {
  const [fileName, setFileName] = useState<string>('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [status, setStatus] = useState<string>('');
  const [error, setError] = useState<string>('');
  const [isSuccess, setIsSuccess] = useState(false);
  const [lastRowInfo, setLastRowInfo] = useState<number>(0);

  const processFile = useCallback(async (file: File) => {
    setIsProcessing(true);
    setStatus('Загрузка файла...');
    setError('');
    setIsSuccess(false);
    setLastRowInfo(0);

    try {
      // Читаем файл как ArrayBuffer
      const arrayBuffer = await file.arrayBuffer();
      
      setStatus('Обработка файла...');
      
      // Загружаем workbook
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer as ArrayBuffer);

      // Ищем лист "Корректировка 1"
      const worksheet = workbook.getWorksheet('Корректировка 1');
      if (!worksheet) {
        throw new Error('Лист "Корректировка 1" не найден в файле. Убедитесь, что лист существует и имеет точное название "Корректировка 1".');
      }

      setStatus('Определение последней заполненной строки...');

      // Определяем последнюю заполненную строку на листе
      let lastRow = 0;
      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > lastRow) {
          lastRow = rowNumber;
        }
      });

      // Если не нашли данных, используем минимум 6
      if (lastRow < 6) {
        lastRow = 6;
      }

      setLastRowInfo(lastRow);
      setStatus(`Последняя заполненная строка: ${lastRow}. Вставка значений и формул...`);

      // Колонки: AS = 45, AT = 46
      const colAS = 45;
      const colAT = 46;

      // Очищаем ячейки перед вставкой (на случай если они заполнены)
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

      // === Вставка слов в AS2, AS3, AS4 ===
      const cellAS2 = worksheet.getCell(2, colAS);
      cellAS2.value = 'СМР';
      cellAS2.font = { bold: true, color: { argb: 'FFFF0000' } };

      const cellAS3 = worksheet.getCell(3, colAS);
      cellAS3.value = 'ТМЦ';
      cellAS3.font = { bold: true, color: { argb: 'FFFF0000' } };

      const cellAS4 = worksheet.getCell(4, colAS);
      cellAS4.value = 'ВСЕГО';
      cellAS4.font = { bold: true, color: { argb: 'FFFF0000' } };

      // === Вставка формул в AT2, AT3, AT4 ===
      // Формулы в формате xlsx (английские имена функций)
      // При открытии в Excel с русской локалью они будут отображаться на русском

      // AT2 - формула для СМР:
      // =(СУММПРОИЗВ($X6:$X[lastRow];R6:R[lastRow];ПРОМЕЖУТОЧНЫЕ.ИТОГИ(3;СМЕЩ($X$6:$X$[lastRow];СТРОКА($X$6:$X$[lastRow])-СТРОКА($X6);;1))))+(СУММПРОИЗВ($Y6:$Y[lastRow];AA6:AA[lastRow];ПРОМЕЖУТОЧНЫЕ.ИТОГИ(3;СМЕЩ($Y$6:$Y$[lastRow];СТРОКА($Y$6:$Y$[lastRow])-СТРОКА($X6);;1))))
      const formulaAT2 = `SUMPRODUCT($X6:$X${lastRow},R6:R${lastRow},SUBTOTAL(3,OFFSET($X$6:$X$${lastRow},ROW($X$6:$X$${lastRow})-ROW($X6),,1)))+SUMPRODUCT($Y6:$Y${lastRow},AA6:AA${lastRow},SUBTOTAL(3,OFFSET($Y$6:$Y$${lastRow},ROW($Y$6:$Y$${lastRow})-ROW($X6),,1)))`;
      
      const cellAT2 = worksheet.getCell(2, colAT);
      cellAT2.value = { formula: formulaAT2 };
      cellAT2.font = { bold: true, color: { argb: 'FFFF0000' } };

      // AT3 - формула для ТМЦ:
      // =(СУММПРОИЗВ($X6:$X[lastRow];Q6:Q[lastRow];ПРОМЕЖУТОЧНЫЕ.ИТОГИ(3;СМЕЩ($X$6:$X$[lastRow];СТРОКА($X$6:$X$[lastRow])-СТРОКА($X6);;1))))+(СУММПРОИЗВ($Y6:$Y[lastRow];Z6:Z[lastRow];ПРОМЕЖУТОЧНЫЕ.ИТОГИ(3;СМЕЩ($Y$6:$Y$[lastRow];СТРОКА($Y$6:$Y$[lastRow])-СТРОКА($X6);;1))))
      const formulaAT3 = `SUMPRODUCT($X6:$X${lastRow},Q6:Q${lastRow},SUBTOTAL(3,OFFSET($X$6:$X$${lastRow},ROW($X$6:$X$${lastRow})-ROW($X6),,1)))+SUMPRODUCT($Y6:$Y${lastRow},Z6:Z${lastRow},SUBTOTAL(3,OFFSET($Y$6:$Y$${lastRow},ROW($Y$6:$Y$${lastRow})-ROW($X6),,1)))`;
      
      const cellAT3 = worksheet.getCell(3, colAT);
      cellAT3.value = { formula: formulaAT3 };
      cellAT3.font = { bold: true, color: { argb: 'FFFF0000' } };

      // AT4 - формула ВСЕГО = AT2 + AT3
      const formulaAT4 = 'AT2+AT3';
      
      const cellAT4 = worksheet.getCell(4, colAT);
      cellAT4.value = { formula: formulaAT4 };
      cellAT4.font = { bold: true, color: { argb: 'FFFF0000' } };

      // Устанавливаем курсор на ячейку AT4
      // В exceljs это делается через worksheet views
      worksheet.views = [
        { 
          state: 'normal' as const, 
          activeCell: 'AT4',
          zoomScale: 100 
        }
      ];

      setStatus('Сохранение файла...');

      // Сохраняем файл
      const output = await workbook.xlsx.writeBuffer();
      const blob = new Blob([output], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      // Формируем имя файла для скачивания
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
  };

  const resetApp = () => {
    setFileName('');
    setStatus('');
    setError('');
    setIsSuccess(false);
    setLastRowInfo(0);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 flex items-center justify-center p-4">
      <div className="w-full max-w-2xl">
        {/* Заголовок */}
        <div className="text-center mb-8">
          <h1 className="text-3xl md:text-4xl font-bold text-white mb-2">
            📊 Формулы для Корректировки
          </h1>
          <p className="text-blue-200 text-lg">
            Автоматическая вставка формул и значений в файл Excel
          </p>
        </div>

        {/* Логотипы */}
        <div className="flex justify-center gap-6 mb-8">
          <div className="flex flex-col items-center group">
            <img src="/logoXLSX.svg" alt="Excel" className="w-14 h-14 rounded-xl shadow-lg group-hover:scale-110 transition-transform" />
            <span className="text-xs text-blue-300 mt-1">Excel</span>
          </div>
          <div className="flex flex-col items-center group">
            <img src="/logoGS.svg" alt="Google Sheets" className="w-14 h-14 rounded-xl shadow-lg group-hover:scale-110 transition-transform" />
            <span className="text-xs text-blue-300 mt-1">Google Sheets</span>
          </div>
          <div className="flex flex-col items-center group">
            <img src="/logo1C.svg" alt="1C" className="w-14 h-14 rounded-xl shadow-lg group-hover:scale-110 transition-transform" />
            <span className="text-xs text-blue-300 mt-1">1С</span>
          </div>
          <div className="flex flex-col items-center group">
            <img src="/logoStiker.svg" alt="Стикер" className="w-14 h-14 rounded-xl shadow-lg group-hover:scale-110 transition-transform" />
            <span className="text-xs text-blue-300 mt-1">Стикер</span>
          </div>
        </div>

        {/* Основная карточка */}
        <div className="bg-white/10 backdrop-blur-lg rounded-2xl shadow-2xl border border-white/20 p-8">
          {/* Зона загрузки */}
          <div
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            className="border-2 border-dashed border-blue-400/50 rounded-xl p-8 text-center hover:border-blue-400 hover:bg-white/5 transition-all cursor-pointer mb-6"
          >
            <div className="mb-4">
              <i className="fas fa-file-excel text-5xl text-green-400"></i>
            </div>
            <p className="text-white text-lg mb-2 font-medium">
              Укажите путь к файлу Корректировка
            </p>
            <p className="text-blue-300 text-sm mb-4">
              Перетащите файл Excel сюда или нажмите кнопку ниже
            </p>
            <label className="inline-block bg-gradient-to-r from-blue-600 to-blue-700 hover:from-blue-700 hover:to-blue-800 text-white font-semibold py-3 px-8 rounded-lg cursor-pointer transition-all shadow-lg hover:shadow-xl active:scale-95">
              <i className="fas fa-folder-open mr-2"></i>
              Выбрать файл Excel
              <input
                type="file"
                accept=".xlsx,.xlsm"
                onChange={handleFileChange}
                className="hidden"
              />
            </label>
          </div>

          {/* Имя файла */}
          {fileName && (
            <div className="bg-white/5 rounded-lg p-4 mb-4 flex items-center justify-between">
              <div className="flex items-center">
                <i className="fas fa-file-excel text-green-400 mr-3 text-xl"></i>
                <div>
                  <p className="text-white font-medium">{fileName}</p>
                  <p className="text-blue-300 text-sm">Файл для обработки</p>
                </div>
              </div>
              {!isProcessing && (
                <button 
                  onClick={resetApp}
                  className="text-blue-400 hover:text-blue-300 transition-colors"
                  title="Сбросить"
                >
                  <i className="fas fa-times-circle text-xl"></i>
                </button>
              )}
            </div>
          )}

          {/* Статус обработки */}
          {isProcessing && (
            <div className="bg-blue-500/20 rounded-lg p-4 mb-4 flex items-center">
              <div className="animate-spin mr-3">
                <i className="fas fa-spinner text-blue-400 text-xl"></i>
              </div>
              <p className="text-blue-200">{status}</p>
            </div>
          )}

          {/* Информация о последней строке */}
          {lastRowInfo > 0 && (
            <div className="bg-indigo-500/20 rounded-lg p-3 mb-4 flex items-center">
              <i className="fas fa-hashtag text-indigo-400 mr-3"></i>
              <p className="text-indigo-200 text-sm">
                Определена последняя заполненная строка: <strong className="text-white">{lastRowInfo}</strong>
              </p>
            </div>
          )}

          {/* Успех */}
          {isSuccess && (
            <div className="bg-green-500/20 rounded-lg p-4 mb-4 flex items-start">
              <i className="fas fa-check-circle text-green-400 mr-3 text-xl mt-0.5"></i>
              <div>
                <p className="text-green-200 font-medium">✅ Файл успешно обработан и скачан!</p>
                <p className="text-green-300 text-sm mt-1">
                  Вставлены значения и формулы. Откройте файл в Excel — курсор установлен на ячейке AT4. 
                  Вы можете проверить изменения и решить, сохранять ли файл.
                </p>
              </div>
            </div>
          )}

          {/* Ошибка */}
          {error && (
            <div className="bg-red-500/20 rounded-lg p-4 mb-4 flex items-start">
              <i className="fas fa-exclamation-triangle text-red-400 mr-3 text-xl mt-0.5"></i>
              <div>
                <p className="text-red-200 font-medium">Ошибка обработки</p>
                <p className="text-red-300 text-sm mt-1">{error}</p>
              </div>
            </div>
          )}

          {/* Информация о том что будет сделано */}
          <div className="bg-white/5 rounded-lg p-5 mt-4">
            <h3 className="text-white font-semibold mb-3 flex items-center">
              <i className="fas fa-list-check text-blue-400 mr-2"></i>
              Что будет сделано с файлом:
            </h3>
            <ul className="text-blue-200 text-sm space-y-3">
              <li className="flex items-start">
                <span className="bg-blue-600/40 text-blue-200 rounded px-2 py-0.5 text-xs font-mono mr-3 mt-0.5 shrink-0">AS2-4</span>
                <span>Вставка значений: <strong className="text-red-400">СМР</strong>, <strong className="text-red-400">ТМЦ</strong>, <strong className="text-red-400">ВСЕГО</strong></span>
              </li>
              <li className="flex items-start">
                <span className="bg-blue-600/40 text-blue-200 rounded px-2 py-0.5 text-xs font-mono mr-3 mt-0.5 shrink-0">AT2</span>
                <span>Формула СУММПРОИЗВ для расчёта <strong className="text-white">СМР</strong> (с автоопределением строки)</span>
              </li>
              <li className="flex items-start">
                <span className="bg-blue-600/40 text-blue-200 rounded px-2 py-0.5 text-xs font-mono mr-3 mt-0.5 shrink-0">AT3</span>
                <span>Формула СУММПРОИЗВ для расчёта <strong className="text-white">ТМЦ</strong> (с автоопределением строки)</span>
              </li>
              <li className="flex items-start">
                <span className="bg-blue-600/40 text-blue-200 rounded px-2 py-0.5 text-xs font-mono mr-3 mt-0.5 shrink-0">AT4</span>
                <span>Формула <strong className="text-white">=AT2+AT3</strong> (ВСЕГО)</span>
              </li>
              <li className="flex items-start mt-2 pt-2 border-t border-white/10">
                <i className="fas fa-paint-brush text-yellow-400 mr-3 mt-0.5"></i>
                <span>Ко всем ячейкам: <strong className="text-red-400">красный цвет шрифта</strong> + <strong className="text-white">полужирное начертание</strong></span>
              </li>
              <li className="flex items-start">
                <i className="fas fa-mouse-pointer text-green-400 mr-3 mt-0.5"></i>
                <span>Курсор устанавливается на ячейку <strong className="text-white">AT4</strong></span>
              </li>
              <li className="flex items-start">
                <i className="fas fa-shield-alt text-green-400 mr-3 mt-0.5"></i>
                <span>Структура файла <strong className="text-white">не нарушается</strong>, скрытые колонки остаются скрытыми</span>
              </li>
            </ul>
          </div>
        </div>

        {/* Футер */}
        <div className="text-center mt-6 text-blue-400/60 text-sm">
          <p>Лист: «Корректировка 1» | Ячейки: AS2–AS4 (значения), AT2–AT4 (формулы)</p>
          <p className="mt-1 text-blue-400/40">Формулы адаптируются под последнюю заполненную строку автоматически</p>
        </div>
      </div>
    </div>
  );
}

export default App;
