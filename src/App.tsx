import { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

interface CategoryTotal {
  category: string;
  total: number;
}

export default function App() {
  const [fileName, setFileName] = useState<string>('');
  const [rawData, setRawData] = useState<{ category: string; amount: number }[]>([]);
  const [resultData, setResultData] = useState<CategoryTotal[]>([]);
  const [error, setError] = useState<string>('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [isDragOver, setIsDragOver] = useState(false);

  const processFile = useCallback((file: File) => {
    setError('');
    setRawData([]);
    setResultData([]);
    setIsProcessing(true);

    const validTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel',
    ];
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

    if (!validTypes.includes(file.type) && !validExtensions.includes(ext)) {
      setError('Пожалуйста, выберите файл Excel (.xlsx, .xls) или CSV (.csv)');
      setIsProcessing(false);
      return;
    }

    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json<(string | number)[]>(worksheet, { header: 1 });

        if (jsonData.length === 0) {
          setError('Файл пустой или не содержит данных');
          setIsProcessing(false);
          return;
        }

        // Parse first 2 columns
        const parsed: { category: string; amount: number }[] = [];
        for (let i = 0; i < jsonData.length; i++) {
          const row = jsonData[i];
          if (!row || row.length === 0) continue;

          const category = row[0];
          const amount = row[1];

          // Skip header rows or empty rows
          if (category === undefined || category === null || category === '') continue;
          if (amount === undefined || amount === null || amount === '') continue;

          const categoryStr = String(category).trim();
          const amountNum = typeof amount === 'number' ? amount : parseFloat(String(amount).replace(',', '.'));

          if (isNaN(amountNum)) continue;

          parsed.push({ category: categoryStr, amount: amountNum });
        }

        if (parsed.length === 0) {
          setError('Не удалось найти данные в первых двух столбцах. Убедитесь, что первый столбец содержит категории, а второй — суммы.');
          setIsProcessing(false);
          return;
        }

        setRawData(parsed);

        // Group by category and sum
        const categoryMap = new Map<string, number>();
        for (const item of parsed) {
          const current = categoryMap.get(item.category) || 0;
          categoryMap.set(item.category, current + item.amount);
        }

        const result: CategoryTotal[] = Array.from(categoryMap.entries())
          .map(([category, total]) => ({ category, total: Math.round(total * 100) / 100 }))
          .sort((a, b) => b.total - a.total);

        setResultData(result);
        setIsProcessing(false);
      } catch (err) {
        console.error(err);
        setError('Ошибка при чтении файла. Убедитесь, что файл является корректным Excel-файлом.');
        setIsProcessing(false);
      }
    };

    reader.onerror = () => {
      setError('Ошибка при чтении файла');
      setIsProcessing(false);
    };

    reader.readAsArrayBuffer(file);
  }, []);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  }, [processFile]);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = () => {
    setIsDragOver(false);
  };

  const downloadResult = () => {
    if (resultData.length === 0) return;

    const wsData: (string | number)[][] = [
      ['Категория', 'Итоговая сумма'],
      ...resultData.map(item => [item.category, item.total]),
    ];

    // Add total row
    const grandTotal = resultData.reduce((sum, item) => sum + item.total, 0);
    wsData.push(['', '']);
    wsData.push(['ИТОГО', Math.round(grandTotal * 100) / 100]);

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Set column widths
    ws['!cols'] = [
      { wch: 40 },
      { wch: 20 },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Итоги по категориям');

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, 'итоги_по_категориям.xlsx');
  };

  const reset = () => {
    setFileName('');
    setRawData([]);
    setResultData([]);
    setError('');
    setIsProcessing(false);
  };

  const formatNumber = (num: number) => {
    return num.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };

  const grandTotal = resultData.reduce((sum, item) => sum + item.total, 0);

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-950 to-slate-900 text-white">
      {/* Header */}
      <header className="border-b border-white/10 backdrop-blur-sm bg-white/5">
        <div className="max-w-6xl mx-auto px-4 py-4 flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-400 to-cyan-500 flex items-center justify-center shadow-lg shadow-emerald-500/20">
            <svg className="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
          </div>
          <div>
            <h1 className="text-xl font-bold bg-gradient-to-r from-emerald-300 to-cyan-300 bg-clip-text text-transparent">
              Excel Обработчик
            </h1>
            <p className="text-xs text-slate-400">Анализ категорий расходов и доходов</p>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8">
        {/* Upload Section */}
        <div className="mb-8">
          <div
            className={`relative border-2 border-dashed rounded-2xl p-12 text-center transition-all duration-300 cursor-pointer
              ${isDragOver
                ? 'border-emerald-400 bg-emerald-400/10 scale-[1.02]'
                : 'border-white/20 hover:border-white/40 hover:bg-white/5'
              }
              ${fileName ? 'opacity-60' : ''}
            `}
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onClick={() => !fileName && document.getElementById('fileInput')?.click()}
          >
            <input
              id="fileInput"
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileChange}
              className="hidden"
            />

            {isProcessing ? (
              <div className="flex flex-col items-center gap-4">
                <div className="w-16 h-16 border-4 border-emerald-400/30 border-t-emerald-400 rounded-full animate-spin"></div>
                <p className="text-lg text-slate-300">Обработка файла...</p>
              </div>
            ) : fileName ? (
              <div className="flex flex-col items-center gap-3">
                <div className="w-16 h-16 rounded-2xl bg-emerald-500/20 flex items-center justify-center">
                  <svg className="w-8 h-8 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                </div>
                <p className="text-lg font-medium text-emerald-300">Файл загружен: {fileName}</p>
                <p className="text-sm text-slate-400">
                  Найдено записей: {rawData.length} | Уникальных категорий: {resultData.length}
                </p>
                <button
                  onClick={(e) => { e.stopPropagation(); reset(); }}
                  className="mt-2 px-4 py-2 text-sm rounded-lg bg-white/10 hover:bg-white/20 transition-colors"
                >
                  Загрузить другой файл
                </button>
              </div>
            ) : (
              <div className="flex flex-col items-center gap-4">
                <div className="w-16 h-16 rounded-2xl bg-white/10 flex items-center justify-center">
                  <svg className="w-8 h-8 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                  </svg>
                </div>
                <div>
                  <p className="text-lg font-medium text-slate-200">
                    Перетащите Excel-файл сюда или{' '}
                    <span className="text-emerald-400 underline underline-offset-2">выберите файл</span>
                  </p>
                  <p className="text-sm text-slate-400 mt-1">
                    Поддерживаются форматы: .xlsx, .xls, .csv
                  </p>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Error */}
        {error && (
          <div className="mb-8 p-4 rounded-xl bg-red-500/10 border border-red-500/30 flex items-start gap-3">
            <svg className="w-6 h-6 text-red-400 flex-shrink-0 mt-0.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
            <div>
              <p className="font-medium text-red-300">Ошибка</p>
              <p className="text-sm text-red-200/80 mt-0.5">{error}</p>
            </div>
          </div>
        )}

        {/* Results */}
        {resultData.length > 0 && (
          <div className="space-y-6">
            {/* Summary Cards */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="rounded-xl bg-white/5 border border-white/10 p-5 backdrop-blur-sm">
                <p className="text-sm text-slate-400">Всего записей</p>
                <p className="text-2xl font-bold text-white mt-1">{rawData.length}</p>
              </div>
              <div className="rounded-xl bg-white/5 border border-white/10 p-5 backdrop-blur-sm">
                <p className="text-sm text-slate-400">Уникальных категорий</p>
                <p className="text-2xl font-bold text-emerald-300 mt-1">{resultData.length}</p>
              </div>
              <div className="rounded-xl bg-white/5 border border-white/10 p-5 backdrop-blur-sm">
                <p className="text-sm text-slate-400">Общая сумма</p>
                <p className="text-2xl font-bold text-cyan-300 mt-1">{formatNumber(grandTotal)}</p>
              </div>
            </div>

            {/* Table */}
            <div className="rounded-xl bg-white/5 border border-white/10 backdrop-blur-sm overflow-hidden">
              <div className="px-6 py-4 border-b border-white/10 flex items-center justify-between">
                <h2 className="text-lg font-semibold text-white">Итоги по категориям</h2>
                <button
                  onClick={downloadResult}
                  className="flex items-center gap-2 px-5 py-2.5 rounded-xl bg-gradient-to-r from-emerald-500 to-cyan-500 hover:from-emerald-400 hover:to-cyan-400 text-white font-medium text-sm transition-all duration-200 shadow-lg shadow-emerald-500/20 hover:shadow-emerald-500/40 hover:scale-105 active:scale-95"
                >
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                  </svg>
                  Скачать Excel
                </button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead>
                    <tr className="border-b border-white/10">
                      <th className="px-6 py-3 text-left text-xs font-semibold text-slate-400 uppercase tracking-wider">
                        №
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-semibold text-slate-400 uppercase tracking-wider">
                        Категория
                      </th>
                      <th className="px-6 py-3 text-right text-xs font-semibold text-slate-400 uppercase tracking-wider">
                        Сумма
                      </th>
                      <th className="px-6 py-3 text-right text-xs font-semibold text-slate-400 uppercase tracking-wider">
                        Доля
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-white/5">
                    {resultData.map((item, index) => {
                      const percentage = grandTotal !== 0 ? (item.total / grandTotal) * 100 : 0;
                      return (
                        <tr key={item.category} className="hover:bg-white/5 transition-colors">
                          <td className="px-6 py-3.5 text-sm text-slate-500">{index + 1}</td>
                          <td className="px-6 py-3.5 text-sm font-medium text-white">{item.category}</td>
                          <td className="px-6 py-3.5 text-sm text-right font-mono text-emerald-300">
                            {formatNumber(item.total)}
                          </td>
                          <td className="px-6 py-3.5 text-sm text-right">
                            <div className="flex items-center justify-end gap-2">
                              <div className="w-20 h-2 rounded-full bg-white/10 overflow-hidden">
                                <div
                                  className="h-full rounded-full bg-gradient-to-r from-emerald-400 to-cyan-400"
                                  style={{ width: `${Math.min(percentage, 100)}%` }}
                                ></div>
                              </div>
                              <span className="text-slate-400 w-14 text-right">{percentage.toFixed(1)}%</span>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr className="border-t-2 border-white/20 bg-white/5">
                      <td className="px-6 py-4"></td>
                      <td className="px-6 py-4 text-sm font-bold text-white">ИТОГО</td>
                      <td className="px-6 py-4 text-sm text-right font-mono font-bold text-cyan-300">
                        {formatNumber(grandTotal)}
                      </td>
                      <td className="px-6 py-4 text-sm text-right font-bold text-white">100%</td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>

            {/* Raw Data Preview */}
            <details className="rounded-xl bg-white/5 border border-white/10 backdrop-blur-sm overflow-hidden">
              <summary className="px-6 py-4 cursor-pointer hover:bg-white/5 transition-colors text-sm font-medium text-slate-300">
                📋 Показать исходные данные ({rawData.length} записей)
              </summary>
              <div className="border-t border-white/10 overflow-x-auto max-h-96 overflow-y-auto">
                <table className="w-full">
                  <thead className="sticky top-0 bg-slate-900/90 backdrop-blur-sm">
                    <tr className="border-b border-white/10">
                      <th className="px-6 py-3 text-left text-xs font-semibold text-slate-400 uppercase">№</th>
                      <th className="px-6 py-3 text-left text-xs font-semibold text-slate-400 uppercase">Категория</th>
                      <th className="px-6 py-3 text-right text-xs font-semibold text-slate-400 uppercase">Сумма</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-white/5">
                    {rawData.map((item, index) => (
                      <tr key={index} className="hover:bg-white/5 transition-colors">
                        <td className="px-6 py-2 text-xs text-slate-500">{index + 1}</td>
                        <td className="px-6 py-2 text-sm text-slate-300">{item.category}</td>
                        <td className="px-6 py-2 text-sm text-right font-mono text-slate-300">
                          {formatNumber(item.amount)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </details>
          </div>
        )}

        {/* Instructions */}
        {resultData.length === 0 && !error && (
          <div className="mt-8 rounded-xl bg-white/5 border border-white/10 p-6 backdrop-blur-sm">
            <h3 className="text-lg font-semibold text-white mb-4 flex items-center gap-2">
              <svg className="w-5 h-5 text-cyan-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              Как это работает
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              <div className="flex gap-3">
                <div className="w-8 h-8 rounded-lg bg-emerald-500/20 flex items-center justify-center flex-shrink-0 text-emerald-400 font-bold text-sm">1</div>
                <div>
                  <p className="font-medium text-white text-sm">Загрузите файл</p>
                  <p className="text-xs text-slate-400 mt-1">Выберите Excel-файл с данными о расходах и доходах</p>
                </div>
              </div>
              <div className="flex gap-3">
                <div className="w-8 h-8 rounded-lg bg-emerald-500/20 flex items-center justify-center flex-shrink-0 text-emerald-400 font-bold text-sm">2</div>
                <div>
                  <p className="font-medium text-white text-sm">Автоматический анализ</p>
                  <p className="text-xs text-slate-400 mt-1">Столбец A — категории, столбец B — суммы. Данные группируются автоматически</p>
                </div>
              </div>
              <div className="flex gap-3">
                <div className="w-8 h-8 rounded-lg bg-emerald-500/20 flex items-center justify-center flex-shrink-0 text-emerald-400 font-bold text-sm">3</div>
                <div>
                  <p className="font-medium text-white text-sm">Скачайте результат</p>
                  <p className="text-xs text-slate-400 mt-1">Получите итоговый Excel-файл с уникальными категориями и суммами</p>
                </div>
              </div>
            </div>

            <div className="mt-6 p-4 rounded-lg bg-white/5 border border-white/10">
              <p className="text-xs text-slate-400">
                <span className="font-semibold text-slate-300">Формат входного файла:</span> Первый столбец (A) — описание категории расхода/дохода, второй столбец (B) — числовая сумма. 
                Строки могут содержать повторяющиеся категории — они будут объединены с суммированием.
              </p>
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-white/10 mt-12">
        <div className="max-w-6xl mx-auto px-4 py-4 text-center text-xs text-slate-500">
          Excel Обработчик — Анализ категорий расходов и доходов
        </div>
      </footer>
    </div>
  );
}
