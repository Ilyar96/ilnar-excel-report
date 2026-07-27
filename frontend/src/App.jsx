import { useEffect, useMemo, useState } from 'react';

const API_BASE = import.meta.env.VITE_API_BASE || 'http://localhost:3000';

function createEntry(file) {
  return {
    id: `${file.name}-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    file,
    status: 'uploading',
    serverName: null,
    error: ''
  };
}

function App() {
  const [files, setFiles] = useState([]);
  const [resultName, setResultName] = useState('');
  const [titleFontSize, setTitleFontSize] = useState('11');
  const [cellFontSize, setCellFontSize] = useState('11');
  const [expensePeriod, setExpensePeriod] = useState(String(new Date().getMonth() + 1));
  const [statusMessage, setStatusMessage] = useState('Выберите Excel-файлы и загрузите их на сервер.');
  const [isUploading, setIsUploading] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [lastResultName, setLastResultName] = useState('');

  const defaultResultName = useMemo(() => {
    const month = new Date().toLocaleString('ru-RU', { month: 'long' });
    return `Акруалы за ${month}`;
  }, []);

  useEffect(() => {
    setResultName((current) => current || defaultResultName);
  }, [defaultResultName]);

  useEffect(() => {
    const cleanupOnUnload = () => {
      fetch(`${API_BASE}/cleanup`, {
        method: 'DELETE',
        keepalive: true
      }).catch(() => {});
    };

    window.addEventListener('unload', cleanupOnUnload);
    return () => window.removeEventListener('unload', cleanupOnUnload);
  }, []);

  const uploadFiles = async (selectedFiles) => {
    if (!selectedFiles.length) return;

    const existingKeys = new Set(files.map((item) => `${item.file.name}|${item.file.size}`));
    const duplicateKeys = new Set();
    const filteredFiles = selectedFiles.filter((file) => {
      const key = `${file.name}|${file.size}`;
      if (existingKeys.has(key) || duplicateKeys.has(key)) {
        duplicateKeys.add(key);
        return false;
      }
      duplicateKeys.add(key);
      return true;
    });

    if (!filteredFiles.length) {
      setStatusMessage('Нет новых файлов для загрузки. Дубликаты не добавляются.');
      return;
    }

    const pendingEntries = filteredFiles.map(createEntry);
    setFiles((prev) => [...prev, ...pendingEntries]);
    setIsUploading(true);
    setStatusMessage('Загружаем файлы на сервер...');

    for (const entry of pendingEntries) {
      const formData = new FormData();
      formData.append('files', entry.file);

      try {
        const response = await fetch(`${API_BASE}/files`, {
          method: 'POST',
          body: formData
        });

        const data = await response.json().catch(() => ({}));

        if (!response.ok) {
          throw new Error(data?.error || 'Не удалось загрузить файл');
        }

        const serverName = data.saved?.[0] || entry.file.name;

        setFiles((prev) =>
          prev.map((item) => (item.id === entry.id ? { ...item, status: 'uploaded', serverName } : item))
        );
      } catch (error) {
        setFiles((prev) =>
          prev.map((item) =>
            item.id === entry.id ? { ...item, status: 'error', error: error.message } : item
          )
        );
      }
    }

    setIsUploading(false);
    setStatusMessage('Загрузка завершена. Можно генерировать акруал.');
  };

  const handleFileSelection = async (event) => {
    const selectedFiles = Array.from(event.target.files || []);
    if (!selectedFiles.length) return;

    await uploadFiles(selectedFiles);
    event.target.value = '';
  };

  const handleRemoveFile = async (id) => {
    const itemToRemove = files.find((item) => item.id === id);
    setFiles((prev) => prev.filter((item) => item.id !== id));

    if (itemToRemove?.serverName && itemToRemove.status === 'uploaded') {
      fetch(`${API_BASE}/files/${encodeURIComponent(itemToRemove.serverName)}`, {
        method: 'DELETE'
      }).catch(() => {});
    }
  };

  const handleGenerate = async () => {
    const uploadedFiles = files.filter((item) => item.status === 'uploaded');

    if (!uploadedFiles.length) {
      setStatusMessage('Сначала загрузите хотя бы один файл.');
      return;
    }

    setIsGenerating(true);
    setStatusMessage('Генерируем акруал...');

    const periodValue = expensePeriod.trim();
    const periodNumber = periodValue === '' ? undefined : Number(periodValue);

    if (periodValue !== '' && (!Number.isInteger(periodNumber) || periodNumber < 1 || periodNumber > 12)) {
      setStatusMessage('Период расхода должен быть целым числом от 1 до 12.');
      setIsGenerating(false);
      return;
    }

    const payload = {
      files: uploadedFiles.map((item) => item.serverName || item.file.name),
      resultName: resultName.trim() || defaultResultName,
      titleFontSize: Number(titleFontSize) || 11,
      cellFontSize: Number(cellFontSize) || 11,
      period: periodNumber
    };

    try {
      const response = await fetch(`${API_BASE}/generate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (!response.ok) {
        const text = await response.text();
        throw new Error(text || 'Не удалось сгенерировать акруал');
      }

      const blob = await response.blob();
      const finalName = `${payload.resultName.replace(/\.xlsx$/i, '')}.xlsx`;
      const downloadUrl = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = finalName;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(downloadUrl);

      setLastResultName(finalName);
      setStatusMessage('Акруал успешно создан. Входные файлы удалены с сервера.');
      setFiles([]);
    } catch (error) {
      setStatusMessage(error.message || 'Ошибка генерации');
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="app-shell">
      <div className="card">
        <h1>Генератор акруалов</h1>
        <p className="subtitle">Загружайте несколько таблиц Excel, собирайте отчёт и получайте готовый файл.</p>

        <div className="form-grid">
          <label>
            <span>Название результата</span>
            <input
              value={resultName}
              onChange={(event) => setResultName(event.target.value)}
              placeholder={defaultResultName}
            />
          </label>
        </div>
        <div className="form-grid">
					<label>
            <span>Размер шрифта заголовков</span>
            <input
              type="number"
              min="8"
              max="32"
              value={titleFontSize}
              onChange={(event) => setTitleFontSize(event.target.value)}
            />
          </label>

          <label>
            <span>Размер шрифта ячеек</span>
            <input
              type="number"
              min="8"
              max="32"
              value={cellFontSize}
              onChange={(event) => setCellFontSize(event.target.value)}
            />
          </label>

          <label>
            <span>Период расхода</span>
            <input
              type="number"
              min="1"
              max="12"
              value={expensePeriod}
              onChange={(event) => setExpensePeriod(event.target.value)}
            />
          </label>
        </div>

        <label className="file-picker">
          <span>Выбрать Excel-файлы</span>
          <input type="file" multiple accept=".xlsx" onChange={handleFileSelection} />
        </label>

        <div className="actions">
          <button onClick={handleGenerate} disabled={isUploading || isGenerating || !files.some((item) => item.status === 'uploaded')}>
            {isGenerating ? 'Генерация...' : 'Сгенерировать и скачать акруал'}
          </button>
        </div>

        <p className="status">{statusMessage}</p>

        {files.length > 0 && (
          <ul className="file-list">
            {files.map((item) => (
              <li key={item.id} className={`file-item ${item.status}`}>
                <div>
                  <span>{item.file.name}</span>
                  <small>
                    {item.status === 'uploading' && 'загружается'}
                    {item.status === 'uploaded' && 'загружено'}
                    {item.status === 'error' && item.error}
                  </small>
                </div>
                <button type="button" onClick={() => handleRemoveFile(item.id)}>
                  Удалить
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}

export default App;
