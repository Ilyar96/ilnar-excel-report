import { useEffect, useMemo, useState } from 'react';

const API_BASE = import.meta.env.VITE_API_BASE || 'http://localhost:3000';

function createEntry(file) {
  return {
    id: `${file.name}-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    file,
    status: 'uploading',
    serverName: null,
    error: '',
    progress: 0
  };
}

function App() {
  const [files, setFiles] = useState([]);
  const [resultName, setResultName] = useState('');
  const [titleFontSize, setTitleFontSize] = useState('11');
  const [cellFontSize, setCellFontSize] = useState('11');
  const [useExpensePeriod, setUseExpensePeriod] = useState(false);
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
      await uploadSingleFile(entry);
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

  const uploadSingleFile = async (entry) => {
    setFiles((prev) =>
      prev.map((item) =>
        item.id === entry.id ? { ...item, status: 'uploading', error: '', progress: 0 } : item
      )
    );

    return new Promise((resolve) => {
      const xhr = new XMLHttpRequest();
      const formData = new FormData();
      formData.append('files', entry.file);

      xhr.open('POST', `${API_BASE}/files`);
      xhr.responseType = 'text';

      xhr.upload.onprogress = (event) => {
        if (!event.lengthComputable) return;
        const percent = Math.round((event.loaded / event.total) * 100);
        setFiles((prev) =>
          prev.map((item) => (item.id === entry.id ? { ...item, progress: percent } : item))
        );
      };

      xhr.onload = () => {
        let data = {};
        try {
          data = JSON.parse(xhr.responseText);
        } catch (err) {
          data = {};
        }

        if (xhr.status >= 200 && xhr.status < 300) {
          const serverName = data.saved?.[0] || entry.file.name;
          setFiles((prev) =>
            prev.map((item) =>
              item.id === entry.id
                ? { ...item, status: 'uploaded', serverName, progress: 100, error: '' }
                : item
            )
          );
        } else {
          const errorMessage = data?.error || `Не удалось загрузить файл (${xhr.status})`;
          setFiles((prev) =>
            prev.map((item) =>
              item.id === entry.id
                ? { ...item, status: 'error', error: errorMessage, progress: 0 }
                : item
            )
          );
        }
        resolve();
      };

      xhr.onerror = () => {
        setFiles((prev) =>
          prev.map((item) =>
            item.id === entry.id
              ? { ...item, status: 'error', error: 'Сетевая ошибка при загрузке файла', progress: 0 }
              : item
          )
        );
        resolve();
      };

      xhr.send(formData);
    });
  };

  const handleRetryUpload = async (item) => {
    setIsUploading(true);
    await uploadSingleFile(item);
    setIsUploading(false);
  };

  const handleGenerate = async () => {
    const uploadedFiles = files.filter((item) => item.status === 'uploaded');

    if (!uploadedFiles.length) {
      setStatusMessage('Сначала загрузите хотя бы один файл.');
      return;
    }

    setIsGenerating(true);
    setStatusMessage('Генерируем акруал...');

    const periodValue = useExpensePeriod ? expensePeriod.trim() : '';
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
      period: useExpensePeriod && periodValue !== '' ? periodNumber : undefined
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

          <div>
						<label className="checkbox-label">
							<input
								className="checkbox-input"
								type="checkbox"
								checked={useExpensePeriod}
								onChange={() => setUseExpensePeriod((current) => !current)}
							/>
							<span>Период расхода</span>
						</label>
            <input
              type="number"
              min="1"
              max="12"
              value={expensePeriod}
              onChange={(event) => setExpensePeriod(event.target.value)}
              disabled={!useExpensePeriod}
            />
          </div>
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
                  <small style={{marginLeft: 5}}>
                    <span style={{color:"blue"}}>
                    	{item.status === 'uploading' && (
	                      item.progress > 0 ? `загружается — ${item.progress}%` : 'загружается...'
	                    )}
                    </span>
                    <span style={{color:"green"}}>{item.status === 'uploaded' && 'загружено'}</span>
                    <span style={{color:"red"}}>{item.status === 'error' && item.error}</span>
                  </small>
                </div>
                <div className="file-actions">
                  {item.status === 'error' && (
                    <button
                      type="button"
                      title="Повторить"
                      className="retry-button"
                      onClick={() => handleRetryUpload(item)}
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M12 5V1L8 5l4 4V6c3.31 0 6 2.69 6 6 0 1.07-.28 2.08-.78 2.96l1.46 1.46C19.74 14.32 20 13.19 20 12c0-4.42-3.58-8-8-8zm-6.34 2.46L4.22 6.32C3.58 7.68 3.2 9.3 3.2 11c0 4.42 3.58 8 8 8v4l4-4-4-4v3c-3.31 0-6-2.69-6-6 0-1.07.28-2.08.78-2.96z" fill="currentColor" />
                      </svg>
                    </button>
                  )}
                  <button type="button" onClick={() => handleRemoveFile(item.id)}>
                    Удалить
                  </button>
                </div>
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}

export default App;
