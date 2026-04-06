'use client';

import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, FileSpreadsheet, Settings, Play, CheckCircle, XCircle, Loader2, Download, Copy, Eye, EyeOff, Sun, Moon, Search, Maximize, Minimize, ChevronUp, ChevronDown } from 'lucide-react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { Prism as SyntaxHighlighter } from 'react-syntax-highlighter';
import { vscDarkPlus } from 'react-syntax-highlighter/dist/cjs/styles/prism';

interface RecordItem {
  id: number;
  url: string;
  metadata?: {
    date?: string;
    mobile_number?: string;
    name?: string;
    duration?: string;
    [key: string]: any;
  };
  status: 'pending' | 'processing' | 'success' | 'error';
  result?: string;
  usage?: {
    promptTokenCount?: number;
    candidatesTokenCount?: number;
    totalTokenCount?: number;
  };
}

export default function QCApp() {
  const [apiKey, setApiKey] = useState('');
  const [model, setModel] = useState('gemini-2.5-flash');
  const [prompt, setPrompt] = useState('Please evaluate this call based on standard quality metrics:\n1. Greeting\n2. Tone and empathy\n3. Issue resolution\n4. Closing');
  const [temperature, setTemperature] = useState(0.7);
  const [workers, setWorkers] = useState(3);
  const [thinkingBudget, setThinkingBudget] = useState<number>(-1);
  
  const [records, setRecords] = useState<RecordItem[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [currentPage, setCurrentPage] = useState(1);
  const [searchQuery, setSearchQuery] = useState('');
  const [showApiKey, setShowApiKey] = useState(false);
  const [theme, setTheme] = useState('dark');
  const [metaPrompt, setMetaPrompt] = useState('Analyze the following quality control transcript evaluations. Identify key trends, common negative patterns, and overall performance metrics.');
  const [metaResult, setMetaResult] = useState('');
  const [isMetaProcessing, setIsMetaProcessing] = useState(false);
  const [metaUsage, setMetaUsage] = useState<{promptTokenCount?: number, candidatesTokenCount?: number}>();
  const [isFullscreen, setIsFullscreen] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const resultsHeaderRef = useRef<HTMLDivElement>(null);
  const mainPanelRef = useRef<HTMLElement>(null);
  const [isAtBottom, setIsAtBottom] = useState(false);

  // Derived values (not hooks — safe after all hooks are declared)
  const itemsPerPage = 10;

  const filteredRecords = records.filter(r => {
    if (!searchQuery.trim()) return true;
    const query = searchQuery.toLowerCase();
    const matchesUrl = r.url.toLowerCase().includes(query);
    const matchesName = r.metadata?.name?.toLowerCase().includes(query);
    const matchesMobile = (r.metadata?.mobile_number || r.metadata?.['mobile number'])?.toString().toLowerCase().includes(query);
    const matchesResult = r.result?.toLowerCase().includes(query);
    return matchesUrl || matchesName || matchesMobile || !!matchesResult;
  });

  const totalPages = Math.ceil(filteredRecords.length / itemsPerPage);
  const paginatedRecords = filteredRecords.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage);

  useEffect(() => {
    if (currentPage > totalPages && totalPages > 0) {
      setCurrentPage(totalPages);
    }
  }, [filteredRecords.length, totalPages, currentPage]);

  useEffect(() => {
    if (resultsHeaderRef.current && records.length > 0) {
      resultsHeaderRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, [currentPage]);

  useEffect(() => {
    const panel = mainPanelRef.current;
    if (!panel) return;
    const onScroll = () => {
      const nearBottom = panel.scrollTop + panel.clientHeight >= panel.scrollHeight - 80;
      setIsAtBottom(nearBottom);
    };
    panel.addEventListener('scroll', onScroll);
    return () => panel.removeEventListener('scroll', onScroll);
  }, []);

  const scrollPanel = () => {
    const panel = mainPanelRef.current;
    if (!panel) return;
    if (isAtBottom) {
      panel.scrollTo({ top: 0, behavior: 'smooth' });
    } else {
      panel.scrollTo({ top: panel.scrollHeight, behavior: 'smooth' });
    }
  };

  const toggleTheme = () => {
    const nextTheme = theme === 'dark' ? 'light' : 'dark';
    setTheme(nextTheme);
    document.documentElement.setAttribute('data-theme', nextTheme);
  };

  const toggleFullscreen = () => {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen();
      setIsFullscreen(true);
    } else {
      document.exitFullscreen();
      setIsFullscreen(false);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const isCsv = file.name.toLowerCase().endsWith('.csv');
    const reader = new FileReader();
    
    reader.onload = (event) => {
      try {
        let workbook;
        
        if (isCsv) {
          const textData = event.target?.result as string;
          workbook = XLSX.read(textData, { type: 'string' });
        } else {
          const data = new Uint8Array(event.target?.result as ArrayBuffer);
          workbook = XLSX.read(data, { type: 'array', cellDates: true });
        }

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        const newRecords: RecordItem[] = [];
        json.forEach((row, index) => {
          let urlValue = '';
          const metadata: any = {};
          
          for (let key in row) {
            const lowerKey = key.toLowerCase().trim();
            if (lowerKey === 'recording_url' || lowerKey.includes('recording url')) {
              urlValue = row[key];
            } else if (['date', 'mobile_number', 'mobile number', 'name', 'duration'].includes(lowerKey)) {
              let val = row[key];
              if (lowerKey === 'date' && typeof val === 'number') {
                try {
                  const parsed = XLSX.SSF.parse_date_code(val);
                  const hh = String(parsed.H).padStart(2, '0');
                  const mm = String(parsed.M).padStart(2, '0');
                  const ss = String(parsed.S).padStart(2, '0');
                  const timeString = (hh === '00' && mm === '00' && ss === '00') ? '' : ` ${hh}:${mm}:${ss}`;
                  val = `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}${timeString}`;
                } catch (err) {}
              } else if (val instanceof Date) {
                const hh = String(val.getUTCHours()).padStart(2, '0');
                const mm = String(val.getUTCMinutes()).padStart(2, '0');
                const ss = String(val.getUTCSeconds()).padStart(2, '0');
                const timeString = (hh === '00' && mm === '00' && ss === '00') ? '' : ` ${hh}:${mm}:${ss}`;
                val = `${val.getUTCFullYear()}-${String(val.getUTCMonth() + 1).padStart(2, '0')}-${String(val.getUTCDate()).padStart(2, '0')}${timeString}`;
              }
              metadata[lowerKey.replace(' ', '_')] = val;
            }
          }
          if (urlValue) {
            newRecords.push({
              id: index,
              url: String(urlValue).trim(),
              metadata: Object.keys(metadata).length > 0 ? metadata : undefined,
              status: 'pending'
            });
          }
        });

        setRecords(newRecords);
      } catch (err) {
        alert('Failed to parse file. Ensure it is a valid spreadsheet/CSV and contains a recording_url column.');
        console.error(err);
      }
    };

    if (isCsv) {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file);
    }
  };

  const startEvaluation = async () => {
    if (!apiKey) {
      alert('Please enter your Gemini API Key first.');
      return;
    }
    if (records.length === 0) {
      alert('Please upload a spreadsheet or CSV file with recording URLs.');
      return;
    }

    setIsProcessing(true);
    setProgress(0);

    let currentIndex = 0;
    let completedCount = 0;

    const processNext = async () => {
      while (true) {
        const index = currentIndex++;
        if (index >= records.length) break;

        const currentRecord = records[index];
        if (currentRecord.status === 'success') {
          completedCount++;
          setProgress(Math.round((completedCount / records.length) * 100));
          continue; 
        }

        setRecords(prev => {
          const next = [...prev];
          next[index] = { ...next[index], status: 'processing' };
          return next;
        });

        try {
          const res = await fetch('/api/evaluate', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({
              url: currentRecord.url,
              prompt,
              apiKey,
              model,
              temperature,
              thinkingBudget,
            })
          });

          const data = await res.json();
          
          setRecords(prev => {
            const next = [...prev];
            next[index] = {
              ...next[index],
              status: data.success ? 'success' : 'error',
              result: data.success ? data.result : (data.error || 'Unknown error'),
              usage: data.usage || undefined
            };
            return next;
          });
        } catch (err: any) {
          setRecords(prev => {
            const next = [...prev];
            next[index] = {
              ...next[index],
              status: 'error',
              result: err.message || 'Failed to communicate with server.'
            };
            return next;
          });
        }

        completedCount++;
        setProgress(Math.round((completedCount / records.length) * 100));
      }
    };

    const workerPool = [];
    const actualWorkers = Math.min(workers, records.length);
    for (let i = 0; i < actualWorkers; i++) {
      workerPool.push(processNext());
    }

    await Promise.all(workerPool);
    setIsProcessing(false);
  };

  const clearData = () => {
    setRecords([]);
    setProgress(0);
    setCurrentPage(1);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const exportResults = () => {
    if (records.length === 0) return;

    const exportData = records.map(r => ({
      'Recording URL': r.url,
      'Name': r.metadata?.name || '',
      'Mobile Number': r.metadata?.mobile_number || r.metadata?.['mobile number'] || '',
      'Date': r.metadata?.date || '',
      'Duration': r.metadata?.duration || '',
      'Status': r.status,
      'Input Tokens': r.usage?.promptTokenCount || 0,
      'Output Tokens': r.usage?.candidatesTokenCount || 0,
      'Result/Transcript': r.result || ''
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Evaluations");
    XLSX.writeFile(workbook, "qc_evaluations_export.xlsx");
  };

  const copyAllResults = () => {
    const outputs = records
      .filter(r => r.result)
      .map(r => `URL: ${r.url}\nResult:\n${r.result}`)
      .join('\n\n=========================================\n\n');
    
    if (outputs) {
      navigator.clipboard.writeText(outputs).then(() => alert('Copied all results to clipboard!'));
    }
  };

  const startMetaAnalysis = async () => {
    if (records.length === 0 || isMetaProcessing) return;
    
    setIsMetaProcessing(true);
    setMetaResult('');
    setMetaUsage(undefined);

    const successfulRecords = records.filter(r => r.status === 'success' && r.result);
    let feed = '';
    successfulRecords.forEach((r, idx) => {
      feed += `\n\n--- CALL #${idx + 1} (${r.url}) ---\n${r.result}`;
    });

    try {
      const res = await fetch('/api/meta', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          feed,
          prompt: metaPrompt,
          apiKey,
          model,
          temperature: 0.4
        })
      });
      const data = await res.json();
      if (data.success) {
        setMetaResult(data.result);
        setMetaUsage(data.usage);
      } else {
        setMetaResult(`**Error**: ${data.error}`);
      }
    } catch (err: any) {
      setMetaResult(`**Failed to reach server**: ${err.message}`);
    } finally {
      setIsMetaProcessing(false);
    }
  };

  const [showThinking, setShowThinking] = useState(false);

  const codeComponents = {
    code({node, inline, className, children, ...props}: any) {
      const match = /language-(\w+)/.exec(className || '');
      return !inline && match ? (
        <SyntaxHighlighter style={vscDarkPlus} language={match[1]} PreTag="div"
          customStyle={{ margin: 0, borderRadius: '4px', background: 'transparent' }} {...props}>
          {String(children).replace(/\n$/, '')}
        </SyntaxHighlighter>
      ) : (
        <code className={className} {...props}>{children}</code>
      );
    }
  };

  return (
    <div className="app-shell animate-fade-in">

      {/* ── FIXED TOP-RIGHT CONTROLS ── */}
      <div style={{ position: 'fixed', top: '1rem', right: '1rem', zIndex: 200, display: 'flex', gap: '0.5rem' }}>
        <button onClick={toggleFullscreen} className="btn btn-secondary" style={{ padding: '0.45rem', borderRadius: '50%' }} title={isFullscreen ? 'Exit fullscreen' : 'Enter fullscreen'}>
          {isFullscreen ? <Minimize size={15} /> : <Maximize size={15} />}
        </button>
        <button onClick={toggleTheme} className="btn btn-secondary" style={{ padding: '0.45rem', borderRadius: '50%' }} title="Toggle theme">
          {theme === 'dark' ? <Sun size={15} /> : <Moon size={15} />}
        </button>
      </div>

      {/* ── LEFT SIDEBAR ── */}
      <aside className="sidebar">

        {/* Branding */}
        <div>
          <h1 className="title" style={{ fontSize: '1.2rem' }}>Scrutio</h1>
          <p className="subtitle" style={{ fontSize: '0.72rem' }}>AI-powered call quality analysis</p>
        </div>

        {/* Configuration */}
        <div className="card">
          <div className="card-header" style={{ padding: '0.75rem 1rem' }}><Settings size={14} /> Configuration</div>
          <div className="card-body" style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem', padding: '0.75rem 1rem' }}>

            <div className="form-group" style={{ marginBottom: 0 }}>
              <label className="label">API Key</label>
              <div style={{ position: 'relative' }}>
                <input type={showApiKey ? 'text' : 'password'} value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)} placeholder="AIzaSy..."
                  style={{ paddingRight: '2.5rem' }} />
                <button type="button" onClick={() => setShowApiKey(!showApiKey)}
                  style={{ position: 'absolute', right: '0.5rem', top: '50%', transform: 'translateY(-50%)', background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', padding: '0.2rem' }}>
                  {showApiKey ? <EyeOff size={15} /> : <Eye size={15} />}
                </button>
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.75rem' }}>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <label className="label">Model</label>
                <select value={model} onChange={(e) => setModel(e.target.value)}>
                  <optgroup label="Latest Preview">
                    <option value="gemini-3.1-pro-preview">3.1 Pro</option>
                    <option value="gemini-3-flash-preview">3 Flash</option>
                    <option value="gemini-3.1-flash-lite-preview">3.1 Lite</option>
                  </optgroup>
                  <optgroup label="Stable">
                    <option value="gemini-2.5-flash">2.5 Flash</option>
                    <option value="gemini-1.5-pro">1.5 Pro</option>
                  </optgroup>
                </select>
              </div>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <label className="label">Temp ({temperature})</label>
                <input type="range" min="0" max="2" step="0.1" value={temperature}
                  onChange={(e) => setTemperature(parseFloat(e.target.value))} />
              </div>
            </div>

            <div className="form-group" style={{ marginBottom: 0 }}>
              <label className="label">Workers ({workers})</label>
              <input type="range" min="1" max="100" step="1" value={workers}
                onChange={(e) => setWorkers(parseInt(e.target.value))} />
            </div>

            {/* Thinking — collapsible */}
            <div style={{ borderTop: '1px solid var(--border-color)', paddingTop: '0.75rem' }}>
              <button type="button" onClick={() => setShowThinking((v: boolean) => !v)}
                style={{ display: 'flex', alignItems: 'center', background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', fontSize: '0.72rem', fontWeight: 500, letterSpacing: '0.05em', textTransform: 'uppercase', padding: 0, width: '100%', justifyContent: 'space-between' }}>
                <span>Thinking {thinkingBudget >= 0 && <span style={{ color: 'var(--text-primary)', textTransform: 'none', letterSpacing: 0 }}>({thinkingBudget === 0 ? 'Off' : thinkingBudget <= 1024 ? 'Low' : thinkingBudget <= 8192 ? 'Med' : 'High'})</span>}</span>
                <span style={{ transform: showThinking ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s', display: 'inline-block' }}>▾</span>
              </button>
              {showThinking && (
                <div style={{ marginTop: '0.75rem' }}>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '0.4rem' }}>
                    {([['Off', 0], ['Low', 1024], ['Med', 8192], ['High', 24576]] as const).map(([label, val]) => (
                      <button key={label} type="button" onClick={() => setThinkingBudget(val === thinkingBudget ? -1 : val)}
                        style={{ padding: '0.35rem 0', fontSize: '0.7rem', fontWeight: 500, border: '1px solid', borderRadius: '4px', cursor: 'pointer', transition: 'all 0.15s',
                          borderColor: thinkingBudget === val ? 'var(--text-primary)' : 'var(--border-color)',
                          background: thinkingBudget === val ? 'var(--text-primary)' : 'var(--bg-surface)',
                          color: thinkingBudget === val ? 'var(--bg-card)' : 'var(--text-secondary)' }}>
                        {label}
                      </button>
                    ))}
                  </div>
                  <span style={{ fontSize: '0.68rem', color: 'var(--text-muted)', marginTop: '0.35rem', display: 'block' }}>
                    {thinkingBudget < 0 ? 'Model default' : thinkingBudget === 0 ? 'Disabled' : `${thinkingBudget.toLocaleString()} tokens`}
                  </span>
                </div>
              )}
            </div>

          </div>
        </div>

        {/* Data Source */}
        <div className="card">
          <div className="card-header" style={{ padding: '0.75rem 1rem' }}><FileSpreadsheet size={14} /> Data Source</div>
          <div className="card-body" style={{ padding: '0.75rem 1rem' }}>
            <div className="dropzone" onClick={() => fileInputRef.current?.click()}>
              <input type="file" accept=".xlsx,.xls,.csv,.xlsm,.ods" style={{ display: 'none' }}
                ref={fileInputRef} onChange={handleFileUpload} />
              <UploadCloud size={26} color="var(--text-secondary)" style={{ marginBottom: '0.6rem' }} />
              <div style={{ fontWeight: 500, fontSize: '0.85rem', marginBottom: '0.15rem' }}>Upload Sheet or CSV</div>
              <div style={{ fontSize: '0.7rem', color: 'var(--text-secondary)' }}>Needs <strong>recording_url</strong> column</div>
            </div>
            {records.length > 0 && (
              <div style={{ marginTop: '0.75rem', padding: '0.55rem 0.75rem', background: 'var(--bg-surface)', border: '1px solid var(--border-color)', borderRadius: '6px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.4rem', fontSize: '0.78rem' }}>
                  <CheckCircle size={13} color="var(--success)" /> {records.length} records
                </div>
                <button onClick={clearData} disabled={isProcessing}
                  style={{ background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: isProcessing ? 'not-allowed' : 'pointer', fontSize: '0.78rem' }}>
                  Clear
                </button>
              </div>
            )}
          </div>
        </div>

      </aside>

      {/* ── RIGHT MAIN PANEL ── */}
      <main className="main-panel" ref={mainPanelRef}>

        {/* Scroll to top/bottom button */}
        <button
          onClick={scrollPanel}
          title={isAtBottom ? 'Scroll to top' : 'Scroll to bottom'}
          style={{
            position: 'fixed',
            bottom: '1.5rem',
            right: '1.5rem',
            zIndex: 200,
            width: '2.25rem',
            height: '2.25rem',
            borderRadius: '50%',
            background: 'var(--bg-card)',
            border: '1px solid var(--border-color)',
            color: 'var(--text-secondary)',
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            boxShadow: '0 2px 8px rgba(0,0,0,0.3)',
            transition: 'all 0.15s ease',
          }}
        >
          {isAtBottom ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
        </button>

        {/* Prompt */}
        <div className="card">
          <div className="card-header">Evaluation Prompt</div>
          <div className="card-body" style={{ padding: '0.5rem 1rem' }}>
            <textarea value={prompt} onChange={(e) => setPrompt(e.target.value)}
              placeholder="What should the AI analyze?"
              style={{ width: '100%', minHeight: '100px', resize: 'vertical', border: 'none', background: 'transparent', padding: '0.5rem 0', outline: 'none', boxShadow: 'none' }} />
          </div>
        </div>

        {/* Action Buttons */}
        <div style={{ display: 'flex', gap: '0.75rem', flexWrap: 'wrap' }}>
          <button className="btn btn-primary" style={{ flex: 1, padding: '0.7rem', minWidth: '140px' }}
            onClick={startEvaluation} disabled={isProcessing || records.length === 0}>
            {isProcessing ? <><Loader2 className="spinner" /> Processing...</> : <><Play size={16} /> Start Evaluation</>}
          </button>
          {!isProcessing && records.some(r => r.status === 'error') && (
            <button className="btn btn-primary" style={{ backgroundColor: 'rgba(229,72,77,0.1)', color: 'var(--error)' }} onClick={startEvaluation}>Retry Failed</button>
          )}
          <button className="btn btn-secondary" onClick={exportResults} disabled={records.length === 0}><Download size={15} /> Export</button>
          <button className="btn btn-secondary" onClick={copyAllResults} disabled={records.filter(r => r.result).length === 0}><Copy size={15} /> Copy All</button>
        </div>

        {/* Progress */}
        {isProcessing && (
          <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.7rem', fontWeight: 500, marginBottom: '0.4rem', color: 'var(--text-secondary)', letterSpacing: '0.06em' }}>
              <span>EVALUATION PROGRESS</span><span>{progress}%</span>
            </div>
            <div className="progress-bar"><div className="progress-fill" style={{ width: `${progress}%` }} /></div>
          </div>
        )}

        {/* Results */}
        {records.length > 0 && (
          <div ref={resultsHeaderRef}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem', flexWrap: 'wrap', gap: '0.75rem' }}>
              <h2 className="title" style={{ fontSize: '1.1rem', marginBottom: 0 }}>Results Matrix</h2>
              <div style={{ position: 'relative' }}>
                <Search size={14} style={{ position: 'absolute', left: '0.65rem', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }} />
                <input type="text" placeholder="Search results..." value={searchQuery}
                  onChange={(e) => { setSearchQuery(e.target.value); setCurrentPage(1); }}
                  style={{ paddingLeft: '2.2rem', background: 'var(--bg-surface)', border: '1px solid var(--border-color)', height: '34px', width: '220px', fontSize: '0.8rem' }} />
              </div>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
              {paginatedRecords.map((rec) => (
                <div key={rec.id} className="card">
                  <div className="card-header" style={{ justifyContent: 'space-between' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', minWidth: 0 }}>
                      <span style={{ color: 'var(--text-muted)', flexShrink: 0 }}>#{rec.id + 1}</span>
                      <a href={rec.url} target="_blank" rel="noreferrer"
                        style={{ textDecoration: 'underline', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontSize: '0.8rem' }}>
                        {rec.url}
                      </a>
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', flexShrink: 0 }}>
                      {rec.usage && (
                        <div className="status-badge" style={{ color: 'var(--text-primary)' }}>
                          {rec.usage.promptTokenCount}↑ {rec.usage.candidatesTokenCount}↓
                        </div>
                      )}
                      <div className={`status-badge ${rec.status}`}>
                        {rec.status === 'processing' && <Loader2 size={11} className="spinner" />}
                        {rec.status === 'success' && <CheckCircle size={11} />}
                        {rec.status === 'error' && <XCircle size={11} />}
                        {rec.status}
                      </div>
                    </div>
                  </div>
                  {rec.result && (
                    <div className="card-body" style={{ position: 'relative', padding: '1.25rem' }}>
                      <button onClick={() => navigator.clipboard.writeText(String(rec.result)).then(() => alert('Copied!'))}
                        className="btn btn-secondary" style={{ position: 'absolute', top: '0.75rem', right: '0.75rem', padding: '0.3rem' }} title="Copy">
                        <Copy size={13} />
                      </button>
                      <div className="markdown-body">
                        <ReactMarkdown remarkPlugins={[remarkGfm]} components={codeComponents}>{String(rec.result)}</ReactMarkdown>
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>

            {totalPages > 1 && (
              <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '1rem', marginTop: '1.5rem' }}>
                <button className="btn btn-secondary" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Previous</button>
                <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>Page <strong style={{ color: 'var(--text-primary)' }}>{currentPage}</strong> of {totalPages}</span>
                <button className="btn btn-secondary" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</button>
              </div>
            )}
          </div>
        )}

        {/* Meta Analysis */}
        {records.length > 0 && records.some(r => r.status === 'success') && (
          <div className="card">
            <div className="card-header">Level 2: Meta Analysis</div>
            <div className="card-body">
              <p className="subtitle" style={{ marginBottom: '1rem' }}>Aggregate all successful outputs to surface macro trends.</p>
              <div className="form-group">
                <label className="label">Meta Prompt</label>
                <textarea value={metaPrompt} onChange={(e) => setMetaPrompt(e.target.value)}
                  style={{ width: '100%', minHeight: '70px', resize: 'vertical' }} />
              </div>
              <button className="btn btn-primary" style={{ width: '100%', padding: '0.7rem' }}
                onClick={startMetaAnalysis} disabled={isMetaProcessing}>
                {isMetaProcessing ? <><Loader2 className="spinner" /> Synthesizing...</> : <><Play size={16} /> Synthesize Macro Trends</>}
              </button>
              {metaResult && (
                <div style={{ marginTop: '1.5rem', paddingTop: '1.25rem', borderTop: '1px solid var(--border-color)', position: 'relative' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.75rem' }}>
                    <h3 className="title" style={{ fontSize: '0.95rem' }}>Synthesis Report</h3>
                    {metaUsage && <div className="status-badge" style={{ color: 'var(--text-primary)' }}>{metaUsage.promptTokenCount}↑ {metaUsage.candidatesTokenCount}↓</div>}
                  </div>
                  <button onClick={() => navigator.clipboard.writeText(metaResult).then(() => alert('Copied!'))}
                    className="btn btn-secondary" style={{ position: 'absolute', top: '1.25rem', right: 0, padding: '0.3rem' }}>
                    <Copy size={13} />
                  </button>
                  <div className="markdown-body">
                    <ReactMarkdown remarkPlugins={[remarkGfm]} components={codeComponents}>{metaResult}</ReactMarkdown>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Empty state */}
        {records.length === 0 && (
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: 'var(--text-muted)', gap: '0.5rem', userSelect: 'none', minHeight: '200px' }}>
            <FileSpreadsheet size={40} strokeWidth={1} />
            <p style={{ fontSize: '0.875rem' }}>Upload a spreadsheet to get started</p>
          </div>
        )}

      </main>
    </div>
  );
}
