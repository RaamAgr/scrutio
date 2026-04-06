'use client';

import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, FileSpreadsheet, Settings, Play, CheckCircle, XCircle, Loader2, Download, Copy, Eye, EyeOff, Sun, Moon, Search, Maximize, Minimize } from 'lucide-react';
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

  return (
    <div className="container animate-fade-in">
      <div style={{ position: 'fixed', top: '1.5rem', right: '1.5rem', zIndex: 100, display: 'flex', gap: '0.5rem' }}>
        <button onClick={toggleFullscreen} className="btn btn-secondary" style={{ padding: '0.5rem', borderRadius: '50%' }} title={isFullscreen ? 'Exit fullscreen' : 'Enter fullscreen'}>
          {isFullscreen ? <Minimize size={18} /> : <Maximize size={18} />}
        </button>
        <button onClick={toggleTheme} className="btn btn-secondary" style={{ padding: '0.5rem', borderRadius: '50%' }} title="Toggle theme">
          {theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}
        </button>
      </div>

      <header className="header" style={{ textAlign: 'center' }}>
        <h1 className="title">Scrutio</h1>
        <p className="subtitle" style={{ maxWidth: '600px', margin: '0 auto' }}>Examine every call with precision. Batch AI-powered quality analysis, surfaced in seconds.</p>
      </header>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '1.5rem', marginBottom: '1.5rem' }}>
        
        {/* Settings Card */}
        <div className="card">
          <div className="card-header">
            <Settings size={18} /> Configuration
          </div>
          <div className="card-body">
            <div className="form-group">
              <label className="label">API Key</label>
              <div style={{ position: 'relative' }}>
                <input 
                  type={showApiKey ? 'text' : 'password'}
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  placeholder="AIzaSy..." 
                  style={{ paddingRight: '2.5rem' }}
                />
                <button
                  type="button"
                  onClick={() => setShowApiKey(!showApiKey)}
                  style={{ position: 'absolute', right: '0.5rem', top: '50%', transform: 'translateY(-50%)', background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', padding: '0.2rem' }}
                >
                  {showApiKey ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1.25rem' }}>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <label className="label">Model</label>
                <select value={model} onChange={(e) => setModel(e.target.value)}>
                  <optgroup label="Latest Preview">
                    <option value="gemini-3.1-pro-preview">3.1 Pro</option>
                    <option value="gemini-3-flash-preview">3 Flash</option>
                    <option value="gemini-3.1-flash-lite-preview">3.1 Flash-Lite</option>
                  </optgroup>
                  <optgroup label="Stable">
                    <option value="gemini-2.5-flash">2.5 Flash</option>
                    <option value="gemini-1.5-pro">1.5 Pro</option>
                  </optgroup>
                </select>
              </div>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <label className="label">Temp ({temperature})</label>
                <input 
                  type="range" 
                  min="0" max="2" step="0.1" 
                  value={temperature}
                  onChange={(e) => setTemperature(parseFloat(e.target.value))}
                />
              </div>
            </div>

            <div className="form-group" style={{ marginBottom: 0 }}>
              <label className="label">Parallel Workers ({workers})</label>
              <input 
                type="range" 
                min="1" max="100" step="1" 
                value={workers}
                onChange={(e) => setWorkers(parseInt(e.target.value))}
              />
            </div>
          </div>
        </div>

        {/* Upload Card */}
        <div className="card">
          <div className="card-header">
            <FileSpreadsheet size={18} /> Data Source
          </div>
          <div className="card-body">
            <div className="dropzone" onClick={() => fileInputRef.current?.click()}>
              <input 
                type="file" 
                accept=".xlsx, .xls, .csv, .xlsm, .ods" 
                style={{ display: 'none' }} 
                ref={fileInputRef}
                onChange={handleFileUpload}
              />
              <UploadCloud size={32} color="var(--text-secondary)" style={{ marginBottom: '1rem' }} />
              <div style={{ fontWeight: 500, marginBottom: '0.25rem' }}>Upload Sheet or CSV</div>
              <div style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>Must contain <strong>recording_url</strong></div>
            </div>

            {records.length > 0 && (
              <div style={{ marginTop: '1rem', padding: '0.75rem 1rem', background: 'var(--bg-surface)', border: '1px solid var(--border-color)', borderRadius: '6px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', fontSize: '0.875rem' }}>
                  <CheckCircle size={16} color="var(--success)" />
                  {records.length} parsed
                </div>
                <button 
                  onClick={clearData} 
                  disabled={isProcessing}
                  style={{ background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: isProcessing ? 'not-allowed' : 'pointer', fontSize: '0.875rem', fontWeight: 500 }}
                >
                  Clear
                </button>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Prompt Card */}
      <div className="card" style={{ marginBottom: '1.5rem' }}>
        <div className="card-header">Evaluation Prompt</div>
        <div className="card-body" style={{ padding: '0.5rem 1rem' }}>
          <textarea 
            value={prompt}
            onChange={(e) => setPrompt(e.target.value)}
            placeholder="What should the AI analyze?" 
            style={{ width: '100%', minHeight: '120px', resize: 'vertical', border: 'none', background: 'transparent', padding: '0.5rem 0', outline: 'none', boxShadow: 'none' }}
          />
        </div>
      </div>

      {/* Action Buttons */}
      <div style={{ display: 'flex', gap: '0.75rem', flexWrap: 'wrap', marginBottom: '2rem' }}>
        <button 
          className="btn btn-primary" 
          style={{ flex: 1, padding: '0.75rem', minWidth: '150px' }}
          onClick={startEvaluation}
          disabled={isProcessing || records.length === 0}
        >
          {isProcessing ? <><Loader2 className="spinner" /> Processing...</> : <><Play size={18} /> Start Evaluation</>}
        </button>

        {!isProcessing && records.some(r => r.status === 'error') && (
          <button 
            className="btn btn-primary" 
            style={{ backgroundColor: 'rgba(229, 72, 77, 0.1)', color: 'var(--error)' }}
            onClick={startEvaluation}
          >
            Retry Failed
          </button>
        )}

        <button className="btn btn-secondary" onClick={exportResults} disabled={records.length === 0}>
          <Download size={16} /> Export
        </button>
        <button className="btn btn-secondary" onClick={copyAllResults} disabled={records.filter(r => r.result).length === 0}>
          <Copy size={16} /> Copy All
        </button>
      </div>
      
      {isProcessing && (
        <div style={{ marginBottom: '2rem' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.75rem', fontWeight: 500, marginBottom: '0.5rem', color: 'var(--text-secondary)', letterSpacing: '0.05em' }}>
            <span>EVALUATION PROGRESS</span>
            <span>{progress}%</span>
          </div>
          <div className="progress-bar">
            <div className="progress-fill" style={{ width: `${progress}%` }} />
          </div>
        </div>
      )}

      {/* Results Section */}
      {records.length > 0 && (
        <div style={{ marginBottom: '3rem' }} ref={resultsHeaderRef}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end', marginBottom: '1.5rem', flexWrap: 'wrap', gap: '1rem' }}>
            <h2 className="title" style={{ fontSize: '1.25rem', marginBottom: 0 }}>Results Matrix</h2>
            <div style={{ position: 'relative', width: '100%', maxWidth: '300px' }}>
              <Search size={16} style={{ position: 'absolute', left: '0.75rem', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }} />
              <input 
                type="text" 
                placeholder="Search name, mobile, url..." 
                value={searchQuery}
                onChange={(e) => { setSearchQuery(e.target.value); setCurrentPage(1); }}
                style={{ width: '100%', paddingLeft: '2.5rem', background: 'var(--bg-surface)', border: '1px solid var(--border-color)', height: '40px' }}
              />
            </div>
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
            {paginatedRecords.map((rec) => (
              <div key={rec.id} className="card">
                <div className="card-header" style={{ justifyContent: 'space-between', background: 'var(--bg-surface)' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
                    <span style={{ color: 'var(--text-muted)' }}>#{rec.id + 1}</span>
                    <a href={rec.url} target="_blank" rel="noreferrer" style={{ textDecoration: 'underline', maxWidth: '250px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {rec.url}
                    </a>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
                    {rec.usage && (
                      <div className="status-badge" style={{ color: 'var(--text-primary)' }}>
                        In: {rec.usage.promptTokenCount} &bull; Out: {rec.usage.candidatesTokenCount}
                      </div>
                    )}
                    <div className={`status-badge ${rec.status}`}>
                      {rec.status === 'processing' && <Loader2 size={12} className="spinner" />}
                      {rec.status === 'success' && <CheckCircle size={12} />}
                      {rec.status === 'error' && <XCircle size={12} />}
                      {rec.status}
                    </div>
                  </div>
                </div>
                {rec.result && (
                  <div className="card-body" style={{ position: 'relative', padding: '1.5rem' }}>
                    <button 
                      onClick={() => navigator.clipboard.writeText(String(rec.result)).then(() => alert('Copied!'))}
                      className="btn btn-secondary"
                      style={{ position: 'absolute', top: '1rem', right: '1rem', padding: '0.4rem' }}
                      title="Copy Output"
                    >
                      <Copy size={14} />
                    </button>
                    <div className="markdown-body">
                      <ReactMarkdown 
                        remarkPlugins={[remarkGfm]}
                        components={{
                          code({node, inline, className, children, ...props}: any) {
                            const match = /language-(\w+)/.exec(className || '');
                            return !inline && match ? (
                              <SyntaxHighlighter
                                style={vscDarkPlus}
                                language={match[1]}
                                PreTag="div"
                                customStyle={{ margin: 0, borderRadius: '4px', background: 'transparent' }}
                                {...props}
                              >
                                {String(children).replace(/\n$/, '')}
                              </SyntaxHighlighter>
                            ) : (
                              <code className={className} {...props}>
                                {children}
                              </code>
                            )
                          }
                        }}
                      >
                        {String(rec.result)}
                      </ReactMarkdown>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
          
          {/* Pagination Controls */}
          {totalPages > 1 && (
            <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '1rem', marginTop: '2rem' }}>
              <button 
                className="btn btn-secondary" 
                disabled={currentPage === 1} 
                onClick={() => setCurrentPage(p => p - 1)}
              >
                Previous
              </button>
              <div style={{ fontSize: '0.875rem', color: 'var(--text-secondary)' }}>
                Page <strong style={{ color: 'var(--text-primary)' }}>{currentPage}</strong> of {totalPages}
              </div>
              <button 
                className="btn btn-secondary" 
                disabled={currentPage === totalPages} 
                onClick={() => setCurrentPage(p => p + 1)}
              >
                Next
              </button>
            </div>
          )}
        </div>
      )}

      {/* Meta Analysis Section */}
      {records.length > 0 && records.some(r => r.status === 'success') && (
        <div className="card" style={{ border: '1px solid var(--text-secondary)' }}>
          <div className="card-header" style={{ background: 'var(--bg-surface)' }}>
            Level 2: Meta Analysis
          </div>
          <div className="card-body">
            <p className="subtitle" style={{ marginBottom: '1.25rem' }}>
              Aggregate all successful outputs into a macro-analysis engine to surface distinct trends.
            </p>

            <div className="form-group">
              <label className="label">Meta Prompt</label>
              <textarea 
                value={metaPrompt}
                onChange={(e) => setMetaPrompt(e.target.value)}
                style={{ width: '100%', minHeight: '80px', resize: 'vertical' }}
              />
            </div>

            <button 
              className="btn btn-primary" 
              style={{ width: '100%', padding: '0.75rem' }}
              onClick={startMetaAnalysis}
              disabled={isMetaProcessing}
            >
              {isMetaProcessing ? <><Loader2 className="spinner" /> Synthesizing Data...</> : <><Play size={18} /> Synthesize Macro Trends</>}
            </button>

            {metaResult && (
              <div style={{ marginTop: '2rem', paddingTop: '1.5rem', borderTop: '1px solid var(--border-color)', position: 'relative' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
                  <h3 className="title" style={{ fontSize: '1rem' }}>Synthesis Report</h3>
                  {metaUsage && (
                    <div className="status-badge" style={{ color: 'var(--text-primary)' }}>
                      Tokens In: {metaUsage.promptTokenCount} &bull; Out: {metaUsage.candidatesTokenCount}
                    </div>
                  )}
                </div>
                <button 
                  onClick={() => navigator.clipboard.writeText(metaResult).then(() => alert('Copied!'))}
                  className="btn btn-secondary"
                  style={{ position: 'absolute', top: '1.5rem', right: '0', padding: '0.4rem' }}
                >
                  <Copy size={14} />
                </button>
                <div className="markdown-body">
                  <ReactMarkdown 
                    remarkPlugins={[remarkGfm]}
                    components={{
                      code({node, inline, className, children, ...props}: any) {
                        const match = /language-(\w+)/.exec(className || '');
                        return !inline && match ? (
                          <SyntaxHighlighter
                            style={vscDarkPlus}
                            language={match[1]}
                            PreTag="div"
                            customStyle={{ margin: 0, borderRadius: '4px', background: 'transparent' }}
                            {...props}
                          >
                            {String(children).replace(/\n$/, '')}
                          </SyntaxHighlighter>
                        ) : (
                          <code className={className} {...props}>
                            {children}
                          </code>
                        )
                      }
                    }}
                  >
                    {metaResult}
                  </ReactMarkdown>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
