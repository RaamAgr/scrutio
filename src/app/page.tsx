'use client';

import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, FileSpreadsheet, Settings, Play, CheckCircle, XCircle, Loader2, Download, Copy, Eye, EyeOff, Sun, Moon, Search, Maximize, Minimize, ChevronUp, ChevronDown, Music2, FileText, X } from 'lucide-react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { Prism as SyntaxHighlighter } from 'react-syntax-highlighter';
import { vscDarkPlus } from 'react-syntax-highlighter/dist/cjs/styles/prism';

interface RecordItem {
  id: number;
  // For spreadsheet mode: the remote URL
  url: string;
  // For audio-file mode: the local File reference
  audioFile?: File;
  // Display label (filename for audio mode, URL for spreadsheet mode)
  label?: string;
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
  originalRow?: any;
  step2Result?: string;
  step2Status?: 'pending' | 'processing' | 'success' | 'error';
  step2Usage?: {
    promptTokenCount?: number;
    candidatesTokenCount?: number;
  };
}

export default function QCApp() {
  const [mounted, setMounted] = useState(false);
  useEffect(() => {
    setMounted(true);
  }, []);

  const [apiKey, setApiKey] = useState('');
  const [anthropicApiKey, setAnthropicApiKey] = useState('');
  const [showAnthropicApiKey, setShowAnthropicApiKey] = useState(false);
  const [vertexServiceAccount, setVertexServiceAccount] = useState('');
  const [model, setModel] = useState('gemini-3-flash-preview');
  const [metaModel, setMetaModel] = useState('gemini-2.5-pro');
  const [prompt, setPrompt] = useState('You are an expert AI transcriptionist specializing in audio-to-text conversion for customer service operations. Your task is to transcribe the provided audio file with high accuracy, strictly separating the dialogue between the Call Center Agent and the Customer.\n\n### Formatting Rules\n\n1. Speaker Labels: Use exactly Agent: and Customer: to identify the speakers. Do not use names or other variations.\n2. Timestamping: Insert a timestamp in the format [HH:MM:SS] at the beginning of every turn or when a significant pause occurs.\n3. Verbatim Accuracy: Transcribe the spoken words exactly as they are uttered. Do not fix grammar, omit filler words (like um, uh, ah), or smooth out sentences.\n4. Unclear Audio: If a word or phrase is completely unintelligible, use [inaudible] instead of guessing.');
  const [temperature, setTemperature] = useState(0.7);
  const [workers, setWorkers] = useState(3);
  const [thinkingBudget, setThinkingBudget] = useState<number>(-1);

  // Input mode: 'spreadsheet' (URLs from xlsx/csv) or 'audio' (direct file upload)
  const [inputMode, setInputMode] = useState<'spreadsheet' | 'audio'>('spreadsheet');

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
  const [metaAnalysisMode, setMetaAnalysisMode] = useState<'aggregate' | 'row'>('aggregate');
  const [metaProgress, setMetaProgress] = useState(0);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [showExportModal, setShowExportModal] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const audioFilesInputRef = useRef<HTMLInputElement>(null);
  const audioFolderInputRef = useRef<HTMLInputElement>(null);
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

  // ── Audio Files Upload Handler ──────────────────────────────────────
  const handleAudioFilesUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const AUDIO_EXTS = /\.(mp3|wav|ogg|m4a|aac|flac|webm|mp4|opus|wma)$/i;
    const audioFiles = Array.from(files).filter(f => AUDIO_EXTS.test(f.name));

    if (audioFiles.length === 0) {
      alert('No supported audio files found. Supported: mp3, wav, ogg, m4a, aac, flac, webm, mp4, opus, wma');
      return;
    }

    const newRecords: RecordItem[] = audioFiles.map((file, idx) => ({
      id: idx,
      url: '',        // not used in audio mode
      audioFile: file,
      label: file.name,
      status: 'pending',
    }));

    setRecords(newRecords);
  };

  // ── Spreadsheet Upload Handler ────────────────────────────────────────
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploading(true);

    const isCsv = file.name.toLowerCase().endsWith('.csv');
    const reader = new FileReader();
    
    reader.onload = (event) => {
      setTimeout(() => {
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

          const isResultKey = (lowerKey: string) => {
            const hasCore = lowerKey.includes('result') || lowerKey.includes('transcript') || lowerKey.includes('output');
            if (!hasCore) return false;
            const hasExcluded = lowerKey.includes('token') || 
                                lowerKey.includes('status') || 
                                lowerKey.includes('count') || 
                                lowerKey.includes('cost') ||
                                lowerKey.includes('length') ||
                                lowerKey.includes('duration') ||
                                lowerKey.includes('date') ||
                                lowerKey.includes('step 2') ||
                                lowerKey.includes('step2') ||
                                lowerKey.includes('extraction');
            return !hasExcluded;
          };

          const isStep2Key = (lowerKey: string) => {
            return lowerKey === 'step 2 output' || 
                   lowerKey.includes('step 2 result') || 
                   lowerKey.includes('step 2 output') || 
                   lowerKey.includes('extraction output') || 
                   lowerKey.includes('extraction result');
          };

          const cleanMobileNumber = (val: any): string => {
            if (val === undefined || val === null) return '';
            const str = String(val).trim();
            const digits = str.replace(/\D/g, '');
            if (digits.length === 12) {
              return digits.slice(2);
            }
            return str;
          };

          const newRecords: RecordItem[] = [];
          json.forEach((row, index) => {
            let urlValue = '';
            let resultValue = '';
            let labelValue = '';
            let step2Value = '';
            const metadata: any = {};
            
            for (let key in row) {
              const lowerKey = key.toLowerCase().trim();
              if (lowerKey === 'recording_url' || lowerKey.includes('recording url')) {
                urlValue = row[key];
              }
              if (lowerKey === 'recording_name' || lowerKey.includes('recording name') || lowerKey === 'recording' || lowerKey === 'name') {
                labelValue = row[key];
                if (!urlValue && labelValue) {
                  urlValue = labelValue;
                }
              }
              if (isResultKey(lowerKey)) {
                resultValue = row[key];
              }
              if (isStep2Key(lowerKey)) {
                step2Value = row[key];
              }
              if (['date', 'mobile_number', 'mobile number', 'name', 'duration'].includes(lowerKey)) {
                let val = row[key];
                if (lowerKey === 'mobile_number' || lowerKey === 'mobile number') {
                  val = cleanMobileNumber(val);
                } else if (lowerKey === 'date' && typeof val === 'number') {
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
            if (urlValue || resultValue || step2Value) {
              newRecords.push({
                id: index,
                url: urlValue ? String(urlValue).trim() : `Record #${index + 1}`,
                label: labelValue ? String(labelValue).trim() : undefined,
                metadata: Object.keys(metadata).length > 0 ? metadata : undefined,
                status: resultValue ? 'success' : 'pending',
                result: resultValue ? String(resultValue).trim() : undefined,
                originalRow: row,
                step2Result: step2Value ? String(step2Value).trim() : undefined,
                step2Status: step2Value ? 'success' : undefined
              });
            }
          });

          setRecords(newRecords);
        } catch (err) {
          alert('Failed to parse file. Ensure it is a valid spreadsheet/CSV and contains a recording_url column.');
          console.error(err);
        } finally {
          setIsUploading(false);
        }
      }, 80);
    };

    reader.onerror = () => {
      setIsUploading(false);
      alert('Error reading the file.');
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
      alert(inputMode === 'audio'
        ? 'Please upload audio files first.'
        : 'Please upload a spreadsheet or CSV file with recording URLs.');
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
          let data: any;

          if (inputMode === 'audio' && currentRecord.audioFile) {
            // Read file as base64
            const base64 = await new Promise<string>((resolve, reject) => {
              const reader = new FileReader();
              reader.onload = () => {
                const result = reader.result as string;
                // Strip data URL prefix
                resolve(result.split(',')[1]);
              };
              reader.onerror = reject;
              reader.readAsDataURL(currentRecord.audioFile!);
            });

            const res = await fetch('/api/evaluate-file', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                fileData: base64,
                fileName: currentRecord.audioFile.name,
                mimeType: currentRecord.audioFile.type || undefined,
                prompt,
                apiKey,
                model: model === 'claude-sonnet-4-6' ? 'gemini-3-flash-preview' : model,
                temperature,
                thinkingBudget,
              }),
            });
            data = await res.json();
          } else {
            const res = await fetch('/api/evaluate', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                url: currentRecord.url,
                prompt,
                apiKey,
                model: model === 'claude-sonnet-4-6' ? 'gemini-3-flash-preview' : model,
                temperature,
                thinkingBudget,
              }),
            });
            data = await res.json();
          }

          setRecords(prev => {
            const next = [...prev];
            next[index] = {
              ...next[index],
              status: data.success ? 'success' : 'error',
              result: data.success ? data.result : (data.error || 'Unknown error'),
              usage: data.usage || undefined,
            };
            return next;
          });
        } catch (err: any) {
          setRecords(prev => {
            const next = [...prev];
            next[index] = {
              ...next[index],
              status: 'error',
              result: err.message || 'Failed to communicate with server.',
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
    if (fileInputRef.current) fileInputRef.current.value = '';
    if (audioFilesInputRef.current) audioFilesInputRef.current.value = '';
    if (audioFolderInputRef.current) audioFolderInputRef.current.value = '';
  };

  const exportResults = () => {
    if (records.length === 0) return;

    let exportData: object[];

    if (inputMode === 'audio') {
      // 2-column output for audio file mode
      exportData = records.map(r => ({
        'Recording Name': r.label || r.audioFile?.name || r.url,
        'Result/Output': r.result || '',
      }));
    } else {
      // Full spreadsheet export for URL mode
      exportData = records.map(r => {
        const rowData = { ...(r.originalRow || {}) };

        let urlKey = 'Recording URL';
        let resultKey = 'Result/Transcript';
        let statusKey = 'Status';
        let inputTokensKey = 'Input Tokens';
        let outputTokensKey = 'Output Tokens';

        if (r.originalRow) {
          for (const key of Object.keys(r.originalRow)) {
            const lowerKey = key.toLowerCase().trim();
            if (lowerKey === 'recording_url' || lowerKey.includes('recording url')) {
              urlKey = key;
            } else if (lowerKey === 'result' || lowerKey === 'transcript' || lowerKey === 'output' || 
                       lowerKey.includes('result/transcript') || lowerKey.includes('result/output') || 
                       lowerKey.includes('result') || lowerKey.includes('output') || lowerKey.includes('transcript')) {
              resultKey = key;
            } else if (lowerKey === 'status') {
              statusKey = key;
            } else if (lowerKey === 'input_tokens' || lowerKey === 'input tokens') {
              inputTokensKey = key;
            } else if (lowerKey === 'output_tokens' || lowerKey === 'output tokens') {
              outputTokensKey = key;
            }
          }
        } else {
          // Fallback if no original row exists
          if (r.metadata?.name) rowData['Name'] = r.metadata.name;
          if (r.metadata?.mobile_number) rowData['Mobile Number'] = r.metadata.mobile_number;
          if (r.metadata?.date) rowData['Date'] = r.metadata.date;
          if (r.metadata?.duration) rowData['Duration'] = r.metadata.duration;
        }

        rowData[urlKey] = r.url;
        rowData[resultKey] = r.result || '';
        rowData[statusKey] = r.status;
        rowData[inputTokensKey] = r.usage?.promptTokenCount || 0;
        rowData[outputTokensKey] = r.usage?.candidatesTokenCount || 0;

        // Add Step 2 output if present
        if (r.step2Result !== undefined) {
          let step2Key = 'Step 2 Output';
          if (r.originalRow) {
            for (const key of Object.keys(r.originalRow)) {
              const lowerKey = key.toLowerCase().trim();
              if (lowerKey === 'step 2 output' || lowerKey.includes('step 2 result') || lowerKey.includes('step 2 output') || lowerKey.includes('extraction output') || lowerKey.includes('extraction result')) {
                step2Key = key;
                break;
              }
            }
          }
          rowData[step2Key] = r.step2Result;
        }

        return rowData;
      });
    }

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Evaluations');
    
    // Safe client-side ArrayBuffer and Blob download
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', 'scrutio_export.xlsx');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const exportAsText = () => {
    if (records.length === 0) return;

    let textContent = '';
    records.forEach((r, idx) => {
      let entryName = '';
      if (inputMode === 'audio') {
        entryName = r.label || r.audioFile?.name || r.url || '';
      } else {
        if (r.url) {
          try {
            const urlObj = new URL(r.url);
            entryName = urlObj.pathname.split('/').pop() || r.url;
          } catch {
            entryName = r.url;
          }
        }
      }

      textContent += `================================================================================\n`;
      textContent += `ENTRY #${idx + 1}${entryName ? `: ${entryName}` : ''}\n`;
      textContent += `================================================================================\n`;
      textContent += `Status: ${r.status}\n`;
      
      if (r.metadata && Object.keys(r.metadata).length > 0) {
        textContent += `Metadata:\n`;
        if (r.metadata.name) textContent += `  Name: ${r.metadata.name}\n`;
        if (r.metadata.mobile_number !== undefined) textContent += `  Mobile: ${r.metadata.mobile_number}\n`;
        if (r.metadata.date) textContent += `  Date: ${r.metadata.date}\n`;
        if (r.metadata.duration) textContent += `  Duration: ${r.metadata.duration}\n`;
      }
      
      textContent += `\n--------------------------------------------------------------------------------\n`;
      textContent += `${r.result || 'No transcript/result available.'}\n`;
      
      if (r.step2Result !== undefined) {
        textContent += `\n--------------------------------------------------------------------------------\n`;
        textContent += `STEP 2 OUTPUT / EXTRACTION:\n`;
        textContent += `--------------------------------------------------------------------------------\n`;
        textContent += `${r.step2Result || 'No output available.'}\n`;
      }
      
      textContent += `\n\n`;
    });
    
    const blob = new Blob([textContent], { type: 'text/plain;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', 'scrutio_export.txt');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
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

  const copyAllMobileNumbers = () => {
    const mobiles = records
      .map(r => r.metadata?.mobile_number || r.metadata?.['mobile number'])
      .filter(m => m !== undefined && m !== null && String(m).trim() !== '')
      .map(m => String(m).trim());
    
    const uniqueMobiles = Array.from(new Set(mobiles));
    
    if (uniqueMobiles.length === 0) {
      alert('No mobile numbers found to copy!');
      return;
    }
    
    const output = uniqueMobiles.join(', ');
    navigator.clipboard.writeText(output).then(() => {
      alert(`Copied ${uniqueMobiles.length} unique mobile numbers to clipboard!`);
    });
  };

  const startMetaAnalysis = async () => {
    if (records.length === 0 || isMetaProcessing) return;
    
    setIsMetaProcessing(true);
    setMetaProgress(0);

    const successfulRecords = records.filter(r => r.status === 'success' && r.result);

    if (metaAnalysisMode === 'aggregate') {
      setMetaResult('');
      setMetaUsage(undefined);

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
            anthropicApiKey,
            vertexServiceAccount,
            model: metaModel,
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
    } else {
      // Row-by-row extraction mode
      if (successfulRecords.length === 0) {
        alert('No successful transcripts/results found from Step 1 to process.');
        setIsMetaProcessing(false);
        return;
      }

      // Initialize all step2Status to pending
      setRecords(prev => {
        return prev.map(r => {
          if (r.status === 'success' && r.result) {
            return { ...r, step2Status: 'pending', step2Result: undefined, step2Usage: undefined };
          }
          return r;
        });
      });

      let currentIndex = 0;
      let completedCount = 0;

      const processNextRow = async () => {
        while (true) {
          const idx = currentIndex++;
          if (idx >= successfulRecords.length) break;

          const record = successfulRecords[idx];
          // Find actual index in records state array
          const mainIndex = records.findIndex(r => r.id === record.id);
          if (mainIndex === -1) continue;

          setRecords(prev => {
            const next = [...prev];
            next[mainIndex] = { ...next[mainIndex], step2Status: 'processing' };
            return next;
          });

          try {
            const interpolatePrompt = (promptText: string, originalRow: any): string => {
              if (!originalRow) return promptText;
              return promptText.replace(/\{\{([^}]+)\}\}/g, (match, columnName) => {
                const targetKey = columnName.trim().toLowerCase();
                const matchingKey = Object.keys(originalRow).find(
                  key => key.toLowerCase().trim() === targetKey
                );
                return matchingKey !== undefined ? String(originalRow[matchingKey]) : '';
              });
            };

            const interpolatedMetaPrompt = interpolatePrompt(metaPrompt, record.originalRow);

            const res = await fetch('/api/meta', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                feed: record.result,
                prompt: interpolatedMetaPrompt,
                apiKey,
                anthropicApiKey,
                vertexServiceAccount,
                model: metaModel,
                temperature: 0.4
              })
            });
            const data = await res.json();

            setRecords(prev => {
              const next = [...prev];
              next[mainIndex] = {
                ...next[mainIndex],
                step2Status: data.success ? 'success' : 'error',
                step2Result: data.success ? data.result : (data.error || 'Unknown error'),
                step2Usage: data.usage || undefined
              };
              return next;
            });
          } catch (err: any) {
            setRecords(prev => {
              const next = [...prev];
              next[mainIndex] = {
                ...next[mainIndex],
                step2Status: 'error',
                step2Result: err.message || 'Failed to communicate with server.'
              };
              return next;
            });
          }

          completedCount++;
          setMetaProgress(Math.round((completedCount / successfulRecords.length) * 100));
        }
      };

      // Run concurrency using the 'workers' state
      const workerPool = [];
      const actualWorkers = Math.min(workers, successfulRecords.length);
      for (let i = 0; i < actualWorkers; i++) {
        workerPool.push(processNextRow());
      }

      await Promise.all(workerPool);
      setIsMetaProcessing(false);
    }
  };

  const downloadMetaReport = () => {
    if (!metaResult) return;
    const blob = new Blob([metaResult], { type: 'text/markdown;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', 'scrutio_synthesis_report.md');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
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

  if (!mounted) {
    return (
      <div style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        minHeight: '100vh',
        background: 'var(--bg-app, #090d16)',
        color: 'var(--text-secondary, #94a3b8)',
        fontFamily: 'system-ui, sans-serif'
      }}>
        <Loader2 size={32} className="spinner" style={{ color: 'var(--accent, #3b82f6)', marginBottom: '1rem', animation: 'spin 1s linear infinite' }} />
        <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>Initializing Scrutio...</span>
      </div>
    );
  }

  return (
    <div className="app-shell animate-fade-in">
      {isUploading && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          background: 'rgba(9, 13, 22, 0.85)',
          backdropFilter: 'blur(5px)',
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
          color: 'var(--text-primary)',
          fontFamily: 'system-ui, sans-serif'
        }}>
          <Loader2 size={40} className="spinner" style={{ color: 'var(--accent, #3b82f6)', marginBottom: '1rem', animation: 'spin 1s linear infinite' }} />
          <span style={{ fontSize: '0.9rem', fontWeight: 500 }}>Parsing spreadsheet and preparing calls...</span>
        </div>
      )}

      {showExportModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          background: 'rgba(9, 13, 22, 0.85)',
          backdropFilter: 'blur(5px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
          fontFamily: 'system-ui, sans-serif'
        }}>
          <div className="card" style={{
            width: '100%',
            maxWidth: '440px',
            background: 'var(--bg-card, #0f172a)',
            border: '1px solid var(--border-color, #1e293b)',
            borderRadius: '12px',
            boxShadow: '0 8px 32px rgba(0,0,0,0.5)',
            position: 'relative',
            overflow: 'hidden',
            animation: 'scale-up 0.2s ease-out'
          }}>
            {/* Header */}
            <div style={{
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              padding: '1.25rem 1.5rem',
              borderBottom: '1px solid var(--border-color)',
            }}>
              <h3 style={{ margin: 0, fontSize: '1.05rem', fontWeight: 600, color: 'var(--text-primary)' }}>Export Format</h3>
              <button 
                onClick={() => setShowExportModal(false)}
                style={{
                  background: 'transparent',
                  border: 'none',
                  color: 'var(--text-secondary)',
                  cursor: 'pointer',
                  padding: '0.2rem',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  borderRadius: '50%',
                  transition: 'background 0.15s'
                }}
                onMouseEnter={(e) => e.currentTarget.style.background = 'rgba(255,255,255,0.05)'}
                onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
              >
                <X size={16} />
              </button>
            </div>

            {/* Body / Selector Cards */}
            <div style={{ padding: '1.5rem', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
              <p style={{ margin: 0, fontSize: '0.82rem', color: 'var(--text-secondary)', marginBottom: '0.25rem' }}>
                Select your preferred file format for the export. The downloaded file name will start with <code style={{ background: 'var(--bg-surface)', padding: '0.1rem 0.3rem', borderRadius: '3px' }}>scrutio_</code>.
              </p>

              {/* Excel Option */}
              <button
                onClick={() => {
                  exportResults();
                  setShowExportModal(false);
                }}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '1rem',
                  width: '100%',
                  padding: '1rem',
                  background: 'var(--bg-surface, #1e293b)',
                  border: '1px solid var(--border-color)',
                  borderRadius: '8px',
                  color: 'var(--text-primary)',
                  cursor: 'pointer',
                  textAlign: 'left',
                  transition: 'all 0.15s ease'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.borderColor = 'var(--accent, #3b82f6)';
                  e.currentTarget.style.background = 'rgba(59, 130, 246, 0.05)';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.borderColor = 'var(--border-color)';
                  e.currentTarget.style.background = 'var(--bg-surface)';
                }}
              >
                <div style={{
                  background: 'rgba(34, 197, 94, 0.15)',
                  color: '#22c55e',
                  padding: '0.6rem',
                  borderRadius: '6px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center'
                }}>
                  <FileSpreadsheet size={22} />
                </div>
                <div>
                  <div style={{ fontWeight: 600, fontSize: '0.85rem', marginBottom: '0.15rem' }}>Excel Spreadsheet (.xlsx)</div>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-secondary)' }}>Full tabular format with all metadata columns & transcript results.</div>
                </div>
              </button>

              {/* Plain Text Option */}
              <button
                onClick={() => {
                  exportAsText();
                  setShowExportModal(false);
                }}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '1rem',
                  width: '100%',
                  padding: '1rem',
                  background: 'var(--bg-surface, #1e293b)',
                  border: '1px solid var(--border-color)',
                  borderRadius: '8px',
                  color: 'var(--text-primary)',
                  cursor: 'pointer',
                  textAlign: 'left',
                  transition: 'all 0.15s ease'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.borderColor = 'var(--accent, #3b82f6)';
                  e.currentTarget.style.background = 'rgba(59, 130, 246, 0.05)';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.borderColor = 'var(--border-color)';
                  e.currentTarget.style.background = 'var(--bg-surface)';
                }}
              >
                <div style={{
                  background: 'rgba(59, 130, 246, 0.15)',
                  color: 'var(--accent, #3b82f6)',
                  padding: '0.6rem',
                  borderRadius: '6px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center'
                }}>
                  <FileText size={22} />
                </div>
                <div>
                  <div style={{ fontWeight: 600, fontSize: '0.85rem', marginBottom: '0.15rem' }}>Plain Text Report (.txt)</div>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-secondary)' }}>Divided entries containing recording URL, metadata, and full transcripts.</div>
                </div>
              </button>
            </div>
          </div>
        </div>
      )}

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
                  onChange={(e) => setApiKey(e.target.value)} placeholder="AI... or AQ..."
                  style={{ paddingRight: '2.5rem' }} />
                <button type="button" onClick={() => setShowApiKey(!showApiKey)}
                  style={{ position: 'absolute', right: '0.5rem', top: '50%', transform: 'translateY(-50%)', background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', padding: '0.2rem' }}>
                  {showApiKey ? <EyeOff size={15} /> : <Eye size={15} />}
                </button>
              </div>
            </div>

            {mounted && (model === 'claude-sonnet-4-6' || metaModel === 'claude-sonnet-4-6') && (
              <div className="form-group" style={{ marginBottom: 0 }}>
                <label className="label">Anthropic API Key (Optional)</label>
                <div style={{ position: 'relative' }}>
                  <input type={showAnthropicApiKey ? 'text' : 'password'} value={anthropicApiKey}
                    onChange={(e) => setAnthropicApiKey(e.target.value)} placeholder="sk-ant-..."
                    style={{ paddingRight: '2.5rem' }} />
                  <button type="button" onClick={() => setShowAnthropicApiKey(!showAnthropicApiKey)}
                    style={{ position: 'absolute', right: '0.5rem', top: '50%', transform: 'translateY(-50%)', background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', padding: '0.2rem' }}>
                    {showAnthropicApiKey ? <EyeOff size={15} /> : <Eye size={15} />}
                  </button>
                </div>
                <span style={{ fontSize: '0.65rem', color: 'var(--text-secondary)', marginTop: '0.2rem', display: 'block' }}>
                  Leave blank to reuse your Google Cloud / Vertex AI API key.
                </span>
                </div>
            )}

            {mounted && (model === 'claude-sonnet-4-6' || metaModel === 'claude-sonnet-4-6') && !anthropicApiKey && (
              <div className="form-group" style={{ marginBottom: 0 }}>
                <label className="label">Vertex Service Account JSON</label>
                <textarea
                  value={vertexServiceAccount}
                  onChange={(e) => setVertexServiceAccount(e.target.value)}
                  placeholder='{"type": "service_account", ...}'
                  style={{
                    width: '100%',
                    height: '60px',
                    resize: 'vertical',
                    fontSize: '0.72rem',
                    fontFamily: 'monospace',
                    padding: '0.4rem 0.5rem',
                    background: 'var(--bg-surface)',
                    border: '1px solid var(--border-color)',
                    borderRadius: '4px',
                    color: 'var(--text-primary)',
                    outline: 'none'
                  }}
                />
                <span style={{ fontSize: '0.65rem', color: 'var(--text-secondary)', marginTop: '0.2rem', display: 'block' }}>
                  Paste the contents of your GCP Service Account JSON key file.
                </span>
              </div>
            )}

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.75rem' }}>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <label className="label">Model</label>
                <select value={model} onChange={(e) => setModel(e.target.value)}>
                  <optgroup label="Anthropic">
                    <option value="claude-sonnet-4-6">Claude Sonnet 4.6</option>
                  </optgroup>
                  <optgroup label="Latest Preview">
                    <option value="gemini-3.5-flash">3.5 Flash</option>
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
          <div className="card-body" style={{ padding: '0.75rem 1rem', display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>

            {/* Mode tabs */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', borderRadius: '6px', overflow: 'hidden', border: '1px solid var(--border-color)' }}>
              <button
                onClick={() => { setInputMode('spreadsheet'); clearData(); }}
                style={{
                  padding: '0.45rem 0', fontSize: '0.72rem', fontWeight: 600, border: 'none', cursor: 'pointer',
                  background: inputMode === 'spreadsheet' ? 'var(--accent)' : 'var(--bg-surface)',
                  color: inputMode === 'spreadsheet' ? '#fff' : 'var(--text-secondary)',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.3rem', transition: 'all 0.15s',
                }}>
                <FileSpreadsheet size={12} /> Sheet / CSV
              </button>
              <button
                onClick={() => { setInputMode('audio'); clearData(); }}
                style={{
                  padding: '0.45rem 0', fontSize: '0.72rem', fontWeight: 600, border: 'none', cursor: 'pointer',
                  background: inputMode === 'audio' ? 'var(--accent)' : 'var(--bg-surface)',
                  color: inputMode === 'audio' ? '#fff' : 'var(--text-secondary)',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.3rem', transition: 'all 0.15s',
                }}>
                <Music2 size={12} /> Audio Files
              </button>
            </div>

            {inputMode === 'spreadsheet' ? (
              <div className="dropzone" onClick={() => fileInputRef.current?.click()}>
                <input type="file" accept=".xlsx,.xls,.csv,.xlsm,.ods" style={{ display: 'none' }}
                  ref={fileInputRef} onChange={handleFileUpload} />
                <UploadCloud size={26} color="var(--text-secondary)" style={{ marginBottom: '0.6rem' }} />
                <div style={{ fontWeight: 500, fontSize: '0.85rem', marginBottom: '0.15rem' }}>Upload Sheet or CSV</div>
                <div style={{ fontSize: '0.65rem', color: 'var(--text-secondary)', textAlign: 'center', padding: '0 0.5rem' }}>
                  Supports raw recording URLs or previously exported Step 1 results to run Step 2 directly!
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                {/* Hidden inputs */}
                <input
                  type="file"
                  accept="audio/*,.mp3,.wav,.ogg,.m4a,.aac,.flac,.webm,.opus,.wma"
                  multiple
                  style={{ display: 'none' }}
                  ref={audioFilesInputRef}
                  onChange={handleAudioFilesUpload}
                />
                <input
                  type="file"
                  accept="audio/*,.mp3,.wav,.ogg,.m4a,.aac,.flac,.webm,.opus,.wma"
                  multiple
                  // @ts-ignore – webkitdirectory is non-standard but widely supported
                  webkitdirectory=""
                  style={{ display: 'none' }}
                  ref={audioFolderInputRef}
                  onChange={handleAudioFilesUpload}
                />
                <div
                  className="dropzone"
                  onClick={() => audioFilesInputRef.current?.click()}
                  style={{ padding: '0.9rem 0.75rem' }}
                >
                  <Music2 size={22} color="var(--text-secondary)" style={{ marginBottom: '0.4rem' }} />
                  <div style={{ fontWeight: 500, fontSize: '0.82rem', marginBottom: '0.1rem' }}>Select Audio Files</div>
                  <div style={{ fontSize: '0.68rem', color: 'var(--text-secondary)' }}>mp3, wav, ogg, m4a, aac…</div>
                </div>
                <button
                  className="btn btn-secondary"
                  style={{ fontSize: '0.75rem', padding: '0.45rem', width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.4rem' }}
                  onClick={() => audioFolderInputRef.current?.click()}
                >
                  <UploadCloud size={13} /> Upload Folder
                </button>
              </div>
            )}

            {records.length > 0 && (
              <div style={{ padding: '0.55rem 0.75rem', background: 'var(--bg-surface)', border: '1px solid var(--border-color)', borderRadius: '6px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.4rem', fontSize: '0.78rem' }}>
                  <CheckCircle size={13} color="var(--success)" /> {records.length} {inputMode === 'audio' ? 'files' : 'records'}
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
              placeholder="Enter your custom evaluation or transcription instructions..."
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
          <button className="btn btn-secondary" onClick={() => setShowExportModal(true)} disabled={records.length === 0}><Download size={15} /> Export</button>
          <button className="btn btn-secondary" onClick={copyAllResults} disabled={records.filter(r => r.result).length === 0}><Copy size={15} /> Copy All</button>
          <button className="btn btn-secondary" onClick={copyAllMobileNumbers} disabled={records.length === 0 || !records.some(r => r.metadata?.mobile_number)}><Copy size={15} /> Copy Mobiles</button>
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
                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', minWidth: 0, flex: 1 }}>
                      <span style={{ color: 'var(--text-muted)', flexShrink: 0 }}>#{rec.id + 1}</span>
                      <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
                        {inputMode === 'audio' ? (
                          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontSize: '0.8rem', display: 'flex', alignItems: 'center', gap: '0.35rem', fontWeight: 500 }}>
                            <Music2 size={12} style={{ flexShrink: 0, color: 'var(--text-muted)' }} />
                            {rec.label || rec.audioFile?.name}
                          </span>
                        ) : (
                          <a href={rec.url} target="_blank" rel="noreferrer"
                            style={{ textDecoration: 'underline', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontSize: '0.8rem', fontWeight: 500 }}>
                            {rec.url}
                          </a>
                        )}
                        {rec.metadata && (
                          <div style={{ display: 'flex', gap: '0.75rem', flexWrap: 'wrap', marginTop: '0.2rem', fontSize: '0.68rem', color: 'var(--text-muted)' }}>
                            {rec.metadata.name && <span><strong>Name:</strong> {rec.metadata.name}</span>}
                            {rec.metadata.mobile_number !== undefined && (
                              <span style={{ display: 'inline-flex', alignItems: 'center', gap: '0.2rem' }}>
                                <strong>Mobile:</strong> {rec.metadata.mobile_number}
                                <button 
                                  onClick={(e) => {
                                    e.preventDefault();
                                    e.stopPropagation();
                                    navigator.clipboard.writeText(String(rec.metadata?.mobile_number || '')).then(() => alert('Mobile number copied!'));
                                  }}
                                  style={{
                                    background: 'transparent',
                                    border: 'none',
                                    color: 'var(--text-muted)',
                                    cursor: 'pointer',
                                    padding: '0.1rem',
                                    display: 'inline-flex',
                                    alignItems: 'center',
                                    outline: 'none'
                                  }}
                                  title="Copy Mobile Number"
                                >
                                  <Copy size={10} />
                                </button>
                              </span>
                            )}
                            {rec.metadata.date && <span><strong>Date:</strong> {rec.metadata.date}</span>}
                            {rec.metadata.duration && <span><strong>Duration:</strong> {rec.metadata.duration}</span>}
                          </div>
                        )}
                      </div>
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

                  {/* Step 2 Output / Extraction Result */}
                  {(rec.step2Status || rec.step2Result) && (
                    <div className="card-body" style={{ position: 'relative', padding: '1.25rem', borderTop: '1px solid var(--border-color)', backgroundColor: 'rgba(255,255,255,0.015)' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.75rem' }}>
                        <h4 style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--accent, #3b82f6)', textTransform: 'uppercase', letterSpacing: '0.05em', margin: 0 }}>
                          Step 2 Output / Extraction
                        </h4>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                          {rec.step2Usage && (
                            <span style={{ fontSize: '0.65rem', color: 'var(--text-muted)' }}>
                              {rec.step2Usage.promptTokenCount}↑ {rec.step2Usage.candidatesTokenCount}↓
                            </span>
                          )}
                          <span className={`status-badge ${rec.step2Status}`} style={{ fontSize: '0.65rem', padding: '0.1rem 0.4rem' }}>
                            {rec.step2Status === 'processing' && <Loader2 size={9} className="spinner" style={{ marginRight: '0.2rem' }} />}
                            {rec.step2Status}
                          </span>
                        </div>
                      </div>
                      {rec.step2Result && (
                        <>
                          <button onClick={() => navigator.clipboard.writeText(String(rec.step2Result)).then(() => alert('Copied!'))}
                            className="btn btn-secondary" style={{ position: 'absolute', top: '1rem', right: '1rem', padding: '0.3rem' }} title="Copy">
                            <Copy size={13} />
                          </button>
                          <div className="markdown-body" style={{ marginTop: '0.5rem' }}>
                            <ReactMarkdown remarkPlugins={[remarkGfm]} components={codeComponents}>{String(rec.step2Result)}</ReactMarkdown>
                          </div>
                        </>
                      )}
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
              <p className="subtitle" style={{ marginBottom: '1rem' }}>Aggregate successful outputs or run batch post-processing/extractions row-by-row.</p>
              
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.75rem', marginBottom: '0.75rem' }}>
                <div className="form-group" style={{ marginBottom: 0 }}>
                  <label className="label">Meta-Analysis Model</label>
                  <select value={metaModel} onChange={(e) => setMetaModel(e.target.value)} style={{ width: '100%', padding: '0.4rem 0.5rem', background: 'var(--bg-surface)', border: '1px solid var(--border-color)', borderRadius: '4px', color: 'var(--text-primary)', outline: 'none' }}>
                    <optgroup label="Gemini (Recommended)">
                      <option value="gemini-2.5-pro">Gemini 2.5 Pro</option>
                      <option value="gemini-2.5-flash">Gemini 2.5 Flash</option>
                      <option value="gemini-3.5-flash">Gemini 3.5 Flash</option>
                      <option value="gemini-3.1-pro-preview">Gemini 3.1 Pro Preview</option>
                      <option value="gemini-3-flash-preview">Gemini 3 Flash Preview</option>
                    </optgroup>
                    <optgroup label="Anthropic (requires sk-ant- key)">
                      <option value="claude-sonnet-4-6">Claude Sonnet 4.6</option>
                    </optgroup>
                  </select>
                </div>
                <div className="form-group" style={{ marginBottom: 0 }}>
                  <label className="label">Analysis Mode</label>
                  <select value={metaAnalysisMode} onChange={(e) => setMetaAnalysisMode(e.target.value as 'aggregate' | 'row')} style={{ width: '100%', padding: '0.4rem 0.5rem', background: 'var(--bg-surface)', border: '1px solid var(--border-color)', borderRadius: '4px', color: 'var(--text-primary)', outline: 'none' }}>
                    <option value="aggregate">Aggregate (All at once)</option>
                    <option value="row">Row-by-Row Extraction</option>
                  </select>
                </div>
              </div>

              <div className="form-group">
                <label className="label">Meta Prompt</label>
                <textarea value={metaPrompt} onChange={(e) => setMetaPrompt(e.target.value)}
                  style={{ width: '100%', minHeight: '70px', resize: 'vertical' }} />
              </div>
              <button className="btn btn-primary" style={{ width: '100%', padding: '0.7rem' }}
                onClick={startMetaAnalysis} disabled={isMetaProcessing}>
                {isMetaProcessing ? (
                  <>
                    <Loader2 className="spinner" /> 
                    {metaAnalysisMode === 'row' ? `Processing Row-by-Row (${metaProgress}%)...` : 'Synthesizing...'}
                  </>
                ) : (
                  <>
                    <Play size={16} /> 
                    {metaAnalysisMode === 'row' ? 'Run Row-by-Row Extraction' : 'Synthesize Macro Trends'}
                  </>
                )}
              </button>

              {isMetaProcessing && metaAnalysisMode === 'row' && (
                <div style={{ marginTop: '0.75rem' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.65rem', fontWeight: 500, marginBottom: '0.25rem', color: 'var(--text-secondary)' }}>
                    <span>EXTRACTION PROGRESS</span><span>{metaProgress}%</span>
                  </div>
                  <div className="progress-bar"><div className="progress-fill" style={{ width: `${metaProgress}%` }} /></div>
                </div>
              )}

              {metaAnalysisMode === 'row' && !isMetaProcessing && records.some(r => r.step2Result) && (
                <div style={{ display: 'flex', gap: '0.75rem', marginTop: '0.75rem' }}>
                  <button className="btn btn-secondary" style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.4rem' }} onClick={() => setShowExportModal(true)}>
                    <Download size={14} /> Export Consolidated Spreadsheet
                  </button>
                </div>
              )}

              {metaAnalysisMode === 'aggregate' && metaResult && (
                <div style={{ marginTop: '1.5rem', paddingTop: '1.25rem', borderTop: '1px solid var(--border-color)', position: 'relative' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.75rem' }}>
                    <h3 className="title" style={{ fontSize: '0.95rem' }}>Synthesis Report</h3>
                    {metaUsage && <div className="status-badge" style={{ color: 'var(--text-primary)' }}>{metaUsage.promptTokenCount}↑ {metaUsage.candidatesTokenCount}↓</div>}
                  </div>
                  <div style={{ position: 'absolute', top: '1.25rem', right: 0, display: 'flex', gap: '0.5rem' }}>
                    <button onClick={() => navigator.clipboard.writeText(metaResult).then(() => alert('Copied!'))}
                      className="btn btn-secondary" style={{ padding: '0.3rem' }} title="Copy">
                      <Copy size={13} />
                    </button>
                    <button onClick={downloadMetaReport}
                      className="btn btn-secondary" style={{ padding: '0.3rem' }} title="Download Report">
                      <Download size={13} />
                    </button>
                  </div>
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
            {inputMode === 'audio' ? <Music2 size={40} strokeWidth={1} /> : <FileSpreadsheet size={40} strokeWidth={1} />}
            <p style={{ fontSize: '0.875rem' }}>
              {inputMode === 'audio' ? 'Upload audio files or a folder to get started' : 'Upload a spreadsheet to get started'}
            </p>
          </div>
        )}

      </main>
    </div>
  );
}
