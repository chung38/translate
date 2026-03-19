/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import mammoth from 'mammoth';
import { Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType } from 'docx';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import { jsPDF } from 'jspdf';
import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import { 
  Upload, 
  FileText, 
  Languages, 
  Download, 
  Loader2, 
  CheckCircle2, 
  AlertCircle,
  ArrowRight,
  FileSpreadsheet,
  File as FileIcon,
  Presentation
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// Set up PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

// DeepSeek API Configuration (Now handled via server proxy)
const DEEPSEEK_PROXY_URL = '/api/translate';

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

const AVAILABLE_LANGUAGES = [
  { id: 'th', name: '泰文', label: 'Thai' },
  { id: 'id', name: '印尼文', label: 'Indonesian' },
  { id: 'vi', name: '越南文', label: 'Vietnamese' },
  { id: 'en', name: '英文', label: 'English' },
];

type TranslationStatus = 'idle' | 'processing' | 'translating' | 'generating' | 'completed' | 'error';

export default function App() {
  const [files, setFiles] = useState<File[]>([]);
  const [selectedLanguages, setSelectedLanguages] = useState<string[]>(['英文']);
  const [industry, setIndustry] = useState('');
  const [status, setStatus] = useState<TranslationStatus>('idle');
  const [statusMessage, setStatusMessage] = useState('');
  const [progress, setProgress] = useState(0);
  const [fileProgress, setFileProgress] = useState<Record<string, number>>({});
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const isCancelledRef = useRef(false);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(e.target.files || []) as File[];
    if (selectedFiles.length > 0) {
      const validExtensions = ['docx', 'xlsx', 'pdf', 'pptx'];
      const newFiles: File[] = [];
      let hasInvalid = false;

      selectedFiles.forEach(f => {
        const extension = f.name.split('.').pop()?.toLowerCase();
        if (validExtensions.includes(extension || '')) {
          newFiles.push(f);
        } else {
          hasInvalid = true;
        }
      });

      if (newFiles.length > 0) {
        setFiles(prev => [...prev, ...newFiles]);
        setError(hasInvalid ? '部分檔案格式不支援，已跳過' : null);
        setStatus('idle');
      } else if (hasInvalid) {
        setError('請上傳有效的 .docx, .xlsx, .pdf 或 .pptx 檔案');
      }
    }
  };

  const removeFile = (index: number) => {
    setFiles(prev => prev.filter((_, i) => i !== index));
    if (files.length === 1) {
      setStatus('idle');
      setProgress(0);
    }
  };

  const toggleLanguage = (langName: string) => {
    setSelectedLanguages(prev => 
      prev.includes(langName) 
        ? prev.filter(l => l !== langName)
        : [...prev, langName]
    );
  };

  const translateBatch = async (texts: string[], targetLangs: string[], retryCount = 0): Promise<Record<string, string>[]> => {
    if (texts.length === 0 || targetLangs.length === 0) return texts.map(() => ({}));
    
    try {
      // Add a delay between batches to stay under RPM limits
      await sleep(1000); 

      const industryContext = industry ? `。這是一個關於「${industry}」行業的文件，請使用該行業的專業術語進行翻譯` : '';
      const prompt = `你是一個專業的翻譯官${industryContext}。請將以下文字陣列中的每一項同時翻譯成以下語言：${targetLangs.join('、')}。
      請嚴格以 JSON 格式回傳結果，格式如下：
      {
        "translations": [
          { "${targetLangs[0]}": "翻譯內容1", "${targetLangs[1]}": "翻譯內容1", ... },
          { "${targetLangs[0]}": "翻譯內容2", "${targetLangs[1]}": "翻譯內容2", ... },
          ...
        ]
      }
      確保回傳的 "translations" 陣列長度與輸入的文字陣列長度完全一致 (${texts.length})。
      注意：翻譯後的文字內容中，絕對不要包含任何語言標籤（例如不要出現 [英文] 或 [English] 等字樣），只需要純粹的翻譯內容。
      不要包含任何 Markdown 標籤（如 \`\`\`json）或額外文字，只回傳純 JSON 字串。
      
      待翻譯內容陣列：
      ${JSON.stringify(texts, null, 2)}`;

      const response = await fetch(DEEPSEEK_PROXY_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ prompt })
      });

      if (!response.ok) {
        let errorMessage = `HTTP error! status: ${response.status}`;
        try {
          const errorData = await response.json();
          errorMessage = errorData.error?.message || errorMessage;
        } catch (e) {
          const textError = await response.text();
          errorMessage = textError || errorMessage;
        }
        throw new Error(errorMessage);
      }

      const data = await response.json();
      const resultText = data.choices[0].message.content.trim();
      const parsed = JSON.parse(resultText);
      
      if (!parsed.translations || !Array.isArray(parsed.translations)) {
        throw new Error('API 回傳格式不正確');
      }

      return parsed.translations;
    } catch (err: any) {
      const isRateLimit = err?.message?.includes('429') || JSON.stringify(err).includes('429');
      const isAuthError = err?.message?.includes('Authentication') || err?.message?.includes('API key') || err?.message?.includes('401') || err?.message?.includes('配置');
      
      if (isRateLimit && retryCount < 5) {
        const waitTime = Math.pow(2, retryCount) * 5000 + Math.random() * 2000;
        console.warn(`DeepSeek Rate limit hit. Waiting ${Math.round(waitTime/1000)}s... (Attempt ${retryCount + 1})`);
        await sleep(waitTime);
        return translateBatch(texts, targetLangs, retryCount + 1);
      }

      if (isAuthError) {
        throw new Error(err.message || 'DeepSeek API Key 驗證失敗，請檢查設定是否正確。');
      }

      console.error(`DeepSeek Batch translation error:`, err);
      // Fallback: return error placeholders for each text
      return texts.map(() => {
        const errorResult: Record<string, string> = {};
        targetLangs.forEach(lang => errorResult[lang] = `(翻譯出錯: ${err?.message || 'API 錯誤'})`);
        return errorResult;
      });
    }
  };

  const cancelTranslation = () => {
    isCancelledRef.current = true;
    setStatus('idle');
    setProgress(0);
    setError('翻譯已取消');
  };

  const processDocx = async (file: File, updateProgress: (p: number, status?: TranslationStatus) => void) => {
    updateProgress(5, 'processing');
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    
    // 1. Read the main document XML
    const documentXmlPath = 'word/document.xml';
    const documentXmlContent = await zip.file(documentXmlPath)?.async('string');
    if (!documentXmlContent) throw new Error('無法讀取 Word 文件內容');

    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(documentXmlContent, 'application/xml');
    const ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // 2. Find all paragraphs
    const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(ns, 'p'));
    const totalParagraphs = paragraphs.length;
    
    updateProgress(10, 'translating');
    const BATCH_SIZE = 10;
    
    for (let i = 0; i < totalParagraphs; i += BATCH_SIZE) {
      if (isCancelledRef.current) throw new Error('翻譯已取消');
      
      const batchParas = paragraphs.slice(i, i + BATCH_SIZE);
      const paraData = batchParas.map(p => {
        // Get all text nodes in this paragraph
        const textNodes = Array.from(p.getElementsByTagNameNS(ns, 't'));
        const fullText = textNodes.map(t => t.textContent || '').join('');
        return { p, fullText };
      }).filter(data => data.fullText.trim().length > 0);

      if (paraData.length > 0) {
        // Translating progress: 10% to 90%
        const currentProgress = 10 + Math.round((Math.min(i + BATCH_SIZE, totalParagraphs) / totalParagraphs) * 80);
        updateProgress(currentProgress, 'translating');
        
        const textsToTranslate = paraData.map(d => d.fullText);
        const batchTranslations = await translateBatch(textsToTranslate, selectedLanguages);
        
        paraData.forEach((data, idx) => {
          const translations = batchTranslations[idx];
          const p = data.p;
          
          // Find the first run to copy its properties
          const firstRun = p.getElementsByTagNameNS(ns, 'r')[0];
          const firstRunPr = firstRun?.getElementsByTagNameNS(ns, 'rPr')[0];

          selectedLanguages.forEach(lang => {
            const translatedText = translations[lang];
            if (!translatedText) return;

            // Create a new run for the translation
            const newRun = xmlDoc.createElementNS(ns, 'w:r');
            
            // Copy properties if they exist
            if (firstRunPr) {
              newRun.appendChild(firstRunPr.cloneNode(true));
            }

            // Add a break before the translation
            const br = xmlDoc.createElementNS(ns, 'w:br');
            newRun.appendChild(br);

            // Add the translated text
            const t = xmlDoc.createElementNS(ns, 'w:t');
            t.textContent = translatedText;
            newRun.appendChild(t);

            // Append the translation to the paragraph
            p.appendChild(newRun);
          });
        });
      }
    }

    // 3. Save the modified XML back to the zip
    updateProgress(95, 'generating');
    const serializer = new XMLSerializer();
    const modifiedXml = serializer.serializeToString(xmlDoc);
    zip.file(documentXmlPath, modifiedXml);

    // 4. Generate the new docx file
    const content = await zip.generateAsync({ type: 'blob' });
    const fileName = file.name.replace(/\.docx$/, `_translated.docx`);
    saveAs(content, fileName);
    updateProgress(100, 'generating');
  };

  const processExcel = async (file: File, updateProgress: (p: number, status?: TranslationStatus) => void) => {
    updateProgress(5, 'processing');
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer);
    const newWorkbook = XLSX.utils.book_new();

    updateProgress(10, 'translating');
    const sheetNames = workbook.SheetNames;
    
    for (let s = 0; s < sheetNames.length; s++) {
      if (isCancelledRef.current) throw new Error('翻譯已取消');
      
      const sheetName = sheetNames[s];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
      
      const newSheetData: any[][] = [];
      const totalRows = jsonData.length;

      // Batch translate rows
      const BATCH_SIZE = 10;
      for (let r = 0; r < totalRows; r += BATCH_SIZE) {
        if (isCancelledRef.current) throw new Error('翻譯已取消');
        
        // Translating progress: 10% to 90%
        const currentProgress = 10 + Math.round(((s * totalRows + r) / (sheetNames.length * totalRows)) * 80);
        updateProgress(currentProgress, 'translating');
        
        const batchRows = jsonData.slice(r, r + BATCH_SIZE);
        const cellsToTranslate: { r: number, c: number, text: string }[] = [];
        
        batchRows.forEach((row, rowIdx) => {
          row.forEach((cellValue, colIdx) => {
            if (cellValue && typeof cellValue === 'string' && cellValue.trim()) {
              cellsToTranslate.push({ r: rowIdx, c: colIdx, text: cellValue });
            }
          });
        });

        if (cellsToTranslate.length > 0) {
          const batchTranslations = await translateBatch(cellsToTranslate.map(c => c.text), selectedLanguages);
          
          cellsToTranslate.forEach((cell, idx) => {
            const translations = batchTranslations[idx];
            let combinedValue = cell.text;
            for (const lang of selectedLanguages) {
              combinedValue += `\n${translations[lang] || '(翻譯失敗)'}`;
            }
            batchRows[cell.r][cell.c] = combinedValue;
          });
        }
        
        newSheetData.push(...batchRows);
      }
      
      const newWorksheet = XLSX.utils.aoa_to_sheet(newSheetData);
      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
    }

    updateProgress(95, 'generating');
    const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
    const excelBlob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(excelBlob, `translated_${file.name}`);
    updateProgress(100, 'generating');
  };

  const processPdf = async (file: File, updateProgress: (p: number, status?: TranslationStatus) => void) => {
    updateProgress(5, 'processing');
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const doc = new jsPDF();
    
    updateProgress(10, 'translating');
    const totalPages = pdf.numPages;

    for (let i = 1; i <= totalPages; i++) {
      if (isCancelledRef.current) throw new Error('翻譯已取消');
      
      // Translating progress: 10% to 90%
      const currentProgress = 10 + Math.round((i / totalPages) * 80);
      updateProgress(currentProgress, 'translating');
      
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      const strings = textContent.items.map((item: any) => item.str);
      const fullText = strings.join(' ').trim();

      if (fullText) {
        // Split text into manageable chunks for translation if it's too long
        const chunks = fullText.match(/.{1,1000}/g) || [fullText];
        const batchTranslations = await translateBatch(chunks, selectedLanguages);
        
        if (i > 1) doc.addPage();
        
        let yPos = 20;
        const margin = 15;
        const pageWidth = doc.internal.pageSize.getWidth();
        const contentWidth = pageWidth - (margin * 2);

        // Header
        doc.setFont("helvetica", "bold");
        doc.setFontSize(14);
        doc.setTextColor(16, 185, 129); // Emerald 600
        doc.text(`Page ${i} Translation`, margin, yPos);
        yPos += 10;

        // Original Text
        doc.setFont("helvetica", "bold");
        doc.setFontSize(10);
        doc.setTextColor(100, 116, 139); // Slate 500
        doc.text("ORIGINAL TEXT", margin, yPos);
        yPos += 6;
        
        doc.setFont("helvetica", "normal");
        doc.setFontSize(9);
        doc.setTextColor(30, 41, 59); // Slate 800
        const splitOriginal = doc.splitTextToSize(fullText, contentWidth);
        
        // Handle page overflow for original text
        for (const line of splitOriginal) {
          if (yPos > 280) {
            doc.addPage();
            yPos = 20;
          }
          doc.text(line, margin, yPos);
          yPos += 5;
        }
        
        yPos += 10;

        // Translations
        for (const lang of selectedLanguages) {
          if (yPos > 260) {
            doc.addPage();
            yPos = 20;
          }
          
          doc.setFont("helvetica", "bold");
          doc.setFontSize(10);
          doc.setTextColor(16, 185, 129); // Emerald 600
          doc.text(`${lang.toUpperCase()}`, margin, yPos);
          yPos += 6;

          doc.setFont("helvetica", "normal");
          doc.setFontSize(10);
          doc.setTextColor(0, 0, 0);
          
          const combinedTrans = batchTranslations.map(t => t[lang] || '').join(' ');
          const splitTrans = doc.splitTextToSize(combinedTrans, contentWidth);
          
          for (const line of splitTrans) {
            if (yPos > 280) {
              doc.addPage();
              yPos = 20;
            }
            doc.text(line, margin, yPos);
            yPos += 6;
          }
          yPos += 8;
        }
      }
    }

    updateProgress(95, 'generating');
    doc.save(`translated_${file.name}`);
    updateProgress(100, 'generating');
  };

  const processPptx = async (file: File, updateProgress: (p: number, status?: TranslationStatus) => void) => {
    updateProgress(5, 'processing');
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    
    // 1. Find all slide files
    const slideFiles = Object.keys(zip.files).filter(path => path.startsWith('ppt/slides/slide') && path.endsWith('.xml'));
    const totalSlides = slideFiles.length;
    
    updateProgress(10, 'translating');
    const nsA = "http://schemas.openxmlformats.org/drawingml/2006/main";

    for (let i = 0; i < totalSlides; i++) {
      if (isCancelledRef.current) throw new Error('翻譯已取消');
      
      // Translating progress: 10% to 90%
      const currentProgress = 10 + Math.round((i / totalSlides) * 80);
      updateProgress(currentProgress, 'translating');
      
      const slidePath = slideFiles[i];
      const slideXmlContent = await zip.file(slidePath)?.async('string');
      if (!slideXmlContent) continue;

      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(slideXmlContent, 'application/xml');
      
      // Find all paragraphs in the slide
      const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(nsA, 'p'));
      
      const paraData = paragraphs.map(p => {
        const textNodes = Array.from(p.getElementsByTagNameNS(nsA, 't'));
        const fullText = textNodes.map(t => t.textContent || '').join('');
        return { p, fullText };
      }).filter(data => data.fullText.trim().length > 0);

      if (paraData.length > 0) {
        const textsToTranslate = paraData.map(d => d.fullText);
        const batchTranslations = await translateBatch(textsToTranslate, selectedLanguages);
        
        paraData.forEach((data, idx) => {
          const translations = batchTranslations[idx];
          const p = data.p;
          
          // Find the first run to copy its properties
          const firstRun = p.getElementsByTagNameNS(nsA, 'r')[0];
          const firstRunPr = firstRun?.getElementsByTagNameNS(nsA, 'rPr')[0];

          selectedLanguages.forEach(lang => {
            const translatedText = translations[lang];
            if (!translatedText) return;

            // Add a break (a:br)
            const br = xmlDoc.createElementNS(nsA, 'a:br');
            if (firstRunPr) {
              br.appendChild(firstRunPr.cloneNode(true));
            }
            p.appendChild(br);

            // Create a new run for the translation
            const newRun = xmlDoc.createElementNS(nsA, 'a:r');
            
            // Copy properties if they exist
            if (firstRunPr) {
              newRun.appendChild(firstRunPr.cloneNode(true));
            }

            // Add the translated text in a new run
            const t = xmlDoc.createElementNS(nsA, 'a:t');
            t.textContent = translatedText;
            newRun.appendChild(t);

            // Append the translation to the paragraph
            p.appendChild(newRun);
          });
        });

        // Save the modified slide XML back to the zip
        const serializer = new XMLSerializer();
        const modifiedXml = serializer.serializeToString(xmlDoc);
        zip.file(slidePath, modifiedXml);
      }
    }

    updateProgress(95, 'generating');
    const content = await zip.generateAsync({ type: 'blob' });
    const fileName = file.name.replace(/\.pptx$/, `_translated.pptx`);
    saveAs(content, fileName);
    updateProgress(100, 'generating');
  };

  const processFiles = async () => {
    if (files.length === 0 || selectedLanguages.length === 0) return;

    try {
      setStatus('processing');
      setError(null);
      setProgress(0);
      setFileProgress({});
      isCancelledRef.current = false;

      for (let i = 0; i < files.length; i++) {
        if (isCancelledRef.current) break;
        
        const currentFile = files[i];
        const extension = currentFile.name.split('.').pop()?.toLowerCase();
        
        setStatusMessage(`正在處理第 ${i + 1} / ${files.length} 份文件: ${currentFile.name}`);
        
        // Update overall progress based on file index
        const baseProgress = (i / files.length) * 100;
        const fileWeight = 100 / files.length;
        
        const updateFileProgress = (p: number, currentStatus?: TranslationStatus) => {
          setFileProgress(prev => ({ ...prev, [currentFile.name]: p }));
          setProgress(Math.round(baseProgress + (p * fileWeight / 100)));
          if (currentStatus) setStatus(currentStatus);
        };

        switch (extension) {
          case 'docx':
            await processDocx(currentFile, updateFileProgress);
            break;
          case 'xlsx':
            await processExcel(currentFile, updateFileProgress);
            break;
          case 'pdf':
            await processPdf(currentFile, updateFileProgress);
            break;
          case 'pptx':
            await processPptx(currentFile, updateFileProgress);
            break;
          default:
            console.warn(`不支援的檔案格式: ${currentFile.name}`);
        }
      }
      
      if (!isCancelledRef.current) {
        setStatus('completed');
        setStatusMessage('所有翻譯已完成！');
        setProgress(100);
      }
    } catch (err) {
      console.error(err);
      setError(err instanceof Error ? err.message : '處理檔案時發生錯誤');
      setStatus('error');
      setStatusMessage('發生錯誤');
    }
  };

  return (
    <div className="min-h-screen bg-[#F5F5F5] font-sans text-[#1A1A1A] p-3 sm:p-4 md:p-8">
      <div className="max-w-3xl mx-auto">
        {/* Header */}
        <header className="mb-8 md:mb-12 text-center">
          <motion.div 
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            className="inline-flex items-center justify-center w-12 h-12 md:w-16 md:h-16 bg-white rounded-2xl shadow-sm mb-4 md:mb-6"
          >
            <Languages className="w-6 h-6 md:w-8 md:h-8 text-emerald-600" />
          </motion.div>
          <motion.h1 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.1 }}
            className="text-2xl md:text-4xl font-light tracking-tight mb-2 md:mb-3"
          >
            全能文件多語翻譯器
          </motion.h1>
          <motion.p 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.2 }}
            className="text-muted text-[10px] md:text-sm uppercase tracking-widest opacity-60"
          >
            Multi-Language Document Translator
          </motion.p>
        </header>

        {/* Main Card */}
        <motion.div 
          initial={{ opacity: 0, scale: 0.98 }}
          animate={{ opacity: 1, scale: 1 }}
          transition={{ delay: 0.3 }}
          className="bg-white rounded-2xl md:rounded-[32px] shadow-sm border border-black/5 overflow-hidden"
        >
          <div className="p-5 md:p-12">
            {/* Upload Section */}
            <div 
              onClick={() => fileInputRef.current?.click()}
              className={`
                relative group cursor-pointer border-2 border-dashed rounded-2xl p-6 md:p-8 transition-all duration-300
                ${files.length > 0 ? 'border-emerald-200 bg-emerald-50/10' : 'border-gray-200 hover:border-emerald-400 hover:bg-gray-50'}
              `}
            >
              <input 
                type="file" 
                ref={fileInputRef}
                onChange={handleFileChange}
                accept=".docx,.xlsx,.pdf,.pptx"
                multiple
                className="hidden"
              />
              
              <div className="flex flex-col items-center text-center">
                <div className={`
                  w-12 h-12 rounded-full flex items-center justify-center mb-3 transition-transform duration-300 group-hover:scale-110
                  ${files.length > 0 ? 'bg-emerald-100 text-emerald-600' : 'bg-gray-100 text-gray-400'}
                `}>
                  <Upload className="w-6 h-6" />
                </div>
                
                <div>
                  <p className="text-base font-medium mb-1">點擊或拖拽上傳文件</p>
                  <p className="text-xs text-gray-400">支援 .docx, .xlsx, .pdf, .pptx 格式 (可多選)</p>
                </div>
              </div>
            </div>

            {/* File List */}
            <AnimatePresence>
              {files.length > 0 && (
                <motion.div 
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="mt-6 space-y-2"
                >
                  <div className="flex justify-between items-center px-1 mb-2">
                    <span className="text-[11px] uppercase tracking-wider font-semibold text-gray-400">待處理檔案 ({files.length})</span>
                    {status === 'idle' && (
                      <button 
                        onClick={() => setFiles([])}
                        className="text-[10px] text-red-500 hover:underline"
                      >
                        全部清除
                      </button>
                    )}
                  </div>
                  {files.map((f, idx) => (
                    <motion.div 
                      key={`${f.name}-${idx}`}
                      initial={{ opacity: 0, x: -10 }}
                      animate={{ opacity: 1, x: 0 }}
                      exit={{ opacity: 0, x: 10 }}
                      className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100"
                    >
                      <div className="flex items-center gap-2 md:gap-3 min-w-0">
                        <div className="text-emerald-600 shrink-0">
                          {f.name.endsWith('.docx') && <FileText className="w-3 h-3 md:w-4 md:h-4" />}
                          {f.name.endsWith('.xlsx') && <FileSpreadsheet className="w-3 h-3 md:w-4 md:h-4" />}
                          {f.name.endsWith('.pdf') && <FileIcon className="w-3 h-3 md:w-4 md:h-4" />}
                          {f.name.endsWith('.pptx') && <Presentation className="w-3 h-3 md:w-4 md:h-4" />}
                        </div>
                        <div className="min-w-0">
                          <p className="text-xs md:text-sm font-medium text-gray-700 truncate">{f.name}</p>
                          <p className="text-[9px] md:text-[10px] text-gray-400">{(f.size / 1024).toFixed(1)} KB</p>
                        </div>
                      </div>
                      
                      <div className="flex items-center gap-2 md:gap-3">
                        {status !== 'idle' && status !== 'error' && (
                          <div className="flex flex-col items-end gap-1">
                            <span className="text-[9px] md:text-[10px] font-mono text-emerald-600">
                              {fileProgress[f.name] || 0}%
                            </span>
                            <div className="w-12 md:w-16 h-1 bg-gray-200 rounded-full overflow-hidden">
                              <div 
                                className="h-full bg-emerald-500 transition-all duration-300" 
                                style={{ width: `${fileProgress[f.name] || 0}%` }}
                              />
                            </div>
                          </div>
                        )}
                        {status === 'idle' && (
                          <button 
                            onClick={() => removeFile(idx)}
                            className="p-1 hover:bg-red-50 rounded-lg text-gray-400 hover:text-red-500 transition-colors"
                          >
                            <AlertCircle className="w-4 h-4 rotate-45" />
                          </button>
                        )}
                        {status === 'completed' && fileProgress[f.name] === 100 && (
                          <CheckCircle2 className="w-4 h-4 text-emerald-500" />
                        )}
                      </div>
                    </motion.div>
                  ))}
                </motion.div>
              )}
            </AnimatePresence>

            {/* Settings */}
            <div className="mt-8 space-y-6">
              <div>
                <label className="block text-[11px] uppercase tracking-wider font-semibold text-gray-400 mb-2 ml-1">
                  工廠行業 (可選，使翻譯更精準)
                </label>
                <input
                  type="text"
                  value={industry}
                  onChange={(e) => setIndustry(e.target.value)}
                  placeholder="例如：電子、紡織、汽車..."
                  className="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-emerald-400 focus:ring-1 focus:ring-emerald-400 outline-none transition-all text-sm"
                />
              </div>

              <div>
                <label className="block text-[11px] uppercase tracking-wider font-semibold text-gray-400 mb-4 ml-1">
                  選擇目標語言 (可多選)
                </label>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                  {AVAILABLE_LANGUAGES.map((lang) => (
                    <button
                      key={lang.id}
                      onClick={() => toggleLanguage(lang.name)}
                      className={`
                        flex items-center justify-between px-4 py-3 rounded-xl border transition-all text-sm
                        ${selectedLanguages.includes(lang.name)
                          ? 'bg-emerald-50 border-emerald-200 text-emerald-700 font-medium'
                          : 'bg-white border-gray-200 text-gray-600 hover:border-emerald-200'}
                      `}
                    >
                      <span>{lang.name}</span>
                      {selectedLanguages.includes(lang.name) && (
                        <CheckCircle2 className="w-4 h-4 text-emerald-500" />
                      )}
                    </button>
                  ))}
                </div>
              </div>
              
              <div className="flex flex-col gap-3">
                <button
                  disabled={files.length === 0 || selectedLanguages.length === 0 || status === 'processing' || status === 'translating' || status === 'generating'}
                  onClick={processFiles}
                  className={`
                    w-full h-[48px] md:h-[56px] rounded-xl font-medium flex items-center justify-center gap-2 transition-all px-4
                    ${files.length === 0 || selectedLanguages.length === 0 || status === 'processing' || status === 'translating' || status === 'generating'
                      ? 'bg-gray-100 text-gray-400 cursor-not-allowed'
                      : 'bg-emerald-600 text-white hover:bg-emerald-700 shadow-lg shadow-emerald-600/20 active:scale-[0.98]'}
                  `}
                >
                  {status === 'idle' && (
                    <>
                      <span className="truncate">
                        開始翻譯 {files.length > 0 && `${files.length} 份`} {selectedLanguages.length > 0 && `(${selectedLanguages.length} 種語言)`}
                      </span>
                      <ArrowRight className="w-4 h-4 shrink-0" />
                    </>
                  )}
                  {(status === 'processing' || status === 'translating' || status === 'generating') && (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin shrink-0" />
                      <span className="truncate">正在處理 {Object.keys(fileProgress).length} / {files.length}...</span>
                    </>
                  )}
                  {status === 'completed' && (
                    <>
                      <CheckCircle2 className="w-4 h-4" />
                      <span>全部翻譯完成</span>
                    </>
                  )}
                  {status === 'error' && (
                    <>
                      <AlertCircle className="w-4 h-4" />
                      <span>重試</span>
                    </>
                  )}
                </button>
                
                {(status === 'processing' || status === 'translating' || status === 'generating') && (
                  <button
                    onClick={cancelTranslation}
                    className="w-full h-[40px] rounded-xl font-medium text-red-500 hover:bg-red-50 transition-all text-sm border border-red-100"
                  >
                    取消翻譯
                  </button>
                )}
              </div>
            </div>

            {/* Progress & Status */}
            <AnimatePresence>
              {(status !== 'idle' && status !== 'error') && (
                <motion.div 
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  exit={{ opacity: 0, height: 0 }}
                  className="mt-8 pt-8 border-top border-gray-100"
                >
                  <div className="flex justify-between items-center mb-2">
                    <span className="text-sm font-medium text-gray-600">
                      {statusMessage}
                    </span>
                    {(status === 'processing' || status === 'translating' || status === 'generating') && (
                      <span className="text-xs font-mono text-emerald-600">{progress}%</span>
                    )}
                  </div>
                  <div className="w-full h-1.5 bg-gray-100 rounded-full overflow-hidden">
                    <motion.div 
                      initial={{ width: 0 }}
                      animate={{ width: `${progress}%` }}
                      className="h-full bg-emerald-500"
                    />
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {/* Error Message */}
            {error && (
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="mt-6 p-4 bg-red-50 rounded-xl flex items-start gap-3 text-red-600 text-sm"
              >
                <AlertCircle className="w-5 h-5 shrink-0" />
                <p>{error}</p>
              </motion.div>
            )}
          </div>

            {/* Footer Info */}
          <div className="bg-gray-50 border-t border-gray-100 p-5 md:p-6 flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="flex items-center gap-4">
              <div className="flex -space-x-2">
                {[1, 2, 3].map(i => (
                  <div key={i} className="w-6 h-6 md:w-8 md:h-8 rounded-full border-2 border-white bg-gray-200 overflow-hidden">
                    <img 
                      src={`https://picsum.photos/seed/user${i}/32/32`} 
                      alt="User" 
                      referrerPolicy="no-referrer"
                    />
                  </div>
                ))}
              </div>
              <span className="text-[10px] md:text-xs text-gray-400 font-medium">已有超過 1,000+ 份文件被翻譯</span>
            </div>
            <div className="flex items-center gap-2 text-[10px] md:text-xs text-gray-400">
              <CheckCircle2 className="w-3 h-3 text-emerald-500" />
              <span>由 DeepSeek AI 強力驅動</span>
            </div>
          </div>
        </motion.div>

        {/* Instructions */}
        <div className="mt-12 grid grid-cols-1 md:grid-cols-3 gap-8">
          {[
            { title: "上傳文件", desc: "支援 Word, Excel, PDF, PPTX 格式，系統會自動提取文字內容。" },
            { title: "智能翻譯", desc: "使用 DeepSeek AI 進行語境感知翻譯，確保翻譯品質。" },
            { title: "多語對照", desc: "翻譯結果將以對照形式呈現，並生成新的翻譯文件。" }
          ].map((item, idx) => (
            <div key={idx} className="text-center">
              <div className="text-2xl font-serif italic text-emerald-600/30 mb-2">0{idx + 1}</div>
              <h3 className="font-medium mb-1">{item.title}</h3>
              <p className="text-xs text-gray-500 leading-relaxed">{item.desc}</p>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
