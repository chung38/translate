/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import mammoth from 'mammoth';
import { Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType } from 'docx';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';
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
  Presentation,
  History
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// Set up PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

// DeepSeek API Configuration (Now handled via server proxy)
const DEEPSEEK_PROXY_URL = '/api/translate';

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

const AVAILABLE_LANGUAGES = [
  { id: 'th', name: '泰文', label: 'Thai', flag: '🇹🇭' },
  { id: 'id', name: '印尼文', label: 'Indonesian', flag: '🇮🇩' },
  { id: 'vi', name: '越南文', label: 'Vietnamese', flag: '🇻🇳' },
  { id: 'en', name: '英文', label: 'English', flag: '🇺🇸' },
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
  const [history, setHistory] = useState<{ name: string, date: string, blob: Blob, type: string }[]>([]);
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
      // Reset input value to allow same file selection
      e.target.value = '';
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
      await sleep(1000 + Math.random() * 500); 

      const industryContext = industry ? `。這是一個關於「${industry}」行業的文件，請使用該行業的專業術語進行翻譯` : '';
      const prompt = `你是一個專業的翻譯官${industryContext}。
      請將以下文字陣列中的每一項同時翻譯成以下語言：${targetLangs.join('、')}。

      請嚴格以 JSON 格式回傳結果，格式如下：
      {
        "translations": [
          { ${targetLangs.map(l => `"${l}": "翻譯內容"`).join(', ')} },
          ...
        ]
      }

      要求：
      1. 確保回傳的 "translations" 陣列長度與輸入的文字陣列長度完全一致 (${texts.length})。
      2. 每個物件的鍵 (Key) 必須完全對應目標語言名稱：${targetLangs.map(l => `"${l}"`).join(', ')}。
      3. 翻譯後的文字內容中，絕對不要包含任何語言標籤（例如不要出現 [英文] 或 [English] 等字樣）。
      4. 確保翻譯內容完全使用目標語言，不要夾雜原始語言。
      5. **重要：如果輸入文字中包含 <color hex="RRGGBB">...</color> 標籤，請在翻譯後的對應單字或片語上也保留這些標籤與相同的 hex 值。**
      6. 不要包含任何 Markdown 標籤（如 \`\`\`json）或額外文字，只回傳純 JSON 字串。

      待翻譯內容陣列：
      ${JSON.stringify(texts)}`;

      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 60000); // 60s timeout

      const response = await fetch(DEEPSEEK_PROXY_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ prompt }),
        signal: controller.signal
      });

      clearTimeout(timeoutId);

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
      
      // Clean up potential Markdown code blocks
      const cleanJson = resultText.replace(/^```json\n?/, '').replace(/\n?```$/, '').trim();
      
      let parsed;
      try {
        parsed = JSON.parse(cleanJson);
      } catch (e) {
        console.error('JSON Parse Error. Raw text:', resultText);
        // Try to find JSON inside the text if parsing failed
        const jsonMatch = resultText.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          parsed = JSON.parse(jsonMatch[0]);
        } else {
          throw new Error('無法解析 API 回傳的 JSON 格式');
        }
      }
      
      if (!parsed.translations || !Array.isArray(parsed.translations)) {
        throw new Error('API 回傳格式不正確 (缺少 translations 陣列)');
      }

      // Normalize keys to match targetLangs exactly
      const normalizedTranslations = parsed.translations.map((item: any) => {
        const newItem: any = {};
        targetLangs.forEach(lang => {
          // Try exact match
          if (item[lang]) {
            newItem[lang] = item[lang];
          } else {
            // Try fuzzy match (e.g. "English" for "英文")
            const keys = Object.keys(item);
            const fuzzyKey = keys.find(k => k.toLowerCase().includes(lang.toLowerCase()) || lang.toLowerCase().includes(k.toLowerCase()));
            if (fuzzyKey) {
              newItem[lang] = item[fuzzyKey];
            } else if (keys.length === 1 && targetLangs.length === 1) {
              // If only one language requested and one returned, assume it's the one
              newItem[lang] = item[keys[0]];
            }
          }
        });
        return newItem;
      });

      return normalizedTranslations;
    } catch (err: any) {
      const isRateLimit = err?.message?.includes('429') || JSON.stringify(err).includes('429');
      const isNetworkError = err?.message?.includes('Load failed') || err?.message?.includes('Failed to fetch') || err?.name === 'TypeError' || err?.message?.includes('NetworkError') || err?.name === 'AbortError';
      const isAuthError = err?.message?.includes('Authentication') || err?.message?.includes('API key') || err?.message?.includes('401') || err?.message?.includes('配置');
      
      // Retry for rate limits or transient network errors
      if ((isRateLimit || isNetworkError) && retryCount < 10) {
        const waitTime = isRateLimit 
          ? (Math.pow(2, retryCount) * 5000 + Math.random() * 2000)
          : (3000 + Math.random() * 2000); // Wait for network errors
          
        console.warn(`DeepSeek ${isRateLimit ? 'Rate limit' : 'Network error'} hit. Waiting ${Math.round(waitTime/1000)}s... (Attempt ${retryCount + 1})`);
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
    const BATCH_SIZE = 5;
    
    for (let i = 0; i < totalParagraphs; i += BATCH_SIZE) {
      if (isCancelledRef.current) throw new Error('翻譯已取消');
      
      const batchParas = paragraphs.slice(i, i + BATCH_SIZE);
      const paraData = batchParas.map(p => {
        // Get all runs in this paragraph to check for color
        const runs = Array.from(p.getElementsByTagNameNS(ns, 'r'));
        let taggedText = '';
        runs.forEach(r => {
          const t = r.getElementsByTagNameNS(ns, 't')[0];
          if (!t) return;
          const text = t.textContent || '';
          const rPr = r.getElementsByTagNameNS(ns, 'rPr')[0];
          const color = rPr?.getElementsByTagNameNS(ns, 'color')[0];
          const hex = color?.getAttribute('w:val');
          
          // If color is present and not default black/auto, wrap it in tags
          if (hex && hex !== '000000' && hex !== 'auto') {
            taggedText += `<color hex="${hex}">${text}</color>`;
          } else {
            taggedText += text;
          }
        });
        return { p, taggedText };
      }).filter(data => data.taggedText.trim().length > 0);

      if (paraData.length > 0) {
        // Translating progress: 10% to 90%
        const currentProgress = 10 + Math.round((Math.min(i + BATCH_SIZE, totalParagraphs) / totalParagraphs) * 80);
        updateProgress(currentProgress, 'translating');
        
        const textsToTranslate = paraData.map(d => d.taggedText);
        const batchTranslations = await translateBatch(textsToTranslate, selectedLanguages);
        
        paraData.forEach((data, idx) => {
          const translations = batchTranslations[idx];
          const p = data.p;
          
          // Find the first run to copy its properties
          const firstRun = p.getElementsByTagNameNS(ns, 'r')[0];
          const firstRunPr = firstRun?.getElementsByTagNameNS(ns, 'rPr')[0];

          // Ensure paragraph spacing is consistent to prevent large gaps in wrapped text (especially for Vietnamese)
          let pPr = p.getElementsByTagNameNS(ns, 'pPr')[0];
          if (!pPr) {
            pPr = xmlDoc.createElementNS(ns, 'w:pPr');
            p.insertBefore(pPr, p.firstChild);
          }
          let spacing = pPr.getElementsByTagNameNS(ns, 'spacing')[0];
          if (!spacing) {
            spacing = xmlDoc.createElementNS(ns, 'w:spacing');
            pPr.appendChild(spacing);
          }
          // Set line spacing to a tighter value (200 twips, approx 0.85x) and remove before/after spacing
          spacing.setAttribute('w:line', '200');
          spacing.setAttribute('w:lineRule', 'auto');
          spacing.setAttribute('w:before', '0');
          spacing.setAttribute('w:after', '0');
          spacing.setAttribute('w:beforeAutospacing', '0');
          spacing.setAttribute('w:afterAutospacing', '0');

          selectedLanguages.forEach(lang => {
            const translatedText = translations[lang];
            if (!translatedText) return;

            // Split by color tags
            const parts = translatedText.split(/(<color hex="[0-9A-Fa-f]{6}">.*?<\/color>)/g);
            
            parts.forEach((part, pIdx) => {
              // Create a new run for each part
              const newRun = xmlDoc.createElementNS(ns, 'w:r');
              
              // Copy properties if they exist
              if (firstRunPr) {
                const newPr = firstRunPr.cloneNode(true) as Element;
                // Remove properties that might cause weird spacing or incorrect colors in translations
                const propsToRemove = ['spacing', 'w', 'kern', 'color'];
                propsToRemove.forEach(prop => {
                  const els = newPr.getElementsByTagNameNS(ns, prop);
                  while (els.length > 0) {
                    newPr.removeChild(els[0]);
                  }
                });
                
                // If this part is a color tag, apply the color
                if (part.startsWith('<color')) {
                  const match = part.match(/<color hex="([0-9A-Fa-f]{6})">(.*?)<\/color>/);
                  if (match) {
                    const hex = match[1];
                    const colorEl = xmlDoc.createElementNS(ns, 'w:color');
                    colorEl.setAttribute('w:val', hex);
                    newPr.appendChild(colorEl);
                  }
                }

                // Force a standard font for translations to avoid metric issues with diacritics
                let rFonts = newPr.getElementsByTagNameNS(ns, 'rFonts')[0];
                if (!rFonts) {
                  rFonts = xmlDoc.createElementNS(ns, 'w:rFonts');
                  newPr.appendChild(rFonts);
                }
                rFonts.setAttribute('w:ascii', 'Arial');
                rFonts.setAttribute('w:hAnsi', 'Arial');
                rFonts.setAttribute('w:cs', 'Arial'); // Complex script font for Vietnamese
                
                // Set language to Vietnamese if applicable to improve rendering
                if (lang === '越南文' || lang === 'Vietnamese' || lang === 'vi') {
                  let langEl = newPr.getElementsByTagNameNS(ns, 'lang')[0];
                  if (!langEl) {
                    langEl = xmlDoc.createElementNS(ns, 'w:lang');
                    newPr.appendChild(langEl);
                  }
                  langEl.setAttribute('w:val', 'vi-VN');
                  langEl.setAttribute('w:eastAsia', 'vi-VN');
                  langEl.setAttribute('w:bidi', 'vi-VN');
                }
                
                newRun.appendChild(newPr);
              }

              // Add a break before the translation (only for the first part of each language)
              if (pIdx === 0) {
                const br = xmlDoc.createElementNS(ns, 'w:br');
                newRun.appendChild(br);
              }

              // Add the text
              const t = xmlDoc.createElementNS(ns, 'w:t');
              // Important: Preserve spaces in Word runs
              t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
              
              if (part.startsWith('<color')) {
                const match = part.match(/<color hex="([0-9A-Fa-f]{6})">(.*?)<\/color>/);
                t.textContent = match ? match[2] : part;
              } else {
                t.textContent = part;
              }
              newRun.appendChild(t);

              // Append the translation to the paragraph
              p.appendChild(newRun);
            });
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
    return { blob: content, name: fileName };
  };

  const processExcel = async (file: File, updateProgress: (p: number, status?: TranslationStatus) => void) => {
    updateProgress(5, 'processing');
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);

      updateProgress(10, 'translating');
      
      const sheetsToProcess: { sheet: ExcelJS.Worksheet, cells: any[] }[] = [];
      
      workbook.eachSheet((worksheet) => {
        const cells: any[] = [];
        // Iterate over all rows that have values
        worksheet.eachRow({ includeEmpty: false }, (row) => {
          row.eachCell({ includeEmpty: false }, (cell) => {
            let text = '';
            const val = cell.value;
            let originalFont = cell.font;
            
            if (typeof val === 'string') {
              text = val;
            } else if (val && typeof val === 'object') {
              if ('richText' in val) {
                const rt = (val as any).richText;
                text = rt.map((rtPart: any) => rtPart.text || '').join('');
                if (rt.length > 0 && rt[0].font) {
                  originalFont = rt[0].font;
                }
              } else if ('result' in val) {
                text = (val as any).result?.toString() || '';
              } else if ('text' in val) {
                text = (val as any).text?.toString() || '';
              } else {
                // Fallback to cell.text which is a getter in exceljs
                text = cell.text;
              }
            } else if (val !== null && val !== undefined) {
              text = val.toString();
            }

            if (text && text.trim()) {
              cells.push({ 
                r: row.number, 
                c: cell.col, 
                text: text.trim(),
                address: cell.address,
                font: originalFont ? JSON.parse(JSON.stringify(originalFont)) : undefined
              });
            }
          });
        });
        
        if (cells.length > 0) {
          sheetsToProcess.push({ sheet: worksheet, cells });
        }
      });

      if (sheetsToProcess.length === 0) {
        throw new Error('找不到可翻譯的文字內容');
      }

      const totalSheets = sheetsToProcess.length;
      for (let s = 0; s < totalSheets; s++) {
        if (isCancelledRef.current) throw new Error('翻譯已取消');
        
        const { sheet, cells } = sheetsToProcess[s];
        const BATCH_SIZE = 10;
        
        for (let i = 0; i < cells.length; i += BATCH_SIZE) {
          if (isCancelledRef.current) throw new Error('翻譯已取消');
          
          const currentProgress = 10 + Math.round(((s * cells.length + i) / (totalSheets * Math.max(1, cells.length))) * 80);
          updateProgress(currentProgress, 'translating');
          
          const batch = cells.slice(i, i + BATCH_SIZE);
          const batchTexts = batch.map(c => c.text);
          
          const batchTranslations = await translateBatch(batchTexts, selectedLanguages);
          
          // Debug: Show first translation in status
          if (batchTranslations.length > 0) {
            const first = batchTranslations[0];
            const firstLang = selectedLanguages[0];
            if (first[firstLang]) {
              setStatusMessage(`已翻譯: ${batchTexts[0].substring(0, 10)}... -> ${first[firstLang].substring(0, 10)}...`);
            }
          }

          batch.forEach((cellInfo, idx) => {
            const translations = batchTranslations[idx];
            const cell = sheet.getCell(cellInfo.address);
            const row = sheet.getRow(cellInfo.r);
            
            // Use original font for the original text
            const originalFont = cellInfo.font || {};
            const richText: any[] = [
              { text: cellInfo.text, font: { ...originalFont } }
            ];
            
            let hasTranslation = false;
            for (const lang of selectedLanguages) {
              const translatedText = translations[lang];
              if (translatedText && !translatedText.includes('翻譯出錯')) {
                // Use original font including color for translation
                richText.push({ 
                  text: `\n${translatedText}`, 
                  font: { ...originalFont } 
                });
                hasTranslation = true;
              } else {
                richText.push({ 
                  text: `\n(待翻譯: ${lang})`, 
                  font: { ...originalFont, size: (originalFont.size || 11) - 1, italic: true, color: { argb: 'FF888888' } } 
                });
              }
            }
            
            // Update the cell value with RichText
            cell.value = { richText };
            
            // Preserve original alignment but ensure wrapText is true
            const originalAlignment = cell.alignment || {};
            cell.alignment = { 
              ...originalAlignment,
              wrapText: true, 
              vertical: 'top'
            };

            // Increase row height to accommodate multiple lines (approx 15pt per line)
            const minHeight = (selectedLanguages.length + 1) * 18;
            if (!row.height || row.height < minHeight) {
              row.height = minHeight;
            }
          });
        }
      }

      // Add a summary sheet as a backup
      const summarySheet = workbook.addWorksheet('翻譯對照表');
      summarySheet.columns = [
        { header: '工作表', key: 'sheetName', width: 15 },
        { header: '儲存格', key: 'address', width: 10 },
        { header: '原始內容', key: 'original', width: 40 },
        ...selectedLanguages.map(lang => ({ header: lang, key: lang, width: 40 }))
      ];

      sheetsToProcess.forEach(({ sheet, cells }) => {
        cells.forEach(cellInfo => {
          // We need to find the translations for this cell again or store them
          // For simplicity, let's just add the original content to the summary
          // Actually, let's skip the summary if it's too complex to re-map
        });
      });

      updateProgress(95, 'generating');
      const buffer = await workbook.xlsx.writeBuffer();
      const excelBlob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const fileName = `translated_${file.name}`;
      saveAs(excelBlob, fileName);
      updateProgress(100, 'generating');
      return { blob: excelBlob, name: fileName };
    } catch (err: any) {
      console.error('Excel processing error:', err);
      throw err;
    }
  };

  const processPdf = async (file: File, updateProgress: (p: number, status?: TranslationStatus) => void) => {
    updateProgress(5, 'processing');
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const doc = new jsPDF();
    
    updateProgress(10, 'translating');
    const totalPages = pdf.numPages;

    // Helper to render text using canvas to avoid CJK font issues in jsPDF
    const renderText = (text: string, x: number, y: number, maxWidth: number, fontSize: number, isBold = false, color = '#000000'): { nextY: number; remainingText: string } => {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) return { nextY: y, remainingText: '' };

      const scale = 4;
      const fontStack = '"Microsoft JhengHei", "微軟正黑體", "Noto Sans TC", "PingFang TC", sans-serif';
      ctx.font = `${isBold ? 'bold ' : ''}${fontSize * scale}px ${fontStack}`;
      
      const words = text.split('');
      let line = '';
      let currentY = y;
      const lineHeight = fontSize * 1.2;

      const renderLine = (lineText: string, ty: number) => {
        const lineCanvas = document.createElement('canvas');
        const lineCtx = lineCanvas.getContext('2d');
        if (!lineCtx) return;
        
        const metrics = ctx.measureText(lineText);
        // Use a more precise height to avoid squashing and extra whitespace
        const canvasHeight = fontSize * scale * 1.4;
        lineCanvas.width = metrics.width;
        lineCanvas.height = canvasHeight;
        
        lineCtx.font = `${isBold ? 'bold ' : ''}${fontSize * scale}px ${fontStack}`;
        lineCtx.fillStyle = color;
        lineCtx.textBaseline = 'top';
        lineCtx.fillText(lineText, 0, 0);
        
        const imgData = lineCanvas.toDataURL('image/png');
        // The height in PDF should be (canvasHeight / scale) to maintain aspect ratio
        doc.addImage(imgData, 'PNG', x, ty, metrics.width / scale, canvasHeight / scale);
      };

      for (let n = 0; n < words.length; n++) {
        const testLine = line + words[n];
        const metrics = ctx.measureText(testLine);
        const testWidth = metrics.width / scale;
        
        if (testWidth > maxWidth && n > 0) {
          renderLine(line, currentY);
          line = words[n];
          currentY += lineHeight;
          
          if (currentY > 275) {
            doc.addPage();
            currentY = 20;
            return { nextY: currentY, remainingText: words.slice(n).join('') };
          }
        } else {
          line = testLine;
        }
      }
      
      if (line) {
        renderLine(line, currentY);
        currentY += lineHeight;
      }
      
      return { nextY: currentY, remainingText: '' };
    };

    const safeRenderText = (text: string, x: number, y: number, maxWidth: number, fontSize: number, isBold = false, color = '#000000') => {
      let result = renderText(text, x, y, maxWidth, fontSize, isBold, color);
      while (result.remainingText) {
        result = renderText(result.remainingText, x, result.nextY, maxWidth, fontSize, isBold, color);
      }
      return result.nextY;
    };

    for (let i = 1; i <= totalPages; i++) {
      if (isCancelledRef.current) throw new Error('翻譯已取消');
      
      const currentProgress = 10 + Math.round((i / totalPages) * 80);
      updateProgress(currentProgress, 'translating');
      
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      const strings = textContent.items.map((item: any) => item.str);
      const fullText = strings.join(' ').trim();

      if (fullText) {
        const chunks = fullText.match(/.{1,500}/g) || [fullText];
        const batchTranslations = await translateBatch(chunks, selectedLanguages);
        
        if (i > 1) doc.addPage();
        
        let yPos = 20;
        const margin = 15;
        const pageWidth = doc.internal.pageSize.getWidth();
        const contentWidth = pageWidth - (margin * 2);

        // Original Text (Restored as requested, but without green headers)
        yPos = safeRenderText(fullText, margin, yPos, contentWidth, 10, false, '#1E293B');
        yPos += 10;

        // Translations - Only render the translated text to remove extra green headers/labels
        for (const lang of selectedLanguages) {
          if (yPos > 260) {
            doc.addPage();
            yPos = 20;
          }
          
          const combinedTrans = batchTranslations.map(t => t[lang] || '').join(' ');
          yPos = safeRenderText(combinedTrans, margin, yPos, contentWidth, 10, false, '#000000');
          yPos += 6;
        }
      }
    }

    updateProgress(95, 'generating');
    const pdfBlob = doc.output('blob');
    const fileName = `translated_${file.name}`;
    saveAs(pdfBlob, fileName);
    updateProgress(100, 'generating');
    return { blob: pdfBlob, name: fileName };
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
        // Get all runs in this paragraph to check for color
        const runs = Array.from(p.getElementsByTagNameNS(nsA, 'r'));
        let taggedText = '';
        runs.forEach(r => {
          const t = r.getElementsByTagNameNS(nsA, 't')[0];
          if (!t) return;
          const text = t.textContent || '';
          const rPr = r.getElementsByTagNameNS(nsA, 'rPr')[0];
          const solidFill = rPr?.getElementsByTagNameNS(nsA, 'solidFill')[0];
          const srgbClr = solidFill?.getElementsByTagNameNS(nsA, 'srgbClr')[0];
          const hex = srgbClr?.getAttribute('val');
          
          // If color is present and not default black/auto, wrap it in tags
          if (hex && hex !== '000000') {
            taggedText += `<color hex="${hex}">${text}</color>`;
          } else {
            taggedText += text;
          }
        });
        return { p, taggedText };
      }).filter(data => data.taggedText.trim().length > 0);

      if (paraData.length > 0) {
        const textsToTranslate = paraData.map(d => d.taggedText);
        const batchTranslations = await translateBatch(textsToTranslate, selectedLanguages);
        
        paraData.forEach((data, idx) => {
          const translations = batchTranslations[idx];
          const p = data.p;
          
          // Find the first run to copy its properties
          const firstRun = p.getElementsByTagNameNS(nsA, 'r')[0];
          const firstRunPr = firstRun?.getElementsByTagNameNS(nsA, 'rPr')[0];

          // Set tighter line spacing for PPTX (85% of standard)
          let pPr = p.getElementsByTagNameNS(nsA, 'pPr')[0];
          if (!pPr) {
            pPr = xmlDoc.createElementNS(nsA, 'a:pPr');
            p.insertBefore(pPr, p.firstChild);
          }
          let lnSpc = pPr.getElementsByTagNameNS(nsA, 'lnSpc')[0];
          if (!lnSpc) {
            lnSpc = xmlDoc.createElementNS(nsA, 'a:lnSpc');
            pPr.appendChild(lnSpc);
          }
          let spcPct = lnSpc.getElementsByTagNameNS(nsA, 'spcPct')[0];
          if (!spcPct) {
            spcPct = xmlDoc.createElementNS(nsA, 'a:spcPct');
            lnSpc.appendChild(spcPct);
          }
          spcPct.setAttribute('val', '85000'); // 85% line spacing

          selectedLanguages.forEach(lang => {
            const translatedText = translations[lang];
            if (!translatedText) return;

            // Split by color tags
            const parts = translatedText.split(/(<color hex="[0-9A-Fa-f]{6}">.*?<\/color>)/g);
            
            parts.forEach((part, pIdx) => {
              // Add a break (a:br) only before the first part of each language
              if (pIdx === 0) {
                const br = xmlDoc.createElementNS(nsA, 'a:br');
                if (firstRunPr) {
                  br.appendChild(firstRunPr.cloneNode(true));
                }
                p.appendChild(br);
              }

              // Create a new run for each part
              const newRun = xmlDoc.createElementNS(nsA, 'a:r');
              
              // Copy properties if they exist
              if (firstRunPr) {
                const newPr = firstRunPr.cloneNode(true) as Element;
                
                // Remove color/fill properties to prevent entire translation from being colored
                // if only part of the original was colored (since we copy from the first run)
                const propsToRemove = ['solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
                propsToRemove.forEach(prop => {
                  const els = newPr.getElementsByTagNameNS(nsA, prop);
                  while (els.length > 0) {
                    newPr.removeChild(els[0]);
                  }
                });

                // If this part is a color tag, apply the color
                if (part.startsWith('<color')) {
                  const match = part.match(/<color hex="([0-9A-Fa-f]{6})">(.*?)<\/color>/);
                  if (match) {
                    const hex = match[1];
                    const solidFill = xmlDoc.createElementNS(nsA, 'a:solidFill');
                    const srgbClr = xmlDoc.createElementNS(nsA, 'a:srgbClr');
                    srgbClr.setAttribute('val', hex);
                    solidFill.appendChild(srgbClr);
                    newPr.appendChild(solidFill);
                  }
                }

                // Force Arial for Vietnamese in PPTX to avoid metric issues
                if (lang === '越南文' || lang === 'Vietnamese' || lang === 'vi') {
                  let latin = newPr.getElementsByTagNameNS(nsA, 'latin')[0];
                  if (!latin) {
                    latin = xmlDoc.createElementNS(nsA, 'a:latin');
                    newPr.appendChild(latin);
                  }
                  latin.setAttribute('typeface', 'Arial');
                  
                  let cs = newPr.getElementsByTagNameNS(nsA, 'cs')[0];
                  if (!cs) {
                    cs = xmlDoc.createElementNS(nsA, 'a:cs');
                    newPr.appendChild(cs);
                  }
                  cs.setAttribute('typeface', 'Arial');
                }
                newRun.appendChild(newPr);
              }

              // Add the text
              const t = xmlDoc.createElementNS(nsA, 'a:t');
              // PowerPoint text nodes generally preserve spaces, but we ensure the content is clean
              if (part.startsWith('<color')) {
                const match = part.match(/<color hex="([0-9A-Fa-f]{6})">(.*?)<\/color>/);
                t.textContent = match ? match[2] : part;
              } else {
                t.textContent = part;
              }
              newRun.appendChild(t);

              // Append the translation to the paragraph
              p.appendChild(newRun);
            });
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
    return { blob: content, name: fileName };
  };

  const processFiles = async () => {
    if (files.length === 0 || selectedLanguages.length === 0) return;

    try {
      setStatus('processing');
      setError(null);
      setProgress(0);
      setFileProgress({});
      isCancelledRef.current = false;

      const filesToProcess = [...files];
      for (let i = 0; i < filesToProcess.length; i++) {
        if (isCancelledRef.current) break;
        
        const currentFile = filesToProcess[i];
        const extension = currentFile.name.split('.').pop()?.toLowerCase();
        
        setStatusMessage(`正在處理第 ${i + 1} / ${filesToProcess.length} 份文件: ${currentFile.name}`);
        
        // Update overall progress based on file index
        const baseProgress = (i / filesToProcess.length) * 100;
        const fileWeight = 100 / filesToProcess.length;
        
        const updateFileProgress = (p: number, currentStatus?: TranslationStatus) => {
          setFileProgress(prev => ({ ...prev, [currentFile.name]: p }));
          setProgress(Math.round(baseProgress + (p * fileWeight / 100)));
          if (currentStatus) setStatus(currentStatus);
        };

        try {
          let result: { blob: Blob, name: string } | undefined;
          switch (extension) {
            case 'docx':
              result = await processDocx(currentFile, updateFileProgress);
              break;
            case 'xlsx':
              result = await processExcel(currentFile, updateFileProgress);
              break;
            case 'pdf':
              result = await processPdf(currentFile, updateFileProgress);
              break;
            case 'pptx':
              result = await processPptx(currentFile, updateFileProgress);
              break;
            default:
              console.warn(`不支援的檔案格式: ${currentFile.name}`);
          }
          
          if (result) {
            const newEntry = { 
              name: result.name, 
              date: new Date().toLocaleTimeString(), 
              blob: result.blob, 
              type: extension || 'unknown' 
            };
            setHistory(prev => [newEntry, ...prev].slice(0, 3));
          }

          // Remove file from list after successful processing
          setFiles(prev => prev.filter(f => f !== currentFile));
        } catch (fileErr) {
          console.error(`處理檔案 ${currentFile.name} 時發生錯誤:`, fileErr);
          // We continue to the next file even if one fails, 
          // but we might want to keep the failed file in the list
          setError(`處理 ${currentFile.name} 時發生錯誤: ${fileErr instanceof Error ? fileErr.message : '未知錯誤'}`);
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
    <div className="min-h-screen bg-[#FAFAFA] font-sans text-[#1A1A1A] relative overflow-hidden">
      {/* Background Accents */}
      <div className="absolute top-0 left-0 w-full h-[500px] bg-gradient-to-b from-emerald-50/50 to-transparent -z-10" />
      <div className="absolute top-[-10%] right-[-10%] w-[40%] h-[40%] bg-emerald-100/20 blur-[120px] rounded-full -z-10" />
      <div className="absolute bottom-[-10%] left-[-10%] w-[40%] h-[40%] bg-blue-100/20 blur-[120px] rounded-full -z-10" />

      <div className="max-w-4xl mx-auto p-4 sm:p-6 md:p-8 relative z-10">
        {/* Header */}
        <header className="mb-8 md:mb-12 text-center">
          <motion.div 
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="inline-flex items-center justify-center mb-4"
          >
            <div className="relative">
              <div className="absolute inset-0 bg-emerald-500/10 blur-xl rounded-full" />
              <div className="relative w-12 h-12 bg-white rounded-2xl shadow-sm border border-gray-100 flex items-center justify-center">
                <Languages className="w-6 h-6 text-emerald-600" />
              </div>
            </div>
          </motion.div>

          <div className="space-y-1">
            <motion.h1 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ delay: 0.1, duration: 0.5 }}
              className="text-2xl md:text-4xl font-bold tracking-tight text-gray-900"
            >
              全能文件<span className="text-emerald-600">多語翻譯器</span>
            </motion.h1>

            <motion.p
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              transition={{ delay: 0.2, duration: 0.5 }}
              className="text-gray-400 text-xs md:text-sm font-medium tracking-wide"
            >
              Professional AI-powered document translation
            </motion.p>
          </div>
        </header>

        {/* Instructions */}
        <div className="mb-16 grid grid-cols-1 md:grid-cols-3 gap-6 md:gap-8">
          {[
            { title: "上傳文件", desc: "支援 Word, Excel, PDF, PPTX 格式，自動提取內容。" },
            { title: "智能翻譯", desc: "使用 AI 進行語境感知翻譯，確保翻譯品質。" },
            { title: "多語對照", desc: "翻譯結果以對照形式呈現，並生成新文件。" }
          ].map((item, idx) => (
            <motion.div 
              key={idx} 
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ delay: 0.3 + idx * 0.1 }}
              className="relative p-6 rounded-2xl bg-white/50 border border-white shadow-sm hover:shadow-md transition-all group"
            >
              <div className="text-3xl font-serif italic text-emerald-600/20 mb-3 group-hover:text-emerald-600/40 transition-colors">0{idx + 1}</div>
              <h3 className="font-semibold text-gray-800 mb-2">{item.title}</h3>
              <p className="text-xs text-gray-500 leading-relaxed font-light">{item.desc}</p>
            </motion.div>
          ))}
        </div>

        {/* Main Card */}
        <motion.div 
          initial={{ opacity: 0, y: 30 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.6 }}
          className="bg-white rounded-[40px] shadow-[0_32px_64px_-16px_rgba(0,0,0,0.08)] border border-black/5 overflow-hidden"
        >
          <div className="p-6 md:p-16">
            {/* Upload Section */}
            <div 
              onClick={() => fileInputRef.current?.click()}
              className={`
                relative group cursor-pointer border-2 border-dashed rounded-[32px] p-8 md:p-12 transition-all duration-500
                ${files.length > 0 ? 'border-emerald-200 bg-emerald-50/20' : 'border-gray-200 hover:border-emerald-400 hover:bg-gray-50/50'}
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
                  w-16 h-16 rounded-full flex items-center justify-center mb-4 transition-all duration-500 group-hover:scale-110
                  ${files.length > 0 ? 'bg-emerald-100 text-emerald-600 shadow-inner' : 'bg-gray-50 text-gray-400'}
                `}>
                  <Upload className="w-7 h-7" />
                </div>
                
                <div className="max-w-xs">
                  <p className="text-lg font-medium text-gray-800 mb-2">點擊或拖拽上傳文件</p>
                  <p className="text-sm text-gray-400 font-light">支援 .docx, .xlsx, .pdf, .pptx 格式 (可多選)</p>
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
                      className="flex items-center justify-between p-4 bg-gray-50/50 rounded-2xl border border-gray-100/50 hover:bg-white hover:shadow-sm transition-all"
                    >
                      <div className="flex items-center gap-4 min-w-0">
                        <div className="w-10 h-10 rounded-xl bg-white flex items-center justify-center text-emerald-600 shadow-sm">
                          {f.name.endsWith('.docx') && <FileText className="w-5 h-5" />}
                          {f.name.endsWith('.xlsx') && <FileSpreadsheet className="w-5 h-5" />}
                          {f.name.endsWith('.pdf') && <FileIcon className="w-5 h-5" />}
                          {f.name.endsWith('.pptx') && <Presentation className="w-5 h-5" />}
                        </div>
                        <div className="min-w-0">
                          <p className="text-sm font-medium text-gray-700 truncate">{f.name}</p>
                          <p className="text-[10px] text-gray-400 font-light">{(f.size / 1024).toFixed(1)} KB</p>
                        </div>
                      </div>
                      
                      <div className="flex items-center gap-4">
                        {status !== 'idle' && status !== 'error' && (
                          <div className="flex flex-col items-end gap-1.5">
                            <span className="text-[10px] font-mono font-bold text-emerald-600">
                              {fileProgress[f.name] || 0}%
                            </span>
                            <div className="w-16 h-1 bg-gray-200 rounded-full overflow-hidden">
                              <div 
                                className="h-full bg-emerald-500 transition-all duration-500" 
                                style={{ width: `${fileProgress[f.name] || 0}%` }}
                              />
                            </div>
                          </div>
                        )}
                        {status === 'idle' && (
                          <button 
                            onClick={() => removeFile(idx)}
                            className="p-2 hover:bg-red-50 rounded-full text-gray-300 hover:text-red-500 transition-all"
                          >
                            <AlertCircle className="w-4 h-4 rotate-45" />
                          </button>
                        )}
                        {status === 'completed' && fileProgress[f.name] === 100 && (
                          <div className="w-8 h-8 rounded-full bg-emerald-50 flex items-center justify-center">
                            <CheckCircle2 className="w-4 h-4 text-emerald-500" />
                          </div>
                        )}
                      </div>
                    </motion.div>
                  ))}
                </motion.div>
              )}
            </AnimatePresence>

            {/* Settings */}
            <div className="mt-12 space-y-10">
              <div>
                <label className="block text-[10px] uppercase tracking-[0.2em] font-bold text-gray-400 mb-4 ml-1">
                  工廠行業 (可選，使翻譯更精準)
                </label>
                <div className="relative">
                  <input
                    type="text"
                    value={industry}
                    onChange={(e) => setIndustry(e.target.value)}
                    placeholder="例如：電子、紡織、汽車..."
                    className="w-full px-5 py-4 rounded-[20px] border border-gray-100 bg-gray-50/30 focus:bg-white focus:border-emerald-400 focus:ring-4 focus:ring-emerald-400/5 outline-none transition-all text-sm placeholder:text-gray-300"
                  />
                </div>
              </div>

              <div>
                <label className="block text-[10px] uppercase tracking-[0.2em] font-bold text-gray-400 mb-6 ml-1">
                  選擇目標語言 (可多選)
                </label>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-4">
                  {AVAILABLE_LANGUAGES.map((lang) => (
                    <button
                      key={lang.id}
                      onClick={() => toggleLanguage(lang.name)}
                      className={`
                        group relative flex items-center justify-between px-5 py-4 rounded-[20px] border transition-all duration-300
                        ${selectedLanguages.includes(lang.name) 
                          ? 'border-emerald-500 bg-emerald-500 text-white shadow-[0_8px_20px_-6px_rgba(16,185,129,0.4)]' 
                          : 'border-gray-100 bg-white hover:border-emerald-200 hover:bg-emerald-50/30 text-gray-600'}
                      `}
                    >
                      <div className="flex items-center gap-3">
                        <span className="text-lg">{lang.flag}</span>
                        <span className="text-xs font-medium tracking-wide">{lang.name}</span>
                      </div>
                      {selectedLanguages.includes(lang.name) && (
                        <CheckCircle2 className="w-4 h-4 text-white" />
                      )}
                    </button>
                  ))}
                </div>
              </div>
              
              <div className="flex flex-col gap-3 pt-4">
                <button
                  disabled={files.length === 0 || selectedLanguages.length === 0 || status === 'processing' || status === 'translating' || status === 'generating'}
                  onClick={processFiles}
                  className={`
                    w-full h-[56px] md:h-[64px] rounded-[24px] font-semibold text-sm tracking-widest uppercase transition-all duration-500 flex items-center justify-center gap-3
                    ${files.length === 0 || selectedLanguages.length === 0 || status === 'processing' || status === 'translating' || status === 'generating'
                      ? 'bg-gray-100 text-gray-400 cursor-not-allowed'
                      : 'bg-gray-900 text-white hover:bg-black hover:shadow-2xl hover:-translate-y-1 active:scale-[0.98]'}
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
                    className="w-full h-[48px] rounded-[20px] font-medium text-red-500 hover:bg-red-50 transition-all text-sm border border-red-100"
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
        </motion.div>

        {/* Translation History */}
        <AnimatePresence>
          {history.length > 0 && (
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="mt-12 bg-white rounded-[32px] shadow-[0_20px_40px_-12px_rgba(0,0,0,0.04)] border border-black/5 overflow-hidden"
            >
              <div className="p-6 md:p-8 border-b border-gray-50 flex items-center justify-between bg-gray-50/30">
                <div className="flex items-center gap-4">
                  <div className="w-10 h-10 rounded-xl bg-emerald-50 flex items-center justify-center">
                    <History className="w-5 h-5 text-emerald-600" />
                  </div>
                  <div>
                    <h2 className="text-sm font-semibold text-gray-800">最近翻譯</h2>
                    <p className="text-[10px] text-gray-400 uppercase tracking-wider font-medium">Translation History</p>
                  </div>
                </div>
              </div>
              <div className="divide-y divide-gray-50">
                {history.map((item, idx) => (
                  <div key={idx} className="p-5 md:p-6 flex items-center justify-between hover:bg-gray-50/50 transition-all group">
                    <div className="flex items-center gap-5">
                      <div className="w-12 h-12 rounded-2xl bg-gray-50 flex items-center justify-center text-gray-400 group-hover:bg-white group-hover:shadow-sm transition-all">
                        {item.type === 'docx' && <FileText className="w-6 h-6 text-blue-500/70" />}
                        {item.type === 'xlsx' && <FileSpreadsheet className="w-6 h-6 text-emerald-500/70" />}
                        {item.type === 'pdf' && <FileIcon className="w-6 h-6 text-red-500/70" />}
                        {item.type === 'pptx' && <Presentation className="w-6 h-6 text-orange-500/70" />}
                        {!['docx', 'xlsx', 'pdf', 'pptx'].includes(item.type) && <FileIcon className="w-6 h-6" />}
                      </div>
                      <div>
                        <h3 className="text-sm font-medium text-gray-700 truncate max-w-[180px] sm:max-w-[400px]">
                          {item.name}
                        </h3>
                        <div className="flex items-center gap-2 mt-1">
                          <span className="text-[10px] text-gray-400 font-light">{item.date}</span>
                          <span className="w-1 h-1 rounded-full bg-gray-200" />
                          <span className="text-[10px] text-emerald-600 font-medium uppercase tracking-tighter">Success</span>
                        </div>
                      </div>
                    </div>
                    <button
                      onClick={() => saveAs(item.blob, item.name)}
                      className="w-10 h-10 rounded-full flex items-center justify-center hover:bg-emerald-50 text-emerald-600 transition-all group/btn border border-transparent hover:border-emerald-100"
                      title="重新下載"
                    >
                      <Download className="w-4 h-4 group-hover/btn:scale-110 transition-transform" />
                    </button>
                  </div>
                ))}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </div>
  );
}
