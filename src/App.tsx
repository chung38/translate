/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import mammoth from 'mammoth';
import { Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType } from 'docx';
import { saveAs } from 'file-saver';
import { 
  Upload, 
  FileText, 
  Languages, 
  Download, 
  Loader2, 
  CheckCircle2, 
  AlertCircle,
  ArrowRight
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

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
  const [file, setFile] = useState<File | null>(null);
  const [selectedLanguages, setSelectedLanguages] = useState<string[]>(['英文']);
  const [status, setStatus] = useState<TranslationStatus>('idle');
  const [progress, setProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile && selectedFile.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
      setFile(selectedFile);
      setError(null);
      setStatus('idle');
    } else {
      setError('請上傳有效的 .docx 檔案');
    }
  };

  const toggleLanguage = (langName: string) => {
    setSelectedLanguages(prev => 
      prev.includes(langName) 
        ? prev.filter(l => l !== langName)
        : [...prev, langName]
    );
  };

  const translateBatch = async (text: string, targetLangs: string[], retryCount = 0): Promise<Record<string, string>> => {
    if (!text.trim() || targetLangs.length === 0) return {};
    
    try {
      // Add a delay between batches to stay under RPM limits
      await sleep(1000); 

      const prompt = `你是一個專業的翻譯官。請將以下文字同時翻譯成以下語言：${targetLangs.join('、')}。
      請嚴格以 JSON 格式回傳結果，格式如下：
      {
        "${targetLangs[0]}": "翻譯內容",
        ...
      }
      不要包含任何 Markdown 標籤（如 \`\`\`json）或額外文字，只回傳純 JSON 字串。
      
      待翻譯內容：
      ${text}`;

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
          // If not JSON, try to get text
          const textError = await response.text();
          errorMessage = textError || errorMessage;
        }
        throw new Error(errorMessage);
      }

      const data = await response.json();
      const resultText = data.choices[0].message.content.trim();
      return JSON.parse(resultText);
    } catch (err: any) {
      const isRateLimit = err?.message?.includes('429') || JSON.stringify(err).includes('429');
      const isAuthError = err?.message?.includes('Authentication') || err?.message?.includes('API key') || err?.message?.includes('401') || err?.message?.includes('配置');
      
      if (isRateLimit && retryCount < 5) {
        const waitTime = Math.pow(2, retryCount) * 5000 + Math.random() * 2000;
        console.warn(`DeepSeek Rate limit hit. Waiting ${Math.round(waitTime/1000)}s... (Attempt ${retryCount + 1})`);
        await sleep(waitTime);
        return translateBatch(text, targetLangs, retryCount + 1);
      }

      if (isAuthError) {
        throw new Error(err.message || 'DeepSeek API Key 驗證失敗，請檢查設定是否正確。');
      }

      console.error(`DeepSeek Batch translation error:`, err);
      const errorResult: Record<string, string> = {};
      targetLangs.forEach(lang => errorResult[lang] = `(翻譯出錯: ${err?.message || 'API 錯誤'})`);
      return errorResult;
    }
  };

  const processFile = async () => {
    if (!file || selectedLanguages.length === 0) return;

    try {
      setStatus('processing');
      setError(null);
      setProgress(0);

      // 1. Convert Docx to HTML (preserves images as base64)
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.convertToHtml({ arrayBuffer });
      const htmlContent = result.value;

      if (!htmlContent.trim()) {
        throw new Error('檔案內容為空');
      }

      // 2. Parse HTML and process elements
      setStatus('translating');
      const parser = new DOMParser();
      const doc = parser.parseFromString(htmlContent, 'text/html');
      
      // Find all text-containing elements (p, h1-h6, li, td)
      const elements = Array.from(doc.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li, td'));
      const totalElements = elements.length;
      
      for (let i = 0; i < totalElements; i++) {
        const el = elements[i];
        const originalText = el.textContent?.trim();
        
        if (originalText && originalText.length > 0) {
          // Update progress
          setProgress(Math.round(((i + 1) / totalElements) * 100));

          // Check for images within this element
          const images = Array.from(el.querySelectorAll('img'));
          const hasImages = images.length > 0;

          // Batch translate all selected languages at once
          const translations = await translateBatch(originalText, selectedLanguages);

          if (hasImages) {
            // Remove images from the original element to move them after the translation
            images.forEach(img => img.remove());
          }

          // Create a container for translations
          const translationContainer = doc.createElement('div');
          translationContainer.className = 'translation-container';
          translationContainer.style.marginTop = '4px';
          translationContainer.style.marginBottom = '12px';

          for (const lang of selectedLanguages) {
            const translatedText = translations[lang] || '(翻譯失敗)';
            const p = doc.createElement('p');
            p.style.color = '#000000';
            p.style.fontStyle = 'normal';
            p.style.fontSize = '10pt';
            p.style.margin = '2px 0';
            p.textContent = `[${lang}] ${translatedText}`;
            translationContainer.appendChild(p);
          }
          
          // Insert translations after the current element
          el.insertAdjacentElement('afterend', translationContainer);

          if (hasImages) {
            // Create a new paragraph for the images and insert it after the translation
            // This ensures the layout is: Original Text -> Translation -> Images
            const imgPara = doc.createElement('p');
            images.forEach(img => imgPara.appendChild(img));
            translationContainer.insertAdjacentElement('afterend', imgPara);
          }
        }
      }

      // 3. Generate new Docx from modified HTML structure
      setStatus('generating');
      const docChildren: Paragraph[] = [];

      // Helper to convert base64 to Uint8Array for docx ImageRun
      const base64ToUint8Array = (base64: string) => {
        const binaryString = window.atob(base64.split(',')[1]);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
        return bytes;
      };

      // Recursive function to process nodes within a paragraph or heading
      const processNode = (node: Node, isTranslation: boolean, parentStyles: any = {}): any[] => {
        const children: any[] = [];
        
        for (const child of Array.from(node.childNodes)) {
          if (child.nodeName === 'IMG') {
            const img = child as HTMLImageElement;
            if (img.src.startsWith('data:image')) {
              try {
                const imageData = base64ToUint8Array(img.src);
                if (imageData && imageData.length > 0) {
                  children.push(new ImageRun({
                    data: imageData,
                    transformation: {
                      width: 400,
                      height: 300,
                    },
                  }));
                }
              } catch (e) {
                console.error("Failed to process image", e);
              }
            }
          } else if (child.nodeType === Node.TEXT_NODE) {
            const text = child.textContent || '';
            if (text.trim() || text === ' ') {
              children.push(new TextRun({
                text: text,
                bold: parentStyles.bold || false,
                italics: (isTranslation ? false : parentStyles.italics) || false,
                color: isTranslation ? "000000" : (parentStyles.color || "000000"),
                size: parentStyles.size || undefined,
              }));
            }
          } else if (child.nodeType === Node.ELEMENT_NODE) {
            const el = child as HTMLElement;
            const newStyles = { ...parentStyles };
            
            if (el.nodeName === 'STRONG' || el.nodeName === 'B') newStyles.bold = true;
            if (el.nodeName === 'EM' || el.nodeName === 'I') newStyles.italics = true;
            
            // Recursively process children of this element
            children.push(...processNode(el, isTranslation, newStyles));
          }
        }
        return children;
      };

      // Iterate through the processed HTML body and convert to docx objects
      const processTopLevelNode = (node: Node): any[] => {
        const nodeName = node.nodeName;
        const el = node as HTMLElement;

        if (nodeName === 'P' || nodeName.match(/^H[1-6]$/)) {
          // Determine heading level
          let headingLevel: any = undefined;
          if (nodeName === 'H1') headingLevel = "Heading1";
          else if (nodeName === 'H2') headingLevel = "Heading2";
          else if (nodeName === 'H3') headingLevel = "Heading3";

          const paragraphChildren = processNode(el, false, {
            bold: nodeName.match(/^H[1-6]$/) ? true : false,
            size: nodeName === 'H1' ? 32 : nodeName === 'H2' ? 28 : nodeName === 'H3' ? 24 : undefined
          });

          if (paragraphChildren.length > 0) {
            return [new Paragraph({
              children: paragraphChildren,
              heading: headingLevel,
              spacing: { before: 120, after: 120 },
            })];
          }
        } else if (nodeName === 'DIV') {
          const isTranslation = el.className === 'translation-container';
          const divChildren: any[] = [];
          
          for (const childNode of Array.from(el.childNodes)) {
            if (isTranslation && childNode.nodeName === 'P') {
              const pEl = childNode as HTMLElement;
              const translationChildren = processNode(pEl, true, { italics: false, color: "000000" });
              if (translationChildren.length > 0) {
                divChildren.push(new Paragraph({
                  children: translationChildren,
                  spacing: { after: 200 },
                  indent: { left: 240 },
                }));
              }
            } else {
              divChildren.push(...processTopLevelNode(childNode));
            }
          }
          return divChildren;
        } else if (nodeName === 'TABLE') {
          const rows: TableRow[] = [];
          const trs = Array.from(el.querySelectorAll('tr'));
          
          for (const tr of trs) {
            const cells: TableCell[] = [];
            const tds = Array.from(tr.childNodes).filter(n => n.nodeName === 'TD' || n.nodeName === 'TH') as HTMLElement[];
            
            for (const td of tds) {
              const cellChildren: any[] = [];
              for (const childNode of Array.from(td.childNodes)) {
                cellChildren.push(...processTopLevelNode(childNode));
              }
              
              cells.push(new TableCell({
                children: cellChildren.length > 0 ? cellChildren : [new Paragraph("")],
                width: { size: 0, type: WidthType.AUTO }, // Use AUTO for better stability
              }));
            }
            
            if (cells.length > 0) {
              rows.push(new TableRow({ children: cells }));
            }
          }
          
          if (rows.length > 0) {
            return [new Table({
              rows: rows,
              width: { size: 0, type: WidthType.AUTO }, // Use AUTO for the whole table too
            })];
          }
        } else if (nodeName === 'UL' || nodeName === 'OL') {
          const listParagraphs: any[] = [];
          for (const childNode of Array.from(el.childNodes)) {
            if (childNode.nodeName === 'LI') {
              const liEl = childNode as HTMLElement;
              const liChildren = processNode(liEl, false);
              if (liChildren.length > 0) {
                listParagraphs.push(new Paragraph({
                  children: liChildren,
                  bullet: { level: 0 },
                }));
              }
            } else {
              // Handle translations or other nodes inside the list
              listParagraphs.push(...processTopLevelNode(childNode));
            }
          }
          return listParagraphs;
        } else if (nodeName === '#text') {
          const text = node.textContent?.trim();
          if (text) {
            return [new Paragraph({ children: [new TextRun(text)] })];
          }
        } else {
          // Fallback for unknown elements: try to process their children
          const fallbackChildren: any[] = [];
          for (const childNode of Array.from(node.childNodes)) {
            fallbackChildren.push(...processTopLevelNode(childNode));
          }
          return fallbackChildren;
        }
        return [];
      };

      for (const node of Array.from(doc.body.childNodes)) {
        const nodes = processTopLevelNode(node);
        if (nodes && nodes.length > 0) {
          docChildren.push(...nodes);
        }
      }

      // Ensure document is not empty
      if (docChildren.length === 0) {
        docChildren.push(new Paragraph("文件內容處理完成"));
      }

      const docxFile = new Document({
        sections: [{
          properties: {},
          children: docChildren,
        }],
      });

      const docxBlob = await Packer.toBlob(docxFile);
      saveAs(docxBlob, `translated_${file.name}`);
      
      setStatus('completed');
    } catch (err) {
      console.error(err);
      setError(err instanceof Error ? err.message : '處理檔案時發生錯誤');
      setStatus('error');
    }
  };

  return (
    <div className="min-h-screen bg-[#F5F5F5] font-sans text-[#1A1A1A] p-4 md:p-8">
      <div className="max-w-3xl mx-auto">
        {/* Header */}
        <header className="mb-12 text-center">
          <motion.div 
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            className="inline-flex items-center justify-center w-16 h-16 bg-white rounded-2xl shadow-sm mb-6"
          >
            <Languages className="w-8 h-8 text-emerald-600" />
          </motion.div>
          <motion.h1 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.1 }}
            className="text-4xl font-light tracking-tight mb-3"
          >
            Word 文書多語翻譯器
          </motion.h1>
          <motion.p 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.2 }}
            className="text-muted text-sm uppercase tracking-widest opacity-60"
          >
            Multi-Language Document Translator
          </motion.p>
        </header>

        {/* Main Card */}
        <motion.div 
          initial={{ opacity: 0, scale: 0.98 }}
          animate={{ opacity: 1, scale: 1 }}
          transition={{ delay: 0.3 }}
          className="bg-white rounded-[32px] shadow-sm border border-black/5 overflow-hidden"
        >
          <div className="p-8 md:p-12">
            {/* Upload Section */}
            <div 
              onClick={() => fileInputRef.current?.click()}
              className={`
                relative group cursor-pointer border-2 border-dashed rounded-2xl p-12 transition-all duration-300
                ${file ? 'border-emerald-200 bg-emerald-50/30' : 'border-gray-200 hover:border-emerald-400 hover:bg-gray-50'}
              `}
            >
              <input 
                type="file" 
                ref={fileInputRef}
                onChange={handleFileChange}
                accept=".docx"
                className="hidden"
              />
              
              <div className="flex flex-col items-center text-center">
                <div className={`
                  w-16 h-16 rounded-full flex items-center justify-center mb-4 transition-transform duration-300 group-hover:scale-110
                  ${file ? 'bg-emerald-100 text-emerald-600' : 'bg-gray-100 text-gray-400'}
                `}>
                  {file ? <FileText className="w-8 h-8" /> : <Upload className="w-8 h-8" />}
                </div>
                
                {file ? (
                  <div>
                    <p className="text-lg font-medium text-emerald-900 mb-1">{file.name}</p>
                    <p className="text-sm text-emerald-600 opacity-70">{(file.size / 1024).toFixed(1)} KB • 已就緒</p>
                  </div>
                ) : (
                  <div>
                    <p className="text-lg font-medium mb-1">點擊或拖拽上傳 Word 檔案</p>
                    <p className="text-sm text-gray-400">僅支援 .docx 格式</p>
                  </div>
                )}
              </div>
            </div>

            {/* Settings */}
            <div className="mt-8 space-y-6">
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
              
              <div className="flex items-end">
                <button
                  disabled={!file || selectedLanguages.length === 0 || status === 'processing' || status === 'translating' || status === 'generating'}
                  onClick={processFile}
                  className={`
                    w-full h-[56px] rounded-xl font-medium flex items-center justify-center gap-2 transition-all
                    ${!file || selectedLanguages.length === 0 || status === 'processing' || status === 'translating' || status === 'generating'
                      ? 'bg-gray-100 text-gray-400 cursor-not-allowed'
                      : 'bg-emerald-600 text-white hover:bg-emerald-700 shadow-lg shadow-emerald-600/20 active:scale-[0.98]'}
                  `}
                >
                  {status === 'idle' && (
                    <>
                      <span>開始翻譯 {selectedLanguages.length > 0 && `(${selectedLanguages.length} 種語言)`}</span>
                      <ArrowRight className="w-4 h-4" />
                    </>
                  )}
                  {(status === 'processing' || status === 'translating' || status === 'generating') && (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>處理中...</span>
                    </>
                  )}
                  {status === 'completed' && (
                    <>
                      <CheckCircle2 className="w-4 h-4" />
                      <span>翻譯完成</span>
                    </>
                  )}
                  {status === 'error' && (
                    <>
                      <AlertCircle className="w-4 h-4" />
                      <span>重試</span>
                    </>
                  )}
                </button>
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
                      {status === 'processing' && '正在讀取文件...'}
                      {status === 'translating' && `正在翻譯句子 (${progress}%)`}
                      {status === 'generating' && '正在生成新文件...'}
                      {status === 'completed' && '處理完成！檔案已下載'}
                    </span>
                    {status === 'translating' && (
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
          <div className="bg-gray-50 border-t border-gray-100 p-6 flex items-center justify-between">
            <div className="flex items-center gap-4">
              <div className="flex -space-x-2">
                {[1, 2, 3].map(i => (
                  <div key={i} className="w-8 h-8 rounded-full border-2 border-white bg-gray-200 overflow-hidden">
                    <img 
                      src={`https://picsum.photos/seed/user${i}/32/32`} 
                      alt="User" 
                      referrerPolicy="no-referrer"
                    />
                  </div>
                ))}
              </div>
              <span className="text-xs text-gray-400 font-medium">已有超過 1,000+ 份文件被翻譯</span>
            </div>
            <div className="flex items-center gap-2 text-xs text-gray-400">
              <CheckCircle2 className="w-3 h-3 text-emerald-500" />
              <span>由 DeepSeek AI 強力驅動</span>
            </div>
          </div>
        </motion.div>

        {/* Instructions */}
        <div className="mt-12 grid grid-cols-1 md:grid-cols-3 gap-8">
          {[
            { title: "上傳文件", desc: "支援標準 Word (.docx) 格式，系統會自動提取文字內容。" },
            { title: "智能翻譯", desc: "使用 Gemini AI 進行語境感知翻譯，確保翻譯品質。" },
            { title: "雙語對照", desc: "翻譯結果將自動插入原句下方，並生成新的對照文件。" }
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
