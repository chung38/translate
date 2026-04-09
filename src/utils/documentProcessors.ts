import React from 'react';
import type { Worksheet } from 'exceljs';

type TranslationStatus = 'idle' | 'processing' | 'translating' | 'generating' | 'completed' | 'error';

export const processDocx = async (
  file: File, 
  targetLanguages: string[],
  industry: string,
  translateBatch: (texts: string[], targetLangs: string[], industry: string) => Promise<Record<string, string>[]>,
  updateProgress: (p: number, status?: TranslationStatus) => void,
  isCancelledRef: React.MutableRefObject<boolean>,
  outputMode: 'combined' | 'separate' = 'combined'
) => {
  updateProgress(10, 'processing');
  
  const JSZip = (await import('jszip')).default;
  const zip = new JSZip();
  const loadedZip = await zip.loadAsync(file);
  
  // Find all XML files in the word/ directory that might contain text
  const docFiles = Object.keys(loadedZip.files).filter(name => 
    name.startsWith('word/') && 
    name.endsWith('.xml') && 
    !name.includes('_rels') && 
    !name.includes('theme') && 
    !name.includes('styles') &&
    !name.includes('settings') &&
    !name.includes('webSettings') &&
    !name.includes('fontTable')
  );
  
  const unescapeXml = (text: string) => text.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  const escapeXml = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

  const textsToTranslate: { file: string, id: string, text: string, pBlock: string, runs?: { rPr: string, text: string }[] }[] = [];
  const fileContents: Record<string, string> = {};

  for (const docFile of docFiles) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    const content = await loadedZip.file(docFile)?.async('text');
    if (content) {
      fileContents[docFile] = content;
      
      const pMatches = content.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g);
      if (pMatches) {
        pMatches.forEach((pBlock, index) => {
          const rMatches = pBlock.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g);
          if (rMatches) {
            let taggedText = '';
            const runs: { rPr: string, text: string }[] = [];
            let runIndex = 0;
            
            rMatches.forEach((rBlock) => {
              const tMatches = rBlock.match(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g);
              if (tMatches) {
                const rPrMatch = rBlock.match(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/);
                const rPr = rPrMatch ? rPrMatch[0] : '';
                
                let runText = '';
                tMatches.forEach(tMatch => {
                  const innerTextMatch = tMatch.match(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/);
                  if (innerTextMatch && innerTextMatch[1]) {
                    runText += unescapeXml(innerTextMatch[1]);
                  }
                });
                
                if (runText) {
                  runs.push({ rPr, text: runText });
                  taggedText += runText;
                  runIndex++;
                }
              }
            });
            
            if (taggedText.trim().length > 0) {
              textsToTranslate.push({ file: docFile, id: `${docFile}_${index}`, text: taggedText, pBlock, runs });
            }
          }
        });
      }
    }
  }

  updateProgress(30, 'translating');

  const batchSize = 10;
  const concurrency = 3;
  const translatedResults: Record<string, string>[] = [];
  
  for (let i = 0; i < textsToTranslate.length; i += batchSize * concurrency) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    
    const promises = [];
    for (let j = 0; j < concurrency && (i + j * batchSize) < textsToTranslate.length; j++) {
      const start = i + j * batchSize;
      const batch = textsToTranslate.slice(start, start + batchSize).map(item => item.text);
      promises.push(translateBatch(batch, targetLanguages, industry));
    }
    
    const results = await Promise.all(promises);
    for (const res of results) {
      translatedResults.push(...res);
    }
    
    const progressIndex = Math.min(i + batchSize * concurrency, textsToTranslate.length);
    updateProgress(30 + (progressIndex / textsToTranslate.length) * 50);
  }

  updateProgress(80, 'generating');

  const generatedFiles: { blob: Blob, name: string }[] = [];
  const langGroups = outputMode === 'separate' ? targetLanguages.map(l => [l]) : [targetLanguages];

  for (const langs of langGroups) {
    const zip = new JSZip();
    const loadedZip = await zip.loadAsync(file);

    for (const docFile of docFiles) {
      let content = fileContents[docFile];
      const fileTexts = textsToTranslate.filter(t => t.file === docFile);
      
      let pIndex = 0;
      content = content.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (pBlock) => {
        const currentItem = fileTexts.find(t => t.id === `${docFile}_${pIndex}`);
        pIndex++;
        
        if (currentItem) {
          const globalIndex = textsToTranslate.findIndex(t => t.id === currentItem.id);
          
          let appendedRuns = '';
          langs.forEach(lang => {
            const translatedText = translatedResults[globalIndex]?.[lang] || '(翻譯失敗)';
            
            appendedRuns += `<w:r><w:br/></w:r>`;
            
            if (currentItem.runs && currentItem.runs.length > 0) {
              // Find the run with the longest text to use as the default formatting
              let longestRun = currentItem.runs[0];
              for (const run of currentItem.runs) {
                if (run.text.length > longestRun.text.length) {
                  longestRun = run;
                }
              }
              const defaultRPr = longestRun.rPr;

              appendedRuns += `<w:r>${defaultRPr}<w:t xml:space="preserve">${escapeXml(translatedText)}</w:t></w:r>`;
            } else {
              appendedRuns += `<w:r><w:t xml:space="preserve">${escapeXml(translatedText)}</w:t></w:r>`;
            }
          });
          
          return pBlock.replace(/<\/w:p>$/, appendedRuns + '</w:p>');
        }
        
        return pBlock;
      });
      
      loadedZip.file(docFile, content);
    }

    const blob = await loadedZip.generateAsync({ type: 'blob' });
    const prefix = outputMode === 'separate' ? `${langs[0]}_` : 'translated_';
    generatedFiles.push({ blob, name: `${prefix}${file.name}` });
  }

  updateProgress(100, 'completed');
  return generatedFiles;
};

export const processExcel = async (
  file: File, 
  targetLanguages: string[],
  industry: string,
  translateBatch: (texts: string[], targetLangs: string[], industry: string) => Promise<Record<string, string>[]>,
  updateProgress: (p: number, status?: TranslationStatus) => void,
  isCancelledRef: React.MutableRefObject<boolean>,
  outputMode: 'combined' | 'separate' = 'combined'
) => {
  updateProgress(10, 'processing');
  const ExcelJS = (await import('exceljs')).default;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  
  const totalWorksheets = workbook.worksheets.length;
  let processedWorksheets = 0;
  
  const allTranslations: { sheet: string, row: number, col: number, original: string, translations: Record<string, string> }[] = [];

  for (const worksheet of workbook.worksheets) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    
    const textsToTranslate: { row: number, col: number, text: string, type: 'string' | 'richText', richText?: any[] }[] = [];
    
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        if (cell.value) {
          if (typeof cell.value === 'string' && cell.value.trim().length > 0) {
            textsToTranslate.push({ row: rowNumber, col: colNumber, text: cell.value, type: 'string' });
          } else if (typeof cell.value === 'object' && (cell.value as any).richText) {
            const richTextArr = (cell.value as any).richText;
            let plainText = '';
            richTextArr.forEach((rt: any) => {
              if (rt.text) {
                plainText += rt.text;
              }
            });
            if (plainText.trim().length > 0) {
              textsToTranslate.push({ row: rowNumber, col: colNumber, text: plainText, type: 'richText', richText: richTextArr });
            }
          }
        }
      });
    });

    updateProgress(10 + (processedWorksheets / totalWorksheets) * 20, 'translating');

    const batchSize = 10;
    const concurrency = 3;
    const translatedResults: Record<string, string>[] = [];
    
    for (let i = 0; i < textsToTranslate.length; i += batchSize * concurrency) {
      if (isCancelledRef.current) throw new Error('Cancelled');
      
      const promises = [];
      for (let j = 0; j < concurrency && (i + j * batchSize) < textsToTranslate.length; j++) {
        const start = i + j * batchSize;
        const batch = textsToTranslate.slice(start, start + batchSize).map(item => item.text);
        promises.push(translateBatch(batch, targetLanguages, industry));
      }
      
      const results = await Promise.all(promises);
      for (const res of results) {
        translatedResults.push(...res);
      }
      
      const progressIndex = Math.min(i + batchSize * concurrency, textsToTranslate.length);
      const sheetProgress = (progressIndex / textsToTranslate.length) * (60 / totalWorksheets);
      updateProgress(30 + (processedWorksheets * (60 / totalWorksheets)) + sheetProgress);
    }

    updateProgress(30 + ((processedWorksheets + 1) / totalWorksheets) * 60, 'generating');

    textsToTranslate.forEach((item, index) => {
      const cell = worksheet.getCell(item.row, item.col);
      const originalText = item.text;
      const translations = translatedResults[index] || {};
      
      allTranslations.push({
        sheet: worksheet.name,
        row: item.row,
        col: item.col,
        original: originalText,
        translations
      });
      
      if (item.type === 'string') {
        let combinedText = originalText + '\n';
        targetLanguages.forEach(lang => {
          const translatedText = translations[lang] || '(翻譯失敗)';
          combinedText += `${translatedText}\n`;
        });
        cell.value = combinedText.trim();
      } else if (item.type === 'richText' && item.richText) {
        const newRichText = [...item.richText];
        
        targetLanguages.forEach(lang => {
          const translatedText = translations[lang] || '(翻譯失敗)';
          
          newRichText.push({ text: '\n', font: item.richText![0]?.font });
          
          // Find the run with the longest text to use as the default formatting
          let longestRun = item.richText![0];
          for (const run of item.richText!) {
            if (run.text && run.text.length > (longestRun.text?.length || 0)) {
              longestRun = run;
            }
          }
          const defaultFont = longestRun.font;

          newRichText.push({ text: translatedText, font: defaultFont });
        });
        cell.value = { richText: newRichText };
      }
      cell.alignment = { 
        ...(cell.alignment || {}), 
        wrapText: true 
      };
    });
    
    processedWorksheets++;
  }

  updateProgress(90, 'generating');
  
  const generatedFiles: { blob: Blob, name: string }[] = [];
  const langGroups = outputMode === 'separate' ? targetLanguages.map(l => [l]) : [targetLanguages];

  for (const langs of langGroups) {
    const ExcelJS = (await import('exceljs')).default;
    const newWorkbook = new ExcelJS.Workbook();
    await newWorkbook.xlsx.load(await file.arrayBuffer());

    for (const worksheet of newWorkbook.worksheets) {
      worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          const translationItem = allTranslations.find(t => t.sheet === worksheet.name && t.row === rowNumber && t.col === colNumber);
          if (translationItem) {
            const originalText = translationItem.original;
            const newRichText: any[] = [];
            
            let defaultFont = {};
            if (typeof cell.value === 'object' && (cell.value as any).richText && (cell.value as any).richText.length > 0) {
              defaultFont = (cell.value as any).richText[0].font || {};
            } else if (cell.font) {
              defaultFont = cell.font;
            }

            newRichText.push({ text: originalText + '\n', font: defaultFont });

            langs.forEach(lang => {
              const translatedText = translationItem.translations[lang] || '(翻譯失敗)';
              newRichText.push({ text: translatedText + '\n', font: defaultFont });
            });
            
            // Remove the last newline
            if (newRichText.length > 0 && newRichText[newRichText.length - 1].text.endsWith('\n')) {
              newRichText[newRichText.length - 1].text = newRichText[newRichText.length - 1].text.slice(0, -1);
            }
            
            cell.value = { richText: newRichText };
          }
          cell.alignment = { 
            ...(cell.alignment || {}), 
            wrapText: true 
          };
        });
      });
    }

    const buffer = await newWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const prefix = outputMode === 'separate' ? `${langs[0]}_` : 'translated_';
    generatedFiles.push({ blob, name: `${prefix}${file.name}` });
  }

  updateProgress(100, 'completed');
  return generatedFiles;
};

export const processPdf = async (
  file: File, 
  targetLanguages: string[],
  industry: string,
  translateBatch: (texts: string[], targetLangs: string[], industry: string) => Promise<Record<string, string>[]>,
  updateProgress: (p: number, status?: TranslationStatus) => void,
  isCancelledRef: React.MutableRefObject<boolean>,
  outputMode: 'combined' | 'separate' = 'combined'
) => {
  updateProgress(10, 'processing');
  const pdfjsLib = await import('pdfjs-dist');
  pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;
  
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ 
    data: arrayBuffer,
    cMapUrl: `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/cmaps/`,
    cMapPacked: true,
    standardFontDataUrl: `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/standard_fonts/`
  }).promise;
  
  let fullText = '';
  const totalPages = pdf.numPages;
  
  // Dynamic import for Tesseract to avoid blocking initial load
  let tesseractWorker: any = null;
  
  for (let i = 1; i <= totalPages; i++) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    const page = await pdf.getPage(i);
    
    let pageText = '';
    
    updateProgress(10 + ((i - 0.5) / totalPages) * 10, 'processing');
    
    // Try native text extraction first
    const textContent = await page.getTextContent();
    const textItems = textContent.items as any[];
    
    let extractedText = '';
    let lastY = null;
    
    // Sort items by Y (descending) and X (ascending) to handle basic layout
    textItems.sort((a, b) => {
      const yDiff = b.transform[5] - a.transform[5];
      if (Math.abs(yDiff) > 5) return yDiff;
      return a.transform[4] - b.transform[4];
    });

    for (const item of textItems) {
      if (!item.str) continue;
      if (lastY !== null && Math.abs(lastY - item.transform[5]) > 5) {
        extractedText += '\n';
      } else if (lastY !== null) {
        extractedText += ' ';
      }
      extractedText += item.str;
      lastY = item.transform[5];
    }
    
    extractedText = extractedText.replace(/ {2,}/g, ' ').trim();

    // If native extraction yields meaningful text, use it. Otherwise, fallback to OCR.
    if (extractedText.length > 20) {
      pageText = extractedText;
    } else {
      console.log(`Page ${i} has little/no text, falling back to OCR...`);
      const viewport = page.getViewport({ scale: 2.0 }); // Higher scale for better OCR
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d');
      if (context) {
        canvas.height = viewport.height;
        canvas.width = viewport.width;
        
        await page.render({ canvasContext: context, viewport: viewport, canvas: canvas }).promise;
        
        if (!tesseractWorker) {
          const { createWorker } = await import('tesseract.js');
          tesseractWorker = await createWorker('chi_tra+eng'); // Load Traditional Chinese and English
          await tesseractWorker.setParameters({
            tessedit_pageseg_mode: '11', // Sparse text. Find as much text as possible in no particular order. Better for tables with empty cells.
          });
        }
        
        const { data: { text } } = await tesseractWorker.recognize(canvas);
        pageText = text;
      }
    }
    
    fullText += pageText + '\n\n';
    updateProgress(10 + (i / totalPages) * 10);
  }
  
  if (tesseractWorker) {
    await tesseractWorker.terminate();
  }

  if (fullText.trim().length === 0) {
    throw new Error('無法從 PDF 中提取文字。即使嘗試了 OCR 辨識，仍無法讀取內容。');
  }

  updateProgress(20, 'translating');
  const paragraphs = fullText.split('\n\n').filter(p => p.trim().length > 0);
  
  const batchSize = 10;
  const concurrency = 3;
  const translatedParagraphs: Record<string, string>[] = [];
  
  for (let i = 0; i < paragraphs.length; i += batchSize * concurrency) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    
    const promises = [];
    for (let j = 0; j < concurrency && (i + j * batchSize) < paragraphs.length; j++) {
      const start = i + j * batchSize;
      const batch = paragraphs.slice(start, start + batchSize);
      promises.push(translateBatch(batch, targetLanguages, industry));
    }
    
    const results = await Promise.all(promises);
    for (const res of results) {
      translatedParagraphs.push(...res);
    }
    
    const progressIndex = Math.min(i + batchSize * concurrency, paragraphs.length);
    updateProgress(20 + (progressIndex / paragraphs.length) * 60);
  }

  updateProgress(80, 'generating');
  
  const { jsPDF } = await import('jspdf');
  const html2canvas = (await import('html2canvas')).default;
  
  const generatedFiles: { blob: Blob, name: string }[] = [];
  const langGroups = outputMode === 'separate' ? targetLanguages.map(l => [l]) : [targetLanguages];

  for (const langs of langGroups) {
    const wrapper = document.createElement('div');
    wrapper.style.position = 'absolute';
    wrapper.style.left = '-9999px';
    wrapper.style.top = '0';
    document.body.appendChild(wrapper);

    const pages: HTMLDivElement[] = [];
    const createPage = () => {
      const page = document.createElement('div');
      page.style.width = '800px';
      page.style.minHeight = '1131px'; // A4 height ratio for 800px width
      page.style.padding = '40px';
      page.style.boxSizing = 'border-box';
      page.style.backgroundColor = '#fff';
      page.style.fontFamily = 'system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, "Noto Sans", "Liberation Sans", sans-serif, "Apple Color Emoji", "Segoe UI Emoji", "Segoe UI Symbol", "Noto Color Emoji", "Microsoft JhengHei", "微軟正黑體", "PingFang TC", "蘋果儷中黑", "Heiti TC", "黑體-繁"';
      page.style.fontSize = '16px';
      page.style.lineHeight = '1.6';
      page.style.color = '#000';
      wrapper.appendChild(page);
      pages.push(page);
      return page;
    };

    let currentPage = createPage();
    
    paragraphs.forEach((original, index) => {
      const block = document.createElement('div');
      
      const p = document.createElement('div');
      p.style.marginBottom = '12px';
      p.innerText = original;
      block.appendChild(p);

      langs.forEach(lang => {
        const translatedText = translatedParagraphs[index]?.[lang] || '(翻譯失敗)';
        const tp = document.createElement('div');
        tp.style.marginBottom = '12px';
        tp.innerText = translatedText;
        block.appendChild(tp);
      });
      
      const spacer = document.createElement('div');
      spacer.style.height = '16px';
      block.appendChild(spacer);

      currentPage.appendChild(block);

      // Check if page exceeded height (1131px) and it's not the only block on the page
      if (currentPage.scrollHeight > 1131 && currentPage.children.length > 1) {
        // Remove block from current page
        currentPage.removeChild(block);
        // Create new page and add block
        currentPage = createPage();
        currentPage.appendChild(block);
      }
    });

    try {
      await document.fonts.ready;
      // Add a small delay to ensure the browser has painted the DOM elements
      await new Promise(resolve => setTimeout(resolve, 500));
      
      const doc = new jsPDF('p', 'pt', 'a4');
      const pdfWidth = doc.internal.pageSize.getWidth();
      
      for (let i = 0; i < pages.length; i++) {
        const canvas = await html2canvas(pages[i], { 
          scale: 2, 
          useCORS: true, 
          logging: false,
          onclone: (clonedDoc) => {
            // Remove all stylesheets in the cloned document to prevent html2canvas 
            // from parsing unsupported CSS functions like "oklch" from Tailwind v4.
            // Since our PDF pages use inline styles exclusively, this won't affect the output.
            const styles = clonedDoc.querySelectorAll('style, link[rel="stylesheet"]');
            styles.forEach(s => s.remove());
          }
        });
        const imgData = canvas.toDataURL('image/jpeg', 0.95);
        
        if (i > 0) {
          doc.addPage();
        }
        
        const imgWidth = pdfWidth;
        const imgHeight = (canvas.height * imgWidth) / canvas.width;
        
        doc.addImage(imgData, 'JPEG', 0, 0, imgWidth, imgHeight);
      }

      const blob = doc.output('blob');
      const prefix = outputMode === 'separate' ? `${langs[0]}_` : 'translated_';
      generatedFiles.push({ blob, name: `${prefix}${file.name.replace('.pdf', '.pdf')}` });
    } finally {
      document.body.removeChild(wrapper);
    }
  }

  updateProgress(100, 'completed');
  return generatedFiles;
};

export const processPptx = async (
  file: File, 
  targetLanguages: string[],
  industry: string,
  translateBatch: (texts: string[], targetLangs: string[], industry: string) => Promise<Record<string, string>[]>,
  updateProgress: (p: number, status?: TranslationStatus) => void,
  isCancelledRef: React.MutableRefObject<boolean>,
  outputMode: 'combined' | 'separate' = 'combined'
) => {
  updateProgress(10, 'processing');
  
  const JSZip = (await import('jszip')).default;
  const zip = new JSZip();
  const loadedZip = await zip.loadAsync(file);
  
  const slideFiles = Object.keys(loadedZip.files).filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml'));
  
  const textsToTranslate: { slide: string, id: string, text: string, pBlock: string, runs?: { rPr: string, text: string }[] }[] = [];
  const slideContents: Record<string, string> = {};

  const unescapeXml = (text: string) => text.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  const escapeXml = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

  for (const slideFile of slideFiles) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    const content = await loadedZip.file(slideFile)?.async('text');
    if (content) {
      slideContents[slideFile] = content;
      
      const pMatches = content.match(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/g);
      if (pMatches) {
        pMatches.forEach((pBlock, index) => {
          const rMatches = pBlock.match(/<a:r\b[^>]*>[\s\S]*?<\/a:r>/g);
          if (rMatches) {
            let taggedText = '';
            const runs: { rPr: string, text: string }[] = [];
            let runIndex = 0;
            
            rMatches.forEach((rBlock) => {
              const tMatch = rBlock.match(/<a:t>([\s\S]*?)<\/a:t>/);
              if (tMatch && tMatch[1]) {
                const text = unescapeXml(tMatch[1]);
                const rPrMatch = rBlock.match(/<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/);
                const rPr = rPrMatch ? rPrMatch[0] : '';
                
                if (text) {
                  runs.push({ rPr, text });
                  taggedText += text;
                  runIndex++;
                }
              }
            });
            
            if (taggedText.trim().length > 0) {
              textsToTranslate.push({ slide: slideFile, id: `${slideFile}_${index}`, text: taggedText, pBlock, runs });
            }
          }
        });
      }
    }
  }

  updateProgress(30, 'translating');

  const batchSize = 10;
  const concurrency = 3;
  const translatedResults: Record<string, string>[] = [];
  
  for (let i = 0; i < textsToTranslate.length; i += batchSize * concurrency) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    
    const promises = [];
    for (let j = 0; j < concurrency && (i + j * batchSize) < textsToTranslate.length; j++) {
      const start = i + j * batchSize;
      const batch = textsToTranslate.slice(start, start + batchSize).map(item => item.text);
      promises.push(translateBatch(batch, targetLanguages, industry));
    }
    
    const results = await Promise.all(promises);
    for (const res of results) {
      translatedResults.push(...res);
    }
    
    const progressIndex = Math.min(i + batchSize * concurrency, textsToTranslate.length);
    updateProgress(30 + (progressIndex / textsToTranslate.length) * 50);
  }

  updateProgress(80, 'generating');

  const generatedFiles: { blob: Blob, name: string }[] = [];
  const langGroups = outputMode === 'separate' ? targetLanguages.map(l => [l]) : [targetLanguages];

  for (const langs of langGroups) {
    const zip = new JSZip();
    const loadedZip = await zip.loadAsync(file);

    for (const slideFile of slideFiles) {
      let content = slideContents[slideFile];
      const slideTexts = textsToTranslate.filter(t => t.slide === slideFile);
      
      let pIndex = 0;
      content = content.replace(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/g, (pBlock) => {
        const currentItem = slideTexts.find(t => t.id === `${slideFile}_${pIndex}`);
        pIndex++;
        
        if (currentItem) {
          const globalIndex = textsToTranslate.findIndex(t => t.id === currentItem.id);
          
          let appendedRuns = '';
          langs.forEach(lang => {
            const translatedText = translatedResults[globalIndex]?.[lang] || '(翻譯失敗)';
            
            appendedRuns += `<a:br/>`;
            
            if (currentItem.runs && currentItem.runs.length > 0) {
              // Find the run with the longest text to use as the default formatting
              let longestRun = currentItem.runs[0];
              for (const run of currentItem.runs) {
                if (run.text.length > longestRun.text.length) {
                  longestRun = run;
                }
              }
              const defaultRPr = longestRun.rPr;

              appendedRuns += `<a:r>${defaultRPr}<a:t>${escapeXml(translatedText)}</a:t></a:r>`;
            } else {
               appendedRuns += `<a:r><a:t>${escapeXml(translatedText)}</a:t></a:r>`;
            }
          });
          
          return pBlock.replace(/<\/a:p>$/, appendedRuns + '</a:p>');
        }
        
        return pBlock;
      });
      
      loadedZip.file(slideFile, content);
    }

    const blob = await loadedZip.generateAsync({ type: 'blob' });
    const prefix = outputMode === 'separate' ? `${langs[0]}_` : 'translated_';
    generatedFiles.push({ blob, name: `${prefix}${file.name}` });
  }

  updateProgress(100, 'completed');
  return generatedFiles;
};
