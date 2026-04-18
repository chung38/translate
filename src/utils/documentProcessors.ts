import React from 'react';
import type { Worksheet } from 'exceljs';

type TranslationStatus = 'idle' | 'processing' | 'translating' | 'generating' | 'completed' | 'error';

const stripTags = (text: string) => {
  if (!text) return text;
  return text.replace(/\[\/?f\d+\]/g, '').replace(/<\/?f[^>]*>/g, '');
};

const isChineseFont = (fontName: string | undefined) => {
  if (!fontName) return false;
  if (/[\u4e00-\u9fff]/.test(fontName)) return true;
  const knownChineseFonts = [
    'mingliu', 'pmingliu', 'dfkai-sb', 'simsun', 'nsimsun', 'simhei',
    'microsoft jhenghei', 'microsoft yahei', 'biaukai', 'kaiti', 'fangsong',
    'times new roman'
  ];
  const lowerName = fontName.toLowerCase();
  return knownChineseFonts.some(f => lowerName.includes(f));
};

const adjustFontForLanguage = (font: any, lang: string) => {
  const isAsianLang = ['zh-TW', 'zh-CN', 'ja', 'ko'].includes(lang);
  let newFont = font ? { ...font } : {};

  const needsStrongFontAdjustment = ['vi', 'th'].includes(lang);

  if (!isAsianLang) {
    if (needsStrongFontAdjustment || !newFont.name || isChineseFont(newFont.name)) {
      newFont.name = 'Arial';
      delete newFont.scheme;
      delete newFont.family;
      delete newFont.theme;
    }
  }

  if (!font && !newFont.name) return undefined;
  return newFont;
};

/**
 * Rewrites <w:rPr> for docx or <a:rPr> for pptx so that Vietnamese/Thai text
 * renders correctly in Word / PowerPoint.
 *
 * KEY FIX (docx):
 *   1. Strip w:hint="eastAsia" — this is what forces Word into the East-Asian
 *      font slot and overrides the explicit font name.
 *   2. Strip any w:rFonts entirely and rebuild with explicit Arial for all slots,
 *      including w:cs (complex script) which covers Vietnamese combining marks.
 *   3. Strip theme/scheme font references that bypass the name attribute.
 *   4. Add <w:lang> so spell-check / glyph selection uses the correct locale.
 */
/**
 * 為越南文/泰文翻譯 run 重建 rPr。
 *
 * 【核心策略】：對 vi/th，完全不繼承原始 rPr 的任何字體設定，
 * 只萃取「顏色、大小、粗體、斜體、底線」等視覺樣式後重新組裝，
 * 字體一律設為 Arial（含 w:cs complex-script slot）。
 *
 * 為何不能只移除 w:hint 或替換 w:rFonts：
 *  - Word 的字體選擇有多個 fallback 層（run → 段落 → 樣式 → 主題）
 *  - 只要段落/樣式層有東亞字體，run 層的修改可能被 override
 *  - 最可靠的方式是從頭建立乾淨的 rPr，不留任何東亞字體殘留
 */
const adjustXmlRPrForLanguage = (rPr: string | undefined, lang: string, docType: 'docx' | 'pptx') => {
  const isAsianLang = ['zh-TW', 'zh-CN', 'ja', 'ko'].includes(lang);
  const needsCleanFont = ['vi', 'th'].includes(lang);

  if (isAsianLang) return rPr || '';

  // 非越南/泰文：只在含中文字體時才調整
  if (!needsCleanFont) {
    if (!rPr) return '';
    const fontRegex = /(?:w:ascii|w:hAnsi|w:eastAsia|w:cs|typeface)="([^"]*)"/g;
    let hasChinese = false;
    let m;
    while ((m = fontRegex.exec(rPr)) !== null) {
      if (isChineseFont(m[1])) { hasChinese = true; break; }
    }
    if (!hasChinese) return rPr;
  }

  if (docType === 'docx') {
    const langCode = lang === 'vi' ? 'vi-VN' : lang === 'th' ? 'th-TH' : 'en-US';
    const arialFonts = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>';
    const langTag = `<w:lang w:val="${langCode}" w:eastAsia="${langCode}" w:bidi="${langCode}"/>`;

    if (!rPr) {
      return `<w:rPr>${arialFonts}${langTag}</w:rPr>`;
    }

    // 從原始 rPr 萃取與字體無關的視覺樣式標籤
    // （這些標籤需要保留：顏色、字號、粗斜體、底線等）
    const visualTags = [
      'w:rStyle',    // 樣式引用（保留，但字體會被我們的 arialFonts override）
      'w:b', 'w:bCs',
      'w:i', 'w:iCs',
      'w:caps', 'w:smallCaps',
      'w:strike', 'w:dstrike',
      'w:color',
      'w:sz', 'w:szCs',
      'w:highlight',
      'w:u',
      'w:vertAlign',
      'w:shd',
      'w:spacing',
      'w:kern',
      'w:position',
      'w:effect',
      'w:outline',
      'w:shadow',
      'w:emboss',
      'w:imprint',
      'w:noProof',
    ];

    let preserved = '';
    for (const tag of visualTags) {
      // 匹配自閉合 <w:tag .../> 或展開形式 <w:tag ...>...</w:tag>
      const re = new RegExp(`<${tag}(?:\\s[^>]*)?\\/?>(?:[\\s\\S]*?<\\/${tag}>)?`, 'g');
      const found = rPr.match(re);
      if (found) {
        for (const f of found) {
          // 過濾 w:b w:val="0" 等「關閉」屬性（這些會取消粗體）
          if (/^<w:[bi]C?s?[\s>]/.test(f) && /w:val="0"/.test(f)) continue;
          preserved += f;
        }
      }
    }

    // 重建：Arial 字體放最前面確保最高優先級，視覺樣式次之，lang 放最後
    return `<w:rPr>${arialFonts}${preserved}${langTag}</w:rPr>`;

  } else if (docType === 'pptx') {
    const langCode = lang === 'vi' ? 'vi-VN' : lang === 'th' ? 'th-TH' : 'en-US';
    const arialLatin = '<a:latin typeface="Arial"/>';
    const arialEa = '<a:ea typeface="Arial"/>';
    const arialCs = '<a:cs typeface="Arial"/>';

    if (!rPr) {
      return `<a:rPr lang="${langCode}" dirty="0">${arialLatin}${arialEa}${arialCs}</a:rPr>`;
    }

    let newRPr = rPr;
    // 移除舊字體標籤
    newRPr = newRPr.replace(/<a:latin[\s\S]*?(?:\/>|<\/a:latin>)/g, '');
    newRPr = newRPr.replace(/<a:ea[\s\S]*?(?:\/>|<\/a:ea>)/g, '');
    newRPr = newRPr.replace(/<a:cs[\s\S]*?(?:\/>|<\/a:cs>)/g, '');
    // 更新 lang 屬性
    if (newRPr.includes(' lang="')) {
      newRPr = newRPr.replace(/ lang="[^"]*"/, ` lang="${langCode}"`);
    } else {
      newRPr = newRPr.replace('<a:rPr', `<a:rPr lang="${langCode}"`);
    }
    // 確保展開形式
    newRPr = newRPr.replace(/(<a:rPr[^>]*)\/>/, '$1></a:rPr>');
    // 插入 Arial 字體
    if (newRPr.includes('</a:rPr>')) {
      newRPr = newRPr.replace('</a:rPr>', `${arialLatin}${arialEa}${arialCs}</a:rPr>`);
    } else {
      newRPr = `<a:rPr lang="${langCode}">${arialLatin}${arialEa}${arialCs}</a:rPr>`;
    }
    return newRPr;
  }

  return rPr || '';
};

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

  const unescapeXml = (text: string) =>
    text.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  const escapeXml = (text: string) =>
    text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

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
                  const lastRun = runs.length > 0 ? runs[runs.length - 1] : null;
                  if (lastRun && lastRun.rPr === rPr) {
                    lastRun.text += runText;
                  } else {
                    runs.push({ rPr, text: runText });
                  }
                }
              }
            });

            runs.forEach((run, idx) => {
              taggedText += `[f${idx}]${run.text}[/f${idx}]`;
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
    for (const res of results) translatedResults.push(...res);

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
              let longestRun = currentItem.runs[0];
              for (const run of currentItem.runs) {
                if (run.text.length > longestRun.text.length) longestRun = run;
              }
              const defaultRPr = adjustXmlRPrForLanguage(longestRun.rPr, lang, 'docx');

              const fRegex = /\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g;
              let match;
              let lastIndex = 0;
              let hasTags = false;
              let prevRunEndedWithoutSpace = false;

              while ((match = fRegex.exec(translatedText)) !== null) {
                hasTags = true;
                const id = parseInt(match[1], 10);
                let text = match[2];

                const originalRPr = currentItem.runs[id] ? currentItem.runs[id].rPr : longestRun.rPr;
                const rPr = adjustXmlRPrForLanguage(originalRPr, lang, 'docx');

                // Handle text between tags (e.g. spaces the AI placed outside [f] markers)
                if (match.index > lastIndex) {
                  const betweenText = stripTags(translatedText.substring(lastIndex, match.index));
                  if (betweenText) {
                    appendedRuns += `<w:r>${defaultRPr}<w:t xml:space="preserve">${escapeXml(betweenText)}</w:t></w:r>`;
                    prevRunEndedWithoutSpace = false;
                  }
                }

                if (text) {
                  let finalText = stripTags(text);

                  // FIX: ensure word boundary space between adjacent translated runs.
                  // If previous run didn't end with a space AND this run doesn't start
                  // with a space AND it's not the very first run, prepend a space so
                  // words don't merge (e.g. "thếgiới" → "thế giới").
                  if (prevRunEndedWithoutSpace && finalText.length > 0 && !finalText.startsWith(' ')) {
                    finalText = ' ' + finalText;
                  }

                  // Preserve trailing space from original run
                  const originalText = currentItem.runs[id]?.text || '';
                  if (originalText.endsWith(' ') && !finalText.endsWith(' ')) {
                    finalText += ' ';
                  }
                  if (originalText.startsWith(' ') && !finalText.startsWith(' ') && lastIndex === 0) {
                    finalText = ' ' + finalText;
                  }

                  appendedRuns += `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(finalText)}</w:t></w:r>`;
                  prevRunEndedWithoutSpace = !finalText.endsWith(' ');
                } else {
                  // Empty tag — don't reset the space tracking
                }

                lastIndex = fRegex.lastIndex;
              }

              if (!hasTags) {
                const cleanText = stripTags(translatedText);
                appendedRuns += `<w:r>${defaultRPr}<w:t xml:space="preserve">${escapeXml(cleanText)}</w:t></w:r>`;
              } else if (lastIndex < translatedText.length) {
                const remainingText = stripTags(translatedText.substring(lastIndex));
                if (remainingText) {
                  appendedRuns += `<w:r>${defaultRPr}<w:t xml:space="preserve">${escapeXml(remainingText)}</w:t></w:r>`;
                }
              }
            } else {
              const fallbackRPr = adjustXmlRPrForLanguage('', lang, 'docx');
              appendedRuns += `<w:r>${fallbackRPr}<w:t xml:space="preserve">${escapeXml(stripTags(translatedText))}</w:t></w:r>`;
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
        if (cell.type === 1) return;

        if (cell.value) {
          if (typeof cell.value === 'string' && cell.value.trim().length > 0) {
            textsToTranslate.push({ row: rowNumber, col: colNumber, text: cell.value, type: 'string' });
          } else if (typeof cell.value === 'object' && (cell.value as any).richText) {
            const richTextArr = (cell.value as any).richText;
            let taggedText = '';
            richTextArr.forEach((rt: any, idx: number) => {
              if (rt.text) taggedText += `[f${idx}]${rt.text}[/f${idx}]`;
            });
            if (taggedText.trim().length > 0) {
              textsToTranslate.push({ row: rowNumber, col: colNumber, text: taggedText, type: 'richText', richText: richTextArr });
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
      for (const res of results) translatedResults.push(...res);

      const progressIndex = Math.min(i + batchSize * concurrency, textsToTranslate.length);
      const sheetProgress = (progressIndex / textsToTranslate.length) * (60 / totalWorksheets);
      updateProgress(30 + (processedWorksheets * (60 / totalWorksheets)) + sheetProgress);
    }

    updateProgress(30 + ((processedWorksheets + 1) / totalWorksheets) * 60, 'generating');

    textsToTranslate.forEach((item, index) => {
      allTranslations.push({
        sheet: worksheet.name,
        row: item.row,
        col: item.col,
        original: item.text,
        translations: translatedResults[index] || {}
      });
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
          if (cell.type === 1) return;

          const translationItem = allTranslations.find(t => t.sheet === worksheet.name && t.row === rowNumber && t.col === colNumber);
          if (translationItem) {
            const newRichText: any[] = [];

            let defaultFont = {};
            if (typeof cell.value === 'object' && (cell.value as any).richText && (cell.value as any).richText.length > 0) {
              defaultFont = (cell.value as any).richText[0].font || {};
            } else if (cell.font) {
              defaultFont = cell.font;
            }

            if (typeof cell.value === 'object' && (cell.value as any).richText) {
              const originalRichText = (cell.value as any).richText;
              originalRichText.forEach((rt: any) => {
                newRichText.push({ text: rt.text, font: rt.font || defaultFont });
              });
            } else {
              newRichText.push({ text: String(cell.value), font: defaultFont });
            }

            langs.forEach(lang => {
              const translatedText = translationItem.translations[lang] || '(翻譯失敗)';
              const adjustedDefaultFont = adjustFontForLanguage(defaultFont, lang);

              if (translationItem.original.includes('[f0]')) {
                newRichText.push({ text: '\n', font: adjustedDefaultFont });

                const fRegex = /\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g;
                let match;
                let lastIndex = 0;
                let hasTags = false;

                while ((match = fRegex.exec(translatedText)) !== null) {
                  hasTags = true;
                  const id = parseInt(match[1], 10);
                  const text = match[2];

                  let originalFont = defaultFont;
                  if (typeof cell.value === 'object' && (cell.value as any).richText) {
                    originalFont = (cell.value as any).richText[id]?.font || defaultFont;
                  }
                  const adjustedOriginalFont = adjustFontForLanguage(originalFont, lang);

                  if (match.index > lastIndex) {
                    const betweenText = stripTags(translatedText.substring(lastIndex, match.index));
                    if (betweenText) newRichText.push({ text: betweenText, font: adjustedDefaultFont });
                  }

                  if (text) {
                    newRichText.push({ text: stripTags(text), font: adjustedOriginalFont });
                  }
                  lastIndex = fRegex.lastIndex;
                }

                if (!hasTags) {
                  newRichText.push({ text: stripTags(translatedText), font: adjustedDefaultFont });
                } else if (lastIndex < translatedText.length) {
                  const remainingText = stripTags(translatedText.substring(lastIndex));
                  if (remainingText) newRichText.push({ text: remainingText, font: adjustedDefaultFont });
                }
              } else {
                newRichText.push({ text: '\n' + stripTags(translatedText), font: adjustedDefaultFont });
              }
            });

            const cleanedRichText = newRichText.map(rt => {
              if (rt.font === undefined) return { text: rt.text };
              return rt;
            });
            cell.value = { richText: cleanedRichText };
          }
          cell.alignment = { ...(cell.alignment || {}), wrapText: true };
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
  let tesseractWorker: any = null;

  for (let i = 1; i <= totalPages; i++) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    const page = await pdf.getPage(i);

    let pageText = '';
    updateProgress(10 + ((i - 0.5) / totalPages) * 10, 'processing');

    const textContent = await page.getTextContent();
    const textItems = textContent.items as any[];

    let extractedText = '';
    let lastY = null;

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

    if (extractedText.length > 20) {
      pageText = extractedText;
    } else {
      console.log(`Page ${i} has little/no text, falling back to OCR...`);
      const viewport = page.getViewport({ scale: 2.0 });
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d');
      if (context) {
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        await page.render({ canvasContext: context, viewport: viewport, canvas: canvas }).promise;

        if (!tesseractWorker) {
          const { createWorker } = await import('tesseract.js');
          tesseractWorker = await createWorker('chi_tra+eng');
          await tesseractWorker.setParameters({ tessedit_pageseg_mode: '11' });
        }

        const { data: { text } } = await tesseractWorker.recognize(canvas);
        pageText = text;
      }
    }

    fullText += pageText + '\n\n';
    updateProgress(10 + (i / totalPages) * 10);
  }

  if (tesseractWorker) await tesseractWorker.terminate();

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
    for (const res of results) translatedParagraphs.push(...res);

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
      page.style.minHeight = '1131px';
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

      if (currentPage.scrollHeight > 1131 && currentPage.children.length > 1) {
        currentPage.removeChild(block);
        currentPage = createPage();
        currentPage.appendChild(block);
      }
    });

    try {
      await document.fonts.ready;
      await new Promise(resolve => setTimeout(resolve, 500));

      const doc = new jsPDF('p', 'pt', 'a4');
      const pdfWidth = doc.internal.pageSize.getWidth();

      for (let i = 0; i < pages.length; i++) {
        const canvas = await html2canvas(pages[i], {
          scale: 2,
          useCORS: true,
          logging: false,
          onclone: (clonedDoc) => {
            const styles = clonedDoc.querySelectorAll('style, link[rel="stylesheet"]');
            styles.forEach(s => s.remove());
          }
        });
        const imgData = canvas.toDataURL('image/jpeg', 0.95);

        if (i > 0) doc.addPage();

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

  const slideFiles = Object.keys(loadedZip.files).filter(name =>
    name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
  );

  const textsToTranslate: { slide: string, id: string, text: string, pBlock: string, runs?: { rPr: string, text: string }[] }[] = [];
  const slideContents: Record<string, string> = {};

  const unescapeXml = (text: string) =>
    text.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  const escapeXml = (text: string) =>
    text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

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

            rMatches.forEach((rBlock) => {
              const tMatch = rBlock.match(/<a:t>([\s\S]*?)<\/a:t>/);
              if (tMatch && tMatch[1]) {
                const text = unescapeXml(tMatch[1]);
                const rPrMatch = rBlock.match(/<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/);
                const rPr = rPrMatch ? rPrMatch[0] : '';

                if (text) {
                  const lastRun = runs.length > 0 ? runs[runs.length - 1] : null;
                  if (lastRun && lastRun.rPr === rPr) {
                    lastRun.text += text;
                  } else {
                    runs.push({ rPr, text });
                  }
                }
              }
            });

            runs.forEach((run, idx) => {
              taggedText += `[f${idx}]${run.text}[/f${idx}]`;
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
    for (const res of results) translatedResults.push(...res);

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
              let longestRun = currentItem.runs[0];
              for (const run of currentItem.runs) {
                if (run.text.length > longestRun.text.length) longestRun = run;
              }
              const defaultRPr = adjustXmlRPrForLanguage(longestRun.rPr, lang, 'pptx');

              const fRegex = /\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g;
              let match;
              let lastIndex = 0;
              let hasTags = false;
              let prevRunEndedWithoutSpace = false;

              while ((match = fRegex.exec(translatedText)) !== null) {
                hasTags = true;
                const id = parseInt(match[1], 10);
                let text = match[2];

                const originalRPr = currentItem.runs[id] ? currentItem.runs[id].rPr : longestRun.rPr;
                const rPr = adjustXmlRPrForLanguage(originalRPr, lang, 'pptx');

                if (match.index > lastIndex) {
                  const betweenText = stripTags(translatedText.substring(lastIndex, match.index));
                  if (betweenText) {
                    appendedRuns += `<a:r>${defaultRPr}<a:t>${escapeXml(betweenText)}</a:t></a:r>`;
                    prevRunEndedWithoutSpace = false;
                  }
                }

                if (text) {
                  let finalText = stripTags(text);

                  if (prevRunEndedWithoutSpace && finalText.length > 0 && !finalText.startsWith(' ')) {
                    finalText = ' ' + finalText;
                  }

                  const originalText = currentItem.runs[id]?.text || '';
                  if (originalText.endsWith(' ') && !finalText.endsWith(' ')) finalText += ' ';
                  if (originalText.startsWith(' ') && !finalText.startsWith(' ') && lastIndex === 0) finalText = ' ' + finalText;

                  appendedRuns += `<a:r>${rPr}<a:t>${escapeXml(finalText)}</a:t></a:r>`;
                  prevRunEndedWithoutSpace = !finalText.endsWith(' ');
                }

                lastIndex = fRegex.lastIndex;
              }

              if (!hasTags) {
                const cleanText = stripTags(translatedText);
                appendedRuns += `<a:r>${defaultRPr}<a:t>${escapeXml(cleanText)}</a:t></a:r>`;
              } else if (lastIndex < translatedText.length) {
                const remainingText = stripTags(translatedText.substring(lastIndex));
                if (remainingText) appendedRuns += `<a:r>${defaultRPr}<a:t>${escapeXml(remainingText)}</a:t></a:r>`;
              }
            } else {
              const fallbackRPr = adjustXmlRPrForLanguage('', lang, 'pptx');
              appendedRuns += `<a:r>${fallbackRPr}<a:t>${escapeXml(stripTags(translatedText))}</a:t></a:r>`;
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
