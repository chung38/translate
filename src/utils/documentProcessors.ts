import React from 'react';
import type { Worksheet } from 'exceljs';

type TranslationStatus = 'idle' | 'processing' | 'translating' | 'generating' | 'completed' | 'error';

const stripTags = (text: string) => {
  if (!text) return text;
  return text.replace(/\[\/?\f\d+\]/g, '').replace(/<\/?f[^>]*>/g, '');
};

const isChineseFont = (fontName: string | undefined) => {
  if (!fontName) return false;
  if (/[\u4e00-\u9fff]/.test(fontName)) return true;
  const knownChineseFonts = [
    'mingliu', 'pmingliu', 'dfkai-sb', 'simsun', 'nsimsun', 'simhei', 
    'microsoft jhenghei', 'microsoft yahei', 'biaukai', 'kaiti', 'fangsong',
    'times new roman' // Sometimes Times New Roman is used as a default for Chinese in older documents but doesn't render Vietnamese well
  ];
  const lowerName = fontName.toLowerCase();
  return knownChineseFonts.some(f => lowerName.includes(f));
};

const adjustFontForLanguage = (font: any, lang: string) => {
  const isAsianLang = ['zh-TW', 'zh-CN', 'ja', 'ko'].includes(lang);
  let newFont = font ? { ...font } : {};
  
  // For Vietnamese and Thai, we strongly prefer Arial because many default fonts 
  // (even non-Chinese ones) have poor support for their specific diacritics in Excel.
  const needsStrongFontAdjustment = ['vi', 'th'].includes(lang);
  
  if (!isAsianLang) {
    if (needsStrongFontAdjustment || !newFont.name || isChineseFont(newFont.name)) {
      newFont.name = 'Arial';
      // CRITICAL: If a font has a 'scheme' (like 'minor' or 'major') or 'theme', Excel will ignore the 'name'
      // and use the theme's default font instead. We must delete it to force Arial.
      delete newFont.scheme;
      delete newFont.family;
      delete newFont.theme;
    }
  }
  
  // If we didn't change anything and it was originally undefined, return undefined
  if (!font && !newFont.name) return undefined;
  
  return newFont;
};

const normalizeDocxRPr = (rPr: string) => {
  if (!rPr) return '';
  return rPr.replace(/<w:rFonts[^>]*>/g, '').replace(/<w:lang[^>]*>/g, '').replace(/<w:hint[^>]*>/g, '');
};

const normalizePptxRPr = (rPr: string) => {
  if (!rPr) return '';
  return rPr.replace(/<a:latin\b[^>]*>/g, '').replace(/<a:ea\b[^>]*>/g, '').replace(/<a:cs\b[^>]*>/g, '').replace(/ lang="[^"]*"/g, '').replace(/ altLang="[^"]*"/g, '');
};

const normalizeExcelFont = (font: any) => {
  if (!font) return '';
  const { name, family, scheme, charset, ...rest } = font;
  return JSON.stringify(rest);
};

const adjustXmlRPrForLanguage = (rPr: string | undefined, lang: string, docType: 'docx' | 'pptx') => {
  const isAsianLang = ['zh-TW', 'zh-CN', 'ja', 'ko'].includes(lang);
  const needsStrongFontAdjustment = ['vi', 'th'].includes(lang);
  
  if (isAsianLang) return rPr || '';
  
  let newRPr = rPr || '';
  let shouldAdjust = needsStrongFontAdjustment;
  
  if (!shouldAdjust && newRPr) {
    // Check if it contains Chinese fonts
    const fontRegex = /(?:w:ascii|w:hAnsi|w:eastAsia|w:cs|typeface)="([^"]*)"/g;
    let match;
    while ((match = fontRegex.exec(newRPr)) !== null) {
      if (isChineseFont(match[1])) {
        shouldAdjust = true;
        break;
      }
    }
  }
  
  if (shouldAdjust) {
    // We disguise Vietnamese as en-US because Word has a long-standing layout engine bug:
    // When Vietnamese diacritics are embedded in a document with a Chinese root theme,
    // Word misclassifies them as East Asian characters. This ignores our Arial request, forces MingLiU,
    // injects wide typography spacing between letters, and breaks words in half at the end of lines.
    // By tagging it as en-US, we force strict Latin text engine rules, perfect Arial, and correct word wrapping.
    const langCode = lang === 'th' ? 'th-TH' : 'en-US';
    
    if (docType === 'docx') {
      const safeRegexMatch = (str: string, regex: RegExp) => (str.match(regex) || [])[0] || '';
      
      const rPrSafe = newRPr;
      
      // The user EXPLICITLY requested to ONLY keep the following core text attributes.
      // We must completely DROP w:rStyle, w:shd, w:vertAlign, w:caps, w:spacing, etc.
      // Doing so severs any tie to Chinese document theme defaults that were causing PMingLiU fallback.
      const b = safeRegexMatch(rPrSafe, /<w:b(?:>|\/>| [^>]*>|<\/w:b>)/);
      const bCs = safeRegexMatch(rPrSafe, /<w:bCs(?:>|\/>| [^>]*>|<\/w:bCs>)/);
      const i = safeRegexMatch(rPrSafe, /<w:i(?:>|\/>| [^>]*>|<\/w:i>)/);
      const iCs = safeRegexMatch(rPrSafe, /<w:iCs(?:>|\/>| [^>]*>|<\/w:iCs>)/);
      const strike = safeRegexMatch(rPrSafe, /<w:strike(?:>|\/>| [^>]*>|<\/w:strike>)/);
      const color = safeRegexMatch(rPrSafe, /<w:color(?:>|\/>| [^>]*>|<\/w:color>)/);
      const sz = safeRegexMatch(rPrSafe, /<w:sz(?:>|\/>| [^>]*>|<\/w:sz>)/);
      const szCs = safeRegexMatch(rPrSafe, /<w:szCs(?:>|\/>| [^>]*>|<\/w:szCs>)/);
      const highlight = safeRegexMatch(rPrSafe, /<w:highlight(?:>|\/>| [^>]*>|<\/w:highlight>)/);
      const u = safeRegexMatch(rPrSafe, /<w:u(?:>|\/>| [^>]*>|<\/w:u>)/);
      
      // Since we stripped East Asian defaults from the document root, we do not need to fight 
      // the East Asian fallback here. Omitting w:eastAsia completely is safest, 
      // preventing Word from invalidating the tags if it evaluates Arial as non-East-Asian.
      const arialFonts = `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:hint="default"/>`;
      const noProof = `<w:noProof/>`;
      const langTag = `<w:lang w:val="${langCode}" w:bidi="ar-SA"/>`;
      
      // Strict standard sequence: rFonts, b, bCs, i, iCs, strike, noProof, color, sz, szCs, highlight, u, lang
      newRPr = `<w:rPr>${arialFonts}${b}${bCs}${i}${iCs}${strike}${noProof}${color}${sz}${szCs}${highlight}${u}${langTag}</w:rPr>`;
    } else if (docType === 'pptx') {
      const sz = (newRPr.match(/ sz="([^"]+)"/) || [])[1];
      const b = (newRPr.match(/ b="([^"]+)"/) || [])[1];
      const i = (newRPr.match(/ i="([^"]+)"/) || [])[1];
      const u = (newRPr.match(/ u="([^"]+)"/) || [])[1];
      const strike = (newRPr.match(/ strike="([^"]+)"/) || [])[1];
      const baseline = (newRPr.match(/ baseline="([^"]+)"/) || [])[1];

      const solidFill = (newRPr.match(/<a:solidFill[^>]*>[\s\S]*?<\/a:solidFill>/) || [])[0];
      const highlight = (newRPr.match(/<a:highlight[^>]*>[\s\S]*?<\/a:highlight>/) || [])[0];

      let attributes = `lang="${langCode}" altLang="en-US" err="0" dirty="0" smtClean="0" noProof="1"`;
      if (sz) attributes += ` sz="${sz}"`;
      if (b) attributes += ` b="${b}"`;
      if (i) attributes += ` i="${i}"`;
      if (u) attributes += ` u="${u}"`;
      if (strike) attributes += ` strike="${strike}"`;
      if (baseline) attributes += ` baseline="${baseline}"`;

      let children = '';
      if (solidFill) children += solidFill;
      if (highlight) children += highlight;

      const arialLatin = `<a:latin typeface="Arial"/>`;
      // For PPTX, standard is just to set latin and ea to Arial, it usually accepts it.
      const arialEa = `<a:ea typeface="Arial"/>`;
      const arialCs = `<a:cs typeface="Arial"/>`;
      
      newRPr = `<a:rPr ${attributes}>${children}${arialLatin}${arialEa}${arialCs}</a:rPr>`;
    }
  }
  
  return newRPr;
};

export const sanitizeOutputText = (text: string, lang: string) => {
  let cleaned = text.normalize('NFC').replace(/[·‧•]/g, ' ').replace(/\u00A0/g, ' ');
  if (lang === 'vi-VN' || lang === 'th-TH' || lang.toLowerCase().startsWith('en')) {
    cleaned = cleaned.replace(/\[f\d+\]\s*\[\/f\d+\]/g, '');
    
    if (lang === 'vi-VN') {
       const VOWELS = /[aAáÁàÀãÃảẢạẠăĂắẮằẰẵẴẳẲặẶâÂấẤầẦẫẪẩẨậẬeEéÉèÈẽẼẻẺẹẸêÊếẾềỀễỄểỂệỆiIíÍìÌĩĨỉỈịỊoOóÓòÒõÕỏỎọỌôÔốỐồỒỗỖổỔộỘơƠớỚờỜỡỠởỞợỢuUúÚùÙũŨủỦụỤưƯứỨừỪữỮửỬựỰyYýÝỳỲỹỸỷỶỵỴ]/;
       const CONSONANTS = /[bBcCdDđĐgGhHkKlLmMnNpPqQrRsStTvVxX]/;
       const TONES = /[áÁàÀãÃảẢạẠắẮằẰẵẴẳẲặẶấẤầẦẫẪẩẨậẬéÉèÈẽẼẻẺẹẸếẾềỀễỄểỂệỆíÍìÌĩĨỉỈịỊóÓòÒõÕỏỎọỌốỐồỒỗỖổỔộỘớỚờỜỡỠởỞợỢúÚùÙũŨủỦụỤứỨừỪữỮửỬựỰýÝỳỲỹỸỷỶỵỴ]/;
       
       let prevCleaned;
       do {
         prevCleaned = cleaned;
         cleaned = cleaned.replace(/([a-zA-ZÀ-ỹ]+[,\.:\?!]?)(\[\/f\d+\])(\[f\d+\])([a-zA-ZÀ-ỹ]+)/g, (match, p1, closeTag, openTag, p2) => {
             const cleanP1 = p1.replace(/[,\.:\?!]/g, '');
             const combined = cleanP1 + p2;
             
             let shouldSeparate = false;
             
             // 1. Two vowels separated by consonant
             const vcv = new RegExp(VOWELS.source + '+' + CONSONANTS.source + '+' + VOWELS.source + '+');
             if (vcv.test(combined)) shouldSeparate = true;
             
             // 2. Two tone marks
             const toneMatch = combined.match(new RegExp(TONES.source, 'g'));
             if (toneMatch && toneMatch.length >= 2) shouldSeparate = true;
             
             // 3. Combined length is implausibly long for a single Vietnamese syllable
             if (combined.length >= 8) shouldSeparate = true;
             
             // 4. Always separate after punctuation
             if (/[,\.:\?!]$/.test(p1)) shouldSeparate = true;
             
             if (shouldSeparate) {
                 return p1 + closeTag + ' ' + openTag + p2;
             }
             return match;
         });
       } while (prevCleaned !== cleaned);
    } else if (lang.toLowerCase().startsWith('en')) {
       let prevCleaned;
       do {
         prevCleaned = cleaned;
         // For English, use heuristic: if both sides are >= 3 letters, they are likely distinct words.
         cleaned = cleaned.replace(/([a-zA-Z]{3,}[,\.:\?!]?)(\[\/f\d+\])(\[f\d+\])([a-zA-Z]{3,})/g, '$1$2 $3$4');
       } while (prevCleaned !== cleaned);
    }
  }
  return cleaned;
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

  const textsToTranslate: { file: string, id: string, text: string, pBlock: string, runs?: { rPr: string, normRPr?: string, text: string }[] }[] = [];
  const fileContents: Record<string, string> = {};

  for (const docFile of docFiles) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    let content = await loadedZip.file(docFile)?.async('text');
    if (content) {
      const isTargetAsian = targetLanguages.some(l => l.includes('zh') || l.includes('ja') || l.includes('ko'));
      if (!isTargetAsian) {
        // NUKE East Asian properties from the entire document XML!
        // This stops Word from aggressively falling back to MingLiU and applying Chinese justified spacing 
        // and mid-word line-wrapping to Vietnamese diacritics.
        content = content.replace(/\s+w:eastAsia="[^"]+"/g, '');
        content = content.replace(/\s+w:eastAsiaTheme="[^"]+"/g, '');
        content = content.replace(/\s+w:hint="eastAsia"/g, '');
        content = content.replace(/<w:eastAsianLayout\b[^>]*\/>/g, '');
      }

      fileContents[docFile] = content;
      
      const pMatches = content.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g);
      if (pMatches) {
        pMatches.forEach((pBlock, index) => {
          const rMatches = pBlock.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g);
          if (rMatches) {
            let taggedText = '';
            const runs: { rPr: string, normRPr?: string, text: string }[] = [];
            
            rMatches.forEach((rBlock) => {
              const tMatches = rBlock.match(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g);
              if (tMatches) {
                const rPrMatch = rBlock.match(/<w:rPr\b[^>]*?(?:\/>|>[\s\S]*?<\/w:rPr>)/);
                const rPr = rPrMatch ? rPrMatch[0] : '';
                const normRPr = normalizeDocxRPr(rPr);
                
                let runText = '';
                tMatches.forEach(tMatch => {
                  const innerTextMatch = tMatch.match(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/);
                  if (innerTextMatch && innerTextMatch[1]) {
                    runText += unescapeXml(innerTextMatch[1]);
                  }
                });
                
                if (runText) {
                  const lastRun = runs.length > 0 ? runs[runs.length - 1] : null;
                  // Merge runs if they have identical stylistic formatting (ignoring language/font hints)
                  if (lastRun && lastRun.normRPr === normRPr) {
                    lastRun.text += runText;
                  } else {
                    runs.push({ rPr, normRPr, text: runText });
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

    const isTargetAsian = langs.some(l => l.includes('zh') || l.includes('ja') || l.includes('ko'));

    if (!isTargetAsian) {
      // Modify styles.xml and settings.xml to globally disable East Asian defaults.
      // This is the absolute core fix for the "MingLiU fallback" and "Vietnamese wide spacing" bugs,
      // as Word derives its East Asian paragraph logic from these root theme configurations.
      const styleFiles = ['word/styles.xml', 'word/settings.xml', 'word/theme/theme1.xml'];
      for (const sf of styleFiles) {
        if (loadedZip.files[sf]) {
          let sContent = await loadedZip.file(sf)?.async('text');
          if (sContent) {
            sContent = sContent.replace(/\s+w:eastAsia="[^"]+"/g, '');
            sContent = sContent.replace(/\s+w:eastAsiaTheme="[^"]+"/g, '');
            sContent = sContent.replace(/\s+w:hint="eastAsia"/g, '');
            sContent = sContent.replace(/<w:eastAsianLayout\b[^>]*\/>/g, '');
            loadedZip.file(sf, sContent);
          }
        }
      }
    }

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
            const rawTranslatedText = translatedResults[globalIndex]?.[lang] || '(翻譯失敗)';
            const translatedText = sanitizeOutputText(rawTranslatedText, lang);
            
            appendedRuns += `<w:r><w:br/></w:r>`;
            
            if (currentItem.runs && currentItem.runs.length > 0) {
              // Find the run with the longest text to use as the default formatting
              let longestRun = currentItem.runs[0];
              for (const run of currentItem.runs) {
                if (run.text.length > longestRun.text.length) {
                  longestRun = run;
                }
              }
              const defaultRPr = adjustXmlRPrForLanguage(longestRun.rPr, lang, 'docx');

              const fRegex = /\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g;
              let match;
              let lastIndex = 0;
              let hasTags = false;
              
              while ((match = fRegex.exec(translatedText)) !== null) {
                hasTags = true;
                const id = parseInt(match[1], 10);
                const text = match[2];
                const originalRPr = currentItem.runs[id] ? currentItem.runs[id].rPr : longestRun.rPr;
                const rPr = adjustXmlRPrForLanguage(originalRPr, lang, 'docx');
                
                if (match.index > lastIndex) {
                   const betweenText = translatedText.substring(lastIndex, match.index);
                   if (betweenText) {
                     const cleanBetween = stripTags(betweenText);
                     appendedRuns += `<w:r>${defaultRPr}<w:t xml:space="preserve">${escapeXml(cleanBetween)}</w:t></w:r>`;
                   }
                } else if (match.index === lastIndex && lastIndex > 0 && (lang === 'vi-VN' || lang.toLowerCase().startsWith('en'))) {
                   // Two tags are perfectly adjacent w/o any text between them: `[/fX][fY]`
                   // If neither ends/starts with a space, add one to prevent "stuck together" words like Mởtủ.
                   // Wait, we just did this globally in sanitizeOutputText (`$1$2 $3$4`), so the space is ALREADY 
                   // injected before the tag begins. It will hit `betweenText === ' '`.
                   // We don't need to double-add here!
                }
                
                if (text) {
                  let finalText = stripTags(text);
                  appendedRuns += `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(finalText)}</w:t></w:r>`;
                }
                lastIndex = fRegex.lastIndex;
              }
              
              if (!hasTags) {
                const cleanText = stripTags(translatedText);
                appendedRuns += `<w:r>${defaultRPr}<w:t xml:space="preserve">${escapeXml(cleanText)}</w:t></w:r>`;
              } else if (lastIndex < translatedText.length) {
                const remainingText = translatedText.substring(lastIndex);
                if (remainingText) {
                  const cleanText = stripTags(remainingText);
                  appendedRuns += `<w:r>${defaultRPr}<w:t xml:space="preserve">${escapeXml(cleanText)}</w:t></w:r>`;
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
        if (cell.type === 1) return; // Skip merged slave cells (ValueType.Merge = 1)
        
        if (cell.value) {
          if (typeof cell.value === 'string' && cell.value.trim().length > 0) {
            textsToTranslate.push({ row: rowNumber, col: colNumber, text: cell.value, type: 'string' });
          } else if (typeof cell.value === 'object' && (cell.value as any).richText) {
            const richTextArr = (cell.value as any).richText;
            
            // Merge fragmented richText to prevent AI translation fragmentation
            const mergedRichText: any[] = [];
            richTextArr.forEach((rt: any) => {
              if (!rt.text) return;
              const normFont = normalizeExcelFont(rt.font);
              const last = mergedRichText.length > 0 ? mergedRichText[mergedRichText.length - 1] : null;
              if (last && last.normFont === normFont) {
                last.text += rt.text;
                // keep the first part's font
              } else {
                mergedRichText.push({ ...rt, normFont });
              }
            });
            
            let taggedText = '';
            mergedRichText.forEach((rt: any, idx: number) => {
              taggedText += `[f${idx}]${rt.text}[/f${idx}]`;
            });
            if (taggedText.trim().length > 0) {
              textsToTranslate.push({ row: rowNumber, col: colNumber, text: taggedText, type: 'richText', richText: mergedRichText });
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
      const originalText = item.text;
      const translations = translatedResults[index] || {};
      
      allTranslations.push({
        sheet: worksheet.name,
        row: item.row,
        col: item.col,
        original: originalText,
        translations
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
          if (cell.type === 1) return; // Skip merged slave cells
          
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

            // Restore original text
            if (typeof cell.value === 'object' && (cell.value as any).richText) {
              const originalRichText = (cell.value as any).richText;
              originalRichText.forEach((rt: any) => {
                newRichText.push({ text: rt.text, font: rt.font || defaultFont });
              });
            } else {
              newRichText.push({ text: String(cell.value), font: defaultFont });
            }

            langs.forEach(lang => {
              const rawTranslatedText = translationItem.translations[lang] || '(翻譯失敗)';
              const translatedText = sanitizeOutputText(rawTranslatedText, lang);
              
              const adjustedDefaultFont = adjustFontForLanguage(defaultFont, lang);

              if (translationItem.original.includes('[f0]')) {
                // It was a rich text item
                newRichText.push({ text: '\n', font: adjustedDefaultFont });
                
                const fRegex = /\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g;
                let match;
                let lastIndex = 0;
                let hasTags = false;
                
                while ((match = fRegex.exec(translatedText)) !== null) {
                  hasTags = true;
                  const id = parseInt(match[1], 10);
                  const text = match[2];
                  
                  // Find original font
                  let originalFont = defaultFont;
                  if (typeof cell.value === 'object' && (cell.value as any).richText) {
                    originalFont = (cell.value as any).richText[id]?.font || defaultFont;
                  }
                  const adjustedOriginalFont = adjustFontForLanguage(originalFont, lang);
                  
                  if (match.index > lastIndex) {
                     const betweenText = translatedText.substring(lastIndex, match.index);
                     if (betweenText) {
                       const cleanBetween = stripTags(betweenText);
                       newRichText.push({ text: cleanBetween, font: adjustedDefaultFont });
                     }
                  }
                  
                  if (text) {
                    let finalText = stripTags(text);
                    newRichText.push({ text: finalText, font: adjustedOriginalFont });
                  }
                  lastIndex = fRegex.lastIndex;
                }
                
                if (!hasTags) {
                  const cleanText = stripTags(translatedText);
                  newRichText.push({ text: cleanText, font: adjustedDefaultFont });
                } else if (lastIndex < translatedText.length) {
                  const remainingText = translatedText.substring(lastIndex);
                  if (remainingText) {
                    const cleanText = stripTags(remainingText);
                    newRichText.push({ text: cleanText, font: adjustedDefaultFont });
                  }
                }
              } else {
                newRichText.push({ text: '\n' + stripTags(translatedText), font: adjustedDefaultFont });
              }
            });
            
            // Always convert to richText to ensure font adjustments are applied correctly
            // even for originally plain text cells
            const cleanedRichText = newRichText.map(rt => {
              if (rt.font === undefined) {
                return { text: rt.text };
              }
              return rt;
            });
            cell.value = { richText: cleanedRichText };
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
): Promise<{ blob: Blob; name: string }[]> => {
  updateProgress(5, 'processing');

  // ── 1. 載入 PDF ──────────────────────────────────────────────────
  const pdfjsLib = await import('pdfjs-dist');
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({
    data: arrayBuffer,
    cMapUrl: `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/cmaps/`,
    cMapPacked: true,
    standardFontDataUrl: `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/standard_fonts/`,
  }).promise;

  const totalPages = pdf.numPages;

  interface TextLine {
    y: number;
    x: number;
    text: string;
    fontSize: number;
  }

  interface PageData {
    pdfWidth: number;
    pdfHeight: number;
    lines: TextLine[];
  }

  // ── Helper：依 Y 分組成邏輯行 ─────────────────────────────────────
  function groupToLines(items: any[]): TextLine[] {
    if (!items?.length) return [];
    const sorted = [...items]
      .filter(i => i.str?.trim())
      .sort((a, b) => {
        const dy = b.transform[5] - a.transform[5];
        return Math.abs(dy) > 2 ? dy : a.transform[4] - b.transform[4];
      });
    const lines: TextLine[] = [];
    for (const item of sorted) {
      const x = item.transform[4];
      const y = item.transform[5];
      const fontSize = Math.abs(item.transform[3]) || 12;
      const match = lines.find(
        l => Math.abs(l.y - y) < Math.max(l.fontSize, fontSize) * 0.6
      );
      if (match) {
        const needSpace =
          x > match.x && !match.text.endsWith(' ') && !item.str.startsWith(' ');
        match.text = x < match.x
          ? item.str + ' ' + match.text
          : match.text + (needSpace ? ' ' : '') + item.str;
        if (x < match.x) match.x = x;
        match.fontSize = Math.max(match.fontSize, fontSize);
      } else {
        lines.push({ y, x, text: item.str, fontSize });
      }
    }
    return lines.sort((a, b) => b.y - a.y);
  }

  // ── 2. 逐頁提取文字（不再 render 截圖） ──────────────────────────
  const pages: PageData[] = [];

  for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    updateProgress(5 + (pageNum / totalPages) * 15, 'processing');

    const page = await pdf.getPage(pageNum);
    const origVp = page.getViewport({ scale: 1 });

    const textContent = await page.getTextContent();
    let lines = groupToLines(textContent.items as any[]);

    // OCR fallback：若文字稀少，render 後 OCR
    if (textContent.items.map((i: any) => i.str || '').join('').trim().length < 20) {
      console.log(`Page ${pageNum}: 文字稀少，改用 OCR`);
      const RENDER_SCALE = 2.0;
      const viewport = page.getViewport({ scale: RENDER_SCALE });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext('2d')!;
      await page.render({ canvasContext: ctx, viewport }).promise;

      const { createWorker } = await import('tesseract.js');
      const worker = await createWorker('chi_tra+eng');
      await (worker as any).setParameters({ tessedit_pageseg_mode: '11' });
      const { data: { text } } = await worker.recognize(canvas);
      await worker.terminate();
      lines = text
        .split('\n')
        .filter(l => l.trim())
        .map((t, i) => ({ y: origVp.height - i * 16, x: 0, text: t, fontSize: 12 }));
    }

    pages.push({
      pdfWidth: origVp.width,
      pdfHeight: origVp.height,
      lines,
    });
  }

  if (pages.every(p => p.lines.length === 0)) {
    throw new Error('無法從 PDF 中提取任何文字，即使嘗試了 OCR 辨識。');
  }

  updateProgress(20, 'translating');

  // ── 3. 批次翻譯所有行 ────────────────────────────────────────────
  // 收集所有頁面的有效行，並建立 globalIndex → (pageIdx, lineIdx) 的對應
  const allLineTexts: string[] = [];
  const linePageIdx: number[] = [];

  for (let pi = 0; pi < pages.length; pi++) {
    for (const line of pages[pi].lines) {
      if (line.text.trim()) {
        allLineTexts.push(line.text);
        linePageIdx.push(pi);
      }
    }
  }

  const batchSize = 15;
  const concurrency = 3;
  const translatedLines: Record<string, string>[] = new Array(allLineTexts.length);

  for (let i = 0; i < allLineTexts.length; i += batchSize * concurrency) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    const promises: Promise<Record<string, string>[]>[] = [];
    for (let j = 0; j < concurrency; j++) {
      const start = i + j * batchSize;
      if (start >= allLineTexts.length) break;
      const end = Math.min(start + batchSize, allLineTexts.length);
      promises.push(translateBatch(allLineTexts.slice(start, end), targetLanguages, industry));
    }
    const results = await Promise.all(promises);
    let fillIdx = i;
    for (const r of results) {
      for (const item of r) {
        translatedLines[fillIdx++] = item;
      }
    }
    const done = Math.min(i + batchSize * concurrency, allLineTexts.length);
    updateProgress(20 + (done / allLineTexts.length) * 50);
  }

  updateProgress(70, 'generating');

  // ── 4. 合成 PDF：純白頁面，原文一行 + 譯文一行交錯排列 ──────────
  //
  // 使用 canvas 繪製文字（支援 CJK / Thai / Vietnamese），
  // 再以 jsPDF addImage 方式輸出，確保字型完整。
  //
  // 排版邏輯：
  //   每個「block」= 原文行 + 每個語言的譯文行
  //   block 高度固定，超出一頁時自動新增 PDF 頁面（純白）
  //
  // 排版常數（pt）
  const PAGE_MARGIN = 40;
  const LINE_GAP = 3;             // 原文行 → 第一行譯文的間距
  const BLOCK_GAP = 10;           // block 與 block 之間的間距
  const ORIG_FONT_SIZE = 9;
  const TRANS_FONT_SIZE = 10;
  const ORIG_LINE_HEIGHT = ORIG_FONT_SIZE * 1.5;
  const TRANS_LINE_HEIGHT = TRANS_FONT_SIZE * 1.6;
  const TEXT_SCALE = 2;           // canvas 超取樣，讓文字清晰

  // 每個 block 的高度（pt）
  const blockHeight =
    ORIG_LINE_HEIGHT + LINE_GAP +
    targetLanguages.length * TRANS_LINE_HEIGHT +
    BLOCK_GAP;

  const TRANS_COLORS = ['#1e3a8a', '#5b21b6', '#065f46', '#7c2d12'];
  const TRANS_BG = [
    'rgba(219,234,254,0.85)',
    'rgba(237,233,254,0.85)',
    'rgba(209,250,229,0.85)',
    'rgba(254,243,199,0.85)',
  ];

  const { jsPDF } = await import('jspdf');

  const langGroups = outputMode === 'separate'
    ? targetLanguages.map(l => [l])
    : [targetLanguages];

  const generatedFiles: { blob: Blob; name: string }[] = [];

  for (const langs of langGroups) {
    if (isCancelledRef.current) throw new Error('Cancelled');

    // 決定輸出頁面尺寸（以第一頁為準）
    const firstPage = pages[0];
    const PW = firstPage.pdfWidth;
    const PH = firstPage.pdfHeight;

    const doc = new jsPDF({
      orientation: PW > PH ? 'landscape' : 'portrait',
      unit: 'pt',
      format: [PW, PH],
    });

    const contentWidth = PW - PAGE_MARGIN * 2;
    const usableHeight = PH - PAGE_MARGIN * 2;
    const blocksPerOutputPage = Math.floor(usableHeight / blockHeight);

    // 建立一個「可繪製 canvas 到 PDF 頁面」的工具函式
    const flushCanvasToPdf = (
      tc: HTMLCanvasElement,
      isFirstPage: boolean,
      pw: number,
      ph: number
    ) => {
      if (!isFirstPage) {
        doc.addPage([pw, ph], pw > ph ? 'landscape' : 'portrait');
      }
      // 白色底
      doc.setFillColor(255, 255, 255);
      doc.rect(0, 0, pw, ph, 'F');
      // 貼文字 canvas
      doc.addImage(tc.toDataURL('image/png'), 'PNG', 0, 0, pw, ph);
    };

    // 走訪所有全域行，依 blocksPerOutputPage 自動分頁
    let isFirstPdfPage = true;
    let blockOnPage = 0;
    let curY = PAGE_MARGIN;

    // 建立初始 canvas
    let tc = document.createElement('canvas');
    tc.width = PW * TEXT_SCALE;
    tc.height = PH * TEXT_SCALE;
    let ctx = tc.getContext('2d')!;
    ctx.scale(TEXT_SCALE, TEXT_SCALE);
    ctx.clearRect(0, 0, PW, PH);

    for (let gi = 0; gi < allLineTexts.length; gi++) {
      // 若本頁已滿，flush 後開新頁
      if (blockOnPage >= blocksPerOutputPage) {
        flushCanvasToPdf(tc, isFirstPdfPage, PW, PH);
        isFirstPdfPage = false;
        blockOnPage = 0;
        curY = PAGE_MARGIN;
        tc = document.createElement('canvas');
        tc.width = PW * TEXT_SCALE;
        tc.height = PH * TEXT_SCALE;
        ctx = tc.getContext('2d')!;
        ctx.scale(TEXT_SCALE, TEXT_SCALE);
        ctx.clearRect(0, 0, PW, PH);
      }

      const origText = allLineTexts[gi];
      const translations = translatedLines[gi] || {};

      // ── 原文行 ──────────────────────────────────────────────────
      const origY = curY + ORIG_LINE_HEIGHT;
      // 淡灰底帶
      ctx.fillStyle = 'rgba(245,245,245,0.9)';
      ctx.fillRect(PAGE_MARGIN - 4, curY, contentWidth + 8, ORIG_LINE_HEIGHT + 2);
      // 原文文字
      ctx.font = `${ORIG_FONT_SIZE}pt Arial, "Noto Sans TC", "Microsoft JhengHei", sans-serif`;
      ctx.fillStyle = '#333333';
      ctx.fillText(origText, PAGE_MARGIN, origY, contentWidth);

      curY += ORIG_LINE_HEIGHT + LINE_GAP;

      // ── 譯文行（每個語言一行）──────────────────────────────────
      langs.forEach((lang, langIdx) => {
        const translatedText = translations[lang] || '';
        if (!translatedText) {
          curY += TRANS_LINE_HEIGHT;
          return;
        }

        const transY = curY + TRANS_LINE_HEIGHT;
        const colorIdx = langIdx % 4;

        // 彩色底帶
        ctx.fillStyle = TRANS_BG[colorIdx];
        ctx.fillRect(PAGE_MARGIN - 4, curY, contentWidth + 8, TRANS_LINE_HEIGHT + 1);

        // 多語言時顯示語言標籤
        let textX = PAGE_MARGIN;
        if (langs.length > 1) {
          ctx.font = `bold ${ORIG_FONT_SIZE - 1}pt Arial, sans-serif`;
          ctx.fillStyle = TRANS_COLORS[colorIdx];
          const label = `[${lang}] `;
          ctx.fillText(label, textX, transY);
          textX += ctx.measureText(label).width + 2;
        }

        ctx.font = `${TRANS_FONT_SIZE}pt Arial, "Noto Sans TC", "Microsoft JhengHei", "Noto Sans Thai", sans-serif`;
        ctx.fillStyle = TRANS_COLORS[colorIdx];
        ctx.fillText(translatedText, textX, transY, contentWidth - (textX - PAGE_MARGIN));

        curY += TRANS_LINE_HEIGHT;
      });

      curY += BLOCK_GAP;
      blockOnPage++;
    }

    // flush 最後一頁（即使不滿也要輸出）
    flushCanvasToPdf(tc, isFirstPdfPage, PW, PH);

    const blob = doc.output('blob');
    const prefix = outputMode === 'separate' ? `${langs[0]}_` : 'translated_';
    generatedFiles.push({ blob, name: `${prefix}${file.name}` });
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
  
  const textsToTranslate: { slide: string, id: string, text: string, pBlock: string, runs?: { rPr: string, normRPr?: string, text: string }[] }[] = [];
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
            const runs: { rPr: string, normRPr?: string, text: string }[] = [];
            
            rMatches.forEach((rBlock) => {
              const tMatch = rBlock.match(/<a:t>([\s\S]*?)<\/a:t>/);
              if (tMatch && tMatch[1]) {
                const text = unescapeXml(tMatch[1]);
                const rPrMatch = rBlock.match(/<a:rPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:rPr>)/);
                const rPr = rPrMatch ? rPrMatch[0] : '';
                const normRPr = normalizePptxRPr(rPr);
                
                if (text) {
                  const lastRun = runs.length > 0 ? runs[runs.length - 1] : null;
                  // Merge runs if they have identical stylistic formatting (ignoring language/font hints)
                  if (lastRun && lastRun.normRPr === normRPr) {
                    lastRun.text += text;
                  } else {
                    runs.push({ rPr, normRPr, text });
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
            const rawTranslatedText = translatedResults[globalIndex]?.[lang] || '(翻譯失敗)';
            const translatedText = sanitizeOutputText(rawTranslatedText, lang);
            
            appendedRuns += `<a:br/>`;
            
            if (currentItem.runs && currentItem.runs.length > 0) {
              // Find the run with the longest text to use as the default formatting
              let longestRun = currentItem.runs[0];
              for (const run of currentItem.runs) {
                if (run.text.length > longestRun.text.length) {
                  longestRun = run;
                }
              }
              const defaultRPr = adjustXmlRPrForLanguage(longestRun.rPr, lang, 'pptx');

              const fRegex = /\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g;
              let match;
              let lastIndex = 0;
              let hasTags = false;
              
              while ((match = fRegex.exec(translatedText)) !== null) {
                hasTags = true;
                const id = parseInt(match[1], 10);
                const text = match[2];
                const originalRPr = currentItem.runs[id] ? currentItem.runs[id].rPr : longestRun.rPr;
                const rPr = adjustXmlRPrForLanguage(originalRPr, lang, 'pptx');
                
                if (match.index > lastIndex) {
                   const betweenText = translatedText.substring(lastIndex, match.index);
                   if (betweenText) {
                     const cleanBetween = stripTags(betweenText);
                     appendedRuns += `<a:r>${defaultRPr}<a:t>${escapeXml(cleanBetween)}</a:t></a:r>`;
                   }
                }
                
                if (text) {
                  let finalText = stripTags(text);
                  const originalText = currentItem.runs[id]?.text || '';
                  if (originalText.endsWith(' ') && !finalText.endsWith(' ')) {
                    finalText += ' ';
                  }
                  if (originalText.startsWith(' ') && !finalText.startsWith(' ')) {
                    finalText = ' ' + finalText;
                  }
                  appendedRuns += `<a:r>${rPr}<a:t>${escapeXml(finalText)}</a:t></a:r>`;
                }
                lastIndex = fRegex.lastIndex;
              }
              
              if (!hasTags) {
                const cleanText = stripTags(translatedText);
                appendedRuns += `<a:r>${defaultRPr}<a:t>${escapeXml(cleanText)}</a:t></a:r>`;
              } else if (lastIndex < translatedText.length) {
                const remainingText = translatedText.substring(lastIndex);
                if (remainingText) {
                  const cleanText = stripTags(remainingText);
                  appendedRuns += `<a:r>${defaultRPr}<a:t>${escapeXml(cleanText)}</a:t></a:r>`;
                }
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
