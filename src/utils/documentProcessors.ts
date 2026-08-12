import React from 'react';
import type { Worksheet } from 'exceljs';

type TranslationStatus = 'idle' | 'processing' | 'translating' | 'generating' | 'completed' | 'error';

// [修正] 原本寫成 /\[\/?\f\d+\]/，其中 \f 是「換頁字元 U+000C」而不是字母 f，
// 這條 replace 永遠不會命中。只要 LLM 回傳的 [f0]...[/f0] 配對壞掉，
// 字面的標籤就會直接印在文件上。
const stripTags = (text: string) => {
  if (!text) return text;
  return text.replace(/\[\/?f\d+\]/g, '').replace(/<\/?f\d+[^>]*>/g, '');
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

const normalizeDocxRPr = (rPr: string) => {
  if (!rPr) return '';
  return rPr.replace(/<w:rFonts[^>]*>/g, '').replace(/<w:lang[^>]*>/g, '').replace(/<w:hint[^>]*>/g, '');
};

const normalizePptxRPr = (rPr: string) => {
  if (!rPr) return '';
  return rPr.replace(/<a:latin\b[^>]*>/g, '').replace(/<a:ea\b[^>]*>/g, '').replace(/<a:cs\b[^>]*>/g, '').replace(/ lang="[^"]*"/g, '').replace(/ altLang="[^"]*"/g, '');
};

// ─── PPTX：OOXML 結構相關 helper ────────────────────────────────────────
//
// [修正] <a:t> 一定要吃屬性。PowerPoint 對任何含前後空白的文字都會寫成
// <a:t xml:space="preserve"> First </a:t>，原本的 /<a:t>/ 完全抓不到，
// 該 run 會被整段跳過，譯文就缺字。
const PPTX_T_RE = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/;
const PPTX_T_RE_SRC = '<a:t(?:\\s[^>]*)?>([\\s\\S]*?)<\\/a:t>';

// 注意：和 Word 的 <w:t> 不同，DrawingML 的 <a:t> 在 schema 裡是純 xsd:string，
// 不接受任何屬性 —— 寫 xml:space="preserve" 反而會違反 schema。
// xsd:string 的 whiteSpace facet 本來就是 preserve，前後空白會自動保留。
// (讀取時仍要容忍屬性，因為 LibreOffice / Google Slides 有時會多寫。)
const makePptxT = (escapedText: string) => `<a:t>${escapedText}</a:t>`;

// [修正] <a:br> 要帶 <a:rPr>，否則換行高度用預設字級，行距會突然變大
const makePptxBr = (rPr: string) =>
  `<a:br>${rPr && rPr.trim() ? rPr : '<a:rPr lang="en-US"/>'}</a:br>`;

// [修正／最關鍵] 譯文必須插在 <a:endParaRPr> 之前。
// OOXML 的 CT_TextParagraph 子元素順序是寫死的：
//     pPr? → (r | br | fld)* → endParaRPr?
// 原本用 pBlock.replace(/<\/a:p>$/, ...) 會把 <a:br/><a:r> 塞到 endParaRPr
// 後面，違反 schema。實測對 ISO-29500 XSD 會多出
//     Element 'br': This element is not expected.
// PowerPoint 遇到就跳「發現無法讀取的內容」，修復的做法是把該段砍掉。
const PPTX_END_PARA_RPR =
  /<a:endParaRPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:endParaRPr>)(?=\s*<\/a:p>$)/;

const insertRunsIntoPptxParagraph = (pBlock: string, runsXml: string) => {
  if (!runsXml) return pBlock;
  return PPTX_END_PARA_RPR.test(pBlock)
    ? pBlock.replace(PPTX_END_PARA_RPR, (m) => runsXml + m)
    : pBlock.replace(/<\/a:p>$/, runsXml + '</a:p>');
};

// [修正] 重算 <a:normAutofit>。
// 原文＋譯文塞進同一個 <a:p>，文字量變 2~3 倍，但 fontScale 是 PowerPoint
// 「存檔當下」算好寫死的值 —— 程式改文字它不會重算，使用者開檔也不會，
// 要人工點進文字方塊才觸發。結果就是字爆出框，這是版面看起來壞掉最直接的原因。
// 文字面積大致與字級平方成正比，所以 fontScale 除以 sqrt(成長倍率)。
const PPTX_TXBODY_RE = /<(p:txBody|a:txBody)>[\s\S]*?<\/\1>/g;

const pptxTextLength = (xml: string) => {
  let n = 0;
  let m: RegExpExecArray | null;
  const re = new RegExp(PPTX_T_RE_SRC, 'g');
  while ((m = re.exec(xml)) !== null) n += m[1].length;
  return n;
};

const rescalePptxAutofit = (newContent: string, oldContent: string) => {
  const oldBodies = oldContent.match(PPTX_TXBODY_RE) || [];
  let i = 0;
  return newContent.replace(PPTX_TXBODY_RE, (body) => {
    const before = pptxTextLength(oldBodies[i++] || '');
    const after = pptxTextLength(body);
    if (!before || after <= before) return body;
    const growth = after / before;

    return body.replace(/<a:normAutofit([^>]*)\/>/, (_m, attrs: string) => {
      const fsRaw = (attrs.match(/fontScale="(\d+)"/) || [])[1];
      const lsRaw = (attrs.match(/lnSpcReduction="(\d+)"/) || [])[1];
      const fs = parseInt(fsRaw || '100000', 10);
      const ls = parseInt(lsRaw || '0', 10);
      const newFs = Math.max(25000, Math.round(fs / Math.sqrt(growth)));
      const newLs = Math.min(20000, ls + 10000);
      return `<a:normAutofit fontScale="${newFs}" lnSpcReduction="${newLs}"/>`;
    });
  });
};

// [修正] 處理範圍不只 ppt/slides/。
// SmartArt 的文字放在 ppt/diagrams/dataN.xml，但畫面上顯示的是
// ppt/diagrams/drawingN.xml 的快取版本 —— 兩邊都要改，只改 data 不會有變化。
const PPTX_TEXT_PART_RE =
  /^ppt\/(slides\/slide\d+\.xml|notesSlides\/notesSlide\d+\.xml|diagrams\/(?:data|drawing)\d+\.xml|charts\/chart\d+\.xml)$/;

const normalizeExcelFont = (font: any) => {
  if (!font) return '';
  const { name, family, scheme, charset, ...rest } = font;
  return JSON.stringify(rest);
};

// ─── PPTX 專用的 rPr 處理 ────────────────────────────────────────────────
// [修正] 原本的 pptx 分支是「整個重組 <a:rPr>」，只保留 solidFill 與 highlight，
// 等於把 <a:ln>(外框字)、<a:effectLst>(陰影/光暈)、<a:gradFill>(漸層字)、
// <a:uLn>/<a:uFill>(底線樣式)、<a:hlinkClick>(超連結)、
// 以及 spc(字距)、cap(全大寫)、kern(字距微調)全部丟掉 —— 視覺上就是格式跑掉。
// 改成「只動字型與語系，其餘原封不動」。
//
// 注意：DrawingML 的 CT_TextCharacterProperties 子元素順序是固定的，
// latin/ea/cs 必須排在 sym / hlinkClick / hlinkMouseOver / rtl / extLst 之前。
const PPTX_FONT_ANCHORS = /<a:sym\b|<a:hlinkClick\b|<a:hlinkMouseOver\b|<a:rtl\b|<a:extLst\b/;

// 想保留原本主題字型的話，把這個設成 null 即可
const PPTX_FORCE_LATIN_FONT: string | null = 'Arial';

const PPTX_LANG_CODES: Record<string, string> = {
  'zh-TW': 'zh-TW', 'zh-CN': 'zh-CN', 'zh': 'zh-CN',
  'ja': 'ja-JP', 'ja-JP': 'ja-JP',
  'ko': 'ko-KR', 'ko-KR': 'ko-KR',
  'vi': 'vi-VN', 'vi-VN': 'vi-VN',
  'th': 'th-TH', 'th-TH': 'th-TH',
  'id': 'id-ID', 'id-ID': 'id-ID',
};

const pptxLangCode = (lang: string) =>
  PPTX_LANG_CODES[lang] || (lang && lang.includes('-') ? lang : 'en-US');

const isPptxAsianLang = (lang: string) =>
  ['zh-TW', 'zh-CN', 'zh', 'ja', 'ja-JP', 'ko', 'ko-KR'].includes(lang);

const adjustPptxRPrForLanguage = (rPr: string | undefined, lang: string) => {
  const langCode = pptxLangCode(lang);

  // 決定要不要換字型：中日韓保留原字型；越南文/泰文，或原字型是中文字型時才換
  let typeface: string | null = null;
  if (!isPptxAsianLang(lang)) {
    if (['vi', 'vi-VN', 'th', 'th-TH'].includes(lang)) {
      typeface = PPTX_FORCE_LATIN_FONT;
    } else if (rPr) {
      const fontRegex = /typeface="([^"]*)"/g;
      let m: RegExpExecArray | null;
      while ((m = fontRegex.exec(rPr)) !== null) {
        if (isChineseFont(m[1])) { typeface = PPTX_FORCE_LATIN_FONT; break; }
      }
    }
  }

  let out = rPr && rPr.trim() ? rPr : '<a:rPr lang="en-US"/>';

  if (typeface) {
    // 自閉合標籤先展開，才有地方放子元素
    if (/\/>\s*$/.test(out)) out = out.replace(/\/>\s*$/, '></a:rPr>');
    out = out
      .replace(/<a:latin\b[^>]*\/>/g, '')
      .replace(/<a:ea\b[^>]*\/>/g, '')
      .replace(/<a:cs\b[^>]*\/>/g, '');
    const fonts =
      `<a:latin typeface="${typeface}"/>` +
      `<a:ea typeface="${typeface}"/>` +
      `<a:cs typeface="${typeface}"/>`;
    out = PPTX_FONT_ANCHORS.test(out)
      ? out.replace(PPTX_FONT_ANCHORS, (m) => fonts + m)
      : out.replace(/<\/a:rPr>\s*$/, fonts + '</a:rPr>');
  }

  // 更新語系（原本中日韓路徑完全不改 lang，會導致字型後援選錯）
  out = /\slang="/.test(out)
    ? out.replace(/\slang="[^"]*"/, ` lang="${langCode}"`)
    : out.replace(/^<a:rPr/, `<a:rPr lang="${langCode}"`);

  // 關掉拼字檢查紅波浪線
  if (!/\snoProof="/.test(out)) out = out.replace(/^<a:rPr/, '<a:rPr noProof="1"');

  return out;
};

const adjustXmlRPrForLanguage = (rPr: string | undefined, lang: string, docType: 'docx' | 'pptx') => {
  if (docType === 'pptx') return adjustPptxRPrForLanguage(rPr, lang);

  const isAsianLang = ['zh-TW', 'zh-CN', 'ja', 'ko'].includes(lang);
  const needsStrongFontAdjustment = ['vi', 'th'].includes(lang);
  
  if (isAsianLang) return rPr || '';
  
  let newRPr = rPr || '';
  let shouldAdjust = needsStrongFontAdjustment;
  
  if (!shouldAdjust && newRPr) {
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
    const langCode = lang === 'th' ? 'th-TH' : 'en-US';
    
    if (docType === 'docx') {
      const safeRegexMatch = (str: string, regex: RegExp) => (str.match(regex) || [])[0] || '';
      
      const rPrSafe = newRPr;
      
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
      
      const arialFonts = `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:hint="default"/>`;
      const noProof = `<w:noProof/>`;
      const langTag = `<w:lang w:val="${langCode}" w:bidi="ar-SA"/>`;
      
      newRPr = `<w:rPr>${arialFonts}${b}${bCs}${i}${iCs}${strike}${noProof}${color}${sz}${szCs}${highlight}${u}${langTag}</w:rPr>`;
    }
  }
  
  return newRPr;
};

// 將段落的 <w:pPr> 中的對齊改為左對齊（覆蓋兩端對齊，避免翻譯後字間距過大）
const forceLeftAlignInPPr = (pBlock: string): string => {
  // 若段落有 <w:pPr>，在其中處理 <w:jc>
  if (/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/.test(pBlock)) {
    return pBlock.replace(/(<w:pPr\b[^>]*>[\s\S]*?)(<\/w:pPr>)/, (match, pPrContent, closeTag) => {
      if (/<w:jc\b/.test(pPrContent)) {
        // 已有 <w:jc>，替換成 left
        return pPrContent.replace(/<w:jc\b[^>]*\/?>/, '<w:jc w:val="left"/>') + closeTag;
      } else {
        // 沒有 <w:jc>，在 </w:pPr> 前插入
        return pPrContent + '<w:jc w:val="left"/>' + closeTag;
      }
    });
  } else if (/<w:p\b[^>]*>/.test(pBlock)) {
    // 段落完全沒有 <w:pPr>，在 <w:p...> 後插入
    return pBlock.replace(/^(<w:p\b[^>]*>)/, '$1<w:pPr><w:jc w:val="left"/></w:pPr>');
  }
  return pBlock;
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
             
             const vcv = new RegExp(VOWELS.source + '+' + CONSONANTS.source + '+' + VOWELS.source + '+');
             if (vcv.test(combined)) shouldSeparate = true;
             
             const toneMatch = combined.match(new RegExp(TONES.source, 'g'));
             if (toneMatch && toneMatch.length >= 2) shouldSeparate = true;
             
             if (combined.length >= 8) shouldSeparate = true;
             
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
         cleaned = cleaned.replace(/([a-zA-Z]{3,}[,\.:\?!]?)(\[\/f\d+\])(\[f\d+\])([a-zA-Z]{3,})/g, '$1$2 $3$4');
       } while (prevCleaned !== cleaned);
    }
  }
  return cleaned;
};

// ─── 統一的段落 run 解析器 ──────────────────────────────────────────────────
// 將一個 <w:p> 區塊解析成「文字區段」陣列。
//
// 關鍵規則：<w:br/> 幾乎永遠包在 <w:r> 內部，例如：
//   <w:r><w:rPr>...</w:rPr><w:br/></w:r>
//   <w:r><w:br/></w:r>
//
// 正確做法：解析每個 <w:r> 時，先看它是否含有 <w:br/>：
//   1. 若 <w:r> 同時有 <w:t> 和 <w:br/>，先輸出文字 run，再輸出 br run。
//   2. 若 <w:r> 只有 <w:br/> 沒有 <w:t>，輸出 br run。
//   3. 若 <w:r> 只有 <w:t>，正常輸出文字 run。
//
// 這樣提取文字和回寫時的 run index 才會完全對應。
type DocxRun = { rPr: string; normRPr: string; text: string; isBr?: boolean };

function parseDocxParagraphRuns(pBlock: string, unescapeXml: (s: string) => string): DocxRun[] {
  const runs: DocxRun[] = [];

  // 只抓 <w:r ...>...</w:r>，不再單獨抓頂層 <w:br/>（因為 br 幾乎都在 w:r 內）
  const rRegex = /<w:r\b[^>]*>[\s\S]*?<\/w:r>/g;
  let match: RegExpExecArray | null;

  while ((match = rRegex.exec(pBlock)) !== null) {
    const rToken = match[0];

    // 取出此 run 的 rPr
    const rPrMatch = rToken.match(/<w:rPr\b[^>]*?(?:\/>|>[\s\S]*?<\/w:rPr>)/);
    const rPr = rPrMatch ? rPrMatch[0] : '';
    const normRPr = normalizeDocxRPr(rPr);

    // 取出文字（可能有多個 <w:t>）
    const tMatches = rToken.match(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/g);
    let runText = '';
    if (tMatches) {
      for (const tTag of tMatches) {
        const inner = tTag.match(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/);
        if (inner && inner[1]) runText += unescapeXml(inner[1]);
      }
    }

    // 偵測此 run 是否含有 <w:br/>
    const hasBr = /<w:br\b[^>]*\/?>/.test(rToken);

    // 先輸出文字部分（若有）
    if (runText) {
      const last = runs.length > 0 ? runs[runs.length - 1] : null;
      if (last && !last.isBr && last.normRPr === normRPr) {
        last.text += runText;
      } else {
        runs.push({ rPr, normRPr, text: runText });
      }
    }

    // 再輸出 br 佔位（若有）
    if (hasBr) {
      const prevRPr = runs.length > 0 ? runs[runs.length - 1].rPr : rPr;
      const prevNormRPr = runs.length > 0 ? runs[runs.length - 1].normRPr : normRPr;
      runs.push({ rPr: prevRPr, normRPr: prevNormRPr, text: '\n', isBr: true });
    }
  }

  return runs;
}

// 將 runs 陣列轉成帶 [f0]...[/f0] tag 的字串
function runsToTaggedText(runs: DocxRun[]): string {
  let tagged = '';
  let fIdx = 0;
  for (const run of runs) {
    if (run.isBr) {
      // br 本身不參與翻譯，但要佔一個 slot 讓 index 對得上
      tagged += `[f${fIdx}]\n[/f${fIdx}]`;
    } else {
      tagged += `[f${fIdx}]${run.text}[/f${fIdx}]`;
    }
    fIdx++;
  }
  return tagged;
}

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
    !name.includes('fontTable') &&
    !name.includes('header') &&
    !name.includes('footer')
  );
  
  const unescapeXml = (text: string) => text.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  const escapeXml = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

  type DocxItem = {
    file: string;
    markerId: string;
    text: string;
    pBlock: string;
    runs: DocxRun[];
  };
  const textsToTranslate: DocxItem[] = [];
  const fileContents: Record<string, string> = {};
  let globalMarkerCounter = 0;

  for (const docFile of docFiles) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    let content = await loadedZip.file(docFile)?.async('text');
    if (content) {
      const isTargetAsian = targetLanguages.some(l => l.includes('zh') || l.includes('ja') || l.includes('ko'));
      if (!isTargetAsian) {
        content = content.replace(/\s+w:eastAsia="[^"]+"/g, '');
        content = content.replace(/\s+w:eastAsiaTheme="[^"]+"/g, '');
        content = content.replace(/\s+w:hint="eastAsia"/g, '');
        content = content.replace(/<w:eastAsianLayout\b[^>]*\/>/g, '');
      }

      // ── 第一步：掃描每個 <w:p>，若含可翻譯文字則加 data-mid ──
      content = content.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (pBlock) => {
        const runs = parseDocxParagraphRuns(pBlock, unescapeXml);
        const hasText = runs.some(r => !r.isBr && r.text.trim().length > 0);
        if (!hasText) return pBlock;
        const mid = `docx_p_${globalMarkerCounter++}`;
        return pBlock.replace(/^(<w:p\b)/, `$1 data-mid="${mid}"`);
      });

      fileContents[docFile] = content;

      // ── 第二步：用同一套 parseDocxParagraphRuns 提取文字 ──
      const pMatches = content.match(/<w:p\b[^>]*data-mid="[^"]+"[^>]*>[\s\S]*?<\/w:p>/g);
      if (pMatches) {
        for (const pBlock of pMatches) {
          const midMatch = pBlock.match(/data-mid="([^"]+)"/);
          if (!midMatch) continue;
          const markerId = midMatch[1];

          const runs = parseDocxParagraphRuns(pBlock, unescapeXml);
          const taggedText = runsToTaggedText(runs);

          if (taggedText.replace(/\[f\d+\][\s\n]*\[\/f\d+\]/g, '').trim().length > 0) {
            textsToTranslate.push({ file: docFile, markerId, text: taggedText, pBlock, runs });
          }
        }
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
      
      content = content.replace(/<w:p\b[^>]*data-mid="([^"]+)"[^>]*>[\s\S]*?<\/w:p>/g, (pBlock, markerId) => {
        const currentItem = fileTexts.find(t => t.markerId === markerId);
        
        if (currentItem) {
          const globalIndex = textsToTranslate.findIndex(t => t.markerId === markerId);
          
          // 找最長 run 作為 fallback 格式
          let longestRun = currentItem.runs.find(r => !r.isBr) || currentItem.runs[0];
          for (const run of currentItem.runs) {
            if (!run.isBr && run.text.length > longestRun.text.length) {
              longestRun = run;
            }
          }

          let appendedRuns = '';
          langs.forEach(lang => {
            const rawTranslatedText = translatedResults[globalIndex]?.[lang] || '(翻譯失敗)';
            const translatedText = sanitizeOutputText(rawTranslatedText, lang);
            
            // 換行分隔
            appendedRuns += `<w:r><w:br/></w:r>`;
            
            const defaultRPr = adjustXmlRPrForLanguage(longestRun.rPr, lang, 'docx');
            const fRegex = /\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g;
            let fMatch: RegExpExecArray | null;
            let lastIndex = 0;
            let hasTags = false;
            
            while ((fMatch = fRegex.exec(translatedText)) !== null) {
              hasTags = true;
              const id = parseInt(fMatch[1], 10);
              const text = fMatch[2];

              // 對應原始 run
              const originalRun = currentItem.runs[id];

              if (fMatch.index > lastIndex) {
                const betweenText = translatedText.substring(lastIndex, fMatch.index);
                const cleanBetween = stripTags(betweenText);
                if (cleanBetween) {
                  appendedRuns += `<w:r>${defaultRPr}<w:t xml:space="preserve">${escapeXml(cleanBetween)}</w:t></w:r>`;
                }
              }

              if (originalRun?.isBr) {
                // 原始是 br，翻譯後若遇到 \n 就插 br，否則略過
                if (text === '\n' || text.includes('\n')) {
                  appendedRuns += `<w:r><w:br/></w:r>`;
                }
              } else if (text && text !== '\n') {
                const rPr = adjustXmlRPrForLanguage(originalRun ? originalRun.rPr : longestRun.rPr, lang, 'docx');
                const finalText = stripTags(text);
                appendedRuns += `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(finalText)}</w:t></w:r>`;
              }

              lastIndex = fRegex.lastIndex;
            }
            
            if (!hasTags) {
              const cleanText = stripTags(translatedText);
              if (cleanText) {
                appendedRuns += `<w:r>${defaultRPr}<w:t xml:space="preserve">${escapeXml(cleanText)}</w:t></w:r>`;
              }
            } else if (lastIndex < translatedText.length) {
              const remainingText = stripTags(translatedText.substring(lastIndex));
              if (remainingText) {
                appendedRuns += `<w:r>${defaultRPr}<w:t xml:space="preserve">${escapeXml(remainingText)}</w:t></w:r>`;
              }
            }
          });
          
          // 移除 data-mid 屬性，並強制段落左對齊（避免兩端對齊造成翻譯文字字間距過大）
          let cleanedPBlock = pBlock.replace(/ data-mid="[^"]+"/, '');
          cleanedPBlock = forceLeftAlignInPPr(cleanedPBlock);
          
          const closeTag = '</w:p>';
          const insertPos = cleanedPBlock.lastIndexOf(closeTag);
          if (insertPos === -1) return cleanedPBlock;
          return cleanedPBlock.slice(0, insertPos) + appendedRuns + closeTag;
        }
        
        // 沒有翻譯結果的段落：同樣移除 data-mid 並強制左對齊
        let cleanedPBlock = pBlock.replace(/ data-mid="[^"]+"/, '');
        cleanedPBlock = forceLeftAlignInPPr(cleanedPBlock);
        return cleanedPBlock;
      });

      content = content.replace(/ data-mid="[^"]+"/g, '');

      // 強制表格 fixed layout，防止翻譯後欄寬撐開
      content = content.replace(/(<w:tbl\b[^>]*>)(\s*<w:tblPr\b[^>]*>[\s\S]*?<\/w:tblPr>)/g, (match, tblOpen, tblPr) => {
        if (/<w:tblLayout\b/.test(tblPr)) {
          return tblOpen + tblPr.replace(/<w:tblLayout\b[^>]*\/?>/g, '<w:tblLayout w:type="fixed"/>');
        }
        return tblOpen + tblPr.replace(/<\/w:tblPr>/, '<w:tblLayout w:type="fixed"/></w:tblPr>');
      });

      content = content.replace(/(<w:tbl\b[^>]*>)(?!\s*<w:tblPr\b)/g, '$1<w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>');
      
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
  
  const allTranslations: {
    sheet: string;
    row: number;
    col: number;
    original: string;
    translations: Record<string, string>;
    mergedRichText?: any[];
    isFormula?: boolean;
  }[] = [];

  for (const worksheet of workbook.worksheets) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    
    const textsToTranslate: {
      row: number;
      col: number;
      text: string;
      type: 'string' | 'richText' | 'formula';
      richText?: any[];
      mergedRichText?: any[];
    }[] = [];
    
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        if (cell.type === 1) return;
        
        if (cell.value) {
          if (typeof cell.value === 'string' && cell.value.trim().length > 0) {
            textsToTranslate.push({ row: rowNumber, col: colNumber, text: cell.value, type: 'string' });
          } else if (typeof cell.value === 'object' && (cell.value as any).richText) {
            const richTextArr = (cell.value as any).richText;
            
            const mergedRichText: any[] = [];
            richTextArr.forEach((rt: any) => {
              if (!rt.text) return;
              const normFont = normalizeExcelFont(rt.font);
              const last = mergedRichText.length > 0 ? mergedRichText[mergedRichText.length - 1] : null;
              if (last && last.normFont === normFont) {
                last.text += rt.text;
              } else {
                mergedRichText.push({ ...rt, normFont });
              }
            });
            
            let taggedText = '';
            mergedRichText.forEach((rt: any, idx: number) => {
              taggedText += `[f${idx}]${rt.text}[/f${idx}]`;
            });
            if (taggedText.trim().length > 0) {
              textsToTranslate.push({
                row: rowNumber,
                col: colNumber,
                text: taggedText,
                type: 'richText',
                richText: richTextArr,
                mergedRichText
              });
            }
          } else if (
            typeof cell.value === 'object' &&
            (cell.value as any).formula &&
            typeof (cell.value as any).result === 'string' &&
            (cell.value as any).result.trim().length > 0
          ) {
            textsToTranslate.push({
              row: rowNumber,
              col: colNumber,
              text: (cell.value as any).result,
              type: 'formula'
            });
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
      const translations = translatedResults[index] || {};
      allTranslations.push({
        sheet: worksheet.name,
        row: item.row,
        col: item.col,
        original: item.text,
        translations,
        mergedRichText: item.mergedRichText,
        isFormula: item.type === 'formula'
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

    const translatedCols: Record<string, Set<number>> = {};

    for (const worksheet of newWorkbook.worksheets) {
      worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          if (cell.type === 1) return;
          
          const translationItem = allTranslations.find(
            t => t.sheet === worksheet.name && t.row === rowNumber && t.col === colNumber
          );
          if (translationItem) {
            let defaultFont: any = {};
            if (typeof cell.value === 'object' && (cell.value as any).richText && (cell.value as any).richText.length > 0) {
              defaultFont = (cell.value as any).richText[0].font || {};
            } else if (cell.font) {
              defaultFont = cell.font;
            }

            const newRichText: any[] = [];

            if (translationItem.isFormula) {
              newRichText.push({ text: String((cell.value as any).result || ''), font: defaultFont });
            } else if (typeof cell.value === 'object' && (cell.value as any).richText) {
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

              const hasMergedRichText = !!(translationItem.mergedRichText && translationItem.mergedRichText.length > 0);

              if (hasMergedRichText && translationItem.original.includes('[f0]')) {
                newRichText.push({ text: '\n', font: adjustedDefaultFont });
                
                const fRegex = /\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g;
                let match;
                let lastIndex = 0;
                let hasTagsInTranslation = false;
                
                while ((match = fRegex.exec(translatedText)) !== null) {
                  hasTagsInTranslation = true;
                  const id = parseInt(match[1], 10);
                  const text = match[2];
                  
                  let originalFont = defaultFont;
                  if (translationItem.mergedRichText && translationItem.mergedRichText[id]) {
                    originalFont = translationItem.mergedRichText[id].font || defaultFont;
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
                
                if (!hasTagsInTranslation) {
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
            
            const cleanedRichText = newRichText.map(rt => {
              if (rt.font === undefined) {
                return { text: rt.text };
              }
              return rt;
            });
            cell.value = { richText: cleanedRichText };

            if (!translatedCols[worksheet.name]) translatedCols[worksheet.name] = new Set();
            translatedCols[worksheet.name].add(colNumber);
          }
          cell.alignment = { 
            ...(cell.alignment || {}), 
            wrapText: true 
          };
        });
      });

      const MAX_COL_WIDTH = 30;
      const sheetTranslatedCols = translatedCols[worksheet.name];
      if (sheetTranslatedCols) {
        sheetTranslatedCols.forEach(colNumber => {
          const col = worksheet.getColumn(colNumber);
          const currentWidth = (col.width && col.width > 0) ? col.width : 12;
          if (currentWidth > MAX_COL_WIDTH) {
            col.width = MAX_COL_WIDTH;
          }
        });
      }
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

  const pages: PageData[] = [];

  for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    updateProgress(5 + (pageNum / totalPages) * 15, 'processing');

    const page = await pdf.getPage(pageNum);
    const origVp = page.getViewport({ scale: 1 });

    const textContent = await page.getTextContent();
    let lines = groupToLines(textContent.items as any[]);

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

  const PAGE_MARGIN = 40;
  const LINE_GAP = 4;
  const BLOCK_GAP = 10;
  const ORIG_FONT_SIZE = 9;
  const TRANS_FONT_SIZE = 10;
  const ORIG_LINE_HEIGHT = ORIG_FONT_SIZE * 1.55;
  const TRANS_LINE_HEIGHT = TRANS_FONT_SIZE * 1.6;
  const TEXT_SCALE = 2;

  const blockHeight =
    ORIG_LINE_HEIGHT + LINE_GAP +
    targetLanguages.length * TRANS_LINE_HEIGHT +
    BLOCK_GAP;

  const TRANS_COLORS = ['#1e3a8a', '#5b21b6', '#065f46', '#7c2d12'];

  const { jsPDF } = await import('jspdf');

  const langGroups = outputMode === 'separate'
    ? targetLanguages.map(l => [l])
    : [targetLanguages];

  const generatedFiles: { blob: Blob; name: string }[] = [];

  for (const langs of langGroups) {
    if (isCancelledRef.current) throw new Error('Cancelled');

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

    const flushCanvasToPdf = (
      tc: HTMLCanvasElement,
      isFirstPage: boolean,
      pw: number,
      ph: number
    ) => {
      if (!isFirstPage) {
        doc.addPage([pw, ph], pw > ph ? 'landscape' : 'portrait');
      }
      doc.setFillColor(255, 255, 255);
      doc.rect(0, 0, pw, ph, 'F');
      doc.addImage(tc.toDataURL('image/png'), 'PNG', 0, 0, pw, ph);
    };

    let isFirstPdfPage = true;
    let blockOnPage = 0;
    let curY = PAGE_MARGIN;

    const makeCanvas = () => {
      const tc = document.createElement('canvas');
      tc.width = PW * TEXT_SCALE;
      tc.height = PH * TEXT_SCALE;
      const ctx = tc.getContext('2d')!;
      ctx.scale(TEXT_SCALE, TEXT_SCALE);
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, PW, PH);
      return { tc, ctx };
    };

    let { tc, ctx } = makeCanvas();

    for (let gi = 0; gi < allLineTexts.length; gi++) {
      if (blockOnPage >= blocksPerOutputPage) {
        flushCanvasToPdf(tc, isFirstPdfPage, PW, PH);
        isFirstPdfPage = false;
        blockOnPage = 0;
        curY = PAGE_MARGIN;
        ({ tc, ctx } = makeCanvas());
      }

      const origText = allLineTexts[gi];
      const translations = translatedLines[gi] || {};

      const origBaseline = curY + ORIG_LINE_HEIGHT;
      ctx.font = `${ORIG_FONT_SIZE}pt Arial, "Noto Sans TC", "Microsoft JhengHei", sans-serif`;
      ctx.fillStyle = '#444444';
      ctx.fillText(origText, PAGE_MARGIN, origBaseline, contentWidth);

      curY += ORIG_LINE_HEIGHT + LINE_GAP;

      langs.forEach((lang, langIdx) => {
        const translatedText = translations[lang] || '';
        if (!translatedText) {
          curY += TRANS_LINE_HEIGHT;
          return;
        }

        const transBaseline = curY + TRANS_LINE_HEIGHT;
        const colorIdx = langIdx % TRANS_COLORS.length;

        let textX = PAGE_MARGIN;

        if (langs.length > 1) {
          ctx.font = `bold ${ORIG_FONT_SIZE - 1}pt Arial, sans-serif`;
          ctx.fillStyle = TRANS_COLORS[colorIdx];
          const label = `[${lang}] `;
          ctx.fillText(label, textX, transBaseline);
          textX += ctx.measureText(label).width + 2;
        }

        ctx.font = `${TRANS_FONT_SIZE}pt Arial, "Noto Sans TC", "Microsoft JhengHei", "Noto Sans Thai", sans-serif`;
        ctx.fillStyle = TRANS_COLORS[colorIdx];
        ctx.fillText(translatedText, textX, transBaseline, contentWidth - (textX - PAGE_MARGIN));

        curY += TRANS_LINE_HEIGHT;
      });

      curY += BLOCK_GAP;
      blockOnPage++;
    }

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
  
  // [修正] 除了投影片，備忘稿、SmartArt(data + drawing 兩份)、圖表也要一起翻，
  // 否則 SmartArt 會維持原文，看起來就像「有些地方沒翻、版面錯亂」。
  const slideFiles = Object.keys(loadedZip.files)
    .filter(name => PPTX_TEXT_PART_RE.test(name))
    .sort();
  
  const textsToTranslate: { slide: string, id: string, text: string, pBlock: string, runs?: { rPr: string, normRPr?: string, text: string }[] }[] = [];
  const slideContents: Record<string, string> = {};

  const unescapeXml = (text: string) => text.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
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
              // [修正] 用會吃屬性的 regex，才抓得到 <a:t xml:space="preserve">
              const tMatch = rBlock.match(PPTX_T_RE);
              if (tMatch && tMatch[1]) {
                const text = unescapeXml(tMatch[1]);
                const rPrMatch = rBlock.match(/<a:rPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:rPr>)/);
                const rPr = rPrMatch ? rPrMatch[0] : '';
                const normRPr = normalizePptxRPr(rPr);
                
                if (text) {
                  const lastRun = runs.length > 0 ? runs[runs.length - 1] : null;
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

  // id -> 在 textsToTranslate 中的位置，避免每個段落都跑一次 findIndex
  const indexById = new Map<string, number>();
  textsToTranslate.forEach((t, i) => indexById.set(t.id, i));

  const batchSize = 10;
  const concurrency = 3;
  const translatedResults: Record<string, string>[] = [];
  
  for (let i = 0; i < textsToTranslate.length; i += batchSize * concurrency) {
    if (isCancelledRef.current) throw new Error('Cancelled');
    
    const promises = [];
    for (let j = 0; j < concurrency && (i + j * batchSize) < textsToTranslate.length; j++) {
      const start = i + j * batchSize;
      const batch = textsToTranslate.slice(start, start + batchSize).map(item => item.text);
      promises.push(
        translateBatch(batch, targetLanguages, industry).then((res) => {
          // 保險：API 少回或多回都會讓後面所有段落錯位，補齊到原長度
          const fixed = res.slice(0, batch.length);
          while (fixed.length < batch.length) fixed.push({});
          return fixed;
        })
      );
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
      const originalContent = slideContents[slideFile];
      if (!originalContent) continue;
      const slideTexts = textsToTranslate.filter(t => t.slide === slideFile);
      
      let pIndex = 0;
      let content = originalContent.replace(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/g, (pBlock) => {
        const currentIndex = pIndex++;
        const currentItem = slideTexts.find(t => t.id === `${slideFile}_${currentIndex}`);
        
        if (currentItem) {
          const globalIndex = indexById.get(currentItem.id) ?? -1;
          
          let appendedRuns = '';
          langs.forEach(lang => {
            const rawTranslatedText = translatedResults[globalIndex]?.[lang] || '(翻譯失敗)';
            const translatedText = sanitizeOutputText(rawTranslatedText, lang);
            
            const runs = currentItem.runs || [];
            let longestRun = runs[0] || { rPr: '', text: '' };
            for (const run of runs) {
              if (run.text.length > longestRun.text.length) longestRun = run;
            }
            const defaultRPr = adjustXmlRPrForLanguage(longestRun.rPr, lang, 'pptx');

            // [修正] 換行要帶 rPr，否則行高用預設字級
            appendedRuns += makePptxBr(defaultRPr);

            const fRegex = /\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g;
            let match;
            let lastIndex = 0;
            let hasTags = false;
            
            while ((match = fRegex.exec(translatedText)) !== null) {
              hasTags = true;
              const id = parseInt(match[1], 10);
              const originalRPr = runs[id] ? runs[id].rPr : longestRun.rPr;
              const rPr = adjustXmlRPrForLanguage(originalRPr, lang, 'pptx');
              
              if (match.index > lastIndex) {
                 const betweenText = stripTags(translatedText.substring(lastIndex, match.index));
                 if (betweenText) {
                   appendedRuns += `<a:r>${defaultRPr}${makePptxT(escapeXml(betweenText))}</a:r>`;
                 }
              }
              
              let finalText = stripTags(match[2]);
              if (finalText) {
                const originalText = runs[id]?.text || '';
                if (originalText.endsWith(' ') && !finalText.endsWith(' ')) finalText += ' ';
                if (originalText.startsWith(' ') && !finalText.startsWith(' ')) finalText = ' ' + finalText;
                appendedRuns += `<a:r>${rPr}${makePptxT(escapeXml(finalText))}</a:r>`;
              }
              lastIndex = fRegex.lastIndex;
            }
            
            // 沒有標籤 → 整段；有標籤 → 收尾剩下的部分。
            // 兩種情況都要過 stripTags，才不會把壞掉的標籤印到投影片上。
            const tail = hasTags ? translatedText.substring(lastIndex) : translatedText;
            const cleanTail = stripTags(tail);
            if (cleanTail) {
              appendedRuns += `<a:r>${defaultRPr}${makePptxT(escapeXml(cleanTail))}</a:r>`;
            }
          });
          
          // [修正] 插在 <a:endParaRPr> 之前，不能直接接在 </a:p> 前
          return insertRunsIntoPptxParagraph(pBlock, appendedRuns);
        }
        
        return pBlock;
      });
      
      // [修正] 依文字增長比例重算自動縮放，避免字爆出框
      content = rescalePptxAutofit(content, originalContent);
      
      loadedZip.file(slideFile, content);
    }

    // [修正] JSZip 預設是 STORE(不壓縮)，輸出檔會膨脹好幾倍
    const blob = await loadedZip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
    const prefix = outputMode === 'separate' ? `${langs[0]}_` : 'translated_';
    generatedFiles.push({ blob, name: `${prefix}${file.name}` });
  }

  updateProgress(100, 'completed');
  return generatedFiles;
};
