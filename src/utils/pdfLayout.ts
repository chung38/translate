// ─────────────────────────────────────────────────────────────
// PDF 版面分析：把 pdf.js 讀到的文字片段整理成「一格一格、一段一段」的翻譯區塊
//
// 為什麼需要：pdf.js 給的是零碎的文字片段。直接把同一高度的片段串成一行，
// 表格同一列的好幾格會被併成一句、直排標籤會被拆成一兩個字、換行的句子會被切斷。
// 這裡依序做：
//   1. 片段 → 同一列 → 依「大空隙」和「表格直線」切成一格一格
//   2. 同一格裡上下相連、前一行寫滿的 → 接回同一段（直排標籤也是這樣接起來）
//   3. 每一段算出譯文可以用的範圍（往右、往下延伸到格線或下一段為止）
//
// 座標一律用「左上角為原點、往下為正」的 pt，跟畫布一樣。
// 這個檔案不碰 DOM，方便寫測試。
// ─────────────────────────────────────────────────────────────

export interface Box { left: number; top: number; right: number; bottom: number }

/** pdf.js getTextContent() 的 item（只用到這幾個欄位） */
export interface RawTextItem { str: string; transform: number[]; width: number }

export interface HLine { y: number; x0: number; x1: number }
export interface VLine { x: number; y0: number; y1: number }
export interface GridLines { h: HLine[]; v: VLine[] }

/** 一格裡同一列的一段文字 */
export interface Segment extends Box {
  text: string;
  fontSize: number;
  /** 要蓋掉的原文位置（每個片段一個框） */
  glyphs: Box[];
}

export interface Cell extends Box {
  /** 每個邊是不是真的有格線（沒有的話是頁面邊界或推估的） */
  hasLeft: boolean; hasRight: boolean; hasTop: boolean; hasBottom: boolean;
}

export interface LayoutBlock {
  /** 送去翻譯的原文（已去掉中文字距的空白） */
  text: string;
  fontSize: number;
  /** 譯文最大字級（一般標籤不要比內文大太多，標題例外） */
  maxFont?: number;
  /** 原文的位置，畫譯文前要先蓋掉 */
  glyphs: Box[];
  /** 原文整段的外框 */
  source: Box;
  /** 譯文可以使用的範圍 */
  box: Box;
  /** 所在的表格格子（蓋白時不要超出，格線才不會被蓋掉） */
  cell?: Cell;
  /** 直排標籤：譯文要轉 90 度寫 */
  vertical: boolean;
  /**
   * 同一格裡由上往下排的清單（□1～□5）：畫的時候整格一起重排，
   * 短的項目省下來的空間給長的用，整格用同一個字級
   */
  stack?: { id: number; region: Box };
  /** 原文是靠右的一行（例：「記錄：陳淑燕 4/29」）：譯文往左延伸、靠右對齊 */
  alignRight?: boolean;
  /** 原文是表格外置中的標題（例：「會議記錄表」）：譯文以原文中心為準左右展開、置中 */
  alignCenter?: boolean;
}

export const LAYOUT = {
  /** 同一列兩個片段之間空隙超過幾個字寬就切開 */
  splitGapEm: 1.4,
  /** 上下兩行之間的空隙在幾個字高以內才算同一段 */
  paragraphGapEm: 0.75,
  /** 前一行的右端離右邊界幾個字寬以內，才算「寫滿了、下一行是接續」 */
  fullLineEm: 3,
  /** 譯文範圍離格線保留的距離 */
  pad: 1.2,
  /** 頁面外圍留白 */
  pageMargin: 14,
};

// ── 字元分類 ─────────────────────────────────────────────────

const CJK = '\\u3400-\\u9fff\\uf900-\\ufaff';
const CJK_PUNCT = '\\u3000-\\u303f\\uff00-\\uffef';
const CJK_RE = new RegExp(`[${CJK}]`);
const LETTER_RE = new RegExp(`[${CJK}\\u0e00-\\u0e7fA-Za-z\\u00c0-\\u024f]`);
const PUA_RE = /[-]/;

/** 有沒有需要翻譯的字（只有符號、數字、勾選框的不翻） */
export const isTranslatable = (text: string): boolean => LETTER_RE.test(text);

/** 圓圈數字、私用區符號（Wingdings 之類）：保留原樣、不要併進文字裡 */
const isStandaloneSymbol = (text: string): boolean => {
  const t = text.trim();
  return t.length > 0 && t.length <= 2 && (PUA_RE.test(t) || /^[○◯●◎①-⑳❶-❿]$/.test(t));
};

/**
 * 去掉中文字距造成的空白：「臺 北 業 務 組」→「臺北業務組」、「第 8 1 條」→「第81條」。
 * 英文單字之間的空白保留。
 */
export function cleanCjkSpacing(text: string): string {
  let t = (text || '').replace(/[ 　]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!t) return t;
  const chars = t.replace(/ /g, '');
  // 字距排版：幾乎每個字後面都跟一個空白
  const spacedOut = chars.length >= 3 && (t.match(/ /g) || []).length >= chars.length * 0.6;
  const wide = `[${CJK}${CJK_PUNCT}]`;
  // 中文、全形標點旁邊的空白都不需要
  t = t.replace(new RegExp(`(${wide}) (?=${wide})`, 'g'), '$1');
  t = t.replace(new RegExp(`(${wide}) (?=[0-9(（)）.,:;%/\\-])`, 'g'), '$1');
  t = t.replace(new RegExp(`([0-9(（)）.,:;%/\\-]) (?=${wide})`, 'g'), '$1');
  if (spacedOut) {
    // 「8 1」「0 2」這種被拆開的數字也接回去
    for (let i = 0; i < 3; i++) t = t.replace(/([0-9(（)）.\-:/]) (?=[0-9(（)）.\-:/])/g, '$1');
  }
  if (spacedOut) {
    // 字距排版裡的網址、英文代號：「h t t p :// w w w . s e f . o r g . t w」→「http://www.sef.org.tw」
    const parts = t.split(' ');
    const shortLatin = (x: string) => /^[A-Za-z0-9.:/：@_\-]{1,4}$/.test(x);
    const merged: string[] = [];
    let prevShort = false;
    for (const part of parts) {
      const isShort = shortLatin(part);
      // 前一個也是短的英數片段 → 接起來（「w w w .」→「www.」）
      if (merged.length && prevShort && /^[A-Za-z0-9.:/：@_\-]/.test(part) &&
          (isShort || /^[A-Za-z0-9]{1,2}[^A-Za-z0-9]/.test(part))) {
        merged[merged.length - 1] += part;
      } else {
        merged.push(part);
      }
      prevShort = isShort;
    }
    t = merged.join(' ');
  }
  // 郵遞區號、電話被拆成「3 2 0 2 1 6」：三個以上的單一數字連在一起才接
  t = t.replace(/(?<![0-9])\d(?: \d){2,}(?![0-9])/g, m => m.replace(/ /g, ''));
  // 括號內側的空白
  t = t.replace(/([(（]) /g, '$1').replace(/ ([)）])/g, '$1');
  return t;
}

/** 新的一項開頭：(一)、一、1.、□、※、註： …… 這種行不要接到上一段 */
export function startsNewItem(text: string): boolean {
  const t = text.trim();
  return (
    /^[(（]?\s*[一二三四五六七八九十]+\s*[)）、]/.test(t) ||
    /^[(（]\s*\d{1,2}\s*[)）]/.test(t) ||
    /^\d{1,2}\s*[.、．)）]/.test(t) ||
    /^[□■☐☑✓✔○●◎※✽＊*•‧·▪►▶◆◇\-–—]/.test(t) ||
    /^(註|備註|說明|附註)\s*[:：\d]/.test(t) ||
    PUA_RE.test(t.charAt(0))
  );
}

const joinText = (a: string, b: string): string => {
  if (!a) return b;
  if (!b) return a;
  const endLatin = /[A-Za-z0-9]$/.test(a);
  const startLatin = /^[A-Za-z0-9]/.test(b);
  return a + (endLatin && startLatin ? ' ' : '') + b;
};

const unionBox = (boxes: Box[]): Box => ({
  left: Math.min(...boxes.map(b => b.left)),
  top: Math.min(...boxes.map(b => b.top)),
  right: Math.max(...boxes.map(b => b.right)),
  bottom: Math.max(...boxes.map(b => b.bottom)),
});

// ── 格線 ─────────────────────────────────────────────────────

/** 兩點之間有沒有直的格線（在 y 這個高度） */
const vLineBetween = (grid: GridLines | undefined, x0: number, x1: number, y: number): boolean =>
  !!grid?.v.some(v => v.x > x0 - 0.5 && v.x < x1 + 0.5 && v.y0 - 1 <= y && v.y1 + 1 >= y);

/** 找出一個範圍所在的表格格子；找不到的邊就用頁面邊界 */
export function findCell(box: Box, grid: GridLines | undefined, pageW: number, pageH: number): Cell {
  const cy = (box.top + box.bottom) / 2;
  const cx = (box.left + box.right) / 2;
  // 字框凸出格線一點點的還是算這一格：Word 預設讓行尾的「，。、」凸出邊界，
  // PDF 給的又是整個字寬（標點的筆畫只在左半邊），一整行就會看起來「壓到」右邊的格線。
  // 沒有這個容許值的話，那一行會被當成橫跨兩欄：自成一段、字級跟同格不一樣、還會畫到隔壁欄去。
  const tolX = Math.min(8, Math.max(1, 0.6 * (box.bottom - box.top)));
  let left = -Infinity, right = Infinity, top = -Infinity, bottom = Infinity;
  for (const v of grid?.v || []) {
    if (v.y0 - 1 > cy || v.y1 + 1 < cy) continue;
    if (v.x <= box.left + tolX && v.x < cx && v.x > left) left = v.x;
    if (v.x >= box.right - tolX && v.x > cx && v.x < right) right = v.x;
  }
  for (const h of grid?.h || []) {
    if (h.x0 - 1 > cx || h.x1 + 1 < cx) continue;
    if (h.y <= box.top + 1 && h.y > top) top = h.y;
    if (h.y >= box.bottom - 1 && h.y < bottom) bottom = h.y;
  }
  return {
    left: Number.isFinite(left) ? left : 0,
    right: Number.isFinite(right) ? right : pageW,
    top: Number.isFinite(top) ? top : 0,
    bottom: Number.isFinite(bottom) ? bottom : pageH,
    hasLeft: Number.isFinite(left),
    hasRight: Number.isFinite(right),
    hasTop: Number.isFinite(top),
    hasBottom: Number.isFinite(bottom),
  };
}

/**
 * 原文在格子裡置中：左右留白差不多，而且留白夠多。
 * 兩邊對齊、寫滿整格的段落左右留白也一樣，但很窄——那不算置中。
 */
export const isCenteredIn = (source: Box, c: Cell) => {
  if (!c.hasLeft || !c.hasRight) return false;
  const lm = source.left - c.left;
  const rm = c.right - source.right;
  return Math.abs(lm - rm) < 3 && lm > Math.max(4, (c.right - c.left) * 0.06);
};

/** 同一欄（左右格線一樣）而且上下之間沒有橫線隔開 */
const sameColumnNoRule = (a: Cell, b: Cell, upper: Box, lower: Box, grid?: GridLines) => {
  if (Math.abs(a.left - b.left) >= 1 || Math.abs(a.right - b.right) >= 1) return false;
  const x0 = Math.max(upper.left, lower.left);
  const x1 = Math.min(upper.right, lower.right);
  const mid = x1 > x0 ? (x0 + x1) / 2 : (upper.left + lower.left) / 2;
  return !grid?.h.some(h =>
    h.y > upper.bottom - 1.5 && h.y < lower.top + 1.5 && h.x0 - 1 <= mid && h.x1 + 1 >= mid
  );
};

/**
 * 從算好的頁面圖片找表格線。
 * 先把文字的位置挖掉（字的筆畫不會被當成線），再找夠長的水平／垂直深色線段。
 */
export function detectGridLines(
  rgba: Uint8ClampedArray | Uint8Array,
  width: number,
  height: number,
  scale: number,
  textBoxes: Box[] = [],
): GridLines {
  const dark = new Uint8Array(width * height);
  for (let i = 0, p = 0; i < dark.length; i++, p += 4) {
    const a = rgba[p + 3];
    if (a < 128) continue;
    const lum = 0.299 * rgba[p] + 0.587 * rgba[p + 1] + 0.114 * rgba[p + 2];
    if (lum < 165) dark[i] = 1;
  }
  // 文字位置挖掉
  for (const b of textBoxes) {
    const x0 = Math.max(0, Math.floor(b.left * scale) + 1);
    const x1 = Math.min(width, Math.ceil(b.right * scale) - 1);
    const y0 = Math.max(0, Math.floor(b.top * scale) + 1);
    const y1 = Math.min(height, Math.ceil(b.bottom * scale) - 1);
    for (let y = y0; y < y1; y++) dark.fill(0, y * width + x0, y * width + Math.max(x0, x1));
  }

  const minH = Math.round(14 * scale);
  const minV = Math.round(8 * scale);
  const bridge = Math.round(2 * scale);   // 線段中間小缺口（交叉點、被字擦到）接起來

  const hRuns: { y: number; x0: number; x1: number }[] = [];
  for (let y = 0; y < height; y++) {
    const row = y * width;
    let start = -1, lastDark = -1;
    for (let x = 0; x <= width; x++) {
      const d = x < width && dark[row + x];
      if (d) {
        if (start < 0) start = x;
        lastDark = x;
      } else if (start >= 0 && x - lastDark > bridge) {
        if (lastDark - start + 1 >= minH) hRuns.push({ y, x0: start, x1: lastDark });
        start = -1;
      }
    }
  }
  const vRuns: { x: number; y0: number; y1: number }[] = [];
  for (let x = 0; x < width; x++) {
    let start = -1, lastDark = -1;
    for (let y = 0; y <= height; y++) {
      const d = y < height && dark[y * width + x];
      if (d) {
        if (start < 0) start = y;
        lastDark = y;
      } else if (start >= 0 && y - lastDark > bridge) {
        if (lastDark - start + 1 >= minV) vRuns.push({ x, y0: start, y1: lastDark });
        start = -1;
      }
    }
  }

  // 粗線會連續好幾排／好幾欄，併成一條
  const mergeRuns = <T extends Record<string, number>>(runs: T[], pos: keyof T, a: keyof T, b: keyof T) => {
    const sorted = [...runs].sort((p, q) => (p[pos] - q[pos]) || (p[a] - q[a]));
    const out: { pos: number; a: number; b: number; n: number; sum: number }[] = [];
    for (const r of sorted) {
      const hit = out.find(o => r[pos] - (o.pos) <= 2 && r[a] <= o.b + bridge && r[b] >= o.a - bridge);
      if (hit) {
        hit.pos = r[pos];
        hit.a = Math.min(hit.a, r[a]);
        hit.b = Math.max(hit.b, r[b]);
        hit.n++;
        hit.sum += r[pos];
      } else {
        out.push({ pos: r[pos], a: r[a], b: r[b], n: 1, sum: r[pos] });
      }
    }
    return out.map(o => ({ pos: o.sum / o.n, a: o.a, b: o.b }));
  };

  const h = mergeRuns(hRuns, 'y', 'x0', 'x1')
    .map(o => ({ y: (o.pos + 0.5) / scale, x0: o.a / scale, x1: (o.b + 1) / scale }));
  const v = mergeRuns(vRuns, 'x', 'y0', 'y1')
    .map(o => ({ x: (o.pos + 0.5) / scale, y0: o.a / scale, y1: (o.b + 1) / scale }));
  return { h, v };
}

// ── 片段 → 格子裡的一段文字 ───────────────────────────────────

interface Piece extends Box { str: string; fontSize: number; baseline: number }

const toPieces = (items: RawTextItem[], pageH: number): Piece[] => {
  const out: Piece[] = [];
  for (const it of items) {
    if (!it || typeof it.str !== 'string' || !it.str.trim()) continue;
    const [a, b, c, d, e, f] = it.transform;
    // 旋轉、上下顛倒、左右鏡像的字不處理（保留原樣）
    if (Math.abs(b) > 0.01 || Math.abs(c) > 0.01 || a <= 0 || d <= 0) continue;
    const fontSize = Math.abs(d) || Math.abs(a) || 10;
    const baseline = pageH - f;
    const w = typeof it.width === 'number' && it.width > 0 ? it.width : fontSize * it.str.length * 0.5;
    out.push({
      str: it.str,
      fontSize,
      baseline,
      left: e,
      right: e + w,
      top: baseline - fontSize * 0.9,
      bottom: baseline + fontSize * 0.18,
    });
  }
  return out;
};

/** 把文字片段的位置（給格線偵測挖掉文字用） */
export const textItemBoxes = (items: RawTextItem[], pageH: number): Box[] =>
  toPieces(items, pageH).map(({ left, top, right, bottom }) => ({ left, top, right, bottom }));

/** 片段 → 一列一列 → 依空隙和直線切成一格一格 */
export function segmentsFromTextItems(
  items: RawTextItem[],
  pageW: number,
  pageH: number,
  grid?: GridLines,
): Segment[] {
  const pieces = toPieces(items, pageH).sort((p, q) => p.baseline - q.baseline || p.left - q.left);

  // 1) 同一條基準線的片段放同一列
  const rows: Piece[][] = [];
  const rowBase: number[] = [];
  for (const p of pieces) {
    let idx = -1;
    for (let r = rows.length - 1; r >= 0 && r >= rows.length - 6; r--) {
      const fs = Math.min(p.fontSize, rows[r][0].fontSize);
      if (Math.abs(rowBase[r] - p.baseline) <= fs * 0.45) { idx = r; break; }
    }
    if (idx < 0) { rows.push([p]); rowBase.push(p.baseline); }
    else rows[idx].push(p);
  }

  // 2) 每一列先依格線分成一格一格，再依空隙、獨立符號切開
  const segments: Segment[] = [];
  const singleCjk = (p: Piece) => /^[㐀-鿿]$/.test(visibleChars(p.str));
  for (const fullRow of rows) {
    fullRow.sort((p, q) => p.left - q.left);
    const cellGroups: Piece[][] = [];
    for (const p of fullRow) {
      const g = cellGroups[cellGroups.length - 1];
      const prev = g?.[g.length - 1];
      const cy = prev ? (Math.min(prev.top, p.top) + Math.max(prev.bottom, p.bottom)) / 2 : 0;
      if (prev && !vLineBetween(grid, prev.right - 1, p.left + 1, cy)) g.push(p);
      else cellGroups.push([p]);
    }
    for (const row of cellGroups) {
    // 分散對齊的標題（「地　　址」「臺 灣 地 區 內」）：整格都是單一個字 → 不管空隙多大都當同一段
    const distributed = row.length >= 2 && row.every(singleCjk);
    let cur: Piece[] = [];
    const flush = () => {
      if (!cur.length) return;
      let text = '';
      for (let i = 0; i < cur.length; i++) {
        const gap = i ? cur[i].left - cur[i - 1].right : 0;
        const sep = i && gap > cur[i].fontSize * 0.15 ? ' ' : '';
        text += sep + cur[i].str;
      }
      const box = unionBox(cur);
      segments.push({
        ...box,
        text,
        fontSize: Math.max(...cur.map(p => p.fontSize)),
        glyphs: cur.map(({ left, top, right, bottom }) => ({ left, top, right, bottom })),
      });
      cur = [];
    };
    for (const p of row) {
      const prev = cur[cur.length - 1];
      if (prev) {
        const fs = Math.max(prev.fontSize, p.fontSize);
        const gap = p.left - prev.right;
        const cy = (Math.min(prev.top, p.top) + Math.max(prev.bottom, p.bottom)) / 2;
        if (
          (!distributed && gap > fs * LAYOUT.splitGapEm) ||
          vLineBetween(grid, prev.right - 1, p.left + 1, cy) ||
          isStandaloneSymbol(p.str) || isStandaloneSymbol(prev.str)
        ) flush();
      }
      cur.push(p);
    }
    flush();
    }
  }
  return joinSpreadPhrases(segments);
}

/** 拆成一個字一段、但其實是固定詞的（例：日期列的「中　華　民　國」） */
const SPREAD_PHRASES = ['中華民國'];

function joinSpreadPhrases(segments: Segment[]): Segment[] {
  const out: Segment[] = [];
  const byRow = [...segments].sort((a, b) => a.top - b.top || a.left - b.left);
  for (let i = 0; i < byRow.length; i++) {
    let joined = false;
    for (const phrase of SPREAD_PHRASES) {
      const run = byRow.slice(i, i + phrase.length);
      if (
        run.length === phrase.length &&
        run.every((s, k) => visibleChars(s.text) === phrase[k] && Math.abs(s.top - run[0].top) < s.fontSize * 0.5)
      ) {
        out.push({
          ...unionBox(run),
          text: phrase,
          fontSize: Math.max(...run.map(s => s.fontSize)),
          glyphs: run.flatMap(s => s.glyphs),
        });
        i += phrase.length - 1;
        joined = true;
        break;
      }
    }
    if (!joined) out.push(byRow[i]);
  }
  return out;
}

// ── 一格一格 → 一段一段 → 譯文範圍 ─────────────────────────────

const visibleChars = (s: string) => s.replace(/\s/g, '');

/**
 * 直排雙欄標籤：「保矯／險正／對機／象關／收／容／人」其實是「保險對象」「矯正機關收容人」兩欄；
 * 「申簽／請章／人」是「申請人」「簽章」。
 * 每行剛好兩個中文字、左邊對齊，連續 3 行以上，或 2 行後面還接著單一個字 → 改成一欄一欄讀。
 * （只有兩行、後面沒有單字的，例如「就醫／情形」，照一般橫的讀）
 */
function readStackedColumns(lines: Segment[]): string {
  const out: string[] = [];
  let i = 0;
  while (i < lines.length) {
    let j = i;
    while (
      j < lines.length &&
      /^[㐀-鿿]{2}$/.test(visibleChars(lines[j].text)) &&
      Math.abs(lines[j].left - lines[i].left) < 1.5
    ) j++;
    let k = j;
    while (k < lines.length && /^[㐀-鿿]$/.test(visibleChars(lines[k].text))) k++;
    const rows = j - i;
    if (rows >= 3 || (rows >= 2 && k > j)) {
      const run = lines.slice(i, j);
      let col1 = run.map(r => visibleChars(r.text)[0]).join('');
      let col2 = run.map(r => visibleChars(r.text)[1]).join('');
      // 後面的單字行看它對齊哪一欄
      const l0 = run[0].left;
      const w = Math.max(...run.map(r => r.right - r.left));
      for (const t of lines.slice(j, k)) {
        const c = (t.left + t.right) / 2;
        if (Math.abs(c - (l0 + w * 0.25)) <= Math.abs(c - (l0 + w * 0.75))) col1 += visibleChars(t.text);
        else col2 += visibleChars(t.text);
      }
      out.push(`${col1}／${col2}`);
      i = k;
    } else {
      out.push(lines[i].text);
      i++;
    }
  }
  return out.reduce(joinText, '');
}

/**
 * 被填寫空格切開的句子接回去：
 * 「保險對象因＿＿＿，無法親自辦理……特委託＿＿＿代辦。」原本會變成三段分開翻，
 * 翻出來的英文接不起來。同一列、同一格，後面那段是以標點開頭（或是很短的結尾「代辦。」）
 * 就接成一句，空格用「＿＿」表示。
 */
function joinBlankFragments(list: { seg: Segment; cell: Cell }[]): { seg: Segment; cell: Cell }[] {
  const items = [...list].sort((p, q) => p.seg.top - q.seg.top || p.seg.left - q.seg.left);
  let changed = true;
  while (changed) {
    changed = false;
    for (let i = 0; i < items.length && !changed; i++) {
      const a = items[i];
      if (!isTranslatable(a.seg.text)) continue;
      const fs = a.seg.fontSize;
      const sameRow = (o: Segment) =>
        Math.abs(o.top - a.seg.top) < fs * 0.5 && Math.abs(o.bottom - a.seg.bottom) < fs * 0.5;
      // 右邊最近的一段（同一列）
      let b: (typeof items)[number] | undefined;
      for (const o of items) {
        if (o === a || !sameRow(o.seg) || o.seg.left < a.seg.right - 0.5) continue;
        if (!b || o.seg.left < b.seg.left) b = o;
      }
      if (!b || !isTranslatable(b.seg.text)) continue;
      if (Math.abs(a.cell.left - b.cell.left) >= 1 || Math.abs(a.cell.right - b.cell.right) >= 1) continue;
      const bt = b.seg.text.trim();
      const continues = /^[，,、；;：:）)]/.test(bt) || (visibleChars(bt).length <= 8 && /[。.]/.test(bt));
      if (!continues) continue;
      const merged: Segment = {
        ...unionBox([a.seg, b.seg]),
        text: `${a.seg.text}＿＿${b.seg.text}`,
        fontSize: Math.max(a.seg.fontSize, b.seg.fontSize),
        glyphs: [...a.seg.glyphs, ...b.seg.glyphs],
      };
      items.splice(items.indexOf(b), 1);
      items[items.indexOf(a)] = { seg: merged, cell: a.cell };
      changed = true;
    }
  }
  return items;
}

export function layoutSegments(
  segments: Segment[],
  pageW: number,
  pageH: number,
  grid?: GridLines,
): LayoutBlock[] {
  const segs = joinBlankFragments(
    segments.map(s => ({ seg: s, cell: findCell(s, grid, pageW, pageH) }))
  ).sort((p, q) => p.seg.top - q.seg.top || p.seg.left - q.seg.left);

  // 沒有格線時，用「長行」右端的中位數當作「寫滿」的參考（不用最大值，避免被一兩行凸出去的拉走）
  const longRights = segs
    .filter(s => s.seg.right - s.seg.left > pageW * 0.55)
    .map(s => s.seg.right)
    .sort((a, b) => a - b);
  const contentRight = longRights.length
    ? longRights[Math.floor(longRights.length / 2)]
    : (segs.length ? Math.max(...segs.map(s => s.seg.right)) : pageW);
  const longLefts = segs
    .filter(s => s.seg.right - s.seg.left > pageW * 0.55)
    .map(s => s.seg.left)
    .sort((a, b) => a - b);
  const contentLeft = longLefts.length ? longLefts[Math.floor(longLefts.length / 2)] : LAYOUT.pageMargin;

  // 1) 接成段落
  type Para = { lines: Segment[]; cell: Cell };
  const paras: Para[] = [];
  const open: Para[] = [];

  for (const { seg, cell } of segs) {
    if (!isTranslatable(seg.text)) continue;
    let target: Para | undefined;
    for (const p of open) {
      const last = p.lines[p.lines.length - 1];
      const fs = Math.max(last.fontSize, seg.fontSize);
      if (!sameColumnNoRule(p.cell, cell, last, seg, grid)) continue;
      const vGap = seg.top - last.bottom;
      if (vGap > fs * LAYOUT.paragraphGapEm || vGap < -fs * 0.4) continue;
      if (Math.abs(last.fontSize - seg.fontSize) > fs * 0.25) continue;
      const first = p.lines[0];
      if (seg.left < first.left - fs * 1.5 || seg.left > first.left + fs * 3) continue;
      if (startsNewItem(seg.text)) continue;
      // 上一行以句號結束、下一行很短：是另一個欄位標籤（例：「……代辦。」下面的「委託人簽章」）
      if (/[。.]$/.test(last.text.trim()) && visibleChars(seg.text).length <= 8) continue;
      const rightLimit = cell.hasRight ? cell.right : contentRight;
      // 下一行只剩一個字（「□5.其」換行到「他」）也算接續
      const orphan = /^[\u3400-\u9fff]$/.test(visibleChars(seg.text)) &&
        Math.abs(seg.left - first.left) < fs * 1.5 && !/[。：:；;，,）)]$/.test(last.text.trim());
      if (!orphan && last.right < rightLimit - fs * LAYOUT.fullLineEm) continue;
      target = p;
      break;
    }
    if (target) {
      target.lines.push(seg);
    } else {
      const p = { lines: [seg], cell };
      paras.push(p);
      open.push(p);
    }
    // 離太遠的段落不用再考慮
    for (let i = open.length - 1; i >= 0; i--) {
      const last = open[i].lines[open[i].lines.length - 1];
      if (seg.top - last.bottom > last.fontSize * 3) open.splice(i, 1);
    }
  }

  // 2) 每一段的譯文範圍
  //    先決定每一段往右能寫到哪裡（到格線或右邊下一段），
  //    再決定往下能寫到哪裡（到格線、或下面那一段「往右延伸後」的範圍），
  //    這樣兩段的譯文範圍不會互相重疊。
  const all = segs.map(s => s.seg);
  const pad = LAYOUT.pad;

  // 這一頁一般內文的字級：標籤譯文不要比它大太多，整頁看起來才一致
  const sizes = paras.map(p => p.lines[0].fontSize).sort((a, b) => a - b);
  const typical = sizes.length ? sizes[Math.floor(sizes.length / 2)] : 10;

  type Draft = {
    p: Para; source: Box; fontSize: number; own: Set<Segment>;
    left: number; top: number; right: number; vertical: boolean; alignRight: boolean; alignCenter: boolean;
  };
  const drafts: Draft[] = paras.map(p => {
    const source = unionBox(p.lines);
    const fontSize = p.lines.map(l => l.fontSize).sort((a, b) => a - b)[Math.floor(p.lines.length / 2)];
    const { cell } = p;
    const own = new Set(p.lines);
    let right = cell.hasRight ? cell.right - pad : pageW - LAYOUT.pageMargin;
    for (const o of all) {
      if (own.has(o)) continue;
      const vOverlap = Math.min(o.bottom, source.bottom) - Math.max(o.top, source.top);
      if (vOverlap > 0 && o.left >= source.right - 0.5 && o.left - pad < right) right = o.left - pad;
    }
    right = Math.max(right, source.right);
    // 有格線的話，譯文絕不寫過格線（原文凸出去的行尾標點不算）
    if (cell.hasRight) right = Math.min(right, cell.right - pad);

    const width = source.right - source.left;
    const height = source.bottom - source.top;
    // 窄而高的格子（左邊那排直排標籤）：譯文轉 90 度寫才放得下
    const cellW = cell.right - cell.left;
    const cellH = cell.bottom - cell.top;
    const narrowCell = cell.hasLeft && cell.hasRight && cell.hasTop && cell.hasBottom &&
      cellW < Math.max(fontSize * 3, 26) && cellH > cellW * 2;
    const vertical = p.lines.length >= 2 &&
      (narrowCell || (p.lines.length >= 3 && width < fontSize * 2.6 && height > width * 2.5));

    // 原文在格子裡置中的（表頭、「陳廠長」這種）：譯文可以用整個格子的寬度，
    // 不然窄欄的「備註」只剩原文那一點寬，英文 Remarks 會被拆成兩半
    let left = Math.max(source.left, cell.hasLeft ? cell.left + pad : 0);
    let alignRight = false;
    let alignCenter = false;
    const sameLineNeighbor = all.some(o => !own.has(o) &&
      Math.min(o.bottom, source.bottom) - Math.max(o.top, source.top) > 0);
    // 表格外置中的標題（例：「會議記錄表」）：每一行的中心都在頁面（或版心）正中間，同一行沒有別的字
    const centerOf = (l: { left: number; right: number }) => (l.left + l.right) / 2;
    const tol = Math.max(3, fontSize * 0.6);
    const pageCenters = [pageW / 2, (contentLeft + contentRight) / 2];
    const centeredOnPage = !vertical && !cell.hasLeft && !cell.hasRight && !sameLineNeighbor &&
      source.right - source.left < (contentRight - contentLeft) * 0.75 &&
      pageCenters.some(c0 => p.lines.every(l => Math.abs(centerOf(l) - c0) < tol));
    if (centeredOnPage) {
      const c0 = centerOf(source);
      const half = Math.max(source.right - c0, Math.min(c0 - LAYOUT.pageMargin, pageW - LAYOUT.pageMargin - c0));
      left = c0 - half;
      right = c0 + half;
      alignCenter = true;
    } else if (isCenteredIn(source, cell)) {
      left = cell.left + pad;
      for (const o of all) {
        if (own.has(o)) continue;
        const vOverlap = Math.min(o.bottom, source.bottom) - Math.max(o.top, source.top);
        if (vOverlap > 0 && o.right <= source.left + 0.5 && o.right + pad > left) left = o.right + pad;
      }
      left = Math.min(left, source.left);
    } else if (p.lines.length === 1 && !vertical) {
      // 靠右的一行字，左邊同一行沒有別的字：譯文往左延伸、靠右對齊，不用擠在原文那一小段寬度裡換好幾行
      const rightEdge = cell.hasRight ? cell.right - pad : contentRight;
      const leftLimit = cell.hasLeft ? cell.left + pad : LAYOUT.pageMargin;
      const width = source.right - source.left;
      const hasLeftNeighbor = all.some(o => !own.has(o) &&
        Math.min(o.bottom, source.bottom) - Math.max(o.top, source.top) > 0 && o.right <= source.left + 0.5);
      if (!hasLeftNeighbor && source.right >= rightEdge - fontSize * 1.5 && source.left - leftLimit > width) {
        left = leftLimit;
        alignRight = true;
      }
    }

    return {
      p, source, fontSize, own, right, vertical, left, alignRight, alignCenter,
      top: Math.max(source.top, cell.hasTop ? cell.top + pad : 0),
    };
  });

  const blocks: LayoutBlock[] = [];
  for (const d of drafts) {
    const { p, source, fontSize, own, left, top, right, vertical, alignRight, alignCenter } = d;
    const { cell } = p;

    const text = cleanCjkSpacing(readStackedColumns(p.lines));
    if (!isTranslatable(text)) continue;

    let box: Box;
    if (vertical) {
      // 直排標籤用整個格子，但要讓開格子裡的其他東西（例如上面的圓圈編號）
      let vTop = cell.hasTop ? cell.top + pad : source.top;
      let vBottom = cell.hasBottom ? cell.bottom - pad : source.bottom;
      for (const o of all) {
        if (own.has(o)) continue;
        const hOverlap = Math.min(o.right, source.right) - Math.max(o.left, source.left);
        if (hOverlap <= 0) continue;
        if (o.bottom <= source.top + 0.5 && o.bottom + pad > vTop) vTop = o.bottom + pad;
        if (o.top >= source.bottom - 0.5 && o.top - pad < vBottom) vBottom = o.top - pad;
      }
      box = {
        left: cell.hasLeft ? cell.left + pad : source.left - 1,
        right: cell.hasRight ? cell.right - pad : source.right + 1,
        top: Math.min(vTop, source.top),
        bottom: Math.max(vBottom, source.bottom),
      };
    } else {
      let bottom = cell.hasBottom ? cell.bottom - pad : pageH - LAYOUT.pageMargin;
      // 下面的原文（不翻的符號、數字也算）
      for (const o of all) {
        if (own.has(o)) continue;
        const hOverlap = Math.min(o.right, right) - Math.max(o.left, left);
        if (hOverlap > 0 && o.top >= source.bottom - 0.5 && o.top - pad < bottom) bottom = o.top - pad;
      }
      // 下面那一段的譯文範圍（往右延伸後的）
      for (const o of drafts) {
        if (o === d || o.vertical) continue;
        const hOverlap = Math.min(o.right, right) - Math.max(o.source.left, left);
        if (hOverlap > 0 && o.source.top >= source.bottom - 0.5 && o.source.top - pad < bottom) {
          bottom = o.source.top - pad;
        }
      }
      box = { left, top, right, bottom: Math.max(bottom, source.bottom) };
    }

    // 標題（明顯比內文大的）保留原本字級；其他的最大不超過內文的 1.15 倍
    const maxFont = fontSize >= typical * 1.35 ? fontSize : Math.min(fontSize, typical * 1.15);

    blocks.push({
      text,
      fontSize,
      maxFont,
      glyphs: p.lines.flatMap(l => l.glyphs),
      source,
      box,
      cell: cell.hasLeft || cell.hasRight || cell.hasTop || cell.hasBottom ? cell : undefined,
      vertical,
      ...(alignRight ? { alignRight: true } : {}),
      ...(alignCenter ? { alignCenter: true } : {}),
    });
  }

  markStacks(blocks, all);
  return blocks;
}

/**
 * 找出「同一格裡由上往下排的清單」：同一個格子、項目之間沒有並排、左邊大致對齊，
 * 而且整個範圍裡沒有其他東西（勾選框、數字之類不翻的原文）。
 */
function markStacks(blocks: LayoutBlock[], all: Segment[]) {
  const byCell = new Map<string, LayoutBlock[]>();
  for (const b of blocks) {
    const c = b.cell;
    if (b.vertical || !c || !c.hasLeft || !c.hasRight || !c.hasTop) continue;
    const key = `${Math.round(c.left)},${Math.round(c.right)},${Math.round(c.top)}`;
    if (!byCell.has(key)) byCell.set(key, []);
    byCell.get(key)!.push(b);
  }
  let id = 0;
  for (const list of byCell.values()) {
    if (list.length < 2) continue;
    list.sort((a, b) => a.source.top - b.source.top);
    const fs = Math.max(...list.map(b => b.fontSize));
    const firstLeft = list[0].source.left;
    let ok = list.every((b, i) =>
      Math.abs(b.source.left - firstLeft) <= fs * 2 &&
      (i === 0 || b.source.top >= list[i - 1].source.bottom - 1)
    );
    if (!ok) continue;
    const region: Box = {
      left: Math.min(...list.map(b => b.box.left)),
      right: Math.min(...list.map(b => b.box.right)),
      top: Math.min(...list.map(b => b.box.top)),
      bottom: Math.max(...list.map(b => b.box.bottom)),
    };
    // 範圍裡有其他原文（不是這些項目自己的字）就不重排，免得蓋到
    const ownGlyph = (o: Segment) => o.glyphs.every(g => list.some(b => b.glyphs.includes(g)));
    ok = !all.some(o =>
      !ownGlyph(o) &&
      Math.min(o.right, region.right) - Math.max(o.left, region.left) > 0.5 &&
      Math.min(o.bottom, region.bottom) - Math.max(o.top, region.top) > 0.5
    );
    if (!ok) continue;
    id++;
    for (const b of list) b.stack = { id, region };
  }
}

/** 有文字層的頁面：片段 → 翻譯區塊 */
export function layoutTextItems(
  items: RawTextItem[],
  pageW: number,
  pageH: number,
  grid?: GridLines,
): LayoutBlock[] {
  return layoutSegments(segmentsFromTextItems(items, pageW, pageH, grid), pageW, pageH, grid);
}

// ── 同一格、同一排用同一個字級 ──────────────────────────────

/**
 * 哪些區塊的譯文要用同一個字級（回傳區塊索引的分組，只回傳兩個以上的組）。
 * - 同一個格子裡的（左右格線相同、上下範圍重疊）：例如「檢附書據」那格的 □1～□5
 * - 表格同一排、一排高的格子（上緣是同一條橫線）：例如表頭「臺灣地區內／臺灣地區外／住院部分負擔超過上限」
 * - 並排的清單欄（同一排、各自是一串 □1～□5 的格子）
 * 原文字級差很多的（標題和註解）不放同一組；直排標籤、不在表格裡的段落不分組。
 * 每段各自配合空間選字級的話，短的字大、長的字小，同一格擺在一起會忽大忽小。
 */
type FontGroupBlock = Pick<LayoutBlock, 'cell' | 'fontSize' | 'vertical' | 'stack'> & { text?: string; source?: Box };

export function fontGroups(blocks: FontGroupBlock[]): number[][] {
  return fontGroupPartitions(blocks).groups;
}

/**
 * 同一格／同一排／並排清單／同一欄標籤要用同一個字級的分組。
 * strict：其中「一定要一樣」的部分（同一格、並排清單）——同一排、同一欄的只是盡量一樣，
 * 某一格特別擠（例：窄窄的「備註」）時，整排不必跟著縮到那麼小。
 */
export function fontGroupPartitions(blocks: FontGroupBlock[]): { groups: number[][]; strict: number[][] } {
  const parent = blocks.map((_, i) => i);
  const find = (i: number): number => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  const join = (a: number, b: number) => { parent[find(a)] = find(b); };
  const sParent = blocks.map((_, i) => i);
  const sFind = (i: number): number => (sParent[i] === i ? i : (sParent[i] = sFind(sParent[i])));

  const inTable = (b: (typeof blocks)[number]) => !b.vertical && !!b.cell && b.cell.hasLeft && b.cell.hasRight;
  const short = (blk: (typeof blocks)[number]) => !!blk.text && blk.text.length <= 10;
  // 表頭列：同一排（上下格線都一樣）的格子全部都是短標籤，至少三格
  const bandKey = (c: Cell) => `${Math.round(c.top)}|${Math.round(c.bottom)}`;
  const bandStats = new Map<string, { n: number; allShort: boolean }>();
  for (const b of blocks) {
    if (!inTable(b) || !b.cell!.hasTop || !b.cell!.hasBottom) continue;
    const st = bandStats.get(bandKey(b.cell!)) || { n: 0, allShort: true };
    st.n++;
    st.allShort = st.allShort && short(b);
    bandStats.set(bandKey(b.cell!), st);
  }
  const headerBand = (c: Cell) => {
    const st = c.hasTop && c.hasBottom ? bandStats.get(bandKey(c)) : undefined;
    return !!st && st.allShort && st.n >= 3;
  };
  // 表格外、同一行並排的短標籤（簽名欄「呈董事長：　總經理：　陳廠長：」）
  const lineLabel = (b: (typeof blocks)[number]) =>
    !b.vertical && !inTable(b) && !!b.source && !!b.text && b.text.length <= 20 &&
    b.source.bottom - b.source.top <= b.fontSize * 1.6;
  for (let i = 0; i < blocks.length; i++) {
    const a = blocks[i];
    if (lineLabel(a)) {
      for (let j = i + 1; j < blocks.length; j++) {
        const b = blocks[j];
        if (!lineLabel(b)) continue;
        const ratio = Math.max(a.fontSize, b.fontSize) / Math.min(a.fontSize, b.fontSize);
        if (ratio <= 1.25 && Math.abs(a.source!.top - b.source!.top) < 2 && Math.abs(a.source!.bottom - b.source!.bottom) < 2) join(i, j);
      }
      continue;
    }
    if (!inTable(a)) continue;
    for (let j = i + 1; j < blocks.length; j++) {
      const b = blocks[j];
      if (!inTable(b)) continue;
      const ratio = Math.max(a.fontSize, b.fontSize) / Math.min(a.fontSize, b.fontSize);
      if (ratio > 1.25) continue;
      const ca = a.cell!, cb = b.cell!;
      const sameColumn = Math.abs(ca.left - cb.left) < 1 && Math.abs(ca.right - cb.right) < 1;
      const vOverlap = Math.min(ca.bottom, cb.bottom) - Math.max(ca.top, cb.top);
      const sameCell = sameColumn && vOverlap > 1;
      // 同一排：只限一排高的格子（表頭、單行欄位），高的格子內容各不相同，不硬拉成一樣
      const oneRow = (c: Cell, f: number) => c.hasBottom && c.bottom - c.top <= f * 2.8;
      // 比較高的表頭（兩行高）：整排都是短標籤的也算同一排（項次／決議事項／執行部門／備註）
      const sameBand = ca.hasTop && cb.hasTop && Math.abs(ca.top - cb.top) < 2 && (
        (oneRow(ca, a.fontSize) && oneRow(cb, b.fontSize)) ||
        (Math.abs(ca.bottom - cb.bottom) < 2 && headerBand(ca) && headerBand(cb))
      );
      // 並排的清單欄（例：「檢附書據」的臺灣地區內／外／住院部分負擔三欄）
      const parallelLists = !!a.stack && !!b.stack && Math.abs(ca.top - cb.top) < 2;
      // 同一欄往下排的欄位名稱（會議名稱／主席／出席人員／列席人員）：都是一排高的格子裡的短標籤
      const isLabel = (blk: typeof a, c: Cell) => oneRow(c, blk.fontSize) && short(blk);
      const labelColumn = sameColumn && !sameCell && isLabel(a, ca) && isLabel(b, cb);
      if (sameCell || sameBand || parallelLists || labelColumn) join(i, j);
      if (sameCell || parallelLists) sParent[sFind(i)] = sFind(j);
    }
  }

  const collect = (root: (i: number) => number) => {
    const m = new Map<number, number[]>();
    blocks.forEach((_, i) => {
      if (!inTable(blocks[i]) && !lineLabel(blocks[i])) return;
      const r = root(i);
      if (!m.has(r)) m.set(r, []);
      m.get(r)!.push(i);
    });
    return [...m.values()].filter(g => g.length >= 2);
  };
  return { groups: collect(find), strict: collect(sFind) };
}

/**
 * 分組後每一段的字級：
 *   同一格／並排清單（strict）→ 取最小的，一定一樣；
 *   同一排／同一欄 → 取最小的，但特別擠的那幾格（不到中位數七成）不算，它們自己用自己的小字級。
 * preferred：每一段自己放得下的最大字級（Infinity = 不參與）
 */
export function groupFontSizes(blocks: FontGroupBlock[], preferred: number[]): Map<number, number> {
  const { groups, strict } = fontGroupPartitions(blocks);
  const own = preferred.slice();
  for (const g of strict) {
    const size = Math.min(...g.map(k => preferred[k]));
    if (Number.isFinite(size)) for (const k of g) own[k] = size;
  }
  const out = new Map<number, number>();
  for (const g of groups) {
    const sizes = g.map(k => own[k]).filter(Number.isFinite).sort((a, b) => a - b);
    if (!sizes.length) continue;
    const median = sizes[Math.floor(sizes.length / 2)];
    const core = sizes.filter(v => v >= median * 0.7);
    const size = Math.min(...core);
    for (const k of g) if (Number.isFinite(own[k])) out.set(k, Math.min(size, own[k]));
  }
  for (const g of strict) for (const k of g) if (!out.has(k) && Number.isFinite(own[k])) out.set(k, own[k]);
  return out;
}

// ── 底色與字色 ──────────────────────────────────────────────

const hex = (r: number, g: number, b: number) =>
  '#' + [r, g, b].map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('');

/**
 * 從頁面圖片取出一段文字周圍的底色（蓋原文用）和字的顏色（寫譯文用）。
 * 有底色的表頭、紅字的提醒都能照原樣保留。
 */
export function sampleBlockColors(
  rgba: Uint8ClampedArray | Uint8Array | null,
  width: number,
  height: number,
  scale: number,
  block: Pick<LayoutBlock, 'glyphs' | 'cell'>,
): { background: string; ink: string } {
  const fallback = { background: '#ffffff', ink: '#111827' };
  if (!rgba || !block.glyphs.length) return fallback;

  const px = (x: number, y: number) => {
    const xi = Math.max(0, Math.min(width - 1, Math.round(x * scale)));
    const yi = Math.max(0, Math.min(height - 1, Math.round(y * scale)));
    const i = (yi * width + xi) * 4;
    return [rgba[i], rgba[i + 1], rgba[i + 2]];
  };
  const lum = (c: number[]) => 0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2];

  // 底色：字框外圍一圈，取最常出現的亮色
  const counts = new Map<string, { n: number; c: number[] }>();
  const cell = block.cell;
  const inside = (x: number, y: number) =>
    !cell || ((!cell.hasLeft || x > cell.left + 1) && (!cell.hasRight || x < cell.right - 1) &&
      (!cell.hasTop || y > cell.top + 1) && (!cell.hasBottom || y < cell.bottom - 1));
  for (const g of block.glyphs.slice(0, 12)) {
    const pts: [number, number][] = [];
    for (let k = 0; k <= 4; k++) {
      const x = g.left + ((g.right - g.left) * k) / 4;
      pts.push([x, g.top - 1.2], [x, g.bottom + 1.2]);
    }
    pts.push([g.left - 1.5, (g.top + g.bottom) / 2], [g.right + 1.5, (g.top + g.bottom) / 2]);
    for (const [x, y] of pts) {
      if (!inside(x, y)) continue;
      const c = px(x, y);
      if (lum(c) < 150) continue;
      const key = c.map(v => v >> 4).join(',');
      const hit = counts.get(key);
      if (hit) hit.n++;
      else counts.set(key, { n: 1, c });
    }
  }
  let background = fallback.background;
  let best = 0;
  for (const { n, c } of counts.values()) {
    if (n > best) { best = n; background = hex(c[0], c[1], c[2]); }
  }

  // 字色：字框裡面深色像素的平均
  let r = 0, g2 = 0, b = 0, n = 0;
  for (const g of block.glyphs.slice(0, 12)) {
    const step = 0.7;
    for (let y = g.top + 1; y < g.bottom - 0.5; y += step) {
      for (let x = g.left + 0.3; x < g.right - 0.3; x += step) {
        const c = px(x, y);
        if (lum(c) < 140) { r += c[0]; g2 += c[1]; b += c[2]; n++; }
      }
    }
  }
  let ink = fallback.ink;
  if (n > 8) {
    r /= n; g2 /= n; b /= n;
    const max = Math.max(r, g2, b), min = Math.min(r, g2, b);
    // 有明顯顏色（紅字、藍字）就沿用，加深一點比較好讀；灰黑色一律用深色
    ink = max - min > 60 ? hex(r * 0.85, g2 * 0.85, b * 0.85) : fallback.ink;
  }
  return { background, ink };
}

// ── 譯文換行與字級 ──────────────────────────────────────────

export interface Measurer { (text: string, fontSize: number): number }

/**
 * 整頁一般文字的字級上限：依字數加權的中位數（大部分文字實際用的字級）。
 * 大標題（原文明顯比內文大的）和直排標籤不算、也不受限。
 * 回傳 null 表示這頁段落太少，不用統一。
 */
export function pageTextSize(
  items: { size: number; sourceSize: number; vertical?: boolean; chars: number }[],
): { size: number; applies: boolean[] } | null {
  const src = items.map(x => x.sourceSize).filter(n => n > 0).sort((a, b) => a - b);
  const typical = src.length ? src[Math.floor(src.length / 2)] : 0;
  const applies = items.map(x => !x.vertical && x.chars > 0 && !(typical && x.sourceSize >= typical * 1.35));
  const pool = items
    .filter((x, k) => applies[k] && Number.isFinite(x.size))
    .sort((a, b) => a.size - b.size);
  if (pool.length < 4) return null;
  const total = pool.reduce((n, x) => n + x.chars, 0);
  let acc = 0;
  for (const x of pool) {
    acc += x.chars;
    if (acc >= total / 2) return { size: x.size, applies };
  }
  return null;
}

// ── 斷行單位 ───────────────────────────────────────────────
// 泰文、寮文、高棉文、緬甸文的字和字之間不空格，要用瀏覽器內建的字典斷詞（Intl.Segmenter），
// 不然會從一個字的中間斷開，甚至把母音、聲調符號跟它的字母拆到兩行。
const SEA_RE = /[\u0E00-\u0EFF\u1000-\u109F\u1780-\u17FF\u19E0-\u19FF]/;
const segmenters = new Map<string, any>();
const seaLocale = (s: string) =>
  /[\u0E80-\u0EFF]/.test(s) ? 'lo' : /[\u1780-\u17FF\u19E0-\u19FF]/.test(s) ? 'km' : /[\u1000-\u109F]/.test(s) ? 'my' : 'th';
/** 一個字母連同疊在上面、下面的符號算一個（泰文聲調、越南文分開寫的聲調） */
export const graphemes = (s: string): string[] => s.match(/\P{M}\p{M}*|\p{M}+/gu) || [];

/** 泰文等一段連續的字 → 一個一個詞 */
export function seaWords(run: string): string[] {
  const loc = seaLocale(run);
  let seg = segmenters.get(loc);
  if (seg === undefined) {
    try {
      const Seg = (Intl as any).Segmenter;
      seg = Seg ? new Seg(loc, { granularity: 'word' }) : null;
    } catch {
      seg = null;
    }
    segmenters.set(loc, seg);
  }
  if (seg) return Array.from(seg.segment(run) as Iterable<{ segment: string }>, x => x.segment);
  return graphemes(run);
}

const OPENERS = /^[(（\[{“‘"'「『《〈【]+$/;
const CLOSERS = /^[,.;:!?)）\]}”’"'、，。；：！？」』》〉】…]+$/;
/**
 * 一段文字（沒有換行）切成斷行單位：英文單字（連著貨幣符號、數字、標點）、中文一個字、
 * 泰文一個詞、空白。左括號黏後面、右括號和標點黏前面——行首不會出現「，」，行尾不會留「(」。
 */
export function lineTokens(para: string): string[] {
  const raw = para.match(
    /(?:[A-Za-z0-9\u00c0-\u024f\u1e00-\u1eff\u0300-\u036f@._%+'’$€£¥₫฿#&*~]|,(?=\d))+[-/]?|[\u0E00-\u0EFF\u1000-\u109F\u1780-\u17FF\u19E0-\u19FF]+|[-/]|\s+|\P{M}\p{M}*|\p{M}+/gu,
  ) || [para];
  const out: string[] = [];
  let pendingOpen = '';
  for (const tk of raw) {
    const pieces = SEA_RE.test(tk) ? seaWords(tk) : [tk];
    for (const piece of pieces) {
      if (OPENERS.test(piece)) { pendingOpen += piece; continue; }
      if (CLOSERS.test(piece) && out.length && !/^\s+$/.test(out[out.length - 1]) && !pendingOpen) {
        out[out.length - 1] += piece;
        continue;
      }
      if (/^\s+$/.test(piece) && pendingOpen) { out.push(pendingOpen); pendingOpen = ''; }
      out.push(pendingOpen + piece);
      pendingOpen = '';
    }
  }
  if (pendingOpen) out.push(pendingOpen);
  return out;
}

/**
 * 依文字的語言決定行高、第一行基準線的位置（都是字級的倍數）。
 * 泰文的母音、聲調會疊在字的上面和下面，越南文一個字母上面可能疊兩個符號，
 * 行距不拉開的話上下行的符號會撞在一起。
 */
export function scriptMetrics(text: string, baseLineHeight = 1.15, lang?: string): { lineHeight: number; ascent: number } {
  // 知道目標語言就照語言決定（同一頁每一格的行距才一致），不知道再看文字本身
  if (lang && /泰|寮|高棉|柬埔寨|緬甸|thai|lao|khmer|myanmar|burmese|ไทย/i.test(lang)) {
    return { lineHeight: baseLineHeight * 1.3, ascent: 1.05 };
  }
  if (lang && /越南|vietnam|tiếng việt/i.test(lang)) return { lineHeight: baseLineHeight * 1.12, ascent: 0.95 };
  if (SEA_RE.test(text)) return { lineHeight: baseLineHeight * 1.3, ascent: 1.05 };
  if (/[\u0300-\u036f\u1ea0-\u1ef9\u01a0\u01a1\u01af\u01b0\u0102\u0103\u0110\u0111]/.test(text)) {
    return { lineHeight: baseLineHeight * 1.12, ascent: 0.95 };
  }
  return { lineHeight: baseLineHeight, ascent: 0.88 };
}

/**
 * 依寬度換行：英數字整個單字不拆（但可以在 - / 後面斷），中文、泰文逐字拆。
 * 單一個字比整行還寬的，才逐字母硬拆。
 */
export function wrapText(text: string, maxWidth: number, fontSize: number, measure: Measurer): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    if (maxWidth <= 0 || measure(para, fontSize) <= maxWidth) { out.push(para); continue; }
    const tokens = lineTokens(para);
    let cur = '';
    const pushToken = (tk: string) => {
      const next = cur + tk;
      if (cur.trim() && measure(next, fontSize) > maxWidth) {
        out.push(cur.replace(/\s+$/, ''));
        cur = tk.replace(/^\s+/, '');
      } else {
        cur = next;
      }
    };
    for (const tk of tokens) {
      if (tk.trim() && measure(tk, fontSize) > maxWidth) {
        for (const ch of graphemes(tk)) pushToken(ch);   // 太長的單字只好逐字母拆（符號不跟字母分開）
      } else {
        pushToken(tk);
      }
    }
    if (cur.trim()) out.push(cur.replace(/\s+$/, ''));
  }
  return out.length ? out : [text];
}

/**
 * 找出塞得進 width × height 的最大字級（不超過 maxFont）。
 * 短的標籤排成一行比較好讀：一行的字級跟換行的差不多（八成以上）就用一行。
 * 上限調低到一個本來就放得下的字級時，結果一定還是那個字級（同一格統一字級要靠這個）。
 * 真的塞不下就用最小字級，回傳 fits = false。
 */
export function fitText(
  text: string,
  width: number,
  height: number,
  maxFont: number,
  minFont: number,
  measure: Measurer,
  lineHeight = 1.15,
  /** 短標籤排成一行時，字級最少要是可換行字級的幾成（跟旁邊統一字級時要接近一點，不然會比鄰居小一截） */
  oneLineRatio = 0.8,
): { fontSize: number; lines: string[]; fits: boolean } {
  const top = Math.max(minFont, maxFont);
  const step = (fs: number) => Math.max(minFont, fs - (fs > 8 ? 0.5 : 0.25));
  // 最長的英文單字、泰文詞：放不下就先縮字，不要把一個字從中間拆開
  const words = text.split('\n').flatMap(lineTokens).filter(w =>
    /[A-Za-z0-9\u00c0-\u024f\u1e00-\u1eff]{2,}/.test(w) || (SEA_RE.test(w) && graphemes(w).length >= 2));
  const fitsAt = (fs: number) => {
    const lines = wrapText(text, width, fs, measure);
    const widest = Math.max(...lines.map(l => measure(l, fs)));
    const wordsFit = fs <= minFont || words.every(w => measure(w, fs) <= width + 0.5);
    return { lines, ok: wordsFit && lines.length * fs * lineHeight <= height + fs * 0.25 && widest <= width + 0.5 };
  };

  // 可以換行時的最大字級
  let general: { fontSize: number; lines: string[]; fits: boolean } | null = null;
  for (let fs = top; ; fs = step(fs)) {
    const { lines, ok } = fitsAt(fs);
    if (ok) { general = { fontSize: fs, lines, fits: true }; break; }
    if (fs <= minFont) { general = { fontSize: minFont, lines, fits: false }; break; }
  }

  // 短標籤：排成一行的最大字級
  const isShort = !text.includes('\n') && text.length <= 28;
  if (isShort && general.lines.length > 1) {
    for (let fs = top; fs >= minFont; fs = step(fs)) {
      if (fs < general.fontSize * oneLineRatio) break;
      if (measure(text, fs) <= width + 0.5 && fs * lineHeight <= height + fs * 0.25) {
        return { fontSize: fs, lines: [text], fits: true };
      }
      if (fs <= minFont) break;
    }
  }
  return general;
}
