// ─────────────────────────────────────────────────────────────
// Excel 翻譯後把表格撐開
//
// 譯文接在原文下一行，字變多了，但列高還是原本的，Excel 只會把多出來的字切掉。
// Excel 自己的「自動調整列高」對合併儲存格沒有作用（表單幾乎都是合併儲存格），
// 所以這裡自己算：依每一格的字、字級、欄寬算出要幾行，需要多高就把列撐高；
// 英文長單字塞不進窄欄時，把欄稍微加寬（有上限，不會把整張表撐得太寬）。
// 最後調整列印設定：表單每張印一頁（Excel 自動縮小），長的資料表一頁寬、高度照常分頁。
// 不碰 DOM：量字寬的函式由呼叫的地方傳進來（瀏覽器用 canvas 量）。
// ─────────────────────────────────────────────────────────────

/** 量一段文字的寬度（pt） */
export type MeasureText = (text: string, sizePt: number, fontName?: string, bold?: boolean) => number;

// ── 符號字型 ───────────────────────────────────────────────
// 原文裡用符號字型（Wingdings 等）的字，例如 Wingdings 2 的「£」其實是勾選框 □。
// 譯文只有這些字照樣用符號字型，其他文字一律用一般字型——
// 不然模型把整句英文放進符號字型的區段，就會整句變成一串符號。
export const SYMBOL_FONT_RE = /wingdings|webdings|^symbol$|marlett|dingbats|mt extra/i;
export const isSymbolFont = (font: any) => !!font?.name && SYMBOL_FONT_RE.test(String(font.name));
export const symbolCharsOf = (runs: any[] | undefined): Map<string, any> => {
  const map = new Map<string, any>();
  for (const rt of runs || []) {
    if (!isSymbolFont(rt?.font)) continue;
    for (const ch of Array.from(String(rt.text || ''))) if (ch.trim()) map.set(ch, rt.font);
  }
  return map;
};
export const splitSymbolRuns = (text: string, font: any, plainFont: any, symbols: Map<string, any>) => {
  const out: { text: string; font: any }[] = [];
  for (const ch of Array.from(text)) {
    const f = symbols.get(ch) ?? (isSymbolFont(font) ? plainFont : font);
    const last = out[out.length - 1];
    if (last && last.font === f) last.text += ch;
    else out.push({ text: ch, font: f });
  }
  return out;
};

// ── 撐開表格 ───────────────────────────────────────────────

export const EXCEL_FIT = {
  /** 行高 = 字級 × 這個倍數（新細明體／標楷體 12pt 在 Excel 是 16.5pt） */
  lineHeight: 1.36,
  /** 儲存格上下留白（pt） */
  paddingV: 3,
  /** 儲存格左右留白（pt） */
  paddingH: 5,
  /** 量出來的字寬再多抓一點，避免實際字型比較寬時又被切掉 */
  safety: 1.08,
  /** 欄寬最多加寬到原本的幾倍 */
  maxWidenRatio: 1.6,
  /** 欄寬上限（字元數） */
  maxColWidth: 45,
  /** Excel 列高上限 */
  maxRowHeight: 409,
  /** 預設字型的數字寬（px），Excel 換算欄寬用 */
  digitPx: 7,
  /** 幾列以內算「表單」：列印存成每張一頁 */
  formMaxRows: 60,
};

export interface Range { top: number; left: number; bottom: number; right: number }

/** Excel 欄寬（字元數）→ pt */
export const colWidthToPt = (w: number) => {
  const px = Math.trunc(((256 * w + Math.trunc(128 / EXCEL_FIT.digitPx)) / 256) * EXCEL_FIT.digitPx);
  return px * 0.75;
};

/** pt → Excel 欄寬（字元數） */
export const ptToColWidth = (pt: number) => Math.ceil((pt / 0.75 / EXCEL_FIT.digitPx) * 100) / 100 + 0.15;

interface Paragraph { text: string; size: number; font?: string; bold?: boolean }

/** 儲存格內容 → 一段一段（依換行切開），每段記下字級和字型 */
export function cellParagraphs(value: any, cellFont: any): Paragraph[] {
  const baseSize = Number(cellFont?.size) || 11;
  const runs: { text: string; font: any }[] = [];
  if (value && typeof value === 'object' && Array.isArray(value.richText)) {
    for (const rt of value.richText) runs.push({ text: String(rt?.text ?? ''), font: rt?.font || cellFont });
  } else if (value && typeof value === 'object' && 'result' in value) {
    runs.push({ text: String(value.result ?? ''), font: cellFont });
  } else if (value !== null && value !== undefined) {
    runs.push({ text: String(value), font: cellFont });
  }

  const out: Paragraph[] = [];
  let cur: Paragraph | null = null;
  for (const run of runs) {
    const size = Number(run.font?.size) || baseSize;
    const parts = run.text.split(/\r?\n/);
    parts.forEach((part, i) => {
      if (i > 0 || !cur) {
        cur = { text: '', size, font: run.font?.name, bold: !!run.font?.bold };
        out.push(cur);
      }
      cur!.text += part;
      cur!.size = Math.max(cur!.size, size);
    });
  }
  return out;
}

const TOKEN_RE = /\s+|[A-Za-z0-9À-ɏḀ-ỿ'’.,:;!?()%°\-/_+&@#*=<>]+|[\s\S]/g;
const isLatinToken = (t: string) => /[A-Za-z0-9À-ɏḀ-ỿ]/.test(t) && !/\s/.test(t);

/** 一段文字在寬度 width（pt）裡要排幾行（照 Excel 的方式：英文在空白處換行、中文逐字換行） */
export function countLines(p: Paragraph, width: number, measure: MeasureText): number {
  if (!p.text) return 1;
  const w = (t: string) => measure(t, p.size, p.font, p.bold) * EXCEL_FIT.safety;
  if (width <= 0) return 1;
  let lines = 1;
  let cur = 0;
  for (const tk of p.text.match(TOKEN_RE) || []) {
    const tw = w(tk);
    if (/^\s+$/.test(tk)) {
      // 行尾的空白不佔位置；行首的空白（標題前面的縮排）照算
      if (cur + tw <= width) cur += tw;
      continue;
    }
    if (cur > 0 && cur + tw > width) {
      lines++;
      cur = 0;
    }
    if (tw > width) {
      // 單字比整格還寬：Excel 會從中間硬拆
      lines += Math.ceil(tw / width) - 1;
      cur = tw % width;
    } else {
      cur += tw;
    }
  }
  return lines;
}

/** 一格內容需要的高度（pt） */
export function requiredHeight(paragraphs: Paragraph[], width: number, measure: MeasureText): number {
  let h = 0;
  for (const p of paragraphs) h += countLines(p, width, measure) * p.size * EXCEL_FIT.lineHeight;
  return h + EXCEL_FIT.paddingV;
}

/** 最長的英文單字寬度（pt）：比欄寬還寬就要加寬欄位，不然會被從中間拆開 */
export function longestWord(paragraphs: Paragraph[], measure: MeasureText): number {
  let max = 0;
  for (const p of paragraphs) {
    for (const tk of p.text.match(TOKEN_RE) || []) {
      if (isLatinToken(tk)) max = Math.max(max, measure(tk, p.size, p.font, p.bold) * EXCEL_FIT.safety);
    }
  }
  return max;
}

export interface FitResult {
  rowsChanged: number;
  colsWidened: number;
  /** 原本沒換行、文字往右延伸的格子，譯文改接在同一行 */
  inlined: number;
  printFit?: 'one-page' | 'one-page-wide';
}

export interface FitOptions {
  /** 翻譯前每一格有沒有開「自動換行」（位址 → true/false） */
  originalWrap?: Map<string, boolean>;
}

/** 一格的值裡的換行都改成空白（譯文接在同一行） */
function flattenValue(value: any): any {
  const join = (t: string) => t.replace(/\s*\r?\n\s*/g, '  ');
  if (value && typeof value === 'object' && Array.isArray(value.richText)) {
    const runs = value.richText.map((rt: any) => ({ ...rt, text: join(String(rt?.text ?? '')) }));
    return { richText: runs };
  }
  return typeof value === 'string' ? join(value) : value;
}

/**
 * 把一張工作表撐開。
 * @param ws       exceljs 的 Worksheet
 * @param cells    要考慮的儲存格（翻譯過的格子的位址，例如 "C9"；合併的話是左上角那格）
 * @param measure  量字寬的函式
 */
export function fitWorksheet(ws: any, cells: Iterable<string>, measure: MeasureText, options: FitOptions = {}): FitResult {
  const result: FitResult = { rowsChanged: 0, colsWidened: 0, inlined: 0 };
  const defaultRow = Number(ws.properties?.defaultRowHeight) || 15;
  const defaultCol = Number(ws.properties?.defaultColWidth) || 8.43;

  // 原本的列高、欄寬（算列印用、加寬上限用）
  const origColWidth = new Map<number, number>();
  const colWidth = (c: number) => {
    const col = ws.getColumn(c);
    if (col.hidden) return 0;
    return Number(col.width) > 0 ? Number(col.width) : defaultCol;
  };
  const rowHeight = (r: number) => {
    const row = ws.getRow(r);
    if (row.hidden) return 0;
    return Number(row.height) > 0 ? Number(row.height) : defaultRow;
  };
  const lastRow = ws.rowCount || 0;
  const lastCol = ws.columnCount || 0;

  const merges: Record<string, Range> = {};
  for (const [master, rng] of Object.entries<any>(ws._merges || {})) {
    const m = rng?.model || rng;
    if (m && m.top) merges[master] = { top: m.top, left: m.left, bottom: m.bottom, right: m.right };
  }

  // 被合併進別格的位置（不是左上角那格）
  const coveredByMerge = (r: number, c: number) =>
    Object.values(merges).some(m => r >= m.top && r <= m.bottom && c >= m.left && c <= m.right);
  const isEmpty = (r: number, c: number) => {
    const v = ws.getCell(r, c).value;
    return v === null || v === undefined || (typeof v === 'string' && !v.trim());
  };

  type Box = { range: Range; paras: Paragraph[] };
  const boxes: Box[] = [];
  for (const addr of cells) {
    const cell = ws.getCell(addr);
    const al = cell.alignment || {};
    if (al.shrinkToFit || (al.textRotation && al.textRotation !== 0)) continue;
    let paras = cellParagraphs(cell.value, cell.font);
    if (!paras.some(p => p.text.trim())) continue;
    const range = merges[cell.address] || { top: cell.row, left: cell.col, bottom: cell.row, right: cell.col };

    // 原本沒開換行、靠左、右邊是空格的（例：「類別：醃漬類」），文字本來就往右延伸顯示。
    // 硬塞在這一格裡會被擠成好幾行，所以譯文改接在同一行、照樣往右延伸。
    const wasWrapped = options.originalWrap?.get(cell.address);
    const horizontal = al.horizontal;
    if (
      wasWrapped === false && !merges[cell.address] &&
      (!horizontal || horizontal === 'left' || horizontal === 'general')
    ) {
      const flat = flattenValue(cell.value);
      const flatParas = cellParagraphs(flat, cell.font);
      const textW = Math.max(...flatParas.map(p => measure(p.text, p.size, p.font, p.bold) * EXCEL_FIT.safety));
      let room = colWidthToPt(colWidth(cell.col)) - EXCEL_FIT.paddingH;
      for (let c = cell.col + 1; c <= Math.max(lastCol, cell.col + 1) && room < textW; c++) {
        if (!isEmpty(cell.row, c) || coveredByMerge(cell.row, c)) break;
        room += colWidthToPt(colWidth(c));
      }
      if (textW <= room) {
        cell.value = flat;
        cell.alignment = { ...al, wrapText: false };
        result.inlined++;
        paras = flatParas;
        continue;   // 一行就夠，不用撐高
      }
    }
    boxes.push({ range, paras });
  }

  const widthPt = (r: Range) => {
    let w = 0;
    for (let c = r.left; c <= r.right; c++) w += colWidthToPt(colWidth(c));
    return Math.max(0, w - EXCEL_FIT.paddingH);
  };

  // 1) 英文長單字塞不下 → 加寬範圍最右邊那一欄（有上限）
  for (const b of boxes) {
    const need = longestWord(b.paras, measure);
    const have = widthPt(b.range);
    if (need <= have + 0.5) continue;
    const c = b.range.right;
    if (!origColWidth.has(c)) origColWidth.set(c, colWidth(c));
    const base = origColWidth.get(c)!;
    if (base <= 0) continue;
    const target = ptToColWidth(colWidthToPt(colWidth(c)) + (need - have));
    const capped = Math.min(target, base * EXCEL_FIT.maxWidenRatio, Math.max(base, EXCEL_FIT.maxColWidth));
    if (capped > colWidth(c) + 0.05) {
      ws.getColumn(c).width = Math.round(capped * 100) / 100;
    }
  }
  result.colsWidened = [...origColWidth.entries()].filter(([c, w]) => colWidth(c) > w + 0.05).length;

  // 2) 只占一列的格子：那一列至少要這麼高
  const need = new Map<number, number>();
  for (const b of boxes) {
    if (b.range.top !== b.range.bottom) continue;
    const h = requiredHeight(b.paras, widthPt(b.range), measure);
    need.set(b.range.top, Math.max(need.get(b.range.top) || 0, h));
  }
  const changed = new Set<number>();
  for (const [r, h] of need) {
    if (h > rowHeight(r) + 0.5 && !ws.getRow(r).hidden) {
      ws.getRow(r).height = Math.min(EXCEL_FIT.maxRowHeight, Math.ceil(h * 4) / 4);
      changed.add(r);
    }
  }

  // 3) 跨好幾列的合併格：合起來不夠高，就把差的平均加到那幾列
  const multi = boxes
    .filter(b => b.range.bottom > b.range.top)
    .sort((a, b) => (a.range.bottom - a.range.top) - (b.range.bottom - b.range.top));
  for (const b of multi) {
    const h = requiredHeight(b.paras, widthPt(b.range), measure);
    const rows: number[] = [];
    for (let r = b.range.top; r <= b.range.bottom; r++) if (!ws.getRow(r).hidden) rows.push(r);
    if (!rows.length) continue;
    const total = rows.reduce((a, r) => a + rowHeight(r), 0);
    if (h <= total + 0.5) continue;
    const add = (h - total) / rows.length;
    for (const r of rows) {
      ws.getRow(r).height = Math.min(EXCEL_FIT.maxRowHeight, Math.ceil((rowHeight(r) + add) * 4) / 4);
      changed.add(r);
    }
  }
  result.rowsChanged = changed.size;

  // 4) 列印：撐開後表格變大，原本的列印設定可能會跑成好幾頁。
  //    表單（60 列以內）存成「每張一頁」，Excel 會自動縮小；更長的資料表只限制一頁寬、高度照常分頁。
  //    原本就有設定「調整成幾頁」的不動。
  if ((changed.size || result.colsWidened) && ws.pageSetup && !ws.pageSetup.fitToPage) {
    const isForm = lastRow <= EXCEL_FIT.formMaxRows;
    ws.pageSetup.fitToPage = true;
    ws.pageSetup.fitToWidth = 1;
    ws.pageSetup.fitToHeight = isForm ? 1 : 0;
    result.printFit = isForm ? 'one-page' : 'one-page-wide';
  }

  return result;
}
