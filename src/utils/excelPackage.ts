// ─────────────────────────────────────────────────────────────
// 修補 exceljs 寫出來的 Excel 檔，讓版面跟原檔一樣
//
// 1) 預設字型：exceljs 會把活頁簿的預設字型（「一般」樣式）換成 Calibri 11。
//    Excel 的欄寬是用預設字型的數字寬當單位——新細明體 12 是 8 像素、Calibri 11 是 7 像素，
//    所以每一欄都會莫名其妙窄掉一成多，字被擠得一直換行。這裡把原檔的預設字型放回去。
// 2) 圖形（線條、文字方塊、圖案、圖表）：exceljs 只保留圖片，其他都會丟掉——
//    例如表頭的斜線（「記錄內容／品項」中間那條）翻完就不見了。
//    這裡從原檔的 zip 把每張工作表的繪圖（xl/drawings/…）連同它用到的檔案整組複製回去。
//    繪圖是用「從第幾欄第幾列到第幾欄第幾列」定位的，列被撐高後圖形會跟著拉長，位置不會跑掉。
//
// 另外把表頭斜線的位置讀出來（readDiagonals），撐開表格時用來調整斜線兩邊文字的位置。
// 只用字串處理 XML，瀏覽器和測試環境都能跑。
// ─────────────────────────────────────────────────────────────
import type JSZip from 'jszip';

const PKG_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DOC_RELS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DRAWING_REL_RE = /\/relationships\/drawing$/;

export interface Rel { id: string; type: string; target: string; external: boolean }
export interface CellBox { top: number; left: number; bottom: number; right: number }

// ── XML 小工具 ──────────────────────────────────────────────

const decodeXml = (s: string) =>
  s.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|quot|apos|amp);/gi, (_, e: string) => {
    const k = e.toLowerCase();
    if (k === 'lt') return '<';
    if (k === 'gt') return '>';
    if (k === 'quot') return '"';
    if (k === 'apos') return "'";
    if (k === 'amp') return '&';
    return String.fromCodePoint(k.startsWith('#x') ? parseInt(k.slice(2), 16) : parseInt(k.slice(1), 10));
  });
const encodeXml = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export const xmlAttr = (tag: string, name: string): string | undefined => {
  const m = tag.match(new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`));
  return m ? decodeXml(m[1] ?? m[2] ?? '') : undefined;
};

export const parseRels = (xml: string | null | undefined): Rel[] => {
  if (!xml) return [];
  const out: Rel[] = [];
  for (const m of xml.match(/<(?:\w+:)?Relationship\b[^>]*>/g) || []) {
    out.push({
      id: xmlAttr(m, 'Id') || '',
      type: xmlAttr(m, 'Type') || '',
      target: xmlAttr(m, 'Target') || '',
      external: (xmlAttr(m, 'TargetMode') || '').toLowerCase() === 'external',
    });
  }
  return out;
};

const writeRels = (rels: Rel[]) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG_RELS_NS}">` +
  rels.map(r =>
    `<Relationship Id="${encodeXml(r.id)}" Type="${encodeXml(r.type)}" Target="${encodeXml(r.target)}"` +
    `${r.external ? ' TargetMode="External"' : ''}/>`
  ).join('') +
  '</Relationships>';

// ── 路徑 ───────────────────────────────────────────────────

const dirOf = (p: string) => (p.includes('/') ? p.slice(0, p.lastIndexOf('/')) : '');
const baseOf = (p: string) => p.slice(p.lastIndexOf('/') + 1);
export const relsPathOf = (part: string) => `${dirOf(part) ? dirOf(part) + '/' : ''}_rels/${baseOf(part)}.rels`;

/** 關聯檔裡的 Target（相對於來源檔案）→ zip 裡的完整路徑 */
export const resolveTarget = (fromPart: string, target: string) => {
  const raw = target.startsWith('/') ? target.slice(1) : `${dirOf(fromPart) ? dirOf(fromPart) + '/' : ''}${target}`;
  const out: string[] = [];
  for (const seg of raw.split('/')) {
    if (!seg || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
};

/** zip 裡的完整路徑 → 從 fromPart 看過去的相對路徑 */
export const relativeTarget = (fromPart: string, toPart: string) => {
  const from = dirOf(fromPart).split('/').filter(Boolean);
  const to = toPart.split('/');
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return [...from.slice(i).map(() => '..'), ...to.slice(i)].join('/');
};

const readText = async (zip: JSZip, path: string) => {
  const f = zip.file(path);
  return f ? f.async('string') : null;
};

// ── 工作表名稱 → 工作表檔案 ────────────────────────────────

export const sheetParts = async (zip: JSZip): Promise<Map<string, string>> => {
  const map = new Map<string, string>();
  const rootRels = parseRels(await readText(zip, '_rels/.rels'));
  const wbRel = rootRels.find(r => /\/officeDocument$/.test(r.type));
  const wbPath = wbRel ? resolveTarget('', wbRel.target) : 'xl/workbook.xml';
  const wbXml = await readText(zip, wbPath);
  if (!wbXml) return map;
  const rels = parseRels(await readText(zip, relsPathOf(wbPath)));
  for (const tag of wbXml.match(/<(?:\w+:)?sheet\b[^>]*>/g) || []) {
    const name = xmlAttr(tag, 'name');
    const rid = xmlAttr(tag, 'r:id');
    const rel = rels.find(r => r.id === rid);
    if (name !== undefined && rel) map.set(name, resolveTarget(wbPath, rel.target));
  }
  return map;
};

// ── [Content_Types].xml ────────────────────────────────────

const CT_PATH = '[Content_Types].xml';
interface ContentTypes { defaults: Map<string, string>; overrides: Map<string, string> }

const parseContentTypes = (xml: string | null): ContentTypes => {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const tag of xml?.match(/<(?:\w+:)?Default\b[^>]*>/g) || []) {
    const ext = xmlAttr(tag, 'Extension');
    const ct = xmlAttr(tag, 'ContentType');
    if (ext && ct) defaults.set(ext.toLowerCase(), ct);
  }
  for (const tag of xml?.match(/<(?:\w+:)?Override\b[^>]*>/g) || []) {
    const part = xmlAttr(tag, 'PartName');
    const ct = xmlAttr(tag, 'ContentType');
    if (part && ct) overrides.set(part, ct);
  }
  return { defaults, overrides };
};

const contentTypeOf = (ct: ContentTypes, part: string) =>
  ct.overrides.get('/' + part) ?? ct.defaults.get((part.split('.').pop() || '').toLowerCase());

// ── 圖形種類 ───────────────────────────────────────────────

/** 繪圖裡有沒有 exceljs 不會保留的東西（線條、形狀、群組、圖表）——只有圖片的話 exceljs 自己會寫 */
export const hasNonPictureShapes = (drawingXml: string) =>
  /<(?:\w+:)?(?:sp|cxnSp|grpSp|graphicFrame)\b/.test(drawingXml);

const EMU_PER_PT = 12700;
/** 錨點落在下一格開頭這麼近以內，就當作停在上一格的邊上 */
const EDGE_EMU = 4 * EMU_PER_PT;

/**
 * 從繪圖讀出跨好幾列的斜線（表頭斜線），回傳它框住的儲存格範圍（從 1 起算）。
 * 例：從 A3 左上角畫到 B4 右下角 → { top: 3, left: 1, bottom: 4, right: 2 }
 */
export const diagonalBoxes = (drawingXml: string): CellBox[] => {
  const boxes: CellBox[] = [];
  const anchorRe = /<((?:\w+:)?)twoCellAnchor\b[^>]*>([\s\S]*?)<\/\1twoCellAnchor>/g;
  let m: RegExpExecArray | null;
  while ((m = anchorRe.exec(drawingXml))) {
    const body = m[2];
    if (/<(?:\w+:)?(?:grpSp|pic|graphicFrame)\b/.test(body)) continue;
    if (!/prst="(?:line|straightConnector1)"/.test(body)) continue;
    const pos = (tag: 'from' | 'to') => {
      const block = body.match(new RegExp(`<(?:\\w+:)?${tag}>([\\s\\S]*?)</(?:\\w+:)?${tag}>`))?.[1] || '';
      const num = (n: string) => Number(block.match(new RegExp(`<(?:\\w+:)?${n}>\\s*(-?\\d+)\\s*<`))?.[1] ?? NaN);
      return { col: num('col'), colOff: num('colOff'), row: num('row'), rowOff: num('rowOff') };
    };
    const a = pos('from');
    const b = pos('to');
    if ([a.col, a.row, b.col, b.row].some(n => !Number.isFinite(n))) continue;
    const top = a.row + 1;
    const left = a.col + 1;
    const bottom = Math.max(top, (b.rowOff || 0) <= EDGE_EMU ? b.row : b.row + 1);
    const right = Math.max(left, (b.colOff || 0) <= EDGE_EMU ? b.col : b.col + 1);
    // 只處理真的斜著跨過好幾列的線（水平線、直線、同一格裡的斜線都不算）
    const spansRows = b.row > a.row || (b.row === a.row && b.rowOff > a.rowOff);
    const spansCols = b.col > a.col || (b.col === a.col && b.colOff > a.colOff);
    if (bottom > top && spansRows && spansCols) boxes.push({ top, left, bottom, right });
  }
  return boxes;
};

/** 每張工作表的表頭斜線範圍（工作表名稱 → 範圍） */
export const readDiagonals = async (zip: JSZip): Promise<Map<string, CellBox[]>> => {
  const out = new Map<string, CellBox[]>();
  for (const [name, sheet] of await sheetParts(zip)) {
    for (const rel of parseRels(await readText(zip, relsPathOf(sheet)))) {
      if (!DRAWING_REL_RE.test(rel.type) || rel.external) continue;
      const xml = await readText(zip, resolveTarget(sheet, rel.target));
      if (!xml) continue;
      const boxes = diagonalBoxes(xml);
      if (boxes.length) out.set(name, [...(out.get(name) || []), ...boxes]);
    }
  }
  return out;
};

// ── 複製繪圖 ───────────────────────────────────────────────

/** worksheet 裡 <drawing> 要放的位置：在這些元素之前（照 Excel 規定的順序） */
const AFTER_DRAWING = [
  'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects',
  'controls', 'webPublishItems', 'tableParts', 'extLst',
];

/** worksheet 最外層那個 <extLst> 的位置（條件式格式裡面也會有 extLst，那些不算）；沒有就 -1 */
const topLevelExtLst = (xml: string, prefix: string) => {
  const end = xml.lastIndexOf(`</${prefix}worksheet>`);
  const re = new RegExp(`<(/?)${prefix}extLst(?=[\\s/>])[^>]*?(/?)>`, 'g');
  const stack: number[] = [];
  let m: RegExpExecArray | null;
  let lastTop = -1;
  let lastTopEnd = -1;
  while ((m = re.exec(xml)) && m.index < end) {
    const tagEnd = m.index + m[0].length;
    if (m[2] === '/') {   // <extLst/>
      if (!stack.length) { lastTop = m.index; lastTopEnd = tagEnd; }
      continue;
    }
    if (!m[1]) stack.push(m.index);
    else {
      const open = stack.pop();
      if (open !== undefined && !stack.length) { lastTop = open; lastTopEnd = tagEnd; }
    }
  }
  if (lastTop < 0) return -1;
  // 最外層的 extLst 一定是 worksheet 的最後一個子元素
  return /^\s*$/.test(xml.slice(lastTopEnd, end)) ? lastTop : -1;
};

export const insertDrawingElement = (sheetXml: string, rid: string) => {
  let xml = sheetXml;
  const root = xml.match(/<((?:\w+:)?)worksheet\b[^>]*>/);
  if (!root) return xml;
  const prefix = root[1];
  // 確保有 r: 命名空間
  let rPrefix = 'r';
  const nsDecl = root[0].match(new RegExp(`xmlns:(\\w+)="${DOC_RELS_NS.replace(/[.]/g, '\\.')}"`));
  if (nsDecl) rPrefix = nsDecl[1];
  else if (!/\/>$/.test(root[0])) xml = xml.replace(root[0], root[0].replace(/\s*>$/, ` xmlns:r="${DOC_RELS_NS}">`));

  const el = `<${prefix}drawing ${rPrefix}:id="${encodeXml(rid)}"/>`;
  let at = -1;
  for (const name of AFTER_DRAWING) {
    const i = name === 'extLst' ? topLevelExtLst(xml, prefix) : xml.search(new RegExp(`<${prefix}${name}[\\s/>]`));
    if (i >= 0 && (at < 0 || i < at)) at = i;
  }
  if (at < 0) at = xml.lastIndexOf(`</${prefix}worksheet>`);
  if (at < 0) return xml;
  return xml.slice(0, at) + el + xml.slice(at);
};

const uniquePath = (zip: JSZip, wanted: string, reserved: Set<string>) => {
  const taken = (p: string) => !!zip.file(p) || reserved.has(p);
  if (!taken(wanted)) return wanted;
  const dir = dirOf(wanted);
  const base = baseOf(wanted);
  const dot = base.lastIndexOf('.');
  const stem = dot > 0 ? base.slice(0, dot) : base;
  const ext = dot > 0 ? base.slice(dot) : '';
  const m = stem.match(/^(.*?)(\d+)$/);
  const head = m ? m[1] : stem + '_';
  let n = m ? Number(m[2]) + 1 : 1;
  for (;;) {
    const p = `${dir ? dir + '/' : ''}${head}${n}${ext}`;
    if (!taken(p)) return p;
    n++;
  }
};

interface CopyCtx {
  src: JSZip;
  dst: JSZip;
  srcCt: ContentTypes;
  dstCt: ContentTypes;
  /** 來源路徑 → 目的路徑（同一個檔案只複製一次） */
  copied: Map<string, string>;
  reserved: Set<string>;
}

/** 把來源的一個檔案連同它關聯到的檔案（圖片、圖表…）整組複製過去 */
async function copyPartTree(ctx: CopyCtx, srcPath: string, dstPath: string): Promise<boolean> {
  const file = ctx.src.file(srcPath);
  if (!file) return false;
  ctx.copied.set(srcPath, dstPath);
  ctx.reserved.add(dstPath);
  ctx.dst.file(dstPath, await file.async('uint8array'));

  const ct = contentTypeOf(ctx.srcCt, srcPath);
  if (ct && contentTypeOf(ctx.dstCt, dstPath) !== ct) ctx.dstCt.overrides.set('/' + dstPath, ct);

  const rels = parseRels(await readText(ctx.src, relsPathOf(srcPath)));
  if (!rels.length) return true;
  const outRels: Rel[] = [];
  for (const rel of rels) {
    if (rel.external) { outRels.push(rel); continue; }
    const srcTarget = resolveTarget(srcPath, rel.target);
    let dstTarget = ctx.copied.get(srcTarget);
    if (!dstTarget) {
      dstTarget = uniquePath(ctx.dst, srcTarget, ctx.reserved);
      if (!(await copyPartTree(ctx, srcTarget, dstTarget))) continue;   // 原檔裡找不到的就略過
    }
    outRels.push({ ...rel, target: relativeTarget(dstPath, dstTarget) });
  }
  ctx.dst.file(relsPathOf(dstPath), writeRels(outRels));
  return true;
}

const writeContentTypes = (ct: ContentTypes, originalXml: string) => {
  const head = originalXml.match(/^[\s\S]*?<(?:\w+:)?Types\b[^>]*>/)?.[0]
    || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`;
  const defaults = [...ct.defaults].map(([e, t]) => `<Default Extension="${encodeXml(e)}" ContentType="${encodeXml(t)}"/>`);
  const overrides = [...ct.overrides].map(([p, t]) => `<Override PartName="${encodeXml(p)}" ContentType="${encodeXml(t)}"/>`);
  return head + defaults.join('') + overrides.join('') + '</Types>';
};

export interface RestoreResult { sheets: number }

/**
 * 把原檔每張工作表的繪圖複製到 exceljs 寫出的檔案裡（直接改 out 這個 zip）。
 * 只有圖片的繪圖 exceljs 自己會寫，不動；有線條、形狀、圖表的才換成原檔那份。
 */
export async function restoreDrawingsInZip(orig: JSZip, out: JSZip): Promise<RestoreResult> {
  const result: RestoreResult = { sheets: 0 };
  const origSheets = await sheetParts(orig);
  const outSheets = await sheetParts(out);
  const ctXml = (await readText(out, CT_PATH)) || '';
  const ctx: CopyCtx = {
    src: orig, dst: out,
    srcCt: parseContentTypes(await readText(orig, CT_PATH)),
    dstCt: parseContentTypes(ctXml),
    copied: new Map(), reserved: new Set(),
  };

  for (const [name, origSheet] of origSheets) {
    const outSheet = outSheets.get(name);
    if (!outSheet) continue;
    const drawingRel = parseRels(await readText(orig, relsPathOf(origSheet)))
      .find(r => DRAWING_REL_RE.test(r.type) && !r.external);
    if (!drawingRel) continue;
    const origDrawing = resolveTarget(origSheet, drawingRel.target);
    const drawingXml = await readText(orig, origDrawing);
    if (!drawingXml || !hasNonPictureShapes(drawingXml)) continue;

    let sheetXml = await readText(out, outSheet);
    if (!sheetXml) continue;
    const outRelsPath = relsPathOf(outSheet);
    const outRels = parseRels(await readText(out, outRelsPath));

    const newDrawing = uniquePath(out, 'xl/drawings/drawing1.xml', ctx.reserved);
    if (!(await copyPartTree(ctx, origDrawing, newDrawing))) continue;

    const existing = outRels.find(r => DRAWING_REL_RE.test(r.type) && !r.external);
    if (existing) {
      // exceljs 已經寫了一份（只有圖片）→ 改指向原檔那份（原檔那份也有這些圖片），舊的拿掉
      const old = resolveTarget(outSheet, existing.target);
      existing.target = relativeTarget(outSheet, newDrawing);
      if (old !== newDrawing) {
        out.remove(old);
        out.remove(relsPathOf(old));
        ctx.dstCt.overrides.delete('/' + old);
      }
    } else {
      let n = outRels.length + 1;
      while (outRels.some(r => r.id === `rId${n}`)) n++;
      const rid = `rId${n}`;
      outRels.push({ id: rid, type: drawingRel.type, target: relativeTarget(outSheet, newDrawing), external: false });
      // exceljs 不會有 <drawing>，保險起見先拿掉
      sheetXml = sheetXml.replace(/<(?:\w+:)?drawing\b[^>]*\/>/g, '');
      sheetXml = insertDrawingElement(sheetXml, rid);
      out.file(outSheet, sheetXml);
    }
    out.file(outRelsPath, writeRels(outRels));
    result.sheets++;
  }

  if (result.sheets) out.file(CT_PATH, writeContentTypes(ctx.dstCt, ctXml));
  return result;
}

// ── 預設字型 ───────────────────────────────────────────────

const stylesPath = async (zip: JSZip) => {
  const rootRels = parseRels(await readText(zip, '_rels/.rels'));
  const wbRel = rootRels.find(r => /\/officeDocument$/.test(r.type));
  const wbPath = wbRel ? resolveTarget('', wbRel.target) : 'xl/workbook.xml';
  const rel = parseRels(await readText(zip, relsPathOf(wbPath))).find(r => /\/styles$/.test(r.type));
  return rel ? resolveTarget(wbPath, rel.target) : 'xl/styles.xml';
};

const FONT_RE = /<(?:\w+:)?font\b[^>]*\/>|<((?:\w+:)?)font\b[^>]*>[\s\S]*?<\/\1font>/g;
const fontsBlock = (xml: string) => xml.match(/<((?:\w+:)?)fonts\b[^>]*>([\s\S]*?)<\/\1fonts>/);
/** 同一個字型的兩種寫法（子元素順序不同）視為相同 */
const fontKey = (font: string) =>
  (font.match(/<(?:\w+:)?(\w+)\b([^>]*)\/>/g) || [])
    .map(t => t.replace(/^<(?:\w+:)?/, '<').replace(/\s+/g, ' ').trim())
    .sort()
    .join('');
const stripPrefixes = (xml: string) => xml.replace(/<(\/?)\w+:/g, '<$1');

/** 活頁簿的預設字型（「一般」樣式用的字型）：名稱、字級 */
export const defaultFontOf = async (zip: JSZip): Promise<{ name?: string; size?: number } | null> => {
  const xml = await readText(zip, await stylesPath(zip));
  const block = xml && fontsBlock(xml);
  const first = block?.[2].match(FONT_RE)?.[0];
  if (!first) return null;
  const val = (tag: string) => {
    const t = first.match(new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*>`))?.[0];
    return t ? xmlAttr(t, 'val') : undefined;
  };
  const size = Number(val('sz'));
  return { name: val('name'), size: Number.isFinite(size) && size > 0 ? size : undefined };
};

/**
 * 把原檔的預設字型放回 exceljs 寫出的檔（直接改 out）。回傳有沒有改。
 * exceljs 的第 0 個字型固定是 Calibri 11：換成原檔的預設字型；
 * 本來就是 Calibri 11 而指到第 0 個字型的格子，改指到放在最後面的 Calibri 11，字型不變。
 */
export async function restoreDefaultFontInZip(orig: JSZip, out: JSZip): Promise<boolean> {
  const origXml = await readText(orig, await stylesPath(orig));
  const outPath = await stylesPath(out);
  const outXml = await readText(out, outPath);
  if (!origXml || !outXml) return false;
  const origFonts = fontsBlock(origXml)?.[2].match(FONT_RE) || [];
  const outBlock = fontsBlock(outXml);
  const outFonts = outBlock?.[2].match(FONT_RE) || [];
  if (!origFonts.length || !outBlock || !outFonts.length) return false;
  if (fontKey(origFonts[0]) === fontKey(outFonts[0])) return false;

  const moved = outFonts.length;   // 原本的第 0 個字型搬到最後
  const fonts = [stripPrefixes(origFonts[0]), ...outFonts.slice(1), outFonts[0]];
  const openTag = outBlock[0].match(/^<[^>]*>/)![0].replace(/\scount="\d+"/, ` count="${fonts.length}"`);
  let xml = outXml.replace(outBlock[0], `${openTag}${fonts.join('')}</${outBlock[1]}fonts>`);

  // cellXfs：第 0 個（沒設定樣式的格子）照樣用預設字型；其他指到第 0 個字型的改指到搬走的 Calibri
  xml = xml.replace(/<((?:\w+:)?)cellXfs\b[^>]*>[\s\S]*?<\/\1cellXfs>/, block => {
    let i = -1;
    return block.replace(/<(?:\w+:)?xf\b[^>]*>/g, tag => {
      i++;
      if (i === 0) return tag;
      return /\sfontId="0"/.test(tag) ? tag.replace(/\sfontId="0"/, ` fontId="${moved}"`) : tag;
    });
  });
  out.file(outPath, xml);
  return true;
}

export interface RepairResult { data: Uint8Array | ArrayBuffer; drawings: number; defaultFont: boolean }

/**
 * 原檔、exceljs 寫出的檔 → 修補後的檔：放回預設字型、補回線條／形狀／圖表。
 * 每一步各自失敗不影響其他步驟；什麼都沒改就原封不動回傳。
 */
export async function repairExcelPackage(
  JSZipCtor: typeof JSZip,
  original: ArrayBuffer | Uint8Array,
  written: ArrayBuffer | Uint8Array,
): Promise<RepairResult> {
  const orig = await JSZipCtor.loadAsync(original);
  const out = await JSZipCtor.loadAsync(written);
  let defaultFont = false;
  let drawings = 0;
  try {
    defaultFont = await restoreDefaultFontInZip(orig, out);
  } catch (e) {
    console.warn('還原 Excel 預設字型失敗', e);
  }
  try {
    drawings = (await restoreDrawingsInZip(orig, out)).sheets;
  } catch (e) {
    console.warn('補回 Excel 圖形失敗', e);
  }
  if (!defaultFont && !drawings) return { data: written, drawings, defaultFont };
  const data = await out.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  return { data, drawings, defaultFont };
}
