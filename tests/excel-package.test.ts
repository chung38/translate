import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  diagonalBoxes, insertDrawingElement, relativeTarget, resolveTarget, restoreDefaultFontInZip,
  defaultFontOf, sheetParts, relsPathOf, parseRels,
} from '../src/utils/excelPackage';
import { processExcel, translateItemsByLanguage } from '../src/utils/documentProcessors';

const MING_FONT = '<font><sz val="12"/><color theme="1"/><name val="新細明體"/><family val="2"/><charset val="136"/><scheme val="minor"/></font>';

/** 表頭斜線：從 A1 左上角畫到 A2 右下角 */
const LINE_DRAWING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:to><xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="2" name="直線接點 1"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></xdr:spPr></xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;

/** 做一份跟使用者表單很像的檔：預設字型新細明體 12、表頭有斜線、廠商／品名重複出現、好幾格共用同一個樣式 */
async function buildForm(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('雞肉');
  ws.getColumn(1).width = 26;
  ws.getColumn(2).width = 8;
  ws.getColumn(3).width = 12;
  const font = { name: '標楷體', size: 14 };
  const set = (addr: string, value: any, alignment: any) => {
    const c = ws.getCell(addr);
    c.value = value;
    c.font = font;
    c.alignment = alignment;
  };
  // 表頭（第 1、2 列，斜線框住 A1:A2）
  ws.getRow(1).height = 40;
  ws.getRow(2).height = 20;
  set('A1', '記錄內容', { horizontal: 'right', vertical: 'middle' });
  set('A2', '品項', { horizontal: 'left', vertical: 'middle' });
  set('B1', '淨重', { horizontal: 'center', vertical: 'middle', wrapText: true });
  // 跟 A1 同一個樣式（exceljs 讀進來會共用同一個樣式物件）
  set('C1', '備註', { horizontal: 'right', vertical: 'middle' });
  // 標籤：廠商（沒開換行）、品名（有開換行），重複兩組
  for (const [r, wrap] of [[3, false], [4, true], [5, false], [6, true]] as const) {
    ws.getRow(r).height = 20;
    set(`A${r}`, r % 2 ? '廠商：' : '品名：           ', { horizontal: 'left', vertical: 'middle', ...(wrap ? { wrapText: true } : {}) });
    ws.getCell(`B${r}`).value = 0;   // 右邊不是空格
  }

  const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
  // 預設字型改成新細明體 12（台灣的 Excel 預設就是這樣）
  const styles = await zip.file('xl/styles.xml')!.async('string');
  zip.file('xl/styles.xml', styles.replace(/(<fonts\b[^>]*>)\s*<font>[\s\S]*?<\/font>/, `$1${MING_FONT}`));
  // 加上表頭斜線
  zip.file('xl/drawings/drawing1.xml', LINE_DRAWING);
  const sheetPath = (await sheetParts(zip)).get('雞肉')!;
  const relsPath = relsPathOf(sheetPath);
  const rels = parseRels(await zip.file(relsPath)?.async('string'));
  zip.file(relsPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${
    rels.map(r => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"/>`).join('')
  }<Relationship Id="rIdDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`);
  const sheet = await zip.file(sheetPath)!.async('string');
  zip.file(sheetPath, insertDrawingElement(sheet, 'rIdDraw'));
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', ct.replace('</Types>',
    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'));
  return zip.generateAsync({ type: 'arraybuffer' });
}

const TRANSLATIONS: Record<string, string> = {
  記錄內容: 'Record Details',
  品項: 'Item',
  淨重: 'Net weight per box measured by the staff on duty',
  備註: 'Remarks',
  '廠商：': 'Supplier:',
  '品名：': 'Product name:',
};

describe('processExcel：翻完版面跟原檔一樣', () => {
  it('預設字型、表頭斜線、標籤排法一致、同一個詞只翻一次', async () => {
    const original = await buildForm();
    const sent: string[] = [];
    const translateBatch = async (texts: string[], langs: string[]) => texts.map(t => {
      sent.push(t);
      const r: Record<string, string> = {};
      for (const l of langs) r[l] = TRANSLATIONS[t.trim()] ?? t;
      return r;
    });
    const file = new File([original], '表單.xlsx');
    const [out] = await processExcel(file, ['英文'], '', translateBatch, () => {}, { current: false } as any, 'combined');
    const buf = await out.blob.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);

    // 1) 預設字型放回新細明體 12（欄寬單位才不會變）
    expect(await defaultFontOf(zip)).toEqual({ name: '新細明體', size: 12 });

    // 2) 斜線還在，而且接在工作表上
    const sheetPath = (await sheetParts(zip)).get('雞肉')!;
    const sheetXml = await zip.file(sheetPath)!.async('string');
    const rels = parseRels(await zip.file(relsPathOf(sheetPath))!.async('string'));
    const drawRel = rels.find(r => /\/drawing$/.test(r.type))!;
    expect(drawRel).toBeTruthy();
    expect(sheetXml).toContain(`<drawing r:id="${drawRel.id}"/>`);
    const drawing = await zip.file(resolveTarget(sheetPath, drawRel.target))!.async('string');
    expect(diagonalBoxes(drawing)).toEqual([{ top: 1, left: 1, bottom: 2, right: 1 }]);
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(resolveTarget(sheetPath, drawRel.target));

    // 3) 重複的「廠商：」「品名：」只送一次翻譯
    expect(sent.filter(t => t.includes('廠商')).length).toBe(1);
    expect(sent.filter(t => t.includes('品名')).length).toBe(1);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const ws = wb.getWorksheet('雞肉')!;
    const text = (a: string) => {
      const v: any = ws.getCell(a).value;
      return v?.richText ? v.richText.map((r: any) => r.text).join('') : String(v);
    };
    // 4) 同一欄的標籤排法一致：「廠商：Supplier:」放得下、「品名：Product name:」放不下 → 全部上下兩行
    for (const a of ['A3', 'A4', 'A5', 'A6']) {
      expect(text(a)).toContain('\n');
      expect(ws.getCell(a).alignment?.wrapText).toBe(true);
    }
    // 5) 表頭被撐高：斜線上面的字靠上、下面的字靠下；跟 A1 同樣式的 C1 不受影響
    expect(Number(ws.getRow(1).height)).toBeGreaterThan(40);
    expect(ws.getCell('A1').alignment?.vertical).toBe('top');
    expect(ws.getCell('A2').alignment?.vertical).toBe('bottom');
    expect(ws.getCell('C1').alignment?.vertical).toBe('middle');
  });
});

describe('restoreDefaultFontInZip', () => {
  const stylesXml = (fonts: string[], xfs: string[]) =>
    `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<fonts count="${fonts.length}">${fonts.join('')}</fonts>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0"/></cellStyleXfs>` +
    `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs></styleSheet>`;
  const calibri = '<font><color theme="1"/><family val="2"/><scheme val="minor"/><sz val="11"/><name val="Calibri"/></font>';
  const kai = '<font><sz val="14"/><name val="標楷體"/></font>';

  it('第 0 個字型換回原檔的；原本真的用 Calibri 的格子改指到搬走的 Calibri', async () => {
    const orig = new JSZip();
    orig.file('xl/styles.xml', stylesXml([MING_FONT, kai], ['<xf fontId="0"/>']));
    const out = new JSZip();
    out.file('xl/styles.xml', stylesXml([calibri, kai], ['<xf numFmtId="0" fontId="0"/>', '<xf numFmtId="0" fontId="1"/>', '<xf numFmtId="0" fontId="0"/>']));
    expect(await restoreDefaultFontInZip(orig, out)).toBe(true);
    const xml = await out.file('xl/styles.xml')!.async('string');
    const fonts = xml.match(/<font>[\s\S]*?<\/font>/g)!;
    expect(fonts[0]).toContain('新細明體');
    expect(fonts[2]).toContain('Calibri');
    expect(xml).toContain('<fonts count="3">');
    const xfs = xml.match(/<cellXfs[\s\S]*<\/cellXfs>/)![0].match(/<xf\b[^>]*>/g)!;
    expect(xfs[0]).toContain('fontId="0"');   // 沒設定樣式的格子：用預設字型
    expect(xfs[1]).toContain('fontId="1"');
    expect(xfs[2]).toContain('fontId="2"');   // 本來就是 Calibri
  });

  it('預設字型本來就一樣（只是寫法順序不同）就不動', async () => {
    const orig = new JSZip();
    orig.file('xl/styles.xml', stylesXml(['<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'], []));
    const out = new JSZip();
    out.file('xl/styles.xml', stylesXml([calibri], []));
    expect(await restoreDefaultFontInZip(orig, out)).toBe(false);
  });
});

describe('XML 小工具', () => {
  it('相對路徑互轉', () => {
    expect(resolveTarget('xl/worksheets/sheet9.xml', '../drawings/drawing1.xml')).toBe('xl/drawings/drawing1.xml');
    expect(resolveTarget('xl/workbook.xml', '/xl/styles.xml')).toBe('xl/styles.xml');
    expect(relativeTarget('xl/worksheets/sheet9.xml', 'xl/drawings/drawing2.xml')).toBe('../drawings/drawing2.xml');
    expect(relativeTarget('xl/drawings/drawing1.xml', 'xl/media/image1.png')).toBe('../media/image1.png');
  });

  it('<drawing> 放在最外層 extLst 前面，不會插進條件式格式裡的 extLst', () => {
    const xml = '<worksheet xmlns="m" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/>' +
      '<conditionalFormatting><cfRule><extLst><ext/></extLst></cfRule></conditionalFormatting>' +
      '<pageSetup/><extLst><ext><extLst/></ext></extLst></worksheet>';
    const out = insertDrawingElement(xml, 'rId9');
    expect(out).toContain('<pageSetup/><drawing r:id="rId9"/><extLst><ext>');
    expect(out.indexOf('<drawing')).toBeGreaterThan(out.indexOf('</conditionalFormatting>'));
  });

  it('水平線、直線不算表頭斜線', () => {
    const flat = LINE_DRAWING.replace('<xdr:row>2</xdr:row><xdr:rowOff>19050', '<xdr:row>0</xdr:row><xdr:rowOff>0');
    expect(diagonalBoxes(flat)).toEqual([]);
  });
});

describe('translateItemsByLanguage', () => {
  it('同一段文字（空白多寡不算）只送一次；跨次呼叫用 memo 共用譯文', async () => {
    const sent: string[] = [];
    const tb = async (texts: string[], langs: string[]) => texts.map(t => {
      sent.push(t);
      return Object.fromEntries(langs.map(l => [l, `EN(${t.trim()})`]));
    });
    const memo = new Map<string, Record<string, string>>();
    const items = [
      { id: 'a', text: '品名：           ' },
      { id: 'b', text: '品名：     ' },
      { id: 'c', text: '批號：' },
    ];
    const r1 = await translateItemsByLanguage(items, ['英文'], '', tb as any, { current: false } as any, undefined, memo);
    expect(sent.length).toBe(2);
    expect(r1.get('a')).toEqual(r1.get('b'));
    const r2 = await translateItemsByLanguage([{ id: 'x', text: '批號：' }], ['英文'], '', tb as any, { current: false } as any, undefined, memo);
    expect(sent.length).toBe(2);
    expect(r2.get('x')).toEqual(r1.get('c'));
  });
});
