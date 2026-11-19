import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import {
  countLines, requiredHeight, fitWorksheet, colWidthToPt, splitSymbolRuns, symbolCharsOf, EXCEL_FIT,
  cellParagraphs, flattenValue,
} from '../src/utils/excelFit';

// 量字寬：中文一個字一個字寬，其他半個字寬（跟實際字型差不多）
const measure = (t: string, size: number) =>
  Array.from(t).reduce((a, ch) => a + (/[⺀-鿿＀-￯]/.test(ch) ? size : size * 0.55), 0);

const newSheet = () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('表單');
  ws.properties.defaultRowHeight = 16.5;
  return ws;
};

describe('countLines / requiredHeight', () => {
  it('英文在空白處換行、中文逐字換行', () => {
    expect(countLines({ text: 'Product Name and Specifications', size: 12 }, 60, measure)).toBeGreaterThanOrEqual(3);
    expect(countLines({ text: '品名規格', size: 12 }, 60, measure)).toBe(1);
    expect(countLines({ text: '品名規格品名規格', size: 12 }, 60, measure)).toBe(2);
  });

  it('行尾空白不多算一行', () => {
    expect(countLines({ text: 'Day   ', size: 12 }, 40, measure)).toBe(1);
  });

  it('高度 = 行數 × 字級 × 行高倍數', () => {
    const h = requiredHeight([{ text: '品名規格', size: 12 }, { text: 'Product', size: 12 }], 200, measure);
    expect(h).toBeCloseTo(2 * 12 * EXCEL_FIT.lineHeight + EXCEL_FIT.paddingV, 1);
  });
});

describe('cellParagraphs / 欄寬', () => {
  it('換行後的那一行用自己的字級（不會被上一段結尾的換行拉大）', () => {
    const paras = cellParagraphs({ richText: [
      { text: 'Drip Loss + Ice\n', font: { size: 14 } },
      { text: '(To be completed by Quality Control)', font: { size: 12 } },
    ] }, { size: 14 });
    expect(paras.map(p => p.size)).toEqual([14, 12]);
  });

  it('欄寬單位跟著預設字型：新細明體 12 的數字寬 8 像素，比 Calibri 寬', () => {
    expect(colWidthToPt(12.6640625, 8)).toBeCloseTo(75.75, 2);
    expect(colWidthToPt(12.6640625)).toBeCloseTo(66.75, 2);
  });

  it('接成一行時，原文尾巴拿來留白的空格一併拿掉', () => {
    const flat = flattenValue({ richText: [{ text: '品名：           ' }, { text: '\nProduct name:' }] });
    expect(flat.richText.map((r: any) => r.text).join('')).toBe('品名：  Product name:');
  });
});

describe('fitWorksheet', () => {
  it('譯文放不下就把列撐高', () => {
    const ws = newSheet();
    ws.getColumn(1).width = 12;
    ws.getRow(1).height = 16.5;
    const c = ws.getCell('A1');
    c.value = '冷凍入庫時間\nTime placed in frozen storage';
    c.font = { size: 12 };
    c.alignment = { wrapText: true };
    const r = fitWorksheet(ws, ['A1'], measure);
    expect(r.rowsChanged).toBe(1);
    expect(ws.getRow(1).height).toBeGreaterThan(50);
  });

  it('跨兩列的合併格：差的高度平均加到兩列', () => {
    const ws = newSheet();
    ws.getColumn(1).width = 10;
    ws.getRow(1).height = 16.5;
    ws.getRow(2).height = 16.5;
    ws.mergeCells('A1:A2');
    const c = ws.getCell('A1');
    c.value = '重量檢測\n(公克)\nWeight Check\n(grams)';
    c.font = { size: 12 };
    c.alignment = { wrapText: true };
    fitWorksheet(ws, ['A1'], measure);
    const total = Number(ws.getRow(1).height) + Number(ws.getRow(2).height);
    expect(total).toBeGreaterThanOrEqual(4 * 12 * EXCEL_FIT.lineHeight);
    expect(Math.abs(Number(ws.getRow(1).height) - Number(ws.getRow(2).height))).toBeLessThan(0.5);
  });

  it('合併好幾欄的格子用合起來的寬度算，不會撐太高', () => {
    const ws = newSheet();
    for (let c = 1; c <= 4; c++) ws.getColumn(c).width = 16;
    ws.getRow(1).height = 40;
    ws.mergeCells('A1:D1');
    ws.getCell('A1').value = '生產管制紀錄表-前處理\nProduction Control Record - Pretreatment';
    ws.getCell('A1').font = { size: 12 };
    ws.getCell('A1').alignment = { wrapText: true, horizontal: 'center' };
    const r = fitWorksheet(ws, ['A1'], measure);
    expect(r.rowsChanged).toBe(0);
    expect(ws.getRow(1).height).toBe(40);
  });

  it('原本沒換行、右邊是空格的：譯文接在同一行往右延伸', () => {
    const ws = newSheet();
    ws.getColumn(1).width = 8;
    const c = ws.getCell('A1');
    c.value = { richText: [{ text: '類別：醃漬類' }, { text: '\nCategory: Pickled Products', font: { name: 'Arial' } }] };
    c.font = { size: 12 };
    c.alignment = { wrapText: true };
    ws.getCell('J1').value = '年';
    const r = fitWorksheet(ws, ['A1'], measure, { originalWrap: new Map([['A1', false]]) });
    expect(r.inlined).toBe(1);
    expect(c.alignment.wrapText).toBe(false);
    expect((c.value as any).richText.map((x: any) => x.text).join('')).toBe('類別：醃漬類  Category: Pickled Products');
    expect(ws.getRow(1).height).toBeUndefined();
  });

  it('右邊沒有空間的就照常換行撐高', () => {
    const ws = newSheet();
    ws.getColumn(1).width = 8;
    ws.getColumn(2).width = 8;
    ws.getCell('B1').value = '有字';
    const c = ws.getCell('A1');
    c.value = '類別：醃漬類\nCategory: Pickled Products';
    c.font = { size: 12 };
    c.alignment = { wrapText: true };
    const r = fitWorksheet(ws, ['A1'], measure, { originalWrap: new Map([['A1', false]]) });
    expect(r.inlined).toBe(0);
    expect(Number(ws.getRow(1).height)).toBeGreaterThan(40);
  });

  it('英文長單字塞不下才加寬欄位，而且有上限', () => {
    const ws = newSheet();
    ws.getColumn(1).width = 6;
    const c = ws.getCell('A1');
    c.value = '規格\nSpecifications';
    c.font = { size: 12 };
    c.alignment = { wrapText: true };
    const r = fitWorksheet(ws, ['A1'], measure);
    expect(r.colsWidened).toBe(1);
    expect(Number(ws.getColumn(1).width)).toBeGreaterThan(6);
    expect(Number(ws.getColumn(1).width)).toBeLessThanOrEqual(6 * EXCEL_FIT.maxWidenRatio + 0.01);
    expect(colWidthToPt(Number(ws.getColumn(1).width))).toBeGreaterThan(colWidthToPt(6));
  });

  it('表單撐開後列印存成每張一頁；原本有設定的不動', () => {
    const ws = newSheet();
    ws.getColumn(1).width = 10;
    ws.getCell('A1').value = '品名規格\nProduct Name and Specifications';
    ws.getCell('A1').font = { size: 12 };
    ws.getCell('A1').alignment = { wrapText: true };
    ws.pageSetup.fitToPage = false;
    const r = fitWorksheet(ws, ['A1'], measure);
    expect(r.printFit).toBe('one-page');
    expect(ws.pageSetup.fitToPage).toBe(true);
    expect(ws.pageSetup.fitToWidth).toBe(1);
    expect(ws.pageSetup.fitToHeight).toBe(1);
  });
});

describe('符號字型', () => {
  const wing = { name: 'Wingdings 2', size: 12 };
  const arial = { name: 'Arial', size: 12 };
  const symbols = symbolCharsOf([{ text: '使用機台: ' }, { text: '£', font: wing }, { text: '真空機C02', font: { name: '標楷體' } }]);

  it('譯文被放進符號字型時，只有原本的符號用符號字型', () => {
    const runs = splitSymbolRuns('£ Vacuum Sealer C02 £ Vacuum Sealer C02-1', wing, arial, symbols);
    expect(runs.map(r => [r.text, r.font.name])).toEqual([
      ['£', 'Wingdings 2'], [' Vacuum Sealer C02 ', 'Arial'], ['£', 'Wingdings 2'], [' Vacuum Sealer C02-1', 'Arial'],
    ]);
  });

  it('譯文用一般字型時，裡面的勾選框照樣顯示成勾選框', () => {
    const runs = splitSymbolRuns('Machine used: £ Sealer', arial, arial, symbols);
    expect(runs.find(r => r.text === '£')?.font.name).toBe('Wingdings 2');
  });
});
