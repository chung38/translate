import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import {
  cleanCjkSpacing, startsNewItem, detectGridLines, layoutTextItems, wrapText, fitText,
  sampleBlockColors, type GridLines, type RawTextItem,
} from '../src/utils/pdfLayout';

// 健保自墊醫療費用核退申請書：表格、直排標籤、字距排版都有，拿來當回歸測試
const fixture = JSON.parse(
  fs.readFileSync(path.join(__dirname, 'fixtures/nhi-reimbursement-form.json'), 'utf8')
);

const pageOf = (i: number) => {
  const pg = fixture.pages[i];
  const items: RawTextItem[] = pg.items.map((r: any[]) => ({
    str: r[0], transform: [r[1], r[2], r[3], r[4], r[5], r[6]], width: r[7],
  }));
  const grid: GridLines = {
    h: pg.grid.h.map((r: number[]) => ({ y: r[0], x0: r[1], x1: r[2] })),
    v: pg.grid.v.map((r: number[]) => ({ x: r[0], y0: r[1], y1: r[2] })),
  };
  return { items, grid, w: pg.w, h: pg.h };
};

// 等寬字的量測：每個字 = 0.55 × 字級（中文算 1 個字寬）
const measure = (t: string, fs: number) =>
  Array.from(t).reduce((a, ch) => a + (/[㐀-鿿　-〿＀-￯]/.test(ch) ? fs : fs * 0.55), 0);

describe('cleanCjkSpacing', () => {
  it('去掉中文字距、接回被拆開的數字', () => {
    expect(cleanCjkSpacing('臺 北 業 務 組')).toBe('臺北業務組');
    expect(cleanCjkSpacing('依 全 民 健 康 保 險 法 第 8 1 條')).toBe('依全民健康保險法第81條');
    expect(cleanCjkSpacing('3 2 0 2 1 6 桃園市中壢區中山東路 3 段 525 號')).toBe('320216桃園市中壢區中山東路3段525號');
  });
  it('字距排版的網址接回去', () => {
    expect(cleanCjkSpacing('網 址 ： h t t p： / / w w w. s e f . o r g . t w ） 申 請')).toContain('www.sef.org.tw');
  });
  it('英文單字之間的空白保留', () => {
    expect(cleanCjkSpacing('Hello world test')).toBe('Hello world test');
    expect(cleanCjkSpacing('健康存摺 APP (醫療查詢)')).toBe('健康存摺 APP (醫療查詢)');
  });
});

describe('startsNewItem', () => {
  it('認得常見的條列開頭', () => {
    for (const t of ['(一)依全民', '（二）本申請書', '一、注意事項', '1.□門診', '□臺灣地區內', '※注意', '註：持收據']) {
      expect(startsNewItem(t)).toBe(true);
    }
    for (const t of ['有重覆申請或受領情事', '人、法定繼承人帳戶', '內(不含例假日)']) {
      expect(startsNewItem(t)).toBe(false);
    }
  });
});

describe('layoutTextItems — 申請書第 1 頁（表格）', () => {
  const { items, grid, w, h } = pageOf(0);
  const blocks = layoutTextItems(items, w, h, grid);
  const texts = blocks.map(b => b.text);
  const find = (t: string) => blocks.find(b => b.text === t);

  it('同一列不同格子分開，不會併成一句', () => {
    expect(texts).toContain('就醫地區');
    expect(texts).toContain('□臺灣地區內 □臺灣地區外');
    expect(texts).toContain('受理號碼');
    expect(texts).toContain('受理日期');
    expect(texts.some(t => t.includes('受理號碼') && t.includes('受理日期'))).toBe(false);
  });

  it('換行的句子接回同一段（中間隔著簽章底線也一樣）', () => {
    expect(texts).toContain(
      '醫療費用收據正本因□遺失、毀損□另有它用而無法檢送，改提供非收據正本申請核退，嗣後如有與本聲明不符之情形、或發現有重覆申請或受領情事，願負法律責任，特此聲明。'
    );
    expect(texts.some(t => t.startsWith('□於就醫時未能及時繳驗健保卡') && t.endsWith('申請退費'))).toBe(true);
    expect(texts).toContain('□5.其他');
  });

  it('直排標籤接成一個詞，雙欄的一欄一欄讀，而且標成要轉向', () => {
    for (const t of ['保險對象／矯正機關收容人', '就醫情形', '申請金額', '自墊費用原因(請於原因欄詳述)', '付款方式', '檢附書據', '申請人／簽章', '委託書']) {
      expect(find(t)?.vertical, t).toBe(true);
    }
    // 不是窄格子的不要轉
    expect(find('縣市')?.vertical).toBe(false);
    expect(find('收據非正本聲明')?.vertical).toBe(false);
  });

  it('分散對齊的表頭整格當一個詞', () => {
    expect(texts).toContain('臺灣地區內');
    expect(texts).toContain('臺灣地區外');
    expect(texts).toContain('中華民國');
    expect(texts).not.toContain('臺');
  });

  it('圓圈編號、純數字、勾選框不送翻譯', () => {
    expect(texts.some(t => /^[-○\d\s]+$/.test(t))).toBe(false);
    expect(texts).not.toContain('□□□□□□');
  });

  it('譯文範圍不超出格子、彼此不重疊', () => {
    for (const b of blocks) {
      if (b.cell?.hasRight) expect(b.box.right).toBeLessThanOrEqual(b.cell.right + 0.01);
      if (b.cell?.hasBottom && !b.vertical) expect(b.box.bottom).toBeLessThanOrEqual(b.cell.bottom + 0.01);
    }
    const flat = blocks.filter(b => !b.vertical);
    let overlaps = 0;
    for (let i = 0; i < flat.length; i++) {
      for (let j = i + 1; j < flat.length; j++) {
        const a = flat[i].box, c = flat[j].box;
        const ix = Math.min(a.right, c.right) - Math.max(a.left, c.left);
        const iy = Math.min(a.bottom, c.bottom) - Math.max(a.top, c.top);
        if (ix > 1 && iy > 1) overlaps++;
      }
    }
    expect(overlaps).toBe(0);
  });

  it('直排標籤讓開上面的圓圈編號', () => {
    const label = find('自墊費用原因(請於原因欄詳述)')!;
    expect(label.box.top).toBeGreaterThan(330);   // ⑤ 在 y≈ 320–332
  });
});

describe('layoutTextItems — 填寫說明第 2 頁（段落）', () => {
  const { items, grid, w, h } = pageOf(1);
  const blocks = layoutTextItems(items, w, h, grid);
  const texts = blocks.map(b => b.text);

  it('整段接起來，下一項不會被併進來', () => {
    const p3 = texts.find(t => t.startsWith('(三)年滿十八歲'))!;
    expect(p3.endsWith('並附身分證明文件及聲明書。')).toBe(true);
    expect(texts.find(t => t.startsWith('五、給付方式'))!.endsWith('所刊載之匯率計算。')).toBe(true);
    expect(texts.some(t => t.includes('(三)年滿') && t.includes('(四)臺灣'))).toBe(false);
  });

  it('表格裡的地址、轄區分開，電話號碼不送翻譯', () => {
    expect(texts).toContain('320216桃園市中壢區中山東路3段525號');
    expect(texts).toContain('桃園、新竹、苗栗');
    expect(texts).toContain('地址');
    expect(texts).toContain('轄區範圍');
    expect(texts.some(t => t.includes('4339111'))).toBe(false);
  });
});

describe('detectGridLines', () => {
  const W = 200, H = 120, S = 2;
  const img = () => {
    const a = new Uint8ClampedArray(W * H * 4).fill(255);
    return a;
  };
  const dot = (a: Uint8ClampedArray, x: number, y: number) => {
    const i = (y * W + x) * 4;
    a[i] = a[i + 1] = a[i + 2] = 0;
  };

  it('找到格線、忽略文字筆畫', () => {
    const a = img();
    for (let x = 10; x < 190; x++) { dot(a, x, 20); dot(a, x, 21); dot(a, x, 100); }   // 兩條橫線（一條粗）
    for (let y = 20; y <= 100; y++) { dot(a, 10, y); dot(a, 100, y); dot(a, 189, y); } // 三條直線
    for (let x = 30; x < 70; x++) dot(a, x, 60);   // 文字裡的長橫畫
    const grid = detectGridLines(a, W, H, S, [{ left: 14, top: 26, right: 36, bottom: 34 }]);
    expect(grid.h.length).toBe(2);
    expect(grid.v.length).toBe(3);
    expect(grid.h[0].y).toBeCloseTo(10.5, 0);
  });
});

describe('wrapText / fitText', () => {
  it('不會把英文單字從中間拆開，寧可縮小字', () => {
    const fit = fitText('Medical Treatment', 40, 30, 9, 4.5, measure);
    expect(fit.lines.join(' ')).toBe('Medical Treatment');
    expect(fit.lines.every(l => !/^[a-z]/.test(l) || l.startsWith('Treatment') || l.startsWith('Medical'))).toBe(true);
  });

  it('短標籤優先排成一行', () => {
    const fit = fitText('Office (', 34, 20, 10, 4.5, measure);
    expect(fit.lines).toEqual(['Office (']);
  });

  it('單字真的放不下時，最小字級才逐字母拆', () => {
    const lines = wrapText('Supercalifragilistic', 20, 4.5, measure);
    expect(lines.length).toBeGreaterThan(1);
  });

  it('塞不下時回傳最小字級', () => {
    const fit = fitText('A very long sentence '.repeat(20), 50, 10, 9, 4.5, measure);
    expect(fit.fits).toBe(false);
    expect(fit.fontSize).toBe(4.5);
  });
});

describe('sampleBlockColors', () => {
  it('紅字沿用紅色、黑字用深色', () => {
    const W = 40, H = 20;
    const make = (r: number, g: number, b: number) => {
      const a = new Uint8ClampedArray(W * H * 4).fill(255);
      for (let y = 5; y < 15; y++) for (let x = 5; x < 35; x++) {
        if ((x + y) % 2) continue;
        const i = (y * W + x) * 4; a[i] = r; a[i + 1] = g; a[i + 2] = b;
      }
      return a;
    };
    const block = { glyphs: [{ left: 5, top: 5, right: 35, bottom: 15 }] };
    expect(sampleBlockColors(make(220, 20, 20), W, H, 1, block).ink).not.toBe('#111827');
    expect(sampleBlockColors(make(30, 30, 30), W, H, 1, block).ink).toBe('#111827');
    expect(sampleBlockColors(make(30, 30, 30), W, H, 1, block).background).toBe('#ffffff');
  });
});
