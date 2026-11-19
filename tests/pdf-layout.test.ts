import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import {
  cleanCjkSpacing, startsNewItem, detectGridLines, layoutTextItems, wrapText, fitText,
  sampleBlockColors, fontGroups, findCell, layoutSegments, pageTextSize, groupFontSizes, isCenteredIn,
  lineTokens, scriptMetrics, graphemes,
  type GridLines, type RawTextItem, type Segment,
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

  it('被填寫空格切開的句子接成一句，下面的簽章欄不會被併進來', () => {
    expect(texts).toContain('保險對象因＿＿，無法親自辦理自墊醫療費用核退相關事宜，特委託＿＿代辦。');
    expect(texts).toContain('委託人簽章');
    expect(texts.some(t => t.startsWith('✽保險對象因＿＿，無法親自辦理將核退費用優先繳付'))).toBe(true);
  });

  it('同一格、同一排、並排清單用同一個字級', () => {
    const groups = fontGroups(blocks).map(g => g.map(i => blocks[i].text));
    const groupOf = (t: string) => groups.find(g => g.includes(t)) || [];
    // 同一格的清單
    expect(groupOf('□1.醫療費用收據正本及費用明細')).toEqual(expect.arrayContaining(['□5.其他', '□4.住院案件者:出院病歷摘要']));
    // 同一排的欄位
    expect(groupOf('姓名')).toEqual(expect.arrayContaining(['身分證字號', '出生日期', '年月日']));
    // 表頭
    expect(groupOf('臺灣地區內')).toEqual(expect.arrayContaining(['臺灣地區外', '住院部分負擔超過上限']));
    // 並排的三欄清單
    expect(groupOf('□3.其他')).toEqual(expect.arrayContaining(['□5.其他']));
    // 窄的標籤格不要跟旁邊的大格綁在一起
    expect(groupOf('收據非正本聲明')).not.toContain('簽章');
  });

  it('同一格由上往下的清單標成可以整格重排', () => {
    const stacked = (t: string) => blocks.find(b => b.text === t)?.stack?.id;
    expect(stacked('□1.醫療費用收據正本及費用明細')).toBeDefined();
    expect(stacked('□1.醫療費用收據正本及費用明細')).toBe(stacked('□5.其他'));
    // 並排的（同一列好幾項）不能重排
    expect(stacked('□1.保險對象簽章')).toBeUndefined();
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

  it('上限調低到本來放得下的字級，結果還是那個字級', () => {
    const free = fitText('Date of Birth', 48, 18, 10.35, 4.5, measure);
    const capped = fitText('Date of Birth', 48, 18, free.fontSize, 4.5, measure);
    expect(capped.fontSize).toBe(free.fontSize);
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

// ── 會議記錄表「一、」那一格：(四)的第二行行尾「，」凸出格線 ──────────────────
// 位置照使用者原檔量的（pt）：決議事項欄 63–376，右邊是執行部門欄 376–473
describe('行尾標點凸出格線', () => {
  const fs = 12;
  const seg = (text: string, left: number, top: number, right?: number): Segment => {
    const r = right ?? left + Array.from(text).length * fs * 0.95;
    return { text, left, right: r, top, bottom: top + fs, fontSize: fs, glyphs: [{ left, right: r, top, bottom: top + fs }] };
  };
  const segs = [
    seg('現況聚泰自有宿舍由內部管理人員自行管理之問題點如下:', 67.5, 240, 374),
    seg('(一)現有宿舍沒有使用執照，建安申報有極大審查不合格的', 67.5, 260, 374),
    seg('風險。', 90, 280),
    seg('(二)因應三期擴建，外籍移工人力將由現在總人數 47 人增', 67.5, 300, 374),
    seg('加到約 67 人(約增加 20 人)，現有宿舍空間不足。', 90, 320, 330),
    seg('(三)夏季用電支出偏高（每月逾 12.5 萬元），加上維修及', 67.5, 340, 374),
    seg('雜支支出，整體營運費用無法管控。', 90, 360),
    seg('(四)缺乏專業宿舍管理制度，包括公共區域清潔、生活規', 67.5, 380, 374),
    seg('範、用電明火規定、緊急聯繫機制等均無有效管理方案，', 90, 400, 380.5),   // 「，」凸出格線 4.5pt
    seg('安全危害風險高。', 90, 420),
  ];
  const grid: GridLines = {
    v: [27, 63, 376, 473, 538, 569].map(x => ({ x, y0: 236, y1: 444 })),
    h: [236, 444].map(y => ({ y, x0: 27, x1: 569 })),
  };

  it('凸出格線不到半個字的那一行還是算在原本那一格', () => {
    const cell = findCell(segs[8], grid, 595, 842);
    expect([cell.left, cell.right]).toEqual([63, 376]);
  });

  it('(四) 三行接成一段、跟同格的段落同一格，譯文不寫過格線', () => {
    const blocks = layoutSegments(segs, 595, 842, grid);
    const four = blocks.find(b => b.text.startsWith('(四)'))!;
    expect(four.text).toContain('均無有效管理方案');
    expect(four.text).toContain('安全危害風險高');
    expect(blocks.filter(b => b.text.includes('安全危害')).length).toBe(1);
    for (const b of blocks) {
      expect(b.cell?.right).toBe(376);
      expect(b.box.right).toBeLessThanOrEqual(376);
    }
    // 整格一起重排（同一個清單）
    expect(new Set(blocks.map(b => b.stack?.id)).size).toBe(1);
    expect(four.stack).toBeDefined();
  });
});

describe('整頁字級一致', () => {
  it('pageTextSize：依字數加權的中位數；大標題不算也不受限', () => {
    const cap = pageTextSize([
      { size: 22, sourceSize: 22, chars: 20 },          // 大標題
      { size: 9, sourceSize: 12, chars: 600 },          // 內文
      { size: 8, sourceSize: 12, chars: 500 },
      { size: 12, sourceSize: 12, chars: 18 },          // 短標籤（空間大）
      { size: 12, sourceSize: 12, chars: 25 },
      { size: 6, sourceSize: 12, chars: 40 },
    ])!;
    expect(cap.size).toBe(9);
    expect(cap.applies).toEqual([false, true, true, true, true, true]);
  });

  it('groupFontSizes：同一排特別擠的一格自己小，其他格不跟著縮；同一格的一定一樣', () => {
    const cell = (left: number, right: number, top = 100, bottom = 120) =>
      ({ left, right, top, bottom, hasLeft: true, hasRight: true, hasTop: true, hasBottom: true });
    const row = [
      { cell: cell(0, 40), fontSize: 12, vertical: false, text: '項次' },
      { cell: cell(40, 350), fontSize: 12, vertical: false, text: '決議事項' },
      { cell: cell(350, 450), fontSize: 12, vertical: false, text: '執行部門（人員）' },
      { cell: cell(450, 480), fontSize: 12, vertical: false, text: '備註' },
    ];
    const sizes = groupFontSizes(row, [9.5, 12, 9, 5.5]);
    expect([0, 1, 2].map(k => sizes.get(k))).toEqual([9, 9, 9]);
    expect(sizes.get(3)).toBe(5.5);

    const sameCell = [
      { cell: cell(0, 300, 100, 300), fontSize: 12, vertical: false, text: '(一)很長的一段' },
      { cell: cell(0, 300, 100, 300), fontSize: 12, vertical: false, text: '(二)短' },
    ];
    const s2 = groupFontSizes(sameCell, [5, 10]);
    expect([s2.get(0), s2.get(1)]).toEqual([5, 5]);
  });

  it('表格外同一行的簽名欄標籤分成同一組', () => {
    const lineBlock = (text: string, left: number) =>
      ({ fontSize: 12, vertical: false, text, source: { left, right: left + 60, top: 700, bottom: 712 } });
    const groups = fontGroups([lineBlock('呈董事長：', 30), lineBlock('總經理：', 160), lineBlock('陳廠長：', 290)]);
    expect(groups).toEqual([[0, 1, 2]]);
  });
});

describe('排版小細節', () => {
  it('左括號不會落單留在行尾', () => {
    const lines = wrapText('Responsible Department (Personnel)', 95, 10, measure);
    for (const l of lines) expect(l.endsWith('(')).toBe(false);
    expect(lines.join(' ')).toContain('(Personnel)');
  });

  it('兩邊對齊寫滿整格的段落不算置中；表頭標籤算', () => {
    const c = { left: 63, right: 376, top: 0, bottom: 100, hasLeft: true, hasRight: true, hasTop: true, hasBottom: true };
    expect(isCenteredIn({ left: 68, right: 371, top: 10, bottom: 90 }, c)).toBe(false);
    expect(isCenteredIn({ left: 180, right: 259, top: 10, bottom: 22 }, c)).toBe(true);
  });

  it('靠右的一行（記錄：陳淑燕 4/29）譯文往左延伸、靠右對齊', () => {
    const s: Segment = { text: '記錄：陳淑燕 4/29', left: 470, right: 565, top: 760, bottom: 772, fontSize: 12, glyphs: [] };
    const body: Segment = { text: '這是一行很長的內文用來決定版心寬度這是一行很長的內文用來決定版心寬度', left: 30, right: 565, top: 700, bottom: 712, fontSize: 12, glyphs: [] };
    const blocks = layoutSegments([body, s], 595, 842);
    const rec = blocks.find(b => b.text.startsWith('記錄'))!;
    expect(rec.alignRight).toBe(true);
    expect(rec.box.left).toBeLessThan(100);
  });

  it('表格外置中的標題（會議記錄表）譯文以頁面中心左右展開、置中', () => {
    const title: Segment = { text: '會議記錄表', left: 242, right: 353, top: 40, bottom: 62, fontSize: 22, glyphs: [] };
    const page: Segment = { text: '第 1 頁', left: 530, right: 565, top: 70, bottom: 82, fontSize: 12, glyphs: [] };
    const body: Segment = { text: '這是一行很長的內文用來決定版心寬度這是一行很長的內文用來決定版心寬度', left: 30, right: 565, top: 120, bottom: 132, fontSize: 12, glyphs: [] };
    const blocks = layoutSegments([title, page, body], 595, 842);
    const t = blocks.find(b => b.text === '會議記錄表')!;
    expect(t.alignCenter).toBe(true);
    expect((t.box.left + t.box.right) / 2).toBeCloseTo(297.5, 0);
    expect(t.box.right - t.box.left).toBeGreaterThan(400);
    // 底下的「第 1 頁」在展開後的範圍裡：譯文高度不會蓋到它
    expect(t.box.bottom).toBeLessThanOrEqual(70);
    // 靠左的一般段落不算置中
    expect(blocks.find(b => b.text.startsWith('這是'))!.alignCenter).toBeFalsy();
  });

  it('置中的窄格標籤可以用整格寬（Remarks 不會被拆成兩半）', () => {
    const grid: GridLines = {
      v: [538, 569].map(x => ({ x, y0: 180, y1: 222 })),
      h: [180, 222].map(y => ({ y, x0: 538, x1: 569 })),
    };
    // 窄欄裡「備／註」上下兩行，每行置中
    const s1: Segment = { text: '備', left: 547.5, right: 559.5, top: 186, bottom: 198, fontSize: 12, glyphs: [] };
    const s2: Segment = { text: '註', left: 547.5, right: 559.5, top: 204, bottom: 216, fontSize: 12, glyphs: [] };
    const [b] = layoutSegments([s1, s2], 595, 842, grid);
    expect(b.text).toBe('備註');
    expect(b.box.left).toBeLessThan(541);
    expect(b.box.right - b.box.left).toBeGreaterThan(26);
  });
});

describe('越南文、印尼文、泰文', () => {
  const TH = 'ขาดระบบบริหารหอพักอย่างมืออาชีพ ทั้งการทำความสะอาดพื้นที่ส่วนกลาง กฎระเบียบการใช้ชีวิต ข้อกำหนดการใช้ไฟฟ้าและเปลวไฟ';

  it('泰文照詞斷行：不從字中間斷開，母音、聲調不會跑到下一行開頭', () => {
    const words = new Set(lineTokens(TH).map(w => w.trim()).filter(Boolean));
    expect(words.has('บริหาร')).toBe(true);
    expect(words.has('หอพัก')).toBe(true);
    for (const width of [60, 90, 140]) {
      const lines = wrapText(TH, width, 10, measure);
      expect(lines.length).toBeGreaterThan(1);
      expect(lines.join('').replace(/\s/g, '')).toBe(TH.replace(/\s/g, ''));
      for (const l of lines) expect(/^\p{M}/u.test(l)).toBe(false);
      // 每一行的頭尾都落在詞的邊界上
      let pos = 0;
      const flat = TH.replace(/\s/g, '');
      const boundaries = new Set<number>([0]);
      for (const w of lineTokens(TH)) { if (!w.trim()) continue; pos += w.replace(/\s/g, '').length; boundaries.add(pos); }
      let at = 0;
      for (const l of lines) { at += l.replace(/\s/g, '').length; expect(boundaries.has(at)).toBe(true); }
      expect(at).toBe(flat.length);
    }
  });

  it('越南文分開寫的聲調符號黏著字母', () => {
    const decomposed = 'Ky\u0301 tu\u0301c xa\u0301'.normalize('NFD');
    expect(graphemes('a\u0301b')).toEqual(['a\u0301', 'b']);
    for (const l of wrapText(decomposed, 18, 10, measure)) expect(/^\p{M}/u.test(l)).toBe(false);
  });

  it('金額不拆開；標點黏前面、左括號黏後面', () => {
    const t = lineTokens('biaya listrik (lebih dari NT$125.000 per bulan), total NT$2,000.');
    expect(t).toContain('NT$125.000');
    expect(t).toContain('NT$2,000.');
    expect(t).toContain('(lebih');
    expect(t).toContain('bulan),');
    expect(lineTokens('風險，建安')).toEqual(['風', '險，', '建', '安']);
  });

  it('行高跟著語言：泰文最寬、越南文稍寬', () => {
    const th = scriptMetrics('หอพัก').lineHeight;
    const vi = scriptMetrics('Thiếu chế độ quản lý').lineHeight;
    // 知道是越南文的頁面：沒有疊兩層符號的「Ghi chú」也用同一個行高
    expect(scriptMetrics('Ghi chú', 1.15, '越南文').lineHeight).toBeCloseTo(vi);
    expect(scriptMetrics('No.', 1.15, '泰文').lineHeight).toBeCloseTo(th);
    const en = scriptMetrics('Dormitory').lineHeight;
    expect(th).toBeGreaterThan(vi);
    expect(vi).toBeGreaterThan(en);
    expect(en).toBeCloseTo(1.15);
    expect(scriptMetrics('Asrama pekerja').lineHeight).toBeCloseTo(1.15);   // 印尼文就是一般拉丁字母
  });
});
