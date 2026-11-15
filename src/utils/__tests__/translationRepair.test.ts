import { describe, it, expect, vi } from 'vitest';

// useTranslation.ts 會連到 firebase，測純函式時擋掉就好
vi.mock('../../firebase', () => ({ db: {}, auth: { currentUser: null } }));

import { rebuildWithSourceTags, repairTranslationTags } from '../../hooks/useTranslation';

const stripTags = (t: string) => t.replace(/\[\/?f\d+\]/g, '');
const HAN = /[㐀-鿿]/;

describe('rebuildWithSourceTags', () => {
  it('泰文（沒有空白分詞）不會把原文中文塞回譯文', () => {
    const source =
      '[f0]工作場所通道、樓梯，應[/f0][f1]保持[/f1][f2]暢通，不可[/f2][f3]堆置器材[/f3][f4]、機件、物料等[/f4]';
    const thai = 'ทางเดินและบันไดในสถานที่ทำงานต้องให้โล่งอยู่เสมอ ห้ามวางอะไหล่ วัสดุ ฯลฯ';
    const out = rebuildWithSourceTags(source, thai);

    expect(HAN.test(stripTags(out))).toBe(false);          // 譯文裡不可以有中文
    expect(stripTags(out).replace(/\s+/g, '')).toBe(thai.replace(/\s+/g, '')); // 泰文一個字都不能少
    expect(out.match(/\[f\d+\]/g)?.length).toBe(5);        // 標籤數量要保持一致
  });

  it('譯文是空的時候回傳空標籤，不退回原文', () => {
    const source = '[f0]禁止吸菸[/f0][f1]違者處分[/f1]';
    const out = rebuildWithSourceTags(source, '');
    expect(stripTags(out).trim()).toBe('');
    expect(HAN.test(stripTags(out))).toBe(false);
  });

  it('越南文（有空白分詞）仍然依比例分配到各標籤', () => {
    const source = '[f0]工作場所通道[/f0][f1]應保持暢通[/f1]';
    const viet = 'Lối đi tại nơi làm việc phải luôn thông thoáng';
    const out = rebuildWithSourceTags(source, viet);
    const parts = [...out.matchAll(/\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g)].map(m => m[2]);
    expect(parts.every(p => p.trim().length > 0)).toBe(true);
    // 跨標籤的詞間空白由 sanitizeOutputText 另外補，這裡只確認字沒有少、順序沒亂
    expect(stripTags(out).replace(/\s+/g, '')).toBe(viet.replace(/\s+/g, ''));
  });

  it('排版用的空白標籤留空，字不會被它吃掉', () => {
    const source = '[f0]使用單位: 生產部[/f0][f1]          [/f1][f2]年[/f2]';
    const viet = 'Đơn vị sử dụng: Bộ phận sản xuất Năm';
    const out = rebuildWithSourceTags(source, viet);
    const parts = [...out.matchAll(/\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g)].map(m => m[2]);
    expect(parts[1].trim()).toBe('');                       // 空白標籤不分到字
    expect(stripTags(out).replace(/\s+/g, '')).toBe(viet.replace(/\s+/g, '')); // 字一個都沒少
  });
});

describe('repairTranslationTags', () => {
  it('標籤本來就對得上時原樣返回', () => {
    const source = '[f0]安全第一[/f0]';
    const translated = '[f0]ความปลอดภัยมาก่อน[/f0]';
    expect(repairTranslationTags(source, translated)).toBe(translated);
  });

  it('原文沒有標籤時不動譯文', () => {
    expect(repairTranslationTags('安全第一', 'An toàn là trên hết')).toBe('An toàn là trên hết');
  });
});
