import { describe, it, expect } from 'vitest';
import {
  langBase,
  alreadyHasLanguage,
  neededLanguages,
  isEchoTranslation,
  stripInvalidXmlChars,
} from '../documentProcessors';

// 這幾組測試專門守著前面實際出過問題的地方。
// 共通點是：型別檢查抓不到，執行時也不會噴錯，只會靜靜地不作用。

describe('langBase — App.tsx 傳的是中文顯示名，不是 ISO 代碼', () => {
  it('中文顯示名要能對應', () => {
    expect(langBase('越南文')).toBe('vi');
    expect(langBase('英文')).toBe('en');
    expect(langBase('繁體中文')).toBe('zh');
    expect(langBase('泰文')).toBe('th');
    expect(langBase('印尼文')).toBe('id');
  });

  it('ISO 代碼也要能對應', () => {
    expect(langBase('vi')).toBe('vi');
    expect(langBase('vi-VN')).toBe('vi');
    expect(langBase('en-US')).toBe('en');
    expect(langBase('zh-TW')).toBe('zh');
  });

  it('認不得的值不要當成某個語言', () => {
    expect(langBase('')).toBe('other');
    expect(langBase('克林貢語')).toBe('other');
  });
});

describe('alreadyHasLanguage — 原檔已經有譯文就不要再翻一次', () => {
  it('中文原文對越南文 → 需要翻譯', () => {
    expect(alreadyHasLanguage('企業文化', '越南文')).toBe(false);
  });

  it('中文原文對中文 → 不需要翻譯', () => {
    expect(alreadyHasLanguage('企業文化', '繁體中文')).toBe(true);
  });

  it('已經是越南文 → 不需要再翻', () => {
    expect(alreadyHasLanguage('Chính sách chất lượng', '越南文')).toBe(true);
  });

  it('越南文不可以被誤判成英文', () => {
    expect(alreadyHasLanguage('Chính sách chất lượng', '英文')).toBe(false);
  });

  it('中英混排：翻英文時跳過，翻越南文時仍要翻', () => {
    const mixed = '多說好話、多做好事、多幫助別人 Say good things, do good deeds, and help others more.';
    expect(alreadyHasLanguage(mixed, '英文')).toBe(true);
    expect(alreadyHasLanguage(mixed, '越南文')).toBe(false);
  });

  it('純數字、日期、空字串不用翻', () => {
    expect(alreadyHasLanguage('2024/10/05', '越南文')).toBe(true);
    expect(alreadyHasLanguage('   ', '越南文')).toBe(true);
    expect(alreadyHasLanguage('123 456', '越南文')).toBe(true);
  });

  it('判斷時要先去掉 [fN] 標籤', () => {
    expect(alreadyHasLanguage('[f0]企業文化[/f0]', '繁體中文')).toBe(true);
  });
});

describe('neededLanguages — 只送真正缺的語言去 API', () => {
  it('已經有越南文時只剩英文要翻', () => {
    expect(neededLanguages('Chính sách chất lượng', ['越南文', '英文'])).toEqual(['英文']);
  });

  it('兩種都缺就兩種都送', () => {
    expect(neededLanguages('品質政策', ['越南文', '英文'])).toEqual(['越南文', '英文']);
  });

  it('都不缺就一個都不送', () => {
    expect(neededLanguages('2024/10/05', ['越南文', '英文'])).toEqual([]);
  });
});

describe('isEchoTranslation — 譯文等於原文就不要重複貼', () => {
  it('型號、代碼被原樣回傳', () => {
    expect(isEchoTranslation('ISO9001:2015', 'ISO9001:2015')).toBe(true);
    expect(isEchoTranslation('CNC', 'CNC')).toBe(true);
  });

  it('只差空白或大小寫也算相同', () => {
    expect(isEchoTranslation('Customer Satisfaction', 'customer  satisfaction')).toBe(true);
  });

  it('空字串視為沒有譯文', () => {
    expect(isEchoTranslation('品質政策', '')).toBe(true);
  });

  it('真的翻譯過就不算', () => {
    expect(isEchoTranslation('品質政策', 'Chính sách chất lượng')).toBe(false);
  });

  it('比對前要先去掉標籤', () => {
    expect(isEchoTranslation('[f0]CNC[/f0]', '[f0]CNC[/f0]')).toBe(true);
  });
});

describe('正則多跳脫一層的回歸測試', () => {
  // 這一類 bug 出現過三次：documentProcessors 的 stripTags 寫成 \f（換頁字元）、
  // useTranslation 的四條標籤正則、以及 insertTranslatedSlides 的 nextFreeIndex。
  it('\\f 是換頁字元，不是字母 f', () => {
    expect(/\[\/?\f\d+\]/.test('[f0]')).toBe(false);   // 壞掉的寫法
    expect(/\[\/?f\d+\]/.test('[f0]')).toBe(true);     // 正確的寫法
  });

  it('樣板字串裡的 \\\\d 會變成「反斜線 + d」', () => {
    expect(new RegExp(`^(\\\\d+)$`).test('123')).toBe(false);  // 壞掉的寫法
    expect(new RegExp(`^(\\d+)$`).test('123')).toBe(true);     // 正確的寫法
  });
});

describe('stripInvalidXmlChars — 譯文裡的非法字元會讓整份檔案打不開', () => {
  it('去掉 XML 1.0 不允許的控制字元', () => {
    expect(stripInvalidXmlChars('Bản\u000Bdịch')).toBe('Bảndịch');
    expect(stripInvalidXmlChars('a\u0001b\u001Fc')).toBe('abc');
    expect(stripInvalidXmlChars('x\uFFFEy\uFFFFz')).toBe('xyz');
  });

  it('保留合法的空白字元', () => {
    expect(stripInvalidXmlChars('a\tb\nc\rd')).toBe('a\tb\nc\rd');
  });

  it('去掉落單的代理對，但保留完整的表情符號', () => {
    expect(stripInvalidXmlChars('a\uD83Db')).toBe('ab');       // 落單高位
    expect(stripInvalidXmlChars('a\uDE00b')).toBe('ab');       // 落單低位
    expect(stripInvalidXmlChars('a\uD83D\uDE00b')).toBe('a\uD83D\uDE00b'); // 完整的 😀
  });

  it('一般文字不受影響', () => {
    expect(stripInvalidXmlChars('Chính sách chất lượng 品質政策')).toBe('Chính sách chất lượng 品質政策');
  });
});
