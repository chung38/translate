import React from 'react';

/**
 * 輸出預覽 —— 用工具實際的排版方式，把「翻譯完會長什麼樣」直接畫出來。
 * 選了哪些語言、選了哪種版面模式，這裡就立刻反映。
 * 尤其是 PPTX 的「同頁對照 / 另加譯文頁」，用看的比用讀的清楚。
 */

// 取自實際會拿來翻譯的那類文件（工安宣導、工作須知）
const SAMPLE_SRC = { h: '安全第一', p: '請確實佩戴個人防護具' };

const SAMPLE: Record<string, { h: string; p: string }> = {
  '越南文':   { h: 'An toàn là trên hết', p: 'Vui lòng đeo đầy đủ thiết bị bảo hộ cá nhân' },
  '泰文':     { h: 'ความปลอดภัยมาก่อน',   p: 'โปรดสวมอุปกรณ์ป้องกันส่วนบุคคลให้ครบถ้วน' },
  '印尼文':   { h: 'Keselamatan yang utama', p: 'Harap kenakan alat pelindung diri dengan lengkap' },
  '英文':     { h: 'Safety first',          p: 'Please wear your personal protective equipment' },
  '繁體中文': { h: '安全第一',              p: '請確實佩戴個人防護具' },
};

const Source = () => (
  <>
    <div className="sheet__h">{SAMPLE_SRC.h}</div>
    <div className="sheet__p">{SAMPLE_SRC.p}</div>
  </>
);

const Translated = ({ lang }: { lang: string }) => {
  const s = SAMPLE[lang];
  if (!s) return null;
  return (
    <>
      <div className="sheet__h--t">{s.h}</div>
      <div className="sheet__p--t">{s.p}</div>
    </>
  );
};

interface Props {
  languages: string[];
  /** 只有 pptx 才有「另加譯文頁」這個選項 */
  layoutMode?: 'append' | 'duplicate-slide';
  hasPptx?: boolean;
}

export const OutputPreview: React.FC<Props> = ({ languages, layoutMode = 'append', hasPptx = false }) => {
  const langs = languages.filter(l => SAMPLE[l]);
  const showTwoPages = hasPptx && layoutMode === 'duplicate-slide';

  return (
    <div className="preview">
      <div className="preview__bar">
        <span>翻譯後長這樣</span>
        <span className="tag">{showTwoPages ? '另加譯文頁' : '同頁對照'}</span>
      </div>

      {langs.length === 0 ? (
        <div className="preview__body">
          <div className="sheet">
            <div className="sheet__tag">原文</div>
            <Source />
            <div className="sheet__gap" />
            <p className="sheet__p" style={{ color: 'var(--muted)' }}>
              選好語言之後，這裡會顯示譯文放進文件的位置與大小。
            </p>
          </div>
        </div>
      ) : showTwoPages ? (
        <div className="preview__body fade-in" key={`dup-${langs.join()}`}>
          <div className="sheet">
            <div className="sheet__tag">原稿頁</div>
            <Source />
          </div>
          <div className="sheet">
            <div className="sheet__tag">延伸頁</div>
            {langs.map((l, i) => (
              <React.Fragment key={l}>
                {i > 0 && <div className="sheet__gap" />}
                <Translated lang={l} />
              </React.Fragment>
            ))}
          </div>
        </div>
      ) : (
        <div className="preview__body fade-in" key={`app-${langs.join()}`}>
          <div className="sheet">
            <div className="sheet__tag">同一段落</div>
            <Source />
            {langs.map(l => (
              <React.Fragment key={l}>
                <div className="sheet__gap" />
                <Translated lang={l} />
              </React.Fragment>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};
