const fs = require('fs');
function adjustXmlRPrForLanguage(rPr, lang, docType) {
  const needsStrongFontAdjustment = ['vi', 'th'].includes(lang);
  let newRPr = rPr || '';
  let shouldAdjust = needsStrongFontAdjustment;
  if (shouldAdjust) {
    const langCode = lang === 'vi' ? 'vi-VN' : (lang === 'th' ? 'th-TH' : 'en-US');
    if (docType === 'docx') {
      const arialFonts = `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>`;
      const langTag = `<w:lang w:val="${langCode}" w:eastAsia="${langCode}" w:bidi="${langCode}"/>`;
      if (!newRPr) {
        newRPr = `<w:rPr>${arialFonts}${langTag}</w:rPr>`;
      } else {
        newRPr = newRPr.replace(/<w:rFonts[\s\S]*?(?:\/>|<\/w:rFonts>)/g, '');
        newRPr = newRPr.replace(/<w:lang[\s\S]*?(?:\/>|<\/w:lang>)/g, '');
        newRPr = newRPr.replace(/<w:hint[\s\S]*?(?:\/>|<\/w:hint>)/g, '');
        if (newRPr.match(/<w:rPr[^>]*\/>/)) {
          newRPr = newRPr.replace(/(<w:rPr[^>]*)\/>/, `$1></w:rPr>`);
        }
        if (newRPr.includes('<w:rStyle')) {
          newRPr = newRPr.replace(/(<w:rStyle[^>]*(?:\/>|<\/w:rStyle>))/, `$1${arialFonts}`);
        } else if (newRPr.includes('<w:rPr>')) {
          newRPr = newRPr.replace('<w:rPr>', `<w:rPr>${arialFonts}`);
        } else if (newRPr.includes('<w:rPr ')) {
          newRPr = newRPr.replace(/(<w:rPr[^>]*>)/, `$1${arialFonts}`);
        } else {
          newRPr = `<w:rPr>${arialFonts}</w:rPr>`;
        }
        if (newRPr.includes('</w:rPr>')) {
          newRPr = newRPr.replace('</w:rPr>', `${langTag}</w:rPr>`);
        } else {
          newRPr = newRPr + langTag;
        }
      }
    }
  }
  return newRPr;
}
console.log(adjustXmlRPrForLanguage('<w:rPr><w:rFonts w:hint="eastAsia"/><w:color w:val="FF0000"/><w:sz w:val="24"/></w:rPr>', 'vi', 'docx'));
console.log(adjustXmlRPrForLanguage('<w:rPr><w:color w:val="000000"/></w:rPr>', 'vi', 'docx'));
