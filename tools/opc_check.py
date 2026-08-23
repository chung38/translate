"""PPTX 封裝層完整性檢查。

XSD 只驗單一 XML 的結構，驗不到「部件之間的關聯」——PowerPoint 跳修復對話框
多半是這一層出問題，LibreOffice 卻通常照開不誤。
"""
import zipfile, re, posixpath, sys
from collections import defaultdict

REL = re.compile(r'<Relationship\b[^>]*/>')
ATTR = lambda tag, name: (re.search(rf'{name}="([^"]*)"', tag) or [None, None])[1]


def check(path):
    z = zipfile.ZipFile(path)
    names = set(n for n in z.namelist() if not n.endswith('/'))
    ct = z.read('[Content_Types].xml').decode()
    defaults = dict(re.findall(r'<Default Extension="([^"]+)" ContentType="([^"]+)"/>', ct))
    problems = []

    def rels_of(part):
        rp = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
        if rp not in names:
            return []
        base = posixpath.dirname(part)
        out = []
        for tag in REL.findall(z.read(rp).decode()):
            t, mode = ATTR(tag, 'Target'), ATTR(tag, 'TargetMode')
            typ = (ATTR(tag, 'Type') or '').rsplit('/', 1)[-1]
            if mode == 'External':
                out.append((typ, None, rp))
            else:
                out.append((typ, posixpath.normpath(posixpath.join(base, t)), rp))
        return out

    # 1. 每個部件都要有 content type
    for n in sorted(names):
        if n == '[Content_Types].xml' or n.endswith('.rels'):
            continue
        if f'PartName="/{n}"' in ct or n.rsplit('.', 1)[-1].lower() in defaults:
            continue
        problems.append(f'[content-type] {n} 沒有登記')

    # 2. 關聯不能指向不存在的檔案
    for n in sorted(names):
        for typ, tgt, rp in rels_of(n):
            if tgt and tgt not in names:
                problems.append(f'[dangling] {rp} 指向不存在的 {tgt}')

    # 3. 備忘稿必須 1:1，而且回指要對得上
    slide_notes = defaultdict(list)
    for n in sorted(names):
        if not re.match(r'ppt/slides/slide\d+\.xml$', n):
            continue
        for typ, tgt, _ in rels_of(n):
            if typ == 'notesSlide' and tgt:
                slide_notes[tgt].append(n)
    for note, slides in slide_notes.items():
        if len(slides) > 1:
            problems.append(f'[notes-1:1] {note} 被多張投影片共用: {slides}')
        back = [t for ty, t, _ in rels_of(note) if ty == 'slide']
        if back and back[0] not in slides:
            problems.append(f'[notes-backref] {note} 回指 {back[0]}，但引用它的是 {slides}')

    # 4. sldIdLst 的每個 r:id 都要解析得到，且 id 唯一、在合法範圍
    pres = z.read('ppt/presentation.xml').decode()
    pres_rels = {ATTR(t, 'Id'): ATTR(t, 'Target')
                 for t in REL.findall(z.read('ppt/_rels/presentation.xml.rels').decode())}
    seen = set()
    for sid, rid in re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"/>', pres):
        if rid not in pres_rels:
            problems.append(f'[sldIdLst] {rid} 在 presentation.xml.rels 找不到')
        if sid in seen:
            problems.append(f'[sldIdLst] sldId {sid} 重複')
        seen.add(sid)
        if not (256 <= int(sid) <= 2147483647):
            problems.append(f'[sldIdLst] sldId {sid} 超出合法範圍')

    # 5. rels 檔裡的 Id 不能重複
    for n in sorted(names):
        if not n.endswith('.rels'):
            continue
        ids = [ATTR(t, 'Id') for t in REL.findall(z.read(n).decode())]
        if len(ids) != len(set(ids)):
            problems.append(f'[dup-rid] {n} 有重複的 Relationship Id')

    # 6. zip 不能有資料夾條目（OPC 每個項目都必須是一個 part）
    for n in z.namelist():
        if n.endswith('/'):
            problems.append(f'[zip-folder] {n} 是資料夾條目，OPC 封裝不允許')

    # 7. docProps/app.xml 的張數統計要跟實際 part 數一致
    if 'docProps/app.xml' in names:
        app = z.read('docProps/app.xml').decode()
        real_slides = len([n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)])
        real_notes = len([n for n in names if re.match(r'ppt/notesSlides/notesSlide\d+\.xml$', n)])
        m = re.search(r'<Slides>(\d+)</Slides>', app)
        if m and int(m.group(1)) != real_slides:
            problems.append(f'[app.xml] <Slides> 寫 {m.group(1)}，實際有 {real_slides} 張')
        m = re.search(r'<Notes>(\d+)</Notes>', app)
        if m and int(m.group(1)) != real_notes:
            problems.append(f'[app.xml] <Notes> 寫 {m.group(1)}，實際有 {real_notes} 份')
        m = re.search(r'<vt:vector size="(\d+)" baseType="lpstr">', app)
        pairs = re.findall(r'<vt:lpstr>投影片標題</vt:lpstr></vt:variant><vt:variant><vt:i4>(\d+)</vt:i4>', app)
        if pairs and int(pairs[0]) != real_slides:
            problems.append(f'[app.xml] TitlesOfParts 宣告 {pairs[0]} 個標題，實際有 {real_slides} 張')

    # 8. SmartArt 的 diagram 部件不該被兩張投影片共用（內容會互相蓋掉）
    diag = defaultdict(list)
    for n in sorted(names):
        if not re.match(r'ppt/slides/slide\d+\.xml$', n):
            continue
        for typ, tgt, _ in rels_of(n):
            if typ.startswith('diagram') and tgt:
                diag[tgt].append(n)
    for part, slides in diag.items():
        if len(slides) > 1:
            problems.append(f'[diagram-shared] {part} 被多張投影片共用: {slides}')

    return problems


if __name__ == '__main__':
    for f in sys.argv[1:]:
        p = check(f)
        print(f'=== {f} ===')
        if not p:
            print('  通過，沒有發現封裝層問題')
        for x in p[:20]:
            print('  ' + x)
        if len(p) > 20:
            print(f'  ...還有 {len(p)-20} 項')
        print()
