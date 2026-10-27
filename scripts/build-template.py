#!/usr/bin/env python3
"""由原始 Word 清冊產生匯出範本。

用法：python3 scripts/build-template.py <原始清冊.docx>

輸出：
  template/template.docx  去除資料列、置入 {{TITLE}} {{ROWS}} {{NOTE}} {{Y1}} {{Y2}} 佔位字的範本
  template/row.xml        單一資料列 XML（{{C0}}~{{C13}} 為 14 個欄位）
  template/template.js    供網頁載入的內嵌範本（base64），請勿手動修改

所有文字一律設為：中文 標楷體、英文數字 Times New Roman。
"""
import re, sys, zipfile, base64, json, os

FONT = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'template')


def norm(x):
    x = re.sub(r'<w:rFonts\b[^>]*/>', '', x)
    x = re.sub(r'<w:rPr>(<w:rStyle [^>]*/>)?', lambda m: '<w:rPr>' + (m.group(1) or '') + FONT, x)
    x = re.sub(r'<w:r( [^>]*)?>(?!<w:rPr>)', lambda m: '<w:r%s><w:rPr>%s</w:rPr>' % (m.group(1) or '', FONT), x)
    return x


def main(src):
    z = zipfile.ZipFile(src)
    files = {n: z.read(n) for n in z.namelist()}
    d = files['word/document.xml'].decode()

    # 標題段落
    k = d.find('<w:t>新增及調整土')
    if k < 0:
        sys.exit('找不到標題「新增及調整土石流潛勢溪流…」段落')
    ts = d.rfind('<w:p ', 0, k); te = d.find('</w:p>', k) + 6
    tpPr = re.search(r'<w:pPr>.*?</w:pPr>', d[ts:te]).group(0)
    title_new = '<w:p>' + tpPr + '<w:r><w:rPr></w:rPr><w:t xml:space="preserve">{{TITLE}}</w:t></w:r></w:p>'

    # 表格：保留 2 列表頭，取第 3 列為資料列範本
    tbl_s = d.find('<w:tbl>'); tbl_e = d.find('</w:tbl>') + 8
    tbl = d[tbl_s:tbl_e]
    rows = [m.group(0) for m in re.finditer(r'<w:tr[ >].*?</w:tr>', tbl, re.S)]
    if len(rows) < 3:
        sys.exit('表格至少需 2 列表頭 + 1 列資料')
    pre = tbl[:tbl.find('<w:tr')]
    h1 = rows[0].replace('<w:t>114</w:t>', '<w:t>{{Y1}}</w:t>')
    h2 = rows[1].replace('<w:t>114</w:t>', '<w:t>{{Y1}}</w:t>').replace('<w:t>113</w:t>', '<w:t>{{Y2}}</w:t>')
    counter = [0]

    def cellrep(m):
        c = m.group(0); i = counter[0]; counter[0] += 1
        return re.sub(r'(</w:pPr>).*?(</w:p>)',
                      r'\1<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">{{C%d}}</w:t></w:r>\2' % i, c, flags=re.S)
    row_tpl = re.sub(r'<w:tc>.*?</w:tc>', cellrep, rows[2], flags=re.S)
    if counter[0] != 14:
        sys.exit('資料列應有 14 欄，實際 %d 欄' % counter[0])

    # 表下註記
    tail = d[tbl_e:]
    ne = tail.find('</w:p>') + 6
    npPr = re.search(r'<w:pPr>.*?</w:pPr>', tail[:ne]).group(0)
    note_new = '<w:p>' + npPr + '<w:r><w:rPr><w:b w:val="0"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">{{NOTE}}</w:t></w:r></w:p>'

    doc = d[:ts] + title_new + d[te:tbl_s] + pre + h1 + h2 + '{{ROWS}}</w:tbl>' + note_new + tail[ne:]
    files['word/document.xml'] = norm(doc).encode()
    row_tpl = norm(row_tpl)
    for n in ('word/header1.xml', 'word/footer1.xml'):
        if n in files:
            files[n] = norm(files[n].decode()).encode()
    s = files['word/styles.xml'].decode()
    s = re.sub(r'(<w:rPrDefault><w:rPr>)<w:rFonts\b[^>]*/>', r'\1' + FONT, s, count=1)
    files['word/styles.xml'] = s.encode()

    # 公開發佈前清除文件中的個人資訊（作者、最後修改者、公司）
    if 'docProps/core.xml' in files:
        c = files['docProps/core.xml'].decode()
        c = re.sub(r'<dc:creator>.*?</dc:creator>', '<dc:creator></dc:creator>', c)
        c = re.sub(r'<cp:lastModifiedBy>.*?</cp:lastModifiedBy>', '<cp:lastModifiedBy></cp:lastModifiedBy>', c)
        files['docProps/core.xml'] = c.encode()
    if 'docProps/app.xml' in files:
        a = files['docProps/app.xml'].decode()
        a = re.sub(r'<Company>.*?</Company>', '<Company></Company>', a)
        files['docProps/app.xml'] = a.encode()

    os.makedirs(OUT, exist_ok=True)
    out_docx = os.path.join(OUT, 'template.docx')
    with zipfile.ZipFile(out_docx, 'w', zipfile.ZIP_DEFLATED) as o:
        for n in z.namelist():
            o.writestr(n, files[n])
    open(os.path.join(OUT, 'row.xml'), 'w', encoding='utf-8').write(row_tpl)
    embed()


def embed():
    """由 template/template.docx 與 template/row.xml 重新產生 template/template.js"""
    b64 = base64.b64encode(open(os.path.join(OUT, 'template.docx'), 'rb').read()).decode()
    row = open(os.path.join(OUT, 'row.xml'), encoding='utf-8').read()
    js = ('/* 自動產生：python3 scripts/build-template.py，請勿手動修改 */\n'
          'window.DOCX_TEMPLATE = ' + json.dumps({'b64': b64, 'row': row}, ensure_ascii=False) + ';\n')
    open(os.path.join(OUT, 'template.js'), 'w', encoding='utf-8').write(js)
    print('已更新 template/template.docx、row.xml、template.js')


if __name__ == '__main__':
    if len(sys.argv) == 2 and sys.argv[1] == '--embed-only':
        embed()
    elif len(sys.argv) == 2:
        main(sys.argv[1])
    else:
        sys.exit(__doc__)
