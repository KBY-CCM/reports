"""출연계약서 docx/pdf → 인물정보 행 추출.  usage: python3 -I extract_contracts.py out.csv file1.docx file2.pdf ...
출력 CSV 컬럼: 이름,분류,필요번호,본명_예금주,은행,계좌번호,게스트여부,촬영일,출연료,정산명의,계약서파일,확인필요
"""
import sys, re, zipfile, html, csv, os, subprocess

MISSING = "모두싸인에서 확인"

def docx_text(p):
    z = zipfile.ZipFile(p); x = z.read('word/document.xml').decode('utf-8', 'ignore')
    t = re.sub(r'</w:p>', '\n', x); t = re.sub(r'<w:tab/>', ' ', t); t = re.sub(r'<[^>]+>', '', t)
    return html.unescape(t)

def pdf_text(p):
    try:
        return subprocess.run(['pdftotext', '-layout', p, '-'], capture_output=True, text=True, timeout=60).stdout
    except Exception:
        return ''

def clean(v):
    v = (v or '').strip()
    v = re.sub(r'^[\[\(\s_]+|[\]\)\s_]+$', '', v)   # [  값  ] / ( ) / 밑줄 제거
    v = v.replace('_', '').strip()
    return v

def grab(text, label_pat, value_pat=r'([^\n\]\)]{1,60})'):
    m = re.search(label_pat + r'\s*[:：]?\s*\[?\s*' + value_pat, text)
    return clean(m.group(1)) if m else ''

def extract(path):
    ext = os.path.splitext(path)[1].lower()
    t = docx_text(path) if ext == '.docx' else pdf_text(path)
    t1 = re.sub(r'[ \t]+', ' ', t)
    r = {k: '' for k in ['이름','분류','필요번호','본명_예금주','은행','계좌번호','게스트여부','촬영일','출연료','정산명의','본명','계약서파일','확인필요']}
    r['계약서파일'] = os.path.basename(path); r['게스트여부'] = '게스트'
    # 이름: 활동명 “X” / 본명 Y  >  과 X(이하 "출연자")  >  출연자: X (인)  >  파일명
    stage, real = '', ''
    m = re.search(r'활동명\s*[“"]([^”"\n]{1,30})[”"]\s*(?:/\s*(?:본명\s*)?([^\(\n/]{1,20}?))?\s*\(이하', t1)
    if m: stage, real = clean(m.group(1)), clean(m.group(2) or '')
    if not stage:
        m = re.search(r'과\s*([^\(\n]{1,30}?)\s*\(이하\s*["“]?출연자', t1)
        if m: stage = clean(m.group(1))
    if not stage:
        m = re.search(r'출연자\s*[:：]\s*([^\(\n]{1,30}?)\s*\(인\)', t1)
        if m: stage = clean(m.group(1))
    if not stage:
        m = re.search(r'계약서[_ ]+([가-힣A-Za-z&,\s]{2,20}?)(?:_|\.docx|\.pdf)', os.path.basename(path))
        if m: stage = clean(m.group(1))
    stage = re.sub(r'님$', '', stage).replace('&', ', ')
    r['이름'] = stage; r['본명'] = real
    # 번호
    rrn = re.search(r'(\d{6})\s*-\s*(\d{7})', t1)
    brn = re.findall(r'(\d{3}-\d{2}-\d{5})', t1)
    brn = [b for b in brn if b != '749-87-01653']  # 금병영 자신 제외
    if rrn:
        r['필요번호'] = rrn.group(1) + '-' + rrn.group(2); r['분류'] = '삼쩜삼'; r['정산명의'] = '개인'
    elif brn:
        r['필요번호'] = brn[0]; r['분류'] = '세금계산서'; r['정산명의'] = '사업자'
    else:
        r['필요번호'] = MISSING; r['분류'] = MISSING
    # 계좌
    r['본명_예금주'] = grab(t1, r'예금주') or MISSING
    r['은행'] = grab(t1, r'은행\s*명?') or MISSING
    acct = grab(t1, r'계좌\s*번호', r'([\d\-\s]{8,30})')
    r['계좌번호'] = re.sub(r'\s', '', acct) if acct and re.search(r'\d{6,}', acct.replace('-', '')) else MISSING
    # 사업자면 예금주 = 사업자명
    if r['정산명의'] == '사업자' and r['본명_예금주'] == MISSING:
        c = re.search(r'소속사\s*[:：]\s*\[?([^\(\n\]]{2,40}?)\]?\s*\(인\)', t1)
        if c: r['본명_예금주'] = clean(c.group(1))
    # 촬영일 / 출연료
    d = re.search(r'촬영\s*일시\s*[:：]\s*(\d{4})[.년\s]*(\d{1,2})[.월\s]*(\d{1,2})', t1)
    d8 = re.search(r'촬영\s*일시\s*[:：]\s*(\d{4})(\d{2})(\d{2})', t1)
    if d8: d = d8
    r['촬영일'] = f"{d.group(1)}-{int(d.group(2)):02d}-{int(d.group(3)):02d}" if d else ''
    f = re.search(r'총\s*출연\s*대금\s*[:：]?\s*₩?\s*([\d,]+)', t1) or re.search(r'출연료\s*[:：]?\s*촬영당\s*₩?\s*([\d,]+)', t1)
    r['출연료'] = f.group(1).replace(',', '') if f else ''
    need = [k for k in ['분류','필요번호','본명_예금주','은행','계좌번호'] if r[k] == MISSING]
    if r['출연료'] == '0': need = []  # 무상 출연은 계좌 불필요
    r['확인필요'] = ', '.join(need)
    return r

if __name__ == '__main__':
    out, files = sys.argv[1], sys.argv[2:]
    rows = [extract(f) for f in files]
    with open(out, 'w', newline='', encoding='utf-8-sig') as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)
    for r in rows: print({k: v for k, v in r.items() if k not in ('필요번호','계좌번호')})
