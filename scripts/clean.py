import openpyxl, re, json, sys
import os
HERE=os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0,HERE)
from fixes import F
exec(open(os.path.join(HERE,'fixes2.py'),encoding='utf-8').read())
from collections import Counter as _C
hits=_C()
used=set()
from collections import Counter
wb=openpyxl.load_workbook(sys.argv[1],read_only=True)
HIRA=r'぀-ゟ'; KANA=r'぀-ヿー'
KAN=r'一-鿿々〆ヶ'
iskana=lambda s: bool(re.fullmatch('['+KANA+r'・\s]+',s))
hasko=lambda s: bool(re.search('[가-힣]',s))
hasjp=lambda s: bool(re.search('['+KANA+KAN+']',s))
def pat(jp):
    s=''
    for ch in jp:
        if re.match('['+KAN+']',ch):
            if not s.endswith('.+'): s+='.+'
        elif re.match('['+KANA+']',ch): s+=re.escape(ch)
        else: s+='.*'
    return '^'+s+'$'
def k2h(s): return ''.join(chr(ord(c)-0x60) if '\u30A1'<=c<='\u30F6' else c for c in s)
def furi_ok(jp,f):
    jp,f=k2h(jp),k2h(f)
    if not f or not iskana(f): return False
    parts=jp.split('・')
    return any(re.match(pat(p.replace(' ','')),f.replace(' ','')) for p in parts)
def clean_furi(jp,f):
    f=f.strip()
    if '->' in f: f=f.split('->')[-1].strip()
    m=re.search(r'[\(（]([^)）]+)[\)）]',f)
    cands=[]
    if m: cands.append(m.group(1).strip())
    cands+= [c.strip(' ?') for c in re.split(r'[/／]',re.sub(r'[\(（][^)）]*[\)）]','',f))]
    for c in cands:
        if furi_ok(jp,c): return c,0
    for c in cands:
        if iskana(c): return c,1
    return cands[0] if cands else f,1
def split_ex(e):
    e=(e or '').strip()
    if not e: return '',''
    m=re.match(r'^(.*?)\s*[\(（]([^()（）]*)[\)）]\s*[。.]?\s*$',e)
    if m:
        a,b=m.group(1).strip(),m.group(2).strip()
        if hasjp(a) and not hasko(a) and hasko(b): return a,b
        if hasko(a) and hasjp(b) and not hasko(b): return b,a
    if hasjp(e) and not hasko(e): return e,''
    # mixed: try to separate
    return e,''
def find_blank(jp,ex):
    if not ex: return -1,0
    forms=[p.strip() for p in jp.split('・') if p.strip()]
    for p in forms:
        i=ex.find(p)
        if i>=0: return i,len(p)
    for p in forms:
        if len(p)>=2 and re.match('['+HIRA+']',p[-1]):
            stem=p[:-1]
            # い-adj / verb: strip trailing okurigana progressively but keep a kanji
            while stem and re.match('['+HIRA+']',stem[-1]) and len(stem)>2: stem=stem[:-1]
            if not re.search('['+KAN+']',stem) and len(stem)<2: continue
            i=ex.find(stem)
            if i>=0:
                j=i+len(stem); k=0
                while j<len(ex) and k<4 and re.match('['+HIRA+']',ex[j]): j+=1;k+=1
                return i,j-i
    return -1,0
CHO='ㄱㄲㄴㄷㄸㄹㅁㅂㅃㅅㅆㅇㅈㅉㅊㅋㅌㅍㅎ'
MERGE={'ㄲ':'ㄱ','ㄸ':'ㄷ','ㅃ':'ㅂ','ㅆ':'ㅅ','ㅉ':'ㅈ'}
def cho(s):
    for ch in s:
        if '가'<=ch<='힣':
            c=CHO[(ord(ch)-0xAC00)//588]; return MERGE.get(c,c)
        if ch.strip() and ch not in '(（[「"\'': break
    return '기타'
out=[];seen=set();stats=Counter()
new_tail=[]
sheets=[w for w in wb.worksheets if w.title!='차,카']+[w for w in wb.worksheets if w.title=='차,카']
for ws in sheets:
    rows=list(ws.iter_rows(values_only=True))
    h=next(i for i,r in enumerate(rows) if r and r[0]=='번호')
    for r in rows[h+1:]:
        if not r or not r[1]: continue
        jp=str(r[1]).strip(); kr=str(r[3] or '').strip(); f=str(r[2] or '').strip()
        if '색인' in jp and '원문 색인' not in kr:
            out.append(['','','','','',-1,0,'기타',0,f'{ws.title} #{r[0]}']); stats['placeholder']+=1; continue
        if 'インデックス' in jp or '원문 색인' in kr: stats['junk']+=1; continue
        if not kr: stats['nokr']+=1; continue
        key=(jp,kr)
        if key in seen: stats['dup']+=1; continue
        seen.add(key)
        src=f'{ws.title} #{r[0]}'
        k1=src+'|'+jp
        hk=next((k for k in (k1,src) if k in F2 or k in F),None)
        if hk:
            fx=(F2[hk] if hk in F2 else F[hk])+(None,None); used.add(hk); hits[hk]+=1
            if fx[1]: jp=fx[1]
            if fx[2]: kr=fx[2]
            furi,flag=fx[0],0
        else:
            furi,flag=clean_furi(jp,f)
        exJ,exK=split_ex(str(r[4] or ''))
        bs,bl=find_blank(jp,exJ)
        furi=furi.replace(' ','').replace('\u3000','')
        if furi==jp: furi=''
        out.append([jp,furi,kr,exJ,exK,bs,bl,cho(kr),flag,src])
        stats['flag']+=flag; stats['blank']+= bs>=0
gc=Counter(w[7] for w in out)
for w in out:
    if gc[w[7]]<20: w[7]='기타'
print(len(out),dict(stats))
print('unused fixes',(set(F)|set(F2))-used)
print('multi-hit',[k for k,v in hits.items() if v>1])
print('applied',len(hits))
print(Counter(w[7] for w in out))
with open(sys.argv[2],'w') as fh:
    fh.write('[\n'+',\n'.join(json.dumps(w,ensure_ascii=False,separators=(',',':')) for w in out)+'\n]\n')
