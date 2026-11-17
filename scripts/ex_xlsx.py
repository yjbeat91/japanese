# 예문 수정 목록 엑셀: fixes_ex.py의 EX를 원본 예문과 나란히 정리한다.
# python3 scripts/ex_xlsx.py "data/일본어 단어장_원본.xlsx" "out/일본어 예문_수정본.xlsx"
import openpyxl, re, sys, os
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from fixes_ex import EX, EX2
KO = re.compile('[가-힣]'); JA = re.compile('[぀-ヿ一-鿿]')
src_wb = openpyxl.load_workbook(sys.argv[1], read_only=True)
rows = []; seen = set()
order = [ws.title for ws in src_wb.worksheets]
for ws in src_wb.worksheets:
    data = list(ws.iter_rows(values_only=True))
    h = next(i for i, r in enumerate(data) if r and r[0] == '번호')
    for r in data[h + 1:]:
        if not r or not r[1]: continue
        jp = str(r[1]).strip(); src = f'{ws.title} #{r[0]}'
        key = next((k for k in (src + '|' + jp, src) if k in EX), None)
        if not key or key in seen: continue
        seen.add(key)
        ej, ek = EX[key]; old = str(r[4] or '').strip()
        kind = '해석에 일본어 섞임' if key in EX2 else '예문 순서·한글 섞임'
        rows.append((ws.title, r[0], jp, str(r[3] or '').strip(), old, ej, ek, kind))
missing = set(EX) - seen
assert not missing, missing
rows.sort(key=lambda x: (order.index(x[0]), x[1] if isinstance(x[1], int) else 99999))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '예문 수정 목록'
F = lambda **k: Font(name='Arial', size=11, **k)
thin = Side(style='thin', color='BFBFBF'); box = Border(top=thin, bottom=thin, left=thin, right=thin)
ws['A1'] = '일본어 예문 수정 목록'; ws['A1'].font = Font(name='Arial', size=14, bold=True)
n1 = sum(r[7] == '예문 순서·한글 섞임' for r in rows); n2 = len(rows) - n1
ws['A2'] = (f'원본 엑셀의 예문 칸을 고친 {len(rows)}줄입니다. 예문 순서·한글 섞임 {n1}줄: "한국어 (일본어)" 순서이거나 일본어 문장에 한글이 섞인 예문. '
            f'해석에 일본어 섞임 {n2}줄: 한국어 해석 안에 일본어 단어가 들어간 예문(일본어 예문은 그대로, 해석만 한국어로). 원본 엑셀은 그대로 두었습니다.')
ws['A2'].font = F(); ws['A2'].alignment = Alignment(wrap_text=True, vertical='top')
ws.merge_cells('A2:H2'); ws.row_dimensions[2].height = 48
heads = ['시트', '번호', '일본어 단어', '뜻', '원래 예문 (원본 그대로)', '고친 일본어 예문', '고친 한국어 해석', '고친 이유']
for i, h in enumerate(heads, 1):
    c = ws.cell(4, i, h); c.font = F(bold=True); c.fill = PatternFill('solid', fgColor='E2F2FC'); c.border = box
    c.alignment = Alignment(vertical='center', wrap_text=True)
YEL = PatternFill('solid', fgColor='FFF2A8')
for ri, row in enumerate(rows, 5):
    for ci, v in enumerate(row, 1):
        c = ws.cell(ri, ci, v); c.font = F(); c.border = box
        c.alignment = Alignment(wrap_text=True, vertical='top', horizontal='right' if ci == 2 else None)
        if ci in (6, 7): c.fill = YEL
for col, w in zip('ABCDEFGH', (9, 7, 18, 22, 46, 40, 36, 18)): ws.column_dimensions[col].width = w
ws.freeze_panes = 'A5'; ws.auto_filter.ref = f'A4:H{4 + len(rows)}'
wb.save(sys.argv[2])
print('rows', len(rows), 'order/mixed', n1, 'ko-with-ja', n2)
