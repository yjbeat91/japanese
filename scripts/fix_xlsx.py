import openpyxl, sys
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.comments import Comment
import os
HERE=os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0,HERE)
from fixes import F
exec(open(os.path.join(HERE,'fixes2.py'),encoding='utf-8').read())
src_path, out_path = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src_path)
YEL = PatternFill('solid', fgColor='FFF2A8')
RED = PatternFill('solid', fgColor='F8C9C9')
changes = []
order = [ws.title for ws in wb.worksheets]
for ws in wb.worksheets:
    key_sheet = ws.title
    hdr = next(r for r in range(1, 10) if ws.cell(r, 1).value == '번호')
    for r in range(hdr + 1, ws.max_row + 1):
        b = ws.cell(r, 2); c = ws.cell(r, 3); d = ws.cell(r, 4)
        if b.value in (None, ''): continue
        no = ws.cell(r, 1).value
        jp = str(b.value).strip(); kr = str(d.value or '').strip()
        src = f'{key_sheet} #{no}'
        if 'インデックス' in jp or '원문 색인' in kr or '색인' in jp:
            for cell in (b, c, d): cell.fill = RED
            b.comment = Comment('단어가 아닌 행입니다. 삭제를 권장합니다.', 'Claude')
            changes.append((ws.title, no, jp, '행 전체', '(색인 행)', '삭제 권장')); continue
        if not kr:
            d.fill = RED; d.comment = Comment('뜻이 비어 있습니다.', 'Claude')
            changes.append((ws.title, no, jp, '뜻', '(비어 있음)', '뜻 입력 필요')); continue
        hk = next((k for k in (src + '|' + jp, src) if k in F2 or k in F), None)
        new_f = None; new_jp = None; new_kr = None
        if hk:
            fx = (F2[hk] if hk in F2 else F[hk]) + (None, None)
            new_f, new_jp, new_kr = fx[0], fx[1], fx[2]
        old_f = '' if c.value is None else str(c.value)
        if new_f is None and ' ' in old_f.strip():
            new_f = old_f.replace(' ', '')
        def setc(cell, label, old, new):
            if new is None or new == old: return
            cell.value = new if new != '' else None
            cell.fill = YEL
            cell.comment = Comment(f'원래 값: {old if old else "(비어 있음)"}', 'Claude')
            changes.append((ws.title, no, new_jp or jp, label, old if old else '(비어 있음)', new if new else '(비움: 가나 단어)'))
        setc(b, '일본어 단어', jp, new_jp)
        target_jp = new_jp or jp
        if new_f == '' : new_f = None if not old_f or old_f == target_jp else ''
        setc(c, '후리가나', old_f, new_f)
        setc(d, '뜻', kr, new_kr)
# change list sheet
ls = wb.create_sheet('수정 목록', 0)
f_hdr = Font(name='Calibri', size=11, bold=True)
f_body = Font(name='Calibri', size=11)
thin = Side(style='thin', color='BFBFBF')
ls['A1'] = '후리가나·단어 수정 목록'; ls['A1'].font = Font(name='Calibri', size=14, bold=True)
ls['A2'] = ('노란 칸 = 고친 칸 (메모에 원래 값), 빨간 칸 = 확인이 필요한 행. '
            '새로 받은 "차,카" 시트(ㅊ·ㅋ 단어 252개)도 검토해 함께 고쳤습니다.')
ls['A2'].font = f_body; ls['A2'].alignment = Alignment(wrap_text=True, vertical='top')
ls.merge_cells('A2:F2'); ls.row_dimensions[2].height = 40
heads = ['시트', '번호', '일본어 단어', '고친 항목', '원래 값', '수정 값']
for i, h in enumerate(heads, 1):
    cell = ls.cell(4, i, h); cell.font = f_hdr; cell.fill = PatternFill('solid', fgColor='E6F0FA')
    cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
changes.sort(key=lambda x: (order.index(x[0]), x[1] if isinstance(x[1], int) else 99999))
for r, row in enumerate(changes, 5):
    for i, v in enumerate(row, 1):
        cell = ls.cell(r, i, v); cell.font = f_body
        cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
        if i == 2: cell.alignment = Alignment(horizontal='right')
for col, w in zip('ABCDEF', (10, 7, 26, 12, 34, 34)):
    ls.column_dimensions[col].width = w
ls.freeze_panes = 'A5'
ls.auto_filter.ref = f'A4:F{4 + len(changes)}'
wb.active = 0
wb.save(out_path)
from collections import Counter
print('changes', len(changes), Counter(c[3] for c in changes), Counter(c[0] for c in changes))
