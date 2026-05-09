from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart, BarChart, LineChart, Reference
from datetime import datetime
import io
import os

app = Flask(__name__)
CORS(app)

# ===== Style helpers =====
def fill(c): return PatternFill('solid', start_color=c)
def border(): return Border(bottom=Side(style='thin', color='2D3250'), right=Side(style='thin', color='2D3250'))
def ca(h='center', v='center', wrap=False): return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
STATUS_COLORS = {'結案':'34D399','轉派技師':'FBBF24','轉派工程師':'A78BFA','客服處理中':'5B8CFF','待客戶寄回':'F87171'}
BC = ['5B8CFF','7C6CFF','34D399','FBBF24','F87171','FB923C','A78BFA','38BDF8','F472B6']

def set_hdr(ws, row, cols):
    for c, val in enumerate(cols, 1):
        cell = ws.cell(row=row, column=c, value=val)
        cell.font = Font(name='Arial', bold=True, color='94A3B8', size=10)
        cell.fill = fill('2D3250')
        cell.alignment = ca()
        cell.border = border()

def title_row(ws, row, text, ncols, bg='5B8CFF'):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name='Arial', bold=True, color='FFFFFF', size=12)
    c.fill = fill(bg)
    c.alignment = ca('left')
    ws.row_dimensions[row].height = 22

def calc_dur(s, e):
    if not s or not e: return ''
    try:
        ms = (datetime.fromisoformat(e) - datetime.fromisoformat(s)).total_seconds()
        if ms <= 0: return ''
        h = int(ms // 3600)
        m = int((ms % 3600) // 60)
        if h >= 24:
            d = h // 24; rh = h % 24
            return f'{d}天{rh}小時' if rh else f'{d}天'
        return f'{h}小時{m}分' if h else f'{m}分鐘'
    except: return ''

def is_dispatch_overdue(r):
    if r.get('status') == '結案': return False
    if not r.get('dispatchDate'): return False
    if r.get('handler') == '客戶': return False
    try: return (datetime.now() - datetime.fromisoformat(r['dispatchDate'])).days > 7
    except: return False

def map_product(p):
    if not p: return '其他'
    if 'DMVR' in p: return 'FMS-DMVR'
    if 'GPS' in p: return 'FMS-GPS'
    if '冷鏈' in p: return 'FMS-冷鏈'
    if '雷達' in p: return 'FMS-雷達'
    return '其他'

# ===== WEEKLY REPORT =====
def generate_weekly(records, from_date, to_date):
    total = len(records)
    closed = sum(1 for r in records if r.get('status') == '結案')
    open_cnt = total - closed
    close_rate = f'{closed/total*100:.1f}%' if total else '0%'
    od = [r for r in records if is_dispatch_overdue(r)]
    durs = []
    for r in records:
        if r.get('status') == '結案' and r.get('date') and r.get('closeDate'):
            try:
                sec = (datetime.fromisoformat(r['closeDate']) - datetime.fromisoformat(r['date'])).total_seconds()
                if sec > 0: durs.append(sec)
            except: pass
    avg_h = round(sum(durs)/len(durs)/3600) if durs else 0
    avg_str = f'{avg_h//24}天{avg_h%24}小時' if avg_h >= 24 else f'{avg_h}小時'
    label = f'{from_date} ～ {to_date}'
    wb = Workbook()

    # ===== 封面 =====
    ws0 = wb.active
    ws0.title = '封面摘要'
    ws0.sheet_view.showGridLines = False
    for col, w in [('A',3),('B',22),('C',18),('D',18),('E',18),('F',18),('G',3)]:
        ws0.column_dimensions[col].width = w

    ws0.merge_cells('B2:F2')
    c = ws0.cell(row=2, column=2, value='⚡ 售服案件週報')
    c.font = Font(name='Arial', bold=True, color='FFFFFF', size=18)
    c.fill = fill('1A1D27'); c.alignment = ca('left')
    ws0.row_dimensions[2].height = 36

    ws0.merge_cells('B3:F3')
    c = ws0.cell(row=3, column=2, value=f'報告期間：{label}')
    c.font = Font(name='Arial', color='94A3B8', size=11)
    c.fill = fill('1A1D27'); c.alignment = ca('left')
    ws0.row_dimensions[3].height = 22
    ws0.row_dimensions[4].height = 12

    kpis = [('進線總件數',total,'5B8CFF','件'),('已結案',closed,'34D399',f'結案率 {close_rate}'),('未結案',open_cnt,'FBBF24','件'),('派工逾7天未結案',len(od),'F87171','件')]
    for i, (lbl, val, color, sub) in enumerate(kpis):
        col = 2 + i
        c = ws0.cell(row=5, column=col, value=lbl)
        c.font = Font(name='Arial', color='94A3B8', size=9); c.fill = fill('22263A'); c.alignment = ca()
        ws0.row_dimensions[5].height = 14
        ws0.merge_cells(start_row=6, start_column=col, end_row=7, end_column=col)
        c = ws0.cell(row=6, column=col, value=val)
        c.font = Font(name='Arial', bold=True, color=color, size=26); c.fill = fill('22263A'); c.alignment = ca()
        ws0.row_dimensions[6].height = 28; ws0.row_dimensions[7].height = 10
        c = ws0.cell(row=8, column=col, value=sub)
        c.font = Font(name='Arial', color='64748B', size=9); c.fill = fill('22263A'); c.alignment = ca()
        ws0.row_dimensions[8].height = 14

    ws0.row_dimensions[9].height = 12
    ws0.merge_cells('B10:F10')
    c = ws0.cell(row=10, column=2, value='👤 本週人員負責件數')
    c.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    c.fill = fill('2D3250'); c.alignment = ca('left')
    ws0.row_dimensions[10].height = 18

    handler_c = {}
    for r in records:
        if r.get('handler'): handler_c[r['handler']] = handler_c.get(r['handler'], 0) + 1
    row = 11
    for h, cnt in sorted(handler_c.items(), key=lambda x: -x[1]):
        bg = '1E2235' if row%2==0 else '161925'
        for col, val in [(2, h),(3, f'{cnt} 件')]:
            c = ws0.cell(row=row, column=col, value=val)
            c.font = Font(name='Arial', size=10, color='E2E8F0')
            c.fill = fill(bg); c.alignment = ca('left') if col==2 else ca()
            c.border = border()
        ws0.row_dimensions[row].height = 16; row += 1

    # ===== ① 進線管道 =====
    ws1 = wb.create_sheet('① 進線管道分析')
    ws1.sheet_view.showGridLines = False
    title_row(ws1, 1, f'📡 進線管道分析　｜　{label}', 3)
    set_hdr(ws1, 2, ['進線管道','件數','佔比'])
    ws1.column_dimensions['A'].width = 18; ws1.column_dimensions['B'].width = 10; ws1.column_dimensions['C'].width = 12

    ch_c = {}
    for r in records: ch_c[r.get('channel') or '未知'] = ch_c.get(r.get('channel') or '未知', 0) + 1
    row = 3
    for ch, cnt in sorted(ch_c.items(), key=lambda x: -x[1]):
        bg = '1E2235' if row%2==0 else '161925'
        for c2, (val, color, bold) in enumerate([(ch,'E2E8F0',False),(cnt,'5B8CFF',True),(f'{cnt/total*100:.1f}%' if total else '0%','94A3B8',False)], 1):
            c = ws1.cell(row=row, column=c2, value=val)
            c.font = Font(name='Arial', bold=bold, color=color, size=10)
            c.fill = fill(bg); c.alignment = ca(); c.border = border()
        row += 1
    for c2, val in enumerate(['合計',total,'100%'], 1):
        c = ws1.cell(row=row, column=c2, value=val)
        c.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
        c.fill = fill('2D3250'); c.alignment = ca(); c.border = border()
    pie = PieChart(); pie.title='進線管道佔比'; pie.style=10; pie.width=14; pie.height=10
    labels = Reference(ws1, min_col=1, min_row=3, max_row=2+len(ch_c))
    data = Reference(ws1, min_col=2, min_row=2, max_row=2+len(ch_c))
    pie.add_data(data, titles_from_data=True); pie.set_categories(labels)
    ws1.add_chart(pie, 'E2')

    # ===== ② 問題分類 =====
    ws2 = wb.create_sheet('② 問題分類統計')
    ws2.sheet_view.showGridLines = False
    title_row(ws2, 1, f'🏷️ 問題分類統計　｜　{label}', 3)
    set_hdr(ws2, 2, ['問題大類','總件數','主要次分類（件數）'])
    ws2.column_dimensions['A'].width = 14; ws2.column_dimensions['B'].width = 10; ws2.column_dimensions['C'].width = 50

    cat_map = {}
    for r in records:
        k = r.get('category') or '其他'
        if k not in cat_map: cat_map[k] = {'total':0,'subs':{}}
        cat_map[k]['total'] += 1
        s = r.get('subcategory') or '其他'
        cat_map[k]['subs'][s] = cat_map[k]['subs'].get(s,0)+1
    sorted_cats = sorted(cat_map.items(), key=lambda x: -x[1]['total'])
    row = 3
    for cat, v in sorted_cats:
        top = '、'.join(f"{s}({n})" for s,n in sorted(v['subs'].items(), key=lambda x:-x[1])[:4])
        bg = '1E2235' if row%2==0 else '161925'
        for c2,(val,color,bold) in enumerate([(cat,'E2E8F0',False),(v['total'],'5B8CFF',True),(top,'94A3B8',False)], 1):
            c = ws2.cell(row=row, column=c2, value=val)
            c.font = Font(name='Arial', bold=bold, color=color, size=10)
            c.fill = fill(bg); c.alignment = ca() if c2<=2 else ca('left',wrap=True); c.border = border()
        ws2.row_dimensions[row].height = 16; row += 1
    bar = BarChart(); bar.type='bar'; bar.title='問題大類件數'; bar.style=10; bar.width=16; bar.height=12
    cats_r = Reference(ws2, min_col=1, min_row=3, max_row=2+len(sorted_cats))
    data_r = Reference(ws2, min_col=2, min_row=2, max_row=2+len(sorted_cats))
    bar.add_data(data_r, titles_from_data=True); bar.set_categories(cats_r)
    ws2.add_chart(bar, 'E2')

    # ===== ③ 處理狀態 =====
    ws3 = wb.create_sheet('③ 處理狀態總覽')
    ws3.sheet_view.showGridLines = False
    title_row(ws3, 1, f'📋 處理狀態總覽　｜　{label}', 4)
    set_hdr(ws3, 2, ['處理狀態','件數','處理人員','重點說明（未結案優先）'])
    ws3.column_dimensions['A'].width = 14; ws3.column_dimensions['B'].width = 8
    ws3.column_dimensions['C'].width = 22; ws3.column_dimensions['D'].width = 60

    status_groups = {}
    for r in records:
        k = r.get('status') or '未知'
        status_groups.setdefault(k, []).append(r)
    row = 3
    for st, rows in sorted(status_groups.items(), key=lambda x: -len(x[1])):
        handlers = '、'.join(set(r['handler'] for r in rows if r.get('handler')))
        unresolved = [r for r in rows if r.get('status') != '結案']
        resolved = [r for r in rows if r.get('status') == '結案']
        priority = (unresolved + resolved)[:3]
        notes = '\n'.join(f"・[{r.get('company','')}] {r.get('subcategory','')}" + (f" → {r['result']}" if r.get('result') else '') + ('' if r.get('status')=='結案' else ' ⚠未結') for r in priority)
        bg = '1E2235' if row%2==0 else '161925'
        for c2, val in enumerate([st, len(rows), handlers, notes], 1):
            c = ws3.cell(row=row, column=c2, value=val)
            c.font = Font(name='Arial', bold=(c2==1), color=STATUS_COLORS.get(st,'E2E8F0') if c2==1 else 'E2E8F0', size=10)
            c.fill = fill(bg); c.alignment = ca() if c2<=2 else ca('left',wrap=True); c.border = border()
        ws3.row_dimensions[row].height = max(45, len(priority)*20); row += 1
    chart_row = row+1
    ws3.cell(row=chart_row, column=1, value='狀態'); ws3.cell(row=chart_row, column=2, value='件數')
    for i,(st,rows) in enumerate(sorted(status_groups.items(), key=lambda x:-len(x[1])),1):
        ws3.cell(row=chart_row+i, column=1, value=st); ws3.cell(row=chart_row+i, column=2, value=len(rows))
    pie2 = PieChart(); pie2.title='處理狀態分佈'; pie2.style=10; pie2.width=14; pie2.height=10
    lb2 = Reference(ws3, min_col=1, min_row=chart_row+1, max_row=chart_row+len(status_groups))
    d2 = Reference(ws3, min_col=2, min_row=chart_row, max_row=chart_row+len(status_groups))
    pie2.add_data(d2, titles_from_data=True); pie2.set_categories(lb2)
    ws3.add_chart(pie2, 'F2')

    # ===== ④ 逾7天未結案 =====
    ws4 = wb.create_sheet('④ 逾7天未結案')
    ws4.sheet_view.showGridLines = False
    max_days = max([(datetime.now()-datetime.fromisoformat(r['dispatchDate'])).days for r in od], default=0) if od else 0
    ws4.merge_cells('A1:H1')
    sc = ws4.cell(row=1, column=1, value=f'⚠️  截至報告日共 {len(od)} 筆派工超過7天未結案　｜　最長已逾 {max_days} 天　｜　報告日期：{to_date}')
    sc.font = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    sc.fill = fill('7F1D1D'); sc.alignment = ca('left')
    ws4.row_dimensions[1].height = 22
    set_hdr(ws4, 2, ['編號','進線日期','派工日期','公司名稱','問題次分類','處理狀態','負責人員','派工已逾天數'])
    for col, w in [('A',18),('B',16),('C',16),('D',14),('E',26),('F',14),('G',12),('H',14)]:
        ws4.column_dimensions[col].width = w
    row = 3
    for r in sorted(od, key=lambda x: x.get('dispatchDate','')):
        try: days = (datetime.now()-datetime.fromisoformat(r['dispatchDate'])).days
        except: days = 0
        day_color = 'F87171' if days>14 else 'FB923C'
        bg = '2A1515' if row%2==0 else '221212'
        vals = [r.get('id',''),r.get('date','')[:10] if r.get('date') else '',r.get('dispatchDate','')[:10] if r.get('dispatchDate') else '',r.get('company',''),r.get('subcategory',''),r.get('status',''),r.get('handler','—'),f'{days}天']
        colors = ['E2E8F0','94A3B8','94A3B8','FFFFFF','94A3B8',STATUS_COLORS.get(r.get('status',''),'E2E8F0'),'E2E8F0',day_color]
        for c2,(val,color) in enumerate(zip(vals,colors),1):
            c = ws4.cell(row=row, column=c2, value=val)
            c.font = Font(name='Arial', bold=(c2==8), color=color, size=10)
            c.fill = fill(bg); c.alignment = ca(); c.border = border()
        ws4.row_dimensions[row].height = 16; row += 1
    if not od:
        ws4.merge_cells('A3:H3')
        c = ws4.cell(row=3, column=1, value='✅ 本週無派工超過7天未結案')
        c.font = Font(name='Arial', bold=True, color='34D399', size=12); c.alignment = ca()

    return wb

# ===== MONTHLY REPORT =====
def generate_monthly(records, from_date, to_date):
    total = len(records)
    closed = sum(1 for r in records if r.get('status') == '結案')
    open_cnt = total - closed
    close_rate = f'{closed/total*100:.1f}%' if total else '0%'
    od = [r for r in records if is_dispatch_overdue(r)]
    durs = []
    for r in records:
        if r.get('status') == '結案' and r.get('date') and r.get('closeDate'):
            try:
                sec = (datetime.fromisoformat(r['closeDate'])-datetime.fromisoformat(r['date'])).total_seconds()
                if sec > 0: durs.append(sec)
            except: pass
    avg_h = round(sum(durs)/len(durs)/3600) if durs else 0
    avg_str = f'{avg_h//24}天{avg_h%24}小時' if avg_h >= 24 else f'{avg_h}小時'
    label = f'{from_date} ～ {to_date}'
    wb = Workbook()

    # ===== 封面 =====
    ws0 = wb.active; ws0.title = '封面摘要'
    ws0.sheet_view.showGridLines = False
    for col, w in [('A',3),('B',22),('C',18),('D',18),('E',18),('F',18),('G',3)]:
        ws0.column_dimensions[col].width = w
    ws0.merge_cells('B2:F2')
    c = ws0.cell(row=2, column=2, value='⚡ 售服案件月報')
    c.font = Font(name='Arial', bold=True, color='FFFFFF', size=18)
    c.fill = fill('1A1D27'); c.alignment = ca('left'); ws0.row_dimensions[2].height = 36
    ws0.merge_cells('B3:F3')
    c = ws0.cell(row=3, column=2, value=f'報告期間：{label}　｜　產製日期：{to_date}')
    c.font = Font(name='Arial', color='94A3B8', size=11)
    c.fill = fill('1A1D27'); c.alignment = ca('left'); ws0.row_dimensions[3].height = 22
    ws0.row_dimensions[4].height = 12

    kpis = [('進線總件數',total,'5B8CFF','件'),('已結案',closed,'34D399',f'結案率 {close_rate}'),('未結案',open_cnt,'FBBF24','件'),('逾7天未結案',len(od),'F87171','件')]
    for i,(lbl,val,color,sub) in enumerate(kpis):
        col = 2+i
        c = ws0.cell(row=5, column=col, value=lbl)
        c.font = Font(name='Arial', color='94A3B8', size=9); c.fill = fill('22263A'); c.alignment = ca()
        ws0.row_dimensions[5].height = 14
        ws0.merge_cells(start_row=6, start_column=col, end_row=7, end_column=col)
        c = ws0.cell(row=6, column=col, value=val)
        c.font = Font(name='Arial', bold=True, color=color, size=22); c.fill = fill('22263A'); c.alignment = ca()
        ws0.row_dimensions[6].height = 24; ws0.row_dimensions[7].height = 10
        c = ws0.cell(row=8, column=col, value=sub)
        c.font = Font(name='Arial', color='64748B', size=9); c.fill = fill('22263A'); c.alignment = ca()
        ws0.row_dimensions[8].height = 14
    ws0.row_dimensions[9].height = 12
    kpis2 = [('結案率',close_rate,'A78BFA'),('平均處理時間',avg_str,'38BDF8'),('最長逾期天數',f'{max((datetime.now()-datetime.fromisoformat(r["dispatchDate"])).days for r in od) if od else 0}天','FB923C')]
    for i,(lbl,val,color) in enumerate(kpis2):
        col = 2+i
        c = ws0.cell(row=10, column=col, value=lbl)
        c.font = Font(name='Arial', color='94A3B8', size=9); c.fill = fill('1A1D27'); c.alignment = ca()
        ws0.row_dimensions[10].height = 14
        ws0.merge_cells(start_row=11, start_column=col, end_row=12, end_column=col)
        c = ws0.cell(row=11, column=col, value=val)
        c.font = Font(name='Arial', bold=True, color=color, size=18); c.fill = fill('1A1D27'); c.alignment = ca()
        ws0.row_dimensions[11].height = 24; ws0.row_dimensions[12].height = 14
    ws0.row_dimensions[13].height = 16
    ws0.merge_cells('B14:F14')
    c = ws0.cell(row=14, column=2, value='👤 本月人員負責件數')
    c.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    c.fill = fill('2D3250'); c.alignment = ca('left'); ws0.row_dimensions[14].height = 18
    handler_c = {}
    for r in records:
        if r.get('handler'): handler_c[r['handler']] = handler_c.get(r['handler'], 0)+1
    row = 15
    for h, cnt in sorted(handler_c.items(), key=lambda x:-x[1]):
        bg = '1E2235' if row%2==0 else '161925'
        for col, val in [(2,h),(3,f'{cnt} 件')]:
            c = ws0.cell(row=row, column=col, value=val)
            c.font = Font(name='Arial', size=10, color='E2E8F0')
            c.fill = fill(bg); c.alignment = ca('left') if col==2 else ca(); c.border = border()
        ws0.row_dimensions[row].height = 16; row += 1

    # ===== ① 進線管道 =====
    ws1 = wb.create_sheet('① 進線管道分析')
    ws1.sheet_view.showGridLines = False
    title_row(ws1, 1, f'📡 進線管道分析　｜　{label}', 3)
    set_hdr(ws1, 2, ['進線管道','件數','佔比'])
    ws1.column_dimensions['A'].width = 18; ws1.column_dimensions['B'].width = 12; ws1.column_dimensions['C'].width = 12
    ch_c = {}
    for r in records: ch_c[r.get('channel') or '未知'] = ch_c.get(r.get('channel') or '未知',0)+1
    row = 3
    for ch, cnt in sorted(ch_c.items(), key=lambda x:-x[1]):
        bg = '1E2235' if row%2==0 else '161925'
        for c2,(val,color,bold) in enumerate([(ch,'E2E8F0',False),(cnt,'5B8CFF',True),(f'{cnt/total*100:.1f}%' if total else '0%','94A3B8',False)],1):
            c = ws1.cell(row=row, column=c2, value=val)
            c.font = Font(name='Arial', bold=bold, color=color, size=10)
            c.fill = fill(bg); c.alignment = ca(); c.border = border()
        row += 1
    for c2, val in enumerate(['合計',total,'100%'],1):
        c = ws1.cell(row=row, column=c2, value=val)
        c.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
        c.fill = fill('2D3250'); c.alignment = ca(); c.border = border()
    pie = PieChart(); pie.title='進線管道佔比'; pie.style=10; pie.width=14; pie.height=10
    lb = Reference(ws1, min_col=1, min_row=3, max_row=2+len(ch_c))
    dt = Reference(ws1, min_col=2, min_row=2, max_row=2+len(ch_c))
    pie.add_data(dt, titles_from_data=True); pie.set_categories(lb)
    ws1.add_chart(pie, 'E2')

    # ===== ② 案件類別 =====
    ws2 = wb.create_sheet('② 案件類別分析')
    ws2.sheet_view.showGridLines = False
    title_row(ws2, 1, f'📦 案件類別分析（產品別）　｜　{label}', 3)
    set_hdr(ws2, 2, ['產品類別','件數','佔比'])
    ws2.column_dimensions['A'].width = 16; ws2.column_dimensions['B'].width = 10; ws2.column_dimensions['C'].width = 12
    prod_map = {'FMS-GPS':0,'FMS-DMVR':0,'FMS-冷鏈':0,'FMS-雷達':0,'其他':0}
    for r in records: prod_map[map_product(r.get('product',''))] += 1
    prod_list = [(k,v) for k,v in prod_map.items() if v>0]
    prod_list.sort(key=lambda x:-x[1])
    total_prod = sum(v for _,v in prod_list)
    PROD_COLORS = ['5B8CFF','34D399','FBBF24','A78BFA','94A3B8']
    row = 3
    for i,(prod,cnt) in enumerate(prod_list):
        bg = '1E2235' if row%2==0 else '161925'
        for c2,(val,color,bold) in enumerate([(prod,PROD_COLORS[i%5],True),(cnt,'FFFFFF',True),(f'{cnt/total_prod*100:.1f}%' if total_prod else '0%','94A3B8',False)],1):
            c = ws2.cell(row=row, column=c2, value=val)
            c.font = Font(name='Arial', bold=bold, color=color, size=10)
            c.fill = fill(bg); c.alignment = ca(); c.border = border()
        ws2.row_dimensions[row].height = 18; row += 1
    for c2,(val,color) in enumerate([('合計','FFFFFF'),(total_prod,'FFFFFF'),('100%','FFFFFF')],1):
        c = ws2.cell(row=row, column=c2, value=val)
        c.font = Font(name='Arial', bold=True, color=color, size=10)
        c.fill = fill('2D3250'); c.alignment = ca(); c.border = border()
    chart_row = row+2
    ws2.cell(row=chart_row, column=1, value='產品類別'); ws2.cell(row=chart_row, column=2, value='件數')
    for i2,(prod,cnt) in enumerate(prod_list,1):
        ws2.cell(row=chart_row+i2, column=1, value=prod); ws2.cell(row=chart_row+i2, column=2, value=cnt)
    pie2 = PieChart(); pie2.title='產品類別佔比'; pie2.style=10; pie2.width=16; pie2.height=12
    lb2 = Reference(ws2, min_col=1, min_row=chart_row+1, max_row=chart_row+len(prod_list))
    d2 = Reference(ws2, min_col=2, min_row=chart_row, max_row=chart_row+len(prod_list))
    pie2.add_data(d2, titles_from_data=True); pie2.set_categories(lb2)
    ws2.add_chart(pie2, 'E2')

    # ===== ③ 客戶進線 TOP5 =====
    ws3 = wb.create_sheet('③ 客戶進線排行')
    ws3.sheet_view.showGridLines = False
    title_row(ws3, 1, f'🏢 客戶進線排行 TOP5　｜　{label}', 4)
    set_hdr(ws3, 2, ['排名','客戶名稱','件數','主要問題'])
    ws3.column_dimensions['A'].width = 8; ws3.column_dimensions['B'].width = 16
    ws3.column_dimensions['C'].width = 10; ws3.column_dimensions['D'].width = 45
    company_c = {}
    company_issues = {}
    for r in records:
        co = r.get('company','')
        if co:
            company_c[co] = company_c.get(co,0)+1
            company_issues.setdefault(co, {})
            cat = r.get('category','其他')
            company_issues[co][cat] = company_issues[co].get(cat,0)+1
    top5 = sorted(company_c.items(), key=lambda x:-x[1])[:5]
    rank_colors = ['FFD700','C0C0C0','CD7F32','E2E8F0','E2E8F0']
    row = 3
    for rank,(co,cnt) in enumerate(top5,1):
        issues = '、'.join(f"{k}({v})" for k,v in sorted(company_issues.get(co,{}).items(), key=lambda x:-x[1])[:2])
        bg = '1E2235' if row%2==0 else '161925'
        for c2,(val,color,bold) in enumerate([(rank,rank_colors[rank-1],rank<=3),(co,'FFFFFF',False),(cnt,'5B8CFF',True),(issues,'94A3B8',False)],1):
            c = ws3.cell(row=row, column=c2, value=val)
            c.font = Font(name='Arial', bold=bold, color=color, size=10)
            c.fill = fill(bg); c.alignment = ca() if c2!=4 else ca('left',wrap=True); c.border = border()
        ws3.row_dimensions[row].height = 16; row += 1
    bar2 = BarChart(); bar2.type='bar'; bar2.title='客戶進線件數 TOP5'; bar2.style=10; bar2.width=16; bar2.height=10
    cats_r = Reference(ws3, min_col=2, min_row=3, max_row=2+len(top5))
    data_r = Reference(ws3, min_col=3, min_row=2, max_row=2+len(top5))
    bar2.add_data(data_r, titles_from_data=True); bar2.set_categories(cats_r)
    ws3.add_chart(bar2, 'F2')

    # ===== ④ 逾7天未結案 =====
    ws4 = wb.create_sheet('④ 逾7天未結案')
    ws4.sheet_view.showGridLines = False
    max_days = max([(datetime.now()-datetime.fromisoformat(r['dispatchDate'])).days for r in od if r.get('dispatchDate')], default=0)
    ws4.merge_cells('A1:H1')
    sc = ws4.cell(row=1, column=1, value=f'⚠️  截至月底共 {len(od)} 筆派工超過7天未結案　｜　最長已逾 {max_days} 天　｜　{label}')
    sc.font = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    sc.fill = fill('7F1D1D'); sc.alignment = ca('left'); ws4.row_dimensions[1].height = 22
    set_hdr(ws4, 2, ['編號','進線日期','派工日期','公司名稱','問題次分類','處理狀態','負責人員','派工已逾天數'])
    for col, w in [('A',18),('B',16),('C',16),('D',14),('E',26),('F',14),('G',12),('H',14)]:
        ws4.column_dimensions[col].width = w
    row = 3
    for r in sorted(od, key=lambda x: x.get('dispatchDate','')):
        try: days = (datetime.now()-datetime.fromisoformat(r['dispatchDate'])).days
        except: days = 0
        day_color = 'F87171' if days>14 else 'FB923C'
        bg = '2A1515' if row%2==0 else '221212'
        vals = [r.get('id',''),r.get('date','')[:10] if r.get('date') else '',r.get('dispatchDate','')[:10] if r.get('dispatchDate') else '',r.get('company',''),r.get('subcategory',''),r.get('status',''),r.get('handler','—'),f'{days}天']
        colors = ['E2E8F0','94A3B8','94A3B8','FFFFFF','94A3B8',STATUS_COLORS.get(r.get('status',''),'E2E8F0'),'E2E8F0',day_color]
        for c2,(val,color) in enumerate(zip(vals,colors),1):
            c = ws4.cell(row=row, column=c2, value=val)
            c.font = Font(name='Arial', bold=(c2==8), color=color, size=10)
            c.fill = fill(bg); c.alignment = ca(); c.border = border()
        ws4.row_dimensions[row].height = 16; row += 1
    if not od:
        ws4.merge_cells('A3:H3')
        c = ws4.cell(row=3, column=1, value='✅ 本月無逾期未結案')
        c.font = Font(name='Arial', bold=True, color='34D399', size=12); c.alignment = ca()

    # ===== ⑤ 問題分類 =====
    ws5 = wb.create_sheet('⑤ 問題分類統計')
    ws5.sheet_view.showGridLines = False
    title_row(ws5, 1, f'🏷️ 問題分類統計　｜　{label}', 3)
    set_hdr(ws5, 2, ['問題大類','總件數','主要次分類（件數）'])
    ws5.column_dimensions['A'].width = 14; ws5.column_dimensions['B'].width = 10; ws5.column_dimensions['C'].width = 50
    cat_map = {}
    for r in records:
        k = r.get('category') or '其他'
        if k not in cat_map: cat_map[k] = {'total':0,'subs':{}}
        cat_map[k]['total'] += 1
        s = r.get('subcategory') or '其他'
        cat_map[k]['subs'][s] = cat_map[k]['subs'].get(s,0)+1
    sorted_cats = sorted(cat_map.items(), key=lambda x:-x[1]['total'])
    row = 3
    for cat, v in sorted_cats:
        top = '、'.join(f"{s}({n})" for s,n in sorted(v['subs'].items(), key=lambda x:-x[1])[:4])
        bg = '1E2235' if row%2==0 else '161925'
        for c2,(val,color,bold) in enumerate([(cat,'E2E8F0',False),(v['total'],'5B8CFF',True),(top,'94A3B8',False)],1):
            c = ws5.cell(row=row, column=c2, value=val)
            c.font = Font(name='Arial', bold=bold, color=color, size=10)
            c.fill = fill(bg); c.alignment = ca() if c2<=2 else ca('left',wrap=True); c.border = border()
        ws5.row_dimensions[row].height = 16; row += 1
    bar3 = BarChart(); bar3.type='bar'; bar3.title='問題大類件數'; bar3.style=10; bar3.width=16; bar3.height=12
    cats_ref = Reference(ws5, min_col=1, min_row=3, max_row=2+len(sorted_cats))
    data_ref = Reference(ws5, min_col=2, min_row=2, max_row=2+len(sorted_cats))
    bar3.add_data(data_ref, titles_from_data=True); bar3.set_categories(cats_ref)
    ws5.add_chart(bar3, 'E2')

    # ===== ⑥ 處理狀態 =====
    ws6 = wb.create_sheet('⑥ 處理狀態總覽')
    ws6.sheet_view.showGridLines = False
    title_row(ws6, 1, f'📋 處理狀態總覽　｜　{label}', 4)
    set_hdr(ws6, 2, ['處理狀態','件數','處理人員','重點說明'])
    ws6.column_dimensions['A'].width = 14; ws6.column_dimensions['B'].width = 8
    ws6.column_dimensions['C'].width = 22; ws6.column_dimensions['D'].width = 60
    status_groups = {}
    for r in records:
        k = r.get('status') or '未知'
        status_groups.setdefault(k, []).append(r)
    row = 3
    for st, rows in sorted(status_groups.items(), key=lambda x:-len(x[1])):
        handlers = '、'.join(set(r['handler'] for r in rows if r.get('handler')))
        unresolved = [r for r in rows if r.get('status') != '結案']
        resolved = [r for r in rows if r.get('status') == '結案']
        priority = (unresolved+resolved)[:3]
        notes = '\n'.join(f"・[{r.get('company','')}] {r.get('subcategory','')}"+(f" → {r['result']}" if r.get('result') else '')+('' if r.get('status')=='結案' else ' ⚠未結') for r in priority)
        bg = '1E2235' if row%2==0 else '161925'
        for c2, val in enumerate([st,len(rows),handlers,notes],1):
            c = ws6.cell(row=row, column=c2, value=val)
            c.font = Font(name='Arial', bold=(c2==1), color=STATUS_COLORS.get(st,'E2E8F0') if c2==1 else 'E2E8F0', size=10)
            c.fill = fill(bg); c.alignment = ca() if c2<=2 else ca('left',wrap=True); c.border = border()
        ws6.row_dimensions[row].height = max(45, len(priority)*20); row += 1
    chart_row = row+1
    ws6.cell(row=chart_row, column=1, value='狀態'); ws6.cell(row=chart_row, column=2, value='件數')
    for i,(st,rows) in enumerate(sorted(status_groups.items(), key=lambda x:-len(x[1])),1):
        ws6.cell(row=chart_row+i, column=1, value=st); ws6.cell(row=chart_row+i, column=2, value=len(rows))
    pie3 = PieChart(); pie3.title='處理狀態分佈'; pie3.style=10; pie3.width=14; pie3.height=10
    lb3 = Reference(ws6, min_col=1, min_row=chart_row+1, max_row=chart_row+len(status_groups))
    d3 = Reference(ws6, min_col=2, min_row=chart_row, max_row=chart_row+len(status_groups))
    pie3.add_data(d3, titles_from_data=True); pie3.set_categories(lb3)
    ws6.add_chart(pie3, 'F2')

    return wb

# ===== API ROUTES =====
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'time': datetime.now().isoformat()})

@app.route('/weekly-report', methods=['POST'])
def weekly_report():
    try:
        data = request.json
        records = data.get('records', [])
        from_date = data.get('from', '')
        to_date = data.get('to', '')
        if not records:
            return jsonify({'error': '無資料'}), 400
        wb = generate_weekly(records, from_date, to_date)
        buf = io.BytesIO()
        wb.save(buf); buf.seek(0)
        filename = f'週報_{from_date}_{to_date}.xlsx'
        return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/monthly-report', methods=['POST'])
def monthly_report():
    try:
        data = request.json
        records = data.get('records', [])
        from_date = data.get('from', '')
        to_date = data.get('to', '')
        if not records:
            return jsonify({'error': '無資料'}), 400
        wb = generate_monthly(records, from_date, to_date)
        buf = io.BytesIO()
        wb.save(buf); buf.seek(0)
        filename = f'月報_{from_date}_{to_date}.xlsx'
        return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
