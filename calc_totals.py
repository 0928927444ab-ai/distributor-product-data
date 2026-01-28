import re
f = open(r'摩瑪已下單產品\門用五金_摩瑪建材\【母表】摩瑪門用五金_完整內部資料.html', 'r', encoding='utf-8')
text = f.read()
f.close()

total = 0
cats = {'lock': [0,0,0], 'hinge': [0,0,0], 'stop': [0,0,0], 'acc': [0,0,0]}

for row in re.findall(r'<tr[^>]*>.*?</tr>', text, re.DOTALL):
    cells = re.findall(r'<td[^>]*>(.*?)</td>', row)
    if not cells: continue
    model = re.sub(r'<[^>]+>', '', cells[0]).strip()
    rmb_i = None
    for i, c in enumerate(cells):
        if chr(0xFFE5) in c:
            rmb_i = i
            break
    if rmb_i is None: continue
    lp_text = cells[rmb_i + 3]
    m = re.search(r'\$([\d,]+)', lp_text)
    if not m: continue
    lp = float(m.group(1).replace(',', ''))
    qty_text = cells[-1]
    m2 = re.search(r'(\d+)', qty_text)
    if not m2: continue
    qty = int(m2.group(1))
    val = lp * qty
    total += val
    if model.startswith(('T86','T80','T61','T8A','T52','A6-','AY-','WY','Y09')):
        cats['lock'][0] += 1; cats['lock'][1] += qty; cats['lock'][2] += val
    elif model.startswith(('V-','ST-')):
        cats['hinge'][0] += 1; cats['hinge'][1] += qty; cats['hinge'][2] += val
    elif model.startswith('W-'):
        cats['stop'][0] += 1; cats['stop'][1] += qty; cats['stop'][2] += val
    elif model.startswith('JYT'):
        cats['acc'][0] += 1; cats['acc'][1] += qty; cats['acc'][2] += val

for k, v in cats.items():
    cnt, qty, val = v
    avg = val/qty if qty else 0
    print(f'{k:6}: {cnt} items, {qty} pcs, avg=${avg:,.0f}, total=~${val/1000:,.0f}K')

all_cnt = sum(v[0] for v in cats.values())
all_qty = sum(v[1] for v in cats.values())
print(f'TOTAL:  {all_cnt} items, {all_qty} pcs, total=~${total/1000:,.0f}K')
print(f'7.5 fold: ~${total*0.75/1000:,.0f}K')
