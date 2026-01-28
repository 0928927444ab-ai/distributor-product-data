#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Update all product tables in 門用五金母表 to new pricing:
- Multiplier: 3.85 (was 5.2)
- No +$300 for locks
- New tiers: 設計師BNI 7.5折 / 建材五金行 6折 / 專案價 5折 / 經銷 3.5折
- Adds one extra column (專案價 5折) to each product row
"""
import re

FILE = r"C:\Users\詹家騏\Desktop\摩瑪建材產品線\摩瑪已下單產品\門用五金_摩瑪建材\【母表】摩瑪門用五金_完整內部資料.html"
MULTIPLIER = 3.85

def fmt(n):
    n = round(n)
    if n >= 1000:
        return f"${n:,}"
    return f"${n}"

def parse_rmb(s):
    m = re.search(r'[￥¥](\d+(?:\.\d+)?)', s)
    return float(m.group(1)) if m else None

def parse_weight(s):
    m = re.search(r'([\d.]+)\s*kg', s)
    return float(m.group(1)) if m else None

def process_row(match):
    tr_html = match.group()
    cells = re.findall(r'<td[^>]*>(.*?)</td>', tr_html)

    if len(cells) < 8:
        return tr_html

    # Count td elements - skip already-updated rows (12+ cells = new format)
    td_matches = list(re.finditer(r'<td[^>]*>.*?</td>', tr_html))
    if len(td_matches) >= 12:
        return tr_html

    # Find RMB and weight
    rmb_val = None
    weight_val = None
    rmb_idx = None

    for i, cell in enumerate(cells):
        if rmb_val is None and ('\uffe5' in cell or '\u00a5' in cell):
            rmb_val = parse_rmb(cell)
            rmb_idx = i
        if weight_val is None and 'kg' in cell:
            weight_val = parse_weight(cell)

    if rmb_val is None or weight_val is None or rmb_idx is None:
        return tr_html

    # Calculate new prices
    cost_twd = rmb_val * 5
    shipping = weight_val * 65

    list_price = round((cost_twd + shipping) * MULTIPLIER)
    designer = round(list_price * 0.75)
    hardware = round(list_price * 0.60)
    project = round(list_price * 0.50)
    dealer = round(list_price * 0.35)

    # Indices: after RMB cell
    ship_idx = rmb_idx + 2
    list_idx = rmb_idx + 3
    old_trade_idx = rmb_idx + 4   # was 同行7折
    old_bni_idx = rmb_idx + 5     # was BNI 6折
    old_dealer_idx = rmb_idx + 6  # was 經銷3.5折

    if old_dealer_idx >= len(td_matches):
        return tr_html

    result = tr_html

    # Replace right-to-left to preserve positions
    # 1. old_dealer -> project + dealer (1 cell becomes 2 cells)
    m = td_matches[old_dealer_idx]
    cls_m = re.search(r'class="([^"]*)"', m.group())
    cls = f' class="{cls_m.group(1)}"' if cls_m else ''
    result = result[:m.start()] + f'<td{cls}>{fmt(project)}</td><td{cls}>{fmt(dealer)}</td>' + result[m.end():]

    # 2. old_bni -> hardware price
    m = td_matches[old_bni_idx]
    cls_m = re.search(r'class="([^"]*)"', m.group())
    cls = f' class="{cls_m.group(1)}"' if cls_m else ''
    result = result[:m.start()] + f'<td{cls}>{fmt(hardware)}</td>' + result[m.end():]

    # 3. old_trade -> designer price
    m = td_matches[old_trade_idx]
    cls_m = re.search(r'class="([^"]*)"', m.group())
    cls = f' class="{cls_m.group(1)}"' if cls_m else ''
    result = result[:m.start()] + f'<td{cls}>{fmt(designer)}</td>' + result[m.end():]

    # 4. list price
    m = td_matches[list_idx]
    cls_m = re.search(r'class="([^"]*)"', m.group())
    cls = f' class="{cls_m.group(1)}"' if cls_m else ''
    result = result[:m.start()] + f'<td{cls}>{fmt(list_price)}</td>' + result[m.end():]

    # 5. shipping
    m = td_matches[ship_idx]
    cls_m = re.search(r'class="([^"]*)"', m.group())
    cls = f' class="{cls_m.group(1)}"' if cls_m else ''
    result = result[:m.start()] + f'<td{cls}>{fmt(round(shipping))}</td>' + result[m.end():]

    return result

def main():
    with open(FILE, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Replace all old-format table headers
    old_hdr = '<th class="amt">\u540c\u884c7\u6298</th><th class="amt">BNI 6\u6298</th><th class="amt">\u7d93\u92b73.5\u6298</th>'
    new_hdr = '<th class="amt">\u8a2d\u8a08\u5e2bBNI 7.5\u6298</th><th class="amt">\u5efa\u6750\u4e94\u91d1\u884c 6\u6298</th><th class="amt">\u5c08\u6848\u50f9 5\u6298</th><th class="amt">\u7d93\u92b73.5\u6298</th>'
    hdr_count = content.count(old_hdr)
    content = content.replace(old_hdr, new_hdr)
    print(f"Replaced {hdr_count} table headers")

    # 2. Process all product rows containing RMB prices
    tr_pattern = r'<tr>(?:(?!</tr>).)*\uffe5(?:(?!</tr>).)*</tr>'

    updated = [0]
    def replace_tr(match):
        result = process_row(match)
        if result != match.group():
            updated[0] += 1
        return result

    content = re.sub(tr_pattern, replace_tr, content, flags=re.DOTALL)
    print(f"Updated {updated[0]} product rows")

    with open(FILE, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Done! All prices updated with 3.85x multiplier.")

    # Verify samples
    with open(FILE, 'r', encoding='utf-8') as f:
        text = f.read()

    # T80-YN: 135*5=675, shipping=130, list=(675+130)*3.85=3099
    if '$3,099' in text:
        print("OK: T80-YN list = $3,099")
    else:
        print("FAIL: T80-YN")

    # V-N3D-HL: 670*5=3350, 3kg, shipping=195, list=(3350+195)*3.85=13649
    if '$13,649' in text:
        print("OK: V-N3D-HL list = $13,649")
    else:
        print("FAIL: V-N3D-HL")

    # W-829-HY: 46*5=230, 0.8kg, shipping=52, list=(230+52)*3.85=1086
    if '$1,086' in text:
        print("OK: W-829-HY list = $1,086")
    else:
        print("FAIL: W-829-HY")

if __name__ == '__main__':
    main()
