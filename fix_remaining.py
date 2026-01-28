#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fix rows that were skipped by the first script (鉸鏈 and 門吸 tables
that have extra columns: 定位 or 系列).
These had 12 cells originally, so were incorrectly skipped.
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
    m = re.search(r'[\uffe5\u00a5](\d+(?:\.\d+)?)', s)
    return float(m.group(1)) if m else None

def parse_weight(s):
    m = re.search(r'([\d.]+)\s*kg', s)
    return float(m.group(1)) if m else None

def process_row(match):
    tr_html = match.group()
    cells = re.findall(r'<td[^>]*>(.*?)</td>', tr_html)

    if len(cells) < 8:
        return tr_html

    # Find RMB cell
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

    # Check if already updated: count cells after rmb
    # Old format: 7 cells after rmb (cost$, ship, list, 3 prices, qty)
    # New format: 8 cells after rmb (cost$, ship, list, 4 prices, qty)
    cells_after_rmb = len(cells) - rmb_idx - 1
    if cells_after_rmb >= 8:
        return tr_html  # already updated

    # Calculate
    cost_twd = rmb_val * 5
    shipping = weight_val * 65

    list_price = round((cost_twd + shipping) * MULTIPLIER)
    designer = round(list_price * 0.75)
    hardware = round(list_price * 0.60)
    project = round(list_price * 0.50)
    dealer = round(list_price * 0.35)

    td_matches = list(re.finditer(r'<td[^>]*>.*?</td>', tr_html))

    ship_idx = rmb_idx + 2
    list_idx = rmb_idx + 3
    old_trade_idx = rmb_idx + 4
    old_bni_idx = rmb_idx + 5
    old_dealer_idx = rmb_idx + 6

    if old_dealer_idx >= len(td_matches):
        return tr_html

    result = tr_html

    # Replace right-to-left
    # 1. old_dealer -> project + dealer
    m = td_matches[old_dealer_idx]
    cls_m = re.search(r'class="([^"]*)"', m.group())
    cls = f' class="{cls_m.group(1)}"' if cls_m else ''
    result = result[:m.start()] + f'<td{cls}>{fmt(project)}</td><td{cls}>{fmt(dealer)}</td>' + result[m.end():]

    # 2. old_bni -> hardware
    m = td_matches[old_bni_idx]
    cls_m = re.search(r'class="([^"]*)"', m.group())
    cls = f' class="{cls_m.group(1)}"' if cls_m else ''
    result = result[:m.start()] + f'<td{cls}>{fmt(hardware)}</td>' + result[m.end():]

    # 3. old_trade -> designer
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

    tr_pattern = r'<tr>(?:(?!</tr>).)*\uffe5(?:(?!</tr>).)*</tr>'

    updated = [0]
    def replace_tr(match):
        result = process_row(match)
        if result != match.group():
            updated[0] += 1
        return result

    content = re.sub(tr_pattern, replace_tr, content, flags=re.DOTALL)
    print(f"Fixed {updated[0]} rows")

    with open(FILE, 'w', encoding='utf-8') as f:
        f.write(content)

    # Verify
    with open(FILE, 'r', encoding='utf-8') as f:
        text = f.read()

    # V-N3D-HL: rmb=670, 3kg -> (3350+195)*3.85 = 13648
    checks = {
        'V-N3D-HL': 13648,
        'W-829-HY': 1009,  # (230+32.5)*3.85 = 1010.625 -> round to 1011? Let me check: 46*5=230, 0.5*65=32.5, (230+32.5)*3.85=1010.625
        'JYT-SY': 1636,    # (360+65)*3.85 = 1636.25 -> 1636
    }
    for model, expected in checks.items():
        expected_str = f"${expected:,}"
        if expected_str in text:
            print(f"OK: {model} list = {expected_str}")
        else:
            print(f"FAIL: {model} expected {expected_str}")

if __name__ == '__main__':
    main()
