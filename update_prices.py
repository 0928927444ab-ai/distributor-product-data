#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Update pricing formula in 門用五金母表:
OLD: 牌價 = 成本(￥)×5×5.2 (+$300 for locks)
NEW: 牌價 = (成本(￥)×5 + 運費)×5.2 (+$300 for locks)
"""
import re

FILE = r"C:\Users\詹家騏\Desktop\摩瑪建材產品線\摩瑪已下單產品\門用五金_摩瑪建材\【母表】摩瑪門用五金_完整內部資料.html"

# Door lock model prefixes (get +$300)
LOCK_PREFIXES = ('T86', 'T80', 'T61', 'T8A', 'T52', 'A6-529', 'AY-', 'WY11', 'Y09')

def is_lock(model):
    return any(model.startswith(p) for p in LOCK_PREFIXES)

def fmt(n):
    """Format number as $X,XXX with comma separator"""
    n = round(n)
    if n >= 1000:
        return f"${n:,}"
    return f"${n}"

def parse_rmb(s):
    """Extract number from ￥150 format"""
    m = re.search(r'[￥¥](\d+(?:\.\d+)?)', s)
    return float(m.group(1)) if m else None

def parse_weight(s):
    """Extract weight from '2 kg' format"""
    m = re.search(r'([\d.]+)\s*kg', s)
    return float(m.group(1)) if m else None

def parse_dollar(s):
    """Extract number from $4,200 format"""
    m = re.search(r'\$([\d,]+)', s)
    return float(m.group(1).replace(',', '')) if m else None

def process_row(tr_html):
    """Process a single <tr> row and update prices."""
    # Extract all <td> cells
    cells = re.findall(r'<td[^>]*>(.*?)</td>', tr_html)
    if len(cells) < 8:
        return tr_html

    # Find the model (first cell text, strip tags)
    model_text = re.sub(r'<[^>]+>', '', cells[0]).strip()

    # Find RMB cost cell and weight cell
    rmb_val = None
    weight_val = None
    rmb_idx = None
    weight_idx = None

    for i, cell in enumerate(cells):
        if rmb_val is None and ('￥' in cell or '¥' in cell):
            rmb_val = parse_rmb(cell)
            rmb_idx = i
        if weight_val is None and 'kg' in cell:
            weight_val = parse_weight(cell)
            weight_idx = i

    if rmb_val is None or weight_val is None:
        return tr_html

    # Calculate new prices
    cost_twd = rmb_val * 5
    shipping = weight_val * 65
    lock = is_lock(model_text)

    base = (cost_twd + shipping) * 5.2
    list_price = base + 300 if lock else base
    trade_price = list_price * 0.7
    bni_price = list_price * 0.6
    dealer_price = list_price * 0.35

    new_shipping = fmt(round(shipping))
    new_list = fmt(list_price)
    new_trade = fmt(trade_price)
    new_bni = fmt(bni_price)
    new_dealer = fmt(dealer_price)

    # Find the price cells to replace (after RMB cell)
    # Pattern: after RMB cost, there's 成本($), 運費, 牌價, 同行, BNI, 經銷
    # We need to find and replace these specific cells

    # Build list of td elements with their full match
    td_pattern = r'<td[^>]*>.*?</td>'
    td_matches = list(re.finditer(td_pattern, tr_html))

    if len(td_matches) < rmb_idx + 5:
        return tr_html

    # The cells after RMB are: 成本($), 運費, 牌價, 同行, BNI, 經銷
    # cost($) idx = rmb_idx + 1
    # shipping idx = rmb_idx + 2
    # list price idx = rmb_idx + 3
    # trade idx = rmb_idx + 4
    # bni idx = rmb_idx + 5
    # dealer idx = rmb_idx + 6

    cost_twd_idx = rmb_idx + 1
    ship_idx = rmb_idx + 2
    list_idx = rmb_idx + 3
    trade_idx = rmb_idx + 4
    bni_idx = rmb_idx + 5
    dealer_idx = rmb_idx + 6

    if dealer_idx >= len(td_matches):
        return tr_html

    # Verify these cells contain $ values
    replacements = {
        ship_idx: new_shipping,
        list_idx: new_list,
        trade_idx: new_trade,
        bni_idx: new_bni,
        dealer_idx: new_dealer,
    }

    result = tr_html
    # Replace from right to left to preserve positions
    for idx in sorted(replacements.keys(), reverse=True):
        if idx >= len(td_matches):
            continue
        match = td_matches[idx]
        old_td = match.group()
        # Extract class attribute
        class_match = re.search(r'class="([^"]*)"', old_td)
        cls = f' class="{class_match.group(1)}"' if class_match else ''
        new_td = f'<td{cls}>{replacements[idx]}</td>'
        result = result[:match.start()] + new_td + result[match.end():]

    return result

def main():
    with open(FILE, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Update formula description
    content = content.replace(
        '匯率 1 RMB = 5 TWD | 牌價倍率 5.2 | 把手類加運費 $300',
        '匯率 1 RMB = 5 TWD | 牌價倍率 5.2 | 門鎖類加 $300 | 運費 $65/kg 計入成本'
    )

    content = content.replace(
        '<b>公式：</b> 把手類牌價 = 成本(￥) &times; 5 &times; 5.2 + $300 | 其他品類牌價 = 成本(￥) &times; 5 &times; 5.2',
        '<b>公式：</b> 門鎖類牌價 = (成本(￥) &times; 5 + 運費) &times; 5.2 + $300 | 其他品類牌價 = (成本(￥) &times; 5 + 運費) &times; 5.2'
    )

    # 2. Update T86 example in pricing table
    content = content.replace(
        '<td class="amt">$4,200</td>\n                    </tr>\n                    <tr><td><b>同行價</b></td><td>7 折</td><td>同業/設計師</td><td class="amt">$2,940</td></tr>\n                    <tr><td><b>BNI會員價</b></td><td>6 折</td><td>BNI綠燈會員</td><td class="amt">$2,520</td></tr>\n                    <tr><td><b>經銷價</b></td><td>3.5 折</td><td>經銷商</td><td class="amt">$1,470</td>',
        '<td class="amt">$4,876</td>\n                    </tr>\n                    <tr><td><b>同行價</b></td><td>7 折</td><td>同業/設計師</td><td class="amt">$3,413</td></tr>\n                    <tr><td><b>BNI會員價</b></td><td>6 折</td><td>BNI綠燈會員</td><td class="amt">$2,926</td></tr>\n                    <tr><td><b>經銷價</b></td><td>3.5 折</td><td>經銷商</td><td class="amt">$1,707</td>'
    )

    # 3. Process all product rows
    # Find all <tr>...</tr> that contain ￥ (RMB prices)
    tr_pattern = r'<tr>(?:(?!</tr>).)*￥(?:(?!</tr>).)*</tr>'

    def replace_tr(match):
        return process_row(match.group())

    content = re.sub(tr_pattern, replace_tr, content, flags=re.DOTALL)

    with open(FILE, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Done! Prices updated.")

    # Verify a few samples
    with open(FILE, 'r', encoding='utf-8') as f:
        text = f.read()

    # Check T86-YN
    if '$4,876' in text:
        print("✓ T86-YN 牌價 = $4,876 (was $4,200)")
    else:
        print("✗ T86-YN check failed")

    # Check V-N3D-HL
    if '$18,434' in text:
        print("✓ V-N3D-HL 牌價 = $18,434 (was $17,420)")
    else:
        print("✗ V-N3D-HL check failed")

    # Check W-829-HY
    if '$1,368' in text:
        print("✓ W-829-HY 牌價 = $1,368 (was $1,196)")
    else:
        print("✗ W-829-HY check failed")

if __name__ == '__main__':
    main()
