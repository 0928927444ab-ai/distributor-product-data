"""為母表 HTML 中缺少運費格子的資料列插入運費欄位"""
import re, decimal

def round_half_up(x):
    return int(decimal.Decimal(str(x)).quantize(decimal.Decimal('1'), rounding=decimal.ROUND_HALF_UP))

filepath = r'C:\Users\詹家騏\Desktop\摩瑪建材產品線\摩瑪已下單產品\門用五金_摩瑪建材\【母表】摩瑪門用五金_完整內部資料.html'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Track current table header column count
current_header_cols = 0
count = 0

for i, line in enumerate(lines):
    stripped = line.strip()

    # Detect table headers to know expected column count
    if '<thead>' in stripped:
        current_header_cols = stripped.count('<th')
        continue

    # Only process data rows with weight
    if '<tr><td>' not in stripped:
        continue

    weight_match = re.search(r'(\d+\.?\d*)\s*kg</td>', stripped)
    if not weight_match:
        continue

    weight = float(weight_match.group(1))
    td_count = stripped.count('<td')

    # If row already has all columns (td_count == current_header_cols), skip
    if td_count >= current_header_cols:
        continue

    # Row is missing the shipping cell
    shipping = round_half_up(weight * 65)
    shipping_str = f'${shipping:,}'

    # Insert shipping cell after 成本($) cell
    # Pattern: ￥NUMBER</td><td class="amt">$NUMBER</td> followed by <td
    # We insert <td class="amt">$SHIPPING</td> after the TWD cost cell
    new_line = re.sub(
        r'(￥\d+</td><td class="amt">\$[\d,]+</td>)(<td)',
        r'\1<td class="amt">' + shipping_str + r'</td>\2',
        line,
        count=1  # Only first occurrence per line
    )

    if new_line != line:
        lines[i] = new_line
        count += 1
        print(f'  Line {i+1}: weight={weight}kg, shipping={shipping_str}')

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print(f'\nTotal: Added shipping to {count} rows')
