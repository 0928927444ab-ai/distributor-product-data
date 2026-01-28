"""移除母表中重複的運費格子（相鄰兩個相同金額格子，去掉一個）"""
import re

filepath = r'C:\Users\詹家騏\Desktop\摩瑪建材產品線\摩瑪已下單產品\門用五金_摩瑪建材\【母表】摩瑪門用五金_完整內部資料.html'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Pattern: two consecutive identical amount cells
# e.g. <td class="amt">$130</td><td class="amt">$130</td>
pattern = r'(<td class="amt">\$[\d,]+</td>)<td class="amt">(\$[\d,]+)</td>'

count = 0
def fix_double(m):
    global count
    cell1_full = m.group(0)
    first_val = re.search(r'\$([\d,]+)', m.group(1)).group(0)
    second_val = m.group(2)
    if first_val == second_val:
        count += 1
        return m.group(1)  # Keep only first
    return cell1_full  # Different values, keep both

content = re.sub(pattern, fix_double, content)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print(f'Removed {count} duplicate shipping cells')
