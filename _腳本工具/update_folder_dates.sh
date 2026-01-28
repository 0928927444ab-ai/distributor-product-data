#!/bin/bash
# 自動更新資料夾日期前綴腳本
cd "C:/Users/詹家騏/Desktop/總代理產品資料"

for dir in [0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]_*/; do
  if [ -d "$dir" ]; then
    # 取得資料夾名稱
    dirname="${dir%/}"
    # 取得舊日期和基礎名稱
    old_date="${dirname:0:8}"
    base_name="${dirname:9}"
    # 取得實際修改日期
    new_date=$(stat -c %y "$dirname" | cut -d' ' -f1 | tr -d '-')
    # 如果日期不同，重新命名
    if [ "$old_date" != "$new_date" ]; then
      new_name="${new_date}_${base_name}"
      mv "$dirname" "$new_name" 2>/dev/null
      echo "更新: $dirname -> $new_name" >> "_腳本工具/日期更新紀錄.log"
    fi
  fi
done
