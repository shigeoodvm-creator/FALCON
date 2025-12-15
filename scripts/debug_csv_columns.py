"""
CSVの列構造を確認するスクリプト
"""

import pandas as pd
from pathlib import Path

csv_path = Path(r"c:\Users\user\Downloads\検定日速報_1219102_202511.csv")

# Shift-JISエンコーディングで読み込む
lines = []
with open(csv_path, 'r', encoding='shift_jis') as f:
    for line in f:
        lines.append(line.strip().split(','))

# DataFrameに変換
max_cols = max(len(row) for row in lines) if lines else 0
for row in lines:
    while len(row) < max_cols:
        row.append('')

df = pd.DataFrame(lines, dtype=str)

# ヘッダー行を取得（7行目、0列目から）
header_row_idx = 7
if len(df) > header_row_idx:
    headers = df.iloc[header_row_idx].tolist()
    
    print("分娩関連の列を検索:")
    for idx, header in enumerate(headers):
        if pd.notna(header):
            header_str = str(header).strip()
            if '分娩' in header_str or '最終' in header_str:
                print(f"  列{idx}: {header_str}")
    
    print("\nデータ行（0980）の確認:")
    if len(df) > header_row_idx + 1:
        row = df.iloc[header_row_idx + 1].tolist()  # 9行目（データ行）
        print(f"  個体識別番号: {row[0]}")
        print(f"  拡大4桁ID: {row[1]}")
        
        # 分娩関連の列の値を確認
        for idx, header in enumerate(headers):
            if pd.notna(header):
                header_str = str(header).strip()
                if '分娩' in header_str or '最終' in header_str:
                    val = row[idx] if len(row) > idx else 'N/A'
                    print(f"  列{idx} ({header_str}): {val}")




