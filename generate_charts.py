#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
from pathlib import Path

# 日本語フォント設定
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial Unicode MS', 'SimHei']
try:
    # Linux環境での日本語フォント
    fm.fontManager.addfont('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc')
    plt.rcParams['font.sans-serif'] = ['Noto Sans CJK JP']
except:
    pass

plt.rcParams['axes.unicode_minus'] = False

# 1. 売上シナリオの棒グラフ
fig1, ax1 = plt.subplots(figsize=(10, 6))

scenarios = ['保守', '中央値', 'モデル店水準']
sales = [250, 390, 550]  # 単位：百万円
colors = ['#3498db', '#2ecc71', '#e74c3c']

bars = ax1.bar(scenarios, sales, color=colors, alpha=0.8, edgecolor='black', linewidth=1.5)

# データラベルの追加
for bar in bars:
    height = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2., height,
             f'{int(height)}百万円',
             ha='center', va='bottom', fontsize=12, fontweight='bold')

ax1.set_ylabel('売上 (百万円)', fontsize=12, fontweight='bold')
ax1.set_xlabel('シナリオ', fontsize=12, fontweight='bold')
ax1.set_title('売上シナリオ比較', fontsize=14, fontweight='bold', pad=20)
ax1.grid(axis='y', alpha=0.3, linestyle='--')
ax1.set_ylim(0, max(sales) * 1.15)

plt.tight_layout()
plt.savefig('sales_chart.png', dpi=300, bbox_inches='tight')
print("✓ sales_chart.png を生成しました")
plt.close()

# 2. 機種構成の円グラフ
fig2, ax2 = plt.subplots(figsize=(10, 8))

machines = ['輪っか', '2本爪', '3本ミニ', '3本', '特殊']
percentages = [44, 21, 21, 12, 2]
colors2 = ['#ff6b6b', '#4ecdc4', '#45b7d1', '#ffd93d', '#a8e6cf']

wedges, texts, autotexts = ax2.pie(percentages, labels=machines, autopct='%1.1f%%',
                                     colors=colors2, startangle=90,
                                     textprops={'fontsize': 11, 'fontweight': 'bold'},
                                     explode=[0.05, 0, 0, 0, 0])

# パーセンテージテキストのカスタマイズ
for autotext in autotexts:
    autotext.set_color('white')
    autotext.set_fontsize(12)
    autotext.set_fontweight('bold')

ax2.set_title('機種構成', fontsize=14, fontweight='bold', pad=20)

plt.tight_layout()
plt.savefig('machine_chart.png', dpi=300, bbox_inches='tight')
print("✓ machine_chart.png を生成しました")
plt.close()

print("\n完了: 両方のチャートが生成されました")
print(f"  - sales_chart.png (売上シナリオ比較)")
print(f"  - machine_chart.png (機種構成)")
