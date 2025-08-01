
# ✅ Go/No-Go Combo Key Validator with Chunked Loading for Large Files

import pandas as pd
import numpy as np
import os
from itertools import combinations
from scipy.stats import chi2_contingency
from datetime import datetime
from colorama import Fore, init
init(autoreset=True)

# --- USER CONFIGURATION --- #
source_filename = "CD_Mexican.csv"
include_business_center = True
base_folder = "/mnt/c/Users/kmor6669/OneDrive - Sysco Corporation/Desktop/Project/FY2026 plan/Stats Tests"
source_file = os.path.join(base_folder, source_filename)

# --- COLUMNS --- #
available_categorical_columns = [
    'Zone Suffix', 'Company Number', 'Attribute Group ID', 'Price Source Type',
    'NPD Cuisine Type', 'NPD DMA', 'Company Region ID', 'Order Channel',
    'Sysco Brand Indicator', 'Customer State', 'Customer Zip Code', 'Business Center ID'
]

default_included = [
    'Zone Suffix', 'Attribute Group ID', 'Price Source Type', 'NPD Cuisine Type'
]
exclude_columns = ['Customer Zip Code']

include_columns = default_included.copy()
if include_business_center:
    include_columns.append('Business Center ID')
active_columns = [col for col in include_columns if col not in exclude_columns]

# --- LOAD DATA IN CHUNKS --- #
print(f"📂 Loading data in chunks from: {source_file}")
chunk_list = []
for chunk in pd.read_csv(source_file, chunksize=500_000, dtype=str, engine='python', on_bad_lines='skip'):
    chunk_list.append(chunk)
df = pd.concat(chunk_list, ignore_index=True)
print(f"✅ Loaded {len(df):,} rows.")

df.columns = df.columns.str.strip()
df[['Company Prefix', 'Zone Suffix']] = df['Price Zone ID'].astype(str).str.split('-', n=1, expand=True)
df['Company Number'] = df['Company Prefix']

df['Delta Pounds YoY'] = pd.to_numeric(df['Delta Pounds YoY'], errors='coerce')
df = df[df['Delta Pounds YoY'].notna()]
df['YoY Change'] = (df['Delta Pounds YoY'] > 0).astype(int)
df['Combo Key'] = df[active_columns].astype(str).agg('|'.join, axis=1)

# --- CRAMER'S V --- #
def cramers_v_from_df(col1, col2):
    confusion_matrix = pd.crosstab(col1, col2)
    if confusion_matrix.shape[0] <= 1 or confusion_matrix.shape[1] <= 1:
        return np.nan
    chi2 = chi2_contingency(confusion_matrix)[0]
    n = confusion_matrix.sum().sum()
    phi2 = chi2 / n
    r, k = confusion_matrix.shape
    return np.sqrt(phi2 / min(k - 1, r - 1)) if min(k - 1, r - 1) > 0 else 0

print("\n🔹 Global Chi-squared test for full Combo Key:")
contingency = pd.crosstab(df['Combo Key'], df['YoY Change'])
if contingency.shape[0] > 1 and contingency.shape[1] == 2:
    chi2, p, dof, _ = chi2_contingency(contingency)
    cramers_v_stat = cramers_v_from_df(df['Combo Key'], df['YoY Change'])
    print(f"📊 Chi-squared = {chi2:.2f}, p = {p:.6f}, Cramér’s V = {cramers_v_stat:.4f}")
    if p < 0.05:
        print("✅ Statistically significant pattern found.")
    else:
        print("❌ Not statistically significant.")
else:
    print("🚫 Not enough variation to test global Combo Key.")

# --- COMBO SCAN 2–4-WAY --- #
combo_results = []
max_combo_size = 4
print(f"\n🔍 Scanning 2–{max_combo_size} variable combinations from: {active_columns}")

for r in range(2, max_combo_size + 1):
    for combo in combinations(active_columns, r):
        combo_name = "+".join(combo)
        temp_key = df[list(combo)].astype(str).agg('|'.join, axis=1)
        v = cramers_v_from_df(temp_key, df['YoY Change'])
        combo_results.append({'Combo': combo_name, "Cramér’s V": round(v, 4)})

combo_df = pd.DataFrame(combo_results).dropna().sort_values("Cramér’s V", ascending=False)
top_combos = combo_df.head(15)

print("\n📊 Top Combo Key Signals by Cramér’s V:")
for _, row in top_combos.iterrows():
    combo = row['Combo']
    v = row["Cramér’s V"]
    if v >= 0.25:
        color = Fore.GREEN
    elif v >= 0.15:
        color = Fore.YELLOW
    else:
        color = Fore.RED
    print(f"{color}{combo:<50}  Cramér’s V = {v:.4f}")

# --- EXPORT RESULTS --- #
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
base_name = os.path.splitext(source_filename)[0]
combo_df.to_csv(f"{base_name}_combo_signals_{timestamp}.csv", index=False)
print(f"\n✅ Full combo scan results saved to: {base_name}_combo_signals_{timestamp}.csv")
print("✅ Script complete.")
