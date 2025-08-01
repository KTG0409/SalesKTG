# ✅ Price Zone Impact Validator (Macro, Micro, Hybrid with Logging)

import pandas as pd
import numpy as np
from scipy.stats import kruskal, ttest_rel
from datetime import datetime
import os
import sys

# --- CONFIG ---
BASE_DIR = "/mnt/c/Users/kmor6669/OneDrive - Sysco Corporation/Desktop/Project/FY2026 plan/ML Project"
FILENAME = "CTX1_32.csv"
ZONE_COL = "Zone Suffix"
MARGIN_COL = "Computer Margin $ per LB"
ITEM_COL = "Item Number"
CUSTOMER_COL = "Company Customer Number"
PRICE_TYPE_COL = "Price Source Type"

# --- LOG FILE SETUP ---
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
base_name = os.path.splitext(FILENAME)[0]
log_filename = f"{base_name}_zone_impact_{timestamp}.txt"
log_file_path = os.path.join(BASE_DIR, log_filename)


class DualWriter:
    def __init__(self, *files):
        self.files = files
    def write(self, text):
        for f in self.files:
            f.write(text)
    def flush(self):
        for f in self.files:
            try: f.flush()
            except: pass
    def close(self):
        for f in self.files:
            try: f.close()
            except: pass

log_file = open(log_file_path, "w", encoding="utf-8")
sys.stdout = DualWriter(sys.stdout, log_file)

# --- LOAD + CLEAN ---
print(f"📂 Loading {FILENAME}...\n")
df = pd.read_csv(os.path.join(BASE_DIR, FILENAME), dtype=str)
df[MARGIN_COL] = pd.to_numeric(df[MARGIN_COL], errors="coerce")
df = df[df[MARGIN_COL].notna()]
df = df[df[PRICE_TYPE_COL] == "CPA"]
df = df[df[ZONE_COL].notna() & df[ITEM_COL].notna() & df[CUSTOMER_COL].notna()]

# Extract Zone and Company Number
df[['Company Prefix', 'Zone Suffix']] = df['Price Zone ID'].astype(str).str.split('-', n=1, expand=True)
df['Company Number'] = df['Company Prefix']

# --- MACRO ANALYSIS ---
print("🔍 STEP 1: MACRO — Margin Differences by Price Zone")
zone_groups = [g[MARGIN_COL].values for _, g in df.groupby(ZONE_COL) if len(g) > 10]
if len(zone_groups) >= 2:
    h_stat, p_val = kruskal(*zone_groups)
    print(f"• Kruskal-Wallis H = {h_stat:.2f}, p = {p_val:.5f}")
    print("✅ Significant margin differences by zone\n" if p_val < 0.05 else "⛔ No significant differences\n")
else:
    print("🚫 Not enough zone data for macro test\n")

# --- MICRO ANALYSIS ---
print("🔬 STEP 2: MICRO — Do Items Track Differently Across Zones?")
item_results = []
for item_id, group in df.groupby(ITEM_COL):
    zone_samples = [z[MARGIN_COL].values for _, z in group.groupby(ZONE_COL)]
    if len(zone_samples) >= 2:
        try:
            h, p = kruskal(*zone_samples)
            item_results.append((item_id, p))
        except:
            continue

sig_items = [r for r in item_results if r[1] < 0.05]
print(f"• Tested {len(item_results)} items")
print(f"✅ {len(sig_items)} items show statistically significant margin differences by zone\n")

# --- HYBRID ANALYSIS ---
print("🧪 STEP 3: HYBRID — Customer+Item Zone Shift Impact")
pivot = df.pivot_table(index=[CUSTOMER_COL, ITEM_COL], columns=ZONE_COL, values=MARGIN_COL, aggfunc="mean")
zone_pairs = [('2', '3'), ('1', '2'), ('3', '4')]

for z1, z2 in zone_pairs:
    if z1 in pivot.columns and z2 in pivot.columns:
        paired = pivot[[z1, z2]].dropna()
        if len(paired) >= 5:
            t_stat, p_val = ttest_rel(paired[z1], paired[z2])
            print(f"• Zone {z1} → {z2}: n={len(paired)} | t={t_stat:.2f}, p={p_val:.5f}")
            print("✅ Margin shift is significant\n" if p_val < 0.05 else "⛔ No significant margin change\n")
        else:
            print(f"🚫 Not enough matched samples for Zone {z1} → {z2}\n")

print("🎯 DONE: Zone Impact Validation Complete")
log_file.close()
