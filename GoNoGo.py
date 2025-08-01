# ✅ Go / No-Go Statistical Validator for Final Combo Key (With CPA Filter Pass)
import pandas as pd
import numpy as np
import os
from scipy.stats import chi2_contingency, entropy
from datetime import datetime
from colorama import Fore, init

init(autoreset=True)
# --- USER CONFIG ---
filename = "All550.csv"
include_business_center = False
include_region_breakdown = True
# --- ACTIVE CUSTOMER CONFIG ---
active_window_weeks = 8  # ⬅️ Default to 8 weeks unless changed
cutoff_week = None

# --- FILE PATH ---
source_folder = "/mnt/c/Users/kmor6669/OneDrive - Sysco Corporation/Desktop/Project/FY2026 plan/ML Project"
source_file = os.path.join(source_folder, filename)
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
base_name = os.path.splitext(filename)[0]

# --- LOG FILE SETUP ---
log_filename = f"{base_name}_analysis_{timestamp}.txt"
log_file_path = os.path.join(source_folder, log_filename)

# Save all print output to this file too
class DualWriter:
    def __init__(self, *files):
        self.files = files
    def write(self, text):
        for f in self.files:
            f.write(text)
    def flush(self):
        for f in self.files:
            try:
                f.flush()
            except ValueError:
                pass
    def close(self):
        for f in self.files:
            try:
                f.close()
            except Exception:
                pass

log_file = open(log_file_path, "w", encoding="utf-8")
import sys
sys.stdout = DualWriter(sys.stdout, log_file)

# --- LOAD + CLEAN ---
print(f"📂 Loading {filename}")
df = pd.read_csv(source_file, low_memory=False, dtype=str)
df.columns = df.columns.str.strip()

# Extract Zone and Company Number
df[['Company Prefix', 'Zone Suffix']] = df['Price Zone ID'].astype(str).str.split('-', n=1, expand=True)
df['Company Number'] = df['Company Prefix']

# Convert values
df['Delta Pounds YoY'] = pd.to_numeric(df['Delta Pounds YoY'], errors='coerce')
df = df[df['Delta Pounds YoY'].notna()]
df['YoY Change'] = (df['Delta Pounds YoY'] > 0).astype(int)

# Use final Combo Key
combo_parts = [
    "Price Source Type",
    "Zone Suffix",
    "NPD Cuisine Type",
    "Attribute Group ID"
]
if include_business_center:
    combo_parts.insert(0, "Business Center ID")

df['Combo Key'] = df[combo_parts].astype(str).agg('|'.join, axis=1)


# --- CHI-SQUARED + CRAMER'S V ---
def cramers_v(conf_matrix):
    chi2 = chi2_contingency(conf_matrix)[0]
    n = conf_matrix.sum().sum()
    phi2 = chi2 / n
    r, k = conf_matrix.shape
    return np.sqrt(phi2 / min(k - 1, r - 1)) if min(k - 1, r - 1) > 0 else 0

def run_analysis(df_subset, label="FULL DATASET"):
    print(f"\n🔹 Running Analysis: {label}")

    # Chi-squared + Cramer's V
    contingency = pd.crosstab(df_subset['Combo Key'], df_subset['YoY Change'])
    if contingency.shape[0] > 1 and contingency.shape[1] == 2:
        chi2, p, dof, _ = chi2_contingency(contingency)
        v = cramers_v(contingency)
        print(f"\n📈 What We Found")
        print(f"• Chi-squared = {chi2:.2f}")
        print(f"• p-value = {p:.6f}")
        print(f"• Cramér’s V = {v:.4f}")
        print("✅ Means the pattern is real and statistically valid" if p < 0.05 else "⛔ Pattern is likely random")
    else:
        v, p = 0, 1
        print("🚫 Not enough data to run test.")
        return

    # Entropy
    print(f"\n🔍 Do groups behave predictably?")
    combo_dist = pd.crosstab(df_subset['Combo Key'], df_subset['YoY Change'], normalize='index')
    row_entropy = combo_dist.apply(lambda x: entropy(x, base=2), axis=1)
    avg_entropy = row_entropy.mean()
    print(f"• Avg Entropy = {avg_entropy:.4f}")
    if avg_entropy > 0.7:
        print("⚠️ High entropy — some group behavior may be inconsistent")
    else:
        print("✅ Lower entropy = customer groups behave consistently")

    # Lift
    print(f"\n📊 Performance spread across groups (Lift)")
    global_rate = df_subset['YoY Change'].mean()
    lift_table = df_subset.groupby('Combo Key')['YoY Change'].mean().div(global_rate)
    lift_summary = lift_table.describe()
    print(f"• Global YoY Positive Rate = {global_rate:.4f}")
    print(f"• Lift Summary:\n{lift_summary}")
    print("✅ Some groups are doing way better/worse than average")

    # Gini
    print(f"\n📉 Volume loss concentration (Gini)")
    gini_score = gini(df_subset.groupby('Combo Key')['Delta Pounds YoY'].sum().abs())
    print(f"• Gini = {gini_score:.4f}")
    print("✅ High Gini = Most loss is coming from a small # of segments")

    # Final call
    print(f"\n🧭 Go / No-Go Decision:")
    if p < 0.05 and v >= 0.15:
        print(Fore.GREEN + "✅ GO: Statistically significant AND structured signal.")
    elif p < 0.05:
        print(Fore.YELLOW + "⚠️ Statistically significant, but weak signal.")
    else:
        print(Fore.RED + "⛔️ NO GO: Not statistically significant.")
    print("✅ Sub-analysis complete.")

# --- GINI FUNCTION ---
def gini(x):
    x = np.sort(np.array(x))
    n = len(x)
    cumx = np.cumsum(x)
    return (2. * np.sum((np.arange(1, n+1) * x))) / (n * cumx[-1]) - (n + 1.) / n

# --- FULL DATASET ANALYSIS ---
run_analysis(df, label="FULL DATASET")

# --- CPA-ONLY SECOND PASS ---
df_cpa = df[df['Price Source Type'] == 'CPA']
run_analysis(df_cpa, label="CPA ONLY")

# --- REGION-LEVEL PRIORITY SCORING (CPA ONLY) ---
if include_region_breakdown:
    print("\n📍 Regional Signal Strength + Volume, Loss & Active Impact (CPA ONLY)")

    if all(col in df_cpa.columns for col in ['Company Region Name', 'Pounds CY', 'Delta Pounds YoY', 'Fiscal Week ID', 'Company Customer Number']):
        df_cpa['Pounds CY'] = pd.to_numeric(df_cpa['Pounds CY'], errors='coerce')
        df_cpa['Delta Pounds YoY'] = pd.to_numeric(df_cpa['Delta Pounds YoY'], errors='coerce')
        df_cpa['Fiscal Week ID'] = pd.to_numeric(df_cpa['Fiscal Week ID'], errors='coerce')
        df_cpa = df_cpa[df_cpa[['Pounds CY', 'Delta Pounds YoY', 'Fiscal Week ID']].notna().all(axis=1)]

        total_cpa_volume = df_cpa['Pounds CY'].sum()
        total_loss_volume = df_cpa[df_cpa['Delta Pounds YoY'] < 0]['Delta Pounds YoY'].abs().sum()

        max_week = df_cpa['Fiscal Week ID'].max()
        cutoff_week = max_week - active_window_weeks

        region_scores = []

        for region, sub in df_cpa.groupby('Company Region Name'):
            region_volume = sub['Pounds CY'].sum()
            region_loss = sub[sub['Delta Pounds YoY'] < 0]['Delta Pounds YoY'].abs().sum()
            volume_pct = region_volume / total_cpa_volume if total_cpa_volume > 0 else 0
            loss_pct = region_loss / total_loss_volume if total_loss_volume > 0 else 0

            all_customers = sub['Company Customer Number'].nunique()
            active_customers = sub[sub['Fiscal Week ID'] >= cutoff_week]['Company Customer Number'].nunique()
            active_pct = active_customers / all_customers if all_customers > 0 else 0

            ct = pd.crosstab(sub['Combo Key'], sub['YoY Change'])
            if ct.shape[0] > 1 and ct.shape[1] == 2:
                try:
                    v = cramers_v(ct)
                    if not np.isnan(v):
                        priority_v_all = v * volume_pct * loss_pct * active_pct
                        region_scores.append({
                            'Region': region,
                            "Cramér’s V": round(v, 4),
                            "% of CPA Volume": round(volume_pct * 100, 2),
                            "% of Loss Volume": round(loss_pct * 100, 2),
                            "% Active Customers": round(active_pct * 100, 2),
                            "Priority Score (V×Vol×Loss×Active)": round(priority_v_all, 6)
                        })
                except:
                    continue

        region_df = pd.DataFrame(region_scores).sort_values("Priority Score (V×Vol×Loss×Active)", ascending=False)

        if not region_df.empty:
            print("\n🏆 Ranked Regions by Expanded Priority Score:")
            print(region_df.to_string(index=False))
            region_df.to_csv(os.path.join(source_folder, f"{base_name}_region_priority_scores.csv"), index=False)
            print(f"\n📁 Exported to: {base_name}_region_priority_scores.csv")
        else:
            print("🚫 No usable region data.")
    else:
        print("ℹ️ Region analysis was skipped due to missing columns.")

# --- FINAL: BEST BUSINESS CENTER + ATTRIBUTE GROUP COMBO  ---
if include_business_center and all(col in df_cpa.columns for col in ['Business Center ID', 'Attribute Group ID', 'Pounds CY', 'Delta Pounds YoY', 'Fiscal Week ID', 'Company Customer Number']):
    df_cpa['Pounds CY'] = pd.to_numeric(df_cpa['Pounds CY'], errors='coerce')
    df_cpa['Delta Pounds YoY'] = pd.to_numeric(df_cpa['Delta Pounds YoY'], errors='coerce')
    df_cpa['Fiscal Week ID'] = pd.to_numeric(df_cpa['Fiscal Week ID'], errors='coerce')
    df_cpa = df_cpa[df_cpa[['Pounds CY', 'Delta Pounds YoY', 'Fiscal Week ID']].notna().all(axis=1)]

    df_cpa['BC|AG'] = df_cpa['Business Center ID'].astype(str) + "|" + df_cpa['Attribute Group ID'].astype(str)

    total_cpa_volume = df_cpa['Pounds CY'].sum()
    total_loss_volume = df_cpa[df_cpa['Delta Pounds YoY'] < 0]['Delta Pounds YoY'].abs().sum()

    max_week = df_cpa['Fiscal Week ID'].max()
    cutoff_week = max_week - active_window_weeks

    bcag_results = []

    for key, group in df_cpa.groupby('BC|AG'):
        if group['YoY Change'].nunique() < 2:
            continue
        try:
            ct = pd.crosstab(group['BC|AG'], group['YoY Change'])
            v = cramers_v(ct)
            if pd.notna(v):
                group_volume = group['Pounds CY'].sum()
                group_loss = group[group['Delta Pounds YoY'] < 0]['Delta Pounds YoY'].abs().sum()
                volume_pct = group_volume / total_cpa_volume if total_cpa_volume > 0 else 0
                loss_pct = group_loss / total_loss_volume if total_loss_volume > 0 else 0

                all_customers = group['Company Customer Number'].nunique()
                active_customers = group[group['Fiscal Week ID'] >= cutoff_week]['Company Customer Number'].nunique()
                active_pct = active_customers / all_customers if all_customers > 0 else 0

                priority_v_all = v * volume_pct * loss_pct * active_pct

                bcag_results.append({
                    'BC|AG': key,
                    "Cramér’s V": round(v, 4),
                    "% of CPA Volume": round(volume_pct * 100, 2),
                    "% of Loss Volume": round(loss_pct * 100, 2),
                    "% Active Customers": round(active_pct * 100, 2),
                    "Priority Score (V×Vol×Loss×Active)": round(priority_v_all, 6)
                })
        except:
            continue

    bcag_df = pd.DataFrame(bcag_results).sort_values("Priority Score (V×Vol×Loss×Active)", ascending=False)

    if not bcag_df.empty:
        print("\n🏆 Top Business Center + AG Combos by Expanded Priority Score:")
        print(bcag_df.head(10).to_string(index=False))
        bcag_df.to_csv(os.path.join(source_folder, f"{base_name}_bcag_priority_scores.csv"), index=False)
        print(f"\n📁 Exported to: {base_name}_bcag_priority_scores.csv")
    else:
        print("🚫 No strong BC|AG signals found.")
else:
    print("ℹ️ Business Center + AG analysis was skipped (toggle is OFF or missing columns).")

if 'region_df' not in locals():
    region_df = pd.DataFrame()
if 'bcag_df' not in locals():
    bcag_df = pd.DataFrame()

print(f"✅ {len(region_df)} regions analyzed.")
print(f"✅ {len(bcag_df)} Business Center + Attribute Group combos analyzed.")
print("\n🎯 All analyses complete.")
log_file.close()

