import pandas as pd
import numpy as np
import os
import time
from datetime import datetime

# --- CONFIGURATION: EASY TO ADJUST --- #
FILENAME = "CTX10.csv"
WEEK_LOOKBACK = 8
PRICE_SOURCE_FILTER = ["CPA"]

# --- Paths --- #
BASE_DIR = "/mnt/c/Users/kmor6669/OneDrive - Sysco Corporation/Desktop/Project/FY2026 plan/ML Project"
source_file = os.path.join(BASE_DIR, FILENAME)
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_path = os.path.join(BASE_DIR, f"{FILENAME.split('.')[0]}_zone_behavior_output_{timestamp}.csv")

# --- START PIPELINE --- #
print("📂 Loading data...")
pipeline_start = time.time()
df = pd.read_csv(source_file, dtype=str)
df.columns = df.columns.str.strip()
df['Customer Key'] = df['Company Customer Number'].astype(str) + '|' + df['Attribute Group ID'].astype(str)


# Fallback: If Price Zone ID is blank or NaN, use Company Number + "-0"
price_zone_fixed = df['Price Zone ID'].copy()
missing_mask = price_zone_fixed.isna() | (price_zone_fixed.astype(str).str.strip() == '')
price_zone_fixed = price_zone_fixed.where(~missing_mask, df['Company Number'] + '-0')

# Extract the 1-digit suffix
df['Zone Suffix'] = (
    price_zone_fixed
    .astype(str)
    .str.extract(r'-(\d)$')[0]
    .fillna('0')
    .astype(float)
)

# --- Convert Fields --- #
print("🔄 Converting types...")
df['Fiscal Week Number'] = pd.to_numeric(df['Fiscal Week Number'], errors='coerce')
df['Fiscal Year Key'] = pd.to_numeric(df['Fiscal Year Key'], errors='coerce')
df['Pounds CY'] = pd.to_numeric(df['Pounds CY'], errors='coerce')
df['Computer Margin $ Per LB PY'] = pd.to_numeric(df['Computer Margin $ Per LB PY'], errors='coerce')
df['Computer Margin $ Per LB CY'] = pd.to_numeric(df['Computer Margin $ Per LB CY'], errors='coerce')
df['Delta Pounds YoY'] = pd.to_numeric(df['Delta Pounds YoY'], errors='coerce')

if PRICE_SOURCE_FILTER:
    df = df[df['Price Source Type'].isin(PRICE_SOURCE_FILTER)]

# --- Create composite key --- #
df['YearWeek'] = df['Fiscal Year Key'] * 100 + df['Fiscal Week Number']
df = df.sort_values(by=['Company Customer Number', 'Attribute Group ID', 'YearWeek'])

# --- Detect Zone Change --- #
print("📈 Detecting zone changes...")
df['Prev Zone'] = df.groupby(['Company Customer Number', 'Attribute Group ID'])['Zone Suffix'].shift()
df['Zone Change'] = df['Zone Suffix'] != df['Prev Zone']
df['Direction'] = df.apply(
    lambda x: 'Up' if pd.notnull(x['Zone Suffix']) and pd.notnull(x['Prev Zone']) and x['Zone Suffix'] > x['Prev Zone']
    else ('Down' if pd.notnull(x['Zone Suffix']) and pd.notnull(x['Prev Zone']) and x['Zone Suffix'] < x['Prev Zone']
    else 'None'),
    axis=1
)

# --- Track Timing --- #
drop_week = df[df['Delta Pounds YoY'] < 0].groupby(['Company Customer Number', 'Attribute Group ID'])['YearWeek'].min().rename('Drop Week')
change_week = df[df['Zone Change']].groupby(['Company Customer Number', 'Attribute Group ID'])['YearWeek'].min().rename('Change Week')
timing = pd.concat([drop_week, change_week], axis=1).reset_index()

def classify_change(row):
    if pd.isna(row['Change Week']):
        return 'No Change'
    elif pd.isna(row['Drop Week']):
        return 'No Drop'
    elif row['Change Week'] < row['Drop Week']:
        return 'Proactive'
    elif row['Change Week'] > row['Drop Week']:
        return 'Reactive'
    else:
        return 'Concurrent'

timing['Change Type'] = timing.apply(classify_change, axis=1)

# --- Summary Function --- #
def summarize_change(group):
    group = group.dropna(subset=['Zone Suffix'])

    if group['Zone Suffix'].nunique() <= 1:
        print("🔕 Skipping group — only one zone seen")
        return None

    change_rows = group[group['Zone Change']]
    if change_rows.empty:
        print("🔕 Skipping group — no zone change")
        return None

    # If it gets here, we are keeping the group:
    print("✅ Keeping group:", group[['Company Customer Number', 'Attribute Group ID']].iloc[0].to_dict())

    change_row = change_rows.iloc[0]
    week_of_change = change_row['YearWeek']
    before = group[group['YearWeek'] < week_of_change]
    after = group[group['YearWeek'] >= week_of_change]

    return pd.Series({
        'Customer': change_row['Company Customer Number'],
        'Week of Change': week_of_change,
        'From Zone': change_row['Prev Zone'],
        'To Zone': change_row['Zone Suffix'],
        'Direction': change_row['Direction'],
        'Volume Before': before['Pounds CY'].sum(),
        'Volume After': after['Pounds CY'].sum(),
        'Delta Pounds': after['Pounds CY'].sum() - before['Pounds CY'].sum(),
        'Margin Before': before['Computer Margin $ Per LB PY'].mean(),
        'Margin After': after['Computer Margin $ Per LB CY'].mean(),
        'Customer Key': change_row['Company Customer Number'] + '|' + change_row['Attribute Group ID'],
        'Delta Margin Rate': after['Computer Margin $ Per LB CY'].mean() - before['Computer Margin $ Per LB PY'].mean()
    })

print("📊 Summarizing behavior...")
# Print diagnostics BEFORE the summary is created
print("🧪 Total valid groups to scan:", df.groupby(['Company Customer Number', 'Attribute Group ID']).ngroups)
print("🧪 Total rows flagged as zone change:", df['Zone Change'].sum())
# 🧪 Apply summarize_change per customer-AG group
summary = (
    df.groupby(['Company Customer Number', 'Attribute Group ID'], group_keys=False)
      .apply(summarize_change)
      .dropna(subset=['Customer'])  # or dropna(how="all") if needed
      .reset_index(drop=True)
)


# 🔍 Check output before merging
print("✅ Summary shape:", summary.shape)
print(summary.head(5))

# 🔄 Rename group index columns to avoid reset_index() conflict
summary.index.names = ['Company Customer Number', 'Attribute Group ID']
summary = summary.reset_index()

# --- Merge customer/cuisine/meta info --- #
print("📎 Merging customer info...")
summary = summary.merge(
    df.groupby(['Company Customer Number', 'Attribute Group ID']).agg({
        'NPD DMA': 'last',
        'NPD Cuisine Type': 'last',
        'Company Number': 'last'
    }).reset_index(),
    on=['Company Customer Number', 'Attribute Group ID'],
    how='left'
)

df['Zone Originally Missing'] = df['Price Zone ID'].isna() | (df['Price Zone ID'].astype(str).str.strip() == '')
df['Normalized Price Zone ID'] = df['Company Number'] + '-' + df['Zone Suffix'].astype(int).astype(str)


summary = summary.merge(timing, on=['Company Customer Number', 'Attribute Group ID'], how='left')

# --- Add recency + win-back flag --- #
print("🕒 Calculating recency and win-back flag...")
latest_yearweek = df['YearWeek'].max()
last_ag_purchase = df[df['Pounds CY'] > 0].groupby(['Company Customer Number', 'Attribute Group ID'])['YearWeek'].max().rename('Last AG Purchase Week')
summary = summary.merge(last_ag_purchase, on=['Company Customer Number', 'Attribute Group ID'], how='left')
summary['Weeks Since Last AG Purchase'] = latest_yearweek - summary['Last AG Purchase Week']

summary['Winback Flag'] = (
    (summary['Volume After'] == 0) &
    (summary['Volume Before'] > 0) &
    (summary['Weeks Since Last AG Purchase'] <= WEEK_LOOKBACK) &
    (summary['Direction'] == 'Up')
)

# --- Priority Scoring --- #
summary['Priority Score'] = (summary['Delta Pounds'].abs() * summary['Margin Before']).fillna(0)

# --- Save and Done --- #
summary.to_csv(output_path, index=False)
print(f"✅ Saved: {output_path}")
print(f"🕒 Runtime: {time.time() - pipeline_start:.2f} sec")
