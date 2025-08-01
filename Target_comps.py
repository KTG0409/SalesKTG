import pandas as pd
import numpy as np
import os

# --- CONFIGURATION --- #
BASE_DIR = "C:/Users/kmor6669/OneDrive - Sysco Corporation/Desktop/Project/FY2026 plan/Stats Tests"
ZONE_FILE = "CTX All_zone_behavior_output_latest.csv"

# --- Load --- #
zone_df = pd.read_csv(os.path.join(BASE_DIR, ZONE_FILE))
zone_df.columns = zone_df.columns.str.strip()

# --- Prep --- #
zone_df['Winback Flag'] = zone_df['Winback Flag'].astype(bool)
zone_df['Delta Pounds'] = pd.to_numeric(zone_df['Delta Pounds'], errors='coerce')
zone_df['Delta Margin Rate'] = pd.to_numeric(zone_df['Delta Margin Rate'], errors='coerce')

# --- Create Segment Flags --- #
zone_df['Was Changed'] = zone_df['From Zone'].notna() & zone_df['To Zone'].notna() & (zone_df['From Zone'] != zone_df['To Zone'])

zone_df['Group'] = np.select(
    [
        zone_df['Winback Flag'] & zone_df['Was Changed'],       # ✅ Should & Did
        zone_df['Winback Flag'] & ~zone_df['Was Changed'],      # ❌ Should but Didn't
        ~zone_df['Winback Flag'] & zone_df['Was Changed'],      # ⚠️ Didn't need but Did
        ~zone_df['Winback Flag'] & ~zone_df['Was Changed']      # 👍 Correctly Ignored
    ],
    [
        'Should & Did',
        'Should Have Changed',
        'Changed but Shouldn’t',
        'Left Alone Correctly'
    ],
    default='Unclassified'
)

# --- Aggregate Performance --- #
summary = (
    zone_df.groupby('Group')[['Delta Pounds', 'Delta Margin Rate']]
    .agg(['count', 'mean', 'sum'])
    .round(2)
    .reset_index()
)

# --- Save Output --- #
output_path = os.path.join(BASE_DIR, "CD_Mexican_feedback_matrix.csv")
summary.to_csv(output_path, index=False)
print("✅ Feedback matrix saved to:", output_path)

# --- Preview --- #
print("\n📊 Feedback Matrix Summary:\n")
print(summary)
