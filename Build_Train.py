import pandas as pd
import numpy as np
import os
import joblib
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import classification_report
from datetime import datetime

# --- CONFIGURATION --- #
BASE_DIR = "/mnt/c/Users/kmor6669/OneDrive - Sysco Corporation/Desktop/Project/FY2026 plan/ML Project"
PRIMARY_FILE = "CD_Mexican.csv"                  # AG 550 data
AG_COUNT_FILE = "customer_ag_breadth.csv"        # AG counts by customer
MODEL_NAME = "customer_lift_engine.pkl"

# --- LOAD FILES --- #
print("📂 Loading dataset...")
primary_df = pd.read_csv(os.path.join(BASE_DIR, PRIMARY_FILE), dtype=str)
ag_counts = pd.read_csv(os.path.join(BASE_DIR, AG_COUNT_FILE), dtype=str)

# --- Merge AG count (XLOOKUP style) --- #
primary_df['Company Customer Number'] = primary_df['Company Customer Number'].astype(str)
ag_counts['Company Customer Number'] = ag_counts['Company Customer Number'].astype(str)

primary_df = primary_df.merge(ag_counts, on='Company Customer Number', how='left')
primary_df['Distinct AG Count'] = pd.to_numeric(primary_df['Distinct AG Count'], errors='coerce').fillna(0)

# --- Convert types --- #
cols_to_float = [
    'Pounds CY', 'Pounds PY', 'Computer Margin $ Per LB CY', 'Computer Margin $ Per LB PY',
    'Delta Pounds YoY', 'Delta Ref Price YoY', 'Delta GP Ext YoY', 'Delta GP Ext per Lb Chg'
]
for col in cols_to_float:
    primary_df[col] = pd.to_numeric(primary_df[col], errors='coerce')

# --- Price Zone Features --- #
primary_df = primary_df[primary_df['Price Zone ID'].notna()]
primary_df[['Company Prefix', 'Zone Suffix']] = primary_df['Price Zone ID'].astype(str).str.split('-', n=1, expand=True)
primary_df['Zone Suffix'] = pd.to_numeric(primary_df['Zone Suffix'], errors='coerce')
primary_df['Prev Zone'] = primary_df.groupby(['Company Customer Number'])['Zone Suffix'].shift()
primary_df['Direction'] = np.where(primary_df['Zone Suffix'] > primary_df['Prev Zone'], 'Up',
                          np.where(primary_df['Zone Suffix'] < primary_df['Prev Zone'], 'Down', 'None'))

primary_df['Dir_Up'] = (primary_df['Direction'] == 'Up').astype(int)
primary_df['Dir_Down'] = (primary_df['Direction'] == 'Down').astype(int)

# --- Target Variable --- #
primary_df['Pct_Change_Pounds'] = ((primary_df['Pounds CY'] - primary_df['Pounds PY']) / primary_df['Pounds PY']).replace([np.inf, -np.inf], 0)
primary_df['Target_Growth'] = (primary_df['Pct_Change_Pounds'] > 0).astype(int)

# --- Feature Selection --- #
features = [
    'Delta Pounds YoY', 'Delta Ref Price YoY', 'Delta GP Ext YoY',
    'Computer Margin $ Per LB PY', 'Computer Margin $ Per LB CY',
    'Distinct AG Count', 'Dir_Up', 'Dir_Down'
]

primary_df = primary_df.dropna(subset=features + ['Target_Growth'])
print(f"✅ Rows after filtering: {len(primary_df)}")

if len(primary_df) == 0:
    raise ValueError("❌ No valid rows left after filtering. Check feature coverage or file content.")

# --- Train/Test Split --- #
X = primary_df[features]
y = primary_df['Target_Growth']

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=42)

# --- Model Training --- #
clf = RandomForestClassifier(n_estimators=100, random_state=42)
clf.fit(X_train, y_train)

# --- Save Model + Features --- #
joblib.dump(clf, os.path.join(BASE_DIR, MODEL_NAME))
with open(os.path.join(BASE_DIR, "customer_lift_features.txt"), "w") as f:
    f.write("\n".join(features))

# --- Evaluate --- #
y_pred = clf.predict(X_test)
print("🔍 Classification Report:")
print(classification_report(y_test, y_pred))

print("✅ Model saved to", MODEL_NAME)
