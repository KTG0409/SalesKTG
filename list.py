import pandas as pd
import os

# --- Step 1: Set filename and path ---
filename = "CTX AGs.csv"  # 🔁 Replace with your file name
folder = "/mnt/c/Users/kmor6669/OneDrive - Sysco Corporation/Desktop/Project/FY2026 plan/ML Project" # 🔁 Update to your actual path
filepath = os.path.join(folder, filename)

# --- Step 2: Load file (sample 1000 rows for performance) ---
df = pd.read_csv(filepath, dtype=str, nrows=1000)  # Read as string for safe type inspection

# --- Step 3: Build Column Inspection Table ---
info = pd.DataFrame({
    'Column Name': df.columns,
    'Data Type (in CSV)': [df[col].dropna().apply(type).mode()[0].__name__ if not df[col].dropna().empty else 'unknown' for col in df.columns],
    'Non-Null Count': df.notnull().sum().values,
    'Sample Value': [df[col].dropna().iloc[0] if df[col].dropna().any() else "" for col in df.columns]
})

# --- Step 4: Print ---
print(info.to_string(index=False))

# --- Step 5: Optional Save ---
save = True  # 🔁 Set to False if you don’t want to save
if save:
    out_path = os.path.join(folder, f"{os.path.splitext(filename)[0]}_column_info.csv")
    info.to_csv(out_path, index=False)
    print(f"\n✅ Saved column data types to: {out_path}")
