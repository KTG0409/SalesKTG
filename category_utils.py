# category_utils.py
import pandas as pd
import numpy as np

def clean_numeric_column(series):
    """Remove $, %, commas from numeric columns before conversion."""
    return pd.to_numeric(
        series.astype(str).str.replace("$", "").str.replace("%", "").str.replace(",", "").str.strip(),
        errors="coerce"
    ).fillna(0)

def _clean_str(x):
    """Clean string but preserve leading zeros."""
    if pd.isna(x): return ""
    return str(x).strip()

def _clean_str_numeric(x):
    """Clean string and strip leading zeros for numeric ID comparison only."""
    if pd.isna(x): return ""
    s = str(x).strip()
    if s.isdigit():
        return s.lstrip("0") or "0"
    return s

def concat_address(addr1, addr2):
    """Combine two address fields."""
    a1 = str(addr1).strip() if pd.notna(addr1) else ""
    a2 = str(addr2).strip() if pd.notna(addr2) else ""
    return (a1 + " " + a2).strip()
# ---------------------------- Core (your originals, kept) ----------------------------
def load_and_prepare(source_path, alignment_path, company=None, attr_groups=None, vendors=None):
    g = pd.read_csv(source_path, dtype=str)
    a = pd.read_csv(alignment_path, dtype=str)

    def _clean_headers(df):
        cols = df.columns.str.replace(r"\s+", " ", regex=True).str.strip()
        df.columns = cols
        return df

    g = _clean_headers(g).fillna("")
    a = _clean_headers(a).fillna("")

    # numerics used downstream
    for col in ["Pounds CY", "Pounds PY", "Fiscal Week Number"]:
        if col in g.columns:
            g[col] = clean_numeric_column(g[col])  

    # Clean but PRESERVE leading zeros for alignment keys
    for col in ["Company Number", "Lot Number", "Item Number", "True Vendor Number"]:
        if col in g.columns:
            g[col] = g[col].map(_clean_str)

    # Alignment key in alignment file
    if "Alignment Key" not in a.columns:
        raise SystemExit("Alignment file must contain 'Alignment Key'.")
    a["Alignment Key"] = a["Alignment Key"].astype(str).str.strip()

    # SUPC and SUVC preserve leading zeros
    for col in ["SUPC", "SUVC"]:
        if col not in a.columns:
            raise SystemExit(f"Alignment file missing required column '{col}'.")
        a[col] = a[col].map(_clean_str)

    # Build alignment key in source
    if "Alignment Key" in g.columns and g["Alignment Key"].astype(str).str.len().gt(0).any():
        g["Alignment Key"] = g["Alignment Key"].astype(str).str.strip()
    else:
        need = [c for c in ["Company Number", "Lot Number"] if c not in g.columns]
        if need:
            raise SystemExit(f"Source missing columns to build Alignment Key: {need}")
        g["Alignment Key"] = g["Company Number"].map(_clean_str) + g["Lot Number"].map(_clean_str)

    # Address concat for leads routing
    g["Customer Street Address"] = g.apply(
        lambda r: concat_address(r.get("Customer Address", ""), r.get("Customer Address 2", "")),
        axis=1
    )

    # Prepare alignment data
    align_keep = [c for c in [
        "Alignment Key", "Lot Number", "Lot Description",
        "Supplier Name", "SUPC", "SUVC",
        "Product Category", "Next events", "OPCO Name", "Award Volume Annualized" 
    ] if c in a.columns]
    for req in ["Alignment Key", "SUPC", "SUVC", "Supplier Name"]:
        if req not in align_keep:
            raise SystemExit(f"Alignment join cannot proceed; missing '{req}'.")

    align = a[align_keep].drop_duplicates("Alignment Key")
    
    # MERGE FIRST (before any filtering)
    print(f"\n🔗 Merging alignment data...")
    print(f"   Source rows: {len(g):,}")
    print(f"   Alignment keys: {len(align):,}")
    
    g = g.merge(align, on="Alignment Key", how="left", suffixes=("", "_ALN"))

    if "SUPC" not in g.columns or g["SUPC"].isna().all():
        print(f"\n❌ NO MATCHES FOUND!")
        print(f"Sample source keys: {g['Alignment Key'].head(5).tolist()}")
        print(f"Sample alignment keys: {align['Alignment Key'].head(5).tolist()}")
        raise SystemExit("No Alignment matches found (all SUPC null). Check keys between files.")

    matches = g["SUPC"].notna().sum()
    print(f"   ✅ Matched: {matches:,} records")

    # NOW FILTER by company (after merge, with smart comparison)
    if company:
        print(f"\n🎯 Applying company filter: '{company}'")
        before_filter = len(g)
        
        # Normalize for comparison (strip leading zeros)
        filter_normalized = _clean_str_numeric(company).lower()
        
        # Create temp normalized columns for comparison
        g["_CompanyNum_Norm"] = g["Company Number"].map(_clean_str_numeric).str.lower()
        g["_CompanyName_Norm"] = g.get("Company Name", "").astype(str).str.strip().str.lower()
        
        # Filter using normalized comparison
        g = g[
            (g["_CompanyNum_Norm"] == filter_normalized) |
            (g["_CompanyName_Norm"] == company.strip().lower())
        ].copy()
        
        # Drop temp columns
        g = g.drop(columns=["_CompanyNum_Norm", "_CompanyName_Norm"])
        
        print(f"   Filtered: {before_filter:,} → {len(g):,} rows")
        
        if len(g) == 0:
            available = pd.read_csv(source_path, dtype=str, nrows=1000)["Company Number"].unique()[:10]
            print(f"\n❌ No data found for company '{company}'")
            print(f"Available company numbers (sample): {available.tolist()}")
            raise SystemExit(f"Company filter '{company}' resulted in zero rows.")
    
    # Filter by attribute groups if specified
    if attr_groups and "Attribute Group ID" in g.columns:
        keep_ag = set(_clean_str(x) for x in attr_groups.split(","))
        before = len(g)
        g = g[g["Attribute Group ID"].map(_clean_str).isin(keep_ag)]
        print(f"   Attribute group filter: {before:,} → {len(g):,} rows")

    # Verify required columns
    for req in ["SUPC", "SUVC", "Supplier Name"]:
        if req not in g.columns:
            raise SystemExit(f"Post-join missing required column '{req}'.")

    # Calculate alignment flags (preserve leading zeros in comparison)
    g["IsAlignedItem"] = (g.get("Item Number", "").map(_clean_str) == g["SUPC"].map(_clean_str)).astype(int)
    g["IsAlignedVendor"] = (g.get("True Vendor Number", "").map(_clean_str) == g["SUVC"].map(_clean_str)).astype(int)

    return g, vendors

def compute_windows(df):
    # current week
    current_week = int(pd.to_numeric(df.get("Fiscal Week Number", 0), errors="coerce").fillna(0).max()) if len(df) else 0

    # Ensure the account-type column exists
    if "Customer Account Type Code" not in df.columns:
        df["Customer Account Type Code"] = "UNKNOWN"

    # REQUIRED grouping keys (DEDENT THIS - same level as the if statement above)
    key = [
        "Company Number","Company Name",
        "Customer Name","Company Customer Number",
        "Customer Account Type Code",
        "Customer DSM",        # ← ADD THIS
        "Customer Territory",
        "Lot Number","Lot Description",
        "Attribute Group ID","Attribute Group Name",
        "Business Center Name",
        "Alignment Key",
        "SUPC","SUVC","Supplier Name"
    ]
    
    missing = [c for c in key if c not in df.columns]
    if missing:
        raise SystemExit(f"compute_windows(): required columns missing: {missing}")

    df = df.copy()
    for col in ["Pounds CY","Pounds PY"]:
        if col not in df.columns: df[col] = 0.0
    # Aligned flags & derived lbs
    both_flag = ((df.get("IsAlignedItem", 0).astype(int)) & (df.get("IsAlignedVendor", 0).astype(int))).astype(int)
    df["ItemAligned_Lbs"] = df["Pounds CY"] * df.get("IsAlignedItem", 0)
    df["VendorAligned_Lbs"] = df["Pounds CY"] * df.get("IsAlignedVendor", 0)
    df["ItemVendorAligned_Lbs"] = df["Pounds CY"] * both_flag

    # YTD aggregates (now per-account-type)
    ytd = df.groupby(key, dropna=False).agg(
        Pounds_CY=("Pounds CY","sum"),
        Pounds_PY=("Pounds PY","sum"),
        ItemAligned_Lbs=("ItemAligned_Lbs","sum"),
        VendorAligned_Lbs=("VendorAligned_Lbs","sum"),
        ItemVendorAligned_Lbs=("ItemVendorAligned_Lbs","sum"),
    ).reset_index()
    ytd["Delta_YoY_Lbs"] = ytd["Pounds_CY"] - ytd["Pounds_PY"]
    ytd["YoY_Pct"] = ytd.apply(lambda r: (r["Delta_YoY_Lbs"] / r["Pounds_PY"]) if r["Pounds_PY"] else np.nan, axis=1)

    # Week-over-week (also per-account-type now)
    lw = df[df["Fiscal Week Number"] == current_week].groupby(key)["Pounds CY"].sum().rename("W_Lbs").reset_index() if current_week else pd.DataFrame(columns=key+["W_Lbs"])
    pw = df[df["Fiscal Week Number"] == current_week - 1].groupby(key)["Pounds CY"].sum().rename("Wm1_Lbs").reset_index() if current_week else pd.DataFrame(columns=key+["Wm1_Lbs"])
    wow = lw.merge(pw, on=key, how="outer") if len(lw) or len(pw) else pd.DataFrame(columns=key+["W_Lbs","Wm1_Lbs"])
    if "W_Lbs" not in wow.columns: wow["W_Lbs"] = 0.0
    if "Wm1_Lbs" not in wow.columns: wow["Wm1_Lbs"] = 0.0
    wow["WoW_Delta_Lbs"] = wow["W_Lbs"] - wow["Wm1_Lbs"]
    wow["WoW_Pct"] = wow.apply(lambda r: (r["WoW_Delta_Lbs"] / r["Wm1_Lbs"]) if r["Wm1_Lbs"] else np.nan, axis=1)

    # 4w vs prior 4w (also per-account-type)
    last4 = df[df["Fiscal Week Number"].between(max(current_week-3, 0), current_week, inclusive="both")] if current_week else df.iloc[0:0]
    prior4 = df[df["Fiscal Week Number"].between(max(current_week-7, 0), max(current_week-4, 0), inclusive="both")] if current_week else df.iloc[0:0]
    l4 = last4.groupby(key)["Pounds CY"].sum().rename("L4_Lbs").reset_index() if len(last4) else pd.DataFrame(columns=key+["L4_Lbs"])
    p4 = prior4.groupby(key)["Pounds CY"].sum().rename("P4_Lbs").reset_index() if len(prior4) else pd.DataFrame(columns=key+["P4_Lbs"])
    m4 = l4.merge(p4, on=key, how="outer") if len(l4) or len(p4) else pd.DataFrame(columns=key+["L4_Lbs","P4_Lbs"])
    if "L4_Lbs" not in m4.columns: m4["L4_Lbs"] = 0.0
    if "P4_Lbs" not in m4.columns: m4["P4_Lbs"] = 0.0
    m4["L4_vs_P4_Delta"] = m4["L4_Lbs"] - m4["P4_Lbs"]
    m4["L4_vs_P4_Pct"] = m4.apply(lambda r: (r["L4_vs_P4_Delta"] / r["P4_Lbs"]) if r["P4_Lbs"] else np.nan, axis=1)

    status = ytd.merge(wow, on=key, how="left").merge(m4, on=key, how="left").fillna(0)
    return status, current_week
