#!/usr/bin/env python3
# category_site_benchmark.py
# Single-site performance vs company benchmark

import argparse
import pandas as pd
import numpy as np
from datetime import datetime
from category_utils import clean_numeric_column, load_and_prepare, compute_windows

def calculate_site_benchmarks(site_data, company_data):
    """
    Compare one site's performance to overall company.
    Shows overall + breakdown by account type + Sysco brand split.
    """
    
    rows = []
    
    def add_benchmark_row(account_type, site_subset, company_subset):
        """Helper to calculate one benchmark row."""
        site_cy = site_subset["Pounds_CY"].sum()
        site_py = site_subset["Pounds_PY"].sum()
        site_delta = site_cy - site_py
        site_yoy_pct = (site_delta / site_py) if site_py > 0 else 0
        
        company_cy = company_subset["Pounds_CY"].sum()
        company_py = company_subset["Pounds_PY"].sum()
        company_delta = company_cy - company_py
        company_yoy_pct = (company_delta / company_py) if company_py > 0 else 0
        
        site_share_cy = (site_cy / company_cy) if company_cy > 0 else 0
        site_share_py = (site_py / company_py) if company_py > 0 else 0
        share_change = site_share_cy - site_share_py
        performance_gap = site_yoy_pct - company_yoy_pct
        
        rows.append({
            "Account_Type": account_type,
            "Site_CY": site_cy,
            "Site_PY": site_py,
            "Site_YoY_%": site_yoy_pct,
            "Company_YoY_%": company_yoy_pct,
            "Gap_vs_Company": performance_gap,
            "Site_Market_Share_CY": site_share_cy,
            "Site_Market_Share_PY": site_share_py,
            "Share_Change": share_change
        })
    
    # === OVERALL ===
    add_benchmark_row("All Accounts", site_data, company_data)
    
    # === BY ACCOUNT TYPE ===
    if "Customer Account Type Code" in site_data.columns:
        for acct_type in ["TRS", "CMU", "LCC"]:
            site_acct = site_data[site_data["Customer Account Type Code"] == acct_type]
            company_acct = company_data[company_data["Customer Account Type Code"] == acct_type]
            
            if not site_acct.empty and not company_acct.empty:
                add_benchmark_row(f"  {acct_type}", site_acct, company_acct)
                
                # === SYSCO BRAND SPLIT ===
                if "Sysco Brand Indicator" in site_data.columns:
                    for brand in ["Y", "N"]:
                        site_brand = site_acct[site_acct["Sysco Brand Indicator"] == brand]
                        company_brand = company_acct[company_acct["Sysco Brand Indicator"] == brand]
                        
                        if not site_brand.empty and not company_brand.empty:
                            brand_label = "Sysco" if brand == "Y" else "Non-Sysco"
                            add_benchmark_row(f"    {acct_type} - {brand_label}", site_brand, company_brand)
    
    return pd.DataFrame(rows)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--alignment", required=True)
    parser.add_argument("--company", required=True, help="Site number or name to benchmark")
    parser.add_argument("--outdir", default=".")
    args = parser.parse_args()
    
    print(f"\n🎯 SITE BENCHMARK REPORT")
    print(f"Site: {args.company}")
    print("="*60)
    
    # Load ALL sites first (for company benchmark)
    company_raw, _ = load_and_prepare(args.source, args.alignment)
    company_status, current_week = compute_windows(company_raw)
    
    # Load FILTERED site data
    site_raw, _ = load_and_prepare(args.source, args.alignment, company=args.company)
    site_status, _ = compute_windows(site_raw)
    
    if site_status.empty:
        print(f"❌ No data found for site '{args.company}'")
        return
    
    # Calculate benchmarks
    benchmarks = calculate_site_benchmarks(site_status, company_status)
    
    # Save to Excel
    output_path = f"{args.outdir}/Site_{args.company}_Benchmark_{datetime.now():%Y%m%d}.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        benchmarks.to_excel(writer, sheet_name="Benchmark", index=False)
        
        wb = writer.book
        ws = wb["Benchmark"]
        
        # Format headers
        for cell in ws[1]:
            cell.font = Font(bold=True, size=11, color="FFFFFF")
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        
        # Format numbers and percentages
        from openpyxl.styles import Font, PatternFill
        for row in range(2, ws.max_row + 1):
            for col_idx in [2, 3]:  # Site_CY, Site_PY
                ws.cell(row, col_idx).number_format = '#,##0'
            for col_idx in [4, 5, 6, 7, 8, 9]:  # Percentage columns
                ws.cell(row, col_idx).number_format = '0.00%'
        
        # Color-code Gap column (red = underperforming, green = outperforming)
        for row in range(2, ws.max_row + 1):
            gap_val = ws.cell(row, 6).value
            if gap_val and float(gap_val) < 0:
                ws.cell(row, 6).fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            elif gap_val and float(gap_val) > 0:
                ws.cell(row, 6).fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        
        # === ADD EXPLAINER TEXT ===
        explain_row = ws.max_row + 3
        
        ws.cell(explain_row, 1, value="HOW TO READ THIS REPORT:").font = Font(bold=True, size=12, color="0066CC")
        explain_row += 1
        
        ws.cell(explain_row, 1, value="Gap vs Company: Negative (red) = site underperforming company average | Positive (green) = outperforming")
        explain_row += 1
        
        ws.cell(explain_row, 1, value="Market Share Change: Negative = losing share of company volume | Positive = gaining share")
        explain_row += 1
        
        ws.cell(explain_row, 1, value="Hierarchy: Indented rows show account type breakdowns and Sysco vs Non-Sysco brand splits")
        explain_row += 2
        
        # === GENERATE TRS LEADS ===
        from category_utils import classify_conversion, build_sales_leads
        
        # Only TRS accounts for site leads
        site_trs = site_status[site_status.get("Customer Account Type Code") == "TRS"].copy()
        
        if not site_trs.empty:
            # Classify conversion status
            site_trs = classify_conversion(site_trs)
            
            # Build leads (TRS only, min 20 lbs/week)
            min_ytd = current_week * 20
            trs_leads = build_sales_leads(site_trs, site_raw, min_ytd=min_ytd)
            
            if not trs_leads.empty:
                trs_leads.to_excel(writer, sheet_name="Site_Leads", index=False)
                
                ws_leads = wb["Site_Leads"]
                ws_leads.insert_rows(1, 3)
                ws_leads.cell(1, 1, value=f"TRS SALES LEADS FOR SITE {args.company}").font = Font(bold=True, size=14, color="0066CC")
                ws_leads.cell(2, 1, value=f"{len(trs_leads)} opportunities | TRS accounts only | Minimum {min_ytd:,.0f} lbs PY").font = Font(size=10, italic=True)
                ws_leads.cell(3, 1, value="'Needs Both' = highest priority (neither item nor vendor aligned)")
                
                print(f"\n📋 TRS Leads: {len(trs_leads)} opportunities")
                
            # === GENERATE VENDOR SPLITS (separate CSV files) ===
            import os
            vendor_dir = os.path.join(args.outdir, "Vendor_Leads")
            os.makedirs(vendor_dir, exist_ok=True)

            if not trs_leads.empty and "Aligned Vendor (SUVC)" in trs_leads.columns:
                for vendor_code in trs_leads["Aligned Vendor (SUVC)"].dropna().unique():
                    vendor_leads = trs_leads[trs_leads["Aligned Vendor (SUVC)"] == vendor_code]
                    vendor_name = vendor_leads["Aligned Supplier Name"].iloc[0] if "Aligned Supplier Name" in vendor_leads.columns else "Unknown"
                    safe_name = "".join(c for c in str(vendor_name) if c.isalnum() or c in (' ', '-', '_')).strip()[:30]
                    
                    csv_path = os.path.join(vendor_dir, f"Site{args.company}_Vendor_{vendor_code}_{safe_name}.csv")
                    vendor_leads.to_csv(csv_path, index=False)
                
                print(f"📦 Created vendor CSV files in {vendor_dir}")

    print(f"\n✅ Report saved: {output_path}")
    print("\n📊 Performance Summary:")
    print(benchmarks.to_string(index=False))
    

if __name__ == "__main__":
    main()
