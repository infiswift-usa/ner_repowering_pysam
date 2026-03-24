import os
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine

# ── Database Connection ───────────────────────────────────────────────────────
# Replace with your actual credentials
DB_URL = "mysql+mysqlconnector://root:params1812@localhost:3306/priceCalci"
engine = create_engine(DB_URL)

# ══════════════════════════════════════════════════════════════════════════════
# 1.  JEPX ── Process Auction Data (Check & Write Append)
# ══════════════════════════════════════════════════════════════════════════════

def process_and_upload_jepx(file_path: str):
    print(f"\n>>> Processing JEPX CSV: {file_path}")
    try:
        df = pd.read_csv(file_path, encoding="shift-jis")
    except UnicodeDecodeError:
        df = pd.read_csv(file_path, encoding="utf-8-sig")

    clean = df[df["商品"] == "非FIT(再エネ指定)"].copy()
    
    # KEEP AS DATETIME: Do not convert to strings. .dt.normalize() safely strips out any hours/minutes!
    clean["約定日"] = pd.to_datetime(clean["約定日"]).dt.normalize()
    clean = clean.sort_values("約定日").reset_index(drop=True)

    # --- Math (Rolling 4-Row Average) ---
    clean['volume_x_price'] = clean['約定価格(円/kWh)'] * clean['約定総量(kWh)']
    rolling_sumproduct = clean['volume_x_price'].rolling(window=4).sum()
    rolling_sum = clean['約定総量(kWh)'].rolling(window=4).sum()
    clean['加重平均値'] = rolling_sumproduct / rolling_sum

    # Drop temporary column
    clean = clean.drop(columns=['volume_x_price'])

    # Force exact column order
    jepx_cols = [
        "年度", "開催回", "商品", "約定日", "約定総量(kWh)", "約定価格(円/kWh)", 
        "約定最高価格(円/kWh)", "約定最低価格(円/kWh)", "入札会員数", "約定会員数", 
        "売り入札量(kWh)", "買い入札量(kWh)", "加重平均値"
    ]
    clean = clean.reindex(columns=jepx_cols)
    clean["約定日"] = pd.to_datetime(clean["約定日"]).dt.strftime("%Y-%m-%d")
    # --- CHECK & WRITE ---
    existing_dates = []
    try:
        df_existing = pd.read_sql("SELECT 約定日 FROM non_fossil", con=engine)
        existing_dates = pd.to_datetime(df_existing['約定日']).dt.strftime("%Y-%m-%d").tolist()    
    except Exception:
        pass 

    # Filter purely by pandas datetime objects
    clean = clean[~clean["約定日"].isin(existing_dates)]

    if clean.empty:
        print("⏭️ No new JEPX dates to add. SQL 'non_fossil' table is already up to date.")
    else:
        clean.to_sql("non_fossil", con=engine, if_exists="append", index=False)
        print(f"✅ Appended {len(clean)} new rows to 'non_fossil'.")


# ══════════════════════════════════════════════════════════════════════════════
# 2.  OCCTO ── Process Regional Data (Maintain Latest Dashboard Table)
# ══════════════════════════════════════════════════════════════════════════════

def process_and_upload_occto(file_path: str):
    print(f"\n>>> Processing OCCTO CSV: {file_path}")
    try:
        df = pd.read_csv(file_path, encoding="shift-jis")
    except UnicodeDecodeError:
        df = pd.read_csv(file_path, encoding="utf-8-sig")

    solar = df[df["電源種別"] == "太陽光"].copy()
    solar["参照価格"] = (
        solar["前年度平均価格"] + solar["当年度月間平均価格"] - solar["前年度月間平均価格"]
    )

    digits = "".join(filter(str.isdigit, os.path.basename(file_path)))
    target_year = int(digits) if digits else pd.to_datetime(solar["年月"], format="%Y/%m").dt.year.max()
    year_col_name = f"{target_year}年度"

    solar["month_str"] = pd.to_datetime(solar["年月"], format="%Y/%m").dt.month.astype(str) + "月"

    regions = ["北海道", "東北", "東京", "中部", "北陸", "関西", "中国", "四国", "九州", "沖縄"]
    months_cols = ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月"]
    
    # --- CHECK DB DASHBOARD ---
    try:
        existing_df = pd.read_sql_table("occto_latest", con=engine)
        if year_col_name in existing_df.columns:
            print(f"📂 Found existing active table for {year_col_name}. Updating...")
            latest_df = existing_df.copy()
        else:
            print(f"✨ New year detected! Replacing old dashboard with fresh table for {year_col_name}...")
            latest_df = pd.DataFrame({year_col_name: regions})
            for m in months_cols:
                latest_df[m] = None
    except Exception:
        print(f"🆕 No existing dashboard found. Creating fresh table for {year_col_name}...")
        latest_df = pd.DataFrame({year_col_name: regions})
        for m in months_cols:
            latest_df[m] = None

    # Fill CSV data
    for _, row in solar.iterrows():
        reg = row["エリア"]
        m_col = row["month_str"]
        price = row["参照価格"]
        if m_col in months_cols:
            latest_df.loc[latest_df[year_col_name] == reg, m_col] = price

    # --- THE FIX: ONLY calculate average if "3月" is present! ---
    # Check if the "3月" column has no empty (NaN/None) values
    if latest_df["3月"].notnull().all():
        latest_df["年度平均"] = latest_df[months_cols].mean(axis=1)
        is_complete = True
    else:
        latest_df["年度平均"] = None  # Leave blank until March arrives
        is_complete = False

    # Overwrite the dashboard with exact column layout
    final_cols = [year_col_name] + months_cols + ["年度平均"]
    latest_df = latest_df[final_cols]
    
    latest_df.to_sql("occto_latest", con=engine, if_exists="replace", index=False)
    print(f"✅ 'occto_latest' dashboard successfully updated with {target_year} data.")

    # --- Update reference_price ONLY if March triggered completion ---
    if is_complete:
        print(f"🎉 3月 (March) data detected! Calculating final averages and updating 'reference_price'...")
        try:
            ref_df = pd.read_sql_table("reference_price", con=engine)
            new_averages = latest_df.set_index(year_col_name)["年度平均"].to_dict()
            ref_df[year_col_name] = ref_df['集計'].map(new_averages)
            ref_df.to_sql("reference_price", con=engine, if_exists="replace", index=False)
            print(f"✅ 'reference_price' table successfully updated with the new {year_col_name} column!")
        except Exception as e:
            print(f"⚠️ Could not automatically update 'reference_price'. Error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# 3.  MAIN RUNNER
# ══════════════════════════════════════════════════════════════════════════════

def run_extractor(
    jepx_csv: str  = "nf_summary_2025.csv",
    occto_csv: str = "FY2025_sansyo_kakaku.csv"
):
    print("=" * 60)
    print("DATA EXTRACTOR (Direct to SQL Pipeline)")
    print("=" * 60)

    if os.path.exists(jepx_csv):
        process_and_upload_jepx(jepx_csv)
    else:
        print(f"[SKIP] {jepx_csv} not found.")

    if os.path.exists(occto_csv):
        process_and_upload_occto(occto_csv)
    else:
        print(f"[SKIP] {occto_csv} not found.")

    print("\n🎉 All database updates complete!")


if __name__ == "__main__":
    run_extractor(
        jepx_csv="nf_summary_2025.csv", 
        occto_csv="FY2025_sansyo_kakaku.csv"
    )