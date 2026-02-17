import pandas as pd
import numpy as np
import PySAM.Pvwattsv8 as pv
import os
from datetime import datetime

# NumPy 2.0 compatibility patch
if not hasattr(np, 'Inf'):
    np.Inf = np.inf
if not hasattr(np, 'int'):
    np.int = int
if not hasattr(np, 'float'):
    np.float = float

print("="*70)
print("PYSAM SIMULATION (PVWATTS V8) - 9-PCS GRANULAR MODEL")
print("METPV-11 CORRECTED DATA")
print("="*70)

# 1. Load the CLEANED weather data
weather_file = r"D:\VS_CODE\Infiswift\metpv-11 working\metpv11_clean_v2.csv"
if not os.path.exists(weather_file):
    print(f"Error: {weather_file} not found. Run the conversion script first.")
    exit(1)

df = pd.read_csv(weather_file)
df['DateTime'] = pd.to_datetime(df['DateTime'])

# Extract metadata
lat = df['Latitude'].iloc[0]
lon = df['Longitude'].iloc[0]
elev = df['Elevation'].iloc[0]
tz = 9 # Japan Standard Time

print(f"\nWeather Data: {len(df)} hours")
print(f"Location: {lat:.3f}N, {lon:.3f}E, {elev:.1f}m")

# 2. Calculate DNI and solar position (Same for all PCS)
print("\n[Calculating DNI using pvlib...]")
import pvlib
location = pvlib.location.Location(lat, lon, tz='Asia/Tokyo')
solar_pos = location.get_solarposition(df['DateTime'])

zenith = solar_pos['zenith'].values
cos_zenith = np.cos(np.radians(zenith))
ghi = df['GHI'].values
dhi = df['DHI'].values

# Avoid division by zero at night and sunrise/sunset (cap at 87 deg)
dni = np.zeros_like(ghi)
daytime = zenith < 87 
dni[daytime] = (ghi[daytime] - dhi[daytime]) / cos_zenith[daytime]

# Physical clipping (PySAM checks for DNI/DHI in [0, 1500])
dni = np.clip(dni, 0, 1500)
dhi = np.clip(dhi, 0, 1500)
dni = np.nan_to_num(dni)

# Prepare Weather Data Dictionary
weather_data = {
    'lat': lat, 'lon': lon, 'tz': tz, 'elev': elev,
    'year': df['DateTime'].dt.year.tolist(),
    'month': df['DateTime'].dt.month.tolist(),
    'day': df['DateTime'].dt.day.tolist(),
    'hour': df['DateTime'].dt.hour.tolist(),
    'minute': [0] * len(df),
    'gh': ghi.tolist(),
    'dn': dni.tolist(),
    'df': dhi.tolist(),
    'tdry': df['Temperature'].tolist(),
    'wspd': df['WindSpeed'].tolist()
}

# 3. Simulate Inverter Groups

def run_pcs_simulation(name, count, dc_kw, ac_kw):
    print(f"\nSimulating {name}: {count} units x {dc_kw}kW DC / {ac_kw}kW AC")
    system = pv.default("PVWattsNone")
    
    # Per unit settings (Matching Hardware Specs)
    system.SystemDesign.system_capacity = dc_kw
    system.SystemDesign.dc_ac_ratio = dc_kw / ac_kw
    system.SystemDesign.inv_eff = 98.4
    system.SystemDesign.losses = 5.0
    system.SystemDesign.tilt = 20
    system.SystemDesign.azimuth = 185
    system.SystemDesign.gcr = 0.81
    
    system.SolarResource.solar_resource_data = weather_data
    system.execute()
    
    # Scale by unit count
    ac_hourly = np.array(system.Outputs.gen) * count # kWh per hour
    return ac_hourly

# Group A: 4 units x 140.0kW DC / 100kW AC (224 modules x 0.625kW)
ac_a = run_pcs_simulation("Group A (PCS 01-04)", 4, 140.0, 100.0)

# Group B: 5 units x 130.0kW DC / 95kW AC (208 modules x 0.625kW)
ac_b = run_pcs_simulation("Group B (PCS 05-09)", 5, 130.0, 95.0)

# Total plant production
ac_total_hourly = ac_a + ac_b
ac_annual = np.sum(ac_total_hourly)

# 4. Results
print("\n" + "="*70)
print("FINAL RESULTS")
print("="*70)
print(f"Annual AC Energy: {ac_annual:,.0f} kWh")

# Comparison with MAXIFIT
maxifit_annual = 1_070_210
diff = ac_annual - maxifit_annual
print(f"MAXIFIT Benchmark: {maxifit_annual:,.0f} kWh")
print(f"Difference:       {diff:,.0f} kWh ({diff/maxifit_annual:+.1%})")

# Monthly Resampling
df_res = pd.DataFrame({'DateTime': df['DateTime'], 'AC_kWh': ac_total_hourly})
df_res.set_index('DateTime', inplace=True)
monthly_ac = df_res['AC_kWh'].resample('ME').sum()

months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
          'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
print("\nMonthly Production (kWh):")
for i, m in enumerate(months):
    print(f"  {m}: {monthly_ac.iloc[i]:>8,.0f} kWh")

# Save results
monthly_ac_df = pd.DataFrame({'Month': months, 'AC_Energy_kWh': monthly_ac.values})
monthly_ac_df.to_csv("metpv11_pysam_results.csv", index=False)
print("\nResults saved to metpv11_pysam_results.csv")
print("="*70)
