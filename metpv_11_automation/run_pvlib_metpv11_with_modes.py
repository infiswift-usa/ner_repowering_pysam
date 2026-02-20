import pandas as pd
import numpy as np

# NumPy 2.0 compatibility patch
if not hasattr(np, 'Inf'):
    np.Inf = np.inf
if not hasattr(np, 'int'):
    np.int = int
if not hasattr(np, 'float'):
    np.float = float

import pvlib
from pvlib.modelchain import ModelChain
from pvlib.pvsystem import Array, FixedMount, PVSystem
from pvlib.temperature import TEMPERATURE_MODEL_PARAMETERS
from pvlib.bifacial import infinite_sheds
import os

print("="*70)
print("PVLIB SIMULATION - 9-PCS GRANULAR MODEL")
print("USING METPV-11 CLEANED DATA")
print("="*70)

# 1. Site Configuration (From D:\test-pvlib-ner1004\specs\plant_config.py)
SITE_LAT = 34.856
SITE_LON = 136.452
SITE_TZ = 'Asia/Tokyo'
SURFACE_TILT = 20
SURFACE_AZIMUTH = 185

# --- 1. SIMULATION MODE SELECTOR ---
# 'LEVEL_1': Mono-facial, Idealized (No bifacial, No shading) -> Target: ~1.5M kWh
# 'LEVEL_2': Bifacial, Idealized (Bifacial gain, No shading)   -> Target: ~1.6M kWh
# 'LEVEL_3': Bifacial, Realistic (Bifacial gain + Shading)      -> Target: ~1.29M kWh
SIM_MODE = 'LEVEL_3'

# Hardware Parameters (from NER132M625E-NGD Datasheet)
MODULE_PARAMS_PVWATTS = {
    "pdc0": 625,
    "gamma_pdc": -0.0029,  # -0.29% / C (Exact from Page 7)
}
BIFACIALITY = 0.80  # Bifaciality Factor
ALBEDO = 0.2        # Ground Reflection Assumption
MODULE_LENGTH = 2.382

# Shading & Height Parameters based on Mode
if SIM_MODE in ['LEVEL_1', 'LEVEL_2']:
    GCR = 0.01      # Effectively Zero shading (MAXIFIT Baseline)
    HUB_HEIGHT = 10.0 # High clearance (Idealized)
else:
    GCR = 0.326      # Realistic row spacing
    HUB_HEIGHT = 1.2  # Realistic ground clearance

PITCH = MODULE_LENGTH / GCR
ENABLE_BIFACIAL = (SIM_MODE != 'LEVEL_1')

INVERTER_PARAMETERS_100KW = {
    "pdc0": 140000,
    "eta_inv_nom": 0.984,
}

INVERTER_PARAMETERS_95KW = {
    "pdc0": 130000,
    "eta_inv_nom": 0.984,
}

# 2. Load and Prepare Weather Data
weather_file = r"D:\VS_CODE\Infiswift\metpv_11_automation\metpv11_clean_v2.csv"
if not os.path.exists(weather_file):
    print(f"Error: {weather_file} not found.")
    exit(1)

df = pd.read_csv(weather_file)
df['DateTime'] = pd.to_datetime(df['DateTime'])
df.set_index('DateTime', inplace=True)
# Localize to Japan Time so pvlib knows the exact solar position
df.index = df.index.tz_localize('Asia/Tokyo')

# Calculate DNI (Normalizing the Horizontal Direct Component)
location = pvlib.location.Location(SITE_LAT, SITE_LON, tz=SITE_TZ)
# Evaluating solar position at the midpoint of each hour
times_mid = df.index + pd.Timedelta(minutes=30)
solar_pos = location.get_solarposition(times_mid)
solar_pos.index = df.index # RE-ALIGN to top-of-hour indices to prevent NANs during transposition
zenith = solar_pos['zenith'].values
cos_zenith = np.cos(np.radians(zenith))

# Create pvlib compatible weather DataFrame
weather = pd.DataFrame(index=df.index)
weather['ghi'] = df['GHI']
weather['dhi'] = df['DHI']
# Element 00002 is Direct Horizontal. DNI = Direct Horizontal / cos(Zenith)
weather['dni'] = (df['DNI_horiz'] / cos_zenith).fillna(0).clip(0, 1500)
weather['temp_air'] = df['Temperature']
weather['wind_speed'] = df['WindSpeed']

# Calculate POA Irradiance
if ENABLE_BIFACIAL:
    print(f"\n[Mode: {SIM_MODE}] Calculating Bifacial POA (Infinite Sheds)...")
    bifacial_irrad = infinite_sheds.get_irradiance(
        SURFACE_TILT, SURFACE_AZIMUTH,
        solar_pos['zenith'].values, solar_pos['azimuth'].values,
        GCR, HUB_HEIGHT, PITCH,
        weather['ghi'].values, weather['dhi'].values, weather['dni'].values,
        ALBEDO,
        bifaciality=BIFACIALITY,
        vectorize=True
    )
    weather['poa_global'] = bifacial_irrad['poa_global']
    weather['poa_direct'] = bifacial_irrad['poa_front_direct']
    print(f"Bifacial Gain Applied: 80% with Albedo {ALBEDO}")
else:
    print(f"\n[Mode: {SIM_MODE}] Calculating Mono-facial POA (Standard)...")
    poa_front = pvlib.irradiance.get_total_irradiance(SURFACE_TILT, SURFACE_AZIMUTH, 
                                                      solar_pos['zenith'], solar_pos['azimuth'], 
                                                      weather['dni'], weather['ghi'], weather['dhi'])
    weather['poa_global'] = poa_front['poa_global']
    weather['poa_direct'] = poa_front['poa_direct']

weather['poa_diffuse'] = weather['poa_global'] - weather['poa_direct']

print(f"Debug Values at Mid-Day:")
print(f"Total POA: {weather['poa_global'].iloc[12]:.2f} W/m2")

# 3. Build the 9 PVSystems (4x100kW + 5x95kW)
mount = FixedMount(surface_tilt=SURFACE_TILT, surface_azimuth=SURFACE_AZIMUTH)
# SAPM temperature model for ground-mounted (open rack)
temp_params = TEMPERATURE_MODEL_PARAMETERS["sapm"]["open_rack_glass_glass"]

array_14 = Array(
    mount=mount,
    module_parameters=MODULE_PARAMS_PVWATTS,
    modules_per_string=16,
    strings=14,
    temperature_model_parameters=temp_params,
)
array_13 = Array(
    mount=mount,
    module_parameters=MODULE_PARAMS_PVWATTS,
    modules_per_string=16,
    strings=13,
    temperature_model_parameters=temp_params,
)

systems = []
# 4 PCS with 14 arrays
for _ in range(4):
    systems.append(PVSystem(arrays=[array_14], inverter_parameters=INVERTER_PARAMETERS_100KW))
# 5 PCS with 13 arrays
for _ in range(5):
    systems.append(PVSystem(arrays=[array_13], inverter_parameters=INVERTER_PARAMETERS_95KW))

# 4. Run ModelChain for each system
ac_plant = pd.Series(0.0, index=weather.index, dtype=float)

print(f"\nRunning pvlib ModelChain for {len(systems)} systems...")
for i, system in enumerate(systems):
    mc = ModelChain.with_pvwatts(system, location)
    # run_model_from_poa uses our pre-calculated bifacial global POA
    mc.run_model_from_poa(weather)
    ac_plant = ac_plant.add(mc.results.ac, fill_value=0)

# Apply 5% system losses (Matching system.SystemDesign.losses in PySAM)
#ac_plant = ac_plant * (1 - 0.05)

# 5. Summary and Results
monthly_kwh = ac_plant.resample("ME").sum() / 1000
yearly_kwh = ac_plant.sum() / 1000

print("\n" + "="*70)
print("FINAL RESULTS (PVLIB WORKFLOW)")
print("="*70)
print(f"Annual AC Energy: {yearly_kwh:,.0f} kWh")

print("\nMonthly Production (kWh):")
months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
for i, m in enumerate(months):
    print(f"  {m}: {monthly_kwh.iloc[i]:>8,.0f} kWh")

# Save results
monthly_kwh.to_csv(r"D:\VS_CODE\Infiswift\metpv_11_automation\metpv11_pvlib_results.csv")
print("\nResults saved to metpv11_pvlib_results.csv")
print("="*70)
