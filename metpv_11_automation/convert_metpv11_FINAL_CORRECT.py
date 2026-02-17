import pandas as pd
import numpy as np
from datetime import datetime, timedelta

metpv_file = r"D:\MaxiFit Installation Files\METPV11\平均年\mea53091.txt"

print("="*70)
print("METPV-11 FIXED-LENGTH PARSER - OFFICIAL SPEC COMPLIANT")
print("="*70)

with open(metpv_file, 'r', encoding='shift-jis') as f:
    lines = f.readlines()

# Record 1: Header (Fixed Length)
# 1-5: ID, 7-26: Name, 27-31: Lat Deg, 32-36: Lat Min, 37-41: Lon Deg, 42-46: Lon Min, 47-54: Elev
header = lines[0]
station_id = header[0:5].strip()
station_name = header[6:26].strip()
lat = float(header[26:31]) + float(header[31:36]) / 60.0
lon = float(header[36:41]) + float(header[41:46]) / 60.0
elevation = float(header[46:54])

print(f"Station: {station_id} {station_name}")
print(f"Coordinates: {lat:.3f}°N, {lon:.3f}°E")
print(f"Elevation: {elevation}m")

# Parse Data (Record 2+)
# 1-5: Element, 7-8: Month, 9-10: Day, 11-15: Space
# 16-20, 21-25...: Hourly Data (4 bytes value + 1 byte remark)
data_rows = {}

for line in lines[1:]:
    if len(line) < 130: continue
    
    element_code = line[0:5].strip()
    
    hourly_values = []
    # 24 hours starting at index 15 (0-indexed)
    for h in range(24):
        start = 15 + (h * 5)
        end = start + 5
        field = line[start:end]
        
        # Byte 1-4 is the value
        val_str = field[0:4].strip()
        # Byte 5 is remark (discarded for calculation)
        
        try:
            val = int(val_str)
            if val == 8888:
                hourly_values.append(np.nan)
            else:
                hourly_values.append(val)
        except:
            hourly_values.append(np.nan)
    
    if element_code not in data_rows:
        data_rows[element_code] = []
    data_rows[element_code].extend(hourly_values)

# Create DataFrame (Typical Year 8760)
start_date = datetime(2016, 1, 1, 0, 0, 0)
timestamps = [start_date + timedelta(hours=i) for i in range(8760)]

def get_param(code, scale):
    vals = data_rows.get(code, [])
    # Ensure 8760 length
    if len(vals) < 8760:
        vals = vals + [np.nan] * (8760 - len(vals))
    return np.array(vals[:8760]) * scale

# Official Units mapping (based on provided doc)
# 00001 (GHI): 0.01 MJ/m2 -> W/m2 (Value * 0.01 * 1e6 / 3600 = Value * 2.7778)
# 00002 (Beam): 0.01 MJ/m2 -> W/m2
# 00003 (Diffuse): 0.01 MJ/m2 -> W/m2
# 00005 (Temp): 0.1 C -> C (Value * 0.1)
# 00007 (Wind): 0.1 m/s -> m/s (Value * 0.1)

df = pd.DataFrame({
    'DateTime': timestamps,
    'GHI': get_param('00001', 2.7778),
    'DNI_horiz': get_param('00002', 2.7778), # This is Beam on horizontal
    'DHI': get_param('00003', 2.7778),
    'Temperature': get_param('00005', 0.1),
    'WindSpeed': get_param('00007', 0.1),
    'Latitude': lat,
    'Longitude': lon,
    'Elevation': elevation
})

# Save for Browser (Strict CRLF, No Bom, Integer values where possible)
output_file = r"D:\VS_CODE\Infiswift\metpv-11 working\metpv11_clean_v2.csv"
df.to_csv(output_file, index=False)

print(f"\nSaved {len(df)} rows to: {output_file}")
print(f"Annual GHI: {df['GHI'].sum()/1000:.1f} kWh/m²/year")
print(f"Peak GHI: {df['GHI'].max():.1f} W/m²")
print(f"Avg Temp: {df['Temperature'].mean():.1f}°C")
print("="*70)
