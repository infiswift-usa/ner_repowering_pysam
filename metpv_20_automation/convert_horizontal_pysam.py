import pandas as pd
import csv
import sys

def convert_horizontal_metpv_to_pysam(input_file, output_file):
    print(f"Converting Horizontal METPV-11: {input_file}")
    
    # Storage for hourly data
    data = {}
    metadata = {}
    
    with open(input_file, 'r', newline='', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        try:
            # First row contains station info: ID, Name, Lat_deg, Lat_min, Lon_deg, Lon_min, Elev
            header = next(reader)
            if len(header) >= 7:
                metadata['lat'] = float(header[2]) + float(header[3])/60.0
                metadata['lon'] = float(header[4]) + float(header[5])/60.0
                metadata['elev'] = float(header[6])
                print(f"Station Metadata Found: Lat={metadata['lat']:.4f}, Lon={metadata['lon']:.4f}, Elev={metadata['elev']:.1f}m")
        except Exception as e:
            print(f"Warning: Could not parse station header: {e}")
        
        for row in reader:
            row = [x.strip() for x in row]
            if not row or len(row) < 30: continue
            
            try:
                # Horizontal Indices: Type=0, Month=1, Day=2, Year=3
                m_type = int(float(row[0]))
                month = int(float(row[1]))
                day = int(float(row[2]))
                year = int(float(row[3]))
                
                # Hourly Data: Cols 4 to 27 (Hours 1..24)
                hourly_vals = [float(x) if x != '' else 0.0 for x in row[4:28]]
                
                # FORCE TYPICAL YEAR (2016 is a leap year but we want 8760 hours)
                # Actually METPV is 365 days usually.
                # If we want exactly 8760, we just map Day 1...365 to Jan 1...Dec 31 2017 (non-leap)
                # or just use an hour offset from Jan 1 2016.
                
                # Let's count days relative to Jan 1
                # The file might not be in order, so let's use a mapping DayOfYear -> Data
                # or just parse and then re-index.
                
                # Simplified: use Day and Month from file but force a specific year 2016
                base_dt = pd.Timestamp(year=2016, month=month, day=day)
                
                for h_idx, val in enumerate(hourly_vals):
                    # Align to hour: 0 to 23
                    dt = base_dt + pd.Timedelta(hours=h_idx)
                    
                    if dt not in data:
                        data[dt] = {'GHI': 0.0, 'DHI': 0.0, 'Temperature': 0.0, 'WindSpeed': 0.0}
                    
                    if m_type == 1:
                        # GHI: 0.01 MJ/m2/h -> W/m2 (Factor 2.777778)
                        data[dt]['GHI'] = val * 10000.0 / 3600.0
                    elif m_type == 2:
                        # DHI
                        data[dt]['DHI'] = val * 10000.0 / 3600.0
                    elif m_type == 5:
                        # Temp: 0.1 C -> C
                        data[dt]['Temperature'] = val * 0.1
                    elif m_type == 7:
                        # Wind: 0.1 m/s -> m/s
                        data[dt]['WindSpeed'] = val * 0.1
                        
            except Exception as e:
                continue

    if not data:
        print("Error: No data records were extracted. Check the input file format.")
        return

    df = pd.DataFrame.from_dict(data, orient='index')
    
    # Ensure all required columns exist
    for c in ['GHI', 'DHI', 'Temperature', 'WindSpeed']:
        if c not in df.columns: df[c] = 0.0
            
    df.index.name = 'DateTime'
    df = df.sort_index().reset_index()
    
    # Add metadata if available
    if metadata:
        df['Latitude'] = metadata['lat']
        df['Longitude'] = metadata['lon']
        df['Elevation'] = metadata['elev']

    df.to_csv(output_file, index=False)
    print(f"Saved {len(df)} rows to {output_file}")
    print(df.head())

if __name__ == "__main__":
    infile = r"D:\VS_CODE\Infiswift\metpv_20_automation\hm53091year.csv"
    outfile = "metpv_horizontal_pysam.csv"
    convert_horizontal_metpv_to_pysam(infile, outfile)
