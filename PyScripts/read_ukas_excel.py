import pandas as pd
import sys
from collections import defaultdict

def read_ukas_excel(file_path):
    try:
        # Read the Excel file
        print(f"Reading Excel file: {file_path}")
        df = pd.read_excel(file_path, engine='openpyxl')
        
        # Clean up the dataframe
        df = df.dropna(how='all')
        
        # Display basic information about the file
        print("\nFile Information:")
        print("-" * 50)
        print(f"Number of rows (after cleaning): {len(df)}")
        print(f"Number of columns: {len(df.columns)}")
        
        # Find and display certificate information
        print("\nCertificate Details:")
        print("-" * 50)
        for idx, row in df.iterrows():
            row_str = str(row.values)
            if 'ISSUED BY' in row_str:
                print(f"Issuer: {row['Unnamed: 2']}")
            elif 'CERTIFICATE NUMBER' in row_str:
                print(f"Certificate Number: {row['Unnamed: 2']}")
            elif 'DATE OF ISSUE' in row_str:
                print(f"Date of Issue: {row['Unnamed: 2']}")
            elif 'TEMPERATURE' in row_str:
                print(f"Temperature: {row['Unnamed: 2']}")
            elif 'HUMIDITY' in row_str:
                print(f"Humidity: {row['Unnamed: 2']}")
        
        # Extract voltage measurements
        print("\nVoltage Measurements Analysis:")
        print("-" * 50)
        
        # Find rows containing A-N measurements
        an_measurements = df[df['CERTIFICATE OF CALIBRATION'] == 'A-N']
        
        if not an_measurements.empty:
            # Organize measurements by voltage range
            measurements = defaultdict(list)
            
            for idx, row in an_measurements.iterrows():
                voltage = row['Unnamed: 1']
                measurement = row['Unnamed: 5']
                if pd.notna(voltage) and pd.notna(measurement):
                    if voltage <= 125:
                        measurements['Low Range (≤125V)'].append((voltage, measurement))
                    elif voltage <= 400:
                        measurements['Mid Range (126V-400V)'].append((voltage, measurement))
                    else:
                        measurements['High Range (>400V)'].append((voltage, measurement))
            
            # Display measurements by range
            for range_name, values in measurements.items():
                print(f"\n{range_name}:")
                print("Voltage (V) | Measurement (V)")
                print("-" * 30)
                
                # Sort by voltage within each range
                values.sort(key=lambda x: x[0])
                
                # Display unique measurements (remove duplicates)
                seen = set()
                for voltage, measurement in values:
                    key = f"{voltage}_{measurement}"
                    if key not in seen:
                        print(f"{voltage:10.2f} | {measurement:10.3f}")
                        seen.add(key)
                
                # Calculate statistics for this range
                voltages = [v[0] for v in values]
                measurements = [v[1] for v in values]
                
                print(f"\nRange Statistics:")
                print(f"Number of test points: {len(set(voltages))}")
                print(f"Voltage range: {min(voltages):.1f}V - {max(voltages):.1f}V")
                print(f"Mean measurement: {sum(measurements)/len(measurements):.3f}V")
                print(f"Min measurement: {min(measurements):.3f}V")
                print(f"Max measurement: {max(measurements):.3f}V")
        else:
            print("No A-N measurements found in the file")
        
        return df
        
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    file_path = "../UKAS STandard cert work in progress 920 (1).xlsx"
    df = read_ukas_excel(file_path)
