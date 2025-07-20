import pandas as pd
import os

def convert_ukas_to_csv():
    # Read the UKAS Excel file
    excel_file = "UKAS STandard cert work in progress 920 (1).xlsx"
    try:
        # Read the UKAS 3150AFX sheet
        print("\nReading UKAS 3150AFX sheet...")
        df = pd.read_excel(excel_file, sheet_name='UKAS 3150AFX', header=None)
        
        # Create CSV content
        csv_content = []
        csv_content.append("# test_point,mode,phase_config,frequency,voltage")
        
        # Function to format test point
        def format_test_point(phase, voltage, mode, freq=50):
            return f'"{phase} {voltage}V Test",{mode},3-PHASE,{freq},{voltage}'
        
        # Process AC voltage tests
        csv_content.append("\n# AC Voltage Tests")
        ac_start = False
        for idx, row in df.iterrows():
            row_text = ' '.join([str(x) for x in row if pd.notna(x)]).lower()
            if '3 phase ac output voltage linearity' in row_text:
                ac_start = True
                continue
            if ac_start and len(row) >= 2:
                phase = row[0] if pd.notna(row[0]) else ''
                setting = row[1] if pd.notna(row[1]) else ''
                if isinstance(phase, str) and isinstance(setting, (int, float)):
                    if phase.startswith('A-N') or phase.startswith('B-N') or phase.startswith('C-N'):
                        csv_content.append(format_test_point(phase, setting, 'AC'))
                if '3 phase dc output voltage' in row_text:
                    ac_start = False
        
        # Process DC voltage tests
        csv_content.append("\n# DC Voltage Tests")
        dc_start = False
        for idx, row in df.iterrows():
            row_text = ' '.join([str(x) for x in row if pd.notna(x)]).lower()
            if '3 phase dc output voltage linearity' in row_text:
                dc_start = True
                continue
            if dc_start and len(row) >= 2:
                phase = row[0] if pd.notna(row[0]) else ''
                setting = row[1] if pd.notna(row[1]) else ''
                if isinstance(phase, str) and isinstance(setting, (int, float)):
                    if phase.startswith('A-N') or phase.startswith('B-N') or phase.startswith('C-N'):
                        csv_content.append(format_test_point(phase, setting, 'DC', 0))
        
        # Write to CSV file
        output_file = "ukas_voltage_tests_new.csv"
        with open(output_file, 'w') as f:
            f.write('\n'.join(csv_content))
        
        print(f"\nSuccessfully converted voltage test data to {output_file}")
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        if isinstance(e, FileNotFoundError):
            print("The Excel file was not found. Please check the file name and path.")
        elif "Sheet" in str(e):
            # List available sheets if sheet not found
            xls = pd.ExcelFile(excel_file)
            print("\nAvailable sheets:")
            for sheet in xls.sheet_names:
                print(f"- {sheet}")

if __name__ == "__main__":
    convert_ukas_to_csv()
