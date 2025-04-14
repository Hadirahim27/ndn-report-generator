
import pandas as pd
from io import BytesIO

def process_ndn_excel(uploaded_file):
    import openpyxl

    xl = pd.ExcelFile(uploaded_file)
    raw_df = xl.parse(xl.sheet_names[0])

    raw_df = raw_df[~raw_df['Status'].str.contains('fault cancel', case=False, na=False)]
    raw_df = raw_df[~raw_df['Fault Type'].str.contains('fault cancelled', case=False, na=False)]

    raw_df.rename(columns={'Hours': 'hour', 'outage': 'duration'}, inplace=True)
    raw_df.rename(columns={'hour': 'outage', 'duration': 'hour'}, inplace=True)

    raw_df.fillna("unknown", inplace=True)
    raw_df.replace(r'^\s*$', 'unknown', regex=True, inplace=True)

    sheets = {
        'Fiber': raw_df[raw_df['Fault Type'].str.lower().isin(['cable fault', 'other', 'others'])],
        'Power': raw_df[raw_df['Fault Type'].str.lower() == 'power problem'],
        'Equipment': raw_df[raw_df['Fault Type'].str.lower() == 'equipment fault'],
        'Valid': raw_df[raw_df['Fault Type'].str.lower() != '3rd party provider'],
        '3rd Party Provider': raw_df[raw_df['Fault Type'].str.lower() == '3rd party provider']
    }

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            if 'hour' in df.columns and 'outage' in df.columns:
                for row_idx in range(2, len(df) + 2):
                    col_hour = df.columns.get_loc("hour") + 1
                    col_outage = df.columns.get_loc("outage") + 1
                    cell = ws.cell(row=row_idx, column=col_hour)
                    ref = ws.cell(row=row_idx, column=col_outage)
                    cell.value = f"={ref.coordinate}*24"
    output.seek(0)
    return output
