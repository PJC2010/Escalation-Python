import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl.utils import get_column_letter
import warnings
warnings.simplefilter(action='ignore', category=UserWarning)
def read_excel_safely(file_path):
   """Safely read Excel file"""
   try:
       return pd.read_excel(file_path, engine='openpyxl')
   except Exception as e1:
       try:
           return pd.read_excel(file_path, engine='openpyxl', data_only=True)
       except Exception as e2:
           print(f"Failed to read {file_path.name}")
           return None
def autofit_columns(worksheet, df):
   """Automatically adjust column widths"""
   for idx, column in enumerate(df.columns):
       column_letter = get_column_letter(idx + 1)
       max_length = max(
           df[column].astype(str).str.len().max(),
           len(str(column))
       )
       adjusted_width = (max_length + 2) * 1.2
       adjusted_width = min(max(adjusted_width, 8), 50)  # Min 8, Max 50
       worksheet.column_dimensions[column_letter].width = adjusted_width
# Set your paths here
BASE_PATH = r"C:/Users/PeteCastillo/OneDrive - VillageMD\Documents - VMD- Quality Leadership- PHI/Med Adherence Exception File Worklists/Week of 11.18"
OUTPUT_FOLDER = r"C:/Users/PeteCastillo/OneDrive - VillageMD/Desktop/Escalation Python/"
# Create output folder if it doesn't exist
output_path = Path(OUTPUT_FOLDER)
output_path.mkdir(parents=True, exist_ok=True)
# Get all Excel files from the folder
folder_path = Path(BASE_PATH)
excel_files = list(folder_path.rglob("*.xlsx"))
# List to store all dataframes
all_dfs = []
# Process each file
for file_path in excel_files:
   df = read_excel_safely(file_path)
   if df is not None:
       # Add file source information
       df['Source_File'] = file_path.name
       df['Folder_Path'] = str(file_path.parent.relative_to(folder_path))
       all_dfs.append(df)
# Combine all dataframes
master_df = pd.concat(all_dfs, ignore_index=True)
# Create output filename with timestamp
timestamp = datetime.now().strftime('%Y%m%d_%H%M')
output_file = output_path / f"Master_Worklist_{timestamp}.xlsx"
# Save to Excel with autofit columns
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
   master_df.to_excel(writer, index=False, sheet_name='Master_Worklist')
   autofit_columns(writer.sheets['Master_Worklist'], master_df)
print(f"Master worklist created with {len(master_df):,} records")
print(f"File saved as: {output_file}")


