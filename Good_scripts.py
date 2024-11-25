import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

class WorklistAnalyzer:
    def __init__(self):

        """

        Initialize the WorklistAnalyzer

        """

        self.base_path = None

        self.output_folder = None

        self.current_date = None

    def set_date(self, date_str):

        """

        Set the date manually for folder path

        Args:

            date_str (str): Date string in 'MM.DD' format (e.g., '11.18')

        """

        try:

            datetime.strptime(date_str, '%m.%d')

            self.current_date = date_str

            print(f"Date set to: {self.current_date}")

        except ValueError:

            raise ValueError("Invalid date format. Please use MM.DD format (e.g., '11.18')")

    def get_next_monday(self):

        """Get the date of the upcoming Monday in mm.dd format"""

        today = datetime.now()

        days_ahead = 7 - today.weekday()

        next_monday = today + timedelta(days=days_ahead)

        return next_monday.strftime('%m.%d')

    def get_week_folder(self):

        """Get the week folder path based on set date"""

        if self.current_date is None:

            self.current_date = datetime.now().strftime('%m.%d')

            print(f"No date set, using current date: {self.current_date}")

        folder_name = f"Week of {self.current_date}"

        folder_path = self.base_path / folder_name

        if not folder_path.exists():

            print(f"Warning: Folder not found: {folder_path}")

        return folder_path

    def get_worklist_files(self):

        """Get list of worklist files"""

        year = "2024"

        return [

            f"AZ.CO.MI Medication Adherence Worklist File Week of {self.current_date}.{year}.xlsx",

            f"TX Medication Adherence Worklist File Week of {self.current_date}.{year}.xlsx",

            f"ATL.KY Medication Adherence Worklist File Week of {self.current_date}.{year}.xlsx"

        ]

    def read_excel_safely(self, file_path):

        """Safely read Excel file with multiple fallback options"""

        try:

            return pd.read_excel(file_path, engine='openpyxl')

        except Exception as e1:

            try:

                return pd.read_excel(file_path, engine='openpyxl', data_only=True)

            except Exception as e2:

                try:

                    return pd.read_excel(file_path, engine='xlrd')

                except Exception as e3:

                    print(f"Failed to read {file_path.name} with all methods:")

                    print(f"Error 1: {str(e1)}")

                    print(f"Error 2: {str(e2)}")

                    print(f"Error 3: {str(e3)}")

                    return None

    def create_pivot_tables(self, df):

        """Create various pivot tables for analysis"""

        pivots = {}

        try:

            # Pivot 1: Escalation Types by Practice

            pivots['Escalation_by_Practice'] = pd.pivot_table(

                df,

                index='Practice Name',

                columns='Escalation Path',

                values='UID',

                aggfunc='count',

                fill_value=0,

                margins=True

            )

            # Pivot 2: Summary by Provider

            pivots['Summary_by_Provider'] = pd.pivot_table(

                df,

                index='PrescribingName',

                values=['UID'],

                columns='Escalation Path',

                aggfunc='count',

                fill_value=0,

                margins=True

            )

            # Create a basic summary

            summary_data = {

                'Metric': [

                    'Total Escalations',

                    'Market/PHO Escalations',

                    'Practice Escalations',

                    'Total Practices',

                    'Total Providers',

                    'Last Updated'

                ],

                'Value': [

                    len(df),

                    len(df[df['Escalation Path'] == 'Market/PHO Escalation']),

                    len(df[df['Escalation Path'] == 'Practice Escalation']),

                    df['Practice Name'].nunique(),

                    df['PrescribingName'].nunique(),

                    datetime.now().strftime('%Y-%m-%d %H:%M')

                ]

            }

            pivots['Summary'] = pd.DataFrame(summary_data)

        except Exception as e:

            print(f"Error creating pivot tables: {str(e)}")

        return pivots

    def process_worklists(self):

        """Process all worklist files and return data by market"""

        folder_path = self.get_week_folder()

        worklist_files = self.get_worklist_files()

        market_dfs = {}

        for file_name in worklist_files:

            file_path = folder_path / file_name

            try:

                if not file_path.exists():

                    print(f"File not found: {file_path}")

                    continue

                print(f"Processing: {file_name}")

                df = self.read_excel_safely(file_path)

                if df is None:

                    continue

                # Clean up column names

                df.columns = [col.strip() for col in df.columns]

                # Find the correct column names (case-insensitive)

                escalation_col = next((col for col in df.columns 

                                     if 'escalation path' in col.lower()), 'Escalation Path')

                market_code_col = next((col for col in df.columns 

                                      if 'market code' in col.lower()), 'MarketCode')
                

                if escalation_col not in df.columns or market_code_col not in df.columns:

                    print(f"Required columns not found in {file_name}")

                    continue

                # Filter for escalation paths

                filtered_df = df[df[escalation_col].isin([

                    'Market/PHO Escalation',

                    'Practice Escalation'

                ])]

                if len(filtered_df) > 0:

                    # Group by MarketCode

                    for market_code in filtered_df[market_code_col].unique():

                        if pd.isna(market_code):

                            continue

                        market_df = filtered_df[filtered_df[market_code_col] == market_code]

                        if market_code in market_dfs:

                            market_dfs[market_code] = pd.concat([market_dfs[market_code], market_df])

                        else:

                            market_dfs[market_code] = market_df

                    print(f"Successfully processed {file_name}")

            except Exception as e:

                print(f"Error processing {file_name}: {str(e)}")

        return market_dfs

    def create_market_files(self, market_dfs):

        """Create separate files for each market with pivot tables"""

        if not market_dfs:

            print("No data available to create files")

            return

        next_monday = self.get_next_monday()

        for market_code, df in market_dfs.items():

            try:

                # Create filename with next Monday's date and market code

                filename = f"{next_monday} {market_code} Med Adherence Escalations.xlsx"

                file_path = self.output_folder / filename

                # Create pivot tables

                pivot_tables = self.create_pivot_tables(df)

                # Write to Excel with multiple sheets

                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:

                    # Raw Data tab

                    df.to_excel(writer, sheet_name='Raw Data', index=False)

                    # Pivot Tables

                    for pivot_name, pivot_df in pivot_tables.items():

                        sheet_name = pivot_name[:31]  # Excel sheet name length limit

                        pivot_df.to_excel(writer, sheet_name=sheet_name)

                print(f"Created file: {filename}")

                print(f"Records for {market_code}: {len(df)}")

            except Exception as e:

                print(f"Error creating file for {market_code}: {str(e)}")

def main():

    # Replace these with your actual paths

    BASE_PATH = r"C:/Users/PeteCastillo/OneDrive - VillageMD\Documents - VMD- Quality Leadership- PHI/Med Adherence Exception File Worklists/"
    OUTPUT_FOLDER = r"C:/Users/PeteCastillo/OneDrive - VillageMD\Desktop/Escalation Python/"

    try:

        # Initialize analyzer and set paths

        analyzer = WorklistAnalyzer()

        analyzer.base_path = Path(BASE_PATH)

        analyzer.output_folder = Path(OUTPUT_FOLDER)

        analyzer.output_folder.mkdir(parents=True, exist_ok=True)

        # Set the date and process files

        analyzer.set_date("11.18")  # Adjust based on your needs

        market_dfs = analyzer.process_worklists()

        if market_dfs:

            analyzer.create_market_files(market_dfs)

    except Exception as e:

        print(f"Main execution error: {str(e)}")

if __name__ == "__main__":

    main() 