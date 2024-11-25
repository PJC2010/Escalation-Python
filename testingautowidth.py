import pandas as pd

from datetime import datetime, timedelta

from pathlib import Path

import warnings

warnings.simplefilter(action='ignore', category=UserWarning)


class AdherenceComparer:

    def __init__(self):
        """Initialize the AdherenceComparer"""

        self.base_path = None

        self.output_folder = None

        self.current_date = None

        self.previous_date = None

    def set_dates(self, current_date):
        """Set current and calculate previous week's date"""

        try:

            # Set current date

            datetime.strptime(current_date, '%m.%d')

            self.current_date = current_date

            # Calculate previous week's date

            current = datetime.strptime(f"2024.{current_date}", '%Y.%m.%d')

            previous = current - timedelta(days=7)

            self.previous_date = previous.strftime('%m.%d')

            print(f"Current week: {self.current_date}")

            print(f"Previous week: {self.previous_date}")

        except ValueError as e:

            raise ValueError(
                "Invalid date format. Please use MM.DD format (e.g., '11.18')")

    def get_folder_paths(self):
        """Get paths for both weeks' folders"""

        current_folder = self.base_path / f"Week of {self.current_date}"

        previous_folder = self.base_path / f"Week of {self.previous_date}"

        return current_folder, previous_folder

    def get_worklist_files(self, date):
        """Get list of worklist files for a specific date"""

        year = "2024"

        return [

            f"AZ.CO.MI Medication Adherence Worklist File Week of {
                date}.{year}.xlsx",

            f"TX Medication Adherence Worklist File Week of {
                date}.{year}.xlsx",

            f"ATL.KY Medication Adherence Worklist File Week of {
                date}.{year}.xlsx"

        ]

    def read_excel_safely(self, file_path):
        """Safely read Excel file with multiple fallback options"""

        try:

            return pd.read_excel(file_path, engine='openpyxl')

        except Exception as e:

            print(f"Error reading {file_path}: {str(e)}")

            return None

    def process_weekly_data(self, folder_path, date):
        """Process all files for a specific week"""

        worklist_files = self.get_worklist_files(date)

        market_dfs = {}

        # Define columns to keep

        key_columns = [

            'PayerMemberId',

            'MarketCode',

            'Practice Name',

            'Provider',

            'PatientName',

            'Escalation Path',

            'Escalation Resolution',

            'Gap Completed',

            'PDCNbr',

            'LastFillDate',

            'NextFillDate'

        ]

        for file_name in worklist_files:

            file_path = folder_path / file_name

            try:

                if not file_path.exists():

                    print(f"File not found: {file_path}")

                    continue

                df = self.read_excel_safely(file_path)

                if df is None:

                    continue

                # Clean column names

                df.columns = [col.strip() for col in df.columns]

                # Find matching columns (case-insensitive)

                column_mapping = {}

                for col in key_columns:

                    matches = [c for c in df.columns if c.lower() ==
                               col.lower()]

                    if matches:

                        column_mapping[col] = matches[0]

                # Select available columns

                available_columns = [
                    col for col in key_columns if col in column_mapping.keys()]

                df = df[[column_mapping[col] for col in available_columns]]

                df.columns = available_columns

                # Filter for escalation paths

                if 'Escalation Path' in df.columns:

                    filtered_df = df[df['Escalation Path'].isin([

                        'Market/PHO Escalation',

                        'Practice Escalation'

                    ])]

                    # Group by MarketCode

                    for market_code in filtered_df['MarketCode'].unique():

                        if pd.isna(market_code):

                            continue

                        market_df = filtered_df[filtered_df['MarketCode'] == market_code].copy(
                        )

                        market_df['Week'] = date  # Add week identifier

                        if market_code in market_dfs:

                            market_dfs[market_code] = pd.concat(
                                [market_dfs[market_code], market_df])

                        else:

                            market_dfs[market_code] = market_df

            except Exception as e:

                print(f"Error processing {file_name}: {str(e)}")

        return market_dfs

    def compare_weeks(self):
        """Compare data between current and previous week"""

        current_folder, previous_folder = self.get_folder_paths()

        # Process both weeks

        current_data = self.process_weekly_data(
            current_folder, self.current_date)

        previous_data = self.process_weekly_data(
            previous_folder, self.previous_date)

        # Compare and create reports

        for market_code in set(list(current_data.keys()) + list(previous_data.keys())):

            try:

                curr_df = current_data.get(market_code, pd.DataFrame())

                prev_df = previous_data.get(market_code, pd.DataFrame())

                # Create comparison file

                filename = f"Week_{self.current_date}_{
                    market_code}_Comparison.xlsx"

                file_path = self.output_folder / filename

                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:

                    # Weekly Data Tabs

                    if not curr_df.empty:

                        curr_df.to_excel(
                            writer, sheet_name='Current Week', index=False)

                    if not prev_df.empty:

                        prev_df.to_excel(
                            writer, sheet_name='Previous Week', index=False)

                    # New Members Tab (in current but not in previous)

                    if not curr_df.empty and not prev_df.empty:

                        new_members = curr_df[~curr_df['PayerMemberId'].isin(
                            prev_df['PayerMemberId'])]

                        if not new_members.empty:

                            new_members.to_excel(
                                writer, sheet_name='New Members', index=False)

                    # Resolved Members Tab (in previous but not in current)

                    if not curr_df.empty and not prev_df.empty:

                        resolved = prev_df[~prev_df['PayerMemberId'].isin(
                            curr_df['PayerMemberId'])]

                        if not resolved.empty:

                            resolved.to_excel(
                                writer, sheet_name='Resolved Members', index=False)

                    # Weekly Summary Tab

                    summary_data = {

                        'Metric': [

                            'Current Week Total Members',

                            'Previous Week Total Members',

                            'New Members This Week',

                            'Resolved Members',

                            'Net Change',

                            'Current Week Practice Count',

                            'Previous Week Practice Count'

                        ],

                        'Value': [

                            len(curr_df) if not curr_df.empty else 0,

                            len(prev_df) if not prev_df.empty else 0,

                            len(new_members) if 'new_members' in locals() else 0,

                            len(resolved) if 'resolved' in locals() else 0,

                            (len(curr_df) if not curr_df.empty else 0) -

                            (len(prev_df) if not prev_df.empty else 0),

                            curr_df['Practice Name'].nunique(
                            ) if not curr_df.empty else 0,

                            prev_df['Practice Name'].nunique(
                            ) if not prev_df.empty else 0

                        ]

                    }

                    pd.DataFrame(summary_data).to_excel(
                        writer, sheet_name='Weekly Summary', index=False)

                print(f"Created comparison report: {filename}")

            except Exception as e:

                print(f"Error creating comparison for {market_code}: {str(e)}")


def main():

    # Replace with your actual paths

    BASE_PATH = r"C:/Users/PeteCastillo/OneDrive - VillageMD\Documents - VMD- Quality Leadership- PHI/Med Adherence Exception File Worklists/"
    OUTPUT_FOLDER = r"C:/Users/PeteCastillo/OneDrive - VillageMD\Desktop/Escalation Python/"

    try:

        comparer = AdherenceComparer()

        comparer.base_path = Path(BASE_PATH)

        comparer.output_folder = Path(OUTPUT_FOLDER)

        comparer.output_folder.mkdir(parents=True, exist_ok=True)

        # Set current week's date - previous week will be calculated automatically

        comparer.set_dates("11.18")

        # Run comparison

        comparer.compare_weeks()

    except Exception as e:

        print(f"Main execution error: {str(e)}")


if __name__ == "__main__":

    main()


def autofit_columns(self, worksheet, df):
    """

    Automatically adjust column widths to fit content

    Args:

        worksheet: openpyxl worksheet object

        df: pandas DataFrame containing the data

    """

    for idx, column in enumerate(df.columns):

        column_letter = get_column_letter(idx + 1)

        # Get maximum length of column content

        max_length = 0

        # Check column header length

        header_length = len(str(column))

        # Check content length for each cell in column

        content_length = df[column].astype(str).str.len().max()

        # Use the maximum of header and content length

        max_length = max(header_length, content_length)

        # Add some padding

        adjusted_width = max_length + 4

        # Set a minimum width

        adjusted_width = max(adjusted_width, 8)

        # Set a maximum width to prevent excessive column widths

        adjusted_width = min(adjusted_width, 50)

        worksheet.column_dimensions[column_letter].width = adjusted_width

# Then modify the create_market_files method to use this:


def create_market_files(self, market_dfs):
    """Create separate files for each market with pivot tables"""

    if not market_dfs:

        print("No data available to create files")

        return

    next_monday = self.get_next_monday()

    for market_code, df in market_dfs.items():

        try:

            filename = f"{next_monday} {market_code} Escalations.xlsx"

            file_path = self.output_folder / filename

            print(f"\nProcessing market: {market_code}")

            print(f"Creating file: {filename}")

            print(f"Total records: {len(df)}")

            # Create pivot tables

            pivot_tables = self.create_pivot_tables(df)

            # Write to Excel with multiple sheets

            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:

                # Raw Data tab

                df.to_excel(writer, sheet_name='Raw Data', index=False)

                self.autofit_columns(writer.sheets['Raw Data'], df)

                # Pivot Tables

                for pivot_name, pivot_df in pivot_tables.items():

                    # Excel sheet name length limit
                    sheet_name = pivot_name[:31]

                    pivot_df.to_excel(writer, sheet_name=sheet_name)

                    print(f"- Created '{sheet_name}' sheet")

                    # Autofit columns for pivot tables

                    self.autofit_columns(writer.sheets[sheet_name], pivot_df)

            print(f"\nSuccessfully created file: {file_path}")

        except Exception as e:

            print(f"\nError creating file for {market_code}: {str(e)}")

            import traceback

            print(traceback.format_exc())

# And for the comparison script, modify the compare_weeks method:


def compare_weeks(self):
    """Compare data between current and previous week"""

    current_folder, previous_folder = self.get_folder_paths()

    # Process both weeks

    current_data = self.process_weekly_data(current_folder, self.current_date)

    previous_data = self.process_weekly_data(
        previous_folder, self.previous_date)

    # Compare and create reports

    for market_code in set(list(current_data.keys()) + list(previous_data.keys())):

        try:

            curr_df = current_data.get(market_code, pd.DataFrame())

            prev_df = previous_data.get(market_code, pd.DataFrame())

            # Create comparison file

            filename = f"Week_{self.current_date}_{
                market_code}_Comparison.xlsx"

            file_path = self.output_folder / filename

            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:

                # Weekly Data Tabs

                if not curr_df.empty:

                    curr_df.to_excel(
                        writer, sheet_name='Current Week', index=False)

                    self.autofit_columns(
                        writer.sheets['Current Week'], curr_df)

                if not prev_df.empty:

                    prev_df.to_excel(
                        writer, sheet_name='Previous Week', index=False)

                    self.autofit_columns(
                        writer.sheets['Previous Week'], prev_df)

                # New Members Tab

                if not curr_df.empty and not prev_df.empty:

                    new_members = curr_df[~curr_df['PayerMemberId'].isin(
                        prev_df['PayerMemberId'])]

                    if not new_members.empty:

                        new_members.to_excel(
                            writer, sheet_name='New Members', index=False)

                        self.autofit_columns(
                            writer.sheets['New Members'], new_members)

                # Resolved Members Tab

                if not curr_df.empty and not prev_df.empty:

                    resolved = prev_df[~prev_df['PayerMemberId'].isin(
                        curr_df['PayerMemberId'])]

                    if not resolved.empty:

                        resolved.to_excel(
                            writer, sheet_name='Resolved Members', index=False)

                        self.autofit_columns(
                            writer.sheets['Resolved Members'], resolved)

                # Weekly Summary Tab

                summary_data = {

                    'Metric': [

                        'Current Week Total Members',

                        'Previous Week Total Members',

                        'New Members This Week',

                        'Resolved Members',

                        'Net Change',

                        'Current Week Practice Count',

                        'Previous Week Practice Count'

                    ],

                    'Value': [

                        len(curr_df) if not curr_df.empty else 0,

                        len(prev_df) if not prev_df.empty else 0,

                        len(new_members) if 'new_members' in locals() else 0,

                        len(resolved) if 'resolved' in locals() else 0,

                        (len(curr_df) if not curr_df.empty else 0) -

                        (len(prev_df) if not prev_df.empty else 0),

                        curr_df['Practice Name'].nunique(
                        ) if not curr_df.empty else 0,

                        prev_df['Practice Name'].nunique(
                        ) if not prev_df.empty else 0

                    ]

                }

                summary_df = pd.DataFrame(summary_data)

                summary_df.to_excel(
                    writer, sheet_name='Weekly Summary', index=False)

                self.autofit_columns(
                    writer.sheets['Weekly Summary'], summary_df)

            print(f"Created comparison report: {filename}")

        except Exception as e:

            print(f"Error creating comparison for {market_code}: {str(e)}")
