import pandas as pd
import matplotlib as plt
import seaborn as sns
import os
import matplotlib.pyplot as plt

from datetime import datetime, timedelta

from pathlib import Path

import warnings

from openpyxl.utils import get_column_letter

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
            raise ValueError(
                "Invalid date format. Please use MM.DD format (e.g., '11.18')")

    def get_this_monday(self):
        today = datetime.now()
        monday = today - timedelta(days=today.weekday())
        return monday.strftime('%m.%d')

    def get_next_monday(self):
        """Get the date of the upcoming Monday in mm.dd format"""
        today = datetime.now()
        days_ahead = 7 - today.weekday()
        next_monday = today + timedelta(days=days_ahead)
        return next_monday.strftime('%m.%d')

    def find_week_folder_by_date(self, date_str):
        """
        Find a folder that contains 'Week of' and the specified date
        
        Args:
            date_str (str): Date string in 'MM.DD' format (e.g., '11.18')
        
        Returns:
            Path: Path to the folder if found, None otherwise
        """
        if not self.base_path.exists():
            print(f"Base path does not exist: {self.base_path}")
            return None
        
        # Format the pattern to search for
        search_pattern = f"Week of {date_str}"
        
        # Search for folders containing the pattern
        matching_folders = [folder for folder in self.base_path.iterdir() 
                           if folder.is_dir() and search_pattern in folder.name]
        
        if not matching_folders:
            print(f"No folder found with pattern: '{search_pattern}'")
            return None
        
        if len(matching_folders) > 1:
            print(f"Warning: Multiple folders found with pattern '{search_pattern}'. Using the first one.")
        
        folder = matching_folders[0]
        print(f"Found folder: {folder.name}")
        return folder

    def get_week_folder(self):
        """Get the week folder path based on set date"""
        if self.current_date is None:
            self.current_date = datetime.now().strftime('%m.%d')
            print(f"No date set, using current date: {self.current_date}")
        
        # Try to find the folder with the specified date
        folder_path = self.find_week_folder_by_date(self.current_date)
        
        # If folder not found, check if the old-style folder name exists
        if folder_path is None:
            folder_name = f"Week of {self.current_date}"
            folder_path = self.base_path / folder_name
            
            if folder_path.exists():
                print(f"Using folder: {folder_path}")
            else:
                print(f"Warning: Folder not found: {folder_path}")
        
        return folder_path

    def find_excel_files_in_folder(self, folder_path):
        """
        Find all Excel files in the specified folder
        
        Args:
            folder_path (Path): Path to the folder to search
            
        Returns:
            list: List of paths to Excel files
        """
        if not folder_path.exists():
            print(f"Folder not found: {folder_path}")
            return []
        
        # Find all Excel files (.xlsx, .xls)
        excel_files = [file for file in folder_path.iterdir() 
                      if file.is_file() and file.suffix.lower() in ['.xlsx', '.xls']]
        
        if not excel_files:
            print(f"No Excel files found in: {folder_path}")
        else:
            print(f"Found {len(excel_files)} Excel files in: {folder_path}")
        
        return excel_files

    def read_excel_safely(self, file_path):
        """Safely read Excel file with multiple fallback options"""
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            return df
        except Exception as e1:
            try:
                df = pd.read_excel(
                    file_path, engine='openpyxl', data_only=True)
                return df
            except Exception as e2:
                try:
                    df = pd.read_excel(file_path, engine='xlrd')
                    return df
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
            pivots['Practice_Escalations'] = pd.pivot_table(
                df,
                index='PracticeName',
                columns='Escalation Path',
                values='PayerMemberId',
                aggfunc='count',
                fill_value=0,
                margins=True,
                margins_name='Total'
            ).sort_values('Total', ascending=False)

            # Pivot 2: Provider Summary
            pivots['Provider_Escalations'] = pd.pivot_table(
                df,
                index='PCP',
                columns='Escalation Path',
                values='PayerMemberId',
                aggfunc='count',
                fill_value=0,
                margins=True,
                margins_name='Total'
            ).sort_values('Total', ascending=False)

            # Create summary
            summary_data = {
                'Metric': [
                    'Total Escalations',
                    'Market/PHO Escalations',
                    'Practice Escalations',
                    'Unique Practices',
                    'Unique Providers',
                    'Report Generated'
                ],
                'Value': [
                    len(df),
                    len(df[df['Escalation Path'] == 'Market/PHO Escalation']),
                    len(df[df['Escalation Path'] == 'Practice Escalation']),
                    df['PracticeName'].nunique(),
                    df['PCP'].nunique(),
                    datetime.now().strftime('%Y-%m-%d %H:%M')
                ]
            }
            pivots['Summary'] = pd.DataFrame(summary_data)
        except Exception as e:
            print(f"Error creating pivot tables: {str(e)}")
        return pivots
    
  
    def create_practice_visualization(self, pivot_df, market_code, output_folder):
        try:
            # Remove the 'Total' row and column for visualization
            viz_df = pivot_df.drop('Total', axis=0).drop('Total', axis=1)

            # Calculate figure height based on number of practices
            num_practices = len(viz_df)
            fig_height = max(8, num_practices * 0.4)

            plt.figure(figsize=(8, fig_height))
            ax = viz_df.plot(kind='barh', stacked=True)

            # Customize the plot
            plt.title(f'{market_code} Practice Escalations', pad=20, fontsize=9)
            plt.xlabel('# of Escalations', fontsize=6)
            plt.ylabel('')  

            # Format practice names with smaller font size
            ax.set_yticklabels(viz_df.index, wrap=True, fontsize=4.5)  # Reduced font size here
            ax.yaxis.set_ticks_position('none')

            # Format x-axis labels
            plt.xticks(ticks=[], fontsize=9)

            # Add value labels on the bars
            for c in ax.containers:
                ax.bar_label(c, label_type='center', fontsize=5)  # Made value labels smaller too

            # Add legend with smaller font
            plt.legend(title='Escalation Type', 
                  bbox_to_anchor=(1.05, 1), 
                  loc='upper left', 
                  fontsize=8,  # Legend text size
                  title_fontsize=9)  # Legend title size

            # Adjust layout
            plt.tight_layout()

            # Save the plot
            plot_filename = f"4.28.25 {market_code}_Practice_Escalations.png"
            plt.savefig(output_folder / plot_filename, dpi=300, bbox_inches='tight')
            plt.close()

            print(f"Created visualization: {plot_filename}")

        except Exception as e:
            print(f"Error creating visualization for {market_code}: {str(e)}") 
  
            try:
                today = datetime.now().strftime('%m.%d')
       
                viz_df = pivot_df.drop('Total', axis=0).drop('Total', axis=1)
       
                num_practices = len(viz_df)
       
                fig_height = max(8, num_practices * 0.4)
                plt.figure(figsize=(12, fig_height))
                ax = viz_df.plot(kind='barh', stacked=True)
       
                plt.title(f'{market_code} Practice Escalations', pad=20)
                plt.xlabel('Number of Escalations')
                plt.ylabel('')  # Remove ylabel since practice names are on y-axis
       
                ax.set_yticklabels(viz_df.index, wrap=True)
       
                for c in ax.containers:
                    ax.bar_label(c, label_type='center')
       
                plt.legend(title='Escalation Type', bbox_to_anchor=(1.05, 1), loc='upper left')
        
                plt.tight_layout()
                
        
                plot_filename = f"{today}_{market_code}_Practice_Escalations.png"
                plt.savefig(output_folder / plot_filename, dpi=300, bbox_inches='tight')
                plt.close()
                print(f"Created visualization: {plot_filename}") 
            except Exception as e: 
                print(f"Error in creating visualization for {market_code}: {str(e)}")
   
    def process_worklists(self):
        """Process all Excel files in the week folder and return data by market"""
        folder_path = self.get_week_folder()
        if folder_path is None or not folder_path.exists():
            print("No valid week folder found.")
            return {}
        
        # Find all Excel files in the folder
        excel_files = self.find_excel_files_in_folder(folder_path)
        
        # Define the columns to keep in the desired order
        desired_columns = [
            'LastImpactableDate',
            'PatientName',
            'DateOfBirth',
            'PracticeName',
            'PCP',
            'Rx Status',
            'Call Disposition',
            'QS Notes',
            'Current Barrier',
            'Action',
            'Escalation Path',
            'Escalation Timeframe',
            'Escalation Deadline',
            'Escalation Resolution',
            'PayerCode',
            'MarketCode',
            'PayerMemberId',
            'PatientPhoneNumber',
            'PatientAddress',
            'DataAsOfDate',
            'EMR ID',
            'United Flag',
            'MedAdherenceMeasureCode',
            'NDCDesc',
            'Impact Category',
            'Gap Priority',
            'PDCNbr',
            'ADRNbr',
            'DaysMissedNbr',
            'Total Fills Column?',
            'Initial Fill Date',
            'LastFillDate',
            'NextFillDate',
            'DrugDispensedQuantityNbr',
            'DrugDispensedDaysSupplyNbr',
            'Last Activity Date',
            'Task Status'
            'OneFillCode',
            'PrescriberNPI',
            'PrescribingName',
            'Prescriber Phone Number',
            'PharmacyStoreName',
            'PharmacyCommunicationNumberText'
        ]

        # Define date columns for formatting
        date_columns = [
            'LastImpactableDate',
            'DateOfBirth',
            'LastFillDate',
            'NextFillDate',
            'Initial Fill Date',
            'Last Activity Date',
            'DataAsOfDate',
            'Escalation Timeframe',
            'Escalation Deadline'
        ]

        market_dfs = {}

        for file_path in excel_files:
            try:
                print(f"\nProcessing file: {file_path.name}")
                
                # Get all sheet names in the Excel file
                excel = pd.ExcelFile(file_path)
                sheet_names = excel.sheet_names
                
                # Skip sheet named "Validation_Lists"
                valid_sheets = [sheet for sheet in sheet_names if sheet != "Validation_Lists"]
                
                if not valid_sheets:
                    print(f"No valid sheets found in: {file_path.name}")
                    continue
                    
                print(f"Processing {len(valid_sheets)} sheets from: {file_path.name}")
                
                # Process each sheet
                for sheet_name in valid_sheets:
                    try:
                        print(f"  Reading sheet: {sheet_name}")
                        
                        # Read the sheet
                        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
                        
                        # If sheet is empty, skip it
                        if df.empty:
                            print(f"  Sheet '{sheet_name}' is empty, skipping.")
                            continue
                            
                        # Clean up column names
                        df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
                        
                        # Find the correct column names (case-insensitive)
                        column_mapping = {}
                        for desired_col in desired_columns:
                            matches = [col for col in df.columns 
                                      if isinstance(col, str) and col.lower() == desired_col.lower()]
                            if matches:
                                column_mapping[desired_col] = matches[0]
                        
                        # If essential columns are missing, skip this sheet
                        required_columns = ['Escalation Path', 'MarketCode']
                        missing_required = [col for col in required_columns if col not in column_mapping]
                        
                        if missing_required:
                            print(f"  Sheet '{sheet_name}' is missing required columns: {missing_required}, skipping.")
                            continue
                        
                        # Select and reorder columns that exist in the file
                        available_columns = [col for col in desired_columns if col in column_mapping.keys()]
                        
                        # If no available columns match, skip this sheet
                        if not available_columns:
                            print(f"  Sheet '{sheet_name}' has no matching columns, skipping.")
                            continue
                            
                        # Select only columns that exist in the file
                        df_selected = df[[column_mapping[col] for col in available_columns]]
                        
                        # Rename columns to desired names
                        df_selected.columns = available_columns
                        
                        # Format date columns
                        for col in date_columns:
                            if col in df_selected.columns:
                                try:
                                    df_selected[col] = pd.to_datetime(df_selected[col], errors='coerce').dt.strftime('%m/%d/%Y')
                                except Exception as e:
                                    print(f"  Could not format date column {col}: {str(e)}")
                        
                        # Filter for escalation paths
                        if 'Escalation Path' in df_selected.columns:
                            filtered_df = df_selected[df_selected['Escalation Path'].isin([
                                'Market/PHO Escalation',
                                'Practice Escalation'
                            ])]
                            
                            if len(filtered_df) > 0:
                                # Group by MarketCode
                                for market_code in filtered_df['MarketCode'].unique():
                                    if pd.isna(market_code):
                                        continue
                                        
                                    market_df = filtered_df[filtered_df['MarketCode'] == market_code]
                                    
                                    if market_code in market_dfs:
                                        market_dfs[market_code] = pd.concat([market_dfs[market_code], market_df])
                                    else:
                                        market_dfs[market_code] = market_df
                                        
                                print(f"  Successfully processed sheet '{sheet_name}' with {len(filtered_df)} records.")
                            else:
                                print(f"  No matching escalations found in sheet '{sheet_name}'")
                        else:
                            print(f"  'Escalation Path' column not found in sheet '{sheet_name}'")
                    
                    except Exception as e:
                        print(f"  Error processing sheet '{sheet_name}': {str(e)}")
                
                print(f"Completed processing file: {file_path.name}")
                
            except Exception as e:
                print(f"Error processing file {file_path.name}: {str(e)}")
                import traceback
                print(traceback.format_exc())

        # Print summary of data collected
        print("\nData collection summary:")
        total_records = 0
        for market_code, df in market_dfs.items():
            records = len(df)
            total_records += records
            print(f"  Market {market_code}: {records} records")
        print(f"Total records collected: {total_records}")
        
        return market_dfs
    
    def create_market_files(self, market_dfs):
        """Create separate files for each market with pivot tables"""
        if not market_dfs:
            print("No data available to create files")
            return

        next_monday = self.get_next_monday()
        this_monday = self.get_this_monday()

        for market_code, df in market_dfs.items():
            try:
                filename = f"{this_monday} {market_code} Med Adherence Escalations.xlsx"
                file_path = self.output_folder / filename

                print(f"\nProcessing market: {market_code}")
                print(f"Creating file: {filename}")
                print(f"Total records: {len(df)}")

                # Create pivot tables
                pivot_tables = self.create_pivot_tables(df)

                # Write to Excel with multiple sheets
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    # Raw Data tab
                    df.to_excel(writer, sheet_name=f"{market_code}", index=False)

                    # Format Raw Data sheet
                    worksheet = writer.sheets[f"{market_code}"]
                    for idx, column in enumerate(df.columns):
                        col_letter = get_column_letter(idx + 1)
                        max_length = max(
                            df[column].astype(str).apply(len).max(),
                            len(str(column))
                        )
                        adjusted_width = (max_length + 2) * 1.2
                        worksheet.column_dimensions[col_letter].width = adjusted_width

                    # Pivot Tables
                    for pivot_name, pivot_df in pivot_tables.items():
                        sheet_name = pivot_name[:31]
                        pivot_df.to_excel(writer, sheet_name=sheet_name)
                        print(f"- Created '{sheet_name}' sheet")

                        # Format pivot sheets
                        pivot_worksheet = writer.sheets[sheet_name]
                        for column in pivot_worksheet.columns:
                            max_length = 0
                            column = [cell for cell in column]
                            try:
                                max_length = max(
                                    len(str(cell.value)) for cell in column if cell.value
                                )
                            except:
                                pass
                            adjusted_width = (max_length + 2)
                            pivot_worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
    
                print(f"Successfully created file: {file_path}")
                if 'Practice_Escalations' in pivot_tables:
                    self.create_practice_visualization(
                        pivot_tables['Practice_Escalations'], 
                        market_code,
                        self.output_folder
                    )

            except Exception as e:
                print(f"\nError creating file for {market_code}: {str(e)}")
                import traceback
                print(traceback.format_exc())


def main():
    # Replace these with your actual paths
    BASE_PATH = r"C:/Users/pcastillo/OneDrive - VillageMD\Documents - VMD- Quality Leadership- PHI/Med Adherence Exception File Worklists/"
    OUTPUT_FOLDER = r"C:/Users/pcastillo/OneDrive - VillageMD\Documents - VMD- Quality Leadership- PHI/Data Updates/MedAdhData Dropzone/Output/EscalationsDropZone/"
    #OUTPUT_FOLDER = r"C:/Users/pcastillo/OneDrive - VillageMD\Desktop/Escalation Python/"
    
    try:
        # Initialize analyzer and set paths
        analyzer = WorklistAnalyzer()
        analyzer.base_path = Path(BASE_PATH)
        analyzer.output_folder = Path(OUTPUT_FOLDER)
        analyzer.output_folder.mkdir(parents=True, exist_ok=True)
        
        # Set the date you're interested in
        analyzer.set_date("4.21")  
        
        # Process worklists for the specified date
        market_dfs = analyzer.process_worklists()
        
        if market_dfs:
            analyzer.create_market_files(market_dfs)
            print("Processing complete!")
        else:
            print("No data was found to process")
    except Exception as e:
        print(f"Main execution error: {str(e)}")


if __name__ == "__main__":
    main()