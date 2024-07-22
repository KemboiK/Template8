import os
import pandas as pd

def process_excel_files(root_dir):
    summary = []
    input_path = os.path.join(root_dir)  # Directly use the provided root_dir
    print("Navigating to:", input_path)  # Debugging print

    if os.path.isdir(input_path):
        for tribunal in os.listdir(input_path):
            tribunal_path = os.path.join(input_path, tribunal)
            print("Checking Tribunal:", tribunal_path)  # Debugging print

            if os.path.isdir(tribunal_path):
                for year in range(2021, 2025):  # Loop through years 2021 to 2024
                    year_path = os.path.join(tribunal_path, str(year))
                    print("Checking Year:", year_path)  # Debugging print

                    if os.path.isdir(year_path):
                        for month in range(1, 13):  # Loop through months 1 to 12
                            month_path = os.path.join(year_path, str(month))
                            print("Checking Month:", month_path)  # Debugging print

                            if os.path.isdir(month_path):
                                for file in os.listdir(month_path):
                                    if file.endswith('.xlsx') and not file.endswith('_deduplicated.xlsx'):
                                        file_path = os.path.join(month_path, file)
                                        # Define the deduplicated file path
                                        new_file_path = file_path.replace('.xlsx', '_deduplicated.xlsx')

                                        print("Processing File:", file_path)  # Debugging print

                                        try:
                                            df = pd.read_excel(file_path, header=4)

                                            # Check if the DataFrame is empty or not
                                            if df.empty:
                                                print(f"File is empty or incorrectly formatted: {file_path}")
                                                continue

                                            # Assuming column names based on your description
                                            relevant_columns = [
                                                'Day', 'Month', 'Year', 'Day.1', 'Month.1', 'Year.1','Code'
                                                'No.', '\n(Select)', 'Name of Judge 1 or Magistrate/DR or Kadhi'
                                            ]

                                            # Drop duplicates based on the relevant columns
                                            if set(relevant_columns).issubset(df.columns):
                                                initial_row_count = len(df)
                                                df.drop_duplicates(subset=relevant_columns, keep='first', inplace=True)
                                                deduplicated_row_count = len(df)

                                                # Only save a new file if there are duplicates to remove
                                                if deduplicated_row_count < initial_row_count:
                                                    df.to_excel(new_file_path, index=False)  # Save to Excel
                                                    summary_file = new_file_path  # Track deduplicated file
                                                else:
                                                    # Remove new deduplicated file if no changes were made
                                                    if os.path.exists(new_file_path):
                                                        os.remove(new_file_path)
                                                    summary_file = file_path  # Track original file
                                            else:
                                                print(f"Relevant columns not found in {file_path}. Available columns: {df.columns.tolist()}")
                                                summary_file = file_path  # Track original file
                                        except Exception as e:
                                            print(f"Error occurred while processing file {file_path}: {e}")
                                            summary_file = file_path  # Track original file

                                        # Append to summary
                                        try:
                                            df_summary = pd.read_excel(summary_file)
                                            not_assigned = df_summary['Name of Judge 1 or Magistrate/DR or Kadhi'].str.contains('Not Yet Assigned', na=False).sum()
                                            assigned = len(df_summary) - not_assigned
                                            
                                            # Extract original file name for summary
                                            original_file_name = file.replace('.xlsx', '_deduplicated.xlsx') if '_deduplicated.xlsx' in file else file

                                            summary.append({
                                                'Tribunal': tribunal,
                                                'Year': int(year),
                                                'Month': int(month),
                                                'File': original_file_name,
                                                'Assigned Cases': assigned,
                                                'Not Assigned Cases': not_assigned,
                                            })
                                        except Exception as e:
                                            print(f"Error occurred while summarizing file {summary_file}: {e}")
    else:
        print(f"Directory not found: {input_path}")
    return pd.DataFrame(summary)

# Define the root directory path
root_dir = 'INPUT\\TEMPLATE 8'

# Process the Excel files and get the summary
summary_df = process_excel_files(root_dir)

# Display the summary DataFrame and save it as CSV
print(summary_df)
summary_df.to_csv('summary.csv', index=False)
