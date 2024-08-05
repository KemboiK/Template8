import os
import pandas as pd

def process_excel_files(root_dir):
    summary = []
    input_path = os.path.join(root_dir)
    print("Navigating to:", input_path)

    if os.path.isdir(input_path):
        for tribunal in os.listdir(input_path):
            tribunal_path = os.path.join(input_path, tribunal)
            print("Checking Tribunal:", tribunal_path)

            if os.path.isdir(tribunal_path):
                for year in range(2021, 2025):
                    year_path = os.path.join(tribunal_path, str(year))
                    print("Checking Year:", year_path)

                    if os.path.isdir(year_path):
                        for month in range(1, 13):
                            month_path = os.path.join(year_path, str(month).zfill(2))  # Ensure month has two digits
                            print("Checking Month:", month_path)

                            if os.path.isdir(month_path):
                                for file in os.listdir(month_path):
                                    if file.endswith('.xlsx') and not file.endswith('_filed.xlsx') and not file.endswith('_deduplicated.xlsx'):
                                        original_file_path = os.path.join(month_path, file)
                                        deduplicated_file_path = os.path.join(month_path, f"{os.path.splitext(file)[0]}_deduplicated.xlsx")
                                        filed_file_path = os.path.join(month_path, f"{os.path.splitext(file)[0]}_filed.xlsx")

                                        print("Processing File:", original_file_path)

                                        try:
                                            df_original = pd.read_excel(original_file_path, header=4)
                                            if df_original.empty:
                                                print(f"File is empty or incorrectly formatted: {original_file_path}")
                                                continue
                                            
                                            relevant_columns = [
                                                'Day', 'Month', 'Year', 'Day.1', 'Month.1', 'Year.1', 'Code',
                                                'No.', '\n(Select)', 'Name of Judge 1 or Magistrate/DR or Kadhi'
                                            ]

                                            if set(relevant_columns).issubset(df_original.columns):
                                                # Create deduplicated file if needed
                                                initial_row_count = len(df_original)
                                                df_deduplicated = df_original.drop_duplicates(subset=relevant_columns, keep='first')
                                                deduplicated_row_count = len(df_deduplicated)

                                                if deduplicated_row_count < initial_row_count and not os.path.exists(deduplicated_file_path):
                                                    df_deduplicated.to_excel(deduplicated_file_path, index=False)

                                                # Add the 'Filed/Not filed' column to the original dataframe
                                                df_original['Filed/Not filed'] = df_original.apply(
                                                    lambda row: 'Filed' if (
                                                        row['Day'] == row['Day.1'] and
                                                        row['Month'] == row['Month.1'] and
                                                        row['Year'] == row['Year.1']
                                                    ) else 'Not filed', axis=1
                                                )
                                                df_original.to_excel(filed_file_path, index=False)
                                            else:
                                                print(f"Relevant columns not found in {original_file_path}. Available columns: {df_original.columns.tolist()}")

                                            try:
                                                not_assigned = df_original['Name of Judge 1 or Magistrate/DR or Kadhi'].str.contains('Not Yet Assigned', na=False).sum()
                                                assigned = len(df_original) - not_assigned
                                                filed_cases = df_original['Filed/Not filed'].str.contains('Filed').sum()

                                                summary.append({
                                                    'Tribunal': tribunal,
                                                    'Year': int(year),
                                                    'Month': int(month),
                                                    'File': f"{tribunal}_{year}_{month}_{file}",
                                                    'Assigned Cases': assigned,
                                                    'Not Assigned Cases': not_assigned,
                                                    'Filed Cases': filed_cases
                                                })
                                            except Exception as e:
                                                print(f"Error occurred while summarizing file {original_file_path}: {e}")
                                        except Exception as e:
                                            print(f"Error occurred while processing file {original_file_path}: {e}")
                                    else:
                                        print(f"Skipping file (already processed or not an Excel file): {file}")
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
