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
                                    if file.endswith('.xlsx'):
                                        file_path = os.path.join(month_path, file)
                                        print("Processing File:", file_path)  # Debugging print

                                        try:
                                            df = pd.read_excel(file_path, header=4)
                                            if 'Name of Judge 1 or Magistrate/DR or Kadhi' in df.columns:
                                                not_assigned = df['Name of Judge 1 or Magistrate/DR or Kadhi'].str.contains('Not Yet Assigned', na=False).sum()
                                                assigned = len(df) - not_assigned
                                                summary.append({
                                                    'Tribunal': tribunal,
                                                    'Year': year,
                                                    'Month': month,
                                                    'File': file,
                                                    'Assigned Cases': assigned,
                                                    'Not Assigned Cases': not_assigned
                                                })
                                            else:
                                                print(f"Column not found in {file_path}. Available columns: {df.columns.tolist()}")
                                        except Exception as e:
                                            print(f"Error processing file {file_path}: {e}")
    else:
        print(f"Directory not found: {input_path}")
    return pd.DataFrame(summary)

# Define the root directory path
root_dir = 'INPUT\TEMPLATE 8'  

# Process the Excel files and get the summary
summary_df = process_excel_files(root_dir)

# Display the summary DataFrame
print(summary_df)
summary_df.to_csv('summary.csv', index=False)