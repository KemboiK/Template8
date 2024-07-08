# Template8
## Excel Files Processor
This script processes Excel files organized in a specific directory structure and summarizes the number of assigned and not assigned cases. The script navigates through a root directory, checks for specific years and months, and processes Excel files within these directories to extract relevant data.

## Directory Structure
The script expects a directory structure that moves from folder name tribunal to subfolder input to subfolder template8 to sufolders the tribunal names each with years 2021 to 2024 and each year with months 1 to 12 in each month is where the excel files are located
## Requirements
- Python 3.x
- pandas
You can install the necessary Python packages using:
pip install pandas
## Usage
1. Clone the repository:
git clone https://github.com/yourusername/excel-files-processor.git
cd excel-files-processor
2. Update the root_dir variable in the script to point to your root directory containing the Excel files.

3. Run the script:
python process_excel_files.py
The script will navigate through the directory structure, process the Excel files, and generate a summary CSV file named summary.csv.

## Script Details
The script performs the following steps:

1. Navigates through the provided root directory.
2. Checks for subdirectories corresponding to tribunals.
3. Loops through the years 2021 to 2024.
4. Loops through the months 1 to 12.
5. Processes Excel files found in the directories.
6. Extracts data from the column 'Name of Judge 1 or Magistrate/DR or Kadhi'.
7. Counts the number of cases that are assigned and not assigned.
8. Appends the summary data to a list and converts it to a DataFrame.
9. Outputs the summary DataFrame to a CSV file named summary.csv.
## Example Output
The output CSV file (summary.csv) will have the following columns:

- Tribunal
- Year
- Month
- File
- Assigned Cases
- Not Assigned Cases
## Debugging
The script includes several print statements to help with debugging and to track the progress of the script. These statements print information about the directories and files being processed.

## Contributions
Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.