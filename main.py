import pandas as pd
import os
import subprocess

# Specify the path to your Excel file
excel_file_path = 'excel.xlsx'

# Load the Excel file
xls = pd.ExcelFile(excel_file_path)

# Create output directory if it doesn't exist
output_directory = './output_json/'
os.makedirs(output_directory, exist_ok=True)

# Git commit message
commit_message = 'Update Excel data'

# Function to commit changes to the Git repository
def git_commit():
    try:
        subprocess.run(['git', 'add', '.'], check=True)
        subprocess.run(['git', 'commit', '-m', commit_message], check=True)
        subprocess.run(['git', 'push'], check=True)
        print(f'Changes committed and pushed to the Git repository.')
    except subprocess.CalledProcessError as e:
        print(f'Error during Git commit: {e}')

# Iterate through each sheet and convert to JSON
for sheet_name in xls.sheet_names:
    # Read the sheet into a DataFrame
    df = pd.read_excel(excel_file_path, sheet_name)
    
    # Convert DataFrame to JSON
    json_output = df.to_json(orient='records', indent=2)
    
    # Write JSON to a file
    json_file_path = os.path.join(output_directory, f'{sheet_name}.json')
    with open(json_file_path, 'w') as json_file:
        json_file.write(json_output)
    
    print(f'Conversion for {sheet_name} complete. JSON file saved at {json_file_path}')

# Commit changes to the Git repository
git_commit()
