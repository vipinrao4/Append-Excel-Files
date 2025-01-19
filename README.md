import pandas as pd
import os

# Input folder containing Excel files
folder_path = r"Input Location "

# Output file path with the file name
output_file = r"Out Put Location"

# Ensure the output directory exists
output_dir = os.path.dirname(output_file)
os.makedirs(output_dir, exist_ok=True)

# Append all Excel files in the folder 
dataframes = [pd.read_excel(os.path.join(folder_path, f)) for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]
appended_df = pd.concat(dataframes, ignore_index=True)

# Save to the output file in Excel format
appended_df.to_excel(output_file, index=False, engine='openpyxl')

print(f"Appended file saved to {output_file}")
