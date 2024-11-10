import os
import pandas as pd

def convert_excel_to_csv_in_current_directory():
    # Get the current directory where the script is being executed
    current_dir = os.getcwd()
    
    # Find all Excel files in the current directory
    excel_files = [f for f in os.listdir(current_dir) if f.endswith('.xlsx') or f.endswith('.xls')]
    
    # Convert each Excel file to CSV
    for file in excel_files:
        file_path = os.path.join(current_dir, file)
        try:
            # Load the Excel file
            workbook = pd.ExcelFile(file_path)
            
            # Process each sheet in the workbook
            for sheet_name in workbook.sheet_names:
                # Read the sheet into a DataFrame
                df = workbook.parse(sheet_name)
                
                # Define the CSV file name based on the Excel file name and sheet name
                csv_file_name = f"{os.path.splitext(file)[0]}_{sheet_name}.csv"
                csv_file_path = os.path.join(current_dir, csv_file_name)
                
                # Save the DataFrame to a CSV file
                df.to_csv(csv_file_path, index=False)
                print(f"Converted {file} - sheet '{sheet_name}' to {csv_file_name}")
        
        except Exception as e:
            print(f"Failed to convert {file}: {e}")

# Run the function
convert_excel_to_csv_in_current_directory()
