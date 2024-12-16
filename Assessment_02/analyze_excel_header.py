import pandas as pd

# Path to the Excel file
file_path = '/Users/tommygrace/Downloads/Assesment/Assessment_02/20230227_Novum_Assesment_1_BADS.xlsx'

# Read the Excel file
try:
    # Load the first sheet of the Excel file
    df = pd.read_excel(file_path, sheet_name=0)
    
    # Display the header of the DataFrame
    print("Header of the Excel file:")
    print(df.head())
except Exception as e:
    print(f"An error occurred while reading the Excel file: {e}")
