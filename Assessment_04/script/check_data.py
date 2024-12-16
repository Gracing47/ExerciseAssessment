import pandas as pd

# Read the Excel file
df = pd.read_excel('/Users/tommygrace/Downloads/Assesment/Assessment_03/20230227_Novum_Assesment_3_BADS.xlsx', skiprows=2)

# Rename columns
df.columns = ['call_type', 'agent', 'total_calls', 'success_rate', 'total_success', 'hours_online', 'avg_duration']

# Display basic information about the dataframe
print("\nDataframe Info:")
print(df.info())

print("\nColumn Names:")
print(df.columns.tolist())

print("\nFirst few rows:")
print(df)

print("\nData types for each column:")
for col in df.columns:
    print(f"\n{col}:")
    print(df[col].head())
