import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# Read the Excel file
df = pd.read_excel('20230227_Novum_Assesment_1_BADS.xlsx')

# Create output directory
import os
if not os.path.exists('analysis_output'):
    os.makedirs('analysis_output')

# 1. Time-based analysis
df['Date'] = df['DateTime'].dt.date
daily_actions = df.groupby('Date').size()

plt.figure(figsize=(15, 6))
daily_actions.plot(kind='line', marker='o')
plt.title('Daily Activity Volume')
plt.xlabel('Date')
plt.ylabel('Number of Actions')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('analysis_output/daily_activity.png')
plt.close()

# 2. Action Type Distribution
plt.figure(figsize=(12, 6))
df['ActionType'].value_counts().plot(kind='bar')
plt.title('Distribution of Action Types')
plt.xlabel('Action Type')
plt.ylabel('Count')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('analysis_output/action_distribution.png')
plt.close()

# 3. Customer Type Analysis (excluding null values)
plt.figure(figsize=(10, 6))
df['CustomerType'].value_counts().plot(kind='pie', autopct='%1.1f%%')
plt.title('Customer Type Distribution')
plt.axis('equal')
plt.savefig('analysis_output/customer_distribution.png')
plt.close()

# 4. Country Code Analysis
plt.figure(figsize=(12, 6))
df['CountryCode'].value_counts().plot(kind='bar')
plt.title('Distribution by Country Code')
plt.xlabel('Country Code')
plt.ylabel('Count')
plt.tight_layout()
plt.savefig('analysis_output/country_distribution.png')
plt.close()

# Generate summary report
with open('analysis_output/detailed_summary.txt', 'w') as f:
    f.write("Detailed Analysis Summary\n")
    f.write("========================\n\n")
    
    f.write("1. Time Period Analysis\n")
    f.write(f"Start Date: {df['DateTime'].min()}\n")
    f.write(f"End Date: {df['DateTime'].max()}\n")
    f.write(f"Total Days: {(df['DateTime'].max() - df['DateTime'].min()).days}\n\n")
    
    f.write("2. Action Types Summary\n")
    f.write(df['ActionType'].value_counts().to_string())
    f.write("\n\n")
    
    f.write("3. Customer Types Summary\n")
    f.write(df['CustomerType'].value_counts().to_string())
    f.write("\n\n")
    
    f.write("4. Country Distribution\n")
    f.write(df['CountryCode'].value_counts().to_string())
    f.write("\n\n")
    
    f.write("5. Change Types Summary\n")
    f.write(df['Change'].value_counts().head().to_string())
    f.write("\n\n")

print("Detailed analysis complete! Check the 'analysis_output' directory for results.")
