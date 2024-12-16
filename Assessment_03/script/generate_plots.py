import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Path to the Excel file
file_path = '/Users/tommygrace/Downloads/Assesment/Assessment_03/20230227_Novum_Assesment_2_BADS.xlsx'

def calculate_agent_availability(value):
    if pd.isna(value):
        return 1  # Fully available
    elif isinstance(value, (int, float)):
        return value  # Partial availability
    elif value.lower() == 'sick' or value.lower() == 'leave':
        return 0  # Not available
    else:
        try:
            return float(value)
        except:
            return 0

try:
    # Load the Excel file
    df = pd.read_excel(file_path)
    
    # Get agent columns (excluding date and day of week)
    agent_columns = df.columns[2:]
    
    # Calculate daily availability considering partial availability
    daily_availability = df[agent_columns].applymap(calculate_agent_availability).sum(axis=1)
    
    # Create date range for x-axis
    dates = pd.to_datetime(df['Unnamed: 0'].iloc[1:])  # Skip the header row
    
    # Plot daily availability
    plt.figure(figsize=(12, 6))
    plt.plot(range(len(daily_availability[1:])), daily_availability[1:], marker='o', linestyle='-', linewidth=2)
    plt.axhline(y=24.5, color='r', linestyle='--', label='Required (24.5)')
    plt.title('Daily Team Availability (August 2022)', pad=20)
    plt.xlabel('Days in August')
    plt.ylabel('Number of Available Agents')
    plt.grid(True, alpha=0.3)
    plt.legend()
    
    # Format x-axis
    plt.xticks(range(len(dates)), [d.strftime('%d-%b') for d in dates], rotation=45)
    
    plt.tight_layout()
    plt.savefig('/Users/tommygrace/Downloads/Assesment/Assessment_03/output/daily_availability_new.png', dpi=300, bbox_inches='tight')
    plt.close()

    # Plot staffing gaps
    required_agents = 24.5
    staffing_gaps = required_agents - daily_availability[1:]
    
    plt.figure(figsize=(12, 6))
    plt.bar(range(len(staffing_gaps)), staffing_gaps, color=['red' if x > 0 else 'green' for x in staffing_gaps])
    plt.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
    plt.title('Daily Staffing Gaps (August 2022)', pad=20)
    plt.xlabel('Days in August')
    plt.ylabel('Staffing Gap (Negative = Surplus)')
    plt.grid(True, alpha=0.3)
    
    # Format x-axis
    plt.xticks(range(len(dates)), [d.strftime('%d-%b') for d in dates], rotation=45)
    
    plt.tight_layout()
    plt.savefig('/Users/tommygrace/Downloads/Assesment/Assessment_03/output/staffing_gaps_new.png', dpi=300, bbox_inches='tight')
    plt.close()

except Exception as e:
    print(f"An error occurred while generating plots: {e}")
