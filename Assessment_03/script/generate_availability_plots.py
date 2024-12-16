import pandas as pd
import matplotlib.pyplot as plt

# Path to the Excel file
file_path = '/Users/tommygrace/Downloads/Assesment/Assessment_03/20230227_Novum_Assesment_2_BADS.xlsx'

# Read the Excel file
try:
    # Load the first sheet of the Excel file
    df = pd.read_excel(file_path, sheet_name=0)
    
    # Extract relevant columns
    date_columns = df.columns[2:]  # Assuming first two columns are date and day
    
    # Calculate daily availability
    daily_availability = df[date_columns].notna().sum(axis=1)
    
    # Plot daily availability
    plt.figure(figsize=(10, 6))
    plt.plot(daily_availability, marker='o', linestyle='-')
    plt.title('Daily Team Availability')
    plt.xlabel('Day')
    plt.ylabel('Number of Available Agents')
    plt.grid(True)
    plt.savefig('/Users/tommygrace/Downloads/Assesment/Assessment_03/output/daily_availability.png')
    plt.close()

    # Plot staffing gaps
    required_agents = 24.5
    staffing_gaps = required_agents - daily_availability
    plt.figure(figsize=(10, 6))
    plt.plot(staffing_gaps, marker='o', linestyle='-', color='red')
    plt.title('Staffing Gaps')
    plt.xlabel('Day')
    plt.ylabel('Gap in Number of Agents')
    plt.grid(True)
    plt.savefig('/Users/tommygrace/Downloads/Assesment/Assessment_03/output/staffing_gaps.png')
    plt.close()

except Exception as e:
    print(f"An error occurred while generating plots: {e}")
