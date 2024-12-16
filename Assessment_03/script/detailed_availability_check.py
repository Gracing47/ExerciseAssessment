import pandas as pd
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
            print(f"Unhandled value: {value}")
            return 0

try:
    # Load the Excel file
    df = pd.read_excel(file_path)
    
    # Print basic information
    print("\nFile Structure:")
    print(f"Total columns: {len(df.columns)}")
    print(f"Column names: {df.columns.tolist()}")
    print(f"\nFirst few rows of data:")
    print(df.head())
    
    # Get agent columns (excluding date and day of week)
    agent_columns = df.columns[2:]
    print(f"\nTotal number of agents: {len(agent_columns)}")
    
    # Calculate daily availability considering partial availability
    daily_availability = df[agent_columns].applymap(calculate_agent_availability).sum(axis=1)
    
    print("\nDaily Availability Statistics:")
    print(f"Average daily available agents: {daily_availability.mean():.2f}")
    print(f"Maximum available agents: {daily_availability.max():.2f}")
    print(f"Minimum available agents: {daily_availability.min():.2f}")
    
    # Service line requirements
    required_agents = 24.5  # (2.5 + 4 + 18)
    days_meeting_requirements = (daily_availability >= required_agents).sum()
    
    print(f"\nDays meeting total requirement ({required_agents} agents):")
    print(f"Meeting requirements: {days_meeting_requirements} days")
    print(f"Below requirements: {len(daily_availability) - days_meeting_requirements} days")
    
    # Print unique values in the dataset to understand different availability states
    print("\nUnique values in the dataset:")
    unique_values = set()
    for col in agent_columns:
        unique_values.update(df[col].dropna().unique())
    print(unique_values)

except Exception as e:
    print(f"An error occurred: {e}")
