import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt

def calculate_agent_availability(value):
    """Calculate agent availability based on cell value."""
    if pd.isna(value):
        return 1.0  # Fully available
    if isinstance(value, (int, float)):
        return float(value)  # Partial availability
    if isinstance(value, str):
        if value.lower() in ['sick', 'leave']:
            return 0.0  # Not available
        try:
            return float(value)  # Try converting string to float
        except ValueError:
            if 'agent' in value.lower():  # Header row
                return 0.0
    return 0.0  # Default case

def analyze_availability(file_path):
    """Analyze team availability from Excel file."""
    # Read Excel file
    df = pd.read_excel(file_path)
    
    # Get agent columns (excluding date and day columns)
    agent_columns = df.columns[2:]
    
    # Calculate daily availability for each agent
    daily_availability = df[agent_columns].applymap(calculate_agent_availability).sum(axis=1)
    
    # Skip the header row for calculations
    daily_availability = daily_availability[1:]  # Skip first row which contains headers
    
    # Calculate statistics
    stats = {
        'total_agents': len(agent_columns),
        'avg_daily_available': daily_availability.mean(),
        'max_available': daily_availability.max(),
        'min_available': daily_availability.min(),
        'days_meeting_req': sum(daily_availability >= 24.5),
        'days_below_req': sum(daily_availability < 24.5),
        'total_working_days': len(daily_availability)
    }
    
    # Create visualization
    plt.figure(figsize=(12, 6))
    plt.plot(range(1, len(daily_availability) + 1), daily_availability, marker='o')
    plt.axhline(y=24.5, color='r', linestyle='--', label='Required (24.5)')
    plt.title('Daily Agent Availability (August 2022)')
    plt.xlabel('Working Day')
    plt.ylabel('Available Agents')
    plt.grid(True)
    plt.legend()
    plt.tight_layout()
    plt.savefig('/Users/tommygrace/Downloads/Assesment/Assessment_03/output/daily_availability.png')
    plt.close()
    
    return stats

if __name__ == "__main__":
    file_path = '/Users/tommygrace/Downloads/Assesment/Assessment_03/20230227_Novum_Assesment_2_BADS.xlsx'
    stats = analyze_availability(file_path)
    
    print("\nTeam Availability Analysis:")
    print(f"Total number of agents: {stats['total_agents']}")
    print(f"Average daily available agents: {stats['avg_daily_available']:.2f}")
    print(f"Maximum available agents: {stats['max_available']:.2f}")
    print(f"Minimum available agents: {stats['min_available']:.2f}")
    print(f"Days meeting requirements (≥24.5 agents): {stats['days_meeting_req']}")
    print(f"Days below requirements: {stats['days_below_req']}")
    print(f"Total working days: {stats['total_working_days']}")
