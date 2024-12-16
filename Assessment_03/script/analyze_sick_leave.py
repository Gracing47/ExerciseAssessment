import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

def count_sick_leave(value):
    """Count if the value represents sick leave."""
    if isinstance(value, str) and value.lower() == 'sick':
        return 1
    return 0

def analyze_sick_leave(file_path):
    """Analyze sick leave patterns from the Excel file."""
    # Read Excel file
    df = pd.read_excel(file_path)
    
    # Skip the header row and get agent columns
    df = df.iloc[1:]  # Skip header
    agent_columns = df.columns[2:]  # Skip date and day columns
    
    # Convert date column to datetime
    df['Date'] = pd.to_datetime(df['Unnamed: 0'])
    
    # Calculate daily sick count
    daily_sick = df[agent_columns].applymap(count_sick_leave).sum(axis=1)
    
    # Create figure with two subplots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
    
    # Plot 1: Daily Sick Leave Count
    ax1.plot(range(1, len(daily_sick) + 1), daily_sick, marker='o', color='red')
    ax1.set_title('Daily Sick Leave Count (August 2022)')
    ax1.set_xlabel('Working Day')
    ax1.set_ylabel('Number of Sick Leaves')
    ax1.grid(True)
    
    # Add day labels
    for i, (idx, count) in enumerate(daily_sick.items(), 1):
        if count > 0:  # Only label days with sick leave
            ax1.annotate(f'Day {i}\n({count})', 
                        (i, count),
                        textcoords="offset points",
                        xytext=(0,10),
                        ha='center')
    
    # Plot 2: Heatmap of Sick Leave Distribution
    sick_matrix = df[agent_columns].applymap(count_sick_leave)
    sick_matrix.index = range(1, len(sick_matrix) + 1)  # Working days
    
    sns.heatmap(sick_matrix.T, cmap='YlOrRd', ax=ax2)
    ax2.set_title('Sick Leave Distribution by Agent and Day')
    ax2.set_xlabel('Working Day')
    ax2.set_ylabel('Agent')
    
    plt.tight_layout()
    plt.savefig('/Users/tommygrace/Downloads/Assesment/Assessment_03/output/sick_leave_analysis.png')
    plt.close()
    
    # Calculate statistics
    total_sick_days = daily_sick.sum()
    max_sick_day = daily_sick.max()
    avg_sick_per_day = daily_sick.mean()
    days_with_sick = (daily_sick > 0).sum()
    
    print("\nSick Leave Statistics:")
    print(f"Total sick leave instances: {total_sick_days}")
    print(f"Maximum sick leaves in one day: {max_sick_day}")
    print(f"Average sick leaves per day: {avg_sick_per_day:.2f}")
    print(f"Days with at least one sick leave: {days_with_sick}")
    
    return {
        'total_sick_days': total_sick_days,
        'max_sick_day': max_sick_day,
        'avg_sick_per_day': avg_sick_per_day,
        'days_with_sick': days_with_sick
    }

if __name__ == "__main__":
    file_path = '/Users/tommygrace/Downloads/Assesment/Assessment_03/20230227_Novum_Assesment_2_BADS.xlsx'
    stats = analyze_sick_leave(file_path)
