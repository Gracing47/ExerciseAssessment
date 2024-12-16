import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os

# Read the Excel file
df = pd.read_excel('../20230227_Novum_Assesment_1_BADS.xlsx')

# Create output directory if it doesn't exist
output_dir = '../output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Prepare datetime features
df['Date'] = df['DateTime'].dt.date
df['Hour'] = df['DateTime'].dt.hour
df['Weekday'] = df['DateTime'].dt.day_name()

# 1. Daily Trend Analysis
daily_volume = df.groupby('Date').size().reset_index()
daily_volume.columns = ['Date', 'Volume']

plt.figure(figsize=(15, 6))
plt.plot(daily_volume['Date'], daily_volume['Volume'], marker='o')
plt.title('Daily Interaction Volume Over Time')
plt.xlabel('Date')
plt.ylabel('Number of Interactions')
plt.xticks(rotation=45)
plt.grid(True, linestyle='--', alpha=0.7)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'daily_trend.png'))
plt.close()

# 2. Workload Distribution Analysis
# Calculate various workload metrics
workload_metrics = pd.DataFrame()
workload_metrics['Total_Actions'] = df['UserName'].value_counts()
workload_metrics['Actions_Per_Day'] = workload_metrics['Total_Actions'] / len(daily_volume)
workload_metrics['Peak_Hour_Volume'] = df.groupby('UserName')['Hour'].value_counts().groupby('UserName').max()
workload_metrics['Unique_Action_Types'] = df.groupby('UserName')['ActionType'].nunique()

# Calculate workload distribution over days of week
weekly_workload = pd.crosstab(df['UserName'], df['Weekday'])

# Visualize workload distribution
plt.figure(figsize=(12, 6))
workload_metrics['Actions_Per_Day'].plot(kind='bar')
plt.title('Average Daily Workload by Agent')
plt.xlabel('Agent')
plt.ylabel('Average Actions per Day')
plt.grid(True, linestyle='--', alpha=0.7)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'daily_workload.png'))
plt.close()

# Weekly workload heatmap
plt.figure(figsize=(12, 6))
sns.heatmap(weekly_workload, annot=True, fmt='d', cmap='YlOrRd')
plt.title('Weekly Workload Distribution by Agent')
plt.xlabel('Day of Week')
plt.ylabel('Agent')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'weekly_workload.png'))
plt.close()

# Generate markdown analysis
trend_analysis = f"""## Trend and Workload Distribution Analysis

### Trend Analysis Findings
1. **Daily Volume Patterns**
   - Average daily volume: {daily_volume['Volume'].mean():.1f} interactions
   - Highest volume day: {daily_volume.loc[daily_volume['Volume'].idxmax(), 'Date']} ({daily_volume['Volume'].max()} interactions)
   - Lowest volume day: {daily_volume.loc[daily_volume['Volume'].idxmin(), 'Date']} ({daily_volume['Volume'].min()} interactions)
   - Volume variability: {daily_volume['Volume'].std():.1f} standard deviation

### Workload Distribution Analysis
1. **Agent Workload Metrics**
   - Highest daily workload: {workload_metrics['Actions_Per_Day'].max():.1f} actions/day ({workload_metrics['Actions_Per_Day'].idxmax()})
   - Lowest daily workload: {workload_metrics['Actions_Per_Day'].min():.1f} actions/day ({workload_metrics['Actions_Per_Day'].idxmin()})
   - Workload spread: {(workload_metrics['Actions_Per_Day'].max() / workload_metrics['Actions_Per_Day'].min()):.1f}x difference between highest and lowest

2. **Key Observations**
   - Peak hour variations among agents suggest {workload_metrics['Peak_Hour_Volume'].max() - workload_metrics['Peak_Hour_Volume'].min()} hour difference in busiest times
   - Agents handle between {workload_metrics['Unique_Action_Types'].min()} to {workload_metrics['Unique_Action_Types'].max()} different types of actions
"""

# Save the analysis
with open(os.path.join(output_dir, 'trend_workload_analysis.md'), 'w') as f:
    f.write(trend_analysis)
