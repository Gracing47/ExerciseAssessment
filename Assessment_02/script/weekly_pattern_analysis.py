import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# Read the Excel file
df = pd.read_excel('../20230227_Novum_Assesment_1_BADS.xlsx')

# Create output directory if it doesn't exist
import os
output_dir = '../output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Add time-based columns
df['Hour'] = df['DateTime'].dt.hour
df['Weekday'] = df['DateTime'].dt.day_name()
df['Date'] = df['DateTime'].dt.date

# Ensure weekdays are in correct order
weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

# 1. Weekly Pattern Analysis
plt.figure(figsize=(15, 8))
weekly_pattern = pd.crosstab(df['Weekday'], df['ActionType'])
weekly_pattern = weekly_pattern.reindex(weekday_order)
weekly_pattern.plot(kind='bar', stacked=True)
plt.title('Task Distribution by Day of Week')
plt.xlabel('Day of Week')
plt.ylabel('Number of Actions')
plt.legend(title='Task Type', bbox_to_anchor=(1.05, 1))
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'weekly_task_distribution.png'), bbox_inches='tight')
plt.close()

# 2. Agent Activity by Day
plt.figure(figsize=(15, 8))
agent_weekly = pd.crosstab(df['Weekday'], df['UserName'])
agent_weekly = agent_weekly.reindex(weekday_order)
agent_weekly.plot(kind='bar', stacked=True)
plt.title('Agent Activity by Day of Week')
plt.xlabel('Day of Week')
plt.ylabel('Number of Actions')
plt.legend(title='Agent', bbox_to_anchor=(1.05, 1))
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'weekly_agent_activity.png'), bbox_inches='tight')
plt.close()

# 3. Analyze consistency of patterns
analysis = """# Weekly Pattern Analysis and Optimization Suggestions

## Weekly Patterns by Day

"""

# Calculate daily statistics
daily_stats = df.groupby('Weekday').agg({
    'ActionType': 'count',
    'UserName': 'nunique'
}).reindex(weekday_order)

for day in weekday_order:
    if day in daily_stats.index:
        actions = daily_stats.loc[day, 'ActionType']
        agents = daily_stats.loc[day, 'UserName']
        analysis += f"\n### {day}\n"
        analysis += f"- Total Actions: {actions}\n"
        analysis += f"- Active Agents: {agents}\n"
        analysis += f"- Actions per Agent: {actions/agents:.1f}\n"

# Calculate busiest times by day
analysis += "\n## Peak Hours by Day\n"
for day in weekday_order:
    day_data = df[df['Weekday'] == day]
    if not day_data.empty:
        peak_hour = day_data.groupby('Hour')['ActionType'].count().idxmax()
        peak_actions = day_data.groupby('Hour')['ActionType'].count().max()
        analysis += f"\n### {day}\n"
        analysis += f"- Peak Hour: {peak_hour:02d}:00\n"
        analysis += f"- Actions at Peak: {peak_actions}\n"

# Calculate task distribution by day
analysis += "\n## Task Distribution Patterns\n"
task_distribution = pd.crosstab(df['Weekday'], df['ActionType'], normalize='index') * 100
task_distribution = task_distribution.reindex(weekday_order)

for day in weekday_order:
    if day in task_distribution.index:
        analysis += f"\n### {day}\n"
        for task in task_distribution.columns:
            percentage = task_distribution.loc[day, task]
            analysis += f"- {task}: {percentage:.1f}%\n"

# Add optimization suggestions
analysis += """
## Optimization Suggestions

1. **Workload Balancing**
"""

# Calculate workload imbalances
agent_workload = df.groupby('UserName')['ActionType'].count()
max_workload = agent_workload.max()
min_workload = agent_workload.min()
workload_ratio = max_workload / min_workload

analysis += f"- Current workload ratio (busiest/least busy): {workload_ratio:.1f}x\n"
analysis += f"- Busiest agent: {agent_workload.idxmax()} ({max_workload} actions)\n"
analysis += f"- Least busy agent: {agent_workload.idxmin()} ({min_workload} actions)\n"

analysis += """
2. **Suggested Schedule Adjustments**
- Stagger shift starts to better cover peak hours
- Assign more agents during identified peak times
- Cross-train agents to handle multiple task types

3. **Task Distribution Optimization**
"""

# Analyze task specialization
for agent in df['UserName'].unique():
    agent_tasks = df[df['UserName'] == agent]['ActionType'].value_counts()
    primary_task = agent_tasks.index[0]
    primary_percentage = (agent_tasks.iloc[0] / agent_tasks.sum()) * 100
    analysis += f"- {agent}: {primary_percentage:.1f}% {primary_task} - "
    if primary_percentage > 80:
        analysis += "Consider diversifying tasks\n"
    elif primary_percentage < 40:
        analysis += "Good task balance\n"
    else:
        analysis += "Moderate specialization\n"

# Save the analysis
with open(os.path.join(output_dir, 'weekly_optimization_analysis.md'), 'w') as f:
    f.write(analysis)

# Print the analysis
print(analysis)
