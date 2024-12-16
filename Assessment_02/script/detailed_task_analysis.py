import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Read the Excel file
df = pd.read_excel('../20230227_Novum_Assesment_1_BADS.xlsx')

# Extract hour from DateTime
df['Hour'] = df['DateTime'].dt.hour

# Create output directory if it doesn't exist
import os
output_dir = '../output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Analysis 1: Task Types by Hour for each Agent
hourly_tasks = pd.crosstab([df['UserName'], df['Hour']], df['ActionType'])

# Analysis 2: Peak Hour Analysis
peak_hours = {
    'Agent 1': 16,
    'Agent 2': 21,
    'Agent 3': 12,
    'Agent 4': 15,
    'Agent 5': 17
}

# Generate detailed analysis
analysis = """# Task and Timing Analysis

## Peak Hour Analysis by Agent\n"""

for agent, peak_hour in peak_hours.items():
    # Get data for this agent's peak hour
    peak_data = df[(df['UserName'] == agent) & (df['Hour'] == peak_hour)]
    task_counts = peak_data['ActionType'].value_counts()
    
    analysis += f"\n### {agent} (Peak Hour: {peak_hour:02d}:00)\n"
    analysis += f"Total Actions: {len(peak_data)}\n\n"
    analysis += "Task Breakdown:\n"
    for task, count in task_counts.items():
        percentage = (count / len(peak_data)) * 100
        analysis += f"- {task}: {count} ({percentage:.1f}%)\n"

# Calculate agent specializations
analysis += "\n## Key Findings\n\n1. **Agent Specializations**\n"
for agent in df['UserName'].unique():
    agent_data = df[df['UserName'] == agent]
    task_counts = agent_data['ActionType'].value_counts()
    main_task = task_counts.index[0]
    main_task_percent = (task_counts.iloc[0] / len(agent_data)) * 100
    analysis += f"\n- {agent}: Primarily handles {main_task} ({main_task_percent:.1f}% of their work)"

# Analyze peak hour patterns
analysis += """

2. **Peak Hour Patterns**
- Morning Peak (12:00): Agent 3 focuses on orders
- Afternoon Peak (15:00-17:00): Agents 1, 4, and 5 handle the highest volume
- Evening Peak (21:00): Agent 2 handles late shift tasks

3. **Task Type Distribution**"""

# Calculate overall task distribution
total_tasks = df['ActionType'].value_counts()
for task, count in total_tasks.items():
    percentage = (count / len(df)) * 100
    analysis += f"\n- {task}: {count} total ({percentage:.1f}% of all work)"

# Create visualization of hourly patterns by task type
plt.figure(figsize=(15, 10))
for i, agent in enumerate(df['UserName'].unique(), 1):
    agent_data = df[df['UserName'] == agent]
    task_by_hour = pd.crosstab(agent_data['Hour'], agent_data['ActionType'])
    
    plt.subplot(3, 2, i)
    task_by_hour.plot(kind='bar', stacked=True)
    plt.title(f'{agent} - Tasks by Hour')
    plt.xlabel('Hour of Day')
    plt.ylabel('Number of Actions')
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize='small')
    plt.xticks(rotation=45)

plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'hourly_task_patterns.png'), bbox_inches='tight')
plt.close()

# Save the analysis
with open(os.path.join(output_dir, 'detailed_task_analysis.md'), 'w') as f:
    f.write(analysis)

# Print the analysis
print(analysis)
