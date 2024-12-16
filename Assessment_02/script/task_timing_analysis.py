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
task_by_hour = pd.crosstab([df['UserName'], df['Hour']], df['ActionType'])

# Analysis 2: Peak Hour Task Breakdown
peak_hours = {
    'Agent 1': 16,
    'Agent 2': 21,
    'Agent 3': 12,
    'Agent 4': 15,
    'Agent 5': 17
}

peak_hour_analysis = []
for agent, peak_hour in peak_hours.items():
    peak_data = df[(df['UserName'] == agent) & (df['Hour'] == peak_hour)]
    task_breakdown = peak_data['ActionType'].value_counts()
    peak_hour_analysis.append({
        'Agent': agent,
        'Peak Hour': peak_hour,
        'Total Actions': len(peak_data),
        'Task Breakdown': task_breakdown.to_dict()
    })

# Visualization 1: Task Types Throughout the Day
plt.figure(figsize=(15, 10))
for agent in df['UserName'].unique():
    agent_data = df[df['UserName'] == agent]
    hourly_counts = agent_data.groupby(['Hour', 'ActionType']).size().unstack(fill_value=0)
    
    plt.subplot(3, 2, int(agent[-1]))
    hourly_counts.plot(kind='bar', stacked=True)
    plt.title(f'{agent} - Task Types by Hour')
    plt.xlabel('Hour')
    plt.ylabel('Number of Actions')
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()

plt.savefig(os.path.join(output_dir, 'task_types_by_hour.png'), bbox_inches='tight')
plt.close()

# Generate detailed analysis
analysis = """# Task and Timing Analysis

## Peak Hour Analysis by Agent

"""

for agent_data in peak_hour_analysis:
    analysis += f"\n### {agent_data['Agent']} (Peak Hour: {agent_data['Peak Hour']:02d}:00)\n"
    analysis += f"Total Actions: {agent_data['Total Actions']}\n\n"
    analysis += "Task Breakdown:\n"
    for task, count in agent_data['Task_Breakdown'].items():
        percentage = (count / agent_data['Total Actions']) * 100
        analysis += f"- {task}: {count} ({percentage:.1f}%)\n"

analysis += """
## Key Findings

1. **Agent Specializations**
"""

# Calculate agent specializations
for agent in df['UserName'].unique():
    agent_data = df[df['UserName'] == agent]
    main_task = agent_data['ActionType'].mode().iloc[0]
    main_task_percent = (agent_data['ActionType'].value_counts().iloc[0] / len(agent_data)) * 100
    analysis += f"\n- {agent}: Primarily handles {main_task} ({main_task_percent:.1f}% of their work)"

analysis += """

2. **Peak Hour Patterns**
- Morning Peak (12:00): Agent 3 focuses mainly on orders
- Afternoon Peak (15:00-17:00): Agents 1, 4, and 5 handle diverse tasks
- Evening Peak (21:00): Agent 2 handles late communications

3. **Task Distribution Insights**
"""

# Calculate overall task distribution
total_tasks = df['ActionType'].value_counts()
for task, count in total_tasks.items():
    percentage = (count / len(df)) * 100
    analysis += f"\n- {task}: {count} total ({percentage:.1f}% of all work)"

# Save the analysis
with open(os.path.join(output_dir, 'task_timing_analysis.md'), 'w') as f:
    f.write(analysis)

# Print the analysis
print(analysis)
