import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os

# Read the Excel file
df = pd.read_excel('/Users/tommygrace/Downloads/Assesment/Assessment_01/20230227_Novum_Assesment_1_BADS.xlsx')

# Create output directory if it doesn't exist
output_dir = '../output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Get date range
date_range = {
    'start': df['DateTime'].min(),
    'end': df['DateTime'].max(),
    'total_days': (df['DateTime'].max() - df['DateTime'].min()).days
}

# 1. Agent Performance Metrics
agent_metrics = pd.DataFrame()
agent_metrics['Total_Actions'] = df['UserName'].value_counts()
agent_metrics['Orders'] = df[df['ActionType'] == 'Order'].groupby('UserName').size()
agent_metrics['Success_Rate'] = df[df['Change'].isin(['Approved', 'Processed'])].groupby('UserName').size() / agent_metrics['Total_Actions'] * 100
agent_metrics['Rejection_Rate'] = df[df['Change'] == 'Rejected'].groupby('UserName').size() / agent_metrics['Total_Actions'] * 100

# 2. Daily Activity Pattern
df['Date'] = df['DateTime'].dt.date
df['Hour'] = df['DateTime'].dt.hour
daily_pattern = df.groupby(['UserName', 'Hour']).size().unstack(fill_value=0)

# 3. Communication Channels Analysis
comm_channels = pd.crosstab(df['UserName'], df['ActionType'])
channel_success = df.groupby('ActionType')['Change'].value_counts().unstack(fill_value=0)

# Visualizations
# 1. Agent Performance Overview
plt.figure(figsize=(12, 6))
agent_metrics[['Total_Actions', 'Orders']].plot(kind='bar')
plt.title('Agent Performance Overview')
plt.xlabel('Agent')
plt.ylabel('Count')
plt.legend(loc='upper right')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'agent_performance.png'))
plt.close()

# 2. Success and Rejection Rates
plt.figure(figsize=(10, 6))
agent_metrics[['Success_Rate', 'Rejection_Rate']].plot(kind='bar')
plt.title('Agent Success and Rejection Rates')
plt.xlabel('Agent')
plt.ylabel('Percentage')
plt.legend(loc='upper right')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'agent_rates.png'))
plt.close()

# 3. Daily Activity Pattern Heatmap
plt.figure(figsize=(12, 6))
sns.heatmap(daily_pattern, cmap='YlOrRd', annot=True, fmt='.0f', cbar_kws={'label': 'Number of Actions'})
plt.title('Hourly Activity Pattern by Agent')
plt.xlabel('Hour of Day')
plt.ylabel('Agent')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'hourly_pattern.png'))
plt.close()

# 4. Communication Channels
plt.figure(figsize=(12, 6))
comm_channels.plot(kind='bar', stacked=True)
plt.title('Communication Channels Used by Agents')
plt.xlabel('Agent')
plt.ylabel('Number of Actions')
plt.legend(title='Channel', bbox_to_anchor=(1.05, 1))
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'communication_channels.png'))
plt.close()

# Generate comprehensive markdown report
markdown_content = f"""# Customer Service Operations Analysis
## Data Overview: {date_range['start'].strftime('%B %d, %Y')} to {date_range['end'].strftime('%B %d, %Y')}

## Executive Summary
This analysis examines {len(df)} customer service interactions over a {date_range['total_days']}-day period. The dataset captures various aspects of customer service operations, including order processing, communication channels, and agent performance metrics.

### Data Scope
- **Time Period**: {date_range['total_days']} days
- **Total Interactions**: {len(df):,}
- **Number of Agents**: {len(agent_metrics)}
- **Types of Actions**: {', '.join(df['ActionType'].unique())}

## Key Findings

### 1. Operational Overview
- **Daily Average**: {len(df)/date_range['total_days']:.0f} interactions per day
- **Peak Hour**: {df.groupby('Hour')['UserName'].count().idxmax()}:00 ({df.groupby('Hour')['UserName'].count().max()} interactions)
- **Most Common Action**: {df['ActionType'].mode()[0]} ({df['ActionType'].value_counts().iloc[0]:,} occurrences)

### 2. Agent Performance
```
{agent_metrics.round(2).to_markdown()}
```

### 3. Communication Channels
```
{comm_channels.to_markdown()}
```

### 4. Success Metrics
- Highest success rate: {agent_metrics['Success_Rate'].max():.1f}% ({agent_metrics['Success_Rate'].idxmax()})
- Lowest rejection rate: {agent_metrics['Rejection_Rate'].min():.1f}% ({agent_metrics['Rejection_Rate'].idxmin()})

## Data Organization Recommendations

### 1. Suggested Pivot Tables
1. **Agent Performance Pivot**
   - Rows: Agent Names
   - Columns: Action Types
   - Values: Count of interactions
   - This shows the distribution of work across agents

2. **Time Analysis Pivot**
   - Rows: Date
   - Columns: Hour of Day
   - Values: Count of interactions
   - This reveals peak operation times

3. **Success Rate Pivot**
   - Rows: Agent Names
   - Columns: Change Status
   - Values: Percentage of total
   - This highlights agent effectiveness

### 2. Recommended Visualizations
1. **Workload Distribution** (agent_performance.png)
   - Bar chart showing total actions and orders per agent
   - Helps identify workload balance

2. **Success Metrics** (agent_rates.png)
   - Comparative view of success and rejection rates
   - Identifies top performers

3. **Activity Patterns** (hourly_pattern.png)
   - Heatmap of hourly activity
   - Optimal for workforce planning

4. **Channel Usage** (communication_channels.png)
   - Stacked bar chart of communication methods
   - Shows preferred customer contact methods

## Conclusions and Recommendations

1. **Workload Distribution**
   - Current distribution shows significant variation between agents
   - Recommendation: Consider workload balancing strategies

2. **Performance Optimization**
   - Success rates vary by agent and channel
   - Recommendation: Implement best practice sharing among agents

3. **Operational Efficiency**
   - Peak hours identified at {df.groupby('Hour')['UserName'].count().idxmax()}:00
   - Recommendation: Adjust staffing to match peak times

4. **Communication Channels**
   - Multiple channels being utilized
   - Recommendation: Focus on channels with highest success rates

## Next Steps
1. Implement suggested pivot tables for regular monitoring
2. Set up automated reporting using provided visualizations
3. Develop training program based on top performer practices
4. Review staffing patterns against peak operation times

---
*Analysis generated on {datetime.now().strftime('%B %d, %Y')}*
"""

# Save markdown report
with open(os.path.join(output_dir, 'agent_analysis.md'), 'w') as f:
    f.write(markdown_content)

print("Updated comprehensive analysis complete! Check the output directory for results.")
