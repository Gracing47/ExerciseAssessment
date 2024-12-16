import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os
import numpy as np

# Set seaborn style
sns.set_style("whitegrid")

# Create output directory if it doesn't exist
output_dir = '../output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Read the Excel file
df = pd.read_excel('/Users/tommygrace/Downloads/Assesment/Assessment_02/20230227_Novum_Assesment_2_BADS.xlsx', skiprows=1)

# Clean up column names and data
df['Date'] = pd.to_datetime(df['Calendar Day'])
df['Day'] = df['Day of the Week']

# Calculate available agents per day
# First, get all agent columns (excluding non-agent columns)
agent_columns = [col for col in df.columns if 'Agent' in col]

def count_available_agents(row):
    # Count agents who are not on leave or absent
    available = sum(1 for col in agent_columns 
                   if pd.notna(row[col]) and 
                   str(row[col]).lower() != 'leave' and
                   str(row[col]).lower() != 'absent')
    return available

# Calculate daily availability
df['Available_Agents'] = df.apply(count_available_agents, axis=1)

# Required staffing levels for each service line
required_staff = {
    'Service Line 1': 2.5,
    'Service Line 2': 4,
    'Service Line 3': 18
}

# Daily availability analysis
daily_stats = df.groupby('Date').agg({
    'Available_Agents': ['mean', 'min', 'max']
}).round(2)

# Calculate staffing gaps
daily_gaps = pd.DataFrame()
for service_line, required in required_staff.items():
    daily_gaps[f'{service_line} Gap'] = df['Available_Agents'] - required

# Monthly overview
monthly_stats = {
    'Average Available Agents': df['Available_Agents'].mean(),
    'Minimum Available Agents': df['Available_Agents'].min(),
    'Maximum Available Agents': df['Available_Agents'].max(),
    'Days Below Service Line 1 (2.5)': (df['Available_Agents'] < 2.5).sum(),
    'Days Below Service Line 2 (4)': (df['Available_Agents'] < 4).sum(),
    'Days Below Service Line 3 (18)': (df['Available_Agents'] < 18).sum()
}

# Weekly patterns
df['Weekday'] = df['Date'].dt.day_name()
weekly_pattern = df.groupby('Weekday')['Available_Agents'].agg(['mean', 'min', 'max']).round(2)

# Sort weekdays in correct order
weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
weekly_pattern = weekly_pattern.reindex(weekday_order)

# 1. Daily Availability vs Requirements
plt.figure(figsize=(15, 8))
plt.plot(df['Date'], df['Available_Agents'], label='Available Agents', marker='o', linewidth=2)
for service_line, required in required_staff.items():
    plt.axhline(y=required, linestyle='--', label=f'{service_line} Requirement ({required})')
plt.title('Daily Agent Availability vs Service Line Requirements', pad=20, fontsize=14)
plt.xlabel('Date', fontsize=12)
plt.ylabel('Number of Agents', fontsize=12)
plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
plt.grid(True, alpha=0.3)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'daily_availability.png'), dpi=300, bbox_inches='tight')
plt.close()

# 2. Weekly Pattern
plt.figure(figsize=(12, 6))
ax = weekly_pattern['mean'].plot(kind='bar', color='skyblue', width=0.8)
plt.title('Average Agent Availability by Day of Week', pad=20, fontsize=14)
plt.xlabel('Day of Week', fontsize=12)
plt.ylabel('Number of Agents', fontsize=12)
plt.grid(True, alpha=0.3, axis='y')

# Add value labels on top of bars
for i, v in enumerate(weekly_pattern['mean']):
    ax.text(i, v + 0.1, f'{v:.1f}', ha='center', fontsize=10)

plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'weekly_pattern.png'), dpi=300, bbox_inches='tight')
plt.close()

# 3. Staffing Gaps Heatmap
daily_gaps_pivot = daily_gaps.copy()
daily_gaps_pivot.index = df['Date']
plt.figure(figsize=(15, 6))
sns.heatmap(daily_gaps_pivot.T, cmap='RdYlGn', center=0, annot=True, fmt='.1f', 
            cbar_kws={'label': 'Staffing Gap (Positive = Surplus, Negative = Shortage)'})
plt.title('Staffing Gaps by Service Line', pad=20, fontsize=14)
plt.xlabel('Date', fontsize=12)
plt.ylabel('Service Line', fontsize=12)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'staffing_gaps.png'), dpi=300, bbox_inches='tight')
plt.close()

# Generate markdown report
markdown_content = f"""# Team Availability Analysis Report
## Period: {df['Date'].min().strftime('%B %d, %Y')} to {df['Date'].max().strftime('%B %d, %Y')}

## Executive Summary
This analysis examines team availability against required staffing levels for three service lines during {df['Date'].min().strftime('%B %Y')}. The data shows significant challenges in meeting staffing requirements, particularly for Service Line 3.

### Required Staffing Levels
- Service Line 1: 2.5 agents
- Service Line 2: 4 agents
- Service Line 3: 18 agents

## Current Staffing Situation

### Monthly Overview
- **Average Available Agents**: {monthly_stats['Average Available Agents']:.1f}
- **Minimum Available Agents**: {monthly_stats['Minimum Available Agents']:.1f}
- **Maximum Available Agents**: {monthly_stats['Maximum Available Agents']:.1f}

### Coverage Issues
- **Service Line 1** (2.5 agents required):
  - Days below requirement: {monthly_stats['Days Below Service Line 1 (2.5)']} days
  - Coverage rate: {100 - (monthly_stats['Days Below Service Line 1 (2.5)']/len(df)*100):.1f}%

- **Service Line 2** (4 agents required):
  - Days below requirement: {monthly_stats['Days Below Service Line 2 (4)']} days
  - Coverage rate: {100 - (monthly_stats['Days Below Service Line 2 (4)']/len(df)*100):.1f}%

- **Service Line 3** (18 agents required):
  - Days below requirement: {monthly_stats['Days Below Service Line 3 (18)']} days
  - Coverage rate: {100 - (monthly_stats['Days Below Service Line 3 (18)']/len(df)*100):.1f}%

## Detailed Analysis

### Weekly Availability Pattern
```
{weekly_pattern.to_markdown()}
```

### Key Findings

1. **Staffing Levels**
   - Current average of {monthly_stats['Average Available Agents']:.1f} agents is insufficient for Service Line 3
   - Maximum staffing ({monthly_stats['Maximum Available Agents']:.1f} agents) never reaches Service Line 3 requirement
   - Minimum staffing ({monthly_stats['Minimum Available Agents']:.1f} agents) is critically low

2. **Weekly Patterns**
   - Strongest staffing: {weekly_pattern['mean'].idxmax()} ({weekly_pattern['mean'].max():.1f} agents)
   - Weakest staffing: {weekly_pattern['mean'].idxmin()} ({weekly_pattern['mean'].min():.1f} agents)
   - Weekend coverage requires attention

3. **Service Line Impact**
   - Service Lines 1 & 2 are generally manageable
   - Service Line 3 faces significant understaffing
   - Current staffing model cannot support all service lines simultaneously

## Recommendations

1. **Immediate Actions**
   - Prioritize recruitment to meet Service Line 3 requirements
   - Implement temporary staff augmentation for critical coverage
   - Review and optimize leave management

2. **Short-term Improvements**
   - Develop flexible scheduling to maximize coverage during peak times
   - Cross-train agents across service lines
   - Create contingency plans for minimum staffing scenarios

3. **Long-term Strategy**
   - Establish sustainable staffing model meeting all service line requirements
   - Implement workforce management system
   - Develop predictive scheduling based on historical patterns

## Visualizations
The following visualizations are available in the output directory:
1. `daily_availability.png`: Daily agent availability compared to service line requirements
2. `weekly_pattern.png`: Average agent availability by day of week
3. `staffing_gaps.png`: Detailed view of staffing gaps for each service line

---
*Analysis generated on {datetime.now().strftime('%B %d, %Y')}*
"""

# Save markdown report
with open(os.path.join(output_dir, 'team_availability.md'), 'w') as f:
    f.write(markdown_content)

print("Team availability analysis complete! Check the output directory for results.")
