import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os
import numpy as np

# Create output directory if it doesn't exist
output_dir = '../output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Read the Excel file
df = pd.read_excel('/Users/tommygrace/Downloads/Assesment/Assessment_03/20230227_Novum_Assesment_3_BADS.xlsx', skiprows=2)

# Rename columns based on their position
df.columns = ['call_type', 'agent', 'total_calls', 'success_rate', 'total_success', 'hours_online', 'avg_duration']

# Convert numeric columns
df['total_calls'] = pd.to_numeric(df['total_calls'], errors='coerce')
df['success_rate'] = pd.to_numeric(df['success_rate'].str.rstrip('%'), errors='coerce') / 100
df['total_success'] = pd.to_numeric(df['total_success'], errors='coerce')
df['hours_online'] = pd.to_numeric(df['hours_online'], errors='coerce')
df['avg_duration'] = pd.to_numeric(df['avg_duration'], errors='coerce')

# Remove any rows with NaN values
df = df.dropna()

# Calculate additional metrics
df['calls_per_hour'] = df['total_calls'] / df['hours_online']
df['conversions_per_hour'] = df['total_success'] / df['hours_online']

# Set seaborn style
sns.set_style("whitegrid")

# 1. Scatter plot: Duration vs Success Rate
plt.figure(figsize=(12, 8))
plt.scatter(df['avg_duration'], df['success_rate'], alpha=0.6, s=100)

# Add agent labels to points
for idx, row in df.iterrows():
    plt.annotate(row['agent'], (row['avg_duration'], row['success_rate']),
                xytext=(5, 5), textcoords='offset points')

plt.title('Average Call Duration vs Success Rate by Agent', pad=20, fontsize=14)
plt.xlabel('Average Call Duration (seconds)', fontsize=12)
plt.ylabel('Success Rate', fontsize=12)

# Add trend line
z = np.polyfit(df['avg_duration'], df['success_rate'], 1)
p = np.poly1d(z)
plt.plot(df['avg_duration'], p(df['avg_duration']), "r--", alpha=0.8)

plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'duration_vs_success.png'), dpi=300, bbox_inches='tight')
plt.close()

# 2. Bar plot: Conversions per Hour
plt.figure(figsize=(12, 6))
df.sort_values('conversions_per_hour', ascending=False).plot(kind='bar', x='agent', y='conversions_per_hour')
plt.title('Conversions per Hour by Agent', pad=20, fontsize=14)
plt.xlabel('Agent', fontsize=12)
plt.ylabel('Conversions per Hour', fontsize=12)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, 'conversions_per_hour.png'), dpi=300, bbox_inches='tight')
plt.close()

# Calculate correlation coefficient
correlation = df['avg_duration'].corr(df['success_rate'])

# Generate markdown report
markdown_content = f"""# Call Efficiency Analysis Report

## Executive Summary
This analysis examines the relationship between call duration and conversion success, comparing different agent approaches to determine optimal efficiency.

## Key Findings

### 1. Duration vs Success Rate Relationship
- Correlation coefficient: {correlation:.2f}
- {'Positive' if correlation > 0 else 'Negative'} correlation between call duration and success rate
- {'Longer calls tend to have higher success rates' if correlation > 0 else 'Shorter calls tend to have higher success rates'}

### 2. Agent Performance Metrics

#### Top Agents by Conversions per Hour:
```
{df.sort_values('conversions_per_hour', ascending=False)[['agent', 'conversions_per_hour', 'success_rate', 'avg_duration']].to_markdown(index=False)}
```

## Detailed Analysis

### Call Duration Statistics
- Average call duration: {df['avg_duration'].mean():.2f} seconds
- Shortest average calls: {df['avg_duration'].min():.2f} seconds (Agent: {df.loc[df['avg_duration'].idxmin(), 'agent']})
- Longest average calls: {df['avg_duration'].max():.2f} seconds (Agent: {df.loc[df['avg_duration'].idxmax(), 'agent']})

### Success Rate Statistics
- Average success rate: {(df['success_rate'].mean() * 100):.1f}%
- Lowest success rate: {(df['success_rate'].min() * 100):.1f}% (Agent: {df.loc[df['success_rate'].idxmin(), 'agent']})
- Highest success rate: {(df['success_rate'].max() * 100):.1f}% (Agent: {df.loc[df['success_rate'].idxmax(), 'agent']})

### Efficiency Metrics
- Average conversions per hour: {df['conversions_per_hour'].mean():.2f}
- Most efficient agent: {df.loc[df['conversions_per_hour'].idxmax(), 'agent']} ({df['conversions_per_hour'].max():.2f} conversions/hour)
- Least efficient agent: {df.loc[df['conversions_per_hour'].idxmin(), 'agent']} ({df['conversions_per_hour'].min():.2f} conversions/hour)

## Analysis of Approaches

### 1. Fast Calls with Lower Success Rate
- Representative agent: {df.loc[df['avg_duration'].idxmin(), 'agent']}
- Average call duration: {df.loc[df['avg_duration'].idxmin(), 'avg_duration']:.0f} seconds
- Success rate: {(df.loc[df['avg_duration'].idxmin(), 'success_rate'] * 100):.1f}%
- Conversions per hour: {df.loc[df['avg_duration'].idxmin(), 'conversions_per_hour']:.2f}

### 2. Longer Calls with Higher Success Rate
- Representative agent: {df.loc[df['avg_duration'].idxmax(), 'agent']}
- Average call duration: {df.loc[df['avg_duration'].idxmax(), 'avg_duration']:.0f} seconds
- Success rate: {(df.loc[df['avg_duration'].idxmax(), 'success_rate'] * 100):.1f}%
- Conversions per hour: {df.loc[df['avg_duration'].idxmax(), 'conversions_per_hour']:.2f}

## Conclusion and Recommendations

### Optimal Approach
Based on the data analysis, {'longer calls with higher success rates' if df.loc[df['conversions_per_hour'].idxmax(), 'avg_duration'] > df['avg_duration'].mean() else 'shorter calls with moderate success rates'} appear to be more efficient in terms of overall success rate per hour.

### Key Success Factors:
1. {'Thorough customer engagement' if correlation > 0 else 'Efficient call handling'}
2. {'Quality over quantity' if correlation > 0 else 'Higher call volume with acceptable success rate'}
3. Balance between duration and success rate

### Recommendations:
1. {'Focus on comprehensive customer interactions' if correlation > 0 else 'Optimize call duration while maintaining quality'}
2. Study and replicate the practices of {df.loc[df['conversions_per_hour'].idxmax(), 'agent']} (highest conversions per hour)
3. {'Provide agents with more time per call' if correlation > 0 else 'Train agents in efficient call handling techniques'}
4. Regular monitoring and feedback on both duration and success metrics

## Visualizations
The following visualizations are available in the output directory:
1. `duration_vs_success.png`: Scatter plot showing the relationship between call duration and success rate
2. `conversions_per_hour.png`: Bar chart showing conversions per hour by agent

---
*Analysis generated on {datetime.now().strftime('%B %d, %Y')}*
"""

# Save markdown report
with open(os.path.join(output_dir, 'call_efficiency.md'), 'w') as f:
    f.write(markdown_content)

print("Call efficiency analysis complete! Check the output directory for results.")
