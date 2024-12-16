import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Read the Excel file
df = pd.read_excel('../20230227_Novum_Assesment_1_BADS.xlsx')

# Extract hour from DateTime
df['Hour'] = df['DateTime'].dt.hour

# Calculate hourly activity for each agent
hourly_activity = df.groupby(['UserName', 'Hour']).size().reset_index()
hourly_activity.columns = ['Agent', 'Hour', 'Number_of_Actions']

# Find peak hours for each agent
peak_hours = hourly_activity.loc[hourly_activity.groupby('Agent')['Number_of_Actions'].idxmax()]
peak_hours = peak_hours.sort_values('Number_of_Actions', ascending=False)

# Create a visualization of hourly patterns
plt.figure(figsize=(12, 6))
for agent in df['UserName'].unique():
    agent_data = hourly_activity[hourly_activity['Agent'] == agent]
    plt.plot(agent_data['Hour'], agent_data['Number_of_Actions'], marker='o', label=agent)

plt.title('Hourly Activity Pattern by Agent')
plt.xlabel('Hour of Day (24-hour format)')
plt.ylabel('Number of Actions')
plt.grid(True, alpha=0.3)
plt.legend()
plt.tight_layout()
plt.savefig('../output/peak_hours_analysis.png')
plt.close()

# Generate a detailed analysis
analysis = """## Peak Hours Analysis

### Peak Activity Hours by Agent
"""

for _, row in peak_hours.iterrows():
    analysis += f"\n- {row['Agent']}: Busiest at {row['Hour']:02d}:00 with {row['Number_of_Actions']} actions"

# Calculate the time spread
earliest_peak = peak_hours['Hour'].min()
latest_peak = peak_hours['Hour'].max()
hour_spread = latest_peak - earliest_peak

analysis += f"""

### Key Findings
1. **Peak Hour Spread**
   - Earliest peak hour: {earliest_peak:02d}:00
   - Latest peak hour: {latest_peak:02d}:00
   - Time difference between earliest and latest peak: {hour_spread} hours

2. **Work Pattern Insights**
   - Different agents have their peak activity at different times of the day
   - This suggests a good coverage of working hours across the team
   - The spread of {hour_spread} hours between peak times helps ensure continuous service throughout the day
"""

# Save the analysis
with open('../output/peak_hours_analysis.md', 'w') as f:
    f.write(analysis)

# Print the analysis to console as well
print(analysis)
