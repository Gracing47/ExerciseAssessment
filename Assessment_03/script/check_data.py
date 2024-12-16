import pandas as pd
import matplotlib.pyplot as plt

# Read the Excel file
df = pd.read_excel('/Users/tommygrace/Downloads/Assesment/Assessment_03/20230227_Novum_Assesment_3_BADS.xlsx', skiprows=1)

# Update columns based on the provided data
columns = ['Call Type', 'Agent', 'Invoices', 'Orders', 'Avg Calls/Day', 'h in phone line', 'Call Duration']
df.columns = columns

# Display basic information about the dataframe
print("\nDataframe Info:")
print(df.info())

print("\nColumn Names:")
print(df.columns.tolist())

print("\nFirst few rows:")
print(df.head())

# Convert 'Call Duration' to numeric types
df['Call Duration'] = pd.to_numeric(df['Call Duration'], errors='coerce')

# Convert 'Avg Calls/Day' to numeric type
df['Avg Calls/Day'] = pd.to_numeric(df['Avg Calls/Day'], errors='coerce')

# Calculate average call length and conversion efficiency
average_call_duration = df['Call Duration'].mean()
average_calls_per_day = df['Avg Calls/Day'].mean()

# Group by agent and calculate their average call duration and calls per day
agent_performance = df.groupby('Agent').agg({'Call Duration': 'mean', 'Avg Calls/Day': 'mean'}).reset_index()

# Determine which approach is better based on average call duration and calls per day
better_approach = agent_performance.loc[agent_performance['Call Duration'].idxmin()]

# Display the updated results
print("\nAverage Call Duration:", average_call_duration)
print("Average Calls/Day:", average_calls_per_day)
print("\nAgent Performance:")
print(agent_performance)
print("\nBetter Approach:")
print(better_approach)

# Plot: Call Duration by Agent
plt.figure(figsize=(10, 6))
plt.bar(agent_performance['Agent'], agent_performance['Call Duration'], color='skyblue')
plt.title('Average Call Duration by Agent')
plt.xlabel('Agent')
plt.ylabel('Call Duration (seconds)')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('/Users/tommygrace/Downloads/Assesment/Assessment_03/output/call_duration_by_agent.png')
plt.close()

# Plot: Average Calls/Day by Agent
plt.figure(figsize=(10, 6))
plt.bar(agent_performance['Agent'], agent_performance['Avg Calls/Day'], color='lightgreen')
plt.title('Average Calls per Day by Agent')
plt.xlabel('Agent')
plt.ylabel('Average Calls/Day')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('/Users/tommygrace/Downloads/Assesment/Assessment_03/output/calls_per_day_by_agent.png')
plt.close()
