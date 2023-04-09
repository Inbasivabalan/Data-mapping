import pandas as pd

# Read input sheet 1
df1 = pd.read_excel('input_sheet1.xlsx')

# Count statements and reasons
df1['num_statements'] = df1.iloc[:, 4:].count(axis=1)
df1['num_reasons'] = df1.iloc[:, 4:].apply(lambda x: x[x.notna()].nunique(), axis=1)

# Sort by leaderboard
df1.sort_values(by=['Team Name', 'num_statements', 'num_reasons'], ascending=[True, False, False], inplace=True)

# Output leaderboard 1
df1_output = df1[['S No', 'Name', 'Team Name', 'User ID', 'num_statements', 'num_reasons']]
print(df1_output.to_string(index=False))
df1_output.to_excel('output_sheet1.xlsx', index=False)

# Read input sheet 2
df2 = pd.read_excel('input_sheet2.xlsx')

# Count statements and reasons
df2['num_statements'] = df2.iloc[:, 4:].count(axis=1)
df2['num_reasons'] = df2.iloc[:, 4:].apply(lambda x: x[x.notna()].nunique(), axis=1)

# Sort by leaderboard
df2.sort_values(by=['name', 'uid', 'num_statements', 'num_reasons'], ascending=[True, True, False, False], inplace=True)

# Output leaderboard 2
df2_output = df2[['S No', 'name', 'uid', 'total_statements', 'total_reasons', 'num_statements', 'num_reasons']]
print(df2_output.to_string(index=False))
df2_output.to_excel('output_sheet2.xlsx', index=False)
