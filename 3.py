import pandas as pd
file1='OldCleansed.xlsx'
file2='NewCleansed.xlsx'
output_file='result.xlsx'
df1=pd.read_excel(file1)
df2=pd.read_excel(file2)
merged_df = df1.copy(deep=True)
merged_df = pd.concat([df1, pd.DataFrame(columns=[f'Column {i+1}' for i in range(6)])], axis=1)
merged_df.columns=['Fund Name', 'Former Units', 'Former Bid Price', 'Former Amount(USD)', 'Later Units', 'Later Bid Price', 'Later Amount(USD)', '% Percentage Change in Bid', '% Percentage Change in Amount', '% Effect of Price Change on the Total Change of AUM']
for i, row in merged_df.iterrows():
    fund_name = row['Fund Name']
    if fund_name in df2['Fund Name'].values:
        merged_df.loc[i, 'Later Units'] = df2.loc[df2['Fund Name'] == fund_name, 'Units Owned'].values[0]
        merged_df.loc[i, 'Later Bid Price'] = df2.loc[df2['Fund Name'] == fund_name, 'Bid Price'].values[0]
        merged_df.loc[i, 'Later Amount(USD)'] = df2.loc[df2['Fund Name'] == fund_name, 'Amount(USD)'].values[0]
    else:
        merged_df.loc[i, 'Later Units'] = 0
        merged_df.loc[i, 'Later Bid Price'] = 0
        merged_df.loc[i, 'Later Amount(USD)'] = 0
for r, row in df2.iterrows():
    fund_name1 = row['Fund Name']
    if fund_name1 in merged_df['Fund Name'].values:
      continue
    else:
      new_row = {
          'Fund Name':fund_name1,
          'Former Units':0,
          'Former Bid Price':0,
          'Former Amount(USD)':0,
          'Later Units':df2.loc[r, 'Units Owned'],
          'Later Bid Price':df2.loc[r, 'Bid Price'],
          'Later Amount(USD)':df2.loc[r, 'Amount(USD)'],
      }
      merged_df = pd.concat([merged_df, pd.DataFrame([new_row])], ignore_index=True)
for s, row in merged_df.iterrows():
    former_bid = row['Former Bid Price']
    later_bid = row['Later Bid Price']
    former_amount = row['Former Amount(USD)']
    later_amount = row['Later Amount(USD)']

    if former_bid == 0 or later_bid == 0:
        merged_df.at[s, '% Percentage Change in Bid'] = 'NA'
    else:
        merged_df.at[s, '% Percentage Change in Bid'] = ((later_bid - former_bid) / former_bid) * 100

    if former_amount == 0 or later_amount == 0:
        merged_df.at[s, '% Percentage Change in Amount'] = 'NA'
    else:
        merged_df.at[s, '% Percentage Change in Amount'] = ((later_amount - former_amount) / former_amount) * 100
def weighted_average(merged_df):
    average = 0
    total = df2['Amount(USD)'].sum()
    for u, row in merged_df.iterrows():
        effect_pct = merged_df.loc[u, '% Effect of Price Change on the Total Change of AUM'] 
        if pd.isnull(effect_pct) or effect_pct == 0:
            continue    
        weight = merged_df.loc[u, 'Later Amount(USD)'] / total
        average += weight * effect_pct  
    return average 
weighted_avg = weighted_average(merged_df)        
merged_df.at[15, 15] = 'Former Total AUM: '
merged_df.at[15, 16] = df1['Amount(USD)'].sum()
merged_df.at[16, 15] = 'Later Total AUM: '
merged_df.at[16, 16] = df2['Amount(USD)'].sum()
merged_df.at[17, 15] = '% Percentage Change in AUM: '
merged_df.at[17, 16] = ((merged_df.at[16, 16]-merged_df.at[15, 16]) / merged_df.at[15, 16])*100
merged_df.at[19, 15] = '* Formula to get the % effect of price change on the total change of AUM=(% change in bid)*(Former Amount)/(Change in total AUM). Negative number indicates AUM change is negatively correlated with the movement of price.'
merged_df.at[18, 15] = 'Weighted average effect of price change on change of total AUM: '
merged_df.at[18, 16] = weighted_avg
merged_df.at[20, 15] = '* Weighted average effect of price change on change of total AUM=(individual later amount)/(later total AUM)*individual effect of price change.'
for t, row in merged_df.iterrows():
    if merged_df.loc[t, '% Percentage Change in Bid'] != 'NA':
        merged_df.loc[t, '% Effect of Price Change on the Total Change of AUM'] = ((merged_df.loc[t, '% Percentage Change in Bid'] * merged_df.loc[t, 'Former Amount(USD)']) / (merged_df.at[16, 16]-merged_df.at[15, 16]))
    else:
        continue
merged_df.to_excel(output_file, index=False)
