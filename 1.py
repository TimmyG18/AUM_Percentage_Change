import pandas as pd

excel_file='previous.xls'
output_file='OldCleansed.xlsx'
df=pd.read_excel(excel_file)
df.drop(index=[0,1,2,3,4,5,6,9], inplace=True)
df.drop(columns=['Unnamed: 2','Unnamed: 3','Unnamed: 0','Unnamed: 4','Unnamed: 5','Unnamed: 6','Unnamed: 7', 'Unnamed: 8','Unnamed: 10','Unnamed: 12','Unnamed: 14','Unnamed: 13'], inplace=True)
df=df.dropna()
df=df.reset_index(drop=True)
df.drop(index=[0], inplace=True)
df.columns=['Fund Name','Units Owned','Bid Price','Amount(USD)']

df['Units Owned']=df.groupby('Fund Name')['Units Owned'].transform('sum')
df=df.drop_duplicates(subset='Fund Name')
#start_index=0
#for index, row in df.iterrows():
  #Fund_Name=row.iloc[0]
  #Units_Owned=row.iloc[1]
  #start_index=start_index+1
  #for i, r in df.iloc[start_index:].iterrows():
    #if Fund_Name==r.iloc[0]:
      #Units_Owned=Units_Owned+r.iloc[1]
      #df.iloc[df.index.get_loc(index), 1]=Units_Owned
      #df.drop(i, inplace=True)
df.sort_values(by=df.columns[0], inplace=True)
df=df.reset_index(drop=True)
df.to_excel(output_file, index=False)
