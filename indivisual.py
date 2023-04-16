import pandas as pd
df = pd.read_excel('input.xlsx','Sheet2',skiprows=6)
df['name'].str.strip()
df.sort_values(by=['name'], ascending=True, inplace=True, key=lambda col: col.str.lower())
df.sort_values(by=['total_statements','total_reasons'], ascending=False, inplace=True)
a=[i for i in range(1,len(df['name'])+1)]
df['Rank']=a
df.rename(columns={"name": "Name","uid":"UID","total_statements":"No. of Statements","total_reasons":"No. of Reasons"}, inplace=True)
df.drop(columns=['S No'])
cols=["Rank","Name","UID","No. of Statements","No. of Reasons"]
df=df[cols]
print(df)

with pd.ExcelWriter('input.xlsx', mode='a') as writer:
    df.to_excel(writer, sheet_name='Sheet5', index=False)