import pandas as pd

# Reading input from Excel file
df1 = pd.read_excel('input.xlsx','Sheet1',skiprows=9)
df2 = pd.read_excel('input.xlsx','Sheet2',skiprows=6)
abc=[]
for i in range(len(df1)):
    if(df1.loc[i,'User ID']==df2.loc[i,'uid']):
        abc.append(df1.loc[i,'Team Name'])

df2['Team Name']=abc
# print(df2)

ans={}
ans['Thinking Teams Leaderboard']=pd.unique(df2['Team Name'])
ans['Average Statements']=[]
ans['Average Reasons']=[]
for i in range(len(pd.unique(df2['Team Name']))):
    ans['Average Statements'].append(round(df2.loc[df2['Team Name']==ans['Thinking Teams Leaderboard'][i],'total_statements'].mean(),2))
    ans['Average Reasons'].append(round(df2.loc[df2['Team Name']==ans['Thinking Teams Leaderboard'][i],'total_reasons'].mean(),2))

# print(ans)
finaldf=pd.DataFrame(ans)
finaldf.sort_values(by=['Average Statements','Average Reasons'], ascending=False, inplace=True)

a=[i for i in range(1,len(pd.unique(df2['Team Name']))+1)]
finaldf['Team Rank']=a
cols = ['Team Rank', 'Thinking Teams Leaderboard', 'Average Statements', 'Average Reasons']
finaldf=finaldf[cols]
print(finaldf)

with pd.ExcelWriter('input.xlsx', mode='a') as writer:
    finaldf.to_excel(writer, sheet_name='Sheet4', index=False)

# writer.save()
