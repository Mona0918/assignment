import pandas as pd
gsheetid='1NRmXI_L0W89yj2uow4UvhVGVEDLDd9szpbE-9v7enpI'
gsheet_url1=f"https://docs.google.com/spreadsheets/d/{gsheetid}/gviz/tq?tqx=out:csv&sheet=data1"
gsheet_url2=f"https://docs.google.com/spreadsheets/d/{gsheetid}/gviz/tq?tqx=out:csv&sheet=data2"
df1=pd.read_csv(gsheet_url1)
df2=pd.read_csv(gsheet_url2)
teams=[]
totalTeam={}
userIDs=[]
flag=0
dict={}
df_list=list(df1.loc[:,"Team Name"])
df_ID=list(df1.loc[:,"User ID"])
teams=list(set(df_list))
for j in range(len(teams)):
    for k in range(len(df_list)):
        if teams[j] == df_list[k]:
            flag=1
            userIDs=userIDs+[df_ID[k]]
        else:
            flag=0
    dict[teams[j]]=userIDs
    userIDs=[]
for i in teams:
    totalTeam[i]=df_list.count(i)
df_list1=list(df2.loc[:,"User ID"])
df_stmt=list(df2.loc[:,"total_statements"])
df_reason=list(df2.loc[:,"total_reasons"])
ID_stmt={}
ID_reason={}
for k in range(len(df_list1)):
    ID_stmt[df_list1[k]]=df_stmt[k]
    ID_reason[df_list1[k]]=df_reason[k]
key1=list(ID_stmt.keys())
key2=list(ID_reason.keys())
sum_stmt=0
avg_stmt=0
sum_reason=0
avg_reason=0
team_avgstmt={}
team_avgreason={}
for i in dict:
    for j in dict[i]:
        for k in key1:
            if j == k:
                sum_stmt=sum_stmt+ID_stmt[j]
        avg_stmt=sum_stmt/totalTeam[i]
    team_avgstmt[i]=avg_stmt
    sum_stmt=0
    avg_stmt=0
for i in dict:
    for j in dict[i]:
        for k in key2:
            if j == k:
                sum_reason=sum_reason+ID_reason[j]
        avg_reason=sum_reason/totalTeam[i]
    team_avgreason[i]=avg_reason
    sum_reason=0
    avg_reason=0
stmt_reason={}
for i in teams:
    sum=team_avgstmt[i]+team_avgreason[i]
    stmt_reason[i]=[team_avgstmt[i]]+[team_avgreason[i]]+[sum]
sorted_stmt_reason=sorted(stmt_reason.items(),reverse=True)
ky=[]
for i in sorted_stmt_reason:
    for j in range(len(i)):
        if j==0:
            ky=ky+[i[j]]
s=[]
r=[]
for i in sorted_stmt_reason:
    for j in range(len(i)):
        if j==1:
            for k in range(len(i[j])):
                if k == 0:
                    str_avg_stmt=f'{i[j][k]:.2f}'
                    s=s+[str_avg_stmt]
                elif k == 1:
                    str_avg_reason=f'{i[j][k]:.2f}'
                    r=r+[str_avg_reason]
Rank=[]
for i in range(1,len(teams)+1):
    Rank=Rank+[i]
Leaderboard=ky
stmt=s
reason=r
columns=['Team Rank','Thinking Teams Leaderboard','Average_Statements','Average_Reasons']
df=pd.DataFrame(list(zip(Rank,Leaderboard,stmt,reason)),columns=columns)
df.to_excel("C:/Users/Mona/PycharmProjects/agile/Output sheets/Leaderboard.xlsx",sheet_name="LeaderBoard Teamwise(Output)",columns=columns, index=False)
