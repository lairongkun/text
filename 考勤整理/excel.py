import pandas as pd

df = pd.read_excel(r"上下班打卡_日报.xlsx",
                   usecols=["姓名", "部门", "时间", "打卡时间", "打卡时间.1"],
                   header=2)
df.head()

df["公司打卡"] = df["打卡时间"] + " \n" + df["打卡时间.1"]
df["姓名-部门"]=df["姓名"] + "-" + df["部门"]
df.head()

result = df.pivot(index='姓名-部门', columns='时间', values='公司打卡').reset_index()

result.head()

result[["姓名", "部门"]] = result['姓名-部门'].str.split('-', expand=True)
result.head()

columns = result.columns.values
columns_r = []
columns_r.extend(columns[-2:])
columns_r.extend(columns[1:-2])
result = result[columns_r]

result.to_excel("打卡记录.xlsx", index=False)
