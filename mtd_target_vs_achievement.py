import pandas as pd
## --------------------------------------------
# # ---------  Months Start Date --------------
# # -------------------------------------------
import datetime
from dateutil.relativedelta import relativedelta
import re
from datetime import date
from pandas.tseries import offsets
import matplotlib.pyplot as plt
import os
import pyodbc as db
import sys

connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=;'                     # Provide server
                        'DATABASE=;'                   # provide database 
                        'UID=;PWD=')                   # provide username and password
cursor = connection.cursor()
dirpath = os.path.dirname(os.path.realpath(__file__))

last_day_chart_df = pd.read_sql_query("""select top 12 (lTRIM(Atask.workfor_name)) as company, ISNULL(sum(Btask.Total),0) as work_hour from 
(select  distinct ltrim(workfor_name) as workfor_name ,
case
when  workfor_name = 'SKF' then 'a' 
when  workfor_name = 'TDCL' then 'b' 
when  workfor_name = 'TCPL' then 'c' 
when  workfor_name = 'TBL' then 'd' 
when  workfor_name = 'PALO' then 'e' 
when  workfor_name = 'ABC' then 'f' 
when  workfor_name = 'TCR' then 'g' 
when  workfor_name = 'TEL' then 'h' 
when  workfor_name = 'BLL' then 'i' 
when  workfor_name = 'BEIL' then 'j' 
when  workfor_name = 'TFL' then 'K' 
when  workfor_name = 'TEA' then 'l'  
else 'm'  end as sqnc
from tasks
where workfor_name <>''
) as Atask
left join 

(select workfor_name,
sum(CAST((DATEDIFF(minute, start_time, end_time)) as  decimal(5, 2))/60) as Total 
from tasks
where Month(task_date)  =  Month(CONVERT(varchar(10), getdate()-1,126))
and YEAR(task_date)  =  YEAR(CONVERT(varchar(10), getdate()-1,126))
group by workfor_name) as Btask
on ltrim(Atask.workfor_name) = ltrim(Btask.workfor_name)
group by Atask.workfor_name,Atask.sqnc
order by sqnc asc""", connection)

# # --------------------------------------------------
# # Current Date and Months Start Date
# # --------------------------------------------------
todayDate = datetime.date.today()
resultDate = todayDate.replace(day=1)
if ((todayDate - resultDate).days > 25):
    resultDate = resultDate + relativedelta(months=0)
month_start_date = int(re.sub("\-", '', str(resultDate)))
current_date = int(re.sub("\-", '', str(todayDate)))
#print('Month Start Date = ', month_start_date)
#print('Current Date = ', current_date)

#-------------------yesterday date-------------------

from datetime import datetime, timedelta
yesterday_date=datetime.strftime(datetime.now() - timedelta(1), '%Y-%m-%d')

last_date=int(re.sub("\-", '', str(yesterday_date)))
#print('yesterday date = ',last_date)

# # ---------------------------------------------------------------
# # Current Months Number of Off Day, Half day and Full Working Day
# # ---------------------------------------------------------------
# Read the dataset
data = pd.read_excel('./working_date.xlsx')
#print(data.columns)
# Sorting data set
total_offday_in_month = data[
    (data.maindate.between(month_start_date, last_date, inclusive=True))
    & (data.working_status == 0)
    ]

Public_off_days = total_offday_in_month.maindate.count()

#print('Total Off Day In this Month = ', Public_off_days)

# # Total Saturday in a month
total_halfday_in_month = data[
    (data.maindate.between(month_start_date, last_date, inclusive=True))
    & (data.working_status == 1)
    ]

weekly_saturdays = total_halfday_in_month.maindate.count()

#print('Total Half Day In this Month = ', weekly_saturdays)
# # Total Full Day in a month


total_Fullday_in_month = data[
    (data.maindate.between(month_start_date, last_date, inclusive=True))
    & (data.working_status == 2)
    ]

full_days = total_Fullday_in_month.maindate.count()

#print('Total Full Day In this Month = ', full_days)

personel_df = pd.read_sql_query("""select count(status) as amount from users
where status=1 and
name != 'Mohammad Shakhawat Hossain Biswas'
""", connection)

total_personel = int(personel_df['amount'])
#total_personel = 45

#print('Total personel ', total_personel)

Total_mtd_working_hour_between_dates = (full_days * 8 * total_personel) + (weekly_saturdays * 4 * total_personel)

#print('Mtd total hour ', Total_mtd_working_hour_between_dates)

target_percentages = [22, 16, 7, 7, 17, 1, 5, 10, 7, 3, 2, 3]

#print('Target percentage per company ', target_percentages)

target_hours = Total_mtd_working_hour_between_dates
the_array = []
for percents in target_percentages:
    # print(percents)
    value = int(round((target_hours * percents) / 100))
    the_array.append(value)
    value = 0
#print('Company wise targets ', the_array)

Company_names = last_day_chart_df['company'].values.tolist()
hours = last_day_chart_df['work_hour'].values.tolist()
#print('Company names ', Company_names)
#print('already achieved work hours ', hours)

for i in range(0, len(hours)):
    hours[i] = int(hours[i])
#print('achived hour into int :',hours)
colors = '#9366c1'

plt.subplots(figsize=(12.81, 4.8))
plt.plot(Company_names, the_array, color='orange', marker='o')

for x, y in zip(Company_names, the_array):
    label = "{:.0f}".format(y)

    plt.annotate(label,  # this is the text
                 (x, y),  # this is the point to label
                 textcoords="offset points",
                 color='red',  # how to position the text
                 xytext=(12, 3),  # distance from text to points (x,y)
                 ha='center')

plt.bar(Company_names, hours, color=colors, align='center', width=.8, alpha=1)

for x, y in zip(Company_names, hours):
    label = "{:.0f}".format(y)

    plt.annotate(label,  # this is the text
                 (x, y),  # this is the point to label
                 textcoords="offset points",  # how to position the text
                 xytext=(0, 3),  # distance from text to points (x,y)
                 ha='center')

plt.title('6. MTD Target vs Achieved Hour', ha='center',color='#3238a8', fontsize=14, fontweight='bold')
#plt.xlabel('Company', fontsize=14)
plt.ylabel('Work Hour', fontsize=14)
plt.tight_layout()
# plt.grid(True)
plt.legend(['Target Hour','Achieved Hour'])
#plt.show()
plt.savefig('mtd_target_vs_achieved_hour.png')
plt.close()
print('6. MTD Target vs Achieved Hour generated')