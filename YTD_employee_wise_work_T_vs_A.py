import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import sys
import pyodbc as db
import openpyxl
import xlrd
import os
import smtplib
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import pytz
import matplotlib.patches as patches
import pyodbc as db
import xlrd
from PIL import Image
import calendar
import sys
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


last_day_chart_df = pd.read_sql_query("""Select distinct LTRIM(a.asign_name) as 'Employee Name',c.short_name as 'Short_Name', ISNULL(b.Total,0) as work_hour
from
(select distinct asign_name,asign_id from tasks
where task_date >= CONVERT(varchar(10), getdate()-50,126)
group by asign_name,asign_id) as a
left join
(select asign_name,asign_id,
sum((DATEDIFF(second, start_time, end_time))/3600) as Total
from tasks
where YEAR (task_date) = YEAR (CONVERT(varchar(10), getdate()-1,126))
group by asign_name,asign_id) as b
on a.asign_id=b.asign_id
left join
(select distinct [name], short_name,[user_id] from employees
where status =1 ) as c
on a.asign_id = c.[user_id]
where a.asign_name<>'Md. Nazmus Sadat'
and a.asign_name<>'Md. Mahmud Ul Islam'
and a.asign_name<>'Md. Al Ameen'
and a.asign_name<>'Md. Shahriar Parvez'
and a.asign_name<>'Naznin Akter'
and a.asign_name<>'Almas Md. Estiak - Suzon'
and a.asign_name<>'Sanchita Saha'
and a.asign_name<>'Goutam Tarafder'
order by work_hour DESC""", connection)

from datetime import date


# # --------------------------------------------------
# # Current Date and Months Start Date
# # --------------------------------------------------
todayDate = datetime.date.today()
resultDate = todayDate.replace(day=1)
if ((todayDate - resultDate).days > 25):
    resultDate = resultDate + relativedelta(months=0)
month_start_date = int(re.sub("\-",'', str(resultDate)))
current_date = int(re.sub("\-",'', str(todayDate)))
#print('Month Start Date = ', month_start_date)
#print('Current Date = ', current_date)

date = datetime.date.today()
years_first_day = int(str(date.year)+'0101')
#print('Years Start Date = ', years_first_day)
# # ---------------------------------------------------------------
# # Current Months Number of Off Day, Half day and Full Working Day
# # ---------------------------------------------------------------
# Read the dataset
data = pd.read_excel('./working_date.xlsx')
#print(data.columns)
# Sorting data set
total_offday_in_month = data[
    (data.maindate.between(years_first_day, current_date, inclusive=True))
    & (data.working_status == 0)
    ]

Public_off_days_in_year = total_offday_in_month.maindate.count()

#print('Total Off Day In this Year = ', Public_off_days_in_year)

# # Total Saturday in a month
total_halfday_in_month = data[
    (data.maindate.between(years_first_day, current_date, inclusive=True))
    & (data.working_status == 1)
    ]

weekly_saturdays_in_year = total_halfday_in_month.maindate.count()

#print('Total Half Day In this Year = ', weekly_saturdays_in_year)
# # Total Full Day in a month


total_Fullday_in_month = data[
    (data.maindate.between(years_first_day, current_date, inclusive=True))
    & (data.working_status == 2)
    ]

full_days_in_year = total_Fullday_in_month.maindate.count()

#print('Total Full Day In this Year = ', full_days_in_year)

personel_df = pd.read_sql_query("""select count(status) as amount from users
where status=1 and
name != 'Mohammad Shakhawat Hossain Biswas'
""", connection)


total_personel = int(personel_df['amount'])
#total_personel=45
target=[]

aii=(8*full_days_in_year)+(4*weekly_saturdays_in_year)
for loop_value in range(0,total_personel):
    target.append(aii)
#print(target)

labels = last_day_chart_df['Short_Name'].values.tolist()
sizes = last_day_chart_df['work_hour'].values.tolist()
#print(labels)
#print(sizes)

plt.subplots(figsize=(12.81, 5))
#colors=['#03254c','#1c3a5d','#35506f','#4e6681','#677c93','#8192a5','#9aa7b7','#9aa7b7','#b3bdc9','#b3bdc9','#ccd3db','#e5e9ed']
#colors=['#4c0099','#5d19a3','#6f32ad','#814cb7','#9366c1','#a57fcc','#a57fcc','#b799d6','#c9b2e0','#c9b2e0','#dbccea','#ede5f4']
colors='#9366c1'
plt.plot(labels, target, color='orange')



plt.bar(labels, sizes,color=colors, align='center',width=.8, alpha=1)

for x,y in zip(labels,sizes):

    label = "{:.1f}".format(y)

    plt.annotate(label, # this is the text
                 (x,y), # this is the point to label
                 textcoords="offset points", # how to position the text
                 xytext=(0,3), # distance from text to points (x,y)
                 rotation=45,
                ha='center')

plt.title('18. YTD Employee Wise Target vs Achieved Hour', ha='center',color='#3238a8', fontsize=14, fontweight='bold')
plt.xticks(labels, rotation='vertical')
#plt.xlabel('Employee', fontsize=14)
plt.ylabel('Work Hour', fontsize=14)
plt.tight_layout()
plt.legend(['Target Hour','Achieved Hour'], loc='upper right')
#plt.grid(True)
plt.savefig('YTD_employee_wise_work_T_vs_A.png')
#plt.show()
plt.close()
print('18. YTD Employee Wise Target vs Achieved Hour generated')

