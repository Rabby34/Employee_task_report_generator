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
  when  workfor_name = 'PALO' then 'b' 
 when  workfor_name = 'TDCL' then 'c' 
  when  workfor_name = 'TEL' then 'd' 
   when  workfor_name = 'BLL' then 'e' 
 when  workfor_name = 'TCPL' then 'f' 
  when  workfor_name = 'TBL' then 'g' 
  when  workfor_name = 'TCR' then 'h' 
  when  workfor_name = 'TEA' then 'i' 
  when  workfor_name = 'BEIL' then 'j' 
  when  workfor_name = 'TFL' then 'K' 
  when  workfor_name = 'ABC' then 'l' 
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

Company_names = last_day_chart_df['company'].values.tolist()
hours = last_day_chart_df['work_hour'].values.tolist()
#print('Company names ', Company_names)
#print('already achieved work hours ', hours)



# Pie chart, where the slices will be ordered and plotted counter-clockwise:
labels = Company_names
sizes = hours

#fig, ax = plt.subplots(figsize=(16,9), subplot_kw=dict(aspect="equal"))

x = np.char.array(labels)
y = np.array(sizes)
explode=(.05,.05,.05,.05,.05,.05,.05,.05,.05,.05,.05,.05)
# colors = ['yellowgreen','red','gold','lightskyblue','white','lightcoral','blue','pink', 'darkgreen','yellow','grey','violet','magenta','cyan']
porcent = 100.*y/y.sum()

patches, texts = plt.pie(y, explode=explode, shadow=False, startangle=90, radius=1,rotatelabels=True,labeldistance=1)
labels = ['{0} - {1:1.1f} %'.format(i,j) for i,j in zip(x, porcent)]
#print(labels)
#sys.exit()
sort_legend = False
if sort_legend:
    patches, labels, dummy =  zip(*sorted(zip(patches, labels, y),
                                          key=lambda x: x[2],
                                          reverse=True))

plt.legend(patches, labels, loc='right center', bbox_to_anchor=(1,.92),
           fontsize=8)
plt.text(0.2, 1.4, '16. MTD Working Hour Percentage Per Company', ha='center',color='#3238a8', fontsize=14, fontweight='bold')
#plt.title('Working Hour Percentage Per Company',fontsize='14',color='white', fontweight='bold',backgroundcolor='#6a9aba')
plt.tight_layout()
#plt.show()
plt.savefig('pie_for_percentage.png')
plt.close()

print('16. MTD Working Hour Percentage Per Company')