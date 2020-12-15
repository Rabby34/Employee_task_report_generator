import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pyodbc as db
import xlrd
from PIL import Image
from matplotlib.patches import Patch
import datetime as dd

from datetime import datetime
import os

import pyodbc as db


connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=;'                     # Provide server
                        'DATABASE=;'                   # provide database 
                        'UID=;PWD=')                   # provide username and password
cursor = connection.cursor()
dirpath = os.path.dirname(os.path.realpath(__file__))

support_df = pd.read_sql_query("""Select sum(CAST((DATEDIFF(minute, start_time, end_time)) as  decimal(5, 2))/60) as 'Support Working Hour'
 from tasks
where task_date =CONVERT(varchar(10), getdate()-1,126)
and taskcategory_name='Support & Maintenance'""", connection)

development_df = pd.read_sql_query("""Select sum(CAST((DATEDIFF(minute, start_time, end_time)) as  decimal(5, 2))/60) as 'Development Working Hour'
 from tasks
where task_date =CONVERT(varchar(10), getdate()-1,126)
and taskcategory_name='Research & Development'""", connection)



# Pie chart, where the slices will be ordered and plotted counter-clockwise:
support = int(support_df['Support Working Hour'])
development = int(development_df['Development Working Hour'])

data = [support, development]
print(data)

legend_element = [Patch(facecolor='#EC6B56', label='Support & \nMaintenance'),
                  Patch(facecolor='#FFC154', label='Research & \nDevelopment')]


colors = ['#EC6B56', '#FFC154']


DataLabel=['Support \n & \n Maintenance','Research \n & \n Development']
explode=(0.05,0)

fig, ax = plt.subplots()
wedges, labels, autopct = ax.pie(data, explode=explode, colors=colors, autopct='%.1f%%', startangle=90, pctdistance=.5)
plt.setp(autopct, fontsize=12, color='black', fontweight='bold')


plt.text(0, 1.28, '1. Last Day Working Area Wise Percentage', ha='center',color='#3238a8', fontsize=14, fontweight='bold')
ax.axis('equal')
plt.legend(handles=legend_element, loc='upper right', fontsize=9)

plt.tight_layout()
plt.savefig('Last_day_SupportvsDevelopment.png')
#plt.show()
plt.close()
print('1. Last Day Working Area Wise Percentage generated')