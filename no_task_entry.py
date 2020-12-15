import pandas as pd
import pyodbc as db

connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=;'                     # Provide server
                        'DATABASE=;'                   # provide database 
                        'UID=;PWD=')                   # provide username and password

no_task_df = pd.read_sql_query(""" 
    Declare @FDate varchar(10)=CONVERT(varchar(8),DATEFROMPARTS(YEAR(GETDATE()-1), 1, 1),112)
    Declare @fromdate varchar(8)= CONVERT(varchar(8), dateAdd(day,0,GETDATE()-1), 112)

     select   asign_id, asign_name as 'Employee Name', max(superviser_name) as 'Supervised By', 
     CONVERT(varchar(12), MAX(task_date),113) as 'Last Entry Date',
     (select
     COUNT  (DISTINCT DATE) as toatlDays  from WorkingDaysActivity
     where LEFT(date,6) between @FDate and @fromdate
     and WorkingStatus <> '0') - COUNT (DISTINCT task_date) as 'No Entry Days'
     from tasks
     where
     LEFT(CONVERT(varchar(8), dateAdd(day,0,task_date), 112),6) between @FDate and @fromdate
     group by asign_id, asign_name
     order by 'No Entry Days' DESC, asign_id ASC  """, connection)

no_task_df = no_task_df[no_task_df.NoEntry >= 1]
no_task_df.to_excel('no_task.xlsx', index=False)
print('No task Excel file created ')



# #  ----------------------------------------------------
# # Egnore this section . This is for counting govt holyday
# # ------------------------------------------------------

# work_data = pd.read_sql_query(""" select * from WorkingDaysActivity
#         where Date <='20200426'""", connection)
# govt_holy_day = work_data[(work_data.WorkingStatus == 0) & (work_data.Days!='Friday')]
#
# print(govt_holy_day.WorkingStatus.count()
# )