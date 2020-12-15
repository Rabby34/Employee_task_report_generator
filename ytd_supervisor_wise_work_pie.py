import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import sys
import os
import pyodbc as db


connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=;'                     # Provide server
                        'DATABASE=;'                   # provide database 
                        'UID=;PWD=')                   # provide username and password
cursor = connection.cursor()
dirpath = os.path.dirname(os.path.realpath(__file__))


Tawhid_df = pd.read_excel('employee_list.xlsx', 'Tawhid_sir')
Names_tawhid = Tawhid_df['Tawhid_sir_employee_list']
Tawhid_sir_people_Name = Names_tawhid.values.tolist()

Zubair_df = pd.read_excel('employee_list.xlsx', 'Zubair_sir')
Names_zubair = Zubair_df['Zubair_sir_employee_list']
Zubair_sir_people_Name = Names_zubair.values.tolist()

Yakub_df = pd.read_excel('employee_list.xlsx', 'Yakub_sir')
Names_yakub = Yakub_df['Yakub_sir_employee_list']
Yakub_sir_people_Name = Names_yakub.values.tolist()



main_yakub_task_count_df=pd.DataFrame()
main_yakub_df = pd.DataFrame()
emplist_of_yakub = Yakub_sir_people_Name
for employ in range(len(emplist_of_yakub)):
    emp_name = (emplist_of_yakub[employ])
    #print(emp_name)
    yakub_ali_df = pd.read_sql_query("""select
ISNULL(sum(DATEDIFF(mm, start_time, end_time))+sum(DATEDIFF(hh, start_time, end_time)),0) as Total 
from tasks
                where  YEAR(task_date) =YEAR(CONVERT(varchar(10), getdate()-1,126)) 
                and asign_name like ?""", connection,params={emp_name})
    main_yakub_df = main_yakub_df.append(yakub_ali_df)

    yakub_task_count_df = pd.read_sql_query("""Select ISNULL(count(time_diff),0) as 'Task Count' 
        from tasks
         where  YEAR(task_date) =YEAR(CONVERT(varchar(10), getdate()-1,126))
         and asign_name like ?""", connection, params={emp_name})

    main_yakub_task_count_df = main_yakub_task_count_df.append(yakub_task_count_df)

main_yakub_df.columns = range(main_yakub_df.shape[1])
main_yakub_task_count_df.columns = range(main_yakub_task_count_df.shape[1])
#print(main_yakub_task_count_df.values)
#print(main_yakub_df.values)
Total_of_yakub=sum(main_yakub_df.values)
#print(Total_of_yakub)
Total_yakub_task=sum(main_yakub_task_count_df.values)

#print(Total_yakub_task)


main_zubair_task_count_df=pd.DataFrame()
main_zubair_df = pd.DataFrame()
emplist_of_zubair = Zubair_sir_people_Name
for employ2 in range(len(emplist_of_zubair)):
    emp_name2 = (emplist_of_zubair[employ2])
    #print(emp_name2)
    Zubair_df = pd.read_sql_query("""select
ISNULL(sum(DATEDIFF(mm, start_time, end_time))+sum(DATEDIFF(hh, start_time, end_time)),0) as Total 
from tasks
                where  YEAR(task_date) =YEAR(CONVERT(varchar(10), getdate()-1,126))
                and asign_name like ?""", connection,params={emp_name2})
    main_zubair_df = main_zubair_df.append(Zubair_df)

    zubair_task_count_df = pd.read_sql_query("""Select count(time_diff) as 'Task Count' 
            from tasks
             where  YEAR(task_date) =YEAR(CONVERT(varchar(10), getdate()-1,126))
             and asign_name like ?""", connection, params={emp_name2})

    main_zubair_task_count_df = main_zubair_task_count_df.append(zubair_task_count_df)

main_zubair_df.columns = range(main_zubair_df.shape[1])

main_zubair_task_count_df.columns = range(main_zubair_task_count_df.shape[1])
#print(main_zubair_df.values)
Total_of_zubair=sum(main_zubair_df.values)
#print(Total_of_zubair)
Total_zubair_task=sum(main_zubair_task_count_df.values)
#print(Total_zubair_task)


main_tawhid_task_count_df=pd.DataFrame()
main_tawhid_df = pd.DataFrame()
emplist_of_tawhid = Tawhid_sir_people_Name
for employ3 in range(len(emplist_of_tawhid)):
    emp_name3 = (emplist_of_tawhid[employ3])
    #print(emp_name3)
    Tawhid_df = pd.read_sql_query("""select
ISNULL(sum(DATEDIFF(mm, start_time, end_time))+sum(DATEDIFF(hh, start_time, end_time)),0) as Total 
from tasks
                where  YEAR(task_date) =YEAR(CONVERT(varchar(10), getdate()-1,126)) 
                and asign_name like ?""", connection,params={emp_name3})
    main_tawhid_df = main_tawhid_df.append(Tawhid_df)

    tawhid_task_count_df = pd.read_sql_query("""Select count(time_diff) as 'Task Count' 
                from tasks
                 where  YEAR(task_date) =YEAR(CONVERT(varchar(10), getdate()-1,126))
                 and asign_name like ?""", connection, params={emp_name3})

    main_tawhid_task_count_df = main_tawhid_task_count_df.append(tawhid_task_count_df)

main_tawhid_df.columns = range(main_tawhid_df.shape[1])

main_tawhid_task_count_df.columns = range(main_tawhid_task_count_df.shape[1])
#print(main_tawhid_df.values)
Total_of_tawhid=sum(main_tawhid_df.values)
#print(Total_of_tawhid)
Total_tawhid_task=sum(main_tawhid_task_count_df.values)
#sys.exit()
#print(Total_tawhid_task)

values_for_pie=[Total_of_yakub[0],Total_of_zubair[0],Total_of_tawhid[0]]
#print(values_for_pie)

Task_count_array=[Total_yakub_task[0],Total_zubair_task[0],Total_tawhid_task[0]]
label_list=['Yakub Ali','Md. Zubair','Md. Tawhid']
#print(label_list)

#sys.exit()

values_of_label = np.char.array(label_list)
values_to_plot = np.array(values_for_pie)
#print(values_to_plot)

labels = ['{0} - \n{1:1.0f} Tasks'.format(i,j) for i,j in zip(values_of_label, Task_count_array)]

explode=(.05,0,0)
porcent = 100.*values_to_plot/values_to_plot.sum()

patches, texts,junk = plt.pie(values_to_plot, explode=explode,autopct='%1.1f%%', startangle=90, radius=1)

plt.setp(junk, **{'color':'white', 'weight':'bold'})
#plt.setp(junk, **{'color':'white', 'weight':'bold', 'fontsize':12.5})
# plt.legend(patches, labels, loc='right center', bbox_to_anchor=(1,.92),
#            fontsize=9)
plt.legend(patches, labels, loc='upper right', bbox_to_anchor=(1.35,.92),
           fontsize=9)
plt.text(0.1, 1.4, '16. YTD Working Hour according to supervisor', ha='center',color='#3238a8', fontsize=14, fontweight='bold')
#plt.title('Working Hour Percentage Per Company',fontsize='14',color='white', fontweight='bold',backgroundcolor='#6a9aba')
plt.tight_layout()

#plt.show()
plt.savefig('YTD_working_hour_according_to_supervisor.png')
plt.close()
print('16. YTD Working Hour according to supervisor generated')

