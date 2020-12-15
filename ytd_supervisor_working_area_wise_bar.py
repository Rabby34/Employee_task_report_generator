import matplotlib
import matplotlib.pyplot as plt
import numpy as np
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


main_yakub_research_df = pd.DataFrame()
main_yakub_support_df = pd.DataFrame()
emplist_of_yakub = Yakub_sir_people_Name
for employ in range(len(emplist_of_yakub)):
    emp_name = (emplist_of_yakub[employ])
    #print(emp_name)
    yakub_ali_research_df = pd.read_sql_query("""select
ISNULL(sum(DATEDIFF(mm, start_time, end_time))+sum(DATEDIFF(hh, start_time, end_time)),0) as Total 
from tasks
where YEAR(task_date) = YEAR(CONVERT(varchar(10), getdate()-1,126))
and taskcategory_name='Research & Development'
and asign_name like ?""", connection,params={emp_name})

    main_yakub_research_df = main_yakub_research_df.append(yakub_ali_research_df)

    yakub_ali_support_df = pd.read_sql_query("""select
    ISNULL(sum(DATEDIFF(mm, start_time, end_time))+sum(DATEDIFF(hh, start_time, end_time)),0) as Total 
    from tasks
    where YEAR(task_date) = YEAR(CONVERT(varchar(10), getdate()-1,126))
    and taskcategory_name='Support & Maintenance'
    and asign_name like ?""", connection, params={emp_name})

    main_yakub_support_df = main_yakub_support_df.append(yakub_ali_support_df)


main_yakub_research_df.columns = range(main_yakub_research_df.shape[1])
main_yakub_support_df.columns = range(main_yakub_support_df.shape[1])

#print(main_yakub_research_df.values)
Total_of_yakub_research=sum(main_yakub_research_df.values)
#print(Total_of_yakub_research)

#print(main_yakub_support_df.values)
Total_of_yakub_support=sum(main_yakub_support_df.values)
#print(Total_of_yakub_support)

#sys.exit()



main_zubair_research_df = pd.DataFrame()
main_zubair_support_df = pd.DataFrame()
emplist_of_zubair = Zubair_sir_people_Name
for employ2 in range(len(emplist_of_zubair)):
    emp_name2 = (emplist_of_zubair[employ2])
    #print(emp_name2)
    zubair_research_df = pd.read_sql_query("""select
    ISNULL(sum(DATEDIFF(mm, start_time, end_time))+sum(DATEDIFF(hh, start_time, end_time)),0) as Total 
    from tasks
    where YEAR(task_date) = YEAR(CONVERT(varchar(10), getdate()-1,126))
    and taskcategory_name='Research & Development'
    and asign_name like ?""", connection, params={emp_name2})

    main_zubair_research_df = main_zubair_research_df.append(zubair_research_df)

    zubair_support_df = pd.read_sql_query("""select
        ISNULL(sum(DATEDIFF(mm, start_time, end_time))+sum(DATEDIFF(hh, start_time, end_time)),0) as Total 
        from tasks
        where YEAR(task_date) = YEAR(CONVERT(varchar(10), getdate()-1,126))
        and taskcategory_name='Support & Maintenance'
        and asign_name like ?""", connection, params={emp_name2})

    main_zubair_support_df = main_zubair_support_df.append(zubair_support_df)

main_zubair_research_df.columns = range(main_zubair_research_df.shape[1])
main_zubair_support_df.columns = range(main_zubair_support_df.shape[1])

#print(main_zubair_research_df.values)
Total_of_zubair_research = sum(main_zubair_research_df.values)
#print(Total_of_zubair_research)

#print(main_zubair_support_df.values)
Total_of_zubair_support = sum(main_zubair_support_df.values)
#print(Total_of_zubair_support)



main_tawhid_research_df = pd.DataFrame()
main_tawhid_support_df = pd.DataFrame()
emplist_of_tawhid = Tawhid_sir_people_Name
for employ3 in range(len(emplist_of_tawhid)):
    emp_name3 = (emplist_of_tawhid[employ3])
    tawhid_research_df = pd.read_sql_query("""select
        ISNULL(sum(DATEDIFF(mm, start_time, end_time))+sum(DATEDIFF(hh, start_time, end_time)),0) as Total 
        from tasks
        where YEAR(task_date) = YEAR(CONVERT(varchar(10), getdate()-1,126))
        and taskcategory_name='Research & Development'
        and asign_name like ?""", connection, params={emp_name3})

    main_tawhid_research_df = main_tawhid_research_df.append(tawhid_research_df)

    tawhid_support_df = pd.read_sql_query("""select
            ISNULL(sum(DATEDIFF(mm, start_time, end_time))+sum(DATEDIFF(hh, start_time, end_time)),0) as Total 
            from tasks
            where YEAR(task_date) = YEAR(CONVERT(varchar(10), getdate()-1,126))
            and taskcategory_name='Support & Maintenance'
            and asign_name like ?""", connection, params={emp_name3})

    main_tawhid_support_df = main_tawhid_support_df.append(tawhid_support_df)

main_tawhid_research_df.columns = range(main_tawhid_research_df.shape[1])
main_tawhid_support_df.columns = range(main_tawhid_support_df.shape[1])

#print(main_tawhid_research_df.values)
Total_of_tawhid_research = sum(main_tawhid_research_df.values)
#print(Total_of_tawhid_research)

#print(main_tawhid_support_df.values)
Total_of_tawhid_support = sum(main_tawhid_support_df.values)
#print(Total_of_tawhid_support)
#sys.exit()


values_for_research_all=[Total_of_yakub_research[0],Total_of_zubair_research[0],Total_of_tawhid_research[0]]
values_for_support_all=[Total_of_yakub_support[0],Total_of_zubair_support[0],Total_of_tawhid_support[0]]
#print(values_for_research_all)
#print(values_for_support_all)
new_array=[((Total_of_yakub_research[0]/(Total_of_yakub_research[0]+Total_of_yakub_support[0]))*100),
                     ((Total_of_zubair_research[0]/(Total_of_zubair_research[0]+Total_of_zubair_support[0]))*100),
                     ((Total_of_tawhid_research[0]/(Total_of_tawhid_research[0]+Total_of_tawhid_support[0]))*100),
           ((Total_of_yakub_support[0] / (Total_of_yakub_research[0] + Total_of_yakub_support[0])) * 100),
           ((Total_of_zubair_support[0] / (Total_of_zubair_research[0] + Total_of_zubair_support[0])) * 100),
           ((Total_of_tawhid_support[0] / (Total_of_tawhid_research[0] + Total_of_tawhid_support[0])) * 100)
           ]
print(new_array)

labels = ['Yakub', 'Zubair', 'Tawhid']

def plot_stacked_bar(data, series_labels, category_labels=None,
                     show_values=False, value_format="{}", y_label=None,
                     colors=None, grid=False, reverse=False):
    """Plots a stacked bar chart with the data and labels provided.

    Keyword arguments:
    data            -- 2-dimensional numpy array or nested list
                       containing data for each series in rows
    series_labels   -- list of series labels (these appear in
                       the legend)
    category_labels -- list of category labels (these appear
                       on the x-axis)
    show_values     -- If True then numeric value labels will
                       be shown on each bar
    value_format    -- Format string for numeric value labels
                       (default is "{}")
    y_label         -- Label for y-axis (str)
    colors          -- List of color labels
    grid            -- If True display grid
    reverse         -- If True reverse the order that the
                       series are displayed (left-to-right
                       or right-to-left)
    """

    ny = len(data[0])
    ind = list(range(ny))

    axes = []
    cum_size = np.zeros(ny)

    data = np.array(data)

    if reverse:
        data = np.flip(data, axis=1)
        category_labels = reversed(category_labels)

    for i, row_data in enumerate(data):
        color = colors[i] if colors is not None else None
        axes.append(plt.bar(ind, row_data, bottom=cum_size,
                            label=series_labels[i], color=color))
        cum_size += row_data

    if category_labels:
        plt.xticks(ind, category_labels)

    if y_label:
        plt.ylabel(y_label)

    plt.legend()

    if grid:
        plt.grid()

    if show_values:
        i = 0
        for axis in axes:
            for bar in axis:
                w, h = bar.get_width(), bar.get_height()
                plt.text(bar.get_x() + w/2, bar.get_y() + h/2,
                         value_format.format(int(h),float(new_array[i])), ha="center",
                         va="center")
                i = i + 1;

plt.figure(figsize=(6.4, 4.8))

series_labels = ['Research & \nDevelopment', 'Support & \nMaintenance']

data = [
    values_for_research_all,
    values_for_support_all
]

#category_labels = ['Cat A', 'Cat B', 'Cat C', 'Cat D']

plot_stacked_bar(
    data,
    series_labels,
    category_labels=labels,
    show_values=True,
    value_format='{0}H - \n{1:1.1f}% ',
    colors=['tab:blue', 'tab:orange']
)

plt.title('17. YTD Supervisors Working Area Wise Hour', ha='center',color='#3238a8', fontsize=14, fontweight='bold')

#plt.xlabel('Employee', fontsize=14)
plt.ylabel('Work Hour', fontsize=14)
plt.tight_layout()
plt.savefig('ytd_supervisor_working_area_wise_bar.png')
#plt.show()
plt.close()
print('17. YTD Supervisors Working Area Wise Hour')