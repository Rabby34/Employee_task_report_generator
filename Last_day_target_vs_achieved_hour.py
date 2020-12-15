import matplotlib.pyplot as plt
import pandas as pd
import os
import pyodbc as db
import sys

connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=;'                     # Provide server
                        'DATABASE=;'                   # provide database 
                        'UID=;PWD=')                   # provide username and password

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
        where task_date  =  CONVERT(varchar(10), getdate()-1,126)
        group by workfor_name) as Btask
        on ltrim(Atask.workfor_name) = ltrim(Btask.workfor_name)
        group by Atask.workfor_name,Atask.sqnc
        order by sqnc asc""", connection)

import datetime

today = datetime.date.today()
#print('Today = ', today)
month_start = datetime.date(today.year, today.month, 1)
#print('Months Start Day = ', month_start)
single_day = datetime.timedelta(days=1)
#print(single_day)

if (today.day - month_start.day == 0):
    import calendar

    total_days_in_month = calendar.monthrange(today.year, today.month - 1)[1]
    #print("Total number of days: ", total_days_in_month)

    total_friday = 0
    total_satur_day = 0

    from datetime import date, datetime, timedelta

    given_date = datetime.today().date()

    #print(given_date)

    end_date = str(given_date.day) + '/' + str(given_date.month) + '/' + str(given_date.year)
    #print(end_date)

    start_date = '1' + '/' + str(given_date.month) + '/' + str(given_date.year)
    #print(start_date)

    end_date2 = str(total_days_in_month) + '/' + str(given_date.month - 1) + '/' + str(given_date.year)
    #print(end_date2)

    start_date2 = '1' + '/' + str((given_date.month) - 1) + '/' + str(given_date.year)
    #print(start_date2)

    import datetime


    def weekday_friday(start, end):
        start_date = datetime.datetime.strptime(start, '%d/%m/%Y')
        end_date = datetime.datetime.strptime(end, '%d/%m/%Y')
        week = {}
        for i in range((end_date - start_date).days):
            day = calendar.day_name[(start_date + datetime.timedelta(days=i + 1)).weekday()]
            week[day] = week[day] + 1 if day in week else 1
        return week['Friday']


    total_friday = weekday_friday(start_date2, end_date2)


    def weekday_saturday(start, end):
        start_date = datetime.datetime.strptime(start, '%d/%m/%Y')
        end_date = datetime.datetime.strptime(end, '%d/%m/%Y')
        week = {}
        for i in range((end_date - start_date).days):
            day = calendar.day_name[(start_date + datetime.timedelta(days=i + 1)).weekday()]
            week[day] = week[day] + 1 if day in week else 1
        return week['Saturday']


    total_satur_day = weekday_saturday(start_date2, end_date2)

    #print('Total Friday:', total_friday)
    #print('Total Saturday:', total_satur_day)

    personel_df = pd.read_sql_query("""select count(status) as amount from users
    where status=1 and
    name != 'Mohammad Shakhawat Hossain Biswas'
    """, connection)

    total_personel = int(personel_df['amount'])
    #total_personel = 45

    average_work_hour_except_friday_saturday = ((total_days_in_month - (
            total_friday + total_satur_day)) * 8 * total_personel) / (
                                                       total_days_in_month - (total_friday + total_satur_day))
    #print('work average except friday and saturday : ', average_work_hour_except_friday_saturday)

    average_work_hour_in_saturday = (total_satur_day * 4 * total_personel) / total_satur_day
    #print('work average in saturday : ', average_work_hour_in_saturday)

    target_percentages = [22, 16, 7, 7, 17, 1, 5, 10, 7, 3, 2, 3]

    yesterday_date = given_date - timedelta(days=1)
    #print(given_date)
    #print(yesterday_date)
    yesterday_name = yesterday_date.strftime("%A")
    #print(yesterday_name)

    # yesterday_name='Saturday'

    if (yesterday_name == 'Saturday'):
        target_hours = average_work_hour_in_saturday
        the_array = []
        for percents in target_percentages:
            #print(percents)
            value = (target_hours * percents) / 100
            the_array.append(value)
            value = 0
        #print(the_array)

    else:
        target_hours = average_work_hour_except_friday_saturday
        the_array = []
        for percents in target_percentages:
            #print(percents)
            value = (target_hours * percents) / 100
            the_array.append(value)
            value = 0
        #print(the_array)

    Company_names = last_day_chart_df['company'].values.tolist()
    hours = last_day_chart_df['work_hour'].values.tolist()
    #print(Company_names)
    #print(hours)

    for i in range(0, len(hours)):
        hours[i] = int(hours[i])
    #print('achived hour into int :', hours)

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

    #plt.title('2. Last Day Target vs Achieved Hour', fontsize=14)
    plt.title('3. Last Day Target vs Achieved Hour', ha='center', color='#3238a8', fontsize=14, fontweight='bold')
    #plt.xlabel('Company', fontsize=14)
    plt.ylabel('Work Hour', fontsize=14)
    plt.tight_layout()
    # plt.grid(True)
    #plt.show()
    plt.legend(['Target Hour', 'Achieved Hour'])
    plt.savefig('last_day_target_vs_achieved_hour.png')
    plt.close()
    print('3. Last Day Target vs Achieved Hour generated')

else:
    import calendar

    total_days_in_month = calendar.monthrange(today.year, today.month)[1]
    #print("Total number of days: ", total_days_in_month)

    # # # -------------------------------------------
    import datetime
    from dateutil.relativedelta import relativedelta
    import re
    from datetime import date
    from pandas.tseries import offsets

    # # --------------------------------------------------
    # # Current Date and Months Start Date
    # # --------------------------------------------------
    todayDate = datetime.date.today()
    resultDate = todayDate.replace(day=1)
    if ((todayDate - resultDate).days > 25):
        resultDate = resultDate + relativedelta(months=0)
    month_start_date = int(re.sub("\-", '', str(resultDate)))
    current_date = int(re.sub("\-", '', str(todayDate)))
    print('Month Start Date = ', month_start_date)
    print('Current Date = ', current_date)

    # # --------------------------------------------------
    # # Months last day
    # # --------------------------------------------------
    import datetime
    import re

    date = datetime.date.today()


    def last_day_of_month(any_day):
        next_month = any_day.replace(day=28) + datetime.timedelta(days=4)  # this will never fail
        return next_month - datetime.timedelta(days=next_month.day)


    months_last_day = last_day_of_month(datetime.date(date.year, date.month, 1))
    months_last_day = int(re.sub("\-", '', str(months_last_day)))
    #print("Month's Last Day = ", months_last_day)
    # # ---------------------------------------------------------------
    # # Current Months Number of Off Day, Half day and Full Working Day
    # # ---------------------------------------------------------------
    # Read the dataset
    data = pd.read_excel('./working_date.xlsx')
    # Sorting data set
    total_offday_in_month = data[
        (data.maindate.between(month_start_date, months_last_day, inclusive=True))
        & (data.working_status == 0)
        ]

    Public_off_days_in_month = total_offday_in_month.maindate.count()

    #print('Total Off Day In this Month = ', Public_off_days_in_month)

    # # Total Saturday in a month
    total_halfday_in_month = data[
        (data.maindate.between(month_start_date, months_last_day, inclusive=True))
        & (data.working_status == 1)
        ]

    weekly_saturdays_in_month = total_halfday_in_month.maindate.count()

    #print('Total Half Day In this Month = ', weekly_saturdays_in_month)
    # # Total Full Day in a month

    total_Fullday_in_month = data[
        (data.maindate.between(month_start_date, months_last_day, inclusive=True))
        & (data.working_status == 2)
        ]

    full_days_in_month = total_Fullday_in_month.maindate.count()

    #print('Total Full Day In this Month = ', full_days_in_month)

    personel_df = pd.read_sql_query("""select count(status) as amount from users
    where status=1 and
    name != 'Mohammad Shakhawat Hossain Biswas'
    """, connection)

    total_personel = int(personel_df['amount'])
    #total_personel = 45

    average_work_hour_except_friday_saturday = ((full_days_in_month) * 8 * total_personel) / (full_days_in_month)
    #print('work average except offday and saturday : ', average_work_hour_except_friday_saturday)

    average_work_hour_in_saturday = (weekly_saturdays_in_month * 4 * total_personel) / weekly_saturdays_in_month
    #print('work average in saturday : ', average_work_hour_in_saturday)

    target_percentages = [22, 16, 7, 7, 17, 1, 5, 10, 7, 3, 2, 3]

    from datetime import date, datetime, timedelta

    given_date = datetime.today().date()
    yesterday_date = given_date - timedelta(days=1)
    #print(given_date)
    #print(yesterday_date)
    yesterday_name = yesterday_date.strftime("%A")
    #print(yesterday_name)

    # yesterday_name='Saturday'

    if (yesterday_name == 'Saturday'):
        target_hours = average_work_hour_in_saturday
        the_array = []
        for percents in target_percentages:
            #print(percents)
            value = (target_hours * percents) / 100
            the_array.append(value)
            value = 0
        #print(the_array)

    else:
        target_hours = average_work_hour_except_friday_saturday
        the_array = []
        for percents in target_percentages:
            #print(percents)
            value = (target_hours * percents) / 100
            the_array.append(value)
            value = 0
        #print(the_array)

    Company_names = last_day_chart_df['company'].values.tolist()
    hours = last_day_chart_df['work_hour'].values.tolist()
    #print(Company_names)
    #print(hours)

    for i in range(0, len(hours)):
        hours[i] = int(hours[i])
    #print('achived hour into int :', hours)

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

    plt.title('3. Last Day Target vs Achieved Hour', ha='center',color='#3238a8', fontsize=14, fontweight='bold')
    #plt.xlabel('Company', fontsize=14)
    plt.ylabel('Work Hour', fontsize=14)
    plt.tight_layout()
    # plt.grid(True)

    plt.legend(['Target Hour', 'Achieved Hour'])
    #plt.show()
    plt.savefig('last_day_target_vs_achieved_hour.png')
    plt.close()
    print('3. Last Day Target vs Achieved Hour generated')
