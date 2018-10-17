import calendar
import datetime

# today = datetime.datetime.now() #текущий день для реальной работы
today = datetime.date(2018, 10, 9)  # текущий день для тестирования
print("Current day: ", type(today))
weekday = (today.weekday())  # порядковый номер дня недели 0-ПН, 6-ВС


# Отображение номера недели текущего года
def week_number():
    weekNumber = datetime.date.today().isocalendar()[1]
    # print('Week number:', weekNumber, "of ", today.year, "Year")
    return weekNumber


# Отображение общего количества дней в месяце
def total_month_days():
    m = today.month
    if m == 1 or m == 3 or m == 5 or m == 7 or m == 8 or m == 10 or m == 12:
        end_day = 31
        print("Days in month: ", end_day)
        return end_day
    elif m == 4 or m == 6 or m == 9 or m == 11:
        end_day = 30
        print("Days in month: ", end_day)
        return end_day
    elif (today.year == 2020 or today.year == 2024 or today.year == 2028) & (m == 2):
        end_day = 29
        print("Days in month: ", end_day)
        return end_day
    elif m == 2:
        end_day = 28
        print("Days in month: ", end_day)
        return end_day


# переменные начала и конца месяца
start = datetime.date(today.year, today.month, 1)
end = datetime.date(today.year, today.month, total_month_days())


# Подсчет количества рабочих дней в текущем месяце (FUNCTION)
def count_total_busyday():
    cal = calendar.Calendar()
    working_days = len([x for x in cal.itermonthdays2(today.year, today.month) if x[0] != 0 and x[1] < 5])
    return working_days


print("Total working days this month: " + str(count_total_busyday()))

if today.day >= 26:
    work_hours = count_total_busyday() * 8
    print("==================== 1 if ===================\n")
    # print("Expected number of reported hours is", work_hours, "hours")

elif weekday >= 4:
    # Fr,Sa,Su
    daydiff = today.weekday() - start.weekday()
    days = (((today - start).days - daydiff) / 7 * 5 + min(daydiff, 5) - (
            max(today.weekday() - 4, 0) % 5)) + 1
    if weekday == 6:
        days += 1
    print("==================== 2 if ===================\n", "today -", today.day, "| weekday:", weekday,
          "| working_days: ", count_total_busyday(), "|")
    work_hours = days * 8

else:
    # Mo,Tu,We,Th
    print("==================== 3 if ===================\n", "today -", today.day, "| weekday: ", weekday,
          "| working_days: ", count_total_busyday(), "|")
    daydiff = today.weekday() - start.weekday()
    days = (((today - start).days - daydiff) / 7 * 5 + min(daydiff, 5) - (max(today.weekday() - 4, 0) % 5)) + 1
    work_hours = (days - weekday - 1) * 8
    print("---  end day  ----", end)
    print("--- start day ----", start)
    print("--- today.day ----", today.day)
print("Expected number of reported hours is", work_hours, "hours")
