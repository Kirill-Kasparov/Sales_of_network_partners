import pandas as pd    # by Kirill Kasparov, 2022
import os
import scipy.stats as stats
import matplotlib.pyplot as plt
import numpy as np

top = 5    # количество Топ партнеров для графиков
data_import = os.getcwd().replace('\\', '/') + '/' + 'month_net.txt'

mode = 'txt'
if os.path.exists(data_import):    # проверяем наличие базы данных
    print('Загружаем базу данных...')
    df = pd.read_csv(data_import, sep='	', encoding='ANSI', header=10)  # загружаем базу
    print('База данных загружена, программа готова к работе.')
elif os.path.exists(os.getcwd().replace('\\', '/') + '/' + 'month_net.xls'):
    data_import = os.getcwd().replace('\\', '/') + '/' + 'month_net.xls'
    mode = 'xls'
    print('Загружаем базу данных...')
    #df = pd.read_csv(data_import, sep='	', encoding='ANSI', header=10)  # загружаем базу
    df = pd.read_excel(data_import, sheet_name='Лист Клиент', header=10, nrows=200000)
    print('База данных загружена, программа готова к работе.')
else:
    print('База данных не найдена.')
    print('Скачайте отчет ХД: "Оборот по отгрузке и КТН по сетевым партнерам" в формате Text и разделителем "Символ табуляции"')
    print('Загруженный файл разместите в директорию:', os.getcwd())
    input()



# Чистим базу
df['Код сети'] = df['Код сети'].astype('str')
df['Код сети'] = df['Код сети'].str.replace(' ', '')
df = df[df['Код сети'].str.isdigit()]    # удаляем все нечисловые значения, вроде "Итого"


if mode == 'txt':
    for i in df.columns[8:]:  # переводим все колонки в число
        df[i] = df[i].str.replace(',', '.')
        df[i] = df[i].str.replace(' ', '')
        df[i] = df[i].astype(float)


# Выводим прирост месяц к месяцу
df['Прирост'] = df['T'] - df['-12']
df = df.sort_values(by='Прирост')

# Выводим прирост по  валовой прибыли месяц к месяцу
df['ВП T'] = df['T'] - (df['T'] / df['Т(КТН)'])
df['ВП T'] = df['ВП T'].fillna(0.0) // 1

for i in range(-1, -13, -1):    # выводим КТН
    columns = 'ВП ' + str(i)
    col_ktn = str(i) + '(КТН)'
    df[columns] = df[str(i)] - (df[str(i)] / df[col_ktn])
    df[columns] = df[columns].fillna(0.0) // 1


# Получаем список уникальных сетей
df_unique = df['Код сети'] + ';' + df['Название сети']
df_unique = pd.DataFrame(df_unique.unique())
df_unique = df_unique[0].str.split(';', expand=True)
df_unique.columns = ['Код сети', 'Название сети']
df_unique.index= df_unique['Код сети']
#print(df_unique)



# Получаем суммы дубликадов по ключу код сети
df.drop(['Название сети', 'Регион', 'Логин ТРП', 'Код ТП', 'ФИО ТП', 'Код партнера', 'Название партнера'], axis=1, inplace=True)
df = df.groupby('Код сети').sum()
df.insert(loc=0, column='Название сети', value=df_unique['Название сети'])    # добавляем название сетей по индексу
df = df.sort_values(by='Прирост')    # сортируем продажи


# Удаляем столбцы с КТН
for i in df.columns:
     if 'КТН' in i:
          del df[i]
#print(df.to_string(max_colwidth=15))

# Выводим маркер партнеров без сделок
df['Нет продаж'] = df['T'] < 10000

# Выводим маркер угрозы прошлого года - major deal
df['Нестандарт'] = df['-11'] > ((df['T'] + df['-1'])/2 * 1.5)

# Добавляем строку суммы всех колонок
d = df.dtypes
df.loc['total'] = df.sum(numeric_only=True)
df.astype(d)

# Нормальное распределение
df['zscore_df'] = stats.zscore(df['Отгрузка'][:-1])    # модуль scipy.stats

# Сохраняем результат
for i in df.columns[1:-2]:    # переводим все колонки в число
     df[i] = df[i].astype(int)
#data_export = os.getcwd().replace('\\', '/') + '/' + 'month_net_export.csv'
#df.to_csv(data_export, sep=';', encoding='windows-1251', index=True, mode='w')

df.to_excel('Отгрузка_по_кодам.xlsx', index=False)


# Статистика
stat_column = df['T'][:-1]

print(stat_column.describe())    # выводим справку по таблице
count = len(stat_column)
print('Количество элементов:', count)
mean = stat_column.mean()    # Получаем среднее значение
print('Среднее значение:', int(mean))
variance = stat_column.var() ** 0.5 # получаем дисперсию
print('Дисперсия:', int(variance))
standard_deviation = variance ** 0.5    # извлекаем квадратный корень, получаем стандартное отклонение
print('Стандартное отклонение:', int(standard_deviation))
minimum = min(stat_column)    # получаем минимальное значение
print('Минимальное значение текущего месяца:', int(minimum))
maximum = max(stat_column)    # получаем максимальное значение
print('Максимальное значение текущего месяца:', int(maximum))
print('Оборот текущего месяца:', df['T'][-1])
print('Оборот в прошлом году:', df['-12'][-1])
print('Прирост оборота в руб.:', df['T'][-1] - df['-12'][-1], 'в процентах:', int((df['T'][-1] / df['-12'][-1] -1) * 100), '%')
print('ВП текущего месяца:', df['ВП T'][-1])
print('ВП в прошлом году:', df['ВП -12'][-1])
print('Прирост ВП в руб.:', df['ВП T'][-1] - df['ВП -12'][-1], 'в процентах:', int((df['ВП T'][-1] / df['ВП -12'][-1] -1) * 100), '%')


# matplotlib
# ----------------------------------------
# Строим график ТОП-5 и АнтиТОП-5
fig, ax = plt.subplots(2, 3, figsize=(15, 9))

# Анти ТОП - график bar
x_bar_anti = list(map(lambda x: x[:10] if len(x) > 10 else x, df['Название сети'][:top])) # Укарачиваем длину названий
y_bar_anti = df['Прирост'][:top] // 1000000

ax[0, 0].bar(x_bar_anti, y_bar_anti)    # сам график
ax[0, 0].set(ylabel='Отток к прошлому году в млн. руб.', title='Анти ТОП-5')
ax[0, 0].tick_params(axis='x', which='major', labelsize=8, rotation=20)    # поворачиваем ось Х, настраиваем шрифт
for i in range(len(x_bar_anti)):    # подписываем столбцы
     ax[0, 0].text(i, y_bar_anti[i], y_bar_anti[i])
print()
# Анти ТОП - график plot
for i in range(top):
     x_plot_anti = df.columns[13:0:-1]
     y_plot_anti = df.iloc[i][13:0:-1]
     ax[1, 0].plot(x_plot_anti, y_plot_anti, label=df['Название сети'][i])
     for n in range(len(y_plot_anti)):  # добавляем подписи элементам
        ax[1, 0].annotate(y_plot_anti[n] // 1000000, xy=(x_plot_anti[n], y_plot_anti[n]), fontsize=6)

ax[1, 0].set(xlabel='Месяц отгрузки', ylabel='Сумма отгрузки в млн. руб.')
ax[1, 0].legend(fontsize=6)
ax[1, 0].grid()

# ТОП - график bar
x_bar_top = list(map(lambda x: x[:10] if len(x) > 10 else x, df['Название сети'][top * -1 - 1:-1]))    # Укарачиваем длину названий
y_bar_top = df['Прирост'][top * -1 - 1:-1] // 1000000    # "top * -1 - 1" - это список с конца +1 строка, так как последняя строка df - ИТОГО

ax[0, 1].bar(x_bar_top, y_bar_top)    # сам график
ax[0, 1].set(ylabel='Прирост к прошлому году в млн. руб.', title='ТОП-5')    # даем имя
ax[0, 1].tick_params(axis='x', which='major', labelsize=8, rotation=20)    # поворачиваем ось Х, настраиваем шрифт
for i in range(len(x_bar_top)):    # подписываем столбцы
     ax[0, 1].text(i, y_bar_top[i], y_bar_top[i])

# ТОП - график plot
for i in range(top * -1 - 1, -1):
     x_plot_anti = df.columns[13:0:-1]
     y_plot_anti = df.iloc[i][13:0:-1]
     ax[1, 1].plot(x_plot_anti, y_plot_anti, label=df['Название сети'][i])
     for n in range(len(y_plot_anti)):  # добавляем подписи элементам
        ax[1, 1].annotate(y_plot_anti[n] // 1000000, xy=(x_plot_anti[n], y_plot_anti[n]), fontsize=6)

ax[1, 1].set(xlabel='Месяц отгрузки', ylabel='Сумма отгрузки в млн. руб.')
ax[1, 1].legend(fontsize=6)
ax[1, 1].grid()

# Нормальное распределение
ax[0, 2].scatter(df['zscore_df'][:-1], df['T'][:-1])
ax[0, 2].set(title='Отгрузки сетевых партнеров', ylabel='Текущий месяц в руб.')
ax[0, 2].grid()
for n in range(len(df['zscore_df'][:-1])):  # добавляем подписи элементам
    mark_x = df['zscore_df'][:-1]
    mark_y = df['T'][:-1]
    if mark_x[n] > 4 or mark_y[n] > 20000000:
        ax[0, 2].annotate(df['Название сети'][n], xy=(mark_x[n], mark_y[n]), fontsize=6)
# регрессия
b1, b0 = np.polyfit(df['zscore_df'][:-1], df['T'][:-1], 1) #  b0 - intercept, b1 - slope
ax[0, 2].plot(df['zscore_df'][:-1], b0 + b1*df['zscore_df'][:-1], color='red')

# График общих итогов в динамике по месяцам
total_x = df.columns[13:0:-1]
total_y = df.loc['total'][13:0:-1]
ax[1, 2].plot(total_x, total_y, color='orange', label='Отгрузка')
for i in range(len(total_y)):    # добавляем подписи элементам
    ax[1, 2].annotate(total_y[i] // 1000000, xy=(total_x[i], total_y[i]), fontsize=8)
ax[1, 2].set(xlabel='Месяц отгрузки', ylabel='Сумма отгрузки в млн. руб.')
ax[1, 2].grid()

# добавляем график динамики ВП
total_y2 = df.loc['total'][29:16:-1]
ax[1, 2].plot(total_x, total_y2, color='dodgerblue', label='ВП')
for i in range(len(total_y2)):    # добавляем подписи элементам
    ax[1, 2].annotate(total_y2[i] // 1000000, xy=(total_x[i], total_y2[i]), fontsize=8)
ax[1, 2].legend(fontsize=6)

plt.show()



