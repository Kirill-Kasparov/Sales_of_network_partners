import pandas as pd    # by Kirill Kasparov, 2022
import os
import datetime
import calendar
import scipy.stats as stats
import dash
from dash import dcc as dcc
from dash import html as html
from dash import Dash, dash_table
import plotly.express as px
import plotly.graph_objs as go


def workind_days():
    # Получаем дату из файла
    creation_time = pd.read_excel(data_import, sheet_name='Лист Клиент', header=1, nrows=1)
    creation_time = creation_time.columns[0].replace('Отчётная дата: ', '').split()
    creation_time = creation_time[0].split('.')
    ddf = datetime.date(int(creation_time[2]), int(creation_time[1]), int(creation_time[0]))

    # Получаем дату по созданию файла
    # creation_time = os.stat(data_import).st_ctime
    # ddf = datetime.datetime.fromtimestamp(creation_time)
    # date = ddf.strftime('%d.%m.%Y')

    # Список праздничных дней
    holidays = pd.read_excel('holidays.xlsx')
    # Конвертируем дату из Excel формата в datetime формат
    for col in holidays.columns:
        holidays[col] = pd.to_datetime(holidays[col]).dt.date

    # Начальная и конечная дата текущего месяца
    start_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')), 1)
    now_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')), int(ddf.strftime('%d')))
    end_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')),
                             calendar.monthrange(now_date.year, now_date.month)[1])

    # Вычисляем количество рабочих дней
    now_working_days = 0
    end_working_days = 0

    current_date = start_date
    while current_date <= end_date:
        # Если текущий день не является выходным или праздничным, увеличиваем счетчик рабочих дней
        if current_date in list(holidays['working_days']):  # рабочие дни исключения
            now_working_days += 1
            end_working_days += 1
        elif current_date.weekday() < 5 and current_date not in list(holidays[int(ddf.strftime('%Y'))]):
            if current_date <= now_date:
                now_working_days += 1
            end_working_days += 1
        current_date += datetime.timedelta(days=1)

    # Получаем данные прошлого года
    start_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')), 1)
    now_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')), int(ddf.strftime('%d')))
    end_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')),
                               calendar.monthrange(now_date_2.year, now_date_2.month)[1])
    now_working_days_2 = 0
    end_working_days_2 = 0

    current_date = start_date_2
    while current_date <= end_date_2:
        # Если текущий день не является выходным или праздничным, увеличиваем счетчик рабочих дней
        if current_date in list(holidays['working_days']):  # рабочие дни исключения
            now_working_days_2 += 1
            end_working_days_2 += 1
        elif current_date.weekday() < 5 and current_date not in list(holidays[int(ddf.strftime('%Y')) - 1]):
            if current_date < now_date_2:
                now_working_days_2 += 1
            end_working_days_2 += 1
        current_date += datetime.timedelta(days=1)
    date_lst = [start_date, start_date_2, now_date, now_date_2, end_date, end_date_2, now_working_days, now_working_days_2, end_working_days, end_working_days_2]
    return date_lst

top = 5    # количество Топ партнеров для графиков
data_import = os.getcwd().replace('\\', '/') + '/' + 'month_net.txt'

mode = 'txt'
if os.path.exists(data_import):    # проверяем наличие базы данных
    df = pd.read_csv(data_import, sep='	', encoding='ANSI', header=10)  # загружаем базу
elif os.path.exists(os.getcwd().replace('\\', '/') + '/' + 'month_net.xls'):
    data_import = os.getcwd().replace('\\', '/') + '/' + 'month_net.xls'
    mode = 'xls'
    df = pd.read_excel(data_import, sheet_name='Лист Клиент', header=10, nrows=200000)
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
        df[i] = df[i].fillna('0')
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

df.to_excel('Отгрузка_по_кодам.xlsx', index=False)

# -----------------------Даты---------------------

start_date, start_date_2, now_date, now_date_2, end_date, end_date_2, now_working_days, now_working_days_2, end_working_days, end_working_days_2 = workind_days()
# print('Начало месяца:', start_date, 'в прошлом году', start_date_2)
# print('Отчетная дата:', now_date, 'в прошлом году', now_date_2)
# print('Конец месяца', end_date, 'в прошлом году', end_date_2)
# print('Рабочих дней сейчас', now_working_days, 'в прошлом году', now_working_days_2)
# print('Всего рабочих дней', end_working_days, 'в прошлом году', end_working_days_2)


# -----------------------DASH---------------------

# Создаем текст
text = html.P(children='Оборот отгрузок текущего месяца: ' + str(df['T'][-1] // 1000000) + ' млн. руб', className="header-description")
text1 = html.P(children='Оборот отгрузок в прошлом году: ' + str(df['-12'][-1] // 1000000) + ' млн. руб', className="header-description")
text2 = html.P(children='Прогноз оборота: ' + str(int((df['T'][-1]) / now_working_days * end_working_days // 1000000)) + ' млн. руб', className="header-description")
text3 = html.P(children='Прогноз прироста по обороту (ср-дн): ' + str(int((((df['T'][-1]) / now_working_days) / (df['-12'][-1] / end_working_days_2) - 1) * 100)) + ' %', className="header-description")

text4 = html.P(children='Оборот ВП текущего месяца: ' + str(df['ВП T'][-1] // 1000000) + ' млн. руб', className="header-description")
text5 = html.P(children='Оборот ВП в прошлом году: ' + str(df['ВП -12'][-1] // 1000000) + ' млн. руб', className="header-description")
text6 = html.P(children='Прогноз ВП: ' + str(int((df['ВП T'][-1]) / now_working_days * end_working_days // 1000000)) + ' млн. руб', className="header-description")
text7 = html.P(children='Прогноз прироста по обороту (ср-дн): ' + str(int((((df['ВП T'][-1]) / now_working_days) / (df['ВП -12'][-1] / end_working_days_2) - 1) * 100)) + ' %', className="header-description")

text8 = html.P(children='Прошло рабочих дней: ' + str(now_working_days) + " из " + str(end_working_days) + ". Всего раб. дней в прошлом году: " + str(end_working_days_2), className="header-description")

# print('Рабочих дней сейчас', now_working_days, 'в прошлом году', now_working_days_2)
# print('Всего рабочих дней', end_working_days, 'в прошлом году', end_working_days_2)

# Создаем графики
# получаем месяца, вместо столбцов Т, -1, -2...
col_for_total_result = [now_date.strftime('%d.%m.%Y')]    # вместо list(df.columns[13:0:-1])
for i in range(1, 13):
    prev_month = now_date - datetime.timedelta(days=28 * i)
    while prev_month.strftime('%m.%Y') in col_for_total_result or prev_month.strftime('%m.%Y') == now_date.strftime('%m.%Y'):   # поправка на дни
        prev_month = prev_month - datetime.timedelta(days=1)
    col_for_total_result.append(prev_month.strftime('%m.%Y'))

# формируем таблицу
total_result = {'Месяц': col_for_total_result[::-1], 'Отгрузка, млн. руб.': list(df.loc['total'][13:0:-1] // 1000000), 'Прибыль, млн. руб.': list(df.loc['total'][29:16:-1] // 1000000)}
# строим график
fig1 = px.line(total_result, x='Месяц', y='Отгрузка, млн. руб.', title='Динамика оборота по месяцам', text='Отгрузка, млн. руб.')
fig1.update_layout(font=dict(size=18))    # увеличиваем шрифт
fig2 = px.bar(total_result, x='Месяц', y='Прибыль, млн. руб.', title='Прибыль по месяцам', text='Прибыль, млн. руб.')
fig2.update_layout(font=dict(size=18))    # увеличиваем шрифт


x_bar_top = list(map(lambda x: x[:30] if len(x) > 30 else x, df['Название сети'][top * -1 - 1:-1]))    # Укарачиваем длину названий
y_bar_top = df['Прирост'][top * -1 - 1:-1] // 1000000    # "top * -1 - 1" - это список с конца +1 строка, так как последняя строка df - ИТОГО
fig3 = px.bar(total_result, x=x_bar_top, y=y_bar_top, title='ТОП сетей текущего месяца (в млн. руб.)', text=y_bar_top)

x_bar_anti = list(map(lambda x: x[:30] if len(x) > 30 else x, df['Название сети'][:top])) # Укарачиваем длину названий
y_bar_anti = df['Прирост'][:top] // 1000000
fig4 = px.bar(total_result, x=x_bar_anti, y=y_bar_anti, title='Анти ТОП сетей текущего месяца (в млн. руб.)', text=y_bar_anti)

# Создаем таблицы
df.drop(['Отгрузка по личным заказам', 'Отгрузка'], axis=1, inplace=True)

# Таблица риска
df_risk = df[df['Нестандарт'] == True]
df_risk = df_risk.iloc[:, 0:14]
df_risk['Прогноз оттока'] = (df_risk['-1'] + df_risk['-2']) / 2 - df_risk['-11']
df_risk = df_risk.sort_values(by='Прогноз оттока').head(10)

for i in range(1, len(df_risk.columns)):    # округляем до млн. руб.
    df_risk.iloc[:, i] = df_risk.iloc[:, i] // 1000000

table = html.Table(
    [html.Tr([html.Th(col) for col in df_risk.columns])] +
    [html.Tr([html.Td(df_risk.iloc[i][col]) for col in df_risk.columns]) for i in range(len(df_risk))], className="table")

# Таблица АнтиТоп
df_anti = df.iloc[:, 0:14].head(5)
df_anti['Отток'] = df_anti['T'] - df_anti['-12']
for i in range(1, len(df_anti.columns)):    # округляем до млн. руб.
    df_anti.iloc[:, i] = df_anti.iloc[:, i] // 1000000
table_anti = html.Table(
    [html.Tr([html.Th(col) for col in df_anti.columns])] +
    [html.Tr([html.Td(df_anti.iloc[i][col]) for col in df_anti.columns]) for i in range(len(df_anti))], className="table")

# Таблица Топ
df = df.sort_values(by='Прирост', ascending=False)
df_top = df.iloc[:, 0:14].head(5)
df_top['Прирост'] = df_top['T'] - df_top['-12']

for i in range(1, len(df_top.columns)):    # округляем до млн. руб.
    df_top.iloc[:, i] = df_top.iloc[:, i] // 1000000
table_top = html.Table(
    [html.Tr([html.Th(col) for col in df_top.columns])] +
    [html.Tr([html.Td(df_top.iloc[i][col]) for col in df_top.columns]) for i in range(len(df_top))], className="table")

# Таблица привлеченные сети
df_stars = df[((df['T'] + df['-1'] + df['-2']) > ((df['-12'] + df['-11'] + df['-10']) * 2)) & ((df['T'] + df['-1'] + df['-2']) > 10000000)]
df_stars = df_stars.iloc[:, 0:14]
for i in range(1, len(df_stars.columns)):    # округляем до млн. руб.
    df_stars.iloc[:, i] = df_stars.iloc[:, i] // 1000000
table_stars = html.Table(
    [html.Tr([html.Th(col) for col in df_stars.columns])] +
    [html.Tr([html.Td(df_stars.iloc[i][col]) for col in df_stars.columns]) for i in range(len(df_stars))], className="table")


# Таблица Потеря продаж 3 мес
df_lost = df[(df['T'] < 1000) & (df['-1'] < 1000) & (df['-2'] < 1000)]
df_lost = df_lost.iloc[:, 0:14]
table_lost = html.Table(
    [html.Tr([html.Th(col) for col in df_lost.columns])] +
    [html.Tr([html.Td(df_lost.iloc[i][col]) for col in df_lost.columns]) for i in range(len(df_lost))], className="table")



# Создаем дашборд
app = dash.Dash(__name__)

app.layout = html.Div(style={'background-image': 'url("/assets/bg.jpg")',
           'background-repeat': 'repeat', 'background-position': 'right top',
           'background-size': '1920px 1152px'}, children=[
    html.H1(children= 'Аналитика продаж сетевых партнеров на ' + str(now_date), className="header-title",),
    html.Br(),
    text, text1, text2, text3, text8,
    dcc.Graph(figure=fig1, className="wrapper",),
    html.Br(),
    text4, text5, text6, text7,
    dcc.Graph(figure=fig2, className="wrapper",),
    html.Br(),
    html.H1(children='Рейтинг по приросту оборота', className="header-title", ),
    html.Br(),
    dcc.Graph(figure=fig3, className="wrapper",),
    table_top,
    html.Br(),
    dcc.Graph(figure=fig4, className="wrapper", ),
    table_anti,
    html.Br(),
    html.H1(children='Риски следующего месяца (оборот в млн. руб.)', className="header-title", ),
    html.Br(),
    table,
    html.Br(),
    html.H1(children='Привлеченные сети (оборот в млн. руб.)', className="header-title", ),
    html.Br(),
    table_stars,
    html.Br(),
    html.H1(children='Потерянные сети', className="header-title", ),
    html.Br(),
    table_lost,
    html.Br(),
    html.H1(children='Продажи в разрезе сетей по месяцам', className="header-title", ),
    html.Br(),
    dash_table.DataTable(
        id='datatable-interactivity',
        columns=[
            {"name": i, "id": i, "deletable": True, "selectable": True} for i in df.columns[:15]
        ],
        data=df.to_dict('records'),
        editable=True,
        filter_action="native",
        sort_action="native",
        sort_mode="multi",
        column_selectable="single",
        row_selectable="multi",
        row_deletable=True,
        selected_columns=[],
        selected_rows=[],
        page_action="native",
        page_current= 0,
        page_size= 20,
    ),
    html.Div(id='datatable-interactivity-container'),
    html.Br(),
    html.Br(),

])

if __name__ == '__main__':
    app.run_server(debug=True)
