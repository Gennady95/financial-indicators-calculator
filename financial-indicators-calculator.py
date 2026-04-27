import pandas as pd, time, re, glob, getpass, platform, telebot, mysql, mysql.connector
from sqlalchemy import create_engine
from mysql.connector import Error
from datetime import datetime, timedelta
from tqdm import tqdm

engine = create_engine(os.getenv("DB_URL"))
bot = telebot.TeleBot(os.getenv("TELEGRAM_TOKEN"))
chat_id = os.getenv("CHAT_ID")
re_1 = r'[^0-9,.;/]'                                                                                                     # Регулярное выражение для отсева букв, пробелов
re_2 = r'[^0-9]'                                                                                                         # Регулярное выражение для отсева знаков

def SendTelegram(status):                                                                                                #Передача сообщения в telegram
	# Получение информации о компьютере
	UserName = getpass.getuser()                                                                                         # Имя пользователя (обычно оно User - не информативно)
	CompName = platform.node()                                                                                           # Имя компьютера
	chat_id = '5249664773'                                                                                               # ID моей телеги
	if status == "try": # Если связь с телегой установлена
		bot.send_message(chat_id, date+" пользователь "+UserName+" ("+CompName+") успешно воспользовался скриптом для выгрузки базовых показателей") # Отправка сообщения
	elif status == "except1": # Если нет подключения к SQL серверу
		bot.send_message(chat_id, "ERROR: "+date+" пользователь "+UserName+" ("+CompName+") неудачно запустил скрипт для выгрузки базовых показателей - не подключил VPN") # Отправка сообщения
def GetSQL():
	global dfSQLmag, dfSQLsto, dfSQLwestcall, dfSQLbitrix
	#Получение данных из SQL
	try:
		lightquery_1 = "SELECT НомерТелефона, Дата, ТТ, Сумма, Прибыль FROM sales_parts"                                 # SQL запрос в базу РОЗНИЦЫ
		lightquery_2 = "SELECT НомерТелефона, Дата, ТТ, Сумма, Прибыль FROM sales_sto"                                   # SQL запрос в базу СТО
		dfSQLmag = pd.read_sql(lightquery_1, engine)                                                                     # Чтение MySQL РОЗНИЦЫ, получение dataframe
		dfSQLsto = pd.read_sql(lightquery_2, engine)                                                                     # Чтение MySQL СТО, получение dataframe
		dfSQg = dfSQLmag.groupby([dfSQLmag.Дата.dt.year, dfSQLmag.Дата.dt.month])['Сумма'].sum()                         # Группировка выручки по годам и месяцам для РОЗНИЦЫ
		dfSQgP = dfSQLmag.groupby([dfSQLmag.Дата.dt.year, dfSQLmag.Дата.dt.month])['Прибыль'].sum()                      # Группировка прибыли по годам и месяцам для РОЗНИЦЫ
		dfSQgP1 = dfSQLmag.groupby([dfSQLmag.Дата.dt.year, dfSQLmag.Дата.dt.month])['Сумма'].count()                     # Группировка прибыли по годам и месяцам для РОЗНИЦЫ
		dfSQgP2 = dfSQLmag.groupby([dfSQLmag.Дата.dt.year, dfSQLmag.Дата.dt.month])['НомерТелефона'].nunique()           # Группировка прибыли по годам и месяцам для РОЗНИЦЫ
		dfSQSg = dfSQLsto.groupby([dfSQLsto.Дата.dt.year, dfSQLsto.Дата.dt.month])['Сумма'].sum()                        # Группировка выручки по годам и месяцам для СТО
		dfSQSgP = dfSQLsto.groupby([dfSQLsto.Дата.dt.year, dfSQLsto.Дата.dt.month])['Прибыль'].sum()                     # Группировка прибыли по годам и месяцам для СТО
		dfSQSgP1 = dfSQLsto.groupby([dfSQLsto.Дата.dt.year, dfSQLsto.Дата.dt.month])['Сумма'].count()                    # Группировка прибыли по годам и месяцам для СТО
		dfSQSgP2 = dfSQLsto.groupby([dfSQLsto.Дата.dt.year, dfSQLsto.Дата.dt.month])['НомерТелефона'].nunique()          # Группировка прибыли по годам и месяцам для СТО
		dfSQMIX1 = pd.concat([dfSQLmag, dfSQLsto], ignore_index=True)
		dfSQMIX2 = dfSQMIX1.groupby([dfSQMIX1.Дата.dt.year, dfSQMIX1.Дата.dt.month])['НомерТелефона'].nunique()
		# Запись вывода в excel файл
		with pd.ExcelWriter('Результат анализа ' + str(datetime.now().strftime('%d.%m_%H-%M-%S')) + ' время - ' + str(datetime.now().strftime('%H-%M-%S')) + '.xlsx') as writer:
			dfSQg.to_excel(writer, sheet_name='Выручка розницы')
			dfSQgP.to_excel(writer, sheet_name='Прибыль розницы')
			dfSQgP1.to_excel(writer, sheet_name='чеки розницы')
			dfSQgP2.to_excel(writer, sheet_name='клиенты розницы')
			dfSQSg.to_excel(writer, sheet_name='Выручка СТО')
			dfSQSgP.to_excel(writer, sheet_name='Прибыль СТО')
			dfSQSgP1.to_excel(writer, sheet_name='чеки СТО')
			dfSQSgP2.to_excel(writer, sheet_name='клиенты СТО')
			dfSQMIX2.to_excel(writer, sheet_name='микс по клиентам')
		SendTelegram("try")
		print("Скрипт закончил работу")
	except Exception as e:
		print(f'Произошла ошибка: {e}')
		print("Не могу подключится к SQL серверу. Проверьте подключение к VPN и перезапустите приложение")
		SendTelegram("except1")
		time.sleep(5); exit()
GetSQL()
