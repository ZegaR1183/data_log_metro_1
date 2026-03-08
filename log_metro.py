# Импорт библиотек
import pandas as pd

# Объявление переменных для чтения и записи файла.
file_in = "result_output"
file_out = "clear_log_new.txt"
file_in = "clear_log_new.txt"
file_out = "dict_all.txt"

# Настройки вывода для DataFrame.
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 300)

# Словари для хранения данных.
list_all = []
list_dict = []
keys_mx = ['name', 'type', 'temp PEM_0', 'temp PEM_1', 'temp RE_0','temp RE_1', 's_fan_1', 's_fan_2', 's_fan_3', 's_fan_4', 's_fan_5']
keys_acx_4000 = ['name', 'type', 'temp PEM_0', 'temp PEM_1', 'temp RE_0', 's_fan_1', 's_fan_2']
keys_acx_2100 = ['name', 'type', 'temp RE_0']