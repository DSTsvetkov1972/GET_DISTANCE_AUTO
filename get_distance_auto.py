from functions import *
from datetime import datetime
import pandas as pd
from win10toast import ToastNotifier
from colorama import init, Fore, Back, Style

class MyToastNotifier(ToastNotifier):
    def __init__(self):
        super().__init__()

    def on_destroy(self, hwnd, msg, wparam, lparam):
        super().on_destroy(hwnd, msg, wparam, lparam)
        return 0

#запускаем colorama
init()
print(Style.BRIGHT)

if input('Если хотите, чтобы программа при работе открывала браузер, введите "ДА": ') == "ДА":
    headless = False
else:
    headless = True

start = datetime.now()

messagebox_title = 'GET_DISTANCE_AUTO version 1'

df = pd.read_excel('Исходник.xlsx', header = None)#, nrows= 5)
routs_in_source_qty = len(df)
df_to_run = df.drop_duplicates()

distinct_routs_qty = len(df_to_run)
#print(distinct_routs_qty )

run_number = 0
run_number_next = 2
routs_qty = distinct_routs_qty
result_false_list_before_len = distinct_routs_qty
total_result_list = []

while True:

    #print(df_to_run)
    if len(df_to_run) == 0:
        break
    run_number += 1
    result_list = Run(headless, run_number, routs_qty, df_to_run)

    result_true_list = result_list[0]
    result_true_list_len = len(result_true_list)    
    #print(result_true_list)

    result_false_list = result_list[1]

    result_false_list_len = len(result_false_list)
    #print(result_false_list)

    total_result_list += result_true_list
 
    if result_true_list_len == routs_qty:
        break   
    elif result_false_list_before_len == result_false_list_len and run_number ==  run_number_next:
        total_result_list += result_false_list
        break
    else:
        run_number_next = run_number + 1
        result_false_list_before_len = result_false_list_len
    #result_true_list_len_before = result_true_list_len
    df_to_run = pd.DataFrame(result_false_list)
    routs_qty = len(df_to_run)
    


result_df = pd.DataFrame(total_result_list, columns =['Откуда','Куда','Сцепленные адреса','Результат','Спарслии для лекового','Для легкового в км','Спарсили для грузовика 12т','Для грузовика 12т в км'])    
result_df.to_excel('Результат %s.xlsx'%datetime.now().strftime("%Y-%m-%d %H-%M-%S"),index=False)   

finish = datetime.now()
print(Fore.CYAN + 
      'Стартовали: %s\nЗакончили: %s\nВыполнялось: %s\nПрогонов: %s'%(start,finish,finish - start,run_number) + 
      Fore.WHITE) 
if len(df_to_run) == 0:
    print(Fore.RED + 
          'Список адресов пуст!' + 
          Fore.WHITE) 
elif result_false_list_len == 0:
    print(Fore.GREEN + 
          'Обработка завершена!\nВ исходной таблице маршрутов: %s\nУникальных: %s'%(routs_in_source_qty,distinct_routs_qty) +
          Fore.WHITE)
else:
    print(Fore.YELLOW + 
          'Обработка завершена!\nВ исходной таблице маршрутов: %s\nУникальных: %s\n\nНе удалось обработать: %s '%(routs_in_source_qty,distinct_routs_qty,result_false_list_len) +
          Fore.WHITE)
"""
toaster = MyToastNotifier()

toaster.show_toast('GET_DISTANCE_AUTO',
                'Обработка завершена!')    
"""    
input()  
