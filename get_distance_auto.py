import os
from functions import *
from logo import logo
from datetime import datetime
import pandas as pd
from colorama import init, Fore, Back, Style
from tkinter import filedialog as fd
from tkinter import messagebox as mb

title = f"GET_DISTANCE_AUTO_(v.{datetime.now().strftime('%Y-%m-%d')})"

init()
print(Style.BRIGHT)
print(Fore.YELLOW)
print(logo)

while True:
    #запускаем colorama

    print(Fore.MAGENTA)
    print('ВЫБЕРИТЕ ЭКСЕЛЕВСКИЙ ФАЙЛ СО СПИСКОМ АДРЕСОВ!')

    print(Fore.BLUE)
    print('Файл должен содержать единственный лист на котором\n'
        'в первой колонке адреса "откуда", во второй - "куда".\n'
        'Колонки должны быть без заголовков, т.е. первая строка - это уже маршрут')

    while True:

        filename = fd.askopenfilename()

        if filename:
            print(Fore.MAGENTA)
            print("Вы выбрали: ", Fore.GREEN, os.path.basename(filename))
            print(Fore.MAGENTA)
            print("Из директории: ", Fore.GREEN, os.path.dirname(filename))
            try:
                df = pd.read_excel(filename, header = None)
                if not df.empty:
                    print(Fore.MAGENTA)
                    print('Вот что содержится в файле:', Fore.RESET)
                    print(df)
                    print(Fore.MAGENTA)
                    print('Продолжить?', Fore.RESET)           
                    if mb.askyesno(title, "Продолжить?"):
                        break
                else:
                    print(Fore.RED)
                    print(f"Лист выбранного Вами файла\n{filename}\nпуст!", Fore.RESET)   
                    mb.showerror(title,
                                f"Лист выбранного Вами файла\n{filename}\nпуст!")     
            except ValueError:
                print(Fore.RED)
                print(f"Не удаются прочитать файл {filename}!", Fore.RESET)   
                mb.showerror(title,
                             f"Не удаются прочитать файл {filename}!")         
        else:
            print(Fore.RED)
            print("Файл не был выбран, повторите попытку!", Fore.RESET)
            mb.showerror(title,
                         "Файл не был выбран, повторите попытку!")    

    print(Fore.MAGENTA)
    print('Нужно ли открывать браузер при работе программы?'.upper(), Fore.RESET)
    headless = mb.askyesno(title, 'Нужно ли открывать браузер при работе программы?')

    start = datetime.now()
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
        


    result_df = pd.DataFrame(total_result_list, columns=['Откуда','Куда','Сцепленные адреса','Результат','Спарслии для лекового','Для легкового в км','Спарсили для грузовика 12т','Для грузовика 12т в км'])    
    resfilename=os.path.join(os.path.dirname(filename), f'{title}_{datetime.now().strftime("%Y-%m-%d %H-%M-%S")}.xlsx')
    result_df.to_excel(resfilename, index=False)   

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
            Fore.RESET)
    else:
        print(Fore.YELLOW + 
            'Обработка завершена!\nВ исходной таблице маршрутов: %s\nУникальных: %s\n\nНе удалось обработать: %s '%(routs_in_source_qty,distinct_routs_qty,result_false_list_len) +
            Fore.WHITE)
        
    print(Fore.MAGENTA)
    print(f"Результат сохранен в файле {resfilename}", Fore.RESET)

    os.startfile(resfilename)     

    if not mb.askyesno(title,'Обработка завершена! Запустить еще раз?'):
        break

 
