import os
import markdown
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options 
from colorama import Fore

def Get_distance(address_from, address_to, sleep_time=0.5, attempt_to_load_time=10, headless=True):
    if str(address_from) == 'nan' or str(address_from) == '':
        return [address_from, 
                address_to, 
                '' + address_to, 
                False,
                'Нет адреса окуда!',
                '',
                '',
                '']
    elif str(address_to) == 'nan' or str(address_to) == '':
        return [address_from, 
                address_to, 
                address_from + '', 
                False,
                'Нет адреса куда!',
                '',
                '',
                '']

    try:
        if not headless:
            options = Options()
            # options.add_argument("--headless=new")
            options.add_argument("--headless")            
            options.add_argument("--window-position=-1920,0")
            driver = webdriver.Chrome(options=options)
            driver.implicitly_wait(1)
        else:
            options = Options()
            options.add_argument("--window-position=-1920,0")
            driver = webdriver.Chrome(options=options)

        
        #driver.set_window_size(1920, 1080)   
        driver.get("https://yandex.ru/maps/")
        try:
            #time.sleep(sleep_time)
            elem = driver.find_element(By.CLASS_NAME,'CheckboxCaptcha-Label')
            print('Капча есть')
            elem.click()    
        except Exception as e:
            print('Капчи нет')         
    except:
        return [address_from, address_to, False, 'Не удалось войти в яндекс-карты!']
#           driver.set_window_size(1920, 1080)      

    try:
        # Открываем поля ввода точек маршрута
        elem = WebDriverWait(driver, attempt_to_load_time).until(EC.element_to_be_clickable((By.CLASS_NAME,"route-control__inner")))
        #time.sleep(t)
        elem.click()
        print("Кликнули на маршруты")
        elems = []

        start =datetime.now()
        while len(elems)<3 and (datetime.now()-start).total_seconds() < attempt_to_load_time:
            elems = driver.find_elements(By.CLASS_NAME,"input__control")
            #print(elems)
        
        #time.sleep(t)
        # Вводим куда
        elem = WebDriverWait(driver, attempt_to_load_time).until(EC.element_to_be_clickable(elems[1]))
        elem.click()
        elem.send_keys(address_from)
        elem.send_keys(Keys.RETURN)
        print("Куда")
        #time.sleep(t)

        # Вводим откуда
        elem = WebDriverWait(driver, attempt_to_load_time).until(EC.element_to_be_clickable(elems[2]))
        elem.click()
        elem.send_keys(address_to)
        elem.send_keys(Keys.RETURN)
        print("Откуда")
        #   time.sleep(t)
    except:
        return[address_from, 
               address_to, 
               address_from + address_to, 
               False,
               'Не удалось загрузить маршруты',
               '',
               '',
               '']
    
    try:
        # Считываем данные для легкового 
        print("Считываем данные для легкового")
        elems = []
        start =datetime.now()
        while len(elems)<1 and (datetime.now()-start).total_seconds() < attempt_to_load_time:
            elems = driver.find_elements(By.CLASS_NAME,"auto-route-snippet-view__route-subtitle")
        elem = elems[0]  
        #elem = driver.find_elements(By.CLASS_NAME,"auto-route-snippet-view__route-subtitle")[0]
        car_distance_raw = elem.text
        if ' км' in car_distance_raw:
            car_distance_in_km = float(car_distance_raw.replace(' км','').replace(',','.'))
        elif ' м' in car_distance_raw:
            car_distance_in_km = float(car_distance_raw.replace(' м','').replace(',','.'))/1000
        else:
            car_distance_in_meters = -777
    except:
        return [address_from, 
                address_to, 
                address_from + address_to, 
                False, 
                'Не загрузился маршрут для легкового автомобиля',
                '',
                '',
                '']

    try:
        print("Считываем данные для тяжёлого грузовика")
        # Открываем параметры
        elem = WebDriverWait(driver, attempt_to_load_time).until(EC.element_to_be_clickable((By.CLASS_NAME,"route-list-view__settings")))
        elem.click()
        #time.sleep(t)

        # Переходим в выбор грузовика
        elem = WebDriverWait(driver, attempt_to_load_time).until(EC.element_to_be_clickable((By.CLASS_NAME,"_view_create")))
        elem.click()
        #print("Добавляем грузовик")
        #time.sleep(t)
        
        # Выбираем тяжёлый грузовик
        elems = []
        start =datetime.now()
        while len(elems)<3 and (datetime.now()-start).total_seconds() < attempt_to_load_time:
            elems = driver.find_elements(By.CLASS_NAME,"cargo-presets-view__item")
        elem = elems[2]
        elem.click()

        elem = WebDriverWait(driver, attempt_to_load_time).until(EC.element_to_be_clickable((By.CLASS_NAME,"_stretched")))
        elem.click()

        #time.sleep(t*3)
        elems = []

        start =datetime.now()
        while len(elems)<1 and (datetime.now()-start).total_seconds() < attempt_to_load_time:
            elems = driver.find_elements(By.CLASS_NAME,"auto-route-snippet-view__route-subtitle")
        elem = elems[0]
        truck_12t_distance_raw = elem.text
        if ' км' in truck_12t_distance_raw:
            truck_12t_distance_in_km = float(truck_12t_distance_raw.replace(' км','').replace(',','.'))
        elif ' м' in elem.text:
            truck_12t_distance_in_km = float(truck_12t_distance_raw.replace(' м','').replace(',','.'))/1000
        else:
            truck_12t_distance_in_meters = -777
        #print(car_distance_raw,car_distance_in_meters,sep=' ')
        return [address_from, 
                address_to, 
                address_from + address_to, 
                True, 
                car_distance_raw, 
                car_distance_in_km, 
                truck_12t_distance_raw, 
                truck_12t_distance_in_km]
    except Exception as e:
        return [address_from,
                address_to, 
                address_from + address_to, 
                True, 
                car_distance_raw, 
                car_distance_in_km,
                'Не удалось получить для грузовика',
                '']


def Run(headless, run_number, routs_qty, df_to_run):
    i = 0
    result_true_list = []
    result_false_list = []
    for row in df_to_run.itertuples():
        i += 1
        address_from = row[1]
        address_to   = row[2]
        print(Fore.BLACK, '-'*100)
        print('Прогон %s Парсим %s из %s Откуда: %s Куда: %s'%(run_number, i, routs_qty, address_from,address_to), sep ='\t')

        row_result = Get_distance(
                address_from,
                address_to,
                headless= headless
                )
        if row_result[3]:
            result_true_list.append(row_result)
        else:
            result_false_list.append(row_result)    
        print(row_result,'\n')
    return (result_true_list, result_false_list)    

def show_manual():
    with open('README.md', 'r', encoding='utf-8') as f:
        markdown_string = f.read()
    
    html_string = "<title>GET_DISTANCE_AUTO</title>\n" +  markdown.markdown(markdown_string)
    
    with open('READMY.html', 'w') as f:
        f.write(html_string)

    os.startfile('READMY.html')