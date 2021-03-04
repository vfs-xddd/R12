from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from datetime import timedelta, datetime
from pyautogui import alert, confirm
import os, sys, psutil, socket
from pynput import keyboard
from R12_TnV_file_reader import file_reader

HOST = 'localhost'  # The remote host
PORT = 55037

def find_process(process_name):
    # находим объект процесса по имени
    for process in psutil.process_iter():
        if process.name() == process_name:
            return process


def send_to_launcher(data='exit', host=HOST, port=PORT):
    with socket.socket() as s:
        s.connect((host, port))
        s.send(data.encode('utf-8'))


def event_full_exit():
    send_to_launcher(data='interupted_exit')
    current_system_pid = os.getpid()
    this_sys = psutil.Process(current_system_pid)
    this_sys.terminate()


def new_win_check(browser, win_hdl_befor_amount, value, description):
    win_hdl_after_amount = len(browser.window_handles)
    if win_hdl_after_amount > win_hdl_befor_amount:
        browser.switch_to.window(browser.window_handles[-1])
        browser.close()
        alert(
            f'Введенное значение\n \'{value}\'\n для\n \'{description}\'\n отсутствует на сервере, короткие тайминги загрузки, в случае разового события возможно просто подзавис браузер.',
            'Ошибка в записи!',
            'Выход')
        send_to_launcher()
        sys.exit()


def form_send_key(browser, id_value_dict, timeout=7):
    # Заполняем поле формы ТнВ
    win_hdl_befor_amount = len(browser.window_handles)
    loading_frame = browser.find_element(By.ID, 'LLAOutput')
    id = id_value_dict['id']
    value = id_value_dict['value']
    description = id_value_dict['descr']

    if browser.find_element(By.ID, id).get_attribute('value'):  # если значение есть по дефолту, то удаляем
        browser.execute_script("arguments[0].setAttribute('value','')", browser.find_element(By.ID, id))

    browser.find_element(By.ID, id).send_keys(value + '\n')

    # если атрибут onchange начинается с return, т.е. обновляем элемент при измен значения
    if browser.find_element(By.ID, id).get_attribute('onchange')[:6] == 'return':
        WebDriverWait(browser, timeout).until(EC.staleness_of(browser.find_element(By.ID, id)))
        WebDriverWait(browser, timeout).until(EC.visibility_of_element_located((By.ID, id)))
        new_win_check(browser, win_hdl_befor_amount, value, description)
        WebDriverWait(browser, timeout).until(EC.text_to_be_present_in_element_value((By.ID, id), value))
    browser.execute_script('arguments[0].scrollIntoView(true);', browser.find_element(By.ID, id))
    browser.execute_script("arguments[0].style.visibility='hidden'", loading_frame)


def check_presence_of_elem(browser, id, alert_text='', alert_title='', alert_button='Выход', timeout=5):
    try:
        WebDriverWait(browser, timeout).until(EC.presence_of_element_located((By.ID, id)))
    except TimeoutException:
        alert(alert_text, alert_title, alert_button)
        send_to_launcher()
        sys.exit()


def main():
    # тело скрипта

    url_tnv = 'http://r12a.ks.rt.ru:8000/OA_HTML/RF.jsp?function_id=53248&resp_id=91792&resp_appl_id=401&security_group_id=0&lang_code=RU&params=klKwI075XGM.CVKHtxvV7-P5eDLBQMIRNDhQ4sKoD3k'

    if find_process('R12_launcher.exe') is None:
        alert('Программа запускается с R12_launcher.exe. Текущий файл не является основным процессом.', 'Ошибка!',
              'Выход')
        sys.exit()
    try:
        data = file_reader('../config_R12_TnV.txt')
    except FileNotFoundError:
        alert('Отсутствует файл конфигурации config_R12_TnV.txt', 'Ошибка!', 'Выход')
        send_to_launcher('error')
        sys.exit()
    except Exception:
        alert('Ошибки обработки файла конфигурации config_R12_TnV.txt', 'Ошибка!', 'Ок')
        sys.exit()
    login = data['login']
    password = data['password']
    template = data['template']  # полное имя файла-шаблона вида файл.xlsx
    req_date_add = int(data['req_date_add'])  # сколько добавить к дате получения
    timeout = int(data['timeout'])  # timeout для всех WebDriverWait()
    tnv_inputs = {
        'oe_sorce': {'id': 'ReqHdrSourceOperUnit', 'value': data['oe_sorce'],
                     'descr': 'ОЕ источник'},
        'osu': {'id': 'ReqHdrSourceOrganization', 'value': data['osu'], 'descr': 'Склад отправки (ОСУ)'},
        'segment': {'id': 'ReqHdrSourceSegment', 'value': data['segment'], 'descr': 'Сегмент источник'},
        'oe_reciever': {'id': 'ReqHdrDestOperatingUnit', 'value': data['oe_reciever'],
                        'descr': 'ОЕ получатель'},
        'cfo': {'id': 'ReqHdrCfo', 'value': data['cfo'], 'descr': 'ЦФО'},
        'date': {'id': 'HdrReqDate', 'value': '', 'descr': 'Дата получения'}}

    template_path = os.path.join(os.getcwd(), template)
    template_exsists = os.path.isfile(template_path)
    if not template_exsists:
        answer = confirm(f'Шаблон ТнВ {template} не обнаружен!\n Продолжаем без шаблона?', 'Внимание!', ['Да', 'Выход'])
        if answer == 'Выход':
            send_to_launcher()
            sys.exit()

    with keyboard.GlobalHotKeys({'<ctrl>+q': event_full_exit}) as listener:
        try:
            browser = webdriver.Ie()
            browser.get(url_tnv)
            browser.maximize_window()
            check_presence_of_elem(browser, 'usernameField', 'Не могу получить окно \'HOMEPAGE\'.', 'Ошибка, не вижу окно входа!', timeout=timeout)
            browser.find_element(By.ID, 'usernameField').send_keys(login + '\n')
            browser.find_element(By.ID, 'passwordField').send_keys(password + '\n')
            browser.find_element(By.ID, 'SubmitButton').click()
            check_presence_of_elem(browser, 'RequestBtn', 'Не могу получить окно \'создание ТнВ\'.', 'Ошибка, не вижу окно ТнВ!', timeout=timeout)
            browser.find_element(By.ID, 'RequestBtn').click()  # Нажали Создание ТнВ

            today = datetime.strptime(browser.find_element(By.ID, 'HdrCreationDate').text, '%d.%m.%Y')
            req_date = today + timedelta(req_date_add)  # считываем дату с формы и добавляем дни
            tnv_inputs['date']['value'] = req_date.strftime("%d.%m.%Y")

            [form_send_key(browser, value, timeout=timeout) for value in tnv_inputs.values()]

            if template_exsists:
                browser.find_element(By.ID, 'DownloadFromExcelBtn').click()  # нажали загрузка шаблона Excel
                WebDriverWait(browser, timeout).until(
                    EC.visibility_of_element_located((By.ID, 'MessageFileUpload_oafileUpload')))
                browser.find_element(By.ID, 'MessageFileUpload_oafileUpload').send_keys(template_path)
                WebDriverWait(browser, timeout).until(
                    EC.text_to_be_present_in_element_value((By.ID, 'MessageFileUpload_oafileUpload'), template_path))
                browser.find_element(By.ID, 'ApplyBtn').click()  # нажали применить, шаблон успешно загрузился
        except TimeoutException:
            alert('Выполнено частично, ошибка в таймингах, возможно следует увеличить значение тайминга',
                  'Ошибка таймингов!',
                  'Ясно')
            send_to_launcher('error')
        except ConnectionResetError: # ok, т.к. вызвана прерывнием ивентом с клавиатуры
            pass
        except Exception as error:
            alert(f'Выполнено частично, неизвестная ошибка.\n{error}', 'С ошибкой!', 'Понял. Принял(')
            send_to_launcher('error')
        else:
            alert('Выполнено', 'Успешно!', 'OK')
            send_to_launcher()
        finally: pass

        listener.join()

main()
