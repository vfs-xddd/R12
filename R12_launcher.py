import os, sys, socket, psutil
from subprocess import PIPE, Popen
from pyautogui import alert
from time import sleep

HOST = ''
PORT = 55037
SW_HIDE = 1  # скроем консольное окно субпроцесса - системная переменная


def find_process(process_name, timeout=1):
    # находим объект процесса по имени
    while timeout:
        for process in psutil.process_iter():
            if process.name() == process_name:
                return process
        if timeout != 1: sleep(1)
        timeout-=1

print('Запускаю скрипт...\n')
if os.getcwd() not in os.environ['Path'].split(';'):
    print('Необходимый файл не добавлен в переменную окрженияя системы.\nДобавляю..')
    os.environ['Path'] = ';'.join([os.environ['Path'], os.getcwd()])
    if os.getcwd() in os.environ['Path'].split(';'):
        print('Успешно добавлен. Необходима перезагрузка!')
file_path = os.path.join(os.getcwd(), 'R12_TnV.exe')
file_exsists = os.path.isfile(file_path)
if not file_exsists:
    alert('Не обнаружен файл R12_TnV.exe!', 'Ошибка!', 'Выход')
    sys.exit()
print('R12_TnV найден, запущен.\nВыполняю основной процесс скрипта...')
#process = Popen([sys.executable, 'R12_TnV.py'], stdout=PIPE, stderr=PIPE)
process = Popen(['R12_TnV.exe' ], stdout=PIPE, stderr=PIPE)
if find_process('R12_TnV.exe', timeout=5) is None:
    alert('Не запустился процесс R12_TnV!', 'Ошибка!', 'Выход')
    print('Ошибка в работе.\nЗавершено.')
    sys.exit()
with socket.socket() as s:
    s.bind((HOST, PORT))
    s.listen(1)
    conn, addr = s.accept()
    with conn:
        print('Connected by', addr)
        while process.poll() is None:
            data = conn.recv(1024)
            if data == b'interupted_exit':
                alert('Прервано пользователем', 'Прервано!', 'Выход')
                break
            if data == b'exit':
                print('Успешно!')
                break
            if data == b'error':
                print('Ошибка в работе.')
                break
            # if not data: break
            # conn.sendall(data)
try:
    find_process('IEDriverServer.exe').terminate()
except AttributeError: pass
print('Завершено.')

'''stdout, stderr = process.communicate()
print('stdout', stdout.decode())
print('stderr', stderr.decode())'''

'''while process.poll() is None:
    output_line = process.stdout.readline()
    print('f', output_line.decode()[:-2]) #обрезаем \r\n
    if output_line.decode()[:-2] == 'exit':
        alert('Прервано пользователем', 'Прервано!', 'Выход')
        #current_system_pid = os.getpid()
        #this_sys = psutil.Process(current_system_pid)
        #this_sys.terminate()
        process.kill()
        sys.exit()'''
'''alert('Прервано пользователем', 'Прервано!', 'Выход')
current_system_pid = os.getpid()
this_sys = psutil.Process(current_system_pid)
this_sys.terminate()'''
# sys.exit()
