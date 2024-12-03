import sys

import os
from traceback import print_exc

import pystray
import subprocess
from threading import Timer
from PIL import Image
from pystray import Menu, MenuItem
import locale


# Получаем текущую локаль
locale.setlocale(locale.LC_ALL, '')
is_russian = True if locale.getlocale()[0].partition("Russian")[1] == "Russian" else False

devices_json = {}
last_devices_json = {}
version = "1.5.1"

russian = {"close": "Закрыть программу", "version": f"Версия {version}"}
english = {"close": "Close application", "version": f"Version {version}"}


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# noinspection PyTypeChecker

def get_audio_devices_list():
    global devices_json
    global last_devices_json

    # Получаем список аудиоустройств
    subprocess.run(['powershell', '-Command',
                    "Get-AudioDevice -List | Out-File -FilePath audio_devices.txt -Encoding utf8"],
                   creationflags=subprocess.CREATE_NO_WINDOW)

    with open('audio_devices.txt', 'r', encoding="utf-8") as f:
        audio_devices = f.read().split('\n')
        # Избавляемся от непонятного /ufufe или как-то так
        audio_devices.pop(0)

        # Избавляемся от пустых строк, переносов
        audio_devices = [string for string in audio_devices if string]

        # Раскидываем инфу о девайсах строго в свои массивы
        temp = []
        splitted_devices = []

        for item in audio_devices:
            temp.append(item)

            if len(temp) == 7:
                splitted_devices.append(temp)
                temp = []

        # Околофинально отделяем записывающие устройства и убираем устройства вывода
        only_recording_devices = []
        for device in splitted_devices:
            if "Type                 : Playback" not in device:
                only_recording_devices.append(device)

        # Окончательно собираем кнопки с нужной информацией и индексами
        menu_items = []

        new_devices_json = {}

        for device in only_recording_devices:
            index = device[0].split()[-1]
            is_default = True if device[1].split('Default              : ')[1] == "True" else False
            name = device[4].split("Name                 : ")[1]
            new_devices_json[index] = is_default
            menu_items.append(MenuItem(f"{index} - {name}", lambda _, *, x=index: set_default_microphone(x),
                                       checked=lambda _, *, x=index: True if devices_json[x] else False,
                                       radio=True))

        # Проверка на различия между новыми и старыми данными
        if new_devices_json == devices_json:
            return False  # Возвращаем False, если наборы отличаются

        # Обновляем last_devices_json только если устройства не изменились
        last_devices_json = devices_json.copy()
        devices_json = new_devices_json

        return menu_items


def set_default_microphone(index: int):
    global devices_json
    subprocess.run(['powershell', '-Command', f'Set-AudioDevice -Index {index}'],
                   creationflags=subprocess.CREATE_NO_WINDOW,
                   capture_output=True, text=True)

    for key, value in devices_json.items():
        devices_json[key] = False

    devices_json[index] = True


def create_icon():
    icon = pystray.Icon("Microchange")
    # image = Image.new('RGB', (32, 32), color='white')
    image = Image.open(resource_path("new_icon.ico"))
    icon.icon = image

    return icon


def create_menu():
    input_devices = get_audio_devices_list()
    if not input_devices:
        return False
    if is_russian:
        close, vers = russian["close"], russian["version"]
    else:
        close, vers = english["close"], english["version"]
    menu = Menu(*input_devices,
                MenuItem(close, stop),
                Menu.SEPARATOR,
                MenuItem(vers, lambda x: True))
    return menu


def stop():
    subprocess.run(["taskkill", "/F", "/IM", "Microchange.exe"], check=True,
                   creationflags=subprocess.CREATE_NO_WINDOW,)


def update_devices():
    # """Обновляет меню с аудиоустройствами."""
    new_menu = create_menu()
    if new_menu:
        icon.menu = new_menu  # Обновляем меню
        icon.update_menu()
        # Запускаем таймер снова через 1 секунду (или любое другое время)
        Timer(1, update_devices).start()
    else:
        Timer(1, update_devices).start()


if __name__ == "__main__":
    try:
        icon = create_icon()
        icon.menu = create_menu()

        update_devices()

        icon.run()

    except Exception:
        with open(os.path.join(os.path.expanduser("~"), "Desktop", "Microchange_error_log.txt"), "w",
                  encoding="utf-8") as file:
            print_exc(file=file)
