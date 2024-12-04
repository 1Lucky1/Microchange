import sys

import os
from traceback import format_exc
import winreg
import logging

import subprocess
import locale
from threading import Timer
from PIL import Image
from pystray import Menu, MenuItem, Icon

# Получаем текущую локаль
locale.setlocale(locale.LC_ALL, '')
is_russian = locale.getlocale()[0].partition("Russian")[1] == "Russian"

devices_json: dict = {}
last_devices_json: dict = {}
VERSION = "1.6.1"

russian = {"close": "Закрыть программу", "version": f"Версия {VERSION}", "autostart": "Автозапуск с системой"}
english = {"close": "Close application", "version": f"Version {VERSION}", "autostart": "Auto start"}

# Настройка логгирования
log_file_path = os.path.join(os.path.expanduser("~"), "Desktop", "Microchange_error_log.txt")

logging.basicConfig(
    filename=log_file_path,
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# noinspection PyTypeChecker


# Путь к исполняемому файлу
APP_NAME: str = "Microchange.exe"
app_path = os.path.join(os.path.dirname(os.path.abspath("Microchange.exe")), "Microchange.exe")


def add_to_startup():
    """Добавляет Microchange.exe в автозагрузку."""
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0,
                         winreg.KEY_SET_VALUE)
    winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, app_path)
    winreg.CloseKey(key)
    # print(f"{APP_NAME} добавлен в автозагрузку.")


def remove_from_startup():
    """Убирает Microchange.exe из автозагрузки."""
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0,
                         winreg.KEY_SET_VALUE)
    try:
        winreg.DeleteValue(key, APP_NAME)
        # print(f"{APP_NAME} убран из автозагрузки.")
    except FileNotFoundError as e:
        logging.error("Cannot remove from startup registry, because it doesn't exist.")
        logging.error(e.with_traceback)
        # print(f"{APP_NAME} не найден в автозагрузке.")
    finally:
        winreg.CloseKey(key)


def is_in_startup(
        MenuItemParam=None):  # Параметр здесь нужен, чтобы избежать исключения, когда MenuItem за каким-то хером передаёт себя в эту функцию!

    """Проверяет, добавлен ли Microchange.exe в автозагрузку."""
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Run")
    try:
        winreg.QueryValueEx(key, APP_NAME)[0]
        return True
    except FileNotFoundError:
        return False
    finally:
        winreg.CloseKey(key)


def toggle_autostart():
    """Переключает состояние Microchange.exe в автозагрузке."""
    if is_in_startup():
        remove_from_startup()
    else:
        add_to_startup()


def get_audio_devices_list():
    """Формирует список кнопок с записывающими устройствами"""
    global devices_json
    global last_devices_json

    # Получаем список аудиоустройств
    subprocess.run(['powershell', '-Command',
                    "Get-AudioDevice -List | Out-File -FilePath audio_devices.txt -Encoding utf8"],
                   creationflags=subprocess.CREATE_NO_WINDOW)

    with open('audio_devices.txt', 'r', encoding="utf-8") as f:

        # Избавляемся от пустых строк, переносов
        audio_devices = [string for string in f.read().split('\n')[1:] if string]

        # Группируем по 7 строк
        splitted_devices = [audio_devices[i:i + 7] for i in range(0, len(audio_devices), 7)]

        # Фильтруем только записывающие устройства
        only_recording_devices = [device for device in splitted_devices if
                                  "Type                 : Playback" not in device]

        menu_items = []
        new_devices_json = {}

        for device in only_recording_devices:
            index = device[0].split()[-1]
            is_default = device[1].split('Default              : ')[1].strip() == "True"
            name = device[4].split("Name                 : ")[1].strip()
            new_devices_json[index] = is_default
            menu_items.append(MenuItem(
                f"{index} - {name}",
                lambda _, *, x=index: set_default_microphone(x),
                checked=lambda _, *, x=index: devices_json[x] is True,
                radio=True
            ))

        # Проверка на различия между новыми и старыми данными
        if new_devices_json == devices_json:
            return False  # Возвращаем False, если наборы одинаковые

        # Обновляем last_devices_json только если устройства не изменились
        last_devices_json = devices_json.copy()
        devices_json = new_devices_json

        return menu_items


def set_default_microphone(index: int):
    """Устанавливает стандартное записывающее устройство"""
    global devices_json
    subprocess.run(['powershell', '-Command', f'Set-AudioDevice -Index {index}'],
                   creationflags=subprocess.CREATE_NO_WINDOW,
                   capture_output=True, text=True)

    for key, value in devices_json.items():
        devices_json[key] = False

    devices_json[index] = True


def create_icon():
    """Создает экземпляр трей приложения"""
    icon = Icon("Microchange")
    # image = Image.new('RGB', (32, 32), color='white')
    image = Image.open(resource_path("new_icon.ico"))
    icon.icon = image

    return icon


def create_menu():
    """Создает меню для иконки в трее"""
    input_devices = get_audio_devices_list()
    if not input_devices:
        return False
    if is_russian:
        close, vers, autostart = russian["close"], russian["version"], russian["autostart"]
    else:
        close, vers, autostart = english["close"], english["version"], english["autostart"]
    menu = Menu(*input_devices,
                Menu.SEPARATOR,
                MenuItem(autostart, toggle_autostart, is_in_startup),
                Menu.SEPARATOR,
                MenuItem(close, stop),
                Menu.SEPARATOR,
                MenuItem(vers, lambda x: True))
    return menu


def stop():
    """Завершает процесс программы"""
    subprocess.run(["taskkill", "/F", "/IM", "Microchange.exe"], check=True,
                   creationflags=subprocess.CREATE_NO_WINDOW, )


def update_devices():
    """Обновляет меню с аудиоустройствами."""
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
        logging.error(f"Main block error: {format_exc()}")
