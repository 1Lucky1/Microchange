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

import win32com.client

# Получаем текущую локаль
locale.setlocale(locale.LC_ALL, '')
is_russian = locale.getlocale()[0].partition("Russian")[1] == "Russian"

devices_json: dict = {}
last_devices_json: dict = {}
VERSION = "1.7"

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
path = sys.argv[0]
APP_NAME = os.path.basename(path)
APP_PATH = os.path.abspath(path)
STARTUP_PATH = os.path.expandvars(r"%Appdata%\Microsoft\Windows\Start Menu\Programs\Startup")


def setup_startup():
    """Создает ярлык на FullMicrochange.exe в папке автозагрузки."""
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut_path = os.path.join(STARTUP_PATH, f"{APP_NAME}.lnk")

    # Удаляем старый ярлык, если он существует
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)

    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.TargetPath = APP_PATH
    shortcut.WorkingDirectory = os.path.dirname(APP_PATH)
    shortcut.IconLocation = APP_PATH
    shortcut.save()


def shortcut_exists(icon_copy=None):  # Параметр - это экземпляр icon, который сюда pystray передает зачем-то
    """Проверяет наличие ярлыка в папке автозагрузки."""
    shortcut_path = os.path.join(STARTUP_PATH, f"{APP_NAME}.lnk")
    return os.path.exists(shortcut_path)


def remove_startup():
    """Удаляет ярлык из автозагрузки."""
    shortcut_path = os.path.join(STARTUP_PATH, f"{APP_NAME}.lnk")
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)
    else:
        logging.error("Ярлык не найден в автозагрузке.")


def toggle_startup():
    if shortcut_exists():
        remove_startup()
    else:
        setup_startup()


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
                MenuItem(autostart, toggle_startup, shortcut_exists),
                Menu.SEPARATOR,
                MenuItem(close, stop),
                Menu.SEPARATOR,
                MenuItem(vers, lambda x: True))
    return menu


def stop():
    """Завершает процесс программы"""
    subprocess.run(["taskkill", "/F", "/IM", APP_NAME], check=True,
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
