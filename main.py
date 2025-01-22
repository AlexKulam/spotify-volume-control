import tkinter as tk
from tkinter import messagebox
import keyboard
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import comtypes
import threading
import pyautogui

# Функция для инициализации COM-объектов в текущем потоке
def initialize_com():
    comtypes.CoInitialize()

# Функция для регулировки громкости Spotify
def adjust_spotify_volume(delta):
    initialize_com()  # Инициализация COM-объектов в текущем потоке
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        # Проверяем, что это процесс Spotify
        if session.Process and session.Process.name() == "Spotify.exe":
            volume = session.SimpleAudioVolume
            current_volume = volume.GetMasterVolume()
            new_volume = max(0.0, min(1.0, current_volume + delta))
            volume.SetMasterVolume(new_volume, None)
            print(f"Spotify volume set to {new_volume * 100:.0f}%")

# Функция для паузы/воспроизведения
def toggle_play_pause():
    pyautogui.press('playpause')  # Эмулируем нажатие клавиши Play/Pause
    print("Play/Pause pressed")

# Функция для запуска adjust_spotify_volume в отдельном потоке
def run_in_thread(delta):
    thread = threading.Thread(target=adjust_spotify_volume, args=(delta,))
    thread.start()
    thread.join()  # Ожидаем завершения потока

# Переменная для отслеживания состояния функции
is_function_active = False

# Функция для включения/выключения функции
def toggle_function():
    global is_function_active
    is_function_active = not is_function_active

    if is_function_active:
        # Включаем горячие клавиши
        keyboard.add_hotkey('shift+alt+1', lambda: run_in_thread(-0.1))  # Уменьшение громкости
        keyboard.add_hotkey('shift+alt+2', lambda: run_in_thread(0.1))   # Увеличение громкости
        keyboard.add_hotkey('shift+alt+3', toggle_play_pause)            # Пауза/воспроизведение
        status_label.config(text="Функция включена", fg="green")
    else:
        # Отключаем горячие клавиши
        keyboard.unhook_all()  # Удаляем все зарегистрированные горячие клавиши
        status_label.config(text="Функция выключена", fg="red")

# Создаем главное окно
root = tk.Tk()
root.title("Управление Spotify")
root.geometry("300x100")  # Размер окна

# Кнопка для включения/выключения функции
toggle_button = tk.Button(root, text="Включить функцию", command=toggle_function)
toggle_button.pack(pady=10)

# Метка для отображения статуса
status_label = tk.Label(root, text="Функция выключена", fg="red")
status_label.pack()

# Запуск основного цикла окна
root.mainloop()