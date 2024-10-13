# Инструкция по развертыванию проекта.
### Cоздание виртуального окружения.
1. `python3 -m venv venv`
### Активация виртуального окружения.
2. `venv\Scripts\activate.bat` - для Windows;
3. `source venv/bin/activate` - для Linux и MacOS.
### Подключить все библиотеки и зависимости проекта.
4. `pip install -r requirements.txt`


# Применение

1. Поместить файл Excel в папку с проектом и переименовать в `script` 
2. Поместить файл Word в папку с проектом и переименовать в `passport`
3. Запустить скрипт командой `python3 main.py`