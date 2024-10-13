import pandas as pd
import logging

def extract_data_from_excel(excel_file):
    try:
        logging.info(f"Чтение данных из Excel файла: {excel_file}")
        df = pd.read_excel(excel_file, sheet_name='Support')

        # Логируем названия колонок для отладки
        logging.info(f"Названия колонок до обработки: {df.columns.tolist()}")

        # Убираем лишние пробелы в названиях колонок
        df.columns = [col.strip() for col in df.columns]

        # Логируем названия колонок после удаления пробелов
        logging.info(f"Названия колонок после обработки: {df.columns.tolist()}")

        # Проверка, что нужные колонки существуют
        required_columns = ['Имя сервера', 'Сайзинг\ncpu/ram/hdd sys/hdd app', 'IP адрес']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            raise ValueError(f"Отсутствуют следующие колонки: {missing_columns}")

        # Переименовываем и отбираем нужные колонки
        df = df[['Имя сервера', 'Сайзинг\ncpu/ram/hdd sys/hdd app', 'IP адрес']]
        df = df.rename(columns={'Сайзинг\ncpu/ram/hdd sys/hdd app': 'Сайзинг', 'IP адрес': 'IP адрес'})
        df = df.dropna()

        logging.info(f"Извлечено строк из Excel: {len(df)}")
        return df

    except Exception as e:
        logging.error(f"Ошибка при обработке Excel файла: {e}")
        return pd.DataFrame()
