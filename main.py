import logging
from word_processing import filter_ip_address_tables
from excel_processing import extract_data_from_excel
from compare import compare_data

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def main():
    try:
        # Пути к файлам
        excel_file = 'script.xlsx'  # Путь к Excel-файлу
        word_file = 'passport.docx'  # Путь к Word-документу

        # Извлечение данных из Word
        df_word = filter_ip_address_tables(word_file)

        # Извлечение данных из Excel
        df_excel = extract_data_from_excel(excel_file)

        # Сравнение данных
        compare_data(df_excel, df_word)

    except Exception as e:
        logging.error(f"Ошибка в основной функции: {e}")


if __name__ == "__main__":
    main()
