import PyPDF2
import pandas as pd
from pathlib import Path
import os
import re
from tkinter import Tk, filedialog

def extract_data_from_pdf(pdf_file):
    try:
        with open(pdf_file, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ''.join([page.extract_text().replace('\n', ' ') for page in reader.pages])

            reestr_match = re.search(r'\b0*(\d{5})/\d{3}\b', text)
            reestr = reestr_match.group(1) if reestr_match else "Не найдено"
            if reestr == "Не найдено":
                print(f"Реестровый номер не найден в файле {pdf_file}")

            fio_start = text.find("документу с") + len("документу с")
            fio_end = text.find(",", fio_start)
            fio = text[fio_start:fio_end].strip() if fio_start != -1 and fio_end != -1 else "Не найдено"
            if fio == "Не найдено":
                print(f"ФИО не найдено в файле {pdf_file}")

            address_start = text.find("местонахождение:") + len("местонахождение:")
            address_end = text.find("в пользу")
            address = text[address_start:address_end].strip() if address_start != -1 and address_end != -1 else "Не найдено"
            if address == "Не найдено":
                print(f"Адрес не найден в файле {pdf_file}")

            parts = address.split(",")
            if len(parts) > 2:
                region = parts[1].strip()
                district = parts[2].strip()
            else:
                region = "Не найдено"
                district = "Не найдено"
                print(f"Область или район не найдены в файле {pdf_file}")

            fio_not_start = text.find("Я, ") + len("Я, ")
            fio_not_end = text.find(",", fio_not_start)
            fio_not = text[fio_not_start:fio_not_end].strip() if fio_not_start != -1 and fio_not_end != -1 else "Не найдено"
            if fio_not == "Не найдено":
                print(f"ФИО отправителя не найдено в файле {pdf_file}")

            regions_indexes = {
                "АБАЙ": "'070000", "АКМОЛИНСКАЯ ОБЛАСТЬ": "'020000", "АКТЮБИНСКАЯ ОБЛАСТЬ": "030000", "АЛМАТИНСКАЯ ОБЛАСТЬ": "", "АТЫРАУСКАЯ ОБЛАСТЬ": "060000",
                "В- КАЗАХСТАНСКАЯ": "070000", "ЖАМБЫЛСКАЯ ОБЛАСТЬ": "080000", "ЖЕТІСУ": "040000", "З-КАЗАХСТАНСКАЯ": "090000",
                "КАРАГАНДИНСКАЯ ОБЛАСТЬ": "100000", "КОСТАНАЙСКАЯ ОБЛАСТЬ": "110000", "КЫЗЫЛОРДИНСКАЯ ОБЛАСТЬ": "120000", "МАНГИСТАУСКАЯ ОБЛАСТЬ": "130000",
                "ПАВЛОДАРСКАЯ ОБЛАСТЬ": "140000", "С-КАЗАХСТАНСКАЯ ОБЛАСТЬ": "150000", "ТУРКЕСТАНСКАЯ ОБЛАСТЬ": "161200", "ҰЛЫТАУ": "101500",
                "АСТАНА": "010000", "ШЫМКЕНТ": "32522", "АЛМАТЫ": "050000"
            }
            
            region_index = regions_indexes.get(region, "Не найдено")
            if region_index == "Не найдено":
                print(f"Индекс области не найден для региона: {region} в файле {pdf_file}")
            
            return {
                'Название файла': reestr,
                'Кому': fio,
                'Телефон': '77000000000',
                'Адрес': address,
                'Город': district,
                'Область': region,
                'Индекс': region_index,
                'От кого': fio_not,
                'Кол-во стр': len(reader.pages)
            }
    except Exception as e:
        print(f"Ошибка при обработке файла {pdf_file}: {e}")
        return None

def create_newgep():
    root = Tk()
    root.withdraw()  # Hide the root window
    folder_path = filedialog.askdirectory(title="Выберите папку с PDF файлами")

    if not folder_path:
        print("Папка не выбрана.")
        return None

    folder_path = Path(folder_path)
    pdf_file_paths = list(folder_path.glob("*.pdf"))

    if not pdf_file_paths:
        print("Нет PDF файлов в выбранной папке.")
        return None

    data_list = []
    for pdf_file_path in pdf_file_paths:
        try:
            data = extract_data_from_pdf(pdf_file_path)
            if data:
                data_list.append(data)
                new_file_name = f"{data['Название файла']}.pdf"
                new_file_path = pdf_file_path.parent / new_file_name
                os.rename(pdf_file_path, new_file_path)
        except Exception as e:
            print(f"Ошибка при извлечении данных из файла {pdf_file_path}: {e}")

    if not data_list:
        print("Нет данных для записи в Excel.")
        return None

    try:
        df = pd.DataFrame(data_list)
        output_file_path = folder_path / 'newgep.xlsx'
        df.to_excel(output_file_path, index=False)
        print(f"Данные извлечены и сохранены в {output_file_path}")
        return output_file_path
    except Exception as e:
        print(f"Ошибка при создании Excel файла: {e}")
        return None

def main():
    try:
        output_file = create_newgep()
        if output_file:
            print(f"Данные извлечены и сохранены в {output_file}")
        else:
            print("Не удалось создать Excel файл.")
    except Exception as e:
        print(f"Общая ошибка: {e}")

if __name__ == '__main__':
    main()
    input("Нажмите Enter для выхода...")  # Чтобы окно не закрывалось сразу
