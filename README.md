# Python Tools — автоматизация и отчёты

Два инструмента для бизнеса и частных лиц.

## 1. Сортировщик файлов
Раскладывает файлы из папки по категориям (Изображения, Документы, Архивы…).
Запуск: `python file_sorter.py`

## 2. Генератор отчётов
Строит красивый отчёт (HTML или Word) по таблице CSV/Excel:
статистика, пропуски, гистограммы, топ-10 категорий.
Запуск: `python report_generator_docx.py`

## Скриншоты
![Сортировщик файлов до и после](screens/do_posle.jpg)
![Пример отчета в Word](screens/word_otchet.jpg)

## Установка
pip install pandas matplotlib openpyxl python-docx pyinstaller
