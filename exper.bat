@echo off
REM Запуск скрипта для проверки и установки openpyxl

echo Проверка наличия модуля openpyxl...

REM Попробуем импортировать openpyxl. Если модуля нет, установим его через pip
python -c "import openpyxl" 2>nul

IF %ERRORLEVEL% NEQ 0 (
    echo Модуль openpyxl не найден. Установка...
    python -m pip install openpyxl

    REM Проверяем, успешно ли установился модуль
    IF %ERRORLEVEL% NEQ 0 (
        echo Ошибка установки openpyxl. Завершение работы.
        exit /b
    )
) ELSE (
    echo Модуль openpyxl найден.
)

REM Запуск основного Python-скрипта
python - <<EOF
import os
import warnings
from openpyxl import load_workbook
warnings.filterwarnings("ignore", message="Data Validation extension is not supported")

def main():
    """
    Главная функция
    """
    file_in_dir = find_file()
    work_type = function_column_value_counting('C', file=file_in_dir)
    work_ed = function_column_value_counting('E', file_in_dir)
    write_date(work_type=work_type, work_ed=work_ed, file=file_in_dir)

def write_date(work_type: list, work_ed: list, file):
    """
    Функция получает на вход два списка, вносит данные из списков в указанные столбцы
    Сохраняет отдельную копию файла с изменёнными данными
    """
    wb = load_workbook(file)
    ws = wb.active
    for i in range(len(work_ed)):
        ws['C' + str(i + 3)] = work_type[i].strip()
        ws['E' + str(i + 3)] = work_ed[i]
    print(f'Запись данных в "new_{file}"')
    wb.save('new_' + file)
    print('Запись завершена')

def function_column_value_counting(column: str, file) -> list:
    """
    Функция получает на вход имя столбца и возвращает список данных из этого столбца
    """
    print(f'Загрузка данных из Справочника столбец "{column}"')

    # Загружаем файл
    wb = load_workbook(file)

    # Выбираем лист в таблице (в данном случае это "Справочник")
    ws = wb['Справочник']

    # Определяем количество строк
    max_row = ws.max_row
    result = []

    for row in range(1, max_row + 1):
        cell = ws[f'{column}{row}']
        if cell.value is not None:
            result.append(cell.value)
        else:
            break
    print(f'Загрузка данных из столбца "{column}" завершена')
    return result[1:]

def find_file() -> str:
    """
    Функция находит файл с расширением '.xlsx' и возвращает его название
    """
    for file in os.listdir():
        if file.endswith('.xlsx'):
            return file
    return None

if __name__ == '__main__':
    main()

EOF

REM Завершение
echo Python скрипт завершён.
pause
