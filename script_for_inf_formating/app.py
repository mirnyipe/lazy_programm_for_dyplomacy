import os
import sys
import re
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Словарь для преобразования номеров месяцев в названия
MONTH_NAMES = {
    '1': 'января', '01': 'января',
    '2': 'февраля', '02': 'февраля',
    '3': 'марта', '03': 'марта',
    '4': 'апреля', '04': 'апреля',
    '5': 'мая', '05': 'мая',
    '6': 'июня', '06': 'июня',
    '7': 'июля', '07': 'июля',
    '8': 'августа', '08': 'августа',
    '9': 'сентября', '09': 'сентября',
    '10': 'октября',
    '11': 'ноября',
    '12': 'декабря'
}


def set_justify_alignment(doc):
    """
    Устанавливает выравнивание текста по ширине во всем документе.
    """
    try:
        # Обрабатываем параграфы в основном документе
        for paragraph in doc.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Обрабатываем таблицы
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        print("✅ Установлено выравнивание текста по ширине")
        return True

    except Exception as e:
        print(f"❌ Ошибка при установке выравнивания текста: {e}")
        return False


def replace_quotes(text):
    """
    Заменяет прямые двойные кавычки на типографские кавычки-лапки.
    """
    # Заменяем открывающие кавычки (кавычка в начале строки или после пробела/знака препинания)
    text = re.sub(r'(^|\s)"', r'\1«', text)
    # Заменяем закрывающие кавычки (кавычка перед пробелом/знаком препинания/концом строки)
    text = re.sub(r'"(\s|[.!?;,]|$)', r'»\1', text)
    # Для оставшихся кавычек предполагаем, что они закрывающие
    text = text.replace('"', '»')
    return text


def process_quotes(doc):
    """
    Обрабатывает замену кавычек во всем документе.
    """
    try:
        # Обрабатываем параграфы в основном документе
        for paragraph in doc.paragraphs:
            process_paragraph_quotes(paragraph)

        # Обрабатываем таблицы
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        process_paragraph_quotes(paragraph)

        print("✅ Заменены прямые кавычки на типографские")
        return True

    except Exception as e:
        print(f"❌ Ошибка при обработке кавычек: {e}")
        return False


def process_paragraph_quotes(paragraph):
    """
    Обрабатывает замену кавычек в параграфе.
    """
    # Собираем весь текст параграфа
    full_text = ''.join([run.text for run in paragraph.runs])

    if not full_text.strip():
        return

    # Заменяем кавычки
    corrected_text = replace_quotes(full_text)

    # Если текст изменился, обновляем параграф
    if corrected_text != full_text:
        # Сохраняем форматирование первого run для применения к новому тексту
        first_run_format = {}
        if paragraph.runs:
            first_run = paragraph.runs[0]
            first_run_format = {
                'font_name': first_run.font.name,
                'font_size': first_run.font.size,
                'bold': first_run.bold,
                'italic': first_run.italic,
                'underline': first_run.underline
            }

        # Очищаем параграф и добавляем текст с исправлениями
        paragraph.clear()
        run = paragraph.add_run(corrected_text)

        # Применяем базовое форматирование
        run.font.name = first_run_format.get('font_name', 'Times New Roman')
        run.font.size = first_run_format.get('font_size', Pt(14))
        if first_run_format.get('bold') is not None:
            run.bold = first_run_format.get('bold')
        if first_run_format.get('italic') is not None:
            run.italic = first_run_format.get('italic')
        if first_run_format.get('underline') is not None:
            run.underline = first_run_format.get('underline')


def replace_special_spaces(text):
    """
    Заменяет специальные пробельные символы на обычные пробелы.
    Также сжимает множественные пробелы в один.
    """
    # Заменяем различные виды пробелов на обычный пробел
    # \u00A0 - неразрывный пробел
    # \u2009 - тонкий пробел
    # \u200A - волосистый пробел
    # \u200B - нулевая ширина пробела
    # \u202F - узкий неразрывный пробел
    # \u205F - средний математический пробел
    # \u3000 - идеографический пробел
    special_spaces = [
        '\u00A0',  # неразрывный пробел
        '\u2009',  # тонкий пробел
        '\u200A',  # волосистый пробел
        '\u200B',  # нулевая ширина пробела
        '\u202F',  # узкий неразрывный пробел
        '\u205F',  # средний математический пробел
        '\u3000',  # идеографический пробел
    ]

    for space in special_spaces:
        text = text.replace(space, ' ')

    # Заменяем множественные пробелы на один
    text = re.sub(r' +', ' ', text)

    return text


def process_special_spaces(doc):
    """
    Обрабатывает замену специальных пробелов на обычные пробелы во всем документе.
    """
    try:
        # Обрабатываем параграфы в основном документе
        for paragraph in doc.paragraphs:
            process_paragraph_special_spaces(paragraph)

        # Обрабатываем таблицы
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        process_paragraph_special_spaces(paragraph)

        print("✅ Заменены специальные пробелы на обычные и сжаты множественные пробелы")
        return True

    except Exception as e:
        print(f"❌ Ошибка при обработке специальных пробелов: {e}")
        return False


def process_paragraph_special_spaces(paragraph):
    """
    Обрабатывает замену специальных пробелов в параграфе.
    """
    # Собираем весь текст параграфа
    full_text = ''.join([run.text for run in paragraph.runs])

    if not full_text.strip():
        return

    # Заменяем специальные пробелы
    corrected_text = replace_special_spaces(full_text)

    # Если текст изменился, обновляем параграф
    if corrected_text != full_text:
        # Сохраняем форматирование первого run для применения к новому тексту
        first_run_format = {}
        if paragraph.runs:
            first_run = paragraph.runs[0]
            first_run_format = {
                'font_name': first_run.font.name,
                'font_size': first_run.font.size,
                'bold': first_run.bold,
                'italic': first_run.italic,
                'underline': first_run.underline
            }

        # Очищаем параграф и добавляем текст с исправлениями
        paragraph.clear()
        run = paragraph.add_run(corrected_text)

        # Применяем базовое форматирование
        run.font.name = first_run_format.get('font_name', 'Times New Roman')
        run.font.size = first_run_format.get('font_size', Pt(14))
        if first_run_format.get('bold') is not None:
            run.bold = first_run_format.get('bold')
        if first_run_format.get('italic') is not None:
            run.italic = first_run_format.get('italic')
        if first_run_format.get('underline') is not None:
            run.underline = first_run_format.get('underline')


def add_space_before_percent(text):
    """
    Добавляет пробел перед знаком процента, если его нет.
    """
    # Паттерн ищет число (целое или десятичное), за которым НЕ следует пробел, и затем знак %
    # Используем отрицательный просмотр назад, чтобы не трогать уже правильные варианты
    pattern = r'(\d+(?:[.,]\d+)?)\s*(?<!\s)%'
    # Заменяем на число + пробел + %
    return re.sub(pattern, r'\1 %', text)


def process_percent_signs(doc):
    """
    Обрабатывает добавление пробелов перед знаками процента во всем документе.
    """
    try:
        # Обрабатываем параграфы в основном документе
        for paragraph in doc.paragraphs:
            process_paragraph_percent_signs(paragraph)

        # Обрабатываем таблицы
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        process_paragraph_percent_signs(paragraph)

        print("✅ Добавлены пробелы перед знаками процента")
        return True

    except Exception as e:
        print(f"❌ Ошибка при обработке знаков процента: {e}")
        return False


def process_paragraph_percent_signs(paragraph):
    """
    Обрабатывает добавление пробелов перед знаками процента в параграфе.
    """
    # Собираем весь текст параграфа
    full_text = ''.join([run.text for run in paragraph.runs])

    if not full_text.strip():
        return

    # Добавляем пробелы перед знаками процента
    corrected_text = add_space_before_percent(full_text)

    # Если текст изменился, обновляем параграф
    if corrected_text != full_text:
        # Сохраняем форматирование первого run для применения к новому тексту
        first_run_format = {}
        if paragraph.runs:
            first_run = paragraph.runs[0]
            first_run_format = {
                'font_name': first_run.font.name,
                'font_size': first_run.font.size,
                'bold': first_run.bold,
                'italic': first_run.italic,
                'underline': first_run.underline
            }

        # Очищаем параграф и добавляем текст с исправлениями
        paragraph.clear()
        run = paragraph.add_run(corrected_text)

        # Применяем базовое форматирование
        run.font.name = first_run_format.get('font_name', 'Times New Roman')
        run.font.size = first_run_format.get('font_size', Pt(14))
        if first_run_format.get('bold') is not None:
            run.bold = first_run_format.get('bold')
        if first_run_format.get('italic') is not None:
            run.italic = first_run_format.get('italic')
        if first_run_format.get('underline') is not None:
            run.underline = first_run_format.get('underline')


def normalize_stanitsa_abbreviations(text):
    """
    Нормализует сокращения слова "станица" к формату "ст-ца".
    """
    # Словарь возможных сокращений и их замен
    abbreviations = {
        # Сначала более специфичные/длинные, потом общие
        r'\bстани(?:ц|цы|цей|ца|це|цам|цами|цах)\b': 'ст-ца',  # станиц, станицы, станице, станицей и т.д.
        r'\bст(?:\.|\b)': 'ст-ца',  # ст. (точка) или ст (без точки) как отдельное слово
        r'\bста(?:н|н\.)\b': 'ст-ца',  # стан, стан.
        r'\bстц\b': 'ст-ца',  # стц
        r'\bстани\b': 'ст-ца',  # стани
    }

    # Применяем замены
    for pattern, replacement in abbreviations.items():
        # Используем флаг re.IGNORECASE для учета регистра
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)

    return text


def process_stanitsa_abbreviations(doc):
    """
    Обрабатывает нормализацию сокращений "станица" во всем документе.
    """
    try:
        # Обрабатываем параграфы в основном документе
        for paragraph in doc.paragraphs:
            process_paragraph_stanitsa_abbreviations(paragraph)

        # Обрабатываем таблицы
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        process_paragraph_stanitsa_abbreviations(paragraph)

        print("✅ Нормализованы сокращения слова 'станица'")
        return True

    except Exception as e:
        print(f"❌ Ошибка при нормализации сокращений 'станица': {e}")
        return False


def process_paragraph_stanitsa_abbreviations(paragraph):
    """
    Обрабатывает нормализацию сокращений "станица" в параграфе.
    """
    # Собираем весь текст параграфа
    full_text = ''.join([run.text for run in paragraph.runs])

    if not full_text.strip():
        return

    # Нормализуем сокращения
    normalized_text = normalize_stanitsa_abbreviations(full_text)

    # Если текст изменился, обновляем параграф
    if normalized_text != full_text:
        # Сохраняем форматирование первого run для применения к новому тексту
        first_run_format = {}
        if paragraph.runs:
            first_run = paragraph.runs[0]
            first_run_format = {
                'font_name': first_run.font.name,
                'font_size': first_run.font.size,
                'bold': first_run.bold,
                'italic': first_run.italic,
                'underline': first_run.underline
            }

        # Очищаем параграф и добавляем нормализованный текст
        paragraph.clear()
        run = paragraph.add_run(normalized_text)

        # Применяем базовое форматирование
        run.font.name = first_run_format.get('font_name', 'Times New Roman')
        run.font.size = first_run_format.get('font_size', Pt(14))
        if first_run_format.get('bold') is not None:
            run.bold = first_run_format.get('bold')
        if first_run_format.get('italic') is not None:
            run.italic = first_run_format.get('italic')
        if first_run_format.get('underline') is not None:
            run.underline = first_run_format.get('underline')


def normalize_dates_in_text(text):
    """
    Преобразует даты в тексте к формату "12 марта 2024 г."
    """

    # Паттерн для дат в формате ДД.ММ.ГГГГ, ДД/ММ/ГГГГ, ДД-ММ-ГГГГ
    def replace_dd_mm_yyyy(match):
        day, month, year = match.groups()
        try:
            # Надежное преобразование номера месяца в название
            month_key = str(int(month))  # Убираем ведущие нули и преобразуем в строку
            month_name = MONTH_NAMES.get(month_key)
            if not month_name:
                return match.group()
            # Преобразуем 2-значный год в 4-значный
            if len(year) == 2:
                year = f"20{year}" if int(year) < 30 else f"19{year}"
            return f"{int(day)} {month_name} {year} г."
        except (ValueError, KeyError):
            return match.group()  # Возвращаем исходную строку если что-то пошло не так

    # Паттерн для дат в формате ГГГГ.ММ.ДД, ГГГГ/ММ/ДД, ГГГГ-ММ-ДД
    def replace_yyyy_mm_dd(match):
        year, month, day = match.groups()
        try:
            # Надежное преобразование номера месяца в название
            month_key = str(int(month))  # Убираем ведущие нули и преобразуем в строку
            month_name = MONTH_NAMES.get(month_key)
            if not month_name:
                return match.group()
            return f"{int(day)} {month_name} {year} г."
        except (ValueError, KeyError):
            return match.group()

    # Паттерн для дат в формате "12 марта 2024" (без "г.")
    def replace_day_month_year(match):
        day, month, year = match.groups()
        return f"{day} {month} {year} г."

    # Паттерн для дат в формате "12 мар. 2024"
    def replace_day_month_abbr_year(match):
        day, month_abbr, year = match.groups()
        # Преобразуем сокращение месяца в полное название
        month_full = {
            'янв': 'января', 'фев': 'февраля', 'мар': 'марта',
            'апр': 'апреля', 'май': 'мая', 'июн': 'июня',
            'июл': 'июля', 'авг': 'августа', 'сен': 'сентября',
            'окт': 'октября', 'ноя': 'ноября', 'дек': 'декабря'
        }.get(month_abbr.lower(), month_abbr)
        return f"{day} {month_full} {year} г."

    # Паттерн для дат в формате "12 марта 2024 г." (с точкой в "г.")
    def replace_day_month_year_g_dot(match):
        day, month, year = match.groups()
        return f"{day} {month} {year} г."  # Убираем точку

    # Паттерн для дат в формате "12 мар. 2024 г." (с точкой в "г.")
    def replace_day_month_abbr_year_g_dot(match):
        day, month_abbr, year = match.groups()
        # Преобразуем сокращение месяца в полное название
        month_full = {
            'янв': 'января', 'фев': 'февраля', 'мар': 'марта',
            'апр': 'апреля', 'май': 'мая', 'июн': 'июня',
            'июл': 'июля', 'авг': 'августа', 'сен': 'сентября',
            'окт': 'октября', 'ноя': 'ноября', 'дек': 'декабря'
        }.get(month_abbr.lower(), month_abbr)
        return f"{day} {month_full} {year} г."  # Убираем точку

    # Применяем преобразования
    # 1. ДД.ММ.ГГГГ или ДД/ММ/ГГГГ или ДД-ММ-ГГГГ (и ДД.ММ.ГГ)
    text = re.sub(r'\b(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b', replace_dd_mm_yyyy, text)

    # 2. ГГГГ.ММ.ДД или ГГГГ/ММ/ДД или ГГГГ-ММ-ДД
    text = re.sub(r'\b(\d{4})[./-](\d{1,2})[./-](\d{1,2})\b', replace_yyyy_mm_dd, text)

    # 4. "12 мар. 2024" -> "12 марта 2024 г."
    text = re.sub(r'\b(\d{1,2})\s+(янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)[.]\s*(\d{4})\b',
                  replace_day_month_abbr_year, text)

    # 6. "12 мар. 2024 г." (с точкой в "г.") -> "12 марта 2024 г." (без точки)
    text = re.sub(r'\b(\d{1,2})\s+(янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)[.]\s*(\d{4})\s+г\.\b',
                  replace_day_month_abbr_year_g_dot, text)

    # 5. "12 марта 2024 г." (с точкой в "г.") -> "12 марта 2024 г." (без точки)
    text = re.sub(
        r'\b(\d{1,2})\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+(\d{4})\s+г\.\b',
        replace_day_month_year_g_dot, text)

    # 3. "12 марта 2024" (без "г.") -> "12 марта 2024 г."
    # Используем более точный паттерн, чтобы не трогать уже существующие "г."
    text = re.sub(
        r'\b(\d{1,2})\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+(\d{4})\b(?!\s*г)',
        replace_day_month_year, text)

    return text


def normalize_dates(doc):
    """
    Нормализует даты во всем документе
    """
    try:
        # Обрабатываем параграфы в основном документе
        for paragraph in doc.paragraphs:
            normalize_paragraph_dates(paragraph)

        # Обрабатываем таблицы
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        normalize_paragraph_dates(paragraph)

        print("✅ Даты нормализованы")
        return True

    except Exception as e:
        print(f"❌ Ошибка при нормализации дат: {e}")
        return False


def normalize_paragraph_dates(paragraph):
    """
    Нормализует даты в параграфе
    """
    # Собираем весь текст параграфа
    full_text = ''.join([run.text for run in paragraph.runs])

    if not full_text.strip():
        return

    # Нормализуем даты в тексте
    normalized_text = normalize_dates_in_text(full_text)

    # Если текст изменился, обновляем параграф
    if normalized_text != full_text:
        # Сохраняем форматирование первого run для применения к новому тексту
        first_run_format = {}
        if paragraph.runs:
            first_run = paragraph.runs[0]
            first_run_format = {
                'font_name': first_run.font.name,
                'font_size': first_run.font.size,
                'bold': first_run.bold,
                'italic': first_run.italic,
                'underline': first_run.underline
            }

        # Очищаем параграф и добавляем нормализованный текст
        paragraph.clear()
        run = paragraph.add_run(normalized_text)

        # Применяем базовое форматирование
        run.font.name = first_run_format.get('font_name', 'Times New Roman')
        run.font.size = first_run_format.get('font_size', Pt(14))
        if first_run_format.get('bold') is not None:
            run.bold = first_run_format.get('bold')
        if first_run_format.get('italic') is not None:
            run.italic = first_run_format.get('italic')
        if first_run_format.get('underline') is not None:
            run.underline = first_run_format.get('underline')


def convert_decimal_separator_in_text(text):
    """
    Преобразует десятичные разделители в числах с точки на запятую
    """

    def replace_decimal_point(match):
        # Заменяем точку на запятую в найденном числе
        return match.group().replace('.', ',')

    # Паттерн для поиска десятичных чисел с точкой
    # Ищем числа, содержащие точку как десятичный разделитель
    # (точка должна быть между цифрами, а не как разделитель групп разрядов)
    # Исключаем даты, добавив отрицательный просмотр назад и вперед
    pattern = r'(?<!\d[.,])\b\d+\.\d+\b(?![.,]\d)'
    text = re.sub(pattern, replace_decimal_point, text)

    return text


def process_decimal_separators(doc):
    """
    Обрабатывает замену десятичных разделителей во всем документе
    """
    try:
        # Обрабатываем параграфы в основном документе
        for paragraph in doc.paragraphs:
            process_paragraph_decimal_separators(paragraph)

        # Обрабатываем таблицы
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        process_paragraph_decimal_separators(paragraph)

        print("✅ Заменены десятичные разделители (точка → запятая)")
        return True

    except Exception as e:
        print(f"❌ Ошибка при замене десятичных разделителей: {e}")
        return False


def process_paragraph_decimal_separators(paragraph):
    """
    Обрабатывает замену десятичных разделителей в параграфе
    """
    # Собираем весь текст параграфа
    full_text = ''.join([run.text for run in paragraph.runs])

    if not full_text.strip():
        return

    # Заменяем десятичные разделители
    converted_text = convert_decimal_separator_in_text(full_text)

    # Если текст изменился, обновляем параграф
    if converted_text != full_text:
        # Сохраняем форматирование первого run для применения к новому тексту
        first_run_format = {}
        if paragraph.runs:
            first_run = paragraph.runs[0]
            first_run_format = {
                'font_name': first_run.font.name,
                'font_size': first_run.font.size,
                'bold': first_run.bold,
                'italic': first_run.italic,
                'underline': first_run.underline
            }

        # Очищаем параграф и добавляем текст с заменами
        paragraph.clear()
        run = paragraph.add_run(converted_text)

        # Применяем базовое форматирование
        run.font.name = first_run_format.get('font_name', 'Times New Roman')
        run.font.size = first_run_format.get('font_size', Pt(14))
        if first_run_format.get('bold') is not None:
            run.bold = first_run_format.get('bold')
        if first_run_format.get('italic') is not None:
            run.italic = first_run_format.get('italic')
        if first_run_format.get('underline') is not None:
            run.underline = first_run_format.get('underline')


def is_likely_date(text):
    """
    Проверяет, является ли текст похожим на дату
    """
    # Основные паттерны дат
    date_patterns = [
        r'\d{1,2}[./-]\d{1,2}[./-]\d{2,4}',  # 12.03.2024, 12/03/2024, 12-03-2024
        r'\d{4}[./-]\d{1,2}[./-]\d{1,2}',  # 2024.03.12, 2024/03/12
        r'\d{1,2}\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+\d{4}',
        # 12 марта 2024
        r'\d{1,2}\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+\d{4}\s*г\.?',
        # 12 марта 2024 г.
        r'\d{1,2}\s+(янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)[.]\s*\d{4}',  # 12 мар. 2024
    ]

    text_lower = text.lower().strip()
    for pattern in date_patterns:
        if re.search(pattern, text_lower, re.IGNORECASE):
            return True
    return False


def contains_month_word(text):
    """
    Проверяет, содержит ли текст название месяца после числа
    """
    month_patterns = [
        r'\d+\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)',
        r'\d+\s+(янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)[.]'
    ]

    text_lower = text.lower()
    for pattern in month_patterns:
        if re.search(pattern, text_lower):
            return True
    return False


def contains_year_word_to_exclude(text, number_text):
    """
    Проверяет, содержит ли текст слово "год" или "г." непосредственно после указанного числа.
    """
    # Паттерны для слов, после которых число НЕ нужно выделять жирным
    # Ищем конкретное число, за которым следует "год" или "г."
    exclude_year_patterns = [
        rf'\b{re.escape(number_text)}\s+год\b',  # "число год" как отдельное слово
        rf'\b{re.escape(number_text)}\s*г\.',  # "число г." с точкой
    ]

    text_lower = text.lower()
    for pattern in exclude_year_patterns:
        if re.search(pattern, text_lower):
            return True
    return False


def make_numbers_bold(doc):
    """
    Выделяет жирным все числа (кроме дат и чисел с "год" и "г.")
    """
    try:
        # Паттерн для поиска чисел (целые, десятичные, отрицательные, с разделителями)
        number_patterns = [
            r'[+-]?\d{1,3}(?:\s\d{3})*(?:[.,]\d+)?',  # Числа с пробелами как разделителями тысяч: 1 000 000,50
            r'[+-]?\d{1,3}(?:\u2009\d{3})*(?:[.,]\d+)?',  # Числа с тонким пробелом
            r'[+-]?\d{1,3}(?:,\d{3})*(?:[.,]\d+)?',  # Числа с запятыми: 1,000,000.50
            r'[+-]?\d{1,3}(?:\.\d{3})*(?:[.,]\d+)?',  # Числа с точками: 1.000.000,50
            r'[+-]?\d+(?:[.,]\d+)?',  # Простые числа: 123, 12.5, -45
        ]

        # Обрабатываем параграфы в основном документе
        for paragraph in doc.paragraphs:
            process_paragraph_numbers(paragraph, number_patterns)

        # Обрабатываем таблицы
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        process_paragraph_numbers(paragraph, number_patterns)

        print("✅ Числа выделены жирным (даты и числа с 'год'/'г.' исключены)")
        return True

    except Exception as e:
        print(f"❌ Ошибка при выделении чисел: {e}")
        return False


def process_paragraph_numbers(paragraph, number_patterns):
    """
    Обрабатывает числа в параграфе
    """
    # Собираем весь текст параграфа
    runs_text = []
    for run in paragraph.runs:
        runs_text.append(run.text)

    full_text = ''.join(runs_text)

    if not full_text.strip():
        return

    # Находим все потенциальные числа
    numbers_found = []
    for pattern in number_patterns:
        for match in re.finditer(pattern, full_text):
            number_text = match.group()
            start_pos = match.start()
            end_pos = match.end()

            # Проверяем контекст вокруг числа
            context_start = max(0, start_pos - 30)
            context_end = min(len(full_text), end_pos + 30)
            context = full_text[context_start:context_end]

            # Если это не похоже на дату, не содержит название месяца и не содержит "год"/"г." непосредственно после числа
            if (not is_likely_date(context) and
                    not contains_month_word(context) and
                    not contains_year_word_to_exclude(context, number_text)):
                numbers_found.append({
                    'text': number_text,
                    'start': start_pos,
                    'end': end_pos
                })

    # Удаляем дубликаты и сортируем по позиции
    if numbers_found:
        # Удаляем пересекающиеся совпадения, оставляем самые длинные
        numbers_found.sort(key=lambda x: x['start'])
        filtered_numbers = []

        for i, current in enumerate(numbers_found):
            is_valid = True
            # Проверяем пересечение с предыдущими
            for prev in filtered_numbers:
                if (current['start'] < prev['end'] and current['end'] > prev['start']):
                    # Есть пересечение, оставляем более длинное
                    if (current['end'] - current['start']) > (prev['end'] - prev['start']):
                        filtered_numbers.remove(prev)
                    else:
                        is_valid = False
                    break

            if is_valid:
                filtered_numbers.append(current)

        numbers_found = filtered_numbers

    if not numbers_found:
        return

    # Очищаем параграф и пересоздаем с новым форматированием
    paragraph.clear()

    # Обрабатываем текст и добавляем числа жирным
    last_pos = 0
    for number_info in numbers_found:
        # Добавляем текст до числа
        if number_info['start'] > last_pos:
            before_text = full_text[last_pos:number_info['start']]
            run = paragraph.add_run(before_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.bold = False  # Явно сбрасываем жирное форматирование

        # Добавляем число жирным
        number_text = number_info['text']
        run = paragraph.add_run(number_text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)

        last_pos = number_info['end']

    # Добавляем оставшийся текст после последнего числа
    if last_pos < len(full_text):
        remaining_text = full_text[last_pos:]
        run = paragraph.add_run(remaining_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.bold = False  # Явно сбрасываем жирное форматирование


def reset_text_formatting_except_bold(doc):
    """
    Сбрасывает все форматирование текста, кроме жирного выделения
    """
    try:
        # Проходим по всем параграфам документа
        for paragraph in doc.paragraphs:
            # Проходим по всем runs (фрагментам текста с одинаковым форматированием)
            for run in paragraph.runs:
                # Сохраняем только жирное форматирование
                is_bold = run.bold

                # Сбрасываем все форматирование
                run.font.name = None
                run.font.size = None
                run.font.bold = None
                run.font.italic = None
                run.font.underline = None
                run.font.color.rgb = None

                # Восстанавливаем жирное форматирование если оно было
                if is_bold is not None:
                    run.bold = is_bold

        # Проходим по всем таблицам (если есть)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            # Сохраняем только жирное форматирование
                            is_bold = run.bold

                            # Сбрасываем все форматирование
                            run.font.name = None
                            run.font.size = None
                            run.font.bold = None
                            run.font.italic = None
                            run.font.underline = None
                            run.font.color.rgb = None

                            # Восстанавливаем жирное форматирование если оно было
                            if is_bold is not None:
                                run.bold = is_bold

        print("✅ Форматирование текста сброшено (сохранено только жирное выделение)")
        return True

    except Exception as e:
        print(f"❌ Ошибка при сбросе форматирования: {e}")
        return False


def apply_uniform_formatting(doc):
    """
    Применяет统一ный стиль ко всему документу:
    - Шрифт: Times New Roman
    - Размер: 14
    - Междустрочный интервал: 1.5
    """
    try:
        # Проходим по всем параграфам и устанавливаем форматирование
        for paragraph in doc.paragraphs:
            # Устанавливаем шрифт для каждого run
            for run in paragraph.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(14)

            # Устанавливаем междустрочный интервал
            paragraph_format = paragraph.paragraph_format
            paragraph_format.line_spacing = 1.5
            paragraph_format.space_before = Pt(0)
            paragraph_format.space_after = Pt(0)

        # Обрабатываем таблицы
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        # Устанавливаем шрифт для каждого run
                        for run in paragraph.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(14)

                        # Устанавливаем междустрочный интервал
                        paragraph_format = paragraph.paragraph_format
                        paragraph_format.line_spacing = 1.5
                        paragraph_format.space_before = Pt(0)
                        paragraph_format.space_after = Pt(0)

        print("✅ Установлен统一ный стиль: Times New Roman, 14pt, интервал 1.5")
        return True

    except Exception as e:
        print(f"❌ Ошибка при применении统一ного стиля: {e}")
        return False


def set_document_margins(doc_path):
    """
    Устанавливает поля документа Word:
    Верх: 1 см, Право: 1.5 см, Низ: 1 см, Лево: 1.5 см
    """
    try:
        # Проверяем существование файла
        if not os.path.exists(doc_path):
            print(f"❌ Файл не найден: {doc_path}")
            return False

        # Открываем документ
        doc = Document(doc_path)

        # Обрабатываем все секции документа
        for i, section in enumerate(doc.sections):
            print(f"Обрабатываем секцию {i + 1}")
            section.top_margin = Cm(1.0)  # Верх
            section.right_margin = Cm(1.5)  # Право
            section.bottom_margin = Cm(1.0)  # Низ
            section.left_margin = Cm(1.5)  # Лево

        # Сбрасываем форматирование текста (сохраняем только жирное)
        if not reset_text_formatting_except_bold(doc):
            return False

        # Применяем统一ный стиль
        if not apply_uniform_formatting(doc):
            return False

        # Заменяем специальные пробелы на обычные
        if not process_special_spaces(doc):
            return False

        # Заменяем прямые кавычки на типографские
        if not process_quotes(doc):
            return False

        # Нормализуем даты
        if not normalize_dates(doc):
            return False

        # Обрабатываем замену десятичных разделителей
        if not process_decimal_separators(doc):
            return False

        # Добавляем пробелы перед знаками процента
        if not process_percent_signs(doc):
            return False

        # Нормализуем сокращения "станица"
        if not process_stanitsa_abbreviations(doc):
            return False

        # Выделяем числа жирным (кроме дат и чисел с "год"/"г.")
        if not make_numbers_bold(doc):
            return False

        # Устанавливаем выравнивание текста по ширине
        if not set_justify_alignment(doc):
            return False

        # Создаем имя для выходного файла
        name, ext = os.path.splitext(doc_path)
        output_path = f"{name}_formatted{ext}"

        # Сохраняем изменения
        doc.save(output_path)
        print(f"✅ Успешно! Документ сохранён как: {output_path}")
        return True

    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return False


def main():
    print("=== Редактор Word документов ===")
    print("Выполняемые действия:")
    print("1. Установка полей: Верх=1см, Право=1.5см, Низ=1см, Лево=1.5см")
    print("2. Сброс форматирования (сохранено только жирное выделение)")
    print("3. Установка стиля: Times New Roman, 14pt, интервал 1.5")
    print("4. Замена специальных пробелов на обычные и сжатие множественных пробелов")
    print("5. Замена прямых кавычек на типографские")
    print("6. Нормализация дат")
    print("7. Замена десятичных разделителей (точка → запятая)")
    print("8. Добавление пробелов перед знаками процента")
    print("9. Нормализация сокращений 'станица'")
    print("10. Выделение чисел жирным (даты и числа с 'год'/'г.' исключены)")
    print("11. Установка выравнивания текста по ширине")
    print("-" * 65)

    # Получаем путь к файлу
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        print(f"Обрабатываем файл: {file_path}")
    else:
        file_path = input("Введите путь к .docx файлу: ").strip()

    # Убираем кавычки если есть
    file_path = file_path.strip('"\'')

    # Проверяем расширение
    if not file_path.lower().endswith('.docx'):
        print("⚠️  Файл должен иметь расширение .docx")
        return

    # Обрабатываем документ
    success = set_document_margins(file_path)

    if not success:
        print("❌ Обработка завершена с ошибками")
    else:
        print("🎉 Обработка завершена успешно!")


if __name__ == "__main__":
    main()