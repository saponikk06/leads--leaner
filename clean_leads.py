#!/usr/bin/env python3
"""
clean_leads.py — приведение "грязной" таблицы заявок к чистому виду.

Использование:
    python clean_leads.py входной_файл.xlsx [--out-clean clean.xlsx] [--out-problems problems.xlsx]

Вход:  Excel/CSV файл с колонками "Имя", "Телефон", "Дата заявки", "Источник"
       (регистр и лишние пробелы в названиях колонок допускаются).
Выход: два файла — чистые уникальные записи и проблемные строки с причиной.
"""

import argparse
import re
import sys
from datetime import datetime

import pandas as pd


# ---------------------------------------------------------------------------
# Телефон
# ---------------------------------------------------------------------------

def normalize_phone(raw):
    """
    Возвращает (номер_в_формате_+7XXXXXXXXXX, ошибка_или_None).
    Правила:
      - оставляем только цифры;
      - 11 цифр, начинается с 8 или 7  -> заменяем первую цифру на 7;
      - 10 цифр (без кода страны)      -> добавляем 7 спереди;
      - иначе -> невалидно ("нет телефона").
    """
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None, "нет телефона"

    digits = re.sub(r"\D", "", str(raw))

    if len(digits) == 11 and digits[0] in ("7", "8"):
        digits = "7" + digits[1:]
    elif len(digits) == 10:
        digits = "7" + digits
    else:
        return None, "нет телефона"

    return "+" + digits, None


# ---------------------------------------------------------------------------
# Дата
# ---------------------------------------------------------------------------

RU_MONTHS = {
    "январ": 1, "феврал": 2, "март": 3, "апрел": 4, "ма": 5, "июн": 6,
    "июл": 7, "август": 8, "сентябр": 9, "октябр": 10, "ноябр": 11, "декабр": 12,
}

DATE_FORMATS = [
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
    "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y",
    "%d.%m.%y", "%d/%m/%y", "%d-%m-%y",
    "%m/%d/%Y",
]


def _parse_russian_text_date(raw):
    """'13 марта 2026' / '04 апреля 2026' -> datetime или None."""
    m = re.match(r"^\s*(\d{1,2})\s+([А-Яа-яё]+)\s+(\d{4})\s*$", raw)
    if not m:
        return None
    day, month_word, year = m.groups()
    month_word_low = month_word.lower()
    month_num = None
    for stem, num in RU_MONTHS.items():
        if month_word_low.startswith(stem):
            month_num = num
            break
    if month_num is None:
        return None
    try:
        return datetime(int(year), month_num, int(day))
    except ValueError:
        return None


def normalize_date(raw):
    """
    Возвращает (дата_ГГГГ-ММ-ДД, ошибка_или_None).
    Строка не выбрасывается при нераспознанной дате — просто помечается.
    """
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None, "нет даты"

    # excel мог отдать pandas.Timestamp / datetime напрямую
    if isinstance(raw, (pd.Timestamp, datetime)):
        return raw.strftime("%Y-%m-%d"), None

    text = str(raw).strip()
    if not text:
        return None, "нет даты"

    # русский текстовый формат: "13 марта 2026"
    dt = _parse_russian_text_date(text)
    if dt:
        return dt.strftime("%Y-%m-%d"), None

    # числовые форматы с одно-/двузначными день/месяц (2026-3-18, 1.4.2026)
    m = re.match(r"^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})$", text)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            return datetime(y, mo, d).strftime("%Y-%m-%d"), None
        except ValueError:
            pass

    m = re.match(r"^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})$", text)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return datetime(y, mo, d).strftime("%Y-%m-%d"), None
        except ValueError:
            pass

    # день месяц год через пробелы: "23 03 2026"
    m = re.match(r"^(\d{1,2})\s+(\d{1,2})\s+(\d{4})$", text)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return datetime(y, mo, d).strftime("%Y-%m-%d"), None
        except ValueError:
            pass

    # стандартные форматы через strptime
    for fmt in DATE_FORMATS:
        try:
            dt = datetime.strptime(text, fmt)
            # для двузначного года считаем, что это 20XX
            if dt.year < 100:
                dt = dt.replace(year=dt.year + 2000)
            return dt.strftime("%Y-%m-%d"), None
        except ValueError:
            continue

    return None, f"нераспознанная дата: '{text}'"


# ---------------------------------------------------------------------------
# Имя
# ---------------------------------------------------------------------------

def normalize_name(raw):
    """Trim + схлопнуть пробелы + Title Case (с учётом дефисов). Пусто -> (None, ошибка)."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None, "нет имени"

    text = re.sub(r"\s+", " ", str(raw).strip())
    if not text:
        return None, "нет имени"

    def cap_word(w):
        return "-".join(part[:1].upper() + part[1:].lower() if part else part
                         for part in w.split("-"))

    normalized = " ".join(cap_word(w) for w in text.split(" "))
    return normalized, None


# ---------------------------------------------------------------------------
# Основная логика
# ---------------------------------------------------------------------------

def find_column(columns, target_lower_variants):
    for col in columns:
        if str(col).strip().lower() in target_lower_variants:
            return col
    return None


def load_table(path):
    if path.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(path)
    elif path.lower().endswith((".csv", ".txt")):
        # пробуем автоматически определить разделитель
        return pd.read_csv(path, sep=None, engine="python")
    else:
        raise ValueError(f"Неподдерживаемый формат файла: {path}")


def save_table(df, path):
    if path.lower().endswith((".xlsx", ".xls")):
        df.to_excel(path, index=False)
    else:
        df.to_csv(path, index=False)


def main():
    parser = argparse.ArgumentParser(description="Очистка файла с заявками")
    parser.add_argument("input", help="Путь к входному файлу (xlsx/csv/txt)")
    parser.add_argument("--out-clean", default="clean_output.xlsx",
                         help="Файл для чистых записей (по умолчанию clean_output.xlsx)")
    parser.add_argument("--out-problems", default="problem_rows.xlsx",
                         help="Файл для проблемных строк (по умолчанию problem_rows.xlsx)")
    args = parser.parse_args()

    df = load_table(args.input)

    name_col = find_column(df.columns, {"имя"})
    phone_col = find_column(df.columns, {"телефон"})
    date_col = find_column(df.columns, {"дата заявки", "дата"})
    source_col = find_column(df.columns, {"источник"})

    missing = [n for n, c in
               [("Имя", name_col), ("Телефон", phone_col),
                ("Дата заявки", date_col), ("Источник", source_col)]
               if c is None]
    if missing:
        print(f"ОШИБКА: не найдены обязательные колонки: {', '.join(missing)}", file=sys.stderr)
        sys.exit(1)

    clean_rows = []
    problem_rows = []

    for idx, row in df.iterrows():
        original_row_number = idx + 2  # +2: заголовок + 1-based

        name, name_err = normalize_name(row[name_col])
        phone, phone_err = normalize_phone(row[phone_col])
        date, date_err = normalize_date(row[date_col])
        source = row[source_col] if pd.notna(row[source_col]) else None

        errors = [e for e in (name_err, phone_err, date_err) if e]

        if errors:
            problem_rows.append({
                "Исходная строка №": original_row_number,
                "Имя (исходное)": row[name_col],
                "Телефон (исходный)": row[phone_col],
                "Дата заявки (исходная)": row[date_col],
                "Источник": source,
                "Причина": "; ".join(errors),
            })
        else:
            clean_rows.append({
                "Имя": name,
                "Телефон": phone,
                "Дата заявки": date,
                "Источник": source,
            })

    clean_df = pd.DataFrame(clean_rows)

    # дедупликация по телефону — оставляем первую встретившуюся запись
    before = len(clean_df)
    if not clean_df.empty:
        clean_df = clean_df.drop_duplicates(subset=["Телефон"], keep="first").reset_index(drop=True)
    removed_dupes = before - len(clean_df)

    problems_df = pd.DataFrame(problem_rows)

    save_table(clean_df, args.out_clean)
    save_table(problems_df, args.out_problems)

    total = len(df)
    print(f"Всего строк на входе: {total}")
    print(f"Чистых уникальных записей: {len(clean_df)}")
    print(f"Удалено дублей по телефону: {removed_dupes}")
    print(f"Проблемных строк: {len(problems_df)}")
    print(f"Проверка целостности (чисто+дубли+проблемы == вход): "
          f"{len(clean_df) + removed_dupes + len(problems_df) == total}")
    print(f"\nСохранено:\n  {args.out_clean}\n  {args.out_problems}")


if __name__ == "__main__":
    main()
