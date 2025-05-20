import os
import xlrd
import openpyxl

from typing import Any, Generator, List, Optional, Dict

from datetime import datetime


HEADER_VARIATIONS_FIN_OPS: Dict[str, list] = {
    "date":      ["дата"],
    "operation": ["операция"],
    "income":    ["сумма зачисления", "зачислено"],
    "expense":   ["сумма списания",   "списано"],
    "comment":   ["примечание",       "назначение", "описание"],
}

def build_col_index_map_from_row(header_row: List[Any], variations: Dict[str, list]) -> Dict[str, int]:
    col_map: Dict[str, int] = {}
    for idx, cell in enumerate(header_row):
        text = str(cell or "").strip().lower()
        for key, variants in variations.items():
            if any(v in text for v in variants):
                col_map[key] = idx
                break
    return col_map

def extract_rows(file_path: str) -> Generator[List[Any], None, None]:
    """
    Чтение строк из файла Excel (форматы .xls или .xlsx).
    Поддерживает как чтение из файлов, так и чтение из байтовых данных.
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".xls":
        sheet = xlrd.open_workbook(file_path).sheet_by_index(0)
        for i in range(sheet.nrows):
            yield sheet.row_values(i)
    elif ext == ".xlsx":
        sheet = openpyxl.load_workbook(file_path, data_only=True).active
        for row in sheet.iter_rows(values_only=True):
            yield list(row)
    else:
        raise ValueError("Неподдерживаемый формат файла")


def parse_date(value: Any) -> Optional[str]:
    """
    Универсальный парсер даты.
    Поддерживает:
    - datetime.datetime
    - Excel float/int дату (как в .xls)
    - Строки в формате 'дд.мм.гггг' или 'дд.мм.гг'
    Возвращает строку в формате 'YYYY-MM-DD' или None.
    """
    if not value:
        return None

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, (int, float)):
        try:
            date = datetime(*xlrd.xldate_as_tuple(value, 0))
            return date.strftime("%Y-%m-%d")
        except Exception:
            return None

    if isinstance(value, str):
        value = value.strip()
        for fmt in ("%d.%m.%Y", "%d.%m.%y"):
            try:
                return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue

    return None

def normalize_str(value: Any) -> str:
    s = str(value).strip() if value is not None else ""
    return "" if s.lower() == "none" else s


def safe_float(value: Any) -> float:
    if value is None:
        return 0.0
    try:
        return float(str(value).replace(',', '.'))
    except (ValueError, TypeError):
        return 0.0

def is_nonzero(value: Any) -> bool:
    """
    Проверка на значение, отличное от нуля.
    """
    try:
        return float(str(value).replace(",", ".").replace(" ", "")) != 0
    except (ValueError, TypeError):
        return False