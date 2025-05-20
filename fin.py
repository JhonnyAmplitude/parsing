import re
from datetime import datetime
from typing import Any, List, Optional, Tuple, Dict

import xlrd

from OperationDTO import OperationDTO
from constants import CURRENCY_DICT, HEADER_VARIATIONS_TRADES
from utils import extract_rows, normalize_str, parse_date,  safe_float


def parse_time(value: Any) -> str:
    if isinstance(value, datetime):
        return value.strftime("%H:%M:%S")
    if isinstance(value, (int, float)):
        try:
            return datetime(*xlrd.xldate_as_tuple(value, 0)).strftime("%H:%M:%S")
        except Exception:
            return "00:00:00"
    if isinstance(value, str):
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(value.strip(), fmt).strftime("%H:%M:%S")
            except ValueError:
                continue
    return "00:00:00"

def normalize_currency(value: Any) -> str:
    value = str(value).strip().upper() if value else ""
    return CURRENCY_DICT.get(value, value)

def build_trade_col_map(header_row: List[Any], trade_type: str) -> Dict[str, List[int]]:
    """
    Построение словаря field -> список индексов на основе HEADER_VARIATIONS_TRADES
    """
    variations = HEADER_VARIATIONS_TRADES.get(trade_type, {})
    col_map: Dict[str, List[int]] = {}
    for idx, cell in enumerate(header_row):
        text = str(cell or "").strip().lower()
        for key, variants in variations.items():
            if any(v in text for v in variants):
                col_map.setdefault(key, []).append(idx)
                break
    print(col_map)
    return col_map


def parse_trade_row(
    row: List[Any],
    trade_type: str,
    ticker: str,
    currency_hint: Optional[str],
    col_idx: Dict[str, List[int]],
    isin: Optional[str] = ""
) -> OperationDTO:
    """
    Динамический разбор строки сделки по map col_idx, поддерживает повторяющиеся названия колонок.
    """
    def safe_index(key: str, default: int = -1, pos: int = 0) -> int:
        if key not in col_idx:
            return default
        idx = col_idx[key]
        if isinstance(idx, list):
            if len(idx) == 1:
                return idx[0]
            if pos < len(idx):
                return idx[pos]
            return idx[0] if idx else default
        return idx

    buy_qty_idx = safe_index('buy_quantity')
    buy_qty = safe_float(row[buy_qty_idx]) if buy_qty_idx >= 0 else 0.0
    is_buy = buy_qty > 0
    op_key = 'buy' if is_buy else 'sell'

    qty_idx = safe_index(f'{op_key}_quantity', pos=0 if is_buy else 1)
    pay_key = f'{op_key}_payment' if is_buy else f'{op_key}_revenue'
    pay_idx = safe_index(pay_key, pos=0 if is_buy else 1)

    # Используем курс сделки как цену
    price_key = f'{op_key}_price'
    price_idx = safe_index(price_key, pos=0 if is_buy else 1)

    aci_idx = safe_index('aci')
    date_idx = safe_index('date')
    time_idx = safe_index('time')
    curr_idx = safe_index('currency')
    comment_idx = safe_index('comment')
    opid_idx = safe_index('operation_id')

    price = safe_float(row[price_idx]) if price_idx >= 0 else 0.0
    quantity = int(safe_float(row[qty_idx])) if qty_idx >= 0 else 0
    payment = safe_float(row[pay_idx]) if pay_idx >= 0 else 0.0

    raw_date = row[date_idx] if date_idx >= 0 and date_idx < len(row) else None
    trade_date = parse_date(raw_date)
    trade_time = parse_time(row[time_idx]) if time_idx >= 0 and time_idx < len(row) else "00:00:00"

    # Явно берём валюту, даже если нет соответствующей колонки
    currency = normalize_currency(row[curr_idx]) if curr_idx >= 0 and curr_idx < len(row) else currency_hint
    comment = normalize_str(row[comment_idx]) if comment_idx >= 0 and comment_idx < len(row) else ""
    aci = safe_float(row[aci_idx]) if aci_idx >= 0 and aci_idx < len(row) else 0.0
    operation_id = str(row[opid_idx]).strip() if opid_idx >= 0 and opid_idx < len(row) else ""

    full_date = f"{trade_date} {trade_time}" if trade_date else ""

    if trade_type == 'currency':
        op_type = 'currency_buy' if is_buy else 'currency_sale'
    elif trade_type in ('stock', 'bond'):
        op_type = 'buy' if is_buy else 'sale'
    else:
        op_type = f"{trade_type}_{op_key}"

    return OperationDTO(
        date=full_date,
        operation_type=op_type,
        payment_sum=payment,
        currency=currency,
        ticker=normalize_str(ticker),
        isin=isin,
        price=price,
        quantity=quantity,
        aci=aci,
        comment=comment,
        operation_id=operation_id,
    )


SECTION_KEYWORDS = {
    'stock': ['акция', 'адр'],
    'bond': ['облигация'],
    'currency': ['иностранная валюта']
}

def parse_trades(filepath: str) -> List[OperationDTO]:
    rows = list(extract_rows(filepath))

    result: List[OperationDTO] = []
    current_ticker = current_isin = current_currency = None
    current_section = None
    parsing_trades = False
    col_idx: Dict[str, List[int]] = {}

    for row in rows:
        row = row[1:]  # Пропускаем первую колонку
        joined_row = ' '.join(map(str, row)).strip().lower()

        # Старт раздела сделок
        if not parsing_trades:
            if '2.1. сделки:' in joined_row:
                parsing_trades = True
            continue

        # Пропуск строк с "итого" или пустых строк
        if 'итого по' in joined_row or not any(cell for cell in row):
            continue

        # Определяем тикер для валютных пар (CNYRUB_TOM, USDRUB_TOM и т.д.)
        if isinstance(row[0], str) and re.match(r'^[A-Z]{3,}RUB_[A-Z]+$', row[0]):
            current_ticker = row[0].strip()
            current_isin = ''
            continue

        # Обработка секции облигаций — тикер и ISIN могут быть в одной строке
        if any(isinstance(cell, str) and 'isin' in cell.lower() for cell in row):
            for i, cell in enumerate(row):
                cell_str = str(cell).strip().upper()
                if cell_str.startswith('ISIN:'):
                    current_isin = cell_str.replace('ISIN:', '').strip()
                elif re.match(r'^RU\d{9}$', cell_str):
                    current_ticker = cell_str
            continue

        # Определение типа секции
        for section, keywords in SECTION_KEYWORDS.items():
            if any(keyword in str(cell).lower() for cell in row for keyword in keywords):
                current_section = section
                col_idx = {}
                break

        # Обработка строки с валютой (только для currency)
        if current_section == 'currency' and not col_idx:
            for i, cell in enumerate(row):
                if isinstance(cell, str):
                    text = cell.lower()
                    if 'валюта лота' in text and i + 1 < len(row):
                        current_currency = str(row[i + 1]).strip()
                    elif 'сопряж' in text and i + 1 < len(row):
                        # при необходимости можно сохранить сопряжённую валюту
                        pass

        # Заголовок таблицы: build map
        if current_section and not col_idx and HEADER_VARIATIONS_TRADES.get(current_section):
            if any(any(v in str(cell).lower() for cell in row) for variants in HEADER_VARIATIONS_TRADES[current_section].values() for v in variants):
                col_idx = build_trade_col_map(row, current_section)
                continue

        # Парсим строки сделок
        if col_idx and any(isinstance(cell, (int, float)) for cell in row):

            try:
                dto = parse_trade_row(
                    row=row,
                    trade_type=current_section,
                    ticker=current_ticker or '',
                    currency_hint=current_currency,
                    col_idx=col_idx,
                    isin=current_isin or ''
                )
                if dto.date and dto.operation_type:
                    result.append(dto)
            except Exception as e:
                print(f"Ошибка при парсинге строки: {row} — {e}")
            continue

    return result



