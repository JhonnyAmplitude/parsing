import re
from datetime import datetime
from typing import Any, List, Optional, Tuple

import xlrd

from OperationDTO import OperationDTO
from constants import CURRENCY_DICT, HEADER_VARIATIONS_TRADES
from utils import extract_rows, normalize_str, parse_date

def normalize_currency(value: Any) -> str:
    value = str(value).strip().upper() if value else ""
    return CURRENCY_DICT.get(value, value)

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

def is_ticker_row(row: List[Any]) -> bool:
    return (
        any(isinstance(cell, str) and "isin" in cell.lower() for cell in row) or
        (isinstance(row[0], str) and re.match(r"^[A-Z]{3,}(_)?[A-Z]{3,}", row[0]) and all(not cell for cell in row[1:]))
    )

def map_headers(row: List[Any], variations: dict) -> dict:
    header_map = {}
    for key, keywords in variations.items():
        for idx, cell in enumerate(row):
            cell_str = str(cell or "").strip().lower()
            if any(keyword.lower() in cell_str for keyword in keywords):
                header_map[key] = idx
                break
    return header_map

def safe_get(row: List[Any], idx: Optional[int], default=None):
    if idx is None or not isinstance(idx, int):
        return default
    return row[idx] if idx < len(row) else default

def to_float(value: Any) -> float:
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0

def normalize_str(value: Any) -> str:
    if not value:
        return ""
    return str(value).strip()


def parse_trade_row(
        row: List[Any],
        trade_type: str,
        header_map: dict,
        ticker: str,
        currency_hint: Optional[str],
        isin: Optional[str] = ""
) -> List[OperationDTO]:
    """
    Парсим строку сделки, возвращая список операций (покупка и/или продажа).
    """
    results = []

    buy_qty_idx = header_map.get("buy_quantity")
    sell_qty_idx = header_map.get("sell_quantity")

    # Для валют — используем отдельные поля для цен (позиционное определение)
    if trade_type == "currency":
        buy_price_idx = header_map.get("buy_price")
        sell_price_idx = header_map.get("sell_price")
    else:
        buy_price_idx = header_map.get("price")
        sell_price_idx = None

    # Покупка
    buy_qty = to_float(safe_get(row, buy_qty_idx))
    if buy_qty > 0:
        buy_payment = to_float(safe_get(row, header_map.get("buy_payment")))
        buy_price   = to_float(safe_get(row, buy_price_idx))
        buy_date    = parse_date(safe_get(row, header_map.get("date")))
        buy_time    = parse_time(safe_get(row, header_map.get("time")))
        buy_currency= normalize_currency(safe_get(row, header_map.get("currency"))) if header_map.get("currency") else currency_hint
        buy_comment = normalize_str(safe_get(row, header_map.get("comment"), ""))
        buy_operation_id = str(safe_get(row, header_map.get("operation_id"), "")).strip()
        buy_aci     = to_float(safe_get(row, header_map.get("aci"), 0.0)) if trade_type == "bond" else 0.0

        full_date = f"{buy_date} {buy_time}" if buy_date else ""

        results.append(OperationDTO(
            date=full_date,
            operation_type="buy",
            payment_sum=buy_payment,
            currency=buy_currency,
            ticker=normalize_str(ticker),
            isin=isin if trade_type != "currency" else "",
            price=buy_price,
            quantity=int(buy_qty),
            aci=buy_aci,
            comment=buy_comment,
            operation_id=buy_operation_id,
        ))

    # Продажа
    sell_qty = to_float(safe_get(row, sell_qty_idx))
    if sell_qty > 0:
        sell_payment = to_float(safe_get(row, header_map.get("sell_payment") or header_map.get("sell_revenue")))

        if trade_type == "bond":
            # как было для облигаций
            sell_price_idx = sell_qty_idx + 1 if sell_qty_idx is not None else None
            sell_price = to_float(safe_get(row, sell_price_idx))
            sell_revenue_idx = header_map.get("sell_revenue")
            sell_aci_idx = sell_revenue_idx + 1 if sell_revenue_idx is not None else None
            sell_aci = to_float(safe_get(row, sell_aci_idx))
        else:
            if trade_type == "currency":
                sell_price = to_float(safe_get(row, sell_price_idx))
            elif trade_type == "stock" and sell_qty_idx is not None:
                sell_price_idx = sell_qty_idx + 1
                sell_price = to_float(safe_get(row, sell_price_idx))
            else:
                sell_price = to_float(safe_get(row, header_map.get("price")))
            sell_aci = 0.0

        sell_date    = parse_date(safe_get(row, header_map.get("date")))
        sell_time    = parse_time(safe_get(row, header_map.get("time")))
        sell_currency= normalize_currency(safe_get(row, header_map.get("currency"))) if header_map.get("currency") else currency_hint
        sell_comment = normalize_str(safe_get(row, header_map.get("comment"), ""))
        sell_operation_id = str(safe_get(row, header_map.get("operation_id"), "")).strip()

        full_date = f"{sell_date} {sell_time}" if sell_date else ""

        results.append(OperationDTO(
            date=full_date,
            operation_type="sale",
            payment_sum=sell_payment,
            currency=sell_currency,
            ticker=normalize_str(ticker),
            isin=isin if trade_type != "currency" else "",
            price=sell_price,
            quantity=int(sell_qty),
            aci=sell_aci,
            comment=sell_comment,
            operation_id=sell_operation_id,
        ))

    return results


SECTION_KEYWORDS = {
    "stock": ["акция", "адр"],
    "bond": ["облигация"],
    "currency": ["иностранная валюта"]
}

def detect_trade_type(header_row: list[str]) -> str:
    header_text = " ".join(header_row).lower()
    if "нкд" in header_text:
        return "bond"
    if "адр" in header_text or "акц" in header_text or "stock" in header_text:
        return "stock"
    # по умолчанию — акции
    return "stock"


def is_section_start(row: List[Any], keywords: List[str]) -> bool:
    return any(keyword in str(cell).lower() for cell in row if cell for keyword in keywords)

def is_valid_trade_row(row: List[Any]) -> bool:
    if not row or not row[0]:
        return False
    if isinstance(row[0], str) and row[0].lower().startswith("итого"):
        return False
    return any(isinstance(cell, (int, float)) for cell in row)

def extract_ticker(row: List[Any]) -> Optional[str]:
    if row and isinstance(row[0], str):
        ticker = row[0].strip()
        if re.match(r"^[A-Z]{3,}(_)?[A-Z]{3,}", ticker):
            return ticker
        return ticker.split()[0]
    return ""

def extract_isin(row: List[Any]) -> Optional[str]:
    for i, cell in enumerate(row):
        if isinstance(cell, str) and "isin" in cell.lower():
            match = re.search(r"ISIN[:\s]*([A-Z0-9]{12})", cell)
            if match:
                return match.group(1)
            next_cell = row[i + 1] if i + 1 < len(row) else ""
            if isinstance(next_cell, str) and re.match(r"^[A-Z0-9]{12}$", next_cell.strip()):
                return next_cell.strip()
    return ""

def parse_trades(filepath: str) -> List[OperationDTO]:
    rows = list(extract_rows(filepath))
    result = []
    current_ticker = None
    current_isin = None
    current_currency = None
    current_section = None
    header_map = {}

    parsing_trades = False
    for row in rows:
        row = row[1:]
        joined_row = " ".join(map(str, row)).lower()

        if not parsing_trades:
            if "2.1. сделки:" in joined_row:
                parsing_trades = True
            continue

        if is_ticker_row(row):
            current_ticker = extract_ticker(row)
            current_isin = extract_isin(row)
            continue

        for section, keywords in SECTION_KEYWORDS.items():
            if is_section_start(row, keywords):
                current_section = section
                header_map = {}
                continue

        if is_section_start(row, ["заем", "овернайт", "цб"]):
            break

        if current_section == "currency" and any("сопряж" in str(cell).lower() for cell in row):
            lot_currency = ""
            pair_currency = ""
            for i, cell in enumerate(row):
                if isinstance(cell, str):
                    cl = cell.lower()
                    if "валюта лота" in cl:
                        lot_currency = normalize_currency(safe_get(row, i + 1))
                    elif "сопряж" in cl:
                        pair_currency = normalize_currency(safe_get(row, i + 1))
            if lot_currency and pair_currency:
                current_ticker = f"{lot_currency}{pair_currency}"
                current_currency = pair_currency
            continue

        # определяем заголовки только один раз
        if not header_map and any("дата" in str(cell).lower() for cell in row):
            if current_section == "currency":
                header_text = [str(cell or "").strip().lower() for cell in row]
                try:
                    buy_price_idx = next(i for i, h in enumerate(header_text) if "курс сделки" in h and "покупка" in h)
                    sell_price_idx = next(i for i, h in enumerate(header_text) if "курс сделки" in h and "продажа" in h)
                    header_map = {
                        "buy_price": buy_price_idx,
                        "buy_quantity": buy_price_idx + 1,
                        "buy_payment": buy_price_idx + 2,
                        "sell_price": sell_price_idx,
                        "sell_quantity": sell_price_idx + 1,
                        "sell_payment": sell_price_idx + 2,
                        "date": next(i for i, h in enumerate(header_text) if h.startswith("дата соверш")),
                        "time": next(i for i, h in enumerate(header_text) if h.startswith("время соверш")),
                        "operation_id": 1
                    }
                except StopIteration:
                    continue
            else:
                header_map = map_headers(row, HEADER_VARIATIONS_TRADES.get(current_section, {}))
            continue

        if not header_map or not is_valid_trade_row(row):
            continue

        ops = parse_trade_row(row, current_section, header_map, current_ticker, current_currency, current_isin)
        result.extend(ops)

    return result
