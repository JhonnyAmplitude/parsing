import json
import re
from datetime import datetime
from typing import Any, Dict, Generator, List, Optional, Tuple
import logging

from OperationDTO import OperationDTO
from constants import (
    CURRENCY_DICT,
    OPERATION_TYPE_MAP,
    SKIP_OPERATIONS,
    SPECIAL_OPERATION_HANDLERS,
    VALID_OPERATIONS,
)
from fin import parse_trades

from utils import (
    parse_date,
    extract_rows,
    build_col_index_map_from_row,
    HEADER_VARIATIONS_FIN_OPS,
    is_nonzero,
)

logger = logging.getLogger(__name__)


def extract_isin(comment: str) -> Optional[str]:
    match = re.search(r'\b[A-Z]{2}[A-Z0-9]{10}\b', comment)
    return match.group(0) if match else ""

def safe_float(value: Any) -> float:
    if value is None:
        return 0.0
    try:
        return float(str(value).replace(',', '.'))
    except (ValueError, TypeError):
        return 0.0

def detect_operation_type(op: str, income: str, expense: str) -> str:
    if not isinstance(op, str):
        return "other"
    if "Покупка" in op:
        return "buy"
    if "Продажа" in op:
        return "sell"
    if op in SPECIAL_OPERATION_HANDLERS:
        return SPECIAL_OPERATION_HANDLERS[op](income, expense)
    return OPERATION_TYPE_MAP.get(op, "other")


def get_cell(row: List[Any], index: int) -> Any:
    if not isinstance(row, list):
        return ""
    if not isinstance(index, int) or index < 0 or index >= len(row):
        return ""
    return row[index]


def process_operation_row(
    data: List[Any],              # = row[1:] из parse_financial_operations
    col_idx: Dict[str, int],      # карта полей → индексы
    current_currency: str,
    ticker: str,
    operation_id: str
) -> Optional[OperationDTO]:
    """
    data        — список ячеек строки, начиная со второго столбца Excel
    col_idx     — {'date': 0, 'operation': 1, 'income': 2, 'expense': 3, 'comment': 4, ...}
    current_currency — валюта из заголовка (например, 'USD', 'RUB')
    """

    # 1) операция
    op_raw = str(data[col_idx["operation"]]).strip()
    if not op_raw or op_raw in SKIP_OPERATIONS:
        return None
    if op_raw not in VALID_OPERATIONS:
        # можно логировать или собирать в unknown_operations
        return None

    # 2) дата
    raw_date = data[col_idx["date"]]
    date = parse_date(raw_date)
    if not date:
        return None

    # 3) суммы
    income  = str(data[col_idx["income"]]).strip()  if "income"  in col_idx else ""
    expense = str(data[col_idx["expense"]]).strip() if "expense" in col_idx else ""
    payment = income if is_nonzero(income) else expense
    payment = safe_float(payment)

    # 4) комментарий + ISIN
    comment = str(data[col_idx["comment"]]).strip() if "comment" in col_idx else ""
    isin_val = extract_isin(comment)

    # 5) тип операции и валюта
    op_type = detect_operation_type(op_raw, income, expense)
    currency = CURRENCY_DICT.get(current_currency, current_currency)

    # 6) собираем DTO
    return OperationDTO(
        date=date,
        operation_type=op_type,
        payment_sum=payment,
        currency=currency,
        ticker=ticker,
        isin=isin_val,
        price=0.0,          # в финансовых операциях цена/кол-во не нужны
        quantity=0,
        aci=0.0,
        comment=comment,
        operation_id=operation_id,
    )


def parse_header_data(row_str: str, header_data: Dict[str, Optional[str]]) -> None:
    if "Генеральное соглашение:" in row_str:
        match = re.search(r"Генеральное соглашение:\s*(\d+)", row_str)
        if match:
            header_data["account_id"] = match.group(1)
        date_match = re.search(r"от\s+(\d{2}\.\d{2}\.\d{4})", row_str)
        if date_match:
            header_data["account_date_start"] = parse_date(date_match.group(1))

    elif "Период:" in row_str and "по" in row_str:
        parts = row_str.split()
        try:
            header_data["date_start"] = parse_date(parts[parts.index("с") + 1])
            header_data["date_end"] = parse_date(parts[parts.index("по") + 1])
        except (ValueError, IndexError, TypeError):
            pass


def is_table_header(row_str: str) -> bool:
    return all(header in row_str for header in ["Дата", "Операция", "Сумма зачисления"])


from utils import (
    parse_date,
    extract_rows,
    build_col_index_map_from_row,
    HEADER_VARIATIONS_FIN_OPS,
    is_nonzero,
)
from constants import CURRENCY_DICT, VALID_OPERATIONS, SKIP_OPERATIONS
from OperationDTO import OperationDTO

from typing import Generator, List, Dict, Optional, Tuple, Any
from OperationDTO import OperationDTO
from constants import CURRENCY_DICT, VALID_OPERATIONS, SKIP_OPERATIONS
from utils import (
    parse_date,
    build_col_index_map_from_row,
    HEADER_VARIATIONS_FIN_OPS,
    is_nonzero,
    safe_float,
)
from final import parse_header_data, detect_operation_type, extract_isin

def parse_financial_operations(
    rows: Generator[List[Any], None, None]
) -> Tuple[Dict[str, Optional[str]], List[OperationDTO]]:
    header_data: Dict[str, Optional[str]] = {
        "account_id": None,
        "account_date_start": None,
        "date_start": None,
        "date_end": None,
        "unknown_operations": []
    }
    operations: List[OperationDTO] = []
    current_currency: Optional[str] = None
    parsing: bool = False
    col_idx: Dict[str, int] = {}

    for row in rows:
        logger.debug(f"row: {row}")
        row_str = " ".join(str(c).strip() for c in row if c).strip()
        if row_str in CURRENCY_DICT:
            current_currency = CURRENCY_DICT[row_str]
            continue

        # 2) Ищем строку-заголовок таблицы
        if not parsing and all(k in row_str.lower() for k in ("дата", "операция", "сумма")):
            # Отсекаем первый служебный столбец
            header_cells = row[1:]
            col_idx = build_col_index_map_from_row(header_cells, HEADER_VARIATIONS_FIN_OPS)
            parsing = True
            continue

        # Собираем метаданные до начала таблицы
        if not parsing:
            parse_header_data(row_str, header_data)
            continue

        # Ждём, пока не построится карта колонок
        if not col_idx:
            continue

        # 3) Разбираем каждую строку таблицы
        data: List[Any] = row[1:]  # смещаемся, чтобы индексы col_idx совпадали
        op_raw = str(data[col_idx["operation"]]).strip()
        if not op_raw or op_raw in SKIP_OPERATIONS:
            continue
        if op_raw not in VALID_OPERATIONS:
            header_data["unknown_operations"].append(op_raw)
            continue

        # Дата
        raw_date = data[col_idx["date"]]
        date = parse_date(raw_date)
        if not date:
            continue

        # Сумма
        income  = str(data[col_idx["income"]]).strip()  if "income"  in col_idx else ""
        expense = str(data[col_idx["expense"]]).strip() if "expense" in col_idx else ""
        payment = safe_float(income if is_nonzero(income) else expense)

        # Комментарий и ISIN
        comment  = str(data[col_idx["comment"]]).strip() if "comment" in col_idx else ""
        isin_val = extract_isin(comment)

        # Тип операции и валюта
        op_type  = detect_operation_type(op_raw, income, expense)
        currency = current_currency or "RUB"

        operations.append(OperationDTO(
            date=date,
            operation_type=op_type,
            payment_sum=payment,
            currency=currency,
            isin=isin_val,
            comment=comment,
            operation_id="",
        ))

    return header_data, operations


def parse_full_statement(file_path: str) -> Dict[str, Any]:
    try:
        rows = list(extract_rows(file_path))
    except Exception as e:
        raise RuntimeError(f"Ошибка при чтении файла {file_path}: {e}")

    if not rows:
        raise ValueError(f"Файл {file_path} пуст или не содержит данных.")

    header_data, financial_operations = parse_financial_operations(iter(rows))
    trade_operations = parse_trades(file_path)
    operations = financial_operations + trade_operations

    operations.sort(key=lambda op: (op._sort_key is None, op._sort_key))

    operations_dict = [
        {k: v for k, v in op.__dict__.items() if not k.startswith("_")}
        for op in operations
    ]

    return {
        "account_id": header_data.get("account_id"),
        "account_date_start": header_data.get("account_date_start"),
        "date_start": header_data.get("date_start"),
        "date_end": header_data.get("date_end"),
        "operations": operations_dict,
    }

