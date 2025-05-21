from utils import is_nonzero

#  Валидные операции, которые обрабатываются
VALID_OPERATIONS = {
    "Вознаграждение компании",
    "Дивиденды",
    "НДФЛ",
    "Погашение купона",
    "Погашение облигации",
    "Приход ДС",
    "Проценты по займам \"овернайт\"",
    "Проценты по займам \"овернайт ЦБ\"",
    "Частичное погашение облигации",
    "Вывод ДС",
}

#  Операции, которые нужно игнорировать
SKIP_OPERATIONS = {
    "Внебиржевая сделка FX (22*)",
    "Займы \"овернайт\"",
    "НКД от операций",
    "Покупка/Продажа",
    "Покупка/Продажа (репо)",
    "Переводы между площадками",
}

#  Маппинг строковых названий операций на типы
OPERATION_TYPE_MAP = {
    "Дивиденды": "dividend",
    "Погашение купона": "coupon",
    "Погашение облигации": "repayment",
    "Приход ДС": "deposit",
    "Частичное погашение облигации": "amortization",
    "Вывод ДС": "withdrawal",
}

#  Обработка операций, тип которых зависит от контекста (доход/расход)
SPECIAL_OPERATION_HANDLERS = {
    'Проценты по займам "овернайт"': lambda i, e: "other_income" if is_nonzero(i) else "other_expense",
    'Проценты по займам "овернайт ЦБ"': lambda i, e: "other_income" if is_nonzero(i) else "other_expense",
    "Вознаграждение компании": lambda i, e: "commission_refund" if is_nonzero(i) else "commission",
    "НДФЛ": lambda i, e: "refund" if is_nonzero(i) else "withholding",
}

HEADER_VARIATIONS_TRADES = {
    "stock": {
        "operation_id": ["номер", "Номер"],
        "buy_quantity": ["куплено", "количеств"],
        "buy_payment": ["сумма платежа", "платеж"],
        "sell_quantity": ["продано", "количеств"],
        "sell_revenue": ["сумма выручки", "выручка"],
        "price": ["цена"],
        "currency": ["валют"],
        "date": ["дата соверш", "совершена"],
        "time": ["время соверш"],
        "comment": ["примеч", "коммент"]
    },
    "bond": {
        "operation_id": ["номер", "Номер"],
        "buy_quantity": ["куплено", "Куплено, шт"],
        "buy_payment": ["сумма платежа", "Сумма платежа"],
        "sell_quantity": ["продано", "Продано, шт"],
        "sell_revenue": ["сумма выручки", "Сумма выручки"],
        "price": ["цена", "Цена, %"],
        "aci": ["нкд", "НКД Продажи", "НКД Покупки"],
        "currency": ["Валюта"],
        "date": ["дата соверш", "совершена"],
        "time": ["время соверш"],
        "comment": ["примеч", "коммент"]
    },
    "currency": {
        "operation_id": ["номер", "Номер"],
        "buy_price": ["курс сделки (покупка)"],
        "buy_quantity": ["объём в валюте лота (в ед. валюты)"],
        "buy_payment": ["объём в сопряж. валюте (в ед. валюты)"],
        "sell_price": ["курс сделки (продажа)"],
        "sell_quantity": ["объём в валюте лота (в ед. валюты)"],
        "sell_payment": ["объём в сопряж. валюте (в ед. валюты)"],
        "date": ["Дата соверш."],
        "time": ["Время соверш"],
        "type": ["тип сделки"],
        "comment": ["примеч", "коммент", "Место сделки"]
    }
}

CURRENCY_DICT = {
    "AED": "AED", "AMD": "AMD", "BYN": "BYN", "CHF": "CHF", "CNY": "CNY",
    "EUR": "EUR", "GBP": "GBP", "HKD": "HKD", "JPY": "JPY", "KGS": "KGS",
    "KZT": "KZT", "NOK": "NOK", "RUB": "RUB", "РУБЛЬ": "RUB", "Рубль": "RUB",
    "SEK": "SEK", "TJS": "TJS", "TRY": "TRY", "USD": "USD", "UZS": "UZS",
    "XAG": "XAG", "XAU": "XAU", "ZAR": "ZAR"
}
