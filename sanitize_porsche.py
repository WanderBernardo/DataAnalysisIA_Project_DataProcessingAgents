#!/usr/bin/env python3
"""
sanitize_porsche.py

Lê um data_sample.xlsx de vendas Porsche (cru) e aplica as regras definidas em
schema.md para gerar um novo arquivo tratado, com cada coluna sanitizada
inserida logo após a sua coluna fonte.

Uso:
    python sanitize_porsche.py [INPUT_XLSX] [OUTPUT_XLSX]

Sem argumentos, usa os defaults definidos em INPUT_PATH e OUTPUT_PATH abaixo.
"""
from __future__ import annotations

import re
import sys
from copy import copy
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

from openpyxl import load_workbook

INPUT_PATH = "data_sample.xlsx"
OUTPUT_PATH = "data_sample_sanitized.xlsx"

INVALID = "INVALID"

# source column -> (output column, sanitizer name)
SOURCE_TO_OUTPUT = {
    "sale_date": "SaleDateSanitized",
    "porsche_model": "PorscheModelSanitized",
    "model_year": "ModelYearSanitized",
    "sale_price": "SalesPriceSanitized",
    "vehicle_mileage": "VehicleMileageSanitized",
    "payment_method": "PayMethodSanitized",
    "city": "CitySanitized",
    "state": "StateSanitized",
    "delivery_status": "DeliveryStatusSanitized",
}

# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
UNITS = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12,
    "thirteen": 13, "fourteen": 14, "fifteen": 15, "sixteen": 16,
    "seventeen": 17, "eighteen": 18, "nineteen": 19,
}
TENS = {
    "twenty": 20, "thirty": 30, "forty": 40, "fifty": 50, "sixty": 60,
    "seventy": 70, "eighty": 80, "ninety": 90,
}


def _clean(value) -> str:
    return re.sub(r"\s+", " ", str(value)).strip()


def _words(text: str) -> list[str]:
    return [w for w in re.split(r"[\s\-]+", text.lower()) if w and w != "and"]


def words_to_int(tokens: list[str]) -> int | None:
    """Converte 'eighty two thousand' -> 82000. None se não for número por extenso."""
    if not tokens:
        return None
    total = current = 0
    for tok in tokens:
        if tok in UNITS:
            current += UNITS[tok]
        elif tok in TENS:
            current += TENS[tok]
        elif tok == "hundred":
            current = max(current, 1) * 100
        elif tok == "thousand":
            total += max(current, 1) * 1000
            current = 0
        elif tok == "million":
            total += max(current, 1) * 1_000_000
            current = 0
        else:
            return None
    return total + current


def parse_number(text: str) -> Decimal | None:
    """
    Interpreta '1,234.56', '1.234,56', '112.750', '12,5', '9800' em Decimal.
    - Dois separadores diferentes: o último é o decimal.
    - Um único tipo de separador com exatamente 3 dígitos depois -> milhar.
    - Caso contrário -> separador decimal.
    """
    text = text.strip().strip(".,")
    if not re.fullmatch(r"\d[\d.,]*", text):
        return None
    has_dot, has_comma = "." in text, "," in text
    if has_dot and has_comma:
        dec = "." if text.rfind(".") > text.rfind(",") else ","
        thou = "," if dec == "." else "."
        text = text.replace(thou, "").replace(dec, ".")
    elif has_dot or has_comma:
        sep = "." if has_dot else ","
        parts = text.split(sep)
        if len(parts) > 2 or len(parts[-1]) == 3:
            text = "".join(parts)  # separador de milhar
        else:
            text = ".".join(parts)  # separador decimal
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def round_half_up(d: Decimal) -> int:
    return int(d.quantize(Decimal("1"), rounding=ROUND_HALF_UP))


# --------------------------------------------------------------------------
# Dates
# --------------------------------------------------------------------------
MONTHS = {}
for i, name in enumerate(
    ["january", "february", "march", "april", "may", "june", "july",
     "august", "september", "october", "november", "december"], start=1):
    MONTHS[name] = i
    MONTHS[name[:3]] = i
MONTHS["sept"] = 9


def _iso(y: int, m: int, d: int) -> str:
    try:
        return date(y, m, d).isoformat()
    except ValueError:
        return INVALID


def sanitize_date(value) -> str:
    if value is None or _clean(value) == "":
        return INVALID
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    s = _clean(value)

    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T]00:00:00)?", s)
    if m:
        return _iso(int(m[1]), int(m[2]), int(m[3]))

    m = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4}|\d{2})", s)
    if m:  # MM/DD/YYYY, MM/DD/YY, MM-DD-YY
        year = int(m[3])
        if len(m[3]) == 2:
            year += 2000
        return _iso(year, int(m[1]), int(m[2]))

    m = re.fullmatch(r"([A-Za-z]+)\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})", s)
    if m and m[1].lower() in MONTHS:
        return _iso(int(m[3]), MONTHS[m[1].lower()], int(m[2]))

    return INVALID


# --------------------------------------------------------------------------
# Model
# --------------------------------------------------------------------------
CANONICAL_MODELS = [
    "911 Carrera", "911 Carrera S", "911 Carrera GTS", "911 Turbo",
    "911 Turbo S", "911 GT3", "911 GT3 RS", "911 Dakar", "911 Targa 4",
    "911 Targa 4S", "718 Cayman", "718 Cayman S", "718 Cayman GT4 RS",
    "718 Boxster", "718 Boxster GTS", "718 Spyder RS", "Cayenne",
    "Cayenne S", "Cayenne Coupe", "Cayenne E-Hybrid", "Cayenne Turbo",
    "Cayenne Turbo GT", "Macan", "Macan S", "Macan T", "Macan GTS",
    "Macan Electric", "Panamera", "Panamera 4", "Panamera 4S",
    "Panamera Turbo", "Panamera Turbo S", "Panamera 4 E-Hybrid", "Taycan",
    "Taycan 4S", "Taycan GTS", "Taycan Turbo", "Taycan Turbo S",
    "Taycan Cross Turismo",
]
_MODEL_LOOKUP = {re.sub(r"[^a-z0-9]", "", m.lower()): m for m in CANONICAL_MODELS}

# Tokens que mantêm caixa especial ao fazer title-case de modelos desconhecidos
_MODEL_UPPER = {"GT3", "GT4", "RS", "GTS", "GT", "RS", "S", "T"}


def _smart_title(text: str) -> str:
    out = []
    for tok in text.split(" "):
        if tok.upper() in _MODEL_UPPER:
            out.append(tok.upper())
        elif re.fullmatch(r"\d+[a-zA-Z]?", tok):
            out.append(tok.upper())
        else:
            out.append("-".join(p.capitalize() for p in tok.split("-")))
    return " ".join(out)


def sanitize_model(value) -> str:
    s = _clean(value)
    if not s:
        return INVALID
    key = re.sub(r"[^a-z0-9]", "", s.lower())
    return _MODEL_LOOKUP.get(key) or _smart_title(s)


# --------------------------------------------------------------------------
# Model year
# --------------------------------------------------------------------------
def sanitize_year(value) -> int | str:
    if value is None or _clean(value) == "":
        return INVALID
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = _clean(value)
    year = None

    if re.fullmatch(r"\d{4}", s):
        year = int(s)
    elif re.fullmatch(r"(\d{2})[-\s](\d{2})", s):
        year = int(s.replace("-", "").replace(" ", ""))
    else:
        toks = _words(s)
        if "thousand" in toks:
            year = words_to_int(toks)
        elif len(toks) >= 2 and toks[0] in ("nineteen", "twenty"):
            rest = words_to_int(toks[1:])
            if rest is not None and rest < 100:
                year = UNITS.get(toks[0], TENS.get(toks[0])) * 100 + rest

    if year is None or not 1990 <= year <= 2035:
        return INVALID
    return year


# --------------------------------------------------------------------------
# Price
# --------------------------------------------------------------------------
def sanitize_price(value) -> str:
    if value is None or _clean(value) == "":
        return INVALID
    if isinstance(value, (int, float)):
        amount = Decimal(str(value))
    else:
        s = _clean(value).lower()
        s = re.sub(r"\b(usd|us\$|dollars?)\b", " ", s).replace("$", " ")
        s = _clean(s)
        amount = None
        if re.fullmatch(r"[a-z\s\-]+", s):
            n = words_to_int(_words(s))
            amount = Decimal(n) if n is not None else None
        else:
            m = re.fullmatch(r"([\d.,]+)\s*(k)?", s)
            if m:
                amount = parse_number(m[1])
                if amount is not None and m[2]:
                    amount *= 1000
    if amount is None or amount < 0:
        return INVALID
    return f"{amount.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP):.2f}"


# --------------------------------------------------------------------------
# Mileage
# --------------------------------------------------------------------------
KM_TO_MI = Decimal("0.621371")


def sanitize_mileage(value) -> int | str:
    if value is None or _clean(value) == "":
        return INVALID
    if isinstance(value, (int, float)):
        return round_half_up(Decimal(str(value))) if value >= 0 else INVALID

    s = _clean(value).lower()
    is_km = bool(re.search(r"\bkm\b|\bkms\b|kilomet", s))

    num = re.search(r"\d[\d.,]*", s)
    if num:
        amount = parse_number(num[0])
    else:
        text = re.sub(r"\b(miles?|mi|km|kms|kilometers?|kilometres?)\b\.?", " ", s)
        text = _clean(text.replace(":", " "))
        if text in ("new", "new car", "brand new", "zero"):
            amount = Decimal(0)
        else:
            n = words_to_int(_words(text))
            amount = Decimal(n) if n is not None else None
    if amount is None or amount < 0:
        return INVALID
    if is_km:
        amount *= KM_TO_MI
    return round_half_up(amount)


# --------------------------------------------------------------------------
# Payment method
# --------------------------------------------------------------------------
def sanitize_payment(value) -> str:
    s = _clean(value) if value is not None else ""
    if not s:
        return INVALID
    k = re.sub(r"[^a-z]", "", s.lower())
    if k in ("creditcard", "creditcardpayment", "credit"):
        return "Credit Card"
    if k in ("debitcard", "debitcardpayment", "debit"):
        return "Debit Card"
    if k in ("banktransfer", "banktransferpayment"):
        return "Bank Transfer"
    if k in ("wiretransfer", "wire", "bankwire", "wiretransferpayment"):
        return "Wire Transfer"
    if k in ("financing", "financingplan", "finance", "financed"):
        return "Financing"
    if k in ("lease", "leasing", "leaseplan"):
        return "Lease"
    if k in ("cash", "cashpayment"):
        return "Cash"
    if k in ("ach", "achpayment"):
        return "ACH Payment"
    if k in ("crypto", "cryptopayment", "cryptocurrency"):
        return "Crypto Payment"
    return re.sub(r"[_\-]+", " ", s).title()


# --------------------------------------------------------------------------
# City
# --------------------------------------------------------------------------
def sanitize_city(value) -> str:
    s = _clean(value) if value is not None else ""
    if not s:
        return INVALID
    s = re.sub(r"^saint\b", "St.", s, flags=re.I)
    s = re.sub(r"^st\.?\s", "St. ", s, flags=re.I)
    return " ".join("-".join(p.capitalize() if not p.lower().startswith("st.") else "St."
                             for p in w.split("-")) for w in s.split(" "))


# --------------------------------------------------------------------------
# State
# --------------------------------------------------------------------------
US_STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT",
    "delaware": "DE", "district of columbia": "DC", "florida": "FL",
    "georgia": "GA", "hawaii": "HI", "idaho": "ID", "illinois": "IL",
    "indiana": "IN", "iowa": "IA", "kansas": "KS", "kentucky": "KY",
    "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN",
    "mississippi": "MS", "missouri": "MO", "montana": "MT",
    "nebraska": "NE", "nevada": "NV", "new hampshire": "NH",
    "new jersey": "NJ", "new mexico": "NM", "new york": "NY",
    "north carolina": "NC", "north dakota": "ND", "ohio": "OH",
    "oklahoma": "OK", "oregon": "OR", "pennsylvania": "PA",
    "rhode island": "RI", "south carolina": "SC", "south dakota": "SD",
    "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT",
    "virginia": "VA", "washington": "WA", "west virginia": "WV",
    "wisconsin": "WI", "wyoming": "WY",
}
_STATE_CODES = set(US_STATES.values())


def sanitize_state(value) -> str:
    s = _clean(value).strip(".").lower() if value is not None else ""
    if not s:
        return INVALID
    if s.upper() in _STATE_CODES:
        return s.upper()
    return US_STATES.get(s, INVALID)


# --------------------------------------------------------------------------
# Delivery status
# --------------------------------------------------------------------------
def sanitize_delivery(value) -> str:
    s = _clean(value) if value is not None else ""
    if not s:
        return INVALID
    k = re.sub(r"[^a-z]", "", s.lower())
    table = {
        "delivered": "Delivered", "deliverd": "Delivered",
        "pending": "Pending",
        "intransit": "In Transit",
        "cancelled": "Cancelled", "canceled": "Cancelled",
        "awaitingdelivery": "Awaiting Delivery",
        "awaitingpickup": "Awaiting Pickup",
        "pendingapproval": "Pending Approval",
        "pendingreview": "Pending Review",
        "shipped": "Shipped",
        "awaitingreview": "Awaiting Review",
    }
    return table.get(k, re.sub(r"[^A-Za-z ]+", " ", s).strip().title() or INVALID)


SANITIZERS = {
    "sale_date": sanitize_date,
    "porsche_model": sanitize_model,
    "model_year": sanitize_year,
    "sale_price": sanitize_price,
    "vehicle_mileage": sanitize_mileage,
    "payment_method": sanitize_payment,
    "city": sanitize_city,
    "state": sanitize_state,
    "delivery_status": sanitize_delivery,
}


# --------------------------------------------------------------------------
# Pipeline
# --------------------------------------------------------------------------
def sanitize_workbook(input_path: str, output_path: str) -> None:
    wb = load_workbook(input_path)
    ws = wb.active

    headers = [c.value for c in ws[1]]
    missing = [c for c in SOURCE_TO_OUTPUT if c not in headers]
    if missing:
        raise SystemExit(f"Colunas obrigatórias ausentes: {missing}")

    original_rows = [[c.value for c in row] for row in ws.iter_rows(min_row=1)]
    max_row = ws.max_row

    # Insere da direita para a esquerda para não deslocar os índices pendentes.
    for src in sorted(SOURCE_TO_OUTPUT, key=headers.index, reverse=True):
        src_idx = headers.index(src) + 1
        out_idx = src_idx + 1
        ws.insert_cols(out_idx)

        head = ws.cell(row=1, column=out_idx, value=SOURCE_TO_OUTPUT[src])
        src_head = ws.cell(row=1, column=src_idx)
        head.font, head.fill = copy(src_head.font), copy(src_head.fill)
        head.border, head.alignment = copy(src_head.border), copy(src_head.alignment)

        fn = SANITIZERS[src]
        for r in range(2, max_row + 1):
            ws.cell(row=r, column=out_idx, value=fn(ws.cell(row=r, column=src_idx).value))

    # insert_cols não desloca larguras: recalcula pelo conteúdo.
    for col in ws.columns:
        longest = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(longest + 2, 12), 40)
    ws.freeze_panes = "A2"
    wb.save(output_path)
    run_quality_checks(original_rows, output_path)


def run_quality_checks(original_rows, output_path: str) -> None:
    ws = load_workbook(output_path).active
    rows = [[c.value for c in r] for r in ws.iter_rows()]
    head = rows[0]
    problems = []

    # 1) colunas originais preservadas e 2) sanitizadas logo após a fonte
    orig_head = original_rows[0]
    for src in orig_head:
        i = head.index(src)
        for r in range(1, len(rows)):
            if rows[r][i] != original_rows[r][orig_head.index(src)]:
                problems.append(f"valor original alterado: {src} linha {r + 1}")
        if src in SOURCE_TO_OUTPUT and head[i + 1] != SOURCE_TO_OUTPUT[src]:
            problems.append(f"{SOURCE_TO_OUTPUT[src]} não está após {src}")

    col = lambda name: [r[head.index(name)] for r in rows[1:]]

    for v in col("SaleDateSanitized"):
        if v != INVALID and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
            problems.append(f"data fora do formato: {v}")
    for v in col("StateSanitized"):
        if v != INVALID and v not in _STATE_CODES:
            problems.append(f"estado inválido: {v}")
    for v in col("VehicleMileageSanitized"):
        if not (isinstance(v, int) or v == INVALID):
            problems.append(f"milhagem não inteira: {v!r}")
    for v in col("SalesPriceSanitized"):
        if v != INVALID and not re.fullmatch(r"\d+\.\d{2}", v):
            problems.append(f"preço fora do formato: {v}")
    for v in col("ModelYearSanitized"):
        if v != INVALID and not (isinstance(v, int) and 1990 <= v <= 2035):
            problems.append(f"ano inválido: {v!r}")
    for out in SOURCE_TO_OUTPUT.values():
        if any(v is None or v == "" for v in col(out)):
            problems.append(f"valores em branco em {out}")

    n = len(rows) - 1
    print(f"{n} linhas processadas -> {output_path}")
    for out in SOURCE_TO_OUTPUT.values():
        print(f"  {out:<26} INVALID: {sum(v == INVALID for v in col(out))}")
    if problems:
        print("FALHAS NOS QUALITY CHECKS:")
        for p in problems:
            print("  -", p)
        raise SystemExit(1)
    print("Quality checks: OK")


if __name__ == "__main__":
    inp = sys.argv[1] if len(sys.argv) > 1 else INPUT_PATH
    out = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_PATH
    sanitize_workbook(inp, out)
