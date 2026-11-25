"""Validation helpers for Spanish invoices: tax IDs (NIF/NIE/CIF), VAT arithmetic and dates."""
from __future__ import annotations

import re
from datetime import date, datetime

DNI_LETTERS = "TRWAGMYFPDXBNJZSQVHLCKE"
CIF_CONTROL_LETTERS = "JABCDEFGHI"
VALID_IVA_RATES = (0.0, 4.0, 10.0, 21.0)


def clean_tax_id(value) -> str:
    """Remove spaces, dots, dashes and slashes and upper-case the value."""
    return re.sub(r"[\s.\-/]", "", value or "").upper()


def _valid_dni(v: str) -> bool:
    m = re.fullmatch(r"(\d{8})([A-Z])", v)
    return bool(m) and DNI_LETTERS[int(m.group(1)) % 23] == m.group(2)


def _valid_nie(v: str) -> bool:
    m = re.fullmatch(r"([XYZ])(\d{7})([A-Z])", v)
    if not m:
        return False
    number = int(str("XYZ".index(m.group(1))) + m.group(2))
    return DNI_LETTERS[number % 23] == m.group(3)


def _valid_cif(v: str) -> bool:
    m = re.fullmatch(r"([ABCDEFGHJKLMNPQRSUVW])(\d{7})([0-9A-J])", v)
    if not m:
        return False
    letter, digits, control = m.groups()
    odd = sum(sum(divmod(int(d) * 2, 10)) for d in digits[0::2])
    even = sum(int(d) for d in digits[1::2])
    check = (10 - (odd + even) % 10) % 10
    expected_digit, expected_letter = str(check), CIF_CONTROL_LETTERS[check]
    if letter in "PQRSNW":  # these entity types must use a control letter
        return control == expected_letter
    if letter in "ABEH":  # these must use a control digit
        return control == expected_digit
    return control in (expected_digit, expected_letter)  # others: either is allowed


def tax_id_status(value) -> str:
    """Return 'missing', 'ok', 'invalid' or 'foreign' (non-Spanish VAT number, not checked)."""
    v = clean_tax_id(value)
    if not v:
        return "missing"
    if re.match(r"^[A-Z]{2}", v):
        if v.startswith("ES"):
            v = v[2:]
        else:
            return "foreign"
    if _valid_dni(v) or _valid_nie(v) or _valid_cif(v):
        return "ok"
    return "invalid"


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def check_totals(inv: dict, line_tol: float = 0.03, total_tol: float = 0.05) -> list[str]:
    """Check VAT lines (base * rate = fee) and that base + VAT - IRPF withholding = total."""
    issues: list[str] = []
    lines = inv.get("desglose_iva") or []
    total = _num(inv.get("total"))
    retention = _num(inv.get("retencion_irpf")) or 0.0
    if total is None:
        issues.append("Falta el total")
    if not lines:
        issues.append("Sin desglose de IVA")
        return issues
    running = 0.0
    for line in lines:
        base, rate, fee = _num(line.get("base")), _num(line.get("tipo")), _num(line.get("cuota"))
        if None in (base, rate, fee):
            issues.append("Línea de IVA incompleta")
            return issues
        if rate not in VALID_IVA_RATES:
            issues.append(f"Tipo de IVA inusual: {rate:g}%")
        if abs(base * rate / 100 - fee) > line_tol:
            issues.append(f"La cuota de IVA al {rate:g}% no cuadra con la base")
        running += base + fee
    if total is not None and abs(running - retention - total) > total_tol:
        issues.append(
            f"Base + IVA - retención ({running - retention:.2f}) no coincide con el total ({total:.2f})"
        )
    return issues


def check_invoice(inv: dict, today: date | None = None) -> list[str]:
    """Return a list of human-readable reasons why this invoice needs review (empty = looks fine)."""
    today = today or date.today()
    issues: list[str] = []
    doc_type = inv.get("tipo_documento") or "factura"
    is_invoice = doc_type == "factura"

    supplier = tax_id_status(inv.get("proveedor_nif"))
    if supplier == "missing" and is_invoice:
        issues.append("Falta NIF del proveedor")
    elif supplier == "invalid":
        issues.append("NIF del proveedor no válido (¿error de lectura?)")

    client = tax_id_status(inv.get("cliente_nif"))
    if client == "missing" and is_invoice:
        issues.append("Falta NIF del cliente")
    elif client == "invalid":
        issues.append("NIF del cliente no válido (¿error de lectura?)")

    if is_invoice and not inv.get("numero_factura"):
        issues.append("Falta número de factura")

    raw_date = inv.get("fecha")
    if not raw_date:
        issues.append("Falta la fecha")
    else:
        try:
            parsed = datetime.strptime(str(raw_date), "%Y-%m-%d").date()
            if parsed > today:
                issues.append("La fecha es futura")
        except ValueError:
            issues.append("Fecha ilegible")

    issues.extend(check_totals(inv))

    if inv.get("confianza") == "baja":
        issues.append("El modelo indica baja confianza en la lectura")
    if doc_type == "otro":
        issues.append("Documento que no parece una factura ni un ticket")
    return issues
