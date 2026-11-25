"""Plain-assert tests. Run with: python test_validators.py"""
from datetime import date

from validators import check_invoice, tax_id_status

TODAY = date(2026, 11, 25)


def good_invoice(**overrides):
    inv = {
        "tipo_documento": "factura",
        "numero_factura": "2025-0142",
        "fecha": "2025-10-14",
        "proveedor_nombre": "Suministros Ejemplo, S.L.",
        "proveedor_nif": "B12345674",
        "cliente_nombre": "Bar Ejemplo",
        "cliente_nif": "12345678Z",
        "desglose_iva": [{"base": 100.0, "tipo": 21, "cuota": 21.0}],
        "retencion_irpf": None,
        "total": 121.0,
        "confianza": "alta",
    }
    inv.update(overrides)
    return inv


def test_tax_ids():
    assert tax_id_status("12345678Z") == "ok"
    assert tax_id_status("12345678A") == "invalid"
    assert tax_id_status("X1234567L") == "ok"
    assert tax_id_status("X1234567A") == "invalid"
    assert tax_id_status("B12345674") == "ok"
    assert tax_id_status("B12345679") == "invalid"
    assert tax_id_status("ESB12345674") == "ok"
    assert tax_id_status("b-123 456 74") == "ok"
    assert tax_id_status("Q1234567D") == "ok"  # Q requires a control letter
    assert tax_id_status("Q12345674") == "invalid"
    assert tax_id_status("B1234567D") == "invalid"  # B requires a control digit
    assert tax_id_status("IE6388047V") == "foreign"
    assert tax_id_status("") == "missing"
    assert tax_id_status(None) == "missing"


def test_good_invoice_has_no_issues():
    assert check_invoice(good_invoice(), today=TODAY) == []


def test_total_mismatch_is_flagged():
    issues = check_invoice(good_invoice(total=131.0), today=TODAY)
    assert any("no coincide con el total" in i for i in issues)


def test_vat_fee_mismatch_is_flagged():
    inv = good_invoice(desglose_iva=[{"base": 100.0, "tipo": 21, "cuota": 25.0}], total=125.0)
    assert any("no cuadra" in i for i in check_invoice(inv, today=TODAY))


def test_unusual_rate_is_flagged():
    inv = good_invoice(desglose_iva=[{"base": 100.0, "tipo": 19, "cuota": 19.0}], total=119.0)
    assert any("inusual" in i for i in check_invoice(inv, today=TODAY))


def test_irpf_withholding():
    inv = good_invoice(retencion_irpf=15.0, total=106.0)
    assert check_invoice(inv, today=TODAY) == []


def test_multiple_vat_rates():
    inv = good_invoice(
        desglose_iva=[
            {"base": 100.0, "tipo": 21, "cuota": 21.0},
            {"base": 50.0, "tipo": 10, "cuota": 5.0},
        ],
        total=176.0,
    )
    assert check_invoice(inv, today=TODAY) == []


def test_invalid_supplier_nif_flagged():
    issues = check_invoice(good_invoice(proveedor_nif="B12345679"), today=TODAY)
    assert any("NIF del proveedor" in i for i in issues)


def test_ticket_without_nif_is_fine():
    ticket = {
        "tipo_documento": "ticket",
        "numero_factura": None,
        "fecha": "2025-11-02",
        "proveedor_nif": None,
        "cliente_nif": None,
        "desglose_iva": [{"base": 9.09, "tipo": 10, "cuota": 0.91}],
        "total": 10.0,
        "confianza": "alta",
    }
    assert check_invoice(ticket, today=TODAY) == []


def test_dates():
    assert any("futura" in i for i in check_invoice(good_invoice(fecha="2027-01-01"), today=TODAY))
    assert any("ilegible" in i for i in check_invoice(good_invoice(fecha="14/10/2025"), today=TODAY))
    assert any("Falta la fecha" in i for i in check_invoice(good_invoice(fecha=None), today=TODAY))


def test_low_confidence_flagged():
    assert any("confianza" in i for i in check_invoice(good_invoice(confianza="baja"), today=TODAY))


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_") and callable(f)]
    for name, fn in tests:
        fn()
        print(f"ok  {name}")
    print(f"\n{len(tests)} tests passed")
