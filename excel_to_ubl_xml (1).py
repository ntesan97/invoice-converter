#!/usr/bin/env python3
"""
Excel → UBL Invoice XML Converter
Reads: General, Edit - Posted Sales Invoice -, Edit - Posted Sales Invoice - 1, Invoicing
Produces: UBL 2.1 Invoice XML (Serbian eFaktura / mfin.gov.rs schema)

Usage:
    python excel_to_ubl_xml.py <input.xlsx> [output.xml]
    python excel_to_ubl_xml.py invoice.xlsx             → invoice.xml
    python excel_to_ubl_xml.py invoice.xlsx result.xml  → result.xml
"""

import sys
import re
from pathlib import Path
from datetime import datetime, date
import pandas as pd
from lxml import etree

# ---------------------------------------------------------------------------
# Namespaces
# ---------------------------------------------------------------------------
NS = {
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "cec": "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "xsd": "http://www.w3.org/2001/XMLSchema",
    "sbt": "http://mfin.gov.rs/srbdt/srbdtext",
    "":    "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2",
}

# ---------------------------------------------------------------------------
# Seller defaults  ← fill in your company's permanent data here
# ---------------------------------------------------------------------------
SELLER = {
    "pib":               "SELLER_PIB",          # 9-digit PIB
    "name":              "Naziv prodavca d.o.o.",
    "street":            "Ulica i broj",
    "city":              "Grad",
    "post_code":         "00000",
    "country":           "RS",
    "mb":                "SELLER_MB",            # 8-digit matični broj
    "email":             "fakture@seller.rs",
    "bank_account":      "XXX-XXXXXXXXXXXXXXXX-XX",
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _fmt_date(val) -> str:
    """Return ISO date string YYYY-MM-DD from various input types."""
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s[:10]  # fallback: first 10 chars


def _str(val) -> str:
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()


def _dec(val, decimals: int = 2) -> str:
    """Format a number with fixed decimal places."""
    try:
        return f"{float(val):.{decimals}f}"
    except (TypeError, ValueError):
        return "0.00"


def _add(parent, tag: str, text: str, **attribs) -> etree.Element:
    """Create a sub-element with the given text and optional attributes."""
    parts = tag.split(":")
    if len(parts) == 2:
        prefix, local = parts
        qname = etree.QName(NS[prefix], local)
    else:
        qname = tag
    el = etree.SubElement(parent, qname, **attribs)
    el.text = text
    return el


def _sub(parent, tag: str) -> etree.Element:
    parts = tag.split(":")
    if len(parts) == 2:
        prefix, local = parts
        qname = etree.QName(NS[prefix], local)
    else:
        qname = tag
    return etree.SubElement(parent, qname)


# ---------------------------------------------------------------------------
# Sheet readers
# ---------------------------------------------------------------------------

def _read_kv(df: pd.DataFrame) -> dict:
    """Read a 2-column key-value sheet, skipping title rows."""
    kv = {}
    for _, row in df.iterrows():
        k = _str(row.iloc[0])
        v = row.iloc[1] if len(row) > 1 else None
        if k and not pd.isna(v) and v is not None:
            kv[k] = v
    return kv


def _read_lines(df: pd.DataFrame) -> list[dict]:
    """Read invoice lines from 'Edit - Posted Sales Invoice -' sheet."""
    # Row 1 (index 1) contains column headers
    headers = [_str(df.iloc[1, c]) for c in range(df.shape[1])]
    lines = []
    for i in range(2, df.shape[0]):
        row = df.iloc[i]
        if _str(row.iloc[0]) == "":
            continue
        record = {headers[c]: row.iloc[c] for c in range(len(headers))}
        lines.append(record)
    return lines


# ---------------------------------------------------------------------------
# XML builder
# ---------------------------------------------------------------------------

def build_xml(xlsx_path: str) -> bytes:
    xl = pd.ExcelFile(xlsx_path)

    # --- Load sheets ---
    gen_df  = pd.read_excel(xl, sheet_name="General",                       header=None)
    inv_df  = pd.read_excel(xl, sheet_name="Edit - Posted Sales Invoice - ", header=None)
    tot_df  = pd.read_excel(xl, sheet_name="Edit - Posted Sales Invoice - 1",header=None)
    inv2_df = pd.read_excel(xl, sheet_name="Invoicing",                      header=None)

    gen  = _read_kv(gen_df)
    tot  = _read_kv(tot_df)
    inv2 = _read_kv(inv2_df)
    lines = _read_lines(inv_df)

    # --- Derived values ---
    invoice_no   = _str(gen.get("No.", ""))
    issue_date   = _fmt_date(gen.get("Document Date"))
    posting_date = _fmt_date(gen.get("Posting Date"))
    vat_date     = _fmt_date(gen.get("VAT Date", gen.get("Posting Date")))
    due_date     = _fmt_date(inv2.get("Due Date"))
    ext_doc_no   = _str(gen.get("External Document No.", ""))

    buyer_name   = _str(gen.get("Sell-to Customer Name", inv2.get("Bill-to Name", "")))
    buyer_pib    = _str(gen.get("Sell-to Customer No.", inv2.get("Bill-to Customer No.", "")))
    buyer_street = _str(gen.get("Sell-to Address", inv2.get("Bill-to Address", "")))
    buyer_city   = _str(gen.get("Sell-to City",    inv2.get("Bill-to City", "")))
    buyer_zip    = _str(gen.get("Sell-to Post Code",inv2.get("Bill-to Post Code", "")))

    discount_total    = float(tot.get("Invoice Discount Amount Excl. VAT", 0) or 0)
    total_excl_vat    = float(tot.get("Total Excl. VAT (RSD)", 0) or 0)
    total_vat         = float(tot.get("Total VAT (RSD)", 0) or 0)
    total_incl_vat    = float(tot.get("Total Incl. VAT (RSD)", 0) or 0)

    # VAT breakdown per rate from lines
    vat_groups: dict[float, dict] = {}
    line_ext_total = 0.0
    for ln in lines:
        disc_pct  = float(ln.get("Line Discount %", 0) or 0)
        qty       = float(ln.get("Quantity", 0) or 0)
        unit_price= float(ln.get("Unit Price Excl. VAT", 0) or 0)
        line_amt  = float(ln.get("Line Amount Excl. VAT", 0) or 0)
        line_ext_total += line_amt

        # Infer VAT rate from posting group name or last column (price incl VAT)
        # Column 14 (index 14) appears to be price incl. VAT
        price_incl = None
        try:
            price_incl = float(ln.get("", None))
        except (TypeError, ValueError):
            pass

        # Derive vat_rate: (incl / excl - 1) * 100, rounded to nearest 5 or 10
        if price_incl and line_amt and line_amt != 0:
            raw_rate = (price_incl / line_amt - 1) * 100
            # round to nearest 5
            vat_rate = round(raw_rate / 5) * 5
        else:
            vat_rate = 10.0  # fallback

        vg = vat_groups.setdefault(vat_rate, {"taxable": 0.0, "tax": 0.0})
        vg["taxable"] += line_amt
        vg["tax"]     += line_amt * (vat_rate / 100)

    # If no groups detected, fall back to single group
    if not vat_groups:
        vat_groups[10.0] = {"taxable": total_excl_vat, "tax": total_vat}

    # ---------------------------------------------------------------------------
    # Build XML tree
    # ---------------------------------------------------------------------------
    nsmap = {
        None:  NS[""],
        "cbc": NS["cbc"],
        "cac": NS["cac"],
        "cec": NS["cec"],
        "xsi": NS["xsi"],
        "xsd": NS["xsd"],
        "sbt": NS["sbt"],
    }
    root = etree.Element(etree.QName(NS[""], "Invoice"), nsmap=nsmap)

    _add(root, "cbc:CustomizationID",
         "urn:cen.eu:en16931:2017#compliant#urn:mfin.gov.rs:srbdt:2022")
    _add(root, "cbc:ID", invoice_no)
    _add(root, "cbc:IssueDate", issue_date)
    if due_date:
        _add(root, "cbc:DueDate", due_date)
    _add(root, "cbc:InvoiceTypeCode", "380")
    if ext_doc_no:
        _add(root, "cbc:Note", ext_doc_no)
    _add(root, "cbc:DocumentCurrencyCode", "RSD")

    # InvoicePeriod
    ip = _sub(root, "cac:InvoicePeriod")
    _add(ip, "cbc:DescriptionCode", "35")

    # --- Supplier ---
    sup = _sub(root, "cac:AccountingSupplierParty")
    sup_party = _sub(sup, "cac:Party")
    ep = _add(sup_party, "cbc:EndpointID", SELLER["pib"])
    ep.set("schemeID", "9948")
    pn = _sub(sup_party, "cac:PartyName")
    _add(pn, "cbc:Name", SELLER["name"])
    pa = _sub(sup_party, "cac:PostalAddress")
    _add(pa, "cbc:StreetName",  SELLER["street"])
    _add(pa, "cbc:CityName",    SELLER["city"])
    _add(pa, "cbc:PostalZone",  SELLER["post_code"])
    co = _sub(pa, "cac:Country")
    _add(co, "cbc:IdentificationCode", SELLER["country"])
    pts = _sub(sup_party, "cac:PartyTaxScheme")
    _add(pts, "cbc:CompanyID", f"RS{SELLER['pib']}")
    ts = _sub(pts, "cac:TaxScheme")
    _add(ts, "cbc:ID", "VAT")
    ple = _sub(sup_party, "cac:PartyLegalEntity")
    _add(ple, "cbc:RegistrationName", SELLER["name"])
    _add(ple, "cbc:CompanyID", SELLER["mb"])
    cnt = _sub(sup_party, "cac:Contact")
    _add(cnt, "cbc:ElectronicMail", SELLER["email"])

    # --- Customer ---
    cust = _sub(root, "cac:AccountingCustomerParty")
    cust_party = _sub(cust, "cac:Party")
    cep = _add(cust_party, "cbc:EndpointID", buyer_pib)
    cep.set("schemeID", "9948")
    cpn = _sub(cust_party, "cac:PartyName")
    _add(cpn, "cbc:Name", buyer_name)
    cpa = _sub(cust_party, "cac:PostalAddress")
    _add(cpa, "cbc:StreetName", buyer_street)
    _add(cpa, "cbc:CityName",   buyer_city)
    if buyer_zip:
        _add(cpa, "cbc:PostalZone", str(buyer_zip))
    cco = _sub(cpa, "cac:Country")
    _add(cco, "cbc:IdentificationCode", "RS")
    cpts = _sub(cust_party, "cac:PartyTaxScheme")
    _add(cpts, "cbc:CompanyID", f"RS{buyer_pib}")
    cts = _sub(cpts, "cac:TaxScheme")
    _add(cts, "cbc:ID", "VAT")
    cple = _sub(cust_party, "cac:PartyLegalEntity")
    _add(cple, "cbc:RegistrationName", buyer_name)

    # --- Delivery ---
    dv = _sub(root, "cac:Delivery")
    _add(dv, "cbc:ActualDeliveryDate", vat_date)

    # --- Payment means ---
    pm = _sub(root, "cac:PaymentMeans")
    _add(pm, "cbc:PaymentMeansCode", "30")
    pfa = _sub(pm, "cac:PayeeFinancialAccount")
    _add(pfa, "cbc:ID", SELLER["bank_account"])

    # --- Document-level discount (if any) ---
    if discount_total and discount_total != 0:
        ac = _sub(root, "cac:AllowanceCharge")
        _add(ac, "cbc:ChargeIndicator", "false")
        disc_amt = _add(ac, "cbc:Amount", _dec(discount_total))
        disc_amt.set("currencyID", "RSD")
        tc = _sub(ac, "cac:TaxCategory")
        _add(tc, "cbc:ID", "S")
        _add(tc, "cbc:Percent", "10")
        tcs = _sub(tc, "cac:TaxScheme")
        _add(tcs, "cbc:ID", "VAT")

    # --- TaxTotal ---
    tt = _sub(root, "cac:TaxTotal")
    ta = _add(tt, "cbc:TaxAmount", _dec(total_vat))
    ta.set("currencyID", "RSD")

    for rate, grp in sorted(vat_groups.items()):
        tst = _sub(tt, "cac:TaxSubtotal")
        taxable_el = _add(tst, "cbc:TaxableAmount", _dec(grp["taxable"]))
        taxable_el.set("currencyID", "RSD")
        tax_el = _add(tst, "cbc:TaxAmount", _dec(grp["tax"]))
        tax_el.set("currencyID", "RSD")
        tc2 = _sub(tst, "cac:TaxCategory")
        _add(tc2, "cbc:ID", "S")
        _add(tc2, "cbc:Percent", str(int(rate)))
        tcs2 = _sub(tc2, "cac:TaxScheme")
        _add(tcs2, "cbc:ID", "VAT")

    # --- LegalMonetaryTotal ---
    lmt = _sub(root, "cac:LegalMonetaryTotal")
    lea = _add(lmt, "cbc:LineExtensionAmount", _dec(line_ext_total))
    lea.set("currencyID", "RSD")
    tea = _add(lmt, "cbc:TaxExclusiveAmount", _dec(total_excl_vat))
    tea.set("currencyID", "RSD")
    tia = _add(lmt, "cbc:TaxInclusiveAmount", _dec(total_incl_vat))
    tia.set("currencyID", "RSD")
    if discount_total and discount_total != 0:
        ata = _add(lmt, "cbc:AllowanceTotalAmount", _dec(discount_total))
        ata.set("currencyID", "RSD")
    pa2 = _add(lmt, "cbc:PrepaidAmount", "0.00")
    pa2.set("currencyID", "RSD")
    pra = _add(lmt, "cbc:PayableRoundingAmount", "0.00")
    pra.set("currencyID", "RSD")
    paya = _add(lmt, "cbc:PayableAmount", _dec(total_incl_vat))
    paya.set("currencyID", "RSD")

    # --- Invoice lines ---
    for idx, ln in enumerate(lines, start=1):
        il = _sub(root, "cac:InvoiceLine")
        _add(il, "cbc:ID", str(idx))

        qty  = _str(ln.get("Quantity", "1"))
        uom  = _str(ln.get("Unit of Measure Code", "XBX"))
        desc = _str(ln.get("Description", ""))
        item_no   = _str(ln.get("No.", ""))
        unit_price= float(ln.get("Unit Price Excl. VAT", 0) or 0)
        line_amt  = float(ln.get("Line Amount Excl. VAT", 0) or 0)
        disc_pct  = float(ln.get("Line Discount %", 0) or 0)

        iq = _add(il, "cbc:InvoicedQuantity", _str(qty))
        iq.set("unitCode", uom)
        lea2 = _add(il, "cbc:LineExtensionAmount", _dec(line_amt))
        lea2.set("currencyID", "RSD")

        # Line discount
        if disc_pct and disc_pct != 0:
            disc_base = unit_price * float(qty)
            disc_amount = disc_base * (disc_pct / 100)
            lac = _sub(il, "cac:AllowanceCharge")
            _add(lac, "cbc:ChargeIndicator", "false")
            _add(lac, "cbc:AllowanceChargeReason", "Popust")
            _add(lac, "cbc:MultiplierFactorNumeric", _dec(disc_pct, 0))
            da = _add(lac, "cbc:Amount", _dec(disc_amount))
            da.set("currencyID", "RSD")
            ba = _add(lac, "cbc:BaseAmount", _dec(disc_base))
            ba.set("currencyID", "RSD")

        # Item
        item = _sub(il, "cac:Item")
        _add(item, "cbc:Name", desc)
        sid = _sub(item, "cac:SellersItemIdentification")
        _add(sid, "cbc:ID", item_no)

        # Determine VAT rate for this line
        price_incl = None
        raw_14 = ln.get("", None)
        try:
            price_incl = float(raw_14)
        except (TypeError, ValueError):
            pass
        if price_incl and line_amt and line_amt != 0:
            raw_rate = (price_incl / line_amt - 1) * 100
            vat_rate = round(raw_rate / 5) * 5
        else:
            vat_rate = 10

        ctc = _sub(item, "cac:ClassifiedTaxCategory")
        _add(ctc, "cbc:ID", "S")
        _add(ctc, "cbc:Percent", str(int(vat_rate)))
        ctcs = _sub(ctc, "cac:TaxScheme")
        _add(ctcs, "cbc:ID", "VAT")

        # Price (gross unit price before discount)
        price_elem = _sub(il, "cac:Price")
        gross_unit = unit_price / (1 - disc_pct / 100) if disc_pct and disc_pct != 100 else unit_price
        pa3 = _add(price_elem, "cbc:PriceAmount", _dec(gross_unit))
        pa3.set("currencyID", "RSD")

    return etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    xlsx_path = sys.argv[1]
    xml_path  = sys.argv[2] if len(sys.argv) > 2 else Path(xlsx_path).with_suffix(".xml")

    xml_bytes = build_xml(xlsx_path)
    with open(xml_path, "wb") as f:
        f.write(xml_bytes)
    print(f"✓  Written: {xml_path}")


if __name__ == "__main__":
    main()
