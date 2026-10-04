# ── Imports ───────────────────────────────────────────────────────────────────
import streamlit as st
import tempfile
import os
from pathlib import Path
from datetime import datetime, date
import pandas as pd
from lxml import etree

# =============================================================================
# SELLER DETAILS
# =============================================================================
SELLER = {
    "pib":          "110014338",
    "name":         "SERVIER doo",
    "street":       "Milutina Milankovića 11a",
    "city":         "Novi Beograd",
    "post_code":    "11070",
    "country":      "RS",
    "mb":           "21285293",
    "email":        "fakture@servier.rs",
    "bank_account": "325-950050031087-338",
}

# =============================================================================
# CONVERSION LOGIC
# =============================================================================
NS = {
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "cec": "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "xsd": "http://www.w3.org/2001/XMLSchema",
    "sbt": "http://mfin.gov.rs/srbdt/srbdtext",
    "":    "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2",
}


def _fmt_date(val) -> str:
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
    return s[:10]


def _str(val) -> str:
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    return str(val).strip()


def _dec(val, decimals: int = 2) -> str:
    try:
        return f"{float(val):.{decimals}f}"
    except (TypeError, ValueError):
        return "0.00"


def _strip_rs_prefix(val: str) -> str:
    """Remove 'RS' prefix from PIB/VAT numbers if present, return digits only."""
    s = val.strip()
    if s.upper().startswith("RS"):
        s = s[2:].strip()
    return s


def _pib(val) -> str:
    """Return a clean PIB/VAT string — no RS prefix, no .0 float suffix."""
    s = _str(val)
    if not s:
        return ""
    s = _strip_rs_prefix(s)
    # Excel often reads integer IDs as floats: "100000266.0" → "100000266"
    try:
        s = str(int(float(s)))
    except (ValueError, TypeError):
        pass
    return s


def _safe_float(val) -> float:
    """Convert to float safely, stripping RS prefix and non-numeric chars."""
    try:
        s = _str(val)
        if not s:
            return 0.0
        # Strip RS prefix if present (e.g. "RS100270693" → "100270693")
        s = _strip_rs_prefix(s)
        return float(s)
    except (TypeError, ValueError):
        return 0.0


def _add(parent, tag: str, text: str, **attribs):
    parts = tag.split(":")
    qname = etree.QName(NS[parts[0]], parts[1]) if len(parts) == 2 else tag
    el = etree.SubElement(parent, qname, **attribs)
    el.text = text
    return el


def _sub(parent, tag: str):
    parts = tag.split(":")
    qname = etree.QName(NS[parts[0]], parts[1]) if len(parts) == 2 else tag
    return etree.SubElement(parent, qname)


def _read_kv(df: pd.DataFrame) -> dict:
    kv = {}
    for _, row in df.iterrows():
        k = _str(row.iloc[0])
        v = row.iloc[1] if len(row) > 1 else None
        if k:
            try:
                is_na = pd.isna(v)
            except (TypeError, ValueError):
                is_na = False
            if not is_na and v is not None:
                kv[k] = v
    return kv


def _read_lines(df: pd.DataFrame) -> list:
    headers = [_str(df.iloc[1, c]) for c in range(df.shape[1])]
    lines = []
    for i in range(2, df.shape[0]):
        row = df.iloc[i]
        if _str(row.iloc[0]) == "":
            continue
        lines.append({headers[c]: row.iloc[c] for c in range(len(headers))})
    return lines


# =============================================================================
# UN/ECE RECOMMENDATION 20 UNIT CODE MAPPING
# Maps common Business Central / local codes → valid UN/ECE 20 codes.
# Full list: https://docs.peppol.eu/poacc/billing/3.0/codelist/UNECERec20/
# =============================================================================
_UNECE20 = {
    # pieces / units
    "KOM":  "C62", "PCS":  "C62", "PC":   "C62", "ST":   "C62",
    "KS":   "C62", "JED":  "C62", "PIECE":"C62", "PIECES":"C62",
    "EA":   "C62", "EACH": "C62", "NOS":  "C62", "NO":   "C62",
    "U":    "C62",
    # already valid UN/ECE codes — pass through
    "C62":  "C62", "XBX":  "XBX", "XBO":  "XBO", "XCA":  "XCA",
    "XPK":  "XPK", "XCT":  "XCT",
    # mass
    "KG":   "KGM", "KGM":  "KGM", "G":    "GRM", "GRM":  "GRM",
    "T":    "TNE", "TNE":  "TNE",
    # volume
    "L":    "LTR", "LTR":  "LTR", "ML":   "MLT", "MLT":  "MLT",
    "M3":   "MTQ", "MTQ":  "MTQ",
    # length
    "M":    "MTR", "MTR":  "MTR", "CM":   "CMT", "CMT":  "CMT",
    "MM":   "MMT", "MMT":  "MMT",
    # area
    "M2":   "MTK", "MTK":  "MTK",
    # time
    "H":    "HUR", "HUR":  "HUR", "MIN":  "MIN", "D":    "DAY",
    "DAY":  "DAY",
    # packaging
    "BOX":  "XBX", "BX":   "XBX", "CUT":  "XCT", "PAK":  "XPK",
    "PAC":  "XPK", "PACK": "XPK", "KUT":  "XBX",
    # pairs / sets
    "PAR":  "PR",  "PR":   "PR",  "SET":  "SET",
}

def _map_uom(code: str) -> str:
    """Return a valid UN/ECE 20 unit code; fall back to C62 (piece) if unknown."""
    c = code.strip().upper()
    return _UNECE20.get(c, "C62")   # C62 = piece, safe universal fallback


def _first(*vals) -> str:
    """First non-empty value as a clean string."""
    for v in vals:
        s = _str(v)
        if s:
            return s
    return ""


def _read_bill_to(df: pd.DataFrame) -> dict:
    """Bill-to block of the 'Shipping and Billing' sheet (new Navision export).

    That sheet repeats keys (Name/Address/City) for Ship-to and Bill-to, so only
    the rows after the 'Bill-to' marker are read.
    """
    out, in_bill_to = {}, False
    for _, row in df.iterrows():
        k = _str(row.iloc[0])
        if k == "Bill-to":
            in_bill_to = True
            continue
        if in_bill_to and k:
            v = _str(row.iloc[1]) if len(row) > 1 else ""
            if v:
                out[k] = v
    return out


def _line_vat_rate(ln: dict, doc_rate: float) -> float:
    """VAT rate for a line.

    Old export: an unnamed column holds the line amount incl. VAT, so the rate is
    derived per line. New export has no such column → use the document rate.
    """
    price_incl = _safe_float(ln.get("", None)) or None
    line_amt = _safe_float(ln.get("Line Amount Excl. VAT", 0))
    if price_incl and line_amt:
        raw_rate = ((price_incl / line_amt) - 1) * 100
        # Snap to an allowed Serbian rate (10 or 20); SEF rejects 5/15/25 for S.
        return 20.0 if raw_rate >= 15.0 else 10.0
    return doc_rate


def _doc_vat_rate(total_excl: float, total_vat: float, has_line_rates: bool) -> float:
    """Single VAT rate of the document, derived from the Excel totals."""
    if not total_excl:
        return 10.0
    ratio = total_vat / total_excl * 100
    for rate in (10.0, 20.0):
        if abs(ratio - rate) < 0.05:
            return rate
    if has_line_rates:
        return 10.0  # old export: rates come from the lines anyway
    raise ValueError(
        f"Stopa PDV-a se ne može odrediti: ukupan PDV je {ratio:.2f}% osnovice. "
        "Faktura verovatno ima stavke sa različitim stopama (10% i 20%), a novi "
        "Excel export nema stopu po stavci. Ovu fakturu treba uneti ručno ili "
        "dodati kolonu sa stopom PDV-a u export."
    )


def build_xml(xlsx_path: str, warnings: list | None = None) -> bytes:
    """Convert a Navision/BC posted sales invoice export to UBL XML.

    Works with both the old and the new export layout. Non-fatal problems are
    appended to `warnings` (if given).
    """
    if warnings is None:
        warnings = []
    xl = pd.ExcelFile(xlsx_path)

    def sheet(name):
        return pd.read_excel(xl, sheet_name=name, header=None) if name in xl.sheet_names else None

    gen_df   = pd.read_excel(xl, sheet_name="General",                        header=None)
    inv_df   = pd.read_excel(xl, sheet_name="Edit - Posted Sales Invoice - ",  header=None)
    tot_df   = pd.read_excel(xl, sheet_name="Edit - Posted Sales Invoice - 1", header=None)
    inv2_df  = sheet("Invoicing")
    reg_df   = sheet("Registration Numbers")
    sb_df    = sheet("Shipping and Billing")   # new export only

    gen   = _read_kv(gen_df)
    tot   = _read_kv(tot_df)
    inv2  = _read_kv(inv2_df) if inv2_df is not None else {}
    reg   = _read_kv(reg_df) if reg_df is not None else {}
    bill  = _read_bill_to(sb_df) if sb_df is not None else {}
    lines = _read_lines(inv_df)

    invoice_no   = _str(gen.get("No.", ""))
    issue_date   = _fmt_date(gen.get("Document Date"))
    vat_date     = _fmt_date(gen.get("VAT Date", gen.get("Posting Date")))
    # New export: Due Date is on "General"; old export: on "Invoicing"
    due_date     = _fmt_date(gen.get("Due Date", inv2.get("Due Date")))
    ext_doc_no   = _str(gen.get("External Document No.", ""))

    # Buyer: old export uses "Sell-to …" keys, new export "Customer/Address/…"
    # on General, with the Bill-to block of "Shipping and Billing" as fallback.
    buyer_name   = _first(gen.get("Sell-to Customer Name"), gen.get("Customer"),
                          bill.get("Name"), inv2.get("Bill-to Name"))
    buyer_street = _first(gen.get("Sell-to Address"), gen.get("Address"),
                          bill.get("Address"), inv2.get("Bill-to Address"))
    buyer_city   = _first(gen.get("Sell-to City"), gen.get("City"),
                          bill.get("City"), inv2.get("Bill-to City"))
    buyer_zip    = _first(gen.get("Sell-to Post Code"), gen.get("Post Code"),
                          bill.get("Post Code"), inv2.get("Bill-to Post Code"))

    # ── Buyer PIB: read from Registration Numbers sheet ──────────────────────
    # "VAT Registration No." may be stored as "100270693" or "RS100270693"
    raw_buyer_pib = reg.get("VAT Registration No.", gen.get("Sell-to Customer No.", inv2.get("Bill-to Customer No.", "")))
    buyer_pib = _pib(raw_buyer_pib)   # pure digits, no RS prefix, no .0

    missing = [label for label, v in (("broj fakture", invoice_no), ("datum fakture", issue_date),
                                      ("naziv kupca", buyer_name), ("PIB kupca", buyer_pib)) if not v]
    if missing:
        raise ValueError("U Excelu nedostaje: " + ", ".join(missing))
    if not due_date:
        warnings.append("Datum dospeća nije pronađen u Excelu — XML je bez DueDate.")

    discount_total = _safe_float(tot.get("Invoice Discount Amount Excl. VAT", 0))
    total_excl_vat = _safe_float(tot.get("Total Excl. VAT (RSD)", 0))
    total_vat      = _safe_float(tot.get("Total VAT (RSD)", 0))
    total_incl_vat = _safe_float(tot.get("Total Incl. VAT (RSD)", 0))
    excel_incl_vat = total_incl_vat

    has_line_rates = any(_safe_float(ln.get("", None)) for ln in lines)
    doc_rate = _doc_vat_rate(total_excl_vat, total_vat, has_line_rates)

    vat_groups = {}
    line_ext_total = 0.0
    for ln in lines:
        line_amt = _safe_float(ln.get("Line Amount Excl. VAT", 0))
        line_ext_total += line_amt
        vat_rate = _line_vat_rate(ln, doc_rate)
        vg = vat_groups.setdefault(vat_rate, {"taxable": 0.0, "tax": 0.0})
        vg["taxable"] += line_amt
        vg["tax"]     += line_amt * (vat_rate / 100)

    if not vat_groups:
        vat_groups[doc_rate] = {"taxable": total_excl_vat, "tax": total_vat}

    # ── BR-S-08 / BR-CO-13 fix ────────────────────────────────────────────────
    # TaxSubtotal taxable amounts (BT-116) must equal line sums MINUS document
    # discount for that VAT rate. Distribute discount proportionally so that:
    #   BT-109 (TaxExclusiveAmount) = Σ BT-131 - BT-107  (BR-CO-13)
    #   BT-116 per rate = Σ BT-131 for that rate - portion of BT-107 (BR-S-08)
    if discount_total and line_ext_total:
        for rate in vat_groups:
            share = vat_groups[rate]["taxable"] / line_ext_total
            reduction = discount_total * share
            vat_groups[rate]["taxable"] -= reduction
            vat_groups[rate]["tax"] = vat_groups[rate]["taxable"] * (rate / 100)
        # Recompute header totals from the adjusted figures for full consistency
        total_excl_vat = line_ext_total - discount_total
        total_vat      = sum(g["tax"] for g in vat_groups.values())
        total_incl_vat = total_excl_vat + total_vat

    nsmap = {None: NS[""], "cbc": NS["cbc"], "cac": NS["cac"],
             "cec": NS["cec"], "xsi": NS["xsi"], "xsd": NS["xsd"], "sbt": NS["sbt"]}
    root = etree.Element(etree.QName(NS[""], "Invoice"), nsmap=nsmap)

    _add(root, "cbc:CustomizationID", "urn:cen.eu:en16931:2017#compliant#urn:mfin.gov.rs:srbdt:2022")
    _add(root, "cbc:ID", invoice_no)
    _add(root, "cbc:IssueDate", issue_date)
    if due_date:
        _add(root, "cbc:DueDate", due_date)
    _add(root, "cbc:InvoiceTypeCode", "380")
    if ext_doc_no:
        _add(root, "cbc:Note", ext_doc_no)
    _add(root, "cbc:DocumentCurrencyCode", "RSD")
    ip = _sub(root, "cac:InvoicePeriod")
    _add(ip, "cbc:DescriptionCode", "35")

    # ── Supplier ──────────────────────────────────────────────────────────────
    sup_party = _sub(_sub(root, "cac:AccountingSupplierParty"), "cac:Party")
    _add(sup_party, "cbc:EndpointID", SELLER["pib"]).set("schemeID", "9948")
    _add(_sub(sup_party, "cac:PartyName"), "cbc:Name", SELLER["name"])
    pa = _sub(sup_party, "cac:PostalAddress")
    _add(pa, "cbc:StreetName", SELLER["street"])
    _add(pa, "cbc:CityName",   SELLER["city"])
    _add(pa, "cbc:PostalZone", SELLER["post_code"])
    _add(_sub(pa, "cac:Country"), "cbc:IdentificationCode", SELLER["country"])
    pts = _sub(sup_party, "cac:PartyTaxScheme")
    _add(pts, "cbc:CompanyID", f"RS{SELLER['pib']}")   # always RS + 9 digits
    _add(_sub(pts, "cac:TaxScheme"), "cbc:ID", "VAT")
    ple = _sub(sup_party, "cac:PartyLegalEntity")
    _add(ple, "cbc:RegistrationName", SELLER["name"])
    _add(ple, "cbc:CompanyID", SELLER["mb"])
    _add(_sub(sup_party, "cac:Contact"), "cbc:ElectronicMail", SELLER["email"])

    # ── Customer ──────────────────────────────────────────────────────────────
    cust_party = _sub(_sub(root, "cac:AccountingCustomerParty"), "cac:Party")
    _add(cust_party, "cbc:EndpointID", buyer_pib).set("schemeID", "9948")
    _add(_sub(cust_party, "cac:PartyName"), "cbc:Name", buyer_name)
    cpa = _sub(cust_party, "cac:PostalAddress")
    _add(cpa, "cbc:StreetName", buyer_street)
    _add(cpa, "cbc:CityName",   buyer_city)
    if buyer_zip:
        _add(cpa, "cbc:PostalZone", buyer_zip)
    _add(_sub(cpa, "cac:Country"), "cbc:IdentificationCode", "RS")
    cpts = _sub(cust_party, "cac:PartyTaxScheme")
    _add(cpts, "cbc:CompanyID", f"RS{buyer_pib}")       # always RS + 9 digits
    _add(_sub(cpts, "cac:TaxScheme"), "cbc:ID", "VAT")
    _add(_sub(cust_party, "cac:PartyLegalEntity"), "cbc:RegistrationName", buyer_name)

    # ── Delivery & payment ────────────────────────────────────────────────────
    _add(_sub(root, "cac:Delivery"), "cbc:ActualDeliveryDate", vat_date)
    pm = _sub(root, "cac:PaymentMeans")
    _add(pm, "cbc:PaymentMeansCode", "30")
    _add(_sub(pm, "cac:PayeeFinancialAccount"), "cbc:ID", SELLER["bank_account"])

    # ── Document-level discount ───────────────────────────────────────────────
    if discount_total:
        ac = _sub(root, "cac:AllowanceCharge")
        _add(ac, "cbc:ChargeIndicator", "false")
        _add(ac, "cbc:Amount", _dec(discount_total)).set("currencyID", "RSD")
        tc = _sub(ac, "cac:TaxCategory")
        _add(tc, "cbc:ID", "S"); _add(tc, "cbc:Percent", str(int(doc_rate)))
        _add(_sub(tc, "cac:TaxScheme"), "cbc:ID", "VAT")

    # ── TaxTotal ──────────────────────────────────────────────────────────────
    # BR-CO-14: header TaxAmount (BT-110) MUST equal sum of subtotal TaxAmounts
    # (BT-117). Always derive it from vat_groups so they are guaranteed equal.
    tt = _sub(root, "cac:TaxTotal")
    tax_total_computed = sum(g["tax"] for g in vat_groups.values())
    _add(tt, "cbc:TaxAmount", _dec(tax_total_computed)).set("currencyID", "RSD")
    for rate, grp in sorted(vat_groups.items()):
        tst = _sub(tt, "cac:TaxSubtotal")
        _add(tst, "cbc:TaxableAmount", _dec(grp["taxable"])).set("currencyID", "RSD")
        _add(tst, "cbc:TaxAmount",     _dec(grp["tax"])).set("currencyID", "RSD")
        tc2 = _sub(tst, "cac:TaxCategory")
        _add(tc2, "cbc:ID", "S"); _add(tc2, "cbc:Percent", str(int(rate)))
        _add(_sub(tc2, "cac:TaxScheme"), "cbc:ID", "VAT")

    # ── LegalMonetaryTotal ────────────────────────────────────────────────────
    # BR-CO-15: TaxInclusiveAmount (BT-112) = TaxExclusiveAmount (BT-109)
    #           + TaxAmount (BT-110). Always derive from computed values so
    #           all three are guaranteed consistent regardless of sheet values.
    tax_incl_computed = total_excl_vat + tax_total_computed
    lmt = _sub(root, "cac:LegalMonetaryTotal")
    _add(lmt, "cbc:LineExtensionAmount", _dec(line_ext_total)).set("currencyID", "RSD")
    _add(lmt, "cbc:TaxExclusiveAmount",  _dec(total_excl_vat)).set("currencyID", "RSD")
    _add(lmt, "cbc:TaxInclusiveAmount",  _dec(tax_incl_computed)).set("currencyID", "RSD")
    if discount_total:
        _add(lmt, "cbc:AllowanceTotalAmount", _dec(discount_total)).set("currencyID", "RSD")
    _add(lmt, "cbc:PrepaidAmount",        "0.00").set("currencyID", "RSD")
    _add(lmt, "cbc:PayableRoundingAmount","0.00").set("currencyID", "RSD")
    _add(lmt, "cbc:PayableAmount",        _dec(tax_incl_computed)).set("currencyID", "RSD")

    # ── Invoice lines ─────────────────────────────────────────────────────────
    for idx, ln in enumerate(lines, start=1):
        qty        = _str(ln.get("Quantity", "1"))
        uom        = _map_uom(_str(ln.get("Unit of Measure Code", "C62")))
        desc       = _str(ln.get("Description", ""))
        item_no    = _str(ln.get("No.", ""))
        unit_price = _safe_float(ln.get("Unit Price Excl. VAT", 0))
        line_amt   = _safe_float(ln.get("Line Amount Excl. VAT", 0))
        disc_pct   = _safe_float(ln.get("Line Discount %", 0))

        qty_f = _safe_float(qty)

        # Unit price before discount (BT-146).
        # New export: "Unit Price Excl. VAT" is already the price BEFORE the line
        #   discount (qty × price × (1 − d%) = line amount).
        # Old export: it was the price AFTER discount, so gross it back up.
        if disc_pct and disc_pct != 100 and qty_f:
            net_if_gross = round(qty_f * unit_price * (1 - disc_pct / 100), 2)
            if abs(net_if_gross - line_amt) <= 0.05:
                gross_unit = unit_price
            else:
                gross_unit = unit_price / (1 - disc_pct / 100)
        else:
            gross_unit = unit_price

        il = _sub(root, "cac:InvoiceLine")
        _add(il, "cbc:ID", str(idx))
        _add(il, "cbc:InvoicedQuantity", qty).set("unitCode", uom)
        _add(il, "cbc:LineExtensionAmount", _dec(line_amt)).set("currencyID", "RSD")

        if disc_pct:
            # BR-LIN-04: line amount = qty × price − allowance. Derive the
            # allowance from the Excel line amount so the equation holds exactly.
            disc_base = round(gross_unit * qty_f, 2)
            lac = _sub(il, "cac:AllowanceCharge")
            _add(lac, "cbc:ChargeIndicator", "false")
            _add(lac, "cbc:AllowanceChargeReason", "Popust")
            _add(lac, "cbc:MultiplierFactorNumeric", _dec(disc_pct, 0))
            _add(lac, "cbc:Amount",     _dec(disc_base - line_amt)).set("currencyID", "RSD")
            _add(lac, "cbc:BaseAmount", _dec(disc_base)).set("currencyID", "RSD")

        item = _sub(il, "cac:Item")
        _add(item, "cbc:Name", desc)
        _add(_sub(item, "cac:SellersItemIdentification"), "cbc:ID", item_no)

        vat_rate = _line_vat_rate(ln, doc_rate)

        ctc = _sub(item, "cac:ClassifiedTaxCategory")
        _add(ctc, "cbc:ID", "S"); _add(ctc, "cbc:Percent", str(int(vat_rate)))
        _add(_sub(ctc, "cac:TaxScheme"), "cbc:ID", "VAT")

        _add(_sub(il, "cac:Price"), "cbc:PriceAmount", _dec(gross_unit)).set("currencyID", "RSD")

    # ── Cross-check against the Excel totals ─────────────────────────────────
    if excel_incl_vat and abs(tax_incl_computed - excel_incl_vat) > 0.05:
        warnings.append(
            f"Ukupno sa PDV-om u XML-u ({tax_incl_computed:,.2f}) se razlikuje od "
            f"Excela ({excel_incl_vat:,.2f}). Proveri fakturu pre slanja."
        )

    return etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")


# =============================================================================
# STREAMLIT UI
# =============================================================================
st.set_page_config(
    page_title="Excel → UBL XML",
    page_icon="📄",
    layout="centered",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

.stApp { background-color: #0f0f0f; color: #e8e8e8; }

#MainMenu, footer, header { visibility: hidden; }

.block-container { max-width: 640px; padding-top: 4rem; padding-bottom: 4rem; }

.app-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.1rem; font-weight: 600;
    letter-spacing: 0.15em; text-transform: uppercase;
    color: #c8f55a; margin-bottom: 0.25rem;
}
.app-subtitle {
    font-size: 0.85rem; color: #555;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 0.05em; margin-bottom: 3rem;
}

[data-testid="stFileUploader"] {
    background: #1a1a1a; border: 1.5px dashed #2e2e2e;
    border-radius: 4px; padding: 1rem; transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover { border-color: #c8f55a; }
[data-testid="stFileUploader"] label {
    color: #888 !important;
    font-family: 'IBM Plex Mono', monospace; font-size: 0.82rem;
}

.stButton > button {
    background: #c8f55a; color: #0f0f0f;
    font-family: 'IBM Plex Mono', monospace; font-weight: 600;
    font-size: 0.85rem; letter-spacing: 0.1em; text-transform: uppercase;
    border: none; border-radius: 3px; padding: 0.65rem 2rem;
    width: 100%; margin-top: 1rem; transition: background 0.15s, transform 0.1s;
}
.stButton > button:hover { background: #d4ff66; transform: translateY(-1px); }
.stButton > button:active { transform: translateY(0); }

.stDownloadButton > button {
    background: transparent; color: #c8f55a;
    font-family: 'IBM Plex Mono', monospace; font-weight: 600;
    font-size: 0.85rem; letter-spacing: 0.1em; text-transform: uppercase;
    border: 1.5px solid #c8f55a; border-radius: 3px;
    padding: 0.65rem 2rem; width: 100%; margin-top: 0.5rem;
    transition: all 0.15s;
}
.stDownloadButton > button:hover { background: #c8f55a; color: #0f0f0f; }

.info-row {
    display: flex; justify-content: space-between;
    font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; color: #444;
    margin-top: 3rem; padding-top: 1rem; border-top: 1px solid #1e1e1e;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="app-title">Invoice Converter</div>', unsafe_allow_html=True)
st.markdown('<div class="app-subtitle">xlsx → ubl xml · serbian e-faktura</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Drop your Excel file here or click to browse",
    type=["xlsx"],
    label_visibility="visible",
)

if uploaded:
    st.markdown(f"**`{uploaded.name}`** — ready to convert")

    if st.button("Convert to XML"):
        with st.spinner("Processing..."):
            tmp_path = None
            try:
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                    tmp.write(uploaded.read())
                    tmp_path = tmp.name

                warnings = []
                xml_bytes = build_xml(tmp_path, warnings)
                os.unlink(tmp_path)

                output_name = Path(uploaded.name).stem + ".xml"
                st.success(f"✓ Converted successfully — {len(xml_bytes):,} bytes")
                for w in warnings:
                    st.warning(w)
                st.download_button(
                    label="⬇  Download XML",
                    data=xml_bytes,
                    file_name=output_name,
                    mime="application/xml",
                )

            except Exception as e:
                if tmp_path:
                    try:
                        os.unlink(tmp_path)
                    except Exception:
                        pass
                st.error(f"Conversion failed:\n\n{e}")

st.markdown("""
<div class="info-row">
    <span>Reads: General · Invoice Lines · Totals · Invoicing · Registration Numbers · Shipping and Billing</span>
    <span>Schema: EN 16931 / mfin.gov.rs 2022</span>
</div>
""", unsafe_allow_html=True)
