"""IAM (Maroc Telecom) multi-page invoice parser.

Reads a PDF invoice and returns a structured dictionary containing:
  - Global summary (page 1 of each invoice block)
  - One entry per contract page with its line items and "TOTAL CONTRAT"
  - A computed `total` field = sum of every contract's `total_contrat`
    (i.e. the intersection of all rows except the first / global-summary
    row with the "Total Contrat" column).
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, BinaryIO

import pdfplumber


_AMOUNT = r"-?\d{1,3}(?:[ \u00a0]\d{3})*(?:[.,]\d+)?"
_DATE = r"\d{2}/\d{2}/\d{4}"

RE_INVOICE_NUMBER = re.compile(r"N°\s*Facture\s*:\s*(\S+)")
RE_CLIENT_NUMBER = re.compile(r"N°\s*Client\s*:\s*(\S+)")
RE_INVOICE_DATE = re.compile(r"Date\s*Facture\s*:\s*(" + _DATE + ")")
# Moroccan mobile phone: exactly "0X XX XX XX XX"
RE_PHONE = re.compile(
    r"N°\s*d['\u2019]Appel\s*:\s*(0\d(?:\s\d{2}){4})"
)
RE_PAGE_NUM = re.compile(r"Page\s+(\d+)/(\d+)")
RE_PAGE_HEADER = re.compile(
    r"^\s*(.+?)\s+Page\s+\d+/\d+\s+YAZAKI", re.MULTILINE
)
RE_GLOBAL_HEADER = re.compile(r"Page\s+globale", re.IGNORECASE)
RE_TOTAL_CONTRAT = re.compile(r"TOTAL\s+CONTRAT\s*:\s*(" + _AMOUNT + ")")
# Period is typically: "Période facturée : <start> <end>" possibly across lines
RE_PERIOD = re.compile(
    r"P[ée]riode\s+factur[ée]e\s*:?\s*[\r\n\s]*(" + _DATE + r")\s+(" + _DATE + r")"
)

RE_ABONNEMENT = re.compile(
    r"Frais\s+d['\u2019]abonnement\s+et\s+services\s*:?\s*(" + _AMOUNT + ")"
)
RE_PONCTUEL_GLOBAL = re.compile(
    r"Frais\s+ponctuels\s+li[ée]s\s+au\s+contrat\s*:?\s*(" + _AMOUNT + ")"
)
RE_HT = re.compile(r"Montant\s+HT\s*:?\s*(" + _AMOUNT + ")")
RE_TVA = re.compile(r"Montant\s+TVA[^:\n]*:?\s*(" + _AMOUNT + ")")
RE_TTC = re.compile(r"Montant\s+TTC\s*:?\s*(" + _AMOUNT + ")")
RE_DU = re.compile(r"Montant\s+d[uû]\s*:?\s*(" + _AMOUNT + ")")

RE_MENSUEL_ROW = re.compile(
    r"^(?P<desc>.+?)\s+(?P<start>" + _DATE + r")\s+(?P<end>" + _DATE + r")\s+"
    r"(?P<amount>" + _AMOUNT + r")\s*$"
)
RE_PONCTUEL_ROW = re.compile(r"^(?P<desc>.+?)\s+(?P<amount>" + _AMOUNT + r")\s*$")


def _to_float(value: str) -> float:
    """Convert French-formatted amount like '23 898,78' or '23898.78' to float."""
    cleaned = value.replace("\u00a0", "").replace(" ", "").replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


@dataclass
class LineItem:
    description: str
    amount: float
    date_start: str | None = None
    date_end: str | None = None


@dataclass
class Contract:
    page_number: int          # physical PDF page index (1-based)
    document_page: str | None # logical page like "2/86"
    contract_type: str
    phone_number: str | None = None
    period_start: str | None = None
    period_end: str | None = None
    frais_mensuels: list[LineItem] = field(default_factory=list)
    frais_ponctuels: list[LineItem] = field(default_factory=list)
    total_contrat: float = 0.0


@dataclass
class GlobalSummary:
    page_number: int = 1
    document_page: str | None = None
    frais_abonnement_services: float = 0.0
    frais_ponctuels: float = 0.0
    montant_ht: float = 0.0
    montant_tva: float = 0.0
    montant_ttc: float = 0.0
    montant_du: float = 0.0


@dataclass
class InvoiceData:
    source_file: str
    client_number: str | None = None
    invoice_number: str | None = None
    invoice_date: str | None = None
    period_start: str | None = None
    period_end: str | None = None
    global_summary: GlobalSummary = field(default_factory=GlobalSummary)
    contracts: list[Contract] = field(default_factory=list)
    total: float = 0.0  # sum of every contract's total_contrat

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _identify_page_kind(text: str) -> tuple[str, str | None]:
    """Return ("global", None) | ("contract", contract_type) | ("unknown", None)."""
    if RE_GLOBAL_HEADER.search(text):
        return "global", None
    m = RE_PAGE_HEADER.search(text)
    if m:
        contract_type = m.group(1).strip()
        # Strip leading noise like "DUPLICATA"
        contract_type = re.sub(
            r"^(?:DUPLICATA|ORIGINAL)\s+", "", contract_type, flags=re.IGNORECASE
        ).strip()
        return "contract", contract_type
    if RE_TOTAL_CONTRAT.search(text):
        return "contract", "Inconnu"
    return "unknown", None


def _document_page(text: str) -> str | None:
    if m := RE_PAGE_NUM.search(text):
        return f"{m.group(1)}/{m.group(2)}"
    return None


def _parse_global_page(text: str) -> GlobalSummary:
    g = GlobalSummary()
    g.document_page = _document_page(text)
    if m := RE_ABONNEMENT.search(text):
        g.frais_abonnement_services = _to_float(m.group(1))
    if m := RE_PONCTUEL_GLOBAL.search(text):
        g.frais_ponctuels = _to_float(m.group(1))
    if m := RE_HT.search(text):
        g.montant_ht = _to_float(m.group(1))
    if m := RE_TVA.search(text):
        g.montant_tva = _to_float(m.group(1))
    if m := RE_TTC.search(text):
        g.montant_ttc = _to_float(m.group(1))
    if m := RE_DU.search(text):
        g.montant_du = _to_float(m.group(1))
    return g


def _parse_contract_page(text: str, page_no: int, contract_type: str) -> Contract:
    contract = Contract(
        page_number=page_no,
        document_page=_document_page(text),
        contract_type=contract_type,
    )

    if m := RE_PHONE.search(text):
        contract.phone_number = m.group(1).strip()
    if m := RE_PERIOD.search(text):
        contract.period_start = m.group(1)
        contract.period_end = m.group(2)
    if m := RE_TOTAL_CONTRAT.search(text):
        contract.total_contrat = _to_float(m.group(1))

    section = None  # "mensuel" | "ponctuel"
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        upper = line.upper()
        if upper.startswith("FRAIS MENSUELS"):
            section = "mensuel"
            continue
        if upper.startswith("FRAIS PONCTUELS"):
            section = "ponctuel"
            continue
        if upper.startswith("TOTAL CONTRAT"):
            section = None
            continue
        if section == "mensuel":
            if m := RE_MENSUEL_ROW.match(line):
                contract.frais_mensuels.append(
                    LineItem(
                        description=m.group("desc").strip(),
                        date_start=m.group("start"),
                        date_end=m.group("end"),
                        amount=_to_float(m.group("amount")),
                    )
                )
        elif section == "ponctuel":
            if m := RE_PONCTUEL_ROW.match(line):
                desc = m.group("desc").strip()
                # Skip false matches that look like rows but aren't
                if desc.upper().startswith(("FRAIS", "TOTAL", "N°", "DATE")):
                    continue
                contract.frais_ponctuels.append(
                    LineItem(
                        description=desc,
                        amount=_to_float(m.group("amount")),
                    )
                )
    return contract


def parse_invoice(
    pdf_source: str | Path | BinaryIO,
    ocr_engine=None,
    progress_callback=None,
) -> InvoiceData:
    """Parse a multi-page IAM invoice PDF into structured data.

    - ``ocr_engine`` : instance OCREngine pour les PDFs scannés.
    - ``progress_callback(current, total, status)`` : appelé à chaque page.
    """
    source_label = (
        str(pdf_source) if isinstance(pdf_source, (str, Path)) else "<uploaded>"
    )
    invoice = InvoiceData(source_file=source_label)

    if isinstance(pdf_source, (str, Path)):
        pdf_bytes = Path(pdf_source).read_bytes()
    else:
        pdf_bytes = pdf_source.read()
        pdf_source = io.BytesIO(pdf_bytes)

    fitz_doc = None
    if ocr_engine is not None:
        try:
            import fitz
            fitz_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        except Exception:
            fitz_doc = None

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total_pages = len(pdf.pages)

        for index, page in enumerate(pdf.pages, start=1):

            # ── Mise à jour progression ──────────────────────────────────
            if progress_callback:
                mode = "OCR" if (ocr_engine is not None) else "Extraction"
                progress_callback(index, total_pages, f"{mode} page {index}/{total_pages}…")

            text = page.extract_text() or ""

            if not text.strip() and fitz_doc is not None:
                try:
                    fitz_page = fitz_doc[index - 1]
                    text = ocr_engine.ocr_page(fitz_page)
                except Exception:
                    text = ""

            if not text.strip():
                continue

            if invoice.invoice_number is None:
                if m := RE_INVOICE_NUMBER.search(text):
                    invoice.invoice_number = m.group(1)
            if invoice.client_number is None:
                if m := RE_CLIENT_NUMBER.search(text):
                    invoice.client_number = m.group(1)
            if invoice.invoice_date is None:
                if m := RE_INVOICE_DATE.search(text):
                    invoice.invoice_date = m.group(1)
            if invoice.period_start is None:
                if m := RE_PERIOD.search(text):
                    invoice.period_start = m.group(1)
                    invoice.period_end = m.group(2)

            kind, contract_type = _identify_page_kind(text)
            if kind == "global":
                summary = _parse_global_page(text)
                summary.page_number = index
                invoice.global_summary = summary
            elif kind == "contract":
                invoice.contracts.append(
                    _parse_contract_page(text, index, contract_type or "Inconnu")
                )

    if fitz_doc is not None:
        fitz_doc.close()

    if progress_callback:
        progress_callback(total_pages, total_pages, "✅ Terminé !")

    invoice.total = round(sum(c.total_contrat for c in invoice.contracts), 2)
    return invoice


if __name__ == "__main__":  # quick manual test
    import json
    import sys

    if len(sys.argv) < 2:
        print("Usage : python invoice_parser.py <chemin-vers-pdf>")
        sys.exit(1)
    data = parse_invoice(sys.argv[1])
    print(f"N° Facture : {data.invoice_number}")
    print(f"Date       : {data.invoice_date}")
    print(f"Période    : {data.period_start} -> {data.period_end}")
    print(f"Contrats   : {len(data.contracts)}")
    print(f"TOTAL      : {data.total:,.2f} MAD  (somme des TOTAL CONTRAT)")
    print(f"Montant HT global (vérification croisée) : "
          f"{data.global_summary.montant_ht:,.2f} MAD")
