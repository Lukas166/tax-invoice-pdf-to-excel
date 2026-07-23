from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Callable, Iterable

import fitz


MONTHS = {
    "januari": 1,
    "februari": 2,
    "maret": 3,
    "april": 4,
    "mei": 5,
    "juni": 6,
    "juli": 7,
    "agustus": 8,
    "september": 9,
    "oktober": 10,
    "november": 11,
    "desember": 12,
}


UNIT_ALIASES = {
    "piece": "PCS",
    "pieces": "PCS",
    "pcs": "PCS",
    "pc": "PCS",
    "buah": "PCS",

    "unit": "UNIT",
    "units": "UNIT",

    "set": "SET",
    "sets": "SET",

    "box": "BOX",
    "boxes": "BOX",

    "pack": "PACK",
    "packs": "PACK",

    "kg": "KG",
    "kilogram": "KG",

    "gram": "GRAM",
    "gr": "GRAM",

    "ton": "TON",

    "liter": "LITER",
    "litre": "LITER",
    "ltr": "LITER",
    "l": "LITER",

    "milliliter": "ML",
    "millilitre": "ML",
    "ml": "ML",

    "meter": "METER",
    "metre": "METER",
    "m": "METER",

    "m2": "M2",
    "m²": "M2",

    "m3": "M3",
    "m³": "M3",

    "roll": "ROLL",
    "rol": "ROLL",

    "lembar": "LBR",
    "lbr": "LBR",
    "sheet": "LBR",
    "sheets": "LBR",

    "batang": "BATANG",

    "pasang": "PASANG",
    "pair": "PASANG",

    "lusin": "LUSIN",
    "dozen": "LUSIN",

    "pallet": "PALLET",
    "palet": "PALLET",

    "drum": "DRUM",
    "pail": "PAIL",

    "sack": "KARUNG",
    "karung": "KARUNG",

    "lot": "LOT",
}


AMOUNT = (
    r"-?"
    r"(?:\d{1,3}(?:[.\s]\d{3})*|\d+)"
    r"(?:,\d+)?"
)


class InvoiceParseError(ValueError):
    pass


@dataclass(slots=True)
class InvoiceItem:
    no: int
    code: str
    name: str

    qty: Decimal
    unit: str
    unit_price: Decimal

    line_amount: Decimal
    discount: Decimal = Decimal("0")
    ppnbm_pdf: Decimal = Decimal("0")

    dpp: Decimal = Decimal("0")
    ppn: Decimal = Decimal("0")
    ppnbm: Decimal = Decimal("0")

    @property
    def weight(self) -> Decimal:
        """
        Nilai netto yang dipakai sebagai bobot pembagian
        DPP, PPN, dan PPnBM.
        """
        return max(
            self.line_amount - self.discount,
            Decimal("0"),
        )


@dataclass(slots=True)
class ParsedInvoice:
    source_name: str

    invoice_date: date
    invoice_number: str
    customer_name: str

    items: list[InvoiceItem]

    gross_total: Decimal
    discount_total: Decimal

    dpp_total: Decimal
    ppn_pdf: Decimal
    ppn_output: Decimal
    ppnbm_total: Decimal

    not_collected: bool
    warnings: list[str] = field(default_factory=list)

    def total(
        self,
        include_ppnbm: bool = True,
    ) -> Decimal:
        result = self.dpp_total + self.ppn_output

        if include_ppnbm:
            result += self.ppnbm_total

        return result


@dataclass(slots=True)
class ParseBatchResult:
    invoices: list[ParsedInvoice]
    errors: list[dict[str, str]]


def normalize_spaces(value: str) -> str:
    return re.sub(
        r"\s+",
        " ",
        value,
    ).strip()


def parse_number(value: str) -> Decimal:
    """
    Mendukung angka seperti:

    1.234.567,89
    1234567,89
    1,234,567.89
    1234567.89
    """
    value = re.sub(
        r"(?i)rp",
        "",
        value,
    )

    value = (
        value
        .replace("\u00a0", "")
        .replace(" ", "")
        .strip()
    )

    if not value:
        return Decimal("0")

    if "," in value and "." in value:
        if value.rfind(",") > value.rfind("."):
            value = (
                value
                .replace(".", "")
                .replace(",", ".")
            )
        else:
            value = value.replace(",", "")

    elif "," in value:
        value = (
            value
            .replace(".", "")
            .replace(",", ".")
        )

    elif "." in value:
        parts = value.split(".")

        if (
            len(parts) > 1
            and all(
                len(part) == 3
                for part in parts[1:]
            )
        ):
            value = "".join(parts)

    try:
        return Decimal(value)

    except InvalidOperation as exc:
        raise InvoiceParseError(
            f"Format angka tidak dikenali: {value!r}"
        ) from exc


def round_rupiah(value: Decimal) -> Decimal:
    return value.quantize(
        Decimal("1"),
        rounding=ROUND_HALF_UP,
    )


def allocate(
    total: Decimal,
    weights: Iterable[Decimal],
) -> list[Decimal]:
    """
    Membagi total berdasarkan bobot item.

    Item terakhir menerima selisih pembulatan agar
    jumlah item sama persis dengan total PDF.
    """
    weights = [
        max(
            Decimal("0"),
            Decimal(weight),
        )
        for weight in weights
    ]

    total = round_rupiah(total)

    if not weights:
        return []

    if len(weights) == 1:
        return [total]

    weight_sum = sum(
        weights,
        Decimal("0"),
    )

    if weight_sum == 0:
        result = [
            round_rupiah(
                total / len(weights)
            )
            for _ in weights[:-1]
        ]

    else:
        result = [
            round_rupiah(
                total * weight / weight_sum
            )
            for weight in weights[:-1]
        ]

    result.append(
        total - sum(
            result,
            Decimal("0"),
        )
    )

    return result


def extract_text(pdf_bytes: bytes) -> str:
    try:
        document = fitz.open(
            stream=pdf_bytes,
            filetype="pdf",
        )

        text = "\n".join(
            page.get_text("text")
            for page in document
        )

        document.close()

    except Exception as exc:
        raise InvoiceParseError(
            f"PDF tidak dapat dibuka: {exc}"
        ) from exc

    text = (
        text
        .replace("\u00a0", " ")
        .replace("\r", "")
    )

    if len(normalize_spaces(text)) < 100:
        raise InvoiceParseError(
            "PDF tidak memiliki text layer "
            "dan membutuhkan OCR."
        )

    return text


def required(
    pattern: str,
    text: str,
    label: str,
    flags: int = 0,
) -> re.Match[str]:
    match = re.search(
        pattern,
        text,
        flags,
    )

    if not match:
        raise InvoiceParseError(
            f"Keyword {label!r} tidak ditemukan."
        )

    return match


def summary_amount(
    text: str,
    label_pattern: str,
    label: str,
    optional: bool = False,
) -> Decimal:
    """
    Mengambil nilai ringkasan terakhir.

    Menggunakan match terakhir karena beberapa label,
    seperti Harga Jual, juga muncul pada header tabel.
    """
    pattern = (
        rf"{label_pattern}"
        rf"\s*"
        rf"(?:\n\s*)?"
        rf"({AMOUNT})"
    )

    matches = list(
        re.finditer(
            pattern,
            text,
            re.I,
        )
    )

    if not matches:
        if optional:
            return Decimal("0")

        raise InvoiceParseError(
            f"Nilai {label!r} tidak ditemukan."
        )

    return round_rupiah(
        parse_number(
            matches[-1].group(1)
        )
    )


def invoice_number(text: str) -> str:
    match = required(
        (
            r"Kode\s+dan\s+Nomor\s+Seri\s+"
            r"Faktur\s+Pajak\s*:\s*"
            r"([0-9.\-]+)"
        ),
        text,
        "Kode dan Nomor Seri Faktur Pajak",
        re.I,
    )

    return match.group(1).strip()


def formatted_invoice_number(value: str) -> str:
    """
    Format nomor faktur lama:

    0100012467950711
    menjadi
    010.001-24.67950711

    Nomor 17 digit Coretax tidak diubah.
    """
    digits = re.sub(
        r"\D",
        "",
        value,
    )

    if len(digits) != 16:
        return value

    return (
        f"{digits[:3]}."
        f"{digits[3:6]}-"
        f"{digits[6:8]}."
        f"{digits[8:]}"
    )


def customer_name(text: str) -> str:
    buyer_section = required(
        (
            r"Pembeli\s+Barang\s+Kena\s+Pajak"
            r"\s*/\s*"
            r"Penerima\s+Jasa\s+Kena\s+Pajak"
            r"\s*:"
        ),
        text,
        "bagian pembeli",
        re.I,
    )

    buyer_text = text[buyer_section.end():]

    name_match = required(
        r"(?:^|\n)\s*Nama\s*:\s*([^\n]+)",
        buyer_text,
        "Nama customer",
        re.I,
    )

    return normalize_spaces(
        name_match.group(1)
    ).upper()


def invoice_date(text: str) -> date:
    month_pattern = "|".join(MONTHS)

    matches = list(
        re.finditer(
            (
                rf"\b(\d{{1,2}})"
                rf"\s+({month_pattern})"
                rf"\s+(\d{{4}})\b"
            ),
            text,
            re.I,
        )
    )

    if not matches:
        raise InvoiceParseError(
            "Tanggal tanda tangan tidak ditemukan."
        )

    # Tanggal tanda tangan berada di bagian bawah,
    # sehingga tanggal terakhir yang digunakan.
    match = matches[-1]

    day = int(match.group(1))
    month = MONTHS[match.group(2).lower()]
    year = int(match.group(3))

    return date(
        year,
        month,
        day,
    )


def normalize_unit(value: str) -> str:
    """
    Unit yang dikenal dinormalisasi.

    Unit yang tidak ada di UNIT_ALIASES tidak dianggap error.
    Nilainya tetap digunakan dalam uppercase.
    """
    cleaned = (
        normalize_spaces(value)
        .strip(" .,:;()[]{}")
        .lower()
    )

    if not cleaned:
        raise InvoiceParseError(
            "Satuan barang kosong."
        )

    return UNIT_ALIASES.get(
        cleaned,
        cleaned.upper(),
    )


def extract_items(text: str) -> list[InvoiceItem]:
    table_header = required(
        (
            r"Nama\s+Barang\s+Kena\s+Pajak"
            r"\s*/\s*"
            r"Jasa\s+Kena\s+Pajak"
            r".*?"
            r"\(Rp\)\s*\n"
        ),
        text,
        "tabel barang",
        re.I | re.S,
    )

    table_tail = text[table_header.end():]

    table_end = re.search(
        (
            r"\n\s*"
            r"Harga\s+Jual\s*/\s*Penggantian"
            r"\s*/\s*Uang\s+Muka\s*/\s*Termin"
            r"\s*\n"
        ),
        table_tail,
        re.I,
    )

    if not table_end:
        raise InvoiceParseError(
            "Batas akhir tabel barang tidak ditemukan."
        )

    item_section = table_tail[
        :table_end.start()
    ].strip()

    item_blocks = list(
        re.finditer(
            (
                r"(?ms)"
                r"^\s*(?P<no>\d+)\s*\n"
                r"\s*(?P<code>[A-Z0-9./\-]+)\s*\n"
                r"(?P<body>.*?)"
                r"(?="
                r"^\s*\d+\s*\n"
                r"\s*[A-Z0-9./\-]+\s*\n"
                r"|\Z"
                r")"
            ),
            item_section,
            re.I,
        )
    )

    if not item_blocks:
        raise InvoiceParseError(
            "Tidak ada item yang berhasil diparse."
        )

    price_pattern = re.compile(
        (
            rf"(?im)"
            rf"^\s*Rp\s*({AMOUNT})"
            rf"\s*[xX×]\s*"
            rf"({AMOUNT})"
            rf"\s+([^\n]+?)\s*$"
        )
    )

    discount_pattern = re.compile(
        (
            rf"Potongan\s+Harga"
            rf"\s*=\s*"
            rf"Rp\s*({AMOUNT})"
        ),
        re.I,
    )

    ppnbm_pattern = re.compile(
        (
            rf"PPnBM\s*"
            rf"\([^)]*\)"
            rf"\s*=\s*"
            rf"Rp\s*({AMOUNT})"
        ),
        re.I,
    )

    standalone_amount_pattern = re.compile(
        rf"(?m)^\s*({AMOUNT})\s*$"
    )

    items: list[InvoiceItem] = []

    for block in item_blocks:
        item_no = int(
            block.group("no")
        )

        body = block.group(
            "body"
        ).strip()

        price_match = price_pattern.search(body)

        if not price_match:
            raise InvoiceParseError(
                "Baris harga x qty tidak ditemukan "
                f"pada item {item_no}."
            )

        # Semua teks sebelum baris Rp harga x qty satuan
        # menjadi JENIS BARANG.
        description = normalize_spaces(
            body[:price_match.start()]
        ).upper()

        if not description:
            raise InvoiceParseError(
                f"Jenis barang kosong pada item {item_no}."
            )

        unit_price = parse_number(
            price_match.group(1)
        )

        quantity = parse_number(
            price_match.group(2)
        )

        unit = normalize_unit(
            price_match.group(3)
        )

        discount_match = discount_pattern.search(body)

        discount = (
            parse_number(
                discount_match.group(1)
            )
            if discount_match
            else Decimal("0")
        )

        ppnbm_match = ppnbm_pattern.search(body)

        ppnbm_value = (
            parse_number(
                ppnbm_match.group(1)
            )
            if ppnbm_match
            else Decimal("0")
        )

        trailing_text = body[
            price_match.end():
        ]

        standalone_amounts = list(
            standalone_amount_pattern.finditer(
                trailing_text
            )
        )

        if standalone_amounts:
            line_amount = parse_number(
                standalone_amounts[-1].group(1)
            )
        else:
            line_amount = (
                quantity * unit_price
            )

        items.append(
            InvoiceItem(
                no=item_no,
                code=normalize_spaces(
                    block.group("code")
                ),
                name=description,
                qty=quantity,
                unit=unit,
                unit_price=unit_price,
                line_amount=line_amount,
                discount=discount,
                ppnbm_pdf=ppnbm_value,
            )
        )

    return items


def parse_invoice_pdf(
    pdf_bytes: bytes,
    source_name: str,
    *,
    zero_ppn_when_not_collected: bool = True,
) -> ParsedInvoice:
    text = extract_text(pdf_bytes)

    items = extract_items(text)

    gross_total = summary_amount(
        text,
        (
            r"Harga\s+Jual\s*/\s*Penggantian"
            r"\s*/\s*Uang\s+Muka\s*/\s*Termin"
        ),
        "Harga Jual",
    )

    discount_total = summary_amount(
        text,
        r"Dikurangi\s+Potongan\s+Harga",
        "Potongan Harga",
        optional=True,
    )

    dpp_total = summary_amount(
        text,
        r"Dasar\s+Pengenaan\s+Pajak",
        "DPP",
    )

    ppn_pdf = summary_amount(
        text,
        (
            r"Jumlah\s+PPN\s*"
            r"\(\s*Pajak\s+Pertambahan\s+Nilai\s*\)"
        ),
        "PPN",
    )

    ppnbm_total = summary_amount(
        text,
        (
            r"Jumlah\s+PPnBM\s*"
            r"\(\s*Pajak\s+Penjualan\s+atas\s+"
            r"Barang\s+Mewah\s*\)"
        ),
        "PPnBM",
        optional=True,
    )

    not_collected = bool(
        re.search(
            r"tidak\s+dipungut",
            text,
            re.I,
        )
    )

    if (
        not_collected
        and zero_ppn_when_not_collected
    ):
        ppn_output = Decimal("0")
    else:
        ppn_output = ppn_pdf

    item_weights = [
        item.weight
        for item in items
    ]

    dpp_parts = allocate(
        dpp_total,
        item_weights,
    )

    # PPN dibagi berdasarkan DPP item,
    # bukan langsung dari harga jual.
    ppn_parts = allocate(
        ppn_output,
        dpp_parts,
    )

    item_ppnbm_values = [
        round_rupiah(
            item.ppnbm_pdf
        )
        for item in items
    ]

    item_ppnbm_total = sum(
        item_ppnbm_values,
        Decimal("0"),
    )

    if item_ppnbm_total == ppnbm_total:
        ppnbm_parts = item_ppnbm_values
    else:
        ppnbm_parts = allocate(
            ppnbm_total,
            item_weights,
        )

    for (
        item,
        item_dpp,
        item_ppn,
        item_ppnbm,
    ) in zip(
        items,
        dpp_parts,
        ppn_parts,
        ppnbm_parts,
        strict=True,
    ):
        item.dpp = item_dpp
        item.ppn = item_ppn
        item.ppnbm = item_ppnbm

    parsed_invoice = ParsedInvoice(
        source_name=source_name,
        invoice_date=invoice_date(text),
        invoice_number=invoice_number(text),
        customer_name=customer_name(text),
        items=items,
        gross_total=gross_total,
        discount_total=discount_total,
        dpp_total=dpp_total,
        ppn_pdf=ppn_pdf,
        ppn_output=ppn_output,
        ppnbm_total=ppnbm_total,
        not_collected=not_collected,
    )

    item_gross_total = round_rupiah(
        sum(
            (
                item.line_amount
                for item in items
            ),
            Decimal("0"),
        )
    )

    item_discount_total = round_rupiah(
        sum(
            (
                item.discount
                for item in items
            ),
            Decimal("0"),
        )
    )

    if abs(
        item_gross_total - gross_total
    ) > Decimal("1"):
        parsed_invoice.warnings.append(
            "Total item "
            f"Rp {item_gross_total:,.0f} "
            "berbeda dengan ringkasan "
            f"Rp {gross_total:,.0f}."
        )

    if abs(
        item_discount_total - discount_total
    ) > Decimal("1"):
        parsed_invoice.warnings.append(
            "Potongan per item "
            f"Rp {item_discount_total:,.0f} "
            "berbeda dengan ringkasan "
            f"Rp {discount_total:,.0f}."
        )

    if (
        not_collected
        and zero_ppn_when_not_collected
        and ppn_pdf
    ):
        parsed_invoice.warnings.append(
            f"PPN PDF Rp {ppn_pdf:,.0f} "
            "dibuat 0.00 karena status tidak dipungut."
        )

    if item_ppnbm_total != ppnbm_total:
        parsed_invoice.warnings.append(
            "PPnBM item berbeda dengan ringkasan; "
            "nilai dialokasikan proporsional."
        )

    return parsed_invoice


def parse_batch(
    files: Iterable[tuple[str, bytes]],
    *,
    zero_ppn_when_not_collected: bool = True,
    skip_duplicates: bool = True,
    progress: (
        Callable[[int, int, str], None]
        | None
    ) = None,
) -> ParseBatchResult:
    file_list = list(files)

    invoices: list[ParsedInvoice] = []
    errors: list[dict[str, str]] = []

    seen_invoice_numbers: set[str] = set()

    total_files = len(file_list)

    for index, (
        filename,
        content,
    ) in enumerate(
        file_list,
        start=1,
    ):
        if progress:
            progress(
                index - 1,
                total_files,
                filename,
            )

        try:
            parsed_invoice = parse_invoice_pdf(
                content,
                filename,
                zero_ppn_when_not_collected=(
                    zero_ppn_when_not_collected
                ),
            )

            if (
                skip_duplicates
                and parsed_invoice.invoice_number
                in seen_invoice_numbers
            ):
                errors.append(
                    {
                        "FILE": filename,
                        "ERROR": (
                            "Nomor faktur duplikat: "
                            f"{parsed_invoice.invoice_number}."
                        ),
                    }
                )

            else:
                seen_invoice_numbers.add(
                    parsed_invoice.invoice_number
                )

                invoices.append(
                    parsed_invoice
                )

        except Exception as exc:
            errors.append(
                {
                    "FILE": filename,
                    "ERROR": str(exc),
                }
            )

        if progress:
            progress(
                index,
                total_files,
                filename,
            )

    return ParseBatchResult(
        invoices=invoices,
        errors=errors,
    )


def excel_number(
    value: Decimal | None,
) -> int | float | None:
    if value is None:
        return None

    if value == value.to_integral_value():
        return int(value)

    return float(value)


def invoices_to_rows(
    invoices: Iterable[ParsedInvoice],
    *,
    include_subtotal: bool = True,
    include_separator: bool = True,
    include_ppnbm: bool = True,
    format_legacy_number: bool = True,
    sort_by_date: bool = True,
) -> list[dict[str, object]]:
    invoice_list = list(invoices)

    if sort_by_date:
        invoice_list.sort(
            key=lambda invoice: (
                invoice.invoice_date,
                invoice.invoice_number,
                invoice.source_name,
            )
        )

    rows: list[dict[str, object]] = []

    for invoice_index, current_invoice in enumerate(
        invoice_list
    ):
        if format_legacy_number:
            output_invoice_number = (
                formatted_invoice_number(
                    current_invoice.invoice_number
                )
            )
        else:
            output_invoice_number = (
                current_invoice.invoice_number
            )

        has_multiple_items = (
            len(current_invoice.items) > 1
        )

        for item in current_invoice.items:
            item_total = (
                item.dpp
                + item.ppn
            )

            if include_ppnbm:
                item_total += item.ppnbm

            rows.append(
                {
                    "_ROW_TYPE": "item",

                    "TGL": current_invoice.invoice_date,

                    "NO. FAKTUR PAJAK": (
                        output_invoice_number
                    ),

                    "NAMA CUSTOMER": (
                        current_invoice.customer_name
                    ),

                    "JENIS BARANG": item.name,

                    "QTY": excel_number(
                        item.qty
                    ),

                    "SATUAN": item.unit,

                    "@ RP": excel_number(
                        item.unit_price
                    ),

                    "DPP": excel_number(
                        item.dpp
                    ),

                    "PPN": excel_number(
                        item.ppn
                    ),

                    "JUMLAH": (
                        None
                        if (
                            has_multiple_items
                            and include_subtotal
                        )
                        else excel_number(
                            item_total
                        )
                    ),
                }
            )

        if (
            has_multiple_items
            and include_subtotal
        ):
            rows.append(
                {
                    "_ROW_TYPE": "subtotal",

                    "TGL": current_invoice.invoice_date,

                    "NO. FAKTUR PAJAK": (
                        output_invoice_number
                    ),

                    "NAMA CUSTOMER": (
                        current_invoice.customer_name
                    ),

                    "JENIS BARANG": None,
                    "QTY": None,
                    "SATUAN": None,
                    "@ RP": None,

                    "DPP": excel_number(
                        current_invoice.dpp_total
                    ),

                    "PPN": excel_number(
                        current_invoice.ppn_output
                    ),

                    "JUMLAH": excel_number(
                        current_invoice.total(
                            include_ppnbm
                        )
                    ),
                }
            )

        is_last_invoice = (
            invoice_index
            == len(invoice_list) - 1
        )

        if (
            include_separator
            and not is_last_invoice
        ):
            rows.append(
                {
                    "_ROW_TYPE": "separator",
                    "TGL": None,
                    "NO. FAKTUR PAJAK": None,
                    "NAMA CUSTOMER": None,
                    "JENIS BARANG": None,
                    "QTY": None,
                    "SATUAN": None,
                    "@ RP": None,
                    "DPP": None,
                    "PPN": None,
                    "JUMLAH": None,
                }
            )

    return rows