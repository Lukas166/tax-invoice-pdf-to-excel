from __future__ import annotations

import csv
import io
import math
from datetime import datetime
from decimal import Decimal
from typing import Iterable, Mapping, Sequence

import xlsxwriter


COLUMNS = [
    "TGL",
    "NO. FAKTUR PAJAK",
    "NAMA CUSTOMER",
    "JENIS BARANG",
    "QTY",
    "SATUAN",
    "@ RP",
    "DPP",
    "PPN",
    "JUMLAH",
]


WIDTHS = {
    "TGL": 12,
    "NO. FAKTUR PAJAK": 26,
    "NAMA CUSTOMER": 28,
    "JENIS BARANG": 82,
    "QTY": 14,
    "SATUAN": 10,
    "@ RP": 16,
    "DPP": 19,
    "PPN": 19,
    "JUMLAH": 20,
}


# Index kolom dimulai dari 0.
# 0 = A, 1 = B.
START_COL = 1


def selected_columns(
    columns: Sequence[str] | None,
) -> list[str]:
    if columns is None:
        return COLUMNS.copy()

    unknown_columns = [
        column
        for column in columns
        if column not in COLUMNS
    ]

    if unknown_columns:
        raise ValueError(
            f"Kolom tidak dikenal: {unknown_columns}"
        )

    result = [
        column
        for column in COLUMNS
        if column in columns
    ]

    if not result:
        raise ValueError(
            "Minimal satu kolom harus dipilih."
        )

    return result


def row_height(
    description: object,
) -> float:
    """
    Menghitung tinggi baris berdasarkan panjang
    deskripsi jenis barang.
    """
    character_count = len(
        str(description or "")
    )

    estimated_lines = max(
        1,
        math.ceil(
            character_count / 105
        ),
    )

    return min(
        120,
        max(
            30,
            estimated_lines * 15 + 6,
        ),
    )


def rows_to_xlsx(
    rows: Iterable[Mapping[str, object]],
    *,
    columns: Sequence[str] | None = None,
) -> bytes:
    columns = selected_columns(columns)
    rows = list(rows)

    output = io.BytesIO()

    workbook = xlsxwriter.Workbook(
        output,
        {
            "in_memory": True,
        },
    )

    worksheet = workbook.add_worksheet(
        "REKAP FAKTUR"
    )

    base_format = {
        "font_name": "Franklin Gothic Book",
        "font_size": 10,
        "border": 1,
        "valign": "vcenter",
    }

    header_format = workbook.add_format(
        {
            **base_format,
            "bold": True,
            "align": "center",
            "bg_color": "#D9D9D9",
            "text_wrap": True,
        }
    )

    text_format = workbook.add_format(
        {
            **base_format,
            "text_wrap": True,
        }
    )

    center_format = workbook.add_format(
        {
            **base_format,
            "align": "center",
        }
    )

    date_format = workbook.add_format(
        {
            **base_format,
            "num_format": "dd-mmm-yy",
            "align": "center",
        }
    )

    # Semua nilai numerik ditampilkan dengan dua angka desimal.
    #
    # Excel English:
    # 1,020.00
    #
    # Excel Indonesia:
    # 1.020,00
    #
    # Nilai tetap disimpan sebagai angka, bukan teks.
    number_code = (
        "#,##0.00;"
        "[Red]-#,##0.00;"
        "0.00"
    )

    number_format = workbook.add_format(
        {
            **base_format,
            "num_format": number_code,
            "align": "right",
        }
    )

    subtotal_number_format = workbook.add_format(
        {
            **base_format,
            "bold": True,
            "num_format": number_code,
            "align": "right",
        }
    )

    blank_format = workbook.add_format(
        base_format
    )

    worksheet.hide_gridlines(2)

    # Membekukan header baris pertama.
    worksheet.freeze_panes(
        1,
        0,
    )

    # Kolom A sengaja kosong.
    worksheet.set_column(
        0,
        0,
        3.43,
    )

    worksheet.set_row(
        0,
        28,
    )

    # Header dimulai dari B1.
    for index, column in enumerate(columns):
        excel_column = START_COL + index

        worksheet.write(
            0,
            excel_column,
            column,
            header_format,
        )

        worksheet.set_column(
            excel_column,
            excel_column,
            WIDTHS[column],
        )

    last_column = (
        START_COL
        + len(columns)
        - 1
    )

    worksheet.autofilter(
        0,
        START_COL,
        0,
        last_column,
    )

    numeric_columns = {
        "QTY",
        "@ RP",
        "DPP",
        "PPN",
        "JUMLAH",
    }

    for excel_row, row in enumerate(
        rows,
        start=1,
    ):
        row_type = row.get(
            "_ROW_TYPE",
            "item",
        )

        if row_type == "separator":
            worksheet.set_row(
                excel_row,
                8,
            )

            for index in range(
                len(columns)
            ):
                worksheet.write_blank(
                    excel_row,
                    START_COL + index,
                    None,
                    blank_format,
                )

            continue

        if row_type == "subtotal":
            worksheet.set_row(
                excel_row,
                22,
            )
        else:
            worksheet.set_row(
                excel_row,
                row_height(
                    row.get("JENIS BARANG")
                ),
            )

        for index, column in enumerate(columns):
            excel_column = START_COL + index
            value = row.get(column)

            if column == "TGL":
                cell_format = date_format

            elif column in {
                "NO. FAKTUR PAJAK",
                "SATUAN",
            }:
                cell_format = center_format

            elif column in numeric_columns:
                if (
                    row_type == "subtotal"
                    and column in {
                        "DPP",
                        "PPN",
                        "JUMLAH",
                    }
                ):
                    cell_format = (
                        subtotal_number_format
                    )
                else:
                    cell_format = number_format

            else:
                cell_format = text_format

            if value is None:
                worksheet.write_blank(
                    excel_row,
                    excel_column,
                    None,
                    cell_format,
                )

            elif column == "TGL":
                if isinstance(
                    value,
                    datetime,
                ):
                    excel_date = value
                else:
                    excel_date = datetime.combine(
                        value,
                        datetime.min.time(),
                    )

                worksheet.write_datetime(
                    excel_row,
                    excel_column,
                    excel_date,
                    cell_format,
                )

            elif column == "NO. FAKTUR PAJAK":
                # Nomor faktur wajib disimpan sebagai teks
                # agar angka 0 di depan tidak hilang.
                worksheet.write_string(
                    excel_row,
                    excel_column,
                    str(value),
                    cell_format,
                )

            elif column in numeric_columns:
                numeric_value = (
                    float(value)
                    if isinstance(
                        value,
                        Decimal,
                    )
                    else value
                )

                worksheet.write_number(
                    excel_row,
                    excel_column,
                    numeric_value,
                    cell_format,
                )

            else:
                worksheet.write(
                    excel_row,
                    excel_column,
                    value,
                    cell_format,
                )

    worksheet.set_landscape()

    worksheet.fit_to_pages(
        1,
        0,
    )

    worksheet.repeat_rows(0)

    if rows:
        worksheet.print_area(
            0,
            START_COL,
            len(rows),
            last_column,
        )

    workbook.close()

    output.seek(0)

    return output.getvalue()


def rows_to_csv(
    rows: Iterable[Mapping[str, object]],
    *,
    columns: Sequence[str] | None = None,
) -> bytes:
    columns = selected_columns(columns)

    output = io.StringIO()

    writer = csv.DictWriter(
        output,
        fieldnames=columns,
        extrasaction="ignore",
        lineterminator="\n",
    )

    writer.writeheader()

    numeric_columns = {
        "QTY",
        "@ RP",
        "DPP",
        "PPN",
        "JUMLAH",
    }

    for row in rows:
        if row.get("_ROW_TYPE") == "separator":
            writer.writerow(
                {
                    column: ""
                    for column in columns
                }
            )

            continue

        clean = {
            column: row.get(column)
            for column in columns
        }

        if clean.get("TGL"):
            clean["TGL"] = clean[
                "TGL"
            ].strftime(
                "%d-%b-%y"
            )

        for column in numeric_columns:
            if (
                column in clean
                and clean[column] is not None
            ):
                clean[column] = (
                    f"{float(clean[column]):.2f}"
                )

        writer.writerow(clean)

    return output.getvalue().encode(
        "utf-8-sig"
    )