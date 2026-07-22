from __future__ import annotations

import csv
import io
from datetime import datetime
from decimal import Decimal
from typing import Iterable, Mapping, Sequence

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name


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


# Lebar kolom mengikuti kurang lebih format
# Rekapan Faktur Pajak yang digunakan sebagai referensi.
WIDTHS = {
    "TGL": 11,
    "NO. FAKTUR PAJAK": 30,
    "NAMA CUSTOMER": 23.5,
    "JENIS BARANG": 31,
    "QTY": 13,
    "SATUAN": 7,
    "@ RP": 15,
    "DPP": 22,
    "PPN": 21,
    "JUMLAH": 22,
}


# Index kolom dimulai dari 0:
# 0 = A
# 1 = B
#
# Hasil dimulai dari B1.
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
    workbook.set_calc_mode("auto")

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
        }
    )

    # Tidak memakai text_wrap.
    # Deskripsi panjang tetap satu baris dan akan
    # terpotong secara visual di sisi kanan.
    text_format = workbook.add_format(
        {
            **base_format,
            "align": "left",
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

    # Nilai disimpan sebagai angka asli.
    #
    # Excel regional English:
    # 1,020.00
    #
    # Excel regional Indonesia:
    # 1.020,00
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

    # JUMLAH untuk faktur yang hanya memiliki satu item.
    total_number_format = workbook.add_format(
        {
            **base_format,
            "bold": True,
            "num_format": number_code,
            "align": "right",
        }
    )

    # DPP, PPN, dan JUMLAH pada subtotal faktur multi-item.
    subtotal_number_format = workbook.add_format(
        {
            **base_format,
            "bold": True,
            "num_format": number_code,
            "align": "right",
        }
    )

    # Baris kosong tetap memiliki border penuh.
    blank_format = workbook.add_format(
        {
            **base_format,
        }
    )

    # Hilangkan gridline bawaan karena tabel sudah
    # menggunakan border sendiri.
    worksheet.hide_gridlines(2)

    # Tinggi default seluruh baris, termasuk baris kosong.
    worksheet.set_default_row(15)

    # Header sedikit lebih tinggi.
    worksheet.set_row(
        0,
        28,
    )

    # Freeze hanya baris header.
    worksheet.freeze_panes(
        1,
        0,
    )

    # Kolom A sengaja dikosongkan.
    worksheet.set_column(
        0,
        0,
        3.43,
    )

    # Pemetaan posisi kolom Excel secara dinamis
    column_positions = {
        column_name: START_COL + index
        for index, column_name in enumerate(columns)
    }

    # Hilangkan warning "number stored as text" jika kolom NO. FAKTUR PAJAK dipilih
    if "NO. FAKTUR PAJAK" in column_positions:
        no_faktur_col = column_positions["NO. FAKTUR PAJAK"]
        worksheet.ignore_errors(
            {
                "number_stored_as_text": (
                    f"{xl_col_to_name(no_faktur_col)}2:"
                    f"{xl_col_to_name(no_faktur_col)}{max(2, len(rows) + 1)}"
                )
            }
        )

    # Tulis header mulai dari B1.
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

    numeric_columns = {
        "QTY",
        "@ RP",
        "DPP",
        "PPN",
        "JUMLAH",
    }

    # Pelacakan batas baris item untuk setiap faktur.
    item_start_excel_row: int | None = None
    item_end_excel_row: int | None = None
    current_invoice_key: tuple[object, object, object] | None = None

    for excel_row, row in enumerate(
        rows,
        start=1,
    ):
        row_type = row.get(
            "_ROW_TYPE",
            "item",
        )

        # Baris separator tetap menggunakan tinggi default
        # dan tetap memiliki border penuh.
        if row_type == "separator":
            for index in range(len(columns)):
                worksheet.write_blank(
                    excel_row,
                    START_COL + index,
                    None,
                    blank_format,
                )
            item_start_excel_row = None
            item_end_excel_row = None
            current_invoice_key = None
            continue

        excel_row_number = excel_row + 1

        if row_type == "item":
            invoice_key = (
                row.get("NO. FAKTUR PAJAK"),
                row.get("TGL"),
                row.get("NAMA CUSTOMER"),
            )

            # Faktur baru harus memulai range subtotal baru,
            # meskipun tidak ada baris separator.
            if invoice_key != current_invoice_key:
                current_invoice_key = invoice_key
                item_start_excel_row = excel_row_number
                item_end_excel_row = excel_row_number
            else:
                if item_start_excel_row is None:
                    item_start_excel_row = excel_row_number
                item_end_excel_row = excel_row_number

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
                # Subtotal faktur multi-item:
                # DPP, PPN, dan JUMLAH dibuat bold.
                if (
                    row_type == "subtotal"
                    and column in {
                        "DPP",
                        "PPN",
                        "JUMLAH",
                    }
                ):
                    cell_format = subtotal_number_format

                # Faktur satu item:
                # JUMLAH langsung berada di baris item
                # dan dibuat bold.
                elif (
                    row_type == "item"
                    and column == "JUMLAH"
                    and value is not None
                ):
                    cell_format = total_number_format

                else:
                    cell_format = number_format

            else:
                cell_format = text_format

            # Penanganan penulisan sel dengan Formula atau Nilai Biasa
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
                # Disimpan sebagai teks agar angka nol
                # di depan tidak hilang.
                worksheet.write_string(
                    excel_row,
                    excel_column,
                    str(value),
                    cell_format,
                )

            elif column in numeric_columns:
                cached_value = (
                    float(value)
                    if isinstance(value, Decimal)
                    else value
                )

                formula_written = False

                # 1. Formula JUMLAH untuk Faktur Satu Item
                if (
                    row_type == "item"
                    and column == "JUMLAH"
                    and "DPP" in column_positions
                    and "PPN" in column_positions
                ):
                    dpp_col_name = xl_col_to_name(column_positions["DPP"])
                    ppn_col_name = xl_col_to_name(column_positions["PPN"])

                    dpp_val = row.get("DPP")
                    ppn_val = row.get("PPN")

                    dpp_num = float(dpp_val) if isinstance(dpp_val, Decimal) else (dpp_val or 0.0)
                    ppn_num = float(ppn_val) if isinstance(ppn_val, Decimal) else (ppn_val or 0.0)

                    ppnbm_comp = cached_value - dpp_num - ppn_num

                    if abs(ppnbm_comp) < 1e-6:
                        formula_str = f"={dpp_col_name}{excel_row_number}+{ppn_col_name}{excel_row_number}"
                    else:
                        ppnbm_str = f"{ppnbm_comp:+.2f}".rstrip("0").rstrip(".")
                        formula_str = f"={dpp_col_name}{excel_row_number}+{ppn_col_name}{excel_row_number}{ppnbm_str}"

                    worksheet.write_formula(
                        excel_row,
                        excel_column,
                        formula_str,
                        cell_format,
                        cached_value,
                    )
                    formula_written = True

                # 2. Formula Subtotal (DPP, PPN, JUMLAH) untuk Faktur Multi-Item
                elif row_type == "subtotal":
                    if column == "DPP" and item_start_excel_row is not None and item_end_excel_row is not None:
                        dpp_col_name = xl_col_to_name(column_positions["DPP"])
                        formula_str = f"=SUM({dpp_col_name}{item_start_excel_row}:{dpp_col_name}{item_end_excel_row})"
                        worksheet.write_formula(
                            excel_row,
                            excel_column,
                            formula_str,
                            cell_format,
                            cached_value,
                        )
                        formula_written = True

                    elif column == "PPN" and item_start_excel_row is not None and item_end_excel_row is not None:
                        ppn_col_name = xl_col_to_name(column_positions["PPN"])
                        formula_str = f"=SUM({ppn_col_name}{item_start_excel_row}:{ppn_col_name}{item_end_excel_row})"
                        worksheet.write_formula(
                            excel_row,
                            excel_column,
                            formula_str,
                            cell_format,
                            cached_value,
                        )
                        formula_written = True

                    elif (
                        column == "JUMLAH"
                        and "DPP" in column_positions
                        and "PPN" in column_positions
                    ):
                        dpp_col_name = xl_col_to_name(column_positions["DPP"])
                        ppn_col_name = xl_col_to_name(column_positions["PPN"])

                        sub_dpp_val = row.get("DPP")
                        sub_ppn_val = row.get("PPN")

                        sub_dpp_num = float(sub_dpp_val) if isinstance(sub_dpp_val, Decimal) else (sub_dpp_val or 0.0)
                        sub_ppn_num = float(sub_ppn_val) if isinstance(sub_ppn_val, Decimal) else (sub_ppn_val or 0.0)

                        ppnbm_comp = cached_value - sub_dpp_num - sub_ppn_num

                        if abs(ppnbm_comp) < 1e-6:
                            formula_str = f"={dpp_col_name}{excel_row_number}+{ppn_col_name}{excel_row_number}"
                        else:
                            ppnbm_str = f"{ppnbm_comp:+.2f}".rstrip("0").rstrip(".")
                            formula_str = f"={dpp_col_name}{excel_row_number}+{ppn_col_name}{excel_row_number}{ppnbm_str}"

                        worksheet.write_formula(
                            excel_row,
                            excel_column,
                            formula_str,
                            cell_format,
                            cached_value,
                        )
                        formula_written = True

                if not formula_written:
                    worksheet.write_number(
                        excel_row,
                        excel_column,
                        cached_value,
                        cell_format,
                    )

            else:
                worksheet.write_string(
                    excel_row,
                    excel_column,
                    str(value),
                    cell_format,
                )

        if row_type == "subtotal":
            item_start_excel_row = None
            item_end_excel_row = None
            current_invoice_key = None

    # Filter mencakup header dan seluruh data.
    worksheet.autofilter(
        0,
        START_COL,
        len(rows),
        last_column,
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