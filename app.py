from __future__ import annotations

from datetime import datetime

import streamlit as st

from exporter import COLUMNS, rows_to_csv, rows_to_xlsx
from parser import invoices_to_rows, parse_batch

st.set_page_config(
    page_title="Bulk Extract Faktur Pajak",
    page_icon="🧾",
    layout="wide",
)

st.title("Bulk Extract Faktur Pajak ke Excel")
st.caption(
    "Upload banyak PDF sekaligus. Setiap item menjadi satu baris "
    "dan hasil Excel dimulai dari B1."
)

COLUMN_KEYS = {
    column: f"column_{index}"
    for index, column in enumerate(COLUMNS)
}

st.session_state.setdefault("all_columns", True)

for key in COLUMN_KEYS.values():
    st.session_state.setdefault(key, True)


def set_all_columns() -> None:
    """
    Mengaktifkan atau menonaktifkan semua checkbox kolom
    berdasarkan checkbox 'Aktifkan semua kolom'.
    """
    for key in COLUMN_KEYS.values():
        st.session_state[key] = st.session_state["all_columns"]


def sync_all_columns() -> None:
    """
    Mengubah status checkbox 'Aktifkan semua kolom'
    ketika salah satu kolom diubah.
    """
    st.session_state["all_columns"] = all(
        st.session_state[key]
        for key in COLUMN_KEYS.values()
    )


def preview_rows(
    rows: list[dict[str, object]],
    columns: list[str],
) -> list[dict[str, object]]:
    """
    Membuat data preview yang tampil seperti format Excel,
    tetapi tidak mengubah tipe data asli yang akan diekspor.
    """
    numeric_columns = {
        "QTY",
        "@ RP",
        "DPP",
        "PPN",
        "JUMLAH",
    }

    result: list[dict[str, object]] = []

    for row in rows:
        current: dict[str, object] = {}

        for column in columns:
            value = row.get(column)

            if column == "TGL" and value:
                current[column] = value.strftime("%d-%b-%y")

            elif column in numeric_columns and value is not None:
                current[column] = f"{float(value):,.2f}"

            else:
                current[column] = "" if value is None else value

        result.append(current)

    return result


with st.sidebar:
    st.header("Pengaturan")

    include_subtotal = st.checkbox(
        "Baris subtotal untuk faktur multi-barang",
        value=True,
        help=(
            "Untuk faktur dengan lebih dari satu barang, "
            "akan dibuat satu baris subtotal DPP, PPN, dan JUMLAH."
        ),
    )

    include_separator = st.checkbox(
        "Baris kosong antar faktur",
        value=True,
        help=(
            "Menambahkan satu baris kosong setelah setiap faktur "
            "agar mudah disalin ke rekapan utama."
        ),
    )

    zero_not_collected = st.checkbox(
        "PPN 'tidak dipungut' menjadi 0.00",
        value=True,
        help=(
            "Nilai PPN dari PDF tetap dibaca, tetapi output PPN "
            "dibuat 0.00 apabila terdapat keterangan tidak dipungut."
        ),
    )

    include_ppnbm = st.checkbox(
        "Sertakan PPnBM dalam JUMLAH",
        value=True,
        help="JUMLAH = DPP + PPN + PPnBM.",
    )

    format_legacy = st.checkbox(
        "Format nomor faktur lama 16 digit",
        value=True,
        help=(
            "Contoh: 0100012467950711 otomatis menjadi "
            "010.001-24.67950711. Nomor Coretax 17 digit tidak diubah."
        ),
    )

    sort_by_date = st.checkbox(
        "Urutkan berdasarkan tanggal dan nomor",
        value=True,
    )

    skip_duplicates = st.checkbox(
        "Lewati nomor faktur duplikat",
        value=True,
        help=(
            "Mencegah nomor faktur yang sama masuk dua kali "
            "dalam satu proses upload."
        ),
    )

    st.divider()
    st.subheader("Kolom yang Dipilih")

    st.checkbox(
        "Aktifkan semua kolom",
        key="all_columns",
        on_change=set_all_columns,
    )

    left_column, right_column = st.columns(2)

    for index, column in enumerate(COLUMNS):
        target_column = (
            left_column
            if index % 2 == 0
            else right_column
        )

        with target_column:
            st.checkbox(
                column,
                key=COLUMN_KEYS[column],
                on_change=sync_all_columns,
            )


selected_columns = [
    column
    for column in COLUMNS
    if st.session_state[COLUMN_KEYS[column]]
]

if not selected_columns:
    st.error("Pilih minimal satu kolom.")
    st.stop()


uploads = st.file_uploader(
    "Pilih PDF Faktur Pajak",
    type=["pdf"],
    accept_multiple_files=True,
    help=(
        "Dapat memilih banyak PDF sekaligus. "
        "Satu file yang gagal tidak menghentikan file lainnya."
    ),
)

if not uploads:
    st.info("Belum ada PDF yang diunggah.")
    st.stop()


progress_bar = st.progress(
    0.0,
    text="Menyiapkan ekstraksi...",
)


def progress(
    done: int,
    total: int,
    filename: str,
) -> None:
    progress_value = done / total if total else 1.0

    progress_bar.progress(
        progress_value,
        text=f"Memproses {done}/{total}: {filename}",
    )


result = parse_batch(
    [
        (uploaded_file.name, uploaded_file.getvalue())
        for uploaded_file in uploads
    ],
    zero_ppn_when_not_collected=zero_not_collected,
    skip_duplicates=skip_duplicates,
    progress=progress,
)

progress_bar.empty()


rows = invoices_to_rows(
    result.invoices,
    include_subtotal=include_subtotal,
    include_separator=include_separator,
    include_ppnbm=include_ppnbm,
    format_legacy_number=format_legacy,
    sort_by_date=sort_by_date,
)


item_count = sum(
    row.get("_ROW_TYPE") == "item"
    for row in rows
)

subtotal_count = sum(
    row.get("_ROW_TYPE") == "subtotal"
    for row in rows
)


metrics = st.columns(4)

metrics[0].metric(
    "PDF Diunggah",
    len(uploads),
)

metrics[1].metric(
    "Faktur Berhasil",
    len(result.invoices),
)

metrics[2].metric(
    "Baris Item",
    item_count,
)

metrics[3].metric(
    "Gagal/Dilewati",
    len(result.errors),
)


if result.errors:
    with st.expander(
        "File gagal atau dilewati",
        expanded=True,
    ):
        st.dataframe(
            result.errors,
            use_container_width=True,
            hide_index=True,
        )


if not result.invoices:
    st.error("Tidak ada faktur yang berhasil diekstrak.")
    st.stop()


warnings = [
    {
        "FILE": invoice.source_name,
        "PERINGATAN": warning,
    }
    for invoice in result.invoices
    for warning in invoice.warnings
]

if warnings:
    with st.expander("Peringatan validasi"):
        st.dataframe(
            warnings,
            use_container_width=True,
            hide_index=True,
        )


st.subheader("Preview")

st.caption(
    f"{item_count} baris item dan "
    f"{subtotal_count} baris subtotal. "
    "Angka Excel memakai 2 digit desimal."
)

st.dataframe(
    preview_rows(
        rows,
        selected_columns,
    ),
    column_order=selected_columns,
    use_container_width=True,
    hide_index=True,
)


timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

excel_bytes = rows_to_xlsx(
    rows,
    columns=selected_columns,
)

csv_bytes = rows_to_csv(
    rows,
    columns=selected_columns,
)


download_excel, download_csv = st.columns(2)

download_excel.download_button(
    label="Download Excel",
    data=excel_bytes,
    file_name=f"rekap_faktur_{timestamp}.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    use_container_width=True,
    type="primary",
)

download_csv.download_button(
    label="Download CSV",
    data=csv_bytes,
    file_name=f"rekap_faktur_{timestamp}.csv",
    mime="text/csv",
    use_container_width=True,
)


with st.expander("Aturan parsing"):
    st.markdown(
        """
- **TGL** dibaca dari tanggal tanda tangan di bagian bawah faktur.
- **NO. FAKTUR PAJAK** dibaca dari `Kode dan Nomor Seri Faktur Pajak`.
- **NAMA CUSTOMER** dibaca langsung dari bagian pembeli tanpa alias nama.
- **JENIS BARANG** mengambil seluruh teks sebelum baris `Rp harga x qty satuan`.
- Teks seperti `SEBANYAK ... DENGAN HARGA JUAL ... PER PIECE` tetap dipertahankan.
- **QTY**, **SATUAN**, dan **@ RP** dibaca dari pola `Rp ... x ... satuan`.
- Satuan yang belum terdapat dalam mapping tetap digunakan dalam bentuk uppercase.
- **Potongan Harga** dibaca per item dan digunakan untuk menentukan nilai netto.
- **DPP** mengikuti total `Dasar Pengenaan Pajak` pada PDF.
- **PPN** mengikuti total `Jumlah PPN` pada PDF.
- **PPnBM** dibaca dari item dan bagian ringkasan faktur.
- Selisih pembulatan diletakkan pada item terakhir agar subtotal sama dengan PDF.
"""
    )


st.caption(
    "PDF harus memiliki text layer. "
    "PDF scan tanpa text layer membutuhkan OCR."
)