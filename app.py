from __future__ import annotations

from datetime import datetime

import streamlit as st

from exporter import COLUMNS, rows_to_csv, rows_to_xlsx
from parser import invoices_to_rows, parse_batch

st.set_page_config(
    page_title="Ekstraktor Faktur Pajak ke Excel",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Injeksi CSS Solid Custom (Headline Lebih Besar, Spasi Seimbang & Bebas Emoji)
st.markdown(
    """
    <style>
    /* Styling Dasar & Tipografi */
    body, .stApp {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        color: #1e293b;
        background-color: #f8fafc;
    }
    
    /* Spasi Atas Halaman Yang Proporsional */
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 2.5rem !important;
    }
    
    /* Header Utama - Headline Lebih Besar (42px) */
    .header-container {
        padding-top: 4px;
        margin-bottom: 8px;
    }
    .app-title {
        font-size: 42px;
        font-weight: 800;
        color: #0f172a;
        margin: 0 0 8px 0;
        letter-spacing: -0.025em;
        line-height: 1.2;
    }
    .app-subtitle {
        font-size: 16px;
        color: #475569;
        margin: 0 0 16px 0;
        line-height: 1.45;
    }
    
    /* Kartu Langkah Penggunaan - Card Polos Berjarak Seimbang */
    .steps-container {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 14px;
        margin-bottom: 20px;
    }
    .step-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 14px 16px;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03);
    }
    .step-number {
        font-size: 11px;
        font-weight: 700;
        color: #64748b;
        text-transform: uppercase;
        margin-bottom: 4px;
    }
    .step-title {
        font-size: 14px;
        font-weight: 600;
        color: #0f172a;
        margin-bottom: 4px;
    }
    .step-desc {
        font-size: 12px;
        color: #64748b;
        margin: 0;
        line-height: 1.4;
    }
    
    /* Seksi Konten & Judul */
    .section-title {
        font-size: 18px;
        font-weight: 600;
        color: #0f172a;
        margin-top: 4px;
        margin-bottom: 10px;
    }
    
    /* Custom Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    .sidebar-header {
        font-size: 18px;
        font-weight: 700;
        color: #0f172a;
        padding-bottom: 6px;
        border-bottom: 2px solid #2563eb;
        margin-bottom: 14px;
    }
    .sidebar-section-label {
        font-size: 12px;
        font-weight: 700;
        color: #334155;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-top: 16px;
        margin-bottom: 6px;
    }
    
    /* Custom Stat Metric Cards Solid Polos */
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px 16px;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03);
    }
    div[data-testid="stMetricLabel"] {
        font-size: 12px;
        font-weight: 600;
        color: #64748b;
    }
    div[data-testid="stMetricValue"] {
        font-size: 22px;
        font-weight: 700;
        color: #0f172a;
    }

    /* Penyesuaian Tombol Solid */
    .stButton button, .stDownloadButton button {
        border-radius: 6px;
        font-weight: 600;
        border: none;
    }
    
    /* Divider & Expander Clean Styling */
    hr {
        margin: 14px 0;
        border-color: #e2e8f0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Header Aplikasi (Headline Ekstra Besar 42px)
st.markdown(
    """
    <div class="header-container">
        <div class="app-title">Ekstraktor Faktur Pajak ke Excel</div>
        <div class="app-subtitle">
            Konversi berkas PDF Faktur Pajak (E-Faktur & Coretax) menjadi rekapitulasi data Excel yang terstruktur, presisi, dan siap pakai.
        </div>
    </div>
    <div class="steps-container">
        <div class="step-card">
            <div class="step-number">Langkah 1</div>
            <div class="step-title">Unggah Berkas PDF</div>
            <div class="step-desc">Pilih atau seret satu atau banyak berkas PDF Faktur Pajak sekaligus.</div>
        </div>
        <div class="step-card">
            <div class="step-number">Langkah 2</div>
            <div class="step-title">Sesuaikan Pengaturan</div>
            <div class="step-desc">Atur susunan kolom, pemisahan baris, serta penanganan nilai PPN.</div>
        </div>
        <div class="step-card">
            <div class="step-number">Langkah 3</div>
            <div class="step-title">Unduh Rekapitulasi</div>
            <div class="step-desc">Dapatkan hasil ekstraksi dalam format Excel (.xlsx) yang tersusun mulai sel B1.</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

COLUMN_KEYS = {
    column: f"column_{index}"
    for index, column in enumerate(COLUMNS)
}

st.session_state.setdefault("all_columns", True)

for key in COLUMN_KEYS.values():
    st.session_state.setdefault(key, True)


def set_all_columns() -> None:
    """Mengaktifkan atau menonaktifkan seluruh kolom laporan."""
    for key in COLUMN_KEYS.values():
        st.session_state[key] = st.session_state["all_columns"]


def sync_all_columns() -> None:
    """Menyelaraskan status checkbox master kolom."""
    st.session_state["all_columns"] = all(
        st.session_state[key]
        for key in COLUMN_KEYS.values()
    )


def preview_rows(
    rows: list[dict[str, object]],
    columns: list[str],
) -> list[dict[str, object]]:
    """Membuat format tampilan data pratinjau yang rapi dan konsisten."""
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


# Pengaturan Sidebar
with st.sidebar:
    st.markdown('<div class="sidebar-header">Pengaturan & Opsi</div>', unsafe_allow_html=True)

    st.markdown('<div class="sidebar-section-label">Format Baris & Rekapan</div>', unsafe_allow_html=True)
    
    include_subtotal = st.checkbox(
        "Baris subtotal faktur multi-item",
        value=True,
        help="Menambahkan baris penutup subtotal DPP, PPN, dan Total pada faktur yang memiliki lebih dari satu barang/jasa.",
    )

    include_separator = st.checkbox(
        "Baris pemisah antar-faktur",
        value=True,
        help="Menyisipkan baris kosong setelah setiap faktur agar data lebih rapi saat disalin ke lembar kerja utama.",
    )

    sort_by_date = st.checkbox(
        "Urutkan berdasarkan tanggal & nomor",
        value=True,
        help="Mengurutkan baris rekapitulasi secara kronologis berdasarkan tanggal transaksi dan nomor seri faktur.",
    )

    st.markdown('<div class="sidebar-section-label">Aturan Pajak & Validasi</div>', unsafe_allow_html=True)

    zero_not_collected = st.checkbox(
        "Ubah PPN tidak dipungut menjadi 0.00",
        value=True,
        help="Mencatat output nilai PPN sebesar 0.00 apabila faktur memuat keterangan fasilitas tidak dipungut (nilai asli PDF tetap dibaca).",
    )

    include_ppnbm = st.checkbox(
        "Sertakan PPnBM dalam Total Jumlah",
        value=True,
        help="Menghitung kolom Total Jumlah dengan rumus: DPP + PPN + PPnBM.",
    )

    format_legacy = st.checkbox(
        "Format standar nomor faktur 16 digit",
        value=True,
        help="Mengubah format angka 16 digit menjadi standar 010.001-24.xxxx. Nomor Coretax 17 digit tidak diubah.",
    )

    skip_duplicates = st.checkbox(
        "Abaikan nomor faktur duplikat",
        value=True,
        help="Mencegah nomor faktur yang sama terproses lebih dari satu kali dalam satu sesi pengunggahan.",
    )

    st.markdown('<div class="sidebar-section-label">Pilihan Kolom Laporan</div>', unsafe_allow_html=True)

    st.checkbox(
        "Pilih semua kolom",
        key="all_columns",
        on_change=set_all_columns,
    )

    left_column, right_column = st.columns(2)

    for index, column in enumerate(COLUMNS):
        target_column = left_column if index % 2 == 0 else right_column

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
    st.error("Pilihlah setidaknya satu kolom laporan pada panel pengaturan.")
    st.stop()


# Area Pengunggahan Berkas
st.markdown('<div class="section-title">Area Pengunggahan Berkas</div>', unsafe_allow_html=True)

uploads = st.file_uploader(
    "Unggah Berkas PDF Faktur Pajak",
    type=["pdf"],
    accept_multiple_files=True,
    help="Anda dapat memilih banyak berkas PDF sekaligus. Kegagalan membaca satu berkas tidak membatalkan berkas lainnya.",
    label_visibility="collapsed",
)

if not uploads:
    st.info("Silakan unggah satu atau beberapa berkas PDF Faktur Pajak untuk memulai proses ekstraksi.")
    st.stop()


progress_bar = st.progress(0.0, text="Menyiapkan proses ekstraksi data...")


def progress(done: int, total: int, filename: str) -> None:
    progress_value = done / total if total else 1.0
    progress_bar.progress(
        progress_value,
        text=f"Memproses berkas {done} dari {total}: {filename}",
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

item_count = sum(row.get("_ROW_TYPE") == "item" for row in rows)
subtotal_count = sum(row.get("_ROW_TYPE") == "subtotal" for row in rows)

# Dashboard Statistik Solid Polos
metrics = st.columns(4)
metrics[0].metric("Total Berkas PDF", len(uploads))
metrics[1].metric("Faktur Berhasil", len(result.invoices))
metrics[2].metric("Total Baris Barang/Jasa", item_count)
metrics[3].metric("Berkas Terkendala", len(result.errors))

# Tampilan Pratinjau & Hasil Ekstraksi
if not result.invoices:
    st.error("Tidak ada data faktur yang berhasil diekstrak dari berkas PDF yang diunggah. Pastikan PDF yang diunggah berisi teks Faktur Pajak resmi.")
    if result.errors:
        st.subheader("Rincian Berkas Terkendala")
        st.dataframe(result.errors, use_container_width=True, hide_index=True)
    st.stop()

tab_preview, tab_errors, tab_warnings = st.tabs(
    [
        "Pratinjau Rekapan Data",
        f"Berkas Terkendala ({len(result.errors)})",
        "Peringatan Validasi",
    ]
)

with tab_preview:
    st.caption(
        f"Menampilkan {item_count} baris item barang/jasa dan {subtotal_count} baris subtotal. "
        "Seluruh nilai numerik diformat dengan 2 digit desimal."
    )

    st.dataframe(
        preview_rows(rows, selected_columns),
        column_order=selected_columns,
        use_container_width=True,
        hide_index=True,
    )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    excel_bytes = rows_to_xlsx(rows, columns=selected_columns)
    csv_bytes = rows_to_csv(rows, columns=selected_columns)

    st.markdown('<div class="section-title">Unduh Hasil Rekapitulasi</div>', unsafe_allow_html=True)

    download_excel, download_csv = st.columns(2)

    download_excel.download_button(
        label="Unduh Rekapitulasi Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"rekap_faktur_pajak_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    download_csv.download_button(
        label="Unduh Rekapitulasi CSV (.csv)",
        data=csv_bytes,
        file_name=f"rekap_faktur_pajak_{timestamp}.csv",
        mime="text/csv",
        use_container_width=True,
    )

with tab_errors:
    if result.errors:
        st.warning("Berikut adalah daftar berkas yang tidak dapat diproses atau dilewati karena nomor faktur duplikat.")
        st.dataframe(result.errors, use_container_width=True, hide_index=True)
    else:
        st.info("Seluruh berkas PDF berhasil diproses tanpa ada kendala.")

with tab_warnings:
    warnings = [
        {"FILE": invoice.source_name, "PERINGATAN": warning}
        for invoice in result.invoices
        for warning in invoice.warnings
    ]
    if warnings:
        st.warning("Beberapa catatan validasi ditemukan selama pemrosesan berkas:")
        st.dataframe(warnings, use_container_width=True, hide_index=True)
    else:
        st.info("Tidak ada peringatan validasi data.")

with st.expander("Petunjuk & Ketentuan Teknis Pemrosesan Data"):
    st.markdown(
        """
        - **Tanggal (TGL)**: Dibaca dari tanggal penandatanganan faktur pada bagian bawah dokumen.
        - **Nomor Faktur Pajak**: Dibaca dari bagian Kode dan Nomor Seri Faktur Pajak (E-Faktur 16 digit atau Coretax 17 digit).
        - **Nama Customer**: Dibaca langsung dari kolom Identitas Pembeli Barang/Penerima Jasa.
        - **Jenis Barang / Jasa**: Mengambil deskripsi lengkap barang/jasa sebelum baris rincian harga.
        - **Kuantitas (QTY) & Satuan**: Dibaca dari pola rincian kuantitas per satuan item.
        - **Dasar Pengenaan Pajak (DPP)**: Mengikuti total Dasar Pengenaan Pajak pada dokumen PDF.
        - **PPN & PPnBM**: Mengikuti perhitungan nilai PPN dan PPnBM resmi sesuai aturan faktur pajak.
        - **Pembulatan Selisih**: Apabila terdapat selisih pembulatan desimal, penyesuaian otomatis diletakkan pada item terakhir agar subtotal presisi sesuai dokumen PDF.
        - **Format Ekspor**: Hasil ekspor Excel secara otomatis disusun mulai dari sel **B1** untuk mempermudah integrasi lembar kerja.
        """
    )

st.caption("Catatan: Berkas PDF harus memiliki lapisan teks (text layer). Dokumen hasil pemindaian (scan) berupa citra murni membutuhkan proses OCR terlebih dahulu.")