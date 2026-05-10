import streamlit as st
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
import io
import re
from itertools import product as iterproduct
import json
import os

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Generator Data Inflasi",
    page_icon="📊",
    layout="centered",
)

# ─────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', sans-serif;
}

.stApp {
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%);
    min-height: 100vh;
}

/* Header */
.main-header {
    text-align: center;
    padding: 2.5rem 0 1.5rem 0;
}
.main-header h1 {
    font-size: 2rem;
    font-weight: 700;
    color: #f1f5f9;
    margin: 0;
    letter-spacing: -0.5px;
}
.main-header p {
    color: #64748b;
    font-size: 0.9rem;
    margin-top: 0.4rem;
}
.badge {
    display: inline-block;
    background: linear-gradient(90deg, #3b82f6, #6366f1);
    color: white;
    font-size: 0.7rem;
    font-weight: 600;
    padding: 3px 10px;
    border-radius: 20px;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-bottom: 0.8rem;
}

/* Card */
.card {
    background: rgba(30, 41, 59, 0.8);
    border: 1px solid rgba(99, 102, 241, 0.2);
    border-radius: 16px;
    padding: 1.5rem;
    margin-bottom: 1.2rem;
    backdrop-filter: blur(10px);
}
.card-title {
    font-size: 0.75rem;
    font-weight: 600;
    color: #6366f1;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    margin-bottom: 1rem;
}

/* Streamlit overrides */
.stTextInput > div > div > input,
.stNumberInput > div > div > input {
    background: rgba(15, 23, 42, 0.8) !important;
    border: 1px solid rgba(99, 102, 241, 0.3) !important;
    border-radius: 8px !important;
    color: #f1f5f9 !important;
    font-family: 'JetBrains Mono', monospace !important;
}
.stSelectbox > div > div {
    background: rgba(15, 23, 42, 0.8) !important;
    border: 1px solid rgba(99, 102, 241, 0.3) !important;
    border-radius: 8px !important;
    color: #f1f5f9 !important;
}
label {
    color: #94a3b8 !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
}

/* Generate button */
.stButton > button {
    width: 100%;
    background: linear-gradient(135deg, #3b82f6, #6366f1) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.75rem 2rem !important;
    font-size: 0.95rem !important;
    font-weight: 600 !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    letter-spacing: 0.5px !important;
    transition: all 0.3s ease !important;
    cursor: pointer !important;
}
.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 25px rgba(99, 102, 241, 0.4) !important;
}

/* Download button */
.stDownloadButton > button {
    width: 100%;
    background: linear-gradient(135deg, #10b981, #059669) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.75rem 2rem !important;
    font-size: 0.95rem !important;
    font-weight: 600 !important;
}

/* Info box */
.info-box {
    background: rgba(59, 130, 246, 0.1);
    border-left: 3px solid #3b82f6;
    border-radius: 0 8px 8px 0;
    padding: 0.8rem 1rem;
    margin: 0.8rem 0;
    color: #93c5fd;
    font-size: 0.83rem;
}
.success-box {
    background: rgba(16, 185, 129, 0.1);
    border-left: 3px solid #10b981;
    border-radius: 0 8px 8px 0;
    padding: 0.8rem 1rem;
    margin: 0.8rem 0;
    color: #6ee7b7;
    font-size: 0.83rem;
}
.error-box {
    background: rgba(239, 68, 68, 0.1);
    border-left: 3px solid #ef4444;
    border-radius: 0 8px 8px 0;
    padding: 0.8rem 1rem;
    margin: 0.8rem 0;
    color: #fca5a5;
    font-size: 0.83rem;
}

.stat-row {
    display: flex;
    gap: 1rem;
    margin-top: 0.8rem;
}
.stat-item {
    flex: 1;
    background: rgba(15, 23, 42, 0.6);
    border-radius: 10px;
    padding: 0.8rem;
    text-align: center;
}
.stat-value {
    font-size: 1.4rem;
    font-weight: 700;
    color: #6366f1;
    font-family: 'JetBrains Mono', monospace;
}
.stat-label {
    font-size: 0.7rem;
    color: #64748b;
    margin-top: 2px;
}

/* Divider */
hr { border-color: rgba(99, 102, 241, 0.15) !important; }

/* Footer */
.footer {
    text-align: center;
    color: #334155;
    font-size: 0.75rem;
    padding: 2rem 0 1rem;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def generate_column_names():
    """Generate Excel-style column names: A, B, ..., Z, AA, AB, ..., ZZ"""
    cols = []
    # Single letters A-Z
    for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        cols.append(c)
    # Double letters AA-ZZ
    for c1 in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        for c2 in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            cols.append(c1 + c2)
    return cols

def col_letter_to_index(col_letter):
    """Convert column letter(s) to 0-based index (A=0, B=1, ..., Z=25, AA=26, ...)"""
    col_letter = col_letter.upper()
    result = 0
    for char in col_letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1  # 0-based

def connect_google_sheet(spreadsheet_url, credentials_json):
    """Connect to Google Sheets using service account credentials"""
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = json.loads(credentials_json)
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(spreadsheet_url)
    return spreadsheet

def get_public_sheet(spreadsheet_url):
    """Connect to public Google Sheet (no auth needed if sheet is public)"""
    # Extract spreadsheet ID from URL
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', spreadsheet_url)
    if not match:
        raise ValueError("URL Google Sheets tidak valid")
    
    spreadsheet_id = match.group(1)
    
    # Try using gspread with anonymous access
    client = gspread.Client(auth=None)
    client.session = gspread.auth.LocalServerFlowWithBackoff  # placeholder
    
    # Use requests-based approach for public sheets
    import requests
    
    sheets_data = []
    for sheet_index in range(5):  # 5 sheets
        # Export sheet as CSV
        csv_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={sheet_index}"
        # We can't use gid=0,1,2... directly; we need actual gids
        # Better: use the sheets API without auth for public sheets
        pass
    
    return sheets_data

def fetch_sheet_as_df(spreadsheet_id, sheet_gid):
    """Fetch a single sheet as DataFrame via CSV export (works for public sheets)"""
    import requests
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={sheet_gid}"
    response = requests.get(url, timeout=30)
    if response.status_code == 200:
        from io import StringIO
        df = pd.read_csv(StringIO(response.text), header=None)
        return df
    else:
        raise Exception(f"Gagal mengambil sheet (HTTP {response.status_code}). Pastikan spreadsheet bersifat publik.")

def get_sheet_gids(spreadsheet_id):
    """Get list of sheet GIDs from Google Sheets"""
    import requests
    # Fetch the HTML of the spreadsheet to extract GIDs
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    response = requests.get(url, timeout=30)
    if response.status_code != 200:
        # Fallback: assume gids are 0,1,2,3,4 (common for new sheets)
        return [0, 1, 2, 3, 4]
    
    # Extract gid values from HTML
    gids = re.findall(r'"gid":(\d+)', response.text)
    # Also try another pattern
    if not gids:
        gids = re.findall(r'gid=(\d+)', response.text)
    
    gids = list(dict.fromkeys(gids))  # deduplicate, preserve order
    gids_int = [int(g) for g in gids[:5]]
    
    if not gids_int:
        gids_int = [0, 1, 2, 3, 4]
    
    return gids_int

KODE_WILAYAH = {
    0: "1800",
    1: "1804",
    2: "1811",
    3: "1871",
    4: "1872",
}

NAMA_WILAYAH = {
    0: "Provinsi Lampung",
    1: "Lampung Barat",
    2: "Lampung Selatan",
    3: "Bandar Lampung",
    4: "Metro",
}

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/15HbcEJwdK9TUo8Wpkgqnfveyp67RLK4B/edit"
SPREADSHEET_ID = "15HbcEJwdK9TUo8Wpkgqnfveyp67RLK4B"

BULAN = [
    ("01", "Januari"), ("02", "Februari"), ("03", "Maret"),
    ("04", "April"), ("05", "Mei"), ("06", "Juni"),
    ("07", "Juli"), ("08", "Agustus"), ("09", "September"),
    ("10", "Oktober"), ("11", "November"), ("12", "Desember"),
]

ALL_COLS = generate_column_names()

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <div class="badge">BPS Provinsi Lampung</div>
    <h1>📊 Generator Data Inflasi</h1>
    <p>Konversi data Google Sheets menjadi file Excel terstruktur</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR - CREDENTIALS (optional)
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Konfigurasi")
    st.markdown("---")
    
    st.markdown("**🔗 Sumber Data**")
    st.markdown(f"""
    <div style="background:rgba(15,23,42,0.8);border-radius:8px;padding:0.6rem 0.8rem;
    border:1px solid rgba(99,102,241,0.2);word-break:break-all;font-size:0.72rem;color:#94a3b8;">
    {SPREADSHEET_URL}
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("**🔑 Autentikasi (Opsional)**")
    st.markdown("<small style='color:#64748b'>Diperlukan jika spreadsheet bersifat privat</small>", unsafe_allow_html=True)
    
    use_credentials = st.toggle("Gunakan Service Account", value=False)
    credentials_json = ""
    if use_credentials:
        credentials_json = st.text_area(
            "Service Account JSON",
            placeholder='{"type": "service_account", ...}',
            height=150,
            help="Paste JSON kredensial Service Account Google"
        )
    else:
        st.info("ℹ️ Mode publik — spreadsheet harus bisa diakses siapa saja", icon=None)
    
    st.markdown("---")
    st.markdown("**📌 Kode Wilayah**")
    for i in range(5):
        st.markdown(f"""
        <div style="display:flex;justify-content:space-between;padding:4px 0;
        border-bottom:1px solid rgba(99,102,241,0.1);font-size:0.78rem;">
            <span style="color:#94a3b8">Sheet {i+1} — {NAMA_WILAYAH[i]}</span>
            <span style="color:#6366f1;font-family:monospace;font-weight:600">{KODE_WILAYAH[i]}</span>
        </div>
        """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# MAIN FORM
# ─────────────────────────────────────────────

# Card: Parameter Waktu
st.markdown('<div class="card"><div class="card-title">📅 Parameter Waktu</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    tahun = st.number_input("Tahun", min_value=2000, max_value=2099, value=2026, step=1)
with col2:
    bulan_options = [f"{kode}. {nama}" for kode, nama in BULAN]
    bulan_selected = st.selectbox("Bulan", bulan_options, index=0)
    kode_bulan = bulan_selected.split(".")[0].strip()
    nama_bulan = bulan_selected.split(". ")[1].strip()
st.markdown('</div>', unsafe_allow_html=True)

# Card: Rentang Kolom
st.markdown('<div class="card"><div class="card-title">📋 Rentang Kolom Data</div>', unsafe_allow_html=True)
st.markdown('<div class="info-box">Pilih kolom awal dan akhir yang akan diambil dari setiap sheet (selain kolom A yang otomatis diambil)</div>', unsafe_allow_html=True)

col3, col4 = st.columns(2)
with col3:
    # Default kolom awal: B
    default_awal = ALL_COLS.index("B") if "B" in ALL_COLS else 1
    kolom_awal = st.selectbox("Kolom Awal", ALL_COLS, index=default_awal)
with col4:
    # Default kolom akhir: Z
    default_akhir = ALL_COLS.index("Z") if "Z" in ALL_COLS else 25
    kolom_akhir = st.selectbox("Kolom Akhir", ALL_COLS, index=default_akhir)

# Validasi rentang kolom
idx_awal = col_letter_to_index(kolom_awal)
idx_akhir = col_letter_to_index(kolom_akhir)

if idx_awal > idx_akhir:
    st.markdown('<div class="error-box">⚠️ Kolom awal harus lebih kecil atau sama dengan kolom akhir</div>', unsafe_allow_html=True)
else:
    jumlah_kolom = idx_akhir - idx_awal + 1
    st.markdown(f'<div class="success-box">✅ Mengambil kolom <strong>{kolom_awal}</strong> s/d <strong>{kolom_akhir}</strong> ({jumlah_kolom} kolom data)</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# Card: Preview output
nama_file = f"data_inflasi_{tahun}_{kode_bulan}.xlsx"
st.markdown(f"""
<div class="card">
    <div class="card-title">📁 Output</div>
    <div style="display:flex;align-items:center;gap:0.8rem;">
        <div style="font-size:2rem;">📄</div>
        <div>
            <div style="color:#f1f5f9;font-weight:600;font-family:'JetBrains Mono',monospace;">{nama_file}</div>
            <div style="color:#64748b;font-size:0.78rem;margin-top:2px;">
                5 sheet → 1 sheet gabungan • Kolom: A + Tahun + Bulan + Kode Wilayah + {kolom_awal}:{kolom_akhir}
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# GENERATE BUTTON
# ─────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
generate_clicked = st.button("⚡ Generate Excel", use_container_width=True)

if generate_clicked:
    if idx_awal > idx_akhir:
        st.error("Kolom awal harus lebih kecil atau sama dengan kolom akhir!")
    else:
        with st.spinner("Mengambil data dari Google Sheets..."):
            try:
                # Get sheet GIDs
                progress = st.progress(0, text="Mendapatkan info sheet...")
                
                try:
                    gids = get_sheet_gids(SPREADSHEET_ID)
                except Exception as e:
                    gids = [0, 1, 2, 3, 4]
                    st.warning(f"Menggunakan GID default (0-4): {e}")
                
                # Ensure we have 5 GIDs
                while len(gids) < 5:
                    gids.append(len(gids))
                
                all_dfs = []
                total_rows = 0
                
                for i in range(5):
                    progress.progress((i + 1) / 6, text=f"Membaca Sheet {i+1} — {NAMA_WILAYAH[i]}...")
                    
                    try:
                        if use_credentials and credentials_json.strip():
                            # Use service account
                            spreadsheet = connect_google_sheet(SPREADSHEET_URL, credentials_json)
                            worksheets = spreadsheet.worksheets()
                            if i < len(worksheets):
                                ws = worksheets[i]
                                raw = ws.get_all_values()
                                df_raw = pd.DataFrame(raw)
                            else:
                                st.warning(f"Sheet {i+1} tidak ditemukan, dilewati.")
                                continue
                        else:
                            # Public sheet via CSV export
                            df_raw = fetch_sheet_as_df(SPREADSHEET_ID, gids[i])
                    except Exception as e:
                        st.warning(f"Sheet {i+1} gagal diambil: {e}")
                        continue
                    
                    if df_raw.empty:
                        st.warning(f"Sheet {i+1} kosong, dilewati.")
                        continue
                    
                    # Pastikan cukup kolom
                    max_needed = max(idx_akhir, 0)
                    if df_raw.shape[1] <= max_needed:
                        # Extend dataframe with empty columns if needed
                        for _ in range(max_needed - df_raw.shape[1] + 1):
                            df_raw[df_raw.shape[1]] = ""
                    
                    # Ambil kolom A (index 0)
                    col_a = df_raw.iloc[:, 0].fillna("")
                    
                    # Ambil kolom rentang yang dipilih
                    actual_end = min(idx_akhir, df_raw.shape[1] - 1)
                    actual_start = min(idx_awal, df_raw.shape[1] - 1)
                    range_cols = df_raw.iloc[:, actual_start:actual_end + 1].copy()
                    
                    # Rename kolom range sesuai huruf
                    range_col_names = ALL_COLS[actual_start:actual_end + 1]
                    range_cols.columns = range_col_names
                    
                    # Bangun DataFrame hasil
                    result_df = pd.DataFrame()
                    result_df["Kode"] = col_a
                    result_df["Tahun"] = str(tahun)
                    result_df["Bulan"] = kode_bulan
                    result_df["Kode_Wilayah"] = KODE_WILAYAH[i]
                    
                    for col_name in range_col_names:
                        result_df[col_name] = range_cols[col_name].values
                    
                    all_dfs.append(result_df)
                    total_rows += len(result_df)
                
                progress.progress(6 / 6, text="Menggabungkan dan menyimpan...")
                
                if not all_dfs:
                    st.error("Tidak ada data yang berhasil diambil dari spreadsheet. Pastikan spreadsheet bersifat publik atau credentials benar.")
                else:
                    # Gabungkan semua sheet
                    final_df = pd.concat(all_dfs, ignore_index=True)
                    
                    # Simpan ke Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='Data Inflasi')
                        
                        # Auto-fit columns (approximate)
                        ws_excel = writer.sheets['Data Inflasi']
                        for column in ws_excel.columns:
                            max_length = 0
                            col_letter = column[0].column_letter
                            for cell in column:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            adjusted_width = min(max_length + 2, 40)
                            ws_excel.column_dimensions[col_letter].width = adjusted_width
                    
                    output.seek(0)
                    progress.empty()
                    
                    # Stats
                    st.markdown(f"""
                    <div class="success-box">
                        ✅ Berhasil! Data siap diunduh.
                    </div>
                    <div class="stat-row">
                        <div class="stat-item">
                            <div class="stat-value">{len(all_dfs)}</div>
                            <div class="stat-label">Sheet Diproses</div>
                        </div>
                        <div class="stat-item">
                            <div class="stat-value">{total_rows:,}</div>
                            <div class="stat-label">Total Baris</div>
                        </div>
                        <div class="stat-item">
                            <div class="stat-value">{final_df.shape[1]}</div>
                            <div class="stat-label">Kolom Output</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Preview
                    with st.expander("👁️ Preview Data (10 baris pertama)", expanded=False):
                        st.dataframe(final_df.head(10), use_container_width=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    # Download button
                    st.download_button(
                        label=f"⬇️ Unduh {nama_file}",
                        data=output.getvalue(),
                        file_name=nama_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                    
            except Exception as e:
                st.error(f"Terjadi kesalahan: {str(e)}")
                st.markdown(f"""
                <div class="error-box">
                    <strong>Troubleshooting:</strong><br>
                    • Pastikan spreadsheet bersifat publik (Anyone with the link can view)<br>
                    • Atau gunakan Service Account JSON di sidebar<br>
                    • Error detail: {str(e)}
                </div>
                """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("""
<div class="footer">
    BPS Provinsi Lampung • Generator Data Inflasi • susenas.my.id
</div>
""", unsafe_allow_html=True)
