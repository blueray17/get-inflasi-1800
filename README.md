# Get Data Inflasi — BPS Lampung

Aplikasi Streamlit untuk mengambil data dari Google Sheets dan menghasilkan file Excel terstruktur.

## Instalasi

```bash
pip install -r requirements.txt
```

## Menjalankan

```bash
streamlit run app_inflasi.py
```

## Cara Penggunaan

1. **Buka aplikasi** di browser (otomatis terbuka di http://localhost:8501)
2. **Isi Parameter Waktu**: Tahun (default 2026) dan pilih Bulan
3. **Isi Rentang Kolom**: Pilih kolom awal dan kolom akhir (A–ZZ)
4. **Klik "⚡ Generate Excel"**
5. **Unduh file** `data_inflasi_<tahun>_<bulan>.xlsx`

## Struktur Output Excel

| Kolom | Sumber |
|-------|--------|
| Kode | Kolom A dari sheet |
| Tahun | Input tahun |
| Bulan | Kode bulan (01–12) |
| Kode_Wilayah | Sheet1=1800, Sheet2=1804, Sheet3=1811, Sheet4=1871, Sheet5=1872 |
| B–Z (dll.) | Kolom sesuai input kolom awal–akhir |

## Catatan Penting

- Spreadsheet Google harus bersifat **publik** (Anyone with the link → Viewer)
- Jika spreadsheet privat, gunakan **Service Account JSON** di sidebar
- 5 sheet pertama yang akan diproses

## Cara Membuat Spreadsheet Publik

1. Buka Google Sheets
2. Klik tombol **Share**
3. Ubah akses menjadi **"Anyone with the link"** → **Viewer**
4. Klik **Done**
