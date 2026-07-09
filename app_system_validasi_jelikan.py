# ==========================================================
# IMPORT
# ==========================================================
import streamlit as st
import pandas as pd
from PIL import Image
from io import BytesIO

# LANDING PAGE
try:
    logo = Image.open("_ MDPI Primary Logo.png")
    st.set_page_config(page_title="Sistem Validasi Data Jelikan", page_icon=logo, layout="wide")
except:
    st.set_page_config(page_title="Sistem Validasi Data Jelikan", layout="wide")

# ==========================================================
# STYLE
# ==========================================================
st.markdown("""
<style>
div[data-testid="metric-container"]{
    background:#f8f9fc;
    border-left:6px solid #0E4C92;
    padding:15px;
    border-radius:10px;
    box-shadow:2px 2px 8px rgba(0,0,0,.1);
}
</style>
""", unsafe_allow_html=True)

# ==========================================================
# HEADER
# ==========================================================
col1, col2 = st.columns([0.01, 3])
with col1:
    try:
        st.image(logo, width=180)
    except:
        pass

with col2:
    st.markdown("""
    <h1 style="color:#0E4C92; text-align:center; margin-bottom:0;">Sistem Validasi Data Jelikan</h1>
    <h3 style="color:#555; text-align:center; margin-top:5px;">Yayasan Masyarakat dan Perikanan Indonesia (MDPI)</h3>
    <h4 style="color:#777; text-align:center;">Validasi Data Otomatis</h4>
    """, unsafe_allow_html=True)

st.markdown("---")

# ==========================================================
# INFORMASI
# ==========================================================
st.info("""
### Informasi Aplikasi
Aplikasi ini dapat melakukan validasi otomatis terhadap seluruh data dari Jelikan.
Validasi meliputi:
- Data kosong
- Kesesuaian data
- Penambahan kolom Catatan Validasi

Output berupa file Excel hasil validasi.
Cek datamu sekarang jangan nanti menumpuk!
""")

# ==========================================================
# UPLOAD
# ==========================================================
uploaded_file = st.file_uploader("Upload File Excel", type=["xlsx"])

if uploaded_file is None:
    st.info("Silakan upload file Excel (.xlsx).")
    st.markdown("---")
    st.markdown("""
    <div style='text-align:center;color:#666'>
    <b>Dashboard Validasi Data Jelikan</b>
    <p>Dikembangkan untuk mendukung pengelolaan kualitas data Yayasan MDPI.</p>
    <p style='font-size:12px'>© 2026 Yayasan MDPI. <i>Happy People Many Fish</i>.</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ==========================================================
# MEMBACA EXCEL
# ==========================================================
all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
validated_sheets = {}
total_data_keseluruhan = 0
total_error_keseluruhan = 0

# ==========================================================
# DICTIONARY BATASAN (STANDARD) YFT
# ==========================================================
# Referensi: loin1_panjang: (berat_min, berat_max)
loin_batas = {
    45: (2, 4.1), 46: (2, 3), 47: (2, 4.2), 48: (2, 3.7), 49: (2, 4.9), 50: (2, 3.46),
    51: (2, 3.6), 52: (2, 3.62), 53: (2, 3.51), 54: (2, 3.72), 55: (2, 3.9), 56: (2, 3.75),
    57: (2, 4.1), 58: (2, 4.32), 59: (2, 4.44), 60: (2.1, 4.6), 61: (2.1, 4.88), 62: (2.1, 5),
    63: (2.1, 5.1), 64: (2.2, 5.4), 65: (2.3, 5.74), 66: (2.7, 6.3), 67: (2.5, 6.3), 68: (2.7, 6.7),
    69: (2.9, 6.84), 70: (3, 7.4), 71: (3.2, 7.58), 72: (3.1, 7.63), 73: (3.2, 8.37), 74: (3.4, 8.5),
    75: (3.5, 8.6), 76: (3.8, 9.1), 77: (3.9, 9.7), 78: (4.1, 9.74), 79: (4.3, 9.92), 80: (4.5, 10.78),
    81: (4.7, 11.12), 82: (5, 11.23), 83: (4.9, 11.34), 84: (5.2, 12), 85: (5.8, 12.6), 86: (5.9, 13.4),
    87: (5.8, 12.9), 88: (5.8, 13.8), 89: (6.1, 13.6), 90: (7.26, 14.3), 91: (6.8, 14.58), 92: (7.7, 14.8),
    93: (8, 14.6), 94: (8.9, 15), 95: (7.5, 14), 96: (8.2, 15.01), 97: (7.6, 15.3), 98: (8.7, 15.5),
    99: (9.3, 15.1), 100: (10.1, 15.1)
}

# Referensi: panjang: (berat_min, berat_max)
bp_batas = {
    80: (10, 16.28), 81: (10, 16.36), 82: (10, 19.4), 83: (10, 14.29), 84: (10, 16), 85: (10, 19.78),
    86: (10, 17.85), 87: (10, 18.01), 88: (10, 19), 89: (10, 19.85), 90: (10, 21), 91: (10, 18.45),
    92: (10, 19), 93: (10, 19.1), 94: (11, 19.52), 95: (11, 21), 96: (11, 22), 97: (10, 21), 98: (12, 25),
    99: (11.3, 22.3), 100: (10, 24), 101: (13, 27), 102: (11.23, 27), 103: (14, 25.85), 104: (15, 28.15),
    105: (14, 27.45), 106: (14, 29.1), 107: (16, 28), 108: (13.63, 30), 109: (14, 34), 110: (15, 34.81),
    111: (16, 30), 112: (14, 34.01), 113: (17, 35.05), 114: (19, 37), 115: (17, 36), 116: (15, 37),
    117: (17, 39), 118: (18, 40.05), 119: (17, 39.3), 120: (20, 41), 121: (20, 42), 122: (22, 44),
    123: (20.22, 42), 124: (26, 45.81), 125: (26, 46.06), 126: (21, 45.45), 127: (29, 48), 128: (27, 50),
    129: (27.5, 50), 130: (29, 49), 131: (31, 48), 132: (31.05, 52.8), 133: (30, 50), 134: (29, 56.15),
    135: (33, 57), 136: (31, 59), 137: (34, 52), 138: (38, 53), 139: (36, 55.75), 140: (39, 59),
    141: (39, 61), 142: (40, 65), 143: (43, 64), 144: (44, 68), 145: (44, 66), 146: (46, 69),
    147: (43.37, 68), 148: (49, 69), 149: (47, 70), 150: (53, 71), 151: (51, 73), 152: (54, 72),
    153: (55, 71), 154: (51.24, 78), 155: (57, 85), 156: (56, 82), 157: (54, 83), 158: (60, 79),
    159: (64, 83), 160: (64, 90), 161: (64, 90), 162: (64, 81), 163: (70, 90), 164: (63, 86),
    165: (70, 89), 166: (72, 90), 167: (70, 85), 168: (81, 92), 169: (85, 97), 170: (80, 87)
}

# ==========================================================
# VALIDASI
# ==========================================================
for sheet_name, df in all_sheets.items():
    catatan = []
    for _, row in df.iterrows():
        error = []
        
        # ======================================================
        # 1. TEMPAT MENAMBAH LOGIKA KHUSUS PER SHEET
        # ======================================================
        if sheet_name == "1-Trip Info":
            # --- Enumerator ---
            enum1 = str(row.get("enumerator1", "")).strip()
            enum2 = str(row.get("enumerator2", "")).strip()
            if (pd.isna(row.get("enumerator1")) or enum1 == "" or enum1.lower() == "nan") and \
               (pd.isna(row.get("enumerator2")) or enum2 == "" or enum2.lower() == "nan"):
                error.append("enumerator 1 kosong")
            
            # --- Satuan Trip & Lama Jam/Hari ---
            satuan = str(row.get("satuan_trip", "")).strip().upper()
            lama_jam = pd.to_numeric(row.get("lama_jam"), errors='coerce')
            lama_hari = pd.to_numeric(row.get("lama_hari"), errors='coerce')
            hari_mancing = pd.to_numeric(row.get("jumlah_hari_memancing"), errors='coerce')

            is_pengecualian_valid = False
            if satuan == 'J':
                if pd.notna(lama_jam) and lama_jam <= 24:
                    if pd.notna(lama_hari) and lama_hari == 0:
                        if pd.notna(hari_mancing) and hari_mancing == 1:
                            is_pengecualian_valid = True

            if satuan == 'J':
                if pd.notna(lama_jam) and lama_jam > 23:
                    if not is_pengecualian_valid:
                        error.append("lama_jam sudah terhitung 1 hari")
                        
            elif satuan == 'H':
                if pd.notna(lama_jam) and lama_jam != 0:
                    error.append("lama_jam harus 0 karena satuan trip H")
                if pd.isna(lama_hari) or lama_hari == 0:
                    error.append("lama_hari tidak boleh 0 atau kosong karena satuan trip H")
            
            # --- Validasi Jumlah Hari Memancing ---
            if pd.notna(hari_mancing) and pd.notna(lama_hari):
                if hari_mancing > lama_hari:
                    if not is_pengecualian_valid:
                        error.append("Jumlah Hari Memancing Lebih Lama dari Jumlah Hari Trip")
            
            # --- Penggunaan Es ---
            es = str(row.get("penggunaan_es", "")).strip()
            if es == "" or es.lower() == "nan":
                error.append("Apakah nelayan tidak membawa ES? Jika Ya beri catatan di deskripsi")
            
            # --- Kapasitas & Panjang ---
            for col in ["kapasitas_kapal", "panjang_kapal", "kapasitas_mesin"]:
                val = pd.to_numeric(row.get(col), errors='coerce')
                if pd.isna(val) or val == 0:
                    error.append(f"{col} tidak boleh 0 dan kosong")
            
            # --- Alat Tangkap ---
            alat = str(row.get("k_alattangkap", "")).strip().upper()
            if alat != "HL":
                error.append("Alat tangkap bukan HL")
            
            # --- Rumpon & Teknik Pencarian & Jumlah ---
            rumpon = str(row.get("rumpon", "")).strip().upper()
            teknik = str(row.get("teknik_pencarian_lokasi_tuna", "")).strip().lower()
            jml_rumpon = pd.to_numeric(row.get("jumlah rumpon"), errors='coerce')
            
            if rumpon == 'F':
                if "rumpon" not in teknik:
                    error.append("kolom rumpon dan teknik pencarian tidak konsisten")
                if pd.isna(jml_rumpon) or jml_rumpon == 0:
                    error.append("jumlah rumpon tidak sesuai (rumpon F tidak boleh kosong/0)")
                    
            elif rumpon == 'N':
                if "rumpon" in teknik:
                    error.append("kolom rumpon dan teknik pencarian tidak konsisten")
                if pd.notna(jml_rumpon) and jml_rumpon != 0:
                    error.append("jumlah rumpon tidak sesuai (rumpon N harusnya jumlah rumpon 0/kosong)")
                    
            elif rumpon == 'X':
                if "rumpon" not in teknik or teknik.replace(" ", "") == "rumpon":
                    error.append("kolom rumpon dan teknik pencarian tidak konsisten")
                if pd.isna(jml_rumpon) or jml_rumpon == 0:
                    error.append("jumlah rumpon tidak sesuai (rumpon X tidak boleh kosong/0)")
            
            # --- Kedalaman ---
            k_min = pd.to_numeric(row.get("kedalaman min"), errors='coerce')
            k_max = pd.to_numeric(row.get("kedalaman max"), errors='coerce')
            
            if pd.isna(k_min) or pd.isna(k_max):
                error.append("kedalaman min dan max tidak boleh kosong")
            else:
                if k_min > k_max:
                    error.append("kedalaman min lebih besar dari kedalaman max")
                if not (0 <= k_min <= 300) or not (0 <= k_max <= 300):
                    error.append("nilai kedalaman mencurigakan (di luar 0 - 300)")
            
            # --- Palka ---
            for col in ["jumlah palka", "kapasitas palka"]:
                val_str = str(row.get(col, "")).strip()
                if val_str == "" or val_str.lower() == "nan":
                    error.append(f"{col} boleh 0 tapi tidak boleh kosong")
                    
        # ======================================================
        # TAMBAHAN: VALIDASI SHEET 7-LargeFish
        # ======================================================
        elif sheet_name == "7-LargeFish":
            k_species = str(row.get("k_species", "")).strip().upper()
            
            if k_species == "YFT":
                b_val = pd.to_numeric(row.get("berat"), errors='coerce')
                p_val = pd.to_numeric(row.get("panjang"), errors='coerce')
                lb_val = pd.to_numeric(row.get("loin1_berat"), errors='coerce')
                lp_val = pd.to_numeric(row.get("loin1_panjang"), errors='coerce')
                
                # Cek apakah ada indikasi data loin (nilai lebih dari 0)
                ada_indikasi_loin = (pd.notna(lb_val) and lb_val > 0) or (pd.notna(lp_val) and lp_val > 0)
                
                if ada_indikasi_loin:
                    # ----------------------------------------------------
                    # FOKUS VALIDASI DATA LOIN
                    # ----------------------------------------------------
                    if pd.isna(lb_val) or pd.isna(lp_val) or lb_val <= 0 or lp_val <= 0:
                        error.append("Data loin tidak lengkap (loin1_berat dan loin1_panjang keduanya harus diisi jika salah satu ada nilainya)")
                    else:
                        lp_int = int(lp_val)
                        if lp_int in loin_batas:
                            lb_min, lb_max = loin_batas[lp_int]
                            if not (lb_min <= lb_val <= lb_max):
                                error.append(f"loin1_berat ({lb_val}) tidak sesuai standard loin1_panjang ({lp_val})")
                        else:
                            error.append(f"loin1_panjang ({lp_val}) di luar batas pengecekan referensi YFT")
                
                else:
                    # ----------------------------------------------------
                    # JIKA BUKAN LOIN, FOKUS VALIDASI IKAN UTUH
                    # ----------------------------------------------------
                    # Loin kosong/0 tidak masalah, asalkan ada berat & panjang
                    has_bp = (pd.notna(b_val) and b_val > 0) and (pd.notna(p_val) and p_val > 0)
                    
                    if not has_bp:
                        error.append("Harus mengisi (berat & panjang) utuh, atau mengisi data loin (jika ikan berupa loin)")
                    else:
                        p_int = int(p_val)
                        if p_int in bp_batas:
                            b_min, b_max = bp_batas[p_int]
                            if not (b_min <= b_val <= b_max):
                                error.append(f"Berat ikan ({b_val}) tidak sesuai standard panjang ({p_val})")
                        else:
                            error.append(f"Panjang ({p_val}) di luar batas pengecekan referensi berat-panjang YFT")

        # ======================================================
        # 2. VALIDASI KOLOM LAINNYA (CEK KOSONG DEFAULT)
        # ======================================================
        for kolom in df.columns:
            # DAFTAR PENGECUALIAN (SANGAT PENTING)
            kolom_dilewati = []
            if sheet_name == "1-Trip Info":
                kolom_dilewati = [
                    "enumerator1", "enumerator2", "penggunaan_es", 
                    "kapasitas_kapal", "panjang_kapal", "kapasitas_mesin",
                    "kedalaman min", "kedalaman max", "jumlah palka", "kapasitas palka",
                    "handline_troll", "alat_tangkap_lain", "tipe_template", "tlc",
                    "spot trace", "flywire", "pds", "gps", "gps merk",
                    "Kapal andon", "asal andon", "Fairtrade name", 
                    "Deskripsi Kesesuaian", "Deskripsi Kendala Nelayan"
                ]
            elif sheet_name == "7-LargeFish":
                # Kolom terkait berat dan panjang diperbolehkan kosong selama memenuhi salah satu logika
                kolom_dilewati = [
                    "berat", "panjang", "loin1_berat", "loin1_panjang", 
                    "karkas_panjang", "karkas_berat"
                ]

            if kolom in kolom_dilewati:
                continue
                
            # Cek kosong untuk kolom-kolom selain pengecualian di atas
            if pd.isna(row[kolom]) or str(row[kolom]).strip() == "":
                error.append(f"{kolom} kosong")
                
        # ======================================================
        # Menyusun Catatan
        # ======================================================
        if len(error) == 0:
            catatan.append("Data Valid")
        else:
            catatan.append("; ".join(error))
            
    df["Catatan Validasi"] = catatan
    validated_sheets[sheet_name] = df
    total_data_keseluruhan += len(df)
    total_error_keseluruhan += (df["Catatan Validasi"] != "Data Valid").sum()

# ==========================================================
# PREVIEW
# ==========================================================
st.subheader("Preview Hasil Validasi")

pilih_sheet = st.selectbox(
    "Pilih Sheet",
    list(validated_sheets.keys())
)

df_tampil = validated_sheets[pilih_sheet]

total_data_sheet = len(df_tampil)
total_error_sheet = (df_tampil["Catatan Validasi"] != "Data Valid").sum()
total_valid_sheet = total_data_sheet - total_error_sheet

if total_data_sheet > 0:
    persen_error_sheet = (total_error_sheet / total_data_sheet) * 100
    persen_valid_sheet = (total_valid_sheet / total_data_sheet) * 100
else:
    persen_error_sheet = 0.0
    persen_valid_sheet = 0.0

sc1, sc3, sc2 = st.columns(3)
sc1.metric("Total Data", f"{total_data_sheet:,}")
sc2.metric("Data Bermasalah", f"{total_error_sheet:,} ({persen_error_sheet:.1f}%)")
sc3.metric("Data Valid", f"{total_valid_sheet:,} ({persen_valid_sheet:.1f}%)")

st.dataframe(df_tampil, use_container_width=True)

# ==========================================================
# RINGKASAN
# ==========================================================
st.subheader("Catatan Validasi Keseluruhan")

if total_error_sheet == 0:
    st.success("Bagus kawand! Seluruh data dari semua sheet ini valid.")
else:
    st.warning(
        f"⚠ Ditemukan {total_error_sheet:,} data bermasalah secara keseluruhan dari sheet ini. "
        f"Silakan cek kolom 'Catatan Validasi' pada tabel di atas untuk detailnya atau download filenya."
    )

# ==========================================================
# DOWNLOAD
# ==========================================================
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    for sheet_name, df in validated_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
output.seek(0)

st.download_button(
    "Download Hasil Validasi",
    output,
    "Hasil_Validasi.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ==========================================================
# FOOTER
# ==========================================================
st.markdown("---")
st.markdown("""
<div style='text-align:center;color:#666'>
<b>Dashboard Validasi Data Jelikan</b>
<p>Dikembangkan untuk mendukung kualitas data perikanan tuna skala kecil dampingan MDPI.</p>
<p style='font-size:12px'>
© 2026 Yayasan MDPI. <i>Happy People Many Fish</i>.
</p>
</div>
""", unsafe_allow_html=True)