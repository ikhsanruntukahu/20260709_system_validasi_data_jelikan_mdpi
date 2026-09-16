# --- Import Library ---
import streamlit as st
import pandas as pd
import calendar
from PIL import Image
from io import BytesIO

# --- Landing Page ---
try:
    logo = Image.open("_ MDPI Primary Logo.png")
    st.set_page_config(page_title="Sistem Validasi Data Jelikan", page_icon=logo, layout="wide")
except:
    st.set_page_config(page_title="Sistem Validasi Data Jelikan", layout="wide")

# --- Kustomisasi CSS ---
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

# --- Header ---
col1, col2, col3 = st.columns([1, 4, 1])
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

with col3:
    st.empty()

st.markdown("---")

# --- Informasi Aplikasi ---
st.info("""
### Informasi Aplikasi
Aplikasi ini dapat melakukan validasi otomatis terhadap seluruh data dari Jelikan.
Validasi meliputi:
- Data kosong
- Kesesuaian data
- Penambahan kolom Catatan Validasi

Output berupa file Excel hasil validasi.
Cek datamu sekarang jangan tunda-tunda!
""")

# --- Upload File ---
uploaded_file = st.file_uploader("Upload File Excel", type=["xlsx"])

if uploaded_file is None:
    st.info("Silakan upload file Excel (.xlsx) hasil download dari jelikan (ExtractDataSampling.xlsx).")
    st.markdown("---")
    st.markdown("""
    <div style='text-align:center;color:#666'>
    <b>Dashboard Validasi Data Jelikan</b>
    <p>Dikembangkan untuk mendukung pengelolaan kualitas data Yayasan MDPI.</p>
    <p style='font-size:12px'>© 2026 Yayasan MDPI. <i>Happy People Many Fish</i>.</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# --- Baca File Excel ---
all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
validated_sheets = {}
total_data_keseluruhan = 0
total_error_keseluruhan = 0

# --- Referensi Batas Berat & Panjang ---
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

bet_bp_batas = {
    74: (10, 11), 75: (10, 12), 76: (10, 11), 77: (10, 11), 78: (10, 12), 
    79: (10, 12), 80: (10, 13), 81: (10, 15), 82: (10, 13), 83: (11, 14), 
    84: (10, 17), 85: (10, 14), 86: (10, 15), 87: (11, 16), 88: (11, 16), 
    89: (11, 20), 90: (11, 20), 91: (13, 20), 92: (13, 18), 93: (13, 18), 
    94: (14, 19), 95: (13, 20), 96: (14, 20), 97: (14, 20), 98: (13, 24), 
    99: (17, 21), 100: (19, 24), 101: (18, 22), 102: (17, 23), 103: (18, 25), 
    104: (19, 27), 105: (11, 30), 106: (20, 28), 107: (22, 26), 108: (21, 29), 
    109: (22, 28), 110: (21, 34), 111: (23, 34), 112: (20, 33), 113: (24, 35), 
    114: (25, 38), 115: (27, 33), 116: (26, 35), 117: (29, 38), 118: (28, 37), 
    119: (26, 39), 120: (28, 39), 121: (32, 44), 122: (32, 43), 123: (29, 40), 
    124: (29, 42), 125: (35, 43), 126: (34, 42), 127: (38, 45), 128: (27, 46), 
    129: (37, 50), 130: (40, 49), 131: (41, 48), 132: (41, 54), 133: (37, 52), 
    134: (41, 54), 135: (41, 56), 136: (42, 55), 137: (46, 56), 138: (43, 57), 
    139: (48, 61), 140: (45, 61), 141: (51, 68), 142: (50, 65), 143: (48, 63), 
    144: (53, 69), 145: (55, 69), 146: (55, 74), 147: (57, 70), 148: (57, 73), 
    149: (58, 71), 150: (55, 78), 151: (57, 83), 152: (63, 79), 153: (64, 75), 
    154: (61, 77), 155: (63, 78), 156: (67, 81), 157: (68, 87), 158: (69, 83), 
    159: (66, 88), 160: (72, 88), 161: (69, 86), 162: (77, 91), 163: (76, 94), 
    164: (82, 95), 165: (79, 95), 166: (80, 95), 167: (90, 100), 168: (86, 88), 
    169: (85, 92), 170: (94, 104), 171: (99, 105), 172: (99, 99), 173: (112, 112)
}

alb_bp_batas = {
    74: (11, 11), 77: (11, 11), 78: (11, 11), 79: (14, 14), 80: (14, 14),
    81: (14, 15), 82: (10, 11), 84: (11, 12), 85: (12, 16), 86: (15, 15),
    87: (12, 14), 88: (12, 14), 89: (13, 15), 90: (13, 15), 91: (13, 15),
    92: (14, 16), 93: (15, 17), 94: (15, 17), 95: (16, 18), 96: (16, 18),
    97: (17, 19), 98: (18, 20), 99: (18, 20), 100: (18, 20), 101: (18, 22),
    102: (19, 23), 103: (21, 23), 104: (21, 23), 105: (21, 23), 106: (21, 23),
    107: (22, 26), 108: (24, 25), 109: (21, 24), 110: (24, 25), 111: (26, 28),
    112: (27, 27), 113: (27, 27), 114: (24, 24), 115: (24, 25), 116: (27, 27),
    117: (30, 30), 119: (28, 28), 120: (30, 33), 122: (32, 32), 123: (33, 33),
    125: (35, 35), 129: (38, 39), 130: (40, 40), 132: (41, 41), 140: (51, 51),
    143: (54, 54), 144: (50, 50), 147: (57, 57), 148: (59, 60), 163: (77, 77),
    164: (80, 80)
}

# --- Proses Validasi Tiap Sheet ---
for sheet_name, df in all_sheets.items():
    catatan = []
    for _, row in df.iterrows():
        error = []
        
        # --- Logika Khusus: 1-Trip Info ---
        if sheet_name == "1-Trip Info":
            enum1 = str(row.get("enumerator1", "")).strip()
            enum2 = str(row.get("enumerator2", "")).strip()
            if (pd.isna(row.get("enumerator1")) or enum1 == "" or enum1.lower() == "nan") and \
               (pd.isna(row.get("enumerator2")) or enum2 == "" or enum2.lower() == "nan"):
                error.append("enumerator 1 kosong")
            
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
            
            if pd.notna(hari_mancing) and pd.notna(lama_hari):
                if hari_mancing > lama_hari:
                    if not is_pengecualian_valid:
                        error.append("Jumlah Hari Memancing Lebih Lama dari Jumlah Hari Trip")
            
            es = str(row.get("penggunaan_es", "")).strip()
            if es == "" or es.lower() == "nan":
                error.append("Apakah nelayan tidak membawa ES? Jika Ya beri catatan di deskripsi")
            
            for col in ["kapasitas_kapal", "panjang_kapal", "kapasitas_mesin"]:
                val = pd.to_numeric(row.get(col), errors='coerce')
                if pd.isna(val) or val == 0:
                    error.append(f"{col} tidak boleh 0 dan kosong")
            
            alat = str(row.get("k_alattangkap", "")).strip().upper()
            if alat != "HL":
                error.append("Alat tangkap bukan HL")
            
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
            
            k_min = pd.to_numeric(row.get("kedalaman min"), errors='coerce')
            k_max = pd.to_numeric(row.get("kedalaman max"), errors='coerce')
            
            if pd.isna(k_min) or pd.isna(k_max):
                error.append("kedalaman min dan max tidak boleh kosong")
            else:
                if k_min > k_max:
                    error.append("kedalaman min lebih besar dari kedalaman max")
                if not (0 <= k_min <= 300) or not (0 <= k_max <= 300):
                    error.append("nilai kedalaman mencurigakan (di luar 0 - 300)")
            
            for col in ["jumlah palka", "kapasitas palka"]:
                val_str = str(row.get(col, "")).strip()
                if val_str == "" or val_str.lower() == "nan":
                    error.append(f"{col} boleh 0 tapi tidak boleh kosong")
                    
        # --- Logika Khusus: 7-LargeFish ---
        elif sheet_name == "7-LargeFish":
            k_species = str(row.get("k_species", "")).strip().upper()
            
            if k_species == "YFT":
                b_val = pd.to_numeric(row.get("berat"), errors='coerce')
                p_val = pd.to_numeric(row.get("panjang"), errors='coerce')
                lb_val = pd.to_numeric(row.get("loin1_berat"), errors='coerce')
                lp_val = pd.to_numeric(row.get("loin1_panjang"), errors='coerce')
                
                ada_indikasi_loin = (pd.notna(lb_val) and lb_val > 0) or (pd.notna(lp_val) and lp_val > 0)
                
                if ada_indikasi_loin:
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

            elif k_species == "BET":
                b_val = pd.to_numeric(row.get("berat"), errors='coerce')
                p_val = pd.to_numeric(row.get("panjang"), errors='coerce')
                lb_val = pd.to_numeric(row.get("loin1_berat"), errors='coerce')
                lp_val = pd.to_numeric(row.get("loin1_panjang"), errors='coerce')
                kb_val = pd.to_numeric(row.get("karkas_berat"), errors='coerce')
                kp_val = pd.to_numeric(row.get("karkas_panjang"), errors='coerce')

                has_bp = (pd.notna(b_val) and b_val > 0) and (pd.notna(p_val) and p_val > 0)
                has_loin = (pd.notna(lb_val) and lb_val > 0) and (pd.notna(lp_val) and lp_val > 0)
                has_karkas = (pd.notna(kb_val) and kb_val > 0) and (pd.notna(kp_val) and kp_val > 0)

                if not (has_bp or has_loin or has_karkas):
                    error.append("Harus melengkapi minimal salah satu pasangan data (berat & panjang) ATAU (loin1) ATAU (karkas)")
                
                if has_bp:
                    p_int = int(p_val)
                    if p_int in bet_bp_batas:
                        b_min, b_max = bet_bp_batas[p_int]
                        if not (b_min <= b_val <= b_max):
                            error.append(f"Berat ikan BET ({b_val}) tidak sesuai standard panjang ({p_val})")
                    else:
                        error.append(f"Panjang BET ({p_val}) di luar batas referensi berat-panjang")
            
            elif k_species == "ALB":
                b_val = pd.to_numeric(row.get("berat"), errors='coerce')
                p_val = pd.to_numeric(row.get("panjang"), errors='coerce')
                lb_val = pd.to_numeric(row.get("loin1_berat"), errors='coerce')
                lp_val = pd.to_numeric(row.get("loin1_panjang"), errors='coerce')
                kb_val = pd.to_numeric(row.get("karkas_berat"), errors='coerce')
                kp_val = pd.to_numeric(row.get("karkas_panjang"), errors='coerce')

                has_bp = (pd.notna(b_val) and b_val > 0) and (pd.notna(p_val) and p_val > 0)
                has_loin = (pd.notna(lb_val) and lb_val > 0) and (pd.notna(lp_val) and lp_val > 0)
                has_karkas = (pd.notna(kb_val) and kb_val > 0) and (pd.notna(kp_val) and kp_val > 0)

                if not (has_bp or has_loin or has_karkas):
                    error.append("Harus melengkapi minimal salah satu pasangan data (berat & panjang) ATAU (loin1) ATAU (karkas)")
                
                if has_bp:
                    p_int = int(p_val)
                    if p_int in alb_bp_batas:
                        b_min, b_max = alb_bp_batas[p_int]
                        if not (b_min <= b_val <= b_max):
                            error.append(f"Berat ikan ALB ({b_val}) tidak sesuai standard panjang ({p_val})")
                    else:
                        error.append(f"Panjang ALB ({p_val}) di luar batas referensi berat-panjang")

        # --- Logika Khusus: 2-Umpan Info ---
        elif sheet_name == "2-Umpan Info":
            kategori_umpan = str(row.get("kategori_umpan", "")).strip().upper()
            t_umpan = pd.to_numeric(row.get("total_umpan"), errors='coerce')
            e_umpan = pd.to_numeric(row.get("estimasi_umpan"), errors='coerce')
            
            has_total = pd.notna(t_umpan) and t_umpan > 0
            has_estimasi = pd.notna(e_umpan) and e_umpan > 0
            
            if kategori_umpan in ['A', 'B', 'C', 'D', 'E', 'G']:
                if has_total and has_estimasi:
                    error.append("pilih berat umpan salah satu (total ril atau estimasi)")
                elif not has_total and not has_estimasi:
                    error.append("berat umpan kosong")
                    
        # --- Logika Khusus: Bycatch Info ---
        elif sheet_name in ["3-Bycatch Info", "Bycatch Info"]:
            b_val = pd.to_numeric(row.get("berat"), errors='coerce')
            if pd.notna(b_val):
                if b_val > 290 or b_val < 0.005:
                    error.append("apakah berat sudah sesuai ?")
            
            harga_val = pd.to_numeric(row.get("harga_per_kg"), errors='coerce')
            if pd.notna(harga_val):
                if harga_val < 500 or harga_val >= 200000:
                    error.append("apakah harga sudah sesuai ?")
                    
        # --- Logika Khusus: 8-ETP ---
        elif sheet_name == "8-ETP":
            interaksi_val = str(row.get("interaksi", "")).strip().lower()
            jml_interaksi = pd.to_numeric(row.get("jml_interaksi"), errors='coerce')
            jml_interaksi_clean = 0 if pd.isna(jml_interaksi) else jml_interaksi
            
            cols_to_check = [
                'jml_interaksi', 'jml_didaratkan', 'didaratkan_mati', 'didaratkan_luka_serius', 
                'didaratkan_luka_ringan', 'didaratkan_tidak_terluka', 'didaratkan_tidak_tahu', 
                'tidak_didaratkan_mati', 'tidak_didaratkan_luka_serius', 'tidak_didaratkan_luka_ringan', 
                'tidak_didaratkan_tidak_terluka', 'tidak_didaratkan_tidak_tahu', 'dibuang', 
                'dimakan', 'dijual', 'diumpan', 'lokasi_interaksi'
            ]
            
            if interaksi_val == "tidak":
                is_invalid = False
                for col in cols_to_check:
                    val = row.get(col)
                    if col == 'lokasi_interaksi':
                        val_str = str(val).strip().lower()
                        if val_str not in ["", "nan", "0", "0.0", "none"]:
                            is_invalid = True
                            break
                    else:
                        val_num = pd.to_numeric(val, errors='coerce')
                        if pd.notna(val_num) and val_num != 0:
                            is_invalid = True
                            break
                            
                if is_invalid:
                    error.append("rincian interaksi tidak sesuai (harus 0/kosong saat interaksi tidak)")
                    
            elif interaksi_val == "ya":
                if pd.isna(jml_interaksi) or jml_interaksi == 0:
                    error.append("jml_interaksi tidak boleh 0 atau kosong")
                else:
                    jml_didaratkan = pd.to_numeric(row.get("jml_didaratkan"), errors='coerce')
                    jml_didaratkan_clean = 0 if pd.isna(jml_didaratkan) else jml_didaratkan

                    cols_didaratkan = [
                        'didaratkan_mati', 'didaratkan_luka_serius', 'didaratkan_luka_ringan', 
                        'didaratkan_tidak_terluka', 'didaratkan_tidak_tahu'
                    ]
                    sum_didaratkan = 0
                    for c in cols_didaratkan:
                        val = pd.to_numeric(row.get(c), errors='coerce')
                        sum_didaratkan += val if pd.notna(val) else 0
                        
                    if sum_didaratkan != jml_didaratkan_clean:
                        error.append("total didaratkan tidak sesuai")
                        
                    cols_dimanfaatkan = ['dibuang', 'dimakan', 'dijual', 'diumpan']
                    sum_dimanfaatkan = 0
                    for c in cols_dimanfaatkan:
                        val = pd.to_numeric(row.get(c), errors='coerce')
                        sum_dimanfaatkan += val if pd.notna(val) else 0
                        
                    if sum_dimanfaatkan != jml_didaratkan_clean:
                        error.append("total dimanfaatkan tidak sesuai")

        # --- Logika Khusus: 9-Unloading ---
        elif sheet_name in ["9-Unloading", "Unloading"]:
            bulan_val = pd.to_numeric(row.get("bulan"), errors='coerce')
            tanggal_val = pd.to_numeric(row.get("tanggal"), errors='coerce')
            tahun_val = pd.to_numeric(row.get("tahun"), errors='coerce')

            bulan_is_valid = False
            
            # Validasi nilai bulan (1-12)
            if pd.notna(bulan_val):
                if 1 <= bulan_val <= 12 and (bulan_val % 1 == 0):
                    bulan_is_valid = True
                else:
                    error.append("bulan harus diisi nilai 1-12")

            # Validasi nilai tanggal (1-31 & penyesuaian jumlah hari per bulan termasuk kabisat)
            if pd.notna(tanggal_val):
                if not (1 <= tanggal_val <= 31 and (tanggal_val % 1 == 0)):
                    error.append("tanggal harus diisi nilai 1-31")
                elif bulan_is_valid:
                    thn = int(tahun_val) if (pd.notna(tahun_val) and (tahun_val % 1 == 0) and tahun_val > 0) else 2026
                    _, max_hari = calendar.monthrange(thn, int(bulan_val))
                    if int(tanggal_val) > max_hari:
                        error.append(f"tanggal bulan {int(bulan_val)} tahun {thn} maksimal {max_hari}")

            # Validasi WPP tidak boleh bernilai 0
            wpp_raw = row.get("wpp")
            if pd.notna(wpp_raw):
                wpp_str = str(wpp_raw).strip()
                wpp_num = pd.to_numeric(wpp_raw, errors='coerce')
                if wpp_str in ["0", "0.0"] or wpp_num == 0:
                    error.append("wpp tidak boleh bernilai 0")

        # --- Cek Kolom Kosong Berdasarkan Pengecualian ---
        for kolom in df.columns:
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
                kolom_dilewati = [
                    "berat", "panjang", "loin1_berat", "loin1_panjang", 
                    "karkas_panjang", "karkas_berat"
                ]

            elif sheet_name == "2-Umpan Info":
                kategori_umpan_cek = str(row.get("kategori_umpan", "")).strip().upper()
                if kategori_umpan_cek == 'F':
                    kolom_dilewati = [
                        "grid1", "grid2", "total_umpan", "estimasi_umpan", 
                        "alattangkap_umpan", "domestic_or_import", "pengadaan_umpan"
                    ]
                else:
                    kolom_dilewati = [
                        "total_umpan", "estimasi_umpan", 
                        "domestic_or_import", "pengadaan_umpan"
                    ]

            elif sheet_name == "8-ETP":
                kolom_dilewati = [
                    'k_species', 'jml_interaksi', 'jml_didaratkan', 'didaratkan_mati', 'didaratkan_luka_serius', 
                    'didaratkan_luka_ringan', 'didaratkan_tidak_terluka', 'didaratkan_tidak_tahu', 
                    'tidak_didaratkan_mati', 'tidak_didaratkan_luka_serius', 'tidak_didaratkan_luka_ringan', 
                    'tidak_didaratkan_tidak_terluka', 'tidak_didaratkan_tidak_tahu', 'dibuang', 
                    'dimakan', 'dijual', 'diumpan', 'lokasi_interaksi'
                ]

            elif sheet_name in ["9-Unloading", "Unloading"]:
                kolom_wajib = ["n_tpi", "n_perusahaan", "nama_kapal", "tempat", "tahun", "bulan", "tanggal", "wpp"]
                kolom_dilewati = [c for c in df.columns if c not in kolom_wajib]

            # Lewati pengecekan jika kolom ada di daftar pengecualian
            if kolom in kolom_dilewati:
                continue
                
            # Jika tidak dilewati, pastikan tidak kosong
            if pd.isna(row[kolom]) or str(row[kolom]).strip() == "":
                error.append(f"data {kolom} kosong")
            
        # --- Susun Catatan Validasi (Alasan di depan, "(data tidak valid)" di belakang) ---
        if len(error) == 0:
            catatan.append("Data Valid")
        else:
            catatan.append("; ".join(error) + " (data tidak valid)")
            
    df["Catatan Validasi"] = catatan
    validated_sheets[sheet_name] = df
    total_data_keseluruhan += len(df)
    total_error_keseluruhan += (df["Catatan Validasi"] != "Data Valid").sum()

# --- Preview Hasil ---
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

# --- Ringkasan Keseluruhan ---
st.subheader("Catatan Validasi Keseluruhan")

if total_error_sheet == 0:
    st.success("Kerja Bagus kawand...! Seluruh data dari sheet ini valid.")
else:
    st.warning(
        f"⚠ Ditemukan {total_error_sheet:,} data bermasalah secara keseluruhan dari sheet ini. "
        f"Silakan cek kolom 'Catatan Validasi' pada tabel di atas untuk detailnya atau download filenya."
    )

# --- Fitur Download ---
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

# --- Footer ---
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