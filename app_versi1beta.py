import streamlit as st
import pandas as pd
import io
import time
import re
import warnings
import win32com.client
import pythoncom
import os

# --- 1. MENGABAIKAN WARNING EXCEL & DEPRECATION ---
# Mengabaikan pesan error cell tanggal yang tidak valid di openpyxl
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# --- FUNGSI BARU: ENKRIPSI EXCEL DENGAN PASSWORD (TAMBAHAN) ---
def to_excel_password(df_save, nama_file_output):
    # Inisialisasi COM untuk sistem Windows
    pythoncom.CoInitialize()
    
    current_dir = os.getcwd()
    # Nama file sementara
    temp_polos = os.path.join(current_dir, "temp_data_polos.xlsx")
    temp_kunci = os.path.join(current_dir, "DATA_HASIL_KURASI.xlsx")
    
    # Simpan ke excel biasa dulu menggunakan engine bawaan pandas
    df_save.to_excel(temp_polos, index=False)
    
    # Panggil engine Excel Windows
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(temp_polos)
        
        # --- TENTUKAN PASSWORD DI SINI ---
        password_rahasia = "AnaMempesona" 
        
        # Simpan ulang dengan password (51 = format .xlsx)
        wb.SaveAs(temp_kunci, 51, password_rahasia)
        wb.Close()
        
        # Baca filenya untuk dikirim ke streamlit
        with open(temp_kunci, "rb") as f:
            data_binari = f.read()
            
        return data_binari
    except Exception as e:
        st.error(f"Gagal mengunci file: {e}")
        return None
    finally:
        excel.Quit()
        # Bersihkan file sementara
        if os.path.exists(temp_polos): os.remove(temp_polos)
        if os.path.exists(temp_kunci): os.remove(temp_kunci)

# 2. KONFIGURASI HALAMAN
st.set_page_config(
    page_title="KURASI", 
    page_icon="image/kurasi-icon.png",  # Menggunakan file logo sebagai favicon browser
    layout="wide"
)

# --- FUNGSI LOGIN ---
def login():
    st.markdown("""
        <style>
        [data-testid="stHorizontalBlock"] {
            align-items: center;
            padding-top: 0rem;
        }
        .login-box {
            padding: 2.5rem;
            border-radius: 15px;
            background-color: #f8f9fa;
            border: 1px solid #e6e9ef;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05);
        }
        .desc-container {
            background-color: #ffffff; 
            padding: 25px; 
            border-radius: 10px; 
            border-left: 5px solid #2e7d32;
        }
        </style>
    """, unsafe_allow_html=True)

    col_left, col_spacer, col_right = st.columns([1.5, 0.2, 1.5])
    
    with col_left:
        try:
            # Gunakan width='stretch' untuk menggantikan use_container_width di 2026
            st.image("image/data.png", width='stretch')
        except:
            st.markdown("<h1 style='color: #2e7d32;'>🛡️ KURASI</h1>", unsafe_allow_html=True)

    with col_right:
        # try:
        #     # Gunakan width='stretch' untuk menggantikan use_container_width di 2026
        #     st.image("image/kurasi.png", width='stretch')
        # except:
        #     st.markdown("<h1 style='color: #2e7d32;'>🛡️ KURASI</h1>", unsafe_allow_html=True)
        # st.markdown('<div class="login-box">', unsafe_allow_html=True)
        st.markdown("""
            <style>
            .desc-container {
                /* Garis samping tipis sebagai pemanis */
                border-left: 3px solid #2e7d32;
                padding-left: 10px;
                margin-top: -100px;
                margin-bottom: 25px;
            }
            .desc-text {
                font-size: 16px; 
                color: #333333; 
                line-height: 1.5; 
                margin: 0;
                text-align: justify;
            }
            </style>
            
            <div class="desc-container">
                <p class="desc-text">
                    <strong>KURASI (Smart Accuracy System)</strong> adalah platform cerdas yang dirancang untuk mengoptimalkan pengelolaan data 
                    <strong>SPM</strong> melalui proses penyaringan otomatis. Sistem ini secara efektif mengeliminasi 
                    data ganda (duplikat) dan memastikan akurasi informasi secara 
                    <span style="color: #2e7d32; font-weight: bold;">instan, transparan, dan akuntabel</span>.
                </p>
            </div>
        """, unsafe_allow_html=True)
        st.subheader("🔐 Login")

        try:
            users = st.secrets["users"]
        except:
            st.error("Konfigurasi Rahasia tidak ditemukan.")
            return

        # SOLUSI: Menambahkan 'key' unik agar tidak error
        username = st.text_input("Username", key="login_user")
        password = st.text_input("Password", type="password", key="login_pass")
        
        # SOLUSI: Menggunakan width='stretch' sesuai peringatan sistem 2026
        btn_login = st.button("Masuk", width='stretch', key="login_btn")
        
        if btn_login:
            if username in users:
                if users[username] == password:
                    st.session_state["logged_in"] = True
                    st.session_state["user_role"] = username.upper()
                    st.success(f"Login Berhasil!")
                    st.rerun()
                else:
                    st.error("❌ Kata Sandi Salah.")
            else:
                st.error(f"⚠️ '{username}' tidak terdaftar.")
        # st.markdown('</div>', unsafe_allow_html=True)

# Inisialisasi Session State (Tetap sama)
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "user_role" not in st.session_state:
    st.session_state["user_role"] = ""

if not st.session_state["logged_in"]:
    login()

else:
    # --- SIDEBAR ---
    with st.sidebar:
        try:
            st.image("image/kurasi-icon.png", width=200)
        except:
            pass
        st.title("👤 Profil Pengguna")
        st.info(f"**{st.session_state['user_role']}**")
        st.write("---")
        
        if st.button("🚪 Keluar"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

    # --- HEADER ---
    col_logo, col_text = st.columns([1, 8])
    with col_logo:
        try:
            st.image("image/kurasi.png", width=300) 
        except:
            st.error("File logo tidak ditemukan")

    with st.expander("📢 BACA PANDUAN FORMAT EXCEL (PENTING)"):
        st.write("""
        1. **Struktur Header**: Judul kolom (NIK, Nama, dll) wajib berada di **Baris Pertama**.
        2. **Kolom Identitas**: Wajib memiliki satu kolom dengan nama **'NIK'**.
        3. **Scrubbing NIK**: Sistem hanya membersihkan kolom NIK.
        4. **Kriteria Lolos**: NIK berjumlah **16 digit** angka, diawali **3277**, dan **tidak ganda**.
        """)
        
        # --- FITUR TAMBAHAN: DOWNLOAD TEMPLATE ---
        st.markdown("---")
        st.write("💡 **Belum punya formatnya?** Unduh template di bawah ini sebagai acuan:")
        
        # Membuat Dataframe Contoh
        df_template = pd.DataFrame({
            'No': [1, 2],
            'NIK': ["3277010101010001", "3277010101010002"],
            'Nama Lengkap': ["Nama Contoh A", "Nama Contoh B"],
            'Tanggal Layanan': ["2025-12-31", "2025-12-30"],
            'Alamat': ["Jl. Contoh No. 123", "Jl. Sampel No. 456"]
        })

        # Konversi ke Excel Binary
        output_template = io.BytesIO()
        with pd.ExcelWriter(output_template, engine='openpyxl') as writer:
            df_template.to_excel(writer, index=False, sheet_name='Template_KURASI')
        
        processed_data = output_template.getvalue()

        # Tombol Download Template
        st.download_button(
            label="📥 Unduh Template Contoh Excel",
            data=processed_data,
            file_name="TEMPLATE_CONTOH_KURASI.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch' # Menyesuaikan gaya 2026
        )

    st.write("---")
    
    # --- INPUT TARGET ---
    target_spm = st.number_input("🎯 Masukkan Target Sasaran (Jiwa)", min_value=1, value=100, step=1)
    
    uploaded_file = st.file_uploader(f"Pilih File Excel BNBA Indikator SPM", type=["xlsx"])

    if uploaded_file:
        start_time = time.time() 

        with st.spinner('⏳ Sedang memvalidasi dan memilah data...'):
            try:
                # 1. BACA DATA (na_filter=False agar alamat/jenis kelamin tidak hilang)
                df_original = pd.read_excel(uploaded_file, dtype=str, na_filter=False)
                
                # 2. FORMAT TANGGAL (Hanya mengubah tampilan yang berbentuk YYYY-MM-DD)
                for col_name in df_original.columns:
                    if 'NIK' not in col_name.upper():
                        mask_tgl = df_original[col_name].str.contains(r'^\d{4}-\d{2}-\d{2}', na=False)
                        if mask_tgl.any():
                            try:
                                df_original.loc[mask_tgl, col_name] = pd.to_datetime(df_original.loc[mask_tgl, col_name], errors='coerce').dt.strftime('%d/%m/%Y')
                            except: pass
            except Exception as e:
                st.error(f"Gagal membaca file: {e}")
                st.stop()
            
            nik_col = [c for c in df_original.columns if 'NIK' in str(c).upper()]
            
            nik_col = [c for c in df_original.columns if 'NIK' in str(c).upper()]
            
            if nik_col:
                col = nik_col[0]
                def mask_nik(val):
                    """Fungsi untuk menyembunyikan digit tengah NIK di pratinjau"""
                    clean = str(val).strip()
                    if len(clean) >= 12:
                        # Menampilkan 4 digit awal, bintang, dan 4 digit akhir
                        return clean[:4] + "********" + clean[-4:]
                    elif len(clean) > 0:
                        return clean[:2] + "********"
                
                # --- 2. SCRUBBING NIK (SANGAT KETAT) ---
                def scrubbing_nik(val):
                    if val is None: return ""
                    val = str(val).strip()
                    
                    # Jika ada E+, LANGSUNG kembalikan nilai asli. 
                    # Jangan lanjut ke bawah (jangan di-re.sub)
                    if 'E+' in val.upper():
                        return val 
                    
                    # Hanya NIK normal yang dibersihkan dari titik/spasi
                    return re.sub(r'\D', '', val)

                # Hasil scrubbing untuk pengecekan (bayangan)
                nik_shadow = df_original[col].apply(scrubbing_nik)

                # --- 3. PROSES PEMILAHAN ---
                # Syarat Dasar: Harus 16 digit angka murni
                mask_16_digit = nik_shadow.str.len() == 16
                mask_ganda = nik_shadow.duplicated(keep='first')
                
                # Syarat Wilayah: Harus diawali '3277'
                mask_wilayah_3277 = nik_shadow.str.startswith('3277')

                # A. DATA LOLOS (16 Digit, Bukan Ganda, Dan Wilayah 3277)
                df_lolos = df_original[mask_16_digit & ~mask_ganda & mask_wilayah_3277].copy()

                # B. DATA LUAR WILAYAH (16 Digit, Bukan Ganda, Tapi BUKAN 3277)
                df_luar_wilayah = df_original[mask_16_digit & ~mask_ganda & ~mask_wilayah_3277].copy()

                # C. DATA ANOMALI (Bukan 16 digit - Termasuk Scientific Format)
                df_anomali = df_original[~mask_16_digit].copy()
                
                # D. DATA GANDA
                df_ganda = df_original[mask_16_digit & mask_ganda].copy()

               # --- 4. DETAIL VERIFIKASI ANOMALI ---
                if not df_anomali.empty:
                    def tentukan_keterangan(val_asli):
                        val_str = str(val_asli).upper()
                        
                        # Deteksi Scientific Format
                        if 'E+' in val_str:
                            return "SCIENTIFIC FORMAT - Ubah kolom NIK menjadi 'TEXT' di Excel!"
                        
                        # Validasi NIK Normal
                        clean_val = scrubbing_nik(val_asli)
                        pjg = len(clean_val)
                        
                        if pjg == 0:
                            return "NIK Kosong"
                        elif pjg < 16:
                            return f"NIK Kurang ({pjg} Digit)"
                        elif pjg > 16:
                            return f"NIK Berlebih ({pjg} Digit)"
                        else:
                            return "Karakter Ilegal (Cek spasi/titik)"
                    
                    df_anomali['KETERANGAN_SISTEM'] = df_anomali[col].apply(tentukan_keterangan)

                if not df_ganda.empty:
                    df_ganda['KETERANGAN_SISTEM'] = "👯 Data Ganda (NIK Duplikat)"

                # --- 5. RE-INDEXING (NOMOR URUT BERURUTAN) ---
                def ksh_nomor(df_input):
                    if df_input is not None and not df_input.empty:
                        # Hapus kolom 'No' lama jika ada agar tidak double
                        if 'No' in df_input.columns:
                            df_input = df_input.drop(columns=['No'])
                        # Masukkan nomor baru di kolom paling depan
                        df_input.insert(0, 'No', range(1, len(df_input) + 1))
                    return df_input

                # Terapkan penomoran ke SEMUA kategori termasuk Luar Wilayah
                df_lolos = ksh_nomor(df_lolos)
                df_luar_wilayah = ksh_nomor(df_luar_wilayah) # Tambahkan ini
                df_anomali = ksh_nomor(df_anomali)
                df_ganda = ksh_nomor(df_ganda)

                # 6. DASHBOARD METRIK
                duration = round(time.time() - start_time, 2)
                st.success(f"💡 Verifikasi Selesai dalam {duration} detik.")
                
                # Hitung Persentase Target
                persen_capaian = round((len(df_lolos) / target_spm) * 100, 2)
                
                # Membuat 5 kolom untuk metrik
                m1, m2, m3, m4, m5 = st.columns(5)

                m1.metric("Data Masuk", f"{len(df_original)} Jiwa")

                m2.metric(
                    "Lolos (3277 Kota Cimahi)", 
                    f"{len(df_lolos)} Jiwa", 
                    f"{persen_capaian}% dari target"
                )

                m3.metric(
                    "Luar Wilayah (Domisili)", 
                    f"{len(df_luar_wilayah)} Jiwa", 
                    delta="Bukan 3277", 
                    delta_color="normal"
                )

                m4.metric(
                    "Anomali", 
                    f"{len(df_anomali)} Jiwa", 
                    delta="Perlu Revisi", 
                    delta_color="inverse"
                )

                m5.metric(
                    "Ganda", 
                    f"{len(df_ganda)} Jiwa", 
                    delta="Disisihkan", 
                    delta_color="inverse"
                )

                # 7. TAB PRATINJAU
                st.write("---")
                t1, t2, t3, t4 = st.tabs(["✅ Lolos (3277 Kota Cimahi)", "📍 Luar Wilayah (Domisili)", "⚠️ Anomali", "🔁 Ganda"])

                with t1: 
                    if not df_lolos.empty:
                        # Buat copy khusus tampilan agar data asli tidak rusak
                        df_display_lolos = df_lolos.copy()
                        df_display_lolos[col] = df_display_lolos[col].apply(mask_nik)
                        st.dataframe(df_display_lolos, width='stretch')
                    else:
                        st.info("Tidak ada data lolos.")
                with t2:
                    if not df_luar_wilayah.empty:
                        df_display_luar = df_luar_wilayah.copy()
                        df_display_luar[col] = df_display_luar[col].apply(mask_nik)
                        st.dataframe(df_display_luar, width='stretch')
                    else: st.info("Tidak ada data luar wilayah.")
                with t3:
                    if not df_anomali.empty:
                        df_display_anomali = df_anomali.copy()
                        # Opsional: Anomali biasanya tidak dimasking agar user tahu digit mana yang salah,
                        # tapi jika ingin tetap aman, gunakan mask_nik:
                        df_display_anomali[col] = df_display_anomali[col].apply(mask_nik)
                        st.dataframe(df_display_anomali, width='stretch')
                    else:
                        st.info("Tidak ada data anomali.")

                with t4:
                    if not df_ganda.empty:
                        df_display_ganda = df_ganda.copy()
                        df_display_ganda[col] = df_display_ganda[col].apply(mask_nik)
                        st.dataframe(df_display_ganda, width='stretch')
                    else:
                        st.info("Tidak ada data ganda.")

                # 8. TOMBOL UNDUH
                # def to_excel(df_save, nama_file_output):
                #     output = io.BytesIO()
                #     with pd.ExcelWriter(output, engine='openpyxl') as writer:
                #         df_save.to_excel(writer, index=False, sheet_name='HASIL')
                #     return output.getvalue()

                # --- 8. TOMBOL UNDUH (VERSI PROTEKSI PASSWORD) ---
                st.write("---")
                st.subheader("📥 Unduh Hasil Akhir")
                
                # Mengambil role user agar muncul di nama file
                role_user = st.session_state.get("user_role", "USER")

                c1, c2, c3, c4 = st.columns(4)

                with c1:
                    if not df_lolos.empty:
                        if st.button("🔒 Proteksi Data Lolos"):
                            with st.spinner("Memproses..."):
                                nama_file = f"1_LOLOS_{role_user}_{uploaded_file.name}"
                                data = to_excel_password(df_lolos, nama_file)
                                if data:
                                    st.download_button("📥 Download Lolos", data=data, file_name=nama_file, key="dl_lolos")

                with c2:
                    if not df_luar_wilayah.empty:
                        if st.button("🔒 Proteksi Data Luar Wilayah"):
                            with st.spinner("Memproses..."):
                                nama_file = f"2_3277+_{role_user}_{uploaded_file.name}"
                                data = to_excel_password(df_luar_wilayah, nama_file)
                                if data:
                                    st.download_button("📥 Download Luar", data=data, file_name=nama_file, key="dl_luar")

                with c3:
                    if not df_anomali.empty:
                        if st.button("🔒 Proteksi Data Anomali"):
                            with st.spinner("Memproses..."):
                                nama_file = f"3_REVISI_{role_user}_{uploaded_file.name}"
                                data = to_excel_password(df_anomali, nama_file)
                                if data:
                                    st.download_button("📥 Download Anomali", data=data, file_name=nama_file, key="dl_anomali")

                with c4:
                    if not df_ganda.empty:
                        if st.button("🔒 Proteksi Data Ganda"):
                            with st.spinner("Memproses..."):
                                nama_file = f"4_GANDA_{role_user}_{uploaded_file.name}"
                                data = to_excel_password(df_ganda, nama_file)
                                if data:
                                    st.download_button("📥 Download Ganda", data=data, file_name=nama_file, key="dl_ganda")
            else:
                st.error("❌ Kolom 'NIK' tidak ditemukan!")