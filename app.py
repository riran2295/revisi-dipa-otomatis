import streamlit as st
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# KONFIGURASI DNS & HALAMAN
# ==========================================
st.set_page_config(page_title="revisi dipa otomatis", page_icon="🚀", layout="centered")

# Injeksi CSS untuk merapikan Layout & Footer
st.markdown("""
    <style>
    /* Mengatur Judul dan Subjudul agar rata tengah */
    .title-text {
        text-align: center;
        font-weight: bold;
        font-size: 2.5rem;
        margin-top: 1rem;
        margin-bottom: 0.5rem;
    }
    .subtitle-text {
        text-align: center;
        color: #666666;
        margin-bottom: 2rem;
    }
    
    /* Mengatur Footer agar nempel di bawah */
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: transparent;
        color: #888888;
        text-align: center;
        padding: 10px;
        font-size: 14px;
        z-index: 100;
    }
    
    /* Menghilangkan padding atas bawaan Streamlit biar lebih rapi */
    .block-container {
        padding-top: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# TAMPILAN HEADER (LOGO & JUDUL)
# ==========================================
# Rasio kolom diubah [1, 3, 1] biar logo di tengah jadi lebih besar (sepanjang teks)
col_logo1, col_logo2, col_logo3 = st.columns([1, 3, 1])
with col_logo2:
    try:
        st.image("logo.png", use_container_width=True)
    except Exception as e:
        st.warning("File 'logo.png' belum ada di folder. Silakan tambahkan dulu gambarnya.")

# Menampilkan Judul & Deskripsi
st.markdown('<div class="title-text">Automasi Reviu Revisi DIPA</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle-text">Upload puluhan file PDF Matriks Perubahan sekaligus, sistem akan otomatis mengekstrak selisih KRO dan meng-generate Kertas Kerja Excel.</div>', unsafe_allow_html=True)


# ==========================================
# LOGIKA UTAMA APLIKASI
# ==========================================

# Membungkus Uploader di dalam kolom biar kotaknya ketengah dan gak kepanjangan
col_up1, col_up2, col_up3 = st.columns([1, 6, 1])
with col_up2:
    uploaded_files = st.file_uploader("Upload File PDF Matriks (Bisa pilih banyak file)", type=["pdf"], accept_multiple_files=True)

# Membungkus tombol proses di tengah
col_btn1, col_btn2, col_btn3 = st.columns([1, 3, 1])
with col_btn2:
    proses_btn = st.button("Proses & Generate Excel", type="primary", use_container_width=True)

if proses_btn:
    if not uploaded_files:
        st.warning("Pilih minimal 1 file PDF dulu ya!")
    else:
        with st.spinner(f'Lagi nge-ekstrak {len(uploaded_files)} file PDF...'):
            data_revisi = []
            pola_angka = r'-?\d+(?:,\d{3})*(?:\.\d+)?'

            for file_pdf in uploaded_files:
                try:
                    with pdfplumber.open(file_pdf) as pdf:
                        satker_aktif = "Belum dapet Satker"
                        jenis_satker = "Daerah" 
                        
                        for halaman in pdf.pages:
                            teks = halaman.extract_text()
                            if not teks: continue
                                
                            baris_teks = teks.split('\n')
                            kro_aktif = None 
                            hitung_baris = 0
                            baris_uang = ""
                            
                            for baris in baris_teks:
                                baris_upper = baris.upper()
                                
                                # HIRARKI DAERAH
                                if any(kw in baris_upper for kw in ["KANTOR PERTANAHAN", "KANTAH", "KANTOR WILAYAH", "KANWIL"]):
                                    nama_bersih = baris
                                    nama_bersih = re.sub(r'^\d+[\s\.\-]*', '', nama_bersih)       
                                    satker_aktif = nama_bersih.strip()
                                    jenis_satker = "Daerah" 
                                    continue

                                # HIRARKI PUSAT (Biro, BPSDM, STPN)
                                elif any(kw in baris_upper for kw in ["BIRO ", "BPSDM", "BADAN PENGEMBANGAN SUMBER DAYA MANUSIA", "SEKOLAH TINGGI PERTANAHAN", "STPN"]):
                                    if "UNIT ORGANISASI" not in baris_upper and "KEMENTERIAN" not in baris_upper:
                                        nama_bersih = re.sub(r'^\d+[\s\.\-]*', '', baris).strip()
                                        satker_aktif = nama_bersih
                                        jenis_satker = "Pusat" 
                                        continue

                                kro_match = re.search(r'^(\d{4}\.[A-Z]{3})', baris)
                                if kro_match:
                                    kro_aktif = kro_match.group(1)
                                    hitung_baris = 0
                                    continue 
                                
                                if kro_aktif:
                                    hitung_baris += 1
                                    if hitung_baris == 1:
                                        baris_uang = baris
                                    elif hitung_baris == 2:
                                        baris_vol = baris
                                        angka_uang = re.findall(pola_angka, baris_uang)
                                        
                                        baris_vol_clean = baris_vol.replace("U M B E R", "UMBER").replace("S U M", "SUM")
                                        angka_vol = re.findall(pola_angka, baris_vol_clean)
                                        
                                        if len(angka_uang) >= 6 and len(angka_vol) >= 3:
                                            selisih_jml = int(angka_uang[4].replace(',', ''))
                                            selisih_blk = int(angka_uang[5].replace(',', ''))
                                            selisih_vol = float(angka_vol[2].replace(',', ''))
                                            
                                            if selisih_jml != 0 or selisih_blk != 0 or selisih_vol != 0:
                                                data_revisi.append({
                                                    'satker': satker_aktif,
                                                    'jenis': jenis_satker, 
                                                    'KRO': kro_aktif,
                                                    'semula_vol': float(angka_vol[0].replace(',', '')),
                                                    'semula_jml': int(angka_uang[0].replace(',', '')),
                                                    'semula_blk': int(angka_uang[1].replace(',', '')),
                                                    'menjadi_vol': float(angka_vol[1].replace(',', '')),
                                                    'menjadi_jml': int(angka_uang[2].replace(',', '')),
                                                    'menjadi_blk': int(angka_uang[3].replace(',', ''))
                                                })
                                        kro_aktif = None
                except Exception as e:
                    st.error(f"Gagal membaca file {file_pdf.name}: {e}")

            if len(data_revisi) > 0:
                # BIKIN EXCEL
                wb = Workbook()
                sheet = wb.active
                sheet.title = "Reviu Revisi DIPA"
                sheet.views.sheetView[0].showGridLines = True

                headers = [
                    "No", "Wilayah", "Nama Satuan Kerja", "KRO", 
                    "Semula Vol", "Semula Jumlah", "Semula Blokir", 
                    "Menjadi Vol", "Menjadi Jumlah", "Menjadi Blokir",
                    "Selisih Vol", "Selisih Jumlah", "Selisih Blokir"
                ]

                font_header = Font(name="Arial", size=10, bold=True)
                kuning_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
                left_align = Alignment(horizontal="left", vertical="center")
                right_align = Alignment(horizontal="right", vertical="center")
                border_tipis = Border(
                    left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
                )

                sheet.row_dimensions[4].height = 28
                for col_idx, header_text in enumerate(headers, 1):
                    cell = sheet.cell(row=4, column=col_idx)
                    cell.value = header_text
                    cell.font = font_header
                    cell.fill = kuning_fill
                    cell.alignment = center_align
                    cell.border = border_tipis

                def tentukan_wilayah(nama_satker, jenis_satker):
                    if jenis_satker == "Pusat": return "Pusat"
                    s = nama_satker.lower()
                    if "jawa timur" in s or "jatim" in s: return "Jawa Timur"
                    if "banten" in s: return "Banten"
                    if "dki jakarta" in s or "jakarta" in s: return "DKI Jakarta"
                    if "jawa barat" in s or "jabar" in s: return "Jawa Barat"
                    if "jawa tengah" in s or "jateng" in s: return "Jawa Tengah"
                    if "sumatera barat" in s or "sumbar" in s: return "Sumatera Barat"
                    if "nusa tenggara timur" in s or "ntt" in s: return "Nusa Tenggara Timur"

                    kamus_wilayah = {
                        "Jawa Timur": ["sumenep", "bangkalan", "surabaya", "malang", "pacitan", "ponorogo", "situbondo", "lumajang", "blitar", "mojokerto", "sidoarjo", "gresik", "banyuwangi", "jember", "kediri", "tuban", "bojonegoro", "ngawi", "magetan", "madiun", "nganjuk", "trenggalek", "tulungagung", "jombang", "pasuruan", "probolinggo", "bondowoso", "lamongan", "pamekasan", "sampang", "batu"],
                        "Banten": ["tangerang", "serang", "cilegon", "pandeglang", "lebak"],
                        "DKI Jakarta": ["kepulauan seribu"],
                        "Jawa Barat": ["bandung", "bogor", "depok", "bekasi", "cimahi", "sukabumi", "cianjur", "garut", "tasikmalaya", "cirebon", "kuningan", "majalengka", "sumedang", "indramayu", "subang", "purwakarta", "karawang", "pangandaran", "banjar"],
                        "Jawa Tengah": ["semarang", "surakarta", "solo", "salatiga", "tegal", "pekalongan", "banyumas", "cilacap", "purbalingga", "banjarnegara", "kebumen", "purworejo", "wonosobo", "boyolali", "klaten", "sukoharjo", "wonogiri", "karanganyar", "sragen", "grobogan", "blora", "rembang", "pati", "kudus", "jepara", "demak", "temanggung", "kendal", "batang", "pemalang", "brebes"],
                        "Sumatera Barat": ["padang", "bukittinggi", "payakumbuh", "solok", "sawahlunto", "pariaman", "pasaman", "agam", "lima puluh kota", "tanah datar", "sijunjung", "dharmasraya", "pesisir selatan", "mentawai"],
                        "Nusa Tenggara Timur": ["kupang", "alor", "belu", "ende", "flores", "lembata", "manggarai", "ngada", "nagekeo", "rote", "sabu", "sikka", "sumba", "timor"]
                    }
                    
                    for provinsi, daftar_kota in kamus_wilayah.items():
                        for kota in daftar_kota:
                            if re.search(fr'\b{re.escape(kota)}\b', s): return provinsi
                                
                    if "prov." in s: return s.split("prov.")[1].strip().title()
                    if "provinsi " in s: return s.split("provinsi ")[1].strip().title()
                    return "Wilayah Belum Diset"

                baris_mulai = 5
                for indeks, d in enumerate(data_revisi):
                    baris = baris_mulai + indeks
                    sheet.row_dimensions[baris].height = 20
                    
                    satker_rapi = d['satker'].title().replace("Kab.", "Kabupaten")
                    wilayah = tentukan_wilayah(d['satker'], d['jenis']) 
                    
                    sheet.cell(row=baris, column=1, value=indeks + 1).alignment = center_align
                    sheet.cell(row=baris, column=2, value=wilayah).alignment = left_align
                    sheet.cell(row=baris, column=3, value=satker_rapi).alignment = left_align
                    sheet.cell(row=baris, column=4, value=d['KRO']).alignment = center_align
                    
                    sheet.cell(row=baris, column=5, value=d['semula_vol'])
                    sheet.cell(row=baris, column=6, value=d['semula_jml'])
                    sheet.cell(row=baris, column=7, value=d['semula_blk'])
                    sheet.cell(row=baris, column=8, value=d['menjadi_vol'])
                    sheet.cell(row=baris, column=9, value=d['menjadi_jml'])
                    sheet.cell(row=baris, column=10, value=d['menjadi_blk'])

                    sheet.cell(row=baris, column=11, value=f"=H{baris}-E{baris}")
                    sheet.cell(row=baris, column=12, value=f"=I{baris}-F{baris}")
                    sheet.cell(row=baris, column=13, value=f"=J{baris}-G{baris}")
                    
                    for col_idx in range(1, 14):
                        c = sheet.cell(row=baris, column=col_idx)
                        c.border = border_tipis
                        c.font = Font(name="Arial", size=10)
                        
                        if col_idx >= 5: 
                            c.number_format = '#,##0'
                            c.alignment = right_align

                for col in sheet.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = re.sub(r'[^A-Z]', '', col[0].coordinate)
                    sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

                output_excel = io.BytesIO()
                wb.save(output_excel)
                output_excel.seek(0)
                
                st.balloons()
                st.success(f"Sukses! Menemukan total {len(data_revisi)} baris selisih dari {len(uploaded_files)} PDF.")
                
                # Menampilkan tombol download di tengah
                col_dl1, col_dl2, col_dl3 = st.columns([1, 3, 1])
                with col_dl2:
                    st.download_button(
                        label="⬇️ Download Kertas Kerja Excel",
                        data=output_excel,
                        file_name="kertas_kerja_OTOMATIS.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
            else:
                st.info("Semua PDF sudah dicek, tapi tidak ditemukan selisih KRO.")

# Menampilkan Footer
st.markdown('<div class="footer">Developed by riranism</div>', unsafe_allow_html=True)