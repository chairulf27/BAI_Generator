import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import datetime

# ==========================================
# FUNGSI PINTAR UNTUK MENGUBAH ANGKA JADI HURUF
# ==========================================
def terbilang(n):
    satuan = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]
    n = int(n)
    if n < 12: return satuan[n]
    elif n < 20: return satuan[n - 10] + " Belas"
    elif n < 100: return (satuan[n // 10] + " Puluh " + satuan[n % 10]).strip()
    elif n < 200: return "Seratus " + terbilang(n - 100)
    elif n < 1000: return (satuan[n // 100] + " Ratus " + terbilang(n % 100)).strip()
    elif n < 2000: return "Seribu " + terbilang(n - 1000)
    elif n < 10000: return (satuan[n // 1000] + " Ribu " + terbilang(n % 1000)).strip()
    return str(n)

HARI = {0: 'Senin', 1: 'Selasa', 2: 'Rabu', 3: 'Kamis', 4: 'Jumat', 5: 'Sabtu', 6: 'Minggu'}
BULAN = {1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April', 5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'}

# ==========================================
# ANTARMUKA APLIKASI
# ==========================================
st.set_page_config(page_title="AutoBAI", layout="wide", page_icon="⚡")
st.title("⚡ Generator BAI Icon+")
st.markdown("Silakan pilih tipe Berita Acara Instalasi yang ingin dibuat.")

# PILIHAN TIPE BAI
tipe_bai = st.radio("Pilih Mode Dokumen:", ["Individu (1 Lokasi)", "Terlampir (Banyak Lokasi / Massal)"], horizontal=True)
st.markdown("---")

# ==========================================
# JALUR 1: MODE INDIVIDU
# ==========================================
if tipe_bai == "Individu (1 Lokasi)":
    with st.form("form_individu"):
        st.subheader("📝 Form BAI Individu")
        col1, col2 = st.columns(2)
        with col1:
            no_pa = st.text_input("Nomor PA", placeholder="Contoh: A221401007204")
            nama_layanan_header = st.text_input("Jenis Layanan", placeholder="Contoh: INTERNET BROADBAND CORPORATE")
            sid_layanan = st.text_input("SID / ID Layanan (Opsional)")
        with col2:
            tgl_input = st.date_input("Pilih Tanggal Instalasi / Generate", value=datetime.date.today())
        
        col3, col4 = st.columns(2)
        with col3:
            nama_pelanggan = st.text_input("Nama Pelanggan", placeholder="Contoh: BUMDES KEMBANG KENANGA")
            koordinat_pelanggan = st.text_input("Titik Koordinat Pelanggan")
        with col4:
            alamat_pelanggan = st.text_area("Alamat Lengkap Pelanggan")

        col5, col6 = st.columns(2)
        with col5:
            perangkat_pelanggan = st.text_input("Nama Perangkat (Pelanggan)")
            port_pelanggan = st.text_input("Kanal Port (Pelanggan)")
        with col6:
            lokasi_pop = st.text_input("Alamat/Lokasi POP")
            perangkat_pop = st.text_input("Nama Perangkat & SN (POP)")
            port_pop = st.text_input("Kanal / Port (POP)")

        btn_generate_individu = st.form_submit_button("🚀 Generate BAI Individu (Word)", use_container_width=True)

    if btn_generate_individu:
        if not (no_pa and nama_pelanggan and nama_layanan_header):
            st.error("⚠️ Mohon lengkapi minimal Nomor PA, Jenis Layanan, dan Nama Pelanggan!")
        else:
            try:
                detail_layanan = f"{nama_layanan_header} ({sid_layanan})" if sid_layanan else nama_layanan_header
                hari_ini = HARI[tgl_input.weekday()]
                bulan_teks = BULAN[tgl_input.month]
                
                data_mapping = {
                    'tanggal_generate': tgl_input.strftime(f"%B %d, %Y"),
                    'nama_layanan_header': nama_layanan_header,
                    'no_pa': no_pa,
                    'hari_ini': hari_ini,
                    'tgl_terbilang': terbilang(tgl_input.day),
                    'bulan_teks': bulan_teks,
                    'tahun_terbilang': terbilang(tgl_input.year),
                    'detail_layanan': detail_layanan,
                    'nama_pelanggan': nama_pelanggan,
                    'alamat_pelanggan': alamat_pelanggan,
                    'koordinat_pelanggan': koordinat_pelanggan,
                    'perangkat_pelanggan': perangkat_pelanggan,
                    'port_pelanggan': port_pelanggan,
                    'lokasi_pop': lokasi_pop,
                    'perangkat_pop': perangkat_pop,
                    'port_pop': port_pop,
                    'tanggal_ttd': f"{tgl_input.day} {bulan_teks} {tgl_input.year}"
                }

                doc = DocxTemplate("template_bai.docx")
                doc.render(data_mapping)
                file_buffer = io.BytesIO()
                doc.save(file_buffer)
                file_buffer.seek(0)
                
                st.success(f"✅ Dokumen berhasil digenerate!")
                st.download_button(
                    label="📥 Download File BAI Individu (.docx)",
                    data=file_buffer,
                    file_name=f"BAI_{nama_pelanggan}_{no_pa}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary", use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ Error: {e}. Pastikan file 'template_bai.docx' tersedia.")

# ==========================================
# JALUR 2: MODE TERLAMPIR (MASSAL)
# ==========================================
elif tipe_bai == "Terlampir (Banyak Lokasi / Massal)":
    st.subheader("📑 Form BAI Massal (Dengan Lampiran)")
    
    # Header Dokumen Utama
    col_a, col_b = st.columns(2)
    with col_a:
        nama_instansi = st.text_input("Nama Instansi Induk", placeholder="Contoh: Dinas Kominfo Kab. XYZ")
        jenis_layanan = st.text_input("Jenis Layanan (Header)")
    with col_b:
        tgl_massal = st.date_input("Pilih Tanggal Instalasi / Generate", key="tgl_massal")

    st.markdown("**Paste Data dari Excel ke Tabel di Bawah Ini (Klik sel pertama lalu Ctrl+V):**")
    
    # Menyiapkan tabel kosong untuk diisi
    df_kosong = pd.DataFrame(columns=["Nomor PA", "Alamat Lokasi", "Nama Perangkat", "SN Perangkat"])
    
    # Tabel interaktif (Bisa tambah baris / copy-paste dari Excel)
    tabel_input = st.data_editor(df_kosong, num_rows="dynamic", use_container_width=True)

    btn_generate_massal = st.button("🚀 Generate BAI Terlampir (Word)", use_container_width=True, type="primary")

    if btn_generate_massal:
        if not nama_instansi or tabel_input.empty:
            st.error("⚠️ Nama Instansi wajib diisi dan tabel tidak boleh kosong!")
        else:
            try:
                # 1. Mengolah data tabel dari layar Streamlit menjadi format untuk Word
                daftar_lokasi = []
                for index, row in tabel_input.iterrows():
                    # Memastikan tidak ada data kosong yang error (NaN)
                    if pd.notna(row['Nomor PA']): 
                        daftar_lokasi.append({
                            'no': index + 1, # Otomatis bikin nomor urut
                            'pa': str(row['Nomor PA']),
                            'alamat': str(row['Alamat Lokasi']),
                            'perangkat': str(row['Nama Perangkat']),
                            'sn': str(row['SN Perangkat'])
                        })

                # 2. Menyiapkan data header
                hari_ini = HARI[tgl_massal.weekday()]
                bulan_teks = BULAN[tgl_massal.month]
                
                data_mapping_massal = {
                    'tanggal_generate': tgl_massal.strftime(f"%B %d, %Y"),
                    'nama_layanan_header': jenis_layanan,
                    'nama_pelanggan': nama_instansi,
                    'hari_ini': hari_ini,
                    'tgl_terbilang': terbilang(tgl_massal.day),
                    'bulan_teks': bulan_teks,
                    'tahun_terbilang': terbilang(tgl_massal.year),
                    'tanggal_ttd': f"{tgl_massal.day} {bulan_teks} {tgl_massal.year}",
                    # Ini variabel penting untuk mengisi tabel di Word!
                    'tabel_lampiran': daftar_lokasi 
                }

                # 3. Merender menggunakan template KEDUA
                doc_massal = DocxTemplate("template_bai_lampiran.docx")
                doc_massal.render(data_mapping_massal)
                
                file_buffer_massal = io.BytesIO()
                doc_massal.save(file_buffer_massal)
                file_buffer_massal.seek(0)
                
                st.success(f"✅ Dokumen massal untuk {nama_instansi} berhasil digenerate dengan {len(daftar_lokasi)} lokasi!")
                st.download_button(
                    label="📥 Download File BAI Terlampir (.docx)",
                    data=file_buffer_massal,
                    file_name=f"BAI_Massal_{nama_instansi}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary", use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ Error: {e}. Pastikan file 'template_bai_lampiran.docx' tersedia dan tag tabel di Word sudah benar.")