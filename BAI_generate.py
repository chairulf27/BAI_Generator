import streamlit as st
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
st.set_page_config(page_title="FAST BAI", layout="wide", page_icon="⚡")
st.title("🔌 BAI NetGen (Network Generator)")
st.markdown("Isi kelengkapan data di bawah ini untuk mengunduh dokumen Berita Acara Instalasi.")

# FORMULIR INPUT
with st.form("form_bai"):
    st.subheader("1. Informasi Layanan & Waktu")
    col1, col2 = st.columns(2)
    with col1:
        no_pa = st.text_input("Nomor PA", placeholder="Contoh: A221401007204")
        nama_layanan_header = st.text_input("Jenis Layanan", placeholder="Contoh: INTERNET BROADBAND CORPORATE")
        sid_layanan = st.text_input("SID / ID Layanan (Opsional)", placeholder="Contoh: 221401003653")
    with col2:
        tgl_input = st.date_input("Pilih Tanggal Instalasi / Generate", value=datetime.date.today())
    
    st.markdown("---")
    
    st.subheader("2. Data Pelanggan")
    col3, col4 = st.columns(2)
    with col3:
        nama_pelanggan = st.text_input("Nama Pelanggan", placeholder="Contoh: BUMDES KEMBANG KENANGA")
        koordinat_pelanggan = st.text_input("Titik Koordinat Pelanggan", placeholder="Contoh: -0.5817, 101.5156")
    with col4:
        alamat_pelanggan = st.text_area("Alamat Lengkap Pelanggan")

    st.markdown("---")

    st.subheader("3. Spesifikasi Teknis Perangkat")
    col5, col6 = st.columns(2)
    with col5:
        st.markdown("**Di Sisi Pelanggan**")
        perangkat_pelanggan = st.text_input("Nama Perangkat (Pelanggan)", placeholder="Contoh: HPE A3100-8 v2 EI Switch")
        port_pelanggan = st.text_input("Kanal Port (Pelanggan)", placeholder="Contoh: Port 2")
    with col6:
        st.markdown("**Di Sisi POP**")
        lokasi_pop = st.text_input("Alamat/Lokasi POP", placeholder="Contoh: SBT-CAMAT.GUNUNG.TOAR-ASR920-UPE-01")
        perangkat_pop = st.text_input("Nama Perangkat & SN (POP)", placeholder="Contoh: ASR-920-24SZ-M & CAT2125V37A")
        port_pop = st.text_input("Kanal / Port (POP)", placeholder="Contoh: GigabitEthernet0/0/0")

    btn_generate = st.form_submit_button("🚀 Generate Berita Acara (Word)", use_container_width=True)

# ==========================================
# PROSES RENDER & DOWNLOAD DOKUMEN (CLOUD READY)
# ==========================================
if btn_generate:
    if not (no_pa and nama_pelanggan and nama_layanan_header):
        st.error("⚠️ Mohon lengkapi minimal Nomor PA, Jenis Layanan, dan Nama Pelanggan!")
    else:
        try:
            with st.spinner("Merakit dokumen..."):
                # Menyiapkan data
                detail_layanan = f"{nama_layanan_header} ({sid_layanan})" if sid_layanan else nama_layanan_header
                hari_ini = HARI[tgl_input.weekday()]
                bulan_teks = BULAN[tgl_input.month]
                tanggal_generate = tgl_input.strftime(f"%B %d, %Y")
                tanggal_ttd = f"{tgl_input.day} {bulan_teks} {tgl_input.year}"
                
                data_mapping = {
                    'tanggal_generate': tanggal_generate,
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
                    'tanggal_ttd': tanggal_ttd
                }

                # Buka dan Proses Template murni di memori (Cloud Safe)
                doc = DocxTemplate("template_bai.docx")
                doc.render(data_mapping)
                
                file_buffer = io.BytesIO()
                doc.save(file_buffer)
                file_buffer.seek(0)
                
                st.success(f"✅ Dokumen berhasil digenerate! Silakan unduh file untuk pelanggan: **{nama_pelanggan}**")
                
                st.download_button(
                    label="📥 Download File BAI (.docx)",
                    data=file_buffer,
                    file_name=f"BAI_{nama_pelanggan}_{no_pa}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
                st.balloons()
                
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan: {e}. Pastikan file 'template_bai.docx' ada di folder aplikasi!")