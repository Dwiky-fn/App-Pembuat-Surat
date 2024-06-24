import streamlit as st
from docx import Document
from io import BytesIO
from docx.shared import Pt
from docx.oxml.ns import qn

# Fungsi untuk mengatur font dan ukuran teks
def set_font(run, font_name='Times New Roman', font_size=12):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), font_name)

# Fungsi untuk mengatur teks menjadi bold
def set_bold(run):
    run.bold = True

# Fungsi untuk membuat surat berdasarkan templat
def create_letter(no_srt, bulan, tanggal, tujuan, acara, hari, tanggal_keg, jam_mulai, jam_selesai, tempat, nama_sekre, nim_sekre):
    # Load template surat
    doc = Document('pemberitahuan.docx')
    
    # Ganti placeholder dengan data yang dimasukkan dan atur font
    for para in doc.paragraphs:
        if '{no_surat}' in para.text:
            para.text = para.text.replace('{no_surat}', no_srt)
            for run in para.runs:
                set_font(run)
        if '{bulan}' in para.text:
            para.text = para.text.replace('{bulan}', bulan)
            for run in para.runs:
                set_font(run)
        if '{tanggal}' in para.text:
            para.text = para.text.replace('{tanggal}', tanggal)
            for run in para.runs:
                set_font(run)
        if '{tujuan}' in para.text:
            for run in para.runs:
                run.text = run.text.replace('{tujuan}', tujuan)
                if tujuan in run.text:
                    set_bold(run)
                set_font(run)
        if '{kegiatan}' in para.text:
            for run in para.runs:
                run.text = run.text.replace('{kegiatan}', acara)
                if tujuan in run.text:
                    set_bold(run)
                set_font(run)
        if '{hari}' in para.text:
            para.text = para.text.replace('{hari}', hari)
            for run in para.runs:
                set_font(run)
        if '{tanggal_keg}' in para.text:
            para.text = para.text.replace('{tanggal_keg}', tanggal_keg)
            for run in para.runs:
                set_font(run)
        if '{jam_mulai}' in para.text:
            para.text = para.text.replace('{jam_mulai}', jam_mulai)
            for run in para.runs:
                set_font(run)
        if '{jam_selesai}' in para.text:
            para.text = para.text.replace('{jam_selesai}', jam_selesai)
            for run in para.runs:
                set_font(run)
        if '{tempat}' in para.text:
            para.text = para.text.replace('{tempat}', tempat)
            for run in para.runs:
                set_font(run)

    for table in doc.tables:
        for row in table.rows:
            cell = row.cells[1]
            for paragraph in cell.paragraphs:
                if '{nama_sekre}' in paragraph.text:
                    for run in paragraph.runs:
                        run.text = run.text.replace('{nama_sekre}', nama_sekre)
                        if nama_sekre in run.text:
                            set_bold(run)
                        set_font(run)
                if '{nim_sekre}' in paragraph.text:
                    for run in paragraph.runs:
                        run.text = run.text.replace('{nim_sekre}', nim_sekre)
                        set_font(run)
    
    # Simpan dokumen ke BytesIO
    output = BytesIO()
    doc.save(output)

    return output

st.title('Aplikasi Pembuat Surat Otomatis')

# Input dari pengguna
no_srt = st.text_input('Nomor Surat', placeholder='001')
bulan = st.text_input('Bulan', placeholder='Bulan dengan Angka Romawi Cth: VI')
tanggal = st.text_input('Tanggal', placeholder='Tanggal Pembuatan Surat Cth: 23 Juni 2004')
tujuan = st.text_input('Tujuan', placeholder='Cth: KEMENDAGRI BEM')
acara = st.text_input('Kegiatan', placeholder='Cth: Rapat Kerja Kepengurusan HMJ Teknik Elektro')
hari = st.text_input('Hari Kegiatan', placeholder='Cth: Senin atau Senin - Rabu')
tanggal_keg = st.text_input('Tanggal Kegiatan', placeholder='Cth: 1 Juli 2024, 1 - 3 Juli 2024, atau 30 Juni 2024 - 2 Juli 2024')
jam_mulai = st.text_input('Jam Mulai', placeholder='Cth: 08.00')
jam_selesai = st.text_input('Jam Selesai', placeholder='Cth: 12.00 WIB atau Selesai')
tempat = st.text_input('Tempat', placeholder='Ruangan Teori TI 4')
nama_sekre = st.text_input('Nama Sekretaris', placeholder='Nama yang membuat surat')
nim_sekre = st.text_input('NIM Sekretaris', placeholder='NIM yang membuat surat')

if st.button('Buat Surat'):
    if no_srt and tanggal and tujuan and acara and tanggal_keg and jam_mulai and jam_selesai and tempat and nama_sekre and nim_sekre:
        # Create letter and get result as BytesIO
        surat = create_letter(no_srt, bulan, tanggal, tujuan, acara, hari, tanggal_keg, jam_mulai, jam_selesai, tempat, nama_sekre, nim_sekre)
        
        if surat:
            # Provide option to download the letter
            st.download_button(
                label="Unduh Surat",
                data=surat.getvalue(),
                file_name=f'{no_srt} Pemberitahuan {tujuan}.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
    else:
        st.warning("Mohon lengkapi semua kolom sebelum membuat surat.")