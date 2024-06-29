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
def srt_pemberitahuan():
    st.title('Surat Pemberitahuan Kegiatan')

    # Load template surat
    doc = Document('pemberitahuan.docx')

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

    # Ganti placeholder dengan data yang dimasukkan dan atur font
    if st.button('Buat Surat'):
        if no_srt and tanggal and tujuan and acara and tanggal_keg and jam_mulai and jam_selesai and tempat and nama_sekre and nim_sekre:
        # Create letter and get result as BytesIO

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
    
        if output:
        # Provide option to download the letter
            st.download_button(
            label = "Unduh Surat",
            data = output.getvalue(),
            file_name = f'{no_srt} Pemberitahuan {tujuan}.docx',
            mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

def srt_peminjaman():
    jenis = st.radio('Jenis Peminjaman:', ('Ruangan', 'Peralatan'))

    if jenis == 'Peralatan':
        st.title('Surat Peminjaman Peralatan')
        doc = Document('peminjaman_barang.docx')

        no_srt = st.text_input('Nomor Surat', placeholder='Cht: 001')
        panitia_keg = st.text_input('Panitia Kegiatan', placeholder='Cth: /PAN-RAIS')
        bulan = st.text_input('Bulan', placeholder='Bulan dengan Angka Romawi Cth: VI')
        tanggal = st.text_input('Tanggal', placeholder='Tanggal Pembuatan Surat Cth: 23 Juni 2004')
        lampiran = st.text_input('Lampiran', placeholder='Cth: 1 atau Berkas')

        tujuan = st.text_input('Tujuan', placeholder='Cth: Kepala Jurusan Teknik Elektro')
        tempat = st.text_input('Tempat Tujuan', placeholder='Politeknik Negeri Pontianak')

        # isi
        kegiatan = st.text_input('Kegiatan', placeholder='Cth: Rapat Kerja Kepengurusan HMJ Teknik Elektro')
        hari = st.text_input('Hari Kegiatan', placeholder='Cth: Senin atau Senin - Rabu')
        tanggal_keg = st.text_input('Tanggal Kegiatan', placeholder='Cth: 1 Juli 2024, 1 - 3 Juli 2024, atau 30 Juni 2024 - 2 Juli 2024')
        jam_mulai = st.text_input('Jam Mulai', placeholder='Cth: 08.00')
        jam_selesai = st.text_input('Jam Selesai', placeholder='Cth: 12.00 WIB atau Selesai')
        tempat_keg = st.text_input('Tempat', placeholder='Ruangan Teori TI 4')

        jab_penanggungjwb = st.text_input('Jabatan Penanggung Jawab', placeholder='Cth: Ketua Panitia')
        nama_penanggungjwb = st.text_input('Nama Penanggung Jawab', placeholder='Cth: Muhammad Ikhwan')
        nim_penanggungjwb = st.text_input('NIM Penanggung Jawab', placeholder='Cth: 3202315001')

        # menyetujui
        nama_tujuan = st.text_input('Nama Tujuan Peminjaman', placeholder='Cth: Hasan')
        nip_tujuan = st.text_input('NIP Tujuan Peminjaman', placeholder='Cth: 197108201999031003')

        # peralatan
        jns_peralatan = jenis
        nama_peralatan = st.text_input('Nama Barang', placeholder='Cth: Proyektor')
        jmlh = st.text_input('Jumlah Barang', placeholder='Cth: 2 Buah')

        if st.button('Buat Surat'):
            if no_srt and panitia_keg and tanggal and tujuan and kegiatan and tanggal_keg and jam_mulai and jam_selesai and tempat_keg and jab_penanggungjwb and nama_penanggungjwb and nim_penanggungjwb and tempat and nama_tujuan and nip_tujuan and jns_peralatan and nama_peralatan and jmlh:
            # Create letter and get result as BytesIO

                for para in doc.paragraphs:
                    if '{no_srt}' in para.text:
                        para.text = para.text.replace('{no_srt}', no_srt)
                        for run in para.runs:
                            set_font(run)
                    if '{panitia_keg}' in para.text:
                        para.text = para.text.replace('{panitia_keg}', panitia_keg)
                        for run in para.runs:
                            set_font(run)
                    if '{bln}' in para.text:
                        para.text = para.text.replace('{bln}', bulan)
                        for run in para.runs:
                            set_font(run)
                    if '{tanggal_srt}' in para.text:
                        para.text = para.text.replace('{tanggal_srt}', tanggal)
                        for run in para.runs:
                            set_font(run)
                    if '{lampiran}' in para.text:
                        para.text = para.text.replace('{lampiran}', lampiran)
                        for run in para.runs:
                            set_font(run)
                    if '{tujuan}' in para.text:
                        for run in para.runs:
                            run.text = run.text.replace('{tujuan}', tujuan)
                            if tujuan in run.text:
                                set_bold(run)
                            set_font(run)
                    if '{tempat}' in para.text:
                        for run in para.runs:
                            run.text = run.text.replace('{tempat}', tempat)
                            if tujuan in run.text:
                                set_bold(run)
                            set_font(run)
                    if '{kegiatan}' in para.text:
                        para.text = para.text.replace('{kegiatan}', kegiatan)
                        for run in para.runs:
                            if tujuan in run.text:
                                set_bold(run)
                            set_font(run)
                    if '{hari_keg}' in para.text:
                        para.text = para.text.replace('{hari_keg}', hari)
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
                    if '{tempat_keg}' in para.text:
                        para.text = para.text.replace('{tempat_keg}', tempat_keg)
                        for run in para.runs:
                            set_font(run)
                    if '{jns_peralatan}' in para.text:
                        para.text = para.text.replace('{jns_peralatan}', jns_peralatan.upper())
                        for run in para.runs:
                            set_font(run)

                for table in doc.tables:
                    for row in table.rows:
                        cell = row.cells[1]
                        for paragraph in cell.paragraphs:
                            if '{jab_penanggungjwb}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{jab_penanggungjwb}', jab_penanggungjwb)
                                    set_font(run)
                            if '{nama_penanggungjwb}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{nama_penanggungjwb}', nama_penanggungjwb)
                                    if nama_penanggungjwb in run.text:
                                        set_bold(run)
                                    set_font(run)
                            if '{nim_penanggungjwb}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{nim_penanggungjwb}', nim_penanggungjwb)
                                    set_font(run)
                            if '{jmlh}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{jmlh}', jmlh)
                                    set_font(run)
                            if '{tujuan}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{tujuan}', tujuan)
                                    set_font(run)
                            if '{nama_tujuan}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{nama_tujuan}', nama_tujuan)
                                    if nama_tujuan in run.text:
                                        set_bold(run)
                                    set_font(run)
                            if '{nip_tujuan}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{nip_tujuan}', nip_tujuan)
                                    set_font(run)

                for table in doc.tables:
                    for row in table.rows:
                        cell = row.cells[0]
                        for paragraph in cell.paragraphs:
                            if '{nama_peralatan}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{nama_peralatan}', nama_peralatan)
                                    set_font(run)

            # Simpan dokumen ke BytesIO
            output = BytesIO()
            doc.save(output)

            if output:
            # Provide option to download the letter
                st.download_button(
                label = "Unduh Surat",
                data = output.getvalue(),
                file_name = f'{no_srt} Peminjaman {nama_peralatan} {tujuan}.docx',
                mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )

    elif jenis == 'Ruangan':
        st.title('Surat Peminjaman Ruangan')
        doc = Document('peminjaman_ruangan.docx')

        no_srt = st.text_input('Nomor Surat', placeholder='Cht: 001')
        panitia_keg = st.text_input('Panitia Kegiatan', placeholder='Cth: /PAN-RAIS')
        bulan = st.text_input('Bulan', placeholder='Bulan dengan Angka Romawi Cth: VI')
        tanggal = st.text_input('Tanggal', placeholder='Tanggal Pembuatan Surat Cth: 23 Juni 2004')
        lampiran = st.text_input('Lampiran', placeholder='Cth: 1 atau Berkas')

        tujuan = st.text_input('Tujuan', placeholder='Cth: Kepala Jurusan Teknik Elektro')
        tempat = st.text_input('Tempat Tujuan', placeholder='Politeknik Negeri Pontianak')

        # isi
        kegiatan = st.text_input('Kegiatan', placeholder='Cth: Rapat Kerja Kepengurusan HMJ Teknik Elektro')
        hari = st.text_input('Hari Kegiatan', placeholder='Cth: Senin atau Senin - Rabu')
        tanggal_keg = st.text_input('Tanggal Kegiatan', placeholder='Cth: 1 Juli 2024, 1 - 3 Juli 2024, atau 30 Juni 2024 - 2 Juli 2024')
        jam_mulai = st.text_input('Jam Mulai', placeholder='Cth: 08.00')
        jam_selesai = st.text_input('Jam Selesai', placeholder='Cth: 12.00 WIB atau Selesai')

        jab_penanggungjwb = st.text_input('Jabatan Penanggung Jawab', placeholder='Cth: Ketua Panitia')
        nama_penanggungjwb = st.text_input('Nama Penanggung Jawab', placeholder='Cth: Muhammad Ikhwan')
        nim_penanggungjwb = st.text_input('NIM Penanggung Jawab', placeholder='Cth: 3202315001')

        # menyetujui
        nama_tujuan = st.text_input('Nama Tujuan Peminjaman', placeholder='Cth: Hasan')
        nip_tujuan = st.text_input('NIP Tujuan Peminjaman', placeholder='Cth: 197108201999031003')

        # peralatan
        jns_peralatan = jenis
        nama_peralatan = st.text_input('Nama Barang', placeholder='Cth: Ruangan Teori TI 1')
        jmlh = st.text_input('Jumlah Barang', placeholder='Cth: 1 Kelas')

        if st.button('Buat Surat'):
            if no_srt and panitia_keg and tanggal and lampiran and tujuan and kegiatan and tanggal_keg and jam_mulai and jam_selesai and jab_penanggungjwb and nama_penanggungjwb and nim_penanggungjwb and tempat and nama_tujuan and nip_tujuan and jns_peralatan and nama_peralatan and jmlh:
            # Create letter and get result as BytesIO

                for para in doc.paragraphs:
                    if '{no_srt}' in para.text:
                        para.text = para.text.replace('{no_srt}', no_srt)
                        for run in para.runs:
                            set_font(run)
                    if '{panitia_keg}' in para.text:
                        para.text = para.text.replace('{panitia_keg}', panitia_keg)
                        for run in para.runs:
                            set_font(run)
                    if '{bln}' in para.text:
                        para.text = para.text.replace('{bln}', bulan)
                        for run in para.runs:
                            set_font(run)
                    if '{tanggal_srt}' in para.text:
                        para.text = para.text.replace('{tanggal_srt}', tanggal)
                        for run in para.runs:
                            set_font(run)
                    if '{lampiran}' in para.text:
                        para.text = para.text.replace('{lampiran}', lampiran)
                        for run in para.runs:
                            set_font(run)
                    if '{tujuan}' in para.text:
                        for run in para.runs:
                            run.text = run.text.replace('{tujuan}', tujuan)
                            if tujuan in run.text:
                                set_bold(run)
                            set_font(run)
                    if '{tempat}' in para.text:
                        for run in para.runs:
                            run.text = run.text.replace('{tempat}', tempat)
                            if tujuan in run.text:
                                set_bold(run)
                            set_font(run)
                    if '{kegiatan}' in para.text:
                        para.text = para.text.replace('{kegiatan}', kegiatan)
                        for run in para.runs:
                            if tujuan in run.text:
                                set_bold(run)
                            set_font(run)
                    if '{hari_keg}' in para.text:
                        para.text = para.text.replace('{hari_keg}', hari)
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
                    if '{jns_peralatan}' in para.text:
                        para.text = para.text.replace('{jns_peralatan}', jns_peralatan.upper())
                        for run in para.runs:
                            set_font(run)

                for table in doc.tables:
                    for row in table.rows:
                        cell = row.cells[1]
                        for paragraph in cell.paragraphs:
                            if '{jab_penanggungjwb}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{jab_penanggungjwb}', jab_penanggungjwb)
                                    set_font(run)
                            if '{nama_penanggungjwb}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{nama_penanggungjwb}', nama_penanggungjwb)
                                    if nama_penanggungjwb in run.text:
                                        set_bold(run)
                                    set_font(run)
                            if '{nim_penanggungjwb}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{nim_penanggungjwb}', nim_penanggungjwb)
                                    set_font(run)
                            if '{jmlh}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{jmlh}', jmlh)
                                    set_font(run)
                            if '{tujuan}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{tujuan}', tujuan)
                                    set_font(run)
                            if '{nama_tujuan}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{nama_tujuan}', nama_tujuan)
                                    if nama_tujuan in run.text:
                                        set_bold(run)
                                    set_font(run)
                            if '{nip_tujuan}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{nip_tujuan}', nip_tujuan)
                                    set_font(run)

                for table in doc.tables:
                    for row in table.rows:
                        cell = row.cells[0]
                        for paragraph in cell.paragraphs:
                            if '{nama_peralatan}' in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace('{nama_peralatan}', nama_peralatan)
                                    set_font(run)
        
            # Simpan dokumen ke BytesIO
            output = BytesIO()
            doc.save(output)

            if output:
            # Provide option to download the letter
                st.download_button(
                label = "Unduh Surat",
                data = output.getvalue(),
                file_name = f'{no_srt} Peminjaman {nama_peralatan} {tujuan}.docx',
                mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )

def main():
    st.sidebar.title('Aplikasi Pembuat Surat Otomatis')

    surat = st.sidebar.radio('Jenis Surat', ('Pemberitahuan', 'Peminjaman'))
    
    if surat == 'Pemberitahuan':
        srt_pemberitahuan()
    elif surat == 'Peminjaman':
        srt_peminjaman()

if __name__ == '__main__':
    main()