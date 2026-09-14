from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


ROOT = Path(r"D:\Project Naskah Karya Perorangan")
OUTPUT = ROOT / "Optimalisasi Kecepatan Klarifikasi Seksi Humas melalui Mekanisme Respons Awal Terverifikasi dalam Penanganan Isu Viral guna Memelihara Kepercayaan Publik.docx"


def set_run_font(run, name="Arial", size=12, bold=None, italic=None):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(0, 0, 0)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_cell_margins(cell, top=90, start=100, bottom=90, end=100):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_table_borders(table, color="BFBFBF", size="6"):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = borders.find(qn(f"w:{edge}"))
        if tag is None:
            tag = OxmlElement(f"w:{edge}")
            borders.append(tag)
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), size)
        tag.set(qn("w:color"), color)


def set_bottom_border(paragraph, color="000000", size="10"):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)


def set_box_border(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "16")
        node.set(qn("w:color"), "000000")
        borders.append(node)
    tc_pr.append(borders)


def add_page_field(paragraph):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])
    set_run_font(run, size=10)


def set_page_number_start(section, start=1):
    sect_pr = section._sectPr
    pg_num = sect_pr.find(qn("w:pgNumType"))
    if pg_num is None:
        pg_num = OxmlElement("w:pgNumType")
        sect_pr.append(pg_num)
    pg_num.set(qn("w:start"), str(start))


doc = Document()

# Core styles derived from the clean-format reference.
normal = doc.styles["Normal"]
normal.font.name = "Arial"
normal._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
normal.font.size = Pt(12)
normal.font.color.rgb = RGBColor(0, 0, 0)
normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
normal.paragraph_format.line_spacing = 1.15
normal.paragraph_format.space_after = Pt(4)

title_style = doc.styles["Title"]
title_style.font.name = "Arial"
title_style._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
title_style._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
title_style.font.size = Pt(15)
title_style.font.bold = True
title_style.font.color.rgb = RGBColor(0, 0, 0)
title_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_style.paragraph_format.space_after = Pt(0)
title_ppr = title_style._element.get_or_add_pPr()
title_border = title_ppr.find(qn("w:pBdr"))
if title_border is not None:
    title_ppr.remove(title_border)

for name, size, align in (("Heading 1", 14, WD_ALIGN_PARAGRAPH.CENTER), ("Heading 2", 12, WD_ALIGN_PARAGRAPH.LEFT), ("Heading 3", 12, WD_ALIGN_PARAGRAPH.LEFT)):
    style = doc.styles[name]
    style.font.name = "Arial"
    style._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
    style._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    style.font.size = Pt(size)
    style.font.bold = True
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.alignment = align
    style.paragraph_format.space_before = Pt(8 if name != "Heading 1" else 0)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.keep_with_next = True


def body(text, indent=True, bold_lead=None):
    p = doc.add_paragraph(style="Normal")
    p.paragraph_format.first_line_indent = Inches(0.5) if indent else None
    if bold_lead and text.startswith(bold_lead):
        r1 = p.add_run(bold_lead)
        set_run_font(r1, bold=True)
        r2 = p.add_run(text[len(bold_lead):])
        set_run_font(r2)
    else:
        r = p.add_run(text)
        set_run_font(r)
    return p


def heading(text, level=2):
    p = doc.add_heading(text, level=level)
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in p.runs:
        set_run_font(run, size=14 if level == 1 else 12, bold=True)
    return p


def item(text, marker, level=0):
    p = doc.add_paragraph(style="Normal")
    p.paragraph_format.left_indent = Inches(0.35 + level * 0.25)
    p.paragraph_format.first_line_indent = Inches(-0.25)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.keep_together = True
    r = p.add_run(f"{marker}  {text}")
    set_run_font(r)
    return p


def table(caption, headers, rows, widths_cm):
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_before = Pt(6)
    cap.paragraph_format.space_after = Pt(4)
    cap.paragraph_format.keep_with_next = True
    r = cap.add_run(caption)
    set_run_font(r, size=11, bold=True)
    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    set_table_borders(tbl)
    for i, (cell, header, width) in enumerate(zip(tbl.rows[0].cells, headers, widths_cm)):
        cell.width = Cm(width)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_margins(cell)
        shade_cell(cell, "D9E2F3")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.keep_with_next = True
        rr = p.add_run(header)
        set_run_font(rr, size=10, bold=True)
    set_repeat_table_header(tbl.rows[0])
    for ri, row_data in enumerate(rows):
        cells = tbl.add_row().cells
        for ci, (cell, value, width) in enumerate(zip(cells, row_data, widths_cm)):
            cell.width = Cm(width)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            if ri % 2 == 1:
                shade_cell(cell, "F7F9FC")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            rr = p.add_run(str(value))
            set_run_font(rr, size=10)
        tr_pr = tbl.rows[-1]._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        tr_pr.append(cant_split)
    after = doc.add_paragraph()
    after.paragraph_format.space_before = Pt(0)
    after.paragraph_format.space_after = Pt(0)
    after.paragraph_format.line_spacing = Pt(1)
    return tbl


# Cover section.
cover_section = doc.sections[0]
cover_section.page_width = Cm(21.0)
cover_section.page_height = Cm(29.7)
cover_section.top_margin = Cm(1.5)
cover_section.bottom_margin = Cm(1.5)
cover_section.left_margin = Cm(2.0)
cover_section.right_margin = Cm(2.0)
cover_section.header.is_linked_to_previous = False
cover_section.footer.is_linked_to_previous = False

code_table = doc.add_table(rows=1, cols=2)
code_table.autofit = False
code_table.columns[0].width = Cm(13.3)
code_table.columns[1].width = Cm(3.7)
right = code_table.cell(0, 1)
set_box_border(right)
right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
right.paragraphs[0].paragraph_format.space_after = Pt(0)
set_run_font(right.paragraphs[0].add_run("NKP : [NOMOR NKP]"), size=12, bold=True)

for _ in range(5):
    doc.add_paragraph()

p = doc.add_paragraph(style="Title")
set_run_font(p.add_run("NASKAH KARYA PERORANGAN ( N K P )"), size=15, bold=True)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(18)
p.paragraph_format.space_after = Pt(4)
set_run_font(p.add_run("TOPIK [NOMOR TOPIK]"), size=14, bold=True)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(24)
set_run_font(p.add_run("[URAIAN TOPIK]"), size=13, bold=True)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(8)
set_run_font(p.add_run("JUDUL"), size=14, bold=True)
p = doc.add_paragraph(style="Title")
title = (
    "OPTIMALISASI KECEPATAN KLARIFIKASI SEKSI HUMAS MELALUI MEKANISME RESPONS AWAL TERVERIFIKASI "
    "DALAM PENANGANAN ISU VIRAL GUNA MEMELIHARA KEPERCAYAAN PUBLIK"
)
set_run_font(p.add_run(title), size=15, bold=True)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(18)
p.paragraph_format.space_after = Pt(18)
set_run_font(p.add_run("[LAMBANG LEMBAGA]"), size=12, bold=True)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_run_font(p.add_run("OLEH"), size=13, bold=True)

author = doc.add_table(rows=3, cols=3)
author.alignment = WD_TABLE_ALIGNMENT.CENTER
author.autofit = False
for row, label, value in (
    (0, "NAMA", "[NAMA PESERTA DIDIK DAN GELAR]"),
    (1, "NOMOR SERDIK", "[NOMOR SERDIK]"),
    (2, "POKJAR", "[POKJAR]"),
):
    author.cell(row, 0).width = Cm(4.2)
    author.cell(row, 1).width = Cm(0.4)
    author.cell(row, 2).width = Cm(7.3)
    author.cell(row, 0).text = label
    author.cell(row, 1).text = ":"
    author.cell(row, 2).text = value
    for ci, cell in enumerate(author.rows[row].cells):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for run in cell.paragraphs[0].runs:
            set_run_font(run, size=12, bold=True)
        cell.paragraphs[0].paragraph_format.space_after = Pt(0)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(20)
p.paragraph_format.space_after = Pt(4)
set_bottom_border(p)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(0)
set_run_font(p.add_run("[PROGRAM PENDIDIKAN]"), size=13, bold=True)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(0)
set_run_font(p.add_run("[ANGKATAN DAN TAHUN ANGGARAN]"), size=13, bold=True)

# Body section.
body_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
body_section.page_width = Cm(21.0)
body_section.page_height = Cm(29.7)
body_section.top_margin = Inches(1)
body_section.bottom_margin = Inches(1)
body_section.left_margin = Inches(1)
body_section.right_margin = Inches(1)
body_section.header_distance = Inches(0.35)
body_section.footer_distance = Inches(0.5)
body_section.header.is_linked_to_previous = False
body_section.footer.is_linked_to_previous = False
set_page_number_start(body_section, 1)

header = body_section.header
header_table = header.add_table(rows=1, cols=3, width=Inches(6.27))
header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
header_table.autofit = False
widths = (Cm(7.0), Cm(3.0), Cm(6.0))
for column, width in zip(header_table.columns, widths):
    column.width = width
for cell, width in zip(header_table.rows[0].cells, widths):
    cell.width = width
left_p = header_table.cell(0, 0).paragraphs[0]
left_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
left_p.paragraph_format.space_after = Pt(0)
set_run_font(left_p.add_run("KEPOLISIAN NEGARA REPUBLIK INDONESIA\nDAERAH [NAMA POLDA]\nRESOR [NAMA POLRES]"), size=9, bold=True)
set_bottom_border(left_p, size="6")
page_p = header_table.cell(0, 1).paragraphs[0]
page_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
page_p.paragraph_format.space_after = Pt(0)
add_page_field(page_p)

# BAB I
heading("BAB I\nPENDAHULUAN", 1)
heading("1. Latar Belakang", 2)
body("Polri menjalankan fungsi pemeliharaan keamanan dan ketertiban, penegakan hukum, serta perlindungan, pengayoman, dan pelayanan kepada masyarakat [1]. Pelaksanaan fungsi tersebut berada dalam perhatian publik. Pada 2026, APJII mencatat 235,26 juta pengguna internet atau 81,72 persen penduduk Indonesia telah terhubung ke internet [3]. Rekaman warga dan komentar saksi karena itu dapat membentuk opini sebelum laporan kedinasan diterima pimpinan.")
body("Arus informasi tersebut disertai risiko kehilangan konteks dan informasi palsu. Komdigi mengidentifikasi 1.923 konten hoaks sepanjang 2024 [4]. Polda Jawa Timur pada Mei 2026 juga mengingatkan bahwa hoaks dan narasi negatif dapat menyebar dalam hitungan menit [5]. Seksi Humas harus merespons cepat, tetapi keterangannya tetap harus terverifikasi, sah untuk dibuka, dan tidak mengganggu pemeriksaan.")
body("Perka Polri Nomor 6 Tahun 2023 menempatkan pemantauan, analisis media, pengelolaan media sosial, dan layanan informasi terpadu sebagai kompetensi fungsi Humas [2]. Naskah ini mengusulkan Mekanisme Respons Awal Terverifikasi atau MRAT, yaitu pernyataan awal berbasis fakta minimum yang telah dipastikan, disertai proses dan waktu pembaruan. Mekanisme tersebut tidak menyimpulkan perkara sebelum pemeriksaan selesai.")

heading("2. Identifikasi Persoalan", 2)
item("Deteksi dan eskalasi isu belum selalu memiliki target waktu serta klasifikasi risiko yang seragam.", "a.")
item("Verifikasi dan otorisasi melibatkan beberapa fungsi, tetapi narahubung serta batas waktunya belum selalu jelas.", "b.")
item("Format pernyataan awal, jadwal pembaruan, prosedur koreksi, dan catatan keputusan belum digunakan pada setiap situasi.", "c.")

heading("3. Rumusan Masalah", 2)
body("Bagaimana mengoptimalkan kecepatan klarifikasi Seksi Humas melalui Mekanisme Respons Awal Terverifikasi dalam penanganan isu viral guna memelihara kepercayaan publik?", indent=False)

heading("4. Ruang Lingkup", 2)
body("Pembahasan dibatasi pada komunikasi awal atas isu viral mengenai pelaksanaan tugas atau dugaan pelanggaran anggota pada tingkat Polres. Pemeriksaan pidana, kode etik, pengamanan siber teknis, dan penindakan konten hanya dibahas sepanjang berkaitan dengan koordinasi informasi publik.")

heading("5. Maksud dan Tujuan", 2)
item("Maksud penulisan adalah menganalisis kesenjangan antara kecepatan informasi publik dan proses verifikasi internal.", "a.")
item("Tujuannya adalah menyusun alur respons yang terukur, pembagian peran, serta dasar evaluasi penanganan isu viral.", "b.")

heading("6. Pendekatan Penulisan", 2)
body("Naskah menggunakan pendekatan deskriptif-analitis melalui telaah peraturan, literatur, data publik, dan praktik komunikasi kepolisian. Penerapannya perlu dilengkapi audit isu enam bulan, wawancara fungsi terkait, serta pengukuran waktu respons untuk mengganti kolom isian yang tersedia.")

# BAB II
heading("BAB II\nPEMBAHASAN", 1)
heading("1. Landasan Hukum dan Konseptual", 2)
body("Landasan hukum MRAT bertumpu pada kewajiban Polri memberikan pelayanan kepada masyarakat [1], pengoordinasian fungsi kehumasan oleh Humas Polri [2], serta kewajiban badan publik menyediakan informasi secara terbuka dengan pengecualian yang ketat dan terbatas [10]. Peraturan Komisi Informasi Nomor 1 Tahun 2021 mengatur standar layanan informasi publik dan memperjelas perlunya pengelolaan informasi melalui sarana elektronik [11]. Dengan demikian, kecepatan respons tidak dapat dipisahkan dari ketepatan klasifikasi informasi.")
body("Coombs menjelaskan komunikasi krisis sebagai bagian dari pengelolaan krisis yang meliputi persiapan, respons, dan pembelajaran setelah kejadian [6]. Situational Crisis Communication Theory menempatkan jenis krisis, atribusi tanggung jawab, riwayat organisasi, dan dampak terhadap pemangku kepentingan sebagai dasar pemilihan respons [7]. Untuk isu yang memuat dugaan pelanggaran anggota, respons defensif yang mendahului pemeriksaan dapat memperbesar ancaman reputasi karena publik menilai institusi dari kesediaannya mengakui masalah, menjelaskan proses, dan memperbaiki kesalahan.")
body("Reynolds dan Seeger menekankan bahwa komunikasi pada tahap awal krisis berlangsung dalam keadaan tidak pasti [8]. Keterbatasan fakta bukan alasan untuk diam. Institusi perlu menyampaikan apa yang telah diketahui, apa yang belum diketahui, tindakan yang sedang dilakukan, dan kapan informasi berikutnya diberikan. Prinsip ini menjadi dasar pernyataan awal dalam MRAT.")
body("Kepercayaan publik dalam naskah ini tidak dipahami sebagai hasil pencitraan jangka pendek. Tyler menunjukkan bahwa legitimasi otoritas hukum berkaitan dengan penilaian masyarakat terhadap keadilan prosedur dan kelayakan otoritas untuk ditaati [9]. Karena itu, respons Humas harus mendukung akuntabilitas tindakan kepolisian, bukan hanya menurunkan jumlah komentar negatif. Survei Litbang Kompas yang dipublikasikan pada Juni 2026 mencatat kepercayaan terhadap Polri sebesar 82,4 persen [12]. Angka agregat tersebut perlu dipelihara melalui konsistensi pelayanan informasi pada setiap peristiwa, sambil tetap membedakan survei nasional dari pengalaman publik pada satu kasus.")

heading("2. Kondisi Faktual", 2)
body("Kondisi komunikasi saat ini memperlihatkan ketimpangan kecepatan. Masyarakat dapat mengunggah bukti visual seketika, sedangkan informasi institusi harus melewati pemeriksaan identitas, waktu, tempat, konteks, dan keterkaitan dengan proses hukum. Divisi Humas Polri pada Agustus 2026 mengingatkan masyarakat untuk memeriksa sumber, waktu, dan konteks video lama yang diedarkan dengan cerita baru [13]. Peringatan tersebut menunjukkan bahwa verifikasi konteks merupakan pekerjaan utama sebelum klarifikasi diterbitkan.")
body("Pada tingkat kesatuan, isu biasanya pertama kali diketahui melalui pemantauan media sosial, laporan anggota, pertanyaan wartawan, atau pengaduan masyarakat. Kelemahan muncul apabila hasil pemantauan tidak langsung disertai klasifikasi risiko dan penetapan penanggung jawab. Semua isu kemudian diproses dengan cara yang sama, padahal pertanyaan layanan biasa tidak membutuhkan jalur otorisasi yang sama dengan video dugaan kekerasan atau peristiwa yang berpotensi mengganggu kamtibmas.")
body("Kesenjangan berikutnya terdapat pada perolehan fakta minimum. Humas tidak selalu menguasai informasi pertama karena fakta berada pada penyidik, fungsi operasional, Propam, atau Polsek tempat kejadian. Di sisi lain, fungsi teknis dapat menunda pemberian informasi karena khawatir mengganggu pemeriksaan. Apabila belum ada daftar fakta yang boleh disampaikan, keputusan sering berubah menjadi pilihan biner antara membuka seluruh informasi dan tidak memberi keterangan. Pilihan tersebut tidak sesuai dengan kebutuhan komunikasi awal.")
body("Kondisi lokal harus dibuktikan sebelum program dilaksanakan. Audit awal sekurang-kurangnya mencatat `[jumlah isu viral enam bulan terakhir]`, median waktu sejak isu terdeteksi sampai dilaporkan kepada pimpinan `[menit]`, median waktu sampai pernyataan pertama `[menit atau jam]`, jumlah koreksi yang diterbitkan `[jumlah]`, serta saluran yang paling sering menjadi sumber isu `[platform]`. Tanpa baseline tersebut, keberhasilan hanya dinilai dari banyaknya konten yang dipublikasikan, bukan dari kualitas respons.")

heading("3. Kondisi yang Diharapkan", 2)
body("Kondisi yang diharapkan adalah tersedianya proses yang memungkinkan Humas merespons cepat tanpa mendahului hasil pemeriksaan. Setiap isu masuk ke satu pencatatan, memperoleh tingkat prioritas, diteruskan kepada narahubung fungsi terkait, dan menghasilkan fakta minimum yang dapat dipublikasikan. Jika fakta belum cukup, Humas tetap dapat mengeluarkan pernyataan awal yang mengakui adanya informasi, menyebut proses verifikasi, dan menetapkan saluran pembaruan.")
body("Kondisi tersebut memerlukan satu sumber informasi resmi pada setiap kejadian. Kesatuan tidak harus membatasi komunikasi pada satu platform, tetapi isi pokok pada situs, akun media sosial, grup media, dan keterangan juru bicara harus konsisten. Perubahan fakta harus dicatat, diperbaiki secara terbuka, dan tidak dihapus tanpa penjelasan apabila informasi sebelumnya telah tersebar.")

heading("4. Faktor yang Memengaruhi", 2)
item("Faktor internal meliputi kompetensi pemantauan dan analisis media; ketersediaan petugas piket; kejelasan narahubung fungsi teknis; kewenangan persetujuan; format fakta minimum; akses terhadap perangkat dan akun resmi; serta budaya organisasi dalam melaporkan masalah sejak awal.", "a.")
item("Faktor eksternal meliputi algoritma platform, akun berpengaruh, potongan video tanpa konteks, tekanan tenggat media, koordinasi dengan pemerintah daerah atau instansi lain, tingkat literasi digital masyarakat, serta kemungkinan penggunaan isu oleh pihak yang berkepentingan.", "b.")
body("Faktor internal lebih dahulu menjadi sasaran perbaikan karena berada dalam kendali kesatuan. Perubahan algoritma atau perilaku warganet tidak dapat dikendalikan oleh Humas, sedangkan daftar narahubung, target waktu, format pernyataan, dan pencatatan keputusan dapat ditetapkan melalui tata kerja internal.")

heading("5. Analisis SWOT", 2)
table(
    "Tabel 1 Analisis SWOT Penanganan Isu Viral",
    ["Unsur", "Temuan Utama", "Implikasi Strategis"],
    [
        ["Kekuatan", "Akun resmi, kewenangan memperoleh informasi kedinasan, jaringan media, dan dasar hukum fungsi Humas.", "Menjadikan akun resmi sebagai rujukan fakta yang konsisten."],
        ["Kelemahan", "Verifikasi lintas fungsi dan otorisasi dapat memerlukan waktu; kompetensi serta piket digital belum selalu merata.", "Menyederhanakan fakta minimum, narahubung, dan jenjang persetujuan."],
        ["Peluang", "Jangkauan internet luas, dukungan media lokal, dan tersedianya alat pemantauan digital.", "Mendistribusikan pembaruan resmi secara cepat dan serentak."],
        ["Ancaman", "Hoaks, potongan video, akun anonim, serangan komentar, kebocoran data, dan politisasi isu.", "Menerapkan klasifikasi risiko, koreksi terbuka, dan perlindungan informasi."],
    ],
    [2.5, 7.2, 7.0],
)
body("Strategi yang dipilih adalah menggunakan kekuatan kelembagaan dan saluran resmi untuk mengatasi kelemahan koordinasi internal. Fokus awal bukan pengadaan aplikasi baru. Kesatuan dapat memakai perangkat pemantauan dan komunikasi yang sudah tersedia, kemudian menambahkan alat apabila evaluasi menunjukkan kebutuhan yang tidak dapat dipenuhi oleh sistem berjalan.")

heading("6. Strategi Pemecahan Masalah", 2)
item("Menetapkan keputusan atau surat perintah tentang MRAT yang memuat klasifikasi isu, petugas piket, narahubung, fakta minimum, kewenangan persetujuan, target waktu, dan kewajiban dokumentasi.", "a.")
item("Menyusun daftar kontak aktif Humas, Satreskrim, Satintelkam, Samapta, Lantas, Propam, SPKT, PPID, Polsek, dan pimpinan. Setiap fungsi menunjuk pejabat utama dan pengganti.", "b.")
item("Melatih petugas melalui simulasi singkat berdasarkan kasus yang pernah muncul. Simulasi menguji kecepatan memperoleh fakta, konsistensi pesan, ketepatan klasifikasi informasi, dan prosedur koreksi.", "c.")
item("Membangun hubungan kerja dengan media melalui satu kanal konfirmasi dan jadwal pembaruan. Wartawan memperoleh kepastian proses meskipun substansi perkara belum seluruhnya dapat disampaikan.", "d.")

heading("7. Mekanisme Respons Awal Terverifikasi", 2)
body("MRAT terdiri atas sembilan tahap. Pertama, petugas mendeteksi isu dan menyimpan tautan, tangkapan layar, waktu unggahan, serta konteks awal. Kedua, isu diklasifikasikan menurut jangkauan, dampak hukum, keselamatan, dan reputasi. Ketiga, petugas mengeskalasi isu kepada Kasi Humas dan narahubung fungsi terkait. Keempat, fungsi teknis mengonfirmasi fakta minimum: waktu, tempat, keterlibatan personel atau sarana, tindakan yang telah dilakukan, dan informasi yang belum dapat dibuka.")
body("Kelima, Humas menyusun pernyataan awal. Keenam, pejabat yang berwenang memberi persetujuan sesuai tingkat isu. Ketujuh, informasi dipublikasikan pada kanal resmi dan disampaikan kepada media. Kedelapan, pembaruan diberikan pada waktu yang telah dinyatakan atau ketika ada fakta penting baru. Kesembilan, seluruh keputusan, versi pesan, dan hasil evaluasi dicatat untuk perbaikan prosedur.")
table(
    "Tabel 2 Klasifikasi Awal Isu",
    ["Tingkat", "Kriteria", "Penanganan Awal"],
    [
        ["1 Rendah", "Pertanyaan layanan, keluhan terbatas, atau informasi salah dengan jangkauan lokal dan risiko rendah.", "Klarifikasi rutin oleh Humas setelah konfirmasi fungsi terkait."],
        ["2 Sedang", "Isu berkembang lintas akun atau media, melibatkan tindakan petugas, dan berpotensi menurunkan kepercayaan.", "Eskalasi kepada Kasi Humas dan pimpinan; pernyataan awal serta pembaruan terjadwal."],
        ["3 Tinggi", "Korban jiwa, dugaan pelanggaran berat, gangguan kamtibmas, perhatian nasional, atau keterlibatan banyak instansi.", "Tim krisis; verifikasi lintas fungsi; juru bicara tunggal; pengawasan Propam atau penyidik sesuai perkara."],
    ],
    [2.6, 8.0, 6.1],
)

heading("8. Pembagian Peran dan Otorisasi", 2)
table(
    "Tabel 3 Pembagian Peran dalam MRAT",
    ["Unsur", "Tanggung Jawab"],
    [
        ["Seksi Humas", "Mendeteksi, mencatat, mengklasifikasikan, menyusun pesan, mengelola media, memublikasikan, dan mengevaluasi respons."],
        ["Fungsi teknis atau penyidik", "Memastikan fakta minimum dan menandai informasi yang belum dapat dibuka karena kepentingan pemeriksaan atau ketentuan hukum."],
        ["Propam", "Memberi konfirmasi proses pemeriksaan internal pada isu dugaan pelanggaran anggota tanpa mendahului hasil pemeriksaan."],
        ["Pimpinan", "Menetapkan arah, menyetujui isu tingkat tinggi, dan menentukan juru bicara."],
        ["PPID", "Memberi pertimbangan klasifikasi informasi publik dan pencatatan permintaan informasi."],
        ["Operator media sosial", "Menerbitkan versi yang telah disetujui, memantau tanggapan, menyimpan bukti publikasi, dan meneruskan pertanyaan penting."],
    ],
    [4.5, 12.2],
)
body("Otorisasi dibuat bertingkat. Isu tingkat rendah dapat disetujui Kasi Humas berdasarkan fakta dari fungsi terkait. Isu tingkat sedang memerlukan persetujuan pejabat yang ditunjuk pimpinan. Isu tingkat tinggi memerlukan arahan Kapolres atau pejabat pengganti dan koordinasi dengan Polda apabila jangkauannya melampaui wilayah. Jenjang tersebut harus dicantumkan dalam surat perintah agar petugas tidak mencari persetujuan baru pada setiap kejadian.")

heading("9. Target Layanan Awal", 2)
body("Target berikut merupakan usulan awal dan harus disesuaikan setelah baseline kesatuan tersedia. Target mengukur proses komunikasi, bukan memaksa penyidik menyelesaikan pemeriksaan dalam waktu singkat.")
table(
    "Tabel 4 Usulan Target Layanan MRAT",
    ["Tahap", "Target Awal", "Bukti"],
    [
        ["Deteksi dan pencatatan", "Paling lambat 15 menit sejak diketahui petugas piket.", "Log isu dan tautan sumber."],
        ["Eskalasi", "Paling lambat 30 menit untuk tingkat sedang dan tinggi.", "Waktu pesan atau tiket internal."],
        ["Pernyataan awal", "Paling lambat 60 menit jika fakta minimum telah dikonfirmasi.", "Tautan publikasi dan lembar persetujuan."],
        ["Pembaruan", "Sesuai waktu yang dijanjikan atau segera setelah fakta penting terverifikasi.", "Riwayat versi dan waktu publikasi."],
        ["Koreksi", "Segera setelah kesalahan diketahui dan disetujui.", "Catatan koreksi terbuka."],
    ],
    [5.0, 6.2, 5.5],
)

heading("10. Rencana Implementasi", 2)
table(
    "Tabel 5 Rencana Aksi",
    ["Waktu", "Kegiatan", "Keluaran"],
    [
        ["0 sampai 3 bulan", "Audit enam bulan isu terdahulu; penetapan tim; daftar narahubung; klasifikasi isu; format pernyataan dan log.", "Baseline, surat perintah, daftar kontak, dan perangkat kerja awal."],
        ["4 sampai 6 bulan", "Uji coba pada isu nyata; dua simulasi lintas fungsi; evaluasi target waktu; penguatan hubungan media.", "Laporan uji coba, hasil simulasi, dan revisi mekanisme."],
        ["7 sampai 12 bulan", "Penerapan penuh; evaluasi triwulanan; integrasi hasil ke analisis kinerja dan kebutuhan pelatihan.", "Laporan kinerja, rekomendasi pimpinan, dan mekanisme yang diperbarui."],
    ],
    [3.3, 8.2, 5.2],
)

heading("11. Monitoring dan Evaluasi", 2)
body("Evaluasi dilakukan setiap bulan dan setelah isu tingkat tinggi. Indikator utama meliputi median waktu deteksi, median waktu eskalasi, persentase pernyataan awal yang memenuhi target, jumlah koreksi karena kesalahan internal, kepatuhan pada jadwal pembaruan, konsistensi pesan antarplatform, jumlah pertanyaan media yang terjawab, dan perubahan pola sentimen setelah fakta dipublikasikan. Sentimen digunakan sebagai petunjuk, bukan satu-satunya ukuran keberhasilan.")
body("Kualitas respons dinilai melalui pemeriksaan sampel: apakah sumber fakta tercatat, informasi yang dikecualikan terlindungi, bahasa tidak menghakimi, pihak terdampak diperlakukan dengan empati, dan koreksi tersedia ketika diperlukan. Hasil evaluasi dibahas oleh Humas bersama fungsi yang paling sering terlibat agar hambatan koordinasi diperbaiki pada sumbernya.")

heading("12. Risiko dan Pengamanan", 2)
item("Kesalahan fakta dikendalikan melalui dua sumber internal untuk isu tingkat tinggi atau satu pejabat pemilik fakta untuk isu lainnya.", "a.")
item("Asas praduga tidak bersalah dijaga dengan membedakan dugaan, fakta terverifikasi, pendapat saksi, dan hasil pemeriksaan.", "b.")
item("Data pribadi, identitas korban tertentu, teknik penyidikan, barang bukti sensitif, dan informasi yang dikecualikan ditelaah bersama penyidik atau PPID [10][11].", "c.")
item("Kritik tidak dihapus hanya karena bernada negatif. Moderasi dibatasi pada spam, ancaman, doxing, pornografi, ujaran kebencian, atau pelanggaran pedoman akun. Laporan warga diarahkan ke kanal pengaduan.", "d.")
item("Setiap versi pernyataan, persetujuan, koreksi, dan waktu publikasi disimpan untuk akuntabilitas dan pembelajaran.", "e.")

# BAB III
doc.add_page_break()
heading("BAB III\nPENUTUP", 1)
heading("1. Kesimpulan", 2)
body("Kecepatan klarifikasi Seksi Humas dapat dioptimalkan apabila kesatuan tidak menunggu seluruh pemeriksaan selesai untuk mengakui adanya isu. Respons awal dapat disampaikan setelah fakta minimum terverifikasi, sepanjang isinya tidak mendahului kesimpulan hukum, membuka informasi yang dikecualikan, atau menutupi kesalahan. Masalah pokok yang harus diselesaikan terletak pada alur deteksi, koordinasi fakta, otorisasi, dan pembaruan yang belum terukur.")
body("MRAT menjawab masalah tersebut melalui sembilan tahap, klasifikasi tiga tingkat, pembagian peran lintas fungsi, target layanan, format pernyataan awal, dan dokumentasi keputusan. Mekanisme ini menempatkan Humas sebagai pengelola proses komunikasi, sedangkan kepastian fakta tetap berasal dari fungsi yang berwenang. Dengan pembagian tersebut, tuntutan kecepatan dan kewajiban akurasi tidak dipertentangkan.")
body("Kepercayaan publik dipelihara melalui keterbukaan yang proporsional, bahasa yang tidak menghakimi, kepastian waktu pembaruan, dan kesediaan melakukan koreksi. Banyaknya konten positif tidak dapat menggantikan akuntabilitas pada saat krisis. Keberhasilan karena itu harus diukur dari ketepatan proses serta kualitas informasi yang diterima masyarakat.")

heading("2. Saran", 2)
item("Kapolres menetapkan MRAT melalui surat perintah atau keputusan internal setelah menyesuaikan target waktu dengan baseline dan struktur kesatuan.", "a.")
item("Kasi Humas melaksanakan audit enam bulan isu terdahulu, memperbarui daftar narahubung setiap bulan, dan menyelenggarakan simulasi lintas fungsi sekurang-kurangnya setiap semester.", "b.")
item("Kasatfung dan Kapolsek menunjuk pejabat pemilik fakta serta pengganti yang dapat dihubungi oleh Humas pada waktu dinas maupun dalam keadaan mendesak.", "c.")
item("PPID dan fungsi hukum memberi pendampingan untuk klasifikasi informasi, perlindungan data pribadi, dan prosedur koreksi agar kecepatan publikasi tetap berada dalam batas hukum.", "d.")
item("Hasil evaluasi MRAT dimasukkan ke dalam analisis kebutuhan pelatihan, pengawasan, dan rencana kerja kesatuan. Pengadaan perangkat baru dilakukan hanya jika evaluasi membuktikan bahwa alat yang tersedia tidak memenuhi kebutuhan pemantauan atau dokumentasi.", "e.")

# Bibliography.
doc.add_page_break()
heading("DAFTAR PUSTAKA", 1)
references = [
    "Undang-Undang Republik Indonesia Nomor 2 Tahun 2002 tentang Kepolisian Negara Republik Indonesia.",
    "Kepolisian Negara Republik Indonesia. Peraturan Kepala Kepolisian Negara Republik Indonesia Nomor 6 Tahun 2023 tentang Penyelenggaraan Kehumasan di Lingkungan Kepolisian Negara Republik Indonesia. https://peraturan.bpk.go.id/Details/326719, diakses 10 September 2026.",
    "Asosiasi Penyelenggara Jasa Internet Indonesia. 2026. Tiga Dekade APJII Dorong Konektivitas Inklusif hingga Wilayah 3T. https://apjii.or.id/berita/d/tiga-dekade-apjii-dorong-konektivitas-inklusif-hingga-wilayah-3t, diakses 10 September 2026.",
    "Kementerian Komunikasi dan Digital Republik Indonesia. 2025. Komdigi Identifikasi 1.923 Konten Hoaks Sepanjang Tahun 2024. https://www.komdigi.go.id/berita/siaranpers/detail/komdigi-identifikasi-1923-konten-hoaks-sepanjang-tahun-2024, diakses 10 September 2026.",
    "Kepolisian Daerah Jawa Timur. 2026. Polda Jatim Gelar Rakernis Humas, Dorong Penguatan Komunikasi Publik dan Manajemen Media. https://tribratanews.jatim.polri.go.id/polda-jatim-gelar-rakernis-humasdorong-penguatan-komunikasi-publik-dan-manajemen-media, diakses 10 September 2026.",
    "Coombs, W. Timothy. 2023. Ongoing Crisis Communication: Planning, Managing, and Responding. Edisi ke-6. Thousand Oaks: SAGE Publications.",
    "Coombs, W. Timothy. 2007. Protecting Organization Reputations During a Crisis: The Development and Application of Situational Crisis Communication Theory. Corporate Reputation Review, 10(3), 163-176. https://doi.org/10.1057/palgrave.crr.1550049.",
    "Reynolds, Barbara, dan Matthew W. Seeger. 2005. Crisis and Emergency Risk Communication as an Integrative Model. Journal of Health Communication, 10(1), 43-55. https://doi.org/10.1080/10810730590904571.",
    "Tyler, Tom R. 2006. Why People Obey the Law. Princeton: Princeton University Press. https://doi.org/10.2307/j.ctv1j66769.",
    "Undang-Undang Republik Indonesia Nomor 14 Tahun 2008 tentang Keterbukaan Informasi Publik. https://peraturan.bpk.go.id/Details/39047/uu-no-14-, diakses 10 September 2026.",
    "Komisi Informasi Republik Indonesia. Peraturan Komisi Informasi Nomor 1 Tahun 2021 tentang Standar Layanan Informasi Publik. https://eppid.komisiinformasi.go.id/uploads/lampiran/22797PerKINo1Tahun2021.pdf, diakses 10 September 2026.",
    "ANTARA. 2026. Litbang Kompas: Kepercayaan Publik pada Polri Naik Jadi 82,4 Persen. https://www.antaranews.com/berita/5624063/litbang-kompas-kepercayaan-publik-pada-polri-naik-jadi-824-persen, diakses 10 September 2026.",
    "Divisi Humas Polri. 2026. Polri Ingatkan Masyarakat Jangan Tertipu Video Lama dengan Cerita Baru. https://tribratanews.polri.go.id/blog/nasional-3/polri-ingatkan-masyarakat-jangan-tertipu-video-lama-dengan-cerita-baru-104934, diakses 10 September 2026.",
]
for idx, ref in enumerate(references, 1):
    item(ref, f"{idx}.")

# Appendices.
doc.add_page_break()
heading("LAMPIRAN 1\nPOLA PIKIR", 1)
table(
    "Pola Pikir Optimalisasi Kecepatan Klarifikasi",
    ["Masukan", "Proses", "Keluaran", "Hasil"],
    [[
        "Isu viral; fakta awal; peraturan; personel; akun resmi; jaringan media.",
        "Deteksi -> klasifikasi -> verifikasi -> otorisasi -> publikasi -> pembaruan -> evaluasi.",
        "Respons awal terverifikasi, konsisten, terdokumentasi, dan tepat waktu.",
        "Kepastian informasi meningkat dan kepercayaan publik terpelihara.",
    ]],
    [4.2, 5.0, 4.4, 3.1],
)

doc.add_page_break()
heading("LAMPIRAN 2\nALUR MEKANISME RESPONS AWAL TERVERIFIKASI", 1)
flow = [
    "1. Isu terdeteksi dan bukti digital disimpan.",
    "2. Petugas piket mencatat sumber, waktu, jangkauan, dan konteks.",
    "3. Isu diklasifikasikan sebagai tingkat rendah, sedang, atau tinggi.",
    "4. Humas menghubungi narahubung fungsi pemilik fakta.",
    "5. Fakta minimum dan informasi yang belum dapat dibuka ditetapkan.",
    "6. Humas menyusun pernyataan awal dan rencana pembaruan.",
    "7. Pejabat berwenang memberi persetujuan sesuai tingkat isu.",
    "8. Pesan dipublikasikan serentak pada saluran resmi dan media.",
    "9. Tanggapan dipantau; fakta baru diperbarui; kesalahan dikoreksi.",
    "10. Penanganan ditutup dengan evaluasi dan penyimpanan arsip.",
]
for line in flow:
    body(line, indent=False)

doc.add_page_break()
heading("LAMPIRAN 3\nTEMPLATE PERNYATAAN AWAL", 1)
body("Seksi Humas Polres [nama kesatuan] telah mengetahui informasi mengenai [uraian singkat isu] yang beredar pada [platform] sejak [waktu]. Berdasarkan verifikasi awal, fakta yang telah dipastikan adalah [fakta minimum yang dapat dibuka].", indent=False)
body("Saat ini [fungsi yang berwenang] sedang memeriksa [aspek yang masih diverifikasi]. Polres [nama kesatuan] meminta masyarakat tidak menyebarkan identitas, rekaman, atau kesimpulan yang belum terverifikasi karena dapat merugikan pihak terkait dan mengganggu proses pemeriksaan.", indent=False)
body("Informasi berikutnya akan disampaikan melalui [kanal resmi] paling lambat [waktu pembaruan] atau segera setelah terdapat fakta penting yang dapat dipublikasikan. Masyarakat yang memiliki informasi dapat menyampaikannya melalui [kanal pengaduan].", indent=False)
body("Catatan penggunaan: hapus bagian yang tidak relevan; jangan mencantumkan identitas yang dilindungi; bedakan dugaan dan fakta; serta pastikan waktu pembaruan dapat dipenuhi.", indent=False)

doc.add_page_break()
heading("LAMPIRAN 4\nFORMAT LOG PEMANTAUAN DAN EVALUASI", 1)
table(
    "Log Penanganan Isu Viral",
    ["Tanggal dan Waktu", "Tahap", "Tindakan atau Fakta", "PIC", "Bukti atau Tautan"],
    [
        ["", "Deteksi", "", "", ""],
        ["", "Klasifikasi", "", "", ""],
        ["", "Verifikasi", "", "", ""],
        ["", "Otorisasi", "", "", ""],
        ["", "Publikasi", "", "", ""],
        ["", "Pembaruan", "", "", ""],
        ["", "Koreksi", "", "", ""],
        ["", "Evaluasi", "", "", ""],
    ],
    [3.0, 2.5, 5.5, 2.3, 3.4],
)

# Keep placeholders visible, request field refresh on open, and set document metadata.
settings = doc.settings._element
update_fields = settings.find(qn("w:updateFields"))
if update_fields is None:
    update_fields = OxmlElement("w:updateFields")
    settings.append(update_fields)
update_fields.set(qn("w:val"), "true")
doc.core_properties.title = "Optimalisasi Kecepatan Klarifikasi Seksi Humas melalui Mekanisme Respons Awal Terverifikasi"
doc.core_properties.subject = "Naskah Karya Perorangan tentang penanganan isu viral oleh Seksi Humas"
doc.core_properties.author = "[NAMA PESERTA DIDIK]"
doc.core_properties.keywords = "Humas Polri; komunikasi krisis; isu viral; kepercayaan publik"

doc.save(OUTPUT)
print(OUTPUT)
