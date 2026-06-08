from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QGridLayout, QHBoxLayout, QLabel, QScrollArea, QVBoxLayout, QWidget

from widgets.card import Card


HELP_SECTIONS = [
    {
        "title": "BOM Comparison",
        "subtitle": "Compare BOM Excel reference dengan BOM source atau sesama BOM Excel TXT.",
        "steps": [
            "Pilih mode TXT vs BOM File atau TXT vs TXT.",
            "Pilih file reference TXT pada panel kiri.",
            "Pilih BOM source .tsv/.xlsx/.xls atau source TXT pada panel kanan sesuai mode.",
            "Klik Run Comparison untuk melihat ADD, CNG, dan DEL.",
            "Gunakan Export Results kalau ada hasil NG yang perlu disimpan.",
        ],
        "notes": [
            "Reference TXT dibaca sebagai pasangan Circuit dan Part Number.",
            "Mode TXT vs BOM File membaca source dari kolom Child dan Designators.",
            "Mode TXT vs TXT membaca kedua file sebagai pasangan Circuit dan Part Number.",
        ],
    },
    {
        "title": "Machine Data Audit",
        "subtitle": "Compare data mesin dengan program TXT.",
        "steps": [
            "Pilih machine type: NPM, CM602, atau BM221.",
            "Import file mesin sesuai tipe: .crb untuk NPM, .POS untuk BM221, atau file mesin lain untuk CM602.",
            "Import Program File .txt.",
            "Klik Run Machine Audit lalu cek panel Machine Data dan Program File.",
            "Export Results bila perlu laporan Excel.",
        ],
        "notes": [
            "ADD berarti data ada di mesin/source tapi tidak ada di program target.",
            "CNG berarti Circuit sama, tapi detail part, koordinat, atau angle berbeda.",
            "DEL berarti data ada di program target tapi tidak ada di mesin/source.",
            "Untuk BM221, Sync to .POS dapat membuat file POS hasil sinkron part CNG non-TRAY.",
        ],
    },
    {
        "title": "PLAN",
        "subtitle": "Generate PLAN baru dari PLAN sebelumnya dan PLAN baru.",
        "steps": [
            "Pilih tipe plan: 1ST PLAN, 2ND PLAN, atau 3RD PLAN.",
            "Pilih PLAN sebelumnya.",
            "Pilih PLAN baru.",
            "Klik Generate PLAN dan tentukan lokasi output.",
        ],
        "notes": [
            "Output disimpan sebagai workbook Excel.",
            "Dialog selesai menampilkan jumlah data match dan yang baru masuk.",
            "Untuk 1ST PLAN, data match di kolom V dan data sejajar di kolom G ikut diberi warna.",
        ],
    },
    {
        "title": "Worksheet Comparator",
        "subtitle": "Verifikasi Worksheet feeder dengan file NPM .crb.",
        "steps": [
            "Buka Other Tools lalu pilih Worksheet Comparator.",
            "Pilih WORKSHEET Excel/CSV.",
            "Pilih file NPM .crb.",
            "Klik Compare Data.",
            "Gunakan filter hasil untuk melihat NG ID/P/N, NG QTY, semua NG, atau MATCH.",
        ],
        "notes": [
            "ID 1401 dan 1402 diabaikan saat compare.",
            "Perbedaan qty untuk table 10 dan 11 ditandai aman karena aturan khusus mesin.",
        ],
    },
    {
        "title": "Worksheet vs BOM Comparator",
        "subtitle": "Compare Worksheet feeder dengan BOM .tsv asalnya.",
        "steps": [
            "Buka Other Tools lalu pilih Worksheet vs BOM Comparator.",
            "Pilih file Worksheet .xlsx.",
            "Pilih BOM .tsv dengan part number yang sama.",
            "Klik Run Comparison.",
            "Export Results kalau ada ADD, CNG, atau DEL.",
        ],
        "notes": [
            "Worksheet dibandingkan memakai total CNT per Part Number.",
            "BOM dibandingkan memakai jumlah designator TOP SMT per Part Number.",
            "Row placeholder PCB ID 1401/1402 pada Worksheet di-skip agar tidak menjadi false NG.",
        ],
    },
    {
        "title": "Feeder Mapping Generator",
        "subtitle": "Convert export TXT dari mesin NPM menjadi Excel feeder mapping.",
        "steps": [
            "Buka Other Tools lalu pilih Feeder Mapping Generator.",
            "Pilih file export .txt dari mesin NPM.",
            "Klik Preview Mapping untuk mengecek Table, Slot, Position, Location Code, dan Part Number.",
            "Klik Generate Excel lalu pilih lokasi output.",
        ],
        "notes": [
            "Data diambil dari section PartsData atau PartsDataEx, lalu FeederData dan FixedFeeder.",
            "Feeder kecil dibuat sebagai posisi L/R.",
            "Feeder besar dibuat sebagai Large (2-Rel) atau Extra Large (3-Rel).",
        ],
    },
    {
        "title": "NPM Feeder Compare",
        "subtitle": "Compare setup feeder antara dua export TXT mesin NPM.",
        "steps": [
            "Buka Other Tools lalu pilih NPM Feeder Compare.",
            "Pilih Program A / Reference dari export TXT NPM lama.",
            "Pilih Program B / Target dari export TXT NPM baru.",
            "Klik Compare Feeders untuk melihat ADD, MOVE, CNG, dan DEL.",
            "Gunakan filter status atau search untuk fokus ke location code dan part tertentu.",
            "Export Excel kalau hasil compare perlu disimpan.",
        ],
        "notes": [
            "Data feeder dibaca dengan parser yang sama seperti Feeder Mapping Generator.",
            "MOVE berarti Part Number yang sama pindah location code.",
            "CNG berarti location code sama, tapi Part Number berubah.",
            "ADD berarti feeder hanya ada di Program B, DEL berarti feeder hanya ada di Program A.",
        ],
    },
    {
        "title": "PCBA Model Used Part Component Collector",
        "subtitle": "Collect dan generate Excel Used Part Component dari file Excel program.",
        "steps": [
            "Buka Other Tools lalu pilih Used Part Component.",
            "Pilih Mode 1 untuk collect berdasarkan list PCB Part Number, atau Mode 2 untuk collect program berbeda dalam satu folder PCB.",
            "Untuk Mode 1, pilih Folder Induk PCB lalu paste list PCB Part Number.",
            "Untuk Mode 2, pilih folder yang berisi file Excel program.",
            "Klik Generate Used Part Excel lalu pilih lokasi output.",
        ],
        "notes": [
            "Data part name diambil dari sheet BOM kolom C.",
            "Row dengan kolom F berisi #N/A, #NA, atau N/A tidak ikut diambil.",
            "Sheet MASTER berisi part unik di kolom MASTER/P/N COMPONENT, lalu matrix penggunaan per part number di kolom berikutnya.",
        ],
    },
    {
        "title": "Used Part Component AVG Insert Collector",
        "subtitle": "Collect component dari Excel program sekaligus hitung rata-rata insert per component.",
        "steps": [
            "Buka Other Tools lalu pilih AVG Insert Component.",
            "Pilih Mode 1 untuk summarize berdasarkan PCB Part Number, atau Mode 2 untuk summarize per Excel program.",
            "Untuk Mode 1, pilih Folder Induk PCB lalu paste list PCB Part Number.",
            "Untuk Mode 2, pilih folder yang berisi file Excel program.",
            "Klik Generate AVG Insert Excel lalu pilih lokasi output.",
        ],
        "notes": [
            "Insert dihitung dari jumlah kemunculan part di sheet BOM kolom C sebelum duplicate dihapus.",
            "Row dengan kolom F berisi #N/A, #NA, atau N/A tidak ikut dihitung.",
            "Sheet detail berisi component dan AVG INSERT, sedangkan sheet MASTER menambahkan kolom AVG INSERT di paling kanan.",
            "Nilai AVG INSERT dibulatkan ke angka bulat tanpa digit di belakang koma.",
        ],
    },
    {
        "title": "Component Usage Finder",
        "subtitle": "Cari component part number dipakai di model dan PCB part number apa saja.",
        "steps": [
            "Buka Other Tools lalu pilih Component Usage Finder.",
            "Isi Component P/N yang ingin dicari.",
            "Pilih Folder Induk PCB yang berisi subfolder program PCB.",
            "Klik Search untuk scan file Excel program di semua subfolder.",
            "Cek hasil utama pada Preview Results.",
            "Gunakan Copy Results untuk copy hasil tab-separated ke Excel.",
            "Gunakan Export Excel untuk menyimpan Preview Result dan Scan Log dalam dua sheet berbeda.",
        ],
        "notes": [
            "Pencarian membaca sheet BOM saja dan dilakukan case-insensitive.",
            "File temporary Excel yang diawali ~$ otomatis dilewati.",
            "Model Part Number diparse dari nama file jika ada beberapa model dipisah tanda +.",
            "PCB Part Number dan revision diparse dari nama file atau nama folder PCB.",
            "Scan Log berisi detail source folder, source file, found row, dan file yang dilewati/error.",
        ],
    },
    {
        "title": "Common Feeder Reuse",
        "subtitle": "Cek candidate substitute component aman atau conflict untuk sharing fixed feeder.",
        "steps": [
            "Buka Other Tools lalu pilih Common Feeder Reuse.",
            "Pilih Folder Induk PCB yang berisi semua Excel program/BOM model family yang ingin dicek.",
            "Pilih Fixed Feeder Source dari export NPM .txt atau Excel Feeder Mapping.",
            "Isi Candidate Component P/N kalau hanya ingin cek part tertentu, atau kosongkan untuk scan semua component non-feeder.",
            "Klik Analyze Reuse lalu gunakan filter SAFE, CONFLICT, atau CHECK.",
            "Export Excel untuk menyimpan Compatibility, Do Not Pair List, Component Usage, dan Usage Matrix.",
        ],
        "notes": [
            "SAFE berarti candidate dan main feeder tidak pernah muncul bareng di PCB/model yang discan.",
            "CONFLICT berarti candidate dan main feeder muncul bareng minimal di satu PCB/model, jadi tidak aman ditaruh satu feeder.",
            "CHECK berarti salah satu P/N tidak ditemukan di folder scan, jadi perlu konfirmasi manual.",
        ],
    },
    {
        "title": "Model Fix Feeder Groups",
        "subtitle": "Kelompokkan PCB/model berdasarkan kemiripan component usage di sheet BOM.",
        "steps": [
            "Buka Other Tools lalu pilih Model Fix Feeder Groups.",
            "Pilih Folder Induk PCB yang berisi semua Excel program/BOM model family.",
            "Atur Min Similarity dan Min Shared Parts sesuai seberapa ketat grouping yang diinginkan.",
            "Klik Analyze Groups untuk melihat rekomendasi group fix feeder.",
            "Gunakan filter GROUPED untuk melihat PCB/model yang cocok digabung, atau SHOW ALL untuk melihat single model.",
            "Export Excel untuk menyimpan Recommended Groups, Group Components, Pair Similarity, Model Components, dan Scan Log.",
        ],
        "notes": [
            "Similarity dihitung dari jumlah component yang sama dibanding model dengan jumlah component lebih kecil.",
            "Recommended Fixed Feeder Parts adalah component yang muncul di semua member group.",
            "Group Components menampilkan COMMON dan PARTIAL supaya pilihan fixed feeder bisa dipertajam manual.",
        ],
    },
    {
        "title": "All In One Comparator",
        "subtitle": "Compare NPM, BM, dan BOM dari satu layar.",
        "steps": [
            "Klik Auto-Import Semua File untuk memilih banyak file sekaligus, atau isi picker satu per satu.",
            "Pastikan file source kiri dan target kanan sudah terpasang pasangannya.",
            "Klik Start Compare.",
            "Gunakan Filter Status untuk fokus ke NG only, beda data, ADD, REMOVE, atau semua data.",
        ],
        "notes": [
            "NPM memakai pasangan .crb dan .txt.",
            "BM memakai pasangan .pos dan .txt.",
            "BOM memakai pasangan .tsv/.csv dan TXT BOM target.",
        ],
    },
    {
        "title": "NEW PCB Excel Creator",
        "subtitle": "Generate workbook program SMT untuk new PCB.",
        "steps": [
            "Buka Other Tools lalu pilih NEW PCB Excel Creator.",
            "Pilih CAD Data .txt, BOM .tsv, Excel Part Library, Excel Referensi, dan gambar Gerber PCB.",
            "Isi semua field Excel Identity: model, program PN, PCB PN, revision, WO supply, creator, dan line.",
            "Klik Generate Excel lalu pilih lokasi output.",
            "Lanjutkan proses di Excel setelah dialog selesai muncul.",
        ],
        "notes": [
            "Tombol Paste pada Gerber PCB Image bisa mengambil gambar langsung dari clipboard.",
            "Output berisi sheet MEMO, CAD, DX, dan BOM.",
        ],
    },
    {
        "title": "Get Insert Point",
        "subtitle": "Ambil data Insert Point dari folder PCB berdasarkan PLAN.",
        "steps": [
            "Buka Other Tools lalu pilih Get Insert Point.",
            "Pilih Excel PLAN.",
            "Pilih Folder Induk PCB.",
            "Atur Start Row dan End Row sesuai range PLAN yang ingin diproses.",
            "Klik Generate Insert Point dan tentukan lokasi output.",
        ],
        "notes": [
            "Dialog selesai menampilkan jumlah berhasil dan error.",
            "Periksa output Excel untuk melihat detail hasil tiap row.",
        ],
    },
    {
        "title": "History & Logs",
        "subtitle": "Melihat dan export history compare.",
        "steps": [
            "Buka History & Logs dari sidebar.",
            "Klik Refresh List untuk memuat ulang daftar history.",
            "Pilih satu history lalu klik View Details / Export Selected untuk export hasil lama.",
            "Klik Clear All History hanya kalau semua log sudah tidak diperlukan.",
        ],
        "notes": [
            "History saat ini menyimpan hasil BOM Comparison dan Machine Data Audit.",
        ],
    },
]


class HelpPage(QWidget):
    def __init__(self, theme_manager, parent=None):
        super().__init__(parent)
        self.theme_manager = theme_manager
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        intro = QLabel("Panduan ringkas untuk setiap fitur aplikasi.")
        intro.setObjectName("MutedLabel")
        header.addWidget(intro)
        header.addStretch(1)
        root.addLayout(header)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        content = QWidget()
        grid = QGridLayout(content)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(12)

        for index, section in enumerate(HELP_SECTIONS):
            grid.addWidget(self._section_card(section), index // 2, index % 2)

        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        scroll.setWidget(content)
        root.addWidget(scroll, 1)

    def _section_card(self, section):
        card = Card()
        title = QLabel(section["title"])
        title.setObjectName("SectionTitle")
        subtitle = QLabel(section["subtitle"])
        subtitle.setObjectName("MutedLabel")
        subtitle.setWordWrap(True)
        card.layout.addWidget(title)
        card.layout.addWidget(subtitle)

        steps_title = self._mini_title("Langkah Pakai")
        card.layout.addWidget(steps_title)
        for number, step in enumerate(section["steps"], 1):
            card.layout.addWidget(self._body_label(f"{number}. {step}"))

        if section.get("notes"):
            notes_title = self._mini_title("Catatan")
            card.layout.addWidget(notes_title)
            for note in section["notes"]:
                card.layout.addWidget(self._body_label(f"- {note}"))

        card.layout.addStretch(1)
        return card

    def _mini_title(self, text):
        label = QLabel(text)
        label.setObjectName("GuideMiniTitle")
        return label

    def _body_label(self, text):
        label = QLabel(text)
        label.setObjectName("GuideBody")
        label.setWordWrap(True)
        label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        return label
