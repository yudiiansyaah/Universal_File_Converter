"""
Universal File Converter — Windows 11 Desktop App
Dibuat dengan PySide6. Mendukung konversi batch untuk:
Dokumen, Spreadsheet, Presentasi, Gambar, Audio, Video, Arsip, Data Terstruktur.
"""
import os
import sys
import traceback
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTableWidget, QTableWidgetItem, QComboBox, QFileDialog,
    QLabel, QLineEdit, QPlainTextEdit, QProgressBar, QHeaderView,
    QMessageBox, QAbstractItemView, QSplitter, QFrame
)

from core import format_map, converters, dependencies

APP_TITLE = "Universal File Converter"
COL_FILE, COL_SRC, COL_TARGET, COL_STATUS = range(4)


class ConversionWorker(QThread):
    progress = Signal(int, str, bool)   # row, message, success
    log_line = Signal(str)
    finished_all = Signal()

    def __init__(self, jobs, output_dir):
        super().__init__()
        self.jobs = jobs  # list of (row, input_path, target_ext)
        self.output_dir = output_dir
        self._stop = False

    def run(self):
        for row, input_path, target_ext in self.jobs:
            if self._stop:
                break
            try:
                out_dir = self.output_dir or str(Path(input_path).parent)
                result_path = converters.convert_file(
                    input_path, target_ext, out_dir, log=self.log_line.emit
                )
                self.progress.emit(row, f"Selesai -> {os.path.basename(result_path)}", True)
            except Exception as e:
                self.log_line.emit(f"ERROR: {e}\n{traceback.format_exc()}")
                self.progress.emit(row, f"Gagal: {e}", False)
        self.finished_all.emit()

    def stop(self):
        self._stop = True


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(980, 640)
        self.file_rows = []  # list of dict: {"path": str, "target_combo": QComboBox}
        self.output_dir = ""
        self.worker = None

        self._build_ui()
        self._build_menu_bar()
        self._check_dependencies_banner()

    # ------------------------------------------------------------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Banner dependensi
        self.dep_banner = QLabel("")
        self.dep_banner.setStyleSheet(
            "background:#3a2f00; color:#ffd76b; padding:8px; border-radius:4px;"
        )
        self.dep_banner.setWordWrap(True)
        self.dep_banner.hide()
        layout.addWidget(self.dep_banner)

        # Toolbar atas
        toolbar = QHBoxLayout()
        btn_add = QPushButton("+ Tambah File")
        btn_add.clicked.connect(self.add_files)
        btn_remove = QPushButton("Hapus Terpilih")
        btn_remove.clicked.connect(self.remove_selected)
        btn_clear = QPushButton("Bersihkan Semua")
        btn_clear.clicked.connect(self.clear_all)
        toolbar.addWidget(btn_add)
        toolbar.addWidget(btn_remove)
        toolbar.addWidget(btn_clear)
        toolbar.addStretch()
        layout.addLayout(toolbar)

        # Tabel file
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["File", "Format Asal", "Konversi ke", "Status"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setSectionResizeMode(COL_FILE, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(COL_SRC, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(COL_TARGET, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(COL_STATUS, QHeaderView.Stretch)
        layout.addWidget(self.table, stretch=3)

        # Output folder
        out_row = QHBoxLayout()
        out_row.addWidget(QLabel("Folder Output:"))
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("(kosong = simpan di folder sumber masing-masing file)")
        out_row.addWidget(self.output_edit, stretch=1)
        btn_browse_out = QPushButton("Pilih Folder...")
        btn_browse_out.clicked.connect(self.choose_output_dir)
        out_row.addWidget(btn_browse_out)
        layout.addLayout(out_row)

        # Convert button + progress
        action_row = QHBoxLayout()
        self.btn_convert = QPushButton("Konversi Semua")
        self.btn_convert.setStyleSheet(
            "font-weight:bold; padding:8px 16px; background:#2f6feb; color:white; border-radius:4px;"
        )
        self.btn_convert.clicked.connect(self.start_conversion)
        action_row.addWidget(self.btn_convert)
        self.progress_bar = QProgressBar()
        action_row.addWidget(self.progress_bar, stretch=1)
        layout.addLayout(action_row)

        # Log
        layout.addWidget(QLabel("Log:"))
        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMaximumBlockCount(2000)
        layout.addWidget(self.log_box, stretch=2)

        # Footer: credit & link sosial media (selalu terlihat)
        layout.addWidget(self._build_footer())

    # ------------------------------------------------------------------
    def _build_menu_bar(self):
        menu_bar = self.menuBar()
        help_menu = menu_bar.addMenu("&Bantuan")

        about_action = QAction("Tentang Aplikasi...", self)
        about_action.triggered.connect(self._show_about_dialog)
        help_menu.addAction(about_action)

    def _show_about_dialog(self):
        QMessageBox.about(
            self,
            "Tentang Universal File Converter",
            "<h3>Universal File Converter</h3>"
            "<p>Versi 1.0.0</p>"
            "<p>Dibuat oleh <b>Yudiansyah</b></p>"
            "<p>"
            '<a href="https://github.com/yudiiansyaah">GitHub</a> &nbsp;|&nbsp; '
            '<a href="https://instagram.com/yudiiansyaah">Instagram</a> &nbsp;|&nbsp; '
            '<a href="https://x.com/yudiiansyaah">X</a> &nbsp;|&nbsp; '
            '<a href="https://t.me/yudiiansyaah">Telegram</a>'
            "</p>"
        )

    # ------------------------------------------------------------------
    def _build_footer(self) -> QWidget:
        SOCIALS = {
            "GitHub": "https://github.com/yudiiansyaah",
            "Instagram": "https://instagram.com/yudiiansyaah",
            "X": "https://x.com/yudiiansyaah",
            "Telegram": "https://t.me/yudiiansyaah",
        }
        links_html = "  •  ".join(
            f'<a href="{url}" style="color:#6fb3ff; text-decoration:none;">{name}</a>'
            for name, url in SOCIALS.items()
        )

        footer = QFrame()
        footer.setStyleSheet("background:#1c1c1c; border-top: 1px solid #333;")
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(12, 6, 12, 6)

        label = QLabel(f"Dibuat oleh <b>Yudiansyah</b>  —  {links_html}")
        label.setOpenExternalLinks(True)
        label.setTextFormat(Qt.RichText)
        label.setStyleSheet("color:#aaaaaa; font-size:11px;")
        footer_layout.addWidget(label)
        footer_layout.addStretch()

        version_label = QLabel("v1.0.0")
        version_label.setStyleSheet("color:#666666; font-size:11px;")
        footer_layout.addWidget(version_label)

        return footer

    # ------------------------------------------------------------------
    def _check_dependencies_banner(self):
        status = dependencies.check_all()
        warnings = []

        # Dokumen/Presentasi hanya butuh SALAH SATU: LibreOffice atau MS Word
        if not status["soffice"] and not status["word"]:
            warnings.append(
                "- Dokumen & Presentasi (docx/doc/odt/pptx -> pdf/dll): install "
                "LibreOffice ATAU pastikan Microsoft Word terpasang."
            )
        elif not status["soffice"] and status["word"]:
            warnings.append(
                "- Dokumen & Presentasi: LibreOffice tidak terdeteksi, akan otomatis "
                "memakai Microsoft Word yang terpasang."
            )

        if not status["ffmpeg"]:
            warnings.append("- Audio & Video: install FFmpeg (ffmpeg.org).")

        if warnings:
            msg = "Info dependensi eksternal:\n" + "\n".join(warnings)
            self.dep_banner.setText(msg)
            self.dep_banner.show()

    # ------------------------------------------------------------------
    def add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Pilih file untuk dikonversi")
        for p in paths:
            self._add_row(p)

    def _add_row(self, path: str):
        ext = Path(path).suffix.lower().lstrip(".")
        if ext == "gz" and Path(path).stem.endswith(".tar"):
            ext = "tar.gz"
        targets = format_map.find_targets(ext)

        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, COL_FILE, QTableWidgetItem(os.path.basename(path)))
        self.table.item(row, COL_FILE).setToolTip(path)
        self.table.setItem(row, COL_SRC, QTableWidgetItem(f".{ext}"))

        combo = QComboBox()
        if targets:
            combo.addItems([f".{t}" for t in targets])
        else:
            combo.addItem("(tidak didukung)")
            combo.setEnabled(False)
        self.table.setCellWidget(row, COL_TARGET, combo)
        self.table.setItem(row, COL_STATUS, QTableWidgetItem("Menunggu"))

        self.file_rows.append({"path": path, "ext": ext})

    def remove_selected(self):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.table.removeRow(r)
            if r < len(self.file_rows):
                del self.file_rows[r]

    def clear_all(self):
        self.table.setRowCount(0)
        self.file_rows.clear()
        self.log_box.clear()

    def choose_output_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Pilih folder output")
        if d:
            self.output_edit.setText(d)

    # ------------------------------------------------------------------
    def start_conversion(self):
        if self.table.rowCount() == 0:
            QMessageBox.information(self, APP_TITLE, "Tambahkan file terlebih dahulu.")
            return

        jobs = []
        for row in range(self.table.rowCount()):
            combo: QComboBox = self.table.cellWidget(row, COL_TARGET)
            if combo is None or not combo.isEnabled():
                self.table.setItem(row, COL_STATUS, QTableWidgetItem("Dilewati (format tidak didukung)"))
                continue
            target_ext = combo.currentText().lstrip(".")
            input_path = self.file_rows[row]["path"]
            jobs.append((row, input_path, target_ext))
            self.table.setItem(row, COL_STATUS, QTableWidgetItem("Antre..."))

        if not jobs:
            QMessageBox.warning(self, APP_TITLE, "Tidak ada file yang bisa dikonversi.")
            return

        self.btn_convert.setEnabled(False)
        self.progress_bar.setMaximum(len(jobs))
        self.progress_bar.setValue(0)
        self.log_box.appendPlainText(f"--- Mulai konversi {len(jobs)} file ---")

        self.worker = ConversionWorker(jobs, self.output_edit.text().strip())
        self.worker.progress.connect(self._on_row_done)
        self.worker.log_line.connect(self.log_box.appendPlainText)
        self.worker.finished_all.connect(self._on_all_done)
        self.worker.start()

    def _on_row_done(self, row: int, message: str, success: bool):
        item = QTableWidgetItem(message)
        if not success:
            item.setForeground(Qt.red)
        else:
            item.setForeground(Qt.darkGreen)
        self.table.setItem(row, COL_STATUS, item)
        self.progress_bar.setValue(self.progress_bar.value() + 1)

    def _on_all_done(self):
        self.btn_convert.setEnabled(True)
        self.log_box.appendPlainText("--- Semua proses selesai ---")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
