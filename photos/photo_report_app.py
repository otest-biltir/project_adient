import math
import os
from datetime import datetime

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QVBoxLayout,
    QWidget,
    QFileDialog,
)

from docx import Document
from docx.enum.text import WD_BREAK
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm


class PhotoReportApp(QMainWindow):
    REPORT_TYPES = [
        ("PRE", "Pre fotoğraf seç"),
        ("POST", "Post fotoğraf seç"),
        ("TEARDOWN", "Teardown fotoğraf seç"),
        ("HANDLE_SIDE_COVER", "Handle Side Cover fotoğraf seç"),
    ]

    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window

        self.setWindowTitle("Photo Report Modülü")
        self.resize(800, 700)

        self.selected_photos = {report_type: [] for report_type, _ in self.REPORT_TYPES}
        self.list_widgets = {}

        root_widget = QWidget()
        self.setCentralWidget(root_widget)
        root_layout = QVBoxLayout(root_widget)

        title = QLabel("Photo Report Oluşturucu")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; margin-bottom: 8px;")
        root_layout.addWidget(title)

        desc = QLabel("Her kategori için birden fazla fotoğraf seçin. 6 fotoğraf başına 1 sayfa oluşturulur.")
        desc.setAlignment(Qt.AlignCenter)
        desc.setStyleSheet("color: #444; margin-bottom: 8px;")
        root_layout.addWidget(desc)

        for report_type, button_label in self.REPORT_TYPES:
            root_layout.addWidget(self._build_category_group(report_type, button_label))

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Hazır")
        root_layout.addWidget(self.progress_bar)

        action_row = QHBoxLayout()

        self.btn_generate = QPushButton("Raporları Oluştur")
        self.btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.btn_generate.clicked.connect(self.generate_reports)
        action_row.addWidget(self.btn_generate)

        self.btn_back = QPushButton("Ana Sayfaya Dön")
        self.btn_back.clicked.connect(self.go_back)
        action_row.addWidget(self.btn_back)

        root_layout.addLayout(action_row)

    def _build_category_group(self, report_type, button_label):
        wrapper = QWidget()
        layout = QVBoxLayout(wrapper)

        header = QLabel(report_type)
        header.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(header)

        button_row = QHBoxLayout()

        btn_select = QPushButton(button_label)
        btn_select.clicked.connect(lambda _, t=report_type: self.select_photos(t))
        button_row.addWidget(btn_select)

        btn_clear = QPushButton("Listeyi temizle")
        btn_clear.clicked.connect(lambda _, t=report_type: self.clear_photos(t))
        button_row.addWidget(btn_clear)

        layout.addLayout(button_row)

        list_widget = QListWidget()
        self.list_widgets[report_type] = list_widget
        layout.addWidget(list_widget)

        return wrapper

    def select_photos(self, report_type):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            f"{report_type} için fotoğrafları seç",
            "",
            "Images (*.jpg *.jpeg *.png)",
        )

        if not files:
            return

        self.selected_photos[report_type] = files
        self.refresh_list(report_type)

    def clear_photos(self, report_type):
        self.selected_photos[report_type] = []
        self.refresh_list(report_type)

    def refresh_list(self, report_type):
        widget = self.list_widgets[report_type]
        widget.clear()
        for path in self.selected_photos[report_type]:
            widget.addItem(os.path.basename(path))

    def _root_dir(self):
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    def _template_path(self, report_type):
        return os.path.join(self._root_dir(), "photos", "templates", f"{report_type}.docx")

    def _output_dir(self):
        directory = os.path.join(self._root_dir(), "tempfiles", "photo_reports")
        os.makedirs(directory, exist_ok=True)
        return directory

    def _unique_output_path(self, report_type):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"{report_type.lower()}_photo_report_{timestamp}.docx"
        output_dir = self._output_dir()
        candidate = os.path.join(output_dir, base_name)

        idx = 1
        while os.path.exists(candidate):
            candidate = os.path.join(output_dir, f"{report_type.lower()}_photo_report_{timestamp}_{idx}.docx")
            idx += 1

        return candidate

    def _iter_image_cells(self, table):
        return [cell for row in table.rows for cell in row.cells]

    def _add_page_table_with_images(self, doc, image_paths):
        table = doc.add_table(rows=2, cols=3)
        table.style = "Table Grid"
        cells = self._iter_image_cells(table)

        for idx, image_path in enumerate(image_paths):
            if idx >= len(cells):
                break

            cell = cells[idx]
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.add_run()
            run.add_picture(image_path, width=Mm(55))

    def _populate_document(self, template_path, image_paths, report_type):
        doc = Document(template_path)
        chunks = [image_paths[i : i + 6] for i in range(0, len(image_paths), 6)]

        for page_index, chunk in enumerate(chunks):
            if page_index > 0:
                doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            self._add_page_table_with_images(doc, chunk)

        output_path = self._unique_output_path(report_type)
        doc.save(output_path)
        return output_path

    def generate_reports(self):
        selected_types = [t for t, _ in self.REPORT_TYPES if self.selected_photos[t]]
        if not selected_types:
            QMessageBox.warning(self, "Uyarı", "En az bir kategori için fotoğraf seçmelisiniz.")
            return

        total_steps = sum(math.ceil(len(self.selected_photos[t]) / 6) for t in selected_types)
        completed_steps = 0
        self.progress_bar.setMaximum(max(total_steps, 1))
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Raporlar oluşturuluyor... %p")

        generated_files = []

        try:
            for report_type in selected_types:
                template_path = self._template_path(report_type)
                if not os.path.exists(template_path):
                    raise FileNotFoundError(f"Template bulunamadı: {template_path}")

                output_path = self._populate_document(template_path, self.selected_photos[report_type], report_type)
                generated_files.append(output_path)

                completed_steps += math.ceil(len(self.selected_photos[report_type]) / 6)
                self.progress_bar.setValue(completed_steps)

            self.progress_bar.setFormat("Tamamlandı")
            QMessageBox.information(
                self,
                "Başarılı",
                "Raporlar oluşturuldu:\n\n" + "\n".join(generated_files),
            )
        except Exception as exc:
            QMessageBox.critical(self, "Hata", f"Rapor oluşturma hatası:\n{exc}")
            self.progress_bar.setFormat("Hata")

    def closeEvent(self, event):
        if self.main_window:
            self.main_window.show()
        super().closeEvent(event)

    def go_back(self):
        self.close()
        if self.main_window:
            self.main_window.show()
