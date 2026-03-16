import os
import tempfile
from copy import deepcopy
from datetime import datetime

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication,
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
    QGroupBox,
)
from docx import Document
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm


class PhotoReportApp(QMainWindow):
    CATEGORIES = {
        "PRE": "PRE",
        "POST": "POST",
        "TEARDOWN": "TEARDOWN",
        "HANDLE_SIDE_COVER": "HANDLE_SIDE_COVER",
    }

    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("Photo Report Modülü")
        self.resize(900, 700)

        self.selected_photos = {key: [] for key in self.CATEGORIES}
        self.list_widgets = {}

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        lbl_title = QLabel("Photo Report Oluşturucu")
        lbl_title.setAlignment(Qt.AlignCenter)
        lbl_title.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(lbl_title)

        category_meta = [
            ("PRE", "Pre fotoğraf seç"),
            ("POST", "Post fotoğraf seç"),
            ("TEARDOWN", "Teardown fotoğraf seç"),
            ("HANDLE_SIDE_COVER", "Handle Side Cover fotoğraf seç"),
        ]

        for category, select_text in category_meta:
            group = QGroupBox(category)
            group_layout = QVBoxLayout(group)

            button_row = QHBoxLayout()

            btn_select = QPushButton(select_text)
            btn_select.clicked.connect(lambda _, c=category: self.select_photos(c))
            button_row.addWidget(btn_select)

            btn_clear = QPushButton("Listeyi temizle")
            btn_clear.clicked.connect(lambda _, c=category: self.clear_photos(c))
            button_row.addWidget(btn_clear)

            group_layout.addLayout(button_row)

            list_widget = QListWidget()
            group_layout.addWidget(list_widget)
            self.list_widgets[category] = list_widget

            layout.addWidget(group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        btn_generate = QPushButton("Seçili Kategoriler İçin Rapor Oluştur")
        btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        btn_generate.clicked.connect(self.generate_reports)
        layout.addWidget(btn_generate)

        btn_back = QPushButton("Ana Menüye Dön")
        btn_back.clicked.connect(self.go_back)
        layout.addWidget(btn_back)

    def select_photos(self, category):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            f"{category} için fotoğraf seç",
            "",
            "Image Files (*.jpg *.jpeg *.png)",
        )
        if not files:
            return

        self.selected_photos[category] = files
        self.refresh_category_list(category)

    def clear_photos(self, category):
        self.selected_photos[category] = []
        self.refresh_category_list(category)

    def refresh_category_list(self, category):
        widget = self.list_widgets[category]
        widget.clear()
        for path in self.selected_photos[category]:
            widget.addItem(os.path.basename(path))

    def _chunked(self, items, size=6):
        for i in range(0, len(items), size):
            yield items[i : i + size]

    def _template_path(self, category):
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(root_dir, "photos", "templates", f"{category}.docx")

    def _temp_output_dir(self):
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        out_dir = os.path.join(root_dir, "tempfiles", "photo_reports")
        os.makedirs(out_dir, exist_ok=True)
        return out_dir

    def _safe_output_path(self, out_dir, category):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"{category.lower()}_photo_report_{timestamp}"
        candidate = os.path.join(out_dir, f"{base_name}.docx")
        idx = 1
        while os.path.exists(candidate):
            candidate = os.path.join(out_dir, f"{base_name}_{idx}.docx")
            idx += 1
        return candidate

    def _build_context(self, doc, photos):
        context = {}
        for i in range(6):
            key_index = i + 1
            if i < len(photos):
                image_obj = InlineImage(doc, photos[i], width=Mm(58))
                filename = os.path.basename(photos[i])
            else:
                image_obj = ""
                filename = ""

            # Farklı template isimlendirme stillerini desteklemek için birden çok alias bırakıyoruz.
            context[f"PHOTO_{key_index}"] = image_obj
            context[f"PHOTO{key_index}"] = image_obj
            context[f"IMG_{key_index}"] = image_obj
            context[f"IMG{key_index}"] = image_obj
            context[f"IMAGE_{key_index}"] = image_obj
            context[f"IMAGE{key_index}"] = image_obj
            context[f"PHOTO_NAME_{key_index}"] = filename
            context[f"PHOTO_NAME{key_index}"] = filename
        return context

    def _merge_documents(self, rendered_paths, output_path):
        if not rendered_paths:
            raise ValueError("Birleştirilecek doküman yok")

        base_doc = Document(rendered_paths[0])
        for next_path in rendered_paths[1:]:
            next_doc = Document(next_path)
            for element in next_doc.element.body:
                if element.tag.endswith("sectPr"):
                    continue
                base_doc.element.body.append(deepcopy(element))

        base_doc.save(output_path)

    def _render_category_report(self, category, photos):
        template_path = self._template_path(category)
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template bulunamadı: {template_path}")

        out_dir = self._temp_output_dir()
        output_path = self._safe_output_path(out_dir, category)

        chunk_list = list(self._chunked(photos, size=6))
        rendered_paths = []

        with tempfile.TemporaryDirectory() as tmp_dir:
            for idx, chunk in enumerate(chunk_list, start=1):
                doc = DocxTemplate(template_path)
                doc.render(self._build_context(doc, chunk))
                rendered_path = os.path.join(tmp_dir, f"{category.lower()}_{idx}.docx")
                doc.save(rendered_path)
                rendered_paths.append(rendered_path)

            self._merge_documents(rendered_paths, output_path)

        return output_path

    def generate_reports(self):
        selected_categories = {
            category: photos
            for category, photos in self.selected_photos.items()
            if photos
        }

        if not selected_categories:
            QMessageBox.warning(self, "Uyarı", "Lütfen en az bir kategori için fotoğraf seçin.")
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(selected_categories))
        self.progress_bar.setValue(0)

        generated_files = []
        try:
            for step, (category, photos) in enumerate(selected_categories.items(), start=1):
                out_path = self._render_category_report(category, photos)
                generated_files.append(out_path)
                self.progress_bar.setValue(step)
                QApplication.processEvents()

            QMessageBox.information(
                self,
                "Başarılı",
                "Raporlar oluşturuldu:\n\n" + "\n".join(generated_files),
            )
        except Exception as exc:
            QMessageBox.critical(self, "Hata", f"Rapor oluşturma hatası:\n{exc}")
        finally:
            self.progress_bar.setVisible(False)

    def go_back(self):
        if self.main_window:
            self.close()
            self.main_window.show()

