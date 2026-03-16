import os
from copy import deepcopy
from io import BytesIO
from datetime import datetime

from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QGroupBox,
    QProgressBar,
)
from PyQt5.QtCore import Qt

from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Mm

import shared.global_data as global_data


class PhotoReportApp(QMainWindow):
    IMAGE_FILTER = "Fotoğraflar (*.jpg *.jpeg *.png)"
    CATEGORIES = {
        "PRE": {
            "title": "Pre",
            "template": "PRE.docx",
            "color": "#03A9F4",
        },
        "POST": {
            "title": "Post",
            "template": "POST.docx",
            "color": "#4CAF50",
        },
        "TEARDOWN": {
            "title": "Teardown",
            "template": "TEARDOWN.docx",
            "color": "#FF9800",
        },
        "HANDLE_SIDE_COVER": {
            "title": "Handle Side Cover",
            "template": "HANDLE_SIDE_COVER.docx",
            "color": "#9C27B0",
        },
    }

    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("Photo Report")
        self.resize(900, 760)

        self.selected_files = {key: [] for key in self.CATEGORIES}
        self.list_widgets = {}

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        title = QLabel("Photo Report Modülü")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; margin-bottom: 10px;")
        main_layout.addWidget(title)

        subtitle = QLabel("Her kategori için çoklu fotoğraf seçebilir ve ayrı Word raporu oluşturabilirsiniz.")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #616161; margin-bottom: 8px;")
        main_layout.addWidget(subtitle)

        for category, cfg in self.CATEGORIES.items():
            main_layout.addWidget(self._build_category_section(category, cfg))

        action_row = QHBoxLayout()

        self.btn_generate = QPushButton("Seçilen Kategoriler İçin Raporları Oluştur")
        self.btn_generate.setStyleSheet(
            "font-size: 15px; padding: 12px; background-color: #1976D2; color: white; font-weight: bold;"
        )
        self.btn_generate.clicked.connect(self.generate_reports)

        self.btn_back = QPushButton("Ana Menüye Dön")
        self.btn_back.setStyleSheet("font-size: 15px; padding: 12px; background-color: #9E9E9E; color: white;")
        self.btn_back.clicked.connect(self.close_and_return)

        action_row.addWidget(self.btn_generate)
        action_row.addWidget(self.btn_back)

        main_layout.addLayout(action_row)

        self.progress = QProgressBar()
        self.progress.setMinimum(0)
        self.progress.setValue(0)
        self.progress.setFormat("Hazır")
        main_layout.addWidget(self.progress)

        self.lbl_output = QLabel("Çıktı klasörü: tempfiles/photo_reports")
        self.lbl_output.setStyleSheet("color: #616161; margin-top: 4px;")
        main_layout.addWidget(self.lbl_output)

    def _build_category_section(self, category, cfg):
        group = QGroupBox(cfg["title"])
        group.setStyleSheet(f"QGroupBox {{ font-weight: bold; color: {cfg['color']}; }}")

        layout = QVBoxLayout(group)

        button_row = QHBoxLayout()

        btn_select = QPushButton(f"{cfg['title']} fotoğraf seç")
        btn_select.clicked.connect(lambda _, c=category: self.select_photos(c))
        button_row.addWidget(btn_select)

        btn_reset = QPushButton("Listeyi temizle")
        btn_reset.clicked.connect(lambda _, c=category: self.clear_category(c))
        button_row.addWidget(btn_reset)

        layout.addLayout(button_row)

        list_widget = QListWidget()
        list_widget.setSelectionMode(QListWidget.ExtendedSelection)
        list_widget.setMinimumHeight(80)
        layout.addWidget(list_widget)

        self.list_widgets[category] = list_widget
        self._refresh_list(category)

        return group

    def select_photos(self, category):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            f"{self.CATEGORIES[category]['title']} fotoğraflarını seç",
            "",
            self.IMAGE_FILTER,
        )

        if not files:
            return

        valid_ext = {".jpg", ".jpeg", ".png"}
        dedup = set(self.selected_files[category])

        for path in files:
            ext = os.path.splitext(path)[1].lower()
            if ext in valid_ext and path not in dedup:
                self.selected_files[category].append(path)
                dedup.add(path)

        self._refresh_list(category)

    def clear_category(self, category):
        self.selected_files[category] = []
        self._refresh_list(category)

    def _refresh_list(self, category):
        list_widget = self.list_widgets[category]
        list_widget.clear()

        files = self.selected_files[category]
        if not files:
            list_widget.addItem(QListWidgetItem("Henüz fotoğraf seçilmedi"))
            list_widget.item(0).setFlags(Qt.NoItemFlags)
            return

        for photo in files:
            list_widget.addItem(QListWidgetItem(os.path.basename(photo)))

    def _ensure_output_dir(self):
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        out_dir = os.path.join(root_dir, "tempfiles", "photo_reports")
        os.makedirs(out_dir, exist_ok=True)
        return out_dir

    def _resolve_template_path(self, category):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_name = self.CATEGORIES[category]["template"]
        return os.path.join(base_dir, "templates", template_name)

    def _safe_output_name(self, category, output_dir):
        test_no = global_data.config.get("TEST_NO") or "UNSPECIFIED"
        normalized_test = str(test_no).replace("/", "_").replace("\\", "_").strip()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"photo_report_{category.lower()}_{normalized_test}_{timestamp}"

        candidate = os.path.join(output_dir, f"{base_name}.docx")
        idx = 1
        while os.path.exists(candidate):
            candidate = os.path.join(output_dir, f"{base_name}_{idx}.docx")
            idx += 1

        return candidate

    def _chunks_of_six(self, photos):
        slots_per_page = 6
        groups = []
        for start in range(0, len(photos), slots_per_page):
            chunk = photos[start : start + slots_per_page]
            while len(chunk) < slots_per_page:
                chunk.append("")
            groups.append(chunk)

        if not groups:
            groups.append(["", "", "", "", "", ""])

        return groups

    def _slot_context(self, doc, photo_chunk, page_no):
        def _inline(path):
            return InlineImage(doc, path, width=Mm(55)) if path else ""

        return {
            "TEST_NO": global_data.config.get("TEST_NO", ""),
            "TEST_DATE": global_data.config.get("TEST_DATE", ""),
            "PROJECT": global_data.config.get("PROJECT", ""),
            "PAGE_NO": page_no,
            # Common placeholder variants used in templates.
            "slot_1": _inline(photo_chunk[0]),
            "slot_2": _inline(photo_chunk[1]),
            "slot_3": _inline(photo_chunk[2]),
            "slot_4": _inline(photo_chunk[3]),
            "slot_5": _inline(photo_chunk[4]),
            "slot_6": _inline(photo_chunk[5]),
            "SLOT_1": _inline(photo_chunk[0]),
            "SLOT_2": _inline(photo_chunk[1]),
            "SLOT_3": _inline(photo_chunk[2]),
            "SLOT_4": _inline(photo_chunk[3]),
            "SLOT_5": _inline(photo_chunk[4]),
            "SLOT_6": _inline(photo_chunk[5]),
            "PHOTO_1": _inline(photo_chunk[0]),
            "PHOTO_2": _inline(photo_chunk[1]),
            "PHOTO_3": _inline(photo_chunk[2]),
            "PHOTO_4": _inline(photo_chunk[3]),
            "PHOTO_5": _inline(photo_chunk[4]),
            "PHOTO_6": _inline(photo_chunk[5]),
        }

    def _render_template_page(self, template_path, photos_for_page, page_no):
        doc_tpl = DocxTemplate(template_path)
        context = self._slot_context(doc_tpl, photos_for_page, page_no)
        doc_tpl.render(context)

        mem = BytesIO()
        doc_tpl.save(mem)
        mem.seek(0)
        return Document(mem)

    def _append_document_body(self, target_doc, source_doc):
        target_body = target_doc.element.body
        source_body = source_doc.element.body

        target_sectpr = target_body.sectPr
        if target_sectpr is not None:
            target_body.remove(target_sectpr)

        for child in source_body:
            if child.tag.endswith("sectPr"):
                continue
            target_body.append(deepcopy(child))

        source_sectpr = source_body.sectPr
        if source_sectpr is not None:
            target_body.append(deepcopy(source_sectpr))

    def _generate_single_category_report(self, category, output_dir):
        template_path = self._resolve_template_path(category)
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template bulunamadı: {template_path}")

        chunks = self._chunks_of_six(list(self.selected_files[category]))
        merged_doc = None

        for page_no, chunk in enumerate(chunks, start=1):
            rendered_page = self._render_template_page(template_path, chunk, page_no)
            if merged_doc is None:
                merged_doc = rendered_page
            else:
                self._append_document_body(merged_doc, rendered_page)

        output_path = self._safe_output_name(category, output_dir)
        merged_doc.save(output_path)
        return output_path

    def generate_reports(self):
        selected_categories = [key for key, files in self.selected_files.items() if files]
        if not selected_categories:
            QMessageBox.warning(self, "Uyarı", "Rapor üretmek için en az bir kategoriye fotoğraf ekleyin.")
            return

        output_dir = self._ensure_output_dir()
        self.progress.setMinimum(0)
        self.progress.setMaximum(len(selected_categories))
        self.progress.setValue(0)

        created_files = []

        try:
            for step, category in enumerate(selected_categories, start=1):
                output_path = self._generate_single_category_report(category, output_dir)
                created_files.append(output_path)

                self.progress.setValue(step)
                self.progress.setFormat(f"%p - {self.CATEGORIES[category]['title']} oluşturuldu")
                QApplication.processEvents()

            result_text = "\n".join(created_files)
            QMessageBox.information(
                self,
                "Başarılı",
                f"{len(created_files)} rapor başarıyla oluşturuldu.\n\nÇıktılar:\n{result_text}",
            )
        except Exception as exc:
            QMessageBox.critical(self, "Hata", f"Rapor oluşturulurken hata oluştu:\n{exc}")

    def close_and_return(self):
        self.close()
        if self.main_window is not None:
            self.main_window.show()
