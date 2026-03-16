import os
import shutil
from datetime import datetime
from uuid import uuid4

from PyQt5.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QGroupBox,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QMessageBox,
)

import shared.global_data as global_data


class FourPhotoModuleWindow(QMainWindow):
    MODULE_NAME = "photo_module"
    SLOT_KEYS = [
        "photo_slot_1",
        "photo_slot_2",
        "photo_slot_3",
        "photo_slot_4",
    ]

    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("4 Fotoğraf Modülü")
        self.resize(700, 420)

        self.slot_paths = {slot_key: "" for slot_key in self.SLOT_KEYS}
        self.slot_labels = {}

        self._build_ui()
        self._load_existing_state()

    def _build_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        info = QLabel(
            "Her alan bağımsız çalışır. Seçilen görseller tempfiles altında bu modüle özel klasöre kopyalanır."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        for slot_key in self.SLOT_KEYS:
            slot_group = QGroupBox(slot_key)
            slot_layout = QVBoxLayout()
            slot_group.setLayout(slot_layout)

            btn_row = QHBoxLayout()
            btn_select = QPushButton("Fotoğraf Seç")
            btn_select.clicked.connect(lambda _, key=slot_key: self.select_photo_for_slot(key))
            btn_row.addWidget(btn_select)

            btn_clear = QPushButton("Temizle")
            btn_clear.clicked.connect(lambda _, key=slot_key: self.clear_slot(key))
            btn_row.addWidget(btn_clear)

            btn_row.addStretch()
            slot_layout.addLayout(btn_row)

            lbl_file = QLabel("Seçilmedi")
            lbl_file.setWordWrap(True)
            lbl_file.setStyleSheet("color: #444;")
            slot_layout.addWidget(lbl_file)
            self.slot_labels[slot_key] = lbl_file

            layout.addWidget(slot_group)

        action_row = QHBoxLayout()
        action_row.addStretch()

        btn_back = QPushButton("Ana Sayfaya Dön")
        btn_back.clicked.connect(self.return_to_main)
        action_row.addWidget(btn_back)

        layout.addLayout(action_row)

    def _load_existing_state(self):
        existing = global_data.config.get("PHOTO_MODULE_SLOTS", {})
        if not isinstance(existing, dict):
            existing = {}

        for slot_key in self.SLOT_KEYS:
            slot_path = existing.get(slot_key, "")
            if slot_path and os.path.exists(slot_path):
                self.slot_paths[slot_key] = slot_path
                self.slot_labels[slot_key].setText(os.path.basename(slot_path))
            elif slot_path:
                # Global durumda path var ama dosya yoksa durum temizlenir
                self.slot_paths[slot_key] = ""
                self.slot_labels[slot_key].setText("Seçilmedi")

        self._sync_global_state()

    def _resolve_session_folder_name(self):
        test_no = str(global_data.config.get("TEST_NO") or "").strip()
        if test_no:
            return test_no.replace("/", "_").replace("\\", "_")

        session_id = global_data.config.get("PHOTO_MODULE_SESSION_ID")
        if not session_id:
            session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
            global_data.config["PHOTO_MODULE_SESSION_ID"] = session_id
        return session_id

    def _module_temp_dir(self):
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        session_folder = self._resolve_session_folder_name()
        module_dir = os.path.join(root_dir, "tempfiles", session_folder, self.MODULE_NAME)
        os.makedirs(module_dir, exist_ok=True)
        return module_dir

    def select_photo_for_slot(self, slot_key):
        selected_path, _ = QFileDialog.getOpenFileName(
            self,
            f"{slot_key} için fotoğraf seç",
            "",
            "Image Files (*.jpg *.jpeg *.png)",
        )
        if not selected_path:
            return

        if not self._is_supported_image(selected_path):
            QMessageBox.warning(self, "Uyarı", "Sadece .jpg, .jpeg, .png dosyaları desteklenir.")
            return

        try:
            saved_path = self._copy_to_tempfiles(selected_path)
            self.slot_paths[slot_key] = saved_path
            self.slot_labels[slot_key].setText(os.path.basename(saved_path))
            self._sync_global_state()
        except Exception as exc:
            QMessageBox.critical(self, "Hata", f"Fotoğraf kaydedilirken hata oluştu:\n{exc}")

    def clear_slot(self, slot_key):
        self.slot_paths[slot_key] = ""
        self.slot_labels[slot_key].setText("Seçilmedi")
        self._sync_global_state()

    def _copy_to_tempfiles(self, source_path):
        module_dir = self._module_temp_dir()
        base_name = os.path.basename(source_path)
        name, ext = os.path.splitext(base_name)
        candidate_name = base_name
        destination = os.path.join(module_dir, candidate_name)

        while os.path.exists(destination):
            unique_suffix = uuid4().hex[:8]
            candidate_name = f"{name}_{unique_suffix}{ext}"
            destination = os.path.join(module_dir, candidate_name)

        shutil.copy2(source_path, destination)
        return destination

    @staticmethod
    def _is_supported_image(path):
        return os.path.splitext(path)[1].lower() in {".jpg", ".jpeg", ".png"}

    def _sync_global_state(self):
        global_data.config["PHOTO_MODULE_SLOTS"] = dict(self.slot_paths)

    def return_to_main(self):
        self.close()
        if self.main_window:
            self.main_window.show()

    def closeEvent(self, event):
        self._sync_global_state()
        if self.main_window:
            self.main_window.show()
        super().closeEvent(event)
