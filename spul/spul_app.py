import sys
import os
import subprocess
import shutil
from pathlib import Path

def _check_and_install_dependencies():
    required_packages = ['pandas', 'numpy', 'PyQt5', 'matplotlib', 'openpyxl', 'xlrd']
    for pkg in required_packages:
        try:
            __import__(pkg)
        except ImportError:
            print(f"Eksik kütüphane tespit edildi, yükleniyor: {pkg}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
                print(f"{pkg} başarıyla yüklendi.")
            except Exception as e:
                print(f"{pkg} yüklenirken hata oluştu: {e}")

_check_and_install_dependencies()

import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel,
                             QMessageBox, QDoubleSpinBox, QGroupBox,
                             QCheckBox, QTableWidget, QTableWidgetItem, QHeaderView,
                             QAbstractItemView, QLineEdit, QFileDialog, QInputDialog,
                             QSizePolicy)
from PyQt5.QtCore import Qt

import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import AutoMinorLocator, FormatStrFormatter, MultipleLocator


MAX_GRAPH_TIME_SEC = 0.14
DATA_INTERVAL_SEC = 0.0004
MS_PER_ROW = DATA_INTERVAL_SEC * 1000.0
ROWS_FOR_14MS = round(14.0 / MS_PER_ROW)
QNAP_TEST_ROOT = r"O:\1_BILTIR_TEST_DOSYALARI\2026\02 - DINIZ-ADIENT"
TEST_FOLDER_PREFIX = "26-"
REPORT_EVA_ACC_RELATIVE = os.path.join("REPORT FILES", "3-EVA-ACC")
TEMPLATE_EXCEL_NAME = "template.xlsx"


class SledAnalyzerApp(QMainWindow):
    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("Sled Test Analyzer (Multi-Graph)")
        self.resize(1280, 960)

        self.data_path = None
        self.export_dir = None
        self.selected_test_name = None
        self.test_locations = []
        self.df_actual = None
        self.df_target = None

        # State
        self.current_graph_idx = 0
        self.graphs = ["Spul", "Acceleration vs Velocity", "Actual vs Target Acceleration"]
        self.local_offsets = [0, 0, 0]

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # --- Left Sidebar Panel ---
        sidebar = QWidget()
        sidebar.setMinimumWidth(360)
        sidebar.setMaximumWidth(440)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(8)

        # --- Right Graph Area ---
        graph_area = QWidget()
        graph_area_layout = QVBoxLayout(graph_area)
        graph_area_layout.setContentsMargins(0, 0, 0, 0)
        graph_area_layout.setSpacing(8)

        # --- Control Panel (Left) ---
        control_group = QGroupBox("Veri Yükleme ve Ayarlar")
        control_layout = QVBoxLayout()
        control_group.setLayout(control_layout)

        # File/Test Selection
        self.btn_select_test = QPushButton("QNAP / Test Klasörü Seç (template.xlsx otomatik)")
        self.btn_select_test.setStyleSheet("background-color: #1976D2; color: white; font-weight: bold; padding: 10px;")
        self.btn_select_test.clicked.connect(self.browse_export_dir)
        control_layout.addWidget(self.btn_select_test)

        self.btn_data = QPushButton("Excel Dosyasını Elle Yükle / Değiştir")
        self.btn_data.clicked.connect(self.load_data_file)
        control_layout.addWidget(self.btn_data)

        self.lbl_data = QLabel("Seçilmedi")
        self.lbl_data.setWordWrap(True)
        control_layout.addWidget(self.lbl_data)

        lbl_format = QLabel("Format: 3. satırdan itibaren A=Time(s), B=Target Acc(g), C=Target Hız(m/s), D=Actual Acc(g), E=Actual Hız(m/s)")
        lbl_format.setWordWrap(True)
        lbl_format.setStyleSheet("color: gray; font-size: 11px;")
        control_layout.addWidget(lbl_format)

        lbl_qnap = QLabel(f"QNAP test kökü: {QNAP_TEST_ROOT}\nTest seçince veri otomatik olarak REPORT FILES/3-EVA-ACC/template.xlsx dosyasından alınır.")
        lbl_qnap.setWordWrap(True)
        lbl_qnap.setStyleSheet("color: #555; font-size: 11px;")
        control_layout.addWidget(lbl_qnap)

        control_layout.addStretch() # Push items up

        # Action Buttons
        self.btn_generate = QPushButton("Oluştur / Güncelle")
        self.btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.btn_generate.clicked.connect(self.generate_plots)
        control_layout.addWidget(self.btn_generate)

        sidebar_layout.addWidget(control_group)

        # --- Offset Table Panel (Right) ---
        offset_group = QGroupBox("Actual Offset Ayarları")
        offset_layout = QVBoxLayout()
        offset_group.setLayout(offset_layout)

        self.table_offset = QTableWidget()
        self.table_offset.setColumnCount(3)
        self.table_offset.setHorizontalHeaderLabels(["Değişken / Grafik", "Offset (ms)", "Satır Karşılığı"])
        self.table_offset.setRowCount(3)
        self.table_offset.verticalHeader().setVisible(False)
        self.table_offset.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table_offset.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table_offset.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        # Tabloyu dikey olarak sıkıştır
        self.table_offset.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table_offset.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table_offset.setMaximumHeight(150) # 3 satırın tam sığacağı ideal yükseklik

        self.table_offset.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_offset.setStyleSheet("QTableWidget { background-color: white; gridline-color: #d3d3d3; } "
                                        "QHeaderView::section { background-color: #f0f0f0; font-weight: bold; }")

        # Populate Table
        labels = ["Spul", "Acceleration vs Velocity", "Actual vs Target Acceleration"]
        used_by = ["0.0 ms", "0.0 ms", "0.0 ms"]

        self.spin_offsets = []
        self.offset_duration_items = []
        for i in range(3):
            # Column 0: Variable
            item_var = QTableWidgetItem(labels[i])
            item_var.setFlags(item_var.flags() ^ Qt.ItemIsEditable)
            self.table_offset.setItem(i, 0, item_var)

            # Column 1: Offset in ms. One UI step = one data row = 0.4 ms.
            spin = QDoubleSpinBox()
            spin.setRange(-4000.0, 4000.0)
            spin.setValue(0.0)
            spin.setSingleStep(MS_PER_ROW)
            spin.setDecimals(1)
            spin.setSuffix(" ms")
            spin.setStyleSheet("border: none; background: transparent;")
            spin.valueChanged.connect(lambda val, idx=i: self.set_local_offset(idx, val))
            self.table_offset.setCellWidget(i, 1, spin)
            self.spin_offsets.append(spin)

            # Column 2: Used By (Blue Text)
            item_used = QTableWidgetItem(used_by[i])
            item_used.setFlags(item_used.flags() ^ Qt.ItemIsEditable)
            item_used.setForeground(Qt.blue)
            self.table_offset.setItem(i, 2, item_used)
            self.offset_duration_items.append(item_used)

        offset_layout.addWidget(self.table_offset)

        # Universal Offset input at bottom of table
        univ_layout = QVBoxLayout()
        univ_layout.addWidget(QLabel("Tüm actual grafiklere aynı offseti uygula:"))
        self.spin_universal = QDoubleSpinBox()
        self.spin_universal.setRange(-4000.0, 4000.0)
        self.spin_universal.setValue(0.0)
        self.spin_universal.setSingleStep(MS_PER_ROW)
        self.spin_universal.setDecimals(1)
        self.spin_universal.setSuffix(" ms")
        self.spin_universal.valueChanged.connect(self.apply_universal_offset)
        univ_layout.addWidget(self.spin_universal)

        # 14 ms tick box: 0.0004 s örnek aralığında 14 ms = 35 satır
        self.check_14ms = QCheckBox(f"Tüm Actual Grafikler İçin 14 ms Sabit Offset ({ROWS_FOR_14MS} satır)")
        self.check_14ms.stateChanged.connect(self.apply_14ms_offset)
        univ_layout.addWidget(self.check_14ms)

        offset_layout.addLayout(univ_layout)
        sidebar_layout.addWidget(offset_group)

        # --- Graph Navigation ---
        nav_layout = QHBoxLayout()
        self.btn_prev = QPushButton("⬅")
        self.btn_prev.setStyleSheet("font-size: 24px; font-weight: bold; width: 60px; height: 40px;")
        self.btn_prev.clicked.connect(self.prev_graph)

        self.lbl_graph_name = QLabel(f"{self.graphs[self.current_graph_idx]}")
        self.lbl_graph_name.setAlignment(Qt.AlignCenter)
        self.lbl_graph_name.setStyleSheet("font-size: 16px; font-weight: bold;")

        self.btn_next = QPushButton("➡")
        self.btn_next.setStyleSheet("font-size: 24px; font-weight: bold; width: 60px; height: 40px;")
        self.btn_next.clicked.connect(self.next_graph)

        nav_layout.addWidget(self.btn_prev)
        nav_layout.addWidget(self.lbl_graph_name)
        nav_layout.addWidget(self.btn_next)

        graph_area_layout.addLayout(nav_layout)

        # --- Plot Area (Matplotlib) ---
        plot_group = QGroupBox("Grafik Ekranı")
        plot_layout = QVBoxLayout()
        plot_group.setLayout(plot_layout)

        self.figure = Figure(figsize=(9.8, 6.6), facecolor="white")
        self.canvas = FigureCanvas(self.figure)
        self.canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.canvas.updateGeometry()
        plot_layout.addWidget(self.canvas, stretch=1)

        # Tablo ayarı
        import matplotlib.gridspec as gridspec
        self.gs = gridspec.GridSpec(2, 1, height_ratios=[4.4, 1.25]) # UI'da sağ panelin içine sığacak yatay yerleşim
        self.ax = self.figure.add_subplot(self.gs[0])
        self.ax_table = self.figure.add_subplot(self.gs[1])
        self.ax_table.axis('off')

        self.ax2 = None # Sağ eksen için

        graph_area_layout.addWidget(plot_group, stretch=1)

        # --- Export Area ---
        export_group = QGroupBox("Export")
        export_layout = QVBoxLayout(export_group)
        export_layout.addWidget(QLabel("Kayıt Dizini:"))
        self.txt_export = QLineEdit(QNAP_TEST_ROOT if os.path.isdir(QNAP_TEST_ROOT) else "")
        export_layout.addWidget(self.txt_export)

        self.btn_export = QPushButton("TEST EXPORT - Tüm Grafikleri Kaydet (.png)")
        self.btn_export.setStyleSheet("background-color: #F57C00; color: white; font-weight: bold; padding: 12px;")
        self.btn_export.clicked.connect(self.export_plots)
        export_layout.addWidget(self.btn_export)
        sidebar_layout.addWidget(export_group)
        sidebar_layout.addStretch()

        # --- Author Info ---
        lbl_author = QLabel("Created by Efe Nakcı")
        lbl_author.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lbl_author.setStyleSheet("color: gray; font-style: italic; font-size: 11px; padding-top: 5px;")
        graph_area_layout.addWidget(lbl_author)

        main_layout.addWidget(sidebar)
        main_layout.addWidget(graph_area, stretch=1)

    def load_data_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Excel Veri Dosyası Seç",
            self.data_path or self.txt_export.text() or "",
            "Excel Dosyaları (*.xlsx *.xls);;Tüm Dosyalar (*)",
        )
        if not path:
            return

        self.data_path = path
        self.lbl_data.setText(path)

        parent_dir = os.path.dirname(path)
        self.export_dir = parent_dir
        self.selected_test_name = os.path.basename(os.path.dirname(parent_dir)) or os.path.splitext(os.path.basename(path))[0]
        self.txt_export.setText(parent_dir)

    def _bundled_template_path(self):
        return Path(__file__).resolve().parent.parent / TEMPLATE_EXCEL_NAME

    def _template_fill_instructions(self):
        return (
            "template.xlsx doldurma formatı:\n\n"
            "• İlk 2 satır başlık/not alanı olarak kalabilir.\n"
            "• 3. satırdan itibaren veri yapıştırın.\n"
            "• A: Time (s)\n"
            "• B: Target Acceleration (g)\n"
            "• C: Target Velocity (m/s)\n"
            "• D: Actual Acceleration (g)\n"
            "• E: Actual Velocity (m/s)\n\n"
            "Target sütunlarına hedef pulse/hız, Actual sütunlarına ölçülen test verisi girilmelidir."
        )

    def _ensure_template_exists(self, template_path):
        if os.path.isfile(template_path):
            return True

        answer = QMessageBox.question(
            self,
            "Template bulunamadı",
            f"Bu test klasöründe {TEMPLATE_EXCEL_NAME} yok:\n{template_path}\n\n"
            "Buraya otomatik boş template.xlsx kurulsun mu?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if answer != QMessageBox.Yes:
            return False

        source_template = self._bundled_template_path()
        if not source_template.is_file():
            QMessageBox.critical(self, "Template kaynağı bulunamadı", f"Repo içindeki template bulunamadı:\n{source_template}")
            return False

        os.makedirs(os.path.dirname(template_path), exist_ok=True)
        shutil.copy2(source_template, template_path)
        QMessageBox.information(
            self,
            "Template kuruldu",
            f"Boş template.xlsx hedef klasöre kopyalandı:\n{template_path}\n\n{self._template_fill_instructions()}",
        )
        return True

    def _resolve_template_from_directory(self, directory):
        directory = os.path.abspath(directory)
        direct_template = os.path.join(directory, TEMPLATE_EXCEL_NAME)
        if os.path.isfile(direct_template) or os.path.basename(directory) == os.path.basename(REPORT_EVA_ACC_RELATIVE):
            return directory, direct_template

        report_dir = os.path.join(directory, REPORT_EVA_ACC_RELATIVE)
        report_template = os.path.join(report_dir, TEMPLATE_EXCEL_NAME)
        if os.path.isdir(report_dir) or os.path.basename(directory).startswith(TEST_FOLDER_PREFIX):
            return report_dir, report_template

        return directory, direct_template

    def apply_selected_directory(self, directory):
        export_dir, template_path = self._resolve_template_from_directory(directory)
        self.txt_export.setText(export_dir)
        self.export_dir = export_dir
        self.selected_test_name = os.path.basename(os.path.normpath(directory)) or "Seçilen klasör"
        if self._ensure_template_exists(template_path):
            self.data_path = template_path
            self.lbl_data.setText(f"{self.selected_test_name} / {TEMPLATE_EXCEL_NAME}")
        else:
            self.data_path = None
            self.lbl_data.setText("template.xlsx bulunamadı")

    def apply_universal_offset(self, val):
        row_offset = self.ms_to_rows(val)
        normalized_ms = row_offset * MS_PER_ROW
        if abs(normalized_ms - val) > 1e-9:
            blocked = self.spin_universal.blockSignals(True)
            self.spin_universal.setValue(normalized_ms)
            self.spin_universal.blockSignals(blocked)
        for spin in self.spin_offsets:
            spin.setValue(normalized_ms)

    def apply_14ms_offset(self, state):
        if state == Qt.Checked:
            # 14 ms, 0.0004 s örnek aralığında 35 satıra karşılık gelir.
            for spin in self.spin_offsets:
                spin.setValue(14.0)
                spin.setEnabled(False)
            self.spin_universal.setValue(14.0)
            self.spin_universal.setEnabled(False)
        else:
            # Re-enable manual edit
            for spin in self.spin_offsets:
                spin.setEnabled(True)
            self.spin_universal.setEnabled(True)

    def browse_export_dir(self):
        tests = self.find_qnap_tests()
        if tests:
            labels = [test["label"] for test in tests]
            selected, ok = QInputDialog.getItem(
                self,
                "QNAP Test Numarası Seç",
                "Kaydedilecek test numarasını seçin:",
                labels,
                0,
                False,
            )
            if ok and selected:
                test_info = tests[labels.index(selected)]
                self.apply_selected_test(test_info)
            return

        directory = QFileDialog.getExistingDirectory(
            self,
            "Test veya REPORT FILES/3-EVA-ACC Klasörü Seç",
            self.txt_export.text() or QNAP_TEST_ROOT,
        )
        if directory:
            self.apply_selected_directory(directory)

    def find_qnap_tests(self):
        root = QNAP_TEST_ROOT
        if not os.path.isdir(root):
            QMessageBox.warning(
                self,
                "QNAP yolu bulunamadı",
                f"QNAP kısayolu/yolu erişilebilir değil:\n{root}\n\nElle kayıt klasörü seçebilirsiniz.",
            )
            return []

        tests = []
        for project_name in sorted(os.listdir(root)):
            project_path = os.path.join(root, project_name)
            if not os.path.isdir(project_path):
                continue
            for test_name in sorted(os.listdir(project_path)):
                test_path = os.path.join(project_path, test_name)
                if not os.path.isdir(test_path) or not test_name.startswith(TEST_FOLDER_PREFIX):
                    continue
                eva_acc_dir = os.path.join(test_path, REPORT_EVA_ACC_RELATIVE)
                if not os.path.isdir(eva_acc_dir):
                    continue
                template_path = os.path.join(eva_acc_dir, TEMPLATE_EXCEL_NAME)
                tests.append({
                    "label": f"{test_name} — {project_name}",
                    "test_name": test_name,
                    "project_name": project_name,
                    "test_path": test_path,
                    "export_dir": eva_acc_dir,
                    "template_path": template_path,
                })
        if not tests:
            QMessageBox.warning(self, "Test bulunamadı", f"{root} altında {TEST_FOLDER_PREFIX}xxx formatında test klasörü bulunamadı.")
        return tests

    def apply_selected_test(self, test_info):
        export_dir = test_info["export_dir"]
        template_path = test_info["template_path"]
        if not os.path.isdir(export_dir):
            os.makedirs(export_dir, exist_ok=True)
        self.txt_export.setText(export_dir)
        self.export_dir = export_dir
        self.selected_test_name = test_info["test_name"]
        if self._ensure_template_exists(template_path):
            self.data_path = template_path
            self.lbl_data.setText(f"{test_info['test_name']} / {TEMPLATE_EXCEL_NAME}")
        else:
            self.data_path = None
            self.lbl_data.setText("template.xlsx bulunamadı")

    def set_local_offset(self, idx, val):
        row_offset = self.ms_to_rows(val)
        self.local_offsets[idx] = row_offset
        normalized_ms = row_offset * MS_PER_ROW
        if abs(normalized_ms - val) > 1e-9:
            spin = self.spin_offsets[idx]
            blocked = spin.blockSignals(True)
            spin.setValue(normalized_ms)
            spin.blockSignals(blocked)
        if idx < len(self.offset_duration_items):
            self.offset_duration_items[idx].setText(self.format_offset_duration(row_offset))
        if self.current_graph_idx == idx and self.df_actual is not None:
            self.draw_current_graph()

    def ms_to_rows(self, offset_ms):
        return int(round(offset_ms / MS_PER_ROW))

    def format_offset_duration(self, row_offset):
        return f"{row_offset} satır × {MS_PER_ROW:.1f} ms"

    def prev_graph(self):
        self.current_graph_idx = (self.current_graph_idx - 1) % len(self.graphs)
        self.update_graph_view()

    def next_graph(self):
        self.current_graph_idx = (self.current_graph_idx + 1) % len(self.graphs)
        self.update_graph_view()

    def update_graph_view(self):
        self.lbl_graph_name.setText(f"{self.graphs[self.current_graph_idx]}")
        if self.df_actual is not None:
            self.draw_current_graph()

    def process_data(self, df):
        df_proc = df.copy()

        # Trim space in column names
        df_proc.columns = df_proc.columns.str.strip()

        # Convert necessary columns to num
        for col in df_proc.columns:
            if col in ['Time', 'Velocity', 'Target Velocity', 'Acceleration', 'Target Acceleration']:
                df_proc[col] = pd.to_numeric(df_proc[col], errors='coerce')

        return df_proc

    def get_current_row_offset(self):
        return int(self.local_offsets[self.current_graph_idx])

    def generate_plots(self):
        if not self.data_path:
            QMessageBox.warning(self, "Uyarı", "Lütfen önce test numarasını seçin. Excel, seçilen testin 3-EVA-ACC/template.xlsx dosyasından otomatik alınır.")
            return False

        try:
            # Tek Excel formatı:
            # 1-2. satırlar atlanır; 3. satırdan itibaren A:E sütunları veri olarak okunur.
            # A=Time(s), B=Target Acceleration(g), C=Target Velocity(m/s),
            # D=Actual Acceleration(g), E=Actual Velocity(m/s).
            df_raw = pd.read_excel(
                self.data_path,
                skiprows=2,
                header=None,
                usecols=[0, 1, 2, 3, 4],
            )
            df_raw.columns = ['Time', 'Target Acceleration', 'Target Velocity', 'Acceleration', 'Velocity']
            df_raw = df_raw.dropna(how='all')
            self.df_actual = self.process_data(df_raw)

            if self.df_actual.empty:
                QMessageBox.warning(self, "Uyarı", "Excel dosyasında 3. satırdan itibaren okunabilir veri bulunamadı.")
                return False

            self.df_target = self.df_actual[['Time', 'Target Acceleration', 'Target Velocity']].copy()
            if 'Target Velocity' in self.df_target.columns and 'Time' in self.df_target.columns:
                self.df_target['Spul_Raw'] = np.where(
                    (self.df_target['Time'] != 0) & (self.df_target['Time'].notna()),
                    (self.df_target['Target Velocity']**2) / self.df_target['Time'],
                    0
                )
                self.df_target['Spul'] = self.df_target['Spul_Raw']

            self.draw_current_graph()
            return True

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Veri işlenirken bir hata oluştu:\n{str(e)}")
            return False

    def apply_offset_to_actual(self, row_offset):
        # Actual hız/ivme verisini zaman eksenini oynatmadan satır bazlı kaydır.
        # Pozitif offset grafiği 0'a yaklaştırır; bu yüzden değerler yukarı kaydırılır.
        df_plot = self.df_actual.copy()
        df_plot['Offset_Time'] = df_plot['Time']
        value_shift = -row_offset
        # SPUL'u hazır seri olarak kaydırma; offsetli Velocity ve mevcut zaman
        # ekseni üzerinden yeniden hesapla ki offset değişimi SPUL değerine de yansısın.
        for col in ['Velocity', 'Acceleration']:
            if col in df_plot.columns:
                df_plot[col] = df_plot[col].shift(value_shift)
        if 'Velocity' in df_plot.columns:
            valid_time = (df_plot['Offset_Time'] > 0) & df_plot['Offset_Time'].notna()
            df_plot['Spul'] = np.nan
            df_plot.loc[valid_time, 'Spul'] = (df_plot.loc[valid_time, 'Velocity'] ** 2) / df_plot.loc[valid_time, 'Offset_Time']
        return df_plot[(df_plot['Offset_Time'] >= 0) & (df_plot['Offset_Time'] <= MAX_GRAPH_TIME_SEC)]

    def apply_offset_to_target(self):
        if self.df_target is None:
            return None
        # Target verisi offsetten etkilenmez; kendi orijinal zaman ekseninde sabit kalır.
        df_plot = self.df_target.copy()
        df_plot['Offset_Time'] = df_plot['Time']
        if 'Target Velocity' in df_plot.columns:
            valid_time = (df_plot['Offset_Time'] > 0) & df_plot['Offset_Time'].notna()
            df_plot['Spul'] = np.nan
            df_plot.loc[valid_time, 'Spul'] = (df_plot.loc[valid_time, 'Target Velocity'] ** 2) / df_plot.loc[valid_time, 'Offset_Time']
        return df_plot[(df_plot['Offset_Time'] >= 0) & (df_plot['Offset_Time'] <= MAX_GRAPH_TIME_SEC)]

    def _series_data(self, df, value_col, trim_trailing_zeros=False):
        series = df[['Offset_Time', value_col]].dropna()
        series = series[series['Offset_Time'] <= MAX_GRAPH_TIME_SEC]
        if trim_trailing_zeros:
            series = self._trim_trailing_zeros(series, value_col)
        return series

    def _trim_trailing_zeros(self, series, value_col):
        non_zero = series[value_col].abs() > 1e-9
        if non_zero.any():
            return series.loc[:non_zero[non_zero].index[-1]]
        return series

    def _set_time_xlim(self, *dfs):
        max_times = []
        for df in dfs:
            if df is not None and 'Offset_Time' in df.columns:
                times = df['Offset_Time'].dropna()
                times = times[(times >= 0) & (times <= MAX_GRAPH_TIME_SEC)]
                if not times.empty:
                    max_times.append(times.max())
        right = min(MAX_GRAPH_TIME_SEC, max(max_times)) if max_times else MAX_GRAPH_TIME_SEC
        self.ax.set_xlim(left=0, right=max(right, 0.001))

    def _max_value_and_time(self, series, value_col):
        if series.empty or series[value_col].dropna().empty:
            return np.nan, 0
        idx = series[value_col].idxmax()
        return series.loc[idx, value_col], series.loc[idx, 'Offset_Time']


    def _style_axes(self, ax, *, zero_line=True):
        ax.set_facecolor('#fbfcfe')
        ax.xaxis.set_major_locator(MultipleLocator(0.02))
        ax.xaxis.set_minor_locator(MultipleLocator(0.005))
        ax.xaxis.set_major_formatter(FormatStrFormatter('%.2f'))
        ax.yaxis.set_minor_locator(AutoMinorLocator(5))
        ax.grid(True, which='major', color='#b0bec5', linewidth=0.95, alpha=0.9)
        ax.grid(True, which='minor', color='#dfe7ec', linewidth=0.55, alpha=0.8)
        for spine in ax.spines.values():
            spine.set_color('#607d8b')
            spine.set_linewidth(1.0)
        ax.tick_params(colors='#263238', labelsize=9)
        ax.tick_params(axis='x', labelrotation=0, pad=4)
        if zero_line:
            ax.axhline(0, color='#111111', linewidth=2.2, alpha=0.95, zorder=1)

    def _set_y_limits_with_zero(self, ax, *series_and_cols, min_span=1.0):
        values = []
        for series, col in series_and_cols:
            if series is not None and col in series.columns:
                clean = series[col].dropna()
                if not clean.empty:
                    values.extend(clean.tolist())
        if not values:
            ax.set_ylim(-min_span, min_span)
            return
        data_min = min(values + [0])
        data_max = max(values + [0])
        span = max(data_max - data_min, min_span)
        pad = span * 0.12
        bottom = data_min - pad
        top = data_max + pad
        if data_min >= 0:
            bottom = -max(pad, min_span * 0.08)
        if data_max <= 0:
            top = max(pad, min_span * 0.08)
        ax.set_ylim(bottom, top)

    def _draw_peak_line(self, ax, x, y, color):
        if pd.isna(y):
            return
        y0, y1 = ax.get_ylim()
        baseline = 0 if y0 <= 0 <= y1 else y0
        ax.vlines(x=x, ymin=baseline, ymax=y, colors=color, linestyles='--', linewidth=1.9, alpha=0.95, zorder=4)
        ax.scatter([x], [y], color=color, edgecolor='white', linewidth=0.9, s=46, zorder=5)

    def _cleanup_axes(self):
        self.ax.clear()
        if self.ax2 is not None:
            self.ax2.remove()
            self.ax2 = None
        self.ax_table.clear()
        self.ax_table.axis('off')

    def draw_current_graph(self):
        if self.df_actual is None:
            return

        row_offset = self.get_current_row_offset()
        df_plot = self.apply_offset_to_actual(row_offset)
        df_target_plot = self.apply_offset_to_target()

        self._cleanup_axes()

        idx = self.current_graph_idx

        if idx == 0:
            self._draw_spul(df_plot, df_target_plot)
        elif idx == 1:
            self._draw_acc_vel(df_plot)
        elif idx == 2:
            self._draw_acc_target_acc(df_plot, df_target_plot)

        self.figure.set_size_inches(9.8, 6.6, forward=True)
        self.figure.subplots_adjust(left=0.07, right=0.98, top=0.96, bottom=0.12, hspace=0.32)
        self.canvas.draw()

    def _draw_spul(self, df_plot, df_target_plot=None):
        if 'Spul' not in df_plot.columns:
            return

        actual_color = '#c77c00'
        target_color = '#2a52be'

        actual_spul = self._series_data(df_plot, 'Spul', trim_trailing_zeros=True)
        self.ax.plot(actual_spul['Offset_Time'].values, actual_spul['Spul'].values, color=actual_color, linewidth=2.8, label="SPUL", zorder=3)
        max_actual_spul, max_actual_time_sec = self._max_value_and_time(actual_spul, 'Spul')


        max_target_spul = "-"
        max_target_time_ms = "-"
        if df_target_plot is not None and 'Spul' in df_target_plot.columns:
            target_spul = self._series_data(df_target_plot, 'Spul', trim_trailing_zeros=True)
            self.ax.plot(target_spul['Offset_Time'].values, target_spul['Spul'].values, color=target_color, linewidth=2.4, linestyle='--', label="Target Spul", alpha=0.9, zorder=3)
            max_target_spul, max_target_time_sec = self._max_value_and_time(target_spul, 'Spul')
            if not pd.isna(max_target_spul):
                max_target_time_ms = max_target_time_sec * 1000.0

        self.ax.set_xlabel("time [s]", labelpad=10)
        self.ax.set_ylabel("Spul [(m/s)²/s]")
        self.ax.legend(
            loc='upper center',
            bbox_to_anchor=(0.5, -0.15),
            ncol=2,
            frameon=False,
            fontsize=14,
            handlelength=2.0
        )
        self._set_time_xlim(df_plot, df_target_plot)
        self._set_y_limits_with_zero(self.ax, (actual_spul, 'Spul'), (target_spul if df_target_plot is not None and 'Spul' in df_target_plot.columns else None, 'Spul'), min_span=1.0)
        self._style_axes(self.ax)
        self._draw_peak_line(self.ax, max_actual_time_sec, max_actual_spul, actual_color)
        if max_target_spul != '-' and not pd.isna(max_target_spul):
            self._draw_peak_line(self.ax, max_target_time_sec, max_target_spul, target_color)

        # Tablo
        actual_val_str = f"{max_actual_spul:.1f} $m^2/s^3$ ({max_actual_time_sec*1000.0:.1f} ms)" if not pd.isna(max_actual_spul) else "-"
        target_val_str = f"{max_target_spul:.1f} $m^2/s^3$ ({max_target_time_ms:.1f} ms)" if not pd.isna(max_target_spul) and max_target_spul != "-" else "-"

        cell_text = [
            ["SPUL", actual_val_str, ""],
            ["Target Spul", target_val_str, ""]
        ]
        self._build_table(cell_text, "SPUL ($f(t)=v^2/t$)")

    def _draw_acc_vel(self, df_plot):
        if 'Acceleration' not in df_plot.columns or 'Velocity' not in df_plot.columns:
            self.ax.text(0.5, 0.5, "Acceleration veya Velocity Sütunu Bulunamadı", ha='center', va='center')
            return

        acc_color = '#1f77b4' # Mavi
        vel_color = '#2ca02c' # Yeşil

        self.ax2 = self.ax.twinx()

        acc_series = self._series_data(df_plot, 'Acceleration')
        vel_series = self._series_data(df_plot, 'Velocity')
        l1 = self.ax.plot(acc_series['Offset_Time'].values, acc_series['Acceleration'].values, color=acc_color, linewidth=2.8, label="Acceleration", zorder=3)
        l2 = self.ax2.plot(vel_series['Offset_Time'].values, vel_series['Velocity'].values, color=vel_color, linewidth=2.6, linestyle='-.', label="Velocity", alpha=0.92, zorder=3)

        max_acc, max_acc_t = self._max_value_and_time(acc_series, 'Acceleration')

        max_vel, max_vel_t = self._max_value_and_time(vel_series, 'Velocity')

        self.ax.set_xlabel("Time, (s)", labelpad=10)
        self.ax.set_ylabel("Acceleration, (g)")
        self.ax2.set_ylabel("Velocity, (m/s)")

        # Legend (Aşağıda ortalanmış bir şekilde iki kutu)
        lines = l1 + l2
        labels = [l.get_label() for l in lines]
        self.ax.legend(
            lines, labels,
            loc='upper center',
            bbox_to_anchor=(0.5, -0.15),
            ncol=2,
            frameon=False,
            fontsize=14,
            handlelength=2.0
        )
        self._set_time_xlim(acc_series, vel_series)
        self.ax2.set_xlim(self.ax.get_xlim())
        self._set_y_limits_with_zero(self.ax, (acc_series, 'Acceleration'), min_span=1.0)
        self._set_y_limits_with_zero(self.ax2, (vel_series, 'Velocity'), min_span=0.2)
        self._style_axes(self.ax)
        self._style_axes(self.ax2, zero_line=False)
        self._draw_peak_line(self.ax, max_acc_t, max_acc, acc_color)
        self._draw_peak_line(self.ax2, max_vel_t, max_vel, vel_color)

        # Tablo
        v_str = f"{max_vel:.2f} $m/s$ ({max_vel_t*1000.0:.1f} ms)" if not pd.isna(max_vel) else "-"
        a_str = f"{max_acc:.2f} g     ({max_acc_t*1000.0:.1f} ms)" if not pd.isna(max_acc) else "-"
        cell_text = [
            ["Sled Velocity", v_str, ""],
            ["Sled Acceleration", a_str, ""]
        ]
        self._build_table(cell_text, "Sled Acceleration and Velocity")

    def _draw_acc_target_acc(self, df_plot, df_target_plot=None):
        if 'Acceleration' not in df_plot.columns:
            self.ax.text(0.5, 0.5, "Actual'da Acceleration Sütunu Bulunamadı", ha='center', va='center')
            return

        acc_color = '#1f77b4'
        target_pulse_color = '#c20078' # Magenta (Morumsı)

        acc_series = self._series_data(df_plot, 'Acceleration')
        l1 = self.ax.plot(acc_series['Offset_Time'].values, acc_series['Acceleration'].values, color=acc_color, linewidth=2.8, label="Acceleration", zorder=3)

        max_acc, max_acc_t = self._max_value_and_time(acc_series, 'Acceleration')

        max_t_acc = "-"
        max_t_acc_t = "-"
        l2 = []
        if df_target_plot is not None and 'Target Acceleration' in df_target_plot.columns:
            target_acc_series = self._series_data(df_target_plot, 'Target Acceleration')
            l2 = self.ax.plot(target_acc_series['Offset_Time'].values, target_acc_series['Target Acceleration'].values, color=target_pulse_color, linewidth=2.4, linestyle='--', label="Target Pulse", alpha=0.9, zorder=3)
            max_t_acc, max_t_acc_t_sec = self._max_value_and_time(target_acc_series, 'Target Acceleration')
            if not pd.isna(max_t_acc):
                max_t_acc_t = max_t_acc_t_sec * 1000.0

        self.ax.set_xlabel("Time, (s)", labelpad=10)
        self.ax.set_ylabel("Acceleration, (g)")

        lines = l1 + l2
        labels = [l.get_label() for l in lines]
        self.ax.legend(
            lines, labels,
            loc='upper center',
            bbox_to_anchor=(0.5, -0.15),
            ncol=2,
            frameon=False,
            fontsize=14,
            handlelength=2.0
        )
        self._set_time_xlim(df_plot, df_target_plot)
        self._set_y_limits_with_zero(self.ax, (acc_series, 'Acceleration'), (target_acc_series if df_target_plot is not None and 'Target Acceleration' in df_target_plot.columns else None, 'Target Acceleration'), min_span=1.0)
        self._style_axes(self.ax)
        self._draw_peak_line(self.ax, max_acc_t, max_acc, acc_color)
        if max_t_acc != '-' and not pd.isna(max_t_acc):
            self._draw_peak_line(self.ax, max_t_acc_t_sec, max_t_acc, target_pulse_color)

        # Tablo
        a_str = f"{max_acc:.2f} g     ({max_acc_t*1000.0:.1f} ms)" if not pd.isna(max_acc) else "-"
        t_str = f"{max_t_acc:.2f} g     ({max_t_acc_t:.1f} ms)" if max_t_acc != "-" else "-"
        cell_text = [
            ["Sled Acceleration", a_str, ""],
            ["Target Acceleration", t_str, ""]
        ]
        self._build_table(cell_text, "Sled vs. Target Acceleration")

    def _build_table(self, cell_text, graph_name_text):
        col_labels = ["", "Max. Value", "Graph Name"]
        table_rows = [row[:] for row in cell_text]
        if table_rows:
            table_rows[0][2] = graph_name_text
        table = self.ax_table.table(
            cellText=table_rows,
            colLabels=col_labels,
            colWidths=[0.25, 0.45, 0.30],
            loc='center',
            cellLoc='center',
            bbox=[0.02, 0.08, 0.96, 0.84],
        )
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1.0, 0.95)

        for (row, col), cell in table.get_celld().items():
            cell.PAD = 0.035
            cell.set_text_props(ha='center', va='center')
            if row == 0:
                cell.set_text_props(weight='bold', ha='center', va='center')
                cell.set_height(0.22)
            else:
                cell.set_height(0.34)

            if col == 2 and row == 2:
                cell.visible_edges = 'BRL'
            if col == 2 and row == 1:
                cell.visible_edges = 'TRL'
            if col == 2 and row == 1:
                cell.get_text().set_wrap(True)
                cell.get_text().set_clip_on(True)

    def export_plots(self):
        save_dir = self.txt_export.text().strip()
        if not save_dir or not os.path.exists(save_dir) or not os.path.isdir(save_dir):
            QMessageBox.warning(self, "Hata", "Lütfen önce test numarasını seçin. Kayıt konumu otomatik olarak testin 3-EVA-ACC klasörü olacaktır.")
            return

        if self.df_actual is None and not self.generate_plots():
            return

        try:
            # Current duruma dokunmadan arkada 3 grafiği çizip kaydedeceğiz
            saved_idx = self.current_graph_idx

            names = ["Spul.png", "Acc_vs_Vel.png", "Acc_vs_Targetacc.png"]

            for i in range(3):
                self.current_graph_idx = i
                self.draw_current_graph()
                path = os.path.join(save_dir, names[i])
                self.figure.set_size_inches(8.27, 11.69, forward=True)
                self.figure.savefig(path, dpi=300, orientation='portrait')

            # Restore
            self.current_graph_idx = saved_idx
            self.update_graph_view()

            QMessageBox.information(self, "Başarılı", f"{self.selected_test_name} testi için tüm 3 grafik kaydedildi:\n{save_dir}\n{names[0]}, {names[1]}, {names[2]}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dışa aktarma hatası:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.callbacks) if hasattr(sys, 'callbacks') else QApplication(sys.argv)
    window = SledAnalyzerApp()
    window.show()
    sys.exit(app.exec_())
