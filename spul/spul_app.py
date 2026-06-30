import sys
import os
import subprocess

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
                             QHBoxLayout, QPushButton, QLabel, QFileDialog,
                             QMessageBox, QGroupBox, QLineEdit,
                             QCheckBox, QSpinBox, QTableWidget, QTableWidgetItem, QHeaderView,
                             QAbstractItemView)
from PyQt5.QtCore import Qt

import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


MAX_GRAPH_TIME_SEC = 0.15
DATA_INTERVAL_SEC = 0.0004
ROWS_FOR_14MS = round(0.014 / DATA_INTERVAL_SEC)


class SledAnalyzerApp(QMainWindow):
    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("Sled Test Analyzer (Multi-Graph)")
        self.resize(1100, 900)

        self.data_path = None
        self.df_actual = None
        self.df_target = None

        # State
        self.current_graph_idx = 0
        self.graphs = ["Spul", "Acceleration vs Velocity", "Actual vs Target Acceleration"]
        self.local_offsets = [0, 0, 0]

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        # --- Top Area Layout ---
        top_layout = QHBoxLayout()

        # --- Control Panel (Left) ---
        control_group = QGroupBox("Veri Yükleme ve Ayarlar")
        control_layout = QVBoxLayout()
        control_group.setLayout(control_layout)

        # File Selection
        self.btn_data = QPushButton("Excel Veri Dosyası Yükle")
        self.btn_data.clicked.connect(self.load_data_file)
        self.lbl_data = QLabel("Seçilmedi")
        self.lbl_data.setWordWrap(True)
        control_layout.addWidget(self.btn_data)
        control_layout.addWidget(self.lbl_data)

        lbl_format = QLabel("Format: 3. satırdan itibaren A=Time(s), B=Target Acc(g), C=Target Hız(m/s), D=Actual Acc(g), E=Actual Hız(m/s)")
        lbl_format.setWordWrap(True)
        lbl_format.setStyleSheet("color: gray; font-size: 11px;")
        control_layout.addWidget(lbl_format)

        control_layout.addStretch() # Push items up

        # Action Buttons
        self.btn_generate = QPushButton("Oluştur / Güncelle")
        self.btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.btn_generate.clicked.connect(self.generate_plots)
        control_layout.addWidget(self.btn_generate)

        top_layout.addWidget(control_group, stretch=1)

        # --- Offset Table Panel (Right) ---
        offset_group = QGroupBox("Actual Satır Offset Ayarları")
        offset_layout = QVBoxLayout()
        offset_group.setLayout(offset_layout)

        self.table_offset = QTableWidget()
        self.table_offset.setColumnCount(3)
        self.table_offset.setHorizontalHeaderLabels(["Değişken / Grafik", "Satır Offset", "Süre Karşılığı"])
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

            # Column 1: Row offset (one step = one data row = 0.0004 s)
            spin = QSpinBox()
            spin.setRange(-10000, 10000)
            spin.setValue(0)
            spin.setSingleStep(1)
            spin.setSuffix(" satır")
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
        univ_layout = QHBoxLayout()
        univ_layout.addWidget(QLabel("Tüm actual grafiklere aynı satır offsetini uygula:"))
        self.spin_universal = QSpinBox()
        self.spin_universal.setRange(-10000, 10000)
        self.spin_universal.setValue(0)
        self.spin_universal.setSingleStep(1)
        self.spin_universal.setSuffix(" satır")
        self.spin_universal.valueChanged.connect(self.apply_universal_offset)
        univ_layout.addWidget(self.spin_universal)

        # 14 ms tick box: 0.0004 s örnek aralığında 14 ms = 35 satır
        self.check_14ms = QCheckBox(f"Tüm Actual Grafikler İçin 14 ms Sabit Offset ({ROWS_FOR_14MS} satır)")
        self.check_14ms.stateChanged.connect(self.apply_14ms_offset)
        univ_layout.addWidget(self.check_14ms)

        offset_layout.addLayout(univ_layout)
        top_layout.addWidget(offset_group, stretch=2)

        main_layout.addLayout(top_layout)

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

        main_layout.addLayout(nav_layout)

        # --- Plot Area (Matplotlib) ---
        plot_group = QGroupBox("Grafik Ekranı")
        plot_layout = QVBoxLayout()
        plot_group.setLayout(plot_layout)

        self.figure = Figure(figsize=(10, 8))
        self.canvas = FigureCanvas(self.figure)
        plot_layout.addWidget(self.canvas)

        # Tablo ayarı
        import matplotlib.gridspec as gridspec
        self.gs = gridspec.GridSpec(2, 1, height_ratios=[4, 1.2]) # Alt tablonun yüksekliğini biraz daha açtım
        self.ax = self.figure.add_subplot(self.gs[0])
        self.ax_table = self.figure.add_subplot(self.gs[1])
        self.ax_table.axis('off')

        self.ax2 = None # Sağ eksen için

        main_layout.addWidget(plot_group)

        # --- Export Area ---
        export_layout = QHBoxLayout()
        export_layout.addWidget(QLabel("Kayıt Dizini:"))
        self.txt_export = QLineEdit(r"c:\Users\pc1\Desktop\adient_data\velocity_acc_target_spul")
        export_layout.addWidget(self.txt_export)

        self.btn_browse = QPushButton("Gözat...")
        self.btn_browse.clicked.connect(self.browse_export_dir)
        export_layout.addWidget(self.btn_browse)

        self.btn_export = QPushButton("Tüm Grafikleri Kaydet (.png)")
        self.btn_export.clicked.connect(self.export_plots)

        export_layout.addWidget(self.btn_export)

        main_layout.addLayout(export_layout)

        # --- Author Info ---
        lbl_author = QLabel("Created by Efe Nakcı")
        lbl_author.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lbl_author.setStyleSheet("color: gray; font-style: italic; font-size: 11px; padding-top: 5px;")
        main_layout.addWidget(lbl_author)

    def apply_universal_offset(self, val):
        for spin in self.spin_offsets:
            spin.setValue(val)

    def apply_14ms_offset(self, state):
        if state == Qt.Checked:
            # 14 ms, 0.0004 s örnek aralığında 35 satıra karşılık gelir.
            for spin in self.spin_offsets:
                spin.setValue(ROWS_FOR_14MS)
                spin.setEnabled(False)
            self.spin_universal.setValue(ROWS_FOR_14MS)
            self.spin_universal.setEnabled(False)
        else:
            # Re-enable manual edit
            for spin in self.spin_offsets:
                spin.setEnabled(True)
            self.spin_universal.setEnabled(True)

    def browse_export_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Kayıt Klasörü Seç", self.txt_export.text())
        if directory:
            self.txt_export.setText(directory)

    def set_local_offset(self, idx, val):
        row_offset = int(val)
        self.local_offsets[idx] = row_offset
        if idx < len(self.offset_duration_items):
            self.offset_duration_items[idx].setText(self.format_offset_duration(row_offset))
        if self.current_graph_idx == idx and self.df_actual is not None:
            self.draw_current_graph()

    def format_offset_duration(self, row_offset):
        offset_ms = row_offset * DATA_INTERVAL_SEC * 1000.0
        return f"{offset_ms:.1f} ms ({row_offset} satır × 0.0004 s)"

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

    def load_data_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excel Veri Dosyası Seç", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.data_path = path
            self.lbl_data.setText(os.path.basename(path))

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
            QMessageBox.warning(self, "Uyarı", "Lütfen tek Excel veri dosyasını yükleyin.")
            return

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
                return

            # Formül gereksinimi kontrol et (Spul = V^2 / t)
            if 'Velocity' in self.df_actual.columns and 'Time' in self.df_actual.columns:
                self.df_actual['Spul_Raw'] = np.where(
                    (self.df_actual['Time'] != 0) & (self.df_actual['Time'].notna()),
                    (self.df_actual['Velocity']**2) / self.df_actual['Time'],
                    0
                )
                self.df_actual['Spul'] = self.df_actual['Spul_Raw']

            self.df_target = self.df_actual[['Time', 'Target Acceleration', 'Target Velocity']].copy()
            if 'Target Velocity' in self.df_target.columns and 'Time' in self.df_target.columns:
                self.df_target['Spul_Raw'] = np.where(
                    (self.df_target['Time'] != 0) & (self.df_target['Time'].notna()),
                    (self.df_target['Target Velocity']**2) / self.df_target['Time'],
                    0
                )
                self.df_target['Spul'] = self.df_target['Spul_Raw']

            self.draw_current_graph()

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Veri işlenirken bir hata oluştu:\n{str(e)}")

    def apply_offset_to_actual(self, row_offset):
        # Actual hız/ivme verisini zaman eksenini oynatmadan satır bazlı kaydır.
        # Pozitif değer veriyi aşağı kaydırır; negatif değer yukarı kaydırır.
        df_plot = self.df_actual.copy()
        df_plot['Offset_Time'] = df_plot['Time']
        for col in ['Velocity', 'Acceleration']:
            if col in df_plot.columns:
                df_plot[col] = df_plot[col].shift(row_offset)
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

    def _series_data(self, df, value_col):
        series = df[['Offset_Time', value_col]].dropna()
        return series[series['Offset_Time'] <= MAX_GRAPH_TIME_SEC]

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

        self.figure.tight_layout()
        self.canvas.draw()

    def _draw_spul(self, df_plot, df_target_plot=None):
        if 'Spul' not in df_plot.columns:
            return

        actual_color = '#FFD700'
        target_color = '#2a52be'

        actual_spul = self._series_data(df_plot, 'Spul')
        self.ax.plot(actual_spul['Offset_Time'].values, actual_spul['Spul'].values, color=actual_color, linewidth=2, label="SPUL")
        max_actual_spul, max_actual_time_sec = self._max_value_and_time(actual_spul, 'Spul')

        if not pd.isna(max_actual_spul):
            self.ax.vlines(x=max_actual_time_sec, ymin=0, ymax=max_actual_spul, colors=actual_color, linestyles='--', linewidth=1, alpha=0.7)

        max_target_spul = "-"
        max_target_time_ms = "-"
        if df_target_plot is not None and 'Spul' in df_target_plot.columns:
            target_spul = self._series_data(df_target_plot, 'Spul')
            self.ax.plot(target_spul['Offset_Time'].values, target_spul['Spul'].values, color=target_color, linewidth=2, linestyle='-', label="Target Spul")
            max_target_spul, max_target_time_sec = self._max_value_and_time(target_spul, 'Spul')
            if not pd.isna(max_target_spul):
                max_target_time_ms = max_target_time_sec * 1000.0
                self.ax.vlines(x=max_target_time_sec, ymin=0, ymax=max_target_spul, colors=target_color, linestyles='--', linewidth=1, alpha=0.7)

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
        self.ax.grid(True)
        self._set_time_xlim(df_plot, df_target_plot)
        self.ax.set_ylim(bottom=0)

        # Tablo
        actual_val_str = f"{max_actual_spul:.1f}  $m^2/s^3$   ({max_actual_time_sec*1000.0:.1f} ms)" if not pd.isna(max_actual_spul) else "-"
        target_val_str = f"{max_target_spul:.1f}  $m^2/s^3$   ({max_target_time_ms:.1f} ms)" if not pd.isna(max_target_spul) and max_target_spul != "-" else "-"

        cell_text = [
            ["SPUL", actual_val_str, ""],
            ["Target Spul", target_val_str, ""]
        ]
        self._build_table(cell_text, "SPUL\nSpecific Accident Capability\n$f(t) = v^2 / t$")

    def _draw_acc_vel(self, df_plot):
        if 'Acceleration' not in df_plot.columns or 'Velocity' not in df_plot.columns:
            self.ax.text(0.5, 0.5, "Acceleration veya Velocity Sütunu Bulunamadı", ha='center', va='center')
            return

        acc_color = '#1f77b4' # Mavi
        vel_color = '#2ca02c' # Yeşil

        self.ax2 = self.ax.twinx()

        acc_series = self._series_data(df_plot, 'Acceleration')
        vel_series = self._series_data(df_plot, 'Velocity')
        l1 = self.ax.plot(acc_series['Offset_Time'].values, acc_series['Acceleration'].values, color=acc_color, linewidth=2, label="Acceleration")
        l2 = self.ax2.plot(vel_series['Offset_Time'].values, vel_series['Velocity'].values, color=vel_color, linewidth=2, label="Velocity")

        max_acc, max_acc_t = self._max_value_and_time(acc_series, 'Acceleration')
        if not pd.isna(max_acc):
            self.ax.vlines(x=max_acc_t, ymin=0, ymax=max_acc, colors=acc_color, linestyles='--', linewidth=1, alpha=0.7)

        max_vel, max_vel_t = self._max_value_and_time(vel_series, 'Velocity')
        if not pd.isna(max_vel):
            self.ax2.vlines(x=max_vel_t, ymin=0, ymax=max_vel, colors=vel_color, linestyles='--', linewidth=1, alpha=0.7)

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
        self.ax.grid(True, alpha=0.5)
        self._set_time_xlim(acc_series, vel_series)
        self.ax2.set_xlim(self.ax.get_xlim())

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
        l1 = self.ax.plot(acc_series['Offset_Time'].values, acc_series['Acceleration'].values, color=acc_color, linewidth=2, label="Acceleration")

        max_acc, max_acc_t = self._max_value_and_time(acc_series, 'Acceleration')
        if not pd.isna(max_acc):
            self.ax.vlines(x=max_acc_t, ymin=0, ymax=max_acc, colors=acc_color, linestyles='--', linewidth=1, alpha=0.7)

        max_t_acc = "-"
        max_t_acc_t = "-"
        l2 = []
        if df_target_plot is not None and 'Target Acceleration' in df_target_plot.columns:
            target_acc_series = self._series_data(df_target_plot, 'Target Acceleration')
            l2 = self.ax.plot(target_acc_series['Offset_Time'].values, target_acc_series['Target Acceleration'].values, color=target_pulse_color, linewidth=2, label="Target Pulse")
            max_t_acc, max_t_acc_t_sec = self._max_value_and_time(target_acc_series, 'Target Acceleration')
            if not pd.isna(max_t_acc):
                max_t_acc_t = max_t_acc_t_sec * 1000.0
                self.ax.vlines(x=max_t_acc_t_sec, ymin=0, ymax=max_t_acc, colors=target_pulse_color, linestyles='--', linewidth=1, alpha=0.7)

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
        self.ax.grid(True, alpha=0.5)
        self._set_time_xlim(df_plot, df_target_plot)

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
        table = self.ax_table.table(cellText=cell_text, colLabels=col_labels, loc='center', cellLoc='center', bbox=[0, 0, 1, 1])
        table.auto_set_font_size(False)
        table.set_fontsize(10)

        for (row, col), cell in table.get_celld().items():
            cell.set_text_props(ha='center', va='center')
            if row == 0:
                cell.set_text_props(weight='bold', ha='center', va='center')

            if col == 2 and row == 2:
                cell.visible_edges = 'BRL'
            if col == 2 and row == 1:
                cell.visible_edges = 'TRL'

        self.ax_table.text(0.833, 0.333, graph_name_text, ha='center', va='center', fontsize=10, transform=self.ax_table.transAxes)

    def export_plots(self):
        save_dir = self.txt_export.text()
        if not os.path.exists(save_dir) or not os.path.isdir(save_dir):
            QMessageBox.warning(self, "Hata", "Geçersiz kayıt dizini.")
            return

        if self.df_actual is None:
            QMessageBox.warning(self, "Hata", "İşlenecek Excel verisi yok!")
            return

        try:
            # Current duruma dokunmadan arkada 3 grafiği çizip kaydedeceğiz
            saved_idx = self.current_graph_idx

            names = ["Spul.png", "Acc_vs_Vel.png", "Acc_vs_Targetacc.png"]

            for i in range(3):
                self.current_graph_idx = i
                self.draw_current_graph()
                path = os.path.join(save_dir, names[i])
                self.figure.savefig(path, dpi=300, bbox_inches='tight')

            # Restore
            self.current_graph_idx = saved_idx
            self.update_graph_view()

            QMessageBox.information(self, "Başarılı", f"Tüm 3 grafik seçilen klasöre kaydedildi:\n{names[0]}, {names[1]}, {names[2]}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dışa aktarma hatası:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.callbacks) if hasattr(sys, 'callbacks') else QApplication(sys.argv)
    window = SledAnalyzerApp()
    window.show()
    sys.exit(app.exec_())
