from __future__ import annotations

import sys
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import json
import re
import numpy as np
import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QDate, QDateTime, QMargins, QTimer, QModelIndex, Qt
from PySide6.QtGui import QAction, QFont, QPainter
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDoubleSpinBox,
    QDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QTabWidget,
    QTableView,
    QVBoxLayout,
    QWidget,
    QFileDialog,
)
from PySide6.QtCharts import (
    QAreaSeries,
    QBarCategoryAxis,
    QBarSeries,
    QBarSet,
    QBoxPlotSeries,
    QBoxSet,
    QChart,
    QChartView,
    QDateTimeAxis,
    QLineSeries,
    QScatterSeries,
    QValueAxis,
)

from core import load_config, save_config, safe_update, load_latest_df


class DataFrameModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame) -> None:
        super().__init__()
        self._df = df

    def set_df(self, df: pd.DataFrame) -> None:
        self.beginResetModel()
        self._df = df
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if self._df is None else len(self._df.index)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if self._df is None else len(self._df.columns)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        if not index.isValid() or self._df is None:
            return None
        if role == Qt.DisplayRole:
            value = self._df.iloc[index.row(), index.column()]
            return "" if pd.isna(value) else str(value)
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> Any:
        if self._df is None:
            return None
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        return str(self._df.index[section])


@dataclass
class ChartConfig:
    chart_type: str
    x_col: Optional[str]
    y_col: Optional[str]
    color_col: Optional[str]
    size_col: Optional[str]
    agg: str
    top_n: Optional[int]
    bins: int
    title: str
    x_cols: Optional[List[str]] = None
    measure_cols: Optional[List[str]] = None
    pivot_rows: Optional[List[str]] = None
    pivot_cols: Optional[List[str]] = None
    pivot_vals: Optional[List[str]] = None
    pivot_agg: Optional[str] = None
    pivot_chart_type: Optional[str] = None


DATA_DIR = Path("data")
TEMPLATES_PATH = DATA_DIR / "templates.json"
COLUMN_META_PATH = DATA_DIR / "column_meta.json"
METRIC_X_LABEL = "(度量)"


def ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def load_templates() -> List[Dict[str, Any]]:
    ensure_data_dir()
    if not TEMPLATES_PATH.exists():
        return []
    try:
        with TEMPLATES_PATH.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def save_templates(templates: List[Dict[str, Any]]) -> None:
    ensure_data_dir()
    with TEMPLATES_PATH.open("w", encoding="utf-8") as f:
        json.dump(templates, f, indent=2, ensure_ascii=False)


def load_column_meta() -> Dict[str, str]:
    ensure_data_dir()
    if not COLUMN_META_PATH.exists():
        return {}
    try:
        with COLUMN_META_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return {str(k): str(v) for k, v in data.items()}
    except Exception:
        return {}
    return {}


def save_column_meta(meta: Dict[str, str]) -> None:
    ensure_data_dir()
    with COLUMN_META_PATH.open("w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2, ensure_ascii=False)


def next_template_name(templates: List[Dict[str, Any]]) -> str:
    existing = {tpl.get("name", "") for tpl in templates}
    idx = 1
    while f"模板{idx}" in existing:
        idx += 1
    return f"模板{idx}"


def normalize_json_value(value: Any) -> Any:
    if isinstance(value, (np.integer, np.floating)):
        return value.item()
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()
    return value


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("CSV Dashboard Desktop")
        self.resize(1400, 900)

        self.df: Optional[pd.DataFrame] = None
        self.filtered_df: Optional[pd.DataFrame] = None
        self.parsed_dates: Dict[str, pd.Series] = {}
        self.filter_controls: Dict[str, Dict[str, Any]] = {}
        self.charts: List[ChartConfig] = []
        self.templates: List[Dict[str, Any]] = []
        self.filter_dialog: Optional[QDialog] = None
        self.filter_summary_label: Optional[QLabel] = None
        self.pivot_df: Optional[pd.DataFrame] = None
        self.column_meta: Dict[str, str] = load_column_meta()
        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.timeout.connect(self.refresh_preview)

        self._build_ui()
        self._load_initial_data()

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)

        splitter = QSplitter(Qt.Horizontal)
        root_layout.addWidget(splitter)

        self._build_left_panel(splitter)
        self._build_right_panel(splitter)
        self._build_filter_dialog()
        self._load_templates()

        self.setCentralWidget(root)

        refresh_action = QAction("重新整理", self)
        refresh_action.triggered.connect(self.refresh_preview)
        self.addAction(refresh_action)

    def _build_left_panel(self, parent: QSplitter) -> None:
        left_content = QWidget()
        left_layout = QVBoxLayout(left_content)
        left_layout.setContentsMargins(8, 8, 8, 8)

        left_layout.addWidget(self._build_data_source_group())
        left_layout.addWidget(self._build_filter_launcher_group())
        left_layout.addWidget(self._build_chart_builder_group())
        left_layout.addWidget(self._build_pivot_group())
        left_layout.addWidget(self._build_template_group())
        left_layout.addStretch(1)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(left_content)
        scroll.setFrameShape(QFrame.NoFrame)

        parent.addWidget(scroll)
        parent.setStretchFactor(0, 0)

    def _build_right_panel(self, parent: QSplitter) -> None:
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(8, 8, 8, 8)

        self.tabs = QTabWidget()
        right_layout.addWidget(self.tabs)

        self._build_preview_tab()
        self._build_pivot_tab()
        self._build_dashboard_tab()

        parent.addWidget(right_widget)
        parent.setStretchFactor(1, 1)

    def _build_data_source_group(self) -> QGroupBox:
        group = QGroupBox("資料來源")
        layout = QVBoxLayout(group)

        path_layout = QHBoxLayout()
        self.path_input = QLineEdit()
        self.browse_button = QPushButton("選擇檔案")
        self.browse_button.clicked.connect(self.select_file)
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(self.browse_button)

        layout.addLayout(path_layout)

        form = QFormLayout()
        self.encoding_combo = QComboBox()
        self.encoding_combo.addItems(["auto", "utf-8-sig", "utf-8", "cp950", "big5", "latin1"])
        form.addRow("編碼", self.encoding_combo)

        self.keep_history_checkbox = QCheckBox("保留每日歷史檔")
        self.keep_history_checkbox.setChecked(True)
        form.addRow("", self.keep_history_checkbox)

        self.use_date_checkbox = QCheckBox("使用日期樣板")
        self.use_date_checkbox.setChecked(False)
        form.addRow("", self.use_date_checkbox)

        self.date_format_input = QLineEdit()
        self.date_format_input.setPlaceholderText("%m%d")
        form.addRow("日期格式", self.date_format_input)

        self.auto_load_checkbox = QCheckBox("啟動時自動更新")
        self.auto_load_checkbox.setChecked(False)
        form.addRow("", self.auto_load_checkbox)

        layout.addLayout(form)

        self.date_help = QLabel("路徑可用 {date} 或 {date:%m%d}；未使用時會嘗試替換檔名最後一段數字")
        self.date_help.setWordWrap(True)
        layout.addWidget(self.date_help)

        button_row = QHBoxLayout()
        self.load_button = QPushButton("從路徑更新")
        self.load_button.clicked.connect(lambda: self.load_from_path())
        button_row.addWidget(self.load_button)

        self.upload_button = QPushButton("選檔並更新")
        self.upload_button.clicked.connect(self.select_and_load)
        button_row.addWidget(self.upload_button)

        self.load_latest_button = QPushButton("載入最新")
        self.load_latest_button.clicked.connect(self.load_latest)
        button_row.addWidget(self.load_latest_button)

        layout.addLayout(button_row)

        self.source_status = QLabel("尚未載入資料")
        self.source_status.setWordWrap(True)
        layout.addWidget(self.source_status)

        return group

    def _build_filter_launcher_group(self) -> QGroupBox:
        group = QGroupBox("篩選器")
        layout = QVBoxLayout(group)

        self.open_filter_button = QPushButton("開啟篩選器視窗")
        self.open_filter_button.clicked.connect(self.open_filter_dialog)
        layout.addWidget(self.open_filter_button)

        self.filter_summary_label = QLabel("篩選後筆數: -")
        self.filter_summary_label.setWordWrap(True)
        layout.addWidget(self.filter_summary_label)

        return group

    def _build_filter_dialog(self) -> None:
        self.filter_dialog = QDialog(self)
        self.filter_dialog.setWindowTitle("篩選器")
        self.filter_dialog.setModal(False)
        self.filter_dialog.resize(420, 720)
        layout = QVBoxLayout(self.filter_dialog)
        layout.addWidget(self._build_filter_group())

        btn_layout = QHBoxLayout()
        self.apply_filter_button = QPushButton("套用篩選")
        self.apply_filter_button.clicked.connect(self.apply_filters)
        btn_layout.addWidget(self.apply_filter_button)

        self.reset_filter_button = QPushButton("重置篩選")
        self.reset_filter_button.clicked.connect(self.reset_filters)
        btn_layout.addWidget(self.reset_filter_button)

        layout.addLayout(btn_layout)

        self.filter_status = QLabel("")
        layout.addWidget(self.filter_status)

    def _build_filter_group(self) -> QGroupBox:
        group = QGroupBox("篩選器")
        layout = QVBoxLayout(group)

        layout.addWidget(QLabel("選擇要啟用篩選的欄位"))
        self.filter_search = QLineEdit()
        self.filter_search.setPlaceholderText("搜尋欄位…")
        self.filter_search.textChanged.connect(self.filter_column_items)
        layout.addWidget(self.filter_search)
        self.filter_column_list = QListWidget()
        self.filter_column_list.itemChanged.connect(self.rebuild_filter_widgets)
        layout.addWidget(self.filter_column_list)

        self.filter_area = QScrollArea()
        self.filter_area.setWidgetResizable(True)
        self.filter_area_widget = QWidget()
        self.filter_area_layout = QVBoxLayout(self.filter_area_widget)
        self.filter_area_layout.setSpacing(20)
        self.filter_area_layout.setContentsMargins(6, 6, 6, 6)
        self.filter_area_layout.addStretch(1)
        self.filter_area.setWidget(self.filter_area_widget)
        layout.addWidget(self.filter_area)

        return group

    def _build_chart_builder_group(self) -> QGroupBox:
        group = QGroupBox("欄位架 / 圖表")
        layout = QVBoxLayout(group)

        form = QFormLayout()
        self.chart_type_combo = QComboBox()
        self.chart_type_combo.addItems(["Bar", "Line", "Scatter", "Histogram", "Histogram+Line", "Box"])
        self.chart_type_combo.currentTextChanged.connect(self.update_builder_options)
        form.addRow("圖表類型", self.chart_type_combo)

        self.x_combo = QComboBox()
        form.addRow("X 軸", self.x_combo)

        self.x2_combo = QComboBox()
        form.addRow("X2", self.x2_combo)

        self.x3_combo = QComboBox()
        form.addRow("X3", self.x3_combo)

        self.y_combo = QComboBox()
        form.addRow("Y 軸/度量", self.y_combo)

        self.color_combo = QComboBox()
        form.addRow("顏色", self.color_combo)

        self.size_combo = QComboBox()
        form.addRow("大小", self.size_combo)

        self.agg_combo = QComboBox()
        self.agg_combo.addItems(["sum", "mean", "median", "min", "max", "count", "none"])
        form.addRow("彙總", self.agg_combo)

        self.topn_spin = QSpinBox()
        self.topn_spin.setRange(0, 1000)
        self.topn_spin.setValue(0)
        form.addRow("Top N", self.topn_spin)

        self.bin_spin = QSpinBox()
        self.bin_spin.setRange(5, 200)
        self.bin_spin.setValue(30)
        form.addRow("直方圖 bins", self.bin_spin)

        self.title_input = QLineEdit()
        form.addRow("標題", self.title_input)

        layout.addLayout(form)

        self.full_labels_checkbox = QCheckBox("完整標籤")
        self.full_labels_checkbox.setChecked(False)
        self.full_labels_checkbox.toggled.connect(self.on_full_labels_toggled)
        layout.addWidget(self.full_labels_checkbox)

        self.metric_group = QGroupBox("度量欄位 (X)")
        metric_layout = QVBoxLayout(self.metric_group)
        self.metric_hint = QLabel("勾選多個數值欄位，會把欄位名稱當成 X，比較其彙總值。")
        self.metric_hint.setWordWrap(True)
        metric_layout.addWidget(self.metric_hint)
        self.metric_list = QListWidget()
        self.metric_list.setSpacing(2)
        self.metric_list.setMinimumHeight(120)
        self.metric_list.itemChanged.connect(self.schedule_preview_refresh)
        metric_layout.addWidget(self.metric_list)
        self.metric_group.setVisible(False)
        layout.addWidget(self.metric_group)

        self.add_chart_button = QPushButton("加入儀表板")
        self.add_chart_button.clicked.connect(self.add_chart)
        layout.addWidget(self.add_chart_button)

        self.builder_hint = QLabel("")
        self.builder_hint.setWordWrap(True)
        layout.addWidget(self.builder_hint)

        def on_builder_change() -> None:
            self.update_metric_mode_state()
            self.update_default_title()
            self.schedule_preview_refresh()

        self.x_combo.currentTextChanged.connect(on_builder_change)
        self.x2_combo.currentTextChanged.connect(on_builder_change)
        self.x3_combo.currentTextChanged.connect(on_builder_change)
        self.y_combo.currentTextChanged.connect(on_builder_change)
        self.color_combo.currentTextChanged.connect(on_builder_change)
        self.size_combo.currentTextChanged.connect(on_builder_change)
        self.agg_combo.currentTextChanged.connect(on_builder_change)
        self.topn_spin.valueChanged.connect(on_builder_change)
        self.bin_spin.valueChanged.connect(on_builder_change)
        self.title_input.textChanged.connect(self.schedule_preview_refresh)

        return group

    def _build_pivot_group(self) -> QGroupBox:
        group = QGroupBox("樞紐")
        layout = QVBoxLayout(group)

        hint = QLabel("篩選使用左上「篩選器」；樞紐使用下方欄位設定")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.pivot_rows_list = QListWidget()
        self.pivot_cols_list = QListWidget()
        self.pivot_vals_list = QListWidget()
        for lst in (self.pivot_rows_list, self.pivot_cols_list, self.pivot_vals_list):
            lst.setSpacing(2)
            lst.setMinimumHeight(100)

        layout.addWidget(QLabel("列（Rows）"))
        layout.addWidget(self.pivot_rows_list)
        layout.addWidget(QLabel("欄（Columns）"))
        layout.addWidget(self.pivot_cols_list)
        layout.addWidget(QLabel("值（Values）"))
        layout.addWidget(self.pivot_vals_list)

        form = QFormLayout()
        self.pivot_agg_combo = QComboBox()
        self.pivot_agg_combo.addItems(["sum", "mean", "median", "min", "max", "count"])
        form.addRow("彙總", self.pivot_agg_combo)

        self.pivot_chart_combo = QComboBox()
        self.pivot_chart_combo.addItems(["長條", "折線", "長條+折線"])
        self.pivot_chart_combo.currentTextChanged.connect(self.apply_pivot)
        form.addRow("圖表", self.pivot_chart_combo)
        layout.addLayout(form)

        btn_row = QHBoxLayout()
        self.apply_pivot_button = QPushButton("套用樞紐")
        self.apply_pivot_button.clicked.connect(self.apply_pivot)
        btn_row.addWidget(self.apply_pivot_button)

        self.add_pivot_button = QPushButton("加入儀表板")
        self.add_pivot_button.clicked.connect(self.add_pivot_to_dashboard)
        btn_row.addWidget(self.add_pivot_button)

        self.reset_pivot_button = QPushButton("清空樞紐")
        self.reset_pivot_button.clicked.connect(self.reset_pivot)
        btn_row.addWidget(self.reset_pivot_button)
        layout.addLayout(btn_row)

        return group

    def _build_template_group(self) -> QGroupBox:
        group = QGroupBox("統計模板")
        layout = QVBoxLayout(group)

        self.template_combo = QComboBox()
        layout.addWidget(self.template_combo)

        button_row = QHBoxLayout()
        self.template_apply_button = QPushButton("套用模板")
        self.template_apply_button.clicked.connect(self.apply_template)
        button_row.addWidget(self.template_apply_button)

        self.template_save_button = QPushButton("新增模板")
        self.template_save_button.clicked.connect(self.save_template)
        button_row.addWidget(self.template_save_button)

        self.template_delete_button = QPushButton("刪除模板")
        self.template_delete_button.clicked.connect(self.delete_template)
        button_row.addWidget(self.template_delete_button)

        layout.addLayout(button_row)

        self.template_status = QLabel("")
        self.template_status.setWordWrap(True)
        layout.addWidget(self.template_status)

        return group

    def _build_preview_tab(self) -> None:
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)

        self.preview_message = QLabel("")
        self.preview_message.setWordWrap(True)
        preview_layout.addWidget(self.preview_message)

        self.preview_chart_view = QChartView()
        self.preview_chart_view.setRenderHint(QPainter.Antialiasing)
        self.preview_chart_view.setMinimumHeight(320)
        preview_layout.addWidget(self.preview_chart_view)

        self.preview_table = QTableView()
        self.preview_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        preview_layout.addWidget(self.preview_table)

        self.tabs.addTab(preview_widget, "預覽")

    def _build_pivot_tab(self) -> None:
        pivot_widget = QWidget()
        pivot_layout = QVBoxLayout(pivot_widget)

        self.pivot_message = QLabel("尚未套用樞紐")
        self.pivot_message.setWordWrap(True)
        pivot_layout.addWidget(self.pivot_message)

        self.pivot_chart_view = QChartView()
        self.pivot_chart_view.setRenderHint(QPainter.Antialiasing)
        self.pivot_chart_view.setMinimumHeight(280)
        pivot_layout.addWidget(self.pivot_chart_view)

        self.pivot_table = QTableView()
        self.pivot_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        pivot_layout.addWidget(self.pivot_table)

        self.tabs.addTab(pivot_widget, "樞紐")

    def _build_dashboard_tab(self) -> None:
        dashboard_widget = QWidget()
        dashboard_layout = QVBoxLayout(dashboard_widget)

        self.dashboard_area = QScrollArea()
        self.dashboard_area.setWidgetResizable(True)
        self.dashboard_container = QWidget()
        self.dashboard_grid = QVBoxLayout(self.dashboard_container)
        self.dashboard_grid.addStretch(1)
        self.dashboard_area.setWidget(self.dashboard_container)

        dashboard_layout.addWidget(self.dashboard_area)

        self.tabs.addTab(dashboard_widget, "儀表板")

    def _load_initial_data(self) -> None:
        cfg = load_config()
        self.path_input.setText(cfg.get("source_path", ""))
        encoding = cfg.get("encoding", "auto")
        if encoding in ["auto", "utf-8-sig", "utf-8", "cp950", "big5", "latin1"]:
            self.encoding_combo.setCurrentText(encoding)
        self.keep_history_checkbox.setChecked(bool(cfg.get("keep_history", True)))
        self.use_date_checkbox.setChecked(bool(cfg.get("use_date_template", False)))
        self.date_format_input.setText(cfg.get("date_format", "%m%d"))
        self.auto_load_checkbox.setChecked(bool(cfg.get("auto_load_on_start", False)))

        if cfg.get("auto_load_on_start") and cfg.get("source_path"):
            loaded = self.load_from_path(show_message=False)
            if loaded:
                return

        df = load_latest_df()
        if df is not None:
            self.set_data(df)
            self.source_status.setText("已載入 data/latest.csv")

    def select_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "選擇 CSV", "", "CSV Files (*.csv)")
        if file_path:
            self.path_input.setText(file_path)

    def select_and_load(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "選擇 CSV", "", "CSV Files (*.csv)")
        if not file_path:
            return
        self.path_input.setText(file_path)
        self.load_from_path()

    def resolve_source_path(self, raw_path: str) -> str:
        if not raw_path:
            return raw_path
        if not self.use_date_checkbox.isChecked():
            return raw_path

        fmt = self.date_format_input.text().strip() or "%m%d"
        today_str = datetime.now().strftime(fmt)

        if "{date" in raw_path:
            def replace(match: re.Match[str]) -> str:
                fmt_override = match.group(1) or fmt
                return datetime.now().strftime(fmt_override)

            return re.sub(r"\{date(?::([^}]+))?\}", replace, raw_path)

        name = Path(raw_path).name
        pattern = rf"(\\d{{{len(today_str)}}})(?!.*\\d)"
        new_name = re.sub(pattern, today_str, name, count=1)
        if new_name != name:
            return str(Path(raw_path).with_name(new_name))
        return raw_path

    def open_filter_dialog(self) -> None:
        if self.filter_dialog is None:
            return
        self.filter_dialog.show()
        self.filter_dialog.raise_()
        self.filter_dialog.activateWindow()

    def load_from_path(self, show_message: bool = True) -> bool:
        raw_path = self.path_input.text().strip()
        path = self.resolve_source_path(raw_path)
        if not path:
            if show_message:
                QMessageBox.warning(self, "CSV", "請輸入檔案路徑")
            return False

        cfg = load_config()
        cfg["source_path"] = raw_path
        cfg["encoding"] = self.encoding_combo.currentText()
        cfg["keep_history"] = self.keep_history_checkbox.isChecked()
        cfg["use_date_template"] = self.use_date_checkbox.isChecked()
        cfg["date_format"] = self.date_format_input.text().strip() or "%m%d"
        cfg["auto_load_on_start"] = self.auto_load_checkbox.isChecked()
        save_config(cfg)

        df, err = safe_update(path, cfg["encoding"], "manual")
        if err:
            if show_message:
                QMessageBox.critical(self, "更新失敗", err)
            return False
        if df is None:
            if show_message:
                QMessageBox.critical(self, "更新失敗", "讀取失敗")
            return False
        self.set_data(df)
        self.source_status.setText(f"已更新: {path}")
        return True

    def load_latest(self) -> None:
        df = load_latest_df()
        if df is None:
            QMessageBox.information(self, "提示", "尚無最新資料")
            return
        self.set_data(df)
        self.source_status.setText("已載入 data/latest.csv")

    def set_data(self, df: pd.DataFrame) -> None:
        self.df = df.copy()
        self.filtered_df = df.copy()
        self.parsed_dates = {}

        model = DataFrameModel(self.filtered_df.head(1000))
        self.preview_table.setModel(model)

        self._refresh_column_lists()
        self.update_builder_options()
        self.refresh_pivot_lists()
        self.apply_filters()

    def _refresh_column_lists(self) -> None:
        self.filter_column_list.blockSignals(True)
        self.filter_column_list.clear()
        if self.df is not None:
            for col in self.df.columns:
                item = QListWidgetItem(col)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                desc = self.column_meta.get(col, "")
                item.setToolTip(desc or "尚未設定欄位說明")
                self.filter_column_list.addItem(item)
        self.filter_column_list.blockSignals(False)

    def filter_column_items(self) -> None:
        keyword = self.filter_search.text().strip().lower()
        for i in range(self.filter_column_list.count()):
            item = self.filter_column_list.item(i)
            text = item.text().lower()
            item.setHidden(bool(keyword) and keyword not in text)

    def refresh_pivot_lists(self) -> None:
        if self.df is None:
            return
        all_cols = list(self.df.columns)
        numeric_cols = [col for col in all_cols if pd.api.types.is_numeric_dtype(self.df[col])]

        def refill(widget: QListWidget, options: List[str]) -> None:
            previous = set()
            for idx in range(widget.count()):
                item = widget.item(idx)
                if item.checkState() == Qt.Checked:
                    previous.add(item.text())
            widget.blockSignals(True)
            widget.clear()
            for col in options:
                item = QListWidgetItem(col)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked if col in previous else Qt.Unchecked)
                widget.addItem(item)
            widget.blockSignals(False)

        refill(self.pivot_rows_list, all_cols)
        refill(self.pivot_cols_list, all_cols)
        refill(self.pivot_vals_list, numeric_cols)

    def get_checked_items(self, widget: QListWidget) -> List[str]:
        values: List[str] = []
        for idx in range(widget.count()):
            item = widget.item(idx)
            if item.checkState() == Qt.Checked:
                values.append(item.text())
        return values

    def get_pivot_selection(self) -> Tuple[List[str], List[str], List[str], str]:
        rows = self.get_checked_items(self.pivot_rows_list)
        cols = self.get_checked_items(self.pivot_cols_list)
        vals = self.get_checked_items(self.pivot_vals_list)
        agg = self.pivot_agg_combo.currentText()
        return rows, cols, vals, agg

    def compute_pivot(
        self,
        df: pd.DataFrame,
        rows: List[str],
        cols: List[str],
        vals: List[str],
        agg: str,
    ) -> Tuple[pd.DataFrame, Union[pd.DataFrame, pd.Series]]:
        if vals:
            pivot = pd.pivot_table(
                df,
                index=rows or None,
                columns=cols or None,
                values=vals,
                aggfunc=agg,
                fill_value=0,
                dropna=False,
            )
        else:
            pivot = pd.pivot_table(
                df,
                index=rows or None,
                columns=cols or None,
                aggfunc="size",
                fill_value=0,
                dropna=False,
            )

        pivot_df = pivot.copy()
        if isinstance(pivot_df, pd.Series):
            pivot_df = pivot_df.to_frame(name="value")
        if isinstance(pivot_df.columns, pd.MultiIndex):
            pivot_df.columns = [
                " / ".join([str(c) for c in col if c != ""]) for col in pivot_df.columns
            ]
        pivot_df = pivot_df.reset_index()
        return pivot_df, pivot

    def edit_column_description(self, col: str, label: QLabel) -> None:
        current = self.column_meta.get(col, "")
        text, ok = QInputDialog.getMultiLineText(
            self,
            "欄位說明",
            f"{col} 的說明",
            current,
        )
        if not ok:
            return
        text = text.strip()
        if text:
            self.column_meta[col] = text
        else:
            self.column_meta.pop(col, None)
        save_column_meta(self.column_meta)
        label.setText(text or "（尚未設定欄位說明）")
        for i in range(self.filter_column_list.count()):
            item = self.filter_column_list.item(i)
            if item.text() == col:
                item.setToolTip(text or "尚未設定欄位說明")
                break

    def rebuild_filter_widgets(self) -> None:
        self.filter_controls = {}
        while self.filter_area_layout.count() > 1:
            child = self.filter_area_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        if self.df is None:
            return

        for idx in range(self.filter_column_list.count()):
            item = self.filter_column_list.item(idx)
            if item.checkState() != Qt.Checked:
                continue
            col = item.text()
            series = self.df[col]

            group = QGroupBox(col)
            group.setStyleSheet(
                "QGroupBox { margin-top: 12px; }"
                "QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; }"
            )
            group_layout = QVBoxLayout(group)
            group_layout.setContentsMargins(12, 24, 12, 12)
            group_layout.setSpacing(10)

            desc = self.column_meta.get(col, "")
            desc_label = QLabel(desc or "（尚未設定欄位說明）")
            desc_label.setWordWrap(True)
            desc_label.setStyleSheet("color: #b5b5b5; font-size: 12px;")
            group_layout.addWidget(desc_label)

            edit_btn = QPushButton("編輯欄位說明")
            edit_btn.setMaximumWidth(140)
            edit_btn.clicked.connect(lambda _, c=col, lbl=desc_label: self.edit_column_description(c, lbl))
            group_layout.addWidget(edit_btn)

            if pd.api.types.is_numeric_dtype(series):
                min_val = float(series.min()) if series.notna().any() else 0.0
                max_val = float(series.max()) if series.notna().any() else 0.0

                min_spin = QDoubleSpinBox()
                min_spin.setRange(-1e18, 1e18)
                min_spin.setValue(min_val)
                max_spin = QDoubleSpinBox()
                max_spin.setRange(-1e18, 1e18)
                max_spin.setValue(max_val)
                min_spin.setMinimumHeight(28)
                max_spin.setMinimumHeight(28)

                form = QFormLayout()
                form.addRow("最小值", min_spin)
                form.addRow("最大值", max_spin)
                group_layout.addLayout(form)

                self.filter_controls[col] = {
                    "type": "numeric",
                    "min": min_spin,
                    "max": max_spin,
                }
            else:
                parsed = self._try_parse_datetime(series)
                if parsed is not None:
                    min_date = parsed.min()
                    max_date = parsed.max()

                    start_edit = QDateEdit()
                    start_edit.setCalendarPopup(True)
                    end_edit = QDateEdit()
                    end_edit.setCalendarPopup(True)
                    start_edit.setMinimumHeight(28)
                    end_edit.setMinimumHeight(28)

                    if pd.notna(min_date):
                        start_edit.setDate(QDate(min_date.year, min_date.month, min_date.day))
                    if pd.notna(max_date):
                        end_edit.setDate(QDate(max_date.year, max_date.month, max_date.day))

                    form = QFormLayout()
                    form.addRow("起始", start_edit)
                    form.addRow("結束", end_edit)
                    group_layout.addLayout(form)

                    self.filter_controls[col] = {
                        "type": "datetime",
                        "start": start_edit,
                        "end": end_edit,
                        "series": parsed,
                    }
                else:
                    values = series.dropna().unique().tolist()
                    if len(values) > 2000:
                        values = pd.Series(values).value_counts().head(2000).index.tolist()
                        group_layout.addWidget(QLabel("只顯示前 2000 個值"))

                    list_widget = QListWidget()
                    list_widget.setSpacing(2)
                    for value in sorted(values, key=lambda x: str(x)):
                        item = QListWidgetItem(str(value))
                        item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                        item.setCheckState(Qt.Checked)
                        item.setData(Qt.UserRole, value)
                        list_widget.addItem(item)
                    list_widget.setMinimumHeight(200)
                    group_layout.addWidget(list_widget)

                    self.filter_controls[col] = {
                        "type": "categorical",
                        "widget": list_widget,
                    }

            self.filter_area_layout.insertWidget(self.filter_area_layout.count() - 1, group)

    def apply_filters(self) -> None:
        if self.df is None:
            return
        df = self.df.copy()
        for col, ctrl in self.filter_controls.items():
            if ctrl["type"] == "numeric":
                min_val = ctrl["min"].value()
                max_val = ctrl["max"].value()
                df = df[df[col].between(min_val, max_val)]
            elif ctrl["type"] == "datetime":
                series = ctrl["series"]
                start_date = ctrl["start"].date().toPython()
                end_date = ctrl["end"].date().toPython()
                mask = (series.dt.date >= start_date) & (series.dt.date <= end_date)
                df = df[mask]
            elif ctrl["type"] == "categorical":
                widget: QListWidget = ctrl["widget"]
                selected = []
                for i in range(widget.count()):
                    item = widget.item(i)
                    if item.checkState() == Qt.Checked:
                        selected.append(item.data(Qt.UserRole))
                if selected:
                    df = df[df[col].isin(selected)]
                else:
                    df = df.iloc[0:0]

        self.filtered_df = df
        self.filter_status.setText(f"篩選後筆數: {len(df):,}")
        if self.filter_summary_label is not None:
            self.filter_summary_label.setText(f"篩選後筆數: {len(df):,}")

        model = self.preview_table.model()
        if isinstance(model, DataFrameModel):
            model.set_df(df.head(1000))

        self.refresh_preview()
        self.refresh_dashboard()

    def reset_pivot(self) -> None:
        for widget in (self.pivot_rows_list, self.pivot_cols_list, self.pivot_vals_list):
            for idx in range(widget.count()):
                widget.item(idx).setCheckState(Qt.Unchecked)
        self.pivot_df = None
        self.pivot_message.setText("尚未套用樞紐")
        self.pivot_chart_view.setChart(QChart())
        self.pivot_table.setModel(DataFrameModel(pd.DataFrame()))

    def apply_pivot(self) -> None:
        if self.filtered_df is None:
            return
        rows, cols, vals, agg = self.get_pivot_selection()

        if not rows and not cols:
            self.pivot_message.setText("請至少選擇「列」或「欄」")
            return

        df = self.filtered_df.copy()
        try:
            pivot_df, pivot = self.compute_pivot(df, rows, cols, vals, agg)
            self.pivot_df = pivot_df

            model = DataFrameModel(pivot_df)
            self.pivot_table.setModel(model)
            self.pivot_message.setText(f"樞紐完成：{len(pivot_df):,} 列")
            self.pivot_chart_view.setChart(
                self.build_pivot_chart(pivot, rows, cols, self.pivot_chart_combo.currentText())
            )
            self.tabs.setCurrentIndex(1)
        except Exception as exc:
            self.pivot_message.setText(f"樞紐失敗: {exc}")
            self.pivot_chart_view.setChart(QChart())

    def add_pivot_to_dashboard(self) -> None:
        if self.filtered_df is None:
            return
        rows, cols, vals, agg = self.get_pivot_selection()
        if not rows and not cols:
            QMessageBox.information(self, "樞紐", "請至少選擇「列」或「欄」")
            return
        try:
            pivot_df, pivot = self.compute_pivot(self.filtered_df, rows, cols, vals, agg)
        except Exception as exc:
            QMessageBox.critical(self, "樞紐", f"樞紐失敗: {exc}")
            return

        title_parts = []
        if rows:
            title_parts.append("列: " + " / ".join(rows))
        if cols:
            title_parts.append("欄: " + " / ".join(cols))
        title = "樞紐" if not title_parts else "樞紐 (" + ", ".join(title_parts) + ")"

        cfg = ChartConfig(
            chart_type="Pivot",
            x_col=None,
            y_col=None,
            color_col=None,
            size_col=None,
            agg=agg,
            top_n=None,
            bins=0,
            title=title,
            x_cols=None,
            measure_cols=None,
            pivot_rows=rows,
            pivot_cols=cols,
            pivot_vals=vals,
            pivot_agg=agg,
            pivot_chart_type=self.pivot_chart_combo.currentText(),
        )
        self.charts.append(cfg)
        self.pivot_df = pivot_df
        self.refresh_dashboard()
        self.tabs.setCurrentIndex(2)

    def reset_filters(self) -> None:
        for i in range(self.filter_column_list.count()):
            item = self.filter_column_list.item(i)
            item.setCheckState(Qt.Unchecked)
        self.rebuild_filter_widgets()
        self.apply_filters()

    def get_filter_state(self) -> Dict[str, Any]:
        columns: List[str] = []
        for i in range(self.filter_column_list.count()):
            item = self.filter_column_list.item(i)
            if item.checkState() == Qt.Checked:
                columns.append(item.text())

        values: Dict[str, Any] = {}
        for col, ctrl in self.filter_controls.items():
            if ctrl["type"] == "numeric":
                values[col] = {
                    "type": "numeric",
                    "min": float(ctrl["min"].value()),
                    "max": float(ctrl["max"].value()),
                }
            elif ctrl["type"] == "datetime":
                start = ctrl["start"].date().toPython()
                end = ctrl["end"].date().toPython()
                values[col] = {
                    "type": "datetime",
                    "start": start.isoformat(),
                    "end": end.isoformat(),
                }
            elif ctrl["type"] == "categorical":
                widget: QListWidget = ctrl["widget"]
                selected: List[Any] = []
                for idx in range(widget.count()):
                    item = widget.item(idx)
                    if item.checkState() == Qt.Checked:
                        selected.append(normalize_json_value(item.data(Qt.UserRole)))
                values[col] = {
                    "type": "categorical",
                    "selected": selected,
                }

        return {"columns": columns, "values": values}

    def apply_filter_state(self, state: Dict[str, Any]) -> None:
        if self.df is None:
            return
        columns = set(state.get("columns", []))
        for i in range(self.filter_column_list.count()):
            item = self.filter_column_list.item(i)
            item.setCheckState(Qt.Checked if item.text() in columns else Qt.Unchecked)

        self.rebuild_filter_widgets()

        values = state.get("values", {})
        for col, cfg in values.items():
            ctrl = self.filter_controls.get(col)
            if not ctrl:
                continue
            if ctrl["type"] == "numeric" and cfg.get("type") == "numeric":
                ctrl["min"].setValue(float(cfg.get("min", ctrl["min"].value())))
                ctrl["max"].setValue(float(cfg.get("max", ctrl["max"].value())))
            elif ctrl["type"] == "datetime" and cfg.get("type") == "datetime":
                try:
                    start_date = QDate.fromString(cfg.get("start", ""), "yyyy-MM-dd")
                    end_date = QDate.fromString(cfg.get("end", ""), "yyyy-MM-dd")
                    if start_date.isValid():
                        ctrl["start"].setDate(start_date)
                    if end_date.isValid():
                        ctrl["end"].setDate(end_date)
                except Exception:
                    pass
            elif ctrl["type"] == "categorical" and cfg.get("type") == "categorical":
                selected_raw = cfg.get("selected", [])
                selected_set = {normalize_json_value(v) for v in selected_raw}
                widget: QListWidget = ctrl["widget"]
                for idx in range(widget.count()):
                    item = widget.item(idx)
                    value = normalize_json_value(item.data(Qt.UserRole))
                    item.setCheckState(Qt.Checked if value in selected_set else Qt.Unchecked)

        self.apply_filters()

    def update_builder_options(self) -> None:
        if self.df is None:
            return

        chart_type = self.chart_type_combo.currentText()
        all_cols = list(self.df.columns)
        numeric_cols = [col for col in all_cols if pd.api.types.is_numeric_dtype(self.df[col])]

        def set_combo(combo: QComboBox, options: List[str]) -> None:
            current = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(options)
            if current in options:
                combo.setCurrentText(current)
            combo.blockSignals(False)

        if chart_type in ("Histogram", "Histogram+Line"):
            set_combo(self.x_combo, numeric_cols)
            set_combo(self.x2_combo, ["(none)"])
            set_combo(self.x3_combo, ["(none)"])
            set_combo(self.y_combo, [])
        elif chart_type == "Scatter":
            set_combo(self.x_combo, numeric_cols)
            set_combo(self.x2_combo, ["(none)"])
            set_combo(self.x3_combo, ["(none)"])
            set_combo(self.y_combo, numeric_cols)
        else:
            if chart_type in ("Bar", "Line"):
                set_combo(self.x_combo, [METRIC_X_LABEL] + all_cols)
            else:
                set_combo(self.x_combo, all_cols)
            multi_options = ["(none)"] + all_cols
            set_combo(self.x2_combo, multi_options)
            set_combo(self.x3_combo, multi_options)
            y_options = ["(count)"] + numeric_cols
            set_combo(self.y_combo, y_options)

        color_options = ["(none)"] + all_cols
        set_combo(self.color_combo, color_options)

        size_options = ["(none)"] + numeric_cols
        set_combo(self.size_combo, size_options)

        self.refresh_metric_list(numeric_cols)
        self.update_metric_mode_state()
        self.update_default_title()
        self.schedule_preview_refresh()

    def refresh_metric_list(self, numeric_cols: List[str]) -> None:
        previous = set(self.get_metric_cols_from_ui())
        self.metric_list.blockSignals(True)
        self.metric_list.clear()
        for col in numeric_cols:
            item = QListWidgetItem(col)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if col in previous else Qt.Unchecked)
            self.metric_list.addItem(item)
        self.metric_list.blockSignals(False)

    def is_metric_mode(self) -> bool:
        return self.x_combo.currentText() == METRIC_X_LABEL

    def update_metric_mode_state(self) -> None:
        supported = self.chart_type_combo.currentText() in ("Bar", "Line")
        if not supported:
            self.metric_group.setVisible(False)
            self.x2_combo.setEnabled(True)
            self.x3_combo.setEnabled(True)
            self.y_combo.setEnabled(True)
            self.color_combo.setEnabled(True)
            self.size_combo.setEnabled(True)
            return

        if self.is_metric_mode():
            self.metric_group.setVisible(True)
            self.x2_combo.setEnabled(False)
            self.x3_combo.setEnabled(False)
            self.y_combo.setEnabled(False)
            self.color_combo.setEnabled(False)
            self.size_combo.setEnabled(False)
        else:
            self.metric_group.setVisible(False)
            self.x2_combo.setEnabled(True)
            self.x3_combo.setEnabled(True)
            self.y_combo.setEnabled(True)
            self.color_combo.setEnabled(True)
            self.size_combo.setEnabled(True)

    def schedule_preview_refresh(self, delay_ms: int = 200) -> None:
        if self._preview_timer.isActive():
            self._preview_timer.stop()
        self._preview_timer.start(delay_ms)

    def on_full_labels_toggled(self) -> None:
        self.schedule_preview_refresh()
        self.refresh_dashboard()
        if self.pivot_df is not None:
            self.apply_pivot()

    def _load_templates(self) -> None:
        self.templates = load_templates()
        self.refresh_template_combo()

    def refresh_template_combo(self) -> None:
        if not hasattr(self, "template_combo"):
            return
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        for tpl in self.templates:
            name = tpl.get("name")
            if name:
                self.template_combo.addItem(name)
        self.template_combo.blockSignals(False)

    def save_template(self) -> None:
        if self.df is None:
            QMessageBox.information(self, "模板", "請先載入資料")
            return
        name = next_template_name(self.templates)
        template = {
            "name": name,
            "charts": [asdict(cfg) for cfg in self.charts],
            "filters": self.get_filter_state(),
        }
        self.templates.append(template)
        save_templates(self.templates)
        self.refresh_template_combo()
        self.template_combo.setCurrentText(name)
        self.template_status.setText(f"已儲存 {name}")

    def apply_template(self) -> None:
        name = self.template_combo.currentText()
        if not name:
            return
        template = next((tpl for tpl in self.templates if tpl.get("name") == name), None)
        if not template:
            return
        chart_data = template.get("charts", [])
        charts: List[ChartConfig] = []
        for cfg in chart_data:
            try:
                charts.append(ChartConfig(**cfg))
            except Exception:
                continue
        self.charts = charts
        self.apply_filter_state(template.get("filters", {}))
        self.refresh_dashboard()
        self.template_status.setText(f"已套用 {name}")

    def delete_template(self) -> None:
        name = self.template_combo.currentText()
        if not name:
            return
        self.templates = [tpl for tpl in self.templates if tpl.get("name") != name]
        save_templates(self.templates)
        self.refresh_template_combo()
        self.template_status.setText(f"已刪除 {name}")

    def update_default_title(self) -> None:
        chart_type = self.chart_type_combo.currentText()
        if self.is_metric_mode():
            metrics = self.get_metric_cols_from_ui()
            x_col = " / ".join(metrics) if metrics else "度量"
        else:
            x_cols = self.get_x_cols_from_ui()
            x_col = " / ".join(x_cols) if x_cols else ""
        y_col = self.y_combo.currentText()
        if y_col == "(count)":
            y_col = "count"
        if not x_col:
            title = f"{chart_type}"
        elif y_col and y_col != "(count)":
            title = f"{chart_type}: {x_col} × {y_col}"
        else:
            title = f"{chart_type}: {x_col}"
        if not self.title_input.text():
            self.title_input.setText(title)

    def get_x_cols_from_ui(self) -> List[str]:
        cols: List[str] = []
        for combo in (self.x_combo, self.x2_combo, self.x3_combo):
            value = combo.currentText().strip()
            if not value or value == "(none)":
                continue
            if value not in cols:
                cols.append(value)
        return cols

    def get_metric_cols_from_ui(self) -> List[str]:
        cols: List[str] = []
        for i in range(self.metric_list.count()):
            item = self.metric_list.item(i)
            if item.checkState() == Qt.Checked:
                cols.append(item.text())
        return cols

    def current_config(self) -> ChartConfig:
        y_col = self.y_combo.currentText()
        if y_col == "(count)" or y_col == "":
            y_col = None
        color_col = self.color_combo.currentText()
        if color_col == "(none)":
            color_col = None
        size_col = self.size_combo.currentText()
        if size_col == "(none)":
            size_col = None

        top_n = self.topn_spin.value()
        top_n = top_n if top_n > 0 else None

        metric_cols = self.get_metric_cols_from_ui() if self.is_metric_mode() else None
        x_cols = [] if metric_cols else self.get_x_cols_from_ui()

        return ChartConfig(
            chart_type=self.chart_type_combo.currentText(),
            x_col=x_cols[0] if x_cols else None,
            y_col=y_col,
            color_col=color_col,
            size_col=size_col,
            agg=self.agg_combo.currentText(),
            top_n=top_n,
            bins=self.bin_spin.value(),
            title=self.title_input.text().strip() or "Chart",
            x_cols=x_cols,
            measure_cols=metric_cols,
        )

    def add_chart(self) -> None:
        if self.filtered_df is None:
            return
        cfg = self.current_config()
        self.charts.append(cfg)
        self.refresh_dashboard()
        self.tabs.setCurrentIndex(1)

    def refresh_preview(self) -> None:
        if self.filtered_df is None:
            return
        cfg = self.current_config()
        chart, err = self.build_chart(self.filtered_df, cfg)
        if err:
            self.preview_message.setText(err)
            self.preview_chart_view.setChart(QChart())
        else:
            self.preview_message.setText("")
            self.preview_chart_view.setChart(chart)

    def refresh_dashboard(self) -> None:
        for i in reversed(range(self.dashboard_grid.count() - 1)):
            item = self.dashboard_grid.takeAt(i)
            if item.widget():
                item.widget().deleteLater()

        if self.filtered_df is None:
            return

        if not self.charts:
            empty_label = QLabel("尚未加入圖表")
            self.dashboard_grid.insertWidget(0, empty_label)
            return

        for idx, cfg in enumerate(self.charts):
            frame = QFrame()
            frame.setFrameShape(QFrame.StyledPanel)
            frame_layout = QVBoxLayout(frame)
            title = QLabel(cfg.title)
            title.setStyleSheet("font-weight: bold")
            frame_layout.addWidget(title)

            if cfg.chart_type == "Pivot":
                rows = cfg.pivot_rows or []
                cols = cfg.pivot_cols or []
                vals = cfg.pivot_vals or []
                agg = cfg.pivot_agg or "sum"
                try:
                    pivot_df, pivot_raw = self.compute_pivot(self.filtered_df, rows, cols, vals, agg)
                    view = QChartView(
                        self.build_pivot_chart(
                            pivot_raw,
                            rows,
                            cols,
                            cfg.pivot_chart_type or "長條",
                        )
                    )
                    view.setMinimumHeight(240)
                    view.setRenderHint(QPainter.Antialiasing)
                    frame_layout.addWidget(view)

                    table = QTableView()
                    table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                    table.setModel(DataFrameModel(pivot_df.head(1000)))
                    table.setMinimumHeight(220)
                    frame_layout.addWidget(table)
                except Exception as exc:
                    frame_layout.addWidget(QLabel(f"樞紐失敗: {exc}"))
            else:
                chart, err = self.build_chart(self.filtered_df, cfg)
                if err:
                    frame_layout.addWidget(QLabel(err))
                else:
                    view = QChartView(chart)
                    view.setMinimumHeight(240)
                    view.setRenderHint(QPainter.Antialiasing)
                    frame_layout.addWidget(view)

            remove_btn = QPushButton("移除圖表")
            remove_btn.clicked.connect(lambda _, i=idx: self.remove_chart(i))
            frame_layout.addWidget(remove_btn)

            self.dashboard_grid.insertWidget(idx, frame)

    def remove_chart(self, idx: int) -> None:
        if 0 <= idx < len(self.charts):
            self.charts.pop(idx)
            self.refresh_dashboard()

    def build_chart(self, df: pd.DataFrame, cfg: ChartConfig) -> Tuple[Optional[QChart], Optional[str]]:
        try:
            if df.empty:
                return None, "篩選後沒有資料"

            chart_type = cfg.chart_type
            if chart_type == "Area":
                chart_type = "Line"
            metric_cols = cfg.measure_cols or []
            metric_mode = bool(metric_cols)
            x_cols = cfg.x_cols or ([cfg.x_col] if cfg.x_col else [])
            if chart_type in ("Histogram", "Histogram+Line", "Scatter"):
                x_cols = x_cols[:1]
            x_col = x_cols[0] if x_cols else None
            y_col = cfg.y_col
            color_col = cfg.color_col
            if color_col and color_col in x_cols:
                color_col = None

            if metric_mode:
                if chart_type not in ("Bar", "Line"):
                    return None, "度量模式僅支援 Bar/Line"
                if not metric_cols:
                    return None, "請勾選度量欄位"
                agg = cfg.agg if cfg.agg != "none" else "sum"
                values: List[float] = []
                for col in metric_cols:
                    series = df[col].dropna()
                    if agg == "count":
                        values.append(float(series.count()))
                    elif agg == "mean":
                        values.append(float(series.mean()) if not series.empty else 0.0)
                    elif agg == "median":
                        values.append(float(series.median()) if not series.empty else 0.0)
                    elif agg == "min":
                        values.append(float(series.min()) if not series.empty else 0.0)
                    elif agg == "max":
                        values.append(float(series.max()) if not series.empty else 0.0)
                    else:
                        values.append(float(series.sum()) if not series.empty else 0.0)

                data = pd.DataFrame({"metric": metric_cols, "value": values})
                if cfg.top_n:
                    data = data.sort_values("value", ascending=False).head(cfg.top_n)
                categories = data["metric"].astype(str).tolist()
                chart = QChart()
                chart.setMargins(QMargins(10, 10, 10, 30))
                if chart_type == "Bar":
                    bar_set = QBarSet(agg)
                    bar_set.append([float(v) for v in data["value"].tolist()])
                    series = QBarSeries()
                    series.append(bar_set)
                    chart.addSeries(series)
                    axis_x = QBarCategoryAxis()
                    axis_x.append(categories)
                    axis_x.setLabelsAngle(-45)
                    chart.addAxis(axis_x, Qt.AlignBottom)
                    series.attachAxis(axis_x)
                    axis_y = QValueAxis()
                    if not data["value"].empty:
                        y_max = float(data["value"].max())
                        axis_y.setRange(0, max(y_max * 1.1, 1))
                    chart.addAxis(axis_y, Qt.AlignLeft)
                    series.attachAxis(axis_y)
                else:
                    series = QLineSeries()
                    for idx, value in enumerate(data["value"].tolist()):
                        series.append(float(idx), float(value))
                    chart.addSeries(series)
                    axis_x = QBarCategoryAxis()
                    axis_x.append(categories)
                    axis_x.setLabelsAngle(-45)
                    chart.addAxis(axis_x, Qt.AlignBottom)
                    series.attachAxis(axis_x)
                    axis_y = QValueAxis()
                    if not data["value"].empty:
                        y_min = float(data["value"].min())
                        y_max = float(data["value"].max())
                        if y_min == y_max:
                            axis_y.setRange(y_min - 1, y_max + 1)
                        else:
                            axis_y.setRange(y_min, y_max)
                    chart.addAxis(axis_y, Qt.AlignLeft)
                    series.attachAxis(axis_y)

                chart.setTitle(cfg.title)
                return chart, None

            if chart_type in ("Histogram", "Histogram+Line"):
                if not x_col:
                    return None, "請選擇欄位"
                if not pd.api.types.is_numeric_dtype(df[x_col]):
                    return None, "直方圖需要數值欄位"
                data = df[x_col].dropna().values
                counts, edges = np.histogram(data, bins=cfg.bins)

                categories = [f"{edges[i]:.2f}-{edges[i+1]:.2f}" for i in range(len(counts))]
                bar_set = QBarSet(x_col)
                bar_set.append(list(counts))
                series = QBarSeries()
                series.append(bar_set)

                chart = QChart()
                chart.addSeries(series)
                axis_x = QBarCategoryAxis()
                axis_x.append(categories)
                axis_x.setLabelsAngle(-45)
                chart.addAxis(axis_x, Qt.AlignBottom)
                series.attachAxis(axis_x)
                axis_y = QValueAxis()
                max_count = int(counts.max()) if len(counts) else 1
                axis_y.setRange(0, max(max_count * 1.1, 1))
                chart.addAxis(axis_y, Qt.AlignLeft)
                series.attachAxis(axis_y)
                if chart_type == "Histogram+Line":
                    line = QLineSeries()
                    for idx, value in enumerate(counts):
                        line.append(float(idx), float(value))
                    chart.addSeries(line)
                    line.attachAxis(axis_x)
                    line.attachAxis(axis_y)
                chart.setTitle(cfg.title)
                chart.setMargins(QMargins(10, 10, 10, 30))
                return chart, None

            if chart_type == "Scatter":
                if not x_col or not y_col:
                    return None, "請選擇 X/Y"
                if not pd.api.types.is_numeric_dtype(df[x_col]) or not pd.api.types.is_numeric_dtype(
                    df[y_col]
                ):
                    return None, "散點圖需要數值欄位"

                chart = QChart()
                chart.setMargins(QMargins(10, 10, 10, 30))
                groups = [None]
                if color_col:
                    groups = df[color_col].dropna().unique().tolist()
                max_points = 8000
                max_per_group = max(500, int(max_points / max(len(groups), 1)))

                for group in groups:
                    series = QScatterSeries()
                    name = "全部" if group is None else str(group)
                    series.setName(name)
                    subset = df if group is None else df[df[color_col] == group]
                    subset = subset[[x_col, y_col]].dropna()
                    if len(subset) > max_per_group:
                        subset = subset.sample(n=max_per_group, random_state=42)
                    if x_col == y_col:
                        for value in subset[x_col].values.tolist():
                            series.append(float(value), float(value))
                    else:
                        for _, row in subset.iterrows():
                            x_val = row[x_col]
                            y_val = row[y_col]
                            if isinstance(x_val, pd.Series):
                                x_val = x_val.iloc[0]
                            if isinstance(y_val, pd.Series):
                                y_val = y_val.iloc[0]
                            series.append(float(x_val), float(y_val))
                    chart.addSeries(series)

                axis_x = QValueAxis()
                axis_y = QValueAxis()
                x_vals = df[x_col].dropna()
                y_vals = df[y_col].dropna()
                if not x_vals.empty and not y_vals.empty:
                    axis_x.setRange(float(x_vals.min()), float(x_vals.max()))
                    axis_y.setRange(float(y_vals.min()), float(y_vals.max()))
                chart.addAxis(axis_x, Qt.AlignBottom)
                chart.addAxis(axis_y, Qt.AlignLeft)
                for s in chart.series():
                    s.attachAxis(axis_x)
                    s.attachAxis(axis_y)
                chart.setTitle(cfg.title)
                return chart, None

            if chart_type == "Box":
                if not x_cols or not y_col:
                    return None, "箱型圖需要 X 與 Y"
                if not pd.api.types.is_numeric_dtype(df[y_col]):
                    return None, "箱型圖需要數值欄位"

                if len(x_cols) > 1:
                    data = df[x_cols + [y_col]].dropna()
                    data = data.copy()
                    data["_x_label"] = data[x_cols].astype(str).agg(" / ".join, axis=1)
                    x_key = "_x_label"
                else:
                    data = df[[x_col, y_col]].dropna()
                    x_key = x_col

                chart = QChart()
                chart.setMargins(QMargins(10, 10, 10, 30))
                series = QBoxPlotSeries()
                for key, group in data.groupby(x_key):
                    values = group[y_col]
                    q1 = values.quantile(0.25)
                    median = values.quantile(0.5)
                    q3 = values.quantile(0.75)
                    vmin = values.min()
                    vmax = values.max()
                    box = QBoxSet(str(key))
                    box.setValue(QBoxSet.LowerExtreme, float(vmin))
                    box.setValue(QBoxSet.LowerQuartile, float(q1))
                    box.setValue(QBoxSet.Median, float(median))
                    box.setValue(QBoxSet.UpperQuartile, float(q3))
                    box.setValue(QBoxSet.UpperExtreme, float(vmax))
                    series.append(box)

                chart.addSeries(series)
                axis_x = QBarCategoryAxis()
                axis_x.append([box.label() for box in series.boxes()])
                chart.addAxis(axis_x, Qt.AlignBottom)
                series.attachAxis(axis_x)
                axis_y = QValueAxis()
                if not df[y_col].dropna().empty:
                    y_min = float(df[y_col].min())
                    y_max = float(df[y_col].max())
                    axis_y.setRange(y_min, y_max)
                chart.addAxis(axis_y, Qt.AlignLeft)
                series.attachAxis(axis_y)
                chart.setTitle(cfg.title)
                return chart, None

            if not x_cols:
                return None, "請選擇 X 軸"

            group_cols = list(x_cols)
            if color_col and color_col not in group_cols:
                group_cols.append(color_col)

            if cfg.agg == "none":
                if y_col is None:
                    return None, "請選擇度量欄位"
                if y_col in group_cols:
                    data = df[group_cols].dropna()
                    y_field = y_col
                else:
                    data = df[group_cols + [y_col]].dropna()
                    y_field = y_col
            else:
                if y_col is None or cfg.agg == "count":
                    data = df.groupby(group_cols, dropna=False).size().reset_index(name="count")
                    y_field = "count"
                else:
                    grouped = df.groupby(group_cols, dropna=False)[y_col].agg(cfg.agg)
                    value_name = y_col
                    if y_col in group_cols:
                        value_name = f"{y_col}_{cfg.agg}"
                        grouped = grouped.rename(value_name)
                    data = grouped.reset_index()
                    y_field = value_name

            use_multi_x = len(x_cols) > 1
            x_display_col = x_col
            if use_multi_x:
                data = data.copy()
                data["_x_label"] = data[x_cols].astype(str).agg(" / ".join, axis=1)
                x_display_col = "_x_label"

            if cfg.top_n:
                data = data.sort_values(y_field, ascending=False).head(cfg.top_n)

            chart = QChart()
            chart.setMargins(QMargins(10, 10, 10, 30))
            if chart_type == "Bar":
                categories = data[x_display_col].astype(str).unique().tolist()
                display_categories = self.format_axis_labels(categories)
                series = QBarSeries()
                if color_col:
                    for key in data[color_col].astype(str).unique().tolist():
                        subset = data[data[color_col].astype(str) == key]
                        values = []
                        for cat in categories:
                            value = subset[subset[x_display_col].astype(str) == cat][y_field]
                            values.append(float(value.iloc[0]) if not value.empty else 0.0)
                        bar_set = QBarSet(key)
                        bar_set.append(values)
                        series.append(bar_set)
                else:
                    values = [
                        float(v)
                        for v in data.set_index(x_display_col)[y_field].reindex(categories).fillna(0)
                    ]
                    bar_set = QBarSet(y_field)
                    bar_set.append(values)
                    series.append(bar_set)

                chart.addSeries(series)
                axis_x = QBarCategoryAxis()
                axis_x.append(display_categories)
                axis_x.setLabelsAngle(-45)
                axis_x.setLabelsFont(self.axis_label_font())
                chart.addAxis(axis_x, Qt.AlignBottom)
                series.attachAxis(axis_x)
                axis_y = QValueAxis()
                if not data[y_field].dropna().empty:
                    y_max = float(data[y_field].max())
                    axis_y.setRange(0, max(y_max * 1.1, 1))
                chart.addAxis(axis_y, Qt.AlignLeft)
                series.attachAxis(axis_y)
            else:
                if color_col:
                    groups = data[color_col].astype(str).unique().tolist()
                else:
                    groups = [None]

                x_mode = "value"
                categories: List[str] = []
                category_map: Dict[str, int] = {}
                parsed_dates: Optional[pd.Series] = None

                if use_multi_x:
                    x_mode = "category"
                    categories = data[x_display_col].astype(str).unique().tolist()
                    category_map = {cat: idx for idx, cat in enumerate(categories)}
                else:
                    if pd.api.types.is_numeric_dtype(data[x_col]):
                        x_mode = "value"
                    else:
                        parsed_dates = self._try_parse_datetime(data[x_col])
                        if parsed_dates is not None:
                            x_mode = "datetime"
                            data = data.copy()
                            data["_x_dt"] = parsed_dates
                        else:
                            x_mode = "category"
                            categories = data[x_col].astype(str).unique().tolist()
                            category_map = {cat: idx for idx, cat in enumerate(categories)}

                for key in groups:
                    subset = data if key is None else data[data[color_col].astype(str) == key]
                    sort_col = x_display_col if x_mode == "category" else x_col
                    subset = subset.sort_values(sort_col)
                    series_name = str(key) if key is not None else y_field

                    def append_points(line_series: QLineSeries) -> None:
                        for _, row in subset.iterrows():
                            y_val = float(row[y_field])
                            if x_mode == "value":
                                line_series.append(float(row[x_col]), y_val)
                            elif x_mode == "datetime":
                                ts = pd.Timestamp(row["_x_dt"])  # type: ignore[index]
                                ms = int(ts.value / 1_000_000)
                                line_series.append(ms, y_val)
                            else:
                                key_val = str(row[x_display_col])
                                line_series.append(category_map.get(key_val, 0), y_val)

                    if chart_type == "Area":
                        line = QLineSeries()
                        line.setName(series_name)
                        append_points(line)
                        if line.count() < 2:
                            chart.addSeries(line)
                        else:
                            area = QAreaSeries(line)
                            area.setName(series_name)
                            chart.addSeries(area)
                    else:
                        line = QLineSeries()
                        line.setName(series_name)
                        append_points(line)
                        chart.addSeries(line)

                if x_mode == "datetime":
                    axis_x = QDateTimeAxis()
                    axis_x.setFormat("yyyy-MM-dd")
                    axis_x.setLabelsAngle(-45)
                    axis_x.setLabelsFont(self.axis_label_font())
                    if parsed_dates is not None and not parsed_dates.dropna().empty:
                        min_ms = int(parsed_dates.min().value / 1_000_000)
                        max_ms = int(parsed_dates.max().value / 1_000_000)
                        axis_x.setRange(
                            QDateTime.fromMSecsSinceEpoch(min_ms),
                            QDateTime.fromMSecsSinceEpoch(max_ms),
                        )
                elif x_mode == "category":
                    axis_x = QBarCategoryAxis()
                    axis_x.append(self.format_axis_labels(categories))
                    axis_x.setLabelsAngle(-45)
                    axis_x.setLabelsFont(self.axis_label_font())
                    if categories:
                        axis_x.setRange(categories[0], categories[-1])
                else:
                    axis_x = QValueAxis()
                    if not data[x_col].dropna().empty:
                        axis_x.setRange(float(data[x_col].min()), float(data[x_col].max()))

                axis_y = QValueAxis()
                if not data[y_field].dropna().empty:
                    y_min = float(data[y_field].min())
                    y_max = float(data[y_field].max())
                    axis_y.setRange(y_min, y_max)
                chart.addAxis(axis_x, Qt.AlignBottom)
                chart.addAxis(axis_y, Qt.AlignLeft)
                for s in chart.series():
                    s.attachAxis(axis_x)
                    s.attachAxis(axis_y)

            chart.setTitle(cfg.title)
            return chart, None
        except Exception as exc:
            return None, f"圖表發生錯誤: {exc}"

    def _try_parse_datetime(self, series: pd.Series) -> Optional[pd.Series]:
        if pd.api.types.is_datetime64_any_dtype(series):
            return series
        parsed = pd.to_datetime(series, errors="coerce")
        if parsed.notna().mean() >= 0.6:
            return parsed
        return None

    def format_axis_labels(self, labels: List[str], max_len: int = 16) -> List[str]:
        full = getattr(self, "full_labels_checkbox", None)
        use_full = bool(full and full.isChecked())
        formatted: List[str] = []
        seen: Dict[str, int] = {}
        for label in labels:
            text = label.replace(" / ", "\n")
            raw = text.replace("\n", " ")
            if not use_full and len(raw) > max_len:
                text = raw[: max_len - 1] + "…"
            if text in seen:
                seen[text] += 1
                text = f"{text}({seen[text]})"
            else:
                seen[text] = 1
            formatted.append(text)
        return formatted

    def axis_label_font(self) -> QFont:
        full = getattr(self, "full_labels_checkbox", None)
        size = 7 if (full and full.isChecked()) else 8
        return QFont("", size)

    def build_pivot_chart(
        self,
        pivot: Union[pd.DataFrame, pd.Series],
        rows: List[str],
        cols: List[str],
        chart_type: str,
    ) -> QChart:
        chart = QChart()
        chart.setMargins(QMargins(10, 10, 10, 30))

        if isinstance(pivot, pd.Series):
            data_df = pivot.to_frame(name="value")
        else:
            data_df = pivot.copy()

        if isinstance(data_df.columns, pd.MultiIndex):
            data_df.columns = [
                " / ".join([str(c) for c in col if c != ""]) for col in data_df.columns
            ]

        if data_df.index.nlevels > 1:
            data_df.index = [" / ".join([str(v) for v in idx]) for idx in data_df.index]

        data_df = data_df.fillna(0)
        if data_df.empty:
            return chart

        categories = [str(idx) for idx in data_df.index.tolist()]
        display_categories = self.format_axis_labels(categories)

        axis_x = QBarCategoryAxis()
        axis_x.append(display_categories)
        axis_x.setLabelsAngle(-45)
        axis_x.setLabelsFont(self.axis_label_font())
        axis_y = QValueAxis()

        if chart_type == "折線":
            max_val = 0.0
            for col in data_df.columns:
                line = QLineSeries()
                line.setName(str(col))
                for idx, value in enumerate(data_df[col].tolist()):
                    line.append(float(idx), float(value))
                chart.addSeries(line)
                max_val = max(max_val, float(np.nanmax(data_df[col].values)))
            chart.addAxis(axis_x, Qt.AlignBottom)
            chart.addAxis(axis_y, Qt.AlignLeft)
            for s in chart.series():
                s.attachAxis(axis_x)
                s.attachAxis(axis_y)
            axis_y.setRange(0, max(max_val * 1.1, 1))
        else:
            bar_series = QBarSeries()
            for col in data_df.columns:
                bar_set = QBarSet(str(col))
                values = [float(v) for v in data_df[col].tolist()]
                bar_set.append(values)
                bar_series.append(bar_set)

            chart.addSeries(bar_series)
            chart.addAxis(axis_x, Qt.AlignBottom)
            chart.addAxis(axis_y, Qt.AlignLeft)
            bar_series.attachAxis(axis_x)
            bar_series.attachAxis(axis_y)

            max_val = float(np.nanmax(data_df.values))
            axis_y.setRange(0, max(max_val * 1.1, 1))

            if chart_type == "長條+折線":
                if data_df.shape[1] > 1:
                    line_values = data_df.sum(axis=1).tolist()
                    line_name = "總和"
                else:
                    line_values = data_df.iloc[:, 0].tolist()
                    line_name = str(data_df.columns[0])
                line = QLineSeries()
                line.setName(line_name)
                for idx, value in enumerate(line_values):
                    line.append(float(idx), float(value))
                chart.addSeries(line)
                line.attachAxis(axis_x)
                line.attachAxis(axis_y)

        title_parts = []
        if rows:
            title_parts.append("列: " + " / ".join(rows))
        if cols:
            title_parts.append("欄: " + " / ".join(cols))
        chart.setTitle("樞紐圖表" if not title_parts else "樞紐圖表 (" + ", ".join(title_parts) + ")")
        return chart


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
