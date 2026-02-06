from __future__ import annotations

import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QDate, QDateTime, QModelIndex, Qt
from PySide6.QtGui import QAction, QPainter
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDoubleSpinBox,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
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

        self._build_ui()
        self._load_initial_data()

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)

        splitter = QSplitter(Qt.Horizontal)
        root_layout.addWidget(splitter)

        self._build_left_panel(splitter)
        self._build_right_panel(splitter)

        self.setCentralWidget(root)

        refresh_action = QAction("重新整理", self)
        refresh_action.triggered.connect(self.refresh_preview)
        self.addAction(refresh_action)

    def _build_left_panel(self, parent: QSplitter) -> None:
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(8, 8, 8, 8)

        left_layout.addWidget(self._build_data_source_group())
        left_layout.addWidget(self._build_filter_group())
        left_layout.addWidget(self._build_chart_builder_group())
        left_layout.addStretch(1)

        parent.addWidget(left_widget)
        parent.setStretchFactor(0, 0)

    def _build_right_panel(self, parent: QSplitter) -> None:
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(8, 8, 8, 8)

        self.tabs = QTabWidget()
        right_layout.addWidget(self.tabs)

        self._build_preview_tab()
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

        layout.addLayout(form)

        button_row = QHBoxLayout()
        self.load_button = QPushButton("從路徑更新")
        self.load_button.clicked.connect(self.load_from_path)
        button_row.addWidget(self.load_button)

        self.load_latest_button = QPushButton("載入最新")
        self.load_latest_button.clicked.connect(self.load_latest)
        button_row.addWidget(self.load_latest_button)

        layout.addLayout(button_row)

        self.source_status = QLabel("尚未載入資料")
        self.source_status.setWordWrap(True)
        layout.addWidget(self.source_status)

        return group

    def _build_filter_group(self) -> QGroupBox:
        group = QGroupBox("篩選器")
        layout = QVBoxLayout(group)

        layout.addWidget(QLabel("選擇要啟用篩選的欄位"))
        self.filter_column_list = QListWidget()
        self.filter_column_list.itemChanged.connect(self.rebuild_filter_widgets)
        layout.addWidget(self.filter_column_list)

        self.filter_area = QScrollArea()
        self.filter_area.setWidgetResizable(True)
        self.filter_area_widget = QWidget()
        self.filter_area_layout = QVBoxLayout(self.filter_area_widget)
        self.filter_area_layout.addStretch(1)
        self.filter_area.setWidget(self.filter_area_widget)
        layout.addWidget(self.filter_area)

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

        return group

    def _build_chart_builder_group(self) -> QGroupBox:
        group = QGroupBox("欄位架 / 圖表")
        layout = QVBoxLayout(group)

        form = QFormLayout()
        self.chart_type_combo = QComboBox()
        self.chart_type_combo.addItems(["Bar", "Line", "Area", "Scatter", "Histogram", "Box"])
        self.chart_type_combo.currentTextChanged.connect(self.update_builder_options)
        form.addRow("圖表類型", self.chart_type_combo)

        self.x_combo = QComboBox()
        form.addRow("X 軸", self.x_combo)

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

        self.add_chart_button = QPushButton("加入儀表板")
        self.add_chart_button.clicked.connect(self.add_chart)
        layout.addWidget(self.add_chart_button)

        self.builder_hint = QLabel("")
        self.builder_hint.setWordWrap(True)
        layout.addWidget(self.builder_hint)

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

        df = load_latest_df()
        if df is not None:
            self.set_data(df)
            self.source_status.setText("已載入 data/latest.csv")

    def select_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "選擇 CSV", "", "CSV Files (*.csv)")
        if file_path:
            self.path_input.setText(file_path)

    def load_from_path(self) -> None:
        path = self.path_input.text().strip()
        if not path:
            QMessageBox.warning(self, "CSV", "請輸入檔案路徑")
            return

        cfg = load_config()
        cfg["source_path"] = path
        cfg["encoding"] = self.encoding_combo.currentText()
        cfg["keep_history"] = self.keep_history_checkbox.isChecked()
        save_config(cfg)

        df, err = safe_update(path, cfg["encoding"], "manual")
        if err:
            QMessageBox.critical(self, "更新失敗", err)
            return
        if df is None:
            QMessageBox.critical(self, "更新失敗", "讀取失敗")
            return
        self.set_data(df)
        self.source_status.setText(f"已更新: {path}")

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
        self.apply_filters()

    def _refresh_column_lists(self) -> None:
        self.filter_column_list.blockSignals(True)
        self.filter_column_list.clear()
        if self.df is not None:
            for col in self.df.columns:
                item = QListWidgetItem(col)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                self.filter_column_list.addItem(item)
        self.filter_column_list.blockSignals(False)

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
            group_layout = QVBoxLayout(group)

            if pd.api.types.is_numeric_dtype(series):
                min_val = float(series.min()) if series.notna().any() else 0.0
                max_val = float(series.max()) if series.notna().any() else 0.0

                min_spin = QDoubleSpinBox()
                min_spin.setRange(-1e18, 1e18)
                min_spin.setValue(min_val)
                max_spin = QDoubleSpinBox()
                max_spin.setRange(-1e18, 1e18)
                max_spin.setValue(max_val)

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
                    for value in sorted(values, key=lambda x: str(x)):
                        item = QListWidgetItem(str(value))
                        item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                        item.setCheckState(Qt.Checked)
                        item.setData(Qt.UserRole, value)
                        list_widget.addItem(item)
                    list_widget.setMaximumHeight(160)
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

        model = self.preview_table.model()
        if isinstance(model, DataFrameModel):
            model.set_df(df.head(1000))

        self.refresh_preview()
        self.refresh_dashboard()

    def reset_filters(self) -> None:
        for i in range(self.filter_column_list.count()):
            item = self.filter_column_list.item(i)
            item.setCheckState(Qt.Unchecked)
        self.rebuild_filter_widgets()
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

        if chart_type == "Histogram":
            set_combo(self.x_combo, numeric_cols)
            set_combo(self.y_combo, [])
        elif chart_type == "Scatter":
            set_combo(self.x_combo, numeric_cols)
            set_combo(self.y_combo, numeric_cols)
        else:
            set_combo(self.x_combo, all_cols)
            y_options = ["(count)"] + numeric_cols
            set_combo(self.y_combo, y_options)

        color_options = ["(none)"] + all_cols
        set_combo(self.color_combo, color_options)

        size_options = ["(none)"] + numeric_cols
        set_combo(self.size_combo, size_options)

        self.update_default_title()
        self.refresh_preview()

    def update_default_title(self) -> None:
        chart_type = self.chart_type_combo.currentText()
        x_col = self.x_combo.currentText()
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

        return ChartConfig(
            chart_type=self.chart_type_combo.currentText(),
            x_col=self.x_combo.currentText() or None,
            y_col=y_col,
            color_col=color_col,
            size_col=size_col,
            agg=self.agg_combo.currentText(),
            top_n=top_n,
            bins=self.bin_spin.value(),
            title=self.title_input.text().strip() or "Chart",
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

            chart, err = self.build_chart(self.filtered_df, cfg)
            if err:
                frame_layout.addWidget(QLabel(err))
            else:
                view = QChartView(chart)
                view.setMinimumHeight(240)
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
        if df.empty:
            return None, "篩選後沒有資料"

        chart_type = cfg.chart_type
        x_col = cfg.x_col
        y_col = cfg.y_col
        color_col = cfg.color_col

        if chart_type == "Histogram":
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
            chart.addAxis(axis_x, Qt.AlignBottom)
            series.attachAxis(axis_x)
            axis_y = QValueAxis()
            max_count = int(counts.max()) if len(counts) else 1
            axis_y.setRange(0, max(max_count * 1.1, 1))
            chart.addAxis(axis_y, Qt.AlignLeft)
            series.attachAxis(axis_y)
            chart.setTitle(cfg.title)
            return chart, None

        if chart_type == "Scatter":
            if not x_col or not y_col:
                return None, "請選擇 X/Y"
            if not pd.api.types.is_numeric_dtype(df[x_col]) or not pd.api.types.is_numeric_dtype(
                df[y_col]
            ):
                return None, "散點圖需要數值欄位"

            chart = QChart()
            groups = [None]
            if color_col:
                groups = df[color_col].dropna().unique().tolist()

            for group in groups:
                series = QScatterSeries()
                name = "全部" if group is None else str(group)
                series.setName(name)
                subset = df if group is None else df[df[color_col] == group]
                for _, row in subset[[x_col, y_col]].dropna().iterrows():
                    series.append(float(row[x_col]), float(row[y_col]))
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
            if not x_col or not y_col:
                return None, "箱型圖需要 X 與 Y"
            if not pd.api.types.is_numeric_dtype(df[y_col]):
                return None, "箱型圖需要數值欄位"

            chart = QChart()
            series = QBoxPlotSeries()
            for key, group in df[[x_col, y_col]].dropna().groupby(x_col):
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

        if not x_col:
            return None, "請選擇 X 軸"

        group_cols = [x_col]
        if color_col:
            group_cols.append(color_col)

        if cfg.agg == "none":
            if y_col is None:
                return None, "請選擇度量欄位"
            data = df[group_cols + [y_col]].dropna()
            y_field = y_col
        else:
            if y_col is None or cfg.agg == "count":
                data = df.groupby(group_cols, dropna=False).size().reset_index(name="count")
                y_field = "count"
            else:
                data = df.groupby(group_cols, dropna=False)[y_col].agg(cfg.agg).reset_index()
                y_field = y_col

        if cfg.top_n:
            data = data.sort_values(y_field, ascending=False).head(cfg.top_n)

        chart = QChart()
        if chart_type == "Bar":
            categories = data[x_col].astype(str).unique().tolist()
            series = QBarSeries()
            if color_col:
                for key in data[color_col].astype(str).unique().tolist():
                    subset = data[data[color_col].astype(str) == key]
                    values = []
                    for cat in categories:
                        value = subset[subset[x_col].astype(str) == cat][y_field]
                        values.append(float(value.iloc[0]) if not value.empty else 0.0)
                    bar_set = QBarSet(key)
                    bar_set.append(values)
                    series.append(bar_set)
            else:
                values = [float(v) for v in data.set_index(x_col)[y_field].reindex(categories).fillna(0)]
                bar_set = QBarSet(y_field)
                bar_set.append(values)
                series.append(bar_set)

            chart.addSeries(series)
            axis_x = QBarCategoryAxis()
            axis_x.append(categories)
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
                subset = subset.sort_values(x_col)
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
                            key_val = str(row[x_col])
                            line_series.append(category_map.get(key_val, 0), y_val)

                if chart_type == "Area":
                    line = QLineSeries()
                    line.setName(series_name)
                    append_points(line)
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
                if parsed_dates is not None and not parsed_dates.dropna().empty:
                    min_ms = int(parsed_dates.min().value / 1_000_000)
                    max_ms = int(parsed_dates.max().value / 1_000_000)
                    axis_x.setRange(QDateTime.fromMSecsSinceEpoch(min_ms), QDateTime.fromMSecsSinceEpoch(max_ms))
            elif x_mode == "category":
                axis_x = QBarCategoryAxis()
                axis_x.append(categories)
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

    def _try_parse_datetime(self, series: pd.Series) -> Optional[pd.Series]:
        if pd.api.types.is_datetime64_any_dtype(series):
            return series
        parsed = pd.to_datetime(series, errors="coerce")
        if parsed.notna().mean() >= 0.6:
            return parsed
        return None


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
