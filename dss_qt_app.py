from __future__ import annotations

import csv
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from fnmatch import fnmatch
from dataclasses import dataclass, replace
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Iterable

from openpyxl import Workbook

import dss_hours_tracker as core

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, QPoint, Qt, QTimer, QUrl, Signal, QItemSelectionModel
from PySide6.QtGui import QAction, QColor, QDesktopServices, QFont, QIcon, QKeySequence, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QBoxLayout,
    QCheckBox,
    QColorDialog,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QProgressBar,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QStyledItemDelegate,
    QTabWidget,
    QTableView,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)


TABLE_DATE_COLUMNS = {"date", "work_date", "week_start", "week_end"}
TABLE_NUMERIC_COLUMNS = {"st", "ot", "dt", "total", "expanded", "days", "similarity", "limit", "actual_total", "delta"}


def _best_contrast_text(hex_color: str) -> str:
    normalized = core.normalize_ui_hex_color(hex_color)
    if not normalized:
        return "#111827"
    body = normalized[1:]
    r = int(body[0:2], 16)
    g = int(body[2:4], 16)
    b = int(body[4:6], 16)
    luminance = (0.299 * r) + (0.587 * g) + (0.114 * b)
    return "#111827" if luminance >= 160 else "#f8fafc"


def _configure_forced_qt_software_rendering() -> None:
    # Keep Qt off the native GPU path so machines with flaky theme/driver combos
    # render the same widget chrome as our known-good environments.
    os.environ["QT_OPENGL"] = "software"
    os.environ["QSG_RHI_BACKEND"] = "software"
    try:
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseSoftwareOpenGL, True)
    except Exception:
        pass


def _build_qt_chrome_stylesheet(theme: core.UiThemeColors) -> str:
    window_bg = theme.window_background
    panel_bg = theme.panel_background
    content_bg = theme.content_chrome_background
    control_bg = theme.control_background
    control_fg = theme.control_foreground
    border = theme.control_border
    stronger_border = core.lighten_hex_color(border, -26)
    hover_bg = theme.control_hover_background
    pressed_bg = theme.control_pressed_background
    disabled_bg = theme.control_disabled_background
    disabled_text = theme.control_disabled_foreground
    table_bg = theme.table_background
    table_header_bg = theme.header_background
    table_header_fg = theme.header_foreground
    text = control_fg
    muted_text = theme.tab_inactive_foreground
    accent = theme.button_primary_background
    accent_text = theme.button_primary_foreground
    danger = theme.button_danger_background
    danger_text = theme.button_danger_foreground
    tab_inactive_bg = theme.tab_inactive_background
    tab_active_bg = theme.tab_active_background
    selection_bg = theme.selection_background
    selection_fg = theme.selection_foreground

    return f"""
        QMainWindow, QWidget {{
            background-color: {window_bg};
            color: {text};
        }}
        QWidget#centralwidget {{
            background-color: {window_bg};
        }}
        QLabel {{
            color: {text};
            background: transparent;
        }}
        QGroupBox {{
            color: {text};
            border: 1px solid {border};
            border-radius: 10px;
            margin-top: 12px;
            padding: 10px;
            padding-top: 14px;
            background-color: {panel_bg};
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 6px;
            color: {text};
            background-color: {panel_bg};
        }}
        QPushButton, QToolButton, QComboBox, QLineEdit, QSpinBox, QPlainTextEdit, QTextEdit, QListWidget, QMenu, QScrollArea {{
            background-color: {control_bg};
            color: {text};
            border: 1px solid {border};
            border-radius: 8px;
        }}
        QPushButton, QToolButton {{
            padding: 6px 12px;
            background-color: {control_bg};
            color: {text};
        }}
        QPushButton:hover, QToolButton:hover, QComboBox:hover, QLineEdit:hover, QSpinBox:hover {{
            background-color: {hover_bg};
            border-color: {stronger_border};
        }}
        QPushButton:pressed, QToolButton:pressed {{
            background-color: {pressed_bg};
        }}
        QPushButton:disabled, QToolButton:disabled, QLabel:disabled, QComboBox:disabled, QLineEdit:disabled, QSpinBox:disabled {{
            color: {disabled_text};
            border-color: {border};
            background-color: {disabled_bg};
        }}
        QPushButton[text="Add DSSs"],
        QPushButton[text="Quick Add"],
        QPushButton[text="Refresh"],
        QPushButton[text="Export View"],
        QPushButton[text="Apply Settings"],
        QPushButton[text="Create Outlook Drafts"],
        QPushButton[text="Sync Outlook Emails"],
        QPushButton[text="Save Templates"],
        QPushButton[text="Save Email"] {{
            background-color: {accent};
            color: {accent_text};
            border-color: {accent};
        }}
        QPushButton[text="Add DSSs"]:hover,
        QPushButton[text="Quick Add"]:hover,
        QPushButton[text="Refresh"]:hover,
        QPushButton[text="Export View"]:hover,
        QPushButton[text="Apply Settings"]:hover,
        QPushButton[text="Create Outlook Drafts"]:hover,
        QPushButton[text="Sync Outlook Emails"]:hover,
        QPushButton[text="Save Templates"]:hover,
        QPushButton[text="Save Email"]:hover {{
            background-color: {core.lighten_hex_color(accent, 12)};
            border-color: {core.lighten_hex_color(accent, -10)};
        }}
        QPushButton[text="Reset All Settings to Default"],
        QPushButton[text="Clear Cached DSSs"],
        QPushButton[text="Clear Stored Emails"],
        QPushButton[text="Clear All Stored Data"],
        QPushButton[text="Clear All"] {{
            background-color: {danger};
            color: {danger_text};
            border-color: {danger};
        }}
        QComboBox::drop-down {{
            border: 0;
            width: 20px;
        }}
        QComboBox QAbstractItemView, QListWidget {{
            background-color: {control_bg};
            color: {text};
            selection-background-color: {selection_bg};
            selection-color: {selection_fg};
        }}
        QTabWidget::pane {{
            border: 1px solid {border};
            border-radius: 12px;
            top: -1px;
            background-color: {content_bg};
        }}
        QTabBar::tab {{
            background-color: {tab_inactive_bg};
            color: {muted_text};
            border: 1px solid {border};
            border-bottom-color: {border};
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            padding: 8px 14px;
            margin-right: 2px;
        }}
        QTabBar::tab:selected {{
            background-color: {tab_active_bg};
            color: {text};
            border-color: {stronger_border};
            border-bottom-color: {tab_active_bg};
        }}
        QTabBar::tab:hover:!selected {{
            background-color: {hover_bg};
            color: {text};
        }}
        QHeaderView::section {{
            background-color: {table_header_bg};
            color: {table_header_fg};
            border: 1px solid {border};
            padding: 6px;
        }}
        QTableView {{
            background-color: {table_bg};
            color: {text};
            gridline-color: {border};
            alternate-background-color: {table_bg};
            border: 1px solid {border};
        }}
        QTableCornerButton::section {{
            background-color: {table_header_bg};
            border: 1px solid {border};
        }}
        QProgressBar {{
            background-color: {panel_bg};
            color: {text};
            border: 1px solid {border};
            border-radius: 5px;
            text-align: center;
        }}
        QProgressBar::chunk {{
            background-color: {accent};
            border-radius: 4px;
        }}
        QToolTip {{
            background-color: {theme.tooltip_background};
            color: {theme.tooltip_foreground};
            border: 1px solid {border};
        }}
    """


def _build_forced_qt_palette(theme: core.UiThemeColors) -> QPalette:
    palette = QPalette()
    content_bg = QColor(theme.content_chrome_background)
    panel_bg = QColor(theme.panel_background)
    table_bg = QColor(theme.table_background)
    border = QColor(theme.control_border)
    text = QColor(theme.control_foreground)
    muted_text = QColor(theme.tab_inactive_foreground)
    disabled_text = QColor(theme.control_disabled_foreground)
    white = QColor(theme.selection_foreground)
    accent = QColor(theme.selection_background)
    accent_soft = QColor(theme.control_disabled_background)

    palette.setColor(QPalette.Window, QColor(theme.window_background))
    palette.setColor(QPalette.WindowText, text)
    palette.setColor(QPalette.Base, QColor(theme.control_background))
    palette.setColor(QPalette.AlternateBase, table_bg)
    palette.setColor(QPalette.ToolTipBase, QColor(theme.tooltip_background))
    palette.setColor(QPalette.ToolTipText, QColor(theme.tooltip_foreground))
    palette.setColor(QPalette.Text, text)
    palette.setColor(QPalette.Button, panel_bg)
    palette.setColor(QPalette.ButtonText, text)
    palette.setColor(QPalette.BrightText, white)
    palette.setColor(QPalette.Highlight, accent)
    palette.setColor(QPalette.HighlightedText, white)
    palette.setColor(QPalette.Light, white)
    palette.setColor(QPalette.Midlight, border)
    palette.setColor(QPalette.Dark, border)
    palette.setColor(QPalette.Mid, border)
    palette.setColor(QPalette.Shadow, QColor("#cbd5e1"))
    palette.setColor(QPalette.Link, accent)
    palette.setColor(QPalette.LinkVisited, QColor("#1d4ed8"))

    palette.setColor(QPalette.Disabled, QPalette.WindowText, disabled_text)
    palette.setColor(QPalette.Disabled, QPalette.Text, disabled_text)
    palette.setColor(QPalette.Disabled, QPalette.ButtonText, disabled_text)
    palette.setColor(QPalette.Disabled, QPalette.Highlight, accent_soft)
    palette.setColor(QPalette.Disabled, QPalette.HighlightedText, muted_text)
    return palette


@dataclass(frozen=True)
class TableSpec:
    table_id: str
    title: str
    columns: tuple[tuple[str, str], ...]


class TableModel(QAbstractTableModel):
    def __init__(
        self,
        columns: list[tuple[str, str]],
        rows: list[dict[str, Any]] | None = None,
        theme: core.UiThemeColors | None = None,
        table_id: str = "",
    ) -> None:
        super().__init__()
        self.columns = columns
        self.rows: list[dict[str, Any]] = rows or []
        self.theme = theme or core.DEFAULT_UI_THEME
        self.table_id = table_id

    def set_rows(self, rows: list[dict[str, Any]]) -> None:
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.columns)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> Any:
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal and 0 <= section < len(self.columns):
            return self.columns[section][1]
        return str(section + 1)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None
        row = self.rows[index.row()]
        key, _label = self.columns[index.column()]
        value = row.get(key, "")
        if role == Qt.DisplayRole:
            if key == "suppressed":
                return "\u2611" if bool(value) else "\u2610"
            return value
        tags = set(row.get("__tags__", ()))
        if role == Qt.BackgroundRole:
            if "missing_email" in tags:
                return QColor(core.MISSING_EMAIL_ROW_BACKGROUND)
            if "alert" in tags:
                return QColor(self.theme.alert_row_background)
            if "crew_total" in tags:
                return QColor(self.theme.crew_total_background)
        if role == Qt.ForegroundRole:
            if "suppressed" in tags:
                return QColor("#7a8699")
            if "missing_email" in tags:
                return QColor(core.MISSING_EMAIL_ROW_FOREGROUND)
            if "alert" in tags:
                return QColor(self.theme.alert_row_foreground)
            if "crew_total" in tags:
                return QColor(self.theme.crew_total_foreground)
        if role == Qt.TextAlignmentRole and key == "suppressed":
            return int(Qt.AlignHCenter | Qt.AlignVCenter)
        if role == Qt.TextAlignmentRole and key in TABLE_NUMERIC_COLUMNS:
            return int(Qt.AlignRight | Qt.AlignVCenter)
        if role == Qt.UserRole:
            return row
        return None

    def sort(self, column: int, order: Qt.SortOrder = Qt.AscendingOrder) -> None:
        if not (0 <= column < len(self.columns)):
            return
        key = self.columns[column][0]
        reverse = order == Qt.DescendingOrder

        if self.table_id == "employee_daily_pf":
            self.layoutAboutToBeChanged.emit()
            self.rows.sort(key=lambda row: self._employee_daily_pf_sort_key(row, key, reverse))
            self.layoutChanged.emit()
            return

        def sort_key(row: dict[str, Any]) -> Any:
            value = row.get(key, "")
            if key in TABLE_NUMERIC_COLUMNS:
                try:
                    return float(value)
                except (TypeError, ValueError):
                    return float("-inf")
            if key in TABLE_DATE_COLUMNS:
                text = str(value).strip()
                try:
                    return datetime.strptime(text, "%Y-%m-%d").date()
                except ValueError:
                    return date.min
            if key == "pf_number":
                return core.pf_number_sort_key(str(value))
            return str(value).casefold()

        self.layoutAboutToBeChanged.emit()
        self.rows.sort(key=sort_key, reverse=reverse)
        self.layoutChanged.emit()

    def _employee_daily_pf_sort_key(self, row: dict[str, Any], key: str, reverse: bool) -> Any:
        employee_key = str(row.get("__employee_full__", row.get("employee", ""))).casefold()
        week_start = self._row_hidden_date(row.get("__week_start__", ""))
        work_date = self._row_hidden_date(row.get("__date_sort__", ""))
        pf_key = core.pf_number_sort_key(str(row.get("pf_number", "")))
        source_index = int(row.get("__source_index__", 0))

        if key == "employee":
            return (self._descending_text_key(employee_key) if reverse else employee_key, week_start, work_date, pf_key, source_index)
        if key == "week_number":
            week_sort = self._descending_date_key(week_start) if reverse else week_start
            return (employee_key, week_sort, work_date, pf_key, source_index)
        if key == "date":
            date_sort = self._descending_date_key(work_date) if reverse else work_date
            return (employee_key, date_sort, week_start, pf_key, source_index)
        if key == "pf_number":
            pf_sort = self._descending_pf_key(pf_key) if reverse else pf_key
            return (employee_key, week_start, work_date, pf_sort, source_index)
        if key in TABLE_NUMERIC_COLUMNS:
            try:
                numeric_value = float(row.get(key, ""))
            except (TypeError, ValueError):
                numeric_value = float("-inf")
            return (employee_key, week_start, work_date, -numeric_value if reverse else numeric_value, pf_key, source_index)
        value = str(row.get(key, "")).casefold()
        value_sort = self._descending_text_key(value) if reverse else value
        return (employee_key, week_start, work_date, value_sort, pf_key, source_index)

    @staticmethod
    def _row_hidden_date(value: Any) -> date:
        text = str(value).strip()
        if not text:
            return date.min
        try:
            return datetime.strptime(text, "%Y-%m-%d").date()
        except ValueError:
            return date.min

    @staticmethod
    def _descending_date_key(value: date) -> int:
        return -value.toordinal()

    @staticmethod
    def _descending_text_key(value: str) -> tuple[int, ...]:
        return tuple(-ord(char) for char in value)

    @staticmethod
    def _descending_pf_key(value: tuple[object, ...]) -> tuple[object, ...]:
        transformed: list[object] = []
        for token in value:
            if isinstance(token, (int, float)):
                transformed.append(-token)
            elif isinstance(token, tuple):
                transformed.append(TableModel._descending_pf_key(token))
            elif isinstance(token, str):
                transformed.append(TableModel._descending_text_key(token))
            else:
                transformed.append(token)
        return tuple(transformed)


class RowAccentDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index) -> None:  # type: ignore[override]
        super().paint(painter, option, index)
        row_payload = index.data(Qt.UserRole) or {}
        tags = set(row_payload.get("__tags__", ()))
        if "employee_break" not in tags and "date_break" not in tags:
            return
        painter.save()
        pen = painter.pen()
        model = index.model()
        theme = getattr(model, "theme", core.DEFAULT_UI_THEME)
        if "employee_break" in tags:
            pen.setColor(QColor(theme.employee_divider))
            pen.setWidth(3)
        else:
            pen.setColor(QColor(theme.date_divider))
            pen.setWidth(2)
        painter.setPen(pen)
        painter.drawLine(option.rect.topLeft(), option.rect.topRight())
        painter.restore()


class DataTablePage(QWidget):
    layoutChanged = Signal(str, list, dict, str, bool)
    selectionChanged = Signal(str, list)

    def __init__(
        self,
        spec: TableSpec,
        theme: core.UiThemeColors,
        config_path: Path,
        row_activate_callback: Callable[[dict[str, Any], str], None] | None = None,
    ) -> None:
        super().__init__()
        self.spec = spec
        self.config_path = config_path
        self._row_activate_callback = row_activate_callback
        self._context_menu_callback: Callable[[list[dict[str, Any]], QPoint], None] | None = None
        self.columns = list(spec.columns)
        self.model = TableModel(self.columns, theme=theme, table_id=spec.table_id)
        self.view = QTableView(self)
        self.view.setModel(self.model)
        self.view.setItemDelegate(RowAccentDelegate(self.view))
        self.view.setSortingEnabled(True)
        self.view.setAlternatingRowColors(False)
        self.view.setSelectionBehavior(QTableView.SelectItems)
        self.view.setSelectionMode(QTableView.ExtendedSelection)
        self.view.verticalHeader().setVisible(False)
        self.view.horizontalHeader().setSectionsMovable(True)
        self.view.horizontalHeader().setStretchLastSection(False)
        self.view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self._apply_selection_stylesheet(theme)

        self.columns_button = QPushButton("Columns", self)
        self.columns_button.clicked.connect(self._show_columns_menu)

        top = QHBoxLayout()
        top.addStretch(1)
        top.addWidget(self.columns_button)
        self.top_bar = top

        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(self.view, 1)

        self._restoring_layout = False
        self._current_rows: list[dict[str, Any]] = []
        self._apply_saved_layout()
        if self.spec.table_id not in core.load_table_layouts(self.config_path):
            self._apply_default_hidden_columns()
        self.view.horizontalHeader().sortIndicatorChanged.connect(self._emit_layout_change)
        self.view.horizontalHeader().sectionMoved.connect(lambda *_: self._emit_layout_change())
        self.view.horizontalHeader().sectionResized.connect(lambda *_: self._emit_layout_change())
        self.view.doubleClicked.connect(self._on_double_clicked)
        self.view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.view.customContextMenuRequested.connect(self._show_context_menu)
        self.view.selectionModel().selectionChanged.connect(lambda *_: self._emit_selection_change())

        copy_action = QAction(self)
        copy_action.setShortcut(QKeySequence.Copy)
        copy_action.triggered.connect(self.copy_selection)
        self.addAction(copy_action)

    def add_toolbar_button(self, text: str, callback: Callable[[], None]) -> QPushButton:
        button = QPushButton(text, self)
        button.clicked.connect(callback)
        self.top_bar.insertWidget(0, button)
        return button

    def add_toolbar_widget(self, widget: QWidget) -> QWidget:
        self.top_bar.insertWidget(0, widget)
        return widget

    def set_context_menu_callback(self, callback: Callable[[list[dict[str, Any]], QPoint], None] | None) -> None:
        self._context_menu_callback = callback

    def set_theme(self, theme: core.UiThemeColors) -> None:
        self.model.theme = theme
        self._apply_selection_stylesheet(theme)
        self.model.layoutChanged.emit()

    def _apply_selection_stylesheet(self, theme: core.UiThemeColors) -> None:
        selection_bg = theme.selection_background
        selection_fg = theme.selection_foreground
        self.view.setStyleSheet(
            f"QTableView::item:selected {{ background-color: {selection_bg}; color: {selection_fg}; }}"
            f"QTableView::item:selected:active {{ background-color: {selection_bg}; color: {selection_fg}; }}"
            f"QTableView::item:selected:!active {{ background-color: {selection_bg}; color: {selection_fg}; }}"
        )

    def set_rows(self, rows: list[dict[str, Any]]) -> None:
        self._current_rows = rows
        self.model.set_rows(rows)
        sort_index = self.view.horizontalHeader().sortIndicatorSection()
        if 0 <= sort_index < len(self.columns):
            self.model.sort(sort_index, self.view.horizontalHeader().sortIndicatorOrder())
        self._emit_selection_change()

    def selected_rows(self) -> list[dict[str, Any]]:
        indexes = self.view.selectionModel().selectedIndexes()
        results: list[dict[str, Any]] = []
        seen_rows: set[int] = set()
        for index in indexes:
            if index.row() in seen_rows:
                continue
            seen_rows.add(index.row())
            results.append(self.model.rows[index.row()])
        return results

    def export_rows(self) -> tuple[list[str], list[dict[str, Any]]]:
        visible_columns = self.visible_columns()
        return visible_columns, list(self.model.rows)

    def copy_selection(self) -> None:
        indexes = self.view.selectionModel().selectedIndexes()
        if not indexes:
            return
        rows = sorted({index.row() for index in indexes})
        cols = sorted({index.column() for index in indexes})
        cell_text: list[str] = []
        index_map = {(index.row(), index.column()): index for index in indexes}
        for row in rows:
            values: list[str] = []
            for col in cols:
                index = index_map.get((row, col))
                values.append("" if index is None else str(index.data(Qt.DisplayRole) or ""))
            cell_text.append("\t".join(values))
        QApplication.clipboard().setText("\n".join(cell_text))

    def visible_columns(self) -> list[str]:
        columns: list[str] = []
        for logical in range(len(self.columns)):
            if not self.view.isColumnHidden(logical):
                columns.append(self.columns[logical][0])
        return columns

    def _show_columns_menu(self) -> None:
        menu = QMenu(self)
        show_all = QAction("Show All", menu)
        show_all.triggered.connect(self._show_all_columns)
        hide_all = QAction("Hide All", menu)
        hide_all.triggered.connect(self._hide_all_columns)
        menu.addAction(show_all)
        menu.addAction(hide_all)
        menu.addSeparator()
        for idx, (key, label) in enumerate(self.columns):
            action = QAction(label, menu)
            action.setCheckable(True)
            action.setChecked(not self.view.isColumnHidden(idx))
            action.toggled.connect(lambda checked, logical=idx: self._toggle_column(logical, checked))
            menu.addAction(action)
        menu.exec(self.columns_button.mapToGlobal(self.columns_button.rect().bottomLeft()))

    def _toggle_column(self, logical: int, checked: bool) -> None:
        if not checked and len(self.visible_columns()) <= 1:
            return
        self.view.setColumnHidden(logical, not checked)
        self._emit_layout_change()

    def _show_all_columns(self) -> None:
        for idx in range(len(self.columns)):
            self.view.setColumnHidden(idx, False)
        self._emit_layout_change()

    def _hide_all_columns(self) -> None:
        for idx in range(1, len(self.columns)):
            self.view.setColumnHidden(idx, True)
        self.view.setColumnHidden(0, False)
        self._emit_layout_change()

    def _apply_saved_layout(self) -> None:
        layouts = core.load_table_layouts(self.config_path)
        layout = layouts.get(self.spec.table_id, {})
        if not layout:
            return
        self._restoring_layout = True
        visible_columns = layout.get("visible_columns", [])
        if visible_columns:
            visible_set = set(visible_columns)
            for idx, (key, _label) in enumerate(self.columns):
                self.view.setColumnHidden(idx, key not in visible_set)
        widths = layout.get("column_widths", {})
        for idx, (key, _label) in enumerate(self.columns):
            if key in widths:
                self.view.setColumnWidth(idx, int(widths[key]))
        sort_column = layout.get("sort_column", "")
        descending = bool(layout.get("sort_descending", False))
        if sort_column:
            for idx, (key, _label) in enumerate(self.columns):
                if key == sort_column:
                    self.view.sortByColumn(idx, Qt.DescendingOrder if descending else Qt.AscendingOrder)
                    break
        self._restoring_layout = False

    def _apply_default_hidden_columns(self) -> None:
        for idx, (key, _label) in enumerate(self.columns):
            if key == "expanded":
                self.view.setColumnHidden(idx, True)

    def _emit_layout_change(self) -> None:
        if self._restoring_layout:
            return
        widths = {key: int(self.view.columnWidth(idx)) for idx, (key, _label) in enumerate(self.columns)}
        sort_column = ""
        descending = False
        sort_index = self.view.horizontalHeader().sortIndicatorSection()
        if 0 <= sort_index < len(self.columns):
            sort_column = self.columns[sort_index][0]
            descending = self.view.horizontalHeader().sortIndicatorOrder() == Qt.DescendingOrder
        self.layoutChanged.emit(self.spec.table_id, self.visible_columns(), widths, sort_column, descending)

    def _emit_selection_change(self) -> None:
        self.selectionChanged.emit(self.spec.table_id, self.selected_rows())

    def _on_double_clicked(self, index: QModelIndex) -> None:
        if not index.isValid():
            return
        row = self.model.rows[index.row()]
        column_key = self.columns[index.column()][0]
        if self._row_activate_callback is not None:
            self._row_activate_callback(row, column_key)

    def _show_context_menu(self, position: QPoint) -> None:
        if self._context_menu_callback is None:
            return
        index = self.view.indexAt(position)
        if index.isValid():
            selected_rows = {selected.row() for selected in self.view.selectionModel().selectedIndexes()}
            if index.row() not in selected_rows:
                self.view.selectionModel().clearSelection()
                self.view.selectionModel().select(index, QItemSelectionModel.Select | QItemSelectionModel.Rows)
                self.view.setCurrentIndex(index)
        self._context_menu_callback(self.selected_rows(), self.view.viewport().mapToGlobal(position))


class CheckListPopup(QFrame):
    selectionChanged = Signal(list)
    CHECKED_MARK = "\u2611"
    UNCHECKED_MARK = "\u2610"

    def __init__(self, parent: QWidget, all_label: str, clear_label: str) -> None:
        super().__init__(parent, Qt.Popup | Qt.FramelessWindowHint)
        self.all_label = all_label
        self.clear_label = clear_label
        self.values: list[str] = []
        self.selected: set[str] = set()
        self.list_widget = QListWidget(self)
        self.list_widget.setSelectionMode(QListWidget.NoSelection)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.addWidget(self.list_widget)
        self.list_widget.itemClicked.connect(self._on_item_clicked)

    def _display_text(self, label: str, checked: bool) -> str:
        return f"{self.CHECKED_MARK if checked else self.UNCHECKED_MARK} {label}"

    def _add_item(self, label: str, data: str, checked: bool) -> None:
        item = QListWidgetItem(self._display_text(label, checked))
        item.setFlags(Qt.ItemIsEnabled)
        item.setData(Qt.UserRole, data)
        item.setData(Qt.UserRole + 1, label)
        self.list_widget.addItem(item)

    def _refresh_item_labels(self) -> None:
        all_checked = len(self.selected) == len(self.values) and len(self.values) > 0
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            data = str(item.data(Qt.UserRole))
            label = str(item.data(Qt.UserRole + 1))
            checked = all_checked if data == "__all__" else (data != "__clear__" and data in self.selected)
            item.setText(self._display_text(label, checked))

    def set_values(self, values: list[str], selected: Iterable[str]) -> None:
        self.values = list(values)
        self.selected = set(selected)
        self.list_widget.clear()
        self._add_item(self.all_label, "__all__", len(self.selected) == len(self.values) and len(self.values) > 0)
        self._add_item(self.clear_label, "__clear__", False)
        for value in self.values:
            self._add_item(value, value, value in self.selected)

    def _on_item_clicked(self, item: QListWidgetItem) -> None:
        data = str(item.data(Qt.UserRole))
        if data == "__all__":
            self.selected = set(self.values)
        elif data == "__clear__":
            self.selected.clear()
        elif data in self.selected:
            self.selected.remove(data)
        else:
            self.selected.add(data)
        self._refresh_item_labels()
        self.selectionChanged.emit([value for value in self.values if value in self.selected])


class CheckListButton(QPushButton):
    selectionChanged = Signal(list)

    def __init__(self, all_text: str, noun: str, parent: QWidget | None = None) -> None:
        super().__init__(all_text, parent)
        self.all_text = all_text
        self.noun = noun
        self.values: list[str] = []
        self.selected: list[str] = []
        self._all_selected_mode = True
        self.popup = CheckListPopup(self, all_text, "Uncheck All")
        self.popup.selectionChanged.connect(self._set_selection)
        self.clicked.connect(self._show_popup)

    def set_choices(self, values: Iterable[str], selected: Iterable[str] | None = None, force_single: bool = False) -> None:
        requested = list(selected) if selected is not None else list(self.selected)
        preserve_all_mode = self._all_selected_mode or self.text() == self.all_text
        self.values = list(values)
        self.setEnabled(bool(self.values))
        if not self.values:
            self.selected = []
            self._all_selected_mode = True
        elif force_single and len(self.values) == 1:
            self.selected = [self.values[0]]
            self._all_selected_mode = True
            self.setEnabled(False)
        else:
            allowed = set(self.values)
            if preserve_all_mode:
                self.selected = list(self.values)
                self._all_selected_mode = True
            else:
                desired = [value for value in requested if value in allowed]
                if desired:
                    self.selected = desired
                    self._all_selected_mode = len(self.selected) == len(self.values)
                else:
                    self.selected = list(self.values)
                    self._all_selected_mode = True
        self._update_text()

    def selected_values(self) -> list[str]:
        if self._all_selected_mode:
            return list(self.values)
        return list(self.selected)

    def _show_popup(self) -> None:
        if not self.values:
            return
        self.popup.set_values(self.values, self.selected)
        self.popup.resize(max(self.width(), 260), 320)
        self.popup.move(self.mapToGlobal(QPoint(0, self.height())))
        self.popup.show()

    def _set_selection(self, values: list[str]) -> None:
        self.selected = values if values else []
        self._all_selected_mode = bool(self.values) and len(self.selected) == len(self.values)
        self._update_text()
        self.selectionChanged.emit(self.selected_values())

    def _update_text(self) -> None:
        if not self.values or len(self.selected) == len(self.values):
            self.setText(self.all_text)
            return
        if not self.selected:
            self.setText(f"No {self.noun}")
            return
        if len(self.selected) == 1:
            self.setText(self.selected[0])
            return
        self.setText(f"{len(self.selected)} {self.noun}")


class QuickDssPickerDialog(QDialog):
    ITEM_KIND_ROLE = Qt.UserRole + 10
    ROOT_PF_ROLE = Qt.UserRole + 11
    FILE_PATH_ROLE = Qt.UserRole

    def __init__(self, parent: QWidget, root_folder: Path, selected_paths: Iterable[Path] | None = None, title: str = "Quick DSS Picker") -> None:
        super().__init__(parent)
        self.root_folder = root_folder
        self.selected_paths = {str(Path(path).resolve()) for path in (selected_paths or [])}
        self._syncing_checks = False
        self._group_children: dict[str, list[QListWidgetItem]] = {}
        self._group_headers: dict[str, QListWidgetItem] = {}
        self.setWindowTitle(title)
        self.resize(980, 620)

        layout = QVBoxLayout(self)
        top = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Search PF, job folder, filename, or path...")
        self.check_all_button = QPushButton("Check All")
        self.uncheck_all_button = QPushButton("Uncheck All")
        self.rescan_button = QPushButton("Rescan")
        top.addWidget(QLabel("Search"))
        top.addWidget(self.search_edit, 1)
        top.addWidget(self.check_all_button)
        top.addWidget(self.uncheck_all_button)
        top.addWidget(self.rescan_button)

        self.summary_label = QLabel("")
        self.list_widget = QListWidget()

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Use Selected")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout.addLayout(top)
        layout.addWidget(self.summary_label)
        layout.addWidget(self.list_widget, 1)
        layout.addWidget(buttons)

        self.search_edit.textChanged.connect(self._apply_filter)
        self.check_all_button.clicked.connect(self._check_all_visible)
        self.uncheck_all_button.clicked.connect(self._uncheck_all_visible)
        self.rescan_button.clicked.connect(self._populate)
        self.list_widget.itemChanged.connect(self._on_item_changed)

        self._populate()

    def _iter_candidate_paths(self) -> list[Path]:
        return core.iter_quick_dss_candidate_paths(self.root_folder)

    def _format_item_label(self, path: Path) -> str:
        pf = core.extract_pf_identifier(path.name)
        if pf and pf != path.stem.strip():
            return f"{pf} | {path.name}"
        return path.name

    def _root_pf_for_path(self, path: Path) -> str:
        pf = core.extract_pf_identifier(path.name).strip()
        if pf.upper().startswith("PF"):
            return pf.split("-", 1)[0]
        return pf or path.stem.strip()

    def _populate(self) -> None:
        previous = set(self.selected_file_paths()) | set(self.selected_paths)
        self.list_widget.clear()
        self._group_children.clear()
        self._group_headers.clear()
        candidates = self._iter_candidate_paths()
        grouped: dict[str, list[Path]] = {}
        for path in candidates:
            grouped.setdefault(self._root_pf_for_path(path), []).append(path)
        self._syncing_checks = True
        try:
            for root_pf in sorted(grouped, key=core.pf_number_sort_key):
                paths = sorted(grouped[root_pf], key=lambda p: core.pf_number_sort_key(core.extract_pf_identifier(p.name)))
                header = QListWidgetItem(root_pf)
                header.setData(self.ITEM_KIND_ROLE, "group")
                header.setData(self.ROOT_PF_ROLE, root_pf)
                header.setFlags(header.flags() | Qt.ItemIsUserCheckable)
                self.list_widget.addItem(header)
                self._group_headers[root_pf] = header
                children: list[QListWidgetItem] = []
                all_checked = True
                for path in paths:
                    item = QListWidgetItem(f"    {self._format_item_label(path)}")
                    item.setData(self.FILE_PATH_ROLE, str(path))
                    item.setData(self.ITEM_KIND_ROLE, "file")
                    item.setData(self.ROOT_PF_ROLE, root_pf)
                    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                    checked = str(path) in previous
                    item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
                    if not checked:
                        all_checked = False
                    self.list_widget.addItem(item)
                    children.append(item)
                self._group_children[root_pf] = children
                header.setCheckState(Qt.Checked if all_checked and children else Qt.Unchecked)
        finally:
            self._syncing_checks = False
        self.summary_label.setText(f"Found {len(candidates)} DSS workbook(s) under {self.root_folder}")
        self._apply_filter()

    def _apply_filter(self) -> None:
        needle = self.search_edit.text().strip().casefold()
        visible = 0
        for root_pf, header in self._group_headers.items():
            header_match = needle in header.text().casefold() if needle else True
            matching_children = 0
            for child in self._group_children.get(root_pf, []):
                child_match = needle in child.text().casefold() if needle else True
                child.setHidden(not child_match)
                if child_match:
                    matching_children += 1
                    visible += 1
            header.setHidden(not header_match and matching_children == 0)
        self.summary_label.setText(f"Showing {visible} DSS workbook(s) from {self.root_folder}")

    def _check_all_visible(self) -> None:
        self._syncing_checks = True
        try:
            for row in range(self.list_widget.count()):
                item = self.list_widget.item(row)
                if item.isHidden():
                    continue
                if item.data(self.ITEM_KIND_ROLE) == "file":
                    item.setCheckState(Qt.Checked)
            for root_pf, header in self._group_headers.items():
                children = self._group_children.get(root_pf, [])
                header.setCheckState(Qt.Checked if children and all(child.checkState() == Qt.Checked for child in children) else Qt.Unchecked)
        finally:
            self._syncing_checks = False

    def _uncheck_all_visible(self) -> None:
        self._syncing_checks = True
        try:
            for row in range(self.list_widget.count()):
                item = self.list_widget.item(row)
                if item.isHidden():
                    continue
                if item.data(self.ITEM_KIND_ROLE) == "file":
                    item.setCheckState(Qt.Unchecked)
            for header in self._group_headers.values():
                header.setCheckState(Qt.Unchecked)
        finally:
            self._syncing_checks = False

    def _on_item_changed(self, item: QListWidgetItem) -> None:
        if self._syncing_checks:
            return
        kind = str(item.data(self.ITEM_KIND_ROLE) or "")
        root_pf = str(item.data(self.ROOT_PF_ROLE) or "")
        if not root_pf:
            return
        self._syncing_checks = True
        try:
            if kind == "group":
                checked = item.checkState() == Qt.Checked
                for child in self._group_children.get(root_pf, []):
                    child.setCheckState(Qt.Checked if checked else Qt.Unchecked)
            elif kind == "file":
                header = self._group_headers.get(root_pf)
                children = self._group_children.get(root_pf, [])
                if header is not None:
                    header.setCheckState(Qt.Checked if children and all(child.checkState() == Qt.Checked for child in children) else Qt.Unchecked)
        finally:
            self._syncing_checks = False

    def selected_file_paths(self) -> list[str]:
        results: list[str] = []
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            if item.data(self.ITEM_KIND_ROLE) != "file":
                continue
            if item.checkState() == Qt.Checked:
                results.append(str(item.data(self.FILE_PATH_ROLE)))
        return results


class LoadWorker(QObject):
    progressChanged = Signal(float, str)
    partialReady = Signal(object, str)
    finished = Signal(object)
    failed = Signal(str)
    cancelled = Signal()

    def __init__(
        self,
        source_paths: list[Path],
        previous_data: core.TrackerData | None,
        cache_dir: Path,
        *,
        max_parallel_parse_workers: int,
        partial_preview_enabled: bool,
        force_reparse: bool = False,
    ) -> None:
        super().__init__()
        self.source_paths = source_paths
        self.previous_data = previous_data
        self.cache_dir = cache_dir
        self.max_parallel_parse_workers = max_parallel_parse_workers
        self.partial_preview_enabled = partial_preview_enabled
        self.force_reparse = force_reparse
        self.cancel_event = threading.Event()

    def cancel(self) -> None:
        self.cancel_event.set()

    def run(self) -> None:
        try:
            data = core.load_tracker_data(
                self.source_paths,
                previous_data=self.previous_data,
                progress_callback=lambda pct, msg: self.progressChanged.emit(float(pct), str(msg)),
                partial_callback=(lambda snapshot, msg: self.partialReady.emit(snapshot, msg)) if self.partial_preview_enabled else None,
                cache_dir=self.cache_dir,
                should_cancel=self.cancel_event.is_set,
                max_parallel_parse_workers=self.max_parallel_parse_workers,
                force_reparse=self.force_reparse,
            )
        except core.OperationCancelled:
            self.cancelled.emit()
            return
        except Exception as exc:  # pragma: no cover - UI surface
            self.failed.emit(str(exc))
            return
        self.finished.emit(data)


class OutlookWorker(QObject):
    finished = Signal(object, object)
    failed = Signal(str)
    cancelled = Signal()
    progressChanged = Signal(int, int, str, object)

    def __init__(self, employee_names: list[str]) -> None:
        super().__init__()
        self.employee_names = employee_names
        self.cancel_event = threading.Event()

    def cancel(self) -> None:
        self.cancel_event.set()

    def run(self) -> None:
        try:
            results, address_book_names = core.query_outlook_emails(
                self.employee_names,
                should_cancel=self.cancel_event.is_set,
                scan_full_address_book=False,
                progress_callback=lambda processed, total, employee, resolution: self.progressChanged.emit(
                    int(processed),
                    int(total),
                    str(employee),
                    resolution,
                ),
            )
        except core.OperationCancelled:
            self.cancelled.emit()
            return
        except Exception as exc:  # pragma: no cover - UI surface
            self.failed.emit(str(exc))
            return
        self.finished.emit(results, address_book_names)


class NameTypoWorker(QObject):
    finished = Signal(object, object)
    failed = Signal(str)
    cancelled = Signal()
    progressChanged = Signal(float, str)

    def __init__(
        self,
        employee_names: list[str],
        daily_records: list[core.DailyRecord],
        employee_emails: dict[str, str],
        missing_email_suppressions: set[str],
        employee_outlook_display_names: dict[str, str],
        outlook_lookup_cache: dict[str, core.OutlookLookupCacheEntry],
    ) -> None:
        super().__init__()
        self.employee_names = list(employee_names)
        self.daily_records = list(daily_records)
        self.employee_emails = dict(employee_emails)
        self.missing_email_suppressions = set(missing_email_suppressions)
        self.employee_outlook_display_names = dict(employee_outlook_display_names)
        self.outlook_lookup_cache = dict(outlook_lookup_cache)
        self.cancel_event = threading.Event()

    def cancel(self) -> None:
        self.cancel_event.set()

    def run(self) -> None:
        try:
            self.progressChanged.emit(0.02, "Preparing name typo review...")
            names_to_check = [
                employee
                for employee in self.employee_names
                if not self.employee_emails.get(employee, "").strip() and employee not in self.missing_email_suppressions
            ]
            self.progressChanged.emit(0.12, f"Checking {len(names_to_check)} unresolved name(s) against local cache...")
            cached_reference_names = [
                employee
                for employee in self.employee_names
                if self.employee_emails.get(employee, "").strip() or self.employee_outlook_display_names.get(employee, "").strip()
            ]
            local_cache_warnings = core.find_cached_employee_name_typos(
                names_to_check,
                cached_reference_names,
                self.daily_records,
            )
            locally_inferred = {warning.employee for warning in local_cache_warnings}
            remaining_names = [employee for employee in names_to_check if employee not in locally_inferred]
            address_book_names: list[str] = []
            skipped_recent_misses: set[str] = set()
            cache_updates = dict(self.outlook_lookup_cache)
            if remaining_names:
                query_names, cached_resolutions, cached_display_names, skipped_recent_misses = core.plan_outlook_query_names(
                    remaining_names,
                    self.employee_emails,
                    self.employee_outlook_display_names,
                    self.outlook_lookup_cache,
                )
                for employee, resolution in cached_resolutions.items():
                    if resolution.email.strip() and not self.employee_emails.get(employee, "").strip():
                        self.employee_emails[employee] = resolution.email.strip()
                    if resolution.display_name.strip() and not self.employee_outlook_display_names.get(employee, "").strip():
                        self.employee_outlook_display_names[employee] = resolution.display_name.strip()
                for employee, display_name in cached_display_names.items():
                    if display_name.strip() and not self.employee_outlook_display_names.get(employee, "").strip():
                        self.employee_outlook_display_names[employee] = display_name.strip()
                if skipped_recent_misses:
                    self.progressChanged.emit(0.2, f"Skipping {len(skipped_recent_misses)} recently checked unresolved name(s)...")
                if query_names:
                    self.progressChanged.emit(0.2, f"Querying Outlook for {len(query_names)} remaining unresolved name(s)...")
                    try:
                        query_results, address_book_names = core.query_outlook_emails(
                            query_names,
                            should_cancel=self.cancel_event.is_set,
                            scan_full_address_book=True,
                            progress_callback=lambda processed, total, employee, _resolution: self.progressChanged.emit(
                                0.2 + (0.5 * (processed / max(total, 1))),
                                f"Checking Outlook: {processed}/{total} ({employee})",
                            ),
                        )
                    except core.OperationCancelled:
                        raise
                    except Exception:
                        address_book_names = []
                        query_results = {}
                    cache_updates = core.update_outlook_lookup_cache(cache_updates, query_names, query_results)
                else:
                    self.progressChanged.emit(0.7, "All unresolved names were satisfied from the local cache.")
            else:
                self.progressChanged.emit(0.7, "All unresolved names were matched against the local cache.")
            effective_names_to_check = [
                employee
                for employee in names_to_check
                if not self.employee_emails.get(employee, "").strip()
            ]
            if self.cancel_event.is_set():
                raise core.OperationCancelled("Cancelled name typo refresh.")

            self.progressChanged.emit(0.8, "Comparing names for likely typos...")
            warnings: list[core.NameTypoWarning] = list(local_cache_warnings)
            if effective_names_to_check:
                warnings.extend(core.find_potential_name_typos(effective_names_to_check, self.employee_names, self.daily_records))
                if address_book_names:
                    warnings.extend(
                        core.find_address_book_name_typos(
                            [employee for employee in remaining_names if employee not in skipped_recent_misses and employee in effective_names_to_check],
                            address_book_names,
                            self.daily_records,
                        )
                    )
            else:
                warnings.extend(core.find_similar_employee_name_pairs(self.employee_names, self.daily_records))
            warnings.extend(
                core.find_outlook_display_name_typos(
                    self.employee_names,
                    self.employee_outlook_display_names,
                    self.daily_records,
                )
            )
        except core.OperationCancelled:
            self.cancelled.emit()
            return
        except Exception as exc:  # pragma: no cover - UI surface
            self.failed.emit(str(exc))
            return
        self.progressChanged.emit(1.0, f"Name typo review complete: {len(warnings)} warning(s)")
        self.finished.emit(warnings, cache_updates)


class EmployeesPage(QWidget):
    changed = Signal()
    syncRequested = Signal()
    mergeRequested = Signal(str, str)

    def __init__(self) -> None:
        super().__init__()
        self.employee_names: list[str] = []
        self.hidden_employee_names: set[str] = set()
        self.employee_emails: dict[str, str] = {}
        self.employee_notes: dict[str, str] = {}
        self.employee_groups: dict[str, list[str]] = {}
        self.missing_email_suppressions: set[str] = set()

        splitter = QSplitter(self)
        left = QWidget()
        left_layout = QVBoxLayout(left)
        roster_box = QGroupBox("Roster")
        roster_layout = QVBoxLayout(roster_box)
        self.employee_list = QListWidget()
        self.employee_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.add_employee_button = QPushButton("Add Employee")
        self.remove_employee_button = QPushButton("Hide Employee")
        self.merge_employee_button = QPushButton("Merge Selected")
        self.sync_button = QPushButton("Sync Outlook Emails")
        self.show_hidden_box = QCheckBox("Show hidden employees")
        self.roster_state_label = QLabel("Manual emails, hidden employees, and merges are saved until you change them.")
        self.roster_state_label.setWordWrap(True)
        roster_layout.addWidget(self.show_hidden_box)
        roster_layout.addWidget(self.employee_list, 1)
        roster_layout.addWidget(self.roster_state_label)

        roster_actions_box = QGroupBox("Roster Actions")
        roster_actions_layout = QGridLayout(roster_actions_box)
        roster_actions_layout.addWidget(self.add_employee_button, 0, 0)
        roster_actions_layout.addWidget(self.remove_employee_button, 0, 1)
        roster_actions_layout.addWidget(self.merge_employee_button, 1, 0)
        roster_actions_layout.addWidget(self.sync_button, 1, 1)
        left_layout.addWidget(roster_box, 1)
        left_layout.addWidget(roster_actions_box)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        detail_box = QGroupBox("Employee Details")
        form = QFormLayout(detail_box)
        self.employee_name_label = QLabel("")
        self.email_edit = QLineEdit()
        self.save_email_button = QPushButton("Save Email")
        self.suppression_box = QCheckBox("Suppress missing email warnings for this employee")
        self.notes_edit = QPlainTextEdit()
        self.group_list = QListWidget()
        self.group_list.setSelectionMode(QListWidget.NoSelection)
        email_row = QWidget()
        email_row_layout = QHBoxLayout(email_row)
        email_row_layout.setContentsMargins(0, 0, 0, 0)
        email_row_layout.addWidget(self.email_edit, 1)
        email_row_layout.addWidget(self.save_email_button)
        form.addRow("Employee", self.employee_name_label)
        form.addRow("Email", email_row)
        form.addRow("", self.suppression_box)
        form.addRow("Notes", self.notes_edit)
        form.addRow("Groups", self.group_list)

        group_box = QGroupBox("Group Membership")
        group_layout = QVBoxLayout(group_box)
        self.groups_list = QListWidget()
        group_buttons = QHBoxLayout()
        self.add_group_button = QPushButton("Add Group")
        self.remove_group_button = QPushButton("Remove Group")
        group_buttons.addWidget(self.add_group_button)
        group_buttons.addWidget(self.remove_group_button)
        group_layout.addWidget(self.groups_list, 1)
        group_layout.addLayout(group_buttons)

        right_layout.addWidget(detail_box, 2)
        right_layout.addWidget(group_box, 1)

        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 4)

        layout = QVBoxLayout(self)
        layout.addWidget(splitter, 1)

        self.employee_list.currentItemChanged.connect(self._populate_details)
        self.employee_list.itemSelectionChanged.connect(self._update_employee_action_states)
        self.groups_list.currentItemChanged.connect(self._group_selected)
        self.add_employee_button.clicked.connect(self._add_employee)
        self.remove_employee_button.clicked.connect(self._remove_employee)
        self.merge_employee_button.clicked.connect(self._merge_selected_employees)
        self.add_group_button.clicked.connect(self._add_group)
        self.remove_group_button.clicked.connect(self._remove_group)
        self.sync_button.clicked.connect(self.syncRequested.emit)
        self.email_edit.editingFinished.connect(self._save_current)
        self.save_email_button.clicked.connect(self._save_current)
        self.suppression_box.toggled.connect(lambda _checked: self._save_current())
        self.notes_edit.textChanged.connect(self._save_current)
        self.group_list.itemChanged.connect(lambda _item: self._save_current())
        self.show_hidden_box.toggled.connect(lambda _checked: self._render_employee_list(self.current_employee()))
        self._update_employee_action_states()

    def set_data(
        self,
        employee_names: list[str],
        hidden_employee_names: set[str],
        employee_emails: dict[str, str],
        employee_notes: dict[str, str],
        employee_groups: dict[str, list[str]],
        missing_email_suppressions: set[str],
    ) -> None:
        current = self.current_employee()
        self.employee_names = sorted(set(employee_names) | set(hidden_employee_names), key=str.casefold)
        self.hidden_employee_names = set(hidden_employee_names)
        self.employee_emails = dict(employee_emails)
        self.employee_notes = dict(employee_notes)
        self.employee_groups = {name: list(values) for name, values in employee_groups.items()}
        self.missing_email_suppressions = set(missing_email_suppressions)
        self._render_employee_list(current)

        self.groups_list.blockSignals(True)
        self.groups_list.clear()
        for group_name in sorted(self.employee_groups, key=str.casefold):
            item = QListWidgetItem(group_name)
            item.setData(Qt.UserRole, group_name)
            self.groups_list.addItem(item)
        self.groups_list.blockSignals(False)
        self._update_employee_action_states()

    def _render_employee_list(self, current: str | None = None) -> None:
        selected = set(self.selected_employees())
        self.employee_list.blockSignals(True)
        self.employee_list.clear()
        for employee in self.employee_names:
            hidden = employee in self.hidden_employee_names
            if hidden and not self.show_hidden_box.isChecked():
                continue
            suppressed = employee in self.missing_email_suppressions
            label, missing = core.build_employee_email_list_label(employee, self.employee_emails.get(employee, ""), suppressed=suppressed)
            if hidden:
                label = f"[Hidden] {label}"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, employee)
            if hidden:
                item.setForeground(QColor("#6b7280"))
            if missing:
                item.setBackground(QColor(core.MISSING_EMAIL_ROW_BACKGROUND))
                item.setForeground(QColor(core.MISSING_EMAIL_ROW_FOREGROUND))
            self.employee_list.addItem(item)
            if employee in selected:
                item.setSelected(True)
        self.employee_list.blockSignals(False)
        if current:
            for row in range(self.employee_list.count()):
                item = self.employee_list.item(row)
                if item.data(Qt.UserRole) == current:
                    self.employee_list.setCurrentItem(item)
                    break
        elif self.employee_list.count():
            self.employee_list.setCurrentRow(0)

    def current_employee(self) -> str:
        item = self.employee_list.currentItem()
        return str(item.data(Qt.UserRole)) if item else ""

    def selected_employees(self) -> list[str]:
        return [str(item.data(Qt.UserRole)) for item in self.employee_list.selectedItems()]

    def snapshot(self) -> tuple[list[str], set[str], dict[str, str], dict[str, str], dict[str, list[str]], set[str]]:
        visible = [name for name in self.employee_names if name not in self.hidden_employee_names]
        return (
            visible,
            set(self.hidden_employee_names),
            dict(self.employee_emails),
            dict(self.employee_notes),
            {name: list(values) for name, values in self.employee_groups.items()},
            set(self.missing_email_suppressions),
        )

    def _populate_details(self) -> None:
        employee = self.current_employee()
        self.employee_name_label.setText(employee)
        self.email_edit.blockSignals(True)
        self.notes_edit.blockSignals(True)
        self.group_list.blockSignals(True)
        self.suppression_box.blockSignals(True)
        self.email_edit.setText(self.employee_emails.get(employee, ""))
        self.notes_edit.setPlainText(self.employee_notes.get(employee, ""))
        self.suppression_box.setChecked(employee in self.missing_email_suppressions)
        self.group_list.clear()
        for group_name in sorted(self.employee_groups, key=str.casefold):
            item = QListWidgetItem(group_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if employee and employee in self.employee_groups.get(group_name, []) else Qt.Unchecked)
            self.group_list.addItem(item)
        self.email_edit.blockSignals(False)
        self.notes_edit.blockSignals(False)
        self.group_list.blockSignals(False)
        self.suppression_box.blockSignals(False)
        self._update_employee_action_states()

    def _save_current(self) -> None:
        employee = self.current_employee()
        if not employee:
            return
        self.employee_emails[employee] = self.email_edit.text().strip()
        self.employee_notes[employee] = self.notes_edit.toPlainText().strip()
        if self.suppression_box.isChecked():
            self.missing_email_suppressions.add(employee)
        else:
            self.missing_email_suppressions.discard(employee)
        for row in range(self.group_list.count()):
            group_name = self.group_list.item(row).text()
            members = set(self.employee_groups.get(group_name, []))
            if self.group_list.item(row).checkState() == Qt.Checked:
                members.add(employee)
            else:
                members.discard(employee)
            self.employee_groups[group_name] = sorted(members)
        self.changed.emit()

    def _add_employee(self) -> None:
        name, ok = QInputDialog.getText(self, "Add Employee", "Employee name:")
        employee = name.strip()
        if not ok or not employee:
            return
        if employee not in self.employee_names:
            self.employee_names.append(employee)
            self.employee_names.sort(key=str.casefold)
        self.hidden_employee_names.discard(employee)
        self.changed.emit()

    def _remove_employee(self) -> None:
        employee = self.current_employee()
        if not employee:
            return
        if employee in self.hidden_employee_names:
            if QMessageBox.question(self, "Restore Employee", f"Restore '{employee}' to the managed roster?") != QMessageBox.Yes:
                return
            self.hidden_employee_names.discard(employee)
        else:
            if QMessageBox.question(self, "Hide Employee", f"Hide '{employee}' from the managed roster?") != QMessageBox.Yes:
                return
            self.hidden_employee_names.add(employee)
        self.changed.emit()

    def _preferred_merge_target(self, names: list[str]) -> str:
        with_email = [name for name in names if self.employee_emails.get(name, "").strip()]
        if len(with_email) == 1:
            return with_email[0]
        if with_email:
            names = with_email
        current = self.current_employee()
        if current in names:
            return current
        return sorted(names, key=str.casefold)[0]

    def _merge_selected_employees(self) -> None:
        selected = self.selected_employees()
        unique = list(dict.fromkeys(selected))
        if len(unique) != 2:
            QMessageBox.information(self, "Merge Employees", "Select exactly two employee names to merge.")
            return
        default_target = self._preferred_merge_target(unique)
        default_source = next(name for name in unique if name != default_target)

        dialog = QDialog(self)
        dialog.setWindowTitle("Merge Employees")
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel("Choose which name should remain after the merge."))
        layout.addWidget(QLabel("Preference: keep the name that already has a real email address."))
        source_combo = QComboBox()
        target_combo = QComboBox()
        for name in unique:
            email = self.employee_emails.get(name, "").strip()
            suffix = f" | {email}" if email else ""
            source_combo.addItem(f"{name}{suffix}", name)
            target_combo.addItem(f"{name}{suffix}", name)
        source_combo.setCurrentIndex(source_combo.findData(default_source))
        target_combo.setCurrentIndex(target_combo.findData(default_target))
        form = QFormLayout()
        form.addRow("Merge this name", source_combo)
        form.addRow("Into this name", target_combo)
        layout.addLayout(form)
        preview = QLabel("")
        preview.setWordWrap(True)
        layout.addWidget(preview)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(buttons)

        def _refresh_preview() -> None:
            source = str(source_combo.currentData())
            target = str(target_combo.currentData())
            ok_button = buttons.button(QDialogButtonBox.Ok)
            if source == target:
                preview.setText("Choose two different names.")
                if ok_button is not None:
                    ok_button.setEnabled(False)
                return
            preview.setText(f"Merge '{source}' into '{target}'. '{target}' will remain.")
            if ok_button is not None:
                ok_button.setEnabled(True)

        source_combo.currentIndexChanged.connect(_refresh_preview)
        target_combo.currentIndexChanged.connect(_refresh_preview)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        _refresh_preview()

        if dialog.exec() != QDialog.Accepted:
            return
        source = str(source_combo.currentData())
        target = str(target_combo.currentData())
        if source == target:
            return
        self.mergeRequested.emit(source, target)

    def _add_group(self) -> None:
        name, ok = QInputDialog.getText(self, "Add Group", "Group name:")
        group_name = name.strip()
        if not ok or not group_name:
            return
        self.employee_groups.setdefault(group_name, [])
        self.changed.emit()

    def _remove_group(self) -> None:
        item = self.groups_list.currentItem()
        if not item:
            return
        group_name = str(item.data(Qt.UserRole))
        if QMessageBox.question(self, "Remove Group", f"Delete employee group '{group_name}'?") != QMessageBox.Yes:
            return
        self.employee_groups.pop(group_name, None)
        self.changed.emit()

    def _group_selected(self) -> None:
        item = self.groups_list.currentItem()
        if not item:
            return
        group_name = str(item.data(Qt.UserRole))
        members = ", ".join(self.employee_groups.get(group_name, [])) or "No members."
        QMessageBox.information(self, "Group Members", f"{group_name}\n\n{members}")

    def _update_employee_action_states(self) -> None:
        selected = self.selected_employees()
        self.merge_employee_button.setEnabled(len(set(selected)) == 2)
        employee = self.current_employee()
        hidden = bool(employee and employee in self.hidden_employee_names)
        self.remove_employee_button.setEnabled(bool(employee))
        self.remove_employee_button.setText("Restore Employee" if hidden else "Hide Employee")
        self.save_email_button.setEnabled(bool(employee))


class FormattingRulesPage(QWidget):
    changed = Signal()

    def __init__(self) -> None:
        super().__init__()
        self.profiles: dict[str, core.FormattingProfile] = {}
        self.current_profile_name = core.DEFAULT_PROFILE_NAME

        layout = QVBoxLayout(self)
        form_box = QGroupBox("Formatting Rules")
        form = QFormLayout(form_box)
        self.profile_combo = QComboBox()
        self.daily_st_edit = QLineEdit()
        self.weekly_st_edit = QLineEdit()
        self.weekly_ot_edit = QLineEdit()
        self.max_hours_edit = QLineEdit()
        self.signin_hours_check_box = QCheckBox()
        self.signin_tolerance_edit = QLineEdit()
        form.addRow("Profile", self.profile_combo)
        form.addRow("Daily ST Alert", self.daily_st_edit)
        form.addRow("Weekly ST Alert", self.weekly_st_edit)
        form.addRow("Weekly OT Alert", self.weekly_ot_edit)
        form.addRow("Max Hours Per Day", self.max_hours_edit)
        form.addRow("Check sign-in time vs entered hours", self.signin_hours_check_box)
        form.addRow("Sign-in mismatch tolerance (hours)", self.signin_tolerance_edit)
        buttons = QHBoxLayout()
        self.new_button = QPushButton("New Profile")
        self.delete_button = QPushButton("Delete Profile")
        self.save_button = QPushButton("Save Profile")
        buttons.addWidget(self.new_button)
        buttons.addWidget(self.delete_button)
        buttons.addWidget(self.save_button)

        layout.addWidget(form_box)
        layout.addLayout(buttons)
        layout.addStretch(1)

        self.profile_combo.currentTextChanged.connect(self._profile_changed)
        self.new_button.clicked.connect(self._new_profile)
        self.delete_button.clicked.connect(self._delete_profile)
        self.save_button.clicked.connect(self._save_profile)

    def set_data(self, profiles: dict[str, core.FormattingProfile], current_profile_name: str) -> None:
        self.profiles = dict(profiles)
        self.current_profile_name = current_profile_name
        self.profile_combo.blockSignals(True)
        self.profile_combo.clear()
        for name in sorted(self.profiles, key=str.casefold):
            self.profile_combo.addItem(name)
        self.profile_combo.setCurrentText(current_profile_name)
        self.profile_combo.blockSignals(False)
        self._load_profile(self.profiles[self.current_profile_name])

    def snapshot(self) -> tuple[dict[str, core.FormattingProfile], str]:
        return dict(self.profiles), self.current_profile_name

    def _profile_changed(self, name: str) -> None:
        if not name or name not in self.profiles:
            return
        self.current_profile_name = name
        self._load_profile(self.profiles[name])
        self.changed.emit()

    def _load_profile(self, profile: core.FormattingProfile) -> None:
        self.daily_st_edit.setText("" if profile.daily_st_threshold is None else core.fmt_hours(profile.daily_st_threshold))
        self.weekly_st_edit.setText("" if profile.st_threshold is None else core.fmt_hours(profile.st_threshold))
        self.weekly_ot_edit.setText("" if profile.ot_threshold is None else core.fmt_hours(profile.ot_threshold))
        self.max_hours_edit.setText("" if profile.max_hours_per_day is None else core.fmt_hours(profile.max_hours_per_day))
        self.signin_hours_check_box.setChecked(profile.signin_hours_check_enabled)
        self.signin_tolerance_edit.setText(
            "" if profile.signin_hours_mismatch_tolerance is None else core.fmt_hours(profile.signin_hours_mismatch_tolerance)
        )

    def _profile_from_form(self, name: str) -> core.FormattingProfile:
        return core.FormattingProfile(
            name=name,
            st_threshold=core.parse_threshold_value(self.weekly_st_edit.text()) if self.weekly_st_edit.text().strip() else None,
            ot_threshold=core.parse_threshold_value(self.weekly_ot_edit.text()) if self.weekly_ot_edit.text().strip() else None,
            daily_st_threshold=core.parse_threshold_value(self.daily_st_edit.text()) if self.daily_st_edit.text().strip() else None,
            max_hours_per_day=core.parse_threshold_value(self.max_hours_edit.text()) if self.max_hours_edit.text().strip() else None,
            signin_hours_check_enabled=self.signin_hours_check_box.isChecked(),
            signin_hours_mismatch_tolerance=core.parse_threshold_value(self.signin_tolerance_edit.text()) if self.signin_tolerance_edit.text().strip() else None,
        )

    def _new_profile(self) -> None:
        name, ok = QInputDialog.getText(self, "New Profile", "Profile name:")
        profile_name = name.strip()
        if not ok or not profile_name:
            return
        self.profiles[profile_name] = core.FormattingProfile(profile_name, None, None, None, None, True, 0.01)
        self.current_profile_name = profile_name
        self.set_data(self.profiles, self.current_profile_name)
        self.changed.emit()

    def _delete_profile(self) -> None:
        if len(self.profiles) <= 1:
            return
        name = self.profile_combo.currentText()
        if QMessageBox.question(self, "Delete Profile", f"Delete profile '{name}'?") != QMessageBox.Yes:
            return
        self.profiles.pop(name, None)
        self.current_profile_name = sorted(self.profiles, key=str.casefold)[0]
        self.set_data(self.profiles, self.current_profile_name)
        self.changed.emit()

    def _save_profile(self) -> None:
        name = self.profile_combo.currentText().strip()
        if not name:
            return
        self.profiles[name] = self._profile_from_form(name)
        self.current_profile_name = name
        self.changed.emit()


class ConfigurationPage(QWidget):
    settingsChanged = Signal()
    applyRequested = Signal()
    resetRequested = Signal()
    clearCacheRequested = Signal()
    clearEmailsRequested = Signal()
    clearAllRequested = Signal()
    exportDiagnosticRequested = Signal()
    testOutlookRequested = Signal()
    showLoadedStatusRequested = Signal()
    submitBugRequested = Signal()
    checkUpdatesRequested = Signal()
    syncOutlookRequested = Signal()
    checkNameTyposRequested = Signal()
    showAppDataRequested = Signal()

    def __init__(self) -> None:
        super().__init__()
        layout = QVBoxLayout(self)
        self.disable_typo_box = QCheckBox()
        self.show_daily_raw_box = QCheckBox()
        self.quickload_box = QCheckBox()
        self.hash_poll_spin = QSpinBox()
        self.hash_poll_spin.setRange(1, 1440)
        self.update_check_delay_spin = QSpinBox()
        self.update_check_delay_spin.setRange(5, 3600)
        self.max_parallel_spin = QSpinBox()
        self.max_parallel_spin.setRange(1, core.MAX_ALLOWED_PARALLEL_PARSE_WORKERS)
        self.partial_preview_box = QCheckBox()
        self.quickload_hotkey_combo = QComboBox()
        self.quickload_hotkey_combo.setEditable(True)
        self.quickload_hotkey_combo.addItems(list(core.QUICKLOAD_CANCEL_HOTKEY_PRESETS))
        self.auto_update_check_box = QCheckBox()
        self.auto_update_download_box = QCheckBox()
        self.signin_hours_check_box = QCheckBox()
        self.library_root_edit = QLineEdit()
        self.library_root_browse_button = QPushButton("Browse...")
        self.version_label = QLabel(f"Application version: {core.APP_VERSION}")
        library_root_row = QWidget()
        library_root_layout = QHBoxLayout(library_root_row)
        library_root_layout.setContentsMargins(0, 0, 0, 0)
        library_root_layout.addWidget(self.library_root_edit, 1)
        library_root_layout.addWidget(self.library_root_browse_button)

        general_box = QGroupBox("General")
        general_form = QFormLayout(general_box)
        general_form.addRow("Show Daily Raw tab", self.show_daily_raw_box)
        general_form.addRow("Application version", self.version_label)

        loading_box = QGroupBox("Loading")
        loading_form = QFormLayout(loading_box)
        loading_form.addRow("Quick load last DSS set on startup", self.quickload_box)
        loading_form.addRow("Check source DSS(s) frequency (minutes)", self.hash_poll_spin)
        loading_form.addRow("Cancel quick-load hotkey", self.quickload_hotkey_combo)
        loading_form.addRow("Delay before automatic update check (seconds)", self.update_check_delay_spin)
        loading_form.addRow("Max parallel workbook parses", self.max_parallel_spin)
        loading_form.addRow("Show partial results while loading", self.partial_preview_box)
        loading_form.addRow("DSS library root folder", library_root_row)

        warnings_box = QGroupBox("Warnings")
        warnings_form = QFormLayout(warnings_box)
        warnings_form.addRow("Disable name typo notifications", self.disable_typo_box)
        warnings_form.addRow("Check sign-in time against entered hours", self.signin_hours_check_box)
        warnings_form.addRow("Automatically check GitHub for updates on startup", self.auto_update_check_box)
        warnings_form.addRow("Automatically download updates on unmetered Wi-Fi", self.auto_update_download_box)

        self.persistence_box = QGroupBox("Persistence")
        persistence_layout = QVBoxLayout(self.persistence_box)
        self.persistence_summary_label = QLabel("")
        self.persistence_summary_label.setWordWrap(True)
        persistence_layout.addWidget(self.persistence_summary_label)

        self.appearance_box = QGroupBox("Appearance")
        self.appearance_box.setCheckable(True)
        self.appearance_box.setChecked(False)
        appearance_layout = QVBoxLayout(self.appearance_box)
        preset_row = QHBoxLayout()
        self.appearance_preset_combo = QComboBox()
        self.appearance_preset_combo.addItems([*core.ui_theme_presets().keys(), "Custom"])
        self.apply_preset_button = QPushButton("Apply Preset")
        preset_row.addWidget(QLabel("Preset"))
        preset_row.addWidget(self.appearance_preset_combo, 1)
        preset_row.addWidget(self.apply_preset_button)
        appearance_layout.addLayout(preset_row)
        self.appearance_content = QWidget()
        self._theme_line_edits: dict[str, QLineEdit] = {}
        appearance_groups_layout = QVBoxLayout(self.appearance_content)
        label_by_attr = {attr: label for label, attr in core.UI_THEME_CONFIG_FIELDS}
        for group_title, attrs in core.UI_THEME_GROUPS:
            group_box = QGroupBox(group_title)
            group_grid = QGridLayout(group_box)
            for row, attr in enumerate(attrs):
                group_grid.addWidget(QLabel(label_by_attr.get(attr, attr)), row, 0)
                edit = self._theme_line_edits.get(attr)
                if edit is None:
                    edit = QLineEdit()
                    self._theme_line_edits[attr] = edit
                pick_button = QPushButton("Pick...")
                pick_button.clicked.connect(lambda _checked=False, key=attr: self._pick_theme_colour(key))
                group_grid.addWidget(edit, row, 1)
                group_grid.addWidget(pick_button, row, 2)
            appearance_groups_layout.addWidget(group_box)
        self.reset_colours_button = QPushButton("Reset colours to sample defaults")
        self.reset_colours_button.clicked.connect(self._reset_theme_defaults)
        appearance_groups_layout.addWidget(self.reset_colours_button, 0, Qt.AlignLeft)
        appearance_layout.addWidget(self.appearance_content)
        self.appearance_box.toggled.connect(self.appearance_content.setVisible)
        self.appearance_content.setVisible(False)

        self.apply_button = QPushButton("Apply Settings")
        self.apply_button.clicked.connect(self.applyRequested.emit)

        maintenance = QGroupBox("Maintenance")
        maintenance_layout = QGridLayout(maintenance)
        self.reset_button = QPushButton("Reset All Settings to Default")
        self.clear_cache_button = QPushButton("Clear Cached DSSs")
        self.clear_emails_button = QPushButton("Clear Stored Emails")
        self.clear_all_button = QPushButton("Clear All Stored Data")
        self.show_app_data_button = QPushButton("Show App Data Folder")
        self.bug_report_button = QPushButton("Submit Bug Report")
        buttons = [
            self.reset_button,
            self.clear_cache_button,
            self.clear_emails_button,
            self.clear_all_button,
        ]
        for idx, button in enumerate(buttons):
            maintenance_layout.addWidget(button, idx // 2, idx % 2)

        diagnostics = QGroupBox("Diagnostics")
        diagnostics_layout = QGridLayout(diagnostics)
        self.export_diagnostic_button = QPushButton("Export Diagnostic Snapshot")
        self.test_outlook_button = QPushButton("Test Outlook Connection")
        self.show_loaded_status_button = QPushButton("Show Loaded DSS Status")
        self.check_updates_button = QPushButton("Check for Updates")
        self.sync_outlook_button = QPushButton("Sync Outlook Emails")
        self.check_name_typos_button = QPushButton("Check Name Typos")
        diagnostics_buttons = [
            self.show_app_data_button,
            self.export_diagnostic_button,
            self.test_outlook_button,
            self.show_loaded_status_button,
            self.bug_report_button,
            self.check_updates_button,
            self.sync_outlook_button,
            self.check_name_typos_button,
        ]
        for idx, button in enumerate(diagnostics_buttons):
            diagnostics_layout.addWidget(button, idx // 2, idx % 2)
        self.update_status_label = QLabel("")
        self.update_status_label.setWordWrap(True)
        diagnostics_layout.addWidget(self.update_status_label, (len(diagnostics_buttons) + 1) // 2, 0, 1, 2)

        layout.addWidget(general_box)
        layout.addWidget(loading_box)
        layout.addWidget(warnings_box)
        layout.addWidget(self.persistence_box)
        layout.addWidget(self.appearance_box)
        layout.addWidget(self.apply_button, 0, Qt.AlignLeft)
        layout.addWidget(maintenance)
        layout.addWidget(diagnostics)
        layout.addStretch(1)

        self.disable_typo_box.toggled.connect(self.settingsChanged.emit)
        self.show_daily_raw_box.toggled.connect(self.settingsChanged.emit)
        self.quickload_box.toggled.connect(self.settingsChanged.emit)
        self.hash_poll_spin.valueChanged.connect(self.settingsChanged.emit)
        self.update_check_delay_spin.valueChanged.connect(self.settingsChanged.emit)
        self.max_parallel_spin.valueChanged.connect(self.settingsChanged.emit)
        self.partial_preview_box.toggled.connect(self.settingsChanged.emit)
        self.quickload_hotkey_combo.lineEdit().editingFinished.connect(self.settingsChanged.emit)
        self.auto_update_check_box.toggled.connect(self.settingsChanged.emit)
        self.auto_update_download_box.toggled.connect(self.settingsChanged.emit)
        self.signin_hours_check_box.toggled.connect(self.settingsChanged.emit)
        self.library_root_edit.editingFinished.connect(self.settingsChanged.emit)
        for edit in self._theme_line_edits.values():
            edit.editingFinished.connect(self.settingsChanged.emit)
            edit.textChanged.connect(lambda _text, field=edit: self._apply_theme_edit_preview(field))
        self.library_root_browse_button.clicked.connect(self._browse_library_root)
        self.apply_preset_button.clicked.connect(self._apply_selected_preset)
        self.reset_button.clicked.connect(self.resetRequested.emit)
        self.clear_cache_button.clicked.connect(self.clearCacheRequested.emit)
        self.clear_emails_button.clicked.connect(self.clearEmailsRequested.emit)
        self.clear_all_button.clicked.connect(self.clearAllRequested.emit)
        self.export_diagnostic_button.clicked.connect(self.exportDiagnosticRequested.emit)
        self.test_outlook_button.clicked.connect(self.testOutlookRequested.emit)
        self.show_loaded_status_button.clicked.connect(self.showLoadedStatusRequested.emit)
        self.bug_report_button.clicked.connect(self.submitBugRequested.emit)
        self.check_updates_button.clicked.connect(self.checkUpdatesRequested.emit)
        self.sync_outlook_button.clicked.connect(self.syncOutlookRequested.emit)
        self.check_name_typos_button.clicked.connect(self.checkNameTyposRequested.emit)
        self.show_app_data_button.clicked.connect(self.showAppDataRequested.emit)
        self._refresh_persistence_summary()

    def _browse_library_root(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Choose DSS Library Root Folder", self.library_root_edit.text().strip() or str(Path.home()))
        if folder:
            self.library_root_edit.setText(folder)
            self.settingsChanged.emit()

    def _pick_theme_colour(self, attr: str) -> None:
        edit = self._theme_line_edits.get(attr)
        if edit is None:
            return
        current = core.normalize_ui_hex_color(edit.text().strip()) or "#ffffff"
        picked = QColorDialog.getColor(QColor(current), self, "Choose colour")
        if picked.isValid():
            edit.setText(picked.name().lower())
            self.settingsChanged.emit()

    def _apply_theme_edit_preview(self, edit: QLineEdit) -> None:
        normalized = core.normalize_ui_hex_color(edit.text().strip())
        if normalized is None:
            edit.setStyleSheet("")
            return
        text_color = _best_contrast_text(normalized)
        edit.setStyleSheet(
            f"QLineEdit {{ background-color: {normalized}; color: {text_color}; border: 1px solid #94a3b8; border-radius: 6px; }}"
        )

    def _reset_theme_defaults(self) -> None:
        defaults = core.DEFAULT_UI_THEME
        for _label, attr in core.UI_THEME_CONFIG_FIELDS:
            edit = self._theme_line_edits.get(attr)
            if edit is not None:
                edit.setText(getattr(defaults, attr))
                self._apply_theme_edit_preview(edit)
        self.appearance_preset_combo.setCurrentText(core.ui_theme_preset_name(defaults))
        self.settingsChanged.emit()

    def _apply_selected_preset(self) -> None:
        preset_name = self.appearance_preset_combo.currentText().strip()
        preset = core.ui_theme_presets().get(preset_name)
        if preset is None:
            return
        for _label, attr in core.UI_THEME_CONFIG_FIELDS:
            edit = self._theme_line_edits.get(attr)
            if edit is not None:
                edit.setText(getattr(preset, attr))
                self._apply_theme_edit_preview(edit)
        self.settingsChanged.emit()

    def _refresh_persistence_summary(self) -> None:
        self.persistence_summary_label.setText(
            "\n".join(
                [
                    f"Warning suppressions are retained for {core.SUPPRESSION_RETENTION_DAYS} days and survive updates.",
                    "Hidden employees remain saved until you restore them and survive updates.",
                    "Employee merges remain saved until you change them and survive updates.",
                    "Manual employee email addresses remain saved until cleared and survive updates.",
                ]
            )
        )

    def set_update_status(self, text: str) -> None:
        self.update_status_label.setText(text)

    def set_settings(self, settings: core.AppSettings) -> None:
        self.disable_typo_box.setChecked(settings.disable_name_typo_notifications)
        self.show_daily_raw_box.setChecked(settings.show_daily_raw_tab)
        self.quickload_box.setChecked(settings.quickload_last_sources_enabled)
        self.hash_poll_spin.setValue(settings.hash_poll_minutes)
        self.update_check_delay_spin.setValue(settings.update_check_delay_seconds)
        self.max_parallel_spin.setValue(settings.max_parallel_parse_workers)
        self.partial_preview_box.setChecked(settings.partial_preview_enabled)
        self.quickload_hotkey_combo.setCurrentText(settings.quickload_cancel_hotkey)
        self.auto_update_check_box.setChecked(settings.auto_update_check_enabled)
        self.auto_update_download_box.setChecked(settings.auto_download_updates_on_unmetered_wifi)
        self.signin_hours_check_box.setChecked(settings.signin_hours_check_enabled)
        self.library_root_edit.setText(settings.dss_library_root)
        for _label, attr in core.UI_THEME_CONFIG_FIELDS:
            edit = self._theme_line_edits.get(attr)
            if edit is not None:
                edit.setText(getattr(settings.ui_theme, attr))
                self._apply_theme_edit_preview(edit)
        self.appearance_preset_combo.setCurrentText(core.ui_theme_preset_name(settings.ui_theme))

    def snapshot(self, current: core.AppSettings) -> core.AppSettings:
        hotkey = core.normalize_quickload_cancel_hotkey(self.quickload_hotkey_combo.currentText().strip())
        if not core.is_allowed_quickload_cancel_hotkey(hotkey):
            raise ValueError("Cancel quick-load hotkey is invalid.")
        theme_payload: dict[str, str] = {
            "reports_outline_background": current.ui_theme.reports_outline_background,
            "reports_outline_foreground": current.ui_theme.reports_outline_foreground,
        }
        for label, attr in core.UI_THEME_CONFIG_FIELDS:
            edit = self._theme_line_edits.get(attr)
            raw = edit.text().strip() if edit is not None else getattr(current.ui_theme, attr)
            normalized = core.normalize_ui_hex_color(raw)
            if normalized is None:
                raise ValueError(f'Invalid colour for "{label}". Use #RRGGBB.')
            theme_payload[attr] = normalized
        return core.AppSettings(
            disable_name_typo_notifications=self.disable_typo_box.isChecked(),
            hash_poll_minutes=int(self.hash_poll_spin.value()),
            update_check_delay_seconds=int(self.update_check_delay_spin.value()),
            show_daily_raw_tab=self.show_daily_raw_box.isChecked(),
            quickload_last_sources_enabled=self.quickload_box.isChecked(),
            quickload_cancel_hotkey=hotkey,
            auto_update_check_enabled=self.auto_update_check_box.isChecked(),
            auto_download_updates_on_unmetered_wifi=self.auto_update_download_box.isChecked(),
            signin_hours_check_enabled=self.signin_hours_check_box.isChecked(),
            dss_library_root=self.library_root_edit.text().strip(),
            max_parallel_parse_workers=int(self.max_parallel_spin.value()),
            partial_preview_enabled=self.partial_preview_box.isChecked(),
            ui_theme=core.parse_ui_theme_payload(theme_payload, defaults=current.ui_theme),
        )


class EmailDraftsPage(QWidget):
    createDraftsRequested = Signal()
    syncRequested = Signal()
    templatesChanged = Signal()

    def __init__(self, theme: core.UiThemeColors, config_path: Path) -> None:
        super().__init__()
        self._selected_week_start: date | None = None
        self.preview_table = DataTablePage(
            TableSpec(
                "email_drafts",
                "Email Drafts",
                (
                    ("employee", "Employee"),
                    ("email", "Email"),
                    ("pf_numbers", "PF#"),
                    ("days", "Rows"),
                    ("st", "ST"),
                    ("ot", "OT"),
                    ("dt", "DT"),
                    ("total", "Total"),
                    ("expanded", "Expanded Hours"),
                ),
            ),
            theme,
            config_path,
        )
        top = QHBoxLayout()
        self.week_combo = QComboBox()
        self.sync_button = QPushButton("Sync Outlook Emails")
        self.create_button = QPushButton("Create Outlook Drafts")
        self.save_templates_button = QPushButton("Save Templates")
        top.addWidget(QLabel("Week Start"))
        top.addWidget(self.week_combo)
        top.addWidget(self.sync_button)
        top.addWidget(self.create_button)
        top.addWidget(self.save_templates_button)
        top.addStretch(1)

        templates = QSplitter(Qt.Vertical)
        subject_box = QGroupBox("Subject Template")
        subject_layout = QVBoxLayout(subject_box)
        self.subject_edit = QTextEdit()
        self.subject_edit.setMaximumHeight(90)
        subject_layout.addWidget(self.subject_edit)

        body_box = QGroupBox("Body Template (HTML)")
        body_layout = QVBoxLayout(body_box)
        self.body_edit = QPlainTextEdit()
        body_layout.addWidget(self.body_edit)

        templates.addWidget(subject_box)
        templates.addWidget(body_box)
        templates.setStretchFactor(0, 1)
        templates.setStretchFactor(1, 2)

        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(templates, 1)
        layout.addWidget(self.preview_table, 2)

        self.sync_button.clicked.connect(self.syncRequested.emit)
        self.create_button.clicked.connect(self.createDraftsRequested.emit)
        self.save_templates_button.clicked.connect(self.templatesChanged.emit)
        self.week_combo.currentIndexChanged.connect(self._remember_selected_week)

    def set_templates(self, subject_template: str, body_template: str) -> None:
        self.subject_edit.setPlainText(subject_template)
        self.body_edit.setPlainText(body_template)

    def set_weeks(self, week_starts: list[date]) -> None:
        current = self._selected_week_start or self.selected_week_start()
        self.week_combo.blockSignals(True)
        self.week_combo.clear()
        for week_start in sorted(week_starts, reverse=True):
            week_end = week_start + core.timedelta(days=6)
            label = f"Week {core.reference_week_number(week_start)} | {core.format_week_label(week_start, week_end)}"
            self.week_combo.addItem(label, week_start)
        if current is not None:
            idx = self.week_combo.findData(current)
            if idx >= 0:
                self.week_combo.setCurrentIndex(idx)
            elif self.week_combo.count():
                self.week_combo.setCurrentIndex(0)
        elif self.week_combo.count():
            self.week_combo.setCurrentIndex(0)
        self.week_combo.blockSignals(False)
        self._remember_selected_week()

    def selected_week_start(self) -> date | None:
        data = self.week_combo.currentData()
        return data if isinstance(data, date) else None

    def _remember_selected_week(self) -> None:
        self._selected_week_start = self.selected_week_start()


class DssQtMainWindow(QMainWindow):
    updateCheckResultReady = Signal(object, object, bool)
    updateCheckErrorRaised = Signal(str, bool)
    updateDownloadSuccessReady = Signal(object, object, bool, bool)
    updateDownloadErrorRaised = Signal(str, bool)
    updateDownloadProgressReady = Signal(int, int)

    def __init__(self, initial_source: list[Path] | None = None) -> None:
        super().__init__()
        self.app_root, self.cache_dir = core.ensure_app_directories()
        self.updates_dir = self.app_root / core.UPDATE_DIRNAME
        self.updates_dir.mkdir(parents=True, exist_ok=True)
        self.config_path = self.app_root / core.CONFIG_FILENAME
        self.table_layouts = core.load_table_layouts(self.config_path)
        self.app_settings = core.load_app_settings(self.config_path)
        self.profiles, self.current_profile_name = core.load_formatting_profiles(self.config_path)
        self.employee_emails = core.load_employee_emails(self.config_path)
        self.employee_outlook_display_names = core.load_employee_outlook_display_names(self.config_path)
        self.outlook_lookup_cache = core.load_outlook_lookup_cache(self.config_path)
        self.employee_notes = core.load_employee_notes(self.config_path)
        self.employee_groups = core.load_employee_groups(self.config_path)
        self.missing_email_suppressions = core.load_missing_email_suppressions(self.config_path)
        self.employee_added_names, self.employee_hidden_names = core.load_employee_name_overrides(self.config_path)
        self.employee_name_merges = core.load_employee_name_merges(self.config_path)
        self.subject_template, self.body_template = core.load_email_templates(self.config_path)
        self.ignored_name_typos = core.load_ignored_name_typos(self.config_path)
        self.suppressed_error_findings = core.load_named_suppressions(self.config_path, "suppressed_error_findings")
        self.suppressed_parse_warnings = core.load_named_suppressions(self.config_path, "suppressed_parse_warnings")
        self.suppressed_workbook_health = core.load_named_suppressions(self.config_path, "suppressed_workbook_health")
        self._cached_name_typo_warnings: list[core.NameTypoWarning] = []
        self._cached_name_typo_key: str | None = None
        self.current_data: core.TrackerData | None = None
        self.source_paths: list[Path] = list(initial_source or [])
        self.load_worker: LoadWorker | None = None
        self.load_thread: threading.Thread | None = None
        self.outlook_worker: OutlookWorker | None = None
        self.outlook_thread: threading.Thread | None = None
        self.name_typo_worker: NameTypoWorker | None = None
        self.name_typo_thread: threading.Thread | None = None
        self._next_load_token = 0
        self._active_load_token = -1
        self._next_outlook_token = 0
        self._active_outlook_token = -1
        self._next_name_typo_token = 0
        self._active_name_typo_token = -1
        self._outlook_partial_updates: dict[str, core.OutlookResolution] = {}
        self._outlook_last_partial_refresh = 0.0
        self._employee_summary_week_start: date | None = None
        self._employee_summary_week_end: date | None = None
        self._show_suppressed_report_rows: dict[str, bool] = {
            "data_review": False,
        }
        self._cancel_shortcut_action: QAction | None = None
        self._quickload_session = False
        self._update_check_in_progress = False
        self._update_download_in_progress = False
        self._auto_update_check_done = False
        self._downloaded_update_path: Path | None = None
        self._update_status_text = f"Installed version: {core.APP_VERSION}"
        self._hash_alerted_paths: set[Path] = set()
        self._refresh_views_queued = False
        self.hash_poll_timer = QTimer(self)
        self.updateCheckResultReady.connect(self._handle_update_check_result)
        self.updateCheckErrorRaised.connect(self._handle_update_check_error)
        self.updateDownloadSuccessReady.connect(self._handle_update_download_success)
        self.updateDownloadErrorRaised.connect(self._handle_update_download_error)
        self.updateDownloadProgressReady.connect(self._handle_update_download_progress)
        self._build_ui()
        self._bind_shortcuts()
        self._apply_window_icon()
        self._sync_ui_from_state()
        self._start_hash_poll_timer()
        if self.source_paths:
            QTimer.singleShot(0, self.reload_data)
        elif self.app_settings.quickload_last_sources_enabled:
            last_paths = [path for path in core.load_last_open_dss_paths(self.config_path) if path.exists()]
            if last_paths:
                self._quickload_session = True
                self.source_paths = last_paths
                QTimer.singleShot(0, self.reload_data)
        QTimer.singleShot(max(5, self.app_settings.update_check_delay_seconds) * 1000, self._auto_check_for_updates)

    def _build_ui(self) -> None:
        self.setWindowTitle(core.DISPLAY_APP_NAME)
        self.setMinimumSize(900, 700)
        self.main_scroll = QScrollArea()
        self.main_scroll.setWidgetResizable(True)
        self.main_scroll.setFrameShape(QScrollArea.NoFrame)
        self.setCentralWidget(self.main_scroll)
        root = QWidget()
        self.main_scroll.setWidget(root)
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setSpacing(10)

        self.add_button = QPushButton("Add DSSs")
        self.quick_add_button = QPushButton("Quick Add")
        self.remove_button = QPushButton("Remove DSSs")
        self.update_button = QPushButton("Refresh")
        self.export_button = QPushButton("Export View")
        self.employee_filter = CheckListButton("All Employees", "Employees")
        self.pf_filter = CheckListButton("All PFs", "PFs")
        self.source_label = QLabel("No workbook loaded")
        self.loading_label = QLabel("")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1000)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        self.percent_label = QLabel("0.0%")
        self.percent_label.setVisible(False)
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setEnabled(False)
        top_button_font = QFont(self.font())
        top_button_font.setPointSize(max(8, top_button_font.pointSize() - 1))
        controls_box = QGroupBox("Workbook Controls")
        controls_layout = QHBoxLayout(controls_box)
        controls_layout.setSpacing(8)
        for widget in (
            self.add_button,
            self.quick_add_button,
            self.remove_button,
            self.update_button,
            self.export_button,
        ):
            widget.setFont(top_button_font)
            widget.setMinimumWidth(112)
            widget.setMinimumHeight(34)
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            controls_layout.addWidget(widget)
        controls_layout.addStretch(1)
        self.controls_box = controls_box
        self.controls_layout = controls_layout
        self.add_button.setToolTip("Add DSS workbooks to the current session without clearing existing ones.")
        self.quick_add_button.setToolTip("Scan the root DSS directory and quickly add matching DSS workbooks.")
        self.remove_button.setToolTip("Remove one or more DSS workbooks from the current session.")
        self.update_button.setToolTip("Re-read the loaded DSS workbooks and refresh all summaries.")
        self.export_button.setToolTip("Export the current table view to Excel.")

        filters_box = QGroupBox("Filters")
        filters_layout = QHBoxLayout(filters_box)
        filters_layout.setSpacing(8)
        filters_layout.addWidget(QLabel("Employee"))
        filters_layout.addWidget(self.employee_filter)
        filters_layout.addWidget(QLabel("PF"))
        filters_layout.addWidget(self.pf_filter)
        filters_layout.addStretch(1)
        self.filters_box = filters_box
        self.filters_layout = filters_layout

        load_box = QGroupBox("Load Status")
        load_layout = QHBoxLayout(load_box)
        load_layout.setSpacing(8)
        load_layout.addWidget(self.source_label, 1)
        load_layout.addWidget(self.loading_label)
        load_layout.addWidget(self.progress_bar, 2)
        load_layout.addWidget(self.cancel_button)
        self.load_box = load_box
        self.load_layout = load_layout

        overview_box = QGroupBox("Overview")
        overview_layout = QHBoxLayout(overview_box)
        self.loaded_files_summary_label = QLabel("Files: 0")
        self.loaded_employees_summary_label = QLabel("Employees: 0")
        self.loaded_pfs_summary_label = QLabel("PFs: 0")
        self.review_summary_label = QLabel("Review: 0")
        self.suppression_summary_label = QLabel("Suppressed: 0")
        self.hidden_summary_label = QLabel("Hidden: 0")
        self.merge_summary_label = QLabel("Merges: 0")
        for widget in (
            self.loaded_files_summary_label,
            self.loaded_employees_summary_label,
            self.loaded_pfs_summary_label,
            self.review_summary_label,
            self.suppression_summary_label,
            self.hidden_summary_label,
            self.merge_summary_label,
        ):
            overview_layout.addWidget(widget)
        overview_layout.addStretch(1)
        self.overview_box = overview_box
        self.overview_layout = overview_layout

        top_groups = QHBoxLayout()
        top_groups.setSpacing(10)
        top_groups.addWidget(controls_box, 3)
        top_groups.addWidget(filters_box, 2)
        top_groups.addWidget(load_box, 3)
        self.top_groups_layout = top_groups
        root_layout.addLayout(top_groups)
        root_layout.addWidget(overview_box)

        status_row = QHBoxLayout()
        self.status_label = QLabel("Load a DSS workbook to view daily and weekly labour summaries.")
        self.status_label.setWordWrap(True)
        self.quickload_hint_label = QLabel("")
        self.quickload_hint_label.setWordWrap(True)
        status_row.addWidget(self.status_label, 1)
        status_row.addWidget(self.quickload_hint_label)
        self.status_row_layout = status_row
        root_layout.addLayout(status_row)

        self.group_tabs = QTabWidget()
        root_layout.addWidget(self.group_tabs, 1)

        self.data_tabs = QTabWidget()
        self.summary_tabs = QTabWidget()
        self.report_tabs = QTabWidget()
        self.settings_tabs = QTabWidget()
        self.group_tabs.addTab(self.data_tabs, "Data")
        self.group_tabs.addTab(self.summary_tabs, "Summaries")
        self.group_tabs.addTab(self.report_tabs, "Reports")
        self.group_tabs.addTab(self.settings_tabs, "Settings")

        theme = self.app_settings.ui_theme
        self.pages: dict[str, DataTablePage] = {}
        self._daily_raw_tab_visible = True
        for parent, spec in [
            (self.data_tabs, TableSpec("daily_raw", "Daily Raw", (("source_file", "Source File"), ("pf_number", "PF#"), ("date", "Date"), ("sheet", "Source Sheet"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours"), ("ranges", "Source Ranges Used")))),
            (self.data_tabs, TableSpec("week_totals", "Week Totals", (("week_start", "Week Start"), ("week_end", "Week End"), ("st", "Whole Crew ST"), ("ot", "Whole Crew OT"), ("dt", "Whole Crew DT"), ("total", "Whole Crew Total"), ("expanded", "Expanded Hours")))),
            (self.summary_tabs, TableSpec("daily_by_pf", "Daily by PF", (("source_file", "Source File"), ("pf_number", "PF#"), ("date", "Date"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours"), ("row_type", "Row Type")))),
            (self.summary_tabs, TableSpec("weekly_by_pf", "Weekly by PF", (("source_file", "Source File"), ("pf_number", "PF#"), ("week_start", "Week Start"), ("week_end", "Week End"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours"), ("row_type", "Row Type")))),
            (self.summary_tabs, TableSpec("pf_totals", "DSS Totals by PF", (("pf_number", "PF#"), ("row_type", "Row Type"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours")))),
            (self.summary_tabs, TableSpec("employee_daily_pf", "Summary by Employee", (("employee", "Employee"), ("week_number", "Week #"), ("date", "Date"), ("pf_number", "PF#"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total")))),
            (self.summary_tabs, TableSpec("combined_daily", "Combined Summary Daily", (("date", "Date"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours")))),
            (self.summary_tabs, TableSpec("combined_weekly", "Combined Summary Weekly", (("week_start", "Week Start"), ("week_end", "Week End"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours")))),
            (self.report_tabs, TableSpec("data_review", "Data Review", (("suppressed", "Suppressed"), ("category", "Category"), ("employee", "Employee"), ("source_file", "Source File"), ("sheet", "Sheet"), ("date", "Date"), ("summary", "Summary"), ("details", "Details")))),
            (None, TableSpec("error_report", "Error Report", (("suppressed", "Suppressed"), ("employee", "Employee"), ("week_start", "Week Start"), ("week_end", "Week End"), ("hour_type", "Rule"), ("limit", "Limit"), ("actual_total", "Actual"), ("delta", "Delta"), ("trigger_date", "Trigger Date"), ("source_files", "Source Files"), ("reason", "Reason"), ("breakdown", "Breakdown")))),
            (None, TableSpec("parse_warnings", "Sheet Parse Warnings", (("suppressed", "Suppressed"), ("source_file", "Source File"), ("sheet", "Sheet"), ("date", "Date"), ("issue", "Issue"), ("details", "Details")))),
            (None, TableSpec("workbook_health", "Workbook Health", (("suppressed", "Suppressed"), ("source_file", "Source File"), ("status", "Status"), ("details", "Details")))),
            (None, TableSpec("name_typos", "Name Typos", (("suppressed", "Suppressed"), ("employee", "Employee"), ("similar_employee", "Suggested Match"), ("similarity", "Similarity"), ("locations", "Locations")))),
            (self.report_tabs, TableSpec("audit_data_trail", "Audit Data Trail", (("source_file", "Source File"), ("pf_number", "PF#"), ("date", "Date"), ("sheet", "Sheet"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours"), ("source_ranges", "Source Ranges"), ("audit", "Audit")))),
        ]:
            page = DataTablePage(spec, theme, self.config_path, row_activate_callback=self._handle_table_row_activated)
            page.layoutChanged.connect(self._save_table_layout)
            if parent is not None:
                parent.addTab(page, spec.title)
            self.pages[spec.table_id] = page
        self._build_report_page_toolbars()
        self._build_employee_summary_toolbar()

        self.email_drafts_page = EmailDraftsPage(theme, self.config_path)
        self.email_drafts_page.preview_table.layoutChanged.connect(self._save_table_layout)
        self.report_tabs.addTab(self.email_drafts_page, "Email Drafts")

        self.configuration_page = ConfigurationPage()
        self.configuration_scroll = QScrollArea()
        self.configuration_scroll.setWidgetResizable(True)
        self.configuration_scroll.setFrameShape(QScrollArea.NoFrame)
        self.configuration_scroll.setWidget(self.configuration_page)
        self.employees_page = EmployeesPage()
        self.formatting_page = FormattingRulesPage()
        self.settings_tabs.addTab(self.configuration_scroll, "Configuration")
        self.settings_tabs.addTab(self.employees_page, "Employees")
        self.settings_tabs.addTab(self.formatting_page, "Formatting Rules")

        self.add_button.clicked.connect(self.add_dss_files)
        self.quick_add_button.clicked.connect(self.quick_add_dss_files)
        self.remove_button.clicked.connect(self.remove_dss_files)
        self.update_button.clicked.connect(lambda: self.reload_data(force_reparse=True))
        self.export_button.clicked.connect(self.export_current_view)
        self.cancel_button.clicked.connect(self.cancel_active_work)
        self.employee_filter.selectionChanged.connect(lambda _values: self.refresh_views())
        self.pf_filter.selectionChanged.connect(lambda _values: self.refresh_views())
        self.email_drafts_page.week_combo.currentIndexChanged.connect(lambda _idx: self.refresh_views())
        self.email_drafts_page.syncRequested.connect(self.sync_outlook_emails)
        self.email_drafts_page.createDraftsRequested.connect(self.create_outlook_drafts)
        self.email_drafts_page.templatesChanged.connect(self.save_email_templates)
        self.employees_page.changed.connect(self._employees_changed)
        self.employees_page.syncRequested.connect(self.sync_outlook_emails)
        self.employees_page.mergeRequested.connect(self._merge_employee_alias)
        self.formatting_page.changed.connect(self._formatting_changed)
        self.configuration_page.settingsChanged.connect(self._settings_changed)
        self.configuration_page.applyRequested.connect(self.apply_settings)
        self.configuration_page.resetRequested.connect(self.reset_settings)
        self.configuration_page.clearCacheRequested.connect(self.clear_cached_dss)
        self.configuration_page.clearEmailsRequested.connect(self.clear_stored_emails)
        self.configuration_page.clearAllRequested.connect(self.clear_all_stored_data)
        self.configuration_page.exportDiagnosticRequested.connect(self.export_diagnostic_snapshot)
        self.configuration_page.testOutlookRequested.connect(self.test_outlook_connection)
        self.configuration_page.showLoadedStatusRequested.connect(self.show_loaded_dss_status)
        self.configuration_page.submitBugRequested.connect(self.submit_bug_report)
        self.configuration_page.checkUpdatesRequested.connect(self.check_for_updates)
        self.configuration_page.syncOutlookRequested.connect(self.sync_outlook_emails)
        self.configuration_page.checkNameTyposRequested.connect(self.check_name_typos_manually)
        self.configuration_page.showAppDataRequested.connect(self.show_app_data_folder)
        self.group_tabs.currentChanged.connect(lambda _index: self._refresh_export_button_state())
        self._apply_responsive_layouts()

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        self._apply_responsive_layouts()

    def _apply_responsive_layouts(self) -> None:
        available_width = self.main_scroll.viewport().width() if hasattr(self, "main_scroll") else self.width()
        stack_top_groups = available_width < 1320
        stack_secondary_rows = available_width < 1080
        stack_overview = available_width < 980

        self.top_groups_layout.setDirection(
            QBoxLayout.TopToBottom if stack_top_groups else QBoxLayout.LeftToRight
        )
        self.filters_layout.setDirection(
            QBoxLayout.TopToBottom if stack_secondary_rows else QBoxLayout.LeftToRight
        )
        self.load_layout.setDirection(
            QBoxLayout.TopToBottom if stack_secondary_rows else QBoxLayout.LeftToRight
        )
        self.status_row_layout.setDirection(
            QBoxLayout.TopToBottom if stack_secondary_rows else QBoxLayout.LeftToRight
        )
        self.overview_layout.setDirection(
            QBoxLayout.TopToBottom if stack_overview else QBoxLayout.LeftToRight
        )

        self.quickload_hint_label.setMaximumWidth(320 if stack_secondary_rows else 260)
        self.data_tabs.currentChanged.connect(lambda _index: self._refresh_export_button_state())
        self.summary_tabs.currentChanged.connect(lambda _index: self._refresh_export_button_state())
        self.report_tabs.currentChanged.connect(lambda _index: self._refresh_export_button_state())

    def _build_employee_summary_toolbar(self) -> None:
        page = self.pages["employee_daily_pf"]
        self.employee_summary_range_box = QCheckBox("Range")
        self.employee_summary_week_start_combo = QComboBox()
        self.employee_summary_to_label = QLabel("to")
        self.employee_summary_week_end_combo = QComboBox()
        self.employee_summary_this_week_button = QPushButton("This Week")
        self.employee_summary_last_week_button = QPushButton("Last Week")
        self.employee_summary_week_end_combo.setEnabled(False)
        self.employee_summary_week_end_combo.setVisible(False)
        self.employee_summary_to_label.setVisible(False)
        self.employee_summary_week_start_combo.setMinimumContentsLength(20)
        self.employee_summary_week_end_combo.setMinimumContentsLength(20)

        for widget in (
            self.employee_summary_last_week_button,
            self.employee_summary_this_week_button,
            self.employee_summary_week_end_combo,
            self.employee_summary_to_label,
            self.employee_summary_week_start_combo,
            self.employee_summary_range_box,
        ):
            page.add_toolbar_widget(widget)

        self.employee_summary_range_box.toggled.connect(self._employee_summary_range_toggled)
        self.employee_summary_week_start_combo.currentIndexChanged.connect(lambda _index: self.refresh_views())
        self.employee_summary_week_end_combo.currentIndexChanged.connect(lambda _index: self.refresh_views())
        self.employee_summary_this_week_button.clicked.connect(self._select_employee_summary_latest_week)
        self.employee_summary_last_week_button.clicked.connect(self._select_employee_summary_previous_week)

    def _build_report_page_toolbars(self) -> None:
        self._report_show_suppressed_boxes: dict[str, QCheckBox] = {}
        self._report_action_buttons: dict[str, dict[str, QPushButton]] = {}
        self._report_selection_rows: dict[str, list[dict[str, Any]]] = {}

        for table_id in ("data_review",):
            page = self.pages[table_id]
            show_box = QCheckBox("Show Suppressed")
            show_box.setChecked(self._show_suppressed_report_rows.get(table_id, False))
            show_box.toggled.connect(lambda checked, page_id=table_id: self._toggle_show_suppressed(page_id, checked))
            page.add_toolbar_widget(show_box)
            self._report_show_suppressed_boxes[table_id] = show_box
            page.selectionChanged.connect(self._report_page_selection_changed)
            page.set_context_menu_callback(lambda rows, global_pos, page_id=table_id: self._open_report_context_menu(page_id, rows, global_pos))

        self._report_action_buttons["data_review"] = {
            "refresh": self.pages["data_review"].add_toolbar_button("Refresh Name Typos", self.check_name_typos_manually),
            "merge": self.pages["data_review"].add_toolbar_button("Merge Selected", self._merge_selected_name_typos),
            "suppress": self.pages["data_review"].add_toolbar_button("Suppress Selected", lambda: self._set_selected_rows_suppressed("data_review", True)),
            "unsuppress": self.pages["data_review"].add_toolbar_button("Unsuppress Selected", lambda: self._set_selected_rows_suppressed("data_review", False)),
        }
        self._update_report_action_states()

    def _toggle_show_suppressed(self, table_id: str, checked: bool) -> None:
        self._show_suppressed_report_rows[table_id] = checked
        self.refresh_views()

    def _report_page_selection_changed(self, table_id: str, rows: list[dict[str, Any]]) -> None:
        self._report_selection_rows[table_id] = rows
        self._update_report_action_states()

    def _update_report_action_states(self) -> None:
        for table_id, buttons in self._report_action_buttons.items():
            rows = self._report_selection_rows.get(table_id, [])
            has_rows = bool(rows)
            has_suppressed = any(bool(row.get("suppressed")) for row in rows)
            has_unsuppressed = any(not bool(row.get("suppressed")) for row in rows)
            has_typo_rows = any(self._report_row_table_id(table_id, row) == "name_typos" for row in rows)
            if "suppress" in buttons:
                buttons["suppress"].setEnabled(has_unsuppressed)
            if "unsuppress" in buttons:
                buttons["unsuppress"].setEnabled(has_suppressed)
            if "merge" in buttons:
                buttons["merge"].setEnabled(has_typo_rows)

    def _open_report_context_menu(self, table_id: str, rows: list[dict[str, Any]], global_pos: QPoint) -> None:
        if not rows:
            return
        menu = QMenu(self)
        has_suppressed = any(bool(row.get("suppressed")) for row in rows)
        has_unsuppressed = any(not bool(row.get("suppressed")) for row in rows)
        has_typo_rows = any(self._report_row_table_id(table_id, row) == "name_typos" for row in rows)
        suppress_action = None
        unsuppress_action = None
        merge_action = None
        if table_id == "data_review":
            suppress_action = menu.addAction("Suppress Selected")
            suppress_action.setEnabled(has_unsuppressed)
            unsuppress_action = menu.addAction("Unsuppress Selected")
            unsuppress_action.setEnabled(has_suppressed)
        if has_typo_rows:
            menu.addSeparator()
            merge_action = menu.addAction("Merge Selected")
            merge_action.setEnabled(True)
        chosen = menu.exec(global_pos)
        if chosen == suppress_action:
            self._set_selected_rows_suppressed(table_id, True)
        elif chosen == unsuppress_action:
            self._set_selected_rows_suppressed(table_id, False)
        elif chosen == merge_action:
            self._merge_selected_name_typos()

    def _bind_shortcuts(self) -> None:
        if self._cancel_shortcut_action is not None:
            self.removeAction(self._cancel_shortcut_action)
            self._cancel_shortcut_action = None
        cancel_shortcut = QKeySequence(self.app_settings.quickload_cancel_hotkey.strip("<>").replace("-", "+"))
        if not cancel_shortcut.isEmpty():
            action = QAction(self)
            action.setShortcut(cancel_shortcut)
            action.triggered.connect(self.cancel_active_work)
            self.addAction(action)
            self._cancel_shortcut_action = action

    def _apply_window_icon(self) -> None:
        icon_path = core.resolve_app_icon_path()
        if icon_path and icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

    def _sync_ui_from_state(self) -> None:
        self.configuration_page.set_settings(self.app_settings)
        self.configuration_page.set_update_status(self._update_status_text)
        self.formatting_page.set_data(self.profiles, self.current_profile_name)
        self.email_drafts_page.set_templates(self.subject_template, self.body_template)
        self._sync_data_tabs_visibility()
        self.hash_poll_timer.setInterval(max(1, self.app_settings.hash_poll_minutes) * 60 * 1000)
        self._refresh_filters()
        self._refresh_employee_page()
        self._refresh_source_status_labels()
        self._refresh_quickload_hint_label()
        self._apply_ui_theme_chrome()
        self._refresh_overview_labels()
        self._queue_refresh_views()
        self._refresh_export_button_state()
        self._set_loading_state(False)

    def _display_employee_name(self, employee: str) -> str:
        return core.resolve_employee_name_merge(employee, self.employee_name_merges)

    def _display_record(self, record: core.DailyRecord) -> core.DailyRecord:
        display_name = self._display_employee_name(record.employee)
        if display_name == record.employee:
            return record
        return replace(record, employee=display_name)

    def _managed_employee_names(self) -> list[str]:
        discovered = {self._display_employee_name(name) for name in (self.current_data.employee_names if self.current_data else [])}
        added = {self._display_employee_name(name) for name in self.employee_added_names}
        hidden = {self._display_employee_name(name) for name in self.employee_hidden_names}
        merged_sources = {source for source in self.employee_name_merges}
        return sorted(((discovered | added) - hidden) - merged_sources, key=str.casefold)

    def _employee_summary_available_weeks(self, records: Iterable[core.DailyRecord]) -> list[date]:
        return sorted({core.monday_week_start(record.work_date) for record in records}, reverse=True)

    def _employee_summary_week_label(self, week_start: date) -> str:
        week_end = week_start + core.timedelta(days=6)
        return f"Week {core.reference_week_number(week_start)} | {core.format_week_label(week_start, week_end)}"

    def _sync_employee_summary_week_controls(self, records: Iterable[core.DailyRecord]) -> None:
        weeks = self._employee_summary_available_weeks(records)
        combos = (self.employee_summary_week_start_combo, self.employee_summary_week_end_combo)
        if not weeks:
            for combo in combos:
                combo.blockSignals(True)
                combo.clear()
                combo.blockSignals(False)
                combo.setEnabled(False)
            self.employee_summary_range_box.setEnabled(False)
            self.employee_summary_this_week_button.setEnabled(False)
            self.employee_summary_last_week_button.setEnabled(False)
            self._employee_summary_week_start = None
            self._employee_summary_week_end = None
            return

        current_start_text = self.employee_summary_week_start_combo.currentData()
        current_end_text = self.employee_summary_week_end_combo.currentData()
        current_start = date.fromisoformat(str(current_start_text)) if current_start_text else self._employee_summary_week_start
        current_end = date.fromisoformat(str(current_end_text)) if current_end_text else self._employee_summary_week_end

        valid_weeks = set(weeks)
        if current_start in valid_weeks:
            self._employee_summary_week_start = current_start
        elif self._employee_summary_week_start not in valid_weeks:
            self._employee_summary_week_start = weeks[0]
        if current_end in valid_weeks:
            self._employee_summary_week_end = current_end
        elif self._employee_summary_week_end not in valid_weeks:
            self._employee_summary_week_end = self._employee_summary_week_start
        if self._employee_summary_week_end and self._employee_summary_week_start and self._employee_summary_week_end < self._employee_summary_week_start:
            older_week = min(self._employee_summary_week_start, self._employee_summary_week_end)
            newer_week = max(self._employee_summary_week_start, self._employee_summary_week_end)
            self._employee_summary_week_start = older_week
            self._employee_summary_week_end = newer_week

        for combo, selected in ((self.employee_summary_week_start_combo, self._employee_summary_week_start), (self.employee_summary_week_end_combo, self._employee_summary_week_end)):
            combo.blockSignals(True)
            combo.clear()
            for week_start in weeks:
                combo.addItem(self._employee_summary_week_label(week_start), week_start.isoformat())
            if selected is not None:
                match_index = next((idx for idx, week_start in enumerate(weeks) if week_start == selected), 0)
                combo.setCurrentIndex(match_index)
            combo.blockSignals(False)
        self.employee_summary_range_box.setEnabled(len(weeks) > 1)
        self.employee_summary_this_week_button.setEnabled(True)
        self.employee_summary_last_week_button.setEnabled(len(weeks) > 1)
        self.employee_summary_week_start_combo.setEnabled(True)
        self.employee_summary_week_end_combo.setEnabled(self.employee_summary_range_box.isChecked())
        self.employee_summary_week_end_combo.setVisible(self.employee_summary_range_box.isChecked())
        self.employee_summary_to_label.setVisible(self.employee_summary_range_box.isChecked())

    def _employee_summary_selected_weeks(self) -> tuple[date | None, date | None]:
        start_text = self.employee_summary_week_start_combo.currentData()
        end_text = self.employee_summary_week_end_combo.currentData()
        start_week = date.fromisoformat(str(start_text)) if start_text else self._employee_summary_week_start
        end_week = date.fromisoformat(str(end_text)) if end_text else self._employee_summary_week_end
        if not self.employee_summary_range_box.isChecked():
            end_week = start_week
        if start_week and end_week and end_week < start_week:
            start_week, end_week = end_week, start_week
        self._employee_summary_week_start = start_week
        self._employee_summary_week_end = end_week
        start = start_week
        end = end_week + core.timedelta(days=6) if end_week else None
        return start, end

    def _employee_summary_range_toggled(self, checked: bool) -> None:
        self.employee_summary_week_end_combo.setEnabled(checked and self.employee_summary_week_end_combo.count() > 0)
        self.employee_summary_week_end_combo.setVisible(checked)
        self.employee_summary_to_label.setVisible(checked)
        if checked and self.employee_summary_week_end is None:
            self._employee_summary_week_end = self._employee_summary_week_start
        self.refresh_views()

    def _select_employee_summary_latest_week(self) -> None:
        if self.employee_summary_week_start_combo.count() <= 0:
            return
        latest_week = date.fromisoformat(str(self.employee_summary_week_start_combo.itemData(0)))
        self._employee_summary_week_start = latest_week
        self._employee_summary_week_end = latest_week
        self.employee_summary_week_start_combo.blockSignals(True)
        self.employee_summary_week_end_combo.blockSignals(True)
        self.employee_summary_range_box.blockSignals(True)
        self.employee_summary_week_start_combo.setCurrentIndex(0)
        self.employee_summary_week_end_combo.setCurrentIndex(0)
        self.employee_summary_range_box.setChecked(False)
        self.employee_summary_week_start_combo.blockSignals(False)
        self.employee_summary_week_end_combo.blockSignals(False)
        self.employee_summary_range_box.blockSignals(False)
        self.employee_summary_week_end_combo.setEnabled(False)
        self.employee_summary_week_end_combo.setVisible(False)
        self.employee_summary_to_label.setVisible(False)
        self.refresh_views()

    def _select_employee_summary_previous_week(self) -> None:
        if self.employee_summary_week_start_combo.count() <= 1:
            return
        previous_week = date.fromisoformat(str(self.employee_summary_week_start_combo.itemData(1)))
        self._employee_summary_week_start = previous_week
        self._employee_summary_week_end = previous_week
        self.employee_summary_week_start_combo.blockSignals(True)
        self.employee_summary_week_end_combo.blockSignals(True)
        self.employee_summary_range_box.blockSignals(True)
        self.employee_summary_week_start_combo.setCurrentIndex(1)
        self.employee_summary_week_end_combo.setCurrentIndex(1)
        self.employee_summary_range_box.setChecked(False)
        self.employee_summary_week_start_combo.blockSignals(False)
        self.employee_summary_week_end_combo.blockSignals(False)
        self.employee_summary_range_box.blockSignals(False)
        self.employee_summary_week_end_combo.setEnabled(False)
        self.employee_summary_week_end_combo.setVisible(False)
        self.employee_summary_to_label.setVisible(False)
        self.refresh_views()

    def _selected_employee_values(self) -> set[str]:
        selected = set(self.employee_filter.selected_values())
        all_values = set(self._managed_employee_names())
        return all_values if not selected or selected == all_values else selected

    def _selected_pf_values(self) -> set[str]:
        available = set(self._available_pfs())
        selected = set(self.pf_filter.selected_values())
        return available if not selected or selected == available else selected

    def _available_pfs(self) -> list[str]:
        if not self.current_data:
            return []
        return sorted(
            {core.display_pf_number(record.pf_number) for record in self.current_data.daily_records},
            key=core.pf_number_sort_key,
        )

    def _sync_data_tabs_visibility(self) -> None:
        daily_raw_page = self.pages["daily_raw"]
        current_index = self.data_tabs.indexOf(daily_raw_page)
        if self.app_settings.show_daily_raw_tab and current_index < 0:
            self.data_tabs.insertTab(0, daily_raw_page, "Daily Raw")
        elif not self.app_settings.show_daily_raw_tab and current_index >= 0:
            self.data_tabs.removeTab(current_index)

    def _refresh_filters(self) -> None:
        employees = self._managed_employee_names()
        self.employee_filter.set_choices(employees, self.employee_filter.selected_values())
        pfs = self._available_pfs()
        self.pf_filter.set_choices(pfs, self.pf_filter.selected_values(), force_single=True)

    def _refresh_employee_page(self) -> None:
        hidden_names = {
            self._display_employee_name(name)
            for name in self.employee_hidden_names
            if name not in self.employee_name_merges
        }
        self.employees_page.set_data(
            self._managed_employee_names(),
            hidden_names,
            self.employee_emails,
            self.employee_notes,
            self.employee_groups,
            self.missing_email_suppressions,
        )

    def _set_update_status(self, text: str) -> None:
        self._update_status_text = text
        self.configuration_page.set_update_status(text)

    def _refresh_source_status_labels(self) -> None:
        if self.current_data and self.current_data.source_paths:
            if len(self.current_data.source_paths) == 1:
                self.source_label.setText(self.current_data.source_paths[0].name)
            else:
                self.source_label.setText(f"{len(self.current_data.source_paths)} DSS workbooks loaded")
            return
        if self.source_paths:
            if len(self.source_paths) == 1:
                self.source_label.setText(self.source_paths[0].name)
            else:
                self.source_label.setText(f"{len(self.source_paths)} DSS workbooks selected")
            return
        self.source_label.setText("No workbook loaded")

    def _refresh_quickload_hint_label(self) -> None:
        if self.load_worker is not None and self._quickload_session:
            hotkey = self.app_settings.quickload_cancel_hotkey.strip() or "<Escape>"
            self.quickload_hint_label.setText(f"Quick load active. Cancel with {hotkey}.")
        else:
            self.quickload_hint_label.setText("")

    def _effective_qt_theme(self) -> core.UiThemeColors:
        return self.app_settings.ui_theme

    def _queue_refresh_views(self) -> None:
        if self._refresh_views_queued:
            return
        self._refresh_views_queued = True

        def run_refresh() -> None:
            self._refresh_views_queued = False
            self.refresh_views()

        QTimer.singleShot(0, run_refresh)

    def _apply_ui_theme_chrome(self) -> None:
        theme = self._effective_qt_theme()
        app = QApplication.instance()
        if app is not None:
            app.setPalette(_build_forced_qt_palette(theme))
        self.setStyleSheet(_build_qt_chrome_stylesheet(theme))

    def _sync_reports_alert_chrome(self, has_errors: bool, has_parse_warnings: bool, has_name_typos: bool = False, has_workbook_health: bool = False) -> None:
        review_index = self.report_tabs.indexOf(self.pages["data_review"])
        reports_index = self.group_tabs.indexOf(self.report_tabs)
        has_review_items = has_errors or has_parse_warnings or has_name_typos or has_workbook_health
        self.report_tabs.setTabText(review_index, "Data Review (!)" if has_review_items else "Data Review")
        self.group_tabs.setTabText(reports_index, "Reports (!)" if has_review_items else "Reports")

    def _invalidate_name_typo_cache(self) -> None:
        self._cached_name_typo_warnings = []
        self._cached_name_typo_key = None

    def _current_name_typo_cache_key(self) -> str | None:
        if self.current_data is None:
            return None
        payload = {
            "file_hashes": sorted((str(path), digest) for path, digest in self.current_data.file_hashes.items()),
            "employees": sorted(self._managed_employee_names(), key=str.casefold),
            "employee_emails": sorted((name, email.strip()) for name, email in self.employee_emails.items() if name.strip()),
            "missing_email_suppressions": sorted(self.missing_email_suppressions, key=str.casefold),
            "employee_outlook_display_names": sorted((name, value.strip()) for name, value in self.employee_outlook_display_names.items() if name.strip()),
        }
        return json.dumps(payload, sort_keys=True, separators=(",", ":"))

    def _refresh_overview_labels(self) -> None:
        if self.current_data is None:
            loaded_files = 0
            employees = 0
            pf_count = 0
            unsuppressed_review = 0
        else:
            loaded_files = len(self.current_data.source_paths)
            employees = len(self._managed_employee_names())
            pf_count = len({core.display_pf_number(record.pf_number) for record in self.current_data.daily_records if core.display_pf_number(record.pf_number)})
            unsuppressed_review = (
                sum(1 for item in core.find_error_findings(self.current_data.daily_records, self._active_profile()) if not self._is_row_suppressed("error_report", core.error_finding_suppression_key(item)))
                + sum(1 for warning in self._current_parse_warnings() if not self._is_row_suppressed("parse_warnings", core.sheet_parse_warning_suppression_key(warning)))
                + sum(1 for item in self.current_data.workbook_health if not self._is_row_suppressed("workbook_health", core.workbook_health_suppression_key(item)))
                + sum(1 for warning in self._current_name_typo_warnings() if not self._is_row_suppressed("name_typos", core.typo_warning_key(warning.employee, warning.similar_employee)))
            )
        suppressed_total = (
            len(self.suppressed_error_findings)
            + len(self.suppressed_parse_warnings)
            + len(self.suppressed_workbook_health)
            + len(self.ignored_name_typos)
            + len(self.missing_email_suppressions)
        )
        self.loaded_files_summary_label.setText(f"Files: {loaded_files}")
        self.loaded_employees_summary_label.setText(f"Employees: {employees}")
        self.loaded_pfs_summary_label.setText(f"PFs: {pf_count}")
        self.review_summary_label.setText(f"Review: {unsuppressed_review}")
        self.suppression_summary_label.setText(f"Suppressed: {suppressed_total}")
        self.hidden_summary_label.setText(f"Hidden: {len(self.employee_hidden_names)}")
        self.merge_summary_label.setText(f"Merges: {len(self.employee_name_merges)}")

    def _active_profile(self) -> core.FormattingProfile:
        return self.profiles.get(self.current_profile_name, next(iter(self.profiles.values())))

    def _current_parse_warnings(self) -> list[core.SheetParseWarning]:
        if not self.current_data:
            return []
        return [
            *self.current_data.parse_warnings,
            *core.build_signin_hours_mismatch_warnings(
                self.current_data.daily_records,
                self.app_settings,
                self._active_profile(),
            ),
        ]

    def _set_cached_name_typo_warnings(self, warnings: list[core.NameTypoWarning]) -> None:
        deduped: list[core.NameTypoWarning] = []
        seen_keys: set[str] = set()
        for warning in warnings:
            key = core.typo_warning_key(warning.employee, warning.similar_employee)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            deduped.append(warning)
        self._cached_name_typo_warnings = deduped
        self._cached_name_typo_key = self._current_name_typo_cache_key()

    def _current_name_typo_warnings(self) -> list[core.NameTypoWarning]:
        allowed_employees = self._selected_employee_values()
        return [
            warning
            for warning in self._cached_name_typo_warnings
            if self._display_employee_name(warning.employee) in allowed_employees
            if self._display_employee_name(warning.employee).casefold() != self._display_employee_name(warning.similar_employee).casefold()
        ]

    def _suppression_set_for_table(self, table_id: str) -> set[str]:
        if table_id == "error_report":
            return self.suppressed_error_findings
        if table_id == "parse_warnings":
            return self.suppressed_parse_warnings
        if table_id == "workbook_health":
            return self.suppressed_workbook_health
        if table_id == "name_typos":
            return self.ignored_name_typos
        return set()

    def _report_row_table_id(self, table_id: str, row: dict[str, Any]) -> str:
        review_table_id = str(row.get("__review_table_id__", "")).strip()
        return review_table_id or table_id

    def _persist_suppression_set(self, table_id: str) -> None:
        if table_id == "error_report":
            core.save_named_suppressions(self.config_path, "suppressed_error_findings", self.suppressed_error_findings)
        elif table_id == "parse_warnings":
            core.save_named_suppressions(self.config_path, "suppressed_parse_warnings", self.suppressed_parse_warnings)
        elif table_id == "workbook_health":
            core.save_named_suppressions(self.config_path, "suppressed_workbook_health", self.suppressed_workbook_health)
        elif table_id == "name_typos":
            self._persist_ignored_name_typos()

    def _suppression_key_for_row(self, table_id: str, row: dict[str, Any]) -> str:
        table_id = self._report_row_table_id(table_id, row)
        key = str(row.get("__suppression_key__", "")).strip()
        if key:
            return key
        if table_id == "name_typos":
            return core.typo_warning_key(str(row.get("employee", "")), str(row.get("similar_employee", "")))
        return ""

    def _is_row_suppressed(self, table_id: str, suppression_key: str) -> bool:
        return bool(suppression_key) and suppression_key in self._suppression_set_for_table(table_id)

    def _set_selected_rows_suppressed(self, table_id: str, suppressed: bool) -> None:
        rows = self._report_selection_rows.get(table_id, [])
        if not rows:
            return
        changed = False
        for row in rows:
            row_table_id = self._report_row_table_id(table_id, row)
            suppressions = self._suppression_set_for_table(row_table_id)
            suppression_key = self._suppression_key_for_row(row_table_id, row)
            if not suppression_key:
                continue
            if suppressed:
                if suppression_key not in suppressions:
                    suppressions.add(suppression_key)
                    changed = True
            elif suppression_key in suppressions:
                suppressions.discard(suppression_key)
                changed = True
        if not changed:
            return
        for changed_table_id in {self._report_row_table_id(table_id, row) for row in rows}:
            self._persist_suppression_set(changed_table_id)
        self.refresh_views()

    def _merge_selected_name_typos(self) -> None:
        rows = [
            row
            for row in self._report_selection_rows.get("data_review", [])
            if self._report_row_table_id("data_review", row) == "name_typos"
        ]
        if not rows:
            return
        for row in rows:
            source_name = str(row.get("employee", "")).strip()
            target_name = str(row.get("similar_employee", "")).strip()
            if source_name and target_name:
                self._merge_employee_alias(source_name, target_name)
        self.refresh_views()

    def _build_name_typo_warnings(self, employee_names: list[str], daily_records: list[core.DailyRecord], address_book_names: list[str] | None = None) -> list[core.NameTypoWarning]:
        names_to_check = [
            employee
            for employee in employee_names
            if not self.employee_emails.get(employee, "").strip() and employee not in self.missing_email_suppressions
        ]
        warnings: list[core.NameTypoWarning] = []
        if names_to_check:
            warnings.extend(core.find_potential_name_typos(names_to_check, employee_names, daily_records))
            if address_book_names:
                warnings.extend(core.find_address_book_name_typos(names_to_check, address_book_names, daily_records))
        else:
            warnings.extend(core.find_similar_employee_name_pairs(employee_names, daily_records))
        warnings.extend(core.find_outlook_display_name_typos(employee_names, self.employee_outlook_display_names, daily_records))
        return warnings

    def _daily_records_filtered(self) -> list[core.DailyRecord]:
        if not self.current_data:
            return []
        allowed_employees = self._selected_employee_values()
        allowed_pfs = self._selected_pf_values()
        results: list[core.DailyRecord] = []
        for record in self.current_data.daily_records:
            display_record = self._display_record(record)
            if display_record.employee not in allowed_employees:
                continue
            if allowed_pfs and core.display_pf_number(display_record.pf_number) not in allowed_pfs:
                continue
            results.append(display_record)
        return results

    def refresh_views(self) -> None:
        theme = self._effective_qt_theme()
        for page in self.pages.values():
            page.set_theme(theme)
        self.email_drafts_page.preview_table.set_theme(theme)
        if not self.current_data:
            self._sync_employee_summary_week_controls([])
            for page in self.pages.values():
                page.set_rows([])
            self.email_drafts_page.preview_table.set_rows([])
            self.email_drafts_page.set_weeks([])
            self._sync_reports_alert_chrome(False, False, False)
            return

        filtered_records = self._daily_records_filtered()
        profile = self._active_profile()
        daily_summary = core.aggregate_daily(filtered_records, combine_sources=False)
        weekly_summary = core.aggregate_weekly(filtered_records, combine_sources=False)
        daily_rollup = core.build_daily_rollup(daily_summary)
        weekly_rollup = core.build_weekly_rollup(weekly_summary)
        employee_week_start, employee_week_end = self._employee_summary_selected_weeks()
        self._sync_employee_summary_week_controls(filtered_records)
        employee_week_start, employee_week_end = self._employee_summary_selected_weeks()
        employee_daily_pf = core.build_employee_day_pf_rows(
            filtered_records,
            week_start=employee_week_start,
            week_end=employee_week_end,
        )
        combined_daily = core.aggregate_daily(filtered_records, combine_sources=True)
        combined_weekly = core.aggregate_weekly(filtered_records, combine_sources=True)
        week_totals = core.build_week_totals(combined_weekly)
        pf_totals = core.build_pf_totals(filtered_records)
        findings = core.build_error_findings(filtered_records, profile)
        name_typo_warnings = self._current_name_typo_warnings()
        week_starts = sorted({record.week_start for record in combined_weekly}, reverse=True)

        self.email_drafts_page.set_weeks(week_starts)
        self.pages["daily_raw"].set_rows([
            {
                "source_file": record.source_file,
                "pf_number": core.display_pf_number(record.pf_number),
                "date": record.work_date.isoformat(),
                "sheet": record.source_sheet,
                "employee": record.employee,
                "st": core.fmt_hours(record.st),
                "ot": core.fmt_hours(record.ot),
                "dt": core.fmt_hours(record.dt),
                "total": core.fmt_hours(record.total),
                "expanded": core.fmt_hours(core.expanded_hours(record.st, record.ot, record.dt)),
                "ranges": record.source_ranges,
            }
            for record in sorted(filtered_records, key=lambda item: (item.work_date, item.source_sheet, item.employee), reverse=True)
        ])
        self.pages["daily_by_pf"].set_rows([
            {
                "source_file": row.source_file,
                "pf_number": core.display_pf_number(row.pf_number),
                "date": row.work_date.isoformat(),
                "employee": row.employee,
                "st": core.fmt_hours(row.st),
                "ot": core.fmt_hours(row.ot),
                "dt": core.fmt_hours(row.dt),
                "total": core.fmt_hours(row.total),
                "expanded": core.fmt_hours(core.expanded_hours(row.st, row.ot, row.dt)),
                "row_type": row.row_type,
                "__tags__": ("crew_total",) if row.row_type == "Crew Total" else (),
            }
            for row in sorted(daily_rollup, key=lambda item: (item.work_date, item.pf_number, item.row_type == "Crew Total", item.employee), reverse=True)
        ])
        self.pages["weekly_by_pf"].set_rows([
            {
                "source_file": row.source_file,
                "pf_number": core.display_pf_number(row.pf_number),
                "week_start": row.week_start.isoformat(),
                "week_end": row.week_end.isoformat(),
                "employee": row.employee,
                "st": core.fmt_hours(row.st),
                "ot": core.fmt_hours(row.ot),
                "dt": core.fmt_hours(row.dt),
                "total": core.fmt_hours(row.total),
                "expanded": core.fmt_hours(core.expanded_hours(row.st, row.ot, row.dt)),
                "row_type": row.row_type,
                "__tags__": (
                    ("crew_total",)
                    if row.row_type == "Crew Total"
                    else ("alert",) if core.is_alert_triggered(row.st, row.ot, row.dt, profile) else ()
                ),
            }
            for row in sorted(weekly_rollup, key=lambda item: (item.week_start, item.pf_number, item.row_type == "Crew Total", item.employee), reverse=True)
        ])
        self.pages["pf_totals"].set_rows([
            {
                "pf_number": core.display_pf_number(row.pf_number),
                "row_type": row.row_type,
                "st": core.fmt_hours(row.st),
                "ot": core.fmt_hours(row.ot),
                "dt": core.fmt_hours(row.dt),
                "total": core.fmt_hours(row.total),
                "expanded": core.fmt_hours(core.expanded_hours(row.st, row.ot, row.dt)),
                "__tags__": ("crew_total",) if row.row_type == "Overall DSS Total" else (),
            }
            for row in pf_totals
        ])
        employee_daily_rows: list[dict[str, str]] = []
        previous_employee = ""
        previous_date = ""
        previous_week_start: date | None = None
        for index, row in enumerate(employee_daily_pf):
            date_text = row.work_date.strftime("%d-%b")
            week_number = str(core.reference_week_number(row.work_date))
            employee_changed = row.employee != previous_employee
            week_changed = row.week_start != previous_week_start
            employee_text = row.employee if employee_changed else ""
            week_cell = week_number if employee_changed or week_changed else ""
            is_date_break = employee_changed or week_changed or date_text != previous_date
            date_cell = date_text if is_date_break else ""
            tags: list[str] = []
            if is_date_break:
                tags.append("date_break")
            if employee_changed:
                tags.append("employee_break")
            employee_daily_rows.append(
                {
                    "employee": employee_text,
                    "week_number": week_cell,
                    "date": date_cell,
                    "pf_number": core.display_pf_number(row.pf_number),
                    "st": core.fmt_hours(row.st),
                    "ot": core.fmt_hours(row.ot),
                    "dt": core.fmt_hours(row.dt),
                    "total": core.fmt_hours(row.total),
                    "__employee_full__": row.employee,
                    "__week_start__": row.week_start.isoformat(),
                    "__date_sort__": row.work_date.isoformat(),
                    "__source_index__": index,
                    "__tags__": tuple(tags),
                }
            )
            previous_employee = row.employee
            previous_date = date_text
            previous_week_start = row.week_start
        self.pages["employee_daily_pf"].set_rows(employee_daily_rows)
        self.pages["combined_daily"].set_rows([
            {
                "date": row.work_date.isoformat(),
                "employee": row.employee,
                "st": core.fmt_hours(row.st),
                "ot": core.fmt_hours(row.ot),
                "dt": core.fmt_hours(row.dt),
                "total": core.fmt_hours(row.total),
                "expanded": core.fmt_hours(core.expanded_hours(row.st, row.ot, row.dt)),
            }
            for row in sorted(combined_daily, key=lambda item: (item.work_date, item.employee), reverse=True)
        ])
        self.pages["combined_weekly"].set_rows([
            {
                "week_start": row.week_start.isoformat(),
                "week_end": row.week_end.isoformat(),
                "employee": row.employee,
                "st": core.fmt_hours(row.st),
                "ot": core.fmt_hours(row.ot),
                "dt": core.fmt_hours(row.dt),
                "total": core.fmt_hours(row.total),
                "expanded": core.fmt_hours(core.expanded_hours(row.st, row.ot, row.dt)),
                "__tags__": ("alert",) if core.is_alert_triggered(row.st, row.ot, row.dt, profile) else (),
            }
            for row in sorted(combined_weekly, key=lambda item: (item.week_start, item.employee), reverse=True)
        ])
        self.pages["week_totals"].set_rows([
            {
                "week_start": row.week_start.isoformat(),
                "week_end": row.week_end.isoformat(),
                "st": core.fmt_hours(row.st),
                "ot": core.fmt_hours(row.ot),
                "dt": core.fmt_hours(row.dt),
                "total": core.fmt_hours(row.total),
                "expanded": core.fmt_hours(core.expanded_hours(row.st, row.ot, row.dt)),
            }
            for row in sorted(week_totals, key=lambda item: item.week_start, reverse=True)
        ])
        error_rows: list[dict[str, Any]] = []
        data_review_rows: list[dict[str, Any]] = []
        for item in sorted(findings, key=lambda finding: (finding.week_start, finding.employee, finding.hour_type), reverse=True):
            suppression_key = core.error_finding_suppression_key(item)
            suppressed = self._is_row_suppressed("error_report", suppression_key)
            if suppressed and not self._show_suppressed_report_rows["data_review"]:
                continue
            tags = ["alert"]
            if suppressed:
                tags.append("suppressed")
            row = {
                "suppressed": suppressed,
                "employee": item.employee,
                "week_start": item.week_start.isoformat(),
                "week_end": item.week_end.isoformat(),
                "hour_type": item.hour_type,
                "limit": core.fmt_hours(item.threshold),
                "actual_total": core.fmt_hours(item.actual_total),
                "delta": core.fmt_hours(item.delta),
                "trigger_date": item.trigger_date.isoformat(),
                "source_files": item.source_files,
                "reason": item.reason,
                "breakdown": item.breakdown,
                "__tags__": tuple(tags),
                "__finding__": item,
                "__suppression_key__": suppression_key,
            }
            error_rows.append(row)
            data_review_rows.append(
                {
                    "suppressed": suppressed,
                    "category": "Hour Count Error",
                    "employee": item.employee,
                    "source_file": item.source_files,
                    "sheet": "",
                    "date": item.trigger_date.isoformat(),
                    "summary": f"{item.hour_type}: {core.fmt_hours(item.actual_total)} vs {core.fmt_hours(item.threshold)}",
                    "details": item.reason,
                    "__tags__": tuple(tags),
                    "__finding__": item,
                    "__suppression_key__": suppression_key,
                    "__review_table_id__": "error_report",
                }
            )
        self.pages["error_report"].set_rows(error_rows)
        filtered_source_files = {record.source_file for record in filtered_records}
        active_parse_warnings = self._current_parse_warnings()
        parse_warning_rows: list[dict[str, Any]] = []
        for warning in active_parse_warnings:
            if warning.source_file not in filtered_source_files:
                continue
            suppression_key = core.sheet_parse_warning_suppression_key(warning)
            suppressed = self._is_row_suppressed("parse_warnings", suppression_key)
            if suppressed and not self._show_suppressed_report_rows["data_review"]:
                continue
            tags = ("suppressed",) if suppressed else ()
            row = {
                "suppressed": suppressed,
                "source_file": warning.source_file,
                "sheet": warning.source_sheet,
                "date": warning.work_date,
                "issue": warning.issue,
                "details": warning.details,
                "__warning__": warning,
                "__suppression_key__": suppression_key,
                "__tags__": tags,
            }
            parse_warning_rows.append(row)
            data_review_rows.append(
                {
                    "suppressed": suppressed,
                    "category": "Sheet Parse Warning",
                    "employee": "",
                    "source_file": warning.source_file,
                    "sheet": warning.source_sheet,
                    "date": warning.work_date,
                    "summary": warning.issue,
                    "details": warning.details,
                    "__warning__": warning,
                    "__suppression_key__": suppression_key,
                    "__tags__": tags,
                    "__review_table_id__": "parse_warnings",
                }
            )
        self.pages["parse_warnings"].set_rows(parse_warning_rows)
        workbook_health_rows: list[dict[str, Any]] = []
        for item in self.current_data.workbook_health:
            if item.source_file not in filtered_source_files:
                continue
            suppression_key = core.workbook_health_suppression_key(item)
            suppressed = self._is_row_suppressed("workbook_health", suppression_key)
            if suppressed and not self._show_suppressed_report_rows["data_review"]:
                continue
            tags = ("suppressed",) if suppressed else ()
            row = {
                "suppressed": suppressed,
                "source_file": item.source_file,
                "status": item.status,
                "details": item.details,
                "__suppression_key__": suppression_key,
                "__tags__": tags,
            }
            workbook_health_rows.append(row)
            data_review_rows.append(
                {
                    "suppressed": suppressed,
                    "category": "Workbook Health",
                    "employee": "",
                    "source_file": item.source_file,
                    "sheet": "",
                    "date": "",
                    "summary": item.status,
                    "details": item.details,
                    "__suppression_key__": suppression_key,
                    "__tags__": tags,
                    "__review_table_id__": "workbook_health",
                }
            )
        self.pages["workbook_health"].set_rows(workbook_health_rows)
        name_typo_rows: list[dict[str, Any]] = []
        for warning in name_typo_warnings:
            suppression_key = core.typo_warning_key(warning.employee, warning.similar_employee)
            suppressed = self._is_row_suppressed("name_typos", suppression_key)
            if suppressed and not self._show_suppressed_report_rows["data_review"]:
                continue
            tags = ("suppressed",) if suppressed else ()
            similarity = f"{warning.similarity * 100:.0f}%"
            locations = "\n".join(warning.locations)
            row = {
                "suppressed": suppressed,
                "employee": warning.employee,
                "similar_employee": warning.similar_employee,
                "similarity": similarity,
                "locations": locations,
                "__warning__": warning,
                "__suppression_key__": suppression_key,
                "__tags__": tags,
            }
            name_typo_rows.append(row)
            data_review_rows.append(
                {
                    "suppressed": suppressed,
                    "category": "Name Typo",
                    "employee": warning.employee,
                    "source_file": "",
                    "sheet": "",
                    "date": "",
                    "summary": f"{warning.employee} -> {warning.similar_employee} ({similarity} similar)",
                    "details": locations,
                    "similar_employee": warning.similar_employee,
                    "__warning__": warning,
                    "__suppression_key__": suppression_key,
                    "__tags__": tags,
                    "__review_table_id__": "name_typos",
                }
            )
        self.pages["name_typos"].set_rows(name_typo_rows)
        self.pages["data_review"].set_rows(data_review_rows)
        self.pages["audit_data_trail"].set_rows([
            {
                "source_file": record.source_file,
                "pf_number": core.display_pf_number(record.pf_number),
                "date": record.work_date.isoformat(),
                "sheet": record.source_sheet,
                "employee": record.employee,
                "st": core.fmt_hours(record.st),
                "ot": core.fmt_hours(record.ot),
                "dt": core.fmt_hours(record.dt),
                "total": core.fmt_hours(record.total),
                "expanded": core.fmt_hours(core.expanded_hours(record.st, record.ot, record.dt)),
                "source_ranges": record.source_ranges,
                "audit": f"{record.source_sheet} -> {record.source_ranges}",
                "__record__": record,
            }
            for record in sorted(filtered_records, key=lambda item: (item.work_date, item.source_sheet, item.employee), reverse=True)
        ])
        self._refresh_email_preview(filtered_records)
        self._refresh_filters()
        unsuppressed_error_count = sum(1 for item in findings if not self._is_row_suppressed("error_report", core.error_finding_suppression_key(item)))
        unsuppressed_parse_count = sum(
            1
            for warning in active_parse_warnings
            if warning.source_file in filtered_source_files
            and not self._is_row_suppressed("parse_warnings", core.sheet_parse_warning_suppression_key(warning))
        )
        unsuppressed_typo_count = sum(
            1
            for warning in name_typo_warnings
            if not self._is_row_suppressed("name_typos", core.typo_warning_key(warning.employee, warning.similar_employee))
        )
        unsuppressed_workbook_health_count = sum(
            1
            for item in self.current_data.workbook_health
            if item.source_file in filtered_source_files
            and not self._is_row_suppressed("workbook_health", core.workbook_health_suppression_key(item))
        )
        self._sync_reports_alert_chrome(
            bool(unsuppressed_error_count),
            bool(unsuppressed_parse_count),
            bool(unsuppressed_typo_count),
            bool(unsuppressed_workbook_health_count),
        )
        self._refresh_overview_labels()

    def _refresh_email_preview(self, filtered_records: list[core.DailyRecord]) -> None:
        week_start = self.email_drafts_page.selected_week_start()
        if week_start is None:
            self.email_drafts_page.preview_table.set_rows([])
            return
        requests = core.build_email_draft_requests(filtered_records, self.employee_emails, week_start)
        rows: list[dict[str, Any]] = []
        for request in requests:
            st_total = round(sum(record.st for record in request.records), 2)
            ot_total = round(sum(record.ot for record in request.records), 2)
            dt_total = round(sum(record.dt for record in request.records), 2)
            pf_numbers = core.pf_numbers_for_records(request.records)
            suppressed = request.employee in self.missing_email_suppressions
            display_email, missing = core.format_email_address_display(request.email, suppressed=suppressed)
            tags = ("missing_email",) if missing else ()
            rows.append(
                {
                    "employee": request.employee,
                    "email": display_email,
                    "pf_numbers": pf_numbers or core.display_pf_number(""),
                    "days": str(len(request.records)),
                    "st": core.fmt_hours(st_total),
                    "ot": core.fmt_hours(ot_total),
                    "dt": core.fmt_hours(dt_total),
                    "total": core.fmt_hours(round(st_total + ot_total + dt_total, 2)),
                    "expanded": core.fmt_hours(core.expanded_hours(st_total, ot_total, dt_total)),
                    "__tags__": tags,
                }
            )
        self.email_drafts_page.preview_table.set_rows(rows)

    def _save_table_layout(self, table_id: str, visible_columns: list[str], widths: dict[str, int], sort_column: str, sort_descending: bool) -> None:
        core.save_table_layout(
            self.config_path,
            table_id,
            visible_columns,
            widths,
            sort_column=sort_column,
            sort_descending=sort_descending,
        )

    def _employees_changed(self) -> None:
        employee_names, hidden_names, emails, notes, groups, suppressions = self.employees_page.snapshot()
        managed = set(employee_names)
        discovered = {self._display_employee_name(name) for name in (self.current_data.employee_names if self.current_data else [])}
        self.employee_added_names = managed - discovered
        self.employee_hidden_names = set(hidden_names)
        self.employee_emails = emails
        self.employee_notes = notes
        self.employee_groups = groups
        self.missing_email_suppressions = suppressions
        core.save_employee_name_overrides(self.config_path, self.employee_added_names, self.employee_hidden_names)
        core.save_employee_emails(self.config_path, self.employee_emails, self.employee_outlook_display_names)
        core.save_employee_notes(self.config_path, self.employee_notes)
        core.save_employee_groups(self.config_path, self.employee_groups)
        core.save_missing_email_suppressions(self.config_path, self.missing_email_suppressions)
        self._invalidate_name_typo_cache()
        self._refresh_filters()
        self._refresh_overview_labels()
        self.refresh_views()

    def _formatting_changed(self) -> None:
        self.profiles, self.current_profile_name = self.formatting_page.snapshot()
        core.save_formatting_profiles(self.config_path, self.profiles, self.current_profile_name)
        self._refresh_overview_labels()
        self.refresh_views()

    def _settings_changed(self) -> None:
        try:
            self.app_settings = self.configuration_page.snapshot(self.app_settings)
        except ValueError as exc:
            self._set_update_status(f"Settings error: {exc}")
            return
        core.save_app_settings(self.config_path, self.app_settings)
        core.save_last_open_dss_paths(self.config_path, self.source_paths)
        self._sync_data_tabs_visibility()
        self.hash_poll_timer.setInterval(max(1, self.app_settings.hash_poll_minutes) * 60 * 1000)
        self._bind_shortcuts()
        self._apply_ui_theme_chrome()
        self._refresh_overview_labels()
        self.refresh_views()

    def apply_settings(self) -> None:
        self._settings_changed()
        QMessageBox.information(self, "Configuration", "Settings applied.")

    def _set_loading_state(self, loading: bool, message: str = "") -> None:
        load_active = self.load_worker is not None
        outlook_active = self.outlook_worker is not None
        typo_active = self.name_typo_worker is not None
        busy = load_active or outlook_active or typo_active
        ui_locked = load_active or typo_active
        has_sources = bool(self.source_paths)
        self.add_button.setEnabled(not busy)
        self.quick_add_button.setEnabled(not busy)
        self.remove_button.setEnabled(has_sources and not busy)
        self.update_button.setEnabled(has_sources and not busy)
        self.employee_filter.setEnabled(not ui_locked)
        self.pf_filter.setEnabled(not ui_locked)
        self.group_tabs.setEnabled(True)
        self.cancel_button.setEnabled(busy)
        self.loading_label.setText("Working..." if busy and not self.loading_label.text().strip() else ("" if not busy else self.loading_label.text()))
        self._refresh_quickload_hint_label()
        self._refresh_source_status_labels()
        self._refresh_export_button_state(busy=ui_locked)
        if message:
            self.status_label.setText(message)

    def _refresh_export_button_state(self, busy: bool | None = None) -> None:
        if busy is None:
            busy = self.load_worker is not None or self.outlook_worker is not None or self.name_typo_worker is not None
        table = self._current_table_page()
        self.export_button.setEnabled(table is not None and not busy)

    def _configured_library_root(self) -> Path | None:
        raw = self.app_settings.dss_library_root.strip()
        if not raw:
            QMessageBox.information(self, "Quick DSS Picker", "Set the DSS library root folder first in Settings -> Configuration.")
            return None
        root = Path(raw).expanduser()
        if not root.exists() or not root.is_dir():
            QMessageBox.warning(self, "Quick DSS Picker", f"The configured DSS library root does not exist:\n{root}")
            return None
        return root.resolve()

    def _show_quick_dss_picker(self, title: str, preselected: Iterable[Path] | None = None) -> list[Path]:
        root = self._configured_library_root()
        if root is None:
            return []
        dialog = QuickDssPickerDialog(self, root, selected_paths=preselected, title=title)
        if dialog.exec() != QDialog.Accepted:
            return []
        return [Path(path).expanduser().resolve() for path in dialog.selected_file_paths()]

    def quick_add_dss_files(self) -> None:
        paths = self._show_quick_dss_picker("Quick Add DSS Workbooks", preselected=self.source_paths)
        if not paths:
            return
        existing = {path.resolve() for path in self.source_paths}
        for path in paths:
            if path.resolve() not in existing:
                self.source_paths.append(path.resolve())
        core.save_last_open_dss_paths(self.config_path, self.source_paths)
        self.reload_data()

    def add_dss_files(self) -> None:
        paths, _filter = QFileDialog.getOpenFileNames(self, "Add DSS Workbook(s)", "", "Excel Workbook (*.xlsx)")
        if not paths:
            return
        existing = {path.resolve() for path in self.source_paths}
        for raw in paths:
            path = Path(raw).expanduser().resolve()
            if path not in existing:
                self.source_paths.append(path)
        core.save_last_open_dss_paths(self.config_path, self.source_paths)
        self.reload_data()

    def remove_dss_files(self) -> None:
        if not self.source_paths:
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Remove DSS(s)")
        layout = QVBoxLayout(dialog)
        label = QLabel("Select DSS workbooks to remove:")
        list_widget = QListWidget()
        for path in self.source_paths:
            item = QListWidgetItem(str(path))
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            list_widget.addItem(item)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        clear_all_button = buttons.addButton("Clear All", QDialogButtonBox.DestructiveRole)
        dialog.setProperty("clear_all_requested", False)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        def _request_clear_all() -> None:
            dialog.setProperty("clear_all_requested", True)
            dialog.accept()
        clear_all_button.clicked.connect(_request_clear_all)
        layout.addWidget(label)
        layout.addWidget(list_widget)
        layout.addWidget(buttons)
        result = dialog.exec()
        if result == 0:
            return
        if bool(dialog.property("clear_all_requested")):
            if QMessageBox.question(self, "Clear All DSSs", "Remove every loaded DSS workbook from the current session?") != QMessageBox.Yes:
                return
            self.source_paths = []
            core.save_last_open_dss_paths(self.config_path, self.source_paths)
            self.current_data = None
            self.progress_bar.setValue(0)
            self.percent_label.setText("0.0%")
            self._refresh_filters()
            self._refresh_employee_page()
            self.loading_label.setText("")
            self._set_loading_state(False, "No DSS workbooks loaded")
            self.refresh_views()
            return
        removed = {list_widget.item(row).text() for row in range(list_widget.count()) if list_widget.item(row).checkState() == Qt.Checked}
        self.source_paths = [path for path in self.source_paths if str(path) not in removed]
        core.save_last_open_dss_paths(self.config_path, self.source_paths)
        if self.source_paths:
            self.reload_data()
        else:
            self.current_data = None
            self.progress_bar.setValue(0)
            self.percent_label.setText("0.0%")
            self._refresh_filters()
            self._refresh_employee_page()
            self.loading_label.setText("")
            self._set_loading_state(False, "No DSS workbooks loaded")
            self.refresh_views()

    def reload_data(self, *, force_reparse: bool = False) -> None:
        if not self.source_paths:
            QMessageBox.information(self, "Open DSS", "Select one or more DSS workbooks first.")
            return
        self.cancel_active_work(abandon_ui=True, reset_ui=False, message="")
        if self.app_settings.quickload_last_sources_enabled and not self.current_data:
            last_paths = [path for path in core.load_last_open_dss_paths(self.config_path) if path.exists()]
            self._quickload_session = bool(last_paths) and self.source_paths == last_paths
        else:
            self._quickload_session = False
        self._hash_alerted_paths.clear()
        self.progress_bar.setValue(0)
        self.percent_label.setText("0.0%")
        self.loading_label.setText("Loading...")
        self._refresh_quickload_hint_label()
        self._next_load_token += 1
        token = self._next_load_token
        self._active_load_token = token
        worker = LoadWorker(
            self.source_paths,
            self.current_data,
            self.cache_dir,
            max_parallel_parse_workers=self.app_settings.max_parallel_parse_workers,
            partial_preview_enabled=self.app_settings.partial_preview_enabled,
            force_reparse=force_reparse,
        )
        self.load_worker = worker
        worker.progressChanged.connect(lambda fraction, message, current_token=token: self._on_load_progress(current_token, fraction, message))
        worker.partialReady.connect(lambda tracker_data, message, current_token=token: self._on_partial_ready(current_token, tracker_data, message))
        worker.finished.connect(lambda tracker_data, current_token=token: self._on_load_finished(current_token, tracker_data))
        worker.failed.connect(lambda message, current_token=token: self._on_load_failed(current_token, message))
        worker.cancelled.connect(lambda current_token=token: self._on_load_cancelled(current_token))
        self._set_loading_state(True, f"{len(self.source_paths)} DSS workbook(s) loading")
        self.load_thread = threading.Thread(target=worker.run, daemon=True)
        self.load_thread.start()

    def cancel_active_work(self, abandon_ui: bool = True, reset_ui: bool = True, message: str = "Cancelling...") -> bool:
        had_active = False
        if self.load_worker is not None:
            had_active = True
            self.load_worker.cancel()
            if abandon_ui:
                self._active_load_token = -1
                self.load_worker = None
                self.load_thread = None
        if self.outlook_worker is not None:
            had_active = True
            self.outlook_worker.cancel()
            if abandon_ui:
                self._active_outlook_token = -1
                self.outlook_worker = None
                self.outlook_thread = None
        if self.name_typo_worker is not None:
            had_active = True
            self.name_typo_worker.cancel()
            if abandon_ui:
                self._active_name_typo_token = -1
                self.name_typo_worker = None
                self.name_typo_thread = None
        if had_active and message:
            self.status_label.setText(message)
        if had_active and abandon_ui and reset_ui:
            self._set_loading_state(False, "Cancelled. Any stuck background copy/read will be ignored.")
        return had_active

    def _on_load_progress(self, token: int, fraction: float, message: str) -> None:
        if token != self._active_load_token:
            return
        value = max(0, min(1000, int(round(fraction * 1000))))
        self.progress_bar.setValue(value)
        self.percent_label.setText(f"{fraction * 100:.1f}%")
        self.loading_label.setText("Loading...")
        self.status_label.setText(message)

    def _on_partial_ready(self, token: int, tracker_data: core.TrackerData, message: str) -> None:
        if token != self._active_load_token:
            return
        self.current_data = tracker_data
        self._invalidate_name_typo_cache()
        self.status_label.setText(message)
        self._refresh_source_status_labels()
        self._refresh_filters()
        self._refresh_employee_page()
        self._refresh_overview_labels()
        if self.group_tabs.currentWidget() != self.settings_tabs:
            self._queue_refresh_views()

    def _on_load_finished(self, token: int, tracker_data: core.TrackerData) -> None:
        if token != self._active_load_token:
            return
        self.current_data = tracker_data
        self._invalidate_name_typo_cache()
        self.progress_bar.setValue(1000)
        self.percent_label.setText("100.0%")
        self.loading_label.setText("")
        self.load_worker = None
        self.load_thread = None
        self._quickload_session = False
        self._set_loading_state(False, f"Loaded {len(tracker_data.source_paths)} DSS workbook(s)")
        self._refresh_source_status_labels()
        self._refresh_filters()
        self._refresh_employee_page()
        self._refresh_overview_labels()
        self._queue_refresh_views()
        if tracker_data.reloaded_paths or tracker_data.reused_paths:
            QMessageBox.information(
                self,
                "Update View",
                f"Reloaded: {len(tracker_data.reloaded_paths)}\nUnchanged: {len(tracker_data.reused_paths)}",
            )

    def _on_load_failed(self, token: int, message: str) -> None:
        if token != self._active_load_token:
            return
        self.load_worker = None
        self.load_thread = None
        self.loading_label.setText("")
        self._quickload_session = False
        self._set_loading_state(False, "Load failed")
        QMessageBox.critical(self, "Failed to open workbook", message)

    def _on_load_cancelled(self, token: int) -> None:
        if token != self._active_load_token:
            return
        self.load_worker = None
        self.load_thread = None
        self.progress_bar.setValue(0)
        self.percent_label.setText("0.0%")
        self.loading_label.setText("")
        self._quickload_session = False
        self._set_loading_state(False, "Load cancelled")

    def _resolve_source_path_by_name(self, source_display_name: str) -> Path | None:
        if not source_display_name or self.current_data is None:
            return None
        for path in self.current_data.source_paths:
            if path.name == source_display_name:
                return path
        return None

    def _handle_table_row_activated(self, row: dict[str, Any], column_key: str) -> None:
        finding = row.get("__finding__")
        if isinstance(finding, core.ErrorFinding):
            self._open_error_finding_in_excel(finding)
            return
        warning = row.get("__warning__")
        if isinstance(warning, core.SheetParseWarning):
            self._open_parse_warning_in_excel(warning)
            return
        if isinstance(warning, core.NameTypoWarning):
            self._open_name_typo_warning_in_excel(warning)
            return
        record = row.get("__record__")
        if isinstance(record, core.DailyRecord):
            self._open_record_ranges_in_excel(record)
            return
        if column_key in {"source_file", "sources"}:
            source_display = str(row.get(column_key, "")).strip()
            if source_display:
                self._open_displayed_source_file(source_display)

    def _open_record_ranges_in_excel(self, record: core.DailyRecord, *, prefer_name_only: bool = False) -> None:
        ranges = core.selection_ranges_for_source_ranges(record.source_ranges, prefer_name_only=prefer_name_only)
        try:
            core.select_excel_workbook_ranges(record.source_path, record.source_sheet, ranges)
        except Exception as exc:
            QMessageBox.critical(self, "Open Excel Range", str(exc))

    def _open_parse_warning_in_excel(self, warning: core.SheetParseWarning) -> None:
        source_path = self._resolve_source_path_by_name(warning.source_file)
        if source_path is None:
            QMessageBox.critical(self, "Open Excel Range", f"Could not find workbook:\n{warning.source_file}")
            return
        ranges: list[str] = []
        if warning.issue == "Revision Indicator AZ2 Mismatch":
            ranges = ["AZ2"]
        elif "(" in warning.details and ")" in warning.details:
            between = warning.details.split("(", 1)[1].split(")", 1)[0]
            ranges = core.selection_ranges_for_source_ranges(between)
        if not ranges:
            self._open_displayed_source_file(warning.source_file)
            return
        try:
            core.select_excel_workbook_ranges(source_path, warning.source_sheet, ranges)
        except Exception as exc:
            QMessageBox.critical(self, "Open Excel Range", str(exc))

    def _open_name_typo_warning_in_excel(self, warning: core.NameTypoWarning) -> None:
        if self.current_data is None:
            return
        target_record: core.DailyRecord | None = None
        for location in warning.locations:
            parts = [part.strip() for part in location.split("|")]
            if len(parts) != 3:
                continue
            date_text, sheet_name, source_file = parts
            for record in self.current_data.daily_records:
                if (
                    record.employee == warning.employee
                    and record.work_date.isoformat() == date_text
                    and record.source_sheet == sheet_name
                    and record.source_file == source_file
                ):
                    target_record = record
                    break
            if target_record is not None:
                break
        if target_record is None:
            for record in self.current_data.daily_records:
                if record.employee == warning.employee:
                    target_record = record
                    break
        if target_record is None:
            QMessageBox.critical(self, "Open Excel Range", f"Could not find source rows for {warning.employee}.")
            return
        self._open_record_ranges_in_excel(target_record, prefer_name_only=True)

    def _open_error_finding_in_excel(self, finding: core.ErrorFinding) -> None:
        if self.current_data is None:
            return
        source_names = {part.strip() for part in finding.source_files.split(",") if part.strip()}
        matching_records = [
            record
            for record in self.current_data.daily_records
            if record.employee == finding.employee
            and record.work_date == finding.trigger_date
            and (not source_names or record.source_file in source_names)
        ]
        if not matching_records:
            QMessageBox.critical(
                self,
                "Open Excel Range",
                f"Could not find source rows for {finding.employee} on {finding.trigger_date.isoformat()}.",
            )
            return
        matching_records.sort(key=lambda record: (record.source_file.casefold(), record.source_sheet.casefold(), record.source_ranges))
        target = matching_records[0]
        same_sheet_records = [
            record
            for record in matching_records
            if record.source_path == target.source_path and record.source_sheet == target.source_sheet
        ]
        ranges: list[str] = []
        if finding.outlook_name_rule:
            for record in same_sheet_records:
                ranges.extend(core.selection_ranges_for_source_ranges(record.source_ranges, prefer_name_only=True))
        else:
            for record in same_sheet_records:
                ranges.extend(core.selection_ranges_for_source_ranges(record.source_ranges))
        deduped_ranges = list(dict.fromkeys(ranges))
        if not deduped_ranges:
            deduped_ranges = core.selection_ranges_for_source_ranges(target.source_ranges, prefer_name_only=finding.outlook_name_rule)
        try:
            core.select_excel_workbook_ranges(target.source_path, target.source_sheet, deduped_ranges)
        except Exception as exc:
            QMessageBox.critical(self, "Open Excel Range", str(exc))

    def save_email_templates(self) -> None:
        self.subject_template = self.email_drafts_page.subject_edit.toPlainText().strip()
        self.body_template = self.email_drafts_page.body_edit.toPlainText().strip()
        core.save_email_templates(self.config_path, self.subject_template, self.body_template)
        QMessageBox.information(self, "Email Drafts", "Email templates saved.")

    def sync_outlook_emails(self) -> None:
        employee_names = [
            employee
            for employee in self._managed_employee_names()
            if not self.employee_emails.get(employee, "").strip() and employee not in self.missing_email_suppressions
        ]
        if not employee_names:
            QMessageBox.information(self, "Outlook Email Sync", "No missing employee emails need syncing right now.")
            return
        query_names, cached_resolutions, cached_display_names, skipped_recent_misses = core.plan_outlook_query_names(
            employee_names,
            self.employee_emails,
            self.employee_outlook_display_names,
            self.outlook_lookup_cache,
        )
        cache_applied = 0
        updated = False
        for employee, resolution in cached_resolutions.items():
            if resolution.email.strip() and not self.employee_emails.get(employee, "").strip():
                self.employee_emails[employee] = resolution.email.strip()
                cache_applied += 1
                updated = True
            if resolution.display_name.strip() and self.employee_outlook_display_names.get(employee, "") != resolution.display_name.strip():
                self.employee_outlook_display_names[employee] = resolution.display_name.strip()
                updated = True
        for employee, display_name in cached_display_names.items():
            if display_name.strip() and self.employee_outlook_display_names.get(employee, "") != display_name.strip():
                self.employee_outlook_display_names[employee] = display_name.strip()
                updated = True
        if updated:
            core.save_employee_emails(self.config_path, self.employee_emails, self.employee_outlook_display_names)
            core.save_outlook_lookup_cache(self.config_path, self.outlook_lookup_cache)
            self._refresh_employee_page()
            self.refresh_views()
        if not query_names:
            still_missing = sum(
                1
                for employee in self._managed_employee_names()
                if not self.employee_emails.get(employee, "").strip() and employee not in self.missing_email_suppressions
            )
            QMessageBox.information(
                self,
                "Outlook Email Sync",
                f"Used cached Outlook results for {cache_applied} employee(s).\n"
                f"Skipped {len(skipped_recent_misses)} recently checked unresolved name(s).\n"
                f"Still missing: {still_missing}",
            )
            return
        self.cancel_active_work(abandon_ui=True, reset_ui=False, message="")
        self._next_outlook_token += 1
        token = self._next_outlook_token
        self._active_outlook_token = token
        self._outlook_partial_updates = {}
        self._outlook_last_partial_refresh = 0.0
        worker = OutlookWorker(query_names)
        self.outlook_worker = worker
        worker.progressChanged.connect(
            lambda processed, total, employee, resolution, current_token=token: self._on_outlook_sync_progress(
                current_token,
                processed,
                total,
                employee,
                resolution,
            )
        )
        worker.finished.connect(
            lambda results, address_book_names, current_token=token, query_names=query_names, cache_applied=cache_applied, skipped_recent_misses=skipped_recent_misses:
            self._on_outlook_sync_finished(current_token, results, address_book_names, query_names, cache_applied, skipped_recent_misses)
        )
        worker.failed.connect(lambda message, current_token=token: self._on_outlook_sync_failed(current_token, message))
        worker.cancelled.connect(lambda current_token=token: self._on_outlook_sync_cancelled(current_token))
        self.outlook_thread = threading.Thread(target=worker.run, daemon=True)
        self.progress_bar.setRange(0, 1000)
        self.progress_bar.setValue(0)
        self.percent_label.setText("0.0%")
        self.loading_label.setText("Syncing emails...")
        self._set_loading_state(True, "Syncing Outlook emails...")
        self.outlook_thread.start()

    def _flush_outlook_partial_updates(self, force: bool = False) -> None:
        if not self._outlook_partial_updates:
            return
        now = time.monotonic()
        if not force and (now - self._outlook_last_partial_refresh) < 1.0:
            return
        updated = False
        for employee, resolution in list(self._outlook_partial_updates.items()):
            if resolution.email and not self.employee_emails.get(employee, "").strip():
                self.employee_emails[employee] = resolution.email
                updated = True
            if resolution.display_name.strip() and self.employee_outlook_display_names.get(employee, "") != resolution.display_name.strip():
                self.employee_outlook_display_names[employee] = resolution.display_name.strip()
                updated = True
        self._outlook_partial_updates.clear()
        self._outlook_last_partial_refresh = now
        if updated:
            self._invalidate_name_typo_cache()
            self._refresh_employee_page()
            self.refresh_views()

    def _on_name_typo_refresh_progress(self, token: int, fraction: float, message: str) -> None:
        if token != self._active_name_typo_token:
            return
        value = max(0, min(1000, int(round(fraction * 1000))))
        self.progress_bar.setRange(0, 1000)
        self.progress_bar.setValue(value)
        self.percent_label.setText(f"{fraction * 100:.1f}%")
        self.loading_label.setText("Checking names...")
        self.status_label.setText(message)

    def _on_outlook_sync_progress(
        self,
        token: int,
        processed: int,
        total: int,
        employee: str,
        resolution: object,
    ) -> None:
        if token != self._active_outlook_token:
            return
        fraction = processed / max(total, 1)
        self.progress_bar.setRange(0, 1000)
        self.progress_bar.setValue(max(0, min(1000, int(round(fraction * 1000)))))
        self.percent_label.setText(f"{fraction * 100:.1f}%")
        self.loading_label.setText("Syncing emails...")
        progress_text = f"Syncing Outlook emails... {processed}/{total}"
        if employee:
            if isinstance(resolution, core.OutlookResolution) and resolution.email.strip():
                progress_text += f" | {employee} -> {resolution.email}"
                self._outlook_partial_updates[employee] = resolution
            else:
                progress_text += f" | {employee}"
        self.status_label.setText(progress_text)
        self._flush_outlook_partial_updates()

    def _on_outlook_sync_finished(
        self,
        token: int,
        results: dict[str, core.OutlookResolution],
        address_book_names: list[str],
        queried_names: list[str],
        cache_applied: int,
        skipped_recent_misses: set[str],
    ) -> None:
        if token != self._active_outlook_token:
            return
        self.outlook_worker = None
        self.outlook_thread = None
        self._flush_outlook_partial_updates(force=True)
        self.outlook_lookup_cache = core.update_outlook_lookup_cache(self.outlook_lookup_cache, queried_names, results)
        updated = 0
        for employee, resolution in results.items():
            if resolution.email and not self.employee_emails.get(employee, "").strip():
                self.employee_emails[employee] = resolution.email
                if resolution.display_name.strip():
                    self.employee_outlook_display_names[employee] = resolution.display_name.strip()
                updated += 1
        if updated:
            self._invalidate_name_typo_cache()
        core.save_employee_emails(self.config_path, self.employee_emails, self.employee_outlook_display_names)
        core.save_outlook_lookup_cache(self.config_path, self.outlook_lookup_cache)
        if self.load_worker is None:
            self.progress_bar.setRange(0, 1000)
            self.progress_bar.setValue(1000)
            self.percent_label.setText("100.0%")
            self.loading_label.setText("")
            self._set_loading_state(False, f"Matched emails: {updated + cache_applied}")
        self._refresh_employee_page()
        warnings: list[core.NameTypoWarning] = []
        if not self.app_settings.disable_name_typo_notifications:
            warnings = self._build_name_typo_warnings(self._managed_employee_names(), self.current_data.daily_records if self.current_data else [], address_book_names)
        self._set_cached_name_typo_warnings(warnings)
        self._refresh_overview_labels()
        self._queue_refresh_views()
        missing_after = sum(
            1
            for employee in self._managed_employee_names()
            if not self.employee_emails.get(employee, "").strip() and employee not in self.missing_email_suppressions
        )
        unsuppressed_warnings = [
            warning
            for warning in warnings
            if core.typo_warning_key(warning.employee, warning.similar_employee) not in self.ignored_name_typos
        ]
        message = (
            f"Matched emails from Outlook: {updated}\n"
            f"Matched emails from cache: {cache_applied}\n"
            f"Skipped recent unresolved names: {len(skipped_recent_misses)}\n"
            f"Still missing: {missing_after}"
        )
        if unsuppressed_warnings and not self.app_settings.disable_name_typo_notifications:
            message += f"\nName typo warnings: {len(unsuppressed_warnings)} (see Reports > Name Typos)"
        QMessageBox.information(self, "Outlook Email Sync", message)

    def _on_outlook_sync_failed(self, token: int, message: str) -> None:
        if token != self._active_outlook_token:
            return
        self.outlook_worker = None
        self.outlook_thread = None
        self._outlook_partial_updates.clear()
        if self.load_worker is None:
            self.loading_label.setText("")
            self._set_loading_state(False, "Outlook sync failed")
        QMessageBox.critical(self, "Outlook Email Sync", message)

    def _on_outlook_sync_cancelled(self, token: int) -> None:
        if token != self._active_outlook_token:
            return
        self.outlook_worker = None
        self.outlook_thread = None
        self._outlook_partial_updates.clear()
        if self.load_worker is None:
            self.loading_label.setText("")
            self._set_loading_state(False, "Outlook sync cancelled")

    def _current_filtered_records(self) -> list[core.DailyRecord]:
        return self._daily_records_filtered()

    def create_outlook_drafts(self) -> None:
        week_start = self.email_drafts_page.selected_week_start()
        if not self.current_data or week_start is None:
            QMessageBox.information(self, "Email Drafts", "Load DSS data and choose a week first.")
            return
        requests = core.build_email_draft_requests(self._current_filtered_records(), self.employee_emails, week_start)
        selected_preview_rows = self.email_drafts_page.preview_table.selected_rows()
        selected_employees = {row.get("employee", "") for row in selected_preview_rows}
        if selected_employees:
            requests = [request for request in requests if request.employee in selected_employees]
        try:
            created, skipped = core.create_outlook_drafts(requests, self.subject_template, self.body_template)
        except Exception as exc:
            QMessageBox.critical(self, "Email Drafts", str(exc))
            return
        message = f"Created drafts: {created}"
        if skipped:
            message += f"\nSkipped missing emails: {', '.join(skipped)}"
        QMessageBox.information(self, "Email Drafts", message)

    def _current_table_page(self) -> DataTablePage | None:
        widget = self.group_tabs.currentWidget()
        if isinstance(widget, QTabWidget):
            inner = widget.currentWidget()
            if isinstance(inner, DataTablePage):
                return inner
            if inner is self.email_drafts_page:
                return self.email_drafts_page.preview_table
        return None

    def export_current_view(self) -> None:
        table = self._current_table_page()
        if table is None:
            QMessageBox.information(self, "Export Current View", "Open a table-based page first.")
            return
        export_path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Export Current View",
            "",
            "Excel Workbook (*.xlsx);;CSV (*.csv)",
        )
        if not export_path:
            return
        columns, rows = table.export_rows()
        visible_pairs = [pair for pair in table.columns if pair[0] in columns]
        if selected_filter.endswith(".csv") or export_path.lower().endswith(".csv"):
            with open(export_path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow([label for _key, label in visible_pairs])
                for row in rows:
                    writer.writerow([row.get(key, "") for key, _label in visible_pairs])
        else:
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = table.spec.title[:31]
            worksheet.append([label for _key, label in visible_pairs])
            for row in rows:
                worksheet.append([row.get(key, "") for key, _label in visible_pairs])
            workbook.save(export_path)
        QMessageBox.information(self, "Export Current View", f"Exported:\n{export_path}")

    def clear_cached_dss(self) -> None:
        deleted = core.clear_cache_files(self.cache_dir)
        if self.current_data is not None:
            self.current_data = core.tracker_data_invalidated_for_cache_clear(self.current_data)
        QMessageBox.information(self, "Configuration", f"Deleted {deleted} cached DSS file(s).")

    def clear_stored_emails(self) -> None:
        self.employee_emails = {}
        self.employee_outlook_display_names = {}
        self.outlook_lookup_cache = {}
        core.remove_config_keys(self.config_path, ["employee_emails", "employee_outlook_display_names", core.OUTLOOK_LOOKUP_CACHE_KEY])
        self._refresh_employee_page()
        self.refresh_views()
        QMessageBox.information(self, "Configuration", "Stored employee emails were cleared.")

    def clear_all_stored_data(self) -> None:
        if QMessageBox.question(self, "Clear All Stored Data", "Delete cached DSS data, emails, groups, notes, templates, layouts, and app settings?") != QMessageBox.Yes:
            return
        core.clear_cache_files(self.cache_dir)
        if self.config_path.exists():
            self.config_path.unlink(missing_ok=True)
        self.current_data = None
        self.source_paths = []
        self.app_settings = core.AppSettings()
        self.profiles = core.default_formatting_profiles()
        self.current_profile_name = core.DEFAULT_PROFILE_NAME
        self.employee_emails = {}
        self.employee_outlook_display_names = {}
        self.outlook_lookup_cache = {}
        self.employee_notes = {}
        self.employee_groups = {}
        self.missing_email_suppressions = set()
        self.employee_added_names = set()
        self.employee_hidden_names = set()
        self.employee_name_merges = {}
        self.subject_template, self.body_template = core.load_email_templates(self.config_path)
        self.ignored_name_typos = set()
        self.suppressed_error_findings = set()
        self.suppressed_parse_warnings = set()
        self.suppressed_workbook_health = set()
        self._invalidate_name_typo_cache()
        self.status_label.setText("No DSS workbooks loaded")
        self.progress_bar.setValue(0)
        self.percent_label.setText("0.0%")
        self._hash_alerted_paths.clear()
        self._set_update_status(f"Installed version: {core.APP_VERSION}")
        self._sync_ui_from_state()

    def reset_settings(self) -> None:
        self.app_settings = core.AppSettings()
        self.profiles = core.default_formatting_profiles()
        self.current_profile_name = core.DEFAULT_PROFILE_NAME
        core.save_app_settings(self.config_path, self.app_settings)
        core.save_formatting_profiles(self.config_path, self.profiles, self.current_profile_name)
        self.configuration_page.set_settings(self.app_settings)
        self.formatting_page.set_data(self.profiles, self.current_profile_name)
        self._sync_data_tabs_visibility()
        self.hash_poll_timer.setInterval(max(1, self.app_settings.hash_poll_minutes) * 60 * 1000)
        self._bind_shortcuts()
        self._apply_ui_theme_chrome()
        self.refresh_views()

    def export_diagnostic_snapshot(self) -> None:
        try:
            export_path = self._write_diagnostic_snapshot()
        except OSError as exc:
            QMessageBox.critical(self, "Diagnostic Snapshot", f"Could not write the diagnostic snapshot.\n\n{exc}")
            return
        QMessageBox.information(self, "Diagnostic Snapshot", f"Saved diagnostic snapshot:\n{export_path}")

    def _write_diagnostic_snapshot(self) -> Path:
        snapshot: dict[str, Any] = {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "app_name": core.DISPLAY_APP_NAME,
            "app_root": str(self.app_root),
            "config_path": str(self.config_path),
            "source_paths": [str(path) for path in self.source_paths],
            "cache_dir": str(self.cache_dir),
            "python_version": sys.version,
            "app_version": core.APP_VERSION,
            "os_name": os.name,
            "app_settings": {
                "disable_name_typo_notifications": self.app_settings.disable_name_typo_notifications,
                "hash_poll_minutes": self.app_settings.hash_poll_minutes,
                "update_check_delay_seconds": self.app_settings.update_check_delay_seconds,
                "show_daily_raw_tab": self.app_settings.show_daily_raw_tab,
                "quickload_last_sources_enabled": self.app_settings.quickload_last_sources_enabled,
                "quickload_cancel_hotkey": self.app_settings.quickload_cancel_hotkey,
                "auto_update_check_enabled": self.app_settings.auto_update_check_enabled,
                "auto_download_updates_on_unmetered_wifi": self.app_settings.auto_download_updates_on_unmetered_wifi,
                "max_parallel_parse_workers": self.app_settings.max_parallel_parse_workers,
                "partial_preview_enabled": self.app_settings.partial_preview_enabled,
                "ui_theme": core.asdict(self.app_settings.ui_theme),
            },
            "formatting_profiles": sorted(self.profiles),
            "current_profile_name": self.current_profile_name,
            "employee_email_count": len([email for email in self.employee_emails.values() if email.strip()]),
            "missing_email_suppression_count": len(self.missing_email_suppressions),
            "employee_group_count": len(self.employee_groups),
            "cache_file_count": len(list(self.cache_dir.glob("*.json"))),
            "update_download_dir": str(self.updates_dir),
            "downloaded_update_path": str(self._downloaded_update_path) if self._downloaded_update_path else "",
            "loaded_dss": [],
        }
        if self.current_data is not None:
            for source_path in self.current_data.source_paths:
                snapshot["loaded_dss"].append(
                    {
                        "path": str(source_path),
                        "hash": self.current_data.file_hashes.get(source_path, ""),
                        "reused": source_path in self.current_data.reused_paths,
                        "reloaded": source_path in self.current_data.reloaded_paths,
                        "hash_alerted": source_path in self._hash_alerted_paths,
                    }
                )
            snapshot["current_data"] = {
                "daily_record_count": len(self.current_data.daily_records),
                "employee_count": len(self.current_data.employee_names),
                "week_count": len(self.current_data.week_totals),
                "source_count": len(self.current_data.source_paths),
            }
        else:
            snapshot["current_data"] = None
        export_path = self.app_root / f"diagnostic_snapshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        export_path.write_text(json.dumps(snapshot, indent=2), encoding="utf-8")
        return export_path

    def submit_bug_report(self) -> None:
        snapshot_path = self._write_diagnostic_snapshot()
        html_body = core.build_bug_report_html(
            self.current_profile_name,
            self.app_root,
            snapshot_path,
            self.source_paths,
            self.current_data.cache_status_by_path if self.current_data else {},
        )
        try:
            warning = core.create_bug_report_draft(
                core.BUG_REPORT_EMAIL,
                f"{core.DISPLAY_APP_NAME} Bug Report",
                html_body,
                attachment_path=snapshot_path,
            )
        except Exception as exc:
            QMessageBox.critical(self, "Submit Bug Report", str(exc))
            return
        message = f"Created an Outlook draft addressed to {core.BUG_REPORT_EMAIL}."
        if warning:
            message += f"\n\n{warning}"
        QMessageBox.information(self, "Submit Bug Report", message)

    def test_outlook_connection(self) -> None:
        if core.pythoncom is None or core.win32com is None:
            QMessageBox.critical(self, "Outlook Connection", "Outlook integration requires pywin32 and desktop Outlook.")
            return
        core.pythoncom.CoInitialize()
        try:
            outlook = core.win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            current_user = getattr(namespace, "CurrentUser", None)
            user_name = str(getattr(current_user, "Name", "")).strip() or "Unknown user"
        except Exception as exc:
            QMessageBox.critical(self, "Outlook Connection", f"Could not connect to Outlook.\n\n{exc}")
            return
        finally:
            core.pythoncom.CoUninitialize()
        QMessageBox.information(self, "Outlook Connection", f"Connected to desktop Outlook successfully.\nUser: {user_name}")

    def show_loaded_dss_status(self) -> None:
        if self.current_data is None or not self.current_data.source_paths:
            QMessageBox.information(self, "Loaded DSS Status", "No DSS workbooks are currently loaded.")
            return
        lines = [
            f"Loaded DSS files: {len(self.current_data.source_paths)}",
            f"Daily records: {len(self.current_data.daily_records)}",
            f"Employees: {len(self.current_data.employee_names)}",
            f"Weeks: {len(self.current_data.week_totals)}",
            "",
        ]
        for source_path in self.current_data.source_paths:
            status_bits: list[str] = []
            if source_path in self.current_data.reloaded_paths:
                status_bits.append("reloaded")
            if source_path in self.current_data.reused_paths:
                status_bits.append("reused")
            if source_path in self._hash_alerted_paths:
                status_bits.append("changed since load")
            if not status_bits:
                status_bits.append("loaded")
            lines.append(str(source_path))
            lines.append(f"Hash: {self.current_data.file_hashes.get(source_path, '')}")
            lines.append(f"Cache: {self.current_data.cache_status_by_path.get(source_path, 'Unknown')}")
            lines.append(f"Status: {', '.join(status_bits)}")
            lines.append("")
        QMessageBox.information(self, "Loaded DSS Status", "\n".join(lines).rstrip())

    def _persist_ignored_name_typos(self) -> None:
        core.save_ignored_name_typos(self.config_path, self.ignored_name_typos)

    def _persist_employee_name_merges(self) -> None:
        core.save_employee_name_merges(self.config_path, self.employee_name_merges)

    def _merge_employee_alias(self, source_name: str, target_name: str) -> None:
        source = source_name.strip()
        target = target_name.strip()
        if not source or not target or source == target:
            return
        resolved_target = core.resolve_employee_name_merge(target, self.employee_name_merges)
        if resolved_target == source:
            return
        self.employee_name_merges[source] = resolved_target
        if not self.employee_emails.get(resolved_target, "").strip() and self.employee_emails.get(source, "").strip():
            self.employee_emails[resolved_target] = self.employee_emails[source]
        if not self.employee_outlook_display_names.get(resolved_target, "").strip() and self.employee_outlook_display_names.get(source, "").strip():
            self.employee_outlook_display_names[resolved_target] = self.employee_outlook_display_names[source]
        source_cache = self.outlook_lookup_cache.get(source)
        target_cache = self.outlook_lookup_cache.get(resolved_target)
        if source_cache is not None:
            if target_cache is None or (not target_cache.email.strip() and source_cache.email.strip()):
                self.outlook_lookup_cache[resolved_target] = source_cache
        if not self.employee_notes.get(resolved_target, "").strip() and self.employee_notes.get(source, "").strip():
            self.employee_notes[resolved_target] = self.employee_notes[source]
        if source in self.missing_email_suppressions:
            self.missing_email_suppressions.add(resolved_target)
            self.missing_email_suppressions.discard(source)
        for group_name, members in list(self.employee_groups.items()):
            member_set = {self._display_employee_name(member) for member in members if member.strip()}
            if source in member_set:
                member_set.discard(source)
            if resolved_target:
                member_set.add(resolved_target)
            self.employee_groups[group_name] = sorted(member_set, key=str.casefold)
        self.employee_added_names = {self._display_employee_name(name) for name in self.employee_added_names if self._display_employee_name(name).strip()}
        self.employee_hidden_names.discard(source)
        self.employee_emails.pop(source, None)
        self.employee_outlook_display_names.pop(source, None)
        self.outlook_lookup_cache.pop(source, None)
        self.employee_notes.pop(source, None)
        self.ignored_name_typos.add(core.typo_warning_key(source, resolved_target))
        self._persist_employee_name_merges()
        core.save_employee_name_overrides(self.config_path, self.employee_added_names, self.employee_hidden_names)
        core.save_employee_emails(self.config_path, self.employee_emails, self.employee_outlook_display_names)
        core.save_outlook_lookup_cache(self.config_path, self.outlook_lookup_cache)
        core.save_employee_notes(self.config_path, self.employee_notes)
        core.save_employee_groups(self.config_path, self.employee_groups)
        core.save_missing_email_suppressions(self.config_path, self.missing_email_suppressions)
        self._persist_ignored_name_typos()
        self._invalidate_name_typo_cache()
        self._refresh_filters()
        self._refresh_employee_page()
        self._refresh_overview_labels()
        self.refresh_views()

    def check_name_typos_manually(self) -> None:
        if self.current_data is None or not self._managed_employee_names():
            QMessageBox.information(self, "Check Name Typos", "Load DSS data first.")
            return
        if self.name_typo_worker is not None:
            return
        cache_key = self._current_name_typo_cache_key()
        if cache_key is not None and cache_key == self._cached_name_typo_key:
            self.refresh_views()
            self._refresh_overview_labels()
            self.group_tabs.setCurrentWidget(self.report_tabs)
            self.report_tabs.setCurrentWidget(self.pages["data_review"])
            self._set_loading_state(False, f"Name typo review reused cached results: {len(self._cached_name_typo_warnings)} warning(s)")
            return
        employee_names = self._managed_employee_names()
        daily_records = self.current_data.daily_records
        self._next_name_typo_token += 1
        token = self._next_name_typo_token
        self._active_name_typo_token = token
        worker = NameTypoWorker(
            employee_names,
            daily_records,
            self.employee_emails,
            self.missing_email_suppressions,
            self.employee_outlook_display_names,
            self.outlook_lookup_cache,
        )
        self.name_typo_worker = worker
        worker.progressChanged.connect(lambda fraction, message, current_token=token: self._on_name_typo_refresh_progress(current_token, fraction, message))
        worker.finished.connect(lambda warnings, cache_updates, current_token=token: self._on_name_typo_refresh_finished(current_token, warnings, cache_updates))
        worker.failed.connect(lambda message, current_token=token: self._on_name_typo_refresh_failed(current_token, message))
        worker.cancelled.connect(lambda current_token=token: self._on_name_typo_refresh_cancelled(current_token))
        self.name_typo_thread = threading.Thread(target=worker.run, daemon=True)
        self.progress_bar.setRange(0, 1000)
        self.progress_bar.setValue(0)
        self.percent_label.setText("0.0%")
        self.loading_label.setText("Checking names...")
        self._set_loading_state(True, "Refreshing name typos...")
        self.name_typo_thread.start()

    def _on_name_typo_refresh_finished(
        self,
        token: int,
        warnings: list[core.NameTypoWarning],
        cache_updates: dict[str, core.OutlookLookupCacheEntry],
    ) -> None:
        if token != self._active_name_typo_token:
            return
        self.name_typo_worker = None
        self.name_typo_thread = None
        self.outlook_lookup_cache = dict(cache_updates)
        core.save_outlook_lookup_cache(self.config_path, self.outlook_lookup_cache)
        self._set_cached_name_typo_warnings(warnings)
        self.refresh_views()
        self._set_loading_state(False, f"Name typo refresh complete: {len(warnings)} warning(s)")
        self.group_tabs.setCurrentWidget(self.report_tabs)
        self.report_tabs.setCurrentWidget(self.pages["data_review"])
        if not warnings:
            QMessageBox.information(self, "Check Name Typos", "No likely name typos were found.")
            return
        unsuppressed = [
            warning
            for warning in warnings
            if core.typo_warning_key(warning.employee, warning.similar_employee) not in self.ignored_name_typos
        ]
        if not unsuppressed:
            QMessageBox.information(self, "Check Name Typos", "All current name typo warnings are already suppressed. Enable Show Suppressed on the Data Review page to review them.")

    def _on_name_typo_refresh_failed(self, token: int, message: str) -> None:
        if token != self._active_name_typo_token:
            return
        self.name_typo_worker = None
        self.name_typo_thread = None
        self._set_loading_state(False, "Name typo refresh failed")
        QMessageBox.critical(self, "Check Name Typos", message)

    def _on_name_typo_refresh_cancelled(self, token: int) -> None:
        if token != self._active_name_typo_token:
            return
        self.name_typo_worker = None
        self.name_typo_thread = None
        self._set_loading_state(False, "Name typo refresh cancelled")

    def show_app_data_folder(self) -> None:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.app_root)))

    def _auto_check_for_updates(self) -> None:
        if self._auto_update_check_done or not self.app_settings.auto_update_check_enabled:
            return
        self._auto_update_check_done = True
        self.check_for_updates(manual=False)

    def check_for_updates(self, manual: bool = True) -> None:
        if self._update_check_in_progress or self._update_download_in_progress:
            return
        self._update_check_in_progress = True
        if manual:
            self._set_update_status(f"Checking GitHub releases from version {core.APP_VERSION}...")
        else:
            self._set_update_status(f"Background update check from version {core.APP_VERSION}...")
        threading.Thread(target=self._check_for_updates_worker, args=(manual,), daemon=True).start()

    def _check_for_updates_worker(self, manual: bool) -> None:
        try:
            release_info = core.fetch_latest_release_info()
            latest_version = str(release_info.get("version", "")).strip()
            network_profile = core.get_windows_network_profile() if latest_version and core.is_newer_version(latest_version, core.APP_VERSION) else None
        except Exception as exc:
            self.updateCheckErrorRaised.emit(str(exc), manual)
            return
        self.updateCheckResultReady.emit(release_info, network_profile, manual)

    def _handle_update_check_error(self, message: str, manual: bool) -> None:
        self._update_check_in_progress = False
        self._set_update_status(f"Installed version: {core.APP_VERSION}")
        if manual:
            QMessageBox.critical(self, "Check for Updates", f"Could not check for updates.\n\n{message}")

    def _handle_update_check_result(self, release_info: dict[str, object], network_profile: dict[str, object] | None, manual: bool) -> None:
        self._update_check_in_progress = False
        latest_version = str(release_info.get("version", "")).strip()
        latest_tag = str(release_info.get("tag_name", "")).strip()
        html_url = str(release_info.get("html_url", "")).strip()
        published_at = str(release_info.get("published_at", "")).strip()
        asset_names = release_info.get("asset_names", [])
        assets_text = ", ".join(asset_names) if isinstance(asset_names, list) and asset_names else "No assets listed"
        if latest_version and core.is_newer_version(latest_version, core.APP_VERSION):
            installer_asset = core.choose_release_installer_asset(release_info)
            network_text = core.describe_network_profile(network_profile or {}) if network_profile is not None else "network state unavailable"
            self._set_update_status(f"Update available: {latest_tag or latest_version} (installed: {core.APP_VERSION}; {network_text})")
            if installer_asset is not None:
                if QMessageBox.question(
                    self,
                    "Update Available",
                    "A newer version is available\n\n"
                    f"Installed: {core.APP_VERSION}\n"
                    f"Latest: {latest_tag or latest_version}\n"
                    f"Published: {published_at or 'Unknown'}\n"
                    f"Assets: {assets_text}\n"
                    f"Network: {network_text}\n\n"
                    "Download the installer to this PC now?",
                ) == QMessageBox.Yes:
                    self._start_update_download(release_info, manual=manual)
                elif manual:
                    QMessageBox.information(self, "Check for Updates", "You can install later from the release page:\n\n" + html_url)
            elif manual:
                QMessageBox.information(
                    self,
                    "Check for Updates",
                    "A newer version is available\n\n"
                    f"Installed: {core.APP_VERSION}\n"
                    f"Latest: {latest_tag or latest_version}\n"
                    f"Published: {published_at or 'Unknown'}\n"
                    f"Assets: {assets_text}\n"
                    f"Network: {network_text}\n\n"
                    "No downloadable installer asset was found on the release.\n\n"
                    f"Release page:\n{html_url}",
                )
            return
        self._set_update_status(f"Installed version: {core.APP_VERSION} (up to date)")
        if manual:
            QMessageBox.information(
                self,
                "Check for Updates",
                "You are up to date.\n\n"
                f"Installed: {core.APP_VERSION}\n"
                f"Latest release: {latest_tag or latest_version or 'Unknown'}",
            )

    def _start_update_download(self, release_info: dict[str, object], manual: bool) -> None:
        if self._update_download_in_progress:
            return
        self._update_download_in_progress = True
        latest_tag = str(release_info.get("tag_name", "")).strip() or str(release_info.get("version", "")).strip() or "update"
        self._set_update_status(f"Downloading update {latest_tag}...")
        self.progress_bar.setRange(0, 1000)
        self.progress_bar.setValue(0)
        self.percent_label.setText("0.0%")
        self.loading_label.setText("Downloading...")
        threading.Thread(target=self._download_update_worker, args=(release_info, manual), daemon=True).start()

    def _download_update_worker(self, release_info: dict[str, object], manual: bool) -> None:
        try:
            installer_asset = core.choose_release_installer_asset(release_info)
            if installer_asset is None:
                raise RuntimeError("No installer asset was found in the latest GitHub release.")
            asset_name = str(installer_asset.get("name", "")).strip()
            download_url = str(installer_asset.get("download_url", "")).strip()
            if not asset_name or not download_url:
                raise RuntimeError("The release installer asset is missing its download URL.")
            destination = self.updates_dir / asset_name
            core.download_release_asset(
                download_url,
                destination,
                progress_callback=lambda downloaded, total: self.updateDownloadProgressReady.emit(downloaded, total or 0),
            )
            checksum_verified = False
            checksum_asset = core.choose_release_checksum_asset(release_info)
            if checksum_asset is not None:
                checksum_url = str(checksum_asset.get("download_url", "")).strip()
                if checksum_url:
                    checksum_text = core.download_url_bytes(checksum_url, timeout=60).decode("utf-8", errors="replace")
                    expected_checksum = core.checksum_for_asset_name(checksum_text, asset_name)
                    if expected_checksum and core.sha256_file(destination) != expected_checksum:
                        raise RuntimeError("Downloaded update failed SHA-256 verification.")
                    checksum_verified = bool(expected_checksum)
            self.updateDownloadSuccessReady.emit(release_info, destination, checksum_verified, manual)
        except Exception as exc:
            self.updateDownloadErrorRaised.emit(str(exc), manual)

    def _handle_update_download_error(self, message: str, manual: bool) -> None:
        self._update_download_in_progress = False
        self._set_update_status(f"Update download failed for installed version {core.APP_VERSION}")
        self.progress_bar.setRange(0, 1000)
        self.progress_bar.setValue(0)
        self.percent_label.setText("0.0%")
        self.loading_label.setText("")
        if manual:
            QMessageBox.critical(self, "Update Download", f"Could not download the update.\n\n{message}")

    def _handle_update_download_progress(self, downloaded: int, total: int) -> None:
        if not self._update_download_in_progress:
            return
        if total > 0:
            fraction = min(max(downloaded / total, 0.0), 1.0)
            self.progress_bar.setRange(0, 1000)
            self.progress_bar.setValue(int(round(fraction * 1000)))
            self.percent_label.setText(f"{fraction * 100:.1f}%")
            self.loading_label.setText("Downloading...")
            self._set_update_status(
                f"Downloading update... {core.fmt_hours(downloaded / (1024 * 1024))} / {core.fmt_hours(total / (1024 * 1024))} MB"
            )
        else:
            self.progress_bar.setRange(0, 0)
            self.percent_label.setText("...")
            self.loading_label.setText("Downloading...")
            self._set_update_status(f"Downloading update... {core.fmt_hours(downloaded / (1024 * 1024))} MB")

    def _handle_update_download_success(self, release_info: dict[str, object], destination: Path, checksum_verified: bool, manual: bool) -> None:
        self._update_download_in_progress = False
        self._downloaded_update_path = destination
        latest_tag = str(release_info.get("tag_name", "")).strip() or str(release_info.get("version", "")).strip() or destination.name
        verification_text = " and verified" if checksum_verified else ""
        self._set_update_status(f"Downloaded update {latest_tag}{verification_text}: {destination.name}")
        self.progress_bar.setRange(0, 1000)
        self.progress_bar.setValue(1000)
        self.percent_label.setText("100.0%")
        self.loading_label.setText("")
        if QMessageBox.question(
            self,
            "Install Update",
            "The update installer has been downloaded.\n\n"
            f"Release: {latest_tag}\n"
            f"File: {destination.name}\n"
            f"Saved to: {destination}\n"
            f"Checksum verified: {'Yes' if checksum_verified else 'No checksum asset found'}\n\n"
            f"Install it now? The {core.DISPLAY_APP_NAME} window will close first.",
        ) == QMessageBox.Yes:
            self._launch_update_installer(destination)
        elif manual:
            QMessageBox.information(self, "Update Download", f"The installer is ready at:\n{destination}")

    def _updater_exe_for_handoff(self) -> Path | None:
        if os.name != "nt":
            return None
        sibling = Path(sys.executable).resolve().parent / core.UPDATER_EXE_NAME
        if sibling.is_file():
            return sibling
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            bundled = Path(sys._MEIPASS) / core.UPDATER_EXE_NAME
            if bundled.is_file():
                dest = self.app_root / core.UPDATER_EXE_NAME
                try:
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(bundled, dest)
                except OSError:
                    return None
                return dest if dest.is_file() else None
        return None

    def _resolve_updater_handoff_argv(self, installer_path: Path) -> list[str] | None:
        inst = installer_path.resolve()
        if os.name != "nt":
            return None
        parent_exe = str(Path(sys.executable).resolve())
        try:
            from dss_tools_updater import get_process_image_path_windows

            live = get_process_image_path_windows(os.getpid())
            if live:
                parent_exe = live
        except Exception:
            pass
        updater_exe = self._updater_exe_for_handoff()
        if updater_exe is not None:
            return [str(updater_exe.resolve()), str(inst), str(os.getpid()), parent_exe]
        if not getattr(sys, "frozen", False):
            dev_script = Path(__file__).resolve().parent / "dss_tools_updater.py"
            if dev_script.is_file():
                return [sys.executable, str(dev_script), str(inst), str(os.getpid()), parent_exe]
        return None

    def _stage_installer_for_updater(self, installer_path: Path) -> Path:
        inst = installer_path.expanduser().resolve()
        try:
            inst.relative_to(self.app_root.resolve())
        except ValueError:
            return inst
        fd, staged = tempfile.mkstemp(prefix="dss_tools_setup_", suffix=".exe", dir=str(tempfile.gettempdir()))
        os.close(fd)
        staged_path = Path(staged)
        try:
            shutil.copy2(inst, staged_path)
        except OSError:
            try:
                staged_path.unlink(missing_ok=True)
            except OSError:
                pass
            raise
        return staged_path

    def _stage_updater_exe_to_temp(self, updater_src: Path) -> Path:
        fd, staged = tempfile.mkstemp(prefix="dss_tools_updater_", suffix=".exe", dir=str(tempfile.gettempdir()))
        os.close(fd)
        dest = Path(staged)
        shutil.copy2(updater_src, dest)
        return dest

    def _launch_update_installer(self, installer_path: Path) -> None:
        if not installer_path.exists():
            QMessageBox.critical(self, "Install Update", f"The downloaded installer could not be found.\n\n{installer_path}")
            return
        try:
            inst = self._stage_installer_for_updater(Path(installer_path)).resolve()
        except OSError as exc:
            QMessageBox.critical(self, "Install Update", f"Could not stage the installer for the update helper.\n\n{exc}")
            return
        argv = self._resolve_updater_handoff_argv(inst)
        if argv is not None:
            if os.name == "nt" and len(argv) >= 1 and Path(argv[0]).suffix.lower() == ".exe" and core._is_updater_executable_path(Path(argv[0])):
                try:
                    staged_u = self._stage_updater_exe_to_temp(Path(argv[0]).resolve())
                except OSError as exc:
                    QMessageBox.critical(self, "Install Update", f"Could not stage the update helper.\n\n{exc}")
                    return
                argv = [str(staged_u), *argv[1:]]
            try:
                if os.name == "nt" and Path(argv[0]).suffix.lower() == ".exe" and core._is_updater_executable_path(Path(argv[0])):
                    core._shell_execute_runas_windows(
                        argv[0],
                        subprocess.list2cmdline(argv[1:]),
                        str(inst.parent),
                        show_cmd=1,
                    )
                else:
                    creationflags = 0
                    if os.name == "nt":
                        creationflags = (
                            getattr(subprocess, "DETACHED_PROCESS", 0)
                            | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
                            | getattr(subprocess, "CREATE_NO_WINDOW", 0)
                        )
                    subprocess.Popen(
                        argv,
                        stdin=subprocess.DEVNULL,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        close_fds=False,
                        cwd=str(inst.parent),
                        creationflags=creationflags,
                    )
            except OSError as exc:
                QMessageBox.critical(self, "Install Update", f"Could not start the update helper.\n\n{exc}")
                return
            QApplication.instance().quit()
            return
        pid = os.getpid()
        escaped_path = str(inst).replace("'", "''")
        command = f"while (Get-Process -Id {pid} -ErrorAction SilentlyContinue) {{ Start-Sleep -Milliseconds 500 }}; Start-Process -FilePath '{escaped_path}'"
        try:
            subprocess.Popen(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command", command],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                stdin=subprocess.DEVNULL,
                close_fds=False,
                **core._windows_hidden_subprocess_options(),
            )
        except OSError as exc:
            QMessageBox.critical(self, "Install Update", f"Could not launch the installer.\n\n{exc}")
            return
        QApplication.instance().quit()

    def _start_hash_poll_timer(self) -> None:
        self.hash_poll_timer.setInterval(max(1, self.app_settings.hash_poll_minutes) * 60 * 1000)
        self.hash_poll_timer.timeout.connect(self._poll_source_hashes)
        self.hash_poll_timer.start()

    def _poll_source_hashes(self) -> None:
        if not self.current_data or self.load_worker is not None or self.outlook_worker is not None:
            return
        changed_paths: list[Path] = []
        for path in self.current_data.source_paths:
            try:
                workbook_bytes = core.read_source_bytes(path)
                content_hash = core.compute_workbook_content_hash(workbook_bytes)
            except Exception:
                continue
            if self.current_data.file_hashes.get(path) and self.current_data.file_hashes.get(path) != content_hash:
                changed_paths.append(path)
        if changed_paths:
            unseen = [path for path in changed_paths if path not in self._hash_alerted_paths]
            if not unseen:
                return
            self._hash_alerted_paths.update(unseen)
            names = ", ".join(path.name for path in unseen[:10])
            self.loading_label.setText("Changed")
            self.status_label.setText(f"DSS changed since last view: {names}. Press Update View to refresh.")
            QMessageBox.information(
                self,
                "Source DSS Changed",
                "One or more loaded DSS workbooks changed since the last inspected state.\n\nPress 'Update View' to refresh the summaries.",
            )


def launch_qt_app(initial_source: list[Path] | None = None) -> int:
    _configure_forced_qt_software_rendering()
    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setPalette(_build_forced_qt_palette(core.DEFAULT_UI_THEME))
    if os.name == "nt" and getattr(sys, "frozen", False):
        core._windows_set_explicit_app_user_model_id()
    window = DssQtMainWindow(initial_source=initial_source)
    window.show()
    return app.exec()
